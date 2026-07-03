"""
厚生労働省「医療用医薬品供給状況」Excelを取得して data.json に変換するスクリプト。
GitHub Actions から毎日実行される。
"""
import requests
import openpyxl
import json
import os
import sys
from io import BytesIO
from datetime import date
from bs4 import BeautifulSoup

MHLW_PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
MHLW_BASE = "https://www.mhlw.go.jp"
OUTPUT_FILE = "data.json"
CHANGES_FILE = "changes.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DataFetcher/1.0)"
}

# 列インデックス（0始まり）
COL_STATUS  = 11   # ⑫出荷対応の状況
COL_BRAND   = 5    # ⑥品名
COL_GENERIC = 2    # ③成分名
COL_YJ      = 4    # ⑤YJコード
COL_MAKER   = 6    # ⑦製造販売業者名


def find_excel_url():
    """MHLWページから最新の xlsx URL を探す"""
    print(f"ページ取得: {MHLW_PAGE}")
    resp = requests.get(MHLW_PAGE, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    for a in soup.find_all("a", href=True):
        href = a["href"]
        if href.endswith(".xlsx") and "iyakuhin" in href.lower():
            full_url = MHLW_BASE + href if href.startswith("/") else href
            print(f"Excel URL 発見: {full_url}")
            return full_url, href.split("/")[-1]

    raise RuntimeError("Excel ファイルのリンクが見つかりませんでした")


def download_excel(url):
    print(f"Excel ダウンロード中...")
    resp = requests.get(url, headers=HEADERS, timeout=60)
    resp.raise_for_status()
    print(f"ダウンロード完了: {len(resp.content):,} bytes")
    return resp.content


def parse_excel(content):
    """Excel を読み込み、データ行の配列を返す"""
    wb = openpyxl.load_workbook(BytesIO(content), data_only=True, read_only=True)
    ws = wb.active

    rows = []
    header_found = False

    for raw_row in ws.iter_rows(values_only=True):
        status_val = str(raw_row[COL_STATUS] or "")

        # データ開始行を検出（出荷状況が①〜⑤で始まる行）
        if not header_found:
            if status_val and status_val[0] in "①②③④⑤":
                header_found = True

        if header_found:
            brand   = str(raw_row[COL_BRAND]   or "").strip()
            generic = str(raw_row[COL_GENERIC]  or "").strip()
            if not brand and not generic:
                continue
            # 各セルを文字列化（日付オブジェクト等も変換）
            row = []
            for v in raw_row[:16]:
                if v is None:
                    row.append("")
                elif hasattr(v, "strftime"):
                    row.append(v.strftime("%Y-%m-%d"))
                else:
                    row.append(str(v))
            rows.append(row)

    wb.close()
    print(f"データ行数: {len(rows):,}")
    return rows


def load_previous_status_map():
    """既存の data.json から YJコード→{status, brand, generic, maker} のマップを返す"""
    if not os.path.exists(OUTPUT_FILE):
        return {}
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            data = json.load(f)
        result = {}
        for row in data.get("rows", []):
            yj = row[COL_YJ] if len(row) > COL_YJ else ""
            if yj:
                result[yj] = {
                    "status":  row[COL_STATUS]  if len(row) > COL_STATUS  else "",
                    "brand":   row[COL_BRAND]   if len(row) > COL_BRAND   else "",
                    "generic": row[COL_GENERIC] if len(row) > COL_GENERIC else "",
                    "maker":   row[COL_MAKER]   if len(row) > COL_MAKER   else "",
                }
        return result
    except Exception as e:
        print(f"既存データ読み込みスキップ: {e}")
        return {}


STATUS_ORDER = {"①": 3, "②": 1, "③": 1, "④": 1, "⑤": 0}

def classify_change(old_status, new_status):
    """変化の種類を分類する"""
    old_ord = STATUS_ORDER.get(old_status[:1] if old_status else "", -1)
    new_ord = STATUS_ORDER.get(new_status[:1] if new_status else "", -1)
    if new_ord > old_ord:
        return "improved"   # 改善（限定→通常 など）
    elif new_ord < old_ord:
        return "worsened"   # 悪化（通常→限定 など）
    return "changed"        # 同カテゴリ内の変化


def detect_changes(old_map, new_rows):
    """新旧データを比較して変化リストを返す"""
    new_map = {}
    for row in new_rows:
        yj = row[COL_YJ] if len(row) > COL_YJ else ""
        if yj:
            new_map[yj] = {
                "status":  row[COL_STATUS]  if len(row) > COL_STATUS  else "",
                "brand":   row[COL_BRAND]   if len(row) > COL_BRAND   else "",
                "generic": row[COL_GENERIC] if len(row) > COL_GENERIC else "",
                "maker":   row[COL_MAKER]   if len(row) > COL_MAKER   else "",
            }

    changes = []

    # ステータス変化を検出
    for yj, new_item in new_map.items():
        if yj in old_map:
            old_status = old_map[yj]["status"]
            new_status = new_item["status"]
            if old_status != new_status:
                changes.append({
                    "yj":      yj,
                    "brand":   new_item["brand"],
                    "generic": new_item["generic"],
                    "maker":   new_item["maker"],
                    "from":    old_status,
                    "to":      new_status,
                    "type":    classify_change(old_status, new_status),
                })
        else:
            # 新規追加品目で通常出荷以外（限定・停止で新規追加）
            new_status = new_item["status"]
            if new_status and new_status[:1] not in ("①", ""):
                changes.append({
                    "yj":      yj,
                    "brand":   new_item["brand"],
                    "generic": new_item["generic"],
                    "maker":   new_item["maker"],
                    "from":    "",
                    "to":      new_status,
                    "type":    "new",
                })

    return changes


def update_changes_file(changes, today):
    """changes.json に今日の変化を追記する（最大180日分保持）"""
    if not changes:
        print("変化なし: changes.json の更新をスキップ")
        return

    history = []
    if os.path.exists(CHANGES_FILE):
        try:
            with open(CHANGES_FILE, encoding="utf-8") as f:
                history = json.load(f).get("history", [])
        except Exception:
            history = []

    # 同じ日付のエントリは上書き
    history = [h for h in history if h.get("date") != today]
    history.insert(0, {"date": today, "changes": changes})
    history = history[:180]  # 最大180日分

    with open(CHANGES_FILE, "w", encoding="utf-8") as f:
        json.dump({"history": history}, f, ensure_ascii=False, separators=(",", ":"))

    print(f"変化記録: {len(changes)} 件 → {CHANGES_FILE}")


def main():
    try:
        # 既存データを先に読み込む（上書き前に差分検出するため）
        old_map = load_previous_status_map()

        xlsx_url, filename = find_excel_url()
        content = download_excel(xlsx_url)
        rows = parse_excel(content)
        today = date.today().isoformat()

        # 差分を検出してchanges.jsonを更新
        if old_map:
            changes = detect_changes(old_map, rows)
            update_changes_file(changes, today)
        else:
            print("初回実行のため差分検出をスキップ")

        result = {
            "fetchDate": today,
            "source": filename,
            "rows": rows
        }

        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, separators=(",", ":"))

        size_kb = len(open(OUTPUT_FILE, encoding="utf-8").read()) // 1024
        print(f"保存完了: {OUTPUT_FILE} ({size_kb:,} KB, {len(rows):,} 件)")

    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
