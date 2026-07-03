"""
厚生労働省「医療用医薬品供給状況」Excelを取得して data.json に変換するスクリプト。
薬価基準収載品目リストから包装情報を取得して package_info.json に保存する。
GitHub Actions から毎日実行される。
"""
import requests
import openpyxl
import json
import os
import sys
from io import BytesIO
from datetime import date, timedelta
from bs4 import BeautifulSoup

MHLW_PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
MHLW_BASE = "https://www.mhlw.go.jp"
OUTPUT_FILE = "data.json"
CHANGES_FILE = "changes.json"
PACKAGE_FILE = "package_info.json"

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


def cell_to_str(v):
    """セルの値を文字列に変換（日付・シリアル数値含む）"""
    if v is None:
        return ""
    if hasattr(v, "strftime"):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, (int, float)) and 40000 < v < 60000:
        return (date(1899, 12, 30) + timedelta(days=int(v))).isoformat()
    return str(v)


def parse_excel(content):
    """Excel を読み込み、データ行の配列を返す"""
    wb = openpyxl.load_workbook(BytesIO(content), data_only=True, read_only=True)
    ws = wb.active

    rows = []
    header_found = False

    for raw_row in ws.iter_rows(values_only=True):
        if len(raw_row) <= COL_STATUS:
            continue
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
            row = [cell_to_str(v) for v in raw_row[:16]]
            # 16列未満の行はパディング
            while len(row) < 16:
                row.append("")
            rows.append(row)

    wb.close()
    print(f"データ行数: {len(rows):,}")
    return rows


# ============================================================
# 薬価基準収載品目リスト → package_info.json
# ============================================================

def _find_yakka_excel_urls():
    """薬価基準品目リストExcelのURLを直接構築して存在確認する。
    ハブページは403が多いため、既知パターンから直接試す。
    _01=内服薬, _02=注射薬, _03=外用薬, _04=歯科用薬, _06=改定品目
    _05=後発品有無情報（包装列なし）は除外。
    """
    today = date.today()
    BASE = "https://www.mhlw.go.jp/topics"
    found = []

    for year in [today.year, today.year - 1, today.year - 2]:
        year_found = []
        for nn in ["01", "02", "03", "04", "06"]:
            url = f"{BASE}/{year}/04/xls/tp{year}0401-01_{nn}.xlsx"
            try:
                # HEADは弾かれるためGET+streamで存在確認のみ
                resp = requests.get(url, headers=HEADERS, timeout=15, stream=True)
                resp.close()
                print(f"  {resp.status_code} {url.split('/')[-1]}")
                if resp.status_code == 200:
                    year_found.append((f"{year}年度_{nn}", url))
            except Exception as e:
                print(f"  エラー {url.split('/')[-1]}: {e}")

        if year_found:
            print(f"  → {year}年度: {len(year_found)}件")
            found.extend(year_found)
            break  # 最新年度が見つかればその年で完了

    if not found:
        print("直接URL確認失敗。ハブページをフォールバック検索中...")
        for year in [today.year, today.year - 1]:
            page_url = f"{BASE}/{year}/04/tp{year}0401-01.html"
            try:
                resp = requests.get(page_url, headers=HEADERS, timeout=20)
                print(f"  {resp.status_code} {page_url}")
                if not resp.ok:
                    continue
                soup = BeautifulSoup(resp.text, "html.parser")
                for a in soup.find_all("a", href=True):
                    href = a["href"]
                    if not href.lower().endswith(".xlsx") or href.endswith("_05.xlsx"):
                        continue
                    full_url = MHLW_BASE + href if href.startswith("/") else href
                    found.append((a.get_text(strip=True), full_url))
                if found:
                    print(f"  → {len(found)}件")
                    break
            except Exception as e:
                print(f"  エラー {page_url}: {e}")

    return found


def _parse_yakka_excel(url):
    """薬価基準ExcelからYJコード→包装のマッピングを抽出する"""
    try:
        print(f"  薬価基準Excel取得: {url}")
        resp = requests.get(url, headers=HEADERS, timeout=120)
        resp.raise_for_status()

        wb = openpyxl.load_workbook(BytesIO(resp.content), data_only=True, read_only=True)
        package_map = {}
        print(f"  シート数: {len(wb.worksheets)}")

        YJ_LABELS = {"YJコード", "ＹＪコード", "ＹＪコ－ド", "YJ코드", "yjコード",
                     "ＹＪ", "YJ", "薬価基準収載医薬品コード"}
        PKG_LABELS = {"包装", "包　装", "包 装"}

        for ws in wb.worksheets:
            yj_col = pkg_col = None
            header_row_idx = 0

            for row_idx, raw_row in enumerate(ws.iter_rows(values_only=True)):
                cells = [str(c or "") for c in raw_row]
                norm_cells = [c.replace("　", "").replace(" ", "").replace("\n", "") for c in cells]

                # ヘッダー行を探す（最初の30行以内）
                if yj_col is None and row_idx < 30:
                    for j, norm in enumerate(norm_cells):
                        if norm in YJ_LABELS:
                            yj_col = j
                        if norm in PKG_LABELS:
                            pkg_col = j
                    if yj_col is not None and pkg_col is not None:
                        header_row_idx = row_idx
                        print(f"  シート「{ws.title}」: YJコード列={yj_col}, 包装列={pkg_col} (行{row_idx})")
                    continue

                if yj_col is None or pkg_col is None:
                    continue
                if len(raw_row) <= max(yj_col, pkg_col):
                    continue

                yj  = norm_cells[yj_col].strip()
                pkg = cells[pkg_col].strip()
                if yj and 10 <= len(yj) <= 14 and pkg:
                    package_map[yj] = pkg

        wb.close()
        print(f"  → {len(package_map):,} 件抽出")
        return package_map

    except Exception as e:
        print(f"  薬価基準Excel解析失敗: {e}")
        return {}


def fetch_package_info():
    """薬価基準収載品目リストから包装情報を取得して package_info.json に保存する"""
    print("\n--- 包装情報取得 ---")

    # 既存ファイルがあれば読み込んでおく（差分更新）
    existing = {}
    if os.path.exists(PACKAGE_FILE):
        try:
            with open(PACKAGE_FILE, encoding="utf-8") as f:
                existing = json.load(f)
            print(f"既存package_info.json: {len(existing):,} 件")
        except Exception:
            pass

    excel_candidates = _find_yakka_excel_urls()
    if not excel_candidates:
        print("薬価基準ExcelのURLが見つかりませんでした（スキップ）")
        return

    merged = dict(existing)
    for label, url in excel_candidates:
        pkg_map = _parse_yakka_excel(url)
        if pkg_map:
            merged.update(pkg_map)

    if not merged:
        print("包装情報取得ゼロ件（スキップ）")
        return

    with open(PACKAGE_FILE, "w", encoding="utf-8") as f:
        json.dump(merged, f, ensure_ascii=False, separators=(",", ":"))

    print(f"package_info.json 保存: {len(merged):,} 件")


# ============================================================

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
        return "improved"
    elif new_ord < old_ord:
        return "worsened"
    return "changed"


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

    history = [h for h in history if h.get("date") != today]
    history.insert(0, {"date": today, "changes": changes})
    history = history[:180]

    with open(CHANGES_FILE, "w", encoding="utf-8") as f:
        json.dump({"history": history}, f, ensure_ascii=False, separators=(",", ":"))

    print(f"変化記録: {len(changes)} 件 → {CHANGES_FILE}")


def main():
    try:
        old_map = load_previous_status_map()

        xlsx_url, filename = find_excel_url()
        content = download_excel(xlsx_url)
        rows = parse_excel(content)
        today = date.today().isoformat()

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

        # 包装情報を薬価基準から取得
        fetch_package_info()

    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
