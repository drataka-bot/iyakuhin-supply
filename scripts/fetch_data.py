"""
厚生労働省「医療用医薬品供給状況」Excelを取得して data.json に変換するスクリプト。
MEDIS 医薬品HOTコードマスターから包装情報を取得して package_info.json に保存する。
GitHub Actions から毎日実行される。
"""
import requests
import openpyxl
import json
import os
import re
import csv
import sys
import zipfile
from io import BytesIO, StringIO
from datetime import date, timedelta
from urllib.parse import urljoin
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
# MEDIS 医薬品HOTコードマスター → package_info.json
#
# 薬価基準収載品目リストには包装列が無いため、包装情報は
# MEDISのHOT13マスター（月次全件・無償公開）から取得する。
# CSVはヘッダーなし・Shift-JIS・24列:
#   列8(idx7)  個別医薬品コード（YJコード）
#   列15(idx14) 包装形態（PTP/バラ等）
#   列16(idx15) 包装単位数   列17(idx16) 包装単位単位
#   列18(idx17) 包装総量数   列19(idx18) 包装総量単位
# ============================================================

MEDIS_HOT_PAGES = [
    "https://www2.medis.or.jp/master/hcode/",
    "https://www2.medis.or.jp/hcode/",
]

YJ_RE = re.compile(r"^\d{7}[A-Za-z]\d{4}$")


def _find_hot_zip_urls():
    """MEDISのダウンロードページから全件マスターzipのURLを収集する"""
    for page_url in MEDIS_HOT_PAGES:
        try:
            resp = requests.get(page_url, headers=HEADERS, timeout=30)
            print(f"  {resp.status_code} {page_url}")
            if not resp.ok:
                continue
            resp.encoding = resp.apparent_encoding
            soup = BeautifulSoup(resp.text, "html.parser")
            found = []
            for a in soup.find_all("a", href=True):
                href = a["href"]
                if not href.lower().endswith(".zip"):
                    continue
                full_url = urljoin(page_url, href)
                label = a.get_text(strip=True)
                found.append((label, full_url))
            if found:
                # HOT13/全件らしきものを優先、HOT9・廃止・日次差分は後回し
                def _priority(item):
                    label, url = item
                    name = url.split("/")[-1].lower() + label
                    score = 0
                    if "13" in name:
                        score -= 2
                    if "全" in label:
                        score -= 1
                    if "9" in name or "廃止" in label or "del" in name.lower():
                        score += 2
                    return score
                found.sort(key=_priority)
                print(f"  → zipリンク {len(found)}件: {[u.split('/')[-1] for _, u in found[:8]]}")
                return found[:10]
        except Exception as e:
            print(f"  エラー {page_url}: {e}")
    return []


def _build_pkg_str(keitai, unit_num, unit_unit, total_num, total_unit):
    """包装形態・単位数・総量から「PTP 10錠×10」形式の文字列を組み立てる"""
    keitai = keitai.strip()
    unit_num = unit_num.strip()
    unit_unit = unit_unit.strip()
    total_num = total_num.strip()
    total_unit = total_unit.strip()

    def _num(s):
        try:
            return float(s.replace(",", ""))
        except ValueError:
            return None

    def _fmt(f):
        return str(int(f)) if f == int(f) else str(f)

    u = _num(unit_num)
    t = _num(total_num)

    if u and t and t > u and (t / u) == int(t / u):
        base = f"{_fmt(u)}{unit_unit}×{int(t / u)}"
    elif t:
        base = f"{_fmt(t)}{total_unit}"
    elif u:
        base = f"{_fmt(u)}{unit_unit}"
    else:
        return ""

    return f"{keitai} {base}".strip() if keitai else base


def _parse_hot_zip(url, yj_filter):
    """HOTマスターzipをダウンロードしてYJコード→包装リスト(dict)を返す"""
    try:
        print(f"  ダウンロード: {url}")
        resp = requests.get(url, headers=HEADERS, timeout=300)
        resp.raise_for_status()
        print(f"  {len(resp.content):,} bytes")

        package_map = {}
        with zipfile.ZipFile(BytesIO(resp.content)) as zf:
            names = [n for n in zf.namelist() if n.lower().endswith((".csv", ".txt"))]
            print(f"  zip内ファイル: {names}")
            for name in names:
                with zf.open(name) as f:
                    text = f.read().decode("cp932", errors="replace")
                sub_map = _parse_hot_rows(csv.reader(StringIO(text)), yj_filter)
                for yj, pkgs in sub_map.items():
                    lst = package_map.setdefault(yj, [])
                    for p in pkgs:
                        if p not in lst:
                            lst.append(p)
                print(f"  {name}: {len(sub_map):,} YJコード")
        return package_map
    except Exception as e:
        print(f"  HOTマスター解析失敗: {e}")
        return {}


# jp-medicine-master-data ミラー（MEDIS HOT13を毎月UTF-8/ヘッダー付きCSVで再配布）
HOT_CATALOG_URL = "https://raw.githubusercontent.com/shiro46mt/jp-medicine-master-data/main/data/data_catalog.json"
HOT_DATA_BASE = "https://raw.githubusercontent.com/shiro46mt/jp-medicine-master-data/main/data/hot13/"


def _parse_hot_rows(reader, yj_filter):
    """CSV行イテレータからYJコード→包装リスト(dict)を抽出する"""
    package_map = {}
    for row in reader:
        if len(row) < 19:
            continue
        yj = (row[7] or "").strip()
        if not YJ_RE.match(yj):
            continue  # ヘッダー行・不正行はここで除外
        if yj_filter and yj not in yj_filter:
            continue
        if (row[14] or "").strip() == "調剤用":
            continue  # 調剤包装単位は表示対象外
        pkg = _build_pkg_str(row[14] or "", row[15] or "", row[16] or "",
                             row[17] or "", row[18] or "")
        if not pkg:
            continue
        lst = package_map.setdefault(yj, [])
        if pkg not in lst:
            lst.append(pkg)
    return package_map


def _fetch_hot13_from_mirror(yj_filter):
    """GitHubミラーからHOT13最新CSVを取得して解析する（動作確認済みの主経路）"""
    resp = requests.get(HOT_CATALOG_URL, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    catalog = resp.json()
    files = next(d for d in catalog["data"] if d["id"] == "hot13")["files"]
    latest = sorted(files)[-1]
    print(f"  HOT13最新ファイル: {latest} (カタログ更新: {catalog.get('update')})")

    resp = requests.get(HOT_DATA_BASE + latest, headers=HEADERS, timeout=300)
    resp.raise_for_status()
    print(f"  ダウンロード完了: {len(resp.content):,} bytes")

    return _parse_hot_rows(csv.reader(StringIO(resp.text)), yj_filter)


def fetch_package_info():
    """MEDIS HOTマスターから包装情報を取得して package_info.json に保存する"""
    print("\n--- 包装情報取得（MEDIS HOTマスター）---")

    # 供給状況データに存在するYJコードだけに絞ってファイルを小さく保つ
    yj_filter = set()
    try:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            for row in json.load(f).get("rows", []):
                if len(row) > COL_YJ and row[COL_YJ]:
                    yj_filter.add(row[COL_YJ])
        print(f"対象YJコード: {len(yj_filter):,} 件")
    except Exception as e:
        print(f"data.json 読み込み失敗（全件対象）: {e}")

    # 主経路: GitHubミラー（ローカルで動作確認済み）
    merged = {}
    try:
        merged = _fetch_hot13_from_mirror(yj_filter)
        print(f"  ミラーから {len(merged):,} YJコード分の包装を取得")
    except Exception as e:
        print(f"  ミラー取得失敗: {e}")

    # フォールバック: MEDIS公式サイトのzip
    if len(merged) < 1000:
        print("  フォールバック: MEDIS公式サイトを試行")
        for label, url in _find_hot_zip_urls():
            pkg_map = _parse_hot_zip(url, yj_filter)
            for yj, pkgs in pkg_map.items():
                lst = merged.setdefault(yj, [])
                for p in pkgs:
                    if p not in lst:
                        lst.append(p)
            if len(merged) > 1000:
                break

    if not merged:
        print("包装情報取得ゼロ件（スキップ）")
        return

    out = {yj: "、".join(pkgs[:6]) for yj, pkgs in merged.items()}
    with open(PACKAGE_FILE, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))
    print(f"package_info.json 保存: {len(out):,} 件")


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
    """changes.json に今日の記録を追記する（変化なしの日も稼働確認として残す・最大180日分保持）"""
    history = []
    if os.path.exists(CHANGES_FILE):
        try:
            with open(CHANGES_FILE, encoding="utf-8") as f:
                history = json.load(f).get("history", [])
        except Exception:
            history = []

    # 同じ日の既存エントリを確認（同日複数回実行への対応）
    todays_existing = next((h for h in history if h.get("date") == today), None)

    if not changes:
        # 変化なし。既に今日「変化あり」で記録済みなら上書きしない
        if todays_existing and todays_existing.get("changes"):
            print("変化なし: 本日は既に変化を記録済みのため据え置き")
            return
        entry = {"date": today, "changes": []}
        print(f"変化なし: 稼働確認として記録 → {CHANGES_FILE}")
    else:
        entry = {"date": today, "changes": changes}
        print(f"変化記録: {len(changes)} 件 → {CHANGES_FILE}")

    history = [h for h in history if h.get("date") != today]
    history.insert(0, entry)
    history = history[:180]

    with open(CHANGES_FILE, "w", encoding="utf-8") as f:
        json.dump({"history": history}, f, ensure_ascii=False, separators=(",", ":"))


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

        # 包装情報をMEDIS HOTマスターから取得
        fetch_package_info()

    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
