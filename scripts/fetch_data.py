"""
厚生労働省「医療用医薬品供給状況」Excelを取得して data.json に変換するスクリプト。
GitHub Actions から毎日実行される。
"""
import requests
import openpyxl
import json
import re
import sys
from io import BytesIO
from datetime import date
from bs4 import BeautifulSoup

MHLW_PAGE = "https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kenkou_iryou/iryou/kouhatu-iyaku/04_00003.html"
MHLW_BASE = "https://www.mhlw.go.jp"
OUTPUT_FILE = "data.json"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (compatible; DataFetcher/1.0)"
}

# 列インデックス（0始まり）
COL_STATUS = 11   # ⑫出荷対応の状況
COL_BRAND  = 5    # ⑥品名
COL_GENERIC = 2   # ③成分名


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


def main():
    try:
        xlsx_url, filename = find_excel_url()
        content = download_excel(xlsx_url)
        rows = parse_excel(content)

        result = {
            "fetchDate": date.today().isoformat(),
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
