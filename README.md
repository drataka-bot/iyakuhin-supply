# 医薬品供給状況検索システム

本番サイト: https://drataka-bot.github.io/iyakuhin-supply/

## 仕組み

| ファイル | 役割 |
|---|---|
| `index.html` | 検索UI（ウォッチリスト・更新履歴・AND/OR検索） |
| `data.json` | 医薬品供給データ（毎日自動更新） |
| `changes.json` | 出荷状況の変化履歴（毎日自動生成） |
| `scripts/fetch_data.py` | 厚労省からデータを取得するスクリプト |

## データ自動更新

GitHub Actionsが毎日11:00 JSTに厚生労働省のExcelを取得し、`data.json`と`changes.json`を自動更新します。
