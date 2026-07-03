# 医薬品供給状況検索システム

## 機能の更新を本番サイトに反映する手順

### 1. このURLを開く

```
https://github.com/drataka-bot/iyakuhin-supply/compare/main...claude/build-nyuka-now-xYKfB
```

### 2. 「Create pull request」ボタンをクリック

### 3. 「Merge pull request」→「Confirm merge」をクリック

以上で https://drataka-bot.github.io/iyakuhin-supply/ に反映されます（数分後）。

---

## 仕組み

| ファイル | 役割 |
|---|---|
| `index.html` | 検索UI（ウォッチリスト・更新履歴・AND/OR検索） |
| `data.json` | 医薬品供給データ（毎日自動更新） |
| `changes.json` | 出荷状況の変化履歴（毎日自動生成） |
| `scripts/fetch_data.py` | 厚労省からデータを取得するスクリプト |

## データ自動更新

GitHub Actionsが毎日11:00 JSTに厚生労働省のExcelを取得し、`data.json`と`changes.json`を自動更新します。
