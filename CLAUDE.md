# 酷澎訂單管理系統（coupang-oms）

**開工前依序讀：**
1. [`coupang-oms/README.md`](coupang-oms/README.md) —— 架構、核心防呆、部署方式
2. [`交接筆記.md`](交接筆記.md) —— 目前狀態、踩過的地雷

這份只放最容易踩錯的幾條，細節不重複。

## ⛔ 這個分支就是正式站的部署來源

**目前分支 `claude/coupang-order-system-design-n17g7z` 是 Render 接的分支。
`git push` 上去會直接自動部署到同事在用的正式站。**

推之前一定要：
1. 兩支測試全過（見下方）
2. 升 `coupang-oms/app.py` 最上面的 `BUILD_VERSION`，否則無法確認部署有沒有生效
3. 推完去 Render 後台看部署是不是綠燈，並打開正式站確認右下角版本號

`main` 不是部署來源，只是參考快照。

## 這是什麼

Flask 應用。把酷澎整合表整份上傳，自動去重、抓出酷澎偷改過的單、記錄誰在何時改了什麼，
一鍵匯出乾淨的 Excel。另有一個**完全獨立**的「採購表轉換」工具（`purchase.py`，網址 `/purchase`）。

正式站跑在 Render + Supabase PostgreSQL，不是本機程式，不需要誰的電腦開著。

⚠️ 這個 repo 根目錄還有 SOP 檢索網站（`index.html`、`img/`），**跟本系統完全無關**，
它有自己的分支與工作目錄 `../sop`。**不要在這個目錄動根目錄那些檔案。**

## 紅線

- **PostgreSQL 模式下 SQL 字串裡不能有裸的 `%`。** `LIKE 'CPG-%'` 會炸成
  `IndexError: tuple index out of range`，而且**本機 SQLite 完全測不出來**。
  LIKE 樣式一定要綁參數：`conn.execute("... LIKE ?", ("CPG-%",))`
- **本機走 SQLite、正式站走 PostgreSQL，行為有差。** 改到 `db.py` 或任何 SQL，
  都要想「這在 Postgres 上會怎樣」。
- **唯一鍵是 `UNIQUE(po_number, sku_id)`**，不能把 `po_number` 單獨設 UNIQUE
  （一張 PO 最多 29 個 SKU，設錯會靜悄悄丟掉 80% 資料）。
- **`warehouse` 和 `delivery_date` 不能進主鍵**，會產生數量翻倍的幽靈列。
- **改了 Tailwind class 一定要重新編譯 CSS**，否則畫面完全沒效果且不報錯：
  `cd coupang-oms/.build-tools && npx tailwindcss -i input.css -o ../static/tailwind.css --config tailwind.config.js --minify`

## 開發與測試

```bash
cd coupang-oms
pip install -r requirements.txt
python app.py            # 或 Windows 雙擊 START.bat，開 http://127.0.0.1:5000

python test_flow.py      # 訂單管理主流程
python test_purchase.py  # 採購表轉換
```

本機模式資料存在程式資料夾**外面**的「資料與設定」資料夾。更新程式可整包覆蓋，
**要備份的是「資料與設定」**。

兩支測試都必須全過才能推。用 `samples/` 的真實資料跑端到端，不是 mock。
