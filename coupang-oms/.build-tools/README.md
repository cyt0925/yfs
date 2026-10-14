# 樣式編譯工具（開發用，OP 平常不需要碰這個資料夾）

`static/tailwind.css`、`static/bootstrap-icons.css` 跟 `static/fonts/`
是編譯好、可以直接用的靜態檔案，已經包含在專案裡，一般更新程式碼
（解壓縮覆蓋 `coupang-oms` 資料夾）不需要重新編譯。

只有在改了 `templates/index.html` 裡用到的 Tailwind class、或改了
`tailwind.config.js` 裡的顏色設定之後，才需要重新編譯：

```bash
cd .build-tools
npm install          # 第一次要先裝一次套件
npx tailwindcss -i input.css -o ../static/tailwind.css \
    --config tailwind.config.js --minify
```

這個資料夾（`.build-tools/`，包含 `node_modules/`）只在開發機器上需要，
不需要跟著 zip 一起發給 OP 用的電腦——那台電腦不需要裝 Node.js，
只要有編譯好的 `static/tailwind.css` 就能正常顯示畫面。
