# OpsPilot PDF Parser

這是 OpsPilot 的後端 PDF 解析服務。目的不是修改使用者 PDF，而是把原始 ERP PDF 在伺服器端轉成穩定 JSON，避免 iPhone Safari / PDF.js 的文字切分差異。

## 已支援
- 庫存調整單
- `B-Breakage`
- `E-Expiry`
- 逐筆 SKU、品名、調整日、Ref #、最後進價、調整量、調整金額、調整前後庫存、差異 %

## Render 部署
1. 建立新的 GitHub repo，把本資料夾 4 個檔案放進 repo 根目錄。
2. Render → New → Blueprint，選該 repo。
3. Render 會讀取 `render.yaml` 自動部署。
4. 部署後測試：`https://你的服務.onrender.com/health`
5. 把服務網址填回 OpsPilot `index.html` 的：
   `<meta name="opspilot-parser-api" content="https://你的服務.onrender.com">`

前端會優先呼叫 `/parse-pdf`；如果 API 沒設定、逾時、或回傳 422（尚未支援的 PDF），就自動退回原本前端解析器。
