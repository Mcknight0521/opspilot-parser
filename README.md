# OpsPilot Universal Parser v3

OpsPilot 後端通用匯入與驗算 API。

## 支援檔案
- PDF
- XLSX / XLS / XLSM / ODS
- CSV / TSV / TXT

## API
- `GET /health`：服務狀態與支援格式
- `POST /parse-file`：所有檔案統一入口（multipart/form-data，欄位名 `file`）
- `POST /parse-pdf`：舊版相容入口

## 目前內建的已知報表模板
1. 庫存調整單 PDF
2. 每日銷售報表－依單品 PDF / XLSX / XLS / XLSM / ODS / CSV / TSV / TXT

已知模板會做欄位與數學交叉驗算；總計不一致時會回 HTTP 422，避免錯誤數字進入正式分析。

其他表格格式會嘗試「通用欄位映射」：日期、SKU、品名、銷售量、營業額、報廢、出清等。若沒有原始總計可核對，會標示 `validation.status = partial` 與 `requiresReview = true`，不假裝成完整驗證。

## Render
Build command:
`pip install -r requirements.txt`

Start command:
`uvicorn app:app --host 0.0.0.0 --port $PORT`
