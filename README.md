# OpsPilot Universal Parser v3.1

統一營運檔案解析與驗算 API。

支援：PDF、XLSX、XLS、XLSM、ODS、CSV、TSV、TXT。

## v3.1 重點
- 修正 Render Free 上傳 Excel 可能因記憶體壓力而 exit 139 / 502。
- 移除 pandas，降低常駐記憶體。
- XLSX/XLSM 改用 openpyxl read-only 串流讀取。
- XLS 直接使用 xlrd。
- ODS 直接使用 odfpy。
- 保留 `/parse-file`、`/parse-pdf` 與驗算/防重複邏輯。

## API
- `GET /health`
- `POST /parse-file`
- `POST /parse-pdf`（舊前端相容）

Render Start Command：
`uvicorn app:app --host 0.0.0.0 --port $PORT`
