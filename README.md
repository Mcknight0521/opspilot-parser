# OpsPilot Parser v4 — Trust Engine

統一支援 PDF / XLSX / XLS / XLSM / ODS / CSV / TSV / TXT。

## v4 重點
- 已知 ERP 報表：專用 parser + 原始總計 / 明細重算驗證
- 陌生表格：多工作表探索、表頭自動定位、中文/英文欄位語意映射
- 資料完整性：逐日報表自動檢查缺日；缺日不補 0、不自行推定原因
- 日層級欄位防重複：來客數 / 工時若同日重複則只計一次；同日值不一致時不擅自加總
- 小計 / 總計排除，避免重複計算
- `/verify-files`：多檔案交叉驗證共同 KPI
- `/history-events`：由後端查詢 DGPA 歷史停班停課，避免瀏覽器 CORS 問題
- 記憶體最佳化：XLSX/XLSM read-only、逐列解析，不使用 pandas

## Endpoints
- `GET /health`
- `POST /parse-file`
- `POST /verify-files`
- `GET /history-events?start=YYYY-MM-DD&end=YYYY-MM-DD&region=屏東縣`
- `POST /parse-pdf`（相容舊版）

## Render
Start command:
`uvicorn app:app --host 0.0.0.0 --port $PORT`
