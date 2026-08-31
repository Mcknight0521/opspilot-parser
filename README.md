# OpsPilot Parser v4.2 — Trust Engine + Auto Event Store

## What changed
- `/history-events` no longer crawls DGPA/CWA during a user request.
- GitHub Actions automatically refreshes `data/official_events.json` every day at 00:17 Taiwan time.
- The first push of the updater automatically bootstraps historical data (default: 10 years).
- Daily runs are incremental: DGPA re-checks the latest 45 days; CWA re-checks the current and previous year.
- Render only reads/filter a small JSON file, so event lookup is fast and low-memory.
- The event store keeps source labels and source URLs for traceability.

## Deploy
Upload the contents of this folder to the root of the same GitHub repository used by Render.
Keep `.github/workflows/update-official-events.yml`, `scripts/`, and `data/`.

1. Push/commit the files.
2. GitHub Actions will automatically run **Update official events** because the updater/workflow changed.
3. That Action commits the populated `data/official_events.json` back to the repository.
4. If Render Auto-Deploy is enabled for this branch, Render redeploys automatically after that commit.
5. `/health` should show version `4.2.0`.

No database is required. No pandas is added.

## API
- `POST /parse-file`
- `POST /verify-files`
- `GET /history-events?start=YYYY-MM-DD&end=YYYY-MM-DD&region=屏東縣`
- `GET /health`

## Automatic event sources
- 停班停課：行政院人事行政總處（DGPA）
- 颱風警報：中央氣象署颱風資料庫（CWA）

If an official source is temporarily unavailable during the scheduled update, the currently committed event store remains available to users; user-facing API requests do not wait for the official site.
