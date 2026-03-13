# SMM-SKU Viewer

Local web viewer for fixed-format township Excel sheets.
You can choose:

- 1 workbook as overview analysis (main view)
- 1 workbook as detail reference (secondary view)

## Features

- No fixed workbook hard-coding: defaults are chosen dynamically using filename heuristics + most recent files.
- Upload new Excel files directly in the UI (`Upload + refresh files`) without restarting the app.
- Workbook pickers let you switch overview/detail files at runtime.
- Main View recreates all sheets from the selected overview workbook as tabs.
- Reference View also has its own tabs from the selected detail workbook (independent from main tabs).
- Main overview table preserves the source sheet formatting (merged headers, cell styles, row/column layout) while filtering month columns.
- Main overview keeps identity columns and shows only selected month groups to avoid long horizontal scrolling.
- Main month-group detection is adaptive (supports variable monthly layouts such as Bot/Lit, PK/Bot/Liter, and single-column month patterns).
- Filter long monthly columns by:
  - `Past N months`
  - `Same month across past N years`
- Main View and Ref Details have separate filter controls (independent mode/N/month selection).
- Metric filter: `Liter`, `PK`, `Bottle`, or all metrics.
- Product name search filter.
- Reference panel shows matching township sheet from the selected detail workbook with the same filter mode.
- Handles township name variants automatically (for example `Pyawbwe` vs `Pyaw Bwe`, `Wantwin` vs `Want Twin`).
- Supported file types: `.xlsx`, `.xlsm`, `.xltx`, `.xltm` (lock files like `~$...` are ignored).

## Run

```bash
cd /Users/min/codex/SMM-SKU
python3 app.py
```

The app now auto-picks an open port from: `5055`, `8000`, `8080`, `5000`.
It prints the exact URL at startup, for example:

```text
Starting viewer at http://127.0.0.1:5055
```

Open that printed URL in your browser.

If you want a fixed port explicitly, run:

```bash
python3 -c "from app import app; app.run(host='127.0.0.1', port=5055, debug=False)"
```

## Dependencies

- `Flask`
- `openpyxl`

Install if needed:

```bash
pip install Flask openpyxl
```

## Deploy On Render

This repo now includes [`render.yaml`](/Users/min/codex/SMM-SKU/render.yaml), so you can deploy quickly.

1. Push this repo to GitHub.
2. In Render, click **New +** -> **Blueprint**.
3. Connect the GitHub repo and select this project.
4. Render will auto-read `render.yaml` and create the web service.
5. Click **Apply** / **Create** and wait for the first deploy.

Render config used:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn --bind 0.0.0.0:$PORT app:app`

### Important note about uploads

Files uploaded via the UI are saved under `uploads/`. On Render's default filesystem, those files are **ephemeral** and can be lost on restart/redeploy.

If you need uploaded files to persist, use one of these:

- Attach a Render Disk and save uploads there.
- Store files in external object storage (for example S3-compatible storage).
