# SMM-SKU Viewer

Local web viewer for fixed-format township Excel sheets.
You can choose:

- 1 workbook as overview analysis (main view)
- 1 workbook as detail reference (secondary view)

## Features

- No fixed workbook hard-coding: defaults are chosen dynamically using filename heuristics + most recent files.
- Upload new Excel files directly in the UI (`Upload + refresh files`) without restarting the app.
- Region-aware uploads: tag uploaded files with a region (or auto-detect from filename).
- RBAC with region/township scoping:
  - `Owner`: full cross-region access, assign RSMs, assign users to RSMs, assign ASMs, set ASM township scopes, and manage uploaded files.
  - `RSM`: view/manage only assigned regions, upload files in assigned regions, assign ASM + township scopes for owned regions.
  - `ASM`: view only assigned townships inside assigned regions.
  - `User`: view only regions inherited from mapped RSM.
- Management UI for:
  - Owner -> assign RSM + region list.
  - Owner -> map user to RSM.
  - Owner/RSM -> assign ASM and township visibility per region.
  - Owner/RSM -> file region update/delete (for uploaded files they have rights to).
- `Current user` selector in the UI lets you switch perspective and verify each role's scoped view.
- Current implementation is role-simulation (no login/auth provider yet): access is enforced from selected `user` context.
- Workbook pickers let you switch overview/detail files at runtime.
- Onboarding, access-management forms, and regional file management are organized directly inside the ribbon tabs.
- Main View recreates all sheets from the selected overview workbook as tabs.
- Reference View also has its own tabs from the selected detail workbook (independent from main tabs).
- Main/Reference sheet tabs are shown at the bottom (Excel-like) and Owner/RSM can rename sheet tabs inline (writes directly to the workbook file).
- Ribbon can be collapsed/expanded (Excel-like) via toggle button, tab double-click, or `Ctrl+F1`.
- Main window chrome now follows Excel desktop style (title bar, menu strip, ribbon band, and formula bar).
- Fixed Excel-style bottom status bar includes active-cell/range info, multi-cell selection stats (Count/Numbers/Sum/Avg/Min/Max), and zoom controls.
- Grid selection supports drag-range selection plus additive multi-area selection using `Ctrl`/`Cmd`.
- Excel-style grid context menu is available via right-click (desktop) and long-press (mobile) with copy/selection/zoom actions.
- Main + Detail panels split left/right on large landscape screens; on smaller/vertical screens the Detail panel becomes a bottom drawer (expand/collapse).
- Main overview table preserves the source sheet formatting (merged headers, cell styles, row/column layout) while filtering month columns.
- Workbook-native hidden rows/columns are respected in the styled main view.
- Workbook `freeze_panes` is respected (frozen rows + columns), with automatic identity-column freeze fallback when no explicit freeze is defined.
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
- Uploaded workbook region metadata is stored in `uploads/workbook_registry.json`.
- Access mappings are stored in `uploads/access_control.json`.

## Frontend Code Structure

The frontend is modularized into separate runtime files:

- `static/modules/app_state.js`: shared app state object, DOM element registry, and month label constants.
- `static/modules/ribbon.js`: ribbon-focused UI sync/utilities (tab switching, mirror controls, summary sync).
- `static/app.js`: main application orchestration (API calls, rendering, RBAC behaviors, event wiring).

`templates/index.html` loads these in order: state module -> ribbon module -> main runtime.

## Excel Parity Research

Feature-by-feature Excel applicability and WYSIWYG gap analysis is documented in [`docs/excel-wysiwyg-gap-analysis.md`](/Users/min/codex/SMM-SKU/docs/excel-wysiwyg-gap-analysis.md).

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

## Template JSON Generator

Generate both template workbooks from normalized JSON while preserving formulas:

```bash
cd /Users/min/codex/SMM-SKU
python3 generate_excel_from_template_json.py \
  --input-json template_generation_input.example.json \
  --mhl-template "MHL 2026 Feb.xlsx" \
  --town-template "7-MTL for Township Summary_4.xlsx" \
  --output-dir /Users/min/codex/SMM-SKU/outputs
```

Default output names:

- `MHL_from_json.xlsx`
- `7-MTL_for_Township_Summary_from_json.xlsx`

You can extend the JSON with:

- normalized facts (`sales_monthly_sku_township`, `sales_monthly_customer`, etc.)
- explicit cell patches (`workbook_patches`)
- explicit row/table patches (`workbook_table_patches`)

## Deploy On Render

This repo now includes [`render.yaml`](/Users/min/codex/SMM-SKU/render.yaml), so you can deploy quickly.

1. Push this repo to GitHub.
2. In Render, click **New +** -> **Blueprint**.
3. Connect the GitHub repo and select this project.
4. Render will auto-read `render.yaml` and create the web service.
5. Click **Apply** / **Create** and wait for the first deploy.

Render config used:

- Build command: `pip install -r requirements.txt`
- Start command: `gunicorn --bind 0.0.0.0:$PORT --workers 1 --timeout 180 app:app`

### Important note about uploads

Files uploaded via the UI are saved under `uploads/`. On Render's default filesystem, those files are **ephemeral** and can be lost on restart/redeploy.

If you need uploaded files to persist, use one of these:

- Attach a Render Disk and save uploads there.
- Store files in external object storage (for example S3-compatible storage).
