# Excel Feature Gap Analysis (SMM-SKU)

Date: 2026-03-14

This document maps major Excel capabilities to what is currently possible in this codebase and what is required for near-Excel or exact-Excel behavior.

## Executive Summary

- The app can now render many workbook-defined visual properties in the main styled table: font, fill, borders, alignment, merged cells, row heights, column widths, hidden rows/columns, and freeze panes.
- Exact WYSIWYG parity with desktop Excel is not achievable with a plain HTML table renderer alone.
- For exact fidelity, use Microsoft 365 for the web integration (WOPI/CSPP) or a high-fidelity commercial spreadsheet engine.

## Current Capability Matrix

| Feature Area | Excel Capability | Current Status in SMM-SKU | Notes |
|---|---|---|---|
| Cell style | Fonts, colors, fills, borders, alignment | Supported (main styled view) | Rendered from workbook styles to CSS. |
| Dimensions | Row height, column width | Supported | Workbook dimensions are applied in generated table HTML. |
| Merge | Merged ranges | Supported (visible ranges) | Merges render when full merged range is visible/selected. |
| Hide/Unhide rows/columns | Manual hide/unhide | Supported (workbook-defined) | Hidden rows/columns are skipped in main styled output. |
| Freeze/Unfreeze panes | Freeze top rows / left columns | Supported (workbook-defined) | `freeze_panes` now maps to frozen row+column behavior. |
| Number formatting | Currency, percent, accounting, locale/date formats | Partial | Values are displayed but not full Excel number-format engine parity. |
| Filters/sort (AutoFilter/Table) | Interactive drop-down filters/sorts | Partial | App has custom filters; Excel AutoFilter semantics are not fully mirrored. |
| Excel tables | Structured references, table styles, totals row behavior | Not yet | No direct parsing/rendering of table objects yet. |
| Conditional formatting | Color scales, icon sets, data bars | Not yet | Requires conditional-format rule evaluation + rendering layer. |
| Data validation | Input lists/rules/messages | Not yet | No validation UI/runtime in rendered grid. |
| Formula engine | Full calculation graph/function compatibility | Not yet | Renderer shows cell values; does not execute workbook calc engine in browser. |
| Pivot/charts/slicers/images/shapes | Interactive objects | Not yet | Not represented in current HTML-table architecture. |
| Comments/notes/threaded comments | Annotations and collaboration | Not yet | Not surfaced in UI. |
| Protection/macros | Sheet/workbook protection, VBA | Not yet | Out of scope for current rendering model. |
| Print/page layout | Print areas, page breaks, headers/footers | Not yet | No print-layout emulation. |

## Hard Boundary for “Exact Excel WYSIWYG”

Exact parity requires Excel’s own rendering and behavior engine. In practice, choose one:

1. Microsoft 365 for the web integration (WOPI/CSPP).
2. Commercial high-fidelity spreadsheet component (accepting some differences).

Why this boundary exists:

- Microsoft documents feature differences between browser and desktop Excel.
- Embedded workbook experiences are constrained and have limits.
- Open-source parsing libraries generally do not provide full Excel UI/behavior parity.

## Recommended Implementation Roadmap

### Phase 1 (Already delivered / low-risk parity)

- Keep style-preserving HTML renderer.
- Respect workbook hidden rows/columns.
- Respect workbook `freeze_panes` rows/columns.

### Phase 2 (Near-term parity upgrades)

- Add Excel-like number format rendering for common patterns (date/time/currency/percent/accounting).
- Add workbook AutoFilter interpretation for initial filtered views.
- Add table-object awareness (header row, total row, table style hints).

### Phase 3 (Advanced Excel semantics)

- Conditional formatting renderer (top used rule types).
- Data validation rendering + optional enforcement in edit flows.
- Formula recalculation strategy (server-side calc service or spreadsheet engine).

### Phase 4 (Exact WYSIWYG target)

- Integrate Microsoft 365 for the web via WOPI/CSPP for native Excel behavior.
- Keep current Python APIs for RBAC, workbook selection, and non-Excel business workflows.

## Architecture Options

### Option A: Keep current HTML renderer (lowest cost)

- Fastest to maintain.
- Best for read-heavy, curated views.
- Will never be fully Excel-identical.

### Option B: Excel-native surface with WOPI/CSPP (highest fidelity)

- Closest to exact Excel behavior and rendering.
- Highest integration and operational complexity.

### Option C: JS spreadsheet engine (middle path)

- Rich interactions and editing in browser.
- Better than plain tables, but still not guaranteed identical to desktop Excel in every edge case.

## Sources

Microsoft support and platform docs:

- https://support.microsoft.com/en-us/office/freeze-panes-to-lock-rows-and-columns-dab2ffc9-020d-4026-8121-67dd25f2508f
- https://support.microsoft.com/en-us/office/hide-or-show-rows-or-columns-659c2cad-802e-44ee-a614-dde8443579f8
- https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
- https://support.microsoft.com/en-us/office/differences-between-using-a-workbook-in-the-browser-and-in-excel-f0dc28ed-b85d-4e1d-be6d-5878005db3b6
- https://support.microsoft.com/en-us/office/embed-your-excel-workbook-on-your-web-page-or-blog-from-sharepoint-or-onedrive-for-business-7af74ce6-e8a0-48ac-ba3b-a1dd627b7773
- https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/overview
- https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/online/hostpage
- https://learn.microsoft.com/en-us/microsoft-365/cloud-storage-partner-program/rest/endpoints
- https://learn.microsoft.com/en-us/graph/excel-concept-overview
- https://learn.microsoft.com/en-us/graph/excel-update-range-format

Library/platform references:

- https://docs.sheetjs.com/docs/getting-started/examples/import
- https://docs.sheetjs.com/docs/miscellany/roadmap
- https://handsontable.com/docs/javascript-data-grid/row-freezing/
- https://handsontable.com/docs/javascript-data-grid/column-freezing/
- https://handsontable.com/docs/javascript-data-grid/row-hiding/
- https://www.ag-grid.com/javascript-data-grid/excel-import/

openpyxl API references:

- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.worksheet.html
- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.dimensions.html
- https://openpyxl.readthedocs.io/en/stable/styles.html
- https://openpyxl.readthedocs.io/en/stable/api/openpyxl.worksheet.merge.html
