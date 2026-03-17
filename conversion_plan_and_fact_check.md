# Conversion Plan + Fact Check (All Excel Files)

Generated from workbook scan in this workspace.

## 1) File Inventory Fact Check

| File | Sheets | Merged Cells | Formula Cells | Styled Cells |
|---|---:|---:|---:|---:|
| 7-MLM for Township Summary.xlsx | 15 | 0 | 0 | 193 |
| 7-MLM for Township Summary_4.xlsx | 12 | 0 | 0 | 690 |
| 7-MTL for Table and DailySales(2026 Jan to Mar).xlsx | 3 | 3 | 100015 | 257467 |
| 7-MTL for Township Summary_4.xlsx | 12 | 647 | 56680 | 230731 |
| 9-MLM Table and DailySales-2026_Feb.xlsx | 12 | 432 | 36916 | 88053 |
| MHL 2026 Feb.xlsx | 25 | 702 | 13334 | 62751 |
| MLM 2026 Feb.xlsx | 25 | 0 | 0 | 3404 |
| MLM_2026_Feb_SKU_Analysis.xlsx | 5 | 0 | 0 | 12334 |
| MLM_2026_Feb_Township_Analysis.xlsx | 6 | 0 | 0 | 2229 |
| MTL_source_to_7MTL_test.xlsx | 12 | 0 | 0 | 138 |
| MTL_source_to_MHL_test.xlsx | 25 | 0 | 0 | 559 |

## 2) Source Workbook Fact Check

Source: `7-MTL for Table and DailySales(2026 Jan to Mar).xlsx`

| Sheet | Rows | Columns | Merged | Formulas |
|---|---:|---:|---:|---:|
| Table | 7427 | 25 | 0 | 0 |
| DailySales | 6174 | 44 | 0 | 100013 |
| Sheet1 | 4 | 7 | 3 | 2 |

### DailySales Column Mapping: Current Parser vs Actual Header

| Field | Current Parser Column | Actual Header Column | Status |
|---|---:|---:|---|
| Date | 3 | 3 | OK |
| VoucherNo | 7 | 7 | OK |
| CarNo | 8 | 8 | OK |
| CustomerID | 9 | 9 | OK |
| CustomerName | 10 | 10 | OK |
| Township | 11 | 11 | OK |
| Team | 12 | 13 | MISMATCH |
| Particular | 13 | 14 | MISMATCH |
| StockID | 14 | 15 | MISMATCH |
| StockName | 15 | 16 | MISMATCH |
| ML | 16 | 17 | MISMATCH |
| Packing | 17 | 18 | MISMATCH |
| Bottle | 18 | 19 | MISMATCH |
| SalesPK | 20 | 21 | MISMATCH |
| SalesBot | 21 | 22 | MISMATCH |
| Liter | 22 | 23 | MISMATCH |
| Price | 23 | 24 | MISMATCH |
| Amount | 24 | 25 | MISMATCH |

### Dynamic Extraction Facts (Header-based)

- Kept rows: **4083**
- Skipped rows: {'no_sales_qty': 972, 'missing_township': 503, 'missing_date': 615}
- Date range: **2026-01-02 to 2026-03-14**
- Unique townships: **11**
- Unique SKUs (StockID+ML+Pack): **54**
- Month distribution:
  - 2026-01: 1780
  - 2026-02: 1648
  - 2026-03: 655
- Team distribution:
  - WholeSales/Contract: 2344
  - Van-1: 672
  - Van-2: 661
  - WholeSales: 313
  - Semi WS: 93

## 3) Goal Template Fact Check

### MHL 2026 Feb.xlsx

- Sheets: **25**
- Merged cells: **702**
- Formula cells: **13334**

| Sheet | Rows | Columns | Merged | Formulas |
|---|---:|---:|---:|---:|
| 7-MTL | 56 | 127 | 66 | 3206 |
| Business Summary | 41 | 22 | 5 | 20 |
| Ws Semi Ws  | 52 | 80 | 37 | 545 |
| Final MTL SKU Wise | 70 | 128 | 68 | 943 |
| MTL Individual | 52 | 122 | 50 | 1266 |
| MTL SKU Analysis | 77 | 156 | 81 | 2308 |
| Township wise Analysis | 20 | 102 | 54 | 408 |
| Final Town SKU wise Analysis | 61 | 52 | 39 | 1522 |
| Outlet Summary | 29 | 30 | 32 | 40 |
| Outlet List  | 319 | 98 | 9 | 0 |
| Way Plan | 85 | 10 | 0 | 86 |
| 3-Van Wise SKU | 39 | 65 | 57 | 502 |
| Competition Information | 27 | 21 | 1 | 49 |
| Top 3 or 4 brands in Township | 55 | 29 | 16 | 375 |
| Meiktila | 31 | 46 | 17 | 242 |
| Tharzi | 21 | 46 | 17 | 170 |
| Pyawbwe | 23 | 46 | 17 | 185 |
| Wantwin | 30 | 46 | 17 | 249 |
| Mahlaing | 21 | 46 | 17 | 199 |
| Yamethin | 23 | 46 | 17 | 197 |
| Kyaukpandaung | 24 | 46 | 17 | 198 |
| Taungthar | 31 | 45 | 17 | 155 |
| Myingan | 28 | 45 | 17 | 125 |
| Bagan | 37 | 45 | 17 | 201 |
| Pakokku | 30 | 45 | 17 | 143 |

### 7-MTL for Township Summary_4.xlsx

- Sheets: **12**
- Merged cells: **647**
- Formula cells: **56680**

| Sheet | Rows | Columns | Merged | Formulas |
|---|---:|---:|---:|---:|
| 7-MTL | 35 | 91 | 141 | 1037 |
| Meiktila | 164 | 127 | 46 | 5647 |
| Tharzi | 164 | 127 | 46 | 5425 |
| Pyaw Bwe | 164 | 127 | 46 | 4852 |
| Want Twin | 164 | 127 | 46 | 5209 |
| Mahaling | 164 | 127 | 46 | 5174 |
| Yamethin | 164 | 127 | 46 | 4922 |
| Kyaukpadaung | 164 | 127 | 46 | 5013 |
| Taungthar | 164 | 127 | 46 | 4747 |
| Myingyan | 164 | 127 | 46 | 4945 |
| Bagan | 164 | 127 | 46 | 4943 |
| Pakokku | 164 | 127 | 46 | 4766 |

## 4) Compatibility Findings

- MHL-like output sheet-name compatibility: 25/25
- 7-MTL-like output sheet-name compatibility: 12/12
- Template fidelity gap: generated outputs are flat tables (no merges/formulas), while templates rely heavily on merged/formula-driven layouts.
- Township naming gap exists mainly in spellings (e.g., Wandwin/Wantwin, Kyukpadaung/Kyaukpadaung, Pakkoku/Pakokku).
- SKU keys contain case variants and aliases; normalization is required before aggregation.

## 5) Conversion Improvement Plan

1. **Parser hardening (mandatory first step)**
   - Replace fixed index extraction with header-based column resolver for `DailySales`.
   - Support bilingual header aliases (e.g., Team=`Sales`, CustomerName=`ကုန်သည်အမည်`, Price=`နှုံး`, Amount=`သင့်ငွေ`).
   - Add fail-fast validation: if required headers missing, stop with explicit error.
2. **Master-table parsing fix**
   - Read first stock block (`A:F`) and customer block (`H:L`) explicitly from `Table` sheet.
   - Avoid accidental right-side duplicate block (`T:Y`) unless intentionally needed.
3. **Normalization layer**
   - Township normalization dictionary for spelling variants.
   - StockID normalization to collapse case/alias variants before aggregation.
   - Optional SKU synonym map (`source stock_id -> template product row label`).
4. **Template-preserving writer**
   - Open goal templates directly, write values into designated data ranges, keep existing merges/formulas/styles intact.
   - Never rebuild goal sheets from scratch for production deliverables.
5. **Data-quality gates before export**
   - Gate checks: non-zero kept rows, month coverage includes target month, township coverage >= expected 11, SKU coverage threshold.
   - Emit reconciliation sheet: source totals vs target totals by month/town/SKU.
6. **Validation + regression tests**
   - Add tests for this source file layout (`2026 Jan to Mar`) and the previous MLM source layout.
   - Snapshot checks for sheet names, key row/column totals, and top 10 SKUs per township.

## 6) Acceptance Criteria

- Parser keeps >0 rows on both source formats.
- Generated output retains template merges/formulas/style counts within template baselines.
- Township and SKU totals reconcile to source within tolerance (exact for bot/liter, near-exact for amount if price fallback used).
- Final files open with unchanged sheet names/order matching goal templates exactly.