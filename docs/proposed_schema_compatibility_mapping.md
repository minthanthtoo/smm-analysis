# Proposed Schema Compatibility Mapping (Source Excel -> Target SQL)

## Scope

This mapping validates compatibility between the proposed normalized schema and these source files:

- `7-MTL for Table and DailySales(2026 Jan to Mar).xlsx`
- `7-MTL for Township Summary_4.xlsx`
- `MHL 2026 Feb.xlsx`
- `9-MLM Table and DailySales-2026_Feb.xlsx` (cross-check for variant layout)

## Compatibility Summary

| Target table | Status | Primary source | Notes |
|---|---|---|---|
| `product` | Compatible | `7-MTL/Table` A:F (+ T:Y) | Needs stock-code alias normalization for DailySales variants |
| `customer` | Compatible | `7-MTL/Table` H:L | 17 customer IDs in DailySales are missing from master list |
| `township` | Partially compatible | `DailySales.K`, `Table.K`, report tabs | Spelling and invalid-value cleanup required (`0`, `#N/A`) |
| `village` | Not source-populatable | N/A | No explicit village column in scanned source structures |
| `vehicle` | Compatible | `Table.P`, `DailySales.H` | DailySales car is sparse but valid |
| `sales_person` | Partially compatible | `Table.R (PG Name)` | No reliable transaction-level linkage in DailySales |
| `route_way` | Partially compatible | `MHL/Outlet List.F`, `MHL/Way Plan.C` | Formatting mismatch: `Van3-1` vs `Van 3-1` |
| `outlet` | Partially compatible | `MHL/Outlet List` | Current-state outlet list exists; no stable outlet key in DailySales |
| `village_township_history` | Not source-populatable | N/A | No village data + no effective dates |
| `customer_territory_history` | Partially compatible | `Table` customer township/channel | Only current assignment available, no historical periods |
| `outlet_territory_history` | Partially compatible | `MHL/Outlet List` township/way | Current assignment only, no historical periods |
| `outlet_salesperson_history` | Not source-populatable | N/A | No dated outlet->salesperson source |
| `route_outlet_assignment` | Partially compatible | `MHL/Outlet List` + `Way Plan` | Current assignment only, no effective dates |
| `sales_invoice` | Partially compatible | `DailySales` | `township_id_at_sale` available; outlet/village/salesperson-at-sale mostly unavailable |
| `sales_invoice_line` | Compatible | `DailySales` | Keep source row/line_no; voucher alone is not a safe unique key |
| `payment_schedule` | Compatible | `DailySales` AF:AM | Expand due-date/due-amount pairs into installments |
| `route_plan` | Compatible | `MHL/Way Plan` | Direct map of date/day/way and outlet counts |
| `sales_target` | Partially compatible | `MHL/Business Summary` + target tabs | Mostly aggregated at report level, not full dimensional target granularity |
| `competitor_product_price` | Compatible | `MHL/Competition Information` | Good direct mapping with minor text normalization |

## Load Order

1. `product`
2. `customer`
3. `township`
4. `vehicle`
5. `sales_person`
6. `route_way`
7. `outlet`
8. `sales_invoice`
9. `sales_invoice_line`
10. `payment_schedule`
11. `route_plan`
12. `sales_target`
13. `competitor_product_price`
14. History tables (seed current rows first; true history after dated sources are available)

## Mapping Rules By Table

## 1) product

Primary source: `7-MTL/Table`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/Table` | `A Stock ID` | `product.stock_code` |
| `7-MTL/Table` | `B Particular` | `product.product_name_mm` or `product.product_name_en` |
| `7-MTL/Table` | `C ml` | `product.size_liter` |
| `7-MTL/Table` | `D Qty` | `product.pack_qty` |
| `7-MTL/Table` | `E Column1` | `product.category` |
| `7-MTL/Table` | `F Sales Price` | `product.default_sales_price` |
| `7-MTL/Table` | derived (`1 * Qty`) | `product.pack_desc` |
| `7-MTL/Table` | `T:Y` (secondary stock block) | optional secondary source / alias ingest |

Transform rules:

- Normalize `stock_code` (trim, collapse repeated spaces, case policy).
- Maintain `product_alias(source_stock_id, product_id)` for known variants from DailySales.
- Keep `brand_name`, `abv_percent` nullable unless another trusted source is introduced.

## 2) customer

Primary source: `7-MTL/Table`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/Table` | `H No` | `customer.customer_code` |
| `7-MTL/Table` | `I ကုန်သည်` | `customer.customer_name_mm` |
| `7-MTL/Table` | `J လိပ်စာ` | `customer.address_text` |

Related assignment fields (current state only):

| Source workbook/sheet | Source column | Target location |
|---|---|---|
| `7-MTL/Table` | `K Township` | seed `customer_territory_history.township_id` (single current row) |
| `7-MTL/Table` | `L Sales` | seed `customer_territory_history.sales_channel` (single current row) |

Transform rules:

- Use alias table for customer code variants (`SM` vs `Sm`, etc.).
- Keep `phone_number` nullable; it is not present in this source.

## 3) township

Primary sources:

- `7-MTL/DailySales` column `K Township`
- `7-MTL/Table` customer township column `K`
- report workbook township labels

| Source text | Target column |
|---|---|
| Township names | `township.township_name_en` (canonicalized) |

Transform rules:

- Reject placeholders: `0`, `#N/A`, blank.
- Normalize spellings with alias map (`Want Twin`/`Wandwin`, `Yamaethin`/`Yamethin`, etc.).

## 4) vehicle

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/Table` | `P Car No.` | `vehicle.car_no` |
| `7-MTL/DailySales` | `H CarNo` | lookup to `vehicle` during transaction load |

## 5) sales_person

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/Table` | `R PG Name` | `sales_person.sales_person_name` |

Note:

- DailySales does not provide reliable per-row salesperson link in this source format.

## 6) route_way

Primary sources:

- `MHL/Outlet List` column `F Way`
- `MHL/Way Plan` column `C Way`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `MHL/Outlet List` | `F Way` | `route_way.way_code` |
| `MHL/Outlet List` | optional route description fields | `route_way.actual_way_name` |
| `MHL/Way Plan` | `C Way` | route reference / validation |
| `MHL/Way Plan` | `D Actual Way Name` | `route_way.actual_way_name` |

Transform rules:

- Canonicalize way code formatting (`Van3-1` -> `Van 3-1`).
- `route_type` can be derived from prefix (`Van`, `WS`, etc.) where available.

## 7) outlet

Primary source: `MHL/Outlet List`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `MHL/Outlet List` | `B ဆိုင်အမည်` | `outlet.outlet_name` |
| `MHL/Outlet List` | `C TYPE` | `outlet.outlet_type` |
| `MHL/Outlet List` | `D လိပ်စာ` | `outlet.address_text` |
| `MHL/Outlet List` | `G ဖုန်းနံပါတ်` | `outlet.phone_number` |
| `MHL/Outlet List` | `H ဘယ်သူ့ Outletလဲ` | `outlet.owner_name` |
| `MHL/Outlet List` | `I ရင်းနှင်းမှု့` | `outlet.investment_level` |

Related assignment fields (current state only):

| Source workbook/sheet | Source column | Target location |
|---|---|---|
| `MHL/Outlet List` | `E Township` | seed `outlet_territory_history.township_id` |
| `MHL/Outlet List` | `F Way` | seed `outlet_territory_history.route_way_id` |

Important limitation:

- DailySales in this source does not carry a stable outlet key, so `sales_invoice.outlet_id` should be nullable in phase 1.

## 8) sales_invoice

Primary source: `7-MTL/DailySales`

Recommended header grouping key:

- `(sale_date, voucher_no, customer_id, sales_channel)` plus deterministic source ordering

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/DailySales` | `C Date` | `sales_invoice.sale_date` |
| `7-MTL/DailySales` | `G VoucherNo` | `sales_invoice.voucher_no` |
| `7-MTL/DailySales` | `I CustomerID` | `sales_invoice.customer_id` (via customer lookup) |
| `7-MTL/DailySales` | `H CarNo` | `sales_invoice.vehicle_id` |
| `7-MTL/DailySales` | `M Sales` | `sales_invoice.channel_type` |
| `7-MTL/DailySales` | `K Township` | `sales_invoice.township_id_at_sale` |

Nullable in phase 1 due source gaps:

- `outlet_id`
- `village_id_at_sale`
- `route_way_id_at_sale`
- `sales_person_id_at_sale`

## 9) sales_invoice_line

Primary source: `7-MTL/DailySales`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `7-MTL/DailySales` | `N Particular` | `sales_invoice_line.transaction_type` |
| `7-MTL/DailySales` | `O StockID` | `sales_invoice_line.product_id` (via product/alias lookup) |
| `7-MTL/DailySales` | `R ပါဝင်မှု` | `sales_invoice_line.pack_qty` |
| `7-MTL/DailySales` | `S bottle` | `sales_invoice_line.bottle_qty` |
| `7-MTL/DailySales` | `U SalesPK` | `sales_invoice_line.sales_pack_qty` |
| `7-MTL/DailySales` | `V SalesBot` | `sales_invoice_line.sales_bottle_qty` |
| `7-MTL/DailySales` | `W Liter` | `sales_invoice_line.liter_qty` |
| `7-MTL/DailySales` | `X နှုံး` | `sales_invoice_line.unit_price` |
| `7-MTL/DailySales` | `Y သင့်ငွေ` | `sales_invoice_line.gross_amount` |
| `7-MTL/DailySales` | `Z စာရင်းဖွင့်` | `sales_invoice_line.opening_balance` |
| `7-MTL/DailySales` | `AA ဈေးဟောင်းလျှော့` | `sales_invoice_line.old_price_discount` |
| `7-MTL/DailySales` | `AB ကော်မရှင်` | `sales_invoice_line.commission_amount` |
| `7-MTL/DailySales` | `AC ဈေးလျှော့` | `sales_invoice_line.discount_amount` |
| `7-MTL/DailySales` | `AD ကားခလျှော့` | `sales_invoice_line.freight_discount` |
| `7-MTL/DailySales` | `AE ကားခ+` | `sales_invoice_line.freight_charge` |
| `7-MTL/DailySales` | `AN အကြွေးကျန်2` | `sales_invoice_line.outstanding_balance` |

Load notes:

- Keep a technical `line_no` from source row ordering.
- Keep source lineage (`source_file`, `source_sheet`, `source_row`) to resolve duplicate voucher+stock cases safely.

## 10) payment_schedule

Primary source: `7-MTL/DailySales`

| Source workbook/sheet | Source columns | Target columns |
|---|---|---|
| `7-MTL/DailySales` | `AF + AG` | installment 1 (`due_date`, `amount_due`) |
| `7-MTL/DailySales` | `AH + AI` | installment 2 (`due_date`, `amount_due`) |
| `7-MTL/DailySales` | `AJ + AK` | installment 3 (`due_date`, `amount_due`) |
| `7-MTL/DailySales` | `AL + AM` | installment 4 (`due_date`, `amount_due`) |

Rule:

- Emit one row per non-null date/amount pair.

## 11) route_plan

Primary source: `MHL/Way Plan`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `MHL/Way Plan` | `A Date` | `route_plan.plan_date` |
| `MHL/Way Plan` | `B Day` | `route_plan.day_name` |
| `MHL/Way Plan` | `C Way` | `route_plan.route_way_id` |
| `MHL/Way Plan` | `D Actual Way Name` | `route_plan.actual_way_name` |
| `MHL/Way Plan` | `E/F/G/H/I` | `outlet_a_count`..`outlet_s_count` |
| `MHL/Way Plan` | `J ToTal` | `total_outlets` |

## 12) sales_target

Primary sources:

- `MHL/Business Summary`
- target fields in other summary tabs

Status:

- Partial compatibility only. Current target inputs are mostly report-level aggregates and may not cover all intended dimensions (`township_id`, `route_way_id`, `metric_type`) in a single clean source.

## 13) competitor_product_price

Primary source: `MHL/Competition Information`

| Source workbook/sheet | Source column | Target column |
|---|---|---|
| `MHL/Competition Information` | `B Region` | `region_name` |
| `MHL/Competition Information` | `C Town` | `town_name` |
| `MHL/Competition Information` | `D Company Name` | `company_name` |
| `MHL/Competition Information` | `E Distributor` | `distributor_name` |
| `MHL/Competition Information` | `F Township` | `township_id` (via lookup) |
| `MHL/Competition Information` | `G Product Name` | `product_name` |
| `MHL/Competition Information` | `K Product` | `focus_budget_flag` |
| `MHL/Competition Information` | `L Size/ML` | `size_text` |
| `MHL/Competition Information` | `M Packing Size` | `pack_size` |
| `MHL/Competition Information` | `N ABV %age` | `abv_percent` |
| `MHL/Competition Information` | `H DB Landing Price` | `db_landing_price` |
| `MHL/Competition Information` | `I DB Selling Price` | `db_selling_price` |
| `MHL/Competition Information` | `J Selling-landing` | `db_margin` |
| `MHL/Competition Information` | `O Buying Price` | `buying_price` |
| `MHL/Competition Information` | `P Freight+Labour` | `freight_labour` |
| `MHL/Competition Information` | `Q Trade promotion` | `trade_promotion` |
| `MHL/Competition Information` | `R Buying-(Freight+Promotion)` | `final_buying_price` |
| `MHL/Competition Information` | `U Remarks` | `notes` |

## Non-populatable or Deferred Fields

From current scanned sources, these fields should remain nullable or deferred:

- `village` master and village-based keys
- effective date ranges for territory history tables (`effective_from`, `effective_to`) beyond seeded current rows
- transaction-level `sales_person_id_at_sale`
- transaction-level `outlet_id` in MTL DailySales

## Recommended ETL Safety Rules

1. Add canonical alias tables: `customer_alias`, `product_alias`, `township_alias`, `route_way_alias`.
2. Keep `sales_invoice_line` lineage keys (`source_file`, `source_sheet`, `source_row`) for reconciliation.
3. Treat report workbooks (`7-MTL summary`, MHL summary tabs) as derived validation targets, not primary fact sources.
4. Seed history tables with single current rows first; ingest real historical periods only when dated reassignment source files are provided.
