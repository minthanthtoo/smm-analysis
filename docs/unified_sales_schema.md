# Unified Sales Schema (9-MLM + 7-MTL)

This schema is a superset that can ingest both:

- `9-MLM Table and DailySales-2026_Feb.xlsx`
- `7-MTL for Table and DailySales(2026 Jan to Mar).xlsx`

## 1) `Table` sheet unified table

```sql
CREATE TABLE IF NOT EXISTS stg_sales_table_unified (
  id BIGSERIAL PRIMARY KEY,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL DEFAULT 'Table',
  source_row INTEGER NOT NULL,
  loaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),

  stock_id TEXT,
  stock_particular TEXT,
  ml NUMERIC(10,3),
  qty NUMERIC(10,3),
  product_group TEXT,          -- Column1 / BE-Whisky grouping
  sales_price NUMERIC(14,2),

  customer_code TEXT,          -- "No"
  customer_name TEXT,          -- "ကုန်သည်"
  customer_address TEXT,       -- "လိပ်စာ"
  township TEXT,
  sales_channel TEXT,          -- Sales / WholeSales etc.
  sale_particular TEXT,        -- Particular (txn type)
  car_no TEXT,
  pg_name TEXT,                -- 7-MTL only

  secondary_stock_id TEXT,     -- 7-MTL col 20+
  secondary_particular TEXT,
  secondary_ml NUMERIC(10,3),
  secondary_qty NUMERIC(10,3),
  secondary_product_group TEXT,
  secondary_sales_price NUMERIC(14,2),

  team_or_van TEXT,            -- sparse unnamed col in 9-MLM Table
  extra_unmapped JSONB NOT NULL DEFAULT '{}'::jsonb
);
```

### `Table` column mapping

| Unified column | 9-MLM `Table` | 7-MTL `Table` |
|---|---|---|
| `stock_id` | col 1 `Stock ID` | col 1 `Stock ID` |
| `stock_particular` | col 2 `Particular` | col 2 `Particular` |
| `ml` | col 3 `ml` | col 3 `ml` |
| `qty` | col 4 `Qty` | col 4 `Qty` |
| `product_group` | col 5 `Column1` | col 5 `Column1` |
| `sales_price` | col 6 `Sales Price` | col 6 `Sales Price` |
| `customer_code` | col 9 `No` | col 8 `No` |
| `customer_name` | col 10 `ကုန်သည်` | col 9 `ကုန်သည်` |
| `customer_address` | col 11 `လိပ်စာ` | col 10 `လိပ်စာ` |
| `township` | col 12 `Township` | col 11 `Township` |
| `sales_channel` | col 13 `Sales` | col 12 `Sales` |
| `sale_particular` | col 15 `Particular` | col 14 `Particular` |
| `car_no` | col 17 `Car No` | col 16 `Car No.` |
| `pg_name` | `NULL` | col 18 `PG Name` |
| `secondary_stock_id` | `NULL` | col 20 `Stock ID` |
| `secondary_particular` | `NULL` | col 21 `Particular` |
| `secondary_ml` | `NULL` | col 22 `ml` |
| `secondary_qty` | `NULL` | col 23 `Qty` |
| `secondary_product_group` | `NULL` | col 24 `Column1` |
| `secondary_sales_price` | `NULL` | col 25 `Sales Price` |
| `team_or_van` | col 19 (unnamed, sparse) | `NULL` |

## 2) `DailySales` sheet unified table

```sql
CREATE TABLE IF NOT EXISTS stg_daily_sales_unified (
  id BIGSERIAL PRIMARY KEY,
  source_file TEXT NOT NULL,
  source_sheet TEXT NOT NULL DEFAULT 'DailySales',
  source_row INTEGER NOT NULL,
  loaded_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),

  year SMALLINT,
  month SMALLINT,
  sale_date TIMESTAMP,
  snapshot_date TIMESTAMP,     -- "Today"
  period_text TEXT,
  product_group TEXT,          -- Column1 / BE/Whisky
  voucher_no TEXT,
  car_no TEXT,
  customer_id TEXT,
  customer_name TEXT,          -- "ကုန်သည်အမည်"
  township TEXT,
  customer_address TEXT,       -- only 7-MTL has direct column
  sales_channel TEXT,          -- WholeSales / Van-1 / etc.
  sale_particular TEXT,
  stock_id TEXT,
  stock_name TEXT,
  ml NUMERIC(10,3),
  pack_size NUMERIC(10,3),     -- "ပါဝင်မှု"
  bottle NUMERIC(14,3),
  parking NUMERIC(14,3),
  sales_pk NUMERIC(14,3),
  sales_bot NUMERIC(14,3),
  sales_liter NUMERIC(14,3),
  unit_price NUMERIC(14,2),    -- "နှုံး"
  line_amount NUMERIC(14,2),   -- "သင့်ငွေ"
  opening_balance NUMERIC(14,2),
  discount_old_price NUMERIC(14,2),
  commission NUMERIC(14,2),
  discount_extra NUMERIC(14,2),   -- 7-MTL only
  transport_discount NUMERIC(14,2), -- 7-MTL only
  transport_plus NUMERIC(14,2),     -- 7-MTL only
  payment_date_1 TIMESTAMP,
  payment_amount_1 NUMERIC(14,2),
  payment_date_2 TIMESTAMP,
  payment_amount_2 NUMERIC(14,2),
  payment_date_3 TIMESTAMP,     -- 7-MTL only
  payment_amount_3 NUMERIC(14,2),
  payment_date_4 TIMESTAMP,     -- 7-MTL only
  payment_amount_4 NUMERIC(14,2),
  outstanding_balance NUMERIC(14,2),

  -- 9-MLM extra duplicate/report block (col 34-44)
  report_date TIMESTAMP,
  report_month TEXT,
  report_year SMALLINT,
  report_customer_name TEXT,
  report_township TEXT,
  report_team TEXT,
  report_particular TEXT,
  report_stock_name TEXT,
  report_sales_ctns NUMERIC(14,3),
  report_sales_bot NUMERIC(14,3),
  report_sales_liter NUMERIC(14,3),

  extra_unmapped JSONB NOT NULL DEFAULT '{}'::jsonb
);
```

### `DailySales` column mapping

| Unified column | 9-MLM `DailySales` | 7-MTL `DailySales` |
|---|---|---|
| `year` | col 1 `Year` | col 1 `Year` |
| `month` | col 2 `Month` | col 2 `Month` |
| `sale_date` | col 3 `Date` | col 3 `Date` |
| `snapshot_date` | col 4 `Today` | col 4 `Today` |
| `period_text` | col 5 `Period` | col 5 `Period` |
| `product_group` | col 6 `Column1` | col 6 `BE/Whisky` |
| `voucher_no` | col 7 `VoucherNo` | col 7 `VoucherNo` |
| `car_no` | col 8 `CarNo` | col 8 `CarNo` |
| `customer_id` | col 9 `CustomerID` | col 9 `CustomerID` |
| `customer_name` | col 10 `ကုန်သည်အမည်` | col 10 `ကုန်သည်အမည်` |
| `township` | col 11 `Township` | col 11 `Township` |
| `customer_address` | `NULL` | col 12 `လိပ်စာ` |
| `sales_channel` | col 12 `WholeSales` | col 13 `Sales` |
| `sale_particular` | col 13 `Particular` | col 14 `Particular` |
| `stock_id` | col 14 `StockID` | col 15 `StockID` |
| `stock_name` | col 15 `StockName` | col 16 `StockName` |
| `ml` | col 16 `ML` | col 17 `ML` |
| `pack_size` | col 17 `ပါဝင်မှု` | col 18 `ပါဝင်မှု` |
| `bottle` | col 18 `Bottle` | col 19 `bottle` |
| `parking` | col 19 `Parking` | col 20 `Parking` |
| `sales_pk` | col 20 `SalesPK` | col 21 `SalesPK` |
| `sales_bot` | col 21 `SalesBot` | col 22 `SalesBot` |
| `sales_liter` | col 22 `Liter` | col 23 `Liter` |
| `unit_price` | col 23 `နှုံး` | col 24 `နှုံး` |
| `line_amount` | col 24 `သင့်ငွေ` | col 25 `သင့်ငွေ` |
| `opening_balance` | col 25 `စာရင်းဖွင့်` | col 26 `စာရင်းဖွင့်` |
| `discount_old_price` | col 26 `ဈေးဟောင်းလျှော့` | col 27 `ဈေးဟောင်းလျှော့` |
| `commission` | col 27 `ကော်မရှင်` | col 28 `ကော်မရှင်` |
| `discount_extra` | `NULL` | col 29 `ဈေးလျှော့` |
| `transport_discount` | `NULL` | col 30 `ကားခလျှော့` |
| `transport_plus` | `NULL` | col 31 `ကားခ+` |
| `payment_date_1` | col 28 `ငွေရသည့်နေ့` | col 32 `ငွေရသည့်နေ့` |
| `payment_amount_1` | col 29 `ကြွေးရငွေ` | col 33 `ကြွေးရငွေ` |
| `payment_date_2` | col 30 `ငွေရသည့်နေ့(၂)` | col 34 `ငွေရသည့်နေ့(၂)` |
| `payment_amount_2` | col 31 `ကြွေးရငွေ2` | col 35 `ကြွေးရငွေ2` |
| `payment_date_3` | `NULL` | col 36 `ငွေရသည့်နေ့(၃)` |
| `payment_amount_3` | `NULL` | col 37 `ကြွေးရငွေ3` |
| `payment_date_4` | `NULL` | col 38 `ငွေရသည့်နေ့ ၄` |
| `payment_amount_4` | `NULL` | col 39 `ကြွေးရငွေ4` |
| `outstanding_balance` | col 32 `အကြွေးကျန်2` | col 40 `အကြွေးကျန်2` |
| `report_date` | col 34 `Date` | `NULL` |
| `report_month` | col 35 `Month` | `NULL` |
| `report_year` | col 36 `Year` | `NULL` |
| `report_customer_name` | col 37 `Customer Name` | `NULL` |
| `report_township` | col 38 `Township` | `NULL` |
| `report_team` | col 39 `Team` | `NULL` |
| `report_particular` | col 40 `Particular` | `NULL` |
| `report_stock_name` | col 41 `Stock Name` | `NULL` |
| `report_sales_ctns` | col 42 `Sales Ctns` | `NULL` |
| `report_sales_bot` | col 43 `Sales Bot;` | `NULL` |
| `report_sales_liter` | col 44 `Sales Liter` | `NULL` |

## 3) Import rules (required before load)

1. Treat empty string / whitespace / `0`-like text carefully; cast to numeric/date only when valid.
2. `Car No` and `Car No.` are the same field (`car_no`).
3. `WholeSales` (9-MLM) and `Sales` (7-MTL) both map to `sales_channel`.
4. Keep source lineage columns (`source_file`, `source_row`) for traceability.
5. Any unexpected non-header columns (for example sparse unnamed columns) should go to `extra_unmapped`.

