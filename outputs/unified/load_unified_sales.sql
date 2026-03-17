-- Optional: run DDL from docs/unified_sales_schema.md first.

-- Load unified Table rows
COPY stg_sales_table_unified (source_file, source_sheet, source_row, stock_id, stock_particular, ml, qty, product_group, sales_price, customer_code, customer_name, customer_address, township, sales_channel, sale_particular, car_no, pg_name, secondary_stock_id, secondary_particular, secondary_ml, secondary_qty, secondary_product_group, secondary_sales_price, team_or_van, extra_unmapped)
FROM '/Users/min/codex/SMM-SKU/outputs/unified/stg_sales_table_unified.csv'
WITH (FORMAT csv, HEADER true, ENCODING 'UTF8');

-- Load unified DailySales rows
COPY stg_daily_sales_unified (source_file, source_sheet, source_row, year, month, sale_date, snapshot_date, period_text, product_group, voucher_no, car_no, customer_id, customer_name, township, customer_address, sales_channel, sale_particular, stock_id, stock_name, ml, pack_size, bottle, parking, sales_pk, sales_bot, sales_liter, unit_price, line_amount, opening_balance, discount_old_price, commission, discount_extra, transport_discount, transport_plus, payment_date_1, payment_amount_1, payment_date_2, payment_amount_2, payment_date_3, payment_amount_3, payment_date_4, payment_amount_4, outstanding_balance, report_date, report_month, report_year, report_customer_name, report_township, report_team, report_particular, report_stock_name, report_sales_ctns, report_sales_bot, report_sales_liter, extra_unmapped)
FROM '/Users/min/codex/SMM-SKU/outputs/unified/stg_daily_sales_unified.csv'
WITH (FORMAT csv, HEADER true, ENCODING 'UTF8');
