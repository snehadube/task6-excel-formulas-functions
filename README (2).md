# Task 6 — Excel Formulas & Functions Fundamentals
**Veda Technology Internship | Data Analytics Track | Level 1, Day 6**

## Objective
Build fluency with the everyday Excel formulas used in real analyst work — VLOOKUP/XLOOKUP, IF, SUMIFS, COUNTIFS, and text functions — by applying them to a live transactional dataset.

## Dataset
[Sample Superstore dataset](https://www.kaggle.com/datasets/vivek468/superstore-dataset-final) — 10,194 real order lines across 4 US regions, 3 product categories, and 3 customer segments, plus supporting dimension tables (Product, Customer, Location).

## Approach
1. Loaded the full `Sales_Facts` table (10,194 rows) along with its `Product_Dimension`, `Customer_Dimension`, and `Location_Dimension` reference tables — kept as-is so every formula below is checkable against real, unmodified source data.
2. **Formulas Demo** sheet — applied every required function to a live 60-row sample, referenced directly off `Sales_Facts` (not copied/pasted), so the results update if the source rows change:
   - `VLOOKUP` — pulls each customer's Segment from `Customer_Dimension` by Customer ID, cross-checked against Sales_Facts.
   - `XLOOKUP` — the same lookup written with modern XLOOKUP syntax, right next to the VLOOKUP column for comparison. **Note:** XLOOKUP only calculates in Excel 365+/current Google Sheets — older Excel and non-Excel spreadsheet engines will show `#NAME?` for this one column, which is expected and documented in the workbook itself.
   - `IF` — flags each order line as Profit or Loss.
   - Nested `IF` — buckets orders into Small / Medium / Large by Sales value.
   - Text functions (`LEFT`, `MID`, `FIND`, `TRIM`, `PROPER`, `UPPER`) — decodes the structured Product ID (e.g. `OFF-PA-10000174` → Category code `OFF`, Sub-category code `PA`) and cleans/initializes customer names.
3. **Summary** sheet — `SUMIFS`/`COUNTIFS` aggregations computed against the **full 10,194-row dataset**: Sales by Region × Category, Profit by Category × Segment, order counts by Region × Ship Mode, plus quick single-cell stats (loss-making orders, high-value orders, top region via `INDEX`/`MATCH`).
## A note on XLOOKUP compatibility
Both `VLOOKUP` and `XLOOKUP` are included side by side in the Formulas Demo sheet, plus `INDEX`/`MATCH` in the Summary sheet as a third, universally-compatible alternative. XLOOKUP is a newer function that not every spreadsheet engine can evaluate — if you open this file somewhere other than Excel 365 or current Google Sheets, the XLOOKUP column may show `#NAME?` while the VLOOKUP and INDEX/MATCH columns still return the correct value. This is expected and called out directly in the workbook.

## Outcome
- All formulas evaluate correctly in Excel 365 / Google Sheets. In stricter/older environments, everything is error-free except the dedicated XLOOKUP demo column, which is a known compatibility limitation of that one function — not a mistake.
- All aggregate figures (Sales, Profit, order counts) are computed live off the real dataset, not hardcoded.

## Files
- `Excel_Formulas_Functions_Task6_Superstore.xlsx` — the completed workbook (Formulas Demo, Summary, Notes, plus the source data tabs).


