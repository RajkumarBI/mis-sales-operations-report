# MIS Sales & Operations Reporting Workbook

A formula-driven Excel MIS model: raw order data goes in, a monthly management report, dashboard and data-quality checks come out. No macros, no hard-coded results.

**File:** `MIS_Sales_Operations_Report.xlsx` · **Currency:** INR · **Data:** synthetic sample (600 orders, Nov-2025 to Oct-2026)

## What's inside

| Sheet | Purpose |
|---|---|
| **Dashboard** | KPI tiles with a region selector, plus 4 charts (monthly trend, margin %, region vs target, category mix) |
| **MIS_Report** | Monthly trend, region vs target, category × region matrix, order status, top 5 customers |
| **Sales_Data** | 600 order lines. Columns A:G are inputs; H:AA are 20 calculated columns |
| **Master** | Customer, product, region/target, discount slab and parameters (GST, SLA days, thresholds) |
| **Formula_Library** | 55 formulas, each with its MIS use case and a live result |
| **Checks** | 13 reconciliation and data-quality tests with PASS / FAIL |
| **README** | In-workbook guide, colour code and refresh steps |

## Formulas covered

- **Lookup & reference:** VLOOKUP (exact and approximate), HLOOKUP, INDEX + MATCH (1-way and 2-way), CHOOSE, INDIRECT, OFFSET, ROWS / COLUMNS
- **Conditional aggregation:** SUMIF(S), COUNTIF(S), AVERAGEIF(S), MAXIFS, MINIFS, SUMPRODUCT, SUBTOTAL
- **Logical:** nested IF, IFS, SWITCH, AND / OR, IFERROR, ISNA
- **Text:** TRIM, PROPER, LEFT / MID / RIGHT, FIND, SUBSTITUTE, TEXT, TEXTJOIN, REPT, VALUE
- **Date & time:** EDATE, EOMONTH, DATEDIF, NETWORKDAYS, WORKDAY, WEEKNUM, fiscal-year logic
- **Math & statistics:** ROUND family, MROUND, MOD, LARGE / SMALL, RANK, MEDIAN, PERCENTILE, STDEV, CAGR, PMT
- **Modelling practice:** named ranges, data validation drop-downs, conditional formatting, reconciliation checks

`MAXIFS`, `MINIFS`, `IFS`, `SWITCH` and `TEXTJOIN` need Excel 2019 or Microsoft 365.

## Report logic

- Net sales = Gross sales − slab discount; invoice value = net sales + GST
- Cancelled orders are excluded from all sales figures
- SLA = delivery within `SLA_Days` working days (weekends excluded)
- Order size: Large / Medium / Small from the thresholds in `Master`
- Region status: On Target ≥ 100%, Near Target ≥ 90%, otherwise Below Target

## Colour code

- **Blue font:** input, safe to edit
- **Black font:** formula, do not overwrite
- **Green font:** formula linking to another sheet

## Refreshing with your own data

1. Paste new transactions into `Sales_Data` columns A:G.
2. Copy the formulas in H:AA down to the last row.
3. Extend the `d_*` named ranges in Name Manager.
4. Update `Master!T7` (report start month) and add any new customers or products.
5. Open `Checks`. Every line should read PASS before you circulate the report.

## Rebuilding the workbook from code

```bash
pip install -r requirements.txt
python build_workbook.py
```

Open the result in Excel and let it recalculate (Ctrl+Alt+F9) if values appear blank.

## Repository layout

```
.
├── MIS_Sales_Operations_Report.xlsx
├── build_workbook.py
├── requirements.txt
├── README.md
└── .gitignore
```

## License

MIT. Free to use and adapt.
