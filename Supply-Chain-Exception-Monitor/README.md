# Supply Chain Exception Monitor

Exception-based planning in Python: instead of reviewing every SKU, a planner sees only what needs attention. The tool projects inventory 8 weeks ahead from forecasts, on-hand stock and open purchase orders, flags problems, ranks them by severity and financial impact, and writes a colour-coded Excel report plus a short weekly summary. An optional step lets an LLM (Claude) word the summary.

## Exceptions detected

| Exception | Trigger | Severity |
|---|---|---|
| `STOCKOUT_RISK` | projected on-hand drops below zero | Critical if A-class or within lead time, otherwise High |
| `BELOW_SAFETY_STOCK` | projected low point is under safety stock | Medium |
| `EXCESS_STOCK` | more than 16 weeks of cover | Medium if excess value is over $20k, otherwise Low |
| `LATE_PO` | purchase order is past its due date | High if A-class or over 7 days late, otherwise Medium |
| `DEMAND_SHIFT` | last 4 weeks of demand differ from the next-4-week forecast by more than 30% | Medium |

Each exception carries an impact in dollars (revenue at risk, excess inventory value or value of late stock) and a recommended action. Thresholds are constants at the top of `exception_monitor.py`.

## Quick start

```bash
pip install -r requirements.txt
python exception_monitor.py            # add --no-ai to skip the LLM summary
```

Sample data is included in `data/`, so it runs immediately. Outputs go to `outputs/`:

| File | Contents |
|---|---|
| `exception_report.xlsx` | Summary, exception counts, exceptions (coloured by severity), projected on-hand by week (negatives in red), snapshot |
| `weekly_summary.md` | Plain-English status update |
| `exceptions.csv` | All exceptions as a flat file |

## Sample output

```
WEEKLY SUPPLY EXCEPTION SUMMARY - week of 16 Nov 2026

Headline: 28 exceptions (2 critical, 9 high, 14 medium, 3 low). 4 SKUs face a projected
stockout with $24,020 of revenue at risk; $516,357 is tied up in excess stock.

Top priorities:
  - SKU031 [Critical] STOCKOUT_RISK: Projected to run out in week 8; shortfall 182 units
    over 8 wks (lead time 2 wks). Action: Expedite open POs; raise new PO for ~417 units
  - SKU010 [Critical] STOCKOUT_RISK: Projected to run out in week 3; shortfall 74 units
    over 8 wks (lead time 4 wks). Action: Expedite open POs; raise new PO for ~89 units
  - SKU021 [High] LATE_PO: PO50027: 874 units overdue by 13 days. Action: Chase supplier ...

Supplier follow-up: 10 purchase orders are overdue ($232,367 of stock in transit late).
```

> The inventory and purchase-order snapshot is **simulated** (`--refresh` regenerates it), so the figures above demonstrate the logic rather than a real business.

## Input files (`data/`)

Replace these CSVs with your own extracts to run the monitor on real data.

| File | Columns |
|---|---|
| `inventory_plan.csv` | `sku, category, ABC, unit_price, unit_cost, lead_time_weeks, avg_weekly_forecast, safety_stock, reorder_point` (extra columns are ignored) |
| `forecast_future.csv` | `sku, week, forecast` (weekly buckets starting on Monday, at least 8 weeks) |
| `sales.csv` | `week, sku, units` (history; the last 4 weeks drive the demand-shift check) |
| `inventory.csv` | `sku, on_hand` |
| `open_orders.csv` | `po_id, sku, qty, due_date` |

The first three files can come from a forecasting and inventory planning run such as [Demand-Forecasting-Inventory-Planning](https://github.com/chandanareddy19/Demand-Forecasting-Inventory-Planning), or from any planning system extract with the same columns.

## How the projection works

For each SKU and each of the next 8 weeks: `projected on-hand = previous on-hand + receipts due that week - forecast`. Purchase orders count as receipts in the week of their due date. **Overdue POs are deliberately excluded** (a conservative assumption) and raised as `LATE_PO` exceptions instead.

## AI-written summary (optional)

```bash
pip install anthropic
export ANTHROPIC_API_KEY=your_key_here        # PowerShell: $env:ANTHROPIC_API_KEY="your_key_here"
python exception_monitor.py
```

The exception numbers always come from the rules above. The LLM only turns the top exceptions into a short status update, and the prompt tells it to use only the supplied figures. If the package, key or network is missing, the script falls back to the rule-based summary automatically (`--no-ai` forces this). The summary file states which one produced the text. Set `ANTHROPIC_MODEL` to choose a different model.

## Limitations

- Deterministic projection: forecast uncertainty and supplier lead-time variability are not modelled.
- Severity rules and thresholds are simple and meant to be tuned.
- The demo snapshot is simulated; real data needs the CSV formats above.

## Possible next steps

- Weekly scheduling and email or Slack delivery of the summary
- A Power BI view over `exceptions.csv`
- Probabilistic stockout risk using forecast error
