"""Supply Chain Exception Monitor (with an optional AI-written weekly summary).

Exception-based planning: instead of reviewing every SKU, a planner only looks at
what needs attention. This script
  1. loads the plan from ./data (forecasts, safety stock, ABC class, recent sales),
  2. loads a current-state snapshot (on-hand inventory + open purchase orders),
  3. projects inventory 8 weeks forward and raises exceptions:
       STOCKOUT_RISK | BELOW_SAFETY_STOCK | EXCESS_STOCK | LATE_PO | DEMAND_SHIFT
  4. writes a colour-coded Excel report and a one-page weekly summary.

The summary is written by an LLM (Claude) if the `anthropic` package is installed and
ANTHROPIC_API_KEY is set; otherwise a rule-based summary is used. Either way the
numbers come from the rules above - the LLM only does the wording.

Run:                python exception_monitor.py            # add --no-ai to skip the LLM
Use your own data:  replace the CSV files in ./data (formats are described in README.md)
"""
import argparse
import os
from pathlib import Path

import numpy as np
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = Path(__file__).parent
DATA, OUT = HERE / "data", HERE / "outputs"

HORIZON = 8            # weeks projected forward
EXCESS_WEEKS = 16      # more than this many weeks of cover = excess
SHIFT_THRESHOLD = 0.30  # recent demand vs forecast deviation that triggers an alert
SEVERITY_ORDER = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}


# --------------------------------------------------------------------------
# Inputs
# --------------------------------------------------------------------------
def load_plan():
    plan = pd.read_csv(DATA / "inventory_plan.csv")
    fc = pd.read_csv(DATA / "forecast_future.csv", parse_dates=["week"])
    sales = pd.read_csv(DATA / "sales.csv", parse_dates=["week"])
    return plan, fc, sales


def simulate_snapshot(plan, as_of, seed=7):
    """Create a realistic current-state snapshot so the demo runs end to end."""
    rng = np.random.default_rng(seed)
    inv, pos = [], []
    for r in plan.itertuples():
        mode = rng.choice(["ok", "low", "high"], p=[0.65, 0.20, 0.15])
        cover = {"ok": rng.uniform(8, 14), "low": rng.uniform(0.5, 5),
                 "high": rng.uniform(20, 35)}[mode]
        inv.append({"sku": r.sku, "on_hand": int(round(r.avg_weekly_forecast * cover))})
        for _ in range(rng.integers(0, 3)):
            pos.append({
                "po_id": f"PO{rng.integers(10000, 99999)}",
                "sku": r.sku,
                "qty": int(round(r.avg_weekly_forecast * rng.uniform(2, 6))),
                "due_date": (as_of + pd.Timedelta(days=int(rng.integers(-14, 42)))).date(),
            })
    return pd.DataFrame(inv), pd.DataFrame(pos)


def load_snapshot(plan, as_of, refresh=False):
    inv_f, po_f = DATA / "inventory.csv", DATA / "open_orders.csv"
    if refresh or not (inv_f.exists() and po_f.exists()):
        DATA.mkdir(exist_ok=True)
        inv, pos = simulate_snapshot(plan, as_of)
        inv.to_csv(inv_f, index=False)
        pos.to_csv(po_f, index=False)
        print(f"Simulated snapshot saved to {DATA} (replace with your own CSVs any time)")
    inv = pd.read_csv(inv_f)
    pos = pd.read_csv(po_f, parse_dates=["due_date"])
    return inv, pos


# --------------------------------------------------------------------------
# Exception engine
# --------------------------------------------------------------------------
def run_rules(plan, fc, sales, inv, pos, as_of):
    plan = plan.merge(inv, on="sku", how="left").fillna({"on_hand": 0}).set_index("sku", drop=False)
    exceptions, projections = [], {}
    late = pos[pos.due_date < as_of]
    ontime = pos[pos.due_date >= as_of]

    def add(sev, typ, sku, detail, impact, action):
        r = plan.loc[sku]
        exceptions.append({"severity": sev, "type": typ, "sku": sku, "category": r.category,
                           "ABC": r.ABC, "detail": detail, "impact_$": round(float(impact)),
                           "recommended_action": action})

    for sku, r in plan.iterrows():
        f = fc[fc.sku == sku].sort_values("week").head(HORIZON)
        weeks, fvals = list(f.week), f.forecast.to_numpy()
        stock, series = float(r.on_hand), []
        for w, d in zip(weeks, fvals):
            rec = ontime[(ontime.sku == sku) & (ontime.due_date >= w)
                         & (ontime.due_date < w + pd.Timedelta(days=7))].qty.sum()
            stock = stock + rec - d
            series.append(round(stock))
        projections[sku] = series
        low = min(series)

        # 1/2. stockout and safety-stock breaches
        if low < 0:
            wk = next(i for i, v in enumerate(series) if v < 0) + 1
            short = -low
            order_qty = int(round(short + r.safety_stock))
            sev = "Critical" if wk <= max(r.lead_time_weeks, 2) or r.ABC == "A" else "High"
            add(sev, "STOCKOUT_RISK", sku,
                f"Projected to run out in week {wk}; shortfall {short:,} units over {HORIZON} wks "
                f"(lead time {r.lead_time_weeks} wks)",
                short * r.unit_price,
                f"Expedite open POs; raise new PO for ~{order_qty:,} units "
                f"(shortfall + safety stock)")
        elif low < r.safety_stock:
            add("Medium", "BELOW_SAFETY_STOCK", sku,
                f"Projected low of {low:,} units is below safety stock of {int(r.safety_stock):,}",
                (r.safety_stock - low) * r.unit_cost,
                f"Bring forward next replenishment by 1-2 weeks; reorder point is "
                f"{int(r.reorder_point):,}")

        # 3. excess
        cover = r.on_hand / r.avg_weekly_forecast if r.avg_weekly_forecast > 0 else 99
        if cover > EXCESS_WEEKS:
            excess = r.on_hand - r.avg_weekly_forecast * EXCESS_WEEKS
            val = excess * r.unit_cost
            add("Medium" if val > 20000 else "Low", "EXCESS_STOCK", sku,
                f"{cover:.0f} weeks of cover vs {EXCESS_WEEKS}-week limit "
                f"(~{int(excess):,} units above limit)",
                val, "Pause replenishment; review promotion, transfer or markdown options")

        # 5. demand signal vs forecast
        recent = sales[sales.sku == sku].sort_values("week").units.tail(4).mean()
        planned = fvals[:4].mean()
        dev = (recent - planned) / planned if planned > 0 else 0
        if abs(dev) > SHIFT_THRESHOLD:
            way = "above" if dev > 0 else "below"
            add("Medium", "DEMAND_SHIFT", sku,
                f"Last 4-wk demand ({recent:,.0f}/wk) is {abs(dev):.0%} {way} next-4-wk forecast "
                f"({planned:,.0f}/wk)",
                abs(recent - planned) * 4 * r.unit_price,
                "Review forecast with sales; check for promo, new-listing or stock-out effects")

    # 4. late purchase orders
    for p in late.itertuples():
        r = plan.loc[p.sku]
        days = (as_of - p.due_date).days
        add("High" if days > 7 or r.ABC == "A" else "Medium", "LATE_PO", p.sku,
            f"{p.po_id}: {p.qty:,} units overdue by {days} days",
            p.qty * r.unit_cost, f"Chase supplier for {p.po_id}; confirm new ETA and expedite if needed")

    exc = pd.DataFrame(exceptions)
    exc["_s"] = exc.severity.map(SEVERITY_ORDER)
    exc = exc.sort_values(["_s", "impact_$"], ascending=[True, False]).drop(columns="_s")
    exc.insert(0, "#", range(1, len(exc) + 1))
    proj = pd.DataFrame(projections, index=[f"W{i}" for i in range(1, HORIZON + 1)]).T
    proj.index.name = "sku"
    return exc.reset_index(drop=True), proj, plan


# --------------------------------------------------------------------------
# Summary (LLM if available, else rule-based)
# --------------------------------------------------------------------------
def rule_summary(exc, as_of):
    sev = exc.severity.value_counts()
    so = exc[exc.type == "STOCKOUT_RISK"]
    ex = exc[exc.type == "EXCESS_STOCK"]
    lines = [
        f"WEEKLY SUPPLY EXCEPTION SUMMARY - week of {as_of:%d %b %Y}", "",
        f"Headline: {len(exc)} exceptions "
        f"({sev.get('Critical', 0)} critical, {sev.get('High', 0)} high, "
        f"{sev.get('Medium', 0)} medium, {sev.get('Low', 0)} low). "
        f"{len(so)} SKUs face a projected stockout with ${so['impact_$'].sum():,.0f} of revenue at risk; "
        f"${ex['impact_$'].sum():,.0f} is tied up in excess stock.", "",
        "Top priorities:",
    ]
    for r in exc.head(5).itertuples():
        lines.append(f"  - {r.sku} [{r.severity}] {r.type}: {r.detail}. Action: {r.recommended_action}")
    late = exc[exc.type == "LATE_PO"]
    if len(late):
        lines += ["", f"Supplier follow-up: {len(late)} purchase orders are overdue "
                      f"(${late['impact_$'].sum():,.0f} of stock in transit late)."]
    return "\n".join(lines)


def llm_summary(exc, as_of):
    """Ask Claude to word the summary. Returns None if unavailable (caller falls back)."""
    try:
        import anthropic
    except ImportError:
        print("  (anthropic package not installed - using rule-based summary)")
        return None
    if not os.getenv("ANTHROPIC_API_KEY"):
        print("  (ANTHROPIC_API_KEY not set - using rule-based summary)")
        return None
    top = exc.head(15).drop(columns=["#"]).to_csv(index=False)
    prompt = (
        f"You are a supply chain planning assistant. Today's planning week starts {as_of:%d %b %Y}.\n"
        "Below are the top exceptions from our weekly inventory projection (CSV). Write a status "
        "update for the planning team lead in under 200 words: one headline sentence, then the 3-5 "
        "most urgent issues with a concrete next step each, then one line on supplier follow-ups. "
        "Use ONLY the numbers given; do not invent data.\n\n" + top)
    try:
        client = anthropic.Anthropic()
        msg = client.messages.create(
            model=os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
            max_tokens=700,
            messages=[{"role": "user", "content": prompt}])
        return msg.content[0].text.strip()
    except Exception as e:  # network/key/model problems should never break the report
        print(f"  (LLM call failed: {e} - using rule-based summary)")
        return None


# --------------------------------------------------------------------------
# Excel report
# --------------------------------------------------------------------------
FILLS = {"Critical": "F8C9C8", "High": "FAD9C4", "Medium": "FBEBB5", "Low": "E3E8E1"}


def write_excel(exc, proj, plan, summary_text, source):
    path = OUT / "exception_report.xlsx"
    snap = plan[["sku", "category", "ABC", "on_hand", "avg_weekly_forecast", "safety_stock",
                 "reorder_point", "lead_time_weeks"]].reset_index(drop=True)
    counts = (exc.groupby(["type", "severity"]).size().unstack(fill_value=0)
              .reindex(columns=list(SEVERITY_ORDER), fill_value=0))
    counts["Total"] = counts.sum(axis=1)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        pd.DataFrame({"Weekly summary (" + source + ")": summary_text.split("\n")}
                     ).to_excel(xw, "Summary", index=False)
        counts.reset_index().to_excel(xw, "Exception_Counts", index=False)
        exc.to_excel(xw, "Exceptions", index=False)
        proj.reset_index().to_excel(xw, "Projected_OnHand", index=False)
        snap.to_excel(xw, "Snapshot", index=False)

    wb = load_workbook(path)
    head = PatternFill("solid", fgColor="2A78D6")
    for ws in wb.worksheets:
        for c in ws[1]:
            c.font, c.fill = Font(bold=True, color="FFFFFF"), head
            c.alignment = Alignment(horizontal="center")
        for i, col in enumerate(ws.columns, 1):
            w = max(len(str(c.value)) if c.value is not None else 0 for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(max(w + 3, 8), 70 if ws.title == "Summary" else 60)
        ws.freeze_panes = "A2"
    ws = wb["Exceptions"]
    sev_col = [c.value for c in ws[1]].index("severity") + 1
    for row in ws.iter_rows(min_row=2):
        sev = row[sev_col - 1].value
        for c in row:
            c.fill = PatternFill("solid", fgColor=FILLS[sev])
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws = wb["Projected_OnHand"]
    red = Font(color="C00000", bold=True)
    for row in ws.iter_rows(min_row=2, min_col=2):
        for c in row:
            if isinstance(c.value, (int, float)) and c.value < 0:
                c.font = red
    wb.save(path)


# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--no-ai", action="store_true", help="skip the LLM summary")
    ap.add_argument("--refresh", action="store_true", help="regenerate the simulated snapshot")
    args = ap.parse_args()

    plan, fc, sales = load_plan()
    as_of = fc.week.min()
    inv, pos = load_snapshot(plan, as_of, args.refresh)
    exc, proj, plan_full = run_rules(plan, fc, sales, inv, pos, as_of)

    text = None if args.no_ai else llm_summary(exc, as_of)
    source = "AI-written" if text else "rule-based"
    text = text or rule_summary(exc, as_of)

    OUT.mkdir(exist_ok=True)
    (OUT / "weekly_summary.md").write_text(f"_Summary source: {source}_\n\n{text}\n")
    exc.to_csv(OUT / "exceptions.csv", index=False)
    write_excel(exc, proj, plan_full, text, source)

    print(f"\n{text}\n")
    print(exc.groupby(["type", "severity"]).size().to_string())
    print(f"\nOutputs written to {OUT}")


if __name__ == "__main__":
    main()
