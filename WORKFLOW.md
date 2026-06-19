# LOKA — Simple Daily Workflow

## The one-message routine (what YOU do)
At end of day, send Claude **one message**: all receipt photos + a short note:

```
DATE: 19-Jun
CARD: 2100   CASH: 1800   TRANSFER: 0
NO-BILL: lechuga 40, gas 350
TRANSFER-TO-ME: 0        (customer transfers paid into Lohith's account)
```

Photos can be HEIC/JPEG/PDF; typed notes work too. Claude reads the receipts, logs
everything, closes the day, and refreshes the dashboard — then you push via Visual Studio.

## Note-line meanings
- CARD / CASH / TRANSFER = the day's revenue split (the cierre).
- NO-BILL = expenses with no receipt (item + amount).
- TRANSFER-TO-ME = customer payments into Lohith's personal account → become revenue AND
  reduce the restaurant's balance owed to Lohith (Owner Ledger entry).
- Paid on Lohith's HSBC **5543 card = say so → goes to Owner Ledger as reimbursable.

## The toolkit — `C:\LOKA\loka.py` (auto-backs-up before every write)
```
python loka.py status                         # print all totals (all-time, per month, last 7 days)
python loka.py add-expense --date 19-Jun --desc "Tortillas" --vendor "Local" --cat "Ingredients - Bread" --amount 52 --paid Restaurant --method Cash
python loka.py add-expenses-json batch.json   # add many expenses at once
python loka.py close-day --date 19-Jun --card 2100 --cash 1800 --transfer 0
python loka.py refresh-dashboard              # regenerate ALL data-driven dashboard sections
```

## End-to-end daily sequence (what Claude runs)
1. Read receipts → build a batch.json of expenses.
2. `add-expenses-json batch.json`     (logs all expenses, auto-backup)
3. `close-day --date .. --card .. --cash .. --transfer ..`   (sets revenue; creates row + auto-SUMIFS)
4. (if any) owner-ledger entries for HSBC-card or transfer-to-me items
5. `refresh-dashboard`                (rebuilds KPIs, banner, cash, bars, weekly, monthly, alert, ledger, footer)
6. You push to GitHub via Visual Studio.

## refresh-dashboard — what it regenerates automatically
Header date · Cash on hand · Banner mini-stats · All-Time KPIs (rev/exp/opnet/commission/trading days) ·
Daily Revenue bars (rolling last 30 days) · Weekly Summary (last 8 weeks) · Monthly Breakdown
(cards + category table, last 6 months) · Daily alert · Owner Ledger card · Footer.
It is **idempotent** — safe to run repeatedly; output is stable.

## refresh-dashboard — manual anchors (edit CFG at top of the DASHBOARD REFRESH section in loka.py)
- `cash_anchor_date` / `cash_anchor_amount` — last physical cash count; cash = anchor + daily nets since.
- `cash_adjust` — net non-restaurant cash flows since the anchor (e.g. customer transfers to Lohith). Reset to 0 when you give a fresh count.
- `commission_rate` — Mercado Pago est. on card revenue (0.0406).
- `week1_start` — 18 May 2026 (weeks are 7-day blocks).

## NOT auto-refreshed (update manually when they change — they rarely do)
Acquisition card · Partner ownership cards · Staff Roster · Break-Even card · Pending Income alert.
These come from the Capital sheet / manual facts. Tell Claude when they change.

## Hard rules (baked into the tools)
Real Excel dates · auto-SUMIFS on new Daily Log rows · 30-backup rotation · never edit_block on the .xlsx ·
dashboard totals recomputed from the Expenses sheet · Claude never pushes to GitHub.
Live: https://mirchimexico-coder.github.io/LOKA/dashboard.html
