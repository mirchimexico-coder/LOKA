# LOKA — HOW TO DO THE BOOKS YOURSELF

You do not need Claude for a normal day. **Double-click `LOKA.bat`** and follow the menu.

---

## THE DAILY 2-MINUTE ROUTINE

1. Double-click **`LOKA.bat`**
2. Press **1** — Notepad opens. Type your numbers. Save & close.
3. Press **2** — it shows you exactly what it will record. **Read it.**
4. Press **3** — confirms, writes everything, runs a health check.
5. Open **VS Code** and **push**. Done.

That's it. Steps 2 and 3 show you the full plan before anything is written, so you
can always back out.

### What you type (this is the whole thing)

```
DATE: 28-Jul
MP: 2455   BBVA: 1200   SOFT: 0   CASH: 3315
TRANSFER-TO-ME: 0    TRANSFER-TO-BANK: 0
BBVA COMMISSION: 1.9%
SOFT COMMISSION: 0
NO-BILL: Pollo 905, Vegetables 150, Oxxo Azucar 24
PAID BY ME: Bistec 290
SALARY: Samu 1200
NOTE: busy Friday
```

Leave any line at `0` or blank if it doesn't apply. **Sunday/closed:** leave the card and
cash lines at 0 and just fill in `NO-BILL` — it records the shopping and skips the sales row.

### It works these out for you
- **Categories** — "Pollo" becomes Ingredients - Meat, "Oxxo Azucar" becomes Pantry, etc.
- **Vendors** — recognises Oxxo, Sam's, Chedraui, Guimar, Brisa del Mar, Tres B, Abastos...
- **MP commission** — always automatic (4.06%). Never type it.
- **BBVA commission** — from the 1.9%, or type the exact pesos if you have the statement.
- **TRANSFER-TO-ME** — adds it to your Owner Ledger *and* keeps it out of till cash.
- **PAID BY ME** — makes it reimbursable to you *and* keeps it out of till cash.
- **Salaries** — filed as Staff - Salary.

**Always read the preview.** If a category looks wrong, just tell Claude later —
it is a 10-second fix and nothing breaks.

---

## THE MENU

| Key | Does |
|---|---|
| 1 | Open today's sheet to type numbers |
| 2 | **Preview** — shows what will be recorded, writes nothing |
| 3 | **Record it** — writes the books, then health-checks |
| 4 | Open the dashboard |
| 5 | Open the workbook in Excel |
| 6 | **Health check** — run this any time you're unsure |
| 7 | Refresh dashboard only |
| 8 | **Teach a category** — fix anything it didn't recognise |
| 9 | Show everything you've taught it |
| 10 | **Report** — this week / this month / all time |
| 11 | **Record propinas** (card tips handed to staff) |
| 12 | **Cash count** — re-anchor to what you physically counted |
| 13 | **Money paid back to Capital** |
| 14 | **Settle the owner ledger** with you |
| 15 | **UNDO** — restore an earlier backup |
| 16 | **Scan receipt photos** — reads them on this PC |
| 17 | Open the receipts drop folder |

---

## READING RECEIPTS (menu 16)

1. Press **17** — the drop folder opens
2. Copy your receipt photos in (iPhone HEIC works, no conversion needed)
3. Press **16**

It reads each photo and tells you the vendor, date and total, with a **confidence**:

| | |
|---|---|
| **HIGH** | it found a line actually labelled "TOTAL" — trust it, but still glance |
| **MEDIUM** | it used the card/cash line instead — check it |
| **LOW** | it guessed the biggest number — **always check this one** |

It also prints a line you can paste straight into `today.txt`.

**This all runs on your PC.** Nothing is uploaded, no internet, no account, nothing to
install — it uses the OCR that is already part of Windows.

**Two honest limits:**
- **Never enter a number you haven't seen with your own eyes.** OCR misreads faded thermal
  paper and creased receipts. The confidence label tells you when to look harder.
- **It reads the total, not the line items.** For a receipt you want broken down by item,
  that is still a Claude job.

Better photos = better reads: flat, straight on, good light, whole receipt in frame.
**If a photo has two receipts side by side, it reads them as one** — photograph them
separately.

Options 10–15 used to need a Claude session. Each one shows you what it will do and asks
you to confirm, takes a backup first, and refreshes everything afterwards.

**Option 15 (UNDO) is your safety net.** It lists the last 15 backups with dates, and backs
up the current state before restoring — so even an undo is undoable.

---

## WHEN IT DOESN'T RECOGNISE SOMETHING (menu 8)

The preview flags anything it couldn't work out:

```
    340.00  Mantel para mesa    Supplies/Other    No bill  <-- ?? not recognised
    ^^ 1 item(s) could not be categorised - they will go to 'Supplies/Other'.
```

**Press 8.** It walks you through each unrecognised item:

1. It shows the item and asks which word to remember (press Enter to use the item name)
2. It lists the categories — type the number
3. Done. It's remembered **forever**

Then press **2** (preview) again and it will be filed correctly.

**It only asks once per word.** Teach it "servilletero" today and every future servilletero
is categorised automatically. Press **9** any time to see everything you've taught it.

You can also press 8 then **A** to add a word by hand, without waiting for it to come up
in a day's entry — useful if you know a new supplier or item is coming.

Everything you teach lives in `categories.json`. Your rules always win over the built-in
guesses, so this is also how you **correct** something it gets wrong: teach it the right
answer and it will never make that mistake again.

---

## HEALTH CHECK (menu 6, or `py doctor.py`)

Green `[OK]` on everything = the books are sound. It checks the things that have
silently broken before: formula ranges running out, blank cells, duplicate days,
bad expense rows, category spelling splits, and that Capital still adds to $600,000.

**If you ever see `[FAIL] blank F/H/I/J/R formula caches`** — that is the common one.
Fix: menu **7** (refresh dashboard). It is harmless and takes 5 seconds.

---

## THINGS ONLY YOU CAN DECIDE (tell Claude, don't guess)

These are rare and need a judgement call, so leave them for a Claude session:

- **Physical cash count** — re-anchoring the cash figure
- **Money moved between Capital and the restaurant** (repayments)
- **Anything on the Capital & Ownership sheet** (partner money, acquisition, renovation)
- **Settling your Owner Ledger** (when the restaurant pays you back)
- **A receipt photo you want read and itemised**
- **A number that looks wrong and you can't explain**

---

## THE 4 RULES THAT MATTER

1. **Customer paid YOUR account** → `TRANSFER-TO-ME`. It's revenue, and the restaurant
   owes you.
2. **Customer paid the RESTAURANT's bank** → `TRANSFER-TO-BANK`. It's revenue and it
   stays in restaurant money. No commission.
3. **You paid from your own pocket** → `PAID BY ME`. Still a real cost, but reimbursable
   to you.
4. **You put it on your card and already paid yourself back with restaurant money**
   → just put it in `NO-BILL`. It's a plain restaurant expense — no ledger.

**Tips (propinas) are not expenses.** They pass through. Only money the restaurant adds
*on top of* collected card tips is a real cost.

---

## IF SOMETHING GOES WRONG

| Problem | Fix |
|---|---|
| Excel cells look blank / margin says 100% | Menu **7** (refresh) |
| "Permission denied" when saving | Close the workbook in Excel, try again |
| Wrong number recorded | Tell Claude — every write is backed up (30 versions in `Backup\`) |
| Recorded the same day twice | Health check catches it; tell Claude to remove one |
| Want to undo everything | `Backup\` has timestamped copies — Claude restores in seconds |

**Every single write makes a backup first.** Nothing is ever lost.

---

## WHAT NOT TO TOUCH

- Don't hand-edit formula cells in Excel (the grey/calculated columns)
- Don't rename sheets (the emoji names must stay exactly as they are)
- Don't delete anything in `Backup\`, `batches\`, or `days\`
- Don't push from the browser — always VS Code

---

# APPENDIX: TYPING STRAIGHT INTO EXCEL

You can absolutely do it this way. `LOKA.bat` is faster and safer, but nothing stops you
editing the workbook by hand. **Two rules make it safe:**

> **RULE 1 — never type over a formula column.**
> **RULE 2 — after saving Excel, always run `LOKA.bat` → 7 (refresh), then 6 (health check).**
> Without the refresh, the dashboard is stale and your new cells may show blank.

### A. Add an expense — sheet `💸 Expenses`

Go to the **first empty row at the bottom** (currently **row 546**) and fill:

| Col | Header | What to type | Example |
|---|---|---|---|
| **B** | Date | a REAL date, not text | `28-Jul-2026` |
| **C** | Description | what it was | `Pollo` |
| **D** | Vendor | shop, or `No bill` | `OXXO` |
| **E** | Category | **must match the list below exactly** | `Ingredients - Meat` |
| **F** | Amount | number only, no `$` | `905` |
| **G** | Paid By | `Restaurant`, `Lohith`, or `Capital` | `Restaurant` |
| **H** | Method | `Cash` or `Card` | `Cash` |
| **J** | Notes | optional | `no receipt` |

**Categories — copy one of these exactly** (a new spelling silently splits your P&L):
```
Ingredients - Vegetables   Ingredients - Meat       Ingredients - Pantry
Ingredients - Bread        Ingredients - Dairy      Ingredients - Beverages
Ingredients - Fruit        Ingredients - Meat & Fish  Ingredients - Eggs
Ingredients - Other        Ingredients - Desserts
Staff - Salary             Staff - Advance          Staff - Propinas
Supermarket/General        Supplies/Other           Kitchen Supplies
Packaging/Disposables      Utilities/Internet       Utilities/Gas
Software/Subscription      Rent                     Office Supplies
Maintenance
```

Tip: click an existing row and drag its formatting down, so the date and money
formatting come with it.

### B. Add a day's sales — sheet `📅 Daily Log`

Next empty row is **74** (row 73 = 25-Jul).

1. **Select the whole previous row (73) and drag/copy it down to row 74.** This carries the
   formulas *and* the formatting. **Do this first** — it's what protects the formula columns.
2. Now overwrite ONLY these cells in row 74:

| Col | What | Leave 0 if none |
|---|---|---|
| **B** | Date | — |
| **C** | Day (Mon/Tue…) | — |
| **D** | Mercado Pago card | 0 |
| **E** | Cash | 0 |
| **G** | Transfers (to-me + to-bank added together) | 0 |
| **U** | Soft Restaurant card | 0 |
| **V** | Soft commission (exact pesos) | 0 |
| **Y** | BBVA card | 0 |
| **Z** | BBVA commission (card × 1.9%) | 0 |
| **L** | Notes | — |

**Never type in F, H, I, J or R** — those are revenue, expenses, net, margin and MP
commission, and they calculate themselves.

**Sunday / closed day:** do NOT add a Daily Log row at all. Just enter the shopping in
the Expenses sheet.

### C. Save, then refresh (this step is not optional)

1. Save and **close** Excel (it must be closed or the tools can't write).
2. `LOKA.bat` → **7** (refresh dashboard)
3. `LOKA.bat` → **6** (health check) — expect **ALL CLEAR**
4. Push in VS Code.

If you skip step 2, the dashboard still shows yesterday and Excel may show blank cells
in the new row.

### D. The three cases you must NOT do by hand in Excel alone

These need a second step, because cash-on-hand can't work them out on its own:

| If this happened | Also run this (one line, in `LOKA.bat` press 7's window or a terminal in C:\LOKA) |
|---|---|
| **Customer paid YOUR account** (say $260) | `py loka.py add-ledger --date 28-Jul-2026 --desc "Transfer-to-me 28-Jul" --transferred 260`<br>`py loka.py cash-adjust --delta -260 --reason "transfer-to-me 28-Jul"` |
| **You paid from your own pocket** (say $234) | put `Lohith` in column G of the expense, then:<br>`py loka.py add-ledger --date 28-Jul-2026 --desc "Pollo paid by Lohith" --spent 234`<br>`py loka.py cash-adjust --delta 234 --reason "owner-paid 28-Jul"` |
| **Money repaid to Capital** (say $2,000) | tell Claude — this also has to move on the Capital sheet |

Honestly: for these three, `LOKA.bat` does it all for you automatically. That's the main
reason to prefer it.

### E. If you make a mistake

Every write makes a backup. `Backup\` holds the last 30 workbooks with timestamps.
Tell Claude what happened and it restores or corrects in seconds. Nothing is ever lost.
