# LOKA — SYSTEM NOTES (read me first)

> **Purpose:** everything an assistant needs to resume LOKA bookkeeping without
> re-discovering it. Read this file FIRST at the start of a session.
> Last full review: **27-Jul-2026**.

---

## 1. BUSINESS FACTS

| | |
|---|---|
| Business | LOKA Cafe Restaurante, Querétaro, MX |
| Partners | **Lohith/Reddy 50%** (managing, does the books), Kashigoud 30%, Shashirekha 20% |
| Committed capital | **$600,000** ($300k / $180k / $120k) — treat as ONE POOLED POT |
| Opening / Week 1 | 18-May-2026 |
| Trading days | Mon–Sat. **Sunday CLOSED** (shopping-only days: expenses, NO Daily Log row) |
| Currency | MXN |

### Card processors & commission
| Processor | Card col | Comm col | Rate | How comm is set |
|---|---|---|---|---|
| Mercado Pago | D (4) | R (18) | 4.06% (`$T$7`) | **auto-formula** `=D*$T$7` |
| Soft Restaurant | U (21) | V (22) | ~2.05% ref | **actual value typed in** (varies daily) |
| BBVA | Y (25) | Z (26) | 1.9% | usually rate×card, but accept actual |

Revenue = MP + Soft + BBVA + cash + transfer(col G).

---

## 2. FILES

```
C:\LOKA\
  LOKA_Restaurant_Manager.xlsx   <- the books
  loka.py                        <- toolkit (ALL writes go through this where possible)
  dashboard.html                 <- published to GitHub Pages
  backup_util.py                 <- 30-version rotating backups
  NOTES.md                       <- this file
  WORKFLOW.md                    <- one-message EOD format
  README.txt                     <- original project readme
  bebidas_recipes.html           <- bar staff drinks reference (internal, not customer-facing)
  batches\                       <- archived batch_*.json expense records
  Backup\                        <- rotating backups: 30 .xlsx + 15 dashboard .html
  Backups\, My Backup\           <- Reddy's OWN manual backups — DO NOT TOUCH
  Bills\                         <- receipt filing scaffold (_DROP_BILLS_HERE + month folders)
  Documentation\                 <- Soft Restaurant setup doc
```

**Folder hygiene (cleaned 27-Jul-2026):** root holds only the 9 live files above.
Never leave `_*.py` temp scripts behind — delete them in the same command that runs them.
Batch JSONs go in `batches\`. Dashboard snapshots now self-rotate at 15 (they had piled
up to 108 because only the .xlsx backups were rotating).

**Claude edits LOCAL FILES ONLY. Reddy pushes to GitHub via VS Code. Never push from the browser.**
Live dashboard: `https://mirchimexico-coder.github.io/LOKA/dashboard.html`

Python: `C:\Progra~1\Python312\python.exe`

---

## 3. WORKBOOK MAP

Sheet indices (openpyxl `wb.worksheets[i]`) — emoji names must be EXACT:

| i | Sheet | Notes |
|---|---|---|
| 0 | `Capital & Ownership` | Sections A–L, see §5 |
| 1 | `📊 Dashboard` | |
| 2 | `📅 Daily Log` | **sheet3.xml** in the zip |
| 3 | `💸 Expenses` | **sheet4.xml**; data from row 8 |
| 4 | `📦 Inventory` | |
| 5 | `👥 Staff & Payroll` | propinas table rows 61–95 |
| 6 | `🚚 Suppliers` | |
| 7 | `📋 Supplier Directory` | |
| 8 | `📈 Monthly P&L` | |
| 9 | `📥 Pending Income` | |
| 10 | `💰 Owner Ledger` | last sheet |

### Daily Log columns
`B`date · `C`day · `D`MP · `E`cash · `F`revenue(formula) · `G`transfer · `H`expenses(SUMIFS)
`I`net · `J`margin · `L`(12)notes · `R`(18)MP comm · `T7`=0.0406
`U`(21)Soft · `V`(22)Soft comm · `Y`(25)BBVA · `Z`(26)BBVA comm

### Expenses columns
`B`date (REAL Excel date, never text) · `C`desc · `D`vendor · `E`category · `F`amount
`G`paid-by · `H`method · `J`(10)notes. Data starts row 8.

### Owner Ledger
Headers row 3, data row 4+. `A`date `B`desc `C`type `D`spent-by-Lohith `E`transferred-to-Lohith
`F`running balance `G`status `H`notes. **TOTALS row is found dynamically** (starts with "TOTALS").
**Positive balance = restaurant owes Lohith.**

---

## 4. ACCOUNTING RULES (confirmed with Reddy — do not re-litigate)

| Situation | Treatment |
|---|---|
| **Transfer-to-me** (customer pays Lohith's personal acct) | revenue in col G **+** Owner Ledger `transferred` entry **+** `cash-adjust −amount` (cash never hit the till) |
| **Transfer to restaurant's own BBVA** | revenue in col G only. Stays in restaurant money. NO ledger, NO cash-adjust, NO commission |
| **Owner-paid expense** (Lohith fronts cash) | Expenses row with `paid="Lohith"` (still a real P&L cost) **+** Owner Ledger `spent` entry **+** `cash-adjust +amount` |
| **Partner/staff card purchase, reimbursed with restaurant money** | `paid="Restaurant"`, normal P&L, **NO ledger entry** |
| **Capital expense** (one-off equipment, e.g. camera cloud) | `paid="Capital"` → carved out of the cash roll via `capf_bd` |
| **Renovation / post-opening capex** | Capital sheet Section I |
| **Propinas (tips)** | Staff & Payroll table rows 61–95. **Pass-through, NOT P&L.** Only restaurant top-up ABOVE collected tips is a real "Staff - Propinas" expense |
| **Chef paid for working late** | real WAGE (Staff - Salary), not a tip |
| **Capital repayment** (ops → capital) | reduce Capital `C147` **+** `cash-adjust −amount`. **NOT a P&L expense** |
| **Soft Restaurant app subscription** | recurring MONTHLY operating expense (Utilities/Internet) |
| **Sunday** | expenses only, **no Daily Log row** |

### Weekly payroll (paid in arrears)
- **Samu** (Ayudante) — Friday, $1,200
- **John** (Head Chef) — Saturday, $3,750
- **Duvi/Debi** (Limpieza) — Saturday, $1,500 full week / $1,000 for 4 days

Weeks are 7-day blocks from 18-May. (Week 10 = 20–26 Jul.)

---

## 5. CAPITAL SHEET (sheet 0) — section map

| Sec | Rows | Contents |
|---|---|---|
| A | 5–8 | Ownership + committed budget ($600k) |
| B | 11–37 | Acquisition $390k. **D31 = balance still unpaid ($2,235)**. Payment breakdown rows 35–37 keeps Kashi's $125k and $20k as SEPARATE transactions (deliberate) |
| C | 42–51 | Kashi pre-opening ($145k transfer + Soft licence + terminal + 2 camera-cloud) |
| D | 56–67 | Lohith pre-opening. **D67 total. D56 = $5,000 advance that ALSO appears in acquisition row 14 — exclude one when summing (`D67−D56`)** |
| E | 71–81 | Shashi |
| F | 84–97 | Grand summary |
| G | 98–108 | Remaining transfers due. **F100 = Kashi still to transfer (~$18,432)** |
| H | ~106–109 | Acquisition split by share |
| I | 114–129 | Renovation / post-opening capex (D129 total) |
| J | 132–140 | Lohith cash position. **D140 = "remaining with me"** |
| K | 142–149 | Operations ↔ Capital. **C147 = advance still owed by Operations** (r149 = cumulative repayments) |
| L | 150–166 | **POOLED CAPITAL POSITION** — reconciles exactly to $600,000 |

### Section L identity (must always tie to $600k)
```
DEPLOYED  = acquisition paid + Lohith preopen (excl $5k advance) + Kashi direct
          + Shashi direct + renovation + operating-costs-funded-from-capital (C147)
AVAILABLE = cash with Lohith (D140) + Kashi still to transfer (F100) + Shashi (F101)
DEPLOYED + AVAILABLE == 600,000
```

---

## 6. TOOLKIT COMMANDS

### The fast path (use these first — costs almost no tokens)
```bash
py eod.py today.txt              # PREVIEW a whole day from Reddy's plain-text format
py eod.py today.txt --apply      # do it all: expenses + daily row + ledger + cash-adjust + refresh
py doctor.py                     # full health check, ALL_CLEAR or FAIL lines
LOKA.bat                         # menu for Reddy (no commands to remember)
```
**`eod.py` is the default answer for a routine day.** Paste Reddy's message into
`today.txt`, run the preview, show him the plan, then `--apply`. It handles
auto-categorisation, both transfer types, owner-paid items, salaries, Sunday
shopping-only days, ledger entries, cash-adjust and the refresh+cache in one shot.
Only drop to the granular commands below when something is unusual.

### Granular commands
```bash
py loka.py status                      # JSON summary
py loka.py backup <label>              # always before a batch of writes

py loka.py add-expenses-json batch.json    # bulk expenses
py loka.py add-expense --date .. --desc .. --vendor .. --cat .. --amount ..
                       [--paid Restaurant|Lohith|Capital] [--method Cash|Card] [--notes ..]

# full multi-card day in ONE command:
py loka.py close-day --date 24-Jul-2026 --card 2455 --cash 3315 \
       [--transfer N] [--soft N --softcomm N] [--bbva N] [--bbvacomm N | --bbvarate 0.019]

py loka.py add-ledger --date .. --desc .. [--spent N] [--transferred N] [--notes ..]
py loka.py cash-adjust --delta -260 --reason "transfer-to-me 24-Jul"

py loka.py refresh-all                 # ***USE THIS*** = refresh-dashboard + inject-cache
py loka.py refresh-dashboard           # dashboard only (leaves blank formula caches!)
py loka.py inject-cache                # formula-cache fix on its own
py loka.py refresh-pl
```

### CANONICAL CATEGORIES (21 — normalised 27-Jul; no spaces around "/")
```
Ingredients - Vegetables    Ingredients - Meat        Ingredients - Pantry
Ingredients - Bread         Ingredients - Dairy       Ingredients - Beverages
Ingredients - Fruit         Ingredients - Meat & Fish Ingredients - Eggs
Ingredients - Other         Ingredients - Desserts
Staff - Salary              Staff - Advance           Staff - Propinas
Supermarket/General         Supplies/Other            Kitchen Supplies
Packaging/Disposables       Utilities/Internet        Utilities/Gas
Software/Subscription       Rent                      Office Supplies
Maintenance
```
**Never invent a new spelling variant** (`Supermarket / General` vs `Supermarket/General`
split the P&L into two lines — merged 55 cells on 27-Jul). `doctor.py` now detects this.

### batch JSON schema
```json
[{"date":"24-Jul-2026","desc":"...","vendor":"...","cat":"...",
  "amount":123.45,"paid":"Restaurant","method":"Cash","notes":"..."}]
```

---

## 7. STANDARD EOD — now ONE command

```bash
# 1. write Reddy's message verbatim into today.txt   2. preview   3. apply
py eod.py today.txt                 # show him this plan first
py eod.py today.txt --apply         # does everything + refresh + cache
py doctor.py                        # confirm ALL CLEAR
```
Then tell Reddy to push via VS Code.

**Token discipline:** for a routine day this is 2 tool calls, not 8. Do NOT hand-write
batch JSON, daily-row openpyxl scripts, ledger scripts or cache-injection scripts any
more — every one of those is now inside `eod.py` / `loka.py`. Reach for bespoke scripts
only for genuinely novel work (Capital sheet edits, receipt OCR, investigations).

### Reddy can now self-serve
`LOKA.bat` gives him a menu (enter numbers -> preview -> record -> health check), and
`HOWTO.md` is his plain-English guide. He only needs a Claude session for: cash-count
re-anchoring, Capital/partner money, ledger settlements, receipt photos, and anything
that looks wrong. Those are listed in HOWTO.md §"THINGS ONLY YOU CAN DECIDE".

---

## 8. HARD-WON TECHNICAL LESSONS (do not regress)

1. **Formula cache (`<v/>`) — THE recurring gotcha.** openpyxl writes formulas with an
   EMPTY cached value and this Excel does NOT honour `fullCalcOnLoad`, so cells render
   BLANK and margin shows 100%. `_inject_cache()` writes real values into `<v>` while
   keeping `<f>`. **`refresh_dashboard` clears caches, so injection must run AFTER it** —
   that's exactly what `refresh-all` does.
2. **SUMIFS row caps.** All ranges were capped at row 500 and silently returned ZERO once
   Expenses passed row 500 (~17-Jul). Swept the whole workbook → **all ranges now 5000**
   (Inventory 2000, Suppliers 3000). Re-extend when Expenses nears 5000.
2b. **The cap REGRESSED once (found 27-Jul).** Fixing the workbook is not enough: the P&L
   generator inside `loka.py` (`refresh_pl`/`_write_pl`) had `$500` hard-coded, and since
   every `refresh-dashboard` re-syncs Monthly P&L, it silently REWROTE the capped formulas
   each run — so Monthly P&L category totals were under-reporting live. Now `$5000` there
   too. **Rule: when fixing a formula in the workbook, also fix the code that GENERATES it,
   or the next refresh undoes you.** Verify with a workbook-wide capped-range scan.
3. **Never start a text label with `=`.** Excel treats it as a formula → **#NAME?**.
   (Bit us at Capital B136/B140/B160/B165.)
4. **Never use `edit_block` on Excel cells** — corrupts the file to 0 bytes. Use openpyxl.
5. **Windows console is cp1252** → never `print()` emoji/accents; the script crashes AFTER
   saving. Write output to a file and `type` it.
6. **Never run inline Python via `powershell -c`** — quoting/`$`/emoji break. Write a
   `.py` file, run it, delete it.
7. **PowerShell `Set-Content` re-encodes UTF-8 as Windows-1252** and destroys Spanish
   characters. Write HTML/text with Python `open(..., encoding='utf-8', newline='\n')`.
8. **Excel file lock**: if the workbook is open, saves fail — ask Reddy to close it.
9. **HEIC receipts**: `pip install pillow_heif --break-system-packages` (container resets
   each turn). The `view` tool often fails on converted images → fall back to
   `tesseract file -l spa+eng --psm 6`. **Bills are often TWO receipts per photo — always
   check for a second one.** If OCR is unsure on a money amount, ASK rather than guess
   (cross-check that line items sum to the printed total).
10. **Self-referencing SUM** (e.g. `=SUM(C154:C160)` written in C160) → circular ref.
11. **Cache-injection regex must handle SELF-CLOSING cells.** `<c r=".."[^>]*>.*?</c>`
    breaks on `<c r="Q74" s="1"/>`: `[^>]*>` eats the `/>`, then `.*?</c>` runs on and
    swallows the NEXT cell, which therefore never gets a cached value. This silently left
    the ENTIRE R (MP-commission) column blank. Pattern is now
    `<c r="(?P<ref>[A-Z]+\d+)"[^>]*?(?:/>|>.*?</c>)`. Found 27-Jul.
12. **`cash_adjust_add` must update `CFG` IN MEMORY, not just the file.** It rewrites
    `loka.py` on disk, but a `refresh_all()` later in the SAME process was still using the
    value loaded at import → Cash on Hand was overstated by the transfer amount ($700 on
    27-Jul). Fixed by also setting `CFG['cash_adjust'] = new`. **General rule: any helper
    that edits `loka.py` must also update the live CFG dict.**
13. **Dashboard cards whose SIGN can flip need direction-aware `sub()` rules.** The Owner
    Ledger went negative for the first time on 27-Jul; the old rule wrote a mangled
    `$-550` under the label "Restaurant owes Lohith" and then could not re-match itself on
    the next refresh (`WARN ledger bal: no match`). Both the label and an ABS value are now
    rewritten. **A `WARN ... no match` is never cosmetic — it means a card is frozen at a
    stale number. Always investigate.**
14. **`refresh-all` must be idempotent.** After any change to a `sub()` rule, run it TWICE
    and confirm zero WARNs both times.
15. **Hard-coded dashboard cards go stale silently — this is the biggest class of dashboard
    bug.** The 3 partner ownership cards were hand-maintained. Shashirekha's still read
    "Deployed $118,895 / Owes Lohith $1,105" long after the -$1,034 reimbursement (Section E
    row 81) had netted her to her exact $120,000 budget, i.e. **owes $0**. Lohith's was also
    stale ($112,573 vs the true $109,206). Kashi's happened to be right. All three now
    regenerate from the Capital sheet via `g['partners']` + the `partner cards` sub() rule:
      - Lohith: base $545k (own $300k + Kashi's $145k + Shashi's $100k transferred TO him);
        remaining = base − preopen(excl $5k advance) − acquisition paid − C147
      - Kashi / Shashi: deployed = their transfer + their direct spend; owes = budget − deployed
    **Any tile showing money must be traced to a workbook cell.** If it can't be, it will
    drift. When Reddy reports "this tile is wrong", check for a hard-coded value first.

---

## 9. DASHBOARD

`refresh_dashboard` regenerates data-driven parts via label-anchored regex `sub()` calls.
**A `WARN ... no match` in the output means a card silently did NOT update — investigate.**

**Auto-updating:** cash, all-time rev/exp/commission, daily bars, weekly grid, monthly
breakdown, ledger card+alert, ops↔capital card, net-after-all-exp, **Ops owe capital**,
**Net Cash Position** (added 27-Jul).

**NOT auto-refreshed (edit by hand, rarely change):** Acquisition card, Staff Roster,
Break-Even card, Pending Income alert.
**Partner ownership cards ARE now auto-refreshed** (fixed 27-Jul — see lesson 15).

### Two different "net" figures — don't conflate (Reddy asked about this)
- **Net after all exp & commission** = P&L: `revenue − expenses − commission`.
  Still carries the ONE-TIME $20,000 rent funded from capital.
- **Net Cash Position** = balance sheet: `cash − Capital advance (C147) − Owner Ledger`.
- **opnet** (console) = operating net EXCLUDING that one-time $20k rent.

---

## 10. STATE AS OF 27-JUL-2026 (end of day)

| Metric | Value |
|---|---|
| Cash anchor | **26-Jul-2026 = $15,395** (physical count, all forms incl. bank) |
| `cash_adjust` | **-700.00** (27-Jul transfer-to-me; was reset to 0 at the count) |
| All-time revenue | $253,931 |
| All-time expenses | $250,770 |
| Card commissions | $4,636 |
| Net after all exp & comm | **-$1,475** |
| Operating net (excl. $20k rent) | **+$23,161** |
| **Net Cash Position** | **+$724** |
| **Cash on Hand** | **$16,902** |
| Operations owe Capital (C147) | $16,727.81 |
| Owner Ledger | **Lohith HOLDS $549.50** (negative balance — first time it flipped) |
| Acquisition unpaid (D31) | $2,235 |
| Kashi still to transfer (F100) | $18,431.66 |
| Pooled: deployed / available | ties to $600,000 |
| Last Daily Log row | **74 = 27-Jul** |
| Last Expenses row | **548** |
| P&L categories | 24 (was 28 before the 27-Jul merge) |
| Propinas slots used | 21 of 30 (rows 63–92) — extend at ~3 left |

### 27-Jul was the first day entered via `eod.py` end-to-end
It exposed bugs 11–13 above. Verified afterwards by independent recompute from raw cells:
revenue, expenses, commission, cash, ledger and Capital all reconcile. `doctor.py` = ALL CLEAR.

### Open / on the horizon
- Propinas table fills around late Aug — extend rows + SUM range then.
- Extend SUMIFS ranges again when Expenses approaches row 5000.
- Consider auto-deducting commission inside the cash calc (proposed, not implemented).
- `cash_adjust` comment grows each `cash-adjust` call (audit trail) — trim if unwieldy.
- **After any day with a transfer-to-me / owner-paid item, sanity-check that Cash on Hand
  moved as expected.** `doctor.py` cannot catch a wrong cash figure — it has no
  independent source of truth for cash. That is what bug 12 hid behind.

---

## 11. WORKING STYLE WITH REDDY

- He gives EOD data in one message; he'll correct categories quickly — best-guess and flag.
- **Flag ambiguities explicitly rather than silently resolving them.**
- He catches real errors (he spotted the bogus "direct-spent under remaining" line) — when
  he pushes back on a number, dig into the root cause, don't paper over it.
- Verify after every change (recompute independently, check the daily bar + cash + net).
- Always mention pushing via VS Code at the end.
