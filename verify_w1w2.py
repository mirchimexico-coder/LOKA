import sys, os
sys.path.insert(0, r'C:\Restaurant')
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)

import openpyxl
wb = openpyxl.load_workbook(r'C:\Restaurant\LOKA_Restaurant_Manager.xlsx')
ws_dl = wb['📅 Daily Log']
ws_exp = wb['💸 Expenses']

# ── WEEK 1 ONLY (18–24 May) ──────────────────────────────────────────────────
w1_dates = ['18-May-2026','19-May-2026','20-May-2026','21-May-2026',
            '22-May-2026','23-May-2026','24-May-2026']

w1_rev = 0; w1_exp = 0; w1_card = 0; w1_cash = 0
for r in range(8, 50):
    b = str(ws_dl[f'B{r}'].value or '')
    if not b: break
    if any(d in b for d in w1_dates):
        card = ws_dl[f'D{r}'].value or 0
        cash = ws_dl[f'E{r}'].value or 0
        exp  = ws_dl[f'H{r}'].value or 0
        w1_rev  += card + cash
        w1_card += card
        w1_cash += cash
        w1_exp  += exp

w1_comm = round(w1_card * 0.0406, 2)

# W1 propinas in expenses
w1_prop_paid    = 1030.00  # John $515 + Debi $515
w1_prop_pending = 168.00   # Victor

# W1 expenses EXCL propinas
w1_exp_excl = w1_exp - w1_prop_paid - w1_prop_pending

# W1 closing balance
w1_mp   = 5553
w1_cash_close = 3170
w1_close = w1_mp + w1_cash_close

# W1 propinas pool = closing (incl commission) - net (excl propinas)
w1_net_excl = w1_rev - w1_exp_excl - w1_comm
w1_pool = (w1_close + w1_comm) - w1_net_excl

print("=" * 55)
print("WEEK 1 PROPINAS VERIFICATION (18–24 May)")
print("=" * 55)
print(f"  Revenue:                     ${w1_rev:>10,}")
print(f"  Card revenue:                ${w1_card:>10,}")
print(f"  Commission (4.06%):         -${w1_comm:>10,.2f}")
print(f"  Expenses excl propinas:     -${w1_exp_excl:>10,.2f}")
print(f"  Net (excl propinas):         ${w1_net_excl:>10,.2f}")
print(f"  W1 Closing (MP+Cash):        ${w1_close:>10,}")
print(f"  + Commission added back:     ${w1_comm:>10,.2f}")
print(f"  Closing incl commission:     ${w1_close+w1_comm:>10,.2f}")
print(f"  W1 PROPINAS POOL:            ${w1_pool:>10,.2f}")
print(f"")
print(f"  Paid to John:               -${w1_prop_paid/2:>10,.2f}")
print(f"  Paid to Debi:               -${w1_prop_paid/2:>10,.2f}")
print(f"  W1 BALANCE (Victor):         ${w1_pool - w1_prop_paid:>10,.2f}")
print(f"  Victor (pending):            ${w1_prop_pending:>10,.2f}")

# ── WEEK 2 ONLY (25–31 May) ──────────────────────────────────────────────────
w2_dates = ['25-May-2026','26-May-2026','27-May-2026','28-May-2026',
            '29-May-2026','30-May-2026','31-May-2026']

w2_rev = 0; w2_exp = 0; w2_card = 0
for r in range(8, 50):
    b = str(ws_dl[f'B{r}'].value or '')
    if not b: break
    if any(d in b for d in w2_dates):
        card = ws_dl[f'D{r}'].value or 0
        cash = ws_dl[f'E{r}'].value or 0
        exp  = ws_dl[f'H{r}'].value or 0
        w2_rev  += card + cash
        w2_card += card
        w2_exp  += exp

w2_comm = round(w2_card * 0.0406, 2)

# W2 closing
w2_mp   = 9448
w2_cash_close = 8300
w2_close = w2_mp + w2_cash_close

# W2 pool = W2 closing incl all commission - W1 net closing incl commission - W2 net from records
# Simpler: total pool - W1 pool used
total_pool = 2091.00
w2_pool = total_pool - w1_pool  # what's left after W1 allocation

print(f"\n{'=' * 55}")
print("WEEK 2 PROPINAS (25–31 May)")
print("=" * 55)
print(f"  W2 Revenue:                  ${w2_rev:>10,}")
print(f"  W2 Card:                     ${w2_card:>10,}")
print(f"  W2 Commission (4.06%):      -${w2_comm:>10,.2f}")
print(f"  W2 Expenses:                -${w2_exp:>10,.2f}")
print(f"  W2 Closing (MP+Cash):        ${w2_close:>10,}")
print(f"")
print(f"  Total pool (W1+W2):          ${total_pool:>10,.2f}")
print(f"  W1 pool used:               -${w1_pool:>10,.2f}")
print(f"  W2 PROPINAS AVAILABLE:       ${total_pool - w1_pool:>10,.2f}")
print(f"  Less Victor (W1 pending):   -${w1_prop_pending:>10,.2f}")
print(f"  W2 FOR JOHN + DEBI:          ${total_pool - w1_pool - w1_prop_pending:>10,.2f}")
