# -*- coding: utf-8 -*-
"""
LOKA - the jobs that used to need a Claude session.

  py tools.py recount        re-anchor cash to a physical count
  py tools.py repay          record money paid back to Capital
  py tools.py tips           record propinas (card tips handed to staff)
  py tools.py settle         settle the Owner Ledger with Lohith
  py tools.py report         weekly / monthly summary
  py tools.py restore        undo - roll back to a previous backup
  py tools.py capex          expense paid with capital money
  py tools.py move           fix a day: move money between payment types
                             (e.g. cash that was really a transfer to me)
"""
import sys, os, io, re, glob, shutil
from copy import copy
from datetime import datetime, date, timedelta
sys.path.insert(0, r'C:\LOKA')
import openpyxl, loka

P = loka.P
def ask(q, default=None):
    v = input(f"   {q}{' ['+str(default)+']' if default is not None else ''}: ").strip()
    return v or (str(default) if default is not None else '')
def money(q, default=None):
    while True:
        v = ask(q, default).replace(',','').replace('$','')
        try: return float(v)
        except ValueError: print("   Please type a number.")
def askdate(q='Date (e.g. 02-Aug or 02-Aug-2026)', default=None):
    default = default or date.today().strftime('%d-%b-%Y')
    while True:
        v = ask(q, default)
        try: return loka.pdate(v)
        except Exception: print("   Try a format like 02-Aug-2026.")
def confirm(msg):
    return ask(f"{msg} (y/n)", 'n').lower().startswith('y')

# ---------------------------------------------------------------- recount
def recount():
    print("\n== RE-ANCHOR CASH TO A PHYSICAL COUNT ==")
    pos = loka.compute()['position']
    reserve = pos.get('capital_reserve', 0.0)
    print(f"   The books currently think you have: ${pos['cash_on_hand']:,.2f}  (operating cash)")
    if reserve:
        print(f"   Capital reserve in the company account: ${reserve:,.2f}  (NOT operating cash)")
        print(f"   -> expected total if you count everything: ${pos['cash_on_hand']+reserve:,.2f}")
    counted = money("Total you actually counted (all forms, incl bank)")
    amt = round(counted - reserve, 2)
    d   = askdate("Date you counted it")
    if reserve:
        print(f"\n   Counted ${counted:,.2f}  -  capital reserve ${reserve:,.2f}  =  operating cash ${amt:,.2f}")
    print(f"\n   New anchor: {d:%d-%b-%Y} = ${amt:,.2f}")
    print("   This also resets the running cash adjustment to 0 (the count absorbs all drift).")
    if not confirm("   Apply?"): print("   cancelled."); return
    loka.backup(f'recount_{d:%b%d}'.lower())
    src = io.open(loka.__file__, encoding='utf-8').read()
    src = re.sub(r'cash_anchor_date=date\([^)]*\), cash_anchor_amount=[\d.]+',
                 f'cash_anchor_date=date({d.year},{d.month},{d.day}), cash_anchor_amount={amt}', src)
    src = re.sub(r'(cash_adjust=)-?[\d.]+(,\s*#)[^\n]*',
                 rf'\g<1>0.00\g<2> RESET at physical count {d:%d-%b-%Y} = ${amt:,.2f}. '
                 'Only add deltas dated AFTER that.', src)
    io.open(loka.__file__,'w',encoding='utf-8',newline='\n').write(src)
    print("   done - re-anchored.")
    import subprocess
    subprocess.call([sys.executable, os.path.join(os.path.dirname(loka.__file__), 'loka.py'), 'refresh-all'])

# ---------------------------------------------------------------- repay capital
def repay():
    print("\n== MONEY PAID BACK TO CAPITAL ==")
    wb = openpyxl.load_workbook(P); cap = wb.worksheets[0]
    owed = float(cap.cell(147,3).value or 0)
    print(f"   Operations currently owes Capital: ${owed:,.2f}")
    amt = money("How much was paid back")
    d   = askdate()
    if amt > owed and not confirm(f"   That is more than the ${owed:,.2f} owed. Continue?"): return
    new = round(owed-amt, 2)
    to_reserve = ask("Did it stay in the COMPANY ACCOUNT as capital reserve? (y/n, n = it went to Lohith)", 'y').lower().startswith('y')
    print(f"\n   Advance owed:  ${owed:,.2f}  ->  ${new:,.2f}")
    print(f"   Operating cash on hand will drop by ${amt:,.2f}.")
    if to_reserve:
        r0 = round(float(cap['C169'].value or 0) + float(cap['C170'].value or 0), 2)
        print(f"   Capital reserve (company acct): ${r0:,.2f}  ->  ${r0+amt:,.2f}")
    else:
        print(f"   The money is now capital held by Lohith.")
    print("   This is NOT a P&L expense - it is a balance-sheet movement.")
    if not confirm("   Apply?"): print("   cancelled."); return
    loka.backup(f'repay_capital_{d:%b%d}'.lower())
    cap.cell(147,3,new)
    if to_reserve:
        cap['C169'] = round(float(cap['C169'].value or 0) + amt, 2)
        cap['E169'] = str(cap['E169'].value or '') + f' | {d:%d-%b-%Y}: +${amt:,.2f} ops repayment kept in company account.'
    prev = cap.cell(149,3).value or 0
    cap.cell(149,2,'  >> Repaid to Capital from operating cash')
    cap.cell(149,3, round(float(prev)-amt,2)); cap.cell(149,3).number_format='$#,##0.00'
    cap.cell(149,5, f'Cumulative repaid from operating cash (latest {d:%d-%b-%Y} ${amt:,.2f}). '
                    'Reduces advance owed & Cash-on-Hand; NOT a P&L expense.')
    wb.calculation.calcMode='auto'; wb.calculation.fullCalcOnLoad=True
    wb.save(P)
    loka.cash_adjust_add(-amt, f'capital repayment {d:%d-%b}')
    print(f"   done - advance now ${new:,.2f}")
    loka.refresh_all(do_backup=False)

# ---------------------------------------------------------------- propinas
def tips():
    print("\n== PROPINAS (card tips handed to staff) ==")
    print("   Reminder: tips are a PASS-THROUGH, not a restaurant expense.")
    print("   Only money the restaurant adds ON TOP of collected tips is a real cost.")
    wb = openpyxl.load_workbook(P); sp = wb['👥 Staff & Payroll']
    LAST = 122                       # table extended 04-Aug from 92 -> 122 (60 slots)
    r = 63
    while r <= LAST and sp.cell(r,1).value not in (None,''): r += 1
    if r > LAST:
        print("   The propinas table is FULL (rows 63-122). Ask Claude to extend it."); return
    print(f"   {LAST-r+1} free slot(s) left.")
    d = askdate()
    label = ask("Which week is this for (e.g. 'Week 11: 03-09 Aug')", f"week to {d:%d-%b}")
    entries = []
    for who in ('John','Duvi/Debi','Samu'):
        v = ask(f"Amount for {who} (blank/0 = skip)", '0').replace(',','').replace('$','')
        try: a = float(v)
        except ValueError: a = 0
        if a: entries.append((who, a))
    while confirm("   Add someone else?"):
        w = ask("Name"); a = money(f"Amount for {w}")
        if a: entries.append((w, a))
    if not entries: print("   nothing to record."); return
    print(f"\n   Will record {len(entries)} entr(y/ies), total ${sum(a for _,a in entries):,.2f}:")
    for w,a in entries: print(f"     {d:%d-%b}  {w} ({label})  ${a:,.2f}")
    if not confirm("   Apply?"): print("   cancelled."); return
    loka.backup(f'propinas_{d:%b%d}'.lower())
    tmpl = r-1 if r > 63 else 63
    for i,(w,a) in enumerate(entries):
        rr = r+i
        if rr > LAST: print("   ran out of slots - stopped early."); break
        for c in (1,2,3,4):
            s_,dd = sp.cell(tmpl,c), sp.cell(rr,c)
            dd.font=copy(s_.font); dd.fill=copy(s_.fill); dd.border=copy(s_.border)
            dd.alignment=copy(s_.alignment); dd.number_format=s_.number_format
        sp.cell(rr,1,d); sp.cell(rr,1).number_format=sp.cell(tmpl,1).number_format
        sp.cell(rr,2,f'{w} ({label})'); sp.cell(rr,4,a)
    wb.calculation.calcMode='auto'; wb.calculation.fullCalcOnLoad=True
    wb.save(P)
    print(f"   done - recorded in rows {r}-{r+len(entries)-1}. Totals are NOT affected (pass-through).")

# ---------------------------------------------------------------- settle ledger
def settle():
    print("\n== SETTLE THE OWNER LEDGER ==")
    g = loka._gather(); bal = g['ol_balance']
    if abs(bal) < 0.01: print("   The ledger is already square."); return
    if bal > 0:
        print(f"   The restaurant owes you ${bal:,.2f}.")
        print("   Settling = the restaurant pays you -> cash goes DOWN.")
    else:
        print(f"   You are holding ${abs(bal):,.2f} of restaurant money.")
        print("   Settling = you hand it back -> cash goes UP.")
    amt = money("Amount settled", f"{abs(bal):.2f}")
    d   = askdate()
    if not confirm("   Apply?"): print("   cancelled."); return
    loka.backup(f'settle_ledger_{d:%b%d}'.lower())
    # typ/status passed explicitly: the add_ledger defaults ('Expense - Personal' /
    # 'Reimburse') made a hand-back read as money OWED to Lohith (fixed 18-Sep).
    if bal > 0:
        loka.add_ledger(d, 'Settlement - restaurant paid Lohith', transferred=amt,
                        typ='Settlement', status='\u2705 Settled',
                        notes='Ledger settled from operating cash', do_backup=False)
        loka.cash_adjust_add(-amt, f'ledger settlement {d:%d-%b}')
    else:
        loka.add_ledger(d, 'Settlement - Lohith returned restaurant money', spent=amt,
                        typ='Settlement', status='\u2705 Settled',
                        notes='Lohith handed back cash he was holding', do_backup=False)
        loka.cash_adjust_add(amt, f'ledger settlement {d:%d-%b}')
    print("   done.")
    loka.refresh_all(do_backup=False)

# ---------------------------------------------------------------- report
def report():
    g = loka._gather()
    days = [(d,c,k,t,e) for d,c,k,t,e in g['days']]
    if not days: print("no data"); return
    last = max(d for d,_,_,_,_ in days)
    def block(title, sel):
        rows=[x for x in days if sel(x[0])]
        rev=sum(c+k+t for _,c,k,t,_ in rows)+sum(g['soft_bd'].get(d,0)+g['bbva_bd'].get(d,0) for d,_,_,_,_ in rows)
        exp=sum(e for *_,e in rows)
        trading=sum(1 for _,c,k,t,_ in rows if c+k+t>0)
        print(f"\n  {title}")
        print(f"    revenue {rev:>12,.2f}   expenses {exp:>12,.2f}   net {rev-exp:>12,.2f}")
        if trading: print(f"    {trading} trading day(s), avg revenue {rev/trading:,.2f}/day")
    print("\n" + "="*60); print(f"  LOKA REPORT  (latest data {last:%d-%b-%Y})"); print("="*60)
    wk = last - timedelta(days=last.weekday())
    block(f"THIS WEEK (from {wk:%d-%b})", lambda d: d >= wk)
    block(f"LAST 7 DAYS", lambda d: d > last - timedelta(days=7))
    block(f"THIS MONTH ({last:%B})", lambda d: d.month == last.month and d.year == last.year)
    block("ALL TIME", lambda d: True)
    s = loka.compute()['position']
    print(f"\n  cash on hand      ${s['cash_on_hand']:,.2f}")
    print(f"  owed to Capital   ${s['owed_to_capital']:,.2f}")
    print(f"  owner ledger      {'restaurant owes you' if s['owner_ledger']>0 else 'you hold'} ${abs(s['owner_ledger']):,.2f}")
    print(f"  net cash position ${s['net_cash_position']:,.2f}")
    if s.get('capital_reserve'):
        print(f"  capital reserve   ${s['capital_reserve']:,.2f}   (company acct, not operating cash)")
    print()

# ---------------------------------------------------------------- restore
def restore():
    print("\n== UNDO - RESTORE A PREVIOUS BACKUP ==")
    b = sorted(glob.glob(r'C:\LOKA\Backup\*.xlsx'), key=os.path.getmtime, reverse=True)[:15]
    if not b: print("   no backups found."); return
    for i,f in enumerate(b,1):
        t = datetime.fromtimestamp(os.path.getmtime(f))
        print(f"     {i:>2}. {t:%d-%b %H:%M}   {os.path.basename(f)[24:-5] or 'auto'}")
    v = ask("Which one to restore (0 = cancel)", '0')
    if not v.isdigit() or int(v) < 1 or int(v) > len(b): print("   cancelled."); return
    pick = b[int(v)-1]
    print(f"\n   Restoring: {os.path.basename(pick)}")
    print("   Your CURRENT workbook will be backed up first, so this is reversible.")
    if not confirm("   Are you sure?"): print("   cancelled."); return
    loka.backup('before_restore')
    shutil.copy(pick, P)
    print("   restored.")
    loka.refresh_all(do_backup=False)

# ---------------------------------------------------------------- capital-paid expense
def capex():
    """An operating cost paid with PARTNER CAPITAL rather than restaurant money."""
    print("\n== EXPENSE PAID FROM CAPITAL ==")
    print("   Use this when partner capital paid an operating bill (rent, electricity...)")
    print("   rather than the restaurant's own money.")
    print("   It stays a real cost in the P&L, does NOT touch cash on hand, and")
    print("   INCREASES what Operations owes Capital.\n")
    import eod
    wb = openpyxl.load_workbook(P); cap = wb.worksheets[0]
    owed = float(cap.cell(147,3).value or 0)
    print(f"   Operations currently owes Capital: ${owed:,.2f}")
    d = askdate()
    items = []
    while True:
        desc = ask("\n   What was it (blank = done)")
        if not desc: break
        amt = money(f"   Amount for '{desc}'")
        guess = eod.guess_cat(desc)
        print(f"   suggested category: {guess}")
        if not confirm("   use that category?"):
            for i,c in enumerate(eod.CATEGORIES,1): print(f"     {i:>2}. {c}")
            v = ask("   Number")
            if v.isdigit() and 1 <= int(v) <= len(eod.CATEGORIES): guess = eod.CATEGORIES[int(v)-1]
        note = ask("   Note (optional)", "Pagado con capital de socios")
        items.append(dict(date=d.strftime('%d-%b-%Y'), desc=desc, vendor=ask("   Vendor", "No bill"),
                          cat=guess, amount=amt, paid='Capital', method='Transfer', notes=note))
    if not items: print("   nothing to record."); return
    tot = sum(i['amount'] for i in items)
    print("\n   " + "-"*54)
    for i in items: print(f"     {i['amount']:>11,.2f}  {i['desc'][:26]:<26} {i['cat']}")
    print(f"     {tot:>11,.2f}  TOTAL")
    print(f"\n   These will be recorded as paid by CAPITAL.")
    print(f"   Operating cash on hand: UNCHANGED (restaurant money did not pay).")
    print(f"   Owed to Capital: ${owed:,.2f}  ->  ${owed+tot:,.2f}")
    reserve = round(float(cap['C169'].value or 0) + float(cap['C170'].value or 0), 2)
    from_reserve = False
    if reserve > 0:
        from_reserve = ask(f"Paid from the CAPITAL RESERVE in the company account (${reserve:,.2f})? "
                           "(y/n, n = a partner paid directly)", 'y').lower().startswith('y')
        if from_reserve:
            if tot > reserve + 0.005:
                print(f"   The reserve only has ${reserve:,.2f}. Split it: record the reserve part here,"
                      " the rest as a separate entry with 'n'."); return
            print(f"   Capital reserve: ${reserve:,.2f}  ->  ${reserve-tot:,.2f}")
    if not confirm("\n   Apply?"): print("   cancelled."); return
    loka.backup(f'capex_{d:%b%d}'.lower())
    n,a0,a1 = loka.add_expenses(items, do_backup=False)
    print(f"   expenses: added {n} row(s) {a0}-{a1}")
    wb2 = openpyxl.load_workbook(P); c2 = wb2.worksheets[0]
    new = round(float(c2.cell(147,3).value or 0) + tot, 2)
    c2.cell(147,3, new)
    c2.cell(147,5, f'Operating costs funded from partner capital, net of repayments. '
                   f'Latest addition {d:%d-%b-%Y} ${tot:,.2f}. Still owed by Operations: ${new:,.2f}.')
    if from_reserve:
        c2['C170'] = round(float(c2['C170'].value or 0) - tot, 2)
        c2['E170'] = str(c2['E170'].value or '') + f' | {d:%d-%b-%Y}: -${tot:,.2f} ' + ', '.join(i['desc'] for i in items)[:80]
    wb2.calculation.calcMode='auto'; wb2.calculation.fullCalcOnLoad=True
    wb2.save(P)
    print(f"   owed to Capital is now ${new:,.2f}")
    loka.refresh_all(do_backup=False)
    print("\n   done. Now push with VS Code.")

# ---------------------------------------------------------------- move between payment types
# key: (label, Daily Log column, commission column or None, CFG rate key or None)
PAYTYPES = {
    '1': ('Cash',                        5,  None, None),
    '2': ('Card - Mercado Pago',         4,  None, None),   # MP commission is a formula on col D
    '3': ('Card - BBVA',                 25, 26,  'bbva_commission_rate'),
    '4': ('Card - Soft Restaurant',      21, 22,  'soft_commission_rate'),
    '5': ('Transfer to restaurant BBVA', 7,  None, None),
    '6': ('Transfer to ME (my account)', 7,  None, None),
}
def move():
    print("\n== MOVE MONEY BETWEEN PAYMENT TYPES ON A DAY ALREADY RECORDED ==")
    print("   Use when you typed an amount under the wrong payment type.")
    print("   The day's TOTAL revenue does not change.\n")
    d = askdate()
    wb = openpyxl.load_workbook(P); dl = wb.worksheets[loka.S_DAILY]
    row = None
    for r in range(8, dl.max_row+1):
        v = dl.cell(r,2).value
        dd = v.date() if isinstance(v, datetime) else v
        if isinstance(dd, date) and dd == d: row = r; break
    if row is None:
        print(f"   {d:%d-%b-%Y} is not in the Daily Log. Record the day first."); return
    val = lambda c: float(dl.cell(row, c).value or 0)
    print(f"\n   Currently recorded for {d:%a %d-%b-%Y}:")
    for k, (lab, col, _, _) in PAYTYPES.items():
        if k == '6': continue
        print(f"     {lab:<30} ${val(col):>10,.2f}" + ("   (both transfer types share this column)" if k == '5' else ''))
    print()
    for k, (lab, *_ ) in PAYTYPES.items(): print(f"     {k}. {lab}")
    src = ask("MOVE FROM (number)"); dst = ask("MOVE TO   (number)")
    if src not in PAYTYPES or dst not in PAYTYPES or src == dst:
        print("   Pick two different numbers from the list."); return
    amt = money("Amount to move")
    if amt <= 0: print("   Amount must be more than 0."); return
    sl, scol, scom, srate = PAYTYPES[src]; tl, tcol, tcom, trate = PAYTYPES[dst]
    if scol != tcol and amt > val(scol) + 0.005:
        print(f"   Only ${val(scol):,.2f} is recorded under {sl}. Nothing changed."); return
    delta = {c: 0.0 for c in (4,5,7,21,22,25,26)}
    if scol != tcol:
        delta[scol] -= amt; delta[tcol] += amt
    for com, rate, sign, lab in ((scom, srate, -1, sl), (tcom, trate, +1, tl)):
        if com:
            dflt = round(amt * loka.CFG.get(rate, 0), 2)
            c = money(f"{lab} commission change (enter 0 if none)", dflt)
            delta[com] += sign * c
    me_in, me_out = dst == '6', src == '6'
    print(f"\n   PLAN for {d:%d-%b-%Y}:")
    print(f"     {sl}  -${amt:,.2f}   ->   {tl}  +${amt:,.2f}")
    for c, lab in ((22,'Soft commission'),(26,'BBVA commission')):
        if delta[c]: print(f"     {lab} {delta[c]:+,.2f}")
    if me_in:  print(f"     Owner Ledger: +${amt:,.2f} transferred to you   |  cash adjust -{amt:,.2f}")
    if me_out: print(f"     Owner Ledger: reverse ${amt:,.2f} transferred   |  cash adjust +{amt:,.2f}")
    if not confirm("   Apply?"): print("   cancelled."); return
    loka.backup(f'move_{d:%b%d}'.lower())
    loka.close_day(d, card=delta[4], cash=delta[5], transfer=delta[7], soft=delta[21],
                   softcomm=delta[22], bbva=delta[25], bbvacomm=delta[26],
                   do_backup=False, mode='add')
    if me_in:
        loka.add_ledger(d, f"Transfer-to-me (customer payment {d:%d-%b}, corrected from {sl})",
                        transferred=amt, notes=f'Moved from {sl} via tools.py move', do_backup=False)
        loka.cash_adjust_add(-amt, f"transfer-to-me {d:%d-%b} (moved from {sl})")
    if me_out:
        loka.add_ledger(d, f"Reversal - transfer-to-me {d:%d-%b} was really {tl}",
                        transferred=-amt, typ='Correction', status='Corrected',
                        notes=f'Moved to {tl} via tools.py move', do_backup=False)
        loka.cash_adjust_add(amt, f"undo transfer-to-me {d:%d-%b} (really {tl})")
    loka.refresh_all(do_backup=False)
    print("\n   done. Now push with VS Code.")

CMDS = dict(recount=recount, repay=repay, tips=tips, settle=settle, report=report,
            restore=restore, capex=capex, move=move)
if __name__ == '__main__':
    a = sys.argv[1] if len(sys.argv) > 1 else ''
    if a in CMDS:
        try: CMDS[a]()
        except KeyboardInterrupt: print("\n   cancelled.")
    else:
        print(__doc__)
