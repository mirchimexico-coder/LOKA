#!/usr/bin/env python
# -*- coding: utf-8 -*-
# C:\LOKA\loka.py — LOKA bookkeeping toolkit (reusable). Never use edit_block on the xlsx.
import sys, os, json, argparse
from datetime import datetime, date
import datetime as dtmod
from copy import copy
import openpyxl

P = r'C:\LOKA\LOKA_Restaurant_Manager.xlsx'
sys.path.insert(0, r'C:\LOKA')
S_DAILY, S_EXP = 2, 3
OWNER = '💰 Owner Ledger'
MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec']

def backup(label='auto'):
    from backup_util import backup_excel
    return backup_excel(label)

def pdate(s):
    if isinstance(s, (datetime, date)):
        return s.date() if isinstance(s, datetime) else s
    s = str(s).strip()
    for fmt in ('%d-%b-%Y', '%d-%b', '%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y'):
        try:
            d = dtmod.datetime.strptime(s, fmt).date()
            if fmt == '%d-%b':
                d = d.replace(year=date.today().year)
            return d
        except ValueError:
            continue
    raise ValueError('Bad date: ' + s)

def _last_data_row(ws, key_col=3, start=8):
    r = ws.max_row
    while r > start and ws.cell(r, key_col).value is None:
        r -= 1
    return r

def _copyfmt(ws, src_row, dst_row, cols):
    for c in cols:
        s, d = ws.cell(src_row, c), ws.cell(dst_row, c)
        d.number_format = s.number_format
        d.font = copy(s.font); d.border = copy(s.border); d.alignment = copy(s.alignment)

def add_expenses(rows, do_backup=True):
    """rows: list of dict(date,desc,vendor,cat,amount,paid,method,notes)."""
    if do_backup: backup('add_expenses')
    wb = openpyxl.load_workbook(P); ex = wb.worksheets[S_EXP]
    last = _last_data_row(ex)
    n = 0
    for i, r in enumerate(rows):
        tr = last + 1 + i
        vals = [pdate(r['date']), r['desc'], r.get('vendor','Local purchase'),
                r['cat'], float(r['amount']), r.get('paid','Restaurant'),
                r.get('method','Cash'), r.get('status','✅ Paid'), r.get('notes','')]
        for c, v in enumerate(vals, start=2):
            ex.cell(tr, c, v)
        ex.cell(tr, 2).number_format = 'dd-mmm-yyyy'
        _copyfmt(ex, last, tr, range(2,11))
        ex.cell(tr,2).number_format = 'dd-mmm-yyyy'
        n += 1
    wb.save(P)
    return n, last+1, last+n

def close_day(d, card=0, cash=0, transfer=0, do_backup=True):
    """Set revenue for date d in Daily Log; create row with auto-SUMIFS if missing."""
    d = pdate(d)
    if do_backup: backup('close_day')
    wb = openpyxl.load_workbook(P); dl = wb.worksheets[S_DAILY]
    EXP = "'💸 Expenses'"
    # find row with this date
    target = None; lastrow = 7
    for r in range(8, dl.max_row+1):
        v = dl.cell(r,2).value
        if v is None: continue
        lastrow = r
        dd = v.date() if isinstance(v, datetime) else v
        if isinstance(dd, date) and dd == d:
            target = r; break
    if target is None:
        target = lastrow + 1
        dl.cell(target,2, d); dl.cell(target,2).number_format='dd-mmm-yyyy'
        dl.cell(target,3, MONTHS[d.month-1] and d.strftime('%a'))
        dl.cell(target,6, f'=IFERROR(D{target}+E{target}+G{target},0)')
        dl.cell(target,8, f"=SUMIFS({EXP}!$F$8:$F$500,{EXP}!$B$8:$B$500,\">=\"&B{target},{EXP}!$B$8:$B$500,\"<\"&(B{target}+1))")
        dl.cell(target,9, f'=IFERROR(F{target}-H{target},0)')
        dl.cell(target,10, f'=IFERROR(I{target}/F{target},0)')
        _copyfmt(dl, lastrow, target, range(2,13))
        dl.cell(target,2).number_format='dd-mmm-yyyy'
    dl.cell(target,4, float(card)); dl.cell(target,5, float(cash)); dl.cell(target,7, float(transfer))
    dl.cell(target,12,'Cierre via loka.py')
    wb.save(P)
    return target, float(card)+float(cash)+float(transfer)

def add_ledger(d, desc, spent=0, transferred=0, typ=None, status=None, notes='', do_backup=True):
    d = pdate(d)
    if do_backup: backup('owner_ledger')
    wb = openpyxl.load_workbook(P); ol = wb[OWNER]
    # find TOTALS row
    tr = None
    for r in range(4, ol.max_row+1):
        if str(ol.cell(r,1).value or '').startswith('TOTALS'):
            tr = r; break
    last = tr - 1            # last data row
    new = tr                 # entry goes where TOTALS was
    _copyfmt(ol, last, new, range(1,9))
    for c in range(1,9): ol.cell(new,c).value=None
    ol.cell(new,1, d); ol.cell(new,1).number_format='dd-mmm-yyyy'
    ol.cell(new,2, desc); ol.cell(new,3, typ or ('Transfer Received' if transferred else 'Expense - Personal'))
    if spent: ol.cell(new,4, float(spent))
    if transferred: ol.cell(new,5, float(transferred))
    ol.cell(new,6, f'=IFERROR(F{last}+D{new}-E{new},0)')
    ol.cell(new,7, status or ('✅ Received' if transferred else '⏳ Reimburse'))
    ol.cell(new,8, notes)
    # rewrite TOTALS one row down
    t = new + 1
    _copyfmt(ol, tr, t, range(1,9))
    ol.cell(t,1,'TOTALS & CURRENT BALANCE')
    ol.cell(t,4, f'=SUM(D4:D{new})'); ol.cell(t,5, f'=SUM(E4:E{new})'); ol.cell(t,6, f'=F{new}')
    ol.cell(t,7, f'=IF(F{new}>0,"Restaurant owes Lohith $"&TEXT(F{new},"#,##0"),IF(F{new}<0,"Lohith holds $"&TEXT(ABS(F{new}),"#,##0"),"Zero balance"))')
    wb.save(P)
    return new

def compute():
    from collections import defaultdict
    wb = openpyxl.load_workbook(P)
    dl = wb.worksheets[S_DAILY]; ex = wb.worksheets[S_EXP]
    last = _last_data_row(ex)
    ebd = defaultdict(float); cat_m = defaultdict(lambda: defaultdict(float))
    for r in range(8, last+1):
        v = ex.cell(r,2).value
        if not isinstance(v, (datetime, date)): continue
        dd = v.date() if isinstance(v, datetime) else v
        amt = float(ex.cell(r,6).value or 0)
        ebd[dd] += amt
        cat_m[(dd.year,dd.month)][str(ex.cell(r,5).value or 'Other')] += amt
    days = []
    inc_m = defaultdict(lambda: defaultdict(float)); dcount = defaultdict(int)
    for r in range(2, dl.max_row+1):
        v = dl.cell(r,2).value
        if v is None: continue
        dd = v.date() if isinstance(v, datetime) else v
        if not isinstance(dd, date): continue
        ca=float(dl.cell(r,4).value or 0); cs=float(dl.cell(r,5).value or 0); tf=float(dl.cell(r,7).value or 0)
        days.append((dd,ca,cs,tf,round(ebd.get(dd,0.0),2)))
        k=(dd.year,dd.month); inc_m[k]['Card']+=ca; inc_m[k]['Cash']+=cs; inc_m[k]['Transfer']+=tf
        if ca+cs+tf>0: dcount[k]+=1
    days.sort()
    rev=sum(c+k+t for _,c,k,t,_ in days); exp=round(sum(ebd.values()),2)
    trading=sum(1 for _,c,k,t,_ in days if c+k+t>0)
    out={'all_time':{'revenue':round(rev,2),'expenses':exp,'net':round(rev-exp,2),
         'net_plus_rent':round(rev-exp+20000,2),'trading_days':trading,
         'card':round(sum(d[1] for d in days),2),'cash':round(sum(d[2] for d in days),2),
         'transfer':round(sum(d[3] for d in days),2)},
         'by_month':{}, 'last_days':[(str(d),round(c+k+t,2),e,round(c+k+t-e,2)) for d,c,k,t,e in days[-7:]]}
    for k in sorted(inc_m):
        ti=sum(inc_m[k].values()); te=sum(cat_m[k].values())
        out['by_month'][f'{MONTHS[k[1]-1]} {k[0]}']={'income':round(ti,2),'expenses':round(te,2),
            'net':round(ti-te,2),'days':dcount[k],
            'top_cats':sorted(((c,round(a,2)) for c,a in cat_m[k].items()), key=lambda x:-x[1])[:6]}
    return out

def main():
    ap = argparse.ArgumentParser(description='LOKA bookkeeping toolkit')
    sub = ap.add_subparsers(dest='cmd')
    sub.add_parser('status')
    sub.add_parser('backup').add_argument('label', nargs='?', default='manual')
    a = sub.add_parser('add-expense')
    for x in ['date','desc','vendor','cat','amount']: a.add_argument('--'+x, required=True)
    a.add_argument('--paid', default='Restaurant'); a.add_argument('--method', default='Cash'); a.add_argument('--notes', default='')
    c = sub.add_parser('close-day')
    c.add_argument('--date', required=True); c.add_argument('--card', default=0); c.add_argument('--cash', default=0); c.add_argument('--transfer', default=0)
    b = sub.add_parser('add-expenses-json'); b.add_argument('file')
    sub.add_parser('refresh-dashboard')
    sub.add_parser('refresh-pl')
    args = ap.parse_args()
    if args.cmd=='status': print(json.dumps(compute(), indent=2, ensure_ascii=False))
    elif args.cmd=='refresh-dashboard': refresh_dashboard()
    elif args.cmd=='refresh-pl': print(refresh_pl())
    elif args.cmd=='backup': print(backup(args.label))
    elif args.cmd=='add-expense':
        n,a0,a1=add_expenses([{'date':args.date,'desc':args.desc,'vendor':args.vendor,'cat':args.cat,'amount':args.amount,'paid':args.paid,'method':args.method,'notes':args.notes}])
        print(f'added rows {a0}-{a1}')
    elif args.cmd=='add-expenses-json':
        n,a0,a1=add_expenses(json.load(open(args.file,encoding='utf-8'))); print(f'added {n} rows {a0}-{a1}')
    elif args.cmd=='close-day':
        row,tot=close_day(args.date,args.card,args.cash,args.transfer); print(f'row {row} revenue {tot}')
    else: ap.print_help()
# (entry point moved to end of file, after all helper defs)


# ============== DASHBOARD REFRESH ==============
DASH = r'C:\LOKA\dashboard.html'
WEEK_COLORS = ['#3b82f6','#22c55e','#f97316','#a855f7','#f59e0b','#ef4444','#06b6d4','#ec4899']
# Manual anchors that change rarely — update here when the situation changes.
CFG = dict(
    cash_anchor_date=date(2026,6,15), cash_anchor_amount=20283.0,
    cash_adjust=-1680.28,               # transfers to Lohith (Jun18 $400 + Jun19 $120 + Jun26 $165 = $685) + Jun3 $178 reimbursement + $817.28 MP commission on post-15-Jun card sales through Jun28 (W5 $467.96 + W6 $349.32) (clears on next physical count)
    commission_rate=0.0406,             # Mercado Pago est. on card revenue
    week1_start=date(2026,5,18),        # W1 begins here; weeks are 7-day blocks
)

def _money(n): return f"${n:,.0f}"
def _signed(n): return (f"+${n:,.0f}" if n>=0 else f"-${abs(n):,.0f}")

def _gather():
    """Full operating dataset from Daily Log + Expenses (raw values, no formulas)."""
    from collections import defaultdict
    wb = openpyxl.load_workbook(P)
    dl = wb.worksheets[S_DAILY]; ex = wb.worksheets[S_EXP]
    last = _last_data_row(ex)
    ebd = defaultdict(float); cat_m = defaultdict(lambda: defaultdict(float))
    for r in range(8, last+1):
        v = ex.cell(r,2).value
        if not isinstance(v,(datetime,date)): continue
        dd = v.date() if isinstance(v,datetime) else v
        amt = float(ex.cell(r,6).value or 0)
        ebd[dd]+=amt; cat_m[(dd.year,dd.month)][str(ex.cell(r,5).value or 'Other')]+=amt
    days=[]; inc_m=defaultdict(lambda: defaultdict(float)); dcount=defaultdict(int)
    for r in range(2, dl.max_row+1):
        v = dl.cell(r,2).value
        if v is None: continue
        dd = v.date() if isinstance(v,datetime) else v
        if not isinstance(dd,date): continue
        ca=float(dl.cell(r,4).value or 0); cs=float(dl.cell(r,5).value or 0); tf=float(dl.cell(r,7).value or 0)
        days.append((dd,ca,cs,tf,round(ebd.get(dd,0.0),2)))
        k=(dd.year,dd.month); inc_m[k]['Card']+=ca; inc_m[k]['Cash']+=cs; inc_m[k]['Transfer']+=tf
        if ca+cs+tf>0: dcount[k]+=1
    days.sort()
    # owner ledger balance (raw sums)
    ol = wb[OWNER]; spent=transferred=0.0
    for r in range(4, ol.max_row+1):
        if str(ol.cell(r,1).value or '').startswith('TOTALS'): break
        spent+=float(ol.cell(r,4).value or 0); transferred+=float(ol.cell(r,5).value or 0)
    return dict(days=days, ebd=ebd, cat_m=cat_m, inc_m=inc_m, dcount=dcount,
                ol_spent=round(spent,2), ol_transferred=round(transferred,2), ol_balance=round(spent-transferred,2))


def _build_bars(days):
    recent = days[-30:]
    maxrev = max((c+k+t for _,c,k,t,_ in recent), default=1) or 1
    out=[]
    for d,c,k,t,e in recent:
        rev=c+k+t; net=rev-e
        widx=(d-CFG['week1_start']).days//7
        col=WEEK_COLORS[widx % len(WEEK_COLORS)]
        w=max(6, round(rev/maxrev*100)) if rev>0 else 3
        cls='up' if net>=0 else 'down'
        out.append(f'<div class="bar-row"><div class="bar-day">{d.strftime("%d-%b")}</div>'
                   f'<div class="bar-track"><div class="bar-fill" style="width:{w}%;background:{col}">{_money(rev)}</div></div>'
                   f'<div class="bar-amt">{_money(rev)}</div><div class="bar-net {cls}">{_signed(net)}</div></div>')
    return ''.join(out)

def _build_weeks(days):
    from collections import defaultdict
    wk=defaultdict(list)
    for row in days: wk[(row[0]-CFG['week1_start']).days//7].append(row)
    idxs=sorted(wk)[-8:]
    maxrev=max((sum(c+k+t for _,c,k,t,_ in wk[i]) for i in idxs), default=1) or 1
    out=[]
    for i in idxs:
        rows=wk[i]; rev=sum(c+k+t for _,c,k,t,_ in rows); exp=sum(e for *_,e in rows); net=rev-exp
        ds=[r[0] for r in rows]; lo,hi=min(ds),max(ds)
        if lo.month==hi.month: rng=f'{lo.day}-{hi.day} {lo.strftime("%b")}'
        else: rng=f'{lo.strftime("%-d %b") if hasattr(lo,"strftime") else lo.day}-{hi.strftime("%d %b")}'
        col=WEEK_COLORS[i % len(WEEK_COLORS)]
        w=max(4, round(rev/maxrev*100))
        cls='up' if net>=0 else 'down'
        out.append(f'<div class="wk"><div class="wk-label">W{i+1} {rng}</div>'
                   f'<div class="wk-rev" style="color:{col}">{_money(rev)}</div>'
                   f'<div style="background:var(--border);height:3px;border-radius:2px;margin:5px 0;overflow:hidden;"><div style="width:{w}%;height:100%;background:{col};"></div></div>'
                   f'<div class="wk-row"><span>Expenses</span><span>{_money(exp)}</span></div>'
                   f'<div class="wk-row"><span>Net</span><span class="{cls}">{_signed(net)}</span></div></div>')
    return ''.join(out)

def _build_monthly(inc_m, cat_m, dcount):
    months=sorted(set(list(inc_m)+list(cat_m))); CUR=months[-1]
    lab={m: date(m[0],m[1],1).strftime('%b %Y') for m in months}
    ordered=list(reversed(months))[:6]
    cards=[]
    for m in ordered:
        ti=sum(inc_m[m].values()); te=sum(cat_m[m].values()); net=ti-te; cur=(m==CUR)
        nc='var(--green)' if net>=0 else 'var(--red)'; bd='var(--green)' if cur else 'var(--border)'
        bg='linear-gradient(135deg,#0d2818,#161b22)' if cur else '#161b22'
        bdg='<span style="font-size:.55rem;background:var(--green);color:#04140a;padding:1px 6px;border-radius:8px;font-weight:800;margin-left:6px;vertical-align:middle;">LIVE</span>' if cur else ''
        cards.append(f'<div style="flex:1;min-width:175px;border:1px solid {bd};border-radius:10px;padding:13px 15px;background:{bg};">'
            f'<div style="font-size:.72rem;text-transform:uppercase;letter-spacing:.05em;color:var(--muted);margin-bottom:8px;">{lab[m]}{bdg}</div>'
            f'<div style="display:flex;justify-content:space-between;font-size:.78rem;margin:3px 0;"><span style="color:var(--muted);">Income</span><span style="color:var(--green);font-weight:700;">{_money(ti)}</span></div>'
            f'<div style="display:flex;justify-content:space-between;font-size:.78rem;margin:3px 0;"><span style="color:var(--muted);">Expenses</span><span style="color:var(--red);font-weight:700;">{_money(te)}</span></div>'
            f'<div style="display:flex;justify-content:space-between;font-size:.9rem;margin:6px 0 2px;border-top:1px solid var(--border);padding-top:6px;"><span style="font-weight:700;">Net</span><span style="color:{nc};font-weight:800;">{_signed(net)}</span></div>'
            f'<div style="font-size:.62rem;color:var(--muted);margin-top:4px;">{dcount[m]} trading days</div></div>')
    cats=sorted({c for m in ordered for c in cat_m[m]}, key=lambda c:-sum(cat_m[m].get(c,0) for m in ordered))
    def tr(lbl,vals,color=None,bold=False,hdr=False):
        td=''
        for v in vals:
            st=('font-weight:700;' if bold else '')+(f'color:{color};' if color else '')
            td+=f'<td style="text-align:right;padding:4px 10px;{st}border-bottom:1px solid var(--border);">{v}</td>'
        ls='font-weight:700;' if (bold or hdr) else ''
        return f'<tr><td style="text-align:left;padding:4px 10px;{ls}border-bottom:1px solid var(--border);">{lbl}</td>{td}</tr>'
    th='<tr><th style="text-align:left;padding:5px 10px;color:var(--muted);font-size:.66rem;text-transform:uppercase;">Line</th>'+''.join(f'<th style="text-align:right;padding:5px 10px;color:var(--muted);font-size:.66rem;text-transform:uppercase;">{lab[m]}</th>' for m in ordered)+'</tr>'
    body=tr('INCOME',['' for _ in ordered],hdr=True)
    for typ in ['Card','Cash','Transfer']: body+=tr(typ,[_money(inc_m[m].get(typ,0)) for m in ordered])
    body+=tr('Total Income',[_money(sum(inc_m[m].values())) for m in ordered],color='var(--green)',bold=True)
    body+=tr('EXPENSES (high \u2192 low)',['' for _ in ordered],hdr=True)
    for c in cats: body+=tr(c,[_money(cat_m[m].get(c,0)) for m in ordered])
    body+=tr('Total Expenses',[_money(sum(cat_m[m].values())) for m in ordered],color='var(--red)',bold=True)
    body+=tr('NET PROFIT / LOSS',[_signed(sum(inc_m[m].values())-sum(cat_m[m].values())) for m in ordered],bold=True)
    table=f'<div style="overflow-x:auto;margin-top:14px;"><table style="width:100%;border-collapse:collapse;font-size:.74rem;"><thead>{th}</thead><tbody>{body}</tbody></table></div>'
    note='<p style="font-size:.66rem;color:var(--muted);margin-top:10px;line-height:1.5;">Showing up to last 6 months. Rent + maintenance + most salaries were front-loaded into June, so June reads as a loss while May (opening fortnight) shows a large profit; from July each month carries its own overhead.</p>'
    return (f'<div class="card" style="margin-bottom:16px;"><h3>&#128197; Monthly Breakdown</h3>'
            f'<div style="display:flex;gap:12px;flex-wrap:wrap;">{"".join(cards)}</div>{table}{note}</div>')


def refresh_dashboard(do_backup=True):
    import re
    g = _gather(); days = g['days']
    rev=sum(c+k+t for _,c,k,t,_ in days); exp=round(sum(g['ebd'].values()),2)
    net=round(rev-exp,2); netrent=round(net+20000,2)
    trading=sum(1 for _,c,k,t,_ in days if c+k+t>0)
    card=sum(d[1] for d in days)
    comm=round(card*CFG['commission_rate'])
    last_d,lc,lk,lt,le = days[-1]
    last_rev=lc+lk+lt; last_net=last_rev-le
    # cash on hand = anchor + nets after anchor + manual adjust
    roll=sum((c+k+t-e) for d,c,k,t,e in days if d>CFG['cash_anchor_date'])
    cash=round(CFG['cash_anchor_amount']+roll+CFG['cash_adjust'])
    datestr=last_d.strftime('%a %d %b %Y')
    if do_backup:
        import shutil, os
        bdir=r'C:\LOKA\Backup'; os.makedirs(bdir,exist_ok=True)
        shutil.copy(DASH, os.path.join(bdir, 'dashboard_'+datetime.now().strftime('%Y%m%d_%H%M%S')+'.html'))
    s=open(DASH,encoding='utf-8').read()
    def sub(pat, repl, label, n=None):
        nonlocal s
        s2,cnt=re.subn(pat, repl, s)
        if n is not None and cnt!=n: print(f'  WARN {label}: expected {n} got {cnt}')
        elif cnt==0: print(f'  WARN {label}: no match')
        s=s2
    M=lambda x: _money(x).replace('$','\\$')  # not used; placeholder
    # --- scalars (label-anchored) ---
    sub(r'Updated: [^<]+', f'Updated: {datestr} (EOD)', 'header')
    sub(r'(<div class="bval">)\$[\d,]+(</div>)', lambda m: m.group(1)+_money(cash)+m.group(2), 'cash')
    sub(r'(All-Time Revenue</div><div[^>]*>)\$[\d,]+', lambda m: m.group(1)+_money(rev), 'banner rev')
    sub(r'(All-Time Expenses</div><div[^>]*>)\$[\d,]+', lambda m: m.group(1)+_money(exp), 'banner exp')
    sub(r'(Commission</div><div[^>]*>)\$[\d,]+', lambda m: m.group(1)+_money(comm), 'banner comm')
    sub(r'(<div class="klabel">Total Revenue</div><div class="kval">)\$[\d,]+(</div><div[^>]*>)\d+( trading days)',
        lambda m: m.group(1)+_money(rev)+m.group(2)+str(trading)+m.group(3), 'kpi rev')
    sub(r'(<div class="klabel">Total Expenses</div><div class="kval">)\$[\d,]+', lambda m: m.group(1)+_money(exp), 'kpi exp')
    sub(r'(<div class="klabel">Operating Net</div><div class="kval">)[+\-]?\$[\d,]+', lambda m: m.group(1)+_signed(netrent), 'kpi net')
    sub(r'(MP Commission</div><div class="kval">)\$[\d,]+', lambda m: m.group(1)+_money(comm), 'kpi comm')
    sub(r'(<div class="alabel">)[^<]+(</div><div class="aval">)Rev [^<]+',
        lambda m: m.group(1)+last_d.strftime('%a %d %b')+m.group(2)+f'Rev {_money(last_rev)} &mdash; Exp {_money(le)} &mdash; Net {_signed(last_net)}', 'daily alert')
    sub(r'Dashboard updated [^<]+',
        f'Dashboard updated {datestr} (EOD) &bull; auto-refreshed by loka.py', 'footer')
    # owner ledger card + alert
    sub(r'(Spent personally</span><span class="sval" style="color:var\(--red\);">)\$[\d,]+', lambda m: m.group(1)+_money(g['ol_spent']), 'ledger spent')
    sub(r'(Transfers received</span><span class="sval" style="color:var\(--green\);">)\$[\d,]+', lambda m: m.group(1)+_money(g['ol_transferred']), 'ledger recv')
    sub(r'(Restaurant owes Lohith</span><span class="sval"[^>]*>)\$[\d,]+', lambda m: m.group(1)+_money(g['ol_balance']), 'ledger bal')
    sub(r'(Lohith Ledger</div><div class="aval">Restaurant owes )\$[\d,]+', lambda m: m.group(1)+_money(g['ol_balance']), 'ledger alert')
    # --- block regens ---
    bars=_build_bars(days)
    s=re.sub(r'(<h3>Daily Revenue &mdash; Last 30 Days</h3>).*?(<div style="display:flex;gap:14px;margin-top:10px;)',
             lambda m: m.group(1)+bars+m.group(2), s, count=1, flags=re.DOTALL)
    weeks=_build_weeks(days)
    s=re.sub(r'(<div class="week-grid">).*?(</div>\s*<p class="note")',
             lambda m: m.group(1)+weeks+m.group(2), s, count=1, flags=re.DOTALL)
    monthly=_build_monthly(g['inc_m'], g['cat_m'], g['dcount'])
    s=re.sub(r'<div class="card" style="margin-bottom:16px;"><h3>&#128197; Monthly Breakdown</h3>.*?(<h3 style="margin-bottom:10px;">&#128202; All-Time)',
             lambda m: monthly+m.group(1), s, count=1, flags=re.DOTALL)
    open(DASH,'w',encoding='utf-8',newline='\n').write(s)
    # keep the Monthly P&L sheet's structure current too (formulas auto-update values themselves)
    try:
        info=refresh_pl(do_backup=False); print(f'  Monthly P&L synced: {info["months"]} months, {info["cats"]} categories')
    except Exception as e:
        print(f'  (Monthly P&L sync skipped: {e})')
    print(f'Dashboard refreshed @ {datestr}: rev {_money(rev)} exp {_money(exp)} opnet {_signed(netrent)} | cash {_money(cash)} | comm {_money(comm)} | ledger {_money(g["ol_balance"])}')


# ============== MONTHLY P&L (LIVE FORMULAS) ==============
def refresh_pl(do_backup=True):
    """Rebuild the Monthly P&L sheet with LIVE SUMIFS/COUNTIFS formulas that
    pull from Daily Log + Expenses, so it auto-updates whenever data changes."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from collections import defaultdict
    if do_backup: backup('monthly_pl')
    wb = openpyxl.load_workbook(P)
    dl = wb.worksheets[S_DAILY]; ex = wb.worksheets[S_EXP]; pl = wb.worksheets[8]
    DLN = "'" + dl.title + "'"; EXN = "'" + ex.title + "'"
    # detect months (from both sheets) and categories (from expenses), in data order
    def parse(v):
        from datetime import datetime as _dt
        if isinstance(v,_dt): return v.date()
        if isinstance(v,date): return v
        return None
    monthset=set(); cat_tot=defaultdict(float)
    for r in range(2, dl.max_row+1):
        d=parse(dl.cell(r,2).value)
        if d: monthset.add((d.year,d.month))
    last=ex.max_row
    while last>7 and ex.cell(last,3).value is None: last-=1
    for r in range(8,last+1):
        d=parse(ex.cell(r,2).value)
        if d:
            monthset.add((d.year,d.month))
            cat_tot[str(ex.cell(r,5).value or 'Other')]+=float(ex.cell(r,6).value or 0)
    months=sorted(monthset)
    cats=[c for c,_ in sorted(cat_tot.items(), key=lambda x:-x[1])]
    label={m: date(m[0],m[1],1).strftime('%b %Y') for m in months}
    ncol=1+len(months)
    def nextm(m):
        y,mo=m; return (y+1,1) if mo==12 else (y,mo+1)
    # date-range criteria builder for a sheet/col
    def rng(sheet,col,m):
        y,mo=m; ny,nm=nextm(m)
        return (f'{sheet}!${col}$8:${col}$500,{sheet}!$B$8:$B$500,">="&DATE({y},{mo},1),'
                f'{sheet}!$B$8:$B$500,"<"&DATE({ny},{nm},1)')
    return _write_pl(wb, pl, months, cats, label, ncol, DLN, EXN, rng, P)


def _write_pl(wb, pl, months, cats, label, ncol, DLN, EXN, rng, Ppath):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    # strip
    for mr in list(pl.merged_cells.ranges): pl.unmerge_cells(str(mr))
    for r in range(1, max(pl.max_row,60)+1):
        for c in range(1, max(pl.max_column,12)+1):
            cell=pl.cell(r,c); cell.value=None
            cell.font=Font(); cell.fill=PatternFill(fill_type=None); cell.border=Border()
            cell.alignment=Alignment(); cell.number_format='General'
    for col in 'ABCDEFGHIJKLMN': pl.column_dimensions[col].width=9.0
    TITLE=Font(bold=True,color='FFFFFF',size=13); TFILL=PatternFill('solid',fgColor='0D2818')
    HDR=Font(bold=True,color='FFFFFF',size=10); HFILL=PatternFill('solid',fgColor='1F4E2C')
    BOLD=Font(bold=True); NORM=Font()
    thin=Side(style='thin',color='D9D9D9'); BORD=Border(left=thin,right=thin,top=thin,bottom=thin)
    MONEY='$#,##0.00'; INTF='#,##0'
    RIGHT=Alignment(horizontal='right'); LEFT=Alignment(horizontal='left'); CTR=Alignment(horizontal='center')
    def cel(r,c,val,font=NORM,fill=None,fmt=None,align=None):
        x=pl.cell(r,c,val); x.font=font; x.border=BORD
        if fill: x.fill=fill
        if fmt: x.number_format=fmt
        x.alignment=align or (RIGHT if c>1 else LEFT)
    def hrow(r,text):
        cel(r,1,text,font=HDR,fill=HFILL,align=LEFT)
        for j,m in enumerate(months): cel(r,2+j,label[m],font=HDR,fill=HFILL,align=RIGHT)
    # title
    pl.merge_cells(start_row=1,start_column=1,end_row=1,end_column=ncol)
    t=pl.cell(1,1,'LOKA — MONTHLY P&L   (live formulas — auto-updates from Daily Log & Expenses)')
    t.font=TITLE; t.fill=TFILL; t.alignment=CTR
    for c in range(1,ncol+1): pl.cell(1,c).fill=TFILL
    pl.row_dimensions[1].height=26
    COL=lambda j: get_column_letter(2+j)
    r=3
    hrow(r,'SUMMARY'); r+=1
    R_INC=r; cel(r,1,'Total Income',font=BOLD)
    for j,m in enumerate(months): cel(r,2+j,f'=SUMIFS({rng(DLN,"F",m)})',fmt=MONEY)
    r+=1
    R_EXP=r; cel(r,1,'Total Expenses',font=BOLD)
    for j,m in enumerate(months): cel(r,2+j,f'=SUMIFS({rng(EXN,"F",m)})',fmt=MONEY)
    r+=1
    cel(r,1,'NET PROFIT / LOSS',font=BOLD)
    for j in range(len(months)): cel(r,2+j,f'={COL(j)}{R_INC}-{COL(j)}{R_EXP}',font=BOLD,fmt=MONEY)
    r+=1
    R_DAYS=r; cel(r,1,'Trading Days')
    for j,m in enumerate(months): cel(r,2+j,f'=COUNTIFS({DLN}!$F$8:$F$500,">0",'+rng(DLN,"F",m).split(",",1)[1]+')',fmt=INTF)
    r+=1
    cel(r,1,'Avg Income / Day')
    for j in range(len(months)): cel(r,2+j,f'=IFERROR({COL(j)}{R_INC}/{COL(j)}{R_DAYS},0)',fmt=MONEY)
    r+=2
    hrow(r,'INCOME BY TYPE'); r+=1
    R_CARD=r
    for col,name in [('D','Card'),('E','Cash'),('G','Transfer')]:
        cel(r,1,name)
        for j,m in enumerate(months): cel(r,2+j,f'=SUMIFS({rng(DLN,col,m)})',fmt=MONEY)
        r+=1
    cel(r,1,'Total Income',font=BOLD)
    for j in range(len(months)): cel(r,2+j,f'={COL(j)}{R_CARD}+{COL(j)}{R_CARD+1}+{COL(j)}{R_CARD+2}',font=BOLD,fmt=MONEY)
    r+=2
    hrow(r,'EXPENSES BY CATEGORY (high \u2192 low)'); r+=1
    for cat in cats:
        cel(r,1,cat)
        for j,m in enumerate(months):
            crit=f'{EXN}!$F$8:$F$500,{EXN}!$E$8:$E$500,$A{r},{EXN}!$B$8:$B$500,">="&DATE({m[0]},{m[1]},1),{EXN}!$B$8:$B$500,"<"&DATE({(m[0]+1) if m[1]==12 else m[0]},{1 if m[1]==12 else m[1]+1},1)'
            cel(r,2+j,f'=SUMIFS({crit})',fmt=MONEY)
        r+=1
    R_TEXP=r; cel(r,1,'TOTAL EXPENSES',font=BOLD)
    for j,m in enumerate(months): cel(r,2+j,f'=SUMIFS({rng(EXN,"F",m)})',font=BOLD,fmt=MONEY)
    r+=1
    cel(r,1,'NET PROFIT / LOSS',font=BOLD)
    for j in range(len(months)): cel(r,2+j,f'={COL(j)}{R_INC}-{COL(j)}{R_TEXP}',font=BOLD,fmt=MONEY)
    pl.column_dimensions['A'].width=34
    for j in range(len(months)): pl.column_dimensions[get_column_letter(2+j)].width=15
    pl.sheet_view.showGridLines=False
    try: wb.calculation.fullCalcOnLoad = True   # force Excel to recompute all formulas on open
    except Exception: pass
    wb.save(Ppath)
    return dict(months=len(months), cats=len(cats), rows=r)


if __name__ == '__main__':
    main()
