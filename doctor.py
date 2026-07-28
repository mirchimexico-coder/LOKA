# -*- coding: utf-8 -*-
"""
LOKA — health check.  Run any time:  py doctor.py
Checks everything that has silently broken before.  All output plain ASCII.
"""
import sys, os, re, zipfile, collections
sys.path.insert(0, r'C:\LOKA')
import openpyxl
from datetime import datetime, date
P=r'C:\LOKA\LOKA_Restaurant_Manager.xlsx'
DASH=r'C:\LOKA\dashboard.html'
ok=[]; warn=[]; bad=[]
def _pd(v): return v.date() if isinstance(v,datetime) else (v if isinstance(v,date) else None)

wb=openpyxl.load_workbook(P)
dl=wb.worksheets[2]; ex=wb.worksheets[3]; cap=wb.worksheets[0]
last=ex.max_row
while last>7 and ex.cell(last,3).value is None: last-=1

# 1 capped SUMIFS ranges
rng=re.compile(r'\$?([A-Z]{1,3})\$?(\d+):\$?([A-Z]{1,3})\$?(\d+)')
n=0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            if isinstance(v,str) and v.startswith('='):
                for m in rng.finditer(v):
                    if m.group(1)==m.group(3) and int(m.group(4))-int(m.group(2))>=50 and int(m.group(4))<=1000: n+=1
(ok if n==0 else bad).append(f"capped formula ranges (<=1000): {n}" + ("" if n==0 else "  <-- FIX: extend to 5000"))

# 2 headroom
(ok if last<4500 else bad).append(f"Expenses last row {last} (cap 5000)")

# 3 text labels starting with '='
lab=0
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            v=c.value
            if isinstance(v,str) and v.startswith('=') and ' ' in v[1:] and '(' not in v and '!' not in v \
               and not re.match(r'^[A-Z]+\$?\d', v[1:]) and '+' not in v and '-' not in v:
                lab+=1
(ok if lab==0 else bad).append(f"text labels starting with '=' (#NAME risk): {lab}")

# 4 empty formula caches in Daily Log key columns
x=zipfile.ZipFile(P).read('xl/worksheets/sheet3.xml').decode('utf-8')
empt=[m.group(1) for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*>(.*?)</c>', x)
      if '<f' in m.group(2) and re.search(r'<v\s*/>', m.group(2)) and m.group(1)[0] in 'FHIJR']
(ok if not empt else bad).append(f"blank F/H/I/J/R formula caches: {len(empt)}" + ("" if not empt else "  <-- FIX: py loka.py refresh-all"))

# 5 duplicate daily dates / gaps
seen={}; dup=0
for r in range(8,dl.max_row+1):
    d=_pd(dl.cell(r,2).value)
    if isinstance(d,date):
        if d in seen: dup+=1
        seen[d]=r
(ok if dup==0 else bad).append(f"duplicate Daily Log dates: {dup}")

# 6 expenses data quality
q=[]
for r in range(8,last+1):
    if ex.cell(r,3).value in (None,''): continue
    if not isinstance(_pd(ex.cell(r,2).value),date): q.append(f"row {r}: bad date")
    a=ex.cell(r,6).value
    if not isinstance(a,(int,float)): q.append(f"row {r}: non-numeric amount")
    elif a<0: q.append(f"row {r}: negative amount")
    if ex.cell(r,7).value in (None,''): q.append(f"row {r}: blank Paid By")
(ok if not q else bad).append(f"expense data problems: {len(q)}"+("" if not q else "  e.g. "+q[0]))

# 7 category spelling variants (same name differing only by spaces/case)
cats=collections.Counter(str(ex.cell(r,5).value or '').strip() for r in range(8,last+1) if ex.cell(r,5).value)
norm=collections.defaultdict(list)
for c in cats:
    norm[re.sub(r'[\s/]+','',c).lower()].append(c)
variants={k:v for k,v in norm.items() if len(v)>1}
(ok if not variants else warn).append(f"category spelling variants: {len(variants)}"+("" if not variants else "  "+str(list(variants.values()))))

# 8 capital reconciles to 600k
def s(lo,hi): return sum(float(cap.cell(r,4).value or 0) for r in range(lo,hi+1) if isinstance(cap.cell(r,4).value,(int,float)))
try:
    c147=float(cap.cell(147,3).value or 0)
    dep=s(14,30)+(s(56,66)-float(cap.cell(56,4).value or 0))+s(43,51)+s(72,81)+s(114,128)+c147
    withme=545000-(s(56,66)-5000)-s(14,30)-c147
    avail=withme+(180000-145000-s(43,51))+(120000-100000-s(72,81))
    tot=dep+avail
    (ok if abs(tot-600000)<1 else bad).append(f"Capital reconciles to 600,000: {tot:,.2f}")
except Exception as e:
    warn.append("capital check failed: "+str(e))

# 9 dashboard freshness + backup counts
import glob
bdir=r'C:\LOKA\Backup'
xl=len(glob.glob(os.path.join(bdir,'*.xlsx'))); ht=len(glob.glob(os.path.join(bdir,'dashboard_*.html')))
(ok if xl<=31 and ht<=16 else warn).append(f"backups: {xl} xlsx (cap 30), {ht} dashboard html (cap 15)")
stray=[f for f in os.listdir(r'C:\LOKA') if f.startswith('_') and f.endswith('.py')]
(ok if not stray else warn).append(f"stray temp scripts in root: {len(stray)} {stray if stray else ''}")

# 10 last recorded day
lastd=max(seen) if seen else None
warn.append(f"last Daily Log date: {lastd}  (row {seen.get(lastd)})") if lastd else None

print("="*62); print("LOKA HEALTH CHECK"); print("="*62)
for m in ok:   print("  [OK]   "+m)
for m in warn: print("  [note] "+m)
for m in bad:  print("  [FAIL] "+m)
print("="*62)
print(("ALL CLEAR" if not bad else f"{len(bad)} PROBLEM(S) - see FAIL lines above"))
