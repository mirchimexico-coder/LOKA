# -*- coding: utf-8 -*-
"""
LOKA — one-shot EOD from a plain-text day file.

USAGE:   py eod.py day.txt            (preview only — shows what it WILL do)
         py eod.py day.txt --apply    (actually write it)

Write day.txt in the same format you already use, e.g.:

    DATE: 24-Jul
    MP: 2455  BBVA: 0  SOFT: 0  CASH: 3315
    TRANSFER-TO-ME: 0   TRANSFER-TO-BANK: 0
    BBVA COMMISSION: 1.9%
    SOFT COMMISSION: 0
    NO-BILL: Pan Molido 72, Oxxo Azucar 24, Queso Panela 590
    PAID BY ME: Pollo 234
    SALARY: Samu 1200, John 3750
    NOTE: anything you want on the row

Every line is optional. Aliases accepted:
    MP / MERCADO PAGO / CARD MERCADO PAGO
    SOFT / CARD SOFT RESTAURANT
    TRANSFER-TO-ME / TRANSFER TO ME
    TRANSFER-TO-BANK / TRANSFER-TO-RESTAURANT-BANK
Sunday/closed day: just omit the card/cash lines — it records expenses only.
"""
import sys, os, re, json, unicodedata
from datetime import date, datetime
sys.path.insert(0, r'C:\LOKA')
import loka

# ---------- vendor / keyword -> category ----------
KEYWORD_CAT = [
 (('pollo','bistec','carne','jamon','chorizo','res','cerdo','milanesa','arrachera','carnic'), 'Ingredients - Meat'),
 (('tilapia','pescado','camaron','mariscos','brisa del mar'),                                 'Ingredients - Meat & Fish'),
 (('verdura','vegetable','lechuga','tomate','jitomate','cebolla','chile','elote','abastos',
   'papa','zanahoria','brocoli','espinaca','aguacate','ajo','pepino','calabaza'),             'Ingredients - Vegetables'),
 (('fruta','fruit','limon','naranja','platano','manzana','fresa','guayaba','piña','papaya'),  'Ingredients - Fruit'),
 (('pan ','pan','bolillo','tortilla','baguette','pan molido','pan reales','masa'),            'Ingredients - Bread'),
 (('queso','leche','crema','yogur','mantequilla','panela','lacteo'),                          'Ingredients - Dairy'),
 (('huevo','egg'),                                                                            'Ingredients - Eggs'),
 (('agua','refresco','cerveza','vino','bebida','jugo','soda','cafe','café','reca'),           'Ingredients - Beverages'),
 (('arroz','azucar','azúcar','aceite','sal ','harina','frijol','pasta','especia','condimento',
   'vinagre','salsa','abarrote'),                                                             'Ingredients - Pantry'),
 (('vaso','bolsa','plato','servilleta','contenedor','empaque','desechable','guimar','stretch',
   'tapa','poliseda','termico','térmico'),                                                    'Packaging/Disposables'),
 (('hielo','clorox','limpieza','detergente','jabon','jabón','escoba','trapo','fibra'),         'Kitchen Supplies'),
 (('renta','rent','arrendador'),                                                              'Rent'),
 (('gas','tomza'),                                                                            'Utilities/Gas'),
 (('internet','telmex','suscripcion','suscripción','subscription','soft restaurant app'),      'Utilities/Internet'),
 (('luz','electricidad','electricity','cfe','recibo de luz'),                                  'Utilities/Electricity'),
 (('propina','tip'),                                                                          'Staff - Propinas'),
 (('sueldo','salario','salary','nomina','nómina','payroll'),                                  'Staff - Salary'),
 (('sam','costco','chedraui','heb','h-e-b','walmart','bodega','soriana','tres b','3b',
   'mercado','super'),                                                                        'Supermarket/General'),
]
VENDOR_HINT = [
 (('oxxo',),'OXXO'), (('sam',),"Sam's Club"), (('costco',),'Costco'),
 (('chedraui','chedruai','chedrahui'),'Chedraui'), (('abastos',),'Central de Abastos'),
 (('guimar',),'Guimar Polietilenos'), (('brisa',),'Brisa del Mar'),
 (('pan reales',),'Pan Reales'), (('tres b','3b'),'Tres B'), (('heb','h-e-b'),'H-E-B'),
 (('soriana',),'Soriana'), (('walmart',),'Walmart'), (('reca',),'Reca'),
]
def _norm(s):
    s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().lower()
    return re.sub(r'\s+',' ',s).strip()

# ---------- learned rules (things Reddy has taught it) ----------
RULES_PATH = r'C:\LOKA\categories.json'
CATEGORIES = [
 'Ingredients - Vegetables','Ingredients - Meat','Ingredients - Meat & Fish',
 'Ingredients - Pantry','Ingredients - Bread','Ingredients - Dairy',
 'Ingredients - Beverages','Ingredients - Fruit','Ingredients - Eggs',
 'Ingredients - Other','Ingredients - Desserts',
 'Staff - Salary','Staff - Advance','Staff - Propinas',
 'Supermarket/General','Kitchen Supplies','Packaging/Disposables','Supplies/Other',
 'Utilities/Internet','Utilities/Gas','Utilities/Electricity','Software/Subscription',
 'Rent','Office Supplies','Maintenance',
]
def load_rules():
    try:
        with open(RULES_PATH, encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}
def save_rule(keyword, category):
    r = load_rules(); r[_norm(keyword)] = category
    with open(RULES_PATH,'w',encoding='utf-8',newline='\n') as f:
        json.dump(r, f, ensure_ascii=False, indent=2, sort_keys=True)
    return r

UNKNOWN = 'Supplies/Other'   # what guess_cat returns when it has no idea

def guess_cat(desc):
    d=_norm(desc)
    # 1) anything Reddy has taught wins outright (longest keyword first = most specific)
    for k in sorted(load_rules(), key=len, reverse=True):
        if k and k in d: return load_rules()[k]
    # supermarket names win first (a store name shouldn't be guessed from its letters).
    # prefix match (no trailing boundary) so "sams"/"sam's" and misspellings still hit.
    for k in ('sam','costco','chedraui','chedruai','chedrahui','h-e-b','heb','walmart',
              'soriana','tres b','3b','bodega','city market','la comer','smart'):
        if re.search(rf'(?<![a-z]){re.escape(k)}', d): return 'Supermarket/General'
    # utensils / tableware: DISPOSABLE (for customers) vs REUSABLE (kitchen tools).
    # Both are called "utensils" in conversation, but they are different lines.
    UTENSIL = ('utensil','utensilio','cubierto','cuchillo','cuchara','tenedor','plato',
               'vaso','taza','charola','bandeja','coladera','colador','sarten','olla',
               'tabla','pinza','espatula','cucharon','batidor','rallador','tupper')
    # Spanish plurals add -s after a vowel but -es after a consonant
    # (vaso->vasos, sarten->sartenes, colador->coladores), so allow (e)?s
    if any(re.search(rf'(?<![a-z]){re.escape(k)}(?:e?s)?(?![a-z])', d) for k in UTENSIL):
        DISPOSABLE = ('desechable','plast','reyma','unicel','biocup','papel','carton',
                      'termico','poliseda','dart')
        return 'Packaging/Disposables' if any(x in d for x in DISPOSABLE) else 'Kitchen Supplies'
    for keys,cat in KEYWORD_CAT:
        for k in keys:
            kn=_norm(k)
            # short keywords must match as whole words to avoid substring collisions
            # (e.g. 'res' inside 'tres b'), but ALLOW a plural 's' -> 'vaso' matches 'vasos'
            if len(kn)<=4:
                if re.search(rf'(?<![a-z]){re.escape(kn)}s?(?![a-z])', d): return cat
            elif kn in d:
                return cat
    return 'Supplies/Other'

def guess_vendor(desc):
    d=_norm(desc)
    for keys,v in VENDOR_HINT:
        for k in keys:
            if re.search(rf'(?<![a-z]){re.escape(_norm(k))}', d): return v
    return 'No bill'

# ---------- parsing ----------
NUM = r'(-?[\d,]+(?:\.\d+)?)'
def _f(x):
    return float(str(x).replace(',','').replace('$','')) if x not in (None,'') else 0.0

def field(txt, *names):
    """Find 'NAME: value' allowing the value to be followed by another FIELD:"""
    for n in names:
        m=re.search(rf'{n}\s*:?\s*\$?{NUM}', txt, re.I)
        if m: return _f(m.group(1))
    return None

def line_items(txt, label):
    """Parse 'LABEL: item 123, item2 45' -> [(desc, amount)]"""
    m=re.search(rf'^{label}\s*:?(.*)$', txt, re.I|re.M)
    if not m: return []
    body=m.group(1).strip()
    if not body or _norm(body) in ('none','no','0','n/a'): return []
    out=[]
    for chunk in re.split(r'[,;]|\s+y\s+', body):
        chunk=chunk.strip()
        if not chunk: continue
        mm=re.match(rf'^(.*?)\s*\$?{NUM}\s*$', chunk)
        if mm and mm.group(1).strip():
            out.append((mm.group(1).strip(), _f(mm.group(2))))
    return out

def parse(txt):
    # ignore everything after a '-----' separator (the how-to block in the template)
    txt = re.split(r'^\s*-{5,}\s*$', txt, maxsplit=1, flags=re.M)[0]
    d={}
    m=re.search(r'DATE\s*:?\s*([0-9]{1,2}[-/ ][A-Za-z]{3,}[-/ ]?[0-9]{0,4})', txt, re.I)
    if not m:
        raise SystemExit(
            "\n  >> The DATE line is empty (or the wrong shape).\n"
            "     Open the file and put a date after DATE:  e.g.   DATE: 28-Jul\n"
            "     Then run the preview again.\n")
    raw=m.group(1).strip().replace('/','-').replace(' ','-')
    parts=raw.split('-')
    if len(parts)==2: raw=f'{parts[0]}-{parts[1]}-{date.today().year}'
    d['date']=loka.pdate(raw)
    d['mp']    = field(txt,'CARD MERCADO PAGO','MERCADO PAGO','MP CARD',r'\bMP\b') or 0
    d['bbva']  = field(txt,'BBVA CARD',r'\bBBVA\b') or 0
    d['soft']  = field(txt,'CARD SOFT RESTAURANT','SOFT RESTAURANT','SOFT CARD',r'\bSOFT\b') or 0
    d['cash']  = field(txt,'CASH','EFECTIVO') or 0
    d['t_me']  = field(txt,'TRANSFER-TO-ME','TRANSFER TO ME','TRANSFERENCIA A MI') or 0
    d['t_bank']= field(txt,'TRANSFER-TO-RESTAURANT-BANK','TRANSFER-TO-BANK','TRANSFER TO BANK') or 0
    # commissions
    sc = field(txt,'SOFT RESTAURANT COMMISSION','SOFT COMMISSION')
    d['softcomm']= sc if sc is not None else 0
    mb=re.search(r'BBVA\s*COMMISSION\s*:?\s*([\d.]+)\s*%', txt, re.I)
    if mb: d['bbvacomm']=round(d['bbva']*float(mb.group(1))/100,2)
    else:  d['bbvacomm']= field(txt,'BBVA COMMISSION') or round(d['bbva']*0.019,2)
    mc = field(txt,'MP COMMISSION','MERCADO PAGO COMMISSION')
    d['mpcomm_note']=mc
    d['no_bill'] = line_items(txt,'NO-BILL') + line_items(txt,'NO BILL')
    d['by_me']   = line_items(txt,'PAID BY ME') + line_items(txt,'PAGADO POR MI')
    d['salary']  = line_items(txt,'SALARY') + line_items(txt,'SUELDO')
    # free-form "Paid 1200 salary to Samu"
    for mm in re.finditer(rf'paid\s+\$?{NUM}\s+salary\s+to\s+(\w+)', txt, re.I):
        d['salary'].append((mm.group(2), _f(mm.group(1))))
    mnote=re.search(r'^NOTE\s*:?(.*)$', txt, re.I|re.M)
    d['note']=mnote.group(1).strip() if mnote else ''
    return d

def build(d):
    """Return (expense_rows, plan_lines)."""
    rows=[]; plan=[]
    ds=d['date'].strftime('%d-%b-%Y')
    for desc,amt in d['no_bill']:
        rows.append(dict(date=ds,desc=desc,vendor=guess_vendor(desc),cat=guess_cat(desc),
                         amount=amt,paid='Restaurant',method='Cash',notes='No receipt; cash'))
    for desc,amt in d['by_me']:
        rows.append(dict(date=ds,desc=desc,vendor=guess_vendor(desc),cat=guess_cat(desc),
                         amount=amt,paid='Lohith',method='Cash',
                         notes='Pagado por Lohith (bolsillo); reembolsable via Owner Ledger'))
    for who,amt in d['salary']:
        rows.append(dict(date=ds,desc=f'{who} - salary',vendor='Payroll',cat='Staff - Salary',
                         amount=amt,paid='Restaurant',method='Cash',notes='Weekly salary'))
    rev=d['mp']+d['bbva']+d['soft']+d['cash']+d['t_me']+d['t_bank']
    exp=sum(r['amount'] for r in rows)
    plan.append(f"DATE            : {d['date']:%a %d-%b-%Y}")
    plan.append(f"REVENUE         : {rev:,.2f}   (MP {d['mp']:,.0f} | BBVA {d['bbva']:,.0f} | Soft {d['soft']:,.0f} | cash {d['cash']:,.0f} | transfer {d['t_me']+d['t_bank']:,.0f})")
    plan.append(f"  commissions   : MP auto {d['mp']*0.0406:,.2f} | BBVA {d['bbvacomm']:,.2f} | Soft {d['softcomm']:,.2f}")
    if d['mpcomm_note']: plan.append(f"  (you wrote MP commission {d['mpcomm_note']:,.2f} - MP col is auto-calculated, informational only)")
    plan.append(f"EXPENSES        : {exp:,.2f}  ({len(rows)} rows)")
    unsure=0
    for r in rows:
        flag=' [PAID BY YOU]' if r['paid']=='Lohith' else ''
        mark='  <-- ?? not recognised' if r['cat']==UNKNOWN else ''
        if mark: unsure+=1
        plan.append(f"    {r['amount']:>9,.2f}  {r['desc'][:34]:<34} {r['cat']:<26} {r['vendor']}{flag}{mark}")
    if unsure:
        plan.append(f"    ^^ {unsure} item(s) could not be categorised - they will go to '{UNKNOWN}'.")
        plan.append(f"       Use menu option 8 (Teach a category) to fix this permanently.")
    plan.append(f"DAY NET         : {rev-exp:,.2f}")
    if d['t_me']:
        plan.append(f"LEDGER          : + transfer received {d['t_me']:,.2f}  (cash-adjust -{d['t_me']:,.2f})")
    byme=sum(a for _,a in d['by_me'])
    if byme:
        plan.append(f"LEDGER          : + you spent {byme:,.2f}          (cash-adjust +{byme:,.2f})")
    if d['t_bank']:
        plan.append(f"NOTE            : {d['t_bank']:,.2f} went to the restaurant bank - revenue, stays in restaurant, no ledger")
    return rows, plan

def main():
    if len(sys.argv)<2:
        print(__doc__); return
    path=sys.argv[1]; apply='--apply' in sys.argv
    txt=open(path,encoding='utf-8').read()
    d=parse(txt); rows,plan=build(d)
    print('\n'.join(plan))
    if not apply:
        print('\n--- PREVIEW ONLY.  Re-run with --apply to write it. ---'); return
    print('\napplying...')
    loka.backup(f"eod_{d['date']:%b%d}".lower())
    if rows:
        n,a0,a1=loka.add_expenses(rows, do_backup=False); print(f'  expenses: added {n} rows {a0}-{a1}')
    has_rev = any([d['mp'],d['bbva'],d['soft'],d['cash'],d['t_me'],d['t_bank']])
    if has_rev:
        r,tot=loka.close_day(d['date'], d['mp'], d['cash'], d['t_me']+d['t_bank'],
                             d['soft'], d['softcomm'], d['bbva'], d['bbvacomm'], do_backup=False)
        print(f'  daily log: row {r}, revenue {tot:,.2f}')
    else:
        print('  no revenue given -> shopping-only day, no Daily Log row')
    if d['t_me']:
        loka.add_ledger(d['date'], f"Transfer-to-me (customer payment {d['date']:%d-%b})",
                        transferred=d['t_me'], notes='Customer paid Lohith personal acct', do_backup=False)
        loka.cash_adjust_add(-d['t_me'], f"transfer-to-me {d['date']:%d-%b}")
        print(f'  ledger + cash-adjust: -{d["t_me"]:,.2f}')
    byme=sum(a for _,a in d['by_me'])
    if byme:
        desc=' + '.join(f'{x} ${a:,.0f}' for x,a in d['by_me'])
        loka.add_ledger(d['date'], f"{desc} (paid by Lohith)", spent=byme,
                        notes='Fronted from own pocket; excluded from till via cash_adjust', do_backup=False)
        loka.cash_adjust_add(byme, f"owner-paid {d['date']:%d-%b}")
        print(f'  ledger + cash-adjust: +{byme:,.2f}')
    loka.refresh_all(do_backup=False)
    print('\nDONE.  Now push with VS Code.')

if __name__=='__main__':
    main()
