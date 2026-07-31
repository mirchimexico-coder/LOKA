# -*- coding: utf-8 -*-
"""
LOKA - read a receipt photo, locally. No internet, nothing uploaded.

  py scan.py                 read every photo in the drop folder
  py scan.py <file|folder>   read a specific photo or folder
  py scan.py <x> --text      also print the full OCR text

Uses the OCR engine already built into Windows, so there is nothing to install.
HEIC (iPhone) photos are handled automatically.

It reports its CONFIDENCE. Anything not marked HIGH should be checked by eye -
never let a guessed number into the books.
"""
import sys, os, re, glob, subprocess, unicodedata
from datetime import datetime, date

HERE      = r'C:\LOKA'
DROP      = os.path.join(HERE, 'Bills', '_DROP_BILLS_HERE')
PS1       = os.path.join(HERE, 'ocr_win.ps1')
WORK      = os.path.join(HERE, '_scan_tmp')
IMG_EXT   = ('.heic','.heif','.jpg','.jpeg','.png','.bmp','.tif','.tiff','.webp')

def to_png(src):
    """Windows OCR can't read HEIC; convert and upscale for small text."""
    os.makedirs(WORK, exist_ok=True)
    out = os.path.join(WORK, os.path.splitext(os.path.basename(src))[0] + '.png')
    from PIL import Image, ImageOps, ImageEnhance
    try:
        import pillow_heif; pillow_heif.register_heif_opener()
    except Exception:
        pass
    im = Image.open(src)
    im = ImageOps.exif_transpose(im).convert('L')       # honour phone rotation
    w, h = im.size
    if max(w, h) < 1600:                                 # upscale small photos
        f = 1600 / max(w, h)
        im = im.resize((int(w*f), int(h*f)), Image.LANCZOS)
    im = ImageOps.autocontrast(im)
    im = ImageEnhance.Sharpness(im).enhance(1.6)
    im.save(out)
    return out

def ocr(png):
    r = subprocess.run(['powershell','-NoProfile','-ExecutionPolicy','Bypass','-File',PS1,png],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    err = (r.stderr or '')
    m = re.search(r'engine=([\w-]+)', err)
    global LAST_ENGINE
    LAST_ENGINE = m.group(1) if m else '?'
    if r.returncode != 0:
        return '', err.strip()[:200]
    return r.stdout or '', ''

LAST_ENGINE = '?'

def _norm(s):
    return unicodedata.normalize('NFKD', s).encode('ascii','ignore').decode().lower()

MONTHS = dict(ene=1,feb=2,mar=3,abr=4,may=5,jun=6,jul=7,ago=8,sep=9,oct=10,nov=11,dic=12,
              jan=1,apr=4,aug=8,dec=12)
def find_date(txt):
    t = _norm(txt)
    # 23/jul/2026 or 23-jul-26  (OCR sometimes drops the separators, so allow spaces too)
    m = re.search(r'(\d{1,2})\s*[/\-. ]\s*([a-z]{3})[a-z]*\s*[/\-. ]\s*(\d{2,4})', t)
    if m and m.group(2) in MONTHS:
        y = int(m.group(3)); y += 2000 if y < 100 else 0
        try: return date(y, MONTHS[m.group(2)], int(m.group(1)))
        except ValueError: pass
    m = re.search(r'(\d{1,2})\s*[/\-.]\s*(\d{1,2})\s*[/\-.]\s*(\d{2,4})', t)
    if m:
        d_, mo, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        y += 2000 if y < 100 else 0
        if mo > 12: d_, mo = mo, d_
        try: return date(y, mo, d_)
        except ValueError: pass
    return None

VENDORS = [('oxxo','OXXO'),('sam','Sam\'s Club'),('costco','Costco'),
           ('chedraui','Chedraui'),('chedruai','Chedraui'),('guimar','Guimar Polietilenos'),
           ('brisa','Brisa del Mar'),('soriana','Soriana'),('walmart','Walmart'),
           ('bodega','Bodega Aurrera'),('heb','H-E-B'),('tres b','Tres B'),
           ('abastos','Central de Abastos'),('pan reales','Pan Reales')]
def find_vendor(txt):
    t = _norm(txt)
    for k, v in VENDORS:
        if k in t: return v
    for line in [l.strip() for l in txt.split('\n') if l.strip()][:4]:
        if len(line) > 3 and not re.search(r'\d{3}', line):
            return line[:34]
    return None

def money_list(s):
    out = []
    for m in re.finditer(r'\$?\s*(\d{1,3}(?:[,\s]\d{3})*(?:\.\d{2})|\d+\.\d{2})', s):
        try: out.append(float(m.group(1).replace(',','').replace(' ','')))
        except ValueError: pass
    return out

def find_total(txt):
    """Return (amount, confidence, why).

    Receipts OCR two ways:
      (a) 'TOTAL   1,481.45'          - label and amount on the same line
      (b) 'TOTAL' ... then a separate block of numbers - Windows OCR often reads
          the label column and the number column as SEPARATE lines. So if the
          label line has no number, we line up the label block with the number
          block that follows and take the matching position.
    """
    lines = [l.strip() for l in txt.split('\n') if l.strip()]
    def is_sub(n):  return re.search(r'\b(sub\s*total|subtotal)\b', n)
    def is_tot(n):  return re.search(r'\b(total|importe|a\s*pagar)\b', n) and not is_sub(n)

    # (a) same-line
    same = []
    for i, l in enumerate(lines):
        n = _norm(l)
        if is_tot(n):
            vals = money_list(l)
            if vals: same.append(max(vals))
    if same:
        return max(same), 'HIGH', 'labelled "total"'

    # (b) label column then number column
    labels = [i for i,l in enumerate(lines) if is_tot(_norm(l)) and not money_list(l)]
    for li in labels:
        # how many label-ish lines lead up to this one (its position in the block)
        start = li
        while start-1 >= 0 and not money_list(lines[start-1]) and len(_norm(lines[start-1]))>2:
            start -= 1
        pos = li - start
        nums = []
        j = li + 1
        while j < len(lines) and len(nums) <= pos + 2:
            v = money_list(lines[j])
            if v: nums.append(max(v))
            elif nums: break
            j += 1
        if len(nums) > pos:
            return nums[pos], 'HIGH', 'label column matched to number column'

    # payment line
    for l in lines:
        if re.search(r'tarjeta|efectivo|debito|credito|cash|card', _norm(l)):
            vals = money_list(l)
            if vals: return max(vals), 'MEDIUM', 'payment line (no "total" found)'
    allv = money_list(txt)
    if allv:
        return max(allv), 'LOW', 'largest number on the receipt - CHECK THIS'
    return None, 'NONE', 'no amounts found'

def scan_one(path, show_text=False):
    name = os.path.basename(path)
    print('\n' + '='*64)
    print(f'  {name}')
    print('='*64)
    try:
        png = to_png(path)
    except Exception as e:
        print(f'  could not open image: {e}'); return None
    txt, err = ocr(png)
    if err:
        print(f'  OCR failed: {err}'); return None
    if not txt.strip():
        print('  no text found - try a straighter, brighter photo.'); return None
    amt, conf, why = find_total(txt)
    d   = find_date(txt)
    ven = find_vendor(txt)
    print(f'  vendor : {ven or "?"}')
    print(f'  date   : {d.strftime("%d-%b-%Y") if d else "?"}')
    print(f'  TOTAL  : {("$%s" % format(amt, ",.2f")) if amt else "?"}   [{conf}]  {why}')
    print(f'  (read with the {LAST_ENGINE} OCR engine)')
    if conf != 'HIGH':
        print('  >> not certain - check the photo before entering this.')
    if show_text:
        print('  ---- text ----')
        for l in txt.split('\n'):
            if l.strip(): print('   ', l)
    if amt and d:
        print(f'\n  paste into today.txt:')
        print(f'    NO-BILL: {ven or "receipt"} {amt:.0f}')
    return dict(file=name, vendor=ven, date=d, total=amt, conf=conf)

def main():
    args = [a for a in sys.argv[1:] if not a.startswith('--')]
    show = '--text' in sys.argv
    target = args[0] if args else DROP
    if os.path.isdir(target):
        files = [f for f in sorted(glob.glob(os.path.join(target,'*')))
                 if f.lower().endswith(IMG_EXT)]
        if not files:
            print(f'\n  No photos in {target}')
            print('  Drop receipt photos there and run this again.\n'); return
    elif os.path.isfile(target):
        files = [target]
    else:
        print(f'  not found: {target}'); return
    res = [scan_one(f, show) for f in files]
    res = [r for r in res if r]
    if len(res) > 1:
        print('\n' + '='*64); print('  SUMMARY'); print('='*64)
        tot = 0
        for r in res:
            print(f"  {r['file'][:26]:<26} {r['vendor'] or '?':<20} "
                  f"{('$%.2f' % r['total']) if r['total'] else '?':>10}  [{r['conf']}]")
            tot += r['total'] or 0
        print(f"  {'TOTAL':<47} ${tot:,.2f}")
        low = [r for r in res if r['conf'] != 'HIGH']
        if low: print(f"\n  {len(low)} receipt(s) not read confidently - check those by eye.")
    print()

if __name__ == '__main__':
    main()
