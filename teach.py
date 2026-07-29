# -*- coding: utf-8 -*-
"""
LOKA - teach the tool a category.

  py teach.py            -> scan today.txt, ask about anything it couldn't recognise
  py teach.py --add      -> just add a word/category by hand
  py teach.py --list     -> show everything you've taught it

Whatever you teach is remembered in categories.json and used from then on.
"""
import sys, os, io, json
sys.path.insert(0, r'C:\LOKA')
import eod

def pick_category(word):
    print(f"\n   Which category does '{word}' belong to?\n")
    for i, c in enumerate(eod.CATEGORIES, 1):
        print(f"     {i:>2}. {c}")
    print("      0. skip this one")
    while True:
        raw = input("\n   Number: ").strip()
        if raw == '0': return None
        if raw.isdigit() and 1 <= int(raw) <= len(eod.CATEGORIES):
            return eod.CATEGORIES[int(raw)-1]
        print("   Please type one of the numbers above.")

def teach_word(word, cat=None):
    cat = cat or pick_category(word)
    if not cat:
        print("   skipped."); return False
    eod.save_rule(word, cat)
    print(f"   SAVED:  anything containing '{word.lower()}'  ->  {cat}")
    return True

def scan_today():
    path = r'C:\LOKA\today.txt'
    if not os.path.exists(path):
        print("\n  There is no today.txt yet - use menu option 1 first,")
        print("  or run:  py teach.py --add   to add a word by hand.\n")
        return
    d = eod.parse(io.open(path, encoding='utf-8').read())
    rows, _ = eod.build(d)
    unknown = [r for r in rows if r['cat'] == eod.UNKNOWN]
    if not unknown:
        print("\n  Good news - everything in today.txt was recognised.")
        print("  (Use  py teach.py --add  if you still want to correct something.)\n")
        return
    print(f"\n  {len(unknown)} item(s) were not recognised. Let's fix them for good.")
    n = 0
    for r in unknown:
        print("\n" + "-"*54)
        print(f"   ITEM:  {r['desc']}   ${r['amount']:,.2f}")
        word = input(f"   Which word should I remember? [{r['desc']}]: ").strip() or r['desc']
        if teach_word(word): n += 1
    print(f"\n  Done - learned {n} new rule(s).")
    print("  Run the PREVIEW again (menu 2) and they will be categorised correctly.\n")

def main():
    args = sys.argv[1:]
    if '--list' in args:
        r = eod.load_rules()
        if not r:
            print("\n  You haven't taught it anything yet.\n"); return
        print(f"\n  You have taught it {len(r)} rule(s):\n")
        for k in sorted(r): print(f"     {k:<28} -> {r[k]}")
        print()
        return
    if '--add' in args:
        w = input("\n   Word or phrase to remember (e.g. 'utensils'): ").strip()
        if not w: print("   nothing entered."); return
        print(f"   (currently '{w}' would be filed as: {eod.guess_cat(w)})")
        teach_word(w); print()
        return
    scan_today()

if __name__ == '__main__':
    main()
