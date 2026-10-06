# -*- coding: utf-8 -*-
"""Copy the important LOKA files to Google Drive (H:\\My Drive\\Loka Tracker).

  py gdrive_backup.py

- Overwrites the main copy in the Drive folder (always the latest).
- Also keeps dated copies in  Loka Tracker\\History  (last 30 of each file).
- Google Drive for desktop syncs the folder to the cloud by itself.
Console is cp1252: plain ASCII output only.
"""
import os, sys, glob, shutil, filecmp
from datetime import datetime

SRC_DIR = r'C:\LOKA'
DEST    = r'H:\My Drive\Loka Tracker'
HIST    = os.path.join(DEST, 'History')
FILES   = ['LOKA_Restaurant_Manager.xlsx', 'dashboard.html']
KEEP    = 60          # History = ONE copy per day (latest of that day), last 60 days
STAMP   = '%Y-%m-%d'

def main():
    print("\n== BACKUP TO GOOGLE DRIVE ==")
    if not os.path.isdir(DEST):
        print(f"   Cannot reach {DEST}")
        print("   Is Google Drive for desktop running and signed in (drive H:)? Nothing copied.")
        return 1
    if os.path.exists(os.path.join(SRC_DIR, '~$LOKA_Restaurant_Manager.xlsx')):
        print("   NOTE: the workbook looks OPEN in Excel. Unsaved changes will NOT be in the backup.")
        print("         Save/close Excel first if you changed anything by hand.")
    os.makedirs(HIST, exist_ok=True)
    stamp = datetime.now().strftime(STAMP)
    ok = True
    for f in FILES:
        src = os.path.join(SRC_DIR, f)
        if not os.path.exists(src):
            print(f"   MISSING {src}"); ok = False; continue
        main_copy = os.path.join(DEST, f)
        stem, ext = os.path.splitext(f)
        dated = os.path.join(HIST, f'{stem}_{stamp}{ext}')
        try:
            shutil.copy2(src, main_copy)
            shutil.copy2(src, dated)
        except PermissionError:
            print(f"   {f}: Drive copy is locked (open on this PC?). Close it and retry."); ok = False; continue
        same = filecmp.cmp(src, main_copy, shallow=False) and filecmp.cmp(src, dated, shallow=False)
        kb = os.path.getsize(src) / 1024
        print(f"   {'OK  ' if same else 'FAIL'} {f:<30} {kb:>8,.0f} KB  -> Loka Tracker + History")
        ok = ok and same
        old = sorted(glob.glob(os.path.join(HIST, f'{stem}_*{ext}')))
        for x in old[:-KEEP]:
            try: os.remove(x)
            except OSError: pass
    n = len(glob.glob(os.path.join(HIST, '*')))
    print(f"\n   History folder: {n} file(s) (one copy per day, last {KEEP} days)")
    print("   Done. Google Drive will sync it to the cloud." if ok else "   Finished WITH PROBLEMS - see above.")
    return 0 if ok else 1

def auto():
    """Called at the end of every loka.refresh_all() (EOD, fixes, counts...).
    One line of output; NEVER raises, so a Drive problem can't break a write."""
    try:
        if not os.path.isdir(DEST):
            print("  google drive: H: not reachable (Drive for desktop off?) - skipped, use menu 22 later")
            return False
        os.makedirs(HIST, exist_ok=True)
        stamp = datetime.now().strftime(STAMP)
        for f in FILES:
            src = os.path.join(SRC_DIR, f); stem, ext = os.path.splitext(f)
            shutil.copy2(src, os.path.join(DEST, f))
            shutil.copy2(src, os.path.join(HIST, f'{stem}_{stamp}{ext}'))
            if not filecmp.cmp(src, os.path.join(DEST, f), shallow=False):
                print(f"  google drive: {f} copy does not match - run menu 22"); return False
            for x in sorted(glob.glob(os.path.join(HIST, f'{stem}_*{ext}')))[:-KEEP]:
                try: os.remove(x)
                except OSError: pass
        print("  google drive: workbook + dashboard backed up to Loka Tracker")
        return True
    except Exception as e:
        print(f"  google drive: backup skipped ({type(e).__name__}: {e}) - use menu 22 later")
        return False

if __name__ == '__main__':
    sys.exit(main())
