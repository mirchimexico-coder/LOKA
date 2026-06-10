"""
LOKA Restaurant — Backup Utility
---------------------------------
Call backup_excel() before ANY update to the Excel file.
Keeps up to MAX_BACKUPS (30) timestamped copies in C:\Restaurant\Backup\
Oldest backups are auto-deleted when the limit is exceeded.

Usage in any script:
    import sys
    sys.path.insert(0, r'C:\Restaurant')
    from backup_util import backup_excel
    backup_excel()   # always call this BEFORE loading/saving the workbook
"""

import os
import shutil
from datetime import datetime

EXCEL_FILE  = r'C:\LOKA\LOKA_Restaurant_Manager.xlsx'
BACKUP_DIR  = r'C:\LOKA\Backup'
MAX_BACKUPS = 30

def backup_excel(label: str = '') -> str:
    """
    Creates a timestamped backup of the Excel file.
    Deletes the oldest backup(s) if count exceeds MAX_BACKUPS.
    Returns the full path of the backup file created.
    """
    # Ensure backup folder exists
    os.makedirs(BACKUP_DIR, exist_ok=True)

    # Build filename:  LOKA_2026-05-26_14-32-05_label.xlsx
    ts    = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    slug  = f'_{label}' if label else ''
    fname = f'LOKA_{ts}{slug}.xlsx'
    dest  = os.path.join(BACKUP_DIR, fname)

    shutil.copy2(EXCEL_FILE, dest)
    print(f'[BACKUP] Created: {dest}')

    # Enforce MAX_BACKUPS — delete oldest if over limit
    backups = sorted(
        [f for f in os.listdir(BACKUP_DIR) if f.endswith('.xlsx')],
        key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f))
    )
    while len(backups) > MAX_BACKUPS:
        oldest = os.path.join(BACKUP_DIR, backups.pop(0))
        os.remove(oldest)
        print(f'[BACKUP] Deleted oldest: {oldest}')

    print(f'[BACKUP] Total backups: {len(backups)}/{MAX_BACKUPS}')
    return dest


if __name__ == '__main__':
    # Run directly to take a manual backup
    import sys
    sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)
    label = sys.argv[1] if len(sys.argv) > 1 else 'manual'
    path  = backup_excel(label)
    print(f'\n✅ Backup complete: {path}')
