"""
LOKA Restaurant — Standard Update Script Template
--------------------------------------------------
Copy this template for every new update operation.
The backup_excel() call at the top is MANDATORY.

Replace the content inside # ── YOUR CHANGES HERE ── with your actual updates.
"""

import sys, os
sys.path.insert(0, r'C:\Restaurant')
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)

# ── STEP 1: Backup FIRST (always) ───────────────────────────────────────────
from backup_util import backup_excel
backup_excel('describe_what_this_update_does')

# ── STEP 2: Load workbook ────────────────────────────────────────────────────
import openpyxl
wb = openpyxl.load_workbook(r'C:\Restaurant\LOKA_Restaurant_Manager.xlsx')

# ── YOUR CHANGES HERE ────────────────────────────────────────────────────────
# ws = wb['📅 Daily Log']
# ws['F8'] = 1234
# etc.

# ── STEP 3: Save ─────────────────────────────────────────────────────────────
wb.save(r'C:\Restaurant\LOKA_Restaurant_Manager.xlsx')
print('✅ Changes saved successfully')
