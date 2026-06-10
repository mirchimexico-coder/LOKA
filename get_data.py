import sys, os
sys.path.insert(0, r'C:\Restaurant')
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)

import openpyxl
from datetime import datetime, date

wb = openpyxl.load_workbook(r'C:\Restaurant\LOKA_Restaurant_Manager.xlsx')
ws_dl = wb['📅 Daily Log']

print("=== Daily Log all rows ===")
for r in range(8, 30):
    b = ws_dl[f'B{r}'].value
    if b is None: break
    c  = ws_dl[f'C{r}'].value
    d  = ws_dl[f'D{r}'].value or 0
    e  = ws_dl[f'E{r}'].value or 0
    h  = ws_dl[f'H{r}'].value or 0
    print(f"  R{r}: {b} {c} | Card={d} Cash={e} Rev={d+e} Exp={h} Net={d+e-h:.2f}")
