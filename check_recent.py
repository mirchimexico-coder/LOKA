import sys, openpyxl
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf8', buffering=1)

wb = openpyxl.load_workbook(r'C:\LOKA\LOKA_Restaurant_Manager.xlsx')
ws = wb[wb.sheetnames[3]]

print("=== Last expense rows (last 12) ===")
last = 7
for r in range(8, 900):
    if ws.cell(r,2).value is None: break
    last = r
for r in range(last-11, last+1):
    b = ws.cell(r,2).value
    desc = ws.cell(r,3).value
    amt  = ws.cell(r,6).value
    print(f"R{r}: {str(b)[:11]:<11} {str(desc)[:40]:<40} ${amt}")

print(f"\nLast row = {last}, next free = {last+1}")

# Check June 11 / June 10 entries
print("\n=== Any June 10/11 entries? ===")
for r in range(8, last+1):
    b = str(ws.cell(r,2).value or '')
    if '06-10' in b or '06-11' in b or '10-Jun' in b or '11-Jun' in b:
        print(f"R{r}: {b} | {ws.cell(r,3).value} | ${ws.cell(r,6).value}")
