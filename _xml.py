# -*- coding: utf-8 -*-
import zipfile, io, re
P=r'C:\LOKA\LOKA_Restaurant_Manager.xlsx'
z=zipfile.ZipFile(P)
# find which sheetN.xml is the Daily Log
wbxml=z.read('xl/workbook.xml').decode('utf-8')
rels=z.read('xl/_rels/workbook.xml.rels').decode('utf-8')
out=[]
# map sheet name -> r:id -> target file
import xml.etree.ElementTree as ET
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
sheets={}
for s in ET.fromstring(wbxml).iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}sheet'):
    sheets[s.get('name')]=s.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
relmap={}
for rel in ET.fromstring(rels):
    relmap[rel.get('Id')]=rel.get('Target')
dl_file='xl/'+relmap[sheets['📅 Daily Log']]
out.append(f"Daily Log file: {dl_file}")
xml=z.read(dl_file).decode('utf-8')
# extract the <c ...>...</c> for H55, H67, H68
for cell in ('H55','H67','H68','F68','I68','J68'):
    m=re.search(r'<c r="'+cell+r'"[^>]*>.*?</c>', xml)
    if not m:
        m=re.search(r'<c r="'+cell+r'"[^>]*/>', xml)
    out.append(f"\n--- {cell} ---")
    out.append(m.group(0) if m else "NOT FOUND")
# check workbook.xml calcPr
mc=re.search(r'<calcPr[^>]*/>', wbxml)
out.append("\n--- calcPr ---")
out.append(mc.group(0) if mc else "NO calcPr")
io.open(r'C:\LOKA\_out.txt','w',encoding='utf-8',newline='\n').write("\n".join(out))
print("done")
