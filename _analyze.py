# -*- coding: utf-8 -*-
import pandas as pd
import zipfile
import os

# === 原始订单 ===
print('=== pet6.25.xls (raw) ===')
df = pd.read_excel(r'C:\Users\Administrator\Desktop\codewhale\pet6.25.xls', dtype=str)
print(f'Rows: {len(df)}, Cols: {len(df.columns)}')

# K col = index 10
print(f'Col 10 name: {df.columns[10]!r}')
print(f'Col 10 values (non-null):')
for i in range(len(df)):
    v = df.iloc[i, 10]
    if pd.notna(v) and v.strip():
        print(f'  Row {i}: K={v!r}')

# AC/AS positions in original
for i in range(len(df.columns)):
    print(f'  Col {i:2d} ({chr(65+i) if i<26 else "A"+chr(65+i-26)}): {df.columns[i]!r}')
    if i >= 50:
        break

# === Converted file XML check ===
print()
print('=== pet6.25-副本.xlsx XML check ===')
src = r'C:\Users\Administrator\Desktop\codewhale\pet6.25-副本.xlsx'
try:
    with zipfile.ZipFile(src, 'r') as zf:
        names = zf.namelist()
        print(f'Files in zip: {len(names)}')
        if 'xl/worksheets/sheet1.xml' in names:
            xml = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
            # Check for duplicate xmlns
            print(f'xmlns count: {xml.count("xmlns=")}')
            print(f'First 500 chars: {xml[:500]}')
            # Find AC/AS cells
            import re
            ac_cells = re.findall(r'AC\d+', xml)
            as_cells = re.findall(r'AS\d+', xml)
            print(f'AC cells: {ac_cells[:10]}')
            print(f'AS cells: {as_cells[:10]}')
except Exception as e:
    print(f'Error: {e}')
