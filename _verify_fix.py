# -*- coding: utf-8 -*-
"""验证 AC/AS 修复效果"""
import sys, os, zipfile, re

sys.path.insert(0, r'C:\Users\Administrator\Desktop\codewhale')
os.chdir(r'C:\Users\Administrator\Desktop\codewhale')

import pandas as pd
from translator import process_orders_preserve_format, load_template, _find_column

# 1. 加载模板（如果有）
template = {}
template_path = r'C:\Users\Administrator\Desktop\codewhale\template_database.json'
try:
    import json
    with open(template_path, 'r', encoding='utf-8') as f:
        db = json.load(f)
    for sku, entries in db.items():
        from translator import parse_rule
        rules = parse_rule(entries[0]['template'])
        if rules:
            template[sku] = rules
    print(f'Loaded template: {len(template)} SKUs')
except Exception as e:
    print(f'No template loaded: {e}')

# 2. 转化
out = r'C:\Users\Administrator\Desktop\codewhale\pet6.25_test_output.xlsx'
src = r'C:\Users\Administrator\Desktop\codewhale\pet6.25.xls'

modified, skipped, errors = process_orders_preserve_format(src, out, template)
print(f'Result: modified={modified}, skipped={skipped}, errors={errors}')

# 3. 检查输出 XML
with zipfile.ZipFile(out, 'r') as zf:
    xml = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')

print(f'xmlns count: {xml.count("xmlns=")}')
# Check for duplicates
import xml.etree.ElementTree as ET
try:
    tree = ET.fromstring(xml)
    print('XML parse: OK (no duplicate attrs)')
except ET.ParseError as e:
    print(f'XML parse ERROR: {e}')

# 4. 检查 AC/AS 列内容
df_out = pd.read_excel(out, dtype=str)
print(f'Output: {len(df_out)} rows x {len(df_out.columns)} cols')

# 找 AC/AS 列
for i, col in enumerate(df_out.columns):
    if i >= 27 and i <= 30:
        vals = df_out.iloc[:5, i].tolist()
        letter = chr(65+i) if i < 26 else 'A'+chr(65+i-26)
        print(f'  Col {i} ({letter}): {str(col)[:30]!r} -> {vals}')
    if i >= 43 and i <= 46:
        vals = df_out.iloc[:5, i].tolist()
        letter = chr(65+i) if i < 26 else 'A'+chr(65+i-26)
        print(f'  Col {i} ({letter}): {str(col)[:30]!r} -> {vals}')

# 5. 检查 K 列（加急）
k_idx = 10
print(f'K col ({df_out.columns[10][:20]!r}) values: {list(df_out.iloc[:5, 10])}')
