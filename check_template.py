import sys
import os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import pandas as pd
from translator import _find_column, load_template, parse_rule

if len(sys.argv) > 1:
    file_path = sys.argv[1]
elif os.path.exists("template.xlsx"):
    file_path = "template.xlsx"
else:
    print("用法: python check_template.py <模板Excel文件路径>")
    print("      python check_template.py")
    print()
    print("说明: 读取模板 Excel 文件，验证每条规则的解析结果，")
    print("      检测解析失败的规则并输出原因。")
    print()
    sys.exit(1)

print(f"检查模板文件: {file_path}")
print()

# 模拟 _load_template_with_errors 的逻辑
parsed = load_template(file_path)
df = pd.read_excel(file_path, dtype=str)
sku_col = _find_column(df, ["SKU", "sku", "商品编码"])
rule_col = _find_column(df, ["翻译模板", "规则", "rule", "template"])
if sku_col is None and len(df.columns) >= 2:
    sku_col = df.columns[1]
if rule_col is None and len(df.columns) >= 3:
    rule_col = df.columns[2]

print(f'sku_col={sku_col!r}, rule_col={rule_col!r}')
print(f'parsed keys: {sorted(parsed.keys())}')
print()

errors = []
for idx, row in df.iterrows():
    sku = str(row[sku_col]).strip() if pd.notna(row[sku_col]) else ""
    rule_str = str(row[rule_col]).strip() if pd.notna(row[rule_col]) else ""
    excel_row = idx + 2
    if sku and rule_str:
        rules = parsed.get(sku)
        status = "OK" if (rules is not None and len(rules) > 0) else f"MISS(rules={rules})"
        print(f'row{excel_row}: SKU={sku!r} status={status}')
        if rules is None or len(rules) == 0:
            # 二次检测
            try:
                test = parse_rule(rule_str)
                if not test:
                    errors.append(f"row{excel_row} SKU={sku}: parse_rule returned empty")
                    print(f'  -> ERROR: parse_rule returned []')
                else:
                    print(f'  -> but 2nd parse_rule succeeded with {len(test)} rules')
            except Exception as e:
                errors.append(f"row{excel_row} SKU={sku}: {e}")
                print(f'  -> ERROR: {e}')

print(f'\nTotal errors: {len(errors)}')
for e in errors:
    print(f'  {e}')
