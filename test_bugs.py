"""Test script - verify 3 bugs"""
import sys, os
os.environ['PYTHONIOENCODING'] = 'utf-8'
sys.path.insert(0, r'D:\codewhale')
from translator import parse_rule, apply_rules

print('=== BUG 1: SKU Change Rule ===')
rules1 = parse_rule('^[Size:L] , -1')
print(f'Parsed: {rules1}')
if rules1:
    sku, cust = apply_rules(rules1, 'CAPS251860', 'Size:L')
    print(f'SKU={sku!r} (expect: CAPS251860-1) {"PASS" if sku=="CAPS251860-1" else "FAIL"}')
else:
    print('PARSE FAILED')

print()
print('=== BUG 2: Delete Rule ===')
rules2 = parse_rule('![None:XYZ-NCT Larkspur]')
print(f'Parsed: {rules2}')
if rules2:
    # apply_rules returns (sku, customization_string), not (sku, lines_list)
    sku, cust_str = apply_rules(rules2, 'SKU123', 'None:XYZ-NCT Larkspur')
    print(f'Customization after delete: {cust_str!r} (expect: "") {"PASS" if cust_str=="" else "FAIL"}')
    # Test with spaces
    sku, cust_str = apply_rules(rules2, 'SKU123', '  None:XYZ-NCT Larkspur  ')
    print(f'With spaces: {cust_str!r} (expect: "") {"PASS" if cust_str=="" else "FAIL"}')
else:
    print('PARSE FAILED')

print()
print('=== BUG 3: Parallel Translate Rule ===')
# User format (no leading =)
rules3 = parse_rule('[Lavender|Purple|Light Pink]=[LavenderPurple|Purple|LightPink]')
print(f'User format parsed: {rules3}')
# Correct format
rules3a = parse_rule('=[Lavender|Purple|Light Pink]=[LavenderPurple|Purple|LightPink]')
print(f'Correct format parsed: {rules3a}')
if rules3a:
    sku, cust = apply_rules(rules3a, 'SKU123', 'Color:Light Pink')
    ok = cust == 'Color:LightPink'
    print(f'Translate result: {cust!r} (expect: Color:LightPink) {"PASS" if ok else "FAIL"}')
else:
    print('PARSE FAILED')

print()
print('=== BUG 3 addon: value-only match test ===')
from translator import KeyValueCondition
c = KeyValueCondition(key='Light Pink', value=None)
# Should match value part of "Color:Light Pink"
print(f'KeyValueCondition("Light Pink", None).matches_line("Color:Light Pink"): {c.matches_line("Color:Light Pink")} (expect: True)')
print(f'KeyValueCondition("Light Pink", None).matches_line("Light Pink:xxx"): {c.matches_line("Light Pink:xxx")} (expect: True)')
