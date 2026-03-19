# -*- coding: utf-8 -*-
"""최종 시스템 테스트"""
import sys
sys.path.insert(0, 'src')
from md_parser import parse_markdown_to_json
from hwpx_generator import HWPXGenerator
from pathlib import Path

print('=== Final System Test ===\n')

# Test 1: Color marker removal
print('[Test 1] Color marker removal')
test_md = '''
# {{green:Title}}
## {{red:Section}}
{{blue:Body text}}
- {{green:List item}}
| {{green:Header}} |
|---|
| {{red:Data}} |
'''
data = parse_markdown_to_json(test_md)
json_str = str(data)
if '{{green:' in json_str or '{{red:' in json_str:
    print('[FAIL] Color markers remain')
else:
    print('[PASS] All color markers removed')

# Test 2: Table style validation
print('\n[Test 2] Table style validation')
gen = HWPXGenerator(styles_path='proposal-styles.json')
parapr_xml = gen._build_table_parapr_xml(99, 'JUSTIFY')
if 'value="0"' in parapr_xml and 'horizontal="JUSTIFY"' in parapr_xml:
    print('[PASS] Table parapr correct (indent=0, align=JUSTIFY)')
else:
    print('[FAIL] Table parapr incorrect')

# Test 3: test_complete.hwpx regeneration
print('\n[Test 3] test_complete.hwpx regeneration')
md_file = Path('test_complete.md')
if md_file.exists():
    md_content = md_file.read_text(encoding='utf-8')
    data = parse_markdown_to_json(md_content)
    gen.generate(data, 'test_complete.hwpx')
    hwpx_file = Path('test_complete.hwpx')
    if hwpx_file.exists():
        size = hwpx_file.stat().st_size
        print(f'[PASS] test_complete.hwpx created ({size:,} bytes)')
    else:
        print('[FAIL] HWPX generation failed')
else:
    print('[SKIP] test_complete.md not found')

print('\n=== Test Complete ===')
