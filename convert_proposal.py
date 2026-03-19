# -*- coding: utf-8 -*-
"""proposal-body.md를 HWPX로 변환"""
import sys
import json
from pathlib import Path

sys.path.insert(0, 'src')
from md_parser import parse_markdown_to_json
from hwpx_generator import HWPXGenerator

# 파일 경로
md_file = Path(r'C:\Users\ubion\Documents\proposals\260319-n2\proposal-body.md')
output_dir = Path(r'C:\Users\ubion\Documents\proposals\260319-n2\output')
json_file = output_dir / 'proposal-body.json'
hwpx_file = output_dir / 'proposal-body.hwpx'

print('Reading MD file:', md_file)

# MD 파일 읽기
raw = md_file.read_bytes()
if raw.startswith(b'\xef\xbb\xbf'):
    md_content = raw.decode('utf-8-sig')
else:
    try:
        md_content = raw.decode('utf-8')
    except:
        md_content = raw.decode('cp949')

print(f'MD file loaded: {len(md_content):,} chars')

# MD → JSON 변환
print('Converting to JSON...')
data = parse_markdown_to_json(md_content, title='')

# JSON 저장
with open(json_file, 'w', encoding='utf-8') as f:
    json.dump(data, f, ensure_ascii=False, indent=2)

print(f'JSON saved: {json_file}')

# HWPX 생성
print('Generating HWPX...')
styles_path = Path(__file__).parent / 'proposal-styles.json'
generator = HWPXGenerator(
    base_dir=str(md_file.parent),
    styles_path=str(styles_path)
)
generator.generate(data, str(hwpx_file))

if hwpx_file.exists():
    size = hwpx_file.stat().st_size
    print(f'HWPX created: {hwpx_file} ({size:,} bytes)')

    # 색상 마커 확인
    with open(json_file, 'r', encoding='utf-8') as f:
        json_str = f.read()
        green_count = json_str.count('{{green:')
        red_count = json_str.count('{{red:')
        if green_count + red_count == 0:
            print('[SUCCESS] Color markers completely removed!')
        else:
            print(f'[WARNING] Color markers remain: green={green_count}, red={red_count}')
else:
    print('[ERROR] HWPX generation failed')
