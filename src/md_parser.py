# -*- coding: utf-8 -*-
"""
마크다운 → HWPX JSON 파서

마크다운 텍스트를 hwpx_generator.py가 소비하는 JSON 구조로 변환합니다.

헤딩 매핑 (# 수 = 레벨):
  #      → 제목체 (섹션 구분, 첫 번째는 문서 제목)
  ##     → 중간 제목 (subtitle level 1) — section_subtitle보다 한 단계 큰 제목체
  ###    → 소제목 (subtitle level 2) — 제목체 스타일 적용
  ####   → level 4 (○)
  #####  → level 5 (―)
  ###### → level 6 (※)

본문 기호 계층: (없음) → □ → ○ → ― → ※

색상 마커 (인라인):
  {{red:텍스트}}, {{green:텍스트}}, {{blue:텍스트}},
  {{yellow:텍스트}}, {{black:텍스트}}

볼드 마커:
  **텍스트** → {{bold:텍스트}}
"""

import re
from typing import Optional


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

def parse_markdown_to_json(
    md_content: str,
    title: str = "",
    date: str = "",
    organization: str = ""
) -> dict:
    """마크다운 텍스트를 HWPX 생성용 JSON으로 변환합니다.

    Args:
        md_content: 마크다운 텍스트
        title: 문서 제목 (생략 시 첫 번째 H1 사용)
        date: 날짜 문자열
        organization: 기관명

    Returns:
        {"metadata": {...}, "content": [...]} 형식의 dict
    """
    lines = md_content.split("\n")

    metadata: dict = {
        "title": title,
        "date": date,
        "organization": organization,
        "include_title": bool(title),
        "include_section_titles": True,
    }

    content: list = []
    current_section: Optional[dict] = None

    i = 0
    while i < len(lines):
        line = lines[i]
        stripped = line.rstrip()

        # ----------------------------------------------------------------
        # 헤딩 파싱 (가장 긴 prefix부터 확인)
        # ----------------------------------------------------------------

        # ###### → level 6 (※)
        if stripped.startswith("###### "):
            text = stripped[7:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            text = _strip_color_markers(text)
            current_section["items"].append(_item(6, text))
            i += 1
            continue

        # ##### → level 5 (―)
        if stripped.startswith("##### "):
            text = stripped[6:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            text = _strip_color_markers(text)
            current_section["items"].append(_item(5, text))
            i += 1
            continue

        # #### → level 4 (○)
        if stripped.startswith("#### "):
            text = stripped[5:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            text = _strip_color_markers(text)
            current_section["items"].append(_item(4, text))
            i += 1
            continue

        # ### → 소제목 (subtitle level 2) — 제목체 스타일 적용
        if stripped.startswith("### "):
            text = stripped[4:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            text = _strip_color_markers(_strip_bold(text))
            current_section["items"].append({
                "type": "subtitle",
                "text": text,
                "subtitle_level": 2,
            })
            i += 1
            continue

        # ## → 중간 제목 (subtitle level 1) — 제목체 스타일 (###보다 한 단계 큼)
        if stripped.startswith("## "):
            text = stripped[3:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            text = _strip_color_markers(_strip_bold(text))
            current_section["items"].append({
                "type": "subtitle",
                "text": text,
                "subtitle_level": 1,
            })
            i += 1
            continue

        # # → 제목체 (첫 번째는 문서 제목, 이후는 섹션 제목)
        if stripped.startswith("# ") and not stripped.startswith("## "):
            h1_text = _strip_color_markers(_strip_bold(stripped[2:].strip()))
            if not metadata["title"]:
                metadata["title"] = h1_text
                metadata["include_title"] = True
            else:
                current_section = _new_section(h1_text)
                content.append(current_section)
            i += 1
            continue

        # ----------------------------------------------------------------
        # 인라인 이미지 ![alt](path) — 목록/텍스트보다 먼저 감지
        # ----------------------------------------------------------------
        img_stripped = stripped.lstrip('-*+ \t')  # 목록 접두사 제거 후에도 이미지 감지
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', stripped)
        if not img_match:
            img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', img_stripped)
        if img_match:
            img_alt = img_match.group(1).strip()
            img_path = img_match.group(2).strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            img_item = {
                "type": "image",
                "path": img_path,
            }
            if img_alt:
                img_item["caption"] = img_alt
            current_section["items"].append(img_item)
            i += 1
            continue

        # ----------------------------------------------------------------
        # 테이블 감지: | 로 시작하고 다음 줄이 구분선
        # ----------------------------------------------------------------
        if stripped.startswith("|"):
            next_line = lines[i + 1].strip() if i + 1 < len(lines) else ""
            if _is_separator_line(next_line):
                table, consumed = _parse_table(lines, i)
                if table is not None:
                    if current_section is None:
                        current_section = _new_section("")
                        content.append(current_section)
                    # 바로 이전 아이템이 짧은 텍스트(level 2)면 표 제목(캡션)으로 승격
                    items = current_section["items"]
                    if (items and "type" not in items[-1]
                            and items[-1].get("level") == 2
                            and len(items[-1].get("text", "")) < 60):
                        caption_item = items.pop()
                        table["title"] = caption_item["text"]
                    current_section["items"].append(table)
                    i += consumed
                    continue

        # ----------------------------------------------------------------
        # 목록 아이템 (- / * / + / 숫자.)
        # 들여쓰기 깊이에 따라: 0 → level3, 1~4 → level4, 5+ → level5
        # ----------------------------------------------------------------
        list_match = re.match(r"^(\s*)([-*+]|\d+\.)\s+(.+)", stripped)
        if list_match:
            indent_len = len(line) - len(line.lstrip())
            text = list_match.group(3).strip()
            # 색상 마커 감지 후 제거
            color = _detect_marker_color(text)
            text = _strip_color_markers(text)

            if indent_len == 0:
                level = 3   # □
            elif indent_len <= 4:
                level = 4   # ○
            else:
                level = 5   # ―

            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(level, text, color))
            i += 1
            continue

        # ----------------------------------------------------------------
        # 수평선 (--- / *** / ___)
        # ----------------------------------------------------------------
        if re.match(r"^[-*_]{3,}\s*$", stripped):
            i += 1
            continue

        # ----------------------------------------------------------------
        # 빈 줄
        # ----------------------------------------------------------------
        if not stripped:
            i += 1
            continue

        # ----------------------------------------------------------------
        # 일반 단락 텍스트 → 맨 앞 기호에 따라 level 자동 결정
        # ----------------------------------------------------------------
        if stripped and not stripped.startswith("#"):
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            clean = _strip_bold(stripped)
            color = _detect_marker_color(clean)
            # 색상 마커 제거
            clean = _strip_color_markers(clean)

            # 맨 앞 기호로 레벨 결정
            level = _detect_level_by_symbol(clean)

            current_section["items"].append({
                "level": level,
                "text": clean,
                "color": color,
            })
            i += 1
            continue

        i += 1

    return {"metadata": metadata, "content": content}


# ---------------------------------------------------------------------------
# 내부 헬퍼
# ---------------------------------------------------------------------------

def _new_section(title: str) -> dict:
    return {"type": "section", "title": title, "items": []}


def _strip_bold(text: str) -> str:
    """**bold** 마커를 {{bold:...}} 인라인 마커로 변환합니다."""
    return re.sub(r'\*\*(.+?)\*\*', r'{{bold:\1}}', text)


def _item(level: int, text: str, color: str = "") -> dict:
    text = _strip_bold(text)
    c = color or _detect_marker_color(text)
    return {"level": level, "text": text, "color": c}


def _detect_marker_color(text: str) -> str:
    """텍스트에서 {{color:...}} 마커 색상을 감지합니다."""
    for color in ("red", "green", "blue", "yellow"):
        if f"{{{{{color}:" in text:
            return color
    return "black"


def _strip_color_markers(text: str) -> str:
    """텍스트에서 {{color:...}} 마커를 제거하고 순수 텍스트만 반환합니다.

    예: "일반 {{green:초록색}} 텍스트" → "일반 초록색 텍스트"
    """
    # {{color:내용}} 패턴을 찾아서 내용만 남김
    return re.sub(r'\{\{(?:red|green|blue|yellow|black):([^}]+)\}\}', r'\1', text)


def _detect_level_by_symbol(text: str) -> int:
    """텍스트 맨 앞 기호로 레벨을 자동 감지합니다.

    기호 매핑:
        □ → level 3
        ○ → level 4
        ― → level 5
        ※ → level 6
        기호 없음 → level 2 (일반 본문)
    """
    text = text.lstrip()
    if not text:
        return 2

    first_char = text[0]
    if first_char == '□':
        return 3
    elif first_char == '○':
        return 4
    elif first_char == '―':
        return 5
    elif first_char == '※':
        return 6
    else:
        return 2  # 기호 없음 = 일반 본문


def _is_separator_line(line: str) -> bool:
    """마크다운 테이블 구분선인지 확인합니다."""
    stripped = line.strip()
    if not stripped:
        return False
    cleaned = stripped.replace("|", "").replace(":", "").replace("-", "").replace(" ", "")
    return len(cleaned) == 0 and "-" in stripped


def _parse_table(lines: list, start: int) -> tuple:
    """마크다운 테이블을 파싱해 JSON table 객체와 소비된 줄 수를 반환합니다."""
    header_line = lines[start].strip()
    if not header_line.startswith("|"):
        return None, 0

    # 선행/후행 | 를 제거하고 내부 셀만 분리 (빈 헤더 허용)
    inner_header = header_line.strip().strip("|")
    headers = [_strip_color_markers(_strip_bold(h.strip())) for h in inner_header.split("|")]

    if not headers or len(headers) < 1:
        return None, 0

    if start + 1 >= len(lines) or not _is_separator_line(lines[start + 1]):
        return None, 0

    col_count = len(headers)

    rows = []
    i = start + 2
    while i < len(lines):
        row_line = lines[i].strip()
        if not row_line.startswith("|"):
            break
        # 선행/후행 | 를 제거하고 내부 셀만 분리 (빈 셀 보존)
        inner = row_line.strip("|")
        cells_raw = [c.strip() for c in inner.split("|")]
        # 열 수 맞추기: 부족하면 "-" 채움, 초과하면 자르기
        while len(cells_raw) < col_count:
            cells_raw.append("-")
        cells_raw = cells_raw[:col_count]

        row = []
        for cell_text in cells_raw:
            if not cell_text:
                cell_text = "-"
            # **볼드** 마커를 {{bold:...}}로 변환
            cell_text = _strip_bold(cell_text)
            # 색상 마커 제거
            cell_text = _strip_color_markers(cell_text)
            # 항상 문자열로 통일
            row.append(cell_text)
        rows.append(row)
        i += 1

    consumed = i - start
    table = {
        "type": "table",
        "headers": headers,
        "rows": rows,
    }
    return table, consumed
