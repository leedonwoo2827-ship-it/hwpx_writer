# -*- coding: utf-8 -*-
"""
마크다운 → HWPX JSON 파서

마크다운 텍스트를 hwpx_generator.py가 소비하는 JSON 구조로 변환합니다.

헤딩 매핑 (# 수 = 레벨):
  #      → 제목체 (섹션 구분, 첫 번째는 문서 제목)
  ##     → level 2 (기호 없음, 일반 본문)
  ###    → level 3 (□)
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
            current_section["items"].append(_item(6, text))
            i += 1
            continue

        # ##### → level 5 (―)
        if stripped.startswith("##### "):
            text = stripped[6:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(5, text))
            i += 1
            continue

        # #### → level 4 (○)
        if stripped.startswith("#### "):
            text = stripped[5:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(4, text))
            i += 1
            continue

        # ### → level 3 (□)
        if stripped.startswith("### "):
            text = stripped[4:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(3, text))
            i += 1
            continue

        # ## → level 2 (기호 없음, 일반 본문)
        if stripped.startswith("## "):
            text = stripped[3:].strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(2, text))
            i += 1
            continue

        # # → 제목체 (첫 번째는 문서 제목, 이후는 섹션 제목)
        if stripped.startswith("# ") and not stripped.startswith("## "):
            h1_text = _strip_bold(stripped[2:].strip())
            if not metadata["title"]:
                metadata["title"] = h1_text
                metadata["include_title"] = True
            else:
                current_section = _new_section(h1_text)
                content.append(current_section)
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
            if indent_len == 0:
                level = 3   # □
            elif indent_len <= 4:
                level = 4   # ○
            else:
                level = 5   # ―

            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append(_item(level, text))
            i += 1
            continue

        # ----------------------------------------------------------------
        # 수평선 (--- / *** / ___)
        # ----------------------------------------------------------------
        if re.match(r"^[-*_]{3,}\s*$", stripped):
            i += 1
            continue

        # ----------------------------------------------------------------
        # 인라인 이미지 ![alt](path)
        # ----------------------------------------------------------------
        img_match = re.match(r'^!\[([^\]]*)\]\(([^)]+)\)', stripped)
        if img_match:
            img_path = img_match.group(2).strip()
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            current_section["items"].append({
                "type": "image",
                "path": img_path,
            })
            i += 1
            continue

        # ----------------------------------------------------------------
        # 빈 줄
        # ----------------------------------------------------------------
        if not stripped:
            i += 1
            continue

        # ----------------------------------------------------------------
        # 일반 단락 텍스트 → level 2 (기호 없음, 일반 본문)
        # ----------------------------------------------------------------
        if stripped and not stripped.startswith("#"):
            if current_section is None:
                current_section = _new_section("")
                content.append(current_section)
            clean = _strip_bold(stripped)
            color = _detect_marker_color(clean)
            current_section["items"].append({
                "level": 2,
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

    raw_headers = [h.strip() for h in header_line.split("|")]
    headers = [h for h in raw_headers if h]

    if not headers:
        return None, 0

    if start + 1 >= len(lines) or not _is_separator_line(lines[start + 1]):
        return None, 0

    rows = []
    i = start + 2
    while i < len(lines):
        row_line = lines[i].strip()
        if not row_line.startswith("|"):
            break
        raw_cells = [c.strip() for c in row_line.split("|")]
        cells_raw = [c for c in raw_cells if c != ""]

        row = []
        for cell_text in cells_raw:
            color_match = re.match(r"^\{\{(red|green|blue|yellow|black):(.+)\}\}$", cell_text.strip())
            if color_match:
                row.append({"text": color_match.group(2), "color": color_match.group(1)})
            else:
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
