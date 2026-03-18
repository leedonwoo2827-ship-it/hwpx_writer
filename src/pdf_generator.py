# -*- coding: utf-8 -*-
"""
PDF Generator — fpdf2 기반 직접 렌더링

HWPX JSON 데이터를 proposal-styles.json 스타일에 맞춰 PDF로 생성합니다.
fpdf2를 사용하여 TTF 한글 폰트를 직접 임베딩합니다.

폰트: Noto Sans KR (제목/고딕) + Noto Serif KR (본문/명조)
라이선스: Google OFL (저작권 자유)
"""

import json
import re
import sys
from pathlib import Path

from fpdf import FPDF


def _log(msg: str) -> None:
    print(msg, file=sys.stderr)


# ---------------------------------------------------------------------------
# 색상 헬퍼
# ---------------------------------------------------------------------------

def _hex_to_rgb(hex_color: str) -> tuple:
    """#RRGGBB → (R, G, B) 튜플."""
    h = hex_color.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))


# ---------------------------------------------------------------------------
# 인라인 마커 파싱
# ---------------------------------------------------------------------------

_MARKER_RE = re.compile(r'\{\{(bold|red|green|blue|yellow|black):(.+?)\}\}')


def _ensure_str(cell) -> str:
    """셀 데이터를 문자열로 보장합니다. (하위호환: dict → 인라인 마커 복원)"""
    if isinstance(cell, dict):
        text = cell.get("text", "")
        color = cell.get("color", "")
        if color and "{{" not in text:
            return f"{{{{{color}:{text}}}}}"
        return text
    return str(cell)


def _parse_segments(text: str, colors: dict) -> list:
    """텍스트를 [{text, bold, color}, ...] 세그먼트 리스트로 파싱합니다."""
    segments = []
    last_end = 0
    for m in _MARKER_RE.finditer(text):
        # 마커 앞 일반 텍스트
        if m.start() > last_end:
            segments.append({"text": text[last_end:m.start()], "bold": False, "color": None})
        marker_type = m.group(1)
        inner_text = m.group(2)
        if marker_type == "bold":
            segments.append({"text": inner_text, "bold": True, "color": None})
        else:
            rgb = _hex_to_rgb(colors.get(marker_type, "#000000"))
            segments.append({"text": inner_text, "bold": False, "color": rgb})
        last_end = m.end()
    # 나머지 텍스트
    if last_end < len(text):
        segments.append({"text": text[last_end:], "bold": False, "color": None})
    return segments


# ---------------------------------------------------------------------------
# PDF 생성기
# ---------------------------------------------------------------------------

_SYMBOLS = {3: "□ ", 4: "○ ", 5: "― ", 6: "※ "}


class ProposalPDF(FPDF):
    """proposal-styles.json 스타일을 적용하는 PDF 생성기."""

    def __init__(self, styles: dict, base_dir: str = ""):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.styles = styles.get("styles", {})
        self.colors = styles.get("colors", {"red": "#dc2626", "green": "#16a34a",
                                              "blue": "#2563eb", "yellow": "#eab308",
                                              "black": "#000000"})
        self.line_spacing = styles.get("lineSpacing", 160) / 100.0
        self.base_dir = base_dir

        # 여백: 상25 하25 좌20 우20 (mm)
        self.set_margins(20, 25, 20)
        self.set_auto_page_break(auto=True, margin=25)

        # 폰트 등록
        self._register_fonts()

    def _register_fonts(self):
        """Noto Sans/Serif KR 폰트를 등록합니다."""
        font_dir = Path("C:/Windows/Fonts")

        # Noto Sans KR (고딕)
        noto_sans = font_dir / "NotoSansKR-VF.ttf"
        if noto_sans.exists():
            self.add_font("NotoSansKR", "", str(noto_sans), uni=True)
            self.add_font("NotoSansKR", "B", str(noto_sans), uni=True)
            self._gothic = "NotoSansKR"
        else:
            # fallback: 맑은 고딕
            malgun = font_dir / "malgun.ttf"
            malgun_bd = font_dir / "malgunbd.ttf"
            if malgun.exists():
                self.add_font("MalgunGothic", "", str(malgun), uni=True)
                if malgun_bd.exists():
                    self.add_font("MalgunGothic", "B", str(malgun_bd), uni=True)
                else:
                    self.add_font("MalgunGothic", "B", str(malgun), uni=True)
                self._gothic = "MalgunGothic"
            else:
                self._gothic = "Helvetica"

        # Noto Serif KR (명조)
        noto_serif = font_dir / "NotoSerifKR-VF.ttf"
        if noto_serif.exists():
            self.add_font("NotoSerifKR", "", str(noto_serif), uni=True)
            self.add_font("NotoSerifKR", "B", str(noto_serif), uni=True)
            self._serif = "NotoSerifKR"
        else:
            self._serif = self._gothic  # fallback

        _log(f"[PDF] 폰트 등록: gothic={self._gothic}, serif={self._serif}")

    def _pt_to_mm(self, pt: float) -> float:
        return pt * 0.3528

    def _set_gothic(self, size: float, bold: bool = False):
        self.set_font(self._gothic, "B" if bold else "", size)

    def _set_serif(self, size: float, bold: bool = False):
        self.set_font(self._serif, "B" if bold else "", size)

    def _write_segments(self, segments: list, font_name: str, size: float):
        """인라인 마커가 적용된 세그먼트를 출력합니다."""
        for seg in segments:
            if seg["color"]:
                self.set_text_color(*seg["color"])
            else:
                self.set_text_color(0, 0, 0)

            style = "B" if seg["bold"] else ""
            self.set_font(font_name, style, size)
            self.write(self._pt_to_mm(size) * self.line_spacing, seg["text"])

        self.set_text_color(0, 0, 0)

    def render_title(self, title: str):
        """문서 제목을 렌더링합니다."""
        s = self.styles.get("title", {})
        size = s.get("size", 22)
        self._set_gothic(size, bold=True)
        self.cell(0, self._pt_to_mm(size) * 1.5, title, align="C", new_x="LMARGIN", new_y="NEXT")
        # 구분선
        y = self.get_y() + 2
        self.set_draw_color(51, 51, 51)
        self.set_line_width(0.5)
        self.line(self.l_margin, y, self.w - self.r_margin, y)
        self.set_y(y + 5)

    def render_section_title(self, title: str):
        """섹션 제목(H1)을 렌더링합니다."""
        s = self.styles.get("level1", {})
        size = s.get("size", 16)
        space_before = self._pt_to_mm(s.get("paragraphSpaceBefore", 20))
        space_after = self._pt_to_mm(s.get("paragraphSpaceAfter", 6))

        self.ln(space_before)
        # 좌측 바
        y = self.get_y()
        h = self._pt_to_mm(size) * 1.4
        self.set_fill_color(51, 51, 51)
        self.rect(self.l_margin, y, 1.2, h, "F")

        self._set_gothic(size, bold=True)
        self.set_x(self.l_margin + 3)
        self.cell(0, h, title, new_x="LMARGIN", new_y="NEXT")
        self.ln(space_after)

    def render_subtitle(self, text: str):
        """소제목(H3)을 렌더링합니다."""
        s = self.styles.get("section_subtitle", {})
        size = s.get("size", 15)
        space_before = self._pt_to_mm(s.get("paragraphSpaceBefore", 15))
        space_after = self._pt_to_mm(s.get("paragraphSpaceAfter", 6))

        self.ln(space_before)
        segments = _parse_segments(text, self.colors)
        self._write_segments(segments, self._gothic, size)
        self.ln(self._pt_to_mm(size) * self.line_spacing + space_after)

    def render_text_item(self, level: int, text: str):
        """일반 텍스트 항목을 렌더링합니다."""
        level_key = f"level{level}"
        s = self.styles.get(level_key, {})
        size = s.get("size", 15)
        left_margin = self._pt_to_mm(s.get("leftMargin", 20))
        space_before = self._pt_to_mm(s.get("paragraphSpaceBefore", 3))
        space_after = self._pt_to_mm(s.get("paragraphSpaceAfter", 2))

        self.ln(space_before)
        self.set_x(self.l_margin + left_margin * 0.5)

        symbol = _SYMBOLS.get(level, "")
        # 텍스트가 이미 기호로 시작하면 중복 방지
        if symbol and text.lstrip().startswith(symbol.strip()):
            full_text = text
        else:
            full_text = f"{symbol}{text}"
        segments = _parse_segments(full_text, self.colors)
        self._write_segments(segments, self._serif, size)
        self.ln(self._pt_to_mm(size) * self.line_spacing + space_after)

    def _calc_cell_height(self, text: str, col_w: float, size: float) -> float:
        """셀 텍스트가 차지할 높이를 계산합니다 (글자 단위 줄바꿈 시뮬레이션)."""
        line_h = self._pt_to_mm(size) * self.line_spacing
        inner_w = col_w - 2
        if inner_w <= 0:
            inner_w = col_w
        if not text or not text.strip():
            return line_h
        self._set_gothic(size)
        # 글자 단위로 줄바꿈 시뮬레이션 (한글 CJK 호환)
        lines = 1
        current_w = 0.0
        for ch in text:
            ch_w = self.get_string_width(ch)
            if current_w + ch_w > inner_w:
                lines += 1
                current_w = ch_w
            else:
                current_w += ch_w
        return line_h * lines

    def _calc_row_height(self, cells_text: list, col_w_list, size: float) -> float:
        """행의 최대 높이를 계산합니다. col_w_list는 float 또는 list."""
        min_h = self._pt_to_mm(size) * 1.8
        max_h = min_h
        for ci, text in enumerate(cells_text):
            cw = col_w_list[ci] if isinstance(col_w_list, list) else col_w_list
            h = self._calc_cell_height(text, cw, size)
            if h > max_h:
                max_h = h
        return max_h

    @staticmethod
    def _strip_markers(text: str) -> str:
        """인라인 마커를 제거하고 순수 텍스트만 반환합니다."""
        return _MARKER_RE.sub(lambda m: m.group(2), text)

    def _detect_cell_color(self, text: str):
        """셀 텍스트에서 첫 번째 색상 마커를 감지하여 RGB 튜플 반환."""
        m = _MARKER_RE.search(text)
        if m and m.group(1) != "bold":
            return _hex_to_rgb(self.colors.get(m.group(1), "#000000"))
        return None

    # ------------------------------------------------------------------
    # 표 셀 전처리: 입구에서 한 번만 처리
    # ------------------------------------------------------------------
    def _prepare_table_cells(self, raw_cells: list) -> list:
        """원본 셀 리스트 → [{"plain": str, "segments": list}, ...] 로 변환.

        이후 코드는 plain(높이 계산용)과 segments(렌더링용)만 사용하므로
        마커가 높이 계산에 섞이거나 렌더링에서 누락될 수 없습니다.
        """
        result = []
        for cell in raw_cells:
            raw = _ensure_str(cell)
            plain = self._strip_markers(raw)
            segments = _parse_segments(raw, self.colors)
            result.append({"plain": plain, "segments": segments})
        return result

    def _render_cell_segments(self, segments: list, x: float, y: float,
                              col_w: float, line_h: float, size: float,
                              align: str = "L", bold: bool = False):
        """셀 내부에 세그먼트를 렌더링합니다. 단색 셀은 multi_cell, 다색은 write."""
        # 색상이 하나뿐인 셀 (대부분): multi_cell로 자동 줄바꿈
        has_mixed = len(segments) > 1 or any(s["bold"] for s in segments)
        first_color = segments[0]["color"] if segments else None

        self.set_xy(x + 1, y + 0.5)

        if not has_mixed:
            # 단일 세그먼트: multi_cell (자동 줄바꿈)
            if first_color:
                self.set_text_color(*first_color)
            else:
                self.set_text_color(0, 0, 0)
            style = "B" if bold else ""
            self.set_font(self._gothic, style, size)
            self.multi_cell(col_w - 2, line_h, segments[0]["text"], align=align)
        else:
            # 다중 세그먼트: write로 인라인 렌더링
            for seg in segments:
                if seg["color"]:
                    self.set_text_color(*seg["color"])
                else:
                    self.set_text_color(0, 0, 0)
                style = "B" if (seg["bold"] or bold) else ""
                self.set_font(self._gothic, style, size)
                self.write(line_h, seg["text"])

        self.set_text_color(0, 0, 0)

    def render_table(self, table: dict):
        """테이블을 렌더링합니다."""
        s = self.styles.get("table", {})
        size = s.get("size", 10)
        headers = table.get("headers", [])
        rows = table.get("rows", [])
        title = table.get("title", "")
        line_h = self._pt_to_mm(size) * self.line_spacing

        # 캡션 (마커 파싱 포함)
        if title:
            self.ln(3)
            cap_segments = _parse_segments(title, self.colors)
            self._set_gothic(11, bold=True)
            self.set_x(self.l_margin)
            self._write_segments(cap_segments, self._gothic, 11)
            self.ln(self._pt_to_mm(11) * self.line_spacing + 1)

        if not headers and not rows:
            return

        # 컬럼 수 & 너비
        col_count = len(headers) if headers else (len(rows[0]) if rows else 1)
        usable_w = self.w - self.l_margin - self.r_margin
        col_w = usable_w / col_count

        # ★ 입구에서 모든 셀을 한 번만 전처리 → 이후 마커 누출 불가능
        header_cells = self._prepare_table_cells(headers)
        header_plains = [c["plain"] for c in header_cells]
        has_visible_header = any(h.strip() for h in header_plains)

        self.set_draw_color(102, 102, 102)
        if has_visible_header:
            self._set_gothic(size, bold=True)
            self.set_fill_color(240, 240, 240)
            row_h = self._calc_row_height(header_plains, col_w, size)
            y0 = self.get_y()
            x0 = self.l_margin
            for ci, cell in enumerate(header_cells):
                x = x0 + ci * col_w
                self.rect(x, y0, col_w, row_h, "FD")
                self._render_cell_segments(
                    cell["segments"], x, y0, col_w, line_h, size,
                    align="C", bold=True)
            self.set_xy(x0, y0 + row_h)

        # 데이터 행
        for row in rows:
            row_cells = self._prepare_table_cells(row)
            row_plains = [c["plain"] for c in row_cells]
            self._set_gothic(size)
            row_h = self._calc_row_height(row_plains, col_w, size)

            # 페이지 넘김
            if self.get_y() + row_h > self.h - self.b_margin:
                self.add_page()

            y0 = self.get_y()
            x0 = self.l_margin
            for ci, cell in enumerate(row_cells):
                x = x0 + ci * col_w
                self.rect(x, y0, col_w, row_h)
                self._render_cell_segments(
                    cell["segments"], x, y0, col_w, line_h, size,
                    align="L", bold=False)
            self.set_xy(x0, y0 + row_h)

        self.ln(2)

    def render_image(self, item: dict):
        """이미지를 렌더링합니다."""
        img_path = item.get("path", "")
        caption = item.get("caption", "")

        if not img_path:
            return

        # 상대경로 → 절대경로
        p = Path(img_path)
        if not p.is_absolute() and self.base_dir:
            p = Path(self.base_dir) / img_path

        if not p.exists():
            self._set_serif(10)
            self.cell(0, 5, f"[이미지 없음: {img_path}]", new_x="LMARGIN", new_y="NEXT")
            return

        usable_w = self.w - self.l_margin - self.r_margin
        max_w = usable_w * 0.8
        self.ln(3)
        x = self.l_margin + (usable_w - max_w) / 2
        self.image(str(p), x=x, w=max_w)

        if caption:
            self.ln(1)
            self._set_gothic(11)
            self.set_text_color(85, 85, 85)
            self.cell(0, 5, caption, align="C", new_x="LMARGIN", new_y="NEXT")
            self.set_text_color(0, 0, 0)
        self.ln(3)


# ---------------------------------------------------------------------------
# 공개 API
# ---------------------------------------------------------------------------

def generate_pdf(data: dict, output_path: str, styles_path: str = "", base_dir: str = "") -> str:
    """HWPX JSON 데이터로부터 PDF를 생성합니다.

    Args:
        data: parse_markdown_to_json()으로 생성된 HWPX JSON 데이터
        output_path: PDF 저장 경로
        styles_path: proposal-styles.json 경로
        base_dir: 이미지 등의 기본 경로

    Returns:
        생성된 PDF 파일 경로
    """
    # 스타일 로드
    if styles_path and Path(styles_path).exists():
        with open(styles_path, "r", encoding="utf-8") as f:
            styles = json.load(f)
    else:
        styles = {"styles": {}, "colors": {}, "lineSpacing": 160}

    metadata = data.get("metadata", {})
    content = data.get("content", [])

    pdf = ProposalPDF(styles, base_dir=base_dir)
    pdf.add_page()

    # 문서 제목
    title = metadata.get("title", "")
    if title and metadata.get("include_title"):
        pdf.render_title(title)

    # 섹션 렌더링
    for section in content:
        if section.get("type") != "section":
            continue

        section_title = section.get("title", "")
        if section_title:
            pdf.render_section_title(section_title)

        for item in section.get("items", []):
            item_type = item.get("type", "")

            if item_type == "subtitle":
                pdf.render_subtitle(item.get("text", ""))
            elif item_type == "table":
                pdf.render_table(item)
            elif item_type == "image":
                pdf.render_image(item)
            else:
                level = item.get("level", 2)
                pdf.render_text_item(level, item.get("text", ""))

    out = Path(output_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    pdf.output(str(out))

    _log(f"[PDF] 생성 완료: {out} ({out.stat().st_size:,} bytes)")
    return str(out)
