# -*- coding: utf-8 -*-
"""
Template Spec — HWPX 양식에서 추출한 모든 스타일/레이아웃 정보

양식 HWPX 파일을 분석하면 이 구조체가 생성되고,
HWPX 생성 시 이 구조체의 값만 사용한다 (하드코딩 제거).
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path


# ---------------------------------------------------------------------------
# 기본 단위 변환
# ---------------------------------------------------------------------------
def pt_to_height(pt: float) -> int:
    """포인트 → charPr height (1pt = 100)"""
    return int(pt * 100)

def pt_to_hwpunit(pt: float) -> int:
    """포인트 → HWPUNIT (1pt = 50)"""
    return int(pt * 50)

def hwpunit_to_pt(hu: int) -> float:
    return hu / 50.0

def height_to_pt(h: int) -> float:
    return h / 100.0

def mm_to_hwpunit(mm: float) -> int:
    """밀리미터 → HWPUNIT"""
    return int(mm * 7200 / 25.4 / 2)


# ---------------------------------------------------------------------------
# 데이터 모델
# ---------------------------------------------------------------------------

@dataclass
class PageSpec:
    """페이지 크기 및 여백 (HWPUNIT)"""
    width: int = 59528              # A4 가로
    height: int = 84186             # A4 세로
    margin_left: int = 8504
    margin_right: int = 8504
    margin_top: int = 5668
    margin_bottom: int = 4252
    margin_header: int = 4252
    margin_footer: int = 4252
    gutter: int = 0
    column_spacing: int = 1134
    tab_stop: int = 8000

    @property
    def body_width(self) -> int:
        """본문 영역 폭 = 페이지폭 - 좌여백 - 우여백"""
        return self.width - self.margin_left - self.margin_right


@dataclass
class TextStyle:
    """텍스트 스타일"""
    font: str = "바탕"
    size_pt: float = 10.0
    bold: bool = False
    italic: bool = False
    color: str = "#000000"
    align: str = "JUSTIFY"
    line_spacing: int = 160
    left_indent_pt: float = 0.0
    hanging_indent_pt: float = 0.0
    space_before_pt: float = 0.0
    space_after_pt: float = 0.0


@dataclass
class TableSpec:
    """표 관련 치수 (HWPUNIT)"""
    cell_margin_left: int = 510
    cell_margin_right: int = 510
    cell_margin_top: int = 141
    cell_margin_bottom: int = 141
    cell_padding_h: int = 1020      # 가로 패딩 (inner_width 계산용)
    cell_padding_v: int = 282       # 세로 패딩 (cell_height 계산용)
    outer_margin: int = 283
    header_bg_color: str = "#D9D9D9"
    border_type: str = "SOLID"
    border_width: str = "0.12 mm"
    border_color: str = "#000000"
    default_line_spacing: int = 160


@dataclass
class ImageSpec:
    """이미지 처리 설정"""
    px_to_hwpunit: int = 75         # 픽셀 → HWPUNIT 변환 비율
    max_width: int = 0              # 0이면 body_width 사용
    default_size: tuple[int, int] = (800, 600)  # 읽기 실패 시 기본 크기


@dataclass
class FootnoteSpec:
    """각주/미주 설정 (HWPUNIT)"""
    fn_between: int = 283
    fn_below_line: int = 567
    fn_above_line: int = 850
    en_line_length: int = 14692344
    en_between: int = 0
    en_below_line: int = 567
    en_above_line: int = 850


@dataclass
class ParagraphMetrics:
    """문단 레이아웃 메트릭스"""
    default_textheight: int = 1000
    default_baseline: int = 850
    default_spacing: int = 600


@dataclass
class ExamSpec:
    """시험 문제집 전용 설정"""
    choice_numbering: str = "circled"
    choice_count: int = 5
    questions_per_page: int = 5
    separator: str = "line"


@dataclass
class HeadingMapping:
    """마크다운 헤딩 → 레벨/기호 매핑"""
    markdown: str = "#"
    level: int = 1
    symbol: str = ""


@dataclass
class ListMapping:
    """리스트 들여쓰기 → 레벨 매핑"""
    indent_min: int = 0
    indent_max: int = 0             # 0이면 무제한
    level: int = 3
    symbol: str = ""


@dataclass
class MarkdownMapping:
    """마크다운 파싱 규칙"""
    headings: list[HeadingMapping] = field(default_factory=lambda: [
        HeadingMapping("#", 1, ""),
        HeadingMapping("##", 2, ""),
        HeadingMapping("###", 3, "◻"),
        HeadingMapping("####", 4, "○"),
        HeadingMapping("#####", 5, "―"),
        HeadingMapping("######", 6, "※"),
    ])
    lists: list[ListMapping] = field(default_factory=lambda: [
        ListMapping(0, 0, 3, ""),
        ListMapping(1, 4, 4, ""),
        ListMapping(5, 999, 5, ""),
    ])
    table_caption_max_chars: int = 60


# ---------------------------------------------------------------------------
# 최상위 Template Spec
# ---------------------------------------------------------------------------

@dataclass
class TemplateSpec:
    """HWPX 양식에서 추출한 전체 스타일 스펙.
    HWPX 생성 시 이 객체의 값만 참조한다."""

    page: PageSpec = field(default_factory=PageSpec)
    image: ImageSpec = field(default_factory=ImageSpec)
    table: TableSpec = field(default_factory=TableSpec)
    footnote: FootnoteSpec = field(default_factory=FootnoteSpec)
    paragraph: ParagraphMetrics = field(default_factory=ParagraphMetrics)
    markdown: MarkdownMapping = field(default_factory=MarkdownMapping)
    exam: ExamSpec = field(default_factory=ExamSpec)

    # 색상 매핑
    colors: dict[str, str] = field(default_factory=lambda: {
        "red": "#DC2626", "green": "#16A34A", "blue": "#2563EB",
        "yellow": "#EAB308", "black": "#000000",
    })

    # 글로벌 줄간격 (styles에서 미지정 시 사용)
    line_spacing: int = 160

    # 스타일 그룹: 용도별 텍스트 스타일
    styles: dict[str, TextStyle] = field(default_factory=lambda: {
        # 공통
        "title": TextStyle(font="맑은 고딕", size_pt=16, bold=True, align="CENTER",
                           space_before_pt=20, space_after_pt=10),
        "body": TextStyle(font="바탕", size_pt=10, align="JUSTIFY", line_spacing=160),
        # 제안서용 (hwpx_writer 호환)
        "level1": TextStyle(font="Noto Sans KR", size_pt=13, bold=True,
                            space_before_pt=10, space_after_pt=6),
        "level2": TextStyle(font="Noto Serif KR", size_pt=10,
                            space_before_pt=0, space_after_pt=3),
        "level3": TextStyle(font="Noto Serif KR", size_pt=10,
                            left_indent_pt=0, hanging_indent_pt=15,
                            space_before_pt=0, space_after_pt=3),
        "level4": TextStyle(font="Noto Serif KR", size_pt=10,
                            left_indent_pt=4, hanging_indent_pt=15.3,
                            space_before_pt=0, space_after_pt=3),
        "level5": TextStyle(font="Noto Serif KR", size_pt=10,
                            left_indent_pt=8, hanging_indent_pt=15.3,
                            space_before_pt=0, space_after_pt=3),
        "level6": TextStyle(font="Noto Serif KR", size_pt=10,
                            left_indent_pt=12, hanging_indent_pt=15.3,
                            space_before_pt=0, space_after_pt=3),
        # 시험 문제집용
        "exam_title": TextStyle(font="맑은 고딕", size_pt=16, bold=True, align="CENTER",
                                space_before_pt=20, space_after_pt=10),
        "exam_info": TextStyle(font="맑은 고딕", size_pt=10, align="CENTER",
                               space_before_pt=2, space_after_pt=2),
        "question_number": TextStyle(font="맑은 고딕", size_pt=11, bold=True, align="LEFT",
                                     space_before_pt=8, space_after_pt=2),
        "question_text": TextStyle(font="바탕", size_pt=10, align="JUSTIFY",
                                   line_spacing=160, space_after_pt=2),
        "choice": TextStyle(font="바탕", size_pt=10, align="LEFT", line_spacing=150,
                            left_indent_pt=10, hanging_indent_pt=12),
        "section_header": TextStyle(font="맑은 고딕", size_pt=14, bold=True, align="CENTER",
                                    space_before_pt=15, space_after_pt=8),
        "explanation_header": TextStyle(font="맑은 고딕", size_pt=11, bold=True, align="LEFT",
                                        space_before_pt=10, space_after_pt=2),
        "explanation_text": TextStyle(font="바탕", size_pt=9.5, align="JUSTIFY",
                                      line_spacing=150, space_after_pt=2),
        "answer_table": TextStyle(font="맑은 고딕", size_pt=10, align="CENTER"),
        # 표 스타일
        "table": TextStyle(font="Noto Sans KR", size_pt=9),
        "table_header": TextStyle(font="Noto Sans KR", size_pt=9, align="CENTER"),
        "table_data": TextStyle(font="Noto Sans KR", size_pt=9, align="LEFT"),
        "table_caption": TextStyle(font="Noto Sans KR", size_pt=11, bold=True, align="CENTER",
                                    space_before_pt=10, space_after_pt=3),
    })

    # -- 편의 메서드 --

    def get_style(self, name: str) -> TextStyle:
        """이름으로 스타일 조회. 없으면 body 반환."""
        return self.styles.get(name, self.styles.get("body", TextStyle()))

    def resolve_color(self, name: str) -> str:
        return self.colors.get(name.lower(), "#000000").upper()

    @property
    def body_width(self) -> int:
        return self.page.body_width

    @property
    def image_max_width(self) -> int:
        return self.image.max_width if self.image.max_width else self.page.body_width

    # -- 직렬화 --

    def to_dict(self) -> dict:
        """JSON 직렬화"""
        def _dc(obj):
            if hasattr(obj, '__dataclass_fields__'):
                return {k: _dc(getattr(obj, k)) for k in obj.__dataclass_fields__}
            elif isinstance(obj, dict):
                return {k: _dc(v) for k, v in obj.items()}
            elif isinstance(obj, list):
                return [_dc(v) for v in obj]
            elif isinstance(obj, tuple):
                return list(obj)
            return obj
        return _dc(self)

    def save(self, path: str | Path):
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def from_dict(cls, d: dict) -> TemplateSpec:
        """JSON 딕셔너리에서 복원"""
        spec = cls()

        # page
        if "page" in d:
            for k, v in d["page"].items():
                if hasattr(spec.page, k):
                    setattr(spec.page, k, v)

        # image
        if "image" in d:
            for k, v in d["image"].items():
                if k == "default_size":
                    spec.image.default_size = tuple(v)
                elif hasattr(spec.image, k):
                    setattr(spec.image, k, v)

        # table
        if "table" in d:
            for k, v in d["table"].items():
                if hasattr(spec.table, k):
                    setattr(spec.table, k, v)

        # footnote
        if "footnote" in d:
            for k, v in d["footnote"].items():
                if hasattr(spec.footnote, k):
                    setattr(spec.footnote, k, v)

        # paragraph
        if "paragraph" in d:
            for k, v in d["paragraph"].items():
                if hasattr(spec.paragraph, k):
                    setattr(spec.paragraph, k, v)

        # colors
        if "colors" in d:
            spec.colors = d["colors"]

        # line_spacing
        if "line_spacing" in d:
            spec.line_spacing = d["line_spacing"]

        # styles
        if "styles" in d:
            for name, sd in d["styles"].items():
                spec.styles[name] = TextStyle(
                    **{k: v for k, v in sd.items() if k in TextStyle.__dataclass_fields__}
                )

        # markdown mapping
        if "markdown" in d:
            md = d["markdown"]
            if "headings" in md:
                spec.markdown.headings = [
                    HeadingMapping(**h) for h in md["headings"]
                ]
            if "lists" in md:
                spec.markdown.lists = [
                    ListMapping(**l) for l in md["lists"]
                ]
            if "table_caption_max_chars" in md:
                spec.markdown.table_caption_max_chars = md["table_caption_max_chars"]

        # exam
        if "exam" in d:
            for k, v in d["exam"].items():
                if hasattr(spec.exam, k):
                    setattr(spec.exam, k, v)

        return spec

    @classmethod
    def load(cls, path: str | Path) -> TemplateSpec:
        with open(path, "r", encoding="utf-8") as f:
            return cls.from_dict(json.load(f))

    @classmethod
    def from_legacy_styles(cls, styles_path: str | Path) -> TemplateSpec:
        """기존 proposal-styles.json에서 TemplateSpec 변환 (하위호환)"""
        with open(styles_path, "r", encoding="utf-8") as f:
            data = json.load(f)

        spec = cls()
        style_config = data.get("styles", {})
        spec.line_spacing = data.get("lineSpacing", 160)

        if "colors" in data:
            spec.colors = data["colors"]

        # level 스타일 변환
        for key, sd in style_config.items():
            ts = TextStyle(
                font=sd.get("font", "바탕"),
                size_pt=sd.get("size", 10),
                bold=sd.get("bold", False),
                align=sd.get("align", "justify").upper(),
                line_spacing=sd.get("lineSpacing", spec.line_spacing),
                left_indent_pt=sd.get("leftMargin", 0),
                hanging_indent_pt=sd.get("hangingIndent", 0),
                space_before_pt=sd.get("paragraphSpaceBefore", 0),
                space_after_pt=sd.get("paragraphSpaceAfter", 0),
            )
            spec.styles[key] = ts

        # Extract symbols from legacy styles into markdown headings
        level_symbols = {}
        for key, sd in style_config.items():
            if key.startswith("level") and "symbol" in sd:
                try:
                    lvl = int(key.replace("level", ""))
                    level_symbols[lvl] = sd["symbol"]
                except ValueError:
                    pass
        for heading in spec.markdown.headings:
            if heading.level in level_symbols:
                heading.symbol = level_symbols[heading.level]

        return spec
