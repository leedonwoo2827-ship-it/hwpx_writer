# -*- coding: utf-8 -*-
"""
HWPX 양식 분석기 — HWPX 파일에서 TemplateSpec을 추출

ZIP 해체 → XML 파싱 → 페이지/폰트/스타일/테이블 치수 추출 → TemplateSpec
"""

import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from template_spec import (
    TemplateSpec, PageSpec, TextStyle, TableSpec, ImageSpec,
    FootnoteSpec, ParagraphMetrics, ExamSpec,
    hwpunit_to_pt, height_to_pt,
)

NS = {
    "hh": "http://www.hancom.co.kr/hwpml/2011/head",
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
}

# 보기/문제 번호 패턴
CHOICE_PATTERN = re.compile(r'^[①②③④⑤⑥⑦⑧⑨⑩]\s')
QUESTION_NUM_PATTERN = re.compile(r'^\d{1,3}[\.\)]\s')


class HWPXAnalyzer:
    """HWPX 파일을 분석하여 TemplateSpec을 추출"""

    def __init__(self, hwpx_path: str):
        self.hwpx_path = Path(hwpx_path)
        self._fonts: dict[str, str] = {}
        self._charpr: dict[str, dict] = {}
        self._parapr: dict[str, dict] = {}

    def analyze(self) -> TemplateSpec:
        """전체 분석 실행 → TemplateSpec 반환"""
        spec = TemplateSpec()

        with zipfile.ZipFile(self.hwpx_path, "r") as zf:
            header = self._read_xml(zf, "Contents/header.xml")
            if header is not None:
                self._parse_fonts(header)
                self._parse_charpr(header)
                self._parse_parapr(header)
                self._extract_table_spec(header, spec)

            section = self._read_xml(zf, "Contents/section0.xml")
            if section is not None:
                self._extract_page_spec(section, spec)
                paragraphs = self._extract_paragraphs(section)
                self._classify_styles(paragraphs, spec)
                self._detect_exam_patterns(paragraphs, spec)

        return spec

    # -- XML 읽기 --

    def _read_xml(self, zf: zipfile.ZipFile, name: str) -> ET.Element | None:
        try:
            return ET.fromstring(zf.read(name))
        except (KeyError, ET.ParseError):
            return None

    # -- header.xml 파싱 --

    def _parse_fonts(self, root: ET.Element):
        for fontface in root.iter(f'{{{NS["hh"]}}}fontface'):
            if fontface.get("lang") == "HANGUL":
                for font in fontface.iter(f'{{{NS["hh"]}}}font'):
                    self._fonts[font.get("id", "")] = font.get("face", "")

    def _parse_charpr(self, root: ET.Element):
        for cp in root.iter(f'{{{NS["hh"]}}}charPr'):
            cid = cp.get("id", "")
            font_ref = cp.find(f'{{{NS["hh"]}}}fontRef')
            font_id = font_ref.get("hangul", "0") if font_ref is not None else "0"
            self._charpr[cid] = {
                "height": int(cp.get("height", "1000")),
                "bold": cp.get("bold", "0") == "1",
                "color": cp.get("textColor", "#000000"),
                "font_id": font_id,
                "font_name": self._fonts.get(font_id, "바탕"),
            }

    def _parse_parapr(self, root: ET.Element):
        for pp in root.iter(f'{{{NS["hh"]}}}paraPr'):
            pid = pp.get("id", "")
            align_el = pp.find(f'{{{NS["hh"]}}}align')
            align = align_el.get("horizontal", "JUSTIFY") if align_el is not None else "JUSTIFY"

            margin = pp.find(f'{{{NS["hh"]}}}margin')
            left = space_before = space_after = indent = 0
            if margin is not None:
                left = self._get_val(margin, "left")
                space_before = self._get_val(margin, "prev")
                space_after = self._get_val(margin, "next")
                indent = self._get_val(margin, "intent")

            ls_el = pp.find(f'{{{NS["hh"]}}}lineSpacing')
            line_spacing = int(ls_el.get("value", "160")) if ls_el is not None else 160

            self._parapr[pid] = {
                "align": align, "left_margin": left,
                "space_before": space_before, "space_after": space_after,
                "indent": indent, "line_spacing": line_spacing,
            }

    def _get_val(self, margin_el: ET.Element, tag: str) -> int:
        el = margin_el.find(f'{{{NS["hc"]}}}{tag}')
        return int(el.get("value", "0")) if el is not None else 0

    def _extract_table_spec(self, root: ET.Element, spec: TemplateSpec):
        """borderFill에서 표 스타일 추출"""
        for bf in root.iter(f'{{{NS["hh"]}}}borderFill'):
            bf_id = bf.get("id", "")
            # id=4가 헤더 셀 (회색 배경)
            if bf_id == "4":
                fb = bf.find(f'{{{NS["hc"]}}}fillBrush')
                if fb is not None:
                    wb = fb.find(f'{{{NS["hc"]}}}winBrush')
                    if wb is not None:
                        color = wb.get("faceColor", "#D9D9D9")
                        if color != "none":
                            spec.table.header_bg_color = color
            # id=3이 본문 셀 (테두리 정보)
            if bf_id == "3":
                lb = bf.find(f'{{{NS["hh"]}}}leftBorder')
                if lb is not None:
                    spec.table.border_type = lb.get("type", "SOLID")
                    spec.table.border_width = lb.get("width", "0.12 mm")
                    spec.table.border_color = lb.get("color", "#000000")

    # -- section0.xml 파싱 --

    def _extract_page_spec(self, root: ET.Element, spec: TemplateSpec):
        for secpr in root.iter(f'{{{NS["hp"]}}}secPr'):
            # 칼럼 간격, 탭
            spec.page.column_spacing = int(secpr.get("spaceColumns", "1134"))
            spec.page.tab_stop = int(secpr.get("tabStop", "8000"))

            page_pr = secpr.find(f'{{{NS["hp"]}}}pagePr')
            if page_pr is not None:
                spec.page.width = int(page_pr.get("width", "59528"))
                spec.page.height = int(page_pr.get("height", "84186"))
                margin = page_pr.find(f'{{{NS["hp"]}}}margin')
                if margin is not None:
                    spec.page.margin_left = int(margin.get("left", "8504"))
                    spec.page.margin_right = int(margin.get("right", "8504"))
                    spec.page.margin_top = int(margin.get("top", "5668"))
                    spec.page.margin_bottom = int(margin.get("bottom", "4252"))
                    spec.page.margin_header = int(margin.get("header", "4252"))
                    spec.page.margin_footer = int(margin.get("footer", "4252"))
                    spec.page.gutter = int(margin.get("gutter", "0"))

            # 각주 설정
            fn = secpr.find(f'{{{NS["hp"]}}}footNotePr')
            if fn is not None:
                ns = fn.find(f'{{{NS["hp"]}}}noteSpacing')
                if ns is not None:
                    spec.footnote.fn_between = int(ns.get("betweenNotes", "283"))
                    spec.footnote.fn_below_line = int(ns.get("belowLine", "567"))
                    spec.footnote.fn_above_line = int(ns.get("aboveLine", "850"))

            en = secpr.find(f'{{{NS["hp"]}}}endNotePr')
            if en is not None:
                nl = en.find(f'{{{NS["hp"]}}}noteLine')
                if nl is not None:
                    spec.footnote.en_line_length = int(nl.get("length", "14692344"))
                ns2 = en.find(f'{{{NS["hp"]}}}noteSpacing')
                if ns2 is not None:
                    spec.footnote.en_between = int(ns2.get("betweenNotes", "0"))
                    spec.footnote.en_below_line = int(ns2.get("belowLine", "567"))
                    spec.footnote.en_above_line = int(ns2.get("aboveLine", "850"))

            # 페이지 보더
            for pbf in secpr.iter(f'{{{NS["hp"]}}}pageBorderFill'):
                offs = pbf.find(f'{{{NS["hp"]}}}offset')
                if offs is not None:
                    # 첫 번째 것만 사용
                    break
            break  # 첫 번째 secPr만

    def _extract_paragraphs(self, root: ET.Element) -> list[dict]:
        """section0.xml에서 문단 텍스트 + 스타일 ID 추출"""
        paragraphs = []
        for p in root.iter(f'{{{NS["hp"]}}}p'):
            parapr_id = p.get("paraPrIDRef", "0")
            text_parts = []
            for run in p.iter(f'{{{NS["hp"]}}}run'):
                charpr_id = run.get("charPrIDRef", "0")
                for t in run.iter(f'{{{NS["hp"]}}}t'):
                    if t.text:
                        text_parts.append((t.text, charpr_id))
            full_text = "".join(t for t, _ in text_parts)
            first_charpr = text_parts[0][1] if text_parts else "0"
            paragraphs.append({
                "text": full_text, "charpr_id": first_charpr,
                "parapr_id": parapr_id,
            })
        return paragraphs

    def _classify_styles(self, paragraphs: list[dict], spec: TemplateSpec):
        """문단 패턴을 분류하여 styles에 반영"""
        question_styles = []
        choice_styles = []
        title_styles = []

        for para in paragraphs:
            text = para["text"].strip()
            if not text:
                continue
            cp = self._charpr.get(para["charpr_id"], {})
            pp = self._parapr.get(para["parapr_id"], {})
            info = self._build_style_info(cp, pp)

            if QUESTION_NUM_PATTERN.match(text):
                question_styles.append(info)
            elif CHOICE_PATTERN.match(text):
                choice_styles.append(info)
            elif height_to_pt(cp.get("height", 1000)) >= 13:
                title_styles.append(info)

        # 대표 스타일 적용
        if question_styles:
            rep = self._most_common(question_styles)
            spec.styles["question_number"] = TextStyle(**{**rep, "bold": True})
            spec.styles["question_text"] = TextStyle(**{**rep, "bold": False})
        if choice_styles:
            rep = self._most_common(choice_styles)
            spec.styles["choice"] = TextStyle(**rep)
        if title_styles:
            rep = self._most_common(title_styles)
            spec.styles["exam_title"] = TextStyle(**rep)
            spec.styles["section_header"] = TextStyle(**rep)

    def _detect_exam_patterns(self, paragraphs: list[dict], spec: TemplateSpec):
        """시험 문제집 패턴 감지"""
        q_count = 0
        has_circled = False
        for p in paragraphs:
            text = p["text"].strip()
            if QUESTION_NUM_PATTERN.match(text):
                q_count += 1
            if CHOICE_PATTERN.match(text):
                has_circled = True

        if has_circled:
            spec.exam.choice_numbering = "circled"
        if q_count > 0:
            total_paras = len([p for p in paragraphs if p["text"].strip()])
            pages_est = max(1, total_paras // 30)
            spec.exam.questions_per_page = max(1, q_count // pages_est)

    def _build_style_info(self, cp: dict, pp: dict) -> dict:
        return {
            "font": cp.get("font_name", "바탕"),
            "size_pt": height_to_pt(cp.get("height", 1000)),
            "bold": cp.get("bold", False),
            "color": cp.get("color", "#000000"),
            "align": pp.get("align", "JUSTIFY"),
            "line_spacing": pp.get("line_spacing", 160),
            "left_indent_pt": hwpunit_to_pt(pp.get("left_margin", 0)),
            "hanging_indent_pt": abs(hwpunit_to_pt(pp.get("indent", 0))),
            "space_before_pt": hwpunit_to_pt(pp.get("space_before", 0)),
            "space_after_pt": hwpunit_to_pt(pp.get("space_after", 0)),
        }

    @staticmethod
    def _most_common(styles: list[dict]) -> dict:
        from collections import Counter
        keys = Counter()
        for s in styles:
            keys[(s["font"], s["size_pt"], s["bold"])] += 1
        top = keys.most_common(1)[0][0]
        for s in styles:
            if (s["font"], s["size_pt"], s["bold"]) == top:
                return s
        return styles[0]


def analyze_hwpx(hwpx_path: str) -> TemplateSpec:
    """편의 함수: HWPX → TemplateSpec"""
    return HWPXAnalyzer(hwpx_path).analyze()


def analyze_and_save(hwpx_path: str, output_path: str) -> str:
    """HWPX 분석 → template-spec.json 저장"""
    spec = analyze_hwpx(hwpx_path)
    spec.save(output_path)
    return output_path
