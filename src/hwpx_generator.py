# -*- coding: utf-8 -*-
"""
HWPX Generator — 문자열 기반 XML 생성

lxml의 자동 네임스페이스 프리픽스(ns0:, ns1:, ...) 문제를 근본적으로 회피하기 위해
XML을 문자열로 직접 생성한다.

한글 오피스가 요구하는 HWPX 구조:
  - mimetype: "application/hwp+zip" (ZIP_STORED, 첫 번째 엔트리)
  - META-INF/container.xml, container.rdf, manifest.xml
  - version.xml, settings.xml
  - Contents/content.hpf, header.xml, section0.xml
  - 모든 Contents/*.xml에 전체 네임스페이스 선언 필수
"""

import json
import os
import re
import sys
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape as xml_escape


def _log(msg: str) -> None:
    print(msg, file=sys.stderr)


# ---------------------------------------------------------------------------
# 네임스페이스 상수
# ---------------------------------------------------------------------------
_ALL_NS = (
    'xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" '
    'xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" '
    'xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" '
    'xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" '
    'xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" '
    'xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" '
    'xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" '
    'xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" '
    'xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:opf="http://www.idpf.org/2007/opf/" '
    'xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" '
    'xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" '
    'xmlns:epub="http://www.idpf.org/2007/ops" '
    'xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"'
)


class HWPXGenerator:
    def __init__(self, base_dir: str = None, styles_path: str = "proposal-styles.json"):
        self.base_dir = Path(base_dir) if base_dir else Path.cwd()

        sp = Path(styles_path)
        self.styles_path = sp if sp.is_absolute() else self.base_dir / sp

        with open(self.styles_path, "r", encoding="utf-8") as f:
            styles_data = json.load(f)
            self.style_config = styles_data["styles"]
            self.colors = styles_data.get("colors", {})

        # CharPr / ParaPr ID 카운터 (0번은 기본용으로 예약)
        self._charpr_list = []   # (id, height, textColor, fontId)
        self._parapr_list = []   # (id, leftMargin, spaceBefore, spaceAfter, align)
        self._next_charpr_id = 1
        self._next_parapr_id = 1
        self._charpr_cache = {}  # (height, textColor, fontId) -> id
        self._parapr_cache = {}  # (leftMargin, spaceBefore, spaceAfter, align) -> id

        # 폰트 매핑: font_name -> id (0번은 기본 폰트)
        self._fonts = []         # [(id, face)] — 순서대로
        self._font_cache = {}    # face -> id

    # ------------------------------------------------------------------
    # 폰트 관리
    # ------------------------------------------------------------------
    def _register_font(self, face: str) -> int:
        if face in self._font_cache:
            return self._font_cache[face]
        fid = len(self._fonts)
        self._fonts.append((fid, face))
        self._font_cache[face] = fid
        return fid

    def _collect_fonts_from_styles(self):
        """스타일에서 사용하는 폰트를 모두 등록"""
        # 기본 폰트 (id=0)
        default_fonts = ["함초롬돋움", "함초롬바탕"]
        for f in default_fonts:
            self._register_font(f)

        # 스타일에서 사용하는 폰트
        for key, style in self.style_config.items():
            font = style.get("font")
            if font:
                self._register_font(font)

    # ------------------------------------------------------------------
    # CharPr 관리
    # ------------------------------------------------------------------
    def _get_charpr_id(self, height: int, text_color: str, font_name: str) -> int:
        font_id = self._register_font(font_name)
        key = (height, text_color.upper(), font_id)
        if key in self._charpr_cache:
            return self._charpr_cache[key]
        cid = self._next_charpr_id
        self._next_charpr_id += 1
        self._charpr_list.append((cid, height, text_color.upper(), font_id))
        self._charpr_cache[key] = cid
        return cid

    # ------------------------------------------------------------------
    # ParaPr 관리
    # ------------------------------------------------------------------
    def _get_parapr_id(self, left_margin: int = 0, space_before: int = 0,
                       space_after: int = 0, align: str = "JUSTIFY") -> int:
        key = (left_margin, space_before, space_after, align)
        if key in self._parapr_cache:
            return self._parapr_cache[key]
        pid = self._next_parapr_id
        self._next_parapr_id += 1
        self._parapr_list.append((pid, left_margin, space_before, space_after, align))
        self._parapr_cache[key] = pid
        return pid

    # ------------------------------------------------------------------
    # pt → HWPX 변환
    # ------------------------------------------------------------------
    @staticmethod
    def _pt_to_height(pt: float) -> int:
        """포인트를 charPr height 단위로 변환 (1pt = 100)"""
        return int(pt * 100)

    @staticmethod
    def _pt_to_hwpunit(pt: float) -> int:
        """포인트를 HWPUNIT로 변환 (1pt ≈ 100 HWPUNIT)"""
        return int(pt * 100)

    # ------------------------------------------------------------------
    # 색상
    # ------------------------------------------------------------------
    def _resolve_color(self, name: str) -> str:
        return self.colors.get(name.lower(), "#000000").upper()

    # ------------------------------------------------------------------
    # 마커 파싱 {{red:텍스트}}
    # ------------------------------------------------------------------
    @staticmethod
    def _parse_markers(text: str):
        """텍스트에서 {{color:...}} 마커를 파싱하여 segments 리스트 반환"""
        # HTML 태그 제거
        text = re.sub(r'<[^>]+>', '', text)
        text = text.replace('&nbsp;', ' ').replace('&lt;', '<').replace('&gt;', '>')
        text = text.replace('&amp;', '&').replace('&quot;', '"')

        segments = []
        pos = 0
        for m in re.finditer(r'\{\{(\w+):([^}]+)\}\}', text):
            if m.start() > pos:
                segments.append((text[pos:m.start()], None))
            segments.append((m.group(2), m.group(1)))
            pos = m.end()
        if pos < len(text):
            segments.append((text[pos:], None))
        return segments if segments else [(text, None)]

    # ==================================================================
    # XML 생성 (문자열 기반)
    # ==================================================================

    def _build_fontfaces_xml(self) -> str:
        """fontfaces XML 생성"""
        langs = ["HANGUL", "LATIN", "HANJA", "JAPANESE", "OTHER", "SYMBOL", "USER"]
        font_cnt = len(self._fonts)

        font_elems = ""
        for fid, face in self._fonts:
            font_elems += (
                f'<hh:font id="{fid}" face="{xml_escape(face)}" type="TTF" isEmbedded="0">'
                f'<hh:typeInfo familyType="FCAT_GOTHIC" weight="6" proportion="4"'
                f' contrast="0" strokeVariation="1" armStyle="1"'
                f' letterform="1" midline="1" xHeight="1"/>'
                f'</hh:font>'
            )

        faces_xml = ""
        for lang in langs:
            faces_xml += (
                f'<hh:fontface lang="{lang}" fontCnt="{font_cnt}">'
                f'{font_elems}'
                f'</hh:fontface>'
            )

        return f'<hh:fontfaces itemCnt="{len(langs)}">{faces_xml}</hh:fontfaces>'

    def _build_charpr_xml(self, cid: int, height: int, text_color: str, font_id: int) -> str:
        fid = str(font_id)
        return (
            f'<hh:charPr id="{cid}" height="{height}" textColor="{text_color}"'
            f' shadeColor="none" useFontSpace="0" useKerning="0"'
            f' symMark="NONE" borderFillIDRef="2">'
            f'<hh:fontRef hangul="{fid}" latin="{fid}" hanja="{fid}"'
            f' japanese="{fid}" other="{fid}" symbol="{fid}" user="{fid}"/>'
            f'<hh:ratio hangul="100" latin="100" hanja="100"'
            f' japanese="100" other="100" symbol="100" user="100"/>'
            f'<hh:spacing hangul="0" latin="0" hanja="0"'
            f' japanese="0" other="0" symbol="0" user="0"/>'
            f'<hh:relSz hangul="100" latin="100" hanja="100"'
            f' japanese="100" other="100" symbol="100" user="100"/>'
            f'<hh:offset hangul="0" latin="0" hanja="0"'
            f' japanese="0" other="0" symbol="0" user="0"/>'
            f'</hh:charPr>'
        )

    def _build_parapr_xml(self, pid, left_margin, space_before, space_after, align) -> str:
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="0" snapToGrid="1"'
            f' suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="KEEP_WORD"'
            f' breakNonLatinWord="KEEP_WORD" widowOrphan="0"'
            f' keepWithNext="0" keepLines="0" pageBreakBefore="0"'
            f' lineWrap="BREAK"/>'
            f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            f'<hh:margin>'
            f'<hc:intent value="0" unit="HWPUNIT"/>'
            f'<hc:left value="{left_margin}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="{space_before}" unit="HWPUNIT"/>'
            f'<hc:next value="{space_after}" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            f' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            f'</hh:paraPr>'
        )

    def _build_borderfills_xml(self) -> str:
        """기본 borderFill + 표용 borderFill 생성"""
        # ID 1~3: 기본 (투명 테두리)
        bfs = ""
        for bid in range(1, 4):
            bfs += (
                f'<hh:borderFill id="{bid}" threeD="0" shadow="0"'
                f' centerLine="NONE" breakCellSeparateLine="0">'
                f'<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
                f'<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
                f'<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
                f'<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
                f'<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
                f'<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
                f'<hh:diagonal type="NONE" width="0.1 mm" color="#000000"/>'
                f'</hh:borderFill>'
            )
        # ID 4: 표 외곽용 (SOLID 테두리)
        bfs += (
            '<hh:borderFill id="4" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:diagonal type="NONE" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
        )
        # ID 5: 셀용 (SOLID 테두리, 연회색 배경)
        bfs += (
            '<hh:borderFill id="5" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#C4C4C4"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#C4C4C4"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#C4C4C4"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#C4C4C4"/>'
            '<hh:diagonal type="NONE" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
        )
        return f'<hh:borderFills itemCnt="5">{bfs}</hh:borderFills>'

    def _build_header_xml(self) -> str:
        """header.xml 전체 생성"""
        fontfaces = self._build_fontfaces_xml()

        # CharPr: id=0 (기본) + 동적 생성분
        charpr_default = self._build_charpr_xml(0, 1100, "#000000", 0)
        charprs = charpr_default
        for cid, height, color, fid in self._charpr_list:
            charprs += self._build_charpr_xml(cid, height, color, fid)
        charpr_cnt = 1 + len(self._charpr_list)

        # ParaPr: id=0 (기본) + 동적 생성분
        parapr_default = self._build_parapr_xml(0, 0, 0, 0, "JUSTIFY")
        paraprs = parapr_default
        for pid, lm, sb, sa, align in self._parapr_list:
            paraprs += self._build_parapr_xml(pid, lm, sb, sa, align)
        parapr_cnt = 1 + len(self._parapr_list)

        borderfills = self._build_borderfills_xml()

        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<hh:head {_ALL_NS} version="1.5" secCnt="1">'
            f'<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>'
            f'<hh:refList>'
            f'{fontfaces}'
            f'{borderfills}'
            f'<hh:charProperties itemCnt="{charpr_cnt}">{charprs}</hh:charProperties>'
            f'<hh:paraProperties itemCnt="{parapr_cnt}">{paraprs}</hh:paraProperties>'
            f'</hh:refList>'
            f'</hh:head>'
        )

    def _build_secpr_xml(self) -> str:
        """페이지 설정 (secPr) — A4, 상하좌우 여백"""
        return (
            '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134"'
            ' tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT"'
            ' outlineShapeIDRef="1" memoShapeIDRef="1"'
            ' textVerticalWidthHead="0" masterPageCnt="0">'
            '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
            '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
            '<hp:visibility hideFirstHeader="0" hideFirstFooter="0"'
            ' hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL"'
            ' hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
            '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
            '<hp:pagePr landscape="WIDELY" width="59528" height="84188" gutterType="LEFT_ONLY">'
            '<hp:margin header="4251" footer="4251" gutter="0"'
            ' left="5669" right="5669" top="4251" bottom="4251"/>'
            '</hp:pagePr>'
            '<hp:footNotePr>'
            '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
            '<hp:noteLine length="-1" type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hp:noteSpacing betweenNotes="283" belowLine="567" aboveLine="850"/>'
            '<hp:numbering type="CONTINUOUS" newNum="1"/>'
            '<hp:placement place="EACH_COLUMN" beneathText="0"/>'
            '</hp:footNotePr>'
            '<hp:endNotePr>'
            '<hp:autoNumFormat type="DIGIT" userChar="" prefixChar="" suffixChar=")" supscript="0"/>'
            '<hp:noteLine length="14692344" type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hp:noteSpacing betweenNotes="0" belowLine="567" aboveLine="850"/>'
            '<hp:numbering type="CONTINUOUS" newNum="1"/>'
            '<hp:placement place="END_OF_DOCUMENT" beneathText="0"/>'
            '</hp:endNotePr>'
            '<hp:pageBorderFill type="BOTH" borderFillIDRef="1" textBorder="PAPER"'
            ' headerInside="0" footerInside="0" fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '<hp:pageBorderFill type="EVEN" borderFillIDRef="1" textBorder="PAPER"'
            ' headerInside="0" footerInside="0" fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '<hp:pageBorderFill type="ODD" borderFillIDRef="1" textBorder="PAPER"'
            ' headerInside="0" footerInside="0" fillArea="PAPER">'
            '<hp:offset left="1417" right="1417" top="1417" bottom="1417"/>'
            '</hp:pageBorderFill>'
            '</hp:secPr>'
        )

    # ------------------------------------------------------------------
    # paragraph / run XML 빌더
    # ------------------------------------------------------------------
    def _run_xml(self, text: str, charpr_id: int) -> str:
        return (
            f'<hp:run charPrIDRef="{charpr_id}">'
            f'<hp:t>{xml_escape(text)}</hp:t>'
            f'</hp:run>'
        )

    def _paragraph_xml(self, runs_xml: str, parapr_id: int = 0,
                       page_break: str = "0", p_id: str = "0") -> str:
        return (
            f'<hp:p id="{p_id}" paraPrIDRef="{parapr_id}"'
            f' styleIDRef="0" pageBreak="{page_break}"'
            f' columnBreak="0" merged="0">'
            f'{runs_xml}'
            f'</hp:p>'
        )

    def _text_paragraph(self, text: str, level: int, font_name: str,
                        font_size_pt: float, left_margin_pt: float = 0,
                        space_before_pt: float = 0, space_after_pt: float = 0,
                        align: str = "JUSTIFY") -> str:
        """마커 색상을 지원하는 텍스트 paragraph 생성"""
        height = self._pt_to_height(font_size_pt)
        parapr_id = self._get_parapr_id(
            self._pt_to_hwpunit(left_margin_pt),
            self._pt_to_hwpunit(space_before_pt),
            self._pt_to_hwpunit(space_after_pt),
            align,
        )

        segments = self._parse_markers(text)
        runs = ""
        for seg_text, seg_color in segments:
            if seg_color:
                color_hex = self._resolve_color(seg_color)
            else:
                color_hex = "#000000"
            cid = self._get_charpr_id(height, color_hex, font_name)
            runs += self._run_xml(seg_text, cid)

        return self._paragraph_xml(runs, parapr_id)

    # ------------------------------------------------------------------
    # 표 XML 빌더
    # ------------------------------------------------------------------
    def _table_cell_xml(self, text: str, charpr_id: int, col_idx: int,
                        row_idx: int, cell_width: int) -> str:
        """표 셀 XML"""
        segments = self._parse_markers(text)

        # 셀 내부 run 들
        cell_runs = ""
        for seg_text, seg_color in segments:
            if seg_color:
                color_hex = self._resolve_color(seg_color)
                table_style = self.style_config.get("table", {})
                font = table_style.get("font", "함초롬돋움")
                size = table_style.get("size", 10)
                cid = self._get_charpr_id(self._pt_to_height(size), color_hex, font)
            else:
                cid = charpr_id
            cell_runs += self._run_xml(seg_text, cid)

        inner_width = max(cell_width - 1020, 1000)

        return (
            f'<hp:tc name="" header="0" hasMargin="0" protect="0"'
            f' editable="0" dirty="0" borderFillIDRef="5">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
            f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
            f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'{cell_runs}'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="1200"'
            f' textheight="1200" baseline="1020" spacing="720"'
            f' horzpos="0" horzsize="{inner_width}" flags="393216"/>'
            f'</hp:linesegarray>'
            f'</hp:p>'
            f'</hp:subList>'
            f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
            f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{cell_width}" height="1765"/>'
            f'<hp:cellMargin left="510" right="510" top="141" bottom="141"/>'
            f'</hp:tc>'
        )

    def _table_xml(self, table_data: dict) -> str:
        """표 XML 생성"""
        headers = table_data.get("headers", [])
        rows = table_data.get("rows", [])
        col_count = len(headers) if headers else 1
        row_count = len(rows) + (1 if headers else 0)

        # 표 글자 스타일
        table_style = self.style_config.get("table", {})
        table_font = table_style.get("font", "함초롬돋움")
        table_size = table_style.get("size", 10)
        table_height = self._pt_to_height(table_size)
        table_charpr = self._get_charpr_id(table_height, "#000000", table_font)

        # A4 본문 폭 기준 (좌우 여백 제외): 59528 - 5669*2 = 48190
        total_width = 48190
        cell_width = total_width // col_count

        # 헤더 행
        header_row = ""
        if headers:
            for ci, h in enumerate(headers):
                h_text = h.get("text", "") if isinstance(h, dict) else str(h)
                header_row += self._table_cell_xml(h_text, table_charpr, ci, 0, cell_width)
            header_row = f'<hp:tr>{header_row}</hp:tr>'

        # 데이터 행
        data_rows = ""
        for ri, row in enumerate(rows):
            cells = ""
            row_idx = ri + (1 if headers else 0)
            for ci, cell in enumerate(row):
                c_text = cell.get("text", "") if isinstance(cell, dict) else str(cell)
                cells += self._table_cell_xml(c_text, table_charpr, ci, row_idx, cell_width)
            data_rows += f'<hp:tr>{cells}</hp:tr>'

        table_h = 1765 * row_count

        return (
            f'<hp:tbl id="0" zOrder="0" numberingType="TABLE"'
            f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
            f' dropcapstyle="None" pageBreak="CELL" repeatHeader="1"'
            f' rowCnt="{row_count}" colCnt="{col_count}"'
            f' cellSpacing="0" borderFillIDRef="4" noAdjust="0">'
            f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
            f' height="{table_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
            f' allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA"'
            f' horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT"'
            f' vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="141" bottom="141"/>'
            f'<hp:inMargin left="510" right="510" top="141" bottom="141"/>'
            f'{header_row}{data_rows}'
            f'</hp:tbl>'
        )

    def _table_paragraph_xml(self, table_data: dict) -> str:
        """표를 담는 paragraph XML"""
        tbl = self._table_xml(table_data)
        row_count = len(table_data.get("rows", [])) + (1 if table_data.get("headers") else 0)
        table_h = 1765 * row_count

        return (
            f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'{tbl}'
            f'<hp:t></hp:t>'
            f'</hp:run>'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="{table_h}" vertsize="1000"'
            f' textheight="1000" baseline="850" spacing="600"'
            f' horzpos="0" horzsize="0" flags="393216"/>'
            f'</hp:linesegarray>'
            f'</hp:p>'
        )

    # ==================================================================
    # 메인 생성
    # ==================================================================
    def generate(self, data: dict, output_path: str) -> str:
        """JSON 데이터를 기반으로 HWPX 문서 생성"""
        _log("[Step 1] Collecting fonts from styles...")
        self._collect_fonts_from_styles()

        _log("[Step 2] Building section content...")
        metadata = data.get("metadata", {})
        title = metadata.get("title", "제목 없음")

        body_paragraphs = ""

        # 첫 번째 paragraph: secPr (페이지 설정)을 포함
        secpr = self._build_secpr_xml()

        # 제목 추가
        include_title = metadata.get("include_title", False)
        if include_title and title:
            title_style = self.style_config.get("title", {})
            title_font = title_style.get("font", "함초롬돋움")
            title_size = title_style.get("size", 25)
            title_align = title_style.get("align", "center").upper()
            if title_align == "CENTER":
                title_align = "CENTER"
            title_height = self._pt_to_height(title_size)
            title_charpr = self._get_charpr_id(title_height, "#000000", title_font)
            title_parapr = self._get_parapr_id(0, 0, self._pt_to_hwpunit(10), title_align)

            runs = self._run_xml(title, title_charpr)
            # 첫 문단에 secPr 포함
            body_paragraphs += (
                f'<hp:p id="0" paraPrIDRef="{title_parapr}"'
                f' styleIDRef="0" pageBreak="0"'
                f' columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}</hp:run>'
                f'{runs}'
                f'</hp:p>'
            )
        else:
            # 제목 없으면 빈 첫 문단에 secPr 포함
            body_paragraphs += (
                f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
                f' pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}<hp:t></hp:t></hp:run>'
                f'</hp:p>'
            )

        # 콘텐츠 처리
        for item in data.get("content", []):
            item_type = item.get("type", "section")

            if item_type == "section":
                # 섹션 제목
                include_section_titles = metadata.get("include_section_titles", False)
                section_title = item.get("title")
                if include_section_titles and section_title:
                    sec_style = self.style_config.get("level1", {})
                    sec_font = sec_style.get("font", "함초롬돋움")
                    sec_size = sec_style.get("size", 18)
                    sec_height = self._pt_to_height(sec_size)
                    sec_charpr = self._get_charpr_id(sec_height, "#000000", sec_font)
                    sec_parapr = self._get_parapr_id(
                        0,
                        self._pt_to_hwpunit(sec_style.get("paragraphSpaceBefore", 25)),
                        self._pt_to_hwpunit(sec_style.get("paragraphSpaceAfter", 8)),
                        "JUSTIFY",
                    )
                    runs = self._run_xml(section_title, sec_charpr)
                    body_paragraphs += self._paragraph_xml(runs, sec_parapr)

                # 섹션 항목
                for sub in item.get("items", []):
                    sub_type = sub.get("type")
                    if sub_type == "table":
                        body_paragraphs += self._table_paragraph_xml(sub)
                        _log(f"[Added] Table: rows={len(sub.get('rows', []))}")
                    else:
                        level = sub.get("level", 1)
                        text = sub.get("text", "")
                        level_key = f"level{level}"
                        style = self.style_config.get(level_key, {})

                        body_paragraphs += self._text_paragraph(
                            text,
                            level,
                            style.get("font", "함초롬돋움"),
                            style.get("size", 11),
                            style.get("leftMargin", 0),
                            style.get("paragraphSpaceBefore", 0),
                            style.get("paragraphSpaceAfter", 3),
                            style.get("align", "justify").upper(),
                        )
                        _log(f"[Added] Level {level}: {text[:50]}...")

            elif item_type == "table":
                body_paragraphs += self._table_paragraph_xml(item)

        # 마지막 빈 문단 (한글 오피스 호환)
        body_paragraphs += self._paragraph_xml(
            '<hp:run charPrIDRef="0"><hp:t></hp:t></hp:run>', 0
        )

        _log("[Step 3] Building header.xml...")
        header_xml = self._build_header_xml()

        _log("[Step 4] Building section0.xml...")
        section_xml = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<hs:sec {_ALL_NS}>'
            f'{body_paragraphs}'
            f'</hs:sec>'
        )

        _log("[Step 5] Packing HWPX archive...")
        self._pack_hwpx(output_path, header_xml, section_xml)

        _log(f"[Success] HWPX generated: {output_path}")
        return output_path

    def _pack_hwpx(self, output_path: str, header_xml: str, section_xml: str):
        """HWPX ZIP 패키지 생성"""

        # -- 고정 파일들 --
        mimetype = "application/hwp+zip"

        version_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version"'
            ' tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1"'
            ' buildNumber="0" os="1" xmlVersion="1.5"'
            ' application="Hancom Office Hangul" appVersion="12, 0, 0, 535 WIN32LEWindows_10"/>'
        )

        container_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"'
            ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
            '<ocf:rootfiles>'
            '<ocf:rootfile full-path="Contents/content.hpf"'
            ' media-type="application/hwpml-package+xml"/>'
            '</ocf:rootfiles>'
            '</ocf:container>'
        )

        container_rdf = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<rdf:RDF xmlns:rdf="http://www.w3.org/1999/02/22-rdf-syntax-ns#">'
            '<rdf:Description rdf:about="">'
            '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
            ' rdf:resource="Contents/header.xml"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="Contents/header.xml">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#HeaderFile"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="">'
            '<ns0:hasPart xmlns:ns0="http://www.hancom.co.kr/hwpml/2016/meta/pkg#"'
            ' rdf:resource="Contents/section0.xml"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="Contents/section0.xml">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#SectionFile"/>'
            '</rdf:Description>'
            '<rdf:Description rdf:about="">'
            '<rdf:type rdf:resource="http://www.hancom.co.kr/hwpml/2016/meta/pkg#Document"/>'
            '</rdf:Description>'
            '</rdf:RDF>'
        )

        manifest_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>'
        )

        settings_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<ha:HWPApplicationSetting'
            ' xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app"'
            ' xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0">'
            '<ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/>'
            '</ha:HWPApplicationSetting>'
        )

        content_hpf = (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<opf:package {_ALL_NS}'
            f' version="" unique-identifier="" id="">'
            f'<opf:metadata>'
            f'<opf:title/>'
            f'<opf:language>ko</opf:language>'
            f'</opf:metadata>'
            f'<opf:manifest>'
            f'<opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>'
            f'<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/>'
            f'<opf:item id="settings" href="settings.xml" media-type="application/xml"/>'
            f'</opf:manifest>'
            f'<opf:spine>'
            f'<opf:itemref idref="header" linear="yes"/>'
            f'<opf:itemref idref="section0" linear="yes"/>'
            f'</opf:spine>'
            f'</opf:package>'
        )

        # -- ZIP 패키징 --
        with zipfile.ZipFile(output_path, "w") as zf:
            # mimetype: 반드시 첫 번째, 비압축(STORED)
            mime_info = zipfile.ZipInfo("mimetype")
            mime_info.compress_type = zipfile.ZIP_STORED
            zf.writestr(mime_info, mimetype)

            # 나머지: DEFLATED
            def _write(name, content):
                zf.writestr(name, content.encode("utf-8"),
                            compress_type=zipfile.ZIP_DEFLATED)

            _write("version.xml", version_xml)
            _write("META-INF/container.xml", container_xml)
            _write("META-INF/container.rdf", container_rdf)
            _write("META-INF/manifest.xml", manifest_xml)
            _write("settings.xml", settings_xml)
            _write("Contents/content.hpf", content_hpf)
            _write("Contents/header.xml", header_xml)
            _write("Contents/section0.xml", section_xml)


if __name__ == "__main__":
    gen = HWPXGenerator(os.getcwd())
    test_data = {
        "metadata": {"title": "테스트 문서", "include_title": True,
                      "include_section_titles": True},
        "content": [
            {
                "type": "section",
                "title": "1. 테스트 섹션",
                "items": [
                    {"level": 1, "text": "레벨 1 텍스트"},
                    {"level": 2, "text": "레벨 2 {{red:빨간색}} 텍스트"},
                    {"level": 3, "text": "레벨 3 {{green:녹색}} 텍스트"},
                    {"level": 4, "text": "레벨 4 일반 텍스트"},
                ]
            },
            {
                "type": "section",
                "title": "2. 표 테스트",
                "items": [
                    {
                        "type": "table",
                        "headers": ["항목", "내용", "비고"],
                        "rows": [
                            ["항목1", "내용1", "비고1"],
                            ["항목2", "{{red:중요}}", "비고2"],
                        ]
                    }
                ]
            }
        ]
    }
    gen.generate(test_data, "test_output.hwpx")
