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
import struct
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
            self.line_spacing = styles_data.get("lineSpacing", 160)

        # CharPr / ParaPr ID 카운터 (0번은 기본용으로 예약)
        self._charpr_list = []   # (id, height, textColor, fontId)
        self._parapr_list = []   # (id, leftMargin, spaceBefore, spaceAfter, align, indent)
        self._next_charpr_id = 1
        self._next_parapr_id = 1
        self._charpr_cache = {}  # (height, textColor, fontId) -> id
        self._parapr_cache = {}  # (leftMargin, spaceBefore, spaceAfter, align) -> id

        # 폰트 매핑: font_name -> id (0번은 기본 폰트)
        self._fonts = []         # [(id, face)] — 순서대로
        self._font_cache = {}    # face -> id

        # 이미지(BinData) 관리
        self._images = []        # [(bin_id, abs_path, zip_name, fmt)]
        self._next_bin_id = 1

    # ------------------------------------------------------------------
    # 폰트 관리
    # ------------------------------------------------------------------
    @staticmethod
    def _font_family_type(face: str) -> str:
        """폰트 이름으로 HWPX 계열 타입 반환 (바탕/명조/Serif → MYEONGJO, 나머지 → GOTHIC)"""
        myeongjo_keywords = ['바탕', '명조', 'Batang', 'Myeongjo', 'Gungsuh', 'Serif']
        if any(k in face for k in myeongjo_keywords):
            return "FCAT_MYEONGJO"
        return "FCAT_GOTHIC"

    def _register_font(self, face: str) -> int:
        if face in self._font_cache:
            return self._font_cache[face]
        fid = len(self._fonts)
        self._fonts.append((fid, face))
        self._font_cache[face] = fid
        return fid

    def _collect_fonts_from_styles(self):
        """스타일에서 사용하는 폰트를 모두 등록"""
        # 기본 폰트: 스타일에서 본문(level2)과 제목(level1) 폰트를 가져옴
        body_font = self.style_config.get("level2", {}).get("font", "Noto Serif KR")
        heading_font = self.style_config.get("level1", {}).get("font", "Noto Sans KR")
        default_fonts = [body_font, heading_font]
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
    def _get_charpr_id(self, height: int, text_color: str, font_name: str,
                       bold: bool = False) -> int:
        font_id = self._register_font(font_name)
        key = (height, text_color.upper(), font_id, bold)
        if key in self._charpr_cache:
            return self._charpr_cache[key]
        cid = self._next_charpr_id
        self._next_charpr_id += 1
        self._charpr_list.append((cid, height, text_color.upper(), font_id, bold))
        self._charpr_cache[key] = cid
        return cid

    # ------------------------------------------------------------------
    # ParaPr 관리
    # ------------------------------------------------------------------
    def _get_parapr_id(self, left_margin: int = 0, space_before: int = 0,
                       space_after: int = 0, align: str = "JUSTIFY",
                       indent: int = 0) -> int:
        key = (left_margin, space_before, space_after, align, indent)
        if key in self._parapr_cache:
            return self._parapr_cache[key]
        pid = self._next_parapr_id
        self._next_parapr_id += 1
        self._parapr_list.append((pid, left_margin, space_before, space_after, align, indent))
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
        """포인트를 HWPUNIT로 변환 (1pt = 50 HWPUNIT)"""
        return int(pt * 50)

    # ------------------------------------------------------------------
    # 이미지 관리
    # ------------------------------------------------------------------
    @staticmethod
    def _read_image_size(path: str) -> tuple:
        """PNG/JPEG 파일에서 (width, height) 픽셀을 읽는다. PIL 불필요."""
        with open(path, "rb") as f:
            header = f.read(32)
            # PNG: 시그니처 8바이트 후 IHDR 청크에 w/h
            if header[:8] == b'\x89PNG\r\n\x1a\n':
                w, h = struct.unpack('>II', header[16:24])
                return (w, h)
            # JPEG: SOI 마커 후 프레임 찾기
            if header[:2] == b'\xff\xd8':
                f.seek(0)
                f.read(2)  # SOI
                while True:
                    marker, = struct.unpack('>H', f.read(2))
                    if marker == 0xFFD9 or marker == 0xFFDA:
                        break
                    length, = struct.unpack('>H', f.read(2))
                    if marker in (0xFFC0, 0xFFC2):
                        f.read(1)  # precision
                        h, w = struct.unpack('>HH', f.read(4))
                        return (w, h)
                    f.read(length - 2)
        # 기본값 (읽기 실패 시)
        return (800, 600)

    def _resolve_image_path(self, image_path: str) -> str:
        """이미지 경로를 절대경로로 변환. 상대경로면 base_dir 기준."""
        p = Path(image_path)
        if p.is_absolute() and p.is_file():
            return str(p)
        # 1) base_dir 기준
        resolved = self.base_dir / p
        if resolved.is_file():
            return str(resolved)
        # 2) output/ 에서 실행된 경우 상위 폴더 기준
        parent_resolved = self.base_dir.parent / p
        if parent_resolved.is_file():
            return str(parent_resolved)
        # 3) 파일명만으로 images/ 폴더에서 탐색
        fname = p.name
        for search_dir in [self.base_dir, self.base_dir.parent]:
            img_dir = search_dir / "images"
            if img_dir.is_dir():
                candidate = img_dir / fname
                if candidate.is_file():
                    _log(f"[Resolve] Found image by filename: {candidate}")
                    return str(candidate)
                # 확장자 무관 탐색 (파일명 stem 매칭)
                stem = p.stem.lower()
                for f in img_dir.iterdir():
                    if f.stem.lower() == stem and f.suffix.lower() in ('.png', '.jpg', '.jpeg', '.gif'):
                        _log(f"[Resolve] Found image by stem match: {f}")
                        return str(f)
        return str(resolved)

    def _register_image(self, image_path: str) -> int:
        """이미지 파일 등록 → BinData ID 반환"""
        bid = self._next_bin_id
        self._next_bin_id += 1
        abs_path = self._resolve_image_path(image_path)
        ext = Path(abs_path).suffix.lower()
        fmt = "PNG" if ext == ".png" else "JPEG" if ext in (".jpg", ".jpeg") else "PNG"
        # 한글 오피스 방식: BinData/imageN.ext (Contents/ 접두사 없음)
        zip_name = f"BinData/image{bid}{ext}"
        self._images.append((bid, abs_path, zip_name, fmt))
        return bid

    def _build_bindatalist_xml(self) -> str:
        """한글 오피스는 header.xml에 binDataList를 넣지 않음 → 빈 문자열"""
        return ""

    def _image_paragraph_xml(self, bin_id: int, width_px: int, height_px: int,
                              max_width_hwpunit: int = 47600) -> str:
        """이미지를 담는 paragraph XML 생성 (한글 오피스 호환 구조).
        실제 한글 오피스가 생성하는 HWPX 구조를 그대로 복제.
        """
        # px → HWPUNIT (75 HWPUNIT/px — 한글 오피스 기준)
        PX_TO_HU = 75
        org_w = int(width_px * PX_TO_HU)
        org_h = int(height_px * PX_TO_HU)

        # 본문 폭 초과 시 비율 축소
        if org_w > max_width_hwpunit:
            ratio = max_width_hwpunit / org_w
            cur_w = max_width_hwpunit
            cur_h = int(org_h * ratio)
        else:
            cur_w = org_w
            cur_h = org_h

        scale_x = cur_w / org_w if org_w else 1
        scale_y = cur_h / org_h if org_h else 1
        center_x = cur_w // 2
        center_y = cur_h // 2

        img_id = f"image{bin_id}"
        return (
            f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'<hp:pic id="{2107322690 + bin_id}" zOrder="0"'
            f' numberingType="PICTURE" textWrap="TOP_AND_BOTTOM"'
            f' textFlow="BOTH_SIDES" lock="0" dropcapstyle="None"'
            f' href="" groupLevel="0" instid="{1033580867 + bin_id}"'
            f' reverse="0">'
            f'<hp:offset x="0" y="0"/>'
            f'<hp:orgSz width="{org_w}" height="{org_h}"/>'
            f'<hp:curSz width="{cur_w}" height="{cur_h}"/>'
            f'<hp:flip horizontal="0" vertical="0"/>'
            f'<hp:rotationInfo angle="0" centerX="{center_x}"'
            f' centerY="{center_y}" rotateimage="1"/>'
            f'<hp:renderingInfo>'
            f'<hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
            f'<hc:scaMatrix e1="{scale_x:.6f}" e2="0" e3="0"'
            f' e4="0" e5="{scale_y:.6f}" e6="0"/>'
            f'<hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>'
            f'</hp:renderingInfo>'
            f'<hp:imgRect>'
            f'<hc:pt0 x="0" y="0"/>'
            f'<hc:pt1 x="{org_w}" y="0"/>'
            f'<hc:pt2 x="{org_w}" y="{org_h}"/>'
            f'<hc:pt3 x="0" y="{org_h}"/>'
            f'</hp:imgRect>'
            f'<hp:imgClip left="0" right="{org_w}" top="0" bottom="{org_h}"/>'
            f'<hp:inMargin left="0" right="0" top="0" bottom="0"/>'
            f'<hc:img binaryItemIDRef="{img_id}" bright="0" contrast="0"'
            f' effect="REAL_PIC" alpha="0"/>'
            f'<hp:effects/>'
            f'<hp:sz width="{cur_w}" widthRelTo="ABSOLUTE"'
            f' height="{cur_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
            f' allowOverlap="0" holdAnchorAndSO="0"'
            f' vertRelTo="PARA" horzRelTo="PARA"'
            f' vertAlign="TOP" horzAlign="LEFT"'
            f' vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="0" right="0" top="0" bottom="0"/>'
            f'</hp:pic>'
            f'</hp:run>'
            f'</hp:p>'
        )

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
            family_type = self._font_family_type(face)
            font_elems += (
                f'<hh:font id="{fid}" face="{xml_escape(face)}" type="TTF" isEmbedded="0">'
                f'<hh:typeInfo familyType="{family_type}" weight="6" proportion="4"'
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

    def _build_charpr_xml(self, cid: int, height: int, text_color: str,
                         font_id: int, bold: bool = False) -> str:
        fid = str(font_id)
        bold_attr = ' bold="1"' if bold else ''
        return (
            f'<hh:charPr id="{cid}" height="{height}" textColor="{text_color}"'
            f' shadeColor="none" useFontSpace="0" useKerning="0"'
            f' symMark="NONE" borderFillIDRef="2"{bold_attr}>'
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

    def _build_parapr_xml(self, pid, left_margin, space_before, space_after, align, indent=0) -> str:
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="0" snapToGrid="1"'
            f' suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="KEEP_WORD"'
            f' breakNonLatinWord="KEEP_WORD" widowOrphan="0"'
            f' keepWithNext="0" keepLines="1" pageBreakBefore="0"'
            f' lineWrap="BREAK"/>'
            f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            f'<hh:margin>'
            f'<hc:intent value="{indent}" unit="HWPUNIT"/>'
            f'<hc:left value="{left_margin}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="{space_before}" unit="HWPUNIT"/>'
            f'<hc:next value="{space_after}" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{self.line_spacing}" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            f' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            f'</hh:paraPr>'
        )

    def _build_table_parapr_xml(self, pid, align="LEFT", left_margin=0, indent=0, line_spacing=130) -> str:
        """표 셀 전용 paragraph style — 글자 단위 줄바꿈 허용, 글꼴에 어울리는 줄 높이"""
        return (
            f'<hh:paraPr id="{pid}" tabPrIDRef="0" condense="0"'
            f' fontLineHeight="1" snapToGrid="1"'
            f' suppressLineNumbers="0" checked="0">'
            f'<hh:align horizontal="{align}" vertical="BASELINE"/>'
            f'<hh:heading type="NONE" idRef="0" level="0"/>'
            f'<hh:breakSetting breakLatinWord="HYPHENATION"'
            f' breakNonLatinWord="BREAK_ALL" widowOrphan="0"'
            f' keepWithNext="0" keepLines="1" pageBreakBefore="0"'
            f' lineWrap="BREAK"/>'
            f'<hh:autoSpacing eAsianEng="0" eAsianNum="0"/>'
            f'<hh:margin>'
            f'<hc:intent value="{indent}" unit="HWPUNIT"/>'
            f'<hc:left value="{left_margin}" unit="HWPUNIT"/>'
            f'<hc:right value="0" unit="HWPUNIT"/>'
            f'<hc:prev value="0" unit="HWPUNIT"/>'
            f'<hc:next value="0" unit="HWPUNIT"/>'
            f'</hh:margin>'
            f'<hh:lineSpacing type="PERCENT" value="{line_spacing}" unit="HWPUNIT"/>'
            f'<hh:border borderFillIDRef="2" offsetLeft="0" offsetRight="0"'
            f' offsetTop="0" offsetBottom="0" connect="0" ignoreMargin="0"/>'
            f'</hh:paraPr>'
        )

    def _get_table_parapr_id(self, is_header: bool = False) -> int:
        """표 셀 전용 parapr ID — JSON 스타일에서 설정 가져오기"""
        cache_key = '_table_parapr_header' if is_header else '_table_parapr_data'
        if not hasattr(self, cache_key):
            pid = self._next_parapr_id
            self._next_parapr_id += 1
            if not hasattr(self, '_table_parapr_list'):
                self._table_parapr_list = []

            # JSON에서 스타일 가져오기
            style_key = "table_header" if is_header else "table_data"
            style = self.style_config.get(style_key, {})
            align = style.get("align", "left").upper()
            left_margin = self._pt_to_hwpunit(style.get("leftMargin", 0))
            indent = 0  # 표 셀은 항상 "보통" (첫줄 들여쓰기/내어쓰기 없음)
            line_spacing = style.get("lineSpacing", 130)

            self._table_parapr_list.append((pid, align, left_margin, indent, line_spacing))
            setattr(self, cache_key, pid)
        return getattr(self, cache_key)

    @staticmethod
    def _ensure_str(cell) -> str:
        """셀 데이터를 문자열로 보장합니다. (하위호환: dict → 인라인 마커 복원)"""
        if isinstance(cell, dict):
            text = cell.get("text", "")
            color = cell.get("color", "")
            if color and "{{" not in text:
                return f"{{{{{color}:{text}}}}}"
            return text
        return str(cell)

    def _build_borderfills_xml(self) -> str:
        """기본 borderFill + 표용 borderFill 생성 (한글 오피스 호환)"""
        # ID 1: 페이지 테두리용 (투명)
        bfs = (
            '<hh:borderFill id="1" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
        )
        # ID 2: 기본 문단 border용 (투명, fillBrush 포함)
        bfs += (
            '<hh:borderFill id="2" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:rightBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:topBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:bottomBorder type="NONE" width="0.1 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '<hc:fillBrush>'
            '<hc:winBrush faceColor="none" hatchColor="#999999" alpha="0"/>'
            '</hc:fillBrush>'
            '</hh:borderFill>'
        )
        # ID 3: 표 본문 셀용 (SOLID 테두리)
        bfs += (
            '<hh:borderFill id="3" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '</hh:borderFill>'
        )
        # ID 4: 표 헤더 셀용 (SOLID 테두리 + 회색 배경)
        bfs += (
            '<hh:borderFill id="4" threeD="0" shadow="0"'
            ' centerLine="NONE" breakCellSeparateLine="0">'
            '<hh:slash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:backSlash type="NONE" Crooked="0" isCounter="0"/>'
            '<hh:leftBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:rightBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:topBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:bottomBorder type="SOLID" width="0.12 mm" color="#000000"/>'
            '<hh:diagonal type="SOLID" width="0.1 mm" color="#000000"/>'
            '<hc:fillBrush>'
            '<hc:winBrush faceColor="#D9D9D9" hatchColor="#999999" alpha="0"/>'
            '</hc:fillBrush>'
            '</hh:borderFill>'
        )
        return f'<hh:borderFills itemCnt="4">{bfs}</hh:borderFills>'

    def _build_tab_properties_xml(self) -> str:
        """tabProperties XML — 한글 오피스 필수 요소"""
        return (
            '<hh:tabProperties itemCnt="3">'
            '<hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>'
            '<hh:tabPr id="1" autoTabLeft="1" autoTabRight="0"/>'
            '<hh:tabPr id="2" autoTabLeft="0" autoTabRight="1"/>'
            '</hh:tabProperties>'
        )

    def _build_numberings_xml(self) -> str:
        """numberings XML — 한글 오피스 필수 요소"""
        return (
            '<hh:numberings itemCnt="1">'
            '<hh:numbering id="1" start="0">'
            '<hh:paraHead start="1" level="1" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295"'
            ' checkable="0">^1.</hh:paraHead>'
            '<hh:paraHead start="1" level="2" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295"'
            ' checkable="0">^2.</hh:paraHead>'
            '<hh:paraHead start="1" level="3" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295"'
            ' checkable="0">^3)</hh:paraHead>'
            '<hh:paraHead start="1" level="4" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295"'
            ' checkable="0">^4)</hh:paraHead>'
            '<hh:paraHead start="1" level="5" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="DIGIT" charPrIDRef="4294967295"'
            ' checkable="0">(^5)</hh:paraHead>'
            '<hh:paraHead start="1" level="6" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="HANGUL_SYLLABLE" charPrIDRef="4294967295"'
            ' checkable="0">(^6)</hh:paraHead>'
            '<hh:paraHead start="1" level="7" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="CIRCLED_DIGIT" charPrIDRef="4294967295"'
            ' checkable="1">^7</hh:paraHead>'
            '<hh:paraHead start="1" level="8" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="CIRCLED_HANGUL_SYLLABLE" charPrIDRef="4294967295"'
            ' checkable="1">^8</hh:paraHead>'
            '<hh:paraHead start="1" level="9" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="HANGUL_JAMO" charPrIDRef="4294967295"'
            ' checkable="0"/>'
            '<hh:paraHead start="1" level="10" align="LEFT" useInstWidth="1"'
            ' autoIndent="1" widthAdjust="0" textOffsetType="PERCENT"'
            ' textOffset="50" numFormat="ROMAN_SMALL" charPrIDRef="4294967295"'
            ' checkable="1"/>'
            '</hh:numbering>'
            '</hh:numberings>'
        )

    def _build_styles_xml(self) -> str:
        """styles XML — 한글 오피스 필수 요소 (기본 스타일 22개)"""
        return (
            '<hh:styles itemCnt="22">'
            '<hh:style id="0" type="PARA" name="바탕글" engName="Normal"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
            '<hh:style id="1" type="PARA" name="본문" engName="Body"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="1" langID="1042" lockForm="0"/>'
            '<hh:style id="2" type="PARA" name="개요 1" engName="Outline 1"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="2" langID="1042" lockForm="0"/>'
            '<hh:style id="3" type="PARA" name="개요 2" engName="Outline 2"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="3" langID="1042" lockForm="0"/>'
            '<hh:style id="4" type="PARA" name="개요 3" engName="Outline 3"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="4" langID="1042" lockForm="0"/>'
            '<hh:style id="5" type="PARA" name="개요 4" engName="Outline 4"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="5" langID="1042" lockForm="0"/>'
            '<hh:style id="6" type="PARA" name="개요 5" engName="Outline 5"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="6" langID="1042" lockForm="0"/>'
            '<hh:style id="7" type="PARA" name="개요 6" engName="Outline 6"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="7" langID="1042" lockForm="0"/>'
            '<hh:style id="8" type="PARA" name="개요 7" engName="Outline 7"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="8" langID="1042" lockForm="0"/>'
            '<hh:style id="9" type="PARA" name="개요 8" engName="Outline 8"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="9" langID="1042" lockForm="0"/>'
            '<hh:style id="10" type="PARA" name="개요 9" engName="Outline 9"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="10" langID="1042" lockForm="0"/>'
            '<hh:style id="11" type="PARA" name="개요 10" engName="Outline 10"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="11" langID="1042" lockForm="0"/>'
            '<hh:style id="12" type="CHAR" name="쪽 번호" engName="Page Number"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langID="1042" lockForm="0"/>'
            '<hh:style id="13" type="PARA" name="머리말" engName="Header"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="13" langID="1042" lockForm="0"/>'
            '<hh:style id="14" type="PARA" name="각주" engName="Footnote"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="14" langID="1042" lockForm="0"/>'
            '<hh:style id="15" type="PARA" name="미주" engName="Endnote"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="15" langID="1042" lockForm="0"/>'
            '<hh:style id="16" type="PARA" name="메모" engName="Memo"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="16" langID="1042" lockForm="0"/>'
            '<hh:style id="17" type="PARA" name="차례 제목" engName="TOC Heading"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="17" langID="1042" lockForm="0"/>'
            '<hh:style id="18" type="PARA" name="차례 1" engName="TOC 1"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="18" langID="1042" lockForm="0"/>'
            '<hh:style id="19" type="PARA" name="차례 2" engName="TOC 2"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="19" langID="1042" lockForm="0"/>'
            '<hh:style id="20" type="PARA" name="차례 3" engName="TOC 3"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="20" langID="1042" lockForm="0"/>'
            '<hh:style id="21" type="PARA" name="캡션" engName="Caption"'
            ' paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="21" langID="1042" lockForm="0"/>'
            '</hh:styles>'
        )

    def _build_header_xml(self) -> str:
        """header.xml 전체 생성 (한글 오피스 호환)"""
        fontfaces = self._build_fontfaces_xml()

        # CharPr: id=0 (기본) + 동적 생성분
        charpr_default = self._build_charpr_xml(0, 1000, "#000000", 0)
        charprs = charpr_default
        for cid, height, color, fid, bold in self._charpr_list:
            charprs += self._build_charpr_xml(cid, height, color, fid, bold)
        charpr_cnt = 1 + len(self._charpr_list)

        # ParaPr: id=0 (기본) + 동적 생성분 + 표 셀 전용
        parapr_default = self._build_parapr_xml(0, 0, 0, 0, "JUSTIFY")
        paraprs = parapr_default
        for pid, lm, sb, sa, align, indent in self._parapr_list:
            paraprs += self._build_parapr_xml(pid, lm, sb, sa, align, indent)
        # 표 셀 전용 parapr 추가
        table_parapr_extra = 0
        if hasattr(self, '_table_parapr_list'):
            for tpid, align, left_margin, indent, line_spacing in self._table_parapr_list:
                paraprs += self._build_table_parapr_xml(tpid, align, left_margin, indent, line_spacing)
                table_parapr_extra += 1
        parapr_cnt = 1 + len(self._parapr_list) + table_parapr_extra

        borderfills = self._build_borderfills_xml()
        tab_properties = self._build_tab_properties_xml()
        numberings = self._build_numberings_xml()
        styles = self._build_styles_xml()

        return (
            f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            f'<hh:head {_ALL_NS} version="1.2" secCnt="1">'
            f'<hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>'
            f'<hh:refList>'
            f'{fontfaces}'
            f'{borderfills}'
            f'<hh:charProperties itemCnt="{charpr_cnt}">{charprs}</hh:charProperties>'
            f'{tab_properties}'
            f'{numberings}'
            f'<hh:paraProperties itemCnt="{parapr_cnt}">{paraprs}</hh:paraProperties>'
            f'{styles}'
            f'</hh:refList>'
            f'<hh:compatibleDocument targetProgram="HWP201X">'
            f'<hh:layoutCompatibility/>'
            f'</hh:compatibleDocument>'
            f'<hh:docOption>'
            f'<hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/>'
            f'</hh:docOption>'
            f'<hh:trackchageConfig flags="56"/>'
            f'</hh:head>'
        )

    def _build_secpr_xml(self) -> str:
        """페이지 설정 (secPr) — A4, 상하좌우 여백 (한글 오피스 호환)"""
        return (
            '<hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134"'
            ' tabStop="8000" outlineShapeIDRef="1" memoShapeIDRef="0"'
            ' textVerticalWidthHead="0" masterPageCnt="0">'
            '<hp:grid lineGrid="0" charGrid="0" wonggojiFormat="0"/>'
            '<hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/>'
            '<hp:visibility hideFirstHeader="0" hideFirstFooter="0"'
            ' hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL"'
            ' hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/>'
            '<hp:lineNumberShape restartType="0" countBy="0" distance="0" startNumber="0"/>'
            '<hp:pagePr landscape="WIDELY" width="59528" height="84186" gutterType="LEFT_ONLY">'
            '<hp:margin header="4252" footer="4252" gutter="0"'
            ' left="8504" right="8504" top="5668" bottom="4252"/>'
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
                        align: str = "JUSTIFY",
                        hanging_indent_pt: float = 0) -> str:
        """마커 색상을 지원하는 텍스트 paragraph 생성"""
        height = self._pt_to_height(font_size_pt)
        # 한글 오피스 내어쓰기: left = leftMargin + hangingIndent (둘째 줄 위치)
        # intent = -hangingIndent (첫째 줄은 기호 시작점)
        actual_left = self._pt_to_hwpunit(left_margin_pt + hanging_indent_pt)
        indent_val = -self._pt_to_hwpunit(hanging_indent_pt) if hanging_indent_pt else 0
        parapr_id = self._get_parapr_id(
            actual_left,
            self._pt_to_hwpunit(space_before_pt),
            self._pt_to_hwpunit(space_after_pt),
            align,
            indent_val,
        )

        segments = self._parse_markers(text)
        runs = ""
        for seg_text, seg_marker in segments:
            is_bold = (seg_marker == "bold")
            if seg_marker and seg_marker != "bold":
                color_hex = self._resolve_color(seg_marker)
            else:
                color_hex = "#000000"
            cid = self._get_charpr_id(height, color_hex, font_name, bold=is_bold)
            runs += self._run_xml(seg_text, cid)

        return self._paragraph_xml(runs, parapr_id)

    # ------------------------------------------------------------------
    # 표 XML 빌더
    # ------------------------------------------------------------------
    def _table_cell_xml(self, text: str, charpr_id: int, col_idx: int,
                        row_idx: int, cell_width: int,
                        is_header: bool = False) -> str:
        """표 셀 XML — 글자 단위 줄바꿈 지원"""
        segments = self._parse_markers(text)

        # 셀 내부 run 들
        cell_runs = ""
        # 표는 항상 Noto Sans KR 9pt 강제 적용
        table_font = "Noto Sans KR"
        table_size = 9
        for seg_text, seg_marker in segments:
            is_bold = (seg_marker == "bold")
            if seg_marker and seg_marker != "bold":
                color_hex = self._resolve_color(seg_marker)
            else:
                color_hex = "#000000"
            cid = self._get_charpr_id(
                self._pt_to_height(table_size), color_hex, table_font, bold=is_bold)
            cell_runs += self._run_xml(seg_text, cid)

        inner_width = max(cell_width - 1020, 1000)
        # 표 셀 전용 parapr (헤더=CENTER, 데이터=LEFT)
        tbl_parapr = self._get_table_parapr_id(is_header=is_header)
        # 헤더 셀: borderFillIDRef=4 (회색 배경), 본문 셀: 3 (흰 배경)
        bf_id = "4" if is_header else "3"

        # 폰트 크기에 맞는 줄 높이 계산 (1pt = 100 HWPUNIT)
        # 9pt 폰트 → 900 높이, 10pt → 1000, 11pt → 1100
        font_height = int(table_size * 100)
        # baseline은 높이의 약 85%, spacing은 높이의 약 60%
        baseline = int(font_height * 0.85)
        spacing = int(font_height * 0.6)

        # lineSpacing 값 적용: 180% → vertsize = font_height * 1.8
        style_key = "table_header" if is_header else "table_data"
        line_spacing_pct = self.style_config.get(style_key, {}).get("lineSpacing", 180)
        vert_size = int(font_height * line_spacing_pct / 100)

        # 셀 높이는 vert_size + 여백(상하 141*2)
        cell_height = vert_size + 282

        return (
            f'<hp:tc name="" header="0" hasMargin="0" protect="0"'
            f' editable="0" dirty="0" borderFillIDRef="{bf_id}">'
            f'<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK"'
            f' vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0"'
            f' textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">'
            f'<hp:p id="0" paraPrIDRef="{tbl_parapr}" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'{cell_runs}'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="{vert_size}"'
            f' textheight="{font_height}" baseline="{baseline}" spacing="{spacing}"'
            f' horzpos="0" horzsize="{inner_width}" flags="393216"/>'
            f'</hp:linesegarray>'
            f'</hp:p>'
            f'</hp:subList>'
            f'<hp:cellAddr colAddr="{col_idx}" rowAddr="{row_idx}"/>'
            f'<hp:cellSpan colSpan="1" rowSpan="1"/>'
            f'<hp:cellSz width="{cell_width}" height="{cell_height}"/>'
            f'<hp:cellMargin left="170" right="170" top="141" bottom="141"/>'
            f'</hp:tc>'
        )

    def _table_xml(self, table_data: dict) -> str:
        """표 XML 생성 (한글 오피스 호환 구조)"""
        headers = table_data.get("headers", [])
        rows = table_data.get("rows", [])
        col_count = len(headers) if headers else (len(rows[0]) if rows else 1)
        # 빈 헤더 감지 (| | | 같은 경우)
        has_visible_header = any(
            (self._ensure_str(h).strip() if isinstance(h, (dict, str)) else str(h).strip())
            for h in headers
        ) if headers else False
        row_count = len(rows) + (1 if has_visible_header else 0)

        # 표 글자 스타일
        table_style = self.style_config.get("table", {})
        table_font = table_style.get("font", "Noto Serif KR")
        table_size = table_style.get("size", 10)
        table_height = self._pt_to_height(table_size)
        table_charpr = self._get_charpr_id(table_height, "#000000", table_font)

        # 본문 폭: 페이지폭(59528) - 좌여백(8504) - 우여백(8504) = 42520
        total_width = 42520
        cell_width = total_width // col_count

        # 헤더 행 (내용이 있는 경우만)
        header_row = ""
        if has_visible_header:
            for ci, h in enumerate(headers):
                h_text = self._ensure_str(h)
                header_row += self._table_cell_xml(
                    h_text, table_charpr, ci, 0, cell_width, is_header=True)
            header_row = f'<hp:tr>{header_row}</hp:tr>'

        # 데이터 행
        data_rows = ""
        for ri, row in enumerate(rows):
            cells = ""
            row_idx = ri + (1 if has_visible_header else 0)
            for ci, cell in enumerate(row):
                c_text = self._ensure_str(cell)
                cells += self._table_cell_xml(c_text, table_charpr, ci, row_idx, cell_width)
            data_rows += f'<hp:tr>{cells}</hp:tr>'

        # 표 전체 높이: 각 행의 셀 높이 합계
        # 셀 높이는 _table_cell_xml에서 계산한 값과 동일하게 계산
        table_style = self.style_config.get("table", {})
        table_size = table_style.get("size", 10)
        font_height = int(table_size * 100)

        # 헤더와 데이터 행의 높이 계산
        header_line_spacing = self.style_config.get("table_header", {}).get("lineSpacing", 180)
        data_line_spacing = self.style_config.get("table_data", {}).get("lineSpacing", 180)

        header_vert_size = int(font_height * header_line_spacing / 100)
        data_vert_size = int(font_height * data_line_spacing / 100)

        header_cell_height = header_vert_size + 282
        data_cell_height = data_vert_size + 282

        # 표 전체 높이 = (헤더 있으면 헤더 높이) + (데이터 행 수 * 데이터 행 높이)
        if has_visible_header:
            table_h = header_cell_height + (data_cell_height * len(rows))
        else:
            table_h = data_cell_height * len(rows)

        return (
            f'<hp:tbl id="0" zOrder="0" numberingType="TABLE"'
            f' textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0"'
            f' dropcapstyle="None" pageBreak="CELL" repeatHeader="1"'
            f' rowCnt="{row_count}" colCnt="{col_count}"'
            f' cellSpacing="0" borderFillIDRef="3" noAdjust="0">'
            f'<hp:sz width="{total_width}" widthRelTo="ABSOLUTE"'
            f' height="{table_h}" heightRelTo="ABSOLUTE" protect="0"/>'
            f'<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1"'
            f' allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA"'
            f' horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT"'
            f' vertOffset="0" horzOffset="0"/>'
            f'<hp:outMargin left="283" right="283" top="283" bottom="283"/>'
            f'<hp:inMargin left="170" right="170" top="141" bottom="141"/>'
            f'{header_row}{data_rows}'
            f'</hp:tbl>'
        )

    def _table_paragraph_xml(self, table_data: dict) -> str:
        """표를 담는 paragraph XML (캡션 포함)"""
        caption_xml = ""
        caption = table_data.get("title", "")
        if caption:
            # 표 제목: 돋움체 가운데 정렬 (마커 파싱 포함)
            cap_style = self.style_config.get("table_caption", {})
            cap_font = cap_style.get("font", "Noto Sans KR")
            cap_size = cap_style.get("size", 11)
            cap_height = self._pt_to_height(cap_size)
            cap_parapr = self._get_parapr_id(
                0,
                self._pt_to_hwpunit(cap_style.get("paragraphSpaceBefore", 10)),
                self._pt_to_hwpunit(cap_style.get("paragraphSpaceAfter", 3)),
                "CENTER",
            )
            segments = self._parse_markers(caption)
            cap_runs = ""
            for seg_text, seg_color in segments:
                if seg_color:
                    color_hex = self._resolve_color(seg_color)
                else:
                    color_hex = "#000000"
                cid = self._get_charpr_id(cap_height, color_hex, cap_font, bold=True)
                cap_runs += self._run_xml(seg_text, cid)
            caption_xml = self._paragraph_xml(cap_runs, cap_parapr)

        tbl = self._table_xml(table_data)

        return (
            caption_xml +
            f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            f' pageBreak="0" columnBreak="0" merged="0">'
            f'<hp:run charPrIDRef="0">'
            f'{tbl}'
            f'<hp:t/>'
            f'</hp:run>'
            f'<hp:linesegarray>'
            f'<hp:lineseg textpos="0" vertpos="0" vertsize="1000"'
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
        # colPr: 단 설정 (한글 오피스 필수)
        colpr = (
            '<hp:ctrl>'
            '<hp:colPr id="" type="NEWSPAPER" layout="LEFT"'
            ' colCount="1" sameSz="1" sameGap="0"/>'
            '</hp:ctrl>'
        )

        if include_title and title:
            title_style = self.style_config.get("title", {})
            title_font = title_style.get("font", "Noto Serif KR")
            title_size = title_style.get("size", 25)
            title_align = title_style.get("align", "center").upper()
            if title_align == "CENTER":
                title_align = "CENTER"
            title_height = self._pt_to_height(title_size)
            title_charpr = self._get_charpr_id(title_height, "#000000", title_font)
            title_parapr = self._get_parapr_id(0, 0, self._pt_to_hwpunit(10), title_align)

            runs = self._run_xml(title, title_charpr)
            # 첫 문단에 secPr + colPr 포함
            body_paragraphs += (
                f'<hp:p id="0" paraPrIDRef="{title_parapr}"'
                f' styleIDRef="0" pageBreak="0"'
                f' columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}{colpr}</hp:run>'
                f'{runs}'
                f'</hp:p>'
            )
        else:
            # 제목 없으면 빈 첫 문단에 secPr + colPr 포함
            body_paragraphs += (
                f'<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
                f' pageBreak="0" columnBreak="0" merged="0">'
                f'<hp:run charPrIDRef="0">{secpr}{colpr}</hp:run>'
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
                    sec_font = sec_style.get("font", "Noto Serif KR")
                    sec_size = sec_style.get("size", 18)
                    sec_height = self._pt_to_height(sec_size)
                    sec_charpr = self._get_charpr_id(sec_height, "#000000", sec_font)
                    sec_parapr = self._get_parapr_id(
                        self._pt_to_hwpunit(sec_style.get("leftMargin", 0)),
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
                    elif sub_type == "image":
                        img_path = sub.get("path", "")
                        resolved = self._resolve_image_path(img_path) if img_path else ""
                        if resolved and os.path.isfile(resolved):
                            bid = self._register_image(img_path)
                            w_px, h_px = self._read_image_size(resolved)
                            body_paragraphs += self._image_paragraph_xml(bid, w_px, h_px)
                            _log(f"[Added] Image: {resolved} ({w_px}x{h_px}px)")
                        else:
                            _log(f"[Warning] Image not found: {img_path}")
                    elif sub_type == "subtitle":
                        # subtitle_level: 1=## (큰 제목), 2=### (소제목)
                        subtitle_level = sub.get("subtitle_level", 2)
                        if subtitle_level == 1:
                            sub_style = self.style_config.get("level1", {})
                        else:
                            sub_style = self.style_config.get("section_subtitle", {})
                        sub_font = sub_style.get("font", "Noto Sans KR")
                        sub_size = sub_style.get("size", 15)
                        sub_bold = sub_style.get("bold", True)
                        sub_charpr = self._get_charpr_id(
                            self._pt_to_height(sub_size), "#000000", sub_font,
                            bold=sub_bold
                        )
                        sub_parapr = self._get_parapr_id(
                            self._pt_to_hwpunit(sub_style.get("leftMargin", 0)),
                            self._pt_to_hwpunit(sub_style.get("paragraphSpaceBefore", 15)),
                            self._pt_to_hwpunit(sub_style.get("paragraphSpaceAfter", 6)),
                            "JUSTIFY",
                        )
                        runs = self._run_xml(sub.get("text", ""), sub_charpr)
                        body_paragraphs += self._paragraph_xml(runs, sub_parapr)
                        _log(f"[Added] Subtitle: {sub.get('text', '')[:50]}")
                    else:
                        level = sub.get("level", 1)
                        text = sub.get("text", "")
                        level_key = f"level{level}"
                        style = self.style_config.get(level_key, {})

                        # 기호(symbol) 적용: 텍스트가 이미 기호로 시작하지 않으면 추가
                        symbol = style.get("symbol", "")
                        if symbol and text and text.startswith(symbol):
                            # 텍스트가 이미 해당 기호로 시작하면 그대로 유지
                            display_text = text
                        else:
                            # 기호가 없거나 텍스트가 다른 문자로 시작하면 기호 추가
                            display_text = f"{symbol} {text}" if symbol else text

                        # ● 항목은 bullet 스타일 적용 (● 유지, 제목체)
                        if text and text.startswith('●'):
                            bullet_style = self.style_config.get("bullet", {})
                            use_style = bullet_style
                            # display_text는 dedup에서 이미 ● 포함 상태 유지
                        else:
                            use_style = style

                        body_paragraphs += self._text_paragraph(
                            display_text,
                            level,
                            use_style.get("font", "Noto Serif KR"),
                            use_style.get("size", 11),
                            use_style.get("leftMargin", 0),
                            use_style.get("paragraphSpaceBefore", 0),
                            use_style.get("paragraphSpaceAfter", 3),
                            use_style.get("align", "justify").upper(),
                            use_style.get("hangingIndent", 0),
                        )
                        _log(f"[Added] Level {level}: {text[:50]}...")

            elif item_type == "table":
                body_paragraphs += self._table_paragraph_xml(item)

            elif item_type == "image":
                img_path = item.get("path", "")
                resolved = self._resolve_image_path(img_path) if img_path else ""
                if resolved and os.path.isfile(resolved):
                    bid = self._register_image(img_path)
                    w_px, h_px = self._read_image_size(resolved)
                    body_paragraphs += self._image_paragraph_xml(bid, w_px, h_px)
                    _log(f"[Added] Image: {resolved} ({w_px}x{h_px}px)")
                    # 이미지 캡션
                    caption = item.get("caption", "")
                    if caption:
                        cap_style = self.style_config.get("image_caption", {})
                        cap_font = cap_style.get("font", "Noto Sans KR")
                        cap_size = cap_style.get("size", 11)
                        cap_height = self._pt_to_height(cap_size)
                        cap_parapr = self._get_parapr_id(
                            0,
                            self._pt_to_hwpunit(cap_style.get("paragraphSpaceBefore", 3)),
                            self._pt_to_hwpunit(cap_style.get("paragraphSpaceAfter", 10)),
                            "CENTER",
                        )
                        segments = self._parse_markers(caption)
                        cap_runs = ""
                        for seg_text, seg_color in segments:
                            color_hex = self._resolve_color(seg_color) if seg_color else "#000000"
                            cid = self._get_charpr_id(cap_height, color_hex, cap_font)
                            cap_runs += self._run_xml(seg_text, cid)
                        body_paragraphs += self._paragraph_xml(cap_runs, cap_parapr)
                else:
                    _log(f"[Warning] Image not found: {img_path} (resolved: {resolved})")

        # 마지막 빈 문단 (한글 오피스 호환 — linesegarray 포함)
        body_paragraphs += (
            '<hp:p id="0" paraPrIDRef="0" styleIDRef="0"'
            ' pageBreak="0" columnBreak="0" merged="0">'
            '<hp:run charPrIDRef="0"/>'
            '<hp:linesegarray>'
            '<hp:lineseg textpos="0" vertpos="0" vertsize="1000"'
            ' textheight="1000" baseline="850" spacing="600"'
            ' horzpos="0" horzsize="42520" flags="393216"/>'
            '</hp:linesegarray>'
            '</hp:p>'
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
            ' tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="0"'
            ' buildNumber="1" os="1" xmlVersion="1.2"'
            ' application="Hancom Office Hangul" appVersion="11, 0, 0, 2129 WIN32LEWindows_8"/>'
        )

        container_xml = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
            '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container"'
            ' xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">'
            '<ocf:rootfiles>'
            '<ocf:rootfile full-path="Contents/content.hpf"'
            ' media-type="application/hwpml-package+xml"/>'
            '<ocf:rootfile full-path="Preview/PrvText.txt"'
            ' media-type="text/plain"/>'
            '<ocf:rootfile full-path="META-INF/container.rdf"'
            ' media-type="application/rdf+xml"/>'
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

        # 한글 오피스는 manifest.xml을 빈 상태로 둠
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

        # content.hpf: 한글 방식 — 이미지는 id="imageN", isEmbeded="1"
        bindata_manifest = ""
        for bid, _, zip_name, fmt in self._images:
            media = "image/png" if fmt == "PNG" else "image/jpeg"
            bindata_manifest += (
                f'<opf:item id="image{bid}" href="{zip_name}"'
                f' media-type="{media}" isEmbeded="1"/>'
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
            f'{bindata_manifest}'
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

            # Preview 파일 (한글 오피스 호환 필수)
            _write("Preview/PrvText.txt", "")

            # 빈 1x1 PNG 이미지 (PrvImage.png)
            import base64
            _EMPTY_PNG = base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4"
                "nGNgYPgPAAEDAQAIicLsAAAABElFTkSuQmCC"
            )
            prv_info = zipfile.ZipInfo("Preview/PrvImage.png")
            prv_info.compress_type = zipfile.ZIP_DEFLATED
            zf.writestr(prv_info, _EMPTY_PNG)

            # 이미지 파일을 ZIP에 추가
            # 한글 오피스 방식: BinData/imageN.png (Contents/ 접두사 없음)
            for bid, abs_path, zip_name, fmt in self._images:
                with open(abs_path, "rb") as img_f:
                    zf.writestr(zip_name, img_f.read(),
                                compress_type=zipfile.ZIP_DEFLATED)


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
