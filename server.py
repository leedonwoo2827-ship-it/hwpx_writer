# -*- coding: utf-8 -*-
"""
HWPX Writer MCP Server

Claude Desktop에서 사용하는 MCP 서버.
마크다운 텍스트 또는 MD 파일을 proposal-styles.json 스타일이 적용된 HWPX 한글 문서로 변환합니다.

실행: python server.py
"""

import json
import os
import sys
from pathlib import Path

BASE_DIR = Path(__file__).parent
sys.path.insert(0, str(BASE_DIR / "src"))

from mcp.server.fastmcp import FastMCP
from md_parser import parse_markdown_to_json
from hwpx_generator import HWPXGenerator

DEFAULT_STYLES_PATH = BASE_DIR / "proposal-styles.json"

mcp = FastMCP("hwpx-writer")


# ---------------------------------------------------------------------------
# 내부 유틸리티
# ---------------------------------------------------------------------------

def _resolve_styles_path(styles_file: str) -> Path:
    if not styles_file:
        return DEFAULT_STYLES_PATH
    p = Path(styles_file)
    return p if p.is_absolute() else BASE_DIR / p


def _build_generator(styles_path: Path) -> HWPXGenerator:
    return HWPXGenerator(base_dir=str(BASE_DIR), styles_path=str(styles_path))


def _run_generator(generator: HWPXGenerator, data: dict, output_path: Path) -> None:
    generator.generate(data, str(output_path))


def _read_md_file(md_path: Path) -> str:
    """MD 파일을 바이트로 한번만 읽고 인코딩을 자동 감지하여 디코딩합니다."""
    raw = md_path.read_bytes()
    # BOM 감지 → UTF-8-SIG
    if raw.startswith(b"\xef\xbb\xbf"):
        return raw.decode("utf-8-sig")
    # UTF-8 시도 → 실패 시 CP949 (한글 Windows)
    try:
        return raw.decode("utf-8")
    except UnicodeDecodeError:
        return raw.decode("cp949")


# ---------------------------------------------------------------------------
# MCP 도구
# ---------------------------------------------------------------------------

@mcp.tool()
def convert_text_to_hwpx(
    text_content: str,
    output_file: str,
    title: str = "",
    styles_file: str = ""
) -> str:
    """Claude가 작성한 마크다운 텍스트를 HWPX 한글 문서 파일로 저장합니다.

    마크다운 헤딩이 자동으로 한글 문서 레벨에 매핑됩니다:
      # → 문서 제목  ## → 섹션  ### → □ 1단계  #### → ○ 2단계
      ##### → ― 3단계  ###### → ※ 4단계

    인라인 색상: {{red:텍스트}}, {{green:텍스트}}

    Args:
        text_content: 변환할 마크다운 형식 텍스트
        output_file: 저장할 HWPX 파일 경로 (예: C:/Users/홍길동/Documents/보고서.hwpx)
        title: 문서 제목 (생략 시 첫 번째 # 헤딩 사용)
        styles_file: 스타일 파일 경로 (생략 시 proposal-styles.json)

    Returns:
        저장된 파일 경로 및 크기
    """
    out_path = Path(output_file)
    if not out_path.is_absolute():
        out_path = Path.home() / "Documents" / output_file

    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        data = parse_markdown_to_json(text_content, title=title)
        styles_path = _resolve_styles_path(styles_file)
        generator = _build_generator(styles_path)
        _run_generator(generator, data, out_path)

        if out_path.exists():
            size = out_path.stat().st_size
            return f"저장 완료!\n경로: {out_path}\n크기: {size:,} bytes"
        return f"오류: 파일 생성에 실패했습니다: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def convert_md_to_hwpx(
    md_file: str,
    output_file: str = "",
    title: str = "",
    styles_file: str = ""
) -> str:
    """마크다운(.md) 파일을 읽어 HWPX 한글 문서로 변환합니다.

    Args:
        md_file: 변환할 마크다운 파일 경로 (예: C:/Users/홍길동/Documents/보고서.md)
        output_file: 저장할 HWPX 파일 경로 (생략 시 md 파일과 같은 위치에 .hwpx 확장자로 저장)
        title: 문서 제목 (생략 시 첫 번째 # 헤딩 사용)
        styles_file: 스타일 파일 경로 (생략 시 proposal-styles.json)

    Returns:
        저장된 파일 경로 및 크기
    """
    md_path = Path(md_file)
    if not md_path.is_absolute():
        md_path = Path.home() / "Documents" / md_file

    if not md_path.exists():
        return f"오류: 마크다운 파일을 찾을 수 없습니다: {md_path}"

    try:
        text_content = _read_md_file(md_path)
    except Exception as e:
        return f"오류: 파일을 읽을 수 없습니다: {type(e).__name__}: {e}"

    if not output_file:
        out_path = md_path.with_suffix(".hwpx")
    else:
        out_path = Path(output_file)
        if not out_path.is_absolute():
            out_path = Path.home() / "Documents" / output_file

    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        data = parse_markdown_to_json(text_content, title=title)
        styles_path = _resolve_styles_path(styles_file)
        generator = _build_generator(styles_path)
        _run_generator(generator, data, out_path)

        if out_path.exists():
            size = out_path.stat().st_size
            return f"저장 완료!\n원본: {md_path}\n경로: {out_path}\n크기: {size:,} bytes"
        return f"오류: 파일 생성에 실패했습니다: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def get_styles(styles_file: str = "") -> str:
    """현재 proposal-styles.json의 스타일 설정(폰트, 크기, 여백 등)을 반환합니다.

    Args:
        styles_file: 스타일 파일 경로 (생략 시 proposal-styles.json)

    Returns:
        현재 스타일 설정 JSON
    """
    styles_path = _resolve_styles_path(styles_file)

    if not styles_path.exists():
        return f"오류: 스타일 파일을 찾을 수 없습니다: {styles_path}"

    with open(styles_path, "r", encoding="utf-8") as f:
        styles = json.load(f)

    return f"스타일 파일: {styles_path}\n\n" + json.dumps(styles, ensure_ascii=False, indent=2)


@mcp.tool()
def update_styles(styles_json: str, styles_file: str = "") -> str:
    """proposal-styles.json의 스타일 규정(폰트, 크기, 여백, 색상)을 업데이트합니다.

    styles_json 예시:
    {
      "styles": {
        "level1": { "font": "HY헤드라인M", "size": 16, "leftMargin": 0 },
        "level2": { "font": "휴먼명조", "size": 14, "leftMargin": 10 }
      },
      "colors": { "red": "#dc2626", "green": "#16a34a" }
    }

    Args:
        styles_json: 새로운 스타일 설정 (JSON 문자열)
        styles_file: 저장할 파일 경로 (생략 시 proposal-styles.json)

    Returns:
        업데이트 결과 메시지
    """
    try:
        styles = json.loads(styles_json)
    except json.JSONDecodeError as e:
        return f"오류: 유효하지 않은 JSON 형식입니다: {e}"

    if "styles" not in styles:
        return "오류: JSON에 'styles' 키가 필요합니다"

    styles_path = _resolve_styles_path(styles_file)

    with open(styles_path, "w", encoding="utf-8") as f:
        json.dump(styles, f, ensure_ascii=False, indent=2)

    return f"스타일이 업데이트되었습니다: {styles_path}\n\n" + json.dumps(styles, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 진입점
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    mcp.run()
