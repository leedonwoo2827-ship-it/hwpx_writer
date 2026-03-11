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


def _resolve_project_output(project_dir: str, output_file: str, fallback_name: str = "output.hwpx") -> Path:
    """project_dir 또는 output_file에서 출력 경로를 결정합니다."""
    if project_dir:
        proj = Path(project_dir)
        out_dir = proj / "output"
        out_dir.mkdir(parents=True, exist_ok=True)
        name = Path(output_file).name if output_file else fallback_name
        return out_dir / name
    if output_file:
        p = Path(output_file)
        return p if p.is_absolute() else Path.home() / "Documents" / output_file
    return Path.home() / "Documents" / fallback_name


def _inject_images(data: dict, image_paths_str: str) -> None:
    """이미지 경로 목록을 data의 content 끝에 image 항목으로 추가합니다."""
    if not image_paths_str:
        return
    paths = [p.strip() for p in image_paths_str.split(",") if p.strip()]
    for p in paths:
        data["content"].append({"type": "image", "path": p})


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
    output_file: str = "",
    title: str = "",
    styles_file: str = "",
    project_dir: str = "",
    image_paths: str = ""
) -> str:
    """Claude가 작성한 마크다운 텍스트를 HWPX 한글 문서 파일로 저장합니다.

    마크다운 헤딩이 자동으로 한글 문서 레벨에 매핑됩니다:
      제목체: # → 문서 제목  ## → 섹션(1.)  ### → 소제목(1.1)
      본문체: #### → □ 본문항목  ##### → ○ 세부항목  ###### → ― 보충항목
      기호 계층: □ → ○ → ― → ※ (들여쓰기 단계적 증가)

    인라인 색상: {{red:텍스트}}, {{green:텍스트}}

    Args:
        text_content: 변환할 마크다운 형식 텍스트
        output_file: 저장할 HWPX 파일명 (project_dir 지정 시 파일명만 필요)
        title: 문서 제목 (생략 시 첫 번째 # 헤딩 사용)
        styles_file: 스타일 파일 경로 (생략 시 proposal-styles.json)
        project_dir: 프로젝트 폴더 경로 (예: C:/Users/ubion/Documents/proposals/260311-n). 지정 시 output/ 하위에 저장
        image_paths: 삽입할 이미지 절대경로 목록 (쉼표 구분, 예: "C:/.../a.png,C:/.../b.png"). 문서 끝에 순서대로 삽입

    Returns:
        저장된 파일 경로 및 크기
    """
    out_path = _resolve_project_output(project_dir, output_file, "output.hwpx")
    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        data = parse_markdown_to_json(text_content, title=title)
        _inject_images(data, image_paths)
        styles_path = _resolve_styles_path(styles_file)
        generator = _build_generator(styles_path)
        _run_generator(generator, data, out_path)

        if out_path.exists():
            size = out_path.stat().st_size
            img_count = len([p for p in image_paths.split(",") if p.strip()]) if image_paths else 0
            img_msg = f"\n이미지: {img_count}개 삽입" if img_count else ""
            return f"저장 완료!\n경로: {out_path}\n크기: {size:,} bytes{img_msg}"
        return f"오류: 파일 생성에 실패했습니다: {out_path}"

    except Exception as e:
        return f"오류: {type(e).__name__}: {e}"


@mcp.tool()
def convert_md_to_hwpx(
    md_file: str,
    output_file: str = "",
    title: str = "",
    styles_file: str = "",
    project_dir: str = "",
    image_paths: str = ""
) -> str:
    """마크다운(.md) 파일을 읽어 HWPX 한글 문서로 변환합니다.

    Args:
        md_file: 변환할 마크다운 파일 경로 (예: C:/Users/홍길동/Documents/보고서.md)
        output_file: 저장할 HWPX 파일명 (project_dir 지정 시 파일명만 필요, 생략 시 md 파일명.hwpx)
        title: 문서 제목 (생략 시 첫 번째 # 헤딩 사용)
        styles_file: 스타일 파일 경로 (생략 시 proposal-styles.json)
        project_dir: 프로젝트 폴더 경로 (예: C:/Users/ubion/Documents/proposals/260311-n). 지정 시 output/ 하위에 저장
        image_paths: 삽입할 이미지 절대경로 목록 (쉼표 구분). 문서 끝에 순서대로 삽입

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

    fallback_name = md_path.with_suffix(".hwpx").name
    if project_dir:
        out_path = _resolve_project_output(project_dir, output_file, fallback_name)
    elif output_file:
        out_path = Path(output_file)
        if not out_path.is_absolute():
            out_path = Path.home() / "Documents" / output_file
    else:
        out_path = md_path.with_suffix(".hwpx")

    out_path.parent.mkdir(parents=True, exist_ok=True)

    try:
        data = parse_markdown_to_json(text_content, title=title)
        _inject_images(data, image_paths)
        styles_path = _resolve_styles_path(styles_file)
        generator = _build_generator(styles_path)
        _run_generator(generator, data, out_path)

        if out_path.exists():
            size = out_path.stat().st_size
            img_count = len([p for p in image_paths.split(",") if p.strip()]) if image_paths else 0
            img_msg = f"\n이미지: {img_count}개 삽입" if img_count else ""
            return f"저장 완료!\n원본: {md_path}\n경로: {out_path}\n크기: {size:,} bytes{img_msg}"
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
