---
name: hwpx_writer
description: "마크다운 텍스트 또는 MD 파일을 proposal-styles.json 스타일이 적용된 HWPX 한글 문서로 변환하는 MCP 스킬."
---

# HWPX Writer — Claude Desktop MCP 스킬

Claude Desktop에서 작성한 글을 **한글(.hwpx) 파일로 자동 저장**합니다.

---

## 제공 도구 (MCP Tools)

### `convert_text_to_hwpx` — 텍스트 → HWPX 변환 (핵심)

Claude가 작성한 마크다운 텍스트를 스타일이 적용된 HWPX 파일로 저장합니다.

| 파라미터 | 설명 |
|---------|------|
| `text_content` | 변환할 마크다운 텍스트 |
| `output_file` | 저장 경로 (예: `C:/Users/홍길동/Documents/보고서.hwpx`) |
| `title` | 문서 제목 (생략 시 첫 번째 `#` 헤딩 사용) |
| `styles_file` | 스타일 파일 경로 (생략 시 `proposal-styles.json`) |

### `convert_md_to_hwpx` — MD 파일 → HWPX 변환

마크다운(.md) 파일을 읽어 스타일이 적용된 HWPX 파일로 변환합니다.

| 파라미터 | 설명 |
|---------|------|
| `md_file` | 변환할 마크다운 파일 경로 (예: `C:/Users/홍길동/Documents/보고서.md`) |
| `output_file` | 저장 경로 (생략 시 MD 파일과 같은 위치에 `.hwpx` 확장자로 저장) |
| `title` | 문서 제목 (생략 시 첫 번째 `#` 헤딩 사용) |
| `styles_file` | 스타일 파일 경로 (생략 시 `proposal-styles.json`) |

### `get_styles` — 스타일 조회

현재 `proposal-styles.json`의 폰트/크기/여백 설정을 확인합니다.

### `update_styles` — 스타일 수정

폰트, 글자 크기, 여백, 색상 등을 수정합니다.

---

## 마크다운 → 한글 문서 레벨 매핑

| 마크다운 | 한글 문서 | 기본 스타일 |
|---------|----------|-----------|
| `# 제목` | 문서 제목 | KoPubWorld돋움체 Bold, 25pt, 가운데 |
| `## 섹션` | 섹션 구분 | — |
| `### 내용` | □ 1단계 | KoPubWorld바탕체 Bold, 16pt |
| `#### 내용` | ○ 2단계 | KoPubWorld바탕체 Medium, 12pt |
| `##### 내용` | ― 3단계 | KoPubWorld바탕체 Light, 11pt |
| `###### 내용` | ※ 4단계 | KoPubWorld돋움체 Medium, 10pt |
| `- 항목` | 3단계 목록 | — |
| `\| 표 \|` | 표 | — |

### 색상 마커

```
{{red:텍스트}}    → 빨간색 (#dc2626)
{{green:텍스트}}  → 녹색   (#16a34a)
```

---

## 사용 예시

```
아래 내용을 C:/Users/홍길동/Documents/보고서.hwpx 로 저장해줘

# 2026년 1분기 업무 보고

## 1. 추진 실적

### 주요 성과
{{green:목표 초과 달성}} - 전년 대비 15% 향상

#### 세부 내역
- 신규 계약 12건
- 기존 계약 갱신 8건
```

---

## 파일 구조

```
hwpx_writer/
├── server.py               ← MCP 서버 진입점
├── proposal-styles.json    ← 글자/문단 스타일 설정
├── requirements.txt        ← Python 패키지 목록
├── SKILL.md                ← 이 문서
├── src/
│   ├── hwpx_generator.py   ← HWPX 생성 엔진
│   └── md_parser.py        ← 마크다운 파서
└── scripts/
    └── fix_namespaces.py   ← 한글 호환성 후처리
```
