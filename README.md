# HWPX Writer MCP — 설치 설명서

Claude Desktop에서 마크다운 텍스트를 **한글(.hwpx) 문서**로 자동 변환하는 MCP 서버입니다.

---

## 1. 설치 방법

### 방법 A: 자동 설치 (권장)

1. `hwpx_writer` 폴더를 원하는 위치에 압축 해제합니다.
   - 예: `D:\mcp\hwpx_writer`
   - **경로에 한글이나 공백이 없는 것을 권장합니다.**

2. `install.bat`을 **더블클릭**합니다.
   - Python이 없으면 자동으로 3.11.9 버전을 설치합니다.
   - 가상환경 생성 및 패키지 설치를 자동으로 수행합니다.
   - 완료 후 Claude Desktop 설정에 붙여넣을 내용을 안내합니다.

3. 안내에 따라 Claude Desktop 설정 파일을 수정합니다. (아래 2번 참고)

4. Claude Desktop을 **재시작**합니다.

### 방법 B: 수동 설치

```
cd D:\mcp\hwpx_writer
python -m venv .venv
.venv\Scripts\pip install -r requirements.txt
```

---

## 2. Claude Desktop 설정

설정 파일 위치:
```
%APPDATA%\Claude\claude_desktop_config.json
```

> Windows 탐색기 주소창에 `%APPDATA%\Claude`를 입력하면 바로 이동됩니다.

### 설정 추가

파일을 열어 `mcpServers` 안에 아래 내용을 추가합니다.
**경로는 실제 설치 위치로 수정하세요.**

```json
{
  "mcpServers": {
    "hwpx-writer": {
      "command": "D:\\mcp\\hwpx_writer\\.venv\\Scripts\\python.exe",
      "args": ["D:\\mcp\\hwpx_writer\\server.py"]
    }
  }
}
```

> 이미 다른 MCP가 등록되어 있다면, `mcpServers` 안에 `"hwpx-writer": { ... }` 부분만 추가하세요.

설정 후 **Claude Desktop을 재시작**하면 적용됩니다.

---

## 3. 사용 방법

Claude Desktop에서 자연어로 요청하면 됩니다.

### 예시 1: 텍스트를 HWPX로 저장

```
아래 내용을 C:/Users/홍길동/Documents/보고서.hwpx 로 저장해줘

# 2026년 1분기 업무 보고

## 1. 추진 실적

### 주요 성과
목표 초과 달성 - 전년 대비 15% 향상

#### 세부 내역
- 신규 계약 12건
- 기존 계약 갱신 8건
```

### 예시 2: 색상 강조

```
{{green:달성}} → 녹색 표시
{{red:미달}}   → 빨간색 표시
```

### 마크다운 → 한글 레벨 매핑

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

---

## 4. 스타일 커스터마이징

`proposal-styles.json` 파일을 수정하면 폰트, 크기, 여백 등을 변경할 수 있습니다.

```json
{
  "styles": {
    "level1": {
      "font": "KoPubWorld돋움체 Bold",
      "size": 18,
      "paragraphSpaceBefore": 25,
      "paragraphSpaceAfter": 8,
      "leftMargin": 0
    }
  },
  "colors": {
    "red": "#dc2626",
    "green": "#16a34a"
  }
}
```

> 폰트는 **PC에 설치된 폰트 이름**을 정확히 입력해야 합니다.

---

## 5. 문제 해결

### Q: Claude Desktop에서 hwpx-writer 도구가 안 보여요.
- `claude_desktop_config.json` 경로가 정확한지 확인하세요.
- 경로의 `\`는 반드시 `\\`로 이중 백슬래시를 사용해야 합니다.
- Claude Desktop을 재시작했는지 확인하세요.

### Q: "Python을 찾을 수 없습니다" 오류가 나요.
- `install.bat`을 다시 실행하세요.
- 또는 [python.org](https://www.python.org/downloads/)에서 직접 설치 후, **PATH에 추가** 옵션을 체크하세요.

### Q: HWPX 파일이 한글에서 안 열려요.
- 한컴오피스 2020 이상 버전이 필요합니다.
- `.hwpx` 확장자가 맞는지 확인하세요.

### Q: 글꼴이 다르게 나와요.
- `proposal-styles.json`에 지정된 폰트가 PC에 설치되어 있어야 합니다.
- 설치되지 않은 폰트는 한글 오피스가 기본 폰트로 대체합니다.

---

## 6. 파일 구조

```
hwpx_writer/
├── install.bat             ← 자동 설치 스크립트
├── README.md               ← 이 문서
├── server.py               ← MCP 서버 진입점
├── proposal-styles.json    ← 글자/문단 스타일 설정
├── requirements.txt        ← Python 패키지 목록
├── src/
│   ├── hwpx_generator.py   ← HWPX 생성 엔진
│   └── md_parser.py        ← 마크다운 파서
└── scripts/
    └── fix_namespaces.py   ← 호환성 후처리 유틸
```
