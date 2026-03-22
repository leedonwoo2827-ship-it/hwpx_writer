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

```bash
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
      "command": "D:\mcp\hwpx_writer\.venv\Scripts\python.exe",
      "args": ["D:\mcp\hwpx_writer\server.py"]
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

## 주요 성과 요약

목표 초과 달성 - 전년 대비 15% 향상

### 신규 계약 추진 현황

신규 계약 12건 체결 완료

#### 주요 계약 내역

A사와 3년 장기 계약 체결 (계약금액: 50억 원)

##### 계약 조건

분기별 납품, 하자보증 2년 적용

###### 참고 사항

계약서는 법무팀 검토 완료
```

### 예시 2: 색상 강조

```
{{green:달성}} → 녹색 표시
{{red:미달}}   → 빨간색 표시
```

## 마크다운 → 한글 레벨 매핑 (최신 2026-03-20)

**마크다운 헤딩 `#`~`######` 6단계를 모두 사용하여 HWPX 스타일 레벨 1~6과 1:1 매핑합니다.**

| 마크다운 | 레벨 | 폰트 | 기호 | 한글 왼쪽 | 한글 내어쓰기 |
|---|---|---|---|---|---|
| `#` 제목 | level 1 | Noto Sans KR Bold 13pt | 없음 | 0pt | — |
| `##` 본문 | level 2 | Noto Serif KR 10pt | 없음 | 14.0pt | — |
| `###` 항목 | level 3 | Noto Sans KR 10pt | □ | 14.0pt | 15.0pt |
| `####` 세부 | level 4 | Noto Serif KR 10pt | ○ | 14.3pt | 15.0pt |
| `#####` 보충 | level 5 | Noto Serif KR 10pt | ― | 14.5pt | 15.0pt |
| `######` 참고 | level 6 | Noto Serif KR 10pt | ※ | 14.8pt | 15.0pt |

### ⚠️ 중요: 마크다운 헤딩 사용 원칙

**기호를 직접 입력하지 마세요!** 헤딩 레벨에 따라 hwpx-writer가 자동으로 기호를 추가합니다.

#### ✅ 올바른 사용법

```markdown
### 데이터 수집 방법
SQL 데이터베이스에서 원시 데이터를 추출합니다.

#### 수집 대상 테이블
user_info, order_history, product_catalog

##### 수집 주기
일 1회 자정 배치 작업으로 수집

###### 참고 사항
과거 데이터는 아카이브 DB에 별도 보관
```

**HWPX 렌더링 결과**:
```
□ 데이터 수집 방법
  SQL 데이터베이스에서 원시 데이터를 추출합니다.

  ○ 수집 대상 테이블
    user_info, order_history, product_catalog

    ― 수집 주기
      일 1회 자정 배치 작업으로 수집

      ※ 참고 사항
        과거 데이터는 아카이브 DB에 별도 보관
```

#### ❌ 잘못된 사용법

```markdown
□ 데이터 수집 방법  ← 기호를 직접 입력하지 마세요!
```

이렇게 하면 정렬이 틀어지고 기호가 중복될 수 있습니다.

### 표 스타일

표는 자동으로 다음 스타일이 적용됩니다:
- **폰트**: Noto Sans KR 9pt (고딕체)
- **줄 간격**: 100%
- **헤더**: 중앙 정렬, 굵게
- **데이터**: 중앙 정렬

> Level 3 기호는 □ (U+25A1)을 사용합니다. `proposal-styles.json`의 `level3.symbol` 값으로 변경 가능합니다.

---

## 4. 스타일 커스터마이징

`proposal-styles.json` 파일을 수정하면 폰트, 크기, 여백 등을 변경할 수 있습니다.
MCP 서버가 매 요청마다 이 파일을 다시 읽으므로 **서버 재시작 없이 즉시 반영**됩니다.

### 문단 여백 속성

| 속성 | 설명 | 한글 오피스 대응 | 비율 |
|------|------|----------------|------|
| `paragraphSpaceBefore` | 문단 위 간격 | 문단 모양 → 문단 위 | JSON × 0.25 = 한글 pt |
| `paragraphSpaceAfter` | 문단 아래 간격 | 문단 모양 → 문단 아래 | JSON × 0.25 = 한글 pt |
| `leftMargin` | 왼쪽 여백 기준값 | 문단 모양 → 왼쪽 | JSON × 0.25 = 한글 pt |
| `hangingIndent` | 내어쓰기 크기 (기호 레벨 전용) | 문단 모양 → 내어쓰기 | JSON × 0.25 = 한글 pt |

> **변환 비율**: JSON 값 × 0.25 = 한글 오피스에 표시되는 pt 값
> 예: `"hangingIndent": 60` → 한글에서 15.0pt 내어쓰기

### leftMargin과 hangingIndent의 관계

기호 레벨(3~6)은 `leftMargin`과 `hangingIndent`를 조합하여 내어쓰기를 구현합니다.

```
한글 오피스 표시:
  왼쪽     = (leftMargin + hangingIndent) × 0.25 pt
  내어쓰기 = hangingIndent × 0.25 pt
  첫줄 시작 = 왼쪽 - 내어쓰기
```

**현재 설정 예시** (level3 기준):

```
leftMargin = -4, hangingIndent = 60

→ 한글 왼쪽     = (-4 + 60) × 0.25 = 14.0 pt
→ 한글 내어쓰기 = 60 × 0.25 = 15.0 pt
→ 첫줄 시작     = 14.0 - 15.0 = -1.0 → 0 pt (페이지 여백에 붙음)
```

이 구조에서 기호(□, ○ 등)는 페이지 왼쪽 여백에서 시작하고,
줄바꿈된 텍스트는 본문(level2)과 같은 위치에 정렬됩니다.

### 전체 줄간격

```json
"lineSpacing": 170       // 본문 줄간격 → 한글 170%
"table_data": {
  "lineSpacing": 100     // 표 줄간격 → 한글 100% (lineSpacing은 % 그대로)
}
```

### 설정 예시

```json
{
  "styles": {
    "level2": {
      "font": "Noto Serif KR",
      "size": 10,
      "paragraphSpaceBefore": 3,
      "paragraphSpaceAfter": 3,
      "align": "justify",
      "leftMargin": 56
    },
    "level3": {
      "symbol": "□",
      "font": "Noto Sans KR",
      "size": 10,
      "align": "left",
      "leftMargin": -4,
      "hangingIndent": 60
    }
  },
  "lineSpacing": 170,
  "colors": {
    "red": "#dc2626",
    "green": "#16a34a"
  }
}
```

> 기본 폰트는 **Noto Sans KR**(제목/고딕) + **Noto Serif KR**(본문/명조)입니다.
> Google 오픈 폰트 라이선스(OFL)로 저작권 자유입니다.
> HWPX 파일은 해당 폰트가 PC에 없으면 한글 오피스 기본 폰트로 대체됩니다.

---

## 5. 문제 해결

### Q: Claude Desktop에서 hwpx-writer 도구가 안 보여요.
- `claude_desktop_config.json` 경로가 정확한지 확인하세요.
- 경로의 `\`는 반드시 `\`로 이중 백슬래시를 사용해야 합니다.
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

### Q: 기호(□, ○, ―, ※)가 정렬이 틀어져요.
- **기호를 직접 입력하지 마세요!** 마크다운 헤딩 레벨(`###`, `####`, `#####`, `######`)을 사용하면 자동으로 기호가 추가됩니다.
- 직접 입력한 기호는 제거하고 헤딩 레벨로 변경하세요.

### Q: 표 안의 텍스트가 겹쳐 보여요.
- 최신 버전에서는 표 줄 간격이 200%로 설정되어 있습니다.
- `proposal-styles.json`의 `table_header.lineSpacing`과 `table_data.lineSpacing` 값을 확인하세요.

---

## 6. 파일 구조

```
hwpx_writer/
├── install.bat             ← 자동 설치 스크립트
├── README.md               ← 이 문서
├── server.py               ← MCP 서버 진입점
├── proposal-styles.json    ← 글자/문단 스타일 설정
├── requirements.txt        ← Python 패키지 목록
├── fonts/                  ← Noto Sans KR + Noto Serif KR 폰트
├── src/
│   ├── hwpx_generator.py   ← HWPX 생성 엔진
│   ├── pdf_generator.py    ← PDF 동시 생성 엔진
│   └── md_parser.py        ← 마크다운 파서
└── scripts/
    └── fix_namespaces.py   ← 호환성 후처리 유틸
```

---

## 7. 변경 이력

### 2026-03-22: 기호 레벨 정렬 완전 수정
- 기호 레벨(3~6) 전용 스타일(항목3~6) 도입으로 LEFT 정렬 문제 해결
- leftMargin + hangingIndent 조합으로 내어쓰기 구현
- 본문(level2) 왼쪽여백 추가로 기호 줄바꿈 텍스트와 정렬 일치
- 표 줄간격 200% → 100%로 축소
- 본문 줄간격 150% → 170%로 조정

### 2026-03-20: 새로운 6단계 레벨 체계 도입
- 마크다운 헤딩 `#`~`######` 6단계를 HWPX 스타일 레벨 1~6과 1:1 매핑
- 기호(□, ○, ―, ※) 자동 적용 (직접 입력 금지)
- 본문 폰트 Noto Serif KR (명조체) 통일
- HWPUNIT 변환 수정 (1pt = 50 HWPUNIT)

### 2026-03-18: 표 스타일 개선
- 표 글자 겹침 문제 해결
- 표 lineSpacing 130% → 180% 증가
- 표 셀 높이 동적 계산

### 2026-03-16: 초기 버전
- Noto Sans KR + Noto Serif KR 폰트 추가
- 기본 마크다운 파싱 지원
- HWPX 생성 기능 구현
