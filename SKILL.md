---
name: hwpx-fill
description: |
  한글(HWPX) 서식에 엑셀 데이터를 자동으로 채워넣는 스킬.
  AI가 서식 구조를 분석하고, 엑셀 헤더와 의미적으로 매핑한 뒤,
  Python 스크립트를 생성하여 대량의 서식 파일을 자동 생성합니다.
  "한글 서식 채워줘", "HWPX 자동 입력", "서식에 데이터 넣어줘",
  "엑셀 데이터 한글에 채워줘" 요청 시 사용합니다.
---

# HWPX 서식 자동 채우기 스킬

한글(HWPX) 서식 파일과 엑셀(.xlsx) 데이터 파일을 받아서, AI가 서식 구조를 분석하고 엑셀 데이터와 매핑한 뒤, 대량의 완성된 서식 파일을 자동 생성한다.

**핵심: 사전 설정(누름틀, 매핑 지정) 없이 AI가 알아서 분석하고 매핑한다.**

---

## 트리거 조건

다음과 같은 요청 시 이 스킬을 사용:
- "한글 서식 채워줘" / "HWPX 자동 입력"
- "서식에 데이터 넣어줘" / "엑셀 데이터 한글에 채워줘"
- "hwpx에 엑셀 데이터 매핑해줘"
- `.hwpx` 파일과 `.xlsx` 파일이 함께 언급된 요청
- "서식 자동 작성" / "신청서 대량 생성"

---

## 유저 입력 확인

작업 시작 전, 다음 정보를 확인한다:

1. **HWPX 서식 경로** — 빈 양식 파일 (`.hwpx`)
   - `.hwp` 파일이면 한글에서 `.hwpx`로 다른이름저장 안내
2. **엑셀 데이터 경로** — 데이터가 들어있는 `.xlsx` 파일
3. **출력 폴더** — 기본값: 엑셀 파일 옆에 `output/` 생성
4. **테스트 건수** — 먼저 몇 건 생성할지 (기본값: 5건)
5. **파일명 패턴** — 출력 파일명에 사용할 컬럼 (예: `{성명}_{번호}.hwpx`)

유저가 "이 서식이랑 이 엑셀 줄게"처럼 간단히 말하면, 작업 폴더의 `.hwpx`와 `.xlsx`를 자동 탐지한다.

---

## 워크플로우

### Step 1: 입력 파일 확인

- 유저가 명시한 경로 또는 작업 폴더에서 `.hwpx`와 `.xlsx` 파일을 찾는다.
- `.hwpx`가 여러 개면 유저에게 선택 요청.
- `.hwp` 파일만 있으면: "한글에서 `.hwpx` 형식으로 다른이름저장 해주세요" 안내.
- 출력 폴더를 확인하고 생성한다.

### Step 2: HWPX 서식 구조 분석

**이 단계가 가장 중요하다. 두 가지 방법을 순서대로 사용한다.**

#### 방법 A: HWPX MCP 도구 (전체 구조 파악)

다음 MCP 도구로 서식의 전체적인 구조를 먼저 파악한다:

```
mcp__hwpx__text_extract_report  → 서식 전체 텍스트 추출
mcp__hwpx__read_text            → 단락 구조와 텍스트
mcp__hwpx__find                 → 특정 라벨 검색
mcp__hwpx__package_get_text     → Contents/section0.xml 원본 XML
mcp__hwpx__package_parts        → 어떤 section이 있는지 확인
```

⚠️ 다음 MCP 도구는 **사용하지 않는다** (XML 파싱 오류로 실패함):
- `save_as`, `set_table_cell_text`, `get_table_cell_map`

#### 방법 B: parse_xml.py 실행 (정확한 셀 좌표 그리드)

MCP 도구로 전체 구조를 파악한 후, 정확한 셀 좌표를 얻기 위해 실행한다:

```bash
python "C:\Users\a\.claude\skills\hwpx-fill\scripts\parse_xml.py" "서식파일.hwpx"
```

JSON 형식이 필요하면:
```bash
python "C:\Users\a\.claude\skills\hwpx-fill\scripts\parse_xml.py" "서식파일.hwpx" --json
```

#### 식별해야 할 것들

parse_xml.py 출력을 분석하여 다음을 식별한다:

1. **라벨 셀**: 텍스트가 있는 셀 (예: `text='성명'`, `text='주소'`)
2. **빈 입력 셀**: `[EMPTY, refs=['18']]`로 표시된 셀 — 데이터를 채울 대상
3. **charPrIDRef 값**: refs에 나오는 숫자 — 서식마다 다름 (18, 7, 17 등)
4. **반복 행 구간**: 같은 열 패턴이 여러 행에 걸쳐 반복되는 영역 (필지, 품목 등)
5. **합계 셀**: 텍스트가 "0"이거나 빈 셀인데, 위치상 합계를 넣어야 하는 곳
6. **서명란**: 표 바깥에 있는 텍스트 (예: `"신청인          (인)"`)

**분석 결과를 다음 형식으로 정리한다:**

```
[서식 구조 분석 결과]
테이블: N열 × M행
헤더 정보 셀:
  (col, row) = "라벨"  ← 옆 빈칸 (col2, row2) [EMPTY]
  ...
반복 데이터 행: row X ~ row Y (최대 Z건)
  각 행 구조: (col, row)=항목1, (col, row)=항목2, ...
합계 셀: (col, row) — 현재 "0" 또는 빈칸
서명란: "원문 텍스트"
charPrIDRef: 주로 사용되는 값 = N
```

### Step 3: 엑셀 데이터 분석

엑셀 데이터의 구조를 파악한다:

1. **헤더 읽기**: 1행의 열 이름 전체 목록
2. **샘플 데이터**: 2~6행의 데이터를 읽어서 각 열의 타입 파악
3. **데이터 카디널리티 판별**:
   - 각 행이 독립적 = **단순 1:1 매핑** (예: 학원별 신청서)
   - 같은 사람이 여러 행 = **그룹핑 필요** (예: 한 사람의 여러 필지)
   - 판별 방법: 이름·주민번호·등록번호 등 식별 컬럼에 중복값이 있는지 확인

**엑셀 MCP 도구를 사용해도 좋다:**
```
mcp__excel__read_data_from_excel → 데이터 읽기
mcp__excel__get_workbook_metadata → 시트 목록, 범위 확인
```

### Step 4: 매핑 제안

서식 라벨과 엑셀 헤더를 **의미적으로 연결**하여 매핑표를 작성하고, 유저에게 확인을 받는다.

#### 매핑 추론 규칙

- **동일 명칭**: "성명" ↔ "성명", "연락처" ↔ "연락처" → 자동 매핑
- **유의어 매핑**: 다음 유의어 쌍을 인식한다:
  - 기관/업체/학원/회사/단체
  - 대표자/대표명/대표
  - 연락처/전화번호/휴대전화/핸드폰
  - 생년월일/생년/주민번호 앞자리
  - 주소/소재지/거주지/도로명주소
  - 면적/경작면적/재배면적
  - 번호/지번/번지
  - 읍면/읍면동/행정구역
  - 리동/리/동

#### 매핑표 형식

유저에게 다음 형식으로 보여준다:

```
=== 매핑 제안 ===
서식 필드 (좌표)           ← 엑셀 컬럼         판단 근거
─────────────────────────────────────────────────────
성명 (col=6, row=1)       ← 성명 (E열)        동일 명칭
기관(업체)명 (col=2, row=2) ← 학원명 (A열)     기관/업체 = 학원
연락처 (col=6, row=3)     ← 연락처 (D열)      동일 명칭
⚠️ 매핑 불가: 지목 (서식에 있지만 엑셀에 해당 컬럼 없음 → 빈칸 유지)

데이터 구조: [단순 1:1] 또는 [그룹핑 - 키: {컬럼명}]
반복 행: row X~Y, 최대 Z건/파일
합계: col=N, row=M (면적 합산)
서명란: "원본텍스트" → "{변수} 치환 패턴"
출력 파일명: {컬럼명}.hwpx
```

**반드시 유저 확인을 받은 후 다음 단계로 진행한다.** 유저가 매핑을 수정할 수 있다.

### Step 5: 채우기 스크립트 생성

유저가 매핑을 확인하면, **작업 폴더에** Python 스크립트를 생성한다.

#### 생성 스크립트 구조

```python
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""[서식명] 자동 채우기 스크립트 - AI 생성"""

import sys, io, os, re, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# hwpx_utils 라이브러리 import
sys.path.insert(0, r"C:\Users\a\.claude\skills\hwpx-fill\scripts")
from pathlib import Path
from hwpx_utils import (
    read_hwpx_xml, write_xml_to_hwpx, fill_cell_by_addr,
    read_excel_data, read_excel_grouped,
    xml_escape, sanitize_filename,
    normalize_phone, normalize_date, normalize_area, area_text,
)

# ── 설정 ──
TEMPLATE_PATH = r"..."
EXCEL_PATH = r"..."
OUTPUT_DIR = Path(r"...")
TEST_LIMIT = 5  # None이면 전체 실행

# ── 매핑 (AI가 결정) ──
# Case A: 단순 1:1 매핑
FIELD_MAP = {
    # (col, row): 'excel_header'
}
SIGNATURE_REPLACEMENTS = [
    # ('원본텍스트', '치환텍스트 패턴')
]

# Case B: 그룹핑 매핑
# GROUP_KEY = '그룹핑키'
# HEADER_FIELDS = {(col, row): 'excel_header', ...}
# ROW_FIELDS = [(col, 'excel_header'), ...]
# ROW_START = 6
# ROW_MAX = 12
# TOTAL_CELL = (col, row)  # None이면 합계 없음

# ── 실행 ──
def main():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    template_xml = read_hwpx_xml(TEMPLATE_PATH)
    headers, rows = read_excel_data(EXCEL_PATH)

    count = 0
    for row_data in rows:
        if TEST_LIMIT and count >= TEST_LIMIT:
            break

        # 파일 생성
        filename = sanitize_filename(f"{row_data.get('키필드', count+1)}") + ".hwpx"
        out_path = OUTPUT_DIR / filename
        shutil.copy2(TEMPLATE_PATH, str(out_path))

        xml = template_xml
        for (col, row), header in FIELD_MAP.items():
            value = str(row_data.get(header, '') or '')
            xml = fill_cell_by_addr(xml, col, row, value)

        for old, new in SIGNATURE_REPLACEMENTS:
            xml = xml.replace(old, new.format(**row_data))

        write_xml_to_hwpx(str(out_path), 'Contents/section0.xml', xml)
        count += 1
        print(f"  [{count}] {filename}")

    print(f"\n완료: {count}건 생성 → {OUTPUT_DIR}")

if __name__ == '__main__':
    main()
```

**위 코드는 뼈대 참고용이다. 실제로는 매핑 분석 결과에 맞게 커스터마이즈하여 생성한다.**

#### 그룹핑 케이스 추가 로직

그룹핑이 필요한 경우 (한 사람 = 여러 행):

```python
headers, groups = read_excel_grouped(EXCEL_PATH, GROUP_KEY)

for key, group_rows in groups.items():
    first = group_rows[0]  # 인적사항은 첫 행에서

    # 헤더 필드 채우기 (성명, 주소 등)
    for (col, row), header in HEADER_FIELDS.items():
        value = str(first.get(header, '') or '')
        xml = fill_cell_by_addr(xml, col, row, value)

    # 반복행 채우기 (필지, 품목 등)
    for idx, row_data in enumerate(group_rows[:ROW_MAX]):
        data_row = ROW_START + idx
        for col, header in ROW_FIELDS:
            value = str(row_data.get(header, '') or '')
            xml = fill_cell_by_addr(xml, col, data_row, value)

    # 합계 계산
    if TOTAL_CELL:
        total = sum(normalize_area(r.get(TOTAL_HEADER)) for r in group_rows[:ROW_MAX])
        xml = fill_cell_by_addr(xml, TOTAL_CELL[0], TOTAL_CELL[1], area_text(total))
```

### Step 6: 테스트 실행

스크립트를 TEST_LIMIT=5 (또는 유저 지정)으로 실행한다:

```bash
python "fill_서식명.py"
```

실행 후:
1. 생성된 파일 수와 목록을 출력한다.
2. 유저에게 "한글에서 1~2개 열어서 확인해주세요" 안내.
3. 문제가 있으면 매핑이나 스크립트를 수정하고 재실행.

### Step 7: 전체 실행

유저가 테스트 결과를 확인하면:
1. `TEST_LIMIT = None`으로 변경
2. 전체 데이터 실행
3. 결과 요약 출력: 총 생성 파일 수, 출력 폴더 경로

---

## 함정과 우회방법

### 1. Windows stdout 인코딩

생성하는 모든 Python 스크립트 최상단에 반드시 추가:
```python
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
```

### 2. WindowsPath vs str

`zipfile.ZipFile()`과 `os.replace()`에 `Path` 객체를 직접 넘기면 에러 발생할 수 있다.
항상 `str(path)`로 변환하여 사용한다. `hwpx_utils.py`는 내부적으로 이를 처리한다.

### 3. charPrIDRef는 서식마다 다르다

빈 셀의 `<hp:run charPrIDRef="N"/>` 에서 N 값은 서식마다 다르다.
절대 하드코딩하지 않는다. `hwpx_utils.py`의 `fill_cell_by_addr`은 범용 regex `r'<hp:run\b[^>]*/>'`로 어떤 ref값이든 매칭한다.

### 4. XML 이스케이프 필수

셀에 넣을 값에 `&`, `<`, `>` 문자가 있으면 XML이 깨진다.
`hwpx_utils.xml_escape()`를 사용하거나, `fill_cell_by_addr()`이 내부적으로 처리한다.

### 5. 기존 텍스트가 있는 셀 (합계 등)

합계 칸에 "0"이 미리 들어있는 경우, 빈 run이 아닌 기존 `<hp:t>0</hp:t>`를 교체해야 한다.
`fill_cell_by_addr()`의 fallback 로직이 이를 처리한다.

### 6. 전화번호가 숫자로 저장됨

엑셀에서 전화번호를 숫자로 저장하면 앞의 `0`이 사라진다.
`normalize_phone()`으로 정규화한다: `1012345678` → `010-1234-5678`

### 7. 날짜가 datetime 객체

openpyxl이 날짜 셀을 `datetime` 객체로 반환할 수 있다.
`normalize_date()`로 문자열 변환한다.

### 8. ZIP 임시파일 패턴

HWPX를 수정할 때 반드시 `.tmp` 파일로 먼저 쓰고 `os.replace()`로 교체한다.
원본에 직접 쓰면 ZIP이 깨진다. `hwpx_utils.write_xml_to_hwpx()`가 이를 처리한다.

### 9. Python 버전

Python 3.11 사용: `C:\Users\a\AppData\Local\Programs\Python\Python311\python.exe`
Python 3.14는 경로 인코딩 문제가 있으므로 사용하지 않는다.

### 10. Bash에서 한글 경로

Claude Code의 Bash는 유닉스 셸이므로 한글 경로를 따옴표로 감싸야 한다:
```bash
python "C:\Users\a\...\fill_script.py"
```

### 11. 여러 section이 있는 서식

복잡한 서식은 `section0.xml`, `section1.xml` 등 여러 섹션을 가질 수 있다.
`mcp__hwpx__package_parts` 또는 HWPX ZIP 내부를 확인하여 어느 section에 표가 있는지 파악한다.

### 12. 엑셀 외부참조 수식

VLOOKUP 등 외부 파일을 참조하는 수식은 `data_only=True`로도 값을 못 읽을 수 있다.
이 경우 유저에게 "엑셀에서 값으로 붙여넣기 후 저장" 안내.

---

## Python 환경

```
Python: C:\Users\a\AppData\Local\Programs\Python\Python311\python.exe
필수 패키지: openpyxl (pip install openpyxl)
스킬 스크립트: C:\Users\a\.claude\skills\hwpx-fill\scripts\
  - hwpx_utils.py: 공용 유틸리티 라이브러리
  - parse_xml.py: HWPX 셀 구조 분석 CLI 도구
```

---

## 출력

- **생성 스크립트**: 작업 폴더에 `fill_서식명.py` 저장 (향후 재사용 가능)
- **출력 파일**: `{output_dir}/` 안에 `.hwpx` 파일 N개
- **콘솔 요약**: 생성 파일 수, 파일명 목록, 출력 폴더 경로
