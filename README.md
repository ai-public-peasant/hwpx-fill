# hwpx-fill

**AI-powered HWPX (Korean word processor) form filler for Claude Code**

한글(HWPX) 서식에 엑셀 데이터를 자동으로 채워넣는 Claude Code 스킬입니다.
AI가 서식 구조를 분석하고, 엑셀 헤더와 의미적으로 매핑한 뒤, Python 스크립트를 생성하여 대량의 서식 파일을 자동 생성합니다.

---

## 핵심: 선작업 제로 (Zero Pre-configuration)

| 기존 방식 | hwpx-fill |
|----------|-----------|
| 서식마다 누름틀·매핑 설정 필요 | **설정 없음** — AI가 알아서 분석 |
| 새 서식이면 처음부터 다시 | 새 서식을 주면 AI가 다시 분석 |
| 서식 10종이면 설정도 10번 | 서식 10종이어도 **그냥 던져주면 끝** |

기존 RPA, 메일머지, 누름틀 매핑 프로그램은 서식 1종마다 사람이 "이 칸 = 이 열" 매핑을 설정해야 했습니다.
이 스킬은 AI가 **"기관(업체)명" 옆 빈칸에는 엑셀의 "학원명"이 들어가겠구나**라고 스스로 판단합니다.

---

## 설치

```bash
# Claude Code의 글로벌 스킬 디렉토리에 클론
cd ~/.claude/skills
git clone https://github.com/ai-public-peasant/hwpx-fill.git
```

### 요구사항
- **Claude Code** (with Claude AI)
- **Python 3.10+**
- **openpyxl** (`pip install openpyxl`)
- **HWPX MCP 서버** (선택, 서식 분석 시 유용)

---

## 사용법

Claude Code에서 다음과 같이 말하면 됩니다:

```
한글 서식 채워줘
서식에 데이터 넣어줘
hwpx에 엑셀 데이터 매핑해줘
```

또는 명시적으로:

```
/hwpx-fill
```

### 워크플로우

```
① 빈 HWPX 서식 + 엑셀 데이터 제공
② AI가 서식 구조 자동 분석 (셀 좌표, 라벨, 빈 칸 식별)
③ AI가 엑셀 헤더와 서식 라벨을 의미적으로 매핑
④ 매핑 제안 → 유저 확인
⑤ 채우기 스크립트 자동 생성
⑥ 테스트 실행 (3~5건) → 확인
⑦ 전체 실행 → 수백 장의 완성 서식 생성
```

---

## 파일 구조

```
hwpx-fill/
├── SKILL.md              # 스킬 정의 + AI 워크플로우 지침
├── README.md             # 이 문서
├── LICENSE               # MIT License
└── scripts/
    ├── hwpx_utils.py     # 공용 유틸리티 라이브러리
    └── parse_xml.py      # HWPX 셀 구조 분석 CLI 도구
```

### scripts/hwpx_utils.py

HWPX 서식 채우기에 필요한 공용 함수 라이브러리:

| 함수 | 설명 |
|------|------|
| `read_hwpx_xml()` | HWPX(ZIP) 안의 XML 파트 읽기 |
| `write_xml_to_hwpx()` | 수정된 XML을 HWPX에 다시 쓰기 |
| `fill_cell_by_addr()` | cellAddr 좌표로 특정 셀에 값 채우기 |
| `read_excel_data()` | 엑셀 데이터 읽기 (단순) |
| `read_excel_grouped()` | 엑셀 데이터를 키로 그룹핑하여 읽기 |
| `normalize_phone()` | 전화번호 정규화 (선행 0 복원 등) |
| `normalize_date()` | 날짜 정규화 |
| `normalize_area()` | 면적 값 정규화 |

### scripts/parse_xml.py

HWPX 파일의 셀 구조를 분석하는 CLI 도구:

```bash
python parse_xml.py "template.hwpx"
python parse_xml.py "template.hwpx" --json
```

출력 예시:
```
총 셀 수: 116
행 범위: 0~22
열 범위: 0~14

  ( 3, 1) span=(3,1) text='성명'
  ( 6, 1) span=(3,1) text='' [EMPTY, refs=['7']]
  ( 9, 1) span=(3,1) text='생년월일'
  (12, 1) span=(3,1) text='' [EMPTY, refs=['7']]
```

---

## 기술 배경

### HWPX란?
한글 문서(.hwp)의 최신 형식으로, ZIP 압축 파일 안에 XML이 들어있는 구조입니다.

### 작동 원리
1. HWPX를 ZIP으로 열어 `Contents/section0.xml` 추출
2. XML에서 `<hp:cellAddr>` 태그로 셀 좌표 식별
3. 빈 셀의 self-closing `<hp:run charPrIDRef="N"/>` 태그에 `<hp:t>값</hp:t>` 삽입
4. 수정된 XML을 다시 ZIP에 써넣기

### AI의 역할 vs 스크립트의 역할

```
AI (Claude):
  ✓ 서식 구조 분석 (라벨, 빈 칸, 반복 행 식별)
  ✓ 엑셀 헤더와 의미적 매핑 ("기관(업체)명" ≈ "학원명")
  ✓ 데이터 그룹핑 전략 결정
  ✓ 채우기 스크립트 생성

Python 스크립트:
  ✓ 엑셀 데이터 읽기 (수백~수천 행)
  ✓ 템플릿 복사 + XML 수정 (기계적 반복)
  ✓ 합계 계산, 파일 저장
```

---

## 한계

| 한계 | 설명 |
|------|------|
| HWPX만 지원 | 구 HWP 형식은 한글에서 HWPX로 변환 필요 |
| 텍스트만 가능 | 이미지/도장 삽입은 미지원 |
| 엑셀에 없는 항목 | 서식에 있지만 엑셀에 없는 항목은 빈칸 유지 |
| 매핑 확인 필수 | AI 판단이 틀릴 수 있으므로 테스트 확인 단계 필수 |

---

## License

MIT License
