#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_xml.py — HWPX 서식의 셀 구조를 분석하는 CLI 도구
=======================================================
HWPX 파일(ZIP) 안의 section XML을 파싱하여
모든 셀의 좌표, 텍스트, 병합 정보, 빈 셀 여부를 출력합니다.

사용법:
    python parse_xml.py "C:\\path\\to\\template.hwpx"
    python parse_xml.py "C:\\path\\to\\template.hwpx" --section Contents/section1.xml
    python parse_xml.py "C:\\path\\to\\template.hwpx" --json

출력 형식 (기본):
    총 셀 수: 138
    행 범위: 0~22
    열 범위: 0~14

      ( 0, 0) span=(1,1) text='번호'
      ( 1, 0) span=(2,1) text='소재지' [HAS TEXT + EMPTY RUN, refs=['18']]
      ( 3, 0) span=(1,1) text='' [EMPTY, refs=['18']]

출력 형식 (--json):
    JSON 배열, 각 요소: {"col":0, "row":0, "colspan":1, "rowspan":1,
                        "text":"번호", "empty":false, "refs":[]}
"""

import sys
import io
import json
import re
import zipfile

# Windows stdout 인코딩 문제 방지
if sys.platform == "win32" and hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def parse_hwpx_cells(hwpx_path, section="Contents/section0.xml"):
    """HWPX 파일의 모든 셀 정보를 파싱하여 리스트로 반환.

    Returns:
        list of dict: 각 셀의 정보
            - col (int): 열 좌표
            - row (int): 행 좌표
            - colspan (int): 열 병합 수
            - rowspan (int): 행 병합 수
            - text (str): 셀 텍스트 내용
            - empty (bool): 빈 입력 셀인지 (empty run이 있는지)
            - has_text_and_empty (bool): 텍스트와 빈 run이 공존하는지
            - refs (list[str]): 빈 run의 charPrIDRef 값 목록
    """
    with zipfile.ZipFile(str(hwpx_path), "r") as zf:
        xml = zf.read(section).decode("utf-8")

    cells_raw = re.findall(r'<hp:cellAddr colAddr="(\d+)" rowAddr="(\d+)"/>', xml)
    result = []

    for col_s, row_s in cells_raw:
        col, row = int(col_s), int(row_s)

        addr_pattern = f'<hp:cellAddr colAddr="{col_s}" rowAddr="{row_s}"/>'
        addr_pos = xml.find(addr_pattern)
        if addr_pos == -1:
            continue

        tc_start = xml.rfind("<hp:tc", 0, addr_pos)
        tc_end = xml.find("</hp:tc>", addr_pos)
        if tc_start == -1 or tc_end == -1:
            continue
        tc_end += len("</hp:tc>")
        cell_xml = xml[tc_start:tc_end]

        # 텍스트 추출
        texts = re.findall(r"<hp:t>(.*?)</hp:t>", cell_xml)
        text_content = "".join(texts).strip()

        # colspan, rowspan
        span_match = re.search(
            r'<hp:cellSpan colSpan="(\d+)" rowSpan="(\d+)"/>', cell_xml
        )
        colspan = int(span_match.group(1)) if span_match else 1
        rowspan = int(span_match.group(2)) if span_match else 1

        # 빈 run 확인
        has_empty_run = bool(re.search(r"<hp:run\b[^>]*/>", cell_xml))
        empty_refs = re.findall(r'<hp:run charPrIDRef="(\d+)"/>', cell_xml)

        result.append(
            {
                "col": col,
                "row": row,
                "colspan": colspan,
                "rowspan": rowspan,
                "text": text_content,
                "empty": not text_content and has_empty_run,
                "has_text_and_empty": bool(text_content) and has_empty_run,
                "refs": empty_refs,
            }
        )

    return result


def print_cells_text(cells):
    """셀 정보를 사람이 읽기 좋은 텍스트로 출력."""
    if not cells:
        print("셀을 찾지 못했습니다.")
        return

    max_row = max(c["row"] for c in cells)
    max_col = max(c["col"] for c in cells)

    print(f"총 셀 수: {len(cells)}")
    print(f"행 범위: 0~{max_row}")
    print(f"열 범위: 0~{max_col}")
    print()

    for c in cells:
        status = ""
        if c["empty"]:
            status = f" [EMPTY, refs={c['refs']}]"
        elif c["has_text_and_empty"]:
            status = f" [HAS TEXT + EMPTY RUN, refs={c['refs']}]"

        line = (
            f"  ({c['col']:>2},{c['row']:>2}) "
            f"span=({c['colspan']},{c['rowspan']}) "
            f"text='{c['text']}'{status}"
        )
        print(line)


def print_cells_json(cells):
    """셀 정보를 JSON 배열로 출력."""
    print(json.dumps(cells, ensure_ascii=False, indent=2))


def main():
    # CLI 인자 처리
    args = sys.argv[1:]

    if not args or args[0] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0)

    hwpx_path = args[0]
    section = "Contents/section0.xml"
    output_json = False

    i = 1
    while i < len(args):
        if args[i] == "--section" and i + 1 < len(args):
            section = args[i + 1]
            i += 2
        elif args[i] == "--json":
            output_json = True
            i += 1
        else:
            print(f"알 수 없는 옵션: {args[i]}", file=sys.stderr)
            i += 1

    cells = parse_hwpx_cells(hwpx_path, section)

    if output_json:
        print_cells_json(cells)
    else:
        print_cells_text(cells)


if __name__ == "__main__":
    main()
