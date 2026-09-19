# UI 설계서 PPTX에서 슬라이드별 UI 정보를 뽑아 CSV/Excel로 변환하는 추출기.
"""사용법
    python ui_spec_extract.py <PPTX파일 또는 폴더> [-o 출력폴더]

출력
    <출력폴더>/ui_list.csv   UI 1개 = 1행 (와이드)
    <출력폴더>/ui_flow.csv   흐름 항목 1개 = 1행 (롱)
    <출력폴더>/ui_element.csv 화면 구성요소 1개 = 1행 (버튼 인벤토리용)
    <출력폴더>/ui_spec.xlsx  위를 시트로 담은 엑셀
"""
import argparse
import csv
import re
import sys
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pptx import Presentation

# ── 템플릿 매핑 ──────────────────────────────────────────────
# 슬라이드 상단 플레이스홀더의 도형 이름
SHAPE_대분류 = "텍스트 개체 틀 3"
SHAPE_중분류 = "텍스트 개체 틀 4"
SHAPE_소분류 = "제목 1"

# 본문 표의 크기와 (라벨셀, 값셀) 위치
TABLE_SIZE = (4, 6)
CELL_MAP = {
    "UI명": ((0, 0), (0, 1), "UI 명"),
    "UI_ID": ((0, 2), (0, 3), "UI ID"),
    "UI유형_원문": ((0, 4), (0, 5), "UI 유형"),
    "UI설명": ((1, 0), (1, 1), "UI 설명"),
}
CELL_UI흐름 = (3, 4)

CHECKED = "■"
LIST_HEADERS = [
    "파일명", "슬라이드", "대분류", "중분류", "소분류",
    "UI명", "UI_ID", "UI유형", "UI유형_원문", "UI설명",
    "화면정의", "업무처리흐름", "기타사항",
]
FLOW_HEADERS = ["파일명", "슬라이드", "UI_ID", "UI명", "섹션", "순번", "내용"]
ELEMENT_HEADERS = ["파일명", "슬라이드", "UI_ID", "UI명", "UI유형",
                   "영역", "영역구분", "순번", "요소명"]
BUTTON_HEADERS = ["버튼명", "사용_화면수", "사용_화면"]

# UI흐름 섹션 제목 → 와이드 시트 컬럼명. 없는 제목은 기타사항으로 모은다.
SECTION_COL = {
    "화면 정의": "화면정의",
    "화면의 업무처리 흐름 정의": "업무처리흐름",
    "기타 제약조건 및 특이사항": "기타사항",
}
SECTION_화면정의 = "화면정의"

# '<영역명> : a, b, c' 를 쪼갤 때 쓰는 구분자
AREA_SEP = re.compile(r"[:：]")
# 영역 이름이 문서마다 달라서(버튼 영역 / 출력 조건 영역 / 보고서결과 영역 …)
# 키워드로 분류한다. 위에서부터 먼저 맞는 것을 쓴다.
AREA_KIND = [
    ("버튼", "버튼"),
    ("조건", "조회조건"),
    ("조회", "조회조건"),
    ("데이터", "데이터"),
    ("결과", "데이터"),
]


def cell_text(table, row, col):
    return table.cell(row, col).text.strip()


def find_shape_text(slide, name):
    for shape in slide.shapes:
        if shape.name == name and shape.has_text_frame:
            return shape.text_frame.text.strip()
    return ""


def find_spec_table(slide):
    """본문 표를 찾는다. 크기와 첫 라벨 셀로 판별한다."""
    for shape in slide.shapes:
        if not shape.has_table:
            continue
        table = shape.table
        if (len(table.rows), len(table.columns)) != TABLE_SIZE:
            continue
        if cell_text(table, 0, 0) == CELL_MAP["UI명"][2]:
            return table
    return None


def normalize_ui_type(raw):
    """'■화면  □팝업  □보고서' 에서 체크된 항목만 뽑는다."""
    picked = []
    for chunk in raw.replace("\n", " ").split():
        if chunk.startswith(CHECKED):
            picked.append(chunk[len(CHECKED):].strip())
    return "/".join(picked)


def parse_flow(table):
    """UI 흐름 셀을 문단 레벨로 읽어 [(섹션, 항목리스트)] 로 만든다.

    레벨0 문단이 섹션 제목, 레벨1 이상이 그 섹션의 항목이다.
    PowerPoint 자동 번호는 텍스트에 남지 않으므로 레벨로만 구분한다.
    """
    cell = table.cell(*CELL_UI흐름)
    sections = []
    for para in cell.text_frame.paragraphs:
        text = para.text.strip()
        if not text:
            continue
        if para.level == 0:
            sections.append((text, []))
        elif sections:
            sections[-1][1].append(text)
        else:
            sections.append(("", [text]))
    return sections


def split_items(text):
    """콤마로 나누되 괄호 안의 콤마는 무시한다.

    '전일(실적, 대비)' 를 '전일(실적' / '대비)' 로 쪼개지 않기 위함이다.
    """
    items, buf, depth = [], [], 0
    for ch in text:
        if ch in "(（":
            depth += 1
        elif ch in ")）":
            depth = max(depth - 1, 0)
        if ch == "," and depth == 0:
            items.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
    items.append("".join(buf))
    # 원본에 '기준일,  사무소' 처럼 공백이 겹친 곳이 있어 한 칸으로 정리한다.
    return [" ".join(i.split()) for i in items if i.strip()]


def classify_area(name):
    for keyword, kind in AREA_KIND:
        if keyword in name:
            return kind
    return "기타"


def parse_areas(lines):
    """화면 정의 항목을 (영역명, 영역구분, [요소]) 로 만든다."""
    areas = []
    for line in lines:
        parts = AREA_SEP.split(line, 1)
        if len(parts) == 2:
            name = " ".join(parts[0].split())
            areas.append((name, classify_area(name), split_items(parts[1])))
        else:
            areas.append(("(미분류)", "기타", [" ".join(line.split())]))
    return areas


def build_button_inventory(element_rows):
    """버튼 요소를 버튼명 기준으로 묶어 어느 화면에서 쓰는지 집계한다."""
    used = OrderedDict()
    for row in element_rows:
        if row["영역구분"] != "버튼":
            continue
        screens = used.setdefault(row["요소명"], [])
        label = f"{row['UI명']}({row['UI_ID']})"
        if label not in screens:
            screens.append(label)
    rows = [{"버튼명": name, "사용_화면수": len(screens), "사용_화면": ", ".join(screens)}
            for name, screens in used.items()]
    rows.sort(key=lambda r: (-r["사용_화면수"], r["버튼명"]))
    return rows


def extract_slide(pptx_name, slide_no, slide, table):
    sections = parse_flow(table)

    row = {h: "" for h in LIST_HEADERS}
    row["파일명"] = pptx_name
    row["슬라이드"] = slide_no
    row["대분류"] = find_shape_text(slide, SHAPE_대분류)
    row["중분류"] = find_shape_text(slide, SHAPE_중분류)
    row["소분류"] = find_shape_text(slide, SHAPE_소분류)

    warnings = []
    for field, (label_pos, value_pos, expected) in CELL_MAP.items():
        actual = cell_text(table, *label_pos)
        if actual != expected:
            warnings.append(f"슬라이드 {slide_no}: 라벨이 '{expected}' 가 아니라 '{actual}' 입니다")
        row[field] = cell_text(table, *value_pos)
    row["UI유형"] = normalize_ui_type(row["UI유형_원문"])

    기타 = []
    for title, items in sections:
        col = SECTION_COL.get(title)
        if col:
            row[col] = "\n".join(items)
        else:
            기타.extend(items)
    if 기타:
        row["기타사항"] = "\n".join(filter(None, [row["기타사항"], *기타]))

    flow_rows = []
    for title, items in sections:
        for i, item in enumerate(items, 1):
            flow_rows.append({
                "파일명": pptx_name, "슬라이드": slide_no,
                "UI_ID": row["UI_ID"], "UI명": row["UI명"],
                "섹션": title, "순번": i, "내용": item,
            })

    element_rows = []
    for title, items in sections:
        if SECTION_COL.get(title) != SECTION_화면정의:
            continue
        for area, kind, elements in parse_areas(items):
            for i, element in enumerate(elements, 1):
                element_rows.append({
                    "파일명": pptx_name, "슬라이드": slide_no,
                    "UI_ID": row["UI_ID"], "UI명": row["UI명"], "UI유형": row["UI유형"],
                    "영역": area, "영역구분": kind, "순번": i, "요소명": element,
                })
    return row, flow_rows, element_rows, warnings


def extract_file(path):
    prs = Presentation(str(path))
    list_rows, flow_rows, element_rows, warnings = [], [], [], []
    for slide_no, slide in enumerate(prs.slides, 1):
        table = find_spec_table(slide)
        if table is None:
            continue
        r, f, e, w = extract_slide(path.name, slide_no, slide, table)
        list_rows.append(r)
        flow_rows.extend(f)
        element_rows.extend(e)
        warnings.extend(w)
    return list_rows, flow_rows, element_rows, warnings


def write_csv(path, headers, rows):
    # 한국어 윈도우 엑셀에서 바로 열리도록 BOM 을 붙인다.
    with open(path, "w", encoding="utf-8-sig", newline="") as fp:
        writer = csv.DictWriter(fp, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def add_sheet(wb, title, headers, rows, widths):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in rows:
        ws.append([row[h] for h in headers])
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_xlsx(path, list_rows, flow_rows, element_rows, button_rows):
    wb = Workbook()
    wb.remove(wb.active)
    add_sheet(wb, "UI목록", LIST_HEADERS, list_rows,
              [28, 8, 16, 16, 20, 26, 14, 10, 22, 40, 50, 50, 40])
    add_sheet(wb, "UI흐름_상세", FLOW_HEADERS, flow_rows,
              [28, 8, 14, 26, 24, 6, 80])
    add_sheet(wb, "화면요소", ELEMENT_HEADERS, element_rows,
              [28, 8, 14, 26, 10, 18, 10, 6, 34])
    add_sheet(wb, "버튼인벤토리", BUTTON_HEADERS, button_rows,
              [18, 12, 70])
    wb.save(path)


def main():
    ap = argparse.ArgumentParser(description="UI 설계서 PPTX → CSV/Excel 추출기")
    ap.add_argument("source", help="PPTX 파일 또는 PPTX 가 들어 있는 폴더")
    ap.add_argument("-o", "--out", default="out", help="출력 폴더 (기본값 out)")
    args = ap.parse_args()

    source = Path(args.source)
    if source.is_dir():
        targets = sorted(p for p in source.glob("*.pptx") if not p.name.startswith("~$"))
    else:
        targets = [source]
    if not targets:
        sys.exit(f"처리할 pptx 가 없습니다. ({source})")

    list_rows, flow_rows, element_rows = [], [], []
    for target in targets:
        rows, flows, elements, warnings = extract_file(target)
        for w in warnings:
            print(f"[경고] {target.name} — {w}", file=sys.stderr)
        print(f"{target.name}: UI {len(rows)}건, 흐름 항목 {len(flows)}건, 화면요소 {len(elements)}건")
        list_rows.extend(rows)
        flow_rows.extend(flows)
        element_rows.extend(elements)
    button_rows = build_button_inventory(element_rows)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    write_csv(out / "ui_list.csv", LIST_HEADERS, list_rows)
    write_csv(out / "ui_flow.csv", FLOW_HEADERS, flow_rows)
    write_csv(out / "ui_element.csv", ELEMENT_HEADERS, element_rows)
    write_xlsx(out / "ui_spec.xlsx", list_rows, flow_rows, element_rows, button_rows)
    print(f"\n완료 → {out.resolve()}")
    print(f"  ui_list.csv    ({len(list_rows)}행)")
    print(f"  ui_flow.csv    ({len(flow_rows)}행)")
    print(f"  ui_element.csv ({len(element_rows)}행, 버튼 {len(button_rows)}종)")
    print("  ui_spec.xlsx   (UI목록 / UI흐름_상세 / 화면요소 / 버튼인벤토리)")


if __name__ == "__main__":
    main()
