# UI 설계서 PPTX에서 슬라이드별 UI 정보를 뽑아 CSV/Excel로 변환하는 추출기.
"""사용법
    python ui_spec_extract.py <PPTX파일 또는 폴더> [-o 출력폴더]

출력
    <출력폴더>/ui_list.csv   UI 1개 = 1행 (와이드)
    <출력폴더>/ui_flow.csv   흐름 항목 1개 = 1행 (롱)
    <출력폴더>/ui_element.csv 화면 구성요소 1개 = 1행 (버튼 인벤토리용)
    <출력폴더>/doc_list.csv  문서(PPTX) 1개 = 1행
    <출력폴더>/ui_spec.xlsx  위를 시트로 담은 엑셀

여러 파일을 한 번에 처리하면 모든 산출물이 하나로 합쳐지며,
UI_KEY 컬럼이 문서를 가로질러 UI 하나를 가리키는 키가 된다.
"""
import argparse
import csv
import re
import sys
import traceback
from collections import OrderedDict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from pptx import Presentation

# 실행 로그에 찍어 어떤 버전이 돌았는지 확인할 수 있게 한다.
# 스크립트를 고칠 때마다 올린다.
VERSION = "2026-09-19e"

# ── 템플릿 매핑 ──────────────────────────────────────────────
# 슬라이드 상단 플레이스홀더의 도형 이름
SHAPE_대분류 = "텍스트 개체 틀 3"
SHAPE_중분류 = "텍스트 개체 틀 4"
SHAPE_소분류 = "제목 1"

# 표지 슬라이드(레이아웃 이름 '표지')의 도형 이름 → 문서 메타 필드.
# 본문 슬라이드에도 같은 이름의 도형이 있으므로 반드시 표지에서만 읽는다.
LAYOUT_표지 = "표지"
COVER_MAP = {
    "제목 1": "문서유형",
    "텍스트 개체 틀 2": "문서명",
    "텍스트 개체 틀 3": "문서번호",
    "텍스트 개체 틀 4": "버전",
    "텍스트 개체 틀 5": "작성일",
}
# 엑셀 셀에 넣을 수 없는 제어문자. 파워포인트에서 Shift+Enter 로 넣은
# 줄바꿈은 \x0b(수직 탭)로 들어오는데 openpyxl 이 이를 거부한다.
RE_제어문자 = re.compile(r"[\x00-\x08\x0e-\x1f]")
# 문서명은 'BI포탈_신용실적 – 04.여신실적(1)' 처럼 뒤에 편 이름이 붙는다.
# 공백으로 둘러싸인 대시만 구분자로 봐야 'NHNIS-BC-DS03' 이 잘리지 않는다.
SEP_문서명 = re.compile(r"\s[–—-]\s")
RE_버전 = re.compile(r"^\d+(\.\d+)*$")
RE_작성일 = re.compile(r"^\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}\.?$")

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
DOC_HEADERS = ["문서키", "파일명", "문서번호", "구분", "문서명", "문서유형",
               "버전", "작성일", "UI건수"]
DUP_HEADERS = ["UI_ID", "중복수", "UI명", "출처"]
LIST_HEADERS = [
    "UI_KEY", "파일명", "문서번호", "버전", "슬라이드", "대분류", "중분류", "소분류",
    "UI명", "UI_ID", "UI유형", "UI유형_원문", "UI설명",
    "조회조건", "입력", "버튼", "데이터", "기타영역",
    "조회조건수", "입력수", "버튼수", "데이터수",
    "업무처리흐름", "기타사항",
]
# 롱 시트는 피벗 소스로 쓰므로 분류 컬럼을 함께 실어 단독으로 집계되게 한다.
CONTEXT_HEADERS = ["UI_KEY", "문서번호", "버전", "슬라이드",
                   "대분류", "중분류", "소분류", "UI_ID", "UI명", "UI유형"]
FLOW_HEADERS = CONTEXT_HEADERS + ["섹션", "순번", "내용"]
ELEMENT_HEADERS = CONTEXT_HEADERS + ["영역", "영역구분", "순번", "요소명"]
BUTTON_HEADERS = ["버튼명", "사용_화면수", "사용_화면"]
# 화면 × 버튼 매트릭스의 고정 컬럼. 이 뒤로 버튼 이름 컬럼이 붙는다.
BTN_MATRIX_BASE = ["UI_KEY", "대분류", "소분류", "UI_ID", "UI명", "UI유형", "버튼수"]
MARK = "●"
MATRIX_BASE_WIDTHS = [42, 16, 22, 14, 30, 8, 8]

# UI흐름 섹션 제목 → 와이드 시트 컬럼명. 없는 제목은 기타사항으로 모은다.
SECTION_COL = {
    "화면 정의": "화면정의",
    "화면의 업무처리 흐름 정의": "업무처리흐름",
    "기타 제약조건 및 특이사항": "기타사항",
}
SECTION_화면정의 = "화면정의"
# 영역구분 → UI목록의 목록 컬럼 / 개수 컬럼
AREA_COL = {
    "조회조건": ("조회조건", "조회조건수"),
    "입력": ("입력", "입력수"),
    "버튼": ("버튼", "버튼수"),
    "데이터": ("데이터", "데이터수"),
    "기타": ("기타영역", None),
}

# '<영역명> : a, b, c' 를 쪼갤 때 쓰는 구분자
AREA_SEP = re.compile(r"[:：]")
# 영역 이름이 문서마다 달라서(버튼 영역 / 출력 조건 영역 / 보고서결과 영역 …)
# 키워드로 분류한다. 위에서부터 먼저 맞는 것을 쓴다.
# 위에서부터 먼저 걸리는 규칙을 쓰므로 순서가 곧 우선순위다.
# 보여 주는 영역(데이터)을 조회조건보다 먼저 두어야 '조회결과 영역' 처럼
# 두 키워드를 모두 가진 이름이 조회조건으로 잘못 빠지지 않는다.
AREA_KIND = [
    ("버튼", "버튼"),
    # 보여 주는 영역
    ("결과", "데이터"),
    ("리스트", "데이터"),
    ("목록", "데이터"),
    # 코드관리 화면의 '그룹코드 영역', '상세코드 영역' 등은 데이터 목록이다.
    ("코드", "데이터"),
    ("카테고리", "데이터"),
    ("데이터", "데이터"),
    # 걸러 보는 영역
    ("조건", "조회조건"),
    ("조회", "조회조건"),
    # 등록·입력은 값을 넣는 곳이라 걸러 보는 조회조건과 성격이 다르다.
    ("등록", "입력"),
    ("입력", "입력"),
]


def clean_text(text):
    """엑셀에 넣을 수 있는 문자만 남긴다.

    파워포인트의 줄바꿈(\x0b)과 페이지 구분(\x0c)은 줄바꿈으로 바꾸고,
    나머지 제어문자는 버린다. 이 처리를 빼면 openpyxl 이
    IllegalCharacterError 로 죽는데, 그 시점이 모든 파일을 다 읽은
    마지막 저장 단계라 작업을 통째로 날리게 된다.
    """
    text = (text.replace("\v", "\n").replace("\f", "\n")
                .replace("\r\n", "\n").replace("\r", "\n"))
    return RE_제어문자.sub("", text)


def cell_text(table, row, col):
    return clean_text(table.cell(row, col).text).strip()


def find_shape_text(slide, name):
    for shape in slide.shapes:
        if shape.name == name and shape.has_text_frame:
            return clean_text(shape.text_frame.text).strip()
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


def extract_cover(prs, fallback_name):
    """표지 슬라이드에서 문서 메타를 읽는다. 표지가 없으면 파일명으로 대체한다."""
    meta = {h: "" for h in DOC_HEADERS}
    warnings = []
    cover = next((s for s in prs.slides if s.slide_layout.name == LAYOUT_표지), None)
    if cover is None:
        warnings.append("표지 슬라이드를 찾지 못해 문서 메타를 비워 둡니다")
    else:
        for shape in cover.shapes:
            field = COVER_MAP.get(shape.name)
            if field and shape.has_text_frame:
                meta[field] = clean_text(shape.text_frame.text).strip()
    # 표지 양식이 다르면 엉뚱한 값이 들어오므로 형식을 가볍게 검사한다.
    if meta["버전"] and not RE_버전.match(meta["버전"]):
        warnings.append(f"버전 형식이 예상과 다릅니다: {meta['버전']!r}")
    if meta["작성일"] and not RE_작성일.match(meta["작성일"]):
        warnings.append(f"작성일 형식이 예상과 다릅니다: {meta['작성일']!r}")
    if not meta["문서번호"]:
        meta["문서번호"] = Path(fallback_name).stem
        warnings.append("문서번호가 비어 있어 파일명으로 대체했습니다")
    # 한 문서번호 아래 여러 편으로 나뉘어 있으므로 편 이름까지 있어야 구분된다.
    parts = SEP_문서명.split(meta["문서명"])
    meta["구분"] = (parts[-1].strip() if len(parts) > 1
                  else meta["문서명"].strip() or Path(fallback_name).stem)
    return meta, warnings


def make_doc_key(meta, seen_docs):
    """문서 하나를 가리키는 키. 예) NHNIS-BC-DS03-v0.7

    같은 문서번호·버전이 두 번 들어오면(같은 파일을 두 번 받은 경우 등)
    뒤에 #2 를 붙여 키가 겹치지 않게 한다. 키가 겹치면 시트 간 조인이
    조용히 어긋나므로, 경고만 내고 넘어가지 않는다.
    """
    base = "-".join(filter(None, [meta["문서번호"], meta["구분"],
                                  f"v{meta['버전']}" if meta["버전"] else ""]))
    key, n = base, 1
    while key in seen_docs:
        n += 1
        key = f"{base}#{n}"
    seen_docs.add(key)
    return key


def make_ui_key(meta, slide_no):
    """문서를 가로질러 UI 하나를 가리키는 키. 예) NHNIS-BC-DS03-v0.7-s4"""
    return f"{meta['문서키']}-s{slide_no}"


def build_duplicates(list_rows):
    """같은 UI_ID 가 여러 문서·슬라이드에 나타나는 경우를 모은다."""
    found = OrderedDict()
    for row in list_rows:
        found.setdefault(row["UI_ID"], []).append(row)
    dups = []
    for ui_id, rows in found.items():
        if len(rows) < 2:
            continue
        names = sorted({r["UI명"] for r in rows})
        dups.append({
            "UI_ID": ui_id,
            "중복수": len(rows),
            "UI명": " / ".join(names),
            "출처": ", ".join(r["UI_KEY"] for r in rows),
        })
    dups.sort(key=lambda r: (-r["중복수"], r["UI_ID"]))
    return dups


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
        text = clean_text(para.text).strip()
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
    """화면 정의 항목을 [영역명, 영역구분, [요소]] 로 만든다.

    콜론이 없는 줄은 앞 영역이 다음 줄로 이어진 것으로 본다. 작성자가
    Shift+Enter 대신 Enter 를 눌러 항목이 문단째로 나뉜 경우인데,
    새 영역으로 세면 이름 없는 '(미분류)' 가 쌓인다.
    """
    areas = []
    for line in lines:
        parts = AREA_SEP.split(line, 1)
        if len(parts) == 2:
            name = " ".join(parts[0].split())
            areas.append([name, classify_area(name), split_items(parts[1])])
        elif areas:
            areas[-1][2].extend(split_items(line))
        else:
            areas.append(["(미분류)", "기타", split_items(line)])
    return areas


def build_button_matrix(list_rows, element_rows, button_rows):
    """화면 × 버튼 매트릭스를 만든다.

    버튼 컬럼은 많이 쓰이는 순으로 왼쪽에 둔다. 그래야 왼쪽 덩어리가
    공통 버튼 세트가 되고, 거기 생긴 구멍이 곧 '이 화면엔 공통 버튼이
    빠졌다' 는 신호로 눈에 들어온다.
    """
    used = {}
    for element in element_rows:
        if element["영역구분"] == "버튼":
            used.setdefault(element["UI_KEY"], set()).add(element["요소명"])

    # 버튼 이름이 고정 컬럼 이름과 겹치면 헤더가 중복돼 엑셀 표가 깨진다.
    taken, columns = set(BTN_MATRIX_BASE), []
    for button in button_rows:
        name = button["버튼명"]
        header = name
        while header in taken:
            header += "_"
        taken.add(header)
        columns.append((name, header))

    headers = BTN_MATRIX_BASE + [header for _, header in columns]
    rows = []
    for row in list_rows:
        have = used.get(row["UI_KEY"], set())
        matrix_row = {h: row[h] for h in BTN_MATRIX_BASE}
        for name, header in columns:
            matrix_row[header] = MARK if name in have else ""
        rows.append(matrix_row)
    return headers, rows


def build_unknown_areas(element_rows):
    """분류 규칙이 잡지 못한 영역 이름을 많이 나온 순으로 모은다."""
    counts = OrderedDict()
    for row in element_rows:
        if row["영역구분"] == "기타":
            counts[row["영역"]] = counts.get(row["영역"], 0) + 1
    return sorted(counts.items(), key=lambda kv: -kv[1])


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


def extract_slide(pptx_name, meta, slide_no, slide, table):
    sections = parse_flow(table)
    ui_key = make_ui_key(meta, slide_no)

    row = {h: "" for h in LIST_HEADERS}
    row["UI_KEY"] = ui_key
    row["파일명"] = pptx_name
    row["문서번호"] = meta["문서번호"]
    row["버전"] = meta["버전"]
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

    # 화면 정의를 뺀 나머지 섹션을 해당 컬럼에 담는다.
    기타 = []
    for title, items in sections:
        col = SECTION_COL.get(title)
        if col == SECTION_화면정의:
            continue
        if col:
            row[col] = "\n".join(items)
        else:
            기타.extend(items)
    if 기타:
        row["기타사항"] = "\n".join(filter(None, [row["기타사항"], *기타]))

    # 롱 시트가 단독으로 피벗되도록 분류 컬럼을 그대로 복사해 붙인다.
    context = {h: row[h] for h in CONTEXT_HEADERS}

    flow_rows = []
    for title, items in sections:
        for i, item in enumerate(items, 1):
            flow_rows.append({**context, "섹션": title, "순번": i, "내용": item})

    element_rows = []
    for title, items in sections:
        if SECTION_COL.get(title) != SECTION_화면정의:
            continue
        for area, kind, elements in parse_areas(items):
            for i, element in enumerate(elements, 1):
                element_rows.append({**context, "영역": area, "영역구분": kind,
                                     "순번": i, "요소명": element})

    # 화면 정의를 영역구분별로 묶어 UI목록에서 한 줄로 훑어볼 수 있게 한다.
    by_kind = OrderedDict()
    for element in element_rows:
        by_kind.setdefault(element["영역구분"], []).append(element["요소명"])
    for kind, (list_col, count_col) in AREA_COL.items():
        picked = by_kind.get(kind, [])
        row[list_col] = ", ".join(picked)
        if count_col:
            row[count_col] = len(picked)

    return row, flow_rows, element_rows, warnings


def extract_file(path, seen_docs):
    prs = Presentation(str(path))
    meta, warnings = extract_cover(prs, path.name)
    meta["파일명"] = path.name
    meta["문서키"] = make_doc_key(meta, seen_docs)
    if meta["문서키"].count("#"):
        warnings.append(f"문서키가 앞선 파일과 겹쳐 {meta['문서키']} 로 구분했습니다. "
                        "표지의 문서명이 같은 파일이 둘 이상인지 확인하세요")
    list_rows, flow_rows, element_rows = [], [], []
    for slide_no, slide in enumerate(prs.slides, 1):
        table = find_spec_table(slide)
        if table is None:
            continue
        r, f, e, w = extract_slide(path.name, meta, slide_no, slide, table)
        list_rows.append(r)
        flow_rows.extend(f)
        element_rows.extend(e)
        warnings.extend(w)
    meta["UI건수"] = len(list_rows)
    return meta, list_rows, flow_rows, element_rows, warnings


def write_csv(path, headers, rows):
    # 한국어 윈도우 엑셀에서 바로 열리도록 BOM 을 붙인다.
    with open(path, "w", encoding="utf-8-sig", newline="") as fp:
        writer = csv.DictWriter(fp, fieldnames=headers)
        writer.writeheader()
        writer.writerows(rows)


def add_sheet(wb, title, table_name, headers, rows, widths, wrap=True):
    ws = wb.create_sheet(title)
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in rows:
        # 읽는 쪽에서 이미 정리하지만, 저장 단계에서 죽으면 전체를 날리므로
        # 여기서 한 번 더 막는다.
        ws.append([clean_text(row[h]) if isinstance(row[h], str) else row[h]
                   for h in headers])
    for i, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = width
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=wrap)
    ws.freeze_panes = "A2"
    if rows:
        # 엑셀 '표' 로 만들어 두면 피벗테이블을 만들 때 범위가 자동으로 잡히고,
        # 행이 늘어도 범위를 다시 지정할 필요가 없다.
        table = Table(displayName=table_name, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2",
                                              showRowStripes=True)
        ws.add_table(table)
    else:
        ws.auto_filter.ref = ws.dimensions


def add_matrix_sheet(wb, title, table_name, headers, rows, base_count):
    """고정 컬럼 + 표시 컬럼으로 된 매트릭스 시트."""
    ws = wb.create_sheet(title)
    ws.append(headers)
    for i, cell in enumerate(ws[1], 1):
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4472C4")
        # 버튼 이름은 세로로 세워야 컬럼을 좁게 두고도 다 읽힌다.
        rotate = 90 if i > base_count else 0
        cell.alignment = Alignment(horizontal="center",
                                   vertical="bottom" if rotate else "center",
                                   textRotation=rotate)
    for row in rows:
        ws.append([row[h] for h in headers])

    ws.row_dimensions[1].height = 95
    for i in range(1, len(headers) + 1):
        letter = get_column_letter(i)
        ws.column_dimensions[letter].width = (
            MATRIX_BASE_WIDTHS[i - 1] if i <= base_count else 4.2)

    mark_fill = PatternFill("solid", fgColor="9DC3E6")
    for row in ws.iter_rows(min_row=2):
        for i, cell in enumerate(row, 1):
            if i <= base_count:
                cell.alignment = Alignment(vertical="center")
                continue
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if cell.value == MARK:
                cell.fill = mark_fill

    ws.freeze_panes = f"{get_column_letter(base_count + 1)}2"
    if rows:
        table = Table(displayName=table_name, ref=ws.dimensions)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleLight1",
                                              showRowStripes=False)
        ws.add_table(table)


def write_xlsx(path, doc_rows, list_rows, matrix, flow_rows, element_rows,
               button_rows, dup_rows):
    wb = Workbook()
    wb.remove(wb.active)
    context_widths = [42, 18, 8, 8, 16, 16, 20, 14, 26, 10]
    add_sheet(wb, "문서", "t_doc", DOC_HEADERS, doc_rows,
              [42, 46, 18, 20, 32, 14, 8, 14, 10])
    _ = MATRIX_BASE_WIDTHS  # 매트릭스 시트가 쓰는 고정 컬럼 폭
    add_sheet(wb, "UI목록", "t_ui", LIST_HEADERS, list_rows,
              [42, 46, 18, 8, 8, 16, 16, 20, 26, 14, 10, 22, 40,
               44, 44, 44, 44, 24, 10, 8, 8, 8, 50, 40])
    matrix_headers, matrix_rows = matrix
    add_matrix_sheet(wb, "화면별버튼", "t_btn_ui", matrix_headers, matrix_rows,
                     len(BTN_MATRIX_BASE))
    # 피벗 소스 두 장은 줄바꿈 없이 한 줄로 둬야 스크롤하며 훑기 좋다.
    add_sheet(wb, "화면요소", "t_element", ELEMENT_HEADERS, element_rows,
              context_widths + [18, 10, 6, 34], wrap=False)
    add_sheet(wb, "UI흐름_상세", "t_flow", FLOW_HEADERS, flow_rows,
              context_widths + [24, 6, 80], wrap=False)
    add_sheet(wb, "버튼인벤토리", "t_button", BUTTON_HEADERS, button_rows,
              [18, 12, 70])
    add_sheet(wb, "중복점검", "t_dup", DUP_HEADERS, dup_rows,
              [16, 8, 40, 90])
    wb.save(path)


def make_console_safe():
    """윈도우 기본 콘솔(cp949)에서 표현 못 하는 글자 때문에 죽지 않게 한다."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except (AttributeError, OSError):
            pass


def main():
    make_console_safe()
    ap = argparse.ArgumentParser(description="UI 설계서 PPTX → CSV/Excel 추출기")
    ap.add_argument("source", help="PPTX 파일 또는 PPTX 가 들어 있는 폴더")
    ap.add_argument("-o", "--out", default="out", help="출력 폴더 (기본값 out)")
    ap.add_argument("-r", "--recursive", action="store_true", help="하위 폴더까지 훑기")
    args = ap.parse_args()

    source = Path(args.source)
    if source.is_dir():
        found = source.rglob("*.pptx") if args.recursive else source.glob("*.pptx")
        targets = sorted(p for p in found if not p.name.startswith("~$"))
    else:
        targets = [source]
    if not targets:
        sys.exit(f"처리할 pptx 가 없습니다. ({source})")

    print(f"ui_spec_extract {VERSION}")
    print(f"대상 {len(targets)}개 파일")
    doc_rows, list_rows, flow_rows, element_rows = [], [], [], []
    seen_docs = set()
    failed = []
    for i, target in enumerate(targets, 1):
        # 어떤 파일을 붙들고 있는지 먼저 알려야 멈춘 것과 느린 것을 구분할 수 있다.
        # 구글드라이브 스트리밍 폴더는 첫 읽기에서 내려받느라 오래 걸린다.
        size_mb = target.stat().st_size / 1024 / 1024
        print(f"  [{i}/{len(targets)}] {target.name} ({size_mb:.1f}MB) 처리 중 ...",
              end="", flush=True)
        try:
            meta, rows, flows, elements, warnings = extract_file(target, seen_docs)
        except Exception as exc:  # 한 파일이 깨져도 나머지는 계속 처리한다
            print(" 실패")
            print(f"[오류] {target.name} — 건너뜁니다", file=sys.stderr)
            print(f"       {type(exc).__name__}: {exc}", file=sys.stderr)
            traceback.print_exc(file=sys.stderr)
            failed.append(target.name)
            continue
        print(f" UI {len(rows)}건, 흐름 {len(flows)}건, 요소 {len(elements)}건")
        for w in warnings:
            print(f"[경고] {target.name} — {w}", file=sys.stderr)
        doc_rows.append(meta)
        list_rows.extend(rows)
        flow_rows.extend(flows)
        element_rows.extend(elements)
    button_rows = build_button_inventory(element_rows)
    dup_rows = build_duplicates(list_rows)
    unknown_areas = build_unknown_areas(element_rows)
    matrix = build_button_matrix(list_rows, element_rows, button_rows)

    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    write_csv(out / "doc_list.csv", DOC_HEADERS, doc_rows)
    write_csv(out / "ui_list.csv", LIST_HEADERS, list_rows)
    write_csv(out / "ui_flow.csv", FLOW_HEADERS, flow_rows)
    write_csv(out / "ui_element.csv", ELEMENT_HEADERS, element_rows)
    write_xlsx(out / "ui_spec.xlsx", doc_rows, list_rows, matrix, flow_rows,
               element_rows, button_rows, dup_rows)
    print(f"\n완료 → {out.resolve()}")
    print(f"  doc_list.csv   ({len(doc_rows)}행)")
    print(f"  ui_list.csv    ({len(list_rows)}행)")
    print(f"  ui_flow.csv    ({len(flow_rows)}행)")
    print(f"  ui_element.csv ({len(element_rows)}행, 버튼 {len(button_rows)}종)")
    print("  ui_spec.xlsx   (문서 / UI목록 / 화면별버튼 / 화면요소 / UI흐름_상세 / "
          "버튼인벤토리 / 중복점검)")
    if unknown_areas:
        total = sum(n for _, n in unknown_areas)
        print(f"\n[확인 필요] 분류하지 못한 영역 {len(unknown_areas)}종 / {total}건 — "
              "아래 이름을 알려주시면 규칙에 넣겠습니다.")
        for name, count in unknown_areas:
            print(f"    {count:5d}건  {name!r}")
    if dup_rows:
        print(f"\n[확인 필요] 같은 UI_ID 가 여러 곳에 있습니다 — {len(dup_rows)}종. 중복점검 시트를 보세요.")
    if failed:
        print(f"\n[확인 필요] 처리하지 못한 파일 {len(failed)}개 — {', '.join(failed)}")


if __name__ == "__main__":
    main()
