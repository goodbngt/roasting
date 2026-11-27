"""데이터 폴더의 기록 파일(.csv)을 모두 읽어 로스팅일지.xlsx 를 만든다.

- 시트: 전체기록 + 일지 날짜별 시트(예: 11-04)
- 특이값(기획서 7장)
  - L값: 기준 L값 ±0.3 밖
  - Agtron: 그 제품의 전체 측정값 중간값 ±5 밖
- 읽기 어려운 값(CSV 확인필요 칸)과 빈 값은 노란색으로 칠함

실행: py 엑셀만들기.py
"""
import csv
import statistics
import sys
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "데이터"  # 이 폴더의 모든 .csv 를 합쳐서 읽음
XLSX_PATH = ROOT / "로스팅일지.xlsx"

HEADERS = ["로스팅 일자", "배치 번호", "제품명", "제품코드", "기준 L값", "L값", "Agtron값", "특이사항", "배전기명"]
WIDTHS = [12, 9, 28, 12, 9, 9, 10, 40, 10]
# CSV 확인필요 칸에 적는 이름 → 엑셀 열 번호(1부터)
CHECK_COL = {"제품명": 3, "제품코드": 4, "기준L": 5, "L": 6, "Agtron": 7}

L_RANGE = 0.3
AGTRON_RANGE = 5

YELLOW = PatternFill("solid", fgColor="FFFF00")
HEADER_FILL = PatternFill("solid", fgColor="D9D9D9")


def to_float(text):
    text = (text or "").strip()
    return float(text) if text else None


def parse_base(text):
    """기준 L값. 일지에 범위로 적힌 경우(예: 19.5~20.0)도 받음 → (표시글자, 아래값, 위값)"""
    text = (text or "").replace(" ", "")
    if not text:
        return None, None, None
    parts = text.split("~")
    low, high = float(parts[0]), float(parts[-1])
    if low == high:
        return low, low, high
    return text, low, high


def load_rows():
    rows = []
    for path in sorted(DATA_DIR.glob("*.csv")):
        with open(path, encoding="utf-8-sig", newline="") as f:
            rows.extend(csv.DictReader(f))
    for r in rows:
        r["기준L"], r["기준L_아래"], r["기준L_위"] = parse_base(r["기준L"])
        r["L"] = to_float(r["L"])
        r["Agtron"] = to_float(r["Agtron"])
        r["배치"] = int(r["배치"])
        r["확인필요"] = {c for c in (r.get("확인필요") or "").split("|") if c}
    rows.sort(key=lambda r: (r["일자"], r["배전기명"], r["배치"]))
    return rows


def mark_outliers(rows):
    agtron_by_code = defaultdict(list)
    for r in rows:
        if r["Agtron"] is not None:
            agtron_by_code[r["제품코드"]].append(r["Agtron"])
    medians = {code: statistics.median(vals) for code, vals in agtron_by_code.items()}

    for r in rows:
        notes = []
        if r["L"] is not None and r["기준L"] is not None:
            # 기준이 범위(예: 19.5~20.0)면 범위 양 끝에서 ±0.3
            below = round(r["기준L_아래"] - r["L"], 2) > L_RANGE
            above = round(r["L"] - r["기준L_위"], 2) > L_RANGE
            if below or above:
                base = r["기준L"] if isinstance(r["기준L"], str) else f"{r['기준L']:g}"
                notes.append(f"L값 기준({base}) ±{L_RANGE} 밖")
        if r["Agtron"] is not None:
            med = medians[r["제품코드"]]
            if round(abs(r["Agtron"] - med), 2) > AGTRON_RANGE:
                notes.append(f"Agtron 중간값({med:g}) ±{AGTRON_RANGE} 밖")
        if r["확인필요"] or r["L"] is None or r["Agtron"] is None:
            notes.append("판독 확인 필요(노란 칸)")
        r["특이사항"] = " / ".join(notes)


def write_sheet(ws, rows):
    ws.append(HEADERS)
    for i, w in enumerate(WIDTHS, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"

    for r in rows:
        ws.append([r["일자"], r["배치"], r["제품명"], r["제품코드"], r["기준L"],
                   r["L"], r["Agtron"], r["특이사항"], r["배전기명"]])
        row_idx = ws.max_row
        yellow_cols = {CHECK_COL[c] for c in r["확인필요"] if c in CHECK_COL}
        if r["L"] is None:
            yellow_cols.add(6)
        if r["Agtron"] is None:
            yellow_cols.add(7)
        for col in yellow_cols:
            ws.cell(row=row_idx, column=col).fill = YELLOW
        ws.cell(row=row_idx, column=6).number_format = "0.00"
        ws.cell(row=row_idx, column=7).number_format = "0.0"


def main():
    rows = load_rows()
    mark_outliers(rows)

    wb = Workbook()
    write_sheet(wb.active, rows)
    wb.active.title = "전체기록"

    by_date = defaultdict(list)
    for r in rows:
        by_date[r["일자"]].append(r)
    for date in sorted(by_date):
        write_sheet(wb.create_sheet(date[5:]), by_date[date])  # 2026-11-04 → 11-04

    try:
        wb.save(XLSX_PATH)
    except PermissionError:
        sys.exit(f"저장 실패: {XLSX_PATH.name} 이 엑셀에서 열려 있으면 닫고 다시 실행하세요.")

    outliers = sum(1 for r in rows if "밖" in r["특이사항"])
    checks = sum(1 for r in rows if "판독" in r["특이사항"])
    print(f"완료: {XLSX_PATH.name}")
    print(f"  기록 {len(rows)}줄, 날짜 {len(by_date)}일, 특이값 {outliers}줄, 판독 확인 {checks}줄")


if __name__ == "__main__":
    main()
