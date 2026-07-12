#!/usr/bin/env python3
"""
Build conferences_<SNAPSHOT>.xlsx from conferences_<SNAPSHOT>.tsv
for the AG Schuck quantum-photonics conference tracker.

Usage: python3 build_snapshot.py YYYY-MM-DD
(defaults to today's date if omitted, expects a matching TSV to already exist)
"""

import csv
import os
import sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

OUTDIR = os.path.dirname(os.path.abspath(__file__))

COLS = [
    "id", "name", "acronym", "type", "topic_tags",
    "start_date", "end_date", "location",
    "abstract_deadline", "paper_deadline",
    "registration_deadline", "early_bird_deadline",
    "website", "relevance_score", "relevance_note",
    "status", "last_verified", "last_change", "source_urls",
]
DEADLINE_COLS = ["abstract_deadline", "paper_deadline",
                 "registration_deadline", "early_bird_deadline"]


def parse_date(s):
    if not s or s in ("TBA", "n/a"):
        return None
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def days_from_today(s, today):
    d = parse_date(s)
    if d is None:
        return None
    return (d - today).days


RED = PatternFill(fill_type="solid", fgColor="FF4444")
AMBER = PatternFill(fill_type="solid", fgColor="FFAA00")


def deadline_fill(s, today):
    n = days_from_today(s, today)
    if n is None:
        return None
    if 0 <= n <= 14:
        return RED
    if 15 <= n <= 60:
        return AMBER
    return None


HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
ARCHIVED_FILL = PatternFill(fill_type="solid", fgColor="E8E8E8")
SCORE_COLORS = {5: "D4EDDA", 4: "D1ECF1", 3: "FFF3CD", 2: "F8D7DA", 1: "F5C6CB"}

CAP = {
    "id": 28, "name": 55, "acronym": 18, "type": 16, "topic_tags": 55,
    "start_date": 13, "end_date": 13, "location": 28,
    "abstract_deadline": 17, "paper_deadline": 14,
    "registration_deadline": 20, "early_bird_deadline": 18,
    "website": 48, "relevance_score": 9, "relevance_note": 52,
    "status": 14, "last_verified": 14, "last_change": 50, "source_urls": 60,
}


def write_xlsx(events, path, today):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conferences"

    ws.append(COLS)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}1"

    col_idx = {c: i + 1 for i, c in enumerate(COLS)}

    for ev in events:
        ws.append([ev.get(c, "") for c in COLS])
        row_num = ws.max_row

        if ev.get("status") == "archived":
            for cell in ws[row_num]:
                cell.fill = ARCHIVED_FILL

        score_cell = ws.cell(row=row_num, column=col_idx["relevance_score"])
        sc = ev.get("relevance_score")
        if sc and str(sc).isdigit() and int(sc) in SCORE_COLORS:
            score_cell.fill = PatternFill(fill_type="solid", fgColor=SCORE_COLORS[int(sc)])

        for dcol in DEADLINE_COLS:
            fill = deadline_fill(str(ev.get(dcol, "")), today)
            if fill:
                ws.cell(row=row_num, column=col_idx[dcol]).fill = fill

    for col_name, cap_w in CAP.items():
        ws.column_dimensions[get_column_letter(col_idx[col_name])].width = cap_w

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=False, vertical="top")

    wb.save(path)
    print(f"XLSX -> {path}  ({len(events)} rows)")


if __name__ == "__main__":
    snapshot = sys.argv[1] if len(sys.argv) > 1 else date.today().isoformat()
    today = date.fromisoformat(snapshot)
    tsv_path = os.path.join(OUTDIR, f"conferences_{snapshot}.tsv")
    xlsx_path = os.path.join(OUTDIR, f"conferences_{snapshot}.xlsx")

    with open(tsv_path, newline="", encoding="utf-8") as f:
        events = list(csv.DictReader(f, delimiter="\t"))

    write_xlsx(events, xlsx_path, today)

    from collections import Counter
    stats = Counter(e["status"] for e in events)
    for s, n in sorted(stats.items()):
        print(f"  {s:15s}: {n}")
