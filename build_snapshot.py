#!/usr/bin/env python3
"""
Build conferences_2026-11-01.tsv  and  conferences_2026-11-01.xlsx
for AG Schuck quantum-photonics conference tracker.
"""

import csv, os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

TODAY     = date(2026, 11, 1)
SNAPSHOT  = "2026-11-01"
OUTDIR    = "/home/user/conferences"

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

# ── Event data ────────────────────────────────────────────────────────────────
# Incremental build: start from the previous snapshot's TSV, then apply this
# run's UPDATES (field changes on existing ids) and NEW rows.
PREV_SNAPSHOT = "2026-10-18"

def load_prev():
    with open(os.path.join(OUTDIR, f"conferences_{PREV_SNAPSHOT}.tsv"),
              newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f, delimiter="\t"))

UPDATES = {
    "it-fab-school-2026": dict(status="archived",
        last_change="event ended 2026-10-30 -> archived", src=None),
    "ofc-2027": dict(status="cfp_closed",
        last_change="abstract deadline 2026-10-20 passed -> cfp_closed (post-deadline 2027-02-23 still possible)", src=None),
    "aps-summit-2027": dict(status="cfp_closed",
        last_change="abstract deadline 2026-10-22 passed -> cfp_closed", src=None),
    "spw-2027": dict(start_date="TBA", end_date="TBA",
        last_change="dates reset to TBA: SPW appears biennial (2024 Edinburgh, 2026 Naples), so earlier 2027-11 dates look unverified; next edition probably 2028 - verify on official site",
        src="https://www.fbk.eu/en/event/37290/single-photon-workshop-2026"),
    "cleo-2027": dict(
        last_change="listings (cantonfair.net, laserfocusworld.com) give Long Beach 2027-05-02..07; unverified third-party hint of submission deadline ~2026-11-23 (not entered; official cleoconference.org not confirmed)",
        src="https://www.laserfocusworld.com/events/event/55361294/optica-cleo-2027"),
    "qps-2027": dict(
        last_change="placeholder; no 2027 announcement found (searched 2026-11-01)", src=None),
    "ecio-2027": dict(
        last_change="placeholder; no 2027 announcement found (searched 2026-11-01); ECIO 2026 was Zurich Jun 15-17", src=None),
}

NEW = []

def build_events():
    rows = load_prev()
    byid = {r["id"]: r for r in rows}
    for id_, upd in UPDATES.items():
        r = byid[id_]
        src = upd.pop("src", None)
        r.update(upd)
        r["last_verified"] = SNAPSHOT
        if src and src not in r["source_urls"]:
            r["source_urls"] += ";" + src
    for n in NEW:
        n = dict(n); n["last_verified"] = SNAPSHOT
        rows.append(n)
    return rows

EVENTS = build_events()



# ── helpers ───────────────────────────────────────────────────────────────────
def parse_date(s):
    """Return date object or None for TBA/n/a."""
    if not s or s in ("TBA", "n/a"):
        return None
    try:
        return date.fromisoformat(s)
    except ValueError:
        return None


def days_from_today(s):
    d = parse_date(s)
    if d is None:
        return None
    return (d - TODAY).days


RED   = PatternFill(fill_type="solid", fgColor="FF4444")
AMBER = PatternFill(fill_type="solid", fgColor="FFAA00")

def deadline_fill(s):
    """Return RED/AMBER fill or None based on days until deadline."""
    n = days_from_today(s)
    if n is None:
        return None
    if 0 <= n <= 14:
        return RED
    if 15 <= n <= 60:
        return AMBER
    return None


# ── write TSV ─────────────────────────────────────────────────────────────────
def write_tsv(events, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter="\t",
                           extrasaction="ignore")
        w.writeheader()
        for e in events:
            w.writerow(e)
    print(f"  TSV  -> {path}  ({len(events)} rows)")


# ── write XLSX ────────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill(fill_type="solid", fgColor="1F4E79")
HEADER_FONT  = Font(bold=True, color="FFFFFF", size=10)
ARCHIVED_FILL = PatternFill(fill_type="solid", fgColor="E8E8E8")
SCORE_COLORS = {5:"D4EDDA", 4:"D1ECF1", 3:"FFF3CD", 2:"F8D7DA", 1:"F5C6CB"}

def write_xlsx(events, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conferences"

    # header
    ws.append(COLS)
    for cell in ws[1]:
        cell.fill   = HEADER_FILL
        cell.font   = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)

    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}1"

    col_idx = {c: i+1 for i, c in enumerate(COLS)}

    for ev in events:
        row_vals = [ev[c] for c in COLS]
        ws.append(row_vals)
        row_num = ws.max_row

        # light-grey background for archived rows
        is_archived = ev["status"] == "archived"
        if is_archived:
            for cell in ws[row_num]:
                cell.fill = ARCHIVED_FILL

        # relevance score colour on that cell
        score_cell = ws.cell(row=row_num, column=col_idx["relevance_score"])
        sc = ev.get("relevance_score")
        if sc and str(sc).isdigit() and int(sc) in SCORE_COLORS:
            score_cell.fill = PatternFill(fill_type="solid",
                                          fgColor=SCORE_COLORS[int(sc)])

        # deadline columns — conditional colour
        for dcol in DEADLINE_COLS:
            val = ev.get(dcol, "")
            fill = deadline_fill(str(val))
            if fill:
                ws.cell(row=row_num, column=col_idx[dcol]).fill = fill

    # auto-size columns (capped)
    CAP = {
        "id": 28, "name": 55, "acronym": 18, "type": 16, "topic_tags": 55,
        "start_date": 13, "end_date": 13, "location": 28,
        "abstract_deadline": 17, "paper_deadline": 14,
        "registration_deadline": 20, "early_bird_deadline": 18,
        "website": 48, "relevance_score": 9, "relevance_note": 52,
        "status": 14, "last_verified": 14, "last_change": 50, "source_urls": 60,
    }
    for col_name, cap_w in CAP.items():
        ci = col_idx[col_name]
        ws.column_dimensions[get_column_letter(ci)].width = cap_w

    # wrap text in long columns
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=False, vertical="top")

    wb.save(path)
    print(f"  XLSX -> {path}")


# ── main ──────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    tsv_path  = os.path.join(OUTDIR, f"conferences_{SNAPSHOT}.tsv")
    xlsx_path = os.path.join(OUTDIR, f"conferences_{SNAPSHOT}.xlsx")

    write_tsv(EVENTS, tsv_path)
    write_xlsx(EVENTS, xlsx_path)
    print(f"\nDone. {len(EVENTS)} events written.")

    # summary stats
    from collections import Counter
    stats = Counter(e["status"] for e in EVENTS)
    for s, n in sorted(stats.items()):
        print(f"  {s:15s}: {n}")
