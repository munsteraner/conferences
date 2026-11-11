#!/usr/bin/env python3
"""
AG Schuck quantum-photonics conference tracker - snapshot generator.

Incremental build:
  1. load the latest previous conferences_YYYY-MM-DD.tsv (the merged snapshot)
  2. apply this run's changes from a JSON file (updates + new rows)
  3. auto-archive passed events, fix stale statuses
  4. write conferences_<SNAPSHOT>.tsv and .xlsx
  5. run the self-check (Step 7 of the routine) and exit non-zero on failure

Usage:
  python3 build_snapshot.py --snapshot 2026-11-11 --changes run_changes.json

run_changes.json format:
  {"updates": [{"id": ..., "fields": {...}, "last_change": "...",
                "source_urls": [...], "verified": true}],
   "new_rows": [{<19 columns minus last_verified>}]}
"""

import argparse, csv, glob, json, os, re, sys
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import FormulaRule

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
DATE_COLS = ["start_date", "end_date"] + DEADLINE_COLS
ACTIVE = {"upcoming", "cfp_open", "cfp_closed", "ongoing", "candidate"}
STATUSES = ACTIVE | {"archived", "cancelled"}
TYPES = {"Conference", "Workshop", "Summer School", "Winter School",
         "Symposium", "Seminar"}


def parse_date(s):
    try:
        return date.fromisoformat(s)
    except (TypeError, ValueError):
        return None


def latest_prev(snapshot):
    files = sorted(f for f in glob.glob(os.path.join(OUTDIR, "conferences_*.tsv"))
                   if os.path.basename(f) != f"conferences_{snapshot}.tsv")
    if not files:
        sys.exit("no previous snapshot found")
    return files[-1]


def load_tsv(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f, delimiter="\t"))


def apply_changes(rows, changes, snapshot):
    byid = {r["id"]: r for r in rows}
    # id corrections (e.g. placeholder had the wrong edition year); row kept
    for old, new_id in changes.get("renames", {}).items():
        if old in byid and new_id not in byid:
            byid[new_id] = byid.pop(old)
            byid[new_id]["id"] = new_id
    for u in changes.get("updates", []):
        u = dict(u, id=changes.get("renames", {}).get(u["id"], u["id"]))
        r = byid.get(u["id"])
        if r is None:
            print(f"WARN: update for unknown id {u['id']}", file=sys.stderr)
            continue
        for k, v in u.get("fields", {}).items():
            if k in COLS and k != "id":
                r[k] = str(v)
        if u.get("last_change"):
            r["last_change"] = u["last_change"]
        if u.get("verified", True):
            r["last_verified"] = snapshot
        srcs = [s for s in r.get("source_urls", "").split(";") if s]
        for s in u.get("source_urls", []):
            if s and s not in srcs:
                srcs.append(s)
        r["source_urls"] = ";".join(srcs)
    for n in changes.get("new_rows", []):
        if n["id"] in byid:
            print(f"WARN: new row id {n['id']} already exists - skipped",
                  file=sys.stderr)
            continue
        row = {c: str(n.get(c, "TBA")) for c in COLS}
        row["last_verified"] = n.get("last_verified", snapshot)
        if isinstance(n.get("source_urls"), list):
            row["source_urls"] = ";".join(n["source_urls"])
        rows.append(row)
        byid[row["id"]] = row
    return rows


def fix_statuses(rows, today, snapshot):
    for r in rows:
        end = parse_date(r["end_date"]) or parse_date(r["start_date"])
        start = parse_date(r["start_date"])
        if r["status"] in ACTIVE and end and end < today:
            r["status"] = "archived"
            r["last_change"] = f"event ended {end.isoformat()} -> archived"
            r["last_verified"] = snapshot
        elif (r["status"] in ACTIVE - {"candidate", "ongoing"}
              and start and end and start <= today <= end):
            r["status"] = "ongoing"
            r["last_change"] = f"event running ({start}..{end}) -> ongoing"
    return rows


# ── XLSX ─────────────────────────────────────────────────────────────────────
HEADER_FILL = PatternFill(fill_type="solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=10)
ARCHIVED_FILL = PatternFill(fill_type="solid", fgColor="E8E8E8")
RED = PatternFill(fill_type="solid", fgColor="FF6B6B", bgColor="FF6B6B")
AMBER = PatternFill(fill_type="solid", fgColor="FFC04D", bgColor="FFC04D")
SCORE_COLORS = {5: "D4EDDA", 4: "D1ECF1", 3: "FFF3CD", 2: "F8D7DA", 1: "F5C6CB"}


def static_fill(s, today):
    d = parse_date(s)
    if d is None:
        return None
    n = (d - today).days
    if 0 <= n <= 14:
        return RED
    if 15 <= n <= 60:
        return AMBER
    return None


def write_xlsx(rows, path, today):
    wb = Workbook()
    ws = wb.active
    ws.title = "Conferences"
    ws.append(COLS)
    for cell in ws[1]:
        cell.fill, cell.font = HEADER_FILL, HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center",
                                   wrap_text=True)
    ws.row_dimensions[1].height = 28
    ws.freeze_panes = "A2"
    col_idx = {c: i + 1 for i, c in enumerate(COLS)}
    for r in rows:
        vals = []
        for c in COLS:
            v = r[c]
            d = parse_date(v) if c in DATE_COLS else None
            vals.append(d if d else (int(v) if c == "relevance_score" and v.isdigit() else v))
        ws.append(vals)
        rn = ws.max_row
        for c in DATE_COLS:
            ws.cell(rn, col_idx[c]).number_format = "yyyy-mm-dd"
        if r["status"] == "archived":
            for cell in ws[rn]:
                cell.fill = ARCHIVED_FILL
        sc = r["relevance_score"]
        if sc.isdigit() and int(sc) in SCORE_COLORS:
            ws.cell(rn, col_idx["relevance_score"]).fill = PatternFill(
                fill_type="solid", fgColor=SCORE_COLORS[int(sc)])
        # static colouring as of snapshot date (for viewers without formula eval)
        if r["status"] not in ("archived", "cancelled"):
            for c in DEADLINE_COLS:
                f = static_fill(r[c], today)
                if f:
                    ws.cell(rn, col_idx[c]).fill = f
    last = ws.max_row
    # live conditional formatting relative to TODAY() when the file is opened
    for c in DEADLINE_COLS:
        L = get_column_letter(col_idx[c])
        rng = f"{L}2:{L}{last}"
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"AND(ISNUMBER({L}2),{L}2-TODAY()>=0,{L}2-TODAY()<=14)"],
            fill=RED, stopIfTrue=True))
        ws.conditional_formatting.add(rng, FormulaRule(
            formula=[f"AND(ISNUMBER({L}2),{L}2-TODAY()>14,{L}2-TODAY()<=60)"],
            fill=AMBER, stopIfTrue=True))
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{last}"
    # autosize columns (content length, capped)
    for c in COLS:
        width = max(len(c), *(len(str(r[c])) for r in rows)) + 2
        ws.column_dimensions[get_column_letter(col_idx[c])].width = min(max(width, 10), 60)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=False, vertical="top")
    wb.save(path)


def write_tsv(rows, path):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=COLS, delimiter="\t",
                           extrasaction="ignore", lineterminator="\n")
        w.writeheader()
        for r in rows:
            w.writerow({c: str(r[c]).replace("\t", " ").replace("\n", " ") for c in COLS})


# ── self-check ───────────────────────────────────────────────────────────────
def self_check(path, prev_count, today):
    errs = []
    with open(path, encoding="utf-8") as f:
        lines = f.read().rstrip("\n").split("\n")
    for i, ln in enumerate(lines):
        if len(ln.split("\t")) != 19:
            errs.append(f"line {i+1}: {len(ln.split(chr(9)))} columns")
    rows = load_tsv(path)
    ids = [r["id"] for r in rows]
    dups = {i for i in ids if ids.count(i) > 1}
    if dups:
        errs.append(f"duplicate ids: {sorted(dups)}")
    for r in rows:
        for c in DATE_COLS:
            v = r[c]
            if v not in ("TBA", "n/a") and not re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
                errs.append(f"{r['id']}: bad date in {c}: {v!r}")
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", r["last_verified"]):
            errs.append(f"{r['id']}: bad last_verified {r['last_verified']!r}")
        if r["status"] not in STATUSES:
            errs.append(f"{r['id']}: bad status {r['status']!r}")
        if r["type"] not in TYPES:
            errs.append(f"{r['id']}: bad type {r['type']!r}")
        end = parse_date(r["end_date"])
        if end and end < today and r["status"] in ACTIVE:
            errs.append(f"{r['id']}: ended {end} but status {r['status']}")
    if len(rows) < prev_count:
        errs.append(f"row count {len(rows)} < previous {prev_count}")
    return rows, errs


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--snapshot", required=True)
    ap.add_argument("--changes")
    a = ap.parse_args()
    today = date.fromisoformat(a.snapshot)
    prev_path = latest_prev(a.snapshot)
    rows = load_tsv(prev_path)
    prev_count = len(rows)
    changes = json.load(open(a.changes)) if a.changes else {}
    rows = apply_changes(rows, changes, a.snapshot)
    rows = fix_statuses(rows, today, a.snapshot)
    tsv = os.path.join(OUTDIR, f"conferences_{a.snapshot}.tsv")
    write_tsv(rows, tsv)
    write_xlsx(rows, os.path.join(OUTDIR, f"conferences_{a.snapshot}.xlsx"), today)
    _, errs = self_check(tsv, prev_count, today)
    print(f"previous: {os.path.basename(prev_path)} ({prev_count} rows) -> "
          f"{len(rows)} rows")
    from collections import Counter
    for s, n in sorted(Counter(r["status"] for r in rows).items()):
        print(f"  {s:12s} {n}")
    if errs:
        print("SELF-CHECK FAILED:")
        for e in errs:
            print("  " + e)
        sys.exit(1)
    print("self-check: OK")


if __name__ == "__main__":
    main()
