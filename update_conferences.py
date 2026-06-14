#!/usr/bin/env python3
"""
Conference list updater — 2026-06-14 run.
Reads conferences_2026-05-13.tsv, applies changes, writes:
  conferences_2026-06-14.tsv
  conferences_2026-06-14.xlsx
"""

import csv
import copy
from datetime import date, datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.filters import AutoFilter

TODAY = "2026-06-14"
INPUT  = "conferences_2026-05-13.tsv"
OUT_TSV  = "conferences_2026-06-14.tsv"
OUT_XLSX = "conferences_2026-06-14.xlsx"

COLS = [
    "id","name","acronym","type","topic_tags",
    "start_date","end_date","location",
    "abstract_deadline","paper_deadline",
    "registration_deadline","early_bird_deadline",
    "website","relevance_score","relevance_note",
    "status","last_verified","last_change","source_urls"
]

# ---------------------------------------------------------------------------
# 1.  Load existing rows
# ---------------------------------------------------------------------------
with open(INPUT, newline="", encoding="utf-8") as f:
    reader = csv.DictReader(f, delimiter="\t")
    rows = [dict(r) for r in reader]

# Index by id for easy lookup
idx = {r["id"]: i for i, r in enumerate(rows)}

def patch(row_id, **kwargs):
    """Apply field patches to a row; always stamp last_verified."""
    i = idx[row_id]
    rows[i].update(kwargs)
    rows[i]["last_verified"] = TODAY

# ---------------------------------------------------------------------------
# 2.  Apply changes
# ---------------------------------------------------------------------------

# ── Events that became ARCHIVED since May 13 ────────────────────────────────

patch("cleo-2026",
      status="archived",
      last_change="archived: CLEO 2026 ran May 17–21 Charlotte NC; next edition CLEO 2027 Long Beach announced")

patch("damop-2026",
      status="archived",
      last_change="archived: DAMOP 2026 ran Jun 1–5 Providence RI")

patch("spie-pfq-2026",
      status="archived",
      last_change="archived: PfQ ran Jun 8–11 Waterloo ON; venue was confirmed Waterloo (not Belfast)")

# ── DATE CORRECTION: GRC Mechanical Systems was January 2026, not July ──────
patch("grc-mech-quant-2026",
      start_date="2026-01-18",
      end_date="2026-01-23",
      status="archived",
      last_change="DATE CORRECTED: event was Jan 18–23 2026 (previously listed as Jul 26–31); now archived",
      source_urls="https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/;https://rakichlab.yale.edu/posts/2026-01-30-cavity-optomechanics-subgroup-attended-the-grc")

# ── Events now ONGOING (started today or this week) ─────────────────────────

patch("cargese-qt-2026",
      status="ongoing",
      last_change="ongoing Jun 8–20 Cargèse (Corsica)")

patch("islc-2026",
      status="ongoing",
      abstract_deadline="2026-02-05",
      last_change="ongoing today Jun 14–17 Tampere; abstract deadline was extended to Feb 5 (closed)")

patch("grc-quantum-science-2026",
      status="ongoing",
      registration_deadline="2026-06-28",
      last_change="ongoing Jun 14–19 Easton MA; GRC application deadline 2026-06-28")

patch("icap-2026",
      status="ongoing",
      last_change="ongoing Jun 14–19 Wuhan")

patch("ecio-2026",
      status="ongoing",
      last_change="ongoing Jun 15–17 ETH Zürich; 253 submitted papers; early-bird passed")

patch("optica-quantum-2026",
      status="ongoing",
      early_bird_deadline="2026-05-01",
      last_change="ongoing Jun 15–18 Glasgow; early-bird deadline was May 1; standard registration applied")

patch("icop-2026",
      status="ongoing",
      last_change="ongoing Jun 15–17 L'Aquila")

# ── cfp_closed (registration / application deadline now past) ────────────────

patch("lake-como-ufqp-2026",
      status="cfp_closed",
      abstract_deadline="2026-05-31",
      registration_deadline="2026-05-31",
      last_change="application deadline May 31 now closed; school Jul 27–31 confirmed (max 40 participants; speakers from ICFO, MPL, Technion)")

patch("spw-2026",
      abstract_deadline="2026-03-25",
      registration_deadline="2026-05-19",
      last_change="abstract deadline Mar 25 (closed); registration deadline May 19 (closed); event Jul 6–10 Naples confirmed")

# ── Other field/deadline updates ─────────────────────────────────────────────

patch("benasque-qsi-2026",
      name="Benasque Quantum Optics 2026",
      acronym="Benasque QO 2026",
      website="https://www.benasque.org/2026qsi/",
      last_change="name confirmed as 'Benasque Quantum Optics 2026' (organizers: Chang/García-Ripoll/Rabl/Romero-Isart); URL benasque.org/2026qsi/",
      source_urls="https://www.benasque.org/2026qsi/;https://www.squad-germany.de/en/benasque-quantum-optics-2/")

patch("apc-2026",
      location="Long Beach, CA, USA",
      start_date="2026-07-27",
      end_date="2026-07-30",
      last_change="location confirmed Long Beach CA (Hilton Long Beach); dates Jul 27–30",
      source_urls="https://www.optica.org/events/congress/advanced_photonics_congress/;https://10times.com/e1h1-7p53-zkfd-k")

patch("eosam-2026",
      paper_deadline="2026-07-15",
      last_change="post-deadline poster submission open until Jul 15 2026; reg open",
      source_urls="https://www.europeanoptics.org/events/eos/eosam2026.html;https://prein.fi/events/eos-annual-meeting-2026/")

patch("ecoc-2026",
      last_change="paper deadline Apr 22 closed; demo paper deadline May 31 closed; registration now open; Sep 20–24 Málaga confirmed",
      source_urls="https://ecoc2026.org/;https://ecoc2026.org/ECOC2026/paper-submission")

patch("qcrypt-2026",
      last_change="registration now open; confirmed Aug 24–28 Ottawa Learning Crossroads CRX",
      source_urls="https://qcrypt.net/2026/;https://qcrypt.net/2026/attend/registration/")

patch("grc-plasmonics-nano-2026",
      registration_deadline="2026-06-21",
      early_bird_deadline="2026-04-12",
      last_change="GRC application deadline Jun 21 2026 (active!); GRS application deadline Jun 20; companion GRS Jul 18–19",
      source_urls="https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/;https://www.grc.org/plasmonics-and-nanophotonics-grs-conference/2026/")

patch("photonics-west-2027",
      abstract_deadline="2026-07-22",
      last_change="abstract deadline 2026-07-22 confirmed; exhibition Feb 2–4; BiOS Expo Jan 30–31",
      source_urls="https://spie.org/conferences-and-exhibitions/photonics-west;https://spie.org/conferences-and-exhibitions/photonics-west/program/browse-program")

# Fill OWTNM placeholder with actual 2026 event (now archived)
patch("owtnm",
      id="owtnm-2026",
      name="XXXII International Workshop on Optical Wave & Waveguide Theory and Numerical Modelling",
      acronym="OWTNM 2026",
      start_date="2026-04-08",
      end_date="2026-04-10",
      location="Lausanne, Switzerland",
      abstract_deadline="2026-02-06",
      paper_deadline="n/a",
      registration_deadline="2026-03-06",
      early_bird_deadline="n/a",
      website="https://owtnm26.epfl.ch/",
      status="archived",
      last_change="placeholder filled: OWTNM 2026 was Apr 8–10 EPFL Lausanne; now archived",
      source_urls="https://owtnm26.epfl.ch/;https://www.iop.org/events/international-workshop-optical-wave-waveguide-theory-and-numerical-modelling")
# Re-index after id change
idx = {r["id"]: i for i, r in enumerate(rows)}

# AMPD 2026 dates confirmed (Apr 28-30 ZIB Berlin) — already archived; just verify
patch("ampd-2026",
      location="Berlin, Germany",
      start_date="2026-04-28",
      end_date="2026-04-30",
      website="https://www.zib.de/workshop-photonic-devices/",
      last_change="confirmed Apr 28–30 ZIB Berlin; 18th edition (AMPD2026); archived",
      source_urls="https://www.zib.de/workshop-photonic-devices/;https://forschungscampus-modal.de/18th-annual-meeting-photonic-devices-ampd2026")

patch("eqtc-2026",
      last_change="confirmed Nov 29–Dec 3 Dublin; EU Quantum Flagship event; abstract/registration TBA",
      source_urls="https://www.eqtc2026.eu/;https://qt.eu/events/eqtc-2026-european-quantum-technologies-conference")

patch("cewqo-2026",
      location="Erlangen, Germany",
      last_change="confirmed Jul 20–24 Erlangen (Kreuz+Quer); organized by Chekhova/Marquardt (FAU/MPL)",
      source_urls="https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/")

patch("qip-2027",
      start_date="2027-02-20",
      end_date="2027-02-26",
      last_change="confirmed Feb 20–26 2027 Singapore CQT (incl. tutorials Feb 20–21)",
      source_urls="https://qipconference.org/2027/;https://qip.iaqi.org/nextqip")

# ---------------------------------------------------------------------------
# 3.  New entries
# ---------------------------------------------------------------------------

new_rows = []

# ── Sheffield Quantum Optics & Nanophotonics Winter School (missed in May run) ─
new_rows.append({
    "id": "wsqon-sheffield-2026",
    "name": "4th Quantum Optics and Nanophotonics Winter School",
    "acronym": "QONWS 2026",
    "type": "Winter School",
    "topic_tags": "quantum optics;single-photon sources;nanophotonics;SNSPDs",
    "start_date": "2026-01-27",
    "end_date": "2026-01-28",
    "location": "Sheffield, UK",
    "abstract_deadline": "2025-12-19",
    "paper_deadline": "n/a",
    "registration_deadline": "2026-01-26",
    "early_bird_deadline": "2025-12-08",
    "website": "https://sites.google.com/sheffield.ac.uk/quantumopticsnanophotonics2026/home",
    "relevance_score": "5",
    "relevance_note": "Dedicated quantum optics & nanophotonics winter school covering single-photon sources, SNSPDs, QDs and light-matter interaction — bullseye for AG Schuck PhDs",
    "status": "archived",
    "last_verified": TODAY,
    "last_change": "new entry (missed in May run); held Jan 27–28 Sheffield; organised by Sheffield Quantum Centre / M4QN; sponsored by ID Quantique, Single Quantum, TOPTICA",
    "source_urls": "https://sites.google.com/sheffield.ac.uk/quantumopticsnanophotonics2026/home;https://www.idquantique.com/nanophotonics-sheffield/;https://m4qn.org/events/winter-school-on-quantum-optics-2026"
})

# ── CLEO 2027 ─────────────────────────────────────────────────────────────────
new_rows.append({
    "id": "cleo-2027",
    "name": "Conference on Lasers and Electro-Optics 2027",
    "acronym": "CLEO 2027",
    "type": "Conference",
    "topic_tags": "integrated quantum photonics;quantum optics;nonlinear optics;silicon photonics;single-photon sources",
    "start_date": "2027-05-02",
    "end_date": "2027-05-06",
    "location": "Long Beach, CA, USA",
    "abstract_deadline": "TBA",
    "paper_deadline": "n/a",
    "registration_deadline": "TBA",
    "early_bird_deadline": "TBA",
    "website": "https://cleoconference.org/",
    "relevance_score": "5",
    "relevance_note": "Premier US laser/photonics conference; major IQP, Si-PIC and quantum optics sessions",
    "status": "upcoming",
    "last_verified": TODAY,
    "last_change": "new entry; May 2–6 2027 Long Beach Convention Center; exhibition May 4–5",
    "source_urls": "https://cleoconference.org/;https://www.optica.org/events/global_calendar/events/2027/conference_on_lasers_and_electro-optics_(1)/;https://10times.com/e13k-xzz2-d2z4"
})

rows.extend(new_rows)

# ---------------------------------------------------------------------------
# 4.  Write TSV
# ---------------------------------------------------------------------------
with open(OUT_TSV, "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLS, delimiter="\t", extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)

print(f"Wrote {len(rows)} rows → {OUT_TSV}")

# ---------------------------------------------------------------------------
# 5.  Write XLSX
# ---------------------------------------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Conferences"

HEADER_FILL   = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT   = Font(bold=True, color="FFFFFF", size=10)
RED_FILL      = PatternFill("solid", fgColor="FF6666")
AMBER_FILL    = PatternFill("solid", fgColor="FFB347")
ARCHIVED_FONT = Font(color="888888", italic=True, size=9)
BODY_FONT     = Font(size=9)

# Human-readable header labels
LABELS = {
    "id": "ID",
    "name": "Event Name",
    "acronym": "Acronym",
    "type": "Type",
    "topic_tags": "Topic Tags",
    "start_date": "Start Date",
    "end_date": "End Date",
    "location": "Location",
    "abstract_deadline": "Abstract Deadline",
    "paper_deadline": "Paper Deadline",
    "registration_deadline": "Reg. Deadline",
    "early_bird_deadline": "Early Bird",
    "website": "Website",
    "relevance_score": "Score",
    "relevance_note": "Relevance Note",
    "status": "Status",
    "last_verified": "Last Verified",
    "last_change": "Last Change",
    "source_urls": "Source URLs",
}

DEADLINE_COLS = {"abstract_deadline", "paper_deadline",
                 "registration_deadline", "early_bird_deadline"}
DATE_COLS = {"start_date", "end_date", "last_verified"}

def parse_iso(s):
    try:
        return date.fromisoformat(s)
    except Exception:
        return None

today = date.fromisoformat(TODAY)
col_idx = {c: i+1 for i, c in enumerate(COLS)}

# Header row
for ci, col in enumerate(COLS, start=1):
    cell = ws.cell(row=1, column=ci, value=LABELS.get(col, col))
    cell.fill = HEADER_FILL
    cell.font = HEADER_FONT
    cell.alignment = Alignment(wrap_text=True, vertical="center")

# Data rows
for ri, row in enumerate(rows, start=2):
    is_archived = row.get("status") == "archived"
    for ci, col in enumerate(COLS, start=1):
        val = row.get(col, "")
        cell = ws.cell(row=ri, column=ci, value=val)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.font = ARCHIVED_FONT if is_archived else BODY_FONT

        # Conditional colour on deadline columns
        if col in DEADLINE_COLS and not is_archived:
            d = parse_iso(val)
            if d is not None:
                delta = (d - today).days
                if delta <= 14:
                    cell.fill = RED_FILL
                elif delta <= 60:
                    cell.fill = AMBER_FILL

# Freeze header row
ws.freeze_panes = "A2"

# Auto-filter on header
ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}1"

# Column widths
WIDTHS = {
    "id": 22, "name": 48, "acronym": 16, "type": 14,
    "topic_tags": 42, "start_date": 12, "end_date": 12,
    "location": 26, "abstract_deadline": 14, "paper_deadline": 14,
    "registration_deadline": 14, "early_bird_deadline": 14,
    "website": 40, "relevance_score": 7, "relevance_note": 46,
    "status": 12, "last_verified": 12, "last_change": 44,
    "source_urls": 48,
}
for col, width in WIDTHS.items():
    ws.column_dimensions[get_column_letter(col_idx[col])].width = width

# Row height for header
ws.row_dimensions[1].height = 30

wb.save(OUT_XLSX)
print(f"Wrote XLSX → {OUT_XLSX}")
print("Done.")
