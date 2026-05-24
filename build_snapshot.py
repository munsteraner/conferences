#!/usr/bin/env python3
"""
Build conferences_2026-05-13.tsv  and  conferences_2026-05-13.xlsx
for AG Schuck quantum-photonics conference tracker.
"""

import csv, os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

TODAY     = date(2026, 5, 24)
SNAPSHOT  = "2026-05-24"
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
# Each entry is a dict keyed by COLS.
# Topic tags drawn from scope vocabulary:
#   integrated quantum photonics | photonic integrated circuits
#   single-photon sources | SNSPDs | quantum optics | nonlinear optics
#   quantum communication | QKD | quantum networking
#   silicon photonics | TFLN | III-V on Si
#   nanophotonics | plasmonics | optomechanics | 2D-material photonics
#   quantum computing hardware | cryogenic electronics

def ev(id, name, acronym, typ, tags, sd, ed, loc,
       abs_dl, pap_dl, reg_dl, eb_dl,
       url, score, note, status, change, srcs):
    return {
        "id": id, "name": name, "acronym": acronym, "type": typ,
        "topic_tags": tags, "start_date": sd, "end_date": ed,
        "location": loc,
        "abstract_deadline": abs_dl, "paper_deadline": pap_dl,
        "registration_deadline": reg_dl, "early_bird_deadline": eb_dl,
        "website": url, "relevance_score": score,
        "relevance_note": note, "status": status,
        "last_verified": SNAPSHOT, "last_change": change,
        "source_urls": srcs,
    }

LV = SNAPSHOT  # last_verified = today for all rows

EVENTS = [

# ════════════════════════════════════════════════════════════════════════
#  ARCHIVED — passed before 2026-05-13
# ════════════════════════════════════════════════════════════════════════
ev("benasque-qnp-2025",
   "Quantum Nanophotonics 2025","Benasque QNP","Workshop",
   "nanophotonics;quantum optics;single-photon sources",
   "2025-03-16","2025-03-22","Benasque, Spain",
   "2025-01-17","n/a","TBA","n/a",
   "https://benasque.org/2025quantumnanophotonics/",5,
   "Directly on nanophotonics and single-photon sources","archived",
   "year roll-over from 2025 edition",
   "https://benasque.org/2025quantumnanophotonics/"),

ev("photonics-north-2025",
   "Photonics North 2025","PN 2025","Conference",
   "silicon photonics;integrated quantum photonics;quantum communication",
   "2025-05-20","2025-05-23","Ottawa, Canada",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Canadian photonics flagship with PIC and quantum sessions","archived",
   "year roll-over from 2025 edition","original list"),

ev("benasque-sqt-2025",
   "Spring School on Superconducting Qubit Technology 2025","Benasque SQT",
   "Summer School",
   "quantum computing hardware;cryogenic electronics",
   "2025-05-21","2025-05-30","Benasque, Spain",
   "TBA","n/a","TBA","n/a",
   "https://benasque.org/2025sqt/",3,
   "Cryogenic electronics for qubit control, adjacent to photonic detector systems",
   "archived","year roll-over from 2025 edition",
   "https://benasque.org/2025sqt/"),

ev("optica-quantum-2025",
   "Optica Quantum 2.0 Conference and Exhibition 2025","Optica Quantum 2025",
   "Conference",
   "integrated quantum photonics;quantum optics;single-photon sources;quantum communication",
   "2025-06-01","2025-06-05","San Francisco, CA, USA",
   "TBA","TBA","TBA","TBA",
   "https://optica.org/events/topical_meetings/quantum/",5,
   "Premier global quantum photonics conference covering PICs, single photons, QKD",
   "archived","year roll-over from 2025 edition",
   "https://optica.org/events/topical_meetings/quantum/"),

ev("icols-2025",
   "26th International Conference on Laser Spectroscopy","ICOLS 2025",
   "Conference","quantum optics;nonlinear optics",
   "2025-06-02","2025-06-07","Elba Island, Italy",
   "2025-01-17","TBA","TBA","TBA",
   "https://icols2025.it/",3,
   "Laser spectroscopy with quantum optics relevance",
   "archived","year roll-over from 2025 edition",
   "https://icols2025.it/"),

ev("ecio-2025",
   "European Conference on Integrated Optics 2025","ECIO 2025","Conference",
   "integrated quantum photonics;photonic integrated circuits;silicon photonics",
   "2025-06-10","2025-06-12","Cardiff, Wales, UK",
   "TBA","TBA","TBA","TBA",
   "https://ecio-conference.org/",5,
   "Largest European dedicated integrated optics conference",
   "archived","year roll-over from 2025 edition",
   "https://ecio-conference.org/"),

ev("benasque-qi-2025",
   "Quantum Information 2025","Benasque QI","Workshop",
   "quantum communication;quantum optics",
   "2025-06-15","2025-07-04","Benasque, Spain",
   "TBA","n/a","TBA","n/a",
   "https://benasque.org/",3,
   "Quantum information including QKD and networking","archived",
   "year roll-over from 2025 edition",
   "https://benasque.org/"),

ev("damop-2025",
   "56th APS Division of Atomic, Molecular and Optical Physics Meeting",
   "DAMOP 2025","Conference","quantum optics;nonlinear optics",
   "2025-06-16","2025-06-20","Portland, OR, USA",
   "2025-01-31","2025-01-31","TBA","TBA",
   "https://aps.org/events/2025/atomic-molecular-optical-physics",3,
   "AMO physics flagship with quantum optics sessions",
   "archived","year roll-over from 2025 edition",
   "https://aps.org/events/2025/atomic-molecular-optical-physics"),

ev("cleo-europe-2025",
   "CLEO/Europe-EQEC 2025","CLEO/Europe 2025","Conference",
   "integrated quantum photonics;quantum optics;nonlinear optics;nanophotonics",
   "2025-06-23","2025-06-27","Munich, Germany",
   "2025-01-27","2025-01-27","TBA","TBA",
   "https://cleoeurope.org/",5,
   "Flagship European photonics and quantum optics conference (biennial)",
   "archived","year roll-over from 2025 edition",
   "https://cleoeurope.org/"),

ev("cewqo-2025",
   "29th Central European Workshop on Quantum Optics","CEWQO 2025",
   "Workshop","quantum optics;single-photon sources;quantum communication",
   "2025-06-23","2025-06-27","Vilnius, Lithuania",
   "TBA","n/a","TBA","n/a",
   "https://ucan.physics.utoronto.ca/conferences/cewqo29/",4,
   "Dedicated European workshop on quantum optics and single-photon physics",
   "archived","year roll-over from 2025 edition",
   "https://ucan.physics.utoronto.ca/conferences/cewqo29/"),

ev("icton-2025",
   "25th International Conference on Transparent Optical Networks",
   "ICTON 2025","Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2025-07-06","2025-07-10","Barcelona, Spain",
   "TBA","TBA","TBA","TBA",
   "https://icton2025.upc.edu/",3,
   "Transparent optical networks with PIC and quantum photonics tracks",
   "archived","year roll-over from 2025 edition",
   "https://icton2025.upc.edu/"),

ev("ieee-sum-2025",
   "IEEE Summer Topicals Meeting Series 2025","IEEE SUM 2025","Conference",
   "integrated quantum photonics;silicon photonics",
   "2025-07-21","2025-07-23","Berlin, Germany",
   "TBA","TBA","TBA","TBA",
   "https://ieee-sum.org/",3,
   "IEEE topical photonics sessions including PICs",
   "archived","year roll-over from 2025 edition",
   "https://ieee-sum.org/"),

ev("qcrypt-2025",
   "QCrypt 2025","QCrypt 2025","Conference",
   "quantum communication;QKD;quantum networking",
   "2025-08-25","2025-08-29","Sanya, China",
   "TBA","TBA","TBA","TBA",
   "https://qcrypt.net/",4,
   "Premier annual quantum cryptography and QKD conference",
   "archived","year roll-over from 2025 edition",
   "https://qcrypt.net/"),

ev("ecoc-2025",
   "51st European Conference on Optical Communication","ECOC 2025",
   "Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2025-09-28","2025-10-02","Copenhagen, Denmark",
   "TBA","TBA","TBA","TBA",
   "https://ecoc2025.org/",4,
   "Europe's biggest optical comms conference with PIC and quantum tracks",
   "archived","year roll-over from 2025 edition",
   "https://ecoc2025.org/"),

ev("moc-2025",
   "30th Microoptics Conference","MOC 2025","Conference",
   "photonic integrated circuits;nanophotonics",
   "2025-10-12","2025-10-15","Tochigi, Japan",
   "2025-05-16","TBA","TBA","TBA","TBA",3,
   "Micro- and nano-optics with PIC sessions",
   "archived","year roll-over from 2025 edition","original list"),

ev("fio-2025",
   "Frontiers in Optics + Laser Science 2025","FiO+LS 2025","Conference",
   "quantum optics;nonlinear optics;nanophotonics",
   "2025-10-27","2025-10-30","Denver, CO, USA",
   "TBA","TBA","TBA","TBA",
   "https://frontiersinoptics.com/",3,
   "Broad US optics conference with quantum photonics components",
   "archived","year roll-over from 2025 edition",
   "https://frontiersinoptics.com/"),

ev("piers-fall-2025",
   "Photonics & Electromagnetics Research Symposium — Fall 2025",
   "PIERS Fall 2025","Symposium",
   "photonic integrated circuits;nanophotonics",
   "2025-11-04","2025-11-08","Chiba, Japan",
   "TBA","TBA","TBA","TBA","TBA",2,
   "General EM/photonics symposium; peripherally relevant",
   "archived","year roll-over from 2025 edition","original list"),

ev("acp-2025",
   "Asia Communications and Photonics Conference 2025","ACP 2025",
   "Conference",
   "photonic integrated circuits;quantum communication",
   "2025-11-05","2025-11-08","Jiangsu, China",
   "2025-08-01","TBA","TBA","TBA","TBA",3,
   "Asia-Pacific photonics flagship with quantum sessions",
   "archived","year roll-over from 2025 edition","original list"),

ev("ipc-2025",
   "IEEE Photonics Conference 2025","IPC 2025","Conference",
   "integrated quantum photonics;photonic integrated circuits;silicon photonics",
   "2025-11-09","2025-11-13","Singapore",
   "TBA","TBA","TBA","TBA",
   "https://ieee-ipc.org/",4,
   "IEEE flagship photonics devices and integrated systems conference",
   "archived","year roll-over from 2025 edition",
   "https://ieee-ipc.org/"),

ev("optima-2025",
   "Optical Communication, Photonics, Telecommunications, and Intelligent Machine Applications",
   "OPTIMA 2025","Conference","photonic integrated circuits",
   "2025-12-04","2025-12-05","Tashkent, Uzbekistan",
   "TBA","TBA","TBA","TBA","TBA",2,
   "Regional optics/photonics conference; limited quantum scope",
   "archived","year roll-over from 2025 edition","original list"),

ev("pgc-2025",
   "Photonics Global Conference 2025","PGC 2025","Conference",
   "photonic integrated circuits;nanophotonics",
   "2025-12-05","2025-12-08","Sanya, China",
   "TBA","TBA","TBA","TBA","TBA",2,
   "Asia-Pacific photonics conference; some PIC tracks",
   "archived","year roll-over from 2025 edition","original list"),

ev("icisct-2025",
   "International Conference on Information Science and Communications Technologies",
   "ICISCT 2025","Conference","quantum communication",
   "2025-12-26","2025-12-27","Kathmandu, Nepal",
   "TBA","TBA","TBA","TBA","TBA",2,
   "Broad ICT conference; limited photonics scope",
   "archived","year roll-over from 2025 edition","original list"),

ev("photonics-west-2026",
   "SPIE Photonics West 2026","PW 2026","Conference",
   "integrated quantum photonics;silicon photonics;nonlinear optics;single-photon sources",
   "2026-01-17","2026-01-22","San Francisco, CA, USA",
   "2025-07-09","TBA","TBA","TBA",
   "https://spie.org/photonics-west",5,
   "SPIE flagship with Quantum West and Si-photonics tracks",
   "archived","year roll-over; Photonics West 2027 now announced (Jan 30–Feb 4)",
   "https://spie.org/photonics-west"),

ev("qip-2026",
   "29th Quantum Information Processing Conference","QIP 2026","Conference",
   "quantum communication;quantum computing hardware;quantum networking",
   "2026-01-26","2026-01-30","Riga, Latvia",
   "2025-09-12","2025-10-10","TBA","TBA",
   "https://qip2026.lu.lv/",4,
   "Premier quantum information conference; networking and photonic hardware sessions",
   "archived","year roll-over; QIP 2027 Singapore announced",
   "https://qip2026.lu.lv/"),

ev("dpg-samop-2026",
   "DPG Frühjahrstagung SAMOP 2026","DPG SAMOP 2026","Conference",
   "quantum optics;nanophotonics;integrated quantum photonics",
   "2026-03-01","2026-03-06","Mainz, Germany",
   "TBA","TBA","TBA","TBA",
   "https://mainz26.dpg-tagungen.de",4,
   "German physics society AMO/photonics spring meeting — easily accessible from Münster",
   "archived","year roll-over; DPG SAMOP 2027 Hannover announced",
   "https://mainz26.dpg-tagungen.de"),

ev("nanolight-2026",
   "Nanolight Conference 2026","Nanolight 2026","Conference",
   "nanophotonics;2D-material photonics;quantum optics;plasmonics",
   "2026-03-02","2026-03-06","Benasque, Spain",
   "2025-12-15","n/a","2026-02-15","n/a",
   "https://benasque.org/2026nanolight/",5,
   "European nanophotonics / near-field / 2D-material flagship",
   "archived","year roll-over",
   "https://benasque.org/2026nanolight/"),

ev("dpg-skm-2026",
   "DPG Frühjahrstagung SKM 2026","DPG SKM 2026","Conference",
   "silicon photonics;nanophotonics;2D-material photonics",
   "2026-03-08","2026-03-13","Dresden, Germany",
   "TBA","TBA","TBA","TBA",
   "https://dresden26.dpg-tagungen.de",3,
   "Condensed matter sessions on photonics materials and devices",
   "archived","year roll-over",
   "https://dresden26.dpg-tagungen.de"),

ev("ofc-2026",
   "Optical Fiber Communication Conference 2026","OFC 2026","Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2026-03-15","2026-03-19","Los Angeles, CA, USA",
   "TBA","TBA","TBA","TBA",
   "https://ofcconference.org/",4,
   "World's largest optical comms conference with PIC and quantum networking sessions",
   "archived","year roll-over; OFC 2027 Los Angeles Mar 7-11 announced",
   "https://ofcconference.org/"),

ev("aps-march-2026",
   "APS Global Physics Summit 2026","APS 2026","Conference",
   "quantum optics;quantum computing hardware;integrated quantum photonics",
   "2026-03-16","2026-03-20","Denver, CO, USA",
   "2025-10-31","2025-10-31","TBA","TBA",
   "https://summit.aps.org/",4,
   "Largest physics meeting; major quantum photonics and computing sessions",
   "archived",
   "year roll-over; confirmed as joint March+April Global Physics Summit; APS 2027 Atlanta Apr 11-16",
   "https://summit.aps.org/; https://march.aps.org/"),

ev("spie-pe-2026",
   "SPIE Photonics Europe 2026","SPIE PE 2026","Conference",
   "integrated quantum photonics;silicon photonics;nanophotonics",
   "2026-04-12","2026-04-16","Strasbourg, France",
   "TBA","TBA","TBA","TBA",
   "https://spie.org/photonics-europe",5,
   "Major European SPIE photonics conference with quantum integrated sessions",
   "archived","year roll-over",
   "https://spie.org/photonics-europe"),

ev("ieee-siphotonics-2026",
   "IEEE Silicon Photonics Conference 2026","IEEE SiPho 2026","Conference",
   "silicon photonics;photonic integrated circuits;III-V on Si",
   "2026-04-13","2026-04-15","Ottawa, ON, Canada",
   "2025-10-03","TBA","TBA","TBA",
   "https://ieee-siphotonics.org/",5,
   "Dedicated silicon photonics conference — direct relevance to PICs",
   "archived","year roll-over",
   "https://ieee-siphotonics.org/"),

ev("ieee-qpain-2026",
   "IEEE 2nd International Conference on Quantum Photonics, AI & Networking",
   "IEEE QPAIN 2026","Conference",
   "integrated quantum photonics;quantum networking",
   "2026-04-16","2026-04-18","Chattogram, Bangladesh",
   "TBA","TBA","TBA","TBA",
   "https://qpain.org/",3,
   "Quantum photonics and networking conference (IEEE-sponsored)",
   "archived","year roll-over",
   "https://qpain.org/"),

ev("inp-webinar-2026",
   "Webinar on InP platform for NIR and MIR applications",
   "InP Webinar","Seminar",
   "photonic integrated circuits;III-V on Si",
   "2026-04-17","2026-04-17","Online",
   "TBA","n/a","TBA","n/a",
   "https://inphomir.eu/1st-webinar/",4,
   "III-V InP PIC platform — directly relevant to quantum light-source integration",
   "archived","year roll-over",
   "https://inphomir.eu/1st-webinar/"),

ev("eftf-2026",
   "European Frequency and Time Forum 2026","EFTF 2026","Conference",
   "quantum optics;quantum communication",
   "2026-04-20","2026-04-23","Noordwijk, The Netherlands",
   "TBA","TBA","TBA","TBA",
   "https://www.eftf.org/home",2,
   "Frequency/time standards; quantum photonic metrology peripherally relevant",
   "archived","year roll-over",
   "https://www.eftf.org/home"),

ev("spie-ds-2026",
   "SPIE Defense + Security 2026","SPIE D+S 2026","Conference",
   "single-photon sources;quantum communication",
   "2026-04-26","2026-04-30","National Harbor, MD, USA",
   "TBA","TBA","TBA","TBA",
   "https://spie.org/defense-and-commercial-sensing",2,
   "Defense photonics with some single-photon sensing sessions",
   "archived","year roll-over",
   "https://spie.org/defense-and-commercial-sensing"),

ev("bqit-2026",
   "Bristol Quantum Information Technologies Workshop 2026","BQIT 2026",
   "Workshop",
   "integrated quantum photonics;quantum communication;quantum computing hardware",
   "2026-04-27","2026-04-29","Bristol, UK",
   "TBA","n/a","TBA","n/a",
   "https://www.bristol.ac.uk/qet-labs/events/bqit-workshop/",5,
   "UK quantum photonics workshop — PICs, single photons, quantum networking",
   "archived","year roll-over",
   "https://www.bristol.ac.uk/qet-labs/events/bqit-workshop/"),

ev("quantumatter-2026",
   "QUANTUMatter 2026","QUANTU 2026","Conference",
   "quantum computing hardware;quantum optics;nanophotonics",
   "2026-04-27","2026-04-30","Barcelona, Spain",
   "2026-02-11","n/a","2026-03-06","2026-02-23",
   "https://www.quantumconf.eu/2026/",3,
   "Quantum materials and devices, including photonic quantum systems",
   "archived","year roll-over",
   "https://www.quantumconf.eu/2026/"),

ev("ampd-2026",
   "Annual Meeting Photonic Devices 2026","AMPD 2026","Conference",
   "photonic integrated circuits;nanophotonics",
   "2026-04-28","2026-04-30","TBA",
   "TBA","TBA","TBA","TBA",
   "https://www.zib.de/workshop-photonic-devices/",4,
   "German photonic devices annual meeting — directly relevant to group's fabrication work",
   "archived","year roll-over",
   "https://www.zib.de/workshop-photonic-devices/"),

# ════════════════════════════════════════════════════════════════════════
#  UPCOMING — 2026 (sorted by start_date)
# ════════════════════════════════════════════════════════════════════════
ev("cleo-2026",
   "Conference on Lasers and Electro-Optics 2026","CLEO 2026","Conference",
   "integrated quantum photonics;quantum optics;nonlinear optics;silicon photonics;single-photon sources",
   "2026-05-17","2026-05-21","Charlotte, NC, USA",
   "2025-11-18","n/a","TBA","TBA",
   "https://cleoconference.org/",5,
   "Premier US laser/photonics conference; major IQP, Si-PIC and quantum optics sessions",
   "archived",
   "archived; concluded May 17–21 2026 Charlotte NC; proceedings via OSA/Optica Publishing",
   "https://cleoconference.org/; https://cleoconference.org/conference-archives/"),

ev("qnetworks-2026",
   "QNetworks 2026 Workshop","QNetworks 2026","Workshop",
   "quantum networking;quantum communication;integrated quantum photonics;QKD",
   "2026-05-12","2026-05-13","Bristol, UK",
   "TBA","n/a","TBA","n/a",
   "https://iqnhub.org/qnetworks-2026/",4,
   "UK quantum networks workshop — integrated photonics for QKD, field trials, entanglement distribution",
   "archived",
   "new entry (archived); held May 12-13 Bristol Megascreen; organised by UK IQN Hub",
   "https://iqnhub.org/qnetworks-2026/; https://www.eventbrite.co.uk/e/qnetworks-workshop-2026-future-quantum-networks-tickets-1981318504984"),

ev("damop-2026",
   "57th APS Division of Atomic, Molecular and Optical Physics Meeting",
   "DAMOP 2026","Conference",
   "quantum optics;nonlinear optics;single-photon sources",
   "2026-06-01","2026-06-05","Providence, RI, USA",
   "2026-01-23","TBA","TBA","TBA",
   "https://aps.org/events/2026/damop-meeting-2026",3,
   "AMO physics flagship with quantum optics and single-photon sessions",
   "upcoming","starts in 8 days (Jun 1-5 Rhode Island Convention Center); registration open",
   "https://aps.org/events/2026/damop-meeting-2026"),

ev("photonics-north-2026",
   "Photonics North 2026","PN 2026","Conference",
   "silicon photonics;integrated quantum photonics;quantum communication",
   "2026-06-01","2026-06-04","TBA",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Canadian national photonics conference with PIC and quantum sessions",
   "upcoming","no change; location TBA",
   "original list"),

ev("cargese-qt-2026",
   "Summer School on Quantum Technologies for Computation and Communication",
   "QT4CC 2026","Summer School",
   "quantum communication;quantum computing hardware;quantum networking",
   "2026-06-08","2026-06-20","Cargèse (Corsica), France",
   "TBA","n/a","TBA","n/a",
   "https://qt4cc.sciencesconf.org/",4,
   "PhD summer school on quantum computation and communication — ideal for group's PhD students",
   "upcoming","new entry; Jun 8-20 at Institut d'Études Scientifiques de Cargèse",
   "https://qt4cc.sciencesconf.org/; https://gdr-teq.cnrs.fr/cevent/629/"),

ev("spie-pfq-2026",
   "SPIE Photonics for Quantum 2026","PfQ 2026","Conference",
   "integrated quantum photonics;single-photon sources;quantum communication;quantum computing hardware",
   "2026-06-08","2026-06-11","Waterloo, ON, Canada",
   "2026-03-01","n/a","TBA","TBA",
   "https://spie.org/conferences-and-exhibitions/photonics-for-quantum",5,
   "Dedicated quantum photonics conference — bullseye for AG Schuck",
   "cfp_closed","confirmed Jun 8-11 Waterloo; post-deadline submissions open",
   "https://spie.org/PFQ26/conferencedetails/photonics-for-quantum"),

ev("islc-2026",
   "30th International Semiconductor Laser Conference","ISLC 2026","Conference",
   "single-photon sources;III-V on Si;photonic integrated circuits",
   "2026-06-14","2026-06-17","Tampere, Finland",
   "2026-02-05","TBA","TBA","TBA",
   "https://events.tuni.fi/islc2026/",4,
   "III-V semiconductor laser conference — key for quantum light source development",
   "cfp_closed",
   "abstract deadline was extended to 2026-02-05 (now closed)",
   "https://events.tuni.fi/islc2026/; https://ieeephotonics.org/event/2026-30th-international-semiconductor-laser-conference-islc/"),

ev("grc-quantum-science-2026",
   "Gordon Research Conference — Quantum Science 2026","GRC QS 2026",
   "Conference",
   "quantum optics;quantum computing hardware;quantum communication;integrated quantum photonics",
   "2026-06-14","2026-06-19","Easton, MA, USA",
   "TBA","n/a","2026-06-28","n/a",
   "https://www.grc.org/quantum-science-conference/2026/",4,
   "GRC on quantum science covering photonic quantum systems and quantum information",
   "upcoming","new entry; application deadline Jun 28",
   "https://www.grc.org/quantum-science-conference/2026/"),

ev("icap-2026",
   "29th International Conference on Atomic Physics","ICAP 2026","Conference",
   "quantum optics;nonlinear optics",
   "2026-06-14","2026-06-19","Wuhan, China",
   "TBA","TBA","TBA","TBA",
   "https://www.icap29.com/home.html",3,
   "Atomic physics conference with quantum optics and light-matter interaction sessions",
   "upcoming","confirmed Jun 14-19 Wuhan",
   "https://www.icap29.com/home.html"),

ev("ecio-2026",
   "European Conference on Integrated Optics 2026","ECIO 2026","Conference",
   "integrated quantum photonics;photonic integrated circuits;silicon photonics;nanophotonics",
   "2026-06-15","2026-06-17","Zürich, Switzerland",
   "TBA","TBA","2026-06-01","2026-05-11",
   "https://ecio-conference.org/",5,
   "Largest European integrated optics conference — core relevance for AG Schuck",
   "cfp_closed",
   "registration deadline Jun 1; early-bird passed May 11; hosted at ETH Zürich",
   "https://www.ecio-conference.org/; https://www.ecio-conference.org/deadlines/"),

ev("optica-quantum-2026",
   "Optica Quantum 2.0 Conference and Exhibition 2026","Optica Quantum 2026",
   "Conference",
   "integrated quantum photonics;quantum optics;single-photon sources;quantum communication;quantum computing hardware",
   "2026-06-15","2026-06-18","Glasgow, UK",
   "2026-02-10","n/a","2026-05-29","2026-05-01",
   "https://www.optica.org/events/topical_meetings/quantum/",5,
   "Premier quantum photonics conference spanning PICs, detectors, QKD, and hardware",
   "cfp_closed",
   "REGISTRATION DEADLINE 2026-05-29 (5 DAYS AWAY); early bird closed May 1; Scottish Event Campus",
   "https://www.optica.org/events/topical_meetings/quantum/; https://www.optica.org/events/topical_meetings/quantum/registration/"),

ev("icop-2026",
   "Italian Conference on Optics and Photonics 2026","ICOP 2026","Conference",
   "integrated quantum photonics;quantum optics;nanophotonics",
   "2026-06-15","2026-06-17","L'Aquila, Italy",
   "TBA","TBA","TBA","TBA",
   "https://www.icop2026.it/",3,
   "Italian national photonics conference with quantum tracks",
   "upcoming","no change",
   "https://www.icop2026.it/"),

ev("ccw-london-2026",
   "Critical Communications World London 2026","CCW 2026","Conference",
   "quantum communication;quantum networking",
   "2026-06-16","2026-06-18","London, UK",
   "TBA","TBA","TBA","TBA",
   "https://www.critical-communications-world.com/",2,
   "Critical comms industry event; QKD sessions of marginal relevance",
   "upcoming","no change",
   "https://www.critical-communications-world.com/call-for-content"),

ev("egas-2026",
   "57th Conference of the European Group on Atomic Systems","EGAS 2026",
   "Conference","quantum optics;nonlinear optics",
   "2026-06-21","2026-06-25","Toruń, Poland",
   "TBA","TBA","TBA","TBA",
   "https://egas57.org/",3,
   "European atomic systems conference with quantum optics and light-matter sessions",
   "upcoming","no change",
   "https://egas57.org/"),

ev("benasque-qsi-2026",
   "Quantum Science: Implementations 2026","Benasque QSI","Workshop",
   "integrated quantum photonics;quantum computing hardware;quantum communication",
   "2026-06-21","2026-07-04","Benasque, Spain",
   "TBA","n/a","TBA","n/a",
   "https://benasque.org/",4,
   "Benasque workshop on quantum implementations including photonic platforms",
   "upcoming","confirmed Jun 21–Jul 4 Benasque",
   "https://benasque.org/; https://www.benasque.org/new_general/cgi-bin/years.pl?ano=2026"),

ev("les-houches-pa-2026",
   "Photons & Atoms 2026 — Les Houches Doctoral School",
   "Les Houches PA 2026","Summer School",
   "quantum optics;nonlinear optics;integrated quantum photonics;single-photon sources",
   "2026-06-22","2026-07-03","Les Houches, France",
   "TBA","n/a","TBA","n/a",
   "https://photon-atoms-26.leshouches.science/",5,
   "Leading European doctoral school on photons, quantum gases and quantum technologies",
   "upcoming","new entry; Jun 22–Jul 3; covers quantum optics, light-matter interaction and QT",
   "https://photon-atoms-26.leshouches.science/; https://first-tf.com/doctoral-training-on-atoms-and-photons-ecole-de-physique-des-houches-france-june-22nd-to-july-3rd-2026/"),

ev("iqt-nordics-2026",
   "IQT Nordics 2026 — Quantum Technology in a Changing World",
   "IQT Nordics 2026","Conference",
   "quantum communication;quantum computing hardware;integrated quantum photonics",
   "2026-06-22","2026-06-24","Oslo, Norway",
   "TBA","n/a","TBA","TBA",
   "https://iqtevent.com/nordics/",3,
   "Nordic quantum industry conference with hardware and QKD sessions",
   "upcoming","new entry; Jun 22-24 OsloMet University",
   "https://iqtevent.com/nordics/; https://www.oslomet.no/en/about/events/iqt-nordics-2026"),

ev("quantum-tech-world-2026",
   "Quantum.Tech World 2026","Q.Tech World 2026","Conference",
   "quantum computing hardware;quantum communication;integrated quantum photonics",
   "2026-06-25","2026-06-26","Boston, MA, USA",
   "TBA","n/a","TBA","TBA",
   "https://www.alphaevents.com/events-quantumtechus",2,
   "Industry quantum technology event; limited deep-science content",
   "upcoming","no change",
   "https://www.alphaevents.com/events-quantumtechus"),

ev("oecc-2026",
   "31st OptoElectronics and Communications Conference","OECC 2026",
   "Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2026-06-28","2026-07-02","Busan, South Korea",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Asia-Pacific optoelectronics and comms conference with PIC sessions",
   "upcoming","no change","original list"),

ev("quantum-korea-2026",
   "Quantum Korea 2026","Quantum Korea 2026","Conference",
   "quantum computing hardware;quantum communication;integrated quantum photonics",
   "2026-07-02","2026-07-04","Seoul, South Korea",
   "TBA","n/a","TBA","TBA",
   "https://quantum-korea.kr/en/main",2,
   "Government-organised Korean quantum technology showcase",
   "upcoming","no change",
   "https://quantum-korea.kr/en/main"),

ev("spw-2026",
   "12th Single Photon Workshop","SPW 2026","Workshop",
   "single-photon sources;SNSPDs;quantum optics;quantum communication",
   "2026-07-06","2026-07-10","Naples, Italy",
   "TBA","n/a","TBA","TBA",
   "https://spw2026.org/",5,
   "THE dedicated single-photon workshop — central to AG Schuck's detector and source work",
   "cfp_closed",
   "confirmed Jul 6-10 Naples (Univ. Federico II); registration open; abstract submission was via EasyChair",
   "https://spw2026.org/; https://iqnhub.org/2026-single-photon-workshop-registration-open/"),

ev("icton-2026",
   "26th International Conference on Transparent Optical Networks",
   "ICTON 2026","Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2026-07-12","2026-07-16","Prague, Czech Republic",
   "TBA","TBA","TBA","TBA",
   "https://www.gov.pl/web/instytut-lacznosci/icton-2026",3,
   "Transparent optical networks with PIC and quantum photonics tracks",
   "upcoming","no change",
   "https://www.gov.pl/web/instytut-lacznosci/icton-2026"),

ev("ieee-sum-2026",
   "IEEE Summer Topicals Meeting Series 2026","IEEE SUM 2026","Conference",
   "photonic integrated circuits;silicon photonics;integrated quantum photonics",
   "2026-07-13","2026-07-15","Tulum, Mexico",
   "TBA","TBA","TBA","TBA",
   "https://ieee-sum.org/",3,
   "IEEE focused topical photonics sessions",
   "upcoming","no change",
   "https://ieee-sum.org/"),

ev("grc-plasmonics-nano-2026",
   "Gordon Research Conference — Plasmonics and Nanophotonics 2026",
   "GRC PNP 2026","Conference",
   "nanophotonics;plasmonics;quantum optics;2D-material photonics",
   "2026-07-19","2026-07-24","Newry, ME, USA",
   "TBA","n/a","2026-06-21","n/a",
   "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/",5,
   "GRC on quantum nanophotonics, plasmonics and 2D photonics — highly relevant to AG Schuck",
   "upcoming","new entry; application deadline Jun 21",
   "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/"),

ev("cewqo-2026",
   "30th Central European Workshop on Quantum Optics","CEWQO30 2026",
   "Workshop",
   "quantum optics;single-photon sources;nonlinear optics;integrated quantum photonics",
   "2026-07-20","2026-07-24","Erlangen, Germany",
   "TBA","n/a","TBA","n/a",
   "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/",5,
   "European quantum optics workshop co-hosted by FAU/MPL Erlangen — close to Münster",
   "upcoming","new entry; Jul 20-24 Erlangen, hosted by Chekhova/Marquardt groups",
   "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/; https://www.oqt.nat.fau.de/events/"),

ev("apc-2026",
   "Optica Advanced Photonics Congress 2026","APC 2026","Conference",
   "integrated quantum photonics;nonlinear optics;nanophotonics",
   "2026-07-26","2026-07-30","TBA",
   "TBA","TBA","TBA","TBA",
   "https://www.optica.org/events/congress/advanced_photonics_congress/",4,
   "Optica congress including nonlinear photonics, waveguides and integrated circuits",
   "upcoming","no change; location TBA",
   "https://www.optica.org/events/congress/advanced_photonics_congress/"),

ev("grc-mech-quant-2026",
   "Gordon Research Conference — Mechanical Systems in the Quantum Regime 2026",
   "GRC MSQR 2026","Conference",
   "optomechanics;quantum computing hardware",
   "2026-07-26","2026-07-31","Lucca (Barga), Italy",
   "TBA","n/a","TBA","n/a",
   "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/",4,
   "Optomechanics GRC — relevant to cavity-optomechanical quantum transducers",
   "upcoming","new entry; Jul 26-31 Lucca Italy",
   "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/"),

ev("lake-como-ufqp-2026",
   "New Frontiers in Ultrafast Quantum Optics — Lake Como School",
   "UFQP 2026","Summer School",
   "quantum optics;nonlinear optics;single-photon sources",
   "2026-07-27","2026-07-31","Como, Italy",
   "2026-05-31","n/a","TBA","n/a",
   "https://ufqp.lakecomoschool.org/",4,
   "Summer school on ultrafast quantum optics — photon correlations and time-frequency entanglement",
   "upcoming",
   "APPLICATION DEADLINE 2026-05-31 (7 DAYS AWAY — STILL OPEN); max 40 places; Villa del Grumello",
   "https://ufqp.lakecomoschool.org/; https://www.quantiki.org/conference/summer-school-new-frontiers-ultrafast-quantum-optics"),

ev("meta-2026",
   "16th International Conference on Metamaterials, Photonic Crystals and Plasmonics",
   "META 2026","Conference",
   "nanophotonics;plasmonics;quantum optics;2D-material photonics",
   "2026-07-14","2026-07-17","Dublin, Ireland",
   "2026-02-13","n/a","TBA","2026-04-24",
   "https://metaconferences.org/META26/index.php/META",4,
   "Metamaterials/nanophotonics conference with dedicated quantum photonics workshop and plasmonics sessions",
   "cfp_closed",
   "new entry; Trinity College Dublin Jul 14-17; oral deadline Feb 13 and poster Mar 13 (both closed); META Quantum Workshop included",
   "https://metaconferences.org/META26/index.php/META; https://metaconferences.org/META26/index.php/META/workshop"),

ev("ieee-rapid-2026",
   "IEEE Research and Applications of Photonics in Defense 2026",
   "IEEE RAPID 2026","Conference",
   "single-photon sources;quantum communication",
   "2026-08-05","2026-08-07","Miramar Beach, FL, USA",
   "TBA","TBA","TBA","TBA",
   "https://ieee-rapid.org/",2,
   "Defense photonics; some single-photon sensing sessions",
   "upcoming","no change",
   "https://ieee-rapid.org/"),

ev("spie-op-2026",
   "SPIE Optics + Photonics 2026","SPIE O+P 2026","Conference",
   "integrated quantum photonics;nanophotonics;quantum optics;single-photon sources",
   "2026-08-23","2026-08-27","San Diego, CA, USA",
   "TBA","TBA","TBA","TBA",
   "https://spie.org/optics-and-photonics",4,
   "SPIE major annual meeting with quantum nanophotonics and quantum technologies tracks",
   "upcoming","dates confirmed Aug 23-27 San Diego; exhibition Aug 25-27; registration opening soon",
   "https://spie.org/optics-and-photonics"),

ev("qcrypt-2026",
   "QCrypt 2026","QCrypt 2026","Conference",
   "quantum communication;QKD;quantum networking",
   "2026-08-24","2026-08-28","Ottawa, Canada",
   "TBA","TBA","TBA","TBA",
   "https://qcrypt.net/2026/",4,
   "Premier annual QKD and quantum cryptography conference",
   "upcoming","confirmed Aug 24-28 Ottawa (Learning Crossroads CRX)",
   "https://qcrypt.net/2026/"),

ev("eosam-2026",
   "European Optical Society Annual Meeting 2026","EOSAM 2026","Conference",
   "integrated quantum photonics;nanophotonics;nonlinear optics",
   "2026-08-24","2026-08-28","Tampere, Finland",
   "2026-04-14","TBA","TBA","2026-06-15",
   "https://www.europeanoptics.org/events/eos/eosam2026.html",4,
   "EOS flagship European photonics meeting with strong quantum photonics programme",
   "cfp_closed","abstract deadline Apr 14 (closed); early bird registration Jun 15; post-deadline poster Jul 15",
   "https://www.europeanoptics.org/events/eos/eosam2026.html; https://www.europeanoptics.org/pages/events/eosam-2026/registration/"),

ev("tqc-2026",
   "Theory of Quantum Computation, Communication and Cryptography 2026",
   "TQC 2026","Conference",
   "quantum communication;quantum computing hardware;quantum networking",
   "2026-08-31","2026-09-04","Sherbrooke, QC, Canada",
   "TBA","TBA","TBA","TBA",
   "https://tqc-conference.org/2026/",3,
   "Theory conference for quantum communication, computation and cryptography",
   "upcoming","new entry",
   "https://tqc-conference.org/2026/"),

ev("asc-2026",
   "Applied Superconductivity Conference 2026","ASC 2026","Conference",
   "SNSPDs;cryogenic electronics;single-photon sources",
   "2026-09-06","2026-09-11","Pittsburgh, PA, USA",
   "2026-02-02","TBA","TBA","TBA",
   "https://appliedsuperconductivity.org/asc2026/",4,
   "Premier superconductivity conference with major SNSPD and cryogenic electronics sessions",
   "cfp_closed","abstract deadline extended to Feb 2 (now closed); acceptance notifications Mar 16; registration open",
   "https://appliedsuperconductivity.org/asc2026/; https://www.appliedsuperconductivity.org/asc2026/important-dates/"),

ev("inphomir-school-2026",
   "INPHOMIR Photonics School 2026","INPHOMIR School","Summer School",
   "photonic integrated circuits;III-V on Si",
   "2026-09-07","2026-09-11","Bari, Italy",
   "TBA","n/a","TBA","n/a",
   "https://inphomir.eu/inphomir-school-program/",4,
   "Dedicated InP/III-V photonics school for NIR/MIR applications",
   "upcoming","no change",
   "https://inphomir.eu/inphomir-school-program/"),

ev("ieee-qce26",
   "IEEE Quantum Week 2026","QCE26","Conference",
   "quantum computing hardware;quantum communication;integrated quantum photonics",
   "2026-09-13","2026-09-18","Toronto, ON, Canada",
   "2026-04-13","2026-04-13","TBA","TBA",
   "https://qce.quantum.ieee.org/2026/",3,
   "IEEE quantum week with photonic quantum computing and QKD sessions",
   "cfp_closed","paper deadline Apr 13 closed; workshop paper abstract Jun 22; poster Phase 1 Jun 1; Sep 13-18 Metro Toronto CC",
   "https://qce.quantum.ieee.org/2026/; https://qce.quantum.ieee.org/2026/submission-deadlines/"),

ev("ecoc-2026",
   "European Conference on Optical Communication 2026","ECOC 2026","Conference",
   "photonic integrated circuits;silicon photonics;quantum communication;integrated quantum photonics",
   "2026-09-20","2026-09-24","Málaga, Spain",
   "2026-04-22","2026-04-22","TBA","TBA",
   "https://ecoc2026.org/",4,
   "Europe's flagship optical comms conference with quantum technologies track",
   "cfp_closed",
   "paper submission closed Apr 22; demo paper deadline May 31; registration open; FYCMA venue Málaga",
   "https://ecoc2026.org/; https://www.ecocexhibition.com/register/"),

ev("mne-2026",
   "International Conference on Micro and Nano Engineering 2026","MNE 2026",
   "Conference","photonic integrated circuits;nanophotonics;2D-material photonics",
   "2026-09-21","2026-09-24","Interlaken, Switzerland",
   "TBA","TBA","TBA","TBA",
   "https://mne2026.imnes.org/",3,
   "Nanofabrication conference relevant to photonic device processing",
   "upcoming","no change",
   "https://mne2026.imnes.org/"),

ev("fio-2026",
   "Frontiers in Optics + Laser Science 2026","FiO+LS 2026","Conference",
   "quantum optics;nonlinear optics;nanophotonics",
   "2026-09-27","2026-09-30","Rochester, NY, USA",
   "TBA","TBA","TBA","TBA",
   "https://frontiersinoptics.com/",3,
   "Broad US photonics conference with quantum optics tracks",
   "upcoming","no change",
   "https://frontiersinoptics.com/"),

ev("hdqs-2026",
   "High-Dimensional Quantum Systems Workshop 2026","HDQS 2026","Workshop",
   "quantum optics;quantum communication;integrated quantum photonics",
   "2026-10-04","2026-10-09","Benasque, Spain",
   "TBA","n/a","TBA","n/a",
   "https://benasque.org/2026hdqs/",4,
   "Benasque workshop on high-dimensional quantum photonics and quantum information",
   "upcoming","no change",
   "https://benasque.org/2026hdqs/"),

ev("photonics-days-2026",
   "Photonics Days Berlin 2026","PDB 2026","Conference",
   "integrated quantum photonics;photonic integrated circuits;nanophotonics",
   "2026-10-07","2026-10-08","Berlin, Germany",
   "TBA","TBA","TBA","TBA",
   "https://photonic-days-berlin.com/",3,
   "German photonics industry and research meeting",
   "upcoming","no change",
   "https://photonic-days-berlin.com/"),

ev("ipc-2026",
   "IEEE Photonics Conference 2026","IPC 2026","Conference",
   "integrated quantum photonics;photonic integrated circuits;silicon photonics",
   "2026-11-08","2026-11-12","Denver, CO, USA",
   "TBA","TBA","TBA","TBA",
   "https://ieee-ipc.org/",4,
   "IEEE flagship photonics devices and integrated systems conference",
   "upcoming","no change",
   "https://ieee-ipc.org/"),

ev("eqtc-2026",
   "European Quantum Technologies Conference 2026","EQTC 2026","Conference",
   "integrated quantum photonics;quantum communication;quantum computing hardware;quantum networking",
   "2026-11-29","2026-12-03","Dublin, Ireland",
   "TBA","TBA","TBA","TBA",
   "https://qt.eu/events/eqtc-2026-european-quantum-technologies-conference",5,
   "EU Quantum Flagship conference — all aspects of quantum technology, high visibility",
   "upcoming","confirmed Nov 29–Dec 3 Dublin",
   "https://qt.eu/events/eqtc-2026-european-quantum-technologies-conference"),

# ════════════════════════════════════════════════════════════════════════
#  UPCOMING — 2027
# ════════════════════════════════════════════════════════════════════════
ev("photonics-west-2027",
   "SPIE Photonics West 2027","PW 2027","Conference",
   "integrated quantum photonics;silicon photonics;nonlinear optics;single-photon sources",
   "2027-01-30","2027-02-04","San Francisco, CA, USA",
   "2026-07-22","TBA","TBA","TBA",
   "https://spie.org/conferences-and-exhibitions/photonics-west",5,
   "SPIE flagship with Quantum West and Si-photonics tracks — abstract deadline Jul 22 2026",
   "upcoming","new entry; Jan 30–Feb 4 2027; abstract deadline 2026-07-22",
   "https://spie.org/conferences-and-exhibitions/photonics-west; https://spie.org/conferences-and-exhibitions/photonics-west/program/browse-program"),

ev("qip-2027",
   "30th Quantum Information Processing Conference","QIP 2027","Conference",
   "quantum communication;quantum computing hardware;quantum networking",
   "2027-02-20","2027-02-26","Singapore",
   "TBA","TBA","TBA","TBA",
   "https://qipconference.org/2027/",4,
   "Premier annual quantum information conference; quantum networking and hardware sessions",
   "upcoming","new entry; Feb 20-26 Singapore (CQT)",
   "https://qipconference.org/2027/; https://qip.iaqi.org/nextqip"),

ev("dpg-samop-2027",
   "DPG Frühjahrstagung SAMOP 2027","DPG SAMOP 2027","Conference",
   "quantum optics;nanophotonics;integrated quantum photonics",
   "2027-02-28","2027-03-05","Hannover, Germany",
   "TBA","TBA","TBA","TBA",
   "https://dpg-physik.de/",4,
   "DPG spring meeting for quantum optics and photonics — easily accessible from Münster",
   "upcoming","no change",
   "https://dpg-physik.de/"),

ev("ofc-2027",
   "Optical Fiber Communication Conference 2027","OFC 2027","Conference",
   "photonic integrated circuits;silicon photonics;quantum communication",
   "2027-03-07","2027-03-11","Los Angeles, CA, USA",
   "TBA","TBA","TBA","TBA",
   "https://ofcconference.org/",4,
   "World's largest optical comms conference with PIC and quantum networking sessions",
   "upcoming","new entry; Mar 7-11 2027 Los Angeles",
   "https://ofcconference.org/"),

ev("spie-oo-2027",
   "SPIE Optics + Optoelectronics 2027","SPIE O+O 2027","Conference",
   "integrated quantum photonics;quantum optics;nanophotonics",
   "2027-04-05","2027-04-08","Prague, Czech Republic",
   "TBA","TBA","TBA","TBA","TBA",4,
   "SPIE European conference with quantum photonics and nanophotonics tracks",
   "upcoming","no change","original list"),

ev("aps-summit-2027",
   "APS Global Physics Summit 2027","APS Summit 2027","Conference",
   "quantum optics;quantum computing hardware;integrated quantum photonics",
   "2027-04-11","2027-04-16","Atlanta, GA, USA",
   "TBA","TBA","TBA","TBA",
   "https://www.aps.org/events/2027/summit",4,
   "Joint APS March+April meeting; major quantum photonics and computing sessions",
   "upcoming","new entry; Apr 11-16 Atlanta",
   "https://www.aps.org/events/2027/summit"),

ev("cleo-europe-2027",
   "CLEO/Europe-EQEC 2027","CLEO/Europe 2027","Conference",
   "integrated quantum photonics;quantum optics;nonlinear optics;nanophotonics",
   "2027-06-20","2027-06-25","Munich, Germany",
   "TBA","TBA","TBA","TBA",
   "https://cleoeurope.org/",5,
   "Flagship European photonics and quantum optics biennial conference",
   "upcoming","corrected end date to Jun 25; part of World of Photonics Congress 2027",
   "https://cleoeurope.org/; https://www.photonics-congress.com/en/about/conferences/cleo-eqec/"),

ev("ieee-rapid-2027",
   "IEEE Research and Applications of Photonics in Defense 2027",
   "IEEE RAPID 2027","Conference",
   "single-photon sources;quantum communication",
   "2027-08-11","2027-08-13","Miramar Beach, FL, USA",
   "TBA","TBA","TBA","TBA",
   "https://ieee-rapid.org/",2,
   "Defense photonics conference",
   "upcoming","no change",
   "https://ieee-rapid.org/"),

ev("ipc-2027",
   "IEEE Photonics Conference 2027","IPC 2027","Conference",
   "integrated quantum photonics;photonic integrated circuits;silicon photonics",
   "2027-11-10","2027-11-14","Eindhoven, Netherlands",
   "TBA","TBA","TBA","TBA",
   "https://ieee-ipc.org/",4,
   "IEEE photonics flagship; European edition in Eindhoven — low travel cost",
   "upcoming","no change",
   "https://ieee-ipc.org/"),

# ════════════════════════════════════════════════════════════════════════
#  PLACEHOLDERS — dates TBA
# ════════════════════════════════════════════════════════════════════════
ev("photonics-switching-2026",
   "Photonics in Switching and Computing 2026","PiS 2026","Conference",
   "photonic integrated circuits;silicon photonics",
   "TBA","TBA","TBA",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Photonic switching and routing conference",
   "upcoming","no change; dates TBA (expected summer 2026)",
   "original list"),

ev("it-fab-school-2026",
   "It-fab Italian Network for Micro and Nano Fabrication School 2026",
   "It-fab School 2026","Summer School",
   "photonic integrated circuits;nanophotonics",
   "TBA","TBA","TBA",
   "TBA","n/a","TBA","n/a","TBA",3,
   "Italian nanofabrication school relevant to photonic device processing",
   "upcoming","no change; dates TBA (expected Oct 2026)",
   "original list"),

ev("opic",
   "Optics & Photonics International Congress","OPIC","Conference",
   "integrated quantum photonics;quantum optics;nanophotonics",
   "TBA","TBA","TBA",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Japanese international photonics congress with quantum tracks",
   "upcoming","placeholder; next edition dates TBA",
   "original list"),

ev("owtnm",
   "Optical Wave & Waveguide Theory and Modelling Workshop","OWTNM",
   "Workshop","photonic integrated circuits;nanophotonics",
   "TBA","TBA","TBA",
   "TBA","n/a","TBA","n/a","TBA",3,
   "Waveguide modelling workshop relevant to PIC simulation",
   "upcoming","placeholder; next edition dates TBA",
   "original list"),

ev("opon-2025",
   "OPON 2025 Münster","OPON 2025","Conference",
   "integrated quantum photonics;nanophotonics",
   "TBA","TBA","Münster, Germany",
   "TBA","TBA","TBA","TBA","TBA",4,
   "Local Münster photonics conference — direct access, no travel",
   "upcoming","no change; dates TBA",
   "original list"),

ev("nrw-nano-2026",
   "NRW NanoConference 2026","NRW Nano 2026","Conference",
   "nanophotonics;photonic integrated circuits",
   "TBA","TBA","North Rhine-Westphalia, Germany",
   "TBA","TBA","TBA","TBA","TBA",3,
   "Regional NRW nanotechnology conference; local access",
   "upcoming","no change; dates TBA",
   "original list"),

ev("icqe",
   "International Conference on Quantum Energy","ICQE","Conference",
   "quantum computing hardware",
   "TBA","TBA","TBA",
   "TBA","TBA","TBA","TBA","TBA",2,
   "Quantum energy applications; adjacent to quantum computing hardware",
   "upcoming","placeholder; next edition TBA",
   "original list"),

]  # end EVENTS


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
    print(f"  TSV  → {path}  ({len(events)} rows)")


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
    print(f"  XLSX → {path}")


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
