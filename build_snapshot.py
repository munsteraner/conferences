#!/usr/bin/env python3
"""
Build conferences_2026-08-09.tsv  and  conferences_2026-08-09.xlsx
for AG Schuck quantum-photonics conference tracker.
"""

import csv, os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

TODAY     = date(2026, 8, 9)
SNAPSHOT  = "2026-08-09"
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

def ev(id, name, acronym, typ, tags, sd, ed, loc,
       abs_dl, pap_dl, reg_dl, eb_dl,
       url, score, note, status, lv, change, srcs):
    return {
        "id": id, "name": name, "acronym": acronym, "type": typ,
        "topic_tags": tags, "start_date": sd, "end_date": ed,
        "location": loc,
        "abstract_deadline": abs_dl, "paper_deadline": pap_dl,
        "registration_deadline": reg_dl, "early_bird_deadline": eb_dl,
        "website": url, "relevance_score": score,
        "relevance_note": note, "status": status,
        "last_verified": lv, "last_change": change,
        "source_urls": srcs,
    }

EVENTS = [
ev("benasque-qnp-2025", "Quantum Nanophotonics 2025", "Benasque QNP", "Workshop", "nanophotonics;quantum optics;single-photon sources", "2025-03-16", "2025-03-22", "Benasque, Spain", "2025-01-17", "n/a", "TBA", "n/a", "https://benasque.org/2025quantumnanophotonics/", "5", "Directly on nanophotonics and single-photon sources", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://benasque.org/2025quantumnanophotonics/"),
ev("photonics-north-2025", "Photonics North 2025", "PN 2025", "Conference", "silicon photonics;integrated quantum photonics;quantum communication", "2025-05-20", "2025-05-23", "Ottawa, Canada", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Canadian photonics flagship with PIC and quantum sessions", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("benasque-sqt-2025", "Spring School on Superconducting Qubit Technology 2025", "Benasque SQT", "Summer School", "quantum computing hardware;cryogenic electronics", "2025-05-21", "2025-05-30", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/2025sqt/", "3", "Cryogenic electronics for qubit control, adjacent to photonic detector systems", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://benasque.org/2025sqt/"),
ev("optica-quantum-2025", "Optica Quantum 2.0 Conference and Exhibition 2025", "Optica Quantum 2025", "Conference", "integrated quantum photonics;quantum optics;single-photon sources;quantum communication", "2025-06-01", "2025-06-05", "San Francisco, CA, USA", "TBA", "TBA", "TBA", "TBA", "https://optica.org/events/topical_meetings/quantum/", "5", "Premier global quantum photonics conference covering PICs, single photons, QKD", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://optica.org/events/topical_meetings/quantum/"),
ev("icols-2025", "26th International Conference on Laser Spectroscopy", "ICOLS 2025", "Conference", "quantum optics;nonlinear optics", "2025-06-02", "2025-06-07", "Elba Island, Italy", "2025-01-17", "TBA", "TBA", "TBA", "https://icols2025.it/", "3", "Laser spectroscopy with quantum optics relevance", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://icols2025.it/"),
ev("ecio-2025", "European Conference on Integrated Optics 2025", "ECIO 2025", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics", "2025-06-10", "2025-06-12", "Cardiff, Wales, UK", "TBA", "TBA", "TBA", "TBA", "https://ecio-conference.org/", "5", "Largest European dedicated integrated optics conference", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://ecio-conference.org/"),
ev("benasque-qi-2025", "Quantum Information 2025", "Benasque QI", "Workshop", "quantum communication;quantum optics", "2025-06-15", "2025-07-04", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/", "3", "Quantum information including QKD and networking", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://benasque.org/"),
ev("damop-2025", "56th APS Division of Atomic, Molecular and Optical Physics Meeting", "DAMOP 2025", "Conference", "quantum optics;nonlinear optics", "2025-06-16", "2025-06-20", "Portland, OR, USA", "2025-01-31", "2025-01-31", "TBA", "TBA", "https://aps.org/events/2025/atomic-molecular-optical-physics", "3", "AMO physics flagship with quantum optics sessions", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://aps.org/events/2025/atomic-molecular-optical-physics"),
ev("cleo-europe-2025", "CLEO/Europe-EQEC 2025", "CLEO/Europe 2025", "Conference", "integrated quantum photonics;quantum optics;nonlinear optics;nanophotonics", "2025-06-23", "2025-06-27", "Munich, Germany", "2025-01-27", "2025-01-27", "TBA", "TBA", "https://cleoeurope.org/", "5", "Flagship European photonics and quantum optics conference (biennial)", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://cleoeurope.org/"),
ev("cewqo-2025", "29th Central European Workshop on Quantum Optics", "CEWQO 2025", "Workshop", "quantum optics;single-photon sources;quantum communication", "2025-06-23", "2025-06-27", "Vilnius, Lithuania", "TBA", "n/a", "TBA", "n/a", "https://ucan.physics.utoronto.ca/conferences/cewqo29/", "4", "Dedicated European workshop on quantum optics and single-photon physics", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://ucan.physics.utoronto.ca/conferences/cewqo29/"),
ev("icton-2025", "25th International Conference on Transparent Optical Networks", "ICTON 2025", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2025-07-06", "2025-07-10", "Barcelona, Spain", "TBA", "TBA", "TBA", "TBA", "https://icton2025.upc.edu/", "3", "Transparent optical networks with PIC and quantum photonics tracks", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://icton2025.upc.edu/"),
ev("ieee-sum-2025", "IEEE Summer Topicals Meeting Series 2025", "IEEE SUM 2025", "Conference", "integrated quantum photonics;silicon photonics", "2025-07-21", "2025-07-23", "Berlin, Germany", "TBA", "TBA", "TBA", "TBA", "https://ieee-sum.org/", "3", "IEEE topical photonics sessions including PICs", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://ieee-sum.org/"),
ev("qcrypt-2025", "QCrypt 2025", "QCrypt 2025", "Conference", "quantum communication;QKD;quantum networking", "2025-08-25", "2025-08-29", "Sanya, China", "TBA", "TBA", "TBA", "TBA", "https://qcrypt.net/", "4", "Premier annual quantum cryptography and QKD conference", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://qcrypt.net/"),
ev("wqed-2025", "Workshop on Waveguide QED 2025", "WQED 2025", "Workshop", "integrated quantum photonics;single-photon sources;quantum optics", "2025-09-01", "2025-09-06", "Erice/Mazara del Vallo, Sicily, Italy", "TBA", "n/a", "TBA", "n/a", "https://quantum.unipa.it/events/wqed2025/", "4", "Waveguide quantum electrodynamics theory/experiment - core to integrated quantum photonics and single-photon-source physics", "archived", "2026-08-09", "new entry (discovered retroactively); event concluded; series roughly biennial (2018,2020,2021,2023,2025)",
   "https://quantum.unipa.it/events/wqed2025/"),
ev("ecoc-2025", "51st European Conference on Optical Communication", "ECOC 2025", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2025-09-28", "2025-10-02", "Copenhagen, Denmark", "TBA", "TBA", "TBA", "TBA", "https://ecoc2025.org/", "4", "Europe's biggest optical comms conference with PIC and quantum tracks", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://ecoc2025.org/"),
ev("nrw-nano-2025", "11th NRW NanoConference 2025", "NRW Nano 2025", "Conference", "nanophotonics;photonic integrated circuits", "2025-09-30", "2025-10-01", "Dortmund, Germany", "TBA", "TBA", "TBA", "TBA", "https://www.nanoconference.de/", "3", "Regional NRW nanotechnology conference; local access", "archived", "2026-08-09", "new entry (discovered retroactively); confirmed 11th edition; supersedes placeholder nrw-nano-2026",
   "https://www.nanoconference.de/; https://www.ivam.com/news/the-program-for-the-11th-nrw-nano-conference-is-online-get-your-ticket-now-at-the-early-bird-rate?lang=en"),
ev("moc-2025", "30th Microoptics Conference", "MOC 2025", "Conference", "photonic integrated circuits;nanophotonics", "2025-10-12", "2025-10-15", "Tochigi, Japan", "2025-05-16", "TBA", "TBA", "TBA", "TBA", "3", "Micro- and nano-optics with PIC sessions", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("fio-2025", "Frontiers in Optics + Laser Science 2025", "FiO+LS 2025", "Conference", "quantum optics;nonlinear optics;nanophotonics", "2025-10-27", "2025-10-30", "Denver, CO, USA", "TBA", "TBA", "TBA", "TBA", "https://frontiersinoptics.com/", "3", "Broad US optics conference with quantum photonics components", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://frontiersinoptics.com/"),
ev("piers-fall-2025", "Photonics & Electromagnetics Research Symposium — Fall 2025", "PIERS Fall 2025", "Symposium", "photonic integrated circuits;nanophotonics", "2025-11-04", "2025-11-08", "Chiba, Japan", "TBA", "TBA", "TBA", "TBA", "TBA", "2", "General EM/photonics symposium; peripherally relevant", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("acp-2025", "Asia Communications and Photonics Conference 2025", "ACP 2025", "Conference", "photonic integrated circuits;quantum communication", "2025-11-05", "2025-11-08", "Jiangsu, China", "2025-08-01", "TBA", "TBA", "TBA", "TBA", "3", "Asia-Pacific photonics flagship with quantum sessions", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("ipc-2025", "IEEE Photonics Conference 2025", "IPC 2025", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics", "2025-11-09", "2025-11-13", "Singapore", "TBA", "TBA", "TBA", "TBA", "https://ieee-ipc.org/", "4", "IEEE flagship photonics devices and integrated systems conference", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "https://ieee-ipc.org/"),
ev("optima-2025", "Optical Communication, Photonics, Telecommunications, and Intelligent Machine Applications", "OPTIMA 2025", "Conference", "photonic integrated circuits", "2025-12-04", "2025-12-05", "Tashkent, Uzbekistan", "TBA", "TBA", "TBA", "TBA", "TBA", "2", "Regional optics/photonics conference; limited quantum scope", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("pgc-2025", "Photonics Global Conference 2025", "PGC 2025", "Conference", "photonic integrated circuits;nanophotonics", "2025-12-05", "2025-12-08", "Sanya, China", "TBA", "TBA", "TBA", "TBA", "TBA", "2", "Asia-Pacific photonics conference; some PIC tracks", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("icisct-2025", "International Conference on Information Science and Communications Technologies", "ICISCT 2025", "Conference", "quantum communication", "2025-12-26", "2025-12-27", "Kathmandu, Nepal", "TBA", "TBA", "TBA", "TBA", "TBA", "2", "Broad ICT conference; limited photonics scope", "archived", "2026-05-13", "year roll-over from 2025 edition",
   "original list"),
ev("photonics-west-2026", "SPIE Photonics West 2026", "PW 2026", "Conference", "integrated quantum photonics;silicon photonics;nonlinear optics;single-photon sources", "2026-01-17", "2026-01-22", "San Francisco, CA, USA", "2025-07-09", "TBA", "TBA", "TBA", "https://spie.org/photonics-west", "5", "SPIE flagship with Quantum West and Si-photonics tracks", "archived", "2026-05-13", "year roll-over; Photonics West 2027 now announced (Jan 30–Feb 4)",
   "https://spie.org/photonics-west"),
ev("qip-2026", "29th Quantum Information Processing Conference", "QIP 2026", "Conference", "quantum communication;quantum computing hardware;quantum networking", "2026-01-26", "2026-01-30", "Riga, Latvia", "2025-09-12", "2025-10-10", "TBA", "TBA", "https://qip2026.lu.lv/", "4", "Premier quantum information conference; networking and photonic hardware sessions", "archived", "2026-05-13", "year roll-over; QIP 2027 Singapore announced",
   "https://qip2026.lu.lv/"),
ev("baltic-qp-2026", "International Conference on Quantum Photonics Development in the Baltic Region", "Baltic QP 2026", "Conference", "integrated quantum photonics;quantum optics;quantum communication", "2026-02-11", "2026-02-13", "Riga, Latvia", "TBA", "TBA", "TBA", "TBA", "https://toeqpl.eu/conference/", "4", "EU ToEQPL-project conference dedicated to quantum optics, integrated photonics, quantum communication and sensing", "archived", "2026-08-09", "new entry (discovered retroactively); event concluded before this snapshot; recurrence unconfirmed",
   "https://toeqpl.eu/conference/"),
ev("dpg-samop-2026", "DPG Frühjahrstagung SAMOP 2026", "DPG SAMOP 2026", "Conference", "quantum optics;nanophotonics;integrated quantum photonics", "2026-03-01", "2026-03-06", "Mainz, Germany", "TBA", "TBA", "TBA", "TBA", "https://mainz26.dpg-tagungen.de", "4", "German physics society AMO/photonics spring meeting — easily accessible from Münster", "archived", "2026-05-13", "year roll-over; DPG SAMOP 2027 Hannover announced",
   "https://mainz26.dpg-tagungen.de"),
ev("nanolight-2026", "Nanolight Conference 2026", "Nanolight 2026", "Conference", "nanophotonics;2D-material photonics;quantum optics;plasmonics", "2026-03-02", "2026-03-06", "Benasque, Spain", "2025-12-15", "n/a", "2026-02-15", "n/a", "https://benasque.org/2026nanolight/", "5", "European nanophotonics / near-field / 2D-material flagship", "archived", "2026-05-13", "year roll-over",
   "https://benasque.org/2026nanolight/"),
ev("dpg-skm-2026", "DPG Frühjahrstagung SKM 2026", "DPG SKM 2026", "Conference", "silicon photonics;nanophotonics;2D-material photonics", "2026-03-08", "2026-03-13", "Dresden, Germany", "TBA", "TBA", "TBA", "TBA", "https://dresden26.dpg-tagungen.de", "3", "Condensed matter sessions on photonics materials and devices", "archived", "2026-05-13", "year roll-over",
   "https://dresden26.dpg-tagungen.de"),
ev("ofc-2026", "Optical Fiber Communication Conference 2026", "OFC 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2026-03-15", "2026-03-19", "Los Angeles, CA, USA", "TBA", "TBA", "TBA", "TBA", "https://ofcconference.org/", "4", "World's largest optical comms conference with PIC and quantum networking sessions", "archived", "2026-05-13", "year roll-over; OFC 2027 Los Angeles Mar 7-11 announced",
   "https://ofcconference.org/"),
ev("aps-march-2026", "APS Global Physics Summit 2026", "APS 2026", "Conference", "quantum optics;quantum computing hardware;integrated quantum photonics", "2026-03-16", "2026-03-20", "Denver, CO, USA", "2025-10-31", "2025-10-31", "TBA", "TBA", "https://summit.aps.org/", "4", "Largest physics meeting; major quantum photonics and computing sessions", "archived", "2026-05-13", "year roll-over; confirmed as joint March+April Global Physics Summit; APS 2027 Atlanta Apr 11-16",
   "https://summit.aps.org/; https://march.aps.org/"),
ev("spie-pe-2026", "SPIE Photonics Europe 2026", "SPIE PE 2026", "Conference", "integrated quantum photonics;silicon photonics;nanophotonics", "2026-04-12", "2026-04-16", "Strasbourg, France", "TBA", "TBA", "TBA", "TBA", "https://spie.org/photonics-europe", "5", "Major European SPIE photonics conference with quantum integrated sessions", "archived", "2026-05-13", "year roll-over",
   "https://spie.org/photonics-europe"),
ev("ieee-siphotonics-2026", "IEEE Silicon Photonics Conference 2026", "IEEE SiPho 2026", "Conference", "silicon photonics;photonic integrated circuits;III-V on Si", "2026-04-13", "2026-04-15", "Ottawa, ON, Canada", "2025-10-03", "TBA", "TBA", "TBA", "https://ieee-siphotonics.org/", "5", "Dedicated silicon photonics conference — direct relevance to PICs", "archived", "2026-05-13", "year roll-over",
   "https://ieee-siphotonics.org/"),
ev("ieee-qpain-2026", "IEEE 2nd International Conference on Quantum Photonics, AI & Networking", "IEEE QPAIN 2026", "Conference", "integrated quantum photonics;quantum networking", "2026-04-16", "2026-04-18", "Chattogram, Bangladesh", "TBA", "TBA", "TBA", "TBA", "https://qpain.org/", "3", "Quantum photonics and networking conference (IEEE-sponsored)", "archived", "2026-05-13", "year roll-over",
   "https://qpain.org/"),
ev("inp-webinar-2026", "Webinar on InP platform for NIR and MIR applications", "InP Webinar", "Seminar", "photonic integrated circuits;III-V on Si", "2026-04-17", "2026-04-17", "Online", "TBA", "n/a", "TBA", "n/a", "https://inphomir.eu/1st-webinar/", "4", "III-V InP PIC platform — directly relevant to quantum light-source integration", "archived", "2026-05-13", "year roll-over",
   "https://inphomir.eu/1st-webinar/"),
ev("eftf-2026", "European Frequency and Time Forum 2026", "EFTF 2026", "Conference", "quantum optics;quantum communication", "2026-04-20", "2026-04-23", "Noordwijk, The Netherlands", "TBA", "TBA", "TBA", "TBA", "https://www.eftf.org/home", "2", "Frequency/time standards; quantum photonic metrology peripherally relevant", "archived", "2026-05-13", "year roll-over",
   "https://www.eftf.org/home"),
ev("spie-ds-2026", "SPIE Defense + Security 2026", "SPIE D+S 2026", "Conference", "single-photon sources;quantum communication", "2026-04-26", "2026-04-30", "National Harbor, MD, USA", "TBA", "TBA", "TBA", "TBA", "https://spie.org/defense-and-commercial-sensing", "2", "Defense photonics with some single-photon sensing sessions", "archived", "2026-05-13", "year roll-over",
   "https://spie.org/defense-and-commercial-sensing"),
ev("bqit-2026", "Bristol Quantum Information Technologies Workshop 2026", "BQIT 2026", "Workshop", "integrated quantum photonics;quantum communication;quantum computing hardware", "2026-04-27", "2026-04-29", "Bristol, UK", "TBA", "n/a", "TBA", "n/a", "https://www.bristol.ac.uk/qet-labs/events/bqit-workshop/", "5", "UK quantum photonics workshop — PICs, single photons, quantum networking", "archived", "2026-05-13", "year roll-over",
   "https://www.bristol.ac.uk/qet-labs/events/bqit-workshop/"),
ev("quantumatter-2026", "QUANTUMatter 2026", "QUANTU 2026", "Conference", "quantum computing hardware;quantum optics;nanophotonics", "2026-04-27", "2026-04-30", "Barcelona, Spain", "2026-02-11", "n/a", "2026-03-06", "2026-02-23", "https://www.quantumconf.eu/2026/", "3", "Quantum materials and devices, including photonic quantum systems", "archived", "2026-05-13", "year roll-over",
   "https://www.quantumconf.eu/2026/"),
ev("ampd-2026", "Annual Meeting Photonic Devices 2026", "AMPD 2026", "Conference", "photonic integrated circuits;nanophotonics", "2026-04-28", "2026-04-30", "TBA", "TBA", "TBA", "TBA", "TBA", "https://www.zib.de/workshop-photonic-devices/", "4", "German photonic devices annual meeting — directly relevant to group's fabrication work", "archived", "2026-05-13", "year roll-over",
   "https://www.zib.de/workshop-photonic-devices/"),
ev("cleo-2026", "Conference on Lasers and Electro-Optics 2026", "CLEO 2026", "Conference", "integrated quantum photonics;quantum optics;nonlinear optics;silicon photonics;single-photon sources", "2026-05-17", "2026-05-21", "Charlotte, NC, USA", "2025-11-18", "n/a", "TBA", "TBA", "https://cleoconference.org/", "5", "Premier US laser/photonics conference; major IQP, Si-PIC and quantum optics sessions", "archived", "2026-08-09", "confirmed occurred as scheduled, Charlotte NC; CLEO 2027 now announced (Long Beach, CA, May 2-7)",
   "https://cleoconference.org/; https://cleoconference.org/registration-2026/"),
ev("damop-2026", "57th APS Division of Atomic, Molecular and Optical Physics Meeting", "DAMOP 2026", "Conference", "quantum optics;nonlinear optics;single-photon sources", "2026-06-01", "2026-06-05", "Providence, RI, USA", "TBA", "TBA", "TBA", "TBA", "https://aps.org/events/2026/damop-meeting-2026", "3", "AMO physics flagship with quantum optics and single-photon sessions", "archived", "2026-08-09", "confirmed occurred, Providence RI; DAMOP 2027 announced (Chicago, IL, May 13-18)",
   "https://aps.org/events/2026/damop-meeting-2026"),
ev("photonics-north-2026", "Photonics North 2026", "PN 2026", "Conference", "silicon photonics;integrated quantum photonics;quantum communication", "2026-06-01", "2026-06-04", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Canadian national photonics conference with PIC and quantum sessions", "archived", "2026-08-09", "corrected dates/location: confirmed Jun 2-5 2026, Québec City (was listed Jun1-4, location TBA); 2027 edition not yet announced",
   "original list"),
ev("cargese-qt-2026", "Summer School on Quantum Technologies for Computation and Communication", "QT4CC 2026", "Summer School", "quantum communication;quantum computing hardware;quantum networking", "2026-06-08", "2026-06-20", "Cargèse (Corsica), France", "TBA", "n/a", "TBA", "n/a", "https://qt4cc.sciencesconf.org/", "4", "PhD summer school on quantum computation and communication — ideal for group's PhD students", "archived", "2026-08-09", "confirmed occurred, Cargèse; next QT4CC edition not yet announced",
   "https://qt4cc.sciencesconf.org/; https://gdr-teq.cnrs.fr/cevent/629/"),
ev("spie-pfq-2026", "SPIE Photonics for Quantum 2026", "PfQ 2026", "Conference", "integrated quantum photonics;single-photon sources;quantum communication;quantum computing hardware", "2026-06-08", "2026-06-11", "Waterloo, ON, Canada", "2026-03-01", "n/a", "TBA", "TBA", "https://spie.org/conferences-and-exhibitions/photonics-for-quantum", "5", "Dedicated quantum photonics conference — bullseye for AG Schuck", "archived", "2026-08-09", "event concluded, Waterloo ON",
   "https://spie.org/PFQ26/conferencedetails/photonics-for-quantum"),
ev("islc-2026", "30th International Semiconductor Laser Conference", "ISLC 2026", "Conference", "single-photon sources;III-V on Si;photonic integrated circuits", "2026-06-14", "2026-06-17", "Tampere, Finland", "2026-02-05", "TBA", "TBA", "TBA", "https://events.tuni.fi/islc2026/", "4", "III-V semiconductor laser conference — key for quantum light source development", "archived", "2026-08-09", "confirmed occurred, Tampere; ISLC is biennial, next (31st, ~2028) not yet announced",
   "https://events.tuni.fi/islc2026/; https://ieeephotonics.org/event/2026-30th-international-semiconductor-laser-conference-islc/"),
ev("grc-quantum-science-2026", "Gordon Research Conference — Quantum Science 2026", "GRC QS 2026", "Conference", "quantum optics;quantum computing hardware;quantum communication;integrated quantum photonics", "2026-06-14", "2026-06-19", "Easton, MA, USA", "TBA", "n/a", "2026-06-28", "n/a", "https://www.grc.org/quantum-science-conference/2026/", "4", "GRC on quantum science covering photonic quantum systems and quantum information", "archived", "2026-08-09", "event concluded; GRC is biennial, 2028 edition not yet announced",
   "https://www.grc.org/quantum-science-conference/2026/"),
ev("icap-2026", "29th International Conference on Atomic Physics", "ICAP 2026", "Conference", "quantum optics;nonlinear optics", "2026-06-14", "2026-06-19", "Wuhan, China", "TBA", "TBA", "TBA", "TBA", "https://www.icap29.com/home.html", "3", "Atomic physics conference with quantum optics and light-matter interaction sessions", "archived", "2026-08-09", "event concluded, Wuhan; next ICAP not yet announced",
   "https://www.icap29.com/home.html"),
ev("ecio-2026", "European Conference on Integrated Optics 2026", "ECIO 2026", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics;nanophotonics", "2026-06-15", "2026-06-17", "Zürich, Switzerland", "TBA", "TBA", "2026-06-01", "2026-05-11", "https://ecio-conference.org/", "5", "Largest European integrated optics conference — core relevance for AG Schuck", "archived", "2026-08-09", "confirmed occurred, ETH Zürich (27th ECIO); ECIO 2027 not yet announced",
   "https://www.ecio-conference.org/; https://www.ecio-conference.org/deadlines/"),
ev("optica-quantum-2026", "Optica Quantum 2.0 Conference and Exhibition 2026", "Optica Quantum 2026", "Conference", "integrated quantum photonics;quantum optics;single-photon sources;quantum communication;quantum computing hardware", "2026-06-15", "2026-06-18", "Glasgow, UK", "2026-02-10", "n/a", "2026-05-29", "n/a", "https://www.optica.org/events/topical_meetings/quantum/", "5", "Premier quantum photonics conference spanning PICs, detectors, QKD, and hardware", "archived", "2026-08-09", "event concluded, Glasgow",
   "https://www.optica.org/events/topical_meetings/quantum/; https://www.optica.org/events/topical_meetings/quantum/registration/"),
ev("icop-2026", "Italian Conference on Optics and Photonics 2026", "ICOP 2026", "Conference", "integrated quantum photonics;quantum optics;nanophotonics", "2026-06-15", "2026-06-17", "L'Aquila, Italy", "TBA", "TBA", "TBA", "TBA", "https://www.icop2026.it/", "3", "Italian national photonics conference with quantum tracks", "archived", "2026-08-09", "event concluded, L'Aquila; ICOP confirmed biennial, next expected ~2028 (not yet announced)",
   "https://www.icop2026.it/"),
ev("ccw-london-2026", "Critical Communications World London 2026", "CCW 2026", "Conference", "quantum communication;quantum networking", "2026-06-16", "2026-06-18", "London, UK", "TBA", "TBA", "TBA", "TBA", "https://www.critical-communications-world.com/", "2", "Critical comms industry event; QKD sessions of marginal relevance", "archived", "2026-08-09", "event concluded, London",
   "https://www.critical-communications-world.com/call-for-content"),
ev("egas-2026", "57th Conference of the European Group on Atomic Systems", "EGAS 2026", "Conference", "quantum optics;nonlinear optics", "2026-06-21", "2026-06-25", "Toruń, Poland", "TBA", "TBA", "TBA", "TBA", "https://egas57.org/", "3", "European atomic systems conference with quantum optics and light-matter sessions", "archived", "2026-08-09", "event concluded, Toruń (EGAS57); EGAS58 (2027) not yet confirmed",
   "https://egas57.org/"),
ev("benasque-qsi-2026", "Quantum Science: Implementations 2026", "Benasque QSI", "Workshop", "integrated quantum photonics;quantum computing hardware;quantum communication", "2026-06-21", "2026-07-04", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/", "4", "Benasque workshop on quantum implementations including photonic platforms", "archived", "2026-08-09", "event concluded, Benasque",
   "https://benasque.org/; https://www.benasque.org/new_general/cgi-bin/years.pl?ano=2026"),
ev("les-houches-pa-2026", "Photons & Atoms 2026 — Les Houches Doctoral School", "Les Houches PA 2026", "Summer School", "quantum optics;nonlinear optics;integrated quantum photonics;single-photon sources", "2026-06-22", "2026-07-03", "Les Houches, France", "TBA", "n/a", "TBA", "n/a", "https://photon-atoms-26.leshouches.science/", "5", "Leading European doctoral school on photons, quantum gases and quantum technologies", "archived", "2026-08-09", "event concluded, Les Houches; next AMO/photonics doctoral school timing is irregular (2020/2023/2024/2026), not yet announced",
   "https://photon-atoms-26.leshouches.science/; https://first-tf.com/doctoral-training-on-atoms-and-photons-ecole-de-physique-des-houches-france-june-22nd-to-july-3rd-2026/"),
ev("iqt-nordics-2026", "IQT Nordics 2026 — Quantum Technology in a Changing World", "IQT Nordics 2026", "Conference", "quantum communication;quantum computing hardware;integrated quantum photonics", "2026-06-22", "2026-06-24", "Oslo, Norway", "TBA", "n/a", "TBA", "TBA", "https://iqtevent.com/nordics/", "3", "Nordic quantum industry conference with hardware and QKD sessions", "archived", "2026-08-09", "event concluded, Oslo",
   "https://iqtevent.com/nordics/; https://www.oslomet.no/en/about/events/iqt-nordics-2026"),
ev("quantum-tech-world-2026", "Quantum.Tech World 2026", "Q.Tech World 2026", "Conference", "quantum computing hardware;quantum communication;integrated quantum photonics", "2026-06-25", "2026-06-26", "Boston, MA, USA", "TBA", "n/a", "TBA", "TBA", "https://www.alphaevents.com/events-quantumtechus", "2", "Industry quantum technology event; limited deep-science content", "archived", "2026-08-09", "event concluded, Boston",
   "https://www.alphaevents.com/events-quantumtechus"),
ev("oecc-2026", "31st OptoElectronics and Communications Conference", "OECC 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2026-06-28", "2026-07-02", "Busan, South Korea", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Asia-Pacific optoelectronics and comms conference with PIC sessions", "archived", "2026-08-09", "event concluded, Busan",
   "original list"),
ev("quantum-korea-2026", "Quantum Korea 2026", "Quantum Korea 2026", "Conference", "quantum computing hardware;quantum communication;integrated quantum photonics", "2026-07-02", "2026-07-04", "Seoul, South Korea", "TBA", "n/a", "TBA", "TBA", "https://quantum-korea.kr/en/main", "2", "Government-organised Korean quantum technology showcase", "archived", "2026-08-09", "event concluded, Seoul",
   "https://quantum-korea.kr/en/main"),
ev("spw-2026", "12th Single Photon Workshop", "SPW 2026", "Workshop", "single-photon sources;SNSPDs;quantum optics;quantum communication", "2026-07-06", "2026-07-10", "Naples, Italy", "TBA", "n/a", "TBA", "TBA", "https://spw2026.org/", "5", "THE dedicated single-photon workshop — central to AG Schuck's detector and source work", "archived", "2026-08-09", "event concluded, Naples (300+ participants); next SPW tentatively reported Nov 2027, Singapore (aggregator-only, UNCONFIRMED - verify with organizers)",
   "https://spw2026.org/; https://iqnhub.org/2026-single-photon-workshop-registration-open/"),
ev("icton-2026", "26th International Conference on Transparent Optical Networks", "ICTON 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2026-07-12", "2026-07-16", "Prague, Czech Republic", "TBA", "TBA", "TBA", "TBA", "https://www.gov.pl/web/instytut-lacznosci/icton-2026", "3", "Transparent optical networks with PIC and quantum photonics tracks", "archived", "2026-08-09", "event concluded, Prague",
   "https://www.gov.pl/web/instytut-lacznosci/icton-2026"),
ev("ieee-sum-2026", "IEEE Summer Topicals Meeting Series 2026", "IEEE SUM 2026", "Conference", "photonic integrated circuits;silicon photonics;integrated quantum photonics", "2026-07-13", "2026-07-15", "Tulum, Mexico", "TBA", "TBA", "TBA", "TBA", "https://ieee-sum.org/", "3", "IEEE focused topical photonics sessions", "archived", "2026-08-09", "event concluded, Tulum",
   "https://ieee-sum.org/"),
ev("meta-2026", "META Conference 2026 (Quantum Workshop track)", "META 2026", "Conference", "nanophotonics;plasmonics;integrated quantum photonics", "2026-07-14", "2026-07-17", "Dublin, Ireland", "TBA", "TBA", "TBA", "TBA", "https://metaconferences.org/META26/", "3", "16th metamaterials/photonic-crystals/plasmonics conference; 2026 Quantum Workshop track covers integrated quantum photonic circuits and plasmonic quantum interfaces", "archived", "2026-08-09", "new entry (discovered retroactively); event concluded",
   "https://metaconferences.org/META26/index.php/META/workshop"),
ev("grc-plasmonics-nano-2026", "Gordon Research Conference — Plasmonics and Nanophotonics 2026", "GRC PNP 2026", "Conference", "nanophotonics;plasmonics;quantum optics;2D-material photonics", "2026-07-19", "2026-07-24", "Newry, ME, USA", "TBA", "n/a", "2026-06-21", "n/a", "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/", "5", "GRC on quantum nanophotonics, plasmonics and 2D photonics — highly relevant to AG Schuck", "archived", "2026-08-09", "event concluded, Newry ME; GRC biennial, 2028 not yet announced",
   "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/"),
ev("cewqo-2026", "30th Central European Workshop on Quantum Optics", "CEWQO30 2026", "Workshop", "quantum optics;single-photon sources;nonlinear optics;integrated quantum photonics", "2026-07-20", "2026-07-24", "Erlangen, Germany", "TBA", "n/a", "TBA", "n/a", "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/", "5", "European quantum optics workshop co-hosted by FAU/MPL Erlangen — close to Münster", "archived", "2026-08-09", "event concluded, Erlangen; CEWQO biennial, next ~2028 not yet announced",
   "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/; https://www.oqt.nat.fau.de/events/"),
ev("apc-2026", "Optica Advanced Photonics Congress 2026", "APC 2026", "Conference", "integrated quantum photonics;nonlinear optics;nanophotonics", "2026-07-26", "2026-07-30", "Long Beach, CA, USA", "TBA", "TBA", "TBA", "TBA", "https://www.optica.org/events/congress/advanced_photonics_congress/", "5", "Optica Advanced Photonics Congress incl. Integrated Photonics Research (IPR) quantum-photonics subtrack, Nonlinear Photonics, and BGPP sub-conferences", "archived", "2026-08-09", "event concluded, Long Beach CA (location corrected from TBA); confirmed includes IPR/NP/BGPP sub-conferences; APC 2027 not yet announced",
   "https://www.optica.org/events/congress/advanced_photonics_congress/"),
ev("grc-mech-quant-2026", "Gordon Research Conference — Mechanical Systems in the Quantum Regime 2026", "GRC MSQR 2026", "Conference", "optomechanics;quantum computing hardware", "2026-07-26", "2026-07-31", "Lucca (Barga), Italy", "TBA", "n/a", "TBA", "n/a", "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/", "4", "Optomechanics GRC — relevant to cavity-optomechanical quantum transducers", "archived", "2026-08-09", "event concluded, Lucca; GRC biennial, 2028 not yet announced",
   "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/"),
ev("lake-como-ufqp-2026", "New Frontiers in Ultrafast Quantum Optics — Lake Como School", "UFQP 2026", "Summer School", "quantum optics;nonlinear optics;single-photon sources", "2026-07-27", "2026-07-31", "Como, Italy", "2026-05-31", "n/a", "TBA", "n/a", "https://ufqp.lakecomoschool.org/", "4", "Summer school on ultrafast quantum optics — photon correlations and time-frequency entanglement", "archived", "2026-08-09", "event concluded, Como; not annual (prior edition 2022), next timing uncertain",
   "https://ufqp.lakecomoschool.org/; https://www.quantiki.org/conference/summer-school-new-frontiers-ultrafast-quantum-optics"),
ev("ieee-rapid-2026", "IEEE Research and Applications of Photonics in Defense 2026", "IEEE RAPID 2026", "Conference", "single-photon sources;quantum communication", "2026-08-05", "2026-08-07", "Miramar Beach, FL, USA", "TBA", "TBA", "TBA", "TBA", "https://ieee-rapid.org/", "2", "Defense photonics; some single-photon sensing sessions", "archived", "2026-08-09", "event concluded, Miramar Beach FL; IEEE RAPID 2027 already tracked",
   "https://ieee-rapid.org/"),
ev("qels-2026", "Quantum Engineering of Levitated Systems 2026", "QELS 2026", "Workshop", "optomechanics;quantum optics", "2026-08-16", "2026-08-22", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/2026qels/", "2", "10th-anniversary Benasque workshop on levitated mesoscopic quantum systems; adjacent to optomechanics", "upcoming", "2026-08-09", "new entry",
   "https://benasque.org/2026qels/"),
ev("spie-op-2026", "SPIE Optics + Photonics 2026", "SPIE O+P 2026", "Conference", "integrated quantum photonics;nanophotonics;quantum optics;single-photon sources", "2026-08-23", "2026-08-27", "San Diego, CA, USA", "TBA", "TBA", "TBA", "TBA", "https://spie.org/optics-and-photonics", "4", "SPIE major annual meeting with quantum nanophotonics and quantum technologies tracks", "cfp_closed", "2026-08-09", "regular abstract deadline has passed; post-deadline submissions currently being accepted",
   "https://spie.org/optics-and-photonics"),
ev("qcrypt-2026", "QCrypt 2026", "QCrypt 2026", "Conference", "quantum communication;QKD;quantum networking", "2026-08-24", "2026-08-28", "Ottawa, Canada", "TBA", "2026-03-13", "TBA", "TBA", "https://qcrypt.net/2026/", "4", "Premier annual QKD and quantum cryptography conference", "cfp_closed", "2026-08-09", "talk submission deadline Mar 13 2026 (closed), poster deadline May 1 2026 (closed); registration open",
   "https://qcrypt.net/2026/"),
ev("eosam-2026", "European Optical Society Annual Meeting 2026", "EOSAM 2026", "Conference", "integrated quantum photonics;nanophotonics;nonlinear optics", "2026-08-24", "2026-08-28", "Tampere, Finland", "2026-04-14", "TBA", "TBA", "2026-06-15", "https://www.europeanoptics.org/events/eos/eosam2026.html", "4", "EOS flagship European photonics meeting with strong quantum photonics programme", "cfp_closed", "2026-08-09", "abstract deadline Apr 14 and early-bird Jun 15 both closed; post-deadline window closed Jul 15",
   "https://www.europeanoptics.org/events/eos/eosam2026.html"),
ev("mwp-2026", "IEEE International Topical Meeting on Microwave Photonics 2026", "MWP 2026", "Conference", "photonic integrated circuits;silicon photonics;III-V on Si", "2026-08-24", "2026-08-27", "Matsue, Japan", "n/a", "2026-03-23", "TBA", "TBA", "https://mwp2026.org/", "3", "IEEE Photonics Society topical meeting on RF/microwave photonics incl. EO modulators (TFLN) and PIC platforms", "cfp_closed", "2026-08-09", "new entry; paper deadline 2026-03-23 (closed)",
   "https://mwp2026.org/; https://ieeephotonics.org/event/2026-international-topical-meeting-on-microwave-photonics-mwp/"),
ev("tqc-2026", "Theory of Quantum Computation, Communication and Cryptography 2026", "TQC 2026", "Conference", "quantum communication;quantum computing hardware;quantum networking", "2026-08-31", "2026-09-04", "Sherbrooke, QC, Canada", "TBA", "2026-04-20", "TBA", "TBA", "https://tqc-conference.org/2026/", "3", "Theory conference for quantum communication, computation and cryptography", "cfp_closed", "2026-08-09", "poster-only submission deadline Apr 20 2026 (closed); main talk deadline unconfirmed",
   "https://tqc-conference.org/2026/"),
ev("asc-2026", "Applied Superconductivity Conference 2026", "ASC 2026", "Conference", "SNSPDs;cryogenic electronics;single-photon sources", "2026-09-06", "2026-09-11", "Pittsburgh, PA, USA", "2026-02-02", "TBA", "TBA", "2026-06-30", "https://appliedsuperconductivity.org/asc2026/", "4", "Premier superconductivity conference with major SNSPD and cryogenic electronics sessions", "cfp_closed", "2026-08-09", "abstract deadline (extended, firm) Feb 2 2026, early registration cutoff Jun 30 2026 (both closed)",
   "https://appliedsuperconductivity.org/asc2026/"),
ev("inphomir-school-2026", "INPHOMIR Photonics School 2026", "INPHOMIR School", "Summer School", "photonic integrated circuits;III-V on Si", "2026-09-07", "2026-09-11", "Bari, Italy", "TBA", "n/a", "TBA", "n/a", "https://inphomir.eu/inphomir-school-program/", "4", "Dedicated InP/III-V photonics school for NIR/MIR applications", "upcoming", "2026-08-09", "no change; note: INPHOMIR Horizon Europe project concludes 2027, further editions uncertain",
   "https://inphomir.eu/inphomir-school-program/"),
ev("ieee-qce26", "IEEE Quantum Week 2026", "QCE26", "Conference", "quantum computing hardware;quantum communication;integrated quantum photonics", "2026-09-13", "2026-09-18", "Toronto, ON, Canada", "2026-04-06", "2026-04-27", "TBA", "TBA", "https://qce.quantum.ieee.org/2026/", "3", "IEEE quantum week with photonic quantum computing and QKD sessions", "cfp_closed", "2026-08-09", "abstract deadline Apr 6, paper deadline extended to Apr 27 (both closed); QCE 2027 location not yet announced",
   "https://qce.quantum.ieee.org/2026/"),
ev("ecoc-2026", "European Conference on Optical Communication 2026", "ECOC 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication;integrated quantum photonics", "2026-09-20", "2026-09-24", "Málaga, Spain", "2026-04-22", "2026-04-22", "TBA", "TBA", "https://ecoc2026.org/", "4", "Europe's flagship optical comms conference with quantum technologies track", "cfp_closed", "2026-08-09", "ECOC 2027 announced (Milan, Oct 10-14); registration cancellation cutoff Jul 20 2026 has passed",
   "https://ecoc2026.org/; https://ecoc2026.org/ECOC2026/paper-submission"),
ev("hdqs-2026", "High-Dimensional Quantum Systems Workshop 2026", "HDQS 2026", "Workshop", "quantum optics;quantum communication;integrated quantum photonics", "2026-09-20", "2026-10-03", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/2026hdqs/", "4", "Benasque workshop on high-dimensional quantum photonics and quantum information", "upcoming", "2026-08-09", "dates corrected per Benasque official listing: Sep 20 - Oct 3 2026 (previously listed Oct 4-9); organizers Kropf & Gebhart (ETH Zürich)",
   "https://benasque.org/2026hdqs/"),
ev("mne-2026", "International Conference on Micro and Nano Engineering 2026", "MNE 2026", "Conference", "photonic integrated circuits;nanophotonics;2D-material photonics", "2026-09-21", "2026-09-24", "Interlaken, Switzerland", "TBA", "TBA", "TBA", "TBA", "https://mne2026.imnes.org/", "3", "Nanofabrication conference relevant to photonic device processing", "upcoming", "2026-05-13", "no change",
   "https://mne2026.imnes.org/"),
ev("baltic-photonics-2026", "Baltic Photonics 2026", "Baltic Photonics 2026", "Conference", "photonic integrated circuits;silicon photonics", "2026-09-24", "2026-09-25", "Vilnius, Lithuania", "TBA", "n/a", "TBA", "TBA", "https://toolas.eu/baltic-photonics-2026/", "2", "Industry-oriented photonics/semiconductor matchmaking event; limited direct overlap with quantum photonics research", "upcoming", "2026-08-09", "new entry",
   "https://toolas.eu/baltic-photonics-2026/"),
ev("fio-2026", "Frontiers in Optics + Laser Science 2026", "FiO+LS 2026", "Conference", "quantum optics;nonlinear optics;nanophotonics", "2026-09-27", "2026-10-01", "Rochester, NY, USA", "2026-01-27", "TBA", "TBA", "TBA", "https://frontiersinoptics.com/", "3", "Broad US photonics conference with quantum optics tracks", "cfp_closed", "2026-08-09", "abstract deadline Jan 27 2026 (closed); confirmed Rochester NY Sep 27-Oct 1 (end date corrected from Sep 30)",
   "https://frontiersinoptics.com/"),
ev("photonics-days-2026", "Photonics Days Berlin 2026", "PDB 2026", "Conference", "integrated quantum photonics;photonic integrated circuits;nanophotonics", "2026-10-07", "2026-10-08", "Berlin, Germany", "TBA", "TBA", "TBA", "TBA", "https://photonic-days-berlin.com/", "3", "German photonics industry and research meeting", "upcoming", "2026-05-13", "no change",
   "https://photonic-days-berlin.com/"),
ev("it-fab-school-2026", "It-fab Italian Network for Micro and Nano Fabrication School 2026", "It-fab School 2026", "Summer School", "photonic integrated circuits;nanophotonics", "2026-10-27", "2026-10-30", "Torino, Italy", "TBA", "n/a", "TBA", "n/a", "TBA", "3", "Italian nanofabrication school relevant to photonic device processing", "upcoming", "2026-08-09", "dates now published: Oct 27-30 2026, hosted by INRiM (first itinerant edition of the school)",
   "original list"),
ev("ipc-2026", "IEEE Photonics Conference 2026", "IPC 2026", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics", "2026-11-08", "2026-11-12", "Denver, CO, USA", "TBA", "2026-05-04", "TBA", "TBA", "https://ieee-ipc.org/", "4", "IEEE flagship photonics devices and integrated systems conference", "cfp_closed", "2026-08-09", "paper submission deadline was 2026-05-04 (now closed); confirmed Denver CO Nov 8-12",
   "https://ieee-ipc.org/"),
ev("eqtc-2026", "European Quantum Technologies Conference 2026", "EQTC 2026", "Conference", "integrated quantum photonics;quantum communication;quantum computing hardware;quantum networking", "2026-11-29", "2026-12-03", "Dublin, Ireland", "2026-08-03", "TBA", "TBA", "TBA", "https://qt.eu/events/eqtc-2026-european-quantum-technologies-conference", "5", "EU Quantum Flagship conference — all aspects of quantum technology, high visibility", "cfp_closed", "2026-08-09", "Science & Technology track abstract deadline Aug 3 2026 (closed)",
   "https://qt.eu/events/eqtc-2026-european-quantum-technologies-conference"),
ev("photonics-west-2027", "SPIE Photonics West 2027", "PW 2027", "Conference", "integrated quantum photonics;silicon photonics;nonlinear optics;single-photon sources", "2027-01-30", "2027-02-04", "San Francisco, CA, USA", "2026-07-22", "TBA", "TBA", "TBA", "https://spie.org/conferences-and-exhibitions/photonics-west", "5", "SPIE flagship with Quantum West and Si-photonics tracks — abstract deadline Jul 22 2026", "cfp_closed", "2026-08-09", "abstract deadline Jul 22 2026 has passed (now closed)",
   "https://spie.org/conferences-and-exhibitions/photonics-west; https://spie.org/conferences-and-exhibitions/photonics-west/program/browse-program"),
ev("qip-2027", "30th Quantum Information Processing Conference", "QIP 2027", "Conference", "quantum communication;quantum computing hardware;quantum networking", "2027-02-20", "2027-02-26", "Singapore", "TBA", "TBA", "TBA", "TBA", "https://qipconference.org/2027/", "4", "Premier annual quantum information conference; quantum networking and hardware sessions", "upcoming", "2026-08-09", "confirmed unchanged; call for submissions not yet released, deadlines still TBA",
   "https://qipconference.org/2027/; https://qip.iaqi.org/nextqip"),
ev("dpg-samop-2027", "DPG Frühjahrstagung SAMOP 2027", "DPG SAMOP 2027", "Conference", "quantum optics;nanophotonics;integrated quantum photonics", "2027-02-28", "2027-03-05", "Hannover, Germany", "TBA", "TBA", "TBA", "TBA", "https://dpg-physik.de/", "4", "DPG spring meeting for quantum optics and photonics — easily accessible from Münster", "upcoming", "2026-05-13", "no change",
   "https://dpg-physik.de/"),
ev("ofc-2027", "Optical Fiber Communication Conference 2027", "OFC 2027", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "2027-03-07", "2027-03-11", "Los Angeles, CA, USA", "TBA", "TBA", "TBA", "TBA", "https://ofcconference.org/", "4", "World's largest optical comms conference with PIC and quantum networking sessions", "upcoming", "2026-05-13", "new entry; Mar 7-11 2027 Los Angeles",
   "https://ofcconference.org/"),
ev("dpg-skm-2027", "DPG Frühjahrstagung SKM 2027", "DPG SKM 2027", "Conference", "silicon photonics;nanophotonics;2D-material photonics", "2027-03-14", "2027-03-19", "Regensburg, Germany", "TBA", "TBA", "TBA", "TBA", "https://www.dpg-physik.de/veranstaltungen/2027/dpg-fruehjahrstagung-regensburg", "3", "Condensed matter sessions on photonics materials and devices; joint with 90th DPG Annual Meeting", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition (Dresden); part of joint DPG Spring Meetings Feb28-Mar19 2027",
   "https://www.dpg-physik.de/aktivitaeten-und-programme/tagungen/fruehjahrstagungen/2027"),
ev("benasque-qnp-2027", "Quantum Nanophotonics 2027", "Benasque QNP 2027", "Workshop", "nanophotonics;quantum optics;single-photon sources", "2027-03-28", "2027-04-02", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/2027quantumnanophotonics/", "5", "Directly on nanophotonics and single-photon sources", "upcoming", "2026-08-09", "new entry; year roll-over from 2025 edition",
   "https://benasque.org/2027quantumnanophotonics/"),
ev("spie-oo-2027", "SPIE Optics + Optoelectronics 2027", "SPIE O+O 2027", "Conference", "integrated quantum photonics;quantum optics;nanophotonics", "2027-04-05", "2027-04-08", "Prague, Czech Republic", "TBA", "TBA", "TBA", "TBA", "TBA", "4", "SPIE European conference with quantum photonics and nanophotonics tracks", "upcoming", "2026-05-13", "no change",
   "original list"),
ev("ampd-2027", "19th Annual Meeting Photonic Devices 2027", "AMPD 2027", "Conference", "photonic integrated circuits;nanophotonics", "2027-04-07", "2027-04-09", "Berlin, Germany", "TBA", "TBA", "TBA", "TBA", "https://www.zib.de/workshop-photonic-devices/", "4", "German photonic devices annual meeting - directly relevant to group's fabrication work", "upcoming", "2026-08-09", "new entry; year roll-over from 2026; LOW-CONFIDENCE (search-snippet only, official page could not be fetched this run - please spot-check)",
   "https://www.zib.de/workshop-photonic-devices/"),
ev("aps-summit-2027", "APS Global Physics Summit 2027", "APS Summit 2027", "Conference", "quantum optics;quantum computing hardware;integrated quantum photonics", "2027-04-11", "2027-04-16", "Atlanta, GA, USA", "TBA", "TBA", "TBA", "TBA", "https://www.aps.org/events/2027/summit", "4", "Joint APS March+April meeting; major quantum photonics and computing sessions", "upcoming", "2026-05-13", "new entry; Apr 11-16 Atlanta",
   "https://www.aps.org/events/2027/summit"),
ev("bqit-2027", "Bristol Quantum Information Technologies Workshop 2027", "BQIT:27", "Workshop", "integrated quantum photonics;quantum communication;quantum computing hardware", "2027-04-26", "2027-04-28", "Bristol, UK", "TBA", "n/a", "TBA", "n/a", "https://www.bristol.ac.uk/qet-labs/events/bqit-workshop/", "5", "UK quantum photonics workshop - PICs, single photons, quantum networking", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition; venue We The Curious, Bristol",
   "https://quantiki.org/conference/bristol-quantum-information-technologies-workshop-2027"),
ev("cleo-2027", "Conference on Lasers and Electro-Optics 2027", "CLEO 2027", "Conference", "integrated quantum photonics;quantum optics;nonlinear optics;silicon photonics;single-photon sources", "2027-05-02", "2027-05-07", "Long Beach, CA, USA", "TBA", "n/a", "TBA", "TBA", "https://cleoconference.org/", "5", "Premier US laser/photonics conference; major IQP, Si-PIC and quantum optics sessions", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition; technical conference May 2-6/7, exhibition May 4-5",
   "https://cleoconference.org/about-us/"),
ev("damop-2027", "58th APS Division of Atomic, Molecular and Optical Physics Meeting", "DAMOP 2027", "Conference", "quantum optics;nonlinear optics;single-photon sources", "2027-05-13", "2027-05-18", "Chicago, IL, USA", "TBA", "TBA", "TBA", "TBA", "https://www.aps.org/events/2027/damop-2027", "3", "AMO physics flagship with quantum optics and single-photon sessions", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition",
   "https://www.aps.org/events/2027/damop-2027"),
ev("cleo-europe-2027", "CLEO/Europe-EQEC 2027", "CLEO/Europe 2027", "Conference", "integrated quantum photonics;quantum optics;nonlinear optics;nanophotonics", "2027-06-20", "2027-06-25", "Munich, Germany", "TBA", "TBA", "TBA", "TBA", "https://cleoeurope.org/", "5", "Flagship European photonics and quantum optics biennial conference", "upcoming", "2026-05-13", "corrected end date to Jun 25; part of World of Photonics Congress 2027",
   "https://cleoeurope.org/; https://www.photonics-congress.com/en/about/conferences/cleo-eqec/"),
ev("benasque-qi-2027", "Quantum Information 2027", "Benasque QI 2027", "Workshop", "quantum communication;quantum optics", "2027-06-20", "2027-07-09", "Benasque, Spain", "TBA", "n/a", "TBA", "n/a", "https://www.benasque.org/2027qi/", "3", "Quantum information including QKD and networking", "upcoming", "2026-08-09", "new entry; year roll-over from 2025 edition",
   "https://www.benasque.org/2027qi/"),
ev("ieee-rapid-2027", "IEEE Research and Applications of Photonics in Defense 2027", "IEEE RAPID 2027", "Conference", "single-photon sources;quantum communication", "2027-08-11", "2027-08-13", "Miramar Beach, FL, USA", "TBA", "TBA", "TBA", "TBA", "https://ieee-rapid.org/", "2", "Defense photonics conference", "upcoming", "2026-05-13", "no change",
   "https://ieee-rapid.org/"),
ev("ecoc-2027", "European Conference on Optical Communication 2027", "ECOC 2027", "Conference", "photonic integrated circuits;silicon photonics;quantum communication;integrated quantum photonics", "2027-10-10", "2027-10-14", "Milan, Italy", "TBA", "TBA", "TBA", "TBA", "https://ecoc2027.org/", "4", "Europe's flagship optical comms conference with quantum technologies track", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition (Málaga); exhibition Oct 11-13",
   "https://ecoc2027.org/; https://www.ecocexhibition.com/future-dates/"),
ev("ipc-2027", "IEEE Photonics Conference 2027", "IPC 2027", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics", "2027-11-10", "2027-11-14", "Eindhoven, Netherlands", "TBA", "TBA", "TBA", "TBA", "https://ieee-ipc.org/", "4", "IEEE photonics flagship; European edition in Eindhoven — low travel cost", "upcoming", "2026-05-13", "no change",
   "https://ieee-ipc.org/"),
ev("spie-pe-2028", "SPIE Photonics Europe 2028", "SPIE PE 2028", "Conference", "integrated quantum photonics;silicon photonics;nanophotonics", "2028-04-02", "2028-04-06", "Strasbourg, France", "TBA", "TBA", "TBA", "TBA", "https://spie.org/conferences-and-exhibitions/photonics-europe", "5", "Major European SPIE photonics conference with quantum integrated sessions; confirmed biennial (even years)", "upcoming", "2026-08-09", "new entry; biennial roll-over from 2026 edition (confirmed 2028 dates already published)",
   "https://spie.org/conferences-and-exhibitions/photonics-europe"),
ev("photonics-switching-2026", "Photonics in Switching and Computing 2026", "PiS 2026", "Conference", "photonic integrated circuits;silicon photonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Photonic switching and routing conference", "upcoming", "2026-05-13", "no change; dates TBA (expected summer 2026)",
   "original list"),
ev("opic", "Optics & Photonics International Congress", "OPIC", "Conference", "integrated quantum photonics;quantum optics;nanophotonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Japanese international photonics congress with quantum tracks", "upcoming", "2026-05-13", "placeholder; next edition dates TBA",
   "original list"),
ev("owtnm", "Optical Wave & Waveguide Theory and Modelling Workshop", "OWTNM", "Workshop", "photonic integrated circuits;nanophotonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "TBA", "3", "Waveguide modelling workshop relevant to PIC simulation", "upcoming", "2026-05-13", "placeholder; next edition dates TBA",
   "original list"),
ev("opon-2025", "OPON 2025 Münster", "OPON 2025", "Conference", "integrated quantum photonics;nanophotonics", "TBA", "TBA", "Münster, Germany", "TBA", "TBA", "TBA", "TBA", "TBA", "4", "Local Münster photonics conference — direct access, no travel", "upcoming", "2026-05-13", "no change; dates TBA",
   "original list"),
ev("nrw-nano-2026", "NRW NanoConference 2026", "NRW Nano 2026", "Conference", "nanophotonics;photonic integrated circuits", "TBA", "TBA", "North Rhine-Westphalia, Germany", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Regional NRW nanotechnology conference; local access", "upcoming", "2026-08-09", "corrected: most recent (11th) edition was confirmed Sep 30-Oct 1 2025, Dortmund (see nrw-nano-2025, archived); next edition dates not yet announced",
   "original list"),
ev("icqe", "International Conference on Quantum Energy", "ICQE", "Conference", "quantum computing hardware", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "2", "Quantum energy applications; adjacent to quantum computing hardware", "upcoming", "2026-05-13", "placeholder; next edition TBA",
   "original list"),
ev("optica-nlo-2027", "Optica Nonlinear Optics Topical Meeting 2027", "NLO 2027", "Conference", "nonlinear optics;quantum optics;thin-film lithium niobate", "TBA", "TBA", "Hawaii, USA", "TBA", "TBA", "TBA", "TBA", "https://www.optica.org/events/topical_meetings/nonlinear_optics/", "3", "Standalone biennial Optica meeting on nonlinear/parametric optics relevant to squeezed-light/photon-pair generation", "upcoming", "2026-08-09", "new entry; biennial (2025 Honolulu, no 2026 edition); 2027 dates not yet announced",
   "https://www.optica.org/events/topical_meetings/nonlinear_optics/"),
ev("wqed-next", "Workshop on Waveguide QED (next edition)", "WQED", "Workshop", "integrated quantum photonics;single-photon sources;quantum optics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "TBA", "4", "Waveguide QED workshop - core to integrated quantum photonics and single-photon-source physics", "upcoming", "2026-08-09", "new entry; placeholder for next edition after WQED 2025 (Erice); not yet announced, expect ~2027",
   "https://quantum.unipa.it/events/wqed2025/"),
ev("erice-qt-school", "Quantum Technology School Erice (next edition)", "Erice QT School", "Summer School", "integrated quantum photonics;single-photon sources", "TBA", "TBA", "Erice, Sicily, Italy", "TBA", "n/a", "TBA", "n/a", "https://sites.google.com/view/quantumtechnologyschoolerice/home", "4", "Ettore Majorana Foundation PhD school on photonic qubits and quantum light sources (organizers Portalupi/Trotta/Bajoni)", "upcoming", "2026-08-09", "new entry; last confirmed edition Jun 2024; next edition dates not yet announced",
   "https://sites.google.com/view/quantumtechnologyschoolerice/home"),
ev("opon-next", "OPON - Optical Properties of Nanostructures (next edition)", "OPON", "Workshop", "nanophotonics;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "TBA", "4", "Invitation-only workshop rotating among Wrocław/Münster/Warsaw/Bayreuth; OPON2025 was hosted in Münster (direct access)", "upcoming", "2026-08-09", "new entry; placeholder for 9th edition after OPON2025 (Münster, Apr 2025); host/dates not yet announced",
   "https://www.uni-muenster.de/Physik.FT/en/opon2025/index.html"),
ev("ieee-qce27", "IEEE Quantum Week 2027", "QCE27", "Conference", "quantum computing hardware;quantum communication;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://qce.quantum.ieee.org/", "3", "IEEE quantum week with photonic quantum computing and QKD sessions", "upcoming", "2026-08-09", "new entry; year roll-over from 2026 edition (Toronto); location not yet announced (TBD)",
   "https://qce.quantum.ieee.org/2026/key-deadlines/"),
ev("photonics-north-2027", "Photonics North (next edition)", "PN 2026", "Conference", "silicon photonics;integrated quantum photonics;quantum communication", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Canadian national photonics conference with PIC and quantum sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after photonics-north-2026; not yet announced",
   "TBA"),
ev("cargese-qt-next", "Summer School on Quantum Technologies for Computation and Communication (next edition)", "QT4CC 2026", "Summer School", "quantum communication;quantum computing hardware;quantum networking", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://qt4cc.sciencesconf.org/", "4", "PhD summer school on quantum computation and communication — ideal for group's PhD students", "upcoming", "2026-08-09", "new entry; placeholder for next edition after cargese-qt-2026; not yet announced",
   "https://qt4cc.sciencesconf.org/"),
ev("spie-pfq-next", "SPIE Photonics for Quantum (next edition)", "PfQ 2026", "Conference", "integrated quantum photonics;single-photon sources;quantum communication;quantum computing hardware", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "TBA", "https://spie.org/conferences-and-exhibitions/photonics-for-quantum", "5", "Dedicated quantum photonics conference — bullseye for AG Schuck", "upcoming", "2026-08-09", "new entry; placeholder for next edition after spie-pfq-2026; not yet announced",
   "https://spie.org/conferences-and-exhibitions/photonics-for-quantum"),
ev("islc-2028", "30th International Semiconductor Laser Conference (next edition)", "ISLC 2026", "Conference", "single-photon sources;III-V on Si;photonic integrated circuits", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://events.tuni.fi/islc2026/", "4", "III-V semiconductor laser conference — key for quantum light source development", "upcoming", "2026-08-09", "new entry; placeholder for next edition after islc-2026; not yet announced",
   "https://events.tuni.fi/islc2026/"),
ev("grc-quantum-science-2028", "Gordon Research Conference — Quantum Science (next edition)", "GRC QS 2026", "Conference", "quantum optics;quantum computing hardware;quantum communication;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://www.grc.org/quantum-science-conference/2026/", "4", "GRC on quantum science covering photonic quantum systems and quantum information", "upcoming", "2026-08-09", "new entry; placeholder for next edition after grc-quantum-science-2026; not yet announced",
   "https://www.grc.org/quantum-science-conference/2026/"),
ev("icap-next", "29th International Conference on Atomic Physics (next edition)", "ICAP 2026", "Conference", "quantum optics;nonlinear optics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://www.icap29.com/home.html", "3", "Atomic physics conference with quantum optics and light-matter interaction sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after icap-2026; not yet announced",
   "https://www.icap29.com/home.html"),
ev("ecio-2027", "European Conference on Integrated Optics (next edition)", "ECIO 2026", "Conference", "integrated quantum photonics;photonic integrated circuits;silicon photonics;nanophotonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://ecio-conference.org/", "5", "Largest European integrated optics conference — core relevance for AG Schuck", "upcoming", "2026-08-09", "new entry; placeholder for next edition after ecio-2026; not yet announced",
   "https://ecio-conference.org/"),
ev("optica-quantum-2027", "Optica Quantum 2.0 Conference and Exhibition (next edition)", "Optica Quantum 2026", "Conference", "integrated quantum photonics;quantum optics;single-photon sources;quantum communication;quantum computing hardware", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://www.optica.org/events/topical_meetings/quantum/", "5", "Premier quantum photonics conference spanning PICs, detectors, QKD, and hardware", "upcoming", "2026-08-09", "new entry; placeholder for next edition after optica-quantum-2026; not yet announced",
   "https://www.optica.org/events/topical_meetings/quantum/"),
ev("icop-2028", "Italian Conference on Optics and Photonics (next edition)", "ICOP 2026", "Conference", "integrated quantum photonics;quantum optics;nanophotonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://www.icop2026.it/", "3", "Italian national photonics conference with quantum tracks", "upcoming", "2026-08-09", "new entry; placeholder for next edition after icop-2026; not yet announced",
   "https://www.icop2026.it/"),
ev("egas-2027", "57th Conference of the European Group on Atomic Systems (next edition)", "EGAS 2026", "Conference", "quantum optics;nonlinear optics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://egas57.org/", "3", "European atomic systems conference with quantum optics and light-matter sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after egas-2026; not yet announced",
   "https://egas57.org/"),
ev("benasque-qsi-next", "Quantum Science: Implementations (next edition)", "Benasque QSI", "Workshop", "integrated quantum photonics;quantum computing hardware;quantum communication", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://benasque.org/", "4", "Benasque workshop on quantum implementations including photonic platforms", "upcoming", "2026-08-09", "new entry; placeholder for next edition after benasque-qsi-2026; not yet announced",
   "https://benasque.org/"),
ev("les-houches-pa-next", "Photons & Atoms (next edition)", "Les Houches PA 2026", "Summer School", "quantum optics;nonlinear optics;integrated quantum photonics;single-photon sources", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://photon-atoms-26.leshouches.science/", "5", "Leading European doctoral school on photons, quantum gases and quantum technologies", "upcoming", "2026-08-09", "new entry; placeholder for next edition after les-houches-pa-2026; not yet announced",
   "https://photon-atoms-26.leshouches.science/"),
ev("iqt-nordics-2027", "IQT Nordics (next edition)", "IQT Nordics 2026", "Conference", "quantum communication;quantum computing hardware;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "TBA", "https://iqtevent.com/nordics/", "3", "Nordic quantum industry conference with hardware and QKD sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after iqt-nordics-2026; not yet announced",
   "https://iqtevent.com/nordics/"),
ev("oecc-2027", "31st OptoElectronics and Communications Conference (next edition)", "OECC 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "3", "Asia-Pacific optoelectronics and comms conference with PIC sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after oecc-2026; not yet announced",
   "TBA"),
ev("icton-2027", "26th International Conference on Transparent Optical Networks (next edition)", "ICTON 2026", "Conference", "photonic integrated circuits;silicon photonics;quantum communication", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://www.gov.pl/web/instytut-lacznosci/icton-2026", "3", "Transparent optical networks with PIC and quantum photonics tracks", "upcoming", "2026-08-09", "new entry; placeholder for next edition after icton-2026; not yet announced",
   "https://www.gov.pl/web/instytut-lacznosci/icton-2026"),
ev("ieee-sum-2027", "IEEE Summer Topicals Meeting Series (next edition)", "IEEE SUM 2026", "Conference", "photonic integrated circuits;silicon photonics;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://ieee-sum.org/", "3", "IEEE focused topical photonics sessions", "upcoming", "2026-08-09", "new entry; placeholder for next edition after ieee-sum-2026; not yet announced",
   "https://ieee-sum.org/"),
ev("grc-plasmonics-nano-2028", "Gordon Research Conference — Plasmonics and Nanophotonics (next edition)", "GRC PNP 2026", "Conference", "nanophotonics;plasmonics;quantum optics;2D-material photonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/", "5", "GRC on quantum nanophotonics, plasmonics and 2D photonics — highly relevant to AG Schuck", "upcoming", "2026-08-09", "new entry; placeholder for next edition after grc-plasmonics-nano-2026; not yet announced",
   "https://www.grc.org/plasmonics-and-nanophotonics-conference/2026/"),
ev("cewqo-2028", "30th Central European Workshop on Quantum Optics (next edition)", "CEWQO30 2026", "Workshop", "quantum optics;single-photon sources;nonlinear optics;integrated quantum photonics", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/", "5", "European quantum optics workshop co-hosted by FAU/MPL Erlangen — close to Münster", "upcoming", "2026-08-09", "new entry; placeholder for next edition after cewqo-2026; not yet announced",
   "https://www.lightmatter.fau.de/2026/04/cewqo30-the-30th-central-european-workshop-on-quantum-optics-2026/"),
ev("apc-2027", "Optica Advanced Photonics Congress (next edition)", "APC 2026", "Conference", "integrated quantum photonics;nonlinear optics;nanophotonics", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "TBA", "https://www.optica.org/events/congress/advanced_photonics_congress/", "5", "Optica Advanced Photonics Congress incl. Integrated Photonics Research (IPR) quantum-photonics subtrack, Nonlinear Photonics, and BGPP sub-conferences", "upcoming", "2026-08-09", "new entry; placeholder for next edition after apc-2026; not yet announced",
   "https://www.optica.org/events/congress/advanced_photonics_congress/"),
ev("grc-mech-quant-2028", "Gordon Research Conference — Mechanical Systems in the Quantum Regime (next edition)", "GRC MSQR 2026", "Conference", "optomechanics;quantum computing hardware", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/", "4", "Optomechanics GRC — relevant to cavity-optomechanical quantum transducers", "upcoming", "2026-08-09", "new entry; placeholder for next edition after grc-mech-quant-2026; not yet announced",
   "https://www.grc.org/mechanical-systems-in-the-quantum-regime-conference/2026/"),
ev("lake-como-ufqp-next", "New Frontiers in Ultrafast Quantum Optics — Lake Como School (next edition)", "UFQP 2026", "Summer School", "quantum optics;nonlinear optics;single-photon sources", "TBA", "TBA", "TBA", "TBA", "n/a", "TBA", "n/a", "https://ufqp.lakecomoschool.org/", "4", "Summer school on ultrafast quantum optics — photon correlations and time-frequency entanglement", "upcoming", "2026-08-09", "new entry; placeholder for next edition after lake-como-ufqp-2026; not yet announced",
   "https://ufqp.lakecomoschool.org/"),
ev("spw-2027", "13th Single Photon Workshop (tentative)", "SPW 2027", "Workshop", "single-photon sources;SNSPDs;quantum optics;quantum communication", "TBA", "TBA", "Singapore", "TBA", "n/a", "TBA", "TBA", "TBA", "5", "THE dedicated single-photon workshop - central to AG Schuck's detector and source work", "upcoming", "2026-08-09", "new entry; TENTATIVE ONLY - aggregator sources report Nov 15-19 2027 Singapore, but NO official confirmation found; verify with organizers before relying on this",
   "https://www.quantum.info/conf/2027.html (aggregator, unofficial)"),
]

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

    for e in events:
        row_vals = [e[c] for c in COLS]
        ws.append(row_vals)
        row_num = ws.max_row

        # light-grey background for archived rows
        is_archived = e["status"] == "archived"
        if is_archived:
            for cell in ws[row_num]:
                cell.fill = ARCHIVED_FILL

        # relevance score colour on that cell
        score_cell = ws.cell(row=row_num, column=col_idx["relevance_score"])
        sc = e.get("relevance_score")
        if sc and str(sc).isdigit() and int(sc) in SCORE_COLORS:
            score_cell.fill = PatternFill(fill_type="solid",
                                          fgColor=SCORE_COLORS[int(sc)])

        # deadline columns — conditional colour
        for dcol in DEADLINE_COLS:
            val = e.get(dcol, "")
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
