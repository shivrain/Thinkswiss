#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique EGIZC (2016–2025)."""

import os
from datetime import date

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "/workspace/output/mozambique"
os.makedirs(OUTPUT_DIR, exist_ok=True)

POLICY_FIELDS = {
    "policy_name": (
        "Estratégia de Gestão Integrada das Zonas Costeiras (2016–2025) — EGIZC"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz213368.pdf",
    "policy_year": 2016,
    "policy_objective": (
        "Establish the Integrated Coastal Zone Management Strategy (EGIZC) 2016–2025 "
        "to guide sustainable development of Mozambique's ~2,700 km coastline and coastal "
        "districts (~42% of national territory, ~572,000 km² ecosystems), balancing "
        "biodiversity conservation, socioeconomic development and marine/coastal pollution "
        "prevention through four thematic areas — natural preservation, sustainable "
        "planning, marine/coastal pollution prevention, and legal/institutional coordination."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Action-plan targets include: at least 10 coastal sanitary landfills (2016–2025); "
        "12 wastewater treatment plants (ETARs) in coastal agglomerations; port waste-reception "
        "equipment at main ports; at least 10 public beaches classified (blue-flag criteria); "
        "250 technicians trained in marine/coastal pollution response; 100 pollution "
        "fiscalization specialists; implementation of National Oil Spill Contingency Plan (PNC); "
        "02 legal instruments on ballast, dredging and waste disposal at sea."
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "National integrated coastal zone management strategy approved by the Council of "
        "Ministers (3rd ordinary session, 9 February 2016, MITADER) — strategic planning "
        "instrument with annexed action plan (~MZN 16 billion indicative budget), not binding "
        "sectoral regulation. Operational rules in Decree 45/2006, Decree 97/2020, Sea Law, "
        "POLMAR, POEM, EDEA, etc."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging, ports"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "consumption, disposal, environmental leakage"
    ),
    "policy_budget": 1.0,
    "policy_budget_text": (
        "Indicative total action-plan budget MZN 16,081,500,000 distributed across four "
        "thematic areas; AT III (Marine and Coastal Pollution Prevention) subtotal "
        "MZN 15,152,000,000 including MZN 3 billion (10 sanitary landfills), MZN 10 billion "
        "(12 ETARs), MZN 2 billion (pollution-response equipment), MZN 20 million (port "
        "waste-reception equipment)."
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Thematic Area III — Marine and Coastal Pollution Prevention: Strategic framework "
            "to reduce, eliminate and mitigate hydrocarbon pollution and land-based sources "
            "(LBSA/Nairobi Protocol); addresses solid waste deposition, direct pollutant "
            "discharges, ship-generated waste (ballast water, grey water, garbage, tank "
            "cleaning residues) and industrial/domestic effluents affecting coastal ecosystems."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'AT III objective: "redução, eliminação, mitigação da poluição por hidrocarbonetos '
            'e por fontes baseadas em terras"; diagnosis cites "depósito de resíduos sólidos", '
            '"lixo", ship discharges and RSU management gaps along the 2,700 km coast.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary pollution-prevention instrument — no explicit 'plástico' but solid waste "
            "and litter on coast/sea are core problems. Not coded separately: hydrocarbon-only "
            "PNC/ESI map actions bundled under monitoring/contingency instrument."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Action Plan III.3.1.3 / RSU context: Construct at least 10 sanitary landfills in "
            "main coastal cities and towns (Pemba, Nacala, Quelimane, Beira, Inhambane, Xai-Xai, "
            "Bilene, Maputo/Matola, Ponta d'Ouro); convert open dumps to sanitary landfills; "
            "waste management definition requires reduction, recycling and reuse of wastes."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'III.3.1.3: "Construir aterros sanitários nas principais cidades e vilas costeiras" '
            '— at least 10 by 2025 (MAEFP, MOPHRH, MITADER; MZN 3 billion). Context: only '
            'Mavoco industrial landfill and Songo controlled dump existed; open dumps predominate.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary solid-waste instrument — plastics included implicitly in RSU. Glossary "
            "defines 'Gestão de Resíduos' with reduction, reciclagem and re-utilização. "
            "References Integrated Urban Solid Waste Management Strategy (MICOA 2012)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Action Plan III.1.2.2: Acquire equipment and materials for waste reception at main "
            "ports (Maputo, Beira, Nacala); strengthen port capacity to receive ship-generated "
            "solid waste, wastewater, tank-cleaning residues and garbage per IMO/MARPOL framework."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'III.1.2.2: "Adquirir equipamentos e material de recepção de resíduos nos principais '
            'portos" — ports capacitated by 2018 (MTC, MITADER, MIMAIP; MZN 20 million). '
            'Context: maritime transport generates "cinza, lixo, outros resíduos sólidos".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Port waste-reception instrument — vessel waste including garbage. Links to Decree "
            "45/2006 and MARPOL Annex V. Not coded separately: pollution-response equipment "
            "(III.1.2.1, MZN 2 billion)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan III.3.1.1 / AT II.3: Construct 12 domestic and industrial wastewater "
            "treatment plants (ETARs) in coastal agglomerations; reduce land-based pollutant "
            "discharges to rivers, estuaries and sea; objective to reduce pollution of surface, "
            "ground and marine waters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'III.3.1.1: 12 ETARs in Ponta d\'Ouro, Matola, Maputo, Xai-Xai, Inhambane, '
            'Inhassoro, Quelimane, Nacala, Moma, Pemba, Palma (2016–2025; MZN 10 billion). '
            'AT II objective 2: "Reduzir a poluição das águas superficiais, subterrâneas e marinhas".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Land-based source reduction — most coastal centres lacked ETARs (except Maputo/Beira). "
            "Reduces nutrient/industrial/domestic pollution reaching sea. Not coded: industrial "
            "effluent standards mobilization (III.3.1.4–5)."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan III.3.2: Establish legal mechanisms for control and management of "
            "ballast water, dredging and waste disposal at sea; ratify/implement IMO ballast-water "
            "guidelines and London Protocol on dumping; reduce sea-based waste dumping and "
            "dredged-material disposal impacts on beaches and ecosystems."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'III.3.2.1: "Elaborar o quadro legal para o controlo e gestão de águas de lastro e '
            'dragagens... deposição de resíduos no mar" — at least 02 legal instruments by 2025. '
            'Text references IMO London Protocol on Garbage Disposal at Sea and MARPOL annexes.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Sea-based waste/disposal control — addresses clandestine hazardous waste dumping. "
            "Glossary defines ship 'Resíduos' as operational waste liable to discharge at sea. "
            "Not coded separately: invasive species control (III.2.1.1)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan II.2.5 / pollution monitoring: Establish public beach classification "
            "criteria and blue-flag attribution for at least 10 beaches; estuary/bay pollution "
            "monitoring programmes; strengthen pollution fiscalization, monitoring and assessment "
            "capacity (100+ technicians; pollution hotspots mapping)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'II.2.5.1: Blue-flag beach classification criteria for at least 10 public beaches '
            '(Ponta d\'Ouro, Maputo, Xai-Xai, Tofo, Vilanculos, Závora, Beira, Nacala, Mossuril, '
            'Pemba). III.1.1: train at least 100 pollution fiscalization specialists (2016–2025).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Beach quality/monitoring instrument — implicit litter and coastal cleanliness via "
            "blue-flag standards. Bundled with pollution supervision system (III.1.1) and water "
            "quality hotspot mapping (III.1.5). Strategy period 2016–2025 completed."
        ),
    },
]

COLUMNS = [
    "policy_name",
    "policy_url",
    "policy_year",
    "policy_objective",
    "policy_target",
    "policy_target_text",
    "policy_type",
    "policy_type_justification",
    "policy_integration",
    "policy_sectors_list",
    "policy_circularity",
    "policy_lifecycle_phases_list",
    "policy_budget",
    "policy_budget_text",
    "policy_score",
    "instrument_type",
    "instrument_lifecycle_stage",
    "instrument_description",
    "instrument_in_force",
    "instrument_implementation",
    "instrument_implementation_text",
    "instrument_score",
    "comments",
]

COLUMN_LABELS = {
    "policy_name": "A — policy_name",
    "policy_url": "B — policy_url",
    "policy_year": "C — policy_year",
    "policy_objective": "D — policy_objective",
    "policy_target": "E — policy_target",
    "policy_target_text": "F — policy_target_text",
    "policy_type": "G — policy_type",
    "policy_type_justification": "H — policy_type_justification",
    "policy_integration": "I — policy_integration",
    "policy_sectors_list": "J — policy_sectors_list",
    "policy_circularity": "K — policy_circularity",
    "policy_lifecycle_phases_list": "L — policy_lifecycle_phases_list",
    "policy_budget": "M — policy_budget",
    "policy_budget_text": "N — policy_budget_text",
    "policy_score": "O — policy_score [auto]",
    "instrument_type": "P — instrument_type",
    "instrument_lifecycle_stage": "Q — instrument_lifecycle_stage",
    "instrument_description": "R — instrument_description",
    "instrument_in_force": "S — instrument_in_force",
    "instrument_implementation": "T — instrument_implementation",
    "instrument_implementation_text": "U — instrument_implementation_text",
    "instrument_score": "V — instrument_score [auto]",
    "comments": "W — comments",
}


def build_dataframe():
    return pd.DataFrame([{**POLICY_FIELDS, **i} for i in INSTRUMENTS], columns=COLUMNS)


def export_csv(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.csv")
    df.rename(columns=COLUMN_LABELS).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def export_xlsx(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.xlsx")
    labelled = df.rename(columns=COLUMN_LABELS)
    labelled.to_excel(path, index=False, sheet_name="4P Index Coding")
    wb = load_workbook(path)
    ws = wb.active
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = min(
            45, max(12, len(str(ws.cell(1, col).value or "")) * 0.9)
        )
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_para(doc, text):
    doc.add_paragraph(text)


def create_translation_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading(
        "Integrated Coastal Zone Management Strategy — EGIZC", 0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("2016–2025 — Republic of Mozambique")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Estratégia de Gestão Integrada das Zonas Costeiras "
        "(EGIZC) 2016–2025\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz213368.pdf")

    sections = [
        ("Approval and Enacting Status", (
            "The Integrated Coastal Zone Management Strategy (EGIZC) was approved by the "
            "3rd Ordinary Session of the Council of Ministers on 9 February 2016.\n\n"
            "Prepared by the Ministry of Land, Environment and Rural Development (MITADER), "
            "Maputo, February 2016. The strategy covers the period 2016–2025 and includes an "
            "annexed Action Plan with an indicative budget of approximately MZN 16,081,500,000."
        )),
        ("Executive Summary", (
            "Mozambique has Africa's third-longest coastline (more than 2,700 km), with coastal "
            "ecosystems covering approximately 572,000 km² (about 42% of the country). Two-thirds "
            "of the population live in coastal districts. Economic activities — ports, maritime "
            "transport, fisheries, aquaculture, tourism, industry, agriculture and mining — depend "
            "heavily on coastal resources.\n\n"
            "Conflicting uses (hydrocarbon exploration, port development, tourism, fisheries, "
            "urban expansion, conservation) cause both positive development and negative impacts: "
            "disordered occupation, unsustainable resource use, pollution from hydrocarbons and "
            "untreated industrial/domestic effluents, unsustainable agriculture, loss of natural "
            "and cultural heritage, and climate-change-exacerbated coastal erosion and salinity intrusion.\n\n"
            "Vision: \"A Mozambican coastal zone rich in biodiversity, free from environmental "
            "degradation, and resilient to climate change, with integrated and sustainable "
            "socioeconomic development.\"\n\n"
            "Mission: \"Promote integrated coastal zone management through articulated and "
            "coordinated implementation of policies and instruments ensuring rational use of "
            "physical space, preservation of natural resources, reduction of community "
            "vulnerability, sustainable development and coastal resilience.\"\n\n"
            "General objective: Ensure balance of marine and coastal ecosystems through "
            "sustainable development, contributing to continuous improvement in quality of life "
            "for coastal communities."
        )),
        ("1. Context and Rationale", [
            ("1.1 Coastal importance", (
                "Mozambique's coastline includes coral reefs, beaches, sand dunes, seagrass beds, "
                "lagoons and estuarine systems. Coastal zones are critical for biodiversity "
                "reproduction and coexistence of terrestrial and marine species. Maritime transport "
                "generates pollutants including ballast water, oily water, tank-cleaning residues, "
                "wastewater, ash, garbage and other solid waste and atmospheric emissions from ships."
            )),
            ("1.2 Need for integrated management", (
                "Growing economic development — coal, hydrocarbons, regional port traffic — "
                "increases pressure on transport and port networks. Isolated sectoral implementation "
                "causes overlapping activities and impacts. Absence of a dedicated marine/coastal "
                "management institution enables environmentally unsustainable actions.\n\n"
                "EGIZC emerges as a ten-year guiding instrument consolidating objectives and "
                "strategic options for integrated coastal management, coordinating public/private "
                "entities, scientific community, civil society and local communities."
            )),
            ("1.3 Legal and institutional framework", (
                "Mozambique is party to UNCLOS, MARPOL, IMO conventions, Nairobi Convention "
                "(LBSA Protocol — Resolution 3/2014), UNFCCC, CBD and African conservation "
                "conventions. National instruments include Environmental Framework Law (1997), "
                "Marine Pollution Prevention Regulation (Decree 45/2006), National Oil Spill "
                "Contingency Plan, Integrated Urban Solid Waste Management Strategy (MICOA 2012), "
                "and institutions including CONDES, CDS-ZC (Coastal Sustainable Development Centre) "
                "and CEPAM (Marine and Coastal Environment Research Centre)."
            )),
        ]),
        ("2. Four Thematic Areas and Strategic Objectives", [
            ("2.1 Thematic Area I — Natural Preservation", (
                "Focus: physical-biological components and living natural resources.\n\n"
                "Strategic objectives:\n"
                "1. Ensure management and preservation of marine and coastal resources through "
                "science-based decisions;\n"
                "2. Protect threatened/endangered species for climate resilience and ecological balance;\n"
                "3. Strengthen ecosystem resilience to climate change impacts.\n\n"
                "Actions include: updating ecosystem use/conservation information; mapping sensitive "
                "areas and critical biodiversity; monitoring threatened species; revising legal "
                "framework for climate vulnerability; promoting marine and coastal protected areas."
            )),
            ("2.2 Thematic Area II — Planning and Sustainable Development", (
                "Focus: human occupation and coastal economic activities.\n\n"
                "Strategic objectives:\n"
                "1. Ensure orderly occupation of spaces;\n"
                "2. Reduce pollution of surface, groundwater and marine waters.\n\n"
                "Actions include: coastal spatial planning instruments; protection of sensitive "
                "coastal areas through light/heavy engineering; coastal requalification and erosion "
                "control; determination of public beach classification standards; implementation of "
                "clean technologies reducing pollutant effects in industry, aquaculture and "
                "artisanal mining."
            )),
            ("2.3 Thematic Area III — Marine and Coastal Pollution Prevention", (
                "Focus: reduction, elimination and mitigation of hydrocarbon pollution and "
                "land-based sources.\n\n"
                "Strategic objectives:\n"
                "1. Strengthen operational capacity for fiscalization, monitoring and assessment of "
                "marine/coastal pollution risks and levels;\n"
                "2. Prevent impacts of land-based and sea-based human activities;\n"
                "3. Implement appropriate clean mechanisms and technologies reducing pollutant "
                "effects in industrial processes.\n\n"
                "Actions include: pollution supervision/control/monitoring systems; acquisition of "
                "pollution-response equipment and trained personnel; National Contingency Plan "
                "implementation; Environmental Sensitivity Index (ESI) maps; water quality "
                "standards and pollution hotspot identification; ballast water, dredging and "
                "sea-waste-disposal legal controls; construction of ETARs and sanitary landfills."
            )),
            ("2.4 Thematic Area IV — Legal Framework and Inter-institutional Coordination", (
                "Focus: legal and institutional aspects of coastal management and cross-cutting "
                "capacity building, technology transfer and financing.\n\n"
                "Strategic objectives include ensuring integrated coastal development, harmonizing "
                "sectoral plans, strengthening institutional coordination, and elaborating guiding "
                "instruments including a National Coastal Management Policy (target: approved "
                "2016–2018)."
            )),
        ]),
        ("3. Marine and Coastal Pollution — Diagnostic (Section 3.2.4)", [
            ("3.1 Land-based pollution", (
                "Coastal pollution results from solid waste deposition and direct pollutant discharges "
                "to coast, sea, rivers, lakes and lagoons from industrial, agricultural and urban "
                "domestic activities and coastal communities.\n\n"
                "Land-based sources (LBSA/Nairobi Protocol) include: solid waste deposition; direct "
                "discharges from urban soils, industrial effluents, agrochemical leachate, refinery "
                "incidents and river transport. Groundwater pollution arises from fertilizers, "
                "pesticides, industrial waste, urban landfill leachate and intensive aquifer exploitation.\n\n"
                "Except Maputo and Beira, most coastal urban centres and tourist agglomerations "
                "(Matola, Xai-Xai, Inhambane, Quelimane, Tete, Nacala, Pemba; Ponta d'Ouro, Tofo, "
                "Inhassoro, Vilanculos, Wimbe, etc.) lack domestic and industrial wastewater treatment "
                "plants (ETARs/ETARIs). For municipal solid waste (RSU), only Mavoco industrial "
                "landfill and Songo controlled dump existed; open dumps were being converted to "
                "sanitary landfills."
            )),
            ("3.2 Sea-based pollution", (
                "Sea-based pollution includes hazardous waste dumping at sea, accidental oil spills "
                "from navigation (deck washing, tanker operations, ballast water, dredging, "
                "offshore platform damage). Clandestine hazardous waste dumping often occurs outside "
                "international conventions.\n\n"
                "Historical incident: Greek tanker Katina-P in Maputo Bay (April 1992). IMO "
                "guidelines address ballast water management and garbage disposal at sea (London "
                "Protocol). Challenges include ratifying ballast-water convention, approving national "
                "pollution-prevention legislation, and operationalizing MARPOL annexes, OPRC/HNS "
                "and Nairobi Convention emergency protocol."
            )),
            ("3.3 Dredging and beach sediment", (
                "Dredging (capital, maintenance, environmental clean-up) can remobilize contaminated "
                "sediments (heavy metals, PAHs, PCBs, pesticides, TBT). Navigation channels act as "
                "sediment sinks, reducing beach nourishment downstream and exacerbating coastal erosion. "
                "EGIZC promotes alternative uses for dredged material including sanitary landfill cover."
            )),
        ]),
        ("4. Key Definitions — Waste and Pollution", (
            "Gestão de Resíduos (Waste Management): All viable procedures to ensure environmentally "
            "safe, sustainable and rational waste management, including reduction, recycling and "
            "reuse, covering separation, collection, handling, transport, storage and/or disposal "
            "and subsequent protection of disposal sites to protect human health and the environment.\n\n"
            "Resíduos (Waste/Ship Garbage): Sanitary sewage and all kinds of domestic and operational "
            "food waste (excluding fresh fish), generated during normal ship operation and liable to "
            "continuous or periodic discharge to sea, lake or navigable rivers.\n\n"
            "Poluição marinha (Marine pollution): Any spill in marine, lacustrine or fluvial "
            "environment caused by any maritime activity or ship circulation.\n\n"
            "Zonas costeiras (Coastal zones): Areas from the interior/terrestrial limit of all "
            "coastal districts (including Lake Niassa and Cahora Bassa shore districts) to 12 "
            "nautical miles seaward."
        )),
        ("5. Action Plan — Plastic-Relevant Elements", [
            ("5.1 Solid waste and sanitary landfills (AT III)", (
                "III.3.1.3 — Construct at least 10 sanitary landfills in Pemba, Nacala Porto, "
                "Quelimane, Beira, Inhambane, Xai-Xai, Praia de Bilene, Maputo/Matola and Ponta "
                "d'Ouro (2016–2025; MZN 3,000,000,000; MAEFP, MOPHRH, MITADER).\n\n"
                "III.3.1.4 — Improve industrial waste/effluent treatment: 1 industrial ETARI and "
                "1 solid waste plan for large industries; 10% of new industries with clean "
                "production technologies.\n\n"
                "III.3.1.5 — Mobilize 20% of new industries to adopt clean standards for waste "
                "and effluent management."
            )),
            ("5.2 Wastewater treatment (AT III)", (
                "III.3.1.1 — Construct 12 ETARs in coastal agglomerations: Ponta d'Ouro, Matola, "
                "Maputo, Xai-Xai, Inhambane, Inhassoro, Quelimane, Nacala Porto, Moma, Pemba, Palma "
                "(2016–2025; MZN 10,000,000,000).\n\n"
                "III.3.1.2 — Approve directive on ETAR classification and construction standards."
            )),
            ("5.3 Ports, pollution response and sea disposal (AT III)", (
                "III.1.2.2 — Acquire port waste-reception equipment at main ports (2016–2018).\n\n"
                "III.1.2.1 — Equip 3 regional pollution-response centres (MZN 2,000,000,000).\n\n"
                "III.1.2.3 — Train 250 technicians in marine/coastal pollution combat (2016–2025).\n\n"
                "III.3.2.1 — Approve at least 2 legal instruments on ballast water, dredging and "
                "waste disposal at sea (2016–2025).\n\n"
                "III.1.3 — Implement National Contingency Plan; produce ESI maps; provincial "
                "contingency plans for Maputo, Beira, Nacala ports."
            )),
            ("5.4 Beaches and monitoring (AT II and III)", (
                "II.2.5 — Establish beach classification criteria and blue-flag attribution for at "
                "least 10 public beaches: Ponta d'Ouro, Maputo, Xai-Xai, Tofo, Vilanculos, Závora, "
                "Beira, Nacala, Mossuril, Pemba.\n\n"
                "III.2.1.2 — Establish pollution monitoring programmes in estuaries and bays "
                "(Maputo, Incomáti, Zambezi, Púngue, Save).\n\n"
                "III.1.1 — Train at least 100 pollution fiscalization specialists; operationalize "
                "supervision, control, monitoring and fiscalization systems (2016–2025)."
            )),
        ]),
        ("6. Institutional Governance", (
            "Coastal management involves MITADER (lead), MIMAIP, MASA, MICTUR, MIREME, MTC, MIC, "
            "MOPHRH, MCTESTP and others. CONDES (National Sustainable Development Council), chaired "
            "by the Prime Minister, harmonizes development policies. CDS-ZC and CEPAM support "
            "coastal research and sustainable development.\n\n"
            "EGIZC articulates with POLMAR, Sea Law, POEM, PNDT, Decree 97/2020 (coastal zone "
            "regulation) and subsequent EDEA (2024–2033). Action Plan AT IV targets a National "
            "Coastal Management Policy (2016–2018)."
        )),
        ("7. Note on Plastics", (
            "EGIZC does not mention plastics or plastic bags explicitly. Plastic-relevant elements "
            "are implicit through:\n\n"
            "— Solid waste (resíduos sólidos) and garbage (lixo) as land-based and ship-based "
            "pollution sources;\n"
            "— Municipal solid waste (RSU) management gaps and sanitary landfill targets;\n"
            "— Waste management definition requiring reduction, recycling and reuse;\n"
            "— Port waste-reception for ship garbage;\n"
            "— London Protocol / MARPOL framework on waste disposal at sea;\n"
            "— Beach blue-flag classification (coastal cleanliness/water quality);\n"
            "— Land-based pollution protocol (LBSA) under Nairobi Convention.\n\n"
            "Binding operational plastic rules developed later in Decree 16/2015 (plastic bags), "
            "Decree 97/2020, EDEA (2024) and POEM. EGIZC strategy period concluded in 2025."
        )),
    ]

    for heading, content in sections:
        add_heading(doc, heading, level=1)
        if isinstance(content, str):
            add_para(doc, content)
        else:
            for subheading, body in content:
                add_heading(doc, subheading, level=2)
                add_para(doc, body)

    doc.add_page_break()
    add_heading(doc, "4P Index Coding Summary", level=1)
    add_para(
        doc,
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — biodiversity/natural "
        f"preservation pillar not overcoded)\n"
        f"See: 4P_Index_Mozambique_EGIZC_2016_2025.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_EGIZC_2016_2025_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique EGIZC 2016–2025 — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — EGIZC (2016–2025)</h1>
  <p>Integrated Coastal Zone Management Strategy. Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_EGIZC_2016_2025_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_egizc_2016_2025.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_EGIZC_2016_2025"
    df = build_dataframe()
    paths = [
        export_csv(df, basename),
        export_xlsx(df, basename),
        create_translation_doc(),
        create_index_html(basename),
    ]
    print("Generated:")
    for p in paths:
        print(f"  {p}")


if __name__ == "__main__":
    main()
