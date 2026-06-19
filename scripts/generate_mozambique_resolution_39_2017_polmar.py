#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 39/2017 (POLMAR)."""

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
        "Política e Estratégia do Mar – POLMAR (Resolução n.º 39/2017)"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz172386.pdf",
    "policy_year": 2017,
    "policy_objective": (
        "Consolidate a national agenda for sustainable, integrated multi-sectoral management "
        "of marine and coastal spaces; develop a profitable and sustainable blue economy; "
        "strengthen state sovereignty over jurisdictional waters; and address weak planning, "
        "monitoring and coordination of marine resource use including marine pollution and waste."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.50,
    "policy_type_justification": (
        "Council of Ministers Resolution n.º 39/2017 under Article 204(f) of the Constitution — "
        "national policy and strategy document (POLMAR), not binding regulation or parliamentary "
        "legislation. Revokes no prior decree; sets strategic framework for sectoral policies."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "consumption, disposal, environmental leakage"
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        "Pillar D strategies D.3–D.4: pollution taxes and user-pays fees for marine resource use; "
        "polluter-pays principle (sec. 15e) for prevention and elimination of pollution costs; "
        "no dedicated plastic budget allocation in resolution text."
    ),
    "policy_score": "[auto]",
}

# Six consolidated instruments — seven pillars not overcoded line-by-line.
INSTRUMENTS = [
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Sec. 15e–f, Pillar D (D.3–D.4): Establishes polluter-pays and user-pays principles "
            "for marine/coastal activities; strategies require pollution from marine resource "
            "loss to be calculated and taxed, and project proponents to bear collateral "
            "environmental damage costs including waste and risk management."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 15e: "Princípio do poluidor pagador" — costs for "prevenção e eliminação '
            'da poluição"; Strategy D.3(3): "taxada a poluição que resulta em perda ou diminuição '
            'do valor dos recursos marinhos"; D.4: user-pays for project environmental costs.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic economic principle — implementation depends on subsequent legislation "
            "(e.g. Decree 45/2006 fees). Not coded separately: general access fees D.3(1)–(2)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pillar A (22b, strategies A.b, A.b2): Government policy to repress marine and "
            "coastal pollution and irregular resource use under national and ratified "
            "international law; strengthen integrated maritime inspection and coastal/lacustrine "
            "police (PRM) and municipal enforcement."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Pillar A 22b(ii): repress "poluição marinha e costeira"; Strategy A.b1: operationalise '
            'Maritime Courts; A.b2: strengthen "PRM - Polícia Costeira, Lacustre e Fluvial" '
            'and municipal police; A.a1: integrated multidisciplinary inspection system.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Governance/enforcement direction — operationalised through Decrees 45/2006, "
            "51/2024 inspection. Not coded separately: piracy/trafficking provisions (A.b i)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pillar C (35d, strategies C.d1–C.d3): Promote management of marine waste from "
            "multiple sources; inventory and map coastal/marine waste concentrations; establish "
            "multisectoral marine waste management, control and inspection; strengthen supervision "
            "of ship/platform waste treatment and hydrocarbon spill response."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Pillar C 35d: "gestão dos resíduos marinhos emanados de várias fontes"; C.d1: '
            'inventory/map waste sources; C.d2: "sistema multissectorial de gestão de resíduos '
            'marinhos"; C.d3: control waste from vessels and platforms and hydrocarbon spills.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary plastic-relevant instrument in POLMAR — explicit marine waste (resíduos "
            "marinhos). Glossary defines resíduos as discarded substances/objects."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pillar C (35c, strategy C.c): Improve management of river basins and land-based "
            "pollution sources affecting marine ecosystems — erosion, water flows, sedimentation "
            "and pollution control from terrestrial activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Pillar C 35c: "gestão das bacias hidrográficas, das fontes de poluição pelas '
            'actividades baseadas em terra"; Strategy C.c: control erosion, flows, sedimentation '
            'and pollution impacting marine ecosystems.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Land-to-sea litter pathway control — complements marine waste instrument C.d. "
            "Links to coastal EGIZC and watershed regulations."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Pillar DA (strategy DA.a2): Ensure application of norms for port security and "
            "prevention of maritime and coastal pollution at national ports as part of modern "
            "port system development."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Strategy DA.a2: "Assegurar a aplicação das normas inerentes (1) à segurança '
            'portuária e (2) à prevenção da poluição marítima e costeira"; responsibility '
            'MTC + MIMAIP with MITADER, MINT, SEPRIV.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Port pollution prevention — aligns with Decree 45/2006 port reception obligations. "
            "Not coded separately: port competitiveness/logistics strategies (DA.a1)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pillar DC (strategies DC.d2–DC.d3): Ecosystem-based fisheries management plans; "
            "combat illegal, unreported and unregulated fishing and destructive fishing practices "
            "that contribute to marine litter including abandoned fishing gear."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'DC.d2: "planos de gestão das pescarias... abordagem ecossistémica"; DC.d3: '
            '"Combater a pesca ilegal, não reportada e não regulamentada e as práticas de '
            'pesca destrutivas".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Fishing gear/litter relevance indirect — no explicit plastic or ghost gear mention. "
            "Bundled fisheries strategies; not coded: aquaculture, post-harvest loss (DC.e1)."
        ),
    },
]

COLUMNS = [
    "policy_name", "policy_url", "policy_year", "policy_objective", "policy_target",
    "policy_target_text", "policy_type", "policy_type_justification", "policy_integration",
    "policy_sectors_list", "policy_circularity", "policy_lifecycle_phases_list",
    "policy_budget", "policy_budget_text", "policy_score",
    "instrument_type", "instrument_lifecycle_stage", "instrument_description",
    "instrument_in_force", "instrument_implementation", "instrument_implementation_text",
    "instrument_score", "comments",
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
        ws.column_dimensions[letter].width = min(45, max(12, len(str(ws.cell(1, col).value or "")) * 0.9))
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

    title = doc.add_heading("Sea Policy and Strategy — POLMAR", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 39/2017 of 14 September")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Política e Estratégia do Mar – POLMAR\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz172386.pdf")

    sections = [
        ("Resolution No. 39/2017", (
            "The Council of Ministers resolves:\n"
            "Sole article. The Sea Policy and Strategy, abbreviated POLMAR, is approved as an "
            "annex forming an integral part of this Resolution.\n"
            "Approved by the Council of Ministers on 15 August 2017."
        )),
        ("Context and focal problem", (
            "POLMAR frames coastal and maritime economic development within a sustainable blue "
            "economy. Focal problem: low economic, social and environmental benefits caused by "
            "inadequate spatial planning, weak monitoring and poor coordination of access to and "
            "exploitation of natural capital at sea and in coastal areas.\n\n"
            "Objectives include strengthening maritime sovereignty, developing a sustainable blue "
            "economy, spatial planning of maritime and coastal zones, integrated governance, and "
            "international cooperation."
        )),
        ("Guiding principles (selected)", (
            "Polluter pays (sec. 15e): polluters bear costs of environmental restoration and "
            "pollution prevention/elimination from marine and coastal activities.\n\n"
            "User pays (sec. 15f): access and use of marine and coastal resources are paid by users.\n\n"
            "Integrated management, precautionary principle, community participation, and "
            "multi-sectoral coordination also guide implementation."
        )),
        ("Seven pillars (overview)", (
            "Pillar A — Governance and legal framework\n"
            "Pillar B — Inter-institutional coordination\n"
            "Pillar C — Marine and coastal environment\n"
            "Pillar D — Economic development\n"
            "Pillar E — Territorial development\n"
            "Pillar F — Human capital development\n"
            "Pillar G — International cooperation\n\n"
            "Sub-pillars under economic development cover ports (DA), maritime transport and "
            "naval industry (DB), fisheries (DC), and culture/tourism/sport (DD)."
        )),
        ("Plastic-relevant policy lines (Pillar C and related)", (
            "Pillar C identifies threats including pollution from multiple sources and coastal "
            "degradation. Policy lines include:\n"
            "c) Improve watershed and land-based pollution management affecting marine ecosystems;\n"
            "d) Promote management of marine waste from various sources.\n\n"
            "Strategies:\n"
            "C.d1 — Inventory and map marine waste concentration sources on coast and at sea.\n"
            "C.d2 — Establish multisectoral marine waste management, control and inspection system.\n"
            "C.d3 — Strengthen supervision of ship/platform waste and hydrocarbon spill response.\n\n"
            "Pillar A represses marine and coastal pollution. Pillar DA requires port pollution "
            "prevention norms. Pillar DC combats destructive and IUU fishing."
        )),
        ("Note on plastics", (
            "POLMAR does not mention plastics explicitly. The glossary defines waste (resíduos) "
            "as discarded substances or objects. Marine waste management strategies and pollution "
            "principles are the primary plastic-relevant elements. Implementation is through "
            "sectoral regulations (e.g. Decree 45/2006, Decree 94/2014, Decree 51/2024)."
        )),
    ]

    for heading, body in sections:
        add_heading(doc, heading, level=1)
        add_para(doc, body)

    doc.add_page_break()
    add_heading(doc, "4P Index Coding Summary", level=1)
    add_para(
        doc,
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — seven pillars not overcoded)\n"
        f"See: 4P_Index_Mozambique_Resolution_39_2017_POLMAR.xlsx",
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Resolution_39_2017_POLMAR_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 39/2017 POLMAR — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 39/2017 (POLMAR)</h1>
  <p>Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Resolution_39_2017_POLMAR_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_39_2017_polmar.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_39_2017_POLMAR"
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
