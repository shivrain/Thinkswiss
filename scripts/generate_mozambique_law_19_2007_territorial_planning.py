#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Law 19/2007 (LOT)."""

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
    "policy_name": "Lei do Ordenamento do Território (Lei n.º 19/2007)",
    "policy_url": (
        "https://sibmoz.gov.mz/content/uploads/2022/01/"
        "Lei-do-ordenamento-territorial.pdf"
    ),
    "policy_year": 2007,
    "policy_objective": (
        "Establish the legal framework for Mozambique's Territorial Planning Policy, "
        "defining principles, objectives, territorial management system and binding "
        "planning instruments at national, provincial, district and municipal levels. "
        "Plastic-relevant scope: General and Partial Urbanisation Plans must define "
        "sanitation (saneamento) networks alongside transport, communications and energy "
        "(Art. 10(5)(b)); environmental precaution, liability and protection principles "
        "apply to territorial interventions affecting pollution pathways. The law contains "
        "no explicit plastics (plástico) or solid-waste (resíduos) provisions — those "
        "appear in the separate Regulamento da LOT."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Parliamentary Law n.º 19/2007 (approved 11 May 2007, promulgated 11 July 2007, "
        "published 18 July 2007, Boletim da República Series I No. 29) under Constitution "
        "Art. 179(1) — primary binding legislation establishing territorial planning "
        "rights, duties and instruments with force of law once published (Art. 11). "
        "Art. 12 delegates instrument regulation to Council of Ministers decree. "
        "Enabling framework for PNDT (Resolution 7/2021), municipal urban plans and "
        "Decree 94/2014 waste implementation at local level."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": "municipalities, water, waste management, environment",
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": "disposal, environmental leakage",
    "policy_budget": 0,
    "policy_budget_text": "",
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Art. 10(5)(b) — General and Partial Urbanisation Plans (Planos Gerais e "
            "Parciais de Urbanização): must structure and qualify urban land, balancing "
            "urban uses and functions, and define networks for transport, communications, "
            "energy and sanitation (saneamento) plus social facilities — requiring "
            "municipal urban plans to integrate sanitation infrastructure, which encompasses "
            "wastewater and municipal environmental sanitation relevant to solid-waste "
            "service planning (resíduos not named in the law)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 10(5)(b): urbanisation plans "definem as redes de transporte, '
            'comunicações, energia e saneamento, os equipamentos sociais"; Art. 13(d): '
            "municipal-level instruments elaborated and approved by competent local planning "
            "bodies after public consultation; Art. 11: published instruments bind public "
            "entities and citizens."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary plastic/waste-relevant instrument — only explicit saneamento "
            "reference in the law. Solid-waste treatment systems are in the Regulamento "
            "da LOT (separate decree), not coded here. Not coded: Urban Structure Plans "
            "(Art. 10(5)(a)) without sanitation-network requirement."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 4(d)–(e) — Precaution and environmental liability principles: territorial "
            "planning instruments must prioritise prevention of environmentally harmful acts "
            "to avoid significant or irreversible negative impacts; public and private "
            "entities are responsible for territorial interventions that damage environmental "
            "quality and must repair harm and compensate citizens — applicable to pollution "
            "including litter and waste-related environmental damage."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 4(d): "priorizar o estabelecimento de sistemas de prevenção de actos '
            'lesivos ao ambiente"; Art. 4(e): "responsabilidade... por qualquer '
            'intervenção sobre o território, que possa ter causado danos ou afectado a '
            'qualidade do ambiente" with obligation of reparation and compensation.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "General environmental-liability framework — no plastic-specific measures. "
            "Not coded separately: participation (4b), equality of access (4c) or "
            "legal certainty (4f) principles."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 5(1)–(2)(d) — Territorial planning objectives: ensure sustainable use "
            "of national space and natural resources, promote quality of life and "
            "environmental protection/conservation; specifically preserve ecological balance, "
            "soil fertility, air purity, ecosystems, forests, water resources, riparian zones "
            "and coastal foreshore (orla marítima) — relevant to land-based and coastal "
            "pollution pathways including plastic leakage to water bodies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 5(1): promotion of quality of life and "protecção e conservação do meio '
            'ambiente"; Art. 5(2)(d): preserve "pureza do ar", water resources, riparian '
            "zones and coastal foreshore while balancing community needs with environmental "
            "safeguard objectives."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Environmental-protection objectives — indirect plastic-pollution relevance via "
            "water/coastal/air quality. Not coded: urban requalification (5(2)(b)) or rural "
            "economic potential (5(2)(c)) without waste linkage."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 24 — Duty of environmental responsibility: materialisation of territorial "
            "development actions must be carried out responsibly in environmental terms, "
            "regardless of the financial origin of the investment — applies to public and "
            "private territorial projects including waste-related infrastructure and "
            "polluting land uses."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 24: "A materialização das acções de desenvolvimento territorial tem de ser '
            'realizada de forma responsável em termos ambientais, independentemente da '
            'origem financeira do investimento."'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Cross-cutting environmental-duty provision in citizens' rights chapter. "
            "Not coded: information rights (Art. 21) or participation rights (Art. 22)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 27 — Environmental fiscalization: administrative and environmental "
            "inspection of norms established for elaboration of territorial planning "
            "instruments is exercised under applicable legislation — enforcement gateway "
            "for compliance with urban plans including required sanitation networks and "
            "environmental safeguards in planning instruments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 27: "A fiscalização administrativa e ambiental das normas estabelecidas '
            'para a elaboração dos instrumentos de ordenamento territorial é exercida nos '
            'termos da legislação em vigor"; Art. 29: non-compliance results in '
            "administrative penalties and fines (details in regulation)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Enforcement/coordination instrument — refers to external environmental "
            "legislation for sanctions. Not coded: monitoring system (Art. 26) for "
            "environmental statistical data collection (not waste operations)."
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

    title = doc.add_heading("Territorial Planning Law", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Law No. 19/2007 of 18 July")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Lei do Ordenamento do Território (Lei n.º 19/2007)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://sibmoz.gov.mz/content/uploads/2022/01/"
        "Lei-do-ordenamento-territorial.pdf"
    )

    sections = [
        ("Preamble and Enacting Formula", (
            "Territory is the physical basis of the State — the spatial reality on which "
            "Mozambican society is established and develops. The Constitution defines "
            "territory as unitary, indivisible and inalienable.\n\n"
            "This Territorial Planning Law provides the legal framework for the Territorial "
            "Planning Policy to achieve rational and sustainable use of natural resources, "
            "preservation of environmental balance, national cohesion, regional potential, "
            "citizens' quality of life, balance between rural and urban living conditions, "
            "improvement of housing, infrastructure and urban systems, and security of "
            "populations vulnerable to natural or human-caused calamities.\n\n"
            "Under Constitution Art. 179(1), the Assembly of the Republic determines:"
        )),
        ("Chapter I — General Principles", [
            ("Article 1 — Definitions (summary)", (
                "Key definitions include: local community; sustainable development; "
                "territorial planning instruments; territorial planning; territorial "
                "planning plan; territorial management system; rural land; urban land; "
                "territory; tangible and intangible goods; social cohesion rupture."
            )),
            ("Articles 2–3 — Object and scope", (
                "Art. 2: The law creates the legal framework for territorial planning and "
                "materialises the Territorial Planning Policy through planning instruments.\n\n"
                "Art. 3: Applies to the entire national territory; regulates relations among "
                "administration levels and between the State and public/private subjects "
                "including local communities."
            )),
            ("Article 4 — Principles", (
                "Territorial planning follows: sustainability; public participation; equality "
                "of access to land, natural resources, infrastructure and public services; "
                "precaution (priority to preventing environmentally harmful acts); "
                "environmental liability and compensation; legal certainty; publicity of "
                "planning instruments."
            )),
            ("Article 5 — Objectives", (
                "Territorial planning ensures organisation of national space and sustainable "
                "use of natural resources, promoting quality of life and environmental "
                "protection.\n\n"
                "Specific objectives include: guarantee current occupation rights; requalify "
                "spontaneous/degraded urban areas; valorise rural economic potential; "
                "preserve ecological balance, soil fertility, air purity, ecosystems, forests, "
                "water resources, riparian zones and coastal foreshore; defend built and "
                "landscape heritage; harmonise environmental and socio-economic policies; "
                "optimise natural resource management; manage conflicts privileging agreements."
            )),
            ("Articles 6–7 — State duty and natural resources", (
                "Art. 6: State and Local Authorities must promote, guide, coordinate and "
                "monitor territorial planning in the public interest.\n\n"
                "Art. 7: Territorial planning must respect land and natural resource use "
                "under applicable law and organise public domain (waters, roads, beaches, "
                "protected areas, ports, airports, monuments, etc.)."
            )),
        ]),
        ("Chapter II — Territorial Management System", [
            ("Articles 8–9 — Levels and characterisation", (
                "Intervention levels: national, provincial, district and municipal "
                "(autárquico). Each level defines strategies, directives and plans within "
                "hierarchical coordination."
            )),
            ("Article 10 — Territorial planning instruments (plastic-relevant)", (
                "National level: National Territorial Development Plan (PNDT); Special "
                "Territorial Planning Plans.\n\n"
                "Provincial level: Provincial Territorial Development Plans.\n\n"
                "District level: District Land-Use Plans.\n\n"
                "Municipal level:\n"
                "a) Urban Structure Plans;\n"
                "b) General and Partial Urbanisation Plans — structure and qualify urban "
                "land, balancing uses and functions, and define networks for transport, "
                "communications, energy and sanitation (saneamento), plus social facilities;\n"
                "c) Detail Plans for specific central urban areas.\n\n"
                "General instruments: land qualification, land classification, National Land "
                "Cadastre, environmental/social/economic inventories, zoning."
            )),
            ("Articles 11–20 — Legal regime (summary)", (
                "Art. 11: Published planning instruments have force of law and bind all "
                "public entities, citizens, communities and private legal persons.\n\n"
                "Arts. 12–18: Council of Ministers approves instrument regulation; "
                "competences for elaboration/approval at each level; ratification; "
                "amendment; review; suspension; publicity.\n\n"
                "Arts. 19–20: Sectoral coordination and compatibility; expropriation for "
                "public-interest territorial projects with fair compensation."
            )),
        ]),
        ("Chapters III–VI — Rights, Monitoring and Final Provisions", [
            ("Articles 21–24 — Citizen rights and duties", (
                "Art. 21: Right to information on planning instruments.\n"
                "Art. 22: Right to participate in elaboration, execution, amendment and "
                "review through public consultation.\n"
                "Art. 23: Guarantees for those harmed (impugnation, popular action, "
                "complaints to Public Prosecutor and Ombudsman).\n"
                "Art. 24: Territorial development actions must be environmentally responsible "
                "regardless of investment financing source."
            )),
            ("Articles 25–27 — Evaluation, monitoring and inspection", (
                "Art. 25: Evaluation reports at end of each legislature.\n"
                "Art. 26: Monitoring with periodic reports; national territorial information "
                "system collecting environmental, social and technical data.\n"
                "Art. 27: Administrative and environmental inspection of planning norms "
                "under applicable legislation."
            )),
            ("Articles 28–31 — Transitional and final provisions", (
                "Art. 28: Existing plans remain in force until reconducted under regulation.\n"
                "Art. 29: Non-compliance results in administrative penalties and fines "
                "(regulated).\n"
                "Art. 30: Government adopts regulatory measures within one year.\n"
                "Art. 31: Law enters into force 90 days after publication.\n\n"
                "Approved by the Assembly of the Republic on 11 May 2007. Promulgated "
                "11 July 2007. Published 18 July 2007.\n"
                "Signed: Armando Emílio Guebuza, President of the Republic."
            )),
        ]),
        ("Note on Plastics and Solid Waste", (
            "Law 19/2007 contains no explicit reference to plastics (plástico) or solid "
            "waste (resíduos sólidos). Plastic-relevant content in the law is limited to:\n\n"
            "— Art. 10(5)(b): General and Partial Urbanisation Plans must define "
            "sanitation (saneamento) networks;\n"
            "— Arts. 4–5: Environmental precaution, liability and protection principles "
            "including water resources and coastal foreshore;\n"
            "— Art. 24: Environmental responsibility for territorial development actions;\n"
            "— Art. 27: Environmental inspection of planning norms.\n\n"
            "Explicit solid-waste treatment systems and reception/processing zones appear "
            "in the separate Regulamento da Lei de Ordenamento do Território (Council of "
            "Ministers decree), not in this law. Operational plastic and urban solid-waste "
            "rules remain in Decree 94/2014, Decree 16/2015, Decree 97/2020 and related "
            "instruments. PNDT (Resolution 7/2021) is approved under Art. 10(2)(a) and "
            "Art. 13(1)(a) of this law."
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
        f"Instruments coded: {len(INSTRUMENTS)} (sanitation/environment relevant only)\n"
        f"E — policy_target: {POLICY_FIELDS['policy_target']}\n"
        f"K — policy_circularity: {POLICY_FIELDS['policy_circularity']} "
        f"({POLICY_FIELDS['policy_lifecycle_phases_list']})\n"
        f"G — policy_type: {POLICY_FIELDS['policy_type']} (parliamentary law)\n"
        f"See: 4P_Index_Mozambique_Law_19_2007_Territorial_Planning.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Law_19_2007_Territorial_Planning_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    base = (
        "https://github.com/shivrain/Thinkswiss/raw/"
        "cursor/mozambique-law-19-2007-territorial-planning-04b6/output/mozambique"
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Law 19/2007 — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Law 19/2007 (Territorial Planning)</h1>
  <p>Lei do Ordenamento do Território. Generated {date.today().isoformat()}.
  Five instruments coded (sanitation networks and environmental provisions only).</p>
  <div class="card">
    <a class="download" href="{base}/{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{base}/{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="{base}/Mozambique_Law_19_2007_Territorial_Planning_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_law_19_2007_territorial_planning.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Law_19_2007_Territorial_Planning"
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
