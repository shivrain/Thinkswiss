#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 7/2021 (PNDT)."""

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
        "Plano Nacional de Desenvolvimento Territorial — PNDT "
        "(Resolução n.º 7/2021)"
    ),
    "policy_url": (
        "https://sibmoz.gov.mz/content/uploads/2022/06/"
        "Plano-Nacional-de-Desenvolvimento-Territorial-1.pdf"
    ),
    "policy_year": 2021,
    "policy_objective": (
        "Approve the National Territorial Development Plan (PNDT) and its Action Plan "
        "under Law 19/2007 to establish Mozambique's long-term territorial vision "
        "(horizon 2040), strategic objectives, territorial model and interministerial "
        "action programme. Plastic-relevant scope: the 2040 sustainability vision targets "
        "drastic reduction of risks from sanitation deficiencies, waste collection and "
        "treatment, and water quality; action items cover urban water/sanitation services, "
        "pollution control for hazardous/polluting industries, and packaging activities "
        "in logistics zones — no explicit plastics (plástico) reference; solid waste "
        "(resíduos) includes plastic fractions."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.50,
    "policy_type_justification": (
        "Parliamentary Resolution n.º 7/2021 (approved 11 November 2021, published "
        "28 December 2021, Boletim da República Series I No. 250) under Constitution "
        "Arts. 110, 117, 181 and Law 19/2007 Art. 13 — national territorial development "
        "plan with annexed action plan (IE items 2020–2040), not binding sectoral waste "
        "or plastic regulation. Art. 6 requires inclusion of PNDT actions in the Economic "
        "and Social Plan (PES). Operational waste rules remain in Decree 94/2014, "
        "Decree 16/2015, etc."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": (
        "municipalities, water, waste management, industry, packaging, environment"
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": (
        "production, disposal, environmental leakage"
    ),
    "policy_budget": 0,
    "policy_budget_text": "",
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "2040 Sustainability and Resilience Challenge 4: drastic reduction of risk "
            "factors associated with deficiencies in sanitation, waste collection and "
            "treatment (recolha e tratamento de resíduos), and water quality — the only "
            "explicit solid-waste reference in the PNDT vision. Resíduos includes plastic "
            "waste fractions though plastics are not named."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Sustainability challenges (OEDT 6 context): "Redução drástica dos factores '
            'de risco associados às deficiências de saneamento, de recolha e tratamento '
            'de resíduos e de qualidade da água" — aspirational 2040 vision without '
            "quantified waste-collection, recycling or plastic-reduction metrics; "
            "operationalised through sectoral IE actions (notably IE 5.2)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary solid-waste instrument — only explicit resíduos reference. Not coded "
            "separately: urban land-use discipline or rural conservation controls without "
            "waste linkage."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Sustainability Challenge 3 — Water resource protection: protect hydrological "
            "resources and drinking-water sources for human consumption, avoiding pollution "
            "and securing ecological flows. IE 5.2 urban water/sanitation programme "
            "articulates service delivery with urban land-use patterns."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Challenge 3: "Protecção dos recursos hídricos e das fontes de água para o '
            'consumo humano, evitando a poluição"; IE 5.2: progressive equity in urban '
            'water and sanitation access, "fixação de metas progressivas para os níveis '
            'de serviço", financing and institutional frameworks for sub-sectors '
            "(2020–2040; MOPHRH, Municipal Councils)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Water/sanitation instrument — pollution prevention and progressive service "
            "targets for urban sanitation (includes wastewater; solid waste collection "
            "addressed in vision challenge 4). Not coded: universal electricity (IE 5.3)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "IE 7.9 — Review of the Delegated Management Framework for water and sanitation "
            "services: redistribute competencies among FIPAG, AIAS and DNAAS to align with "
            "2018 constitutional decentralisation and PNDT urban hierarchy; strengthen CRA "
            "mandates — institutional framework governing urban sanitation service delivery "
            "relevant to municipal solid waste and wastewater management."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'IE 7.9: "Revisão do Quadro de Gestão Delegada dos serviços de água e '
            'saneamento" — competency redistribution for decentralised water/sanitation '
            "governance in cities and towns (2020–2024; MOPHRH, Council of Ministers); "
            "IE 5.2 references adequacy of institutional and financing frameworks for "
            "sub-sectors."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Governance/financing instrument for urban sanitation — no plastic-specific "
            "provisions. Not coded separately: IE 7.8 territorial-management capacity "
            "programme."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "IE 7.24 — Action plan for situations of grave environmental or public-health "
            "risk: cadastre of hazardous or heavily polluting industrial activities; "
            "implementation of mitigation action plans; monitoring with pollution "
            "prevention/control measures and possible relocation of activities or "
            "settlements; community campaigns on critical environment and public-health "
            "problems — covers industrial pollution pathways including potential plastic "
            "manufacturing/processing facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'IE 7.24: cadastro of "actividades e indústrias perigosas ou fortemente '
            'poluidoras"; "medidas de prevenção e controlo da poluição"; priority areas '
            "Moatize/Tete, Matola/Maputo, hydrocarbon and mineral port zones (2020–2024 "
            "1st priority; 2025–2029 2nd priority; Environment sector)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Pollution-control instrument — no explicit plastic industry targeting. "
            "Bundled with disaster/extreme-event response. Not coded: IE 7.25 military "
            "servitude areas."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "IE 3.7 — Logistics and consolidation zones: create a network of logistics "
            "platforms at intermodal nodes, strategic transport junctions, airports and "
            "peripheries of main cities for consolidation, storage, packaging (embalagem) "
            "and cargo distribution — packaging activities imply plastic packaging "
            "materials in freight/logistics chains."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'IE 3.7: "actividades de grupagem, armazenamento, embalagem e distribuição '
            'de carga" at intermodal logistics platforms linked to economic motor spaces '
            "(VT4); national scope 2020–2024; Industry and Trade sector."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Packaging/production-side instrument in logistics context — no waste "
            "management or EPR provisions. Not coded: agro-logistics cold-chain centres "
            "(IE 3.8) without packaging mention."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "OEDT 6 / IE 2.3 — Prevent natural and anthropogenic risks and safeguard "
            "ecosystem sustainability: National Urban Rehabilitation and Qualification "
            "Programme (PNRC) develops urban and environmental qualification methodology "
            "for city centres with long-term (20-year) infrastructure and public-space "
            "improvement financing — indirect relevance to urban environmental quality "
            "including waste-service conditions in central urban areas."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'IE 2.3 (PNRC): "qualificação urbana e ambiental das áreas centrais das '
            'cidades"; catalytic financing mechanism for 20-year municipal qualification '
            "programmes — priority Grande Maputo, Beira, Nampula, Tete (2020–2040; "
            "Territorial Planning, Environment, Finance sectors)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Urban environmental-qualification instrument — no explicit waste/plastic "
            "measures. Coded for ambiental qualification of cities affecting overall "
            "urban service conditions. Not coded: housing strategy (IE 2.6)."
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

    title = doc.add_heading("National Territorial Development Plan — PNDT", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 7/2021 — Horizon 2040")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Plano Nacional de Desenvolvimento Territorial — PNDT "
        "(Resolução n.º 7/2021)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://sibmoz.gov.mz/content/uploads/2022/06/"
        "Plano-Nacional-de-Desenvolvimento-Territorial-1.pdf"
    )

    sections = [
        ("Resolution No. 7/2021 — Enacting Provisions", (
            "Whereas it is necessary to approve the National Territorial Development Plan "
            "and its Action Plan under paragraphs 1 and 2 of Article 110, paragraph 2(e) "
            "of Article 117 and Article 181 of the Constitution, together with paragraph "
            "(a) of paragraph 1 of Article 13 of Law No. 19/2007 of 18 July (Territorial "
            "Planning Legal Framework), the Assembly of the Republic determines:\n\n"
            "Article 1 (Approval). The National Territorial Development Plan (PNDT) and "
            "its Action Plan, attached hereto, are approved as an integral part of this "
            "Resolution.\n\n"
            "Article 2 (Principles). Preparation and execution of the PNDT are guided by: "
            "legality; justice and impartiality; sustainability (equitable capacity at all "
            "territorial levels without compromising environmental balance and ecosystems); "
            "participation (mandatory public consultations); and transparency.\n\n"
            "Article 3 (Government competences). The Government shall ensure PNDT "
            "implementation; promote private sector, civil society and community "
            "involvement; and ensure technical capacity.\n\n"
            "Article 5 (Protection of pre-existing rights). Implementation must respect "
            "pre-existing rights; resettlement requires fair and transparent compensation "
            "by concessionaires.\n\n"
            "Article 6 (Inclusion in Economic and Social Plan). The Government must include "
            "PNDT implementation actions and budgetary impact in Economic and Social Plan "
            "proposals.\n\n"
            "Article 7 (Intersectoral coordination). CIOT (Interministerial Commission for "
            "Territorial Planning) coordinates implementation, monitoring and evaluation.\n\n"
            "Article 8 (Territorial Planning Observatory). Government shall create OOT as "
            "an inclusive participatory platform.\n\n"
            "Articles 9–10. Each government sector executes its Action Plan items; "
            "provinces, districts and municipalities shall create Technical Implementation "
            "Offices (GTI).\n\n"
            "Approved by the Assembly of the Republic on 11 November 2021. Published "
            "28 December 2021, Boletim da República Series I No. 250.\n\n"
            "Signed: Esperança Laurinda Francisco Nhiuane Bias, President of the Assembly."
        )),
        ("Part I — What is the PNDT?", (
            "The PNDT is simultaneously a Vision, a Strategy, a Framework of Guiding "
            "Principles and an Instrument for public-policy efficiency.\n\n"
            "Its first mission is to establish the long-term vision for occupation, use "
            "and transformation of national territory aligned with strategic country "
            "objectives, preventing negative effects of resource exploitation driven by "
            "global markets.\n\n"
            "The PNDT interprets ENDE 2015–2035 and Agenda 2025 objectives with a spatial "
            "dimension, interrelating three purposes: prosperous and competitive "
            "Mozambique; safe and inclusive Mozambique; sustainable Mozambique.\n\n"
            "Horizon: 2040 ('Moçambique 2040')."
        )),
        ("Strategic Framework — Nine Objectives (OEDT)", (
            "OEDT 1: Integrate territory, build dynamic internal market, reduce regional "
            "disparities.\n"
            "OEDT 2: Increase national wealth — human, built and institutional capital.\n"
            "OEDT 3: Promote catalytic spaces for economic growth.\n"
            "OEDT 4: Mobilise territorial potential and diversify rural economies.\n"
            "OEDT 5: Balanced urban network with inclusive, creative cities.\n"
            "OEDT 6: Prevent natural and anthropogenic risks; safeguard ecosystems and "
            "biodiversity.\n"
            "OEDT 7: Combat poverty and promote social inclusion.\n"
            "OEDT 8: Value Mozambique's geographic position and regional cooperation.\n"
            "OEDT 9: Strengthen national cohesion, citizenship and territorial governance."
        )),
        ("Sustainability and Resilience Challenges (plastic-relevant)", (
            "Territorial ordering challenges include:\n"
            "1. Discipline urban settlements and requalify urban areas.\n"
            "2. Control rural interventions near conservation and biodiversity areas.\n"
            "3. Protect water resources and drinking-water sources, avoiding pollution.\n"
            "4. Reallocate land to socially productive uses.\n"
            "5. Prior assessment and coordination of sectoral interventions.\n"
            "6. Finance urban infrastructure through equitable land-value capture.\n\n"
            "Sustainability and resilience challenges include:\n"
            "1. Better scientific knowledge of natural and anthropogenic hazards.\n"
            "2. Disaster prevention and mitigation.\n"
            "3. Reduce hydrological vulnerability in some provinces.\n"
            "4. Drastic reduction of risk factors from deficiencies in sanitation, "
            "waste collection and treatment, and water quality.\n"
            "5. Preserve ecological diversity and ecoregions.\n"
            "6. Adequate landscape management.\n"
            "7. Articulate biodiversity conservation with economic development."
        )),
        ("Part II — Action Plan: Key Plastic-Relevant Items", [
            ("IE 5.2 — Urban safe water and sanitation services", (
                "Organise and implement urban water supply and sanitation services "
                "articulated with urban land-use dynamics, involving:\n"
                "• Progressive equity in service access;\n"
                "• Progressive targets for minimum service levels;\n"
                "• Adequate governance and quality-control frameworks;\n"
                "• Institutional and financing framework adequacy for sub-sectors;\n"
                "• Review/elaboration of master plans for systems aligned with territorial "
                "planning instruments.\n\n"
                "Priority: five urban-system levels for water; N1–N3 levels for sanitation. "
                "Period: 2020–2040. Entities: MOPHRH, Municipal Councils."
            )),
            ("IE 7.9 — Delegated management framework for water and sanitation", (
                "Review the Delegated Management Framework for water and sanitation "
                "services; redistribute competencies among FIPAG, AIAS and DNAAS to align "
                "with 2018 constitutional decentralisation and PNDT urban hierarchy; "
                "strengthen CRA mandates. Cities and towns, 2020–2024."
            )),
            ("IE 7.24 — Action plan for grave environmental/public-health risks", (
                "Cadastre hazardous or heavily polluting industrial activities; implement "
                "action plans to mitigate or eliminate risks and respond to extreme events "
                "or catastrophes. Interventions include monitoring and pollution "
                "prevention/control measures, possible relocation of activities or "
                "settlements, and community campaigns on critical environment and public "
                "health problems.\n\n"
                "1st priority (2020–2024): Moatize/Tete, Matola/Maputo, hydrocarbon and "
                "mineral port areas.\n"
                "2nd priority (2025–2029): other areas with high polluting-industry "
                "concentration."
            )),
            ("IE 3.7 — Logistics zones with packaging activities", (
                "Create a network of logistics platforms at intermodal nodes, strategic "
                "transport junctions, airports and peripheries of main cities for "
                "consolidation, storage, packaging (embalagem) and cargo distribution, "
                "linked to economic motor spaces (VT4). National, 2020–2024."
            )),
            ("IE 2.3 — Urban environmental qualification (PNRC)", (
                "National Programme for Rehabilitation and Urban and Environmental "
                "Qualification of city centres: develop qualification methodology; create "
                "catalytic financing mechanism for 20-year municipal infrastructure and "
                "public-space improvement programmes. Priority: Grande Maputo, Beira, "
                "Nampula, Tete (2020–2040)."
            )),
        ]),
        ("Note on Plastics and Solid Waste", (
            "The PNDT contains no explicit reference to plastics (plástico). Plastic-relevant "
            "content is limited to:\n\n"
            "— One explicit solid-waste reference in the 2040 vision: 'recolha e tratamento "
            "de resíduos' (waste collection and treatment) alongside sanitation and water "
            "quality deficiencies;\n"
            "— Urban water and sanitation service programmes (IE 5.2, IE 7.9);\n"
            "— Pollution control for hazardous/heavily polluting industries (IE 7.24);\n"
            "— Packaging (embalagem) activities in logistics zones (IE 3.7);\n"
            "— Urban environmental qualification of city centres (IE 2.3 / OEDT 6).\n\n"
            "The plan sets no quantified plastic-reduction, recycling, EPR or dedicated "
            "waste-management budget lines. IE 5.2 establishes progressive sanitation "
            "service targets but not plastic-specific metrics. Art. 6 requires PNDT budget "
            "inclusion in the PES but without waste/plastic line-items in the PNDT text.\n\n"
            "Operational plastic and solid-waste measures remain in Decree 94/2014, "
            "Decree 16/2015, Decree 97/2020, EDEA (Resolution 53/2024) and related "
            "instruments."
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
        f"Instruments coded: {len(INSTRUMENTS)} (plastic/solid-waste relevant only)\n"
        f"E — policy_target: {POLICY_FIELDS['policy_target']} (no quantified waste/plastic targets)\n"
        f"K — policy_circularity: {POLICY_FIELDS['policy_circularity']} "
        f"({POLICY_FIELDS['policy_lifecycle_phases_list']})\n"
        f"M — policy_budget: {POLICY_FIELDS['policy_budget']} (no waste/plastic budget)\n"
        f"See: 4P_Index_Mozambique_Resolution_7_2021_PNDT.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Resolution_7_2021_PNDT_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    base = (
        "https://github.com/shivrain/Thinkswiss/raw/"
        "cursor/mozambique-resolution-7-2021-pndt-04b6/output/mozambique"
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 7/2021 PNDT — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 7/2021 (PNDT)</h1>
  <p>National Territorial Development Plan — horizon 2040. Generated {date.today().isoformat()}.
  Six instruments coded (waste, sanitation, pollution, packaging relevant only).</p>
  <div class="card">
    <a class="download" href="{base}/{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{base}/{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="{base}/Mozambique_Resolution_7_2021_PNDT_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_7_2021_pndt.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_7_2021_PNDT"
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
