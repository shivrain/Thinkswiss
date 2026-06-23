#!/usr/bin/env python3
"""Generate 4P Index coding for ANAC Tourism Concessions Manual (2012)."""

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
        "Concessões de Turismo nas Áreas Protegidas em Moçambique: "
        "Manual para Operadores e Concessionários (2012)"
    ),
    "policy_url": (
        "https://www.anac.gov.mz/wp-content/uploads/2017/08/"
        "Manual-Concessoes-para-Sector-Privado_2012.pdf"
    ),
    "policy_year": 2012,
    "policy_objective": (
        "Guide private tourism operators and concessionaires through the process "
        "of establishing and running tourism concessions in Mozambique's protected "
        "areas (USAID SPEED Program; published April 2012, authors Rita Casimiro "
        "and Anna Spenceley). Plastic-relevant scope: requires environmental "
        "management plans with waste-management provisions; guides operators on "
        "packaging reduction, disposable plastic bottle avoidance, litter "
        "('pack it in, pack it out'), wildlife-proof waste bins, waste-disposal "
        "area siting, garbage removal procurement, and solid/liquid waste in EIAs — "
        "voluntary best-practice guidance hosted by ANAC, not binding legislation."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "Technical guidance manual (Manual para Operadores e Concessionários) "
        "prepared under USAID/Mozambique SPEED contract (DAI and Nathan Associates, "
        "published 23 April 2012, SPEED-Reports-2012-005) — operational handbook "
        "for tourism concession applicants and operators in protected areas, hosted "
        "on ANAC website. Not parliamentary law or executive regulation; implements "
        "best-practice expectations referenced in concession contracts, EIA requirements "
        "(Decree 45/2004) and MITUR/ANAC management frameworks."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": (
        "tourism, waste management, municipalities, industry, environment, packaging"
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": (
        "consumption, disposal, environmental leakage"
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
            "Sec. 5.2.2 Development stage — Environmental Management Plan (EMP): "
            "concessionaires must develop a comprehensive environmental management plan "
            "complying with Mozambican legislation; the EMP must include plans for "
            "waste management, water management and energy management to minimise waste "
            "production and resource use — core waste-planning requirement for tourism "
            "operations in protected areas."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 5.2.2: "The environmental management plan must include plans for waste '
            'management, water management, and energy management to minimise waste '
            "production and use of water and energy\"; linked to EIA/licensing under "
            "Law 20/97 and Decree 45/2004 (Sec. 2.1.5)."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary waste-management planning instrument — mandatory in development "
            "stage per manual. Not coded separately: energy or biodiversity sections."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Sec. 5.2.2 — Environmental Management System (EMS, e.g. ISO 14001): "
            "recommended for larger operations to audit consumption of packaging "
            "resources and food, and waste production including sewage, wastewater, "
            "organic and inorganic waste — systematic review to reduce waste volumes "
            "including plastic packaging and solid waste streams."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Sec. 5.2.2: EMS \"examines the amount of consumption (e.g. energy, water, "
            "packaging resources, food) and waste production (e.g. sewage, waste water, "
            "organic and inorganic waste)\"; recommended not mandatory — voluntary "
            "certification pathway for concessionaires."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Voluntary EMS instrument — packaging and inorganic waste explicitly listed. "
            "Not coded: local-materials sourcing targets (30% within 50 km)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Sec. 5.2.2(d) / lifecycle framework — Development and decommissioning: "
            "siting of access roads, materials storage sites and waste disposal areas "
            "must be planned carefully in relation to the EIA to minimise negative "
            "impacts; project lifecycle table includes waste disposal at construction "
            "and materials use/disposal during operation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 5.2.2(d): "Considering the siting of access roads, materials storage '
            'sites, and waste disposal areas carefully in relation to the EIA"; '
            "lifecycle diagram: Construction — waste disposal; Operation — materials "
            "use and disposal; Decommissioning — waste disposal and remediation."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Spatial planning for waste infrastructure — EIA-linked. Table 4 business "
            "plan template also requires environmental assessment of waste (solid and "
            "liquid)."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Sec. 5.2.3 Operations — Waste best practices: guidance to reduce packaging "
            "(bulk purchasing, supplier engagement on packaging disposal costs); avoid "
            "disposable plastic water bottles (use refillable bottles); wildlife-proof "
            "bins for food waste; refillable soap dispensers; paper recycling; compost "
            "and vermiculture for organic waste — explicit plastic consumption reduction "
            "measure for guest services."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 5.2.3 Waste: "use purified water in re-fillable bottles rather than '
            'water in disposable plastic bottles, to reduce waste"; bulk purchasing '
            '"has less packaging"; mongoose/monkey/baboon/hyena proof bins for food '
            "waste; compost heap and worm farm for vegetable waste."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Only explicit 'plastic' reference in manual — disposable plastic bottles. "
            "Not coded separately: kitchen food-waste minimisation (operational detail "
            "bundled here)."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Sec. 5.2.3 Operations — Litter and visitor waste: advise tourists not to "
            "litter and adopt a 'pack it in, pack it out' policy in conservation areas "
            "with wildlife; educate clients before arrival on what can be disposed of "
            "on site — marine/coastal protected-area litter prevention aligned with "
            "Decree 45/2006 coastal pollution framework referenced in Sec. 2.1.5."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 5.2.3: "Advise tourists not to litter and instead adopt a \'pack it in, '
            "pack it out' policy – especially in conservation areas and where wildlife "
            'is around"; Sec. 2.1.5 references Decree 45/2006 on coastal/marine '
            "pollution prevention for protected areas."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Visitor litter/leakage prevention — critical for PA plastic pollution. "
            "Not coded: carbon-offset guidance (Box 3)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Secs. 5.2.3 and 5.3.2 — Concession contract compliance and local "
            "procurement: protected-area authority may provide waste-management "
            "guidance that concessionaires must follow under concession contracts; "
            "economic best practice includes procuring garbage removal and local "
            "services from communities — operational waste-service delivery in "
            "remote protected areas."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Sec. 5.2.3: "The PA may provide its own guidance on waste management, '
            'which you will need to follow in terms of your concession contract"; '
            "Sec. 5.3.2: local procurement examples include \"garbage removal\" "
            "among goods transport, furniture and consumables."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Contractual/procurement waste-services instrument — links PA management "
            "plans to operator obligations. Not coded: craft packaging guidance (Sec. 5.3)."
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
        "Tourism Concessions in Protected Areas in Mozambique", 0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Manual for Operators and Concessionaires — April 2012")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English title of: Concessões de Turismo nas Áreas Protegidas em Moçambique: "
        "Manual para Operadores e Concessionários (2012)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://www.anac.gov.mz/wp-content/uploads/2017/08/"
        "Manual-Concessoes-para-Sector-Privado_2012.pdf"
    )

    sections = [
        ("Publication Information", (
            "Program: Mozambique Support Program for Economic and Enterprise Development "
            "(SPEED), USAID/Mozambique.\n"
            "Contractor: DAI and Nathan Associates.\n"
            "Publication date: 23 April 2012 (SPEED-Reports-2012-005).\n"
            "Authors: Rita Casimiro and Dr Anna Spenceley.\n\n"
            "The manual guides private operators through tourism concession processes in "
            "Mozambique's protected areas, covering context, stakeholders, procurement, "
            "contracts, environmental/social/economic best practice, and annexes with "
            "templates and contacts."
        )),
        ("Section 2.1.5 — Legal and Environmental Framework", (
            "Tourism projects in protected areas are subject to the Environmental Law "
            "(Law 20/97) and Environmental Impact Assessment Regulation (Decree 45/2004) "
            "requiring a fully fledged EIA and management plan.\n\n"
            "The manual also references Decree 45/2006 on preventing pollution and "
            "protecting the coastal and marine environment for activities in coastal "
            "protected areas (e.g. Maputo Special Reserve, Quirimbas, Lake Niassa).\n\n"
            "MITUR/DNAC (now ANAC) manages protected areas and grants tourism concessions."
        )),
        ("Section 5.2 — Environmental Best Practice (plastic-relevant)", [
            ("5.2.2 Development stage", (
                "Concessionaires must develop a comprehensive environmental management plan "
                "(EMP) with waste management, water management and energy management "
                "plans to minimise waste production.\n\n"
                "Larger operations may adopt an EMS (e.g. ISO 14001) auditing packaging "
                "resources, food consumption and organic/inorganic waste production.\n\n"
                "Sustainable design includes careful siting of waste disposal areas in "
                "relation to the EIA, minimising construction material wastage, and "
                "preserving vegetation."
            )),
            ("5.2.3 Operations stage — Waste", (
                "Main operational environmental issues: energy, water, waste and biodiversity.\n\n"
                "Waste management guidance:\n"
                "• Reduce packaging through bulk purchasing and supplier engagement;\n"
                "• Use wildlife-proof bins for food waste in wildlife areas;\n"
                "• Use refillable dispensers instead of individual soap tablets;\n"
                "• Use refillable bottles rather than disposable plastic water bottles;\n"
                "• Recycle paper and printer cartridges;\n"
                "• Compost garden waste; worm farms for vegetable waste;\n"
                "• Educate visitors: 'pack it in, pack it out' — no littering in "
                "conservation areas;\n"
                "• Follow PA waste-management guidance in concession contracts.\n\n"
                "Wastewater: reed-bed systems and recycling for irrigation/toilets referenced."
            )),
            ("Business plan / EIA template (Table 4)", (
                "Feasibility study environmental impact section requires consideration of "
                "landscape, soil, air, water, biodiversity, waste (solid and liquid), energy, "
                "access and infrastructure footprint."
            )),
        ]),
        ("Section 5.3.2 — Economic Best Practice (waste-relevant)", (
            "Local procurement during construction and operations should include garbage "
            "removal among services sourced from local communities, alongside transport, "
            "furniture, uniforms and consumables — supporting waste-service delivery in "
            "remote protected areas."
        )),
        ("Note on Plastics and Solid Waste", (
            "This manual is voluntary operational guidance, not binding legislation.\n\n"
            "Explicit plastic reference: disposable plastic water bottles — operators "
            "advised to use refillable bottles instead.\n\n"
            "Other waste/plastic-relevant content:\n"
            "— EMP must include waste-management plans;\n"
            "— EMS covers packaging resources and inorganic waste;\n"
            "— Waste disposal area siting in EIA;\n"
            "— Packaging reduction through bulk purchasing;\n"
            "— 'Pack it in, pack it out' litter policy for tourists;\n"
            "— Garbage removal in local procurement;\n"
            "— Solid and liquid waste in business-plan EIA template;\n"
            "— Reference to Decree 45/2006 coastal/marine pollution regulation.\n\n"
            "No quantified national plastic-reduction targets. Binding waste rules for "
            "operators remain in concession contracts, EIA licences, Decree 94/2014, "
            "Decree 97/2020 and PA management plans."
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
        f"Instruments coded: {len(INSTRUMENTS)}\n"
        f"E — policy_target: {POLICY_FIELDS['policy_target']}\n"
        f"K — policy_circularity: {POLICY_FIELDS['policy_circularity']} "
        f"({POLICY_FIELDS['policy_lifecycle_phases_list']})\n"
        f"G — policy_type: {POLICY_FIELDS['policy_type']} (guidance manual)\n"
        f"See: 4P_Index_Mozambique_ANAC_Tourism_Concessions_Manual_2012.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR,
        "Mozambique_ANAC_Tourism_Concessions_Manual_2012_English_Translation.docx",
    )
    doc.save(path)
    return path


def create_index_html(basename):
    base = (
        "https://github.com/shivrain/Thinkswiss/raw/"
        "cursor/mozambique-anac-tourism-concessions-manual-2012-04b6/output/mozambique"
    )
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>ANAC Tourism Concessions Manual 2012 — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — ANAC Tourism Concessions Manual (2012)</h1>
  <p>Manual for Operators and Concessionaires. Generated {date.today().isoformat()}.
  Six instruments coded (waste, packaging, litter, plastic bottles).</p>
  <div class="card">
    <a class="download" href="{base}/{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{base}/{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="{base}/Mozambique_ANAC_Tourism_Concessions_Manual_2012_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_anac_tourism_concessions_manual_2012.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_ANAC_Tourism_Concessions_Manual_2012"
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
