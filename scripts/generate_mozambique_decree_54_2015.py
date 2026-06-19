#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Decree 54/2015."""

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
        "Regulamento da Avaliação do Impacto Ambiental (Decreto n.º 54/2015)"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz174566.pdf",
    "policy_year": 2015,
    "policy_objective": (
        "Regulate the environmental impact assessment process for public and private activities "
        "that may influence the environment; classify projects by impact category; require "
        "environmental studies, mitigation measures and three-stage environmental licensing "
        "before waste treatment, landfill, industrial and other high-impact facilities may operate."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (Regulamento) approved by Council of Ministers Decree n.º 54/2015 "
        "under Article 33 of the Environment Law (Law n.º 20/97) — sub-legislative instrument, "
        "not parliamentary legislation. Revokes Decrees 45/2004 and 42/2008."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "waste management, industry, municipalities, packaging, fisheries, tourism, water, agriculture"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, disposal, environmental leakage"
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        "Art. 27: graduated environmental licensing fees by study category (e.g. 30,000–500,000+ MT); "
        "Art. 28–29: administrative fines from 50,000 MT to 10,000,000+ MT; Art. 30: fee revenues "
        "50% State / 50% Environment Fund; fine revenues 60% State / 40% FUNAB."
    ),
    "policy_score": "[auto]",
}

# Six consolidated instruments — EIA gatekeeping framework; annex line-items not overcoded.
INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 4–5, Annexes I–IV: Classifies activities into categories A+, A, B and C requiring "
            "different levels of environmental assessment; Annexes list plastic-relevant activities "
            "including elastomer manufacturing, waste treatment/disposal, sanitary landfills "
            "(>150,000 t), incineration, and large industrial/port facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 4: categories A+, A, B, C tied to Annex lists; Art. 5: exemptions only for '
            'emergency activities approved by MICOA; Annex II 2.6: "Aterros sanitários com '
            'capacidade para mais de 150 000 toneladas"; Annex II 2.5.3: elastomer products.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Screening instrument — determines whether a project needs full EIA. Not coded per "
            "annex line-item; waste/landfill/incineration/elastomer entries consolidated here."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 8–12, 14–19: Requires pre-assessment (EAP), full EIA (categories A+/A) or "
            "Simplified Environmental Study — EAS (category B) before licensing; studies must assess "
            "waste generation, residual/cumulative impacts and alternatives; includes mandatory "
            "public participation and consultation for categories A+, A and B (Art. 15)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 8: EAP for A+/A, EAS for B; Art. 11: EIA report must cover "impactos residuais '
            'e cumulativos" and waste-related impacts; Art. 15: public participation mandatory '
            'for A+, A, B; Art. 19: processing deadlines by category.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Single procedural instrument for all study types and public participation. Not coded "
            "separately: expert review commissions (Art. 16–18), consultant registration (Art. 23–24)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Art. 7, 11, 21, 25: Approved projects must implement mitigation measures and an "
            "Environmental Management Plan (PGA); proponents must maintain compliance with license "
            "conditions, implement approved waste/impact controls, and bear civil/criminal liability "
            "for unauthorized operation or failure to apply mitigation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 7(f): licensing documentation includes "Plano de Gestão Ambiental"; Art. 11: '
            'EIA must define mitigation strategy and measures; Art. 25: proponent must implement '
            'mitigation, maintain records, and comply with reassessment/resettlement obligations.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Substantive outcome obligation — converts EIA findings into binding management duties "
            "for waste streams and pollution controls at licensed facilities."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 20–22: Three-stage environmental licensing — Provisional Licence (optional, "
            "after EAP approval), Installation Licence (after EIA/EAS approval and resettlement "
            "plan if required), Operation Licence (after installation compliance verified); "
            "operating without valid licence subject to fine (Art. 20(3))."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 20(1): "Licença Ambiental Provisória", "Licença Ambiental de Instalação", '
            '"Licença Ambiental de Operação"; Art. 20(3): operation without installation licence '
            'is fined; Art. 22: licence validity and renewal rules.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Core permit gate for landfills, waste plants and industrial facilities. Fee payment "
            "triggered after Installation Licence approval (per Art. 20 and energypedia summary)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 26: MICOA may inspect whether licensed activities comply with approved EIA, "
            "mitigation measures and licence conditions; all category A+ and A projects must undergo "
            "inspection verifying waste-management and environmental controls are implemented."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 26(1): inspection verifies "cumprimento das medidas de mitigação" and licence '
            'conditions; Art. 26(2): "Todos os projectos de categoria A + e A devem ser sujeitos '
            'a inspecção".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Post-approval compliance check. Complements Decree 51/2024 environmental inspection "
            "framework but is the EIA-specific inspection obligation in this regulation."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 27–30: Graduated licensing fees by category; administrative fines for operating "
            "without licence, failing mitigation obligations, false information, or consultant "
            "violations (Art. 28–29, up to millions of MT); fee revenue split 50/50 State and "
            "Environment Fund; fine revenue 60% State / 40% FUNAB."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 28: fines for "exploração da actividade sem observância" of Art. 25, unauthorised '
            'operation, or failure to implement mitigation; Art. 29: fine graduation by severity; '
            'Art. 30: "60% para o Fundo do Ambiente" (fines).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Single economic enforcement instrument; not duplicated per fine tier. Applies to "
            "underlying substantive violations including unlicensed waste facility operation."
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

    title = doc.add_heading("Regulation on Environmental Impact Assessment", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Decree No. 54/2015 of 31 December")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Regulamento da Avaliação do Impacto Ambiental\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz174566.pdf")

    sections = [
        ("Decree No. 54/2015", (
            "The Council of Ministers decrees:\n"
            "Article 1. The Regulation on the Environmental Impact Assessment Process is approved "
            "as an annex forming an integral part of this Decree.\n"
            "Article 2. The Minister responsible for the environment shall approve complementary "
            "norms for operationalising this Decree.\n"
            "Article 3. Decrees No. 45/2004 of 29 September and No. 42/2008 of 4 November are repealed.\n"
            "Article 4. This Decree enters into force ninety days after publication.\n"
            "Approved by the Council of Ministers on 30 November 2015."
        )),
        ("Regulation — Chapter I. General provisions", (
            "Article 1 (Object). To establish rules on the environmental impact assessment process.\n\n"
            "Article 2 (Scope). Applies to public and private activities that may directly or "
            "indirectly influence the environment.\n\n"
            "Article 3 (Principle). Environmental impact assessment is mandatory for activities "
            "listed in the annexes.\n\n"
            "Article 4 (Categories). Activities classified as:\n"
            "a) Category A+ — full EIA by independent Expert Reviewers;\n"
            "b) Category A — full EIA;\n"
            "c) Category B — Simplified Environmental Study (EAS);\n"
            "d) Category C — Preliminary Environmental Information only.\n\n"
            "Article 5. Emergency activities may be exempted subject to MICOA approval."
        )),
        ("Chapter II — Studies and public participation", (
            "Article 8. Pre-assessment (EAP) required for categories A+ and A; EAS for category B.\n\n"
            "Articles 10–12. Content requirements for EAP, full EIA and EAS including project "
            "description, alternatives, impact identification, residual/cumulative impacts, and "
            "mitigation measures.\n\n"
            "Article 15. Public participation and consultation mandatory for categories A+, A and B; "
            "proponent bears costs of public consultation.\n\n"
            "Articles 16–18. Technical commissions and expert review procedures for EIA reports.\n\n"
            "Article 19. Processing deadlines by category and study type."
        )),
        ("Chapter III — Environmental licensing", (
            "Article 20. Three-stage licensing:\n"
            "a) Provisional Environmental Licence (optional, after EAP approval);\n"
            "b) Installation Environmental Licence (after EIA/EAS approval and resettlement plan if required);\n"
            "c) Operation Environmental Licence (after verified installation compliance).\n"
            "Operating without installation licence is subject to administrative fine.\n\n"
            "Article 21. Environmental viability analysis before licence issuance.\n\n"
            "Article 25. Proponent obligations: implement mitigation, maintain records, comply with "
            "licence conditions, bear liability for unauthorised operation."
        )),
        ("Chapter V — Inspection, fees and sanctions", (
            "Article 26. MICOA inspects compliance with approved EIA and mitigation measures; "
            "all A+ and A projects must be inspected.\n\n"
            "Article 27. Graduated licensing fees by category (published in Boletim da República).\n\n"
            "Articles 28–29. Administrative fines for licence violations, false information, "
            "consultant misconduct — graduated by severity (50,000 MT to 10,000,000+ MT).\n\n"
            "Article 30. Fee revenues: 50% State, 50% Environment Fund; fine revenues: 60% State, 40% FUNAB."
        )),
        ("Annexes — plastic-relevant activities (selected)", (
            "Annex I (A+): large-scale projects including major industrial and biodiversity-sensitive developments.\n\n"
            "Annex II (A): includes waste treatment and disposal, sanitary landfills (>150,000 tonnes), "
            "incineration plants, elastomer product manufacturing, large port/industrial facilities.\n\n"
            "Annex III (B): smaller waste treatment, recycling, water/sanitation systems, food processing.\n\n"
            "Annex IV (C): low-impact activities requiring preliminary environmental information only.\n\n"
            "Note: No explicit mention of plastics; relevance is through EIA/licensing requirements "
            "for waste, landfill, incineration and elastomer/packaging-related industries."
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
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — EIA gatekeeping framework)\n"
        f"See: 4P_Index_Mozambique_Decree_54_2015.xlsx",
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Decree_54_2015_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Decree 54/2015 — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Decree 54/2015 (Environmental Impact Assessment)</h1>
  <p>Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Decree_54_2015_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_decree_54_2015.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Decree_54_2015"
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
