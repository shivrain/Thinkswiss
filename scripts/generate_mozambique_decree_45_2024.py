#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Decree 45/2024."""

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
        "Regulamento sobre o Processo de Auditoria Ambiental (Decreto n.º 45/2024)"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz227797.pdf",
    "policy_year": 2024,
    "policy_objective": (
        "Establish principles and rules for environmental audit of public and private "
        "activities across implementation, deactivation and restoration phases; verify "
        "compliance with environmental legislation, Environmental Management Plans, "
        "monitoring performance and corrective action on waste and pollution controls."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (Regulamento) approved by Council of Ministers Decree n.º 45/2024 "
        "under Articles 18 and 33 of the Environment Law (Law n.º 20/97) — sub-legislative "
        "instrument, not parliamentary legislation. Revokes Decree n.º 25/2011."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "waste management, industry, municipalities, packaging, fisheries, tourism, water, agriculture"
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": (
        "production, disposal, environmental leakage"
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        "Art. 14: auditor registration fees 8,250–55,000 MT; Art. 15–16: fines 50,000–15,000,000 MT "
        "by category/violation; Art. 18: fee revenues 60% State / 40% environment sector; "
        "fine revenues 40% State / 60% environment sector."
    ),
    "policy_score": "[auto]",
}

# Five consolidated instruments — audit compliance framework; not overcoded per article.
INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 4–7, 5: Mandates environmental audit (public and private) at least once "
            "per year for all EIA categories A+, A, B and C; audit scope includes legal "
            "compliance, Environmental Management Plan and counterbalance plan implementation, "
            "environmental performance, monitoring reports, prior audits and action plans."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 7(4): "auditorias ambientais... são realizadas pelo menos uma vez por ano"; '
            'Art. 5(c)-(h): PGA, Plano de Contrabalanço, desempenho, monitorização, Planos de Acção; '
            'Art. 6: MICOA registry, validation, conformity opinions for licence renewal.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Horizontal compliance multiplier for waste/pollution laws via PGA and monitoring "
            "verification. Not coded separately: auditor registration (Art. 10–11), cost rules "
            "(Art. 9), MICOA competency list (Art. 6) beyond audit mandate."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 8, 13(1)–(2): Auditors must issue conformity reports with mandatory "
            "recommendations; audited entities must prepare and submit an Action Plan within "
            "30 business days addressing audit findings; annual audit report submitted to MICOA "
            "in physical and electronic format; non-compliance with recommendations is sanctioned."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 8(6): recommendations are "cumprimento obrigatório"; Art. 8(7): Action Plan '
            'within "30 dias úteis"; Art. 13(1): minimum one annual audit report; Art. 13(2): '
            'Action Plan submission within 30 days of report receipt.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Core corrective-action instrument — converts audit findings into binding remediation "
            "for waste management and pollution non-conformities identified in PGA/monitoring."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 13(3)–(5): Proponents must submit semi-annual Environmental Performance Reports "
            "(or as set in licence), annual Environmental Monitoring Reports, and monthly "
            "combined performance/monitoring reports during exploration/prospecting phases."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 13(3): "Semestralmente" performance reports; Art. 13(4): annual monitoring '
            'report; Art. 13(5): monthly reports in prospection/research phase; Art. 13(6): '
            'non-submission penalised under this Regulation.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Ongoing disclosure obligation supporting detection of plastic/waste leakage from "
            "licensed operations. Bundled reporting frequencies — not coded per report type."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 12: Audited entities must facilitate audits, provide Environmental Management "
            "Plans, performance and monitoring reports, prior audit reports and action plans; "
            "private auditors must notify MICOA 15 days before auditing and submit reports "
            "within 15 days after completion."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.0,
        "instrument_implementation_text": (
            'Art. 12(c): provide "Plano de Gestão Ambiental, Relatórios de Desempenho, '
            'Relatórios de Monitorização"; Art. 12(g): monitoring and private audit reports '
            'within "quinze dias"; Art. 12(h): 15-day advance notice to MICOA.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Cooperation duty enabling verification of waste-management compliance. Fines for "
            "non-cooperation in Art. 15(8)–(10) coded under sanctions instrument."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 14–18: Auditor registration fees (11,000–55,000 MT); graduated fines for "
            "obstructing public audit (200,000–3,000,000 MT by category), unregistered private "
            "auditing, failure to submit audit/action/monitoring reports (50,000–500,000 MT), "
            "and non-compliance with audit recommendations (500,000–15,000,000 MT by category)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 15(1): obstruction fines A+ 3,000,000 MT to C 200,000 MT; Art. 16: '
            'recommendation non-compliance A+ 15,000,000 MT, A 5,000,000 MT, B 800,000 MT, '
            'C 500,000 MT; Art. 17: aggravating/attenuating circumstances.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Single economic enforcement instrument; not duplicated per fine tier. Applies when "
            "entities fail audit-derived waste/pollution obligations."
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

    title = doc.add_heading("Regulation on the Environmental Audit Process", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Decree No. 45/2024 of 26 June")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Regulamento sobre o Processo de Auditoria Ambiental\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz227797.pdf")

    sections = [
        ("Decree No. 45/2024", (
            "The Council of Ministers decrees:\n"
            "Article 1. The Regulation on the Environmental Audit Process is approved as an annex "
            "forming an integral part of this Decree.\n"
            "Article 2. The Minister responsible for the environment shall approve general and "
            "specific directives on environmental audit and other implementation norms.\n"
            "Article 3. Decree No. 25/2011 of 15 June is repealed.\n"
            "Article 4. This Regulation enters into force ninety days after publication.\n"
            "Approved by the Council of Ministers on 7 May 2024."
        )),
        ("Regulation — General provisions", (
            "Article 1 (Object). To establish principles and rules governing environmental audit.\n\n"
            "Article 2 (Scope). Applies to public and private activities in implementation, "
            "deactivation and restoration phases that may affect environmental components.\n\n"
            "Article 4. Environmental audit is a systematic management instrument evaluating "
            "environmental control and protection processes.\n\n"
            "Article 5 (Audit focus). Includes:\n"
            "a) compliance with environmental legislation;\n"
            "b) mitigation measure implementation;\n"
            "c) Environmental Management Plan implementation;\n"
            "d) counterbalance plan implementation;\n"
            "e) environmental performance;\n"
            "f) performance and monitoring reports;\n"
            "g) prior public and private audit reports; and\n"
            "h) action plan implementation."
        )),
        ("Audit types and frequency", (
            "Article 7.\n"
            "1. Two types: public audit (by environment sector) and private audit (by registered "
            "auditors independent of the project's EIA consultant).\n"
            "4. Public and private audits for categories A+, A, B and C at least once per year.\n"
            "5. A+ and A public audits conducted by central authority.\n"
            "6. B and C audits conducted by provincial authority.\n\n"
            "Article 8. Audit reports must assess legal conformity, project impacts, mitigation "
            "implementation and performance. Recommendations are mandatory for audited entities. "
            "Action Plan required within 30 business days. Reports retained minimum 10 years."
        )),
        ("Reporting obligations (Article 13)", (
            "1. Annually: at least one Environmental Audit Report to MICOA.\n"
            "2. Within 30 days of audit report: Action Plan in response.\n"
            "3. Semi-annually (or as per licence): Environmental Performance Report.\n"
            "4. Annually: Environmental Monitoring Report.\n"
            "5. During exploration/prospecting: monthly combined performance and monitoring report.\n"
            "6. Non-compliance subject to fines under this Regulation."
        )),
        ("Sanctions and fees (selected)", (
            "Article 14. Auditor registration fees: individual 11,000 MT; collective 55,000 MT.\n\n"
            "Article 15. Fines for obstructing public audit: A+ 3,000,000 MT; A 800,000 MT; "
            "B 500,000 MT; C 200,000 MT. Other fines for unregistered auditing, late/missing "
            "reports and cooperation failures (50,000–800,000 MT).\n\n"
            "Article 16. Fines for non-compliance with audit recommendations: A+ 15,000,000 MT; "
            "A 5,000,000 MT; B 800,000 MT; C 500,000 MT.\n\n"
            "Article 18. Fee revenues: 60% State, 40% environment sector; fine revenues: "
            "40% State, 60% environment sector."
        )),
        ("Note on plastic pollution relevance", (
            "Glossary defines pollutant as including chemical products and waste (resíduos). "
            "No explicit plastic provisions — relevance is through mandatory audit and monitoring "
            "of Environmental Management Plans and legal compliance for licensed waste/industrial "
            "facilities (links to Decrees 54/2015 EIA and 51/2024 inspection)."
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
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — audit compliance framework)\n"
        f"See: 4P_Index_Mozambique_Decree_45_2024.xlsx",
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Decree_45_2024_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Decree 45/2024 — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Decree 45/2024 (Environmental Audit)</h1>
  <p>Generated {date.today().isoformat()}. Five consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Decree_45_2024_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_decree_45_2024.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Decree_45_2024"
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
