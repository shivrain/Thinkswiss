#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 13/2021 (Health Policy)."""

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
        "Política de Saúde e Estratégia de Implementação "
        "(Resolução n.º 13/2021)"
    ),
    "policy_url": (
        "https://sisma.misau.gov.mz/docs/Politica%20Nacional%20de%20Saude.pdf"
    ),
    "policy_year": 2021,
    "policy_objective": (
        "Approve the National Health Policy and its Implementation Strategy to align "
        "Mozambique's health-sector normative framework with current socio-economic realities "
        "and social determinants of health, promoting universal health coverage and SDG "
        "alignment. Plastic-relevant scope is limited to Pilar 3 (Health in All Policies): "
        "the policy recognises urbanisation pressures on solid waste management and "
        "environmental sanitation as health determinants, and promotes intersectoral "
        "coordination on water, sanitation and environmental risk reduction — solid waste "
        "includes plastic fractions but plastics are not named explicitly."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "Council of Ministers Resolution n.º 13/2021 (approved 30 March 2021, published "
        "16 April 2021, Boletim da República Series I No. 72) under Constitution Art. 203(1)(f) — "
        "national health policy and implementation strategy; revokes Resolution 4/95. "
        "Strategic planning instrument for MISAU, not binding sectoral waste or plastic "
        "regulation. Operational waste rules remain in Decree 94/2014, Decree 16/2015, etc."
    ),
    "policy_integration": 0.25,
    "policy_sectors_list": "municipalities, water, waste management",
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
            "Pilar 3, Strategy 3.1.2 — Increase Access to Basic Sanitation: recognises "
            "accelerated urban population growth increasing pressure on solid waste management "
            "and environmental sanitation; promotes expanded access to basic sanitation "
            "infrastructure as a social determinant of health. Solid waste (resíduos sólidos) "
            "includes plastic waste fractions though plastics are not named explicitly."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Pilar 3, 3.1.2: "aumentando a pressão sobre a gestão dos resíduos sólidos e '
            'saneamento do meio no geral"; policy promotes increased access to basic '
            "sanitation infrastructure — qualitative strategic commitment only; no "
            "quantified waste-collection, landfill or plastic-reduction targets; "
            "implementation delegated to intersectoral coordination and municipal actors."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Only explicit solid-waste reference in the policy — primary plastic-relevant "
            "instrument. Not coded separately: potable water access (3.1.1) or housing "
            "(3.1.3) strategies without waste linkage."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pilar 3, Strategy 3.1.4 — Reduce Negative Impacts of Environmental Risks: "
            "promotes prevention, management and mitigation of environmental health risks; "
            "notes industrialisation accompanied by environmental impacts; addresses "
            "cyclone/flood vulnerability affecting communities and health-service networks. "
            "Indirect relevance to plastic pollution via general environmental-risk and "
            "industrial-pollution framing."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Pilar 3, 3.1.4: "promover intervenções orientadas para prevenção, gestão e '
            'mitigação dos impactos negativos dos riscos ambientais"; notes industrialisation '
            'is "acompanhada de impactos sobre o meio ambiente" — strategic prevention '
            "framework without plastic-specific measures or enforcement mechanisms."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "General environmental-risk instrument — no explicit waste or plastic mention. "
            "Coded because industrial environmental impacts may include plastic pollution "
            "pathways. Not coded: tobacco/alcohol/drug strategies (Pilar 1)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Pilar 3 — Health in All Policies and National Commission on Social Determinants "
            "of Health (CN-DSS): formalises intersectoral collaboration on social "
            "determinants including water, sanitation, housing, food, employment and "
            "environment; third commission level promotes multissectoral policies improving "
            "material living conditions — enabling coordinated action on solid waste and "
            "sanitation as health determinants across MISAU, municipalities and partner "
            "ministries."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Pilar 3 objective: "Saúde em Todas as Políticas" — determinants include '
            '"ambiente socio-económico" and "ambiente físico"; CN-DSS third level: '
            '"água, saneamento, habitação, alimentação saudável, emprego, ambiente, '
            'serviços de saúde e de educação de qualidade"; implementation through '
            "interministerial programmatic coordination and strategic-plan sharing."
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Governance/coordination instrument for sanitation and waste as health "
            "determinants — no operational waste-management rules. Not coded separately: "
            "16 health-service strengthening strategies (Pilar 2) or medical-products "
            "strategy (1.1.10) without waste-disposal provisions."
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
        "National Health Policy and Implementation Strategy", 0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 13/2021")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Política de Saúde e Estratégia de Implementação "
        "(Resolução n.º 13/2021)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://sisma.misau.gov.mz/docs/Politica%20Nacional%20de%20Saude.pdf"
    )

    sections = [
        ("Resolution No. 13/2021 — Enacting Provisions", (
            "Whereas it is necessary to adjust the normative framework of the health sector to "
            "the current socio-economic reality of the country and to the approach of social "
            "determinants of health, in particular alignment with the Sustainable Development "
            "Goals, providing it with principles, objectives and guidance for adequate health "
            "management, it is urgent to approve the Health Policy.\n\n"
            "In these terms and in exercise of the powers established in paragraph (f) of "
            "paragraph 1 of Article 203 of the Constitution of the Republic of Mozambique, "
            "the Council of Ministers determines:\n\n"
            "Article 1. The Health Policy and the Strategy for its Implementation, attached "
            "hereto, are approved as an integral part of this Resolution.\n\n"
            "Article 2. Resolution No. 4/95 of 11 July is repealed.\n\n"
            "Approved by the Council of Ministers on 30 March 2021. Published in the Official "
            "Gazette (Boletim da República), Series I, No. 72, 16 April 2021.\n\n"
            "Signed: Carlos Agostinho do Rosário, Prime Minister."
        )),
        ("I. Health Policy — Context", [
            ("1.1 Population profile", (
                "According to the 2017 population census, Mozambique's total population is "
                "27,909,798 inhabitants — 52% female, 48% male. Age structure: 0–14 years "
                "26.6%, 15–64 years 50.1%, 65+ years 3.3%. Gross birth rate 37.9%; total "
                "fertility rate 5.2 children per woman (15–49 years)."
            )),
            ("1.2 Socio-economic factors", (
                "With GDP per capita of USD 1,328 in 2018 (World Development Indicators), "
                "Mozambique remains among low-income countries; approximately 62% of the "
                "population lived in poverty (less than USD 1.90/day) in 2014. Government "
                "health expenditure was estimated at 3% of GDP in 2016 (Global Health Observatory)."
            )),
            ("1.3 Population health status", (
                "Life expectancy at birth was 60 years (2016). Maternal mortality 489/100,000 "
                "live births; neonatal mortality 27/1,000 live births. Main causes of death: "
                "communicable and other diseases (65%), non-communicable diseases (27%), "
                "injuries (8%)."
            )),
            ("1.4 Social and health determinants", (
                "IMASIDA 2015 results show high poverty (70%), chronic malnutrition (43%), "
                "low schooling especially among women (29%), weak access to potable water "
                "(63%, especially rural), and deficient environmental sanitation (27%)."
            )),
            ("1.5 Health service coverage and emerging challenges", (
                "Vaccine-preventable and childhood disease burdens have declined over two "
                "decades, but HIV/AIDS, malaria and tuberculosis remain severe; non-communicable "
                "diseases (hypertension, diabetes, cancers), trauma, alcohol, tobacco and drug "
                "abuse are rising. Objectives must align with universal health coverage and SDGs."
            )),
            ("1.6 Opportunities and challenges", (
                "Macroeconomic constraints on health financing (Abuja recommendations), "
                "decentralisation impacts on universal access, and system resilience to shocks, "
                "disasters and public-health emergencies without compromising essential services "
                "are key challenges."
            )),
        ]),
        ("2. General Policy Provisions", [
            ("2.1–2.3 Vision, mission and values", (
                'Vision: "A healthy population contributing to the development of Mozambique."\n\n'
                'Mission: "Work to improve health conditions for all Mozambicans through a system '
                'that recognises, empowers and coordinates all stakeholders."\n\n'
                "Values: solidarity; professional ethics and deontology; respect for diversity; "
                "humanisation; professionalism; commitment."
            )),
            ("2.4 Guiding principles (summary)", (
                "Multi-sectorality — public policies across sectors impact population health; "
                "social participation; partnerships with civil society, industry and development "
                "partners; decentralised governance; equity; comprehensiveness of care; "
                "accountability through legal instruments; inclusion without discrimination."
            )),
            ("2.5 Policy objectives", (
                "General objective: Promote improvement of living conditions and lifestyles, "
                "strengthening measures for individual and collective health and reducing "
                "vulnerabilities and health risks from social determinants.\n\n"
                "Specific objectives: stimulate health promotion as part of integrated National "
                "Health System care; promote equitable health practices; offer better services "
                "at all levels; promote health-promotion education and training; stimulate "
                "research and innovation; articulate public policies with national and "
                "international health agendas."
            )),
        ]),
        ("II. Implementation Strategy — Three Pillars", [
            ("Pilar 1 — Well-being and healthy lifestyles", (
                "Strategic objective: Promote well-being and adoption of healthy lifestyles.\n\n"
                "Strategy 1 — Increase food availability and dietary diversification (collaboration "
                "with agriculture and relevant sectors; nutritional education; micronutrient "
                "deficiency protection).\n\n"
                "Strategy 2 — Combat tobacco and derivatives, abusive alcohol and other drugs "
                "(especially among youth).\n\n"
                "Strategy 3 — Encourage safe and responsible sexual behaviour (reproductive health, "
                "unintended pregnancies, STIs including HIV)."
            )),
            ("Pilar 2 — Access and use of health services (summary)", (
                "Strategic objective: Strengthen the National Health System focusing on access, "
                "utilisation and quality.\n\n"
                "The National Health System comprises three subsystems: (a) Public Health "
                "Subsystem (National Health Service); (b) Community Health Subsystem; "
                "(c) Private Health Subsystem (for-profit and non-profit).\n\n"
                "Nineteen strategies cover: system resilience; public-private-community "
                "partnerships; community participation; lifecycle-oriented service delivery; "
                "health promotion and prevention; disease detection and control; quality and "
                "access improvement; decentralisation; emergency medical systems; rational use "
                "of quality medicines and medical products; health infrastructure; health "
                "technologies; human resources; financial resource efficiency; ICT; health "
                "information systems; research; traditional/alternative medicine coordination; "
                "and legislation updating."
            )),
            ("Pilar 3 — Health in All Policies (plastic-relevant section)", (
                "Strategic objective: Ensure implementation of the Health in All Policies "
                "approach — social determinants of health.\n\n"
                "Health problems are influenced by socio-economic environment, health system, "
                "physical environment and biological attributes outside the health sector, "
                "requiring responsive intersectoral collaboration mechanisms.\n\n"
                "Strategy 3.1.1 — Increase access to potable water (49% of households had "
                "access in 2017 census, up from 35% in 2007; ~half the population still lacks "
                "potable water).\n\n"
                "Strategy 3.1.2 — Increase access to basic sanitation: "
                '"Mozambique has witnessed accelerated population growth and greater attraction '
                'to urban areas, increasing pressure on solid waste management and environmental '
                'sanitation in general." The policy promotes expanded access to basic sanitation '
                "infrastructure.\n\n"
                "Strategy 3.1.3 — Increase access to adequate housing (housing materials linked "
                "to respiratory diseases; 2017 census: ~47% palhota-type dwellings, 5.8% "
                "conventional houses).\n\n"
                "Strategy 3.1.4 — Reduce negative impacts of environmental risks: promote "
                "prevention, management and mitigation of environmental health risks; Mozambique "
                "is vulnerable to cyclones and floods; industrialisation is accompanied by "
                "environmental impacts."
            )),
        ]),
        ("III. Implementation Mechanisms", (
            "Implementation is coordinated among:\n"
            "a) National Commission on Social Determinants of Health (CN-DSS);\n"
            "b) Ministry of Health through the National Health System;\n"
            "c) Provincial and district State representation bodies;\n"
            "d) Local authority and decentralised governance bodies;\n"
            "e) Health partners.\n\n"
            "The CN-DSS coordinates multisectoral response to reduce socio-economic inequalities "
            "and improve population health across four action levels:\n"
            "— Level 1: comprehensive policies promoting healthy lifestyles;\n"
            "— Level 2: community and civil-society involvement (including vulnerable groups);\n"
            "— Level 3: multissectoral policies improving material and psychosocial living "
            "conditions (water, sanitation, housing, healthy food, employment, environment, "
            "quality health and education services);\n"
            "— Level 4: macroeconomic, labour-market, environmental-protection and peace "
            "policies promoting sustainable development and reducing inequalities.\n\n"
            "MISAU will create an institutional policy-management mechanism for operationalisation "
            "and strategic-plan alignment."
        )),
        ("IV–V. Communication, Monitoring and Evaluation", (
            "Communication is a main implementation component to motivate institutions and the "
            "population toward policy goals and constitutional duty to promote personal health "
            "through healthy lifestyles. A central communication plan will prioritise mass "
            "dissemination at all levels.\n\n"
            "Monitoring and evaluation will be routine and continuous (quarterly, semi-annual, "
            "annual) using a national indicator framework. MISAU is responsible with "
            "intersectoral collaboration as integral to the determinants-based health approach."
        )),
        ("6. Note on Plastics and Solid Waste", (
            "This health policy contains no explicit reference to plastics (plástico). "
            "Plastic-relevant content is limited to:\n\n"
            "— One mention of solid waste management (gestão dos resíduos sólidos) as an "
            "urbanisation pressure alongside environmental sanitation (Pilar 3, Strategy 3.1.2);\n"
            "— General environmental-risk reduction including industrial environmental impacts "
            "(Strategy 3.1.4);\n"
            "— Intersectoral Health-in-All-Policies framework covering water, sanitation and "
            "environment as health determinants (Pilar 3; CN-DSS Level 3).\n\n"
            "The policy sets no quantified plastic-reduction, waste-collection, recycling or "
            "EPR targets and allocates no waste/plastic-specific budget. Strategy 1.1.14 on "
            "health financing refers to general resource mobilisation, not waste management. "
            "Medical products and medications (Strategy 1.1.10) are addressed for rational use "
            "and local production but not for medical-plastic or healthcare-waste disposal.\n\n"
            "Operational plastic and solid-waste measures remain in Decree 94/2014 (Urban Solid "
            "Waste), Decree 16/2015 (Plastic Bags), Decree 97/2020 (Coastal Zones), EDEA "
            "(Resolution 53/2024) and related instruments."
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
        f"See: 4P_Index_Mozambique_Resolution_13_2021_Health_Policy.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Resolution_13_2021_Health_Policy_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 13/2021 Health Policy — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 13/2021 (National Health Policy)</h1>
  <p>National Health Policy and Implementation Strategy. Generated {date.today().isoformat()}.
  Three instruments coded (solid waste / sanitation / environmental determinants only).</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Resolution_13_2021_Health_Policy_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_13_2021_health_policy.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_13_2021_Health_Policy"
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
