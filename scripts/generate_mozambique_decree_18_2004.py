#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Decree 18/2004 (as amended 67/2010)."""

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
        "Regulamento sobre Padrões de Qualidade Ambiental e de Emissão de Efluentes "
        "(Decreto n.º 18/2004, alterado pelo Decreto n.º 67/2010)"
    ),
    "policy_url": (
        "https://biblioteca.biofund.org.mz/biblioteca_virtual/"
        "decreto-n-o-182004-de-2-de-junho-regulamento-sobre-padroes-de-qualidade-ambiental-e-de-emissao-de-efluentes/"
    ),
    "policy_year": 2010,
    "policy_objective": (
        "Establish national environmental quality standards and effluent/emission limits for air, water "
        "and industrial discharges to control pollutant concentrations, including standards and prohibitions "
        "relevant to waste burning, incineration, marine discharges, and industrial emissions affecting "
        "plastic and other waste streams."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (Regulamento) approved by Council of Ministers under Decree n.º 18/2004 "
        "and amended by Decree n.º 67/2010, issued under Article 33 of the Environment Law — "
        "sub-legislative instrument, not parliamentary legislation."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "industry, waste management, water, chemicals, municipalities, agriculture, fisheries"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'Art. 23 (as amended): special authorization fees between 50,000.00 MT and 500,000.00 MT; '
        '"40% para o Fundo do Ambiente"; Art. 24: administrative fines between 1,000,000.00 MT and '
        '10,000,000.00 MT graded by EIA activity category.'
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 3: Regulation applies to all public and private activities that directly or indirectly "
            "may interfere with environmental components (air, water, soil), covering industrial and "
            "waste-related emission sources."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 3 (per Decree 18/2004, cited in national EIA/legal summaries): broad application to '
            'activities affecting environmental components; linked to environmental licensing and inspection framework.'
        ),
        "instrument_score": "[auto]",
        "comments": "Scope instrument enabling application to plastic manufacturing and waste-disposal activities.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 7 and Annex I (as amended by Decree 67/2010): National ambient air quality standards "
            "for SO₂, CO, O₃, particulate matter and related parameters with primary/secondary limits "
            "across defined averaging periods."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Annex I (Decree 67/2010): numeric concentration limits (µg/m³) for air pollutants including '
            'particulates; Art. 24 fines for violations; special authorization regime Art. 22-23.'
        ),
        "instrument_score": "[auto]",
        "comments": "Particulate standards relevant to open burning and incineration of waste including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 8: Establishes atmospheric pollutant emission standards per industrial establishment, "
            "limiting releases from stationary industrial sources."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 8 (Decree 18/2004): emission standards by industrial establishment; Art. 10 chimney '
            'discharge rules; Art. 24 fines; authorization fees Art. 23.'
        ),
        "instrument_score": "[auto]",
        "comments": "Applies to industrial facilities including plastic production and processing plants.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 10: Requires appropriate mechanisms for atmospheric pollutant discharges through chimneys/stacks."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 10 (Decree 18/2004): mechanisms for atmospheric discharges via appropriate chimneys; '
            'enforcement via general penalty framework Art. 24.'
        ),
        "instrument_score": "[auto]",
        "comments": "Controls point-source air emissions from industrial stacks.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Annex V (air quality section, Decree 67/2010): Prohibits open-air burning of solid waste, "
            "liquid waste, or any other combustible material that causes environmental degradation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"é proibida a queima ao ar livre de resíduos sólidos, líquidos ou de qualquer outro material '
            'combustível, desde que cause degradação ambiental"; Art. 24: fines 1–10 million MT.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Directly applicable to open burning of plastic and mixed solid waste. Mandatory prohibition language."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Annex V: Prohibits installation and operation of domestic and industrial incinerators, "
            "except hospital incinerators."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"proíbe-se a instalação e o funcionamento de incineradores domiciliares e industriais, '
            'excepto os hospitalares"; Art. 24: graduated fines by EIA category.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Major end-of-life instrument for plastic waste thermal treatment; hospital exemption reduces "
            "unconditionality score. Cross-reference Decree 94/2014 for municipal waste incineration rules."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Annex V: Authorises/requirement power to mandate automatic equipment measuring quantities "
            "and qualities of emitted pollutants when necessary."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"Em caso necessário, poderá ser exigida a instalação e operação de equipamentos automáticos '
            'para medição das quantidades e qualidades dos poluentes emitidos."'
        ),
        "instrument_score": "[auto]",
        "comments": "Enabling power ('poderá ser exigida'); scored in_force=0 per coding Rule 12.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Annex V: Prohibits emission of odorous substances in quantities perceptible beyond the "
            "emitting property boundary; detection by credentialed agents."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Fica proibida a emissão de substâncias odoríferas na atmosfera em quantidades que possam '
            'ser perceptíveis fora dos limites da área de propriedade da fonte emissora"; credentialed agents.'
        ),
        "instrument_score": "[auto]",
        "comments": "Annex IB lists odorous substance thresholds including chemicals used in plastic production.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Annex IA (Decree 67/2010): Air quality limits for inorganic and organic carcinogenic pollutants "
            "including benzene, formaldehyde, styrene, toluene, tetrachloroethylene and related substances."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Annex IA: numeric limits (µg/m³) for lead, mercury, arsenic, benzene, formaldehyde, styrene, '
            'toluene, tetrachloroethylene, etc.; Art. 24 enforcement.'
        ),
        "instrument_score": "[auto]",
        "comments": "Styrene and toluene are directly relevant to plastics/polymers manufacturing emissions.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 16 and Annex V: Standards for liquid industrial effluent quality and emissions before "
            "final discharge to receiving bodies; marine/ocean receptor criteria require near-absence of "
            "floating materials, oils, deposits, and artificial colorants."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 16 (Decree 18/2004): industrial effluent standards before receptor discharge; '
            'Annex V marine criteria: "Materias flutuantes: virtualmente ausentes"; chemical tables with '
            'min/max limits (mg/l) for metals and organic substances.'
        ),
        "instrument_score": "[auto]",
        "comments": "Marine floating materials standard directly relevant to plastic leakage to ocean.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex V Tables I and IA: Maximum and minimum concentration limits for potentially harmful "
            "chemical substances and pesticides in receiving waters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Annex V: detailed mg/l limits for aluminium, cadmium, lead, mercury, phenols, surfactants, '
            'organochlorine pesticides, etc.; industrial/agricultural discharges permitted only if criteria met.'
        ),
        "instrument_score": "[auto]",
        "comments": "Includes surfactants and phenols relevant to plastic-related industrial effluents.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 22-23 (as amended): Special authorization required for emissions exceeding standards; "
            "polluter pays authorization fee of 50,000–500,000 MT."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 23(1): "é devida uma taxa a ser paga pelo poluidor, num valor compreendido entre '
            '50 000,00 MT e 500 000,00 MT"; Art. 23(3): 40% to Environment Fund.'
        ),
        "instrument_score": "[auto]",
        "comments": "Regulatory licensing with economic fee; also scored 0.60 economic aspect in comments.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 24 (as amended): Administrative fines of 1,000,000–10,000,000 MT for violations, "
            "graduated by EIA activity category (A/B/C)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 24(1): fines "entre 1 000 000,00 MT e 10 000 000,00 MT"; Art. 24(2): graduated by '
            'EIA categories; Art. 3: payment via competent tax directorate.'
        ),
        "instrument_score": "[auto]",
        "comments": "Primary enforcement instrument for all standards including waste burning and incineration bans.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 20: Noise level standards to be established by MICOA considering emission source, "
            "after consultation with sector authorities."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Art. 20 (Decree 18/2004): "os níveis de ruídos serão estabelecidos tendo em conta a fonte '
            'emissora do ruído, cujos padrões e limites o MICOA estabelecerá, em legislação especial".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Enabling power — secondary legislation on noise limits reportedly not published as of major EIA "
            "reviews; in_force=0."
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


def create_translation_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading(
        "Regulation on Environmental Quality and Effluent Standards",
        0,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph(
        "Decree No. 18/2004, as amended by Decree No. 67/2010"
    )
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation\n"
        "Regulamento sobre Padrões de Qualidade Ambiental e de Emissão de Efluentes\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}")

    doc.add_heading("Decree No. 67/2010 (Amendment)", level=1)
    doc.add_paragraph(
        "of 31 December\n\n"
        "Given the need to review and update environmental quality standards and revise applicable "
        "fees and fines, pursuant to Article 33 of the Environment Law combined with Article 6 of "
        "Decree No. 18/2004 of 2 June, the Council of Ministers decrees:"
    )

    sections = [
        ("Article 1", (
            "1. Annexes I and V referred to in Article 7 and paragraph 3 of Article 16 of the Regulation "
            "on Environmental Quality and Effluent Emission Standards, approved by Decree No. 18/2004 of "
            "2 June, are amended and replaced by Annexes I and V to this Decree.\n"
            "2. Annexes IA and IB to this Decree are approved and incorporated into the Regulation on "
            "Environmental Quality and Effluent Emission Standards."
        )),
        ("Article 2 — Articles 23 and 24 (as amended)", (
            "Article 23 (Special authorization issuance fees)\n"
            "1. For issuance of the authorization provided in paragraph 2 of Article 22 of the Regulation "
            "on Environmental Quality and Effluent Emission Standards, a fee shall be paid by the polluter "
            "in an amount between 50,000.00 MT and 500,000.00 MT.\n"
            "3. Revenue from fees collected under this Regulation shall be allocated as follows:\n"
            "   a) 60% to the State Budget;\n"
            "   b) 40% to the Environment Fund.\n\n"
            "Article 24 (Violations and fines)\n"
            "1. Without prejudice to other sanctions under applicable law, the following constitute "
            "violations punishable by fines between 1,000,000.00 MT and 10,000,000.00 MT:\n"
            "   [subparagraphs a–c as in original decree]\n"
            "2. Fines shall be graduated as follows:\n"
            "   a) 1,000,000.00–2,000,000.00 MT for Category C activities under the EIA Regulation;\n"
            "   b) 2,000,000.00–5,000,000.00 MT for Category B activities;\n"
            "   c) 5,000,000.00–10,000,000.00 MT for Category A activities."
        )),
        ("Article 3 (Payment of fees and fines)", (
            "Revenues collected under this Regulation shall be paid to the competent Tax Directorate "
            "using the appropriate official form."
        )),
        ("Article 4 (Updating fees and fines)", (
            "The Ministers overseeing the Environment and Finance sectors shall update fee and fine "
            "amounts provided in this Regulation."
        )),
        ("Decree No. 18/2004 — Key provisions (base regulation, summary)", (
            "Approved 2 June 2004. Implements Article 10 of the Environment Law (Law No. 20/97).\n\n"
            "Article 3 — Scope: Applies to all public and private activities that directly or indirectly "
            "may interfere with environmental components.\n\n"
            "Article 7 — Sets parameters for maintenance of air quality.\n\n"
            "Article 8 — Provides atmospheric pollutant emission standards per industrial establishment.\n\n"
            "Article 10 — Mechanisms for atmospheric pollutant discharges through appropriate chimneys.\n\n"
            "Article 16 — Establishes emission standards for liquid industrial effluent discharges and "
            "quality requirements before final discharge to the receiving environment.\n\n"
            "Article 20 — Noise levels shall be established by MICOA considering the noise source, "
            "through special legislation after sector consultation.\n\n"
            "Article 22 — Special authorization regime for emissions (referenced in amended Art. 23)."
        )),
        ("Annex I — Air Quality Standards (as amended 2010)", (
            "Numeric limits (µg/m³) for SO₂, CO, O₃ and particulate matter across 10-minute, 15-minute, "
            "30-minute, 1-hour, 8-hour, 24-hour and annual averaging periods (primary and secondary standards)."
        )),
        ("Annex IA — Inorganic and Organic Carcinogenic Air Pollutants", (
            "Limits for lead, manganese, mercury, arsenic, chromium, nickel, benzene, formaldehyde, "
            "styrene, toluene, tetrachloroethylene and related substances across defined sampling periods."
        )),
        ("Annex IB — Substances with Odorous Properties", (
            "Concentration limits (ppm/vol) for ammonia, bromine, chlorine, methylene chloride, "
            "carbon disulfide, phenol, perchloroethylene, carbon tetrachloride and related substances."
        )),
        ("Annex V — Air quality, marine receptor and chemical standards (extract)", (
            "Air quality management:\n"
            "• Open-air burning of solid waste, liquid waste or any other combustible material causing "
            "environmental degradation is prohibited.\n"
            "• Installation and operation of domestic and industrial incinerators is prohibited, except hospital incinerators.\n"
            "• Where necessary, automatic equipment measuring emitted pollutant quantities and qualities may be required.\n"
            "• Emission of odorous substances perceptible beyond the emitting property boundary is prohibited; "
            "detection by credentialed agents.\n\n"
            "Marine/ocean receiving body standards:\n"
            "Pollutant levels must remain below maximum values for potentially harmful substances. "
            "Industrial and agro-livestock discharges are permitted only if:\n"
            "a) Floating materials: virtually absent;\n"
            "b) Oils and greases: virtually absent;\n"
            "c) Substances causing colour, odour and turbidity: virtually absent;\n"
            "d) Artificial colorants: virtually absent;\n"
            "e) Substances forming objectionable deposits: virtually absent;\n"
            "f) Substances/conditions facilitating undesirable aquatic life: virtually absent;\n"
            "g) BOD₅ (20°C) ≤ 5 mg/l;\n"
            "h) Dissolved oxygen ≥ 6 mg/l;\n"
            "i) pH between 6.5 and 8.5; no change greater than 0.2 units from normal.\n\n"
            "Tables I and IA set minimum/maximum concentration limits (mg/l) for metals, phenols, "
            "surfactants, organochlorine pesticides, herbicides and other chemical substances."
        )),
    ]

    for heading, body in sections:
        doc.add_heading(heading, level=2)
        doc.add_paragraph(body)

    doc.add_page_break()
    doc.add_heading("4P Index Coding Summary", level=1)
    doc.add_paragraph(
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Year (last amendment): {POLICY_FIELDS['policy_year']}\n"
        f"Instruments coded: {len(INSTRUMENTS)}\n"
        f"Full table: 4P_Index_Mozambique_Decree_18_2004.xlsx"
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Decree_18_2004_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Decree 18/2004 — 4P Index Deliverables</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; line-height: 1.6; }}
    h1 {{ color: #1F4E79; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem 1.25rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem; border-radius: 4px; text-decoration: none; }}
  </style>
</head>
<body>
  <h1>4P Index — Decree 18/2004 (as amended 67/2010)</h1>
  <p>Generated {date.today().isoformat()}.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel coding table</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Decree_18_2004_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_decree_18_2004.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Decree_18_2004"
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
