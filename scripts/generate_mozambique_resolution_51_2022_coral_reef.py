#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 51/2022 (ECOR)."""

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
        "Estratégia Nacional de Gestão e Conservação dos Recifes de Coral "
        "(2022–2032) — Resolução n.º 51/2022"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/moz214814.pdf",
    "policy_year": 2022,
    "policy_objective": (
        "Approve the National Coral Reef Management and Conservation Strategy (ECOR) "
        "2022–2032 under Sea Law (Law 20/2019) to ensure ecological integrity and resilience "
        "of Mozambique's coral reefs (~1,800 km², 300+ coral species) by eliminating "
        "anthropogenic degradation and promoting sustainable use — including mitigation of "
        "land-based pollution (sedimentation, coastal discharges) and coastal water pollution "
        "from plastics, chemicals and organic sources identified as key local threats."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'Annex II explicit threat: "Poluição das águas costeiras (plásticos, produtos químicos '
        'e orgânicos)". Strategic targets: lose no more than 5% live coral cover and 5% reef '
        'extent from human causes by 2032 (baseline 2020); conserve ≥30% of reef area in Marine '
        'Protected Areas by 2030; abandon illegal/destructive reef fishing by 2026; national '
        'reef monitoring programme operational by 2026; sustainable financing mechanism by 2025.'
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "Council of Ministers Resolution n.º 51/2022 (approved 15 November 2022, published "
        "28 December 2022, Boletim da República Series I No. 250) under Sea Law Art. 79 — "
        "national coral-reef strategy with annexed action plan (ECOR), not binding sectoral "
        "regulation. Coordinated by MIMAIP; articulates with POLMAR, POEM, REPMAR, REJUEM, "
        "REICIM and Nairobi Convention Regional Coral Reef Action Plan."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging"
    ),
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": (
        "consumption, environmental leakage"
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        "Action plan assigns line-item budgets across three pillars (e.g. MZN 10 million reef "
        "mapping/validation; MZN 2.1 million fisheries fiscalization and good-practices guide; "
        "MZN 1.1 million extractive-industry mitigation guide; MZN 8.9 million environmental "
        "education pilot; MZN 1.7 million communication plan). Pilar 3 target: sustainable "
        "financing mechanism by 2025 (ProAzul, government funds)."
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Annex II / strategic context: Coastal water pollution explicitly identified as a "
            "threat to coral reefs, including plastics, chemical and organic pollutants; local "
            "threats also include sedimentation and pollution from land-based discharges; "
            "unplanned coastal and marine development (ports, pipelines, offshore hydrocarbons) "
            "aggravates pollution pressure on reef ecosystems."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Annex II (Desenvolvimento costeiro e marinho): "Poluição das águas costeiras '
            '(plásticos, produtos químicos e orgânicos)"; Sec. 1.1: "sedimentação e poluição '
            'com origem em descargas que provêm do meio terrestre".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary explicit plastic instrument — only direct 'plásticos' reference in strategy. "
            "Not coded separately: climate-change/ocean warming threats (global factor)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Strategy objective and principles: Ensure ecological integrity and resilience by "
            "eliminating anthropogenic degradation forms and promoting sustainable use; polluter-pays "
            "and user-pays principles — actors responsible for reef degradation must repair and "
            "compensate damages in proportion to harm caused."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Resolution enacting text: "eliminação das formas de degradação antropogénica e '
            'promoção do uso sustentável"; Princípio do utilizador e poluidor pagador — '
            '"responsabilização de todas as pessoas singulares ou colectivas pela reparação e '
            'compensação na mesma proporção dos danos causados aos recifes de coral".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic liability framework — applicable to land-based pollution including litter. "
            "Not coded separately: precaution, ecosystem-approach or participation principles."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan 1.2: Abandon illegal and destructive reef fishing practices by 2026; "
            "strengthen fiscalization of fishing gear in coral ecosystems; develop and validate "
            "Good Practices Guide for fisheries and tourism on coral reefs — addresses abandoned "
            "nets and materials damaging reefs (ghost-gear relevance)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action 1.2: "Práticas de pesca ilegal e destrutiva em recifes de coral são '
            'abandonadas e é reforçada a fiscalização"; text notes industrial fishing may damage '
            'reefs "por meio de redes e outros materiais abandonados"; Guião de boas práticas de '
            'pesca e turismo (2022–2025; MZN 2,100,000).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Fisheries/ghost-gear instrument — no explicit 'plástico' but abandoned nets cited. "
            "Bundled with IUU risk (ship groundings on reefs). Not coded: coral restoration "
            "guidelines (action 1.4)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan 1.3: Good Practices Guide for mitigating extractive-industry impacts "
            "on coral reefs (oil & gas, mining, ports, pipelines); applied in Environmental "
            "Impact Assessment and post-assessment environmental management and monitoring — "
            "reduces pollution and physical damage from coastal/marine development on reefs."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action 1.3.1: "Elaborar Guião de boas práticas para a mitigação de impactos da '
            'indústria extractiva nos recifes de coral" — workshop with environmental authorities, '
            'Oil & Gas and mining regulators and companies (2022–2024; MZN 1,100,000; MTA, MIREME).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Development-impact mitigation — ports, pipelines and offshore platforms listed as "
            "reef threats in Annex II. Not coded separately: ACM management strengthening (2.2)."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan Meta 4 / 1.4.2: National Coral Reef Monitoring Programme — map reef "
            "distribution, assess ecological and socioeconomic state, identify degraded reef areas "
            "and causes of degradation; publish national conservation status reports to inform "
            "management decisions (2023–2026)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action 1.4.2.1: "Identificar áreas recifais degradadas e as causas de degradação"; '
            'Action 4.1.2: National Monitoring Plan with ecological/socioeconomic indicators; '
            'reef mapping and field validation (MZN 10,000,000 for 4.1.1).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Monitoring/degradation-cause identification — pollution (including plastics) tracked "
            "as threat category. Not coded separately: Allen Coral Atlas validation or scholarship "
            "programme (3.3)."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Action Plan Meta 2: Integrate coral reefs into POEM and complementary legal "
            "instruments as conservation areas; conserve at least 30% of Mozambique's reef area "
            "effectively in Marine Protected Areas by 2030; systematic conservation planning for "
            "new reef protected areas — reduces uncontrolled development and pollution pressures."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Meta 2: "conservar pelo menos 30% da área de recifes de coral em Moçambique de forma '
            'efectiva em Áreas de Conservação Marinhas" by 2030; Actions 2.1.1–2.1.2 integrate reef '
            'maps into POEM zoning; Action 2.3.1 systematic conservation planning for new areas.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Spatial protection instrument — indirect plastic-pollution reduction by limiting "
            "incompatible uses. Links to SDG 14.5 / CBD GBF targets. Not coded: legal framework "
            "review (Meta 7, Pilar 3)."
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
        "National Coral Reef Management and Conservation Strategy — ECOR", 0
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 51/2022 — Period 2022–2032")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Estratégia Nacional de Gestão e Conservação dos Recifes "
        "de Coral (2022–2032) — Resolução n.º 51/2022\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://faolex.fao.org/docs/pdf/moz214814.pdf")

    sections = [
        ("Resolution No. 51/2022 — Enacting Provisions", (
            "Having regard to the need to ensure the ecological integrity of coral reefs in order "
            "to increase their resilience, through elimination of anthropogenic forms of degradation "
            "and promotion of sustainable use, under the combined provisions of paragraph 1(f) and "
            "paragraph 2 of Article 79 of Law No. 20/2019 of 8 November (Sea Law), the Council of "
            "Ministers resolves:\n\n"
            "Sole Article. The Coral Reef Management and Conservation Strategy 2022–2032 is "
            "approved as an annex forming an integral part of this Resolution.\n\n"
            "Approved by the Council of Ministers on 15 November 2022. Published in the Official "
            "Gazette (Boletim da República), Series I, No. 250, 28 December 2022.\n\n"
            "Signed: Adriano Maleiane, Prime Minister."
        )),
        ("1. Context and Importance", [
            ("1.1 Global and national significance", (
                "Coral reefs occupy only 0.2% of ocean surface but shelter more than 25% of known "
                "marine species. They support essential ecosystem services for approximately one "
                "billion people and a global economy worth USD 2.7 trillion annually (USD 36 billion "
                "from tourism alone).\n\n"
                "Mozambique has approximately 2,700 km of coastline with reef area estimated at "
                "1,800 km² (2001 data), at least 300 coral species (~40% of global nominal species), "
                "800+ reef fish species and extensive associated biodiversity distributed across "
                "three biogeographic zones: Northern Coral Coast (Rovuma to Primeiras e Segundas), "
                "Central zone, and Southern Parabolic Dune Coast (Bazaruto to Ponta do Ouro)."
            )),
            ("1.2 Threats and rationale", (
                "Global threats include climate-change-driven warming, acidification and mass "
                "bleaching events (1998, 2004, 2010, 2014–2017). Local threats include unplanned "
                "coastal development, destructive fishing, sedimentation and pollution from "
                "land-based discharges.\n\n"
                "Approximately 60% of Mozambique's population lives in coastal areas; artisanal "
                "fisheries (~80% of total catch) depend significantly on reef resources. Industrial "
                "fishing may damage reefs through abandoned nets and materials. IUU fishing poses "
                "grounding and food-chain risks.\n\n"
                "Mitigating local impacts can increase reef resilience to climate change and, in "
                "some cases, reverse decline within a decade. The strategy aligns with the Nairobi "
                "Convention Regional Coral Reef Action Plan and POLMAR Pillar C on ecosystem "
                "conservation."
            )),
        ]),
        ("2. Strategic Framework", [
            ("2.1 Vision, mission and objective", (
                'Vision: "A country with resilient and developing coral reef ecosystems, ensuring '
                'ecosystem services for current and future generations."\n\n'
                'Mission: "Conserve and protect coral reef ecosystems to ensure their ecological '
                'integrity and quality of ecosystem services for Mozambican society."\n\n'
                'Objective: "Ensure ecological integrity of coral reefs to increase resilience '
                'through elimination of anthropogenic degradation forms and promotion of sustainable use."'
            )),
            ("2.2 Guiding principles (summary)", (
                "Rational use and management of finite living resources; recognition of local "
                "knowledge; evidence-based interventions; precaution; balance between economic "
                "development and conservation; polluter-pays and user-pays (users pay for access "
                "and those who degrade reefs must compensate proportionally); integrated intersectoral "
                "management; broad citizen participation; sustainability; international cooperation."
            )),
            ("2.3 Three intervention pillars and strategic targets", (
                "Pilar 1 — Protection and conservation:\n"
                "• By 2032, lose no more than 5% of live coral cover and 5% of reef extent from "
                "human causes (baseline 2020/2022);\n"
                "• By 2030, create legal instruments/programmes and conserve ≥30% of reef area "
                "effectively in Marine Protected Areas.\n\n"
                "Pilar 2 — Education, capacity and scientific knowledge:\n"
                "• By 2032, build national technical/scientific capacity for sustainable management, "
                "research and effective monitoring;\n"
                "• By 2030, reef monitoring programme information supports CBD Global Biodiversity "
                "Framework targets;\n"
                "• By 2032, continuous systematic environmental education for reef users and managers.\n\n"
                "Pilar 3 — Coordination and financing:\n"
                "• By 2032, effective management through clear legal responsibilities and operational "
                "inter-institutional coordination;\n"
                "• By 2025, sustainable financing mechanism for reef conservation and monitoring."
            )),
        ]),
        ("3. Annex II — Threats to Coral Reefs (Plastic-Relevant)", (
            "Climate change: ocean warming and acidification; increased cyclone frequency.\n\n"
            "Unsustainable fisheries: destructive gear and practices on reefs; coral removal for "
            "ornamentation/construction; aquarium collection; sport fishing; illegal/unmonitored fishing.\n\n"
            "Coastal and marine development:\n"
            "• Lack of fishing zoning around reefs;\n"
            "• Negative practices: sport fishing, unplanned coastal development, coral removal for "
            "tourist craft markets, coastal mining, offshore hydrocarbons, port development, cables "
            "and pipelines;\n"
            "• Coastal water pollution (plastics, chemical and organic products);\n"
            "• Insufficient protected reef area (total and strict protection).\n\n"
            "Capacity gaps: limited diving capacity, monitoring gaps inside and outside protected "
            "areas, insufficient equipped fiscalization, limited financing.\n\n"
            "Institutional gaps: previously no unified national reef strategy; limited coordination; "
            "weak local community involvement; restrictive research regulations (REICIM)."
        )),
        ("4. Governance and Implementation", [
            ("4.1 Institutional coordination", (
                "MIMAIP coordinates strategy implementation. A Coordination Unit (UC) monitors, "
                "harmonizes and evaluates progress; reports to the National Sea Council (CNM) on "
                "structural matters. Provincial and district/municipal teams include environment, "
                "maritime, public works, mining, transport, agriculture, tourism, education, "
                "universities, NGOs, communities and private sector representatives.\n\n"
                "A central Technical-Scientific Unit (UTC) advises the UC on monitoring plans, data "
                "review and best-practice recommendations. Activities must be integrated into "
                "sectoral medium/long-term fiscal scenarios and Economic and Social Plans (PES)."
            )),
            ("4.2 Implementation instruments", (
                "An Implementation Guide will describe processes, procedures, methodologies and "
                "indicators for each action-plan activity. Legal framework and inter-institutional "
                "coordination will be reviewed. Reef activities in conservation areas follow ANAC "
                "responsibilities coordinated with UC; research follows REICIM. Activities must "
                "respect POLMAR, REJUEM, REPMAR, REICIM and POEM.\n\n"
                "Monitoring: biannual interim evaluations; final ten-year evaluation by independent "
                "entity through public tender. Adaptive management approach with lessons learned "
                "and mitigation recommendations."
            )),
        ]),
        ("5. Action Plan — Key Measures (Plastic-Relevant)", [
            ("5.1 Pilar 1 — Protection", (
                "1.1 Communication and implementation guide dissemination (2022–2024).\n\n"
                "1.2 Abandon illegal/destructive reef fishing by 2026; strengthen gear fiscalization; "
                "Good Practices Guide for fisheries and tourism on reefs; fiscalization gap analysis "
                "(MZN 2,100,000).\n\n"
                "1.3 Good Practices Guide for mitigating extractive-industry impacts (oil & gas, "
                "mining) in EIA and post-EIA environmental management (MZN 1,100,000).\n\n"
                "1.4 Ecological restoration: national restoration guidelines; identify degraded reef "
                "areas and causes of degradation; train 20 technicians in restoration methods.\n\n"
                "2.1–2.3 Integrate reefs into POEM and legal instruments; strengthen existing ACM "
                "management; systematic conservation planning for new protected reef areas (30% "
                "target by 2030)."
            )),
            ("5.2 Pilar 2 — Monitoring and education", (
                "4.1 Map and validate national reef distribution (Allen Coral Atlas validation; "
                "MZN 10,000,000); implement National Coral Reef Monitoring Programme with "
                "ecological/socioeconomic indicators; publish national conservation status reports.\n\n"
                "5.1 Integrate reef environmental education into school curricula and pilot "
                "programmes (50 schools, 10 CSOs by 2029; MZN 8,900,000 pilot); train teachers "
                "and community fisheries councils (CCPs) on reef management and monitoring."
            )),
            ("5.3 Pilar 3 — Financing and legal framework", (
                "6.1 Define detailed ECOR budget and fundraising strategy; establish government "
                "financing mechanisms (ProAzul, relevant sector funds) by 2025.\n\n"
                "Meta 7 (by 2032): Effective reef management through fiscalization of legal "
                "framework with clear entity responsibilities and operational coordination."
            )),
        ]),
        ("6. Note on Plastics", (
            "ECOR contains one explicit plastics reference in Annex II threat taxonomy: coastal "
            "water pollution including plastics, chemical and organic products. Additional "
            "plastic-relevant elements are implicit:\n\n"
            "— Land-based sedimentation and pollution from terrestrial discharges (Section 1.1);\n"
            "— Abandoned fishing nets and materials damaging reefs;\n"
            "— Polluter-pays principle for degradation compensation;\n"
            "— Extractive-industry and port/pipeline development impact mitigation;\n"
            "— National monitoring and degraded-area/cause identification;\n"
            "— Marine Protected Area expansion reducing incompatible coastal uses.\n\n"
            "The action plan does not include dedicated plastic-reduction, EPR or waste-management "
            "targets. Operational plastic and marine-litter measures remain in EGIZC, EDEA, POEM, "
            "Decree 45/2006, Decree 97/2020 and Decree 16/2015.\n\n"
            "Note on period: The approved strategy covers 2022–2032 (Resolution 51/2022). Some "
            "action-plan baseline years reference 2020."
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
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — education/restoration "
        f"actions not overcoded)\n"
        f"See: 4P_Index_Mozambique_Resolution_51_2022_Coral_Reef.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Resolution_51_2022_Coral_Reef_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 51/2022 ECOR — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 51/2022 (ECOR)</h1>
  <p>National Coral Reef Management and Conservation Strategy 2022–2032. Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Resolution_51_2022_Coral_Reef_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_51_2022_coral_reef.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_51_2022_Coral_Reef"
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
