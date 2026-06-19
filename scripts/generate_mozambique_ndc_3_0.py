#!/usr/bin/env python3
"""Generate 4P Index coding table and documents for Mozambique Provisional NDC 3.0."""

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
    "policy_name": "NDC 3.0 Provisória de Moçambique – Período: 2026 - 2035 (Mozambique's Provisional NDC 3.0)",
    "policy_url": "https://unfccc.int/sites/default/files/2025-11/Mozambique%20ProvNDC_ENG.pdf",
    "policy_year": 2025,
    "policy_objective": (
        "Communicate Mozambique's updated Paris Agreement climate commitment for 2026–2035, "
        "including a conditional economy-wide GHG mitigation target, sectoral mitigation coverage "
        "(including waste/landfill), large-scale adaptation priorities, and frameworks linking "
        "international climate finance to national green and circular-economy initiatives."
    ),
    "policy_target": 1,
    "policy_target_text": (
        '"Target of reducing emissions by between 15% and 25% of baseline (BUR2 2020)" (Annex 1, '
        'para. 1(d)); "Type of target: Emissions reduction target below a BAU scenario between 15 to '
        '25% of baseline (BUR2 2020)"; implementation period "January 1, 2026 to December 31, 2035" '
        '(Annex 1, para. 2(a)). Target noted as under final technical review.'
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "International treaty commitment (Paris Agreement NDC) and supreme national climate strategic "
        "policy — not domestic legislation or executive regulation. Classified 0.50 because the document "
        "incorporates quantifiable mitigation targets (15–25% below 2020 BAU by 2035), sectoral coverage, "
        "and planned operational/investment instruments."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "waste management, recycling, municipalities, industry, agriculture, fisheries, water, "
        "energy, tourism, retail, packaging"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NDC implementation is "highly conditional" on international climate finance (Section 5, '
        'Climate finance); aligned with National Climate Finance Strategy 2025-2034 (ENFC); '
        'planned NDC 3.0 Investment Plan and operational plan to follow full NDC communication; '
        'ValoRE and other programmes already receiving/drawing on climate and development finance.'
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex 1 paras. 1(b), 1(d), 3(b): Conditional economy-wide mitigation pledge to reduce "
            "national GHG emissions by 15–25% below 2020 BAU by 2035, explicitly covering the Waste "
            "sector (IPCC sub-categories 4.A.2, 4.C.1, 4.D.1) alongside Energy and AFOLU."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Annex 1, 1(d): quantified "Target of reducing emissions by between 15% and 25% of baseline"; '
            'DINAMC-led process (Section 2); monitoring via BTR/ETF and developing MRV system (Annex 1, 4(a), 5(b)); '
            'conditional on international support (Annex 1, 6(d)).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Provisional NDC submitted to UNFCCC Nov 2025; final target subject to modelling review. "
            "Waste-sector landfill methane reductions are part of covered IPCC categories. No plastic-specific sub-target."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex 1, 3(b): Formal inclusion of Waste sector in NDC 3.0 mitigation scope using IPCC "
            "classification — sub-categories 4.A.2 (solid waste disposal), 4.C.1 (incineration/open burning), "
            "4.D.1 (wastewater treatment) — governing national reporting and mitigation planning for waste streams."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Annex 1, 3(b): "Waste: including (sub-)categories 4.A (4.A.2.); 4.C (4.C.1.); 4.D (4.D.1.)"; '
            'IPCC 2006 inventory methods (Annex 1, 5(a)-(d)); sector consultations include Waste (Annex 1, 4(a)).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Governance/planning instrument defining sector coverage; applies to municipal solid waste "
            "including plastics in landfills and open burning. No explicit 'plastics' term in document."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Annex 1, 3(d) / Section on mitigation co-benefits: ValoRE (Programme for Sustainable Waste "
            "Management in Mozambique) strengthens waste infrastructure and circular value chains; "
            "implemented in three municipalities with improved systems and recyclable recovery, reducing "
            "landfill emissions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Activities already implemented in three municipalities have improved waste management systems '
            'and established circular value chain opportunities, with measurable mitigation co-benefits through '
            'reduced landfill emissions and more efficient resource use of available resources (recyclables...)"'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary concrete waste/circular-economy programme cited in NDC; directly relevant to plastic "
            "and recyclable waste streams though plastics not named explicitly."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Annex 1, 3(d): Promoting Employment for Women for Green Transformation initiative supports "
            "green business models in value chains including renewable energy, circular economy and blue economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"promotes green employment opportunities for women and environmentally sustainable business '
            'models in value chains such as renewable energy, circular economy and blue economy"'
        ),
        "instrument_score": "[auto]",
        "comments": "Voluntary/green economy programme; circular economy reference enables plastic-relevant value-chain actions.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex 1, 4(a): DINAMC-led participatory NDC planning and sectoral consultation process, "
            "including dedicated Waste sector consultations and integration of NDC 2.0 evaluation, BUR2, "
            "NAP, ENDE and PQG into NDC 3.0."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'DINAMC and GIIMC lead stocktaking (Annex 1, 4(a)); sectoral consultations include "Waste"; '
            'provincial focal points engaged; Quality Control team provides technical guidance (Section 2).'
        ),
        "instrument_score": "[auto]",
        "comments": "Coordination instrument for embedding waste/plastic measures in national climate planning.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex 1, 4(a): Planned NDC 3.0 operational plan, Investment Plan and capacity-building "
            "programme to be developed after full NDC communication to UNFCCC."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"The development of these key documents will follow the full NDC 3.0 communication to the UNFCCC."'
        ),
        "instrument_score": "[auto]",
        "comments": "Enabling/planning power not yet exercised; will operationalise finance and waste investments.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Annex 1, 4(a) / 5(b): Development and operationalization of Strengthened National Transparency "
            "Framework (ETF) and MRV system (including Carbon Registry and climate finance tracking) to monitor "
            "mitigation programme implementation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"the country\'s Measurement, Reporting, and Verification (MRV) system (currently developing)"; '
            '"strengthening of implementation will be done through the development and operationalization of '
            'a Strengthened National Transparency Framework (ETF), based on the National Climate Change '
            'Monitoring and Evaluation System (SNMAMC)".'
        ),
        "instrument_score": "[auto]",
        "comments": "Monitoring instrument under development; scored in_force=0 per enabling-power rule.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Section 5 (Environmental protection): NDC 3.0 reinforces pollution control and sustainable "
            "environmental management, addressing national challenges in air quality and waste treatment "
            "on the path to low-carbon development."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"NDC 3.0 reinforces pollution control and sustainable environmental management as a cornerstone '
            'on the path towards low-carbon development"; references Environment Law No. 20/97.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic commitment rather than binding standard; explicitly cites waste treatment deficits. "
            "Scored 0.40 as information/strategic voluntary framing without mandatory provisions in this document."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Section 5 (Communication and awareness-raising): Expand use of community radios, school curricula, "
            "civil society networks and social media to mainstream climate action and mobilize engagement across sectors."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"Communication and awareness-raising are essential to mainstream the implementation of the NDC"; '
            '"Mozambique already makes use of tools such as community radios, school curricula, civil society '
            'networks and social media".'
        ),
        "instrument_score": "[auto]",
        "comments": "Awareness instrument applicable to consumer waste/plastic pollution behaviour change.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Section 5 (Climate finance) / Section 2: NDC 3.0 conditioned on international climate finance "
            "aligned with National Climate Finance Strategy 2025-2034 (ENFC); prioritizes pledged/committed/disbursed funding."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Access to climate finance and investment is of utmost importance"; feasibility criteria include '
            '"Availability of international climate finance against definitions and standards outlined in the '
            'National Climate Finance Strategy 2025-2034 (ENFC)" and "Priority given to actions for which '
            'funding is pledged, committed or disbursed".'
        ),
        "instrument_score": "[auto]",
        "comments": "Economic/finance instrument enabling waste and circular-economy project funding including ValoRE.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Annex 1, 5(g): Mozambique intends to explore Article 6 voluntary cooperation (ITMOs, Non-Market "
            "Approaches) to achieve mitigation targets cost-effectively, potentially including waste-sector outcomes."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            '"Mozambique intends to explore the use of voluntary cooperation mechanisms (e.g., Internationally '
            'Transferred Mitigation Results (ITMOs), Non-Market Approach (NMAs)) under Article 6 of the Paris Agreement".'
        ),
        "instrument_score": "[auto]",
        "comments": "Enabling power ('intends to explore'); in_force=0 until operationalized.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Section 3 / Annex 1, 4(c): Adaptation component structured across six sectors (Agriculture/Forestry/Blue "
            "Economy, Health, WASH/Water, Infrastructure, Education, Early Warning) aligned with NAP 2023-2032, "
            "supporting coastal and ecosystem resilience relevant to marine litter pathways."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Section 3: "provisional NDC 3.0 structures the adaptation component in six sectors"; '
            'reaffirms NAP objectives on integrating adaptation into planning/budgeting and district-level resilience.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Adaptation governance; Blue Economy and coastal infrastructure indirectly relevant to marine plastics. "
            "Not plastic-specific."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Section 5 (Institutional strengthening): Strengthen government capacity and inter-sectoral coordination "
            "to integrate climate action and NDC implementation, including ETF reporting under Paris Agreement MPGs."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Government readiness and capacity are essential"; "Strengthening institutional coordination will enable '
            'government entities to report in line with the Strengthened Transparency Framework (ETF)".'
        ),
        "instrument_score": "[auto]",
        "comments": "Cross-cutting governance enabling waste-sector NDC implementation.",
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
    rows = [{**POLICY_FIELDS, **inst} for inst in INSTRUMENTS]
    return pd.DataFrame(rows, columns=COLUMNS)


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


def create_english_doc():
    """Formatted English policy document (source is already English)."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("Mozambique's Provisional NDC 3.0", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Period: 2026 – 2035")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run("Republic of Mozambique | Ministry of Agriculture, Environment and Fisheries\n").italic = True
    meta.add_run("National Directorate for Environment and Climate Change (DINAMC)\n")
    meta.add_run(f"Formatted policy text for 4P Index analysis — {date.today().isoformat()}\n")
    meta.add_run("Source: https://unfccc.int/sites/default/files/2025-11/Mozambique%20ProvNDC_ENG.pdf")

    sections = [
        ("1. Introduction", (
            "Mozambique submitted its INDC in 2015, which became NDC 1.0 (2020–2030) upon ratification of the "
            "Paris Agreement in 2018. NDC 2.0 was submitted in 2021 for 2021–2025. Mozambique is now developing "
            "NDC 3.0 for 2026–2035 as a provisional communication under the Paris Agreement, using a participatory "
            "whole-of-society approach led by DINAMC/MAAP with government, private sector, civil society, academia "
            "and international partners."
        )),
        ("2. From NDC 2.0 to NDC 3.0", (
            "Preparation is guided by lessons from NDC 2.0, a Quality Control team, and an NDC Evaluation and "
            "Implementation Report. Sectoral consultations ensure stakeholder ownership. Improvements include: "
            "robustness via updated GHG inventory (1990–2020, BUR2); structured adaptation across six sectors; "
            "and realism based on climate finance availability (ENFC 2025–2034), national priorities, pledged "
            "funding, technology, and implementation capacity."
        )),
        ("3. Adaptation (Article 7)", (
            "Adaptation is Mozambique's top priority. NDC 3.0 reaffirms the National Adaptation Plan (NAP) 2023–2032 "
            "objectives: integrate adaptation into planning/budgeting; improve data, technology and finance access; "
            "and implement district-level resilience actions. Adaptation is structured across: (i) Agriculture, Forestry "
            "and Blue Economy; (ii) Health; (iii) WASH and Water Resources; (iv) Infrastructure; (v) Education; and "
            "(vi) Early Warning and Social Protection. Loss & Damage under Article 8 is being explored."
        )),
        ("4. Mitigation (Article 4)", (
            "As an LDC with minimal global emissions, Mozambique nonetheless sets mitigation targets relative to BAU "
            "(BUR2 2020). NDC 3.0 aligns with the developing LT-LEDS to 2050 and uses IPCC 2006 sectoral classifications."
        )),
        ("5. Cross-cutting themes", (
            "Institutional strengthening; gender equality and vulnerable groups; conflict sensitivity and human mobility; "
            "knowledge, information and technology; communication and awareness-raising; environmental protection "
            "(including pollution control and waste treatment challenges); climate finance; and loss and damage."
        )),
        ("6. ICTU guidance", (
            "Mozambique applies Decision 4/CMA.1 guidance. Annex 1 provides structured ICTU information."
        )),
        ("Annex 1 — Key mitigation information", (
            "• Target type: 15–25% emissions reduction below 2020 BAU by 2035 (under finalization)\n"
            "• Base year: 2020 (30,600 Gg CO₂eq excluding LULUCF; 91,300 Gg including LULUCF)\n"
            "• Period: 1 January 2026 – 31 December 2035\n"
            "• Sectors: Energy, AFOLU (incl. Blue Economy measures), IPPU (modelled, no selected measures), "
            "Waste (4.A.2, 4.C.1, 4.D.1)\n"
            "• Gases: CO₂, CH₄, N₂O, SF₆\n"
            "• Key waste initiative: ValoRE — sustainable waste management and circular value chains in three municipalities\n"
            "• Circular economy referenced in women's green employment programme\n"
            "• Conditional on international finance, technology transfer and capacity building (Arts. 9–11)\n"
            "• Article 6 ITMO cooperation to be explored\n"
            "• MRV/ETF system under development; operational plan and investment plan to follow final NDC"
        )),
    ]
    for heading, body in sections:
        doc.add_heading(heading, level=1)
        doc.add_paragraph(body)

    doc.add_page_break()
    doc.add_heading("4P Index Coding Summary", level=1)
    doc.add_paragraph(
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Instruments coded: {len(INSTRUMENTS)}\n"
        f"See: 4P_Index_Mozambique_NDC_3_0_Provisional.xlsx"
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_NDC_3_0_Provisional_English.docx")
    doc.save(path)
    return path


def create_portuguese_doc():
    """Portuguese translation of the provisional NDC 3.0."""
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("NDC 3.0 Provisória de Moçambique", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Período: 2026 – 2035")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run("República de Moçambique | Ministério da Agricultura, Ambiente e Pescas\n").italic = True
    meta.add_run("Direcção Nacional do Ambiente e das Alterações Climáticas (DINAMC)\n")
    meta.add_run(f"Tradução para português — {date.today().isoformat()}\n")
    meta.add_run("Fonte: https://unfccc.int/sites/default/files/2025-11/Mozambique%20ProvNDC_ENG.pdf")

    sections = [
        ("1. Introdução", (
            "Moçambique apresentou a sua Contribuição Intencionalmente Determinada (INDC) em 2015, que se tornou "
            "a NDC 1.0 (2020–2030) com a ratificação do Acordo de Paris em 2018. A NDC 2.0 foi submetida em 2021 "
            "para o período 2021–2025. Moçambique está agora a desenvolver a NDC 3.0 para 2026–2035 como comunicação "
            "provisória ao abrigo do Acordo de Paris, através de uma abordagem participativa de sociedade inteira "
            "liderada pela DINAMC/MAAP com governo, sector privado, sociedade civil, academia e parceiros internacionais."
        )),
        ("2. Da NDC 2.0 à NDC 3.0", (
            "A preparação é orientada pelas lições da NDC 2.0, uma equipa de Controlo de Qualidade e um Relatório de "
            "Avaliação e Implementação da NDC. Consultas sectoriais garantem apropriação pelas partes interessadas. "
            "Melhorias incluem: robustez via inventário actualizado de GEE (1990–2020, BUR2); adaptação estruturada "
            "em seis sectores; e realismo com base na disponibilidade de financiamento climático (ENFC 2025–2034), "
            "prioridades nacionais, financiamento comprometido, tecnologia e capacidade de implementação."
        )),
        ("3. Adaptação (Artigo 7)", (
            "A adaptação é a principal prioridade de Moçambique. A NDC 3.0 reafirma os objectivos do Plano Nacional "
            "de Adaptação (PNA) 2023–2032: integrar a adaptação no planeamento e orçamentação; melhorar o acesso a "
            "dados, tecnologia e financiamento; e implementar acções de resiliência ao nível distrital. A adaptação "
            "estrutura-se em: (i) Agricultura, Silvicultura e Economia Azul; (ii) Saúde; (iii) WASH e Recursos Hídricos; "
            "(iv) Infra-estruturas; (v) Educação; e (vi) Alerta Precoce e Protecção Social. Perdas e Danos ao abrigo "
            "do Artigo 8 estão a ser explorados."
        )),
        ("4. Mitigação (Artigo 4)", (
            "Como PLD com emissões globais mínimas, Moçambique define nonetheless metas de mitigação relativas ao "
            "cenário BAU (BUR2 2020). A NDC 3.0 alinha-se com a LT-LEDS em desenvolvimento até 2050 e usa classificações "
            "sectoriais do IPCC 2006."
        )),
        ("5. Temas transversais", (
            "Fortalecimento institucional; igualdade de género e grupos vulneráveis; sensibilidade a conflitos e mobilidade "
            "humana; conhecimento, informação e tecnologia; comunicação e sensibilização; protecção ambiental (incluindo "
            "controlo da poluição e desafios de tratamento de resíduos); financiamento climático; e perdas e danos."
        )),
        ("6. Orientação ICTU", (
            "Moçambique aplica a orientação da Decisão 4/CMA.1. O Anexo 1 fornece informação ICTU estruturada."
        )),
        ("Anexo 1 — Informação chave de mitigação", (
            "• Tipo de meta: redução de emissões de 15–25% abaixo do BAU de 2020 até 2035 (em finalização)\n"
            "• Ano de referência: 2020 (30.600 Gg CO₂eq sem UTCATF; 91.300 Gg com UTCATF)\n"
            "• Período: 1 de Janeiro de 2026 – 31 de Dezembro de 2035\n"
            "• Sectores: Energia, AFOLU (incl. medidas de Economia Azul), IPPU (modelado, sem medidas seleccionadas), "
            "Resíduos (4.A.2, 4.C.1, 4.D.1)\n"
            "• Gases: CO₂, CH₄, N₂O, SF₆\n"
            "• Iniciativa chave de resíduos: ValoRE — gestão sustentável de resíduos e cadeias de valor circulares em três municípios\n"
            "• Economia circular referenciada no programa de emprego verde para mulheres\n"
            "• Condicional ao apoio internacional, transferência de tecnologia e reforço de capacidades (Arts. 9–11)\n"
            "• Cooperação ITMO do Artigo 6 a explorar\n"
            "• Sistema MRV/ETF em desenvolvimento; plano operacional e plano de investimento a seguir à NDC final"
        )),
    ]
    for heading, body in sections:
        doc.add_heading(heading, level=1)
        doc.add_paragraph(body)

    path = os.path.join(OUTPUT_DIR, "Mozambique_NDC_3_0_Provisional_Portuguese.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique NDC 3.0 — 4P Index Deliverables</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; line-height: 1.6; }}
    h1 {{ color: #1F4E79; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem 1.25rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem; border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
    a.download:hover {{ background: #163a5c; }}
    .meta {{ color: #555; font-size: 0.95rem; }}
  </style>
</head>
<body>
  <h1>Plastic Pollution Policy Index (4P Index)</h1>
  <p class="meta">Mozambique — Provisional NDC 3.0 (2026–2035)</p>
  <p>Generated on {date.today().isoformat()}.</p>

  <div class="card">
    <h2>4P Index Coding Table (Excel)</h2>
    <p>{len(INSTRUMENTS)} instrument rows (columns A–W).</p>
    <a class="download" href="{basename}.xlsx" download>Download Excel (.xlsx)</a>
  </div>

  <div class="card">
    <h2>4P Index Coding Table (CSV)</h2>
    <a class="download" href="{basename}.csv" download>Download CSV</a>
  </div>

  <div class="card">
    <h2>English Policy Document (Word)</h2>
    <p>Formatted English text (source document is already in English).</p>
    <a class="download" href="Mozambique_NDC_3_0_Provisional_English.docx" download>Download English (.docx)</a>
  </div>

  <div class="card">
    <h2>Portuguese Translation (Word)</h2>
    <p>Tradução para português da NDC 3.0 Provisória.</p>
    <a class="download" href="Mozambique_NDC_3_0_Provisional_Portuguese.docx" download>Download Portuguese (.docx)</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_ndc_3_0.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_NDC_3_0_Provisional"
    df = build_dataframe()
    csv_path = export_csv(df, basename)
    xlsx_path = export_xlsx(df, basename)
    eng_path = create_english_doc()
    pt_path = create_portuguese_doc()
    html_path = create_index_html(basename)
    print("Generated:")
    for p in [csv_path, xlsx_path, eng_path, pt_path, html_path]:
        print(f"  {p}")


if __name__ == "__main__":
    main()
