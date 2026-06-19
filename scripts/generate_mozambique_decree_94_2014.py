#!/usr/bin/env python3
"""Generate 4P Index coding table and English translation for Mozambique Decree 94/2014."""

import os
from datetime import date

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "/workspace/output/mozambique"
os.makedirs(OUTPUT_DIR, exist_ok=True)

POLICY_FIELDS = {
    "policy_name": "Regulamento sobre a Gestão de Resíduos Sólidos Urbanos (Decreto n.º 94/2014)",
    "policy_url": "https://faolex.fao.org/docs/pdf/moz148514.pdf",
    "policy_year": 2014,
    "policy_objective": (
        "Establish national rules for the environmentally sound management of urban solid waste, "
        "including prevention and reduction of waste generation, segregation (including plastics), "
        "selective collection, recycling/valorization, and safe final disposal to protect public "
        "health and the environment."
    ),
    "policy_target": 1,
    "policy_target_text": (
        '"Os Conselhos Municipais ou vilas não municipalizadas devem encerrar as lixeiras a céu aberto '
        'nas suas áreas de jurisdição e garantir a construção de aterros sanitários ou controlados no '
        'prazo de três anos após a publicação do presente regulamento." (Art. 25); '
        'Annex I requires plans to include "Objectivos e metas do Plano durante os cincos anos de '
        'vigência do mesmo"; Annex II requires reporting on "Taxa de cobertura de recolha de resíduos" '
        'and "Percentagem e tonelagem de resíduos destinados para reciclagem".'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (Regulamento) approved by Council of Ministers Decree (Decreto n.º 94/2014), "
        "issued under Article 33 of the Environment Law — sub-legislative instrument, not parliamentary legislation."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "production, consumption, retail, waste management, recycling, municipalities, industry"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'Art. 19: "As taxas de limpeza urbana são estabelecidas e cobradas pelos Conselhos Municipais ou '
        'Governos Distritais..."; Art. 22: "60% para o FUNAB (Fundo do Ambiente)" and "60% dos valores '
        'recebidos pelo FUNAB devem ser aplicados em actividades de promoção de boas práticas de gestão '
        'de resíduos sólidos urbanos e de melhoria das actividades de monitoria e fiscalização do '
        'cumprimento do presente regulamento."'
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5: Assigns urban solid waste management competencies to the Ministry of Environment "
            "(rule-making, inspections, national registry, monitoring, sanctions) and to Municipal/District "
            "Councils (local waste services, by-laws, collection procedures, local registry, penalties)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 5(1)(h)-(i): "Penalizar os gestores dos Conselhos Municipais..."; '
            '"Monitorar e fiscalizar o cumprimento das disposições do presente Regulamento." '
            'Art. 5(2)(i): municipal penalty authority.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Core multi-level governance framework applicable to all waste streams including plastics. "
            "Coordination instrument scored 0.20 per coding rules."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 4(d): Mandates the waste management hierarchy — prevention/reduction, reuse, recycling, "
            "other valorization, then disposal — using best available technologies at sustainable cost."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 4(d): "a gestão de resíduos sólidos urbanos deve respeitar a seguinte ordem de prioridades..."; '
            'Art. 5(1)(i): Ministry monitors compliance; Art. 20(2): violations of Art. 4 are fineable.'
        ),
        "instrument_score": "[auto]",
        "comments": "Foundational circular economy principle; not plastic-specific but directly governs plastic waste prioritization.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 4(c): Establishes prevention and reduction of urban solid waste generation and harmful "
            "character as a priority management objective."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 4(c): "constitui objectivo prioritário... evitar e reduzir a sua produção"; '
            'Art. 11(a): producers must "Minimizar a produção de resíduos sólidos urbanos"; '
            'Art. 20(2): fines for non-compliance with Art. 4.'
        ),
        "instrument_score": "[auto]",
        "comments": "Linked to Art. 11(a) producer minimization obligation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 6(a)-(b): Municipal/District Councils must prevent waste dumping in beaches, sea, waterways, "
            "and other hazardous locations; prohibit open burning and disposal in unlicensed facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 6: "Garantir que os resíduos sólidos urbanos não sejam lançados em praias, no mar..."; '
            '"Assegurar que os resíduos sólidos não sejam depositados ou queimados a céu aberto..."; '
            'Art. 20(2): fines of 240,000 MT; Art. 6(d): annual waste registry.'
        ),
        "instrument_score": "[auto]",
        "comments": "Directly addresses marine/coastal plastic leakage pathways.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 8: All public and private entities managing urban solid waste must prepare and implement "
            "5-year Integrated Urban Solid Waste Management Plans based on the waste hierarchy (Annex I)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 8(1): "devem elaborar e implementar um plano de gestão integrada"; '
            'Art. 8(2): plans approved by Municipal Assemblies/District Governments; '
            'Art. 10(1): annual reporting; Art. 20(2): general fines.'
        ),
        "instrument_score": "[auto]",
        "comments": "Annex I requires plan objectives/targets and awareness campaigns but no national plastic-specific targets.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 9: Treatment and final disposal facilities for urban solid waste require prior "
            "environmental licensing under the Environmental Impact Assessment Regulation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 9(1): "estão sujeitas a prévio licenciamento ambiental"; '
            'licensing process under EIA Regulation; Art. 6(b): unlicensed disposal prohibited.'
        ),
        "instrument_score": "[auto]",
        "comments": "Applies to landfills and treatment plants receiving plastic fractions.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 10: Municipalities submit annual waste management registers to the Ministry; maintain "
            "entity registries; mandatory 24-hour notification of accidental waste spills."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 10(1): annual register "até ao final do primeiro trimestre de cada ano"; '
            'Art. 10(3)-(4): 24-hour spill reporting; Art. 5(1)(d): information access; '
            'Annex II includes recycling percentages/tonnage.'
        ),
        "instrument_score": "[auto]",
        "comments": "Annex II explicitly requires data on selective collection and recycling tonnage.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 11: Binding obligations on producers, transporters and operators to minimize waste, "
            "segregate/condition waste per Art. 14, treat waste before final disposal, ensure safe transport "
            "without dispersal, and maintain annual waste records."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 11: "São obrigações dos produtores, transportadores e operadores..."; '
            'Art. 11(c): segregation per Art. 14 (includes plastic); Art. 20(2): fines of 240,000 MT; '
            'Art. 11(h): annual records.'
        ),
        "instrument_score": "[auto]",
        "comments": "Closest analogue to producer responsibility; applies to all waste categories including plastics.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 12: Municipal/District Councils establish collection routes, frequency and schedules; "
            "transport must use appropriate vehicles to minimize environmental and public health risks."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 12(1): methods "serão estabelecidos pelos Conselhos Municipais ou Governos Distritais"; '
            'Art. 12(3): "O transporte de resíduos deve ser feito em veículos apropriados"; '
            'Art. 11(f): no dispersal along transport routes.'
        ),
        "instrument_score": "[auto]",
        "comments": "Municipal implementation discretion on collection systems.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 13: Selective collection systems must be approved by municipalities, separate waste per "
            "Art. 14 categories, promote cooperatives of recyclable material collectors, and be implemented "
            "by municipalities, private sector, or cooperatives."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 13(1): "O sistema de recolha selectiva deve ser aprovado pelos Conselhos Municipais..."; '
            'Art. 13(2): promote cooperatives; Annex II: selective collection reporting; Art. 20(2): fines.'
        ),
        "instrument_score": "[auto]",
        "comments": "Primary instrument enabling plastic recycling at source; plastic is an explicit Art. 14 category.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 14: Mandates segregation of urban solid waste into categories including plastic (d); "
            "requires producers/handlers to provide adequate leak-proof conditioning containers with clear identification."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 14(1)(d): "Plástico" as mandatory segregation category; '
            'Art. 14(2)-(3): adequate conditioning to prevent dispersal; '
            'Art. 11(c): linked mandatory obligation; Art. 20(2): fines.'
        ),
        "instrument_score": "[auto]",
        "comments": "Only article explicitly naming plastics as a segregated waste category. Art. 1 also defines contaminated plastics as special waste.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 15: Municipalities establish and approve treatment and valorization systems specifying "
            "mechanical, physical, thermal, chemical or biological processes, and forms of reuse, recycling, "
            "material recovery, or energy co-processing."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 15(1): systems "são estabelecidos e aprovados pelos Conselhos Municipais ou Governos Distritais"; '
            'Art. 15(2): must specify valorization processes including recycling.'
        ),
        "instrument_score": "[auto]",
        "comments": "Municipal-level planning instrument for recycling infrastructure.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 16: Final disposal must follow Ministry operational norms and occur only in sanitary "
            "or controlled landfills."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 16(1)-(2): "deposição final... em aterros sanitários ou controlados"; '
            'Art. 9: licensing required; Art. 20(2): fines for violations of Art. 16.'
        ),
        "instrument_score": "[auto]",
        "comments": "Infrastructure standard for residual waste including non-recycled plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 17: Municipal/District Councils responsible for post-closure maintenance and environmental "
            "monitoring of landfills per Ministry-approved closure plans."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 17(1): municipal responsibility for "manutenção e... monitoria ambiental após o encerramento"; '
            'Art. 17(2): closure plan approved by Ministry; Art. 20(2): fines.'
        ),
        "instrument_score": "[auto]",
        "comments": "Addresses long-term leakage risk from landfilled plastics.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 18: Municipalities must promote environmental education on waste reduction, pollution "
            "prevention, reuse and recycling, involving communities, schools, media, and civil society."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 18(a)-(b): "Promover programas educativos de consciencialização pública"; '
            '"Proceder a divulgação de boas práticas de gestão de resíduos sólidos urbanos".'
        ),
        "instrument_score": "[auto]",
        "comments": "Awareness instrument; no dedicated enforcement mechanism in Art. 18 itself.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 19: Municipal/District Councils establish and collect urban cleaning fees for collection, "
            "transport, treatment and disposal of urban solid waste via municipal by-laws."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Art. 19(1): "As taxas de limpeza urbana são estabelecidas e cobradas pelos Conselhos Municipais '
            'ou Governos Distritais"; Art. 5(2)(e): municipal authority to set fees.'
        ),
        "instrument_score": "[auto]",
        "comments": "Economic instrument; fee levels set locally, not nationally ring-fenced.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 20-22: Administrative fines (150,000 MT and 240,000 MT) for obstructing inspections and "
            "violating key provisions; 30% increase for repeat offences; 60% of national fines allocated to FUNAB."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'Art. 20: specific fine amounts; Art. 21: 20-day payment period, fiscal execution; '
            'Art. 22: "60% para o FUNAB"; Art. 5(1)(h), 5(2)(i): penalty authority designated.'
        ),
        "instrument_score": "[auto]",
        "comments": "Art. 20(2) explicitly covers violations of Arts. 4, 6, 11, 16, 17.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 25: Municipalities and non-municipalized towns must close open-air dumps and ensure "
            "construction of sanitary or controlled landfills within 3 years of publication."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'Art. 25: "no prazo de três anos após a publicação do presente regulamento"; '
            'Art. 24: temporary fine exemption for municipalities showing diligence; '
            'Art. 5(1)(b): Ministry inspections.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Quantifiable 3-year infrastructure target. Art. 24 creates conditional exemption reducing "
            "unconditionality score. Deadline was March 2017 (90 days after 31 Dec 2014 publication + 3 years)."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 1 (Resíduos especiais): Defines contaminated plastics among special household hazardous "
            "waste requiring differentiated management."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Art. 1: "Resíduos especiais – resíduos com características perigosas... tais como... plásticos '
            'contaminados e outros"; Art. 14(1)(j): special waste category.'
        ),
        "instrument_score": "[auto]",
        "comments": "Explicit mention of plastics; classification instrument enabling differentiated handling.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Decreto Art. 2: Authorises the Environment Minister to approve general/specific directives and "
            "implementation norms for the Regulation (enabling power)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'Decreto Art. 2: "Compete ao Ministro que superintende o Sector do Ambiente aprovar as directivas '
            'gerais e específicas e outras normas para a implementação do presente Regulamento."'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Enabling power using 'compete... aprovar' — scored in_force=0 unless subordinate directives "
            "are confirmed operationalised. No implementing directives cited in this decree text."
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
    rows = []
    for inst in INSTRUMENTS:
        row = {**POLICY_FIELDS, **inst}
        rows.append(row)
    return pd.DataFrame(rows, columns=COLUMNS)


def export_csv(df):
    path = os.path.join(OUTPUT_DIR, "4P_Index_Mozambique_Decree_94_2014.csv")
    df.rename(columns=COLUMN_LABELS).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def export_xlsx(df):
    path = os.path.join(OUTPUT_DIR, "4P_Index_Mozambique_Decree_94_2014.xlsx")
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


def add_para(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = bold
    return p


def create_translation_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading(
        "Regulation on Urban Solid Waste Management (Decree No. 94/2014)",
        0,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Regulamento sobre a Gestão de Resíduos Sólidos Urbanos "
        "(Decreto n.º 94/2014)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}")

    doc.add_paragraph(
        "Source: Boletim da República, I Série, No. 105, 31 December 2014. "
        "Official URL: https://faolex.fao.org/docs/pdf/moz148514.pdf"
    )

    add_heading(doc, "Decree No. 94/2014", 1)
    add_para(doc, "of 31 December", bold=True)
    doc.add_paragraph(
        "Given the need to review the norms and procedures relating to the proper management "
        "of urban solid waste resulting from human activities, given the harmful consequences "
        "that poor management entails for public health and the environment, pursuant to the "
        "provisions of Article 33 of the Environment Law, the Council of Ministers decrees:"
    )
    add_para(doc, "Article 1", bold=True)
    doc.add_paragraph(
        "The Regulation on the Management of Urban Solid Waste and its annexes, which form "
        "an integral part of this Decree, are hereby approved."
    )
    add_para(doc, "Article 2", bold=True)
    doc.add_paragraph(
        "It is the responsibility of the Minister overseeing the Environment Sector to approve "
        "general and specific directives and other norms for the implementation of this Regulation."
    )
    add_para(doc, "Article 3", bold=True)
    doc.add_paragraph(
        "Decree No. 13/2006 of 15 June, which approved the Regulation on Waste Management, is hereby repealed."
    )
    add_para(doc, "Article 4", bold=True)
    doc.add_paragraph("This decree shall enter into force ninety days after its publication.")
    doc.add_paragraph(
        "Approved by the Council of Ministers on 11 November 2014.\n"
        "Published.\n"
        "The Prime Minister, Alberto Clementino António Vaquina."
    )

    add_heading(doc, "Regulation on the Management of Urban Solid Waste", 1)

    add_heading(doc, "Chapter I — General Provisions", 2)
    add_para(doc, "Article 1 — Definitions", bold=True)
    doc.add_paragraph("For the purposes of this regulation, the following are defined:")
    definitions = [
        ("Packaging/Conditioning", "Placement of waste in containers with conditions of tightness and hygiene to prevent its dispersal."),
        ("Recovery or Valorization", "Use of waste or its components through recycling or reuse processes aimed at obtaining secondary raw materials for reintroduction into production and/or consumption circuits in analogous use, without alteration."),
        ("Aquaculture", "Production of aquatic organisms, such as fish, crustaceans, amphibians, reptiles, and cultivation of aquatic plants for human use."),
        ("Storage", "Temporary and controlled deposition of waste prior to treatment, recovery or disposal."),
        ("Sanitary landfill", "Infrastructure whose purpose is the safe deposition of urban solid waste in soil, using engineering principles to eliminate environmental impacts and confine waste in the smallest possible volume."),
        ("Controlled landfill", "Infrastructure for deposition of waste in soil according to management plans that lacks leachate control, impermeabilization and gas management systems."),
        ("Composting", "Method for decomposition of organic material in waste under adequate conditions to obtain organic compost."),
        ("Compost", "Fertilizing matter resulting from controlled decomposition of organic waste, obtained by composting or anaerobic digestion followed by composting."),
        ("Environmentally adequate final disposal", "Placement of waste in sanitary landfills, observing specific operational norms to avoid harm or risks to public health and safety and to minimize adverse environmental impacts."),
        ("Waste holder", "Person or entity that controls or holds waste in its possession."),
        ("Sorting stations", "Infrastructure where waste is separated, by manual or mechanical processes, into materials destined for valorization."),
        ("Transfer stations", "Transitory installations to consolidate, prepare and transport waste to treatment, valorization or final disposal sites."),
        ("Waste management", "All viable procedures to ensure environmentally safe, sustainable and rational waste management, taking into account the need for reduction, recycling and reuse, including separation, collection, handling, transport, storage and/or disposal of waste as well as subsequent protection of disposal sites, in order to protect human health and the environment against harmful effects."),
        ("Risk management", "Systematic identification of hazards and development of control measures to manage risks associated with each identified hazard."),
        ("Incineration", "Controlled burning of solid waste in furnaces designed to fully transform waste into inert material, also achieving volume and weight reduction."),
        ("Waste operator", "Entity that carries out activities related to waste management."),
        ("Hazard", "Potential to degrade environmental quality, harm health and life of persons or damage property."),
        ("Integrated Urban Solid Waste Management Plan", "Document containing systematized technical information on collection, transport, handling, storage, treatment, valorization or disposal operations, including monitoring of discharge sites during and after closure of respective installations, as well as planning of those operations."),
        ("Waste producer", "Natural or legal persons, public or private, that generate waste as a result of their activities."),
        ("Recycling", "Process of transforming solid waste involving alteration of physical, physicochemical or biological properties, with a view to transforming them into inputs or new products."),
        ("Collection", "Collection operation including sorting of waste for transport."),
        ("Selective collection", "Differentiated collection system from the generating source of waste segregated into its various components."),
        ("Waste", "Substances or objects that are discarded, intended to be discarded, or that we are legally obliged to discard, also designated as refuse."),
        ("Special waste", "Waste with hazardous characteristics produced in households in small quantities such as electrical and electronic equipment, used oils, contaminated plastics and others."),
        ("Biowaste", "Biodegradable waste from green spaces, namely gardens, parks, sports fields, as well as biodegradable food waste from households, meal supply units or similar waste from food processing units."),
        ("Bulky household waste", "Waste from households whose removal is not possible by normal means due to volume, shape or dimensions, or whose deposition in existing containers is considered inconvenient by the Municipality."),
        ("Commercial solid waste", "Waste of commercial origin with characteristics of domestic solid waste, such as from commercial establishments, offices, restaurants and similar."),
        ("Industrial solid waste equated to urban waste", "Industrial waste with characteristics of domestic urban solid waste such as from canteens and offices."),
        ("Hospital solid waste equated to urban waste", "Waste from hospital units with characteristics of domestic urban solid waste such as from canteens and offices."),
        ("Urban solid waste", "Waste originating from domestic and commercial activities in population agglomerations."),
        ("Risk", "Probability of occurrence of an accident and the consequences resulting from that occurrence."),
        ("Segregation", "Processes of separation of urban solid waste based on constituent materials for subsequent recycling, composting, incineration and final disposal."),
        ("Waste transport", "Any physical transfer operation of waste using road, rail, air or maritime means."),
        ("Waste treatment", "Any valorization or disposal operation, including preparation prior to valorization or disposal, comprising mechanical, physical, thermal, chemical or biological processes that alter waste characteristics to reduce volume or hazardousness."),
    ]
    for term, definition in definitions:
        p = doc.add_paragraph(style="List Bullet")
        p.add_run(f"{term} — ").bold = True
        p.add_run(definition)

    add_para(doc, "Article 2 — Object", bold=True)
    doc.add_paragraph(
        "This Regulation has as its object the establishment of rules for the management of urban solid waste on national territory."
    )
    add_para(doc, "Article 3 — Scope of Application", bold=True)
    doc.add_paragraph(
        "1. This regulation applies to all natural and legal persons, public and private, involved in:\n"
        "a) The production and management of urban solid waste;\n"
        "b) The production and management of industrial and hospital waste equated to urban waste.\n"
        "2. The rules established by this regulation do not apply to the management of:\n"
        "a) Hazardous industrial waste;\n"
        "b) Biomedical waste;\n"
        "c) Radioactive waste;\n"
        "d) Emissions and effluent discharges;\n"
        "e) Wastewater;\n"
        "f) Other waste subject to specific regulation."
    )
    add_para(doc, "Article 4 — General principles of waste management", bold=True)
    doc.add_paragraph(
        "Under this regulation, the general principles of waste management are as follows:\n"
        "a) Self-sufficiency — waste management operations should preferably take place on national territory, minimizing cross-border movements;\n"
        "b) Management responsibility — waste management is part of the material life cycle and is the responsibility of the respective producer and/or holder;\n"
        "c) Prevention and reduction — priority objective to avoid and reduce waste production and harmful character;\n"
        "d) Waste hierarchy — management must respect priority order: prevention/reduction, reuse, recycling, other forms of valorization, and disposal, using best available technologies at economically sustainable cost;\n"
        "e) Citizen responsibility — citizens must adopt preventive behaviors and practices facilitating reuse and valorization;\n"
        "f) Protection of human health and the environment — priority to avoid and reduce risks, ensuring processes do not generate adverse environmental effects including water, air, soil pollution, impacts on fauna and flora, noise, odors or landscape damage;\n"
        "g) Polluter pays — the polluter must bear costs of repairing environmental damage caused."
    )
    add_para(doc, "Article 5 — Competencies in Urban Solid Waste Management", bold=True)
    doc.add_paragraph(
        "1. The Ministry overseeing the Environment Sector shall: issue rules; conduct inspections; ensure institutional involvement in inspections; "
        "guarantee access to information; promote good practices (recycling, composting, selective collection, sanitary landfills); maintain a National "
        "Register of entities handling urban solid waste; suspend illegal or dangerous storage/transport/disposal; penalize municipal managers for "
        "inadequate management; and monitor compliance.\n"
        "2. Municipal and District Councils shall, within their jurisdiction: ensure adequate waste management; adopt municipal by-laws and regulations; "
        "define collection, transport, treatment and final destination procedures; promote good practices; set fees for collection, transport, treatment "
        "and disposal; register entities; suspend illegal operations; ensure compliance; and penalize offenders."
    )
    add_para(doc, "Article 6 — Obligations of Municipal and District Councils", bold=True)
    doc.add_paragraph(
        "Municipal and District Councils must:\n"
        "a) Ensure urban solid waste is not dumped on beaches, at sea, in watercourses, or other hazardous locations;\n"
        "b) Ensure waste is not deposited or burned in the open air or in unlicensed facilities;\n"
        "c) Ensure compliance with producer, transporter and operator obligations under Article 11;\n"
        "d) Maintain an annual register of origins, quantities and types of waste handled, transported, treated, valorized or disposed of (Annex II);\n"
        "e) Ensure compliance with other provisions of this Regulation."
    )
    add_para(doc, "Article 7 — Classification of Urban Solid Waste", bold=True)
    doc.add_paragraph(
        "Urban solid waste is classified according to Mozambican Standard NM339 — Solid Waste — Classification."
    )

    add_heading(doc, "Chapter II — Management of Urban Solid Waste", 2)
    add_para(doc, "Article 8 — Integrated Urban Solid Waste Management Plan", bold=True)
    doc.add_paragraph(
        "1. All public and/or private entities carrying out activities related to urban solid waste management must prepare and implement "
        "an integrated management plan based on the waste hierarchy (Article 4(d)) containing at minimum the information in Annex I.\n"
        "2. Plans are valid for five (5) years from approval by Municipal Assemblies or District Governments and may be updated when justified."
    )
    add_para(doc, "Article 9 — Environmental licensing of treatment and final disposal facilities", bold=True)
    doc.add_paragraph(
        "1. Facilities for treatment and final disposal of urban solid waste are subject to prior environmental licensing under the Environmental Impact Assessment Regulation.\n"
        "2. Licensing requests shall be submitted to competent bodies under the EIA Regulation.\n"
        "3. Assessment shall be conducted under the EIA Regulation."
    )
    add_para(doc, "Article 10 — Duty of information", bold=True)
    doc.add_paragraph(
        "1. Municipal/District Councils must submit annual waste management registers to the Ministry by end of Q1 each year (Annex II).\n"
        "2. Annual cadastre of entities handling solid waste must be provided to the Ministry.\n"
        "3. Entities must inform Councils of accidental spills within 24 hours and keep Councils informed of measures taken.\n"
        "4. Councils must inform the Ministry of spills within 24 hours of receiving notification."
    )
    add_para(doc, "Article 11 — Obligations of producers, transporters and operators", bold=True)
    doc.add_paragraph(
        "a) Minimize production of urban solid waste;\n"
        "b) Train workers on health, occupational safety and environment;\n"
        "c) Ensure segregation and conditioning per Article 14;\n"
        "d) Ensure treatment before adequate final disposal;\n"
        "e) Protect workers against contamination risks;\n"
        "f) Ensure adequate transport without dispersal along routes;\n"
        "g) Ensure disposal does not negatively impact environment or public health;\n"
        "h) Maintain detailed annual records of waste handled."
    )
    add_para(doc, "Article 12 — Collection and transport", bold=True)
    doc.add_paragraph(
        "1. Collection and transport methods shall be established by Municipal/District Councils.\n"
        "2. Competent entities may adopt technically appropriate systems per situation and waste type.\n"
        "3. Transport must use appropriate vehicles to minimize risks.\n"
        "4. Collection shall follow routes, frequency and schedules approved by Councils.\n"
        "5. Councils shall inform residents of placement and collection times."
    )
    add_para(doc, "Article 13 — Selective collection", bold=True)
    doc.add_paragraph(
        "1. Selective collection systems must be approved by Councils and separate waste per Article 14 categories.\n"
        "2. Systems must promote participation of cooperatives or associations of collectors of reusable and recyclable materials.\n"
        "3. Implementation shall be carried out by Councils, private sector, or cooperatives/associations."
    )
    add_para(doc, "Article 14 — Segregation and conditioning", bold=True)
    doc.add_paragraph(
        "1. Urban solid waste shall be segregated into:\n"
        "a) Organic matter; b) Paper or cardboard; c) Rubble; d) Plastic; e) Glass; f) Metal; g) Textiles; h) Rubber; "
        "i) Bulky household waste; j) Special waste.\n"
        "2. Producing/handling entities must provide adequate conditioning to prevent dispersal.\n"
        "3. Conditioning forms must allow clear identification of containers and locations per category."
    )
    add_para(doc, "Article 15 — Treatment and valorization", bold=True)
    doc.add_paragraph(
        "1. Treatment and valorization systems are established and approved by Municipal/District Councils.\n"
        "2. Systems must specify processes (mechanical, physical, thermal, chemical, biological) and forms of reuse, recycling, material recovery or energy co-processing."
    )
    add_para(doc, "Article 16 — Final disposal", bold=True)
    doc.add_paragraph(
        "1. Final disposal must follow Ministry operational norms to avoid harm to health, safety and environment.\n"
        "2. Final disposal shall be carried out in sanitary or controlled landfills."
    )
    add_para(doc, "Article 17 — Closure of old dumps and sanitary landfills", bold=True)
    doc.add_paragraph(
        "1. Municipal/District Councils are responsible for maintenance and environmental monitoring after closure.\n"
        "2. Maintenance and monitoring shall follow a closure plan approved by the Ministry."
    )
    add_para(doc, "Article 18 — Environmental education", bold=True)
    doc.add_paragraph(
        "Councils must promote educational programs on adequate waste management emphasizing reduction, pollution prevention, reuse and recycling; "
        "disseminate good practices involving communities, schools, universities, media, private sector and civil society; and publish urban cleaning activity calendars."
    )

    add_heading(doc, "Chapter III — Fees, Infractions and Penalties", 2)
    add_para(doc, "Article 19 — Fees", bold=True)
    doc.add_paragraph(
        "1. Urban cleaning fees are established and collected by Councils per municipal by-laws and paid at respective treasuries.\n"
        "2. Use of fee revenues is determined by approved municipal by-laws/regulations."
    )
    add_para(doc, "Article 20 — Infractions and penalties", bold=True)
    doc.add_paragraph(
        "1. Obstructing inspection activities without just cause is punishable by a fine of 150,000.00 MT.\n"
        "2. Non-compliance with Articles 4, 6, Article 11(d)-(h), 16 and 17 is punishable by a fine of 240,000.00 MT.\n"
        "3. Fines increase by 30% cumulatively for repeat offences."
    )
    add_para(doc, "Article 21 — Collection of fines", bold=True)
    doc.add_paragraph(
        "1. Ministry fines are paid at the State Treasury.\n"
        "2. Offenders have 20 calendar days to pay from notification.\n"
        "3. Unpaid fines are sent to competent Fiscal Execution Court.\n"
        "4. Council fines are determined by Municipal Assemblies or District Governments."
    )
    add_para(doc, "Article 22 — Destination of fine revenues", bold=True)
    doc.add_paragraph(
        "1. Fines under Article 20(1): 40% to State Budget; 60% to FUNAB (Environment Fund).\n"
        "2. 60% of FUNAB receipts must be applied to promoting good waste management practices and improving monitoring.\n"
        "3. Destination of council fines determined by Assemblies."
    )
    add_para(doc, "Article 23 — Updating fine and fee values", bold=True)
    doc.add_paragraph(
        "1. National fine/fee values updated by joint ministerial diploma of Finance and Environment Ministers when necessary.\n"
        "2. Municipal fine/fee values updated by Municipal Assemblies or District Governments."
    )

    add_heading(doc, "Chapter IV — Final Provisions", 2)
    add_para(doc, "Article 24 — Temporary exemptions", bold=True)
    doc.add_paragraph(
        "Councils demonstrating necessary diligence and progress in closing dumps and constructing sanitary landfills are temporarily exempt from fines for final disposal in dumps."
    )
    add_para(doc, "Article 25 — Transitional provision", bold=True)
    doc.add_paragraph(
        "Municipal Councils or non-municipalized towns must close open-air dumps in their jurisdictions and ensure construction of sanitary or controlled landfills "
        "within three years of publication of this regulation."
    )

    add_heading(doc, "Annex I — Minimum Requirements of an Integrated Urban Solid Waste Management Plan", 2)
    doc.add_paragraph(
        "Plans must describe current waste management situation and measures to improve environmentally correct treatment and disposal. "
        "Minimum elements: (a) characterization of Municipality/District; (b) objectives and targets for the five-year plan period; "
        "(c) organizational aspects including responsibility sharing and revenue options; (d) current waste management situation; "
        "(e) SWOT analysis; (f) proposals for adequate management; (g) awareness campaign actions; (h) annexes; (i) bibliography."
    )

    add_heading(doc, "Annex II — Annual Urban Solid Waste Register Form", 2)
    doc.add_paragraph(
        "Information to be produced by Municipal/District Councils includes: municipality/district name; contacts; NUIT; data on involved entities; "
        "total population; estimated waste production (tonnes/year); per capita generation (kg/year); collection coverage rate; waste deposited in landfill (tonnes/year); "
        "selective collection information; percentage and tonnage destined for recycling; main treatment forms; main final disposal methods; waste management accidents."
    )

    doc.add_page_break()
    add_heading(doc, "4P Index Coding Summary", 1)
    doc.add_paragraph(
        f"Policy coded: {POLICY_FIELDS['policy_name']}\n"
        f"Year: {POLICY_FIELDS['policy_year']}\n"
        f"Instruments identified: {len(INSTRUMENTS)}\n"
        f"Policy URL: {POLICY_FIELDS['policy_url']}\n\n"
        "Full coded table available in accompanying spreadsheet: "
        "4P_Index_Mozambique_Decree_94_2014.xlsx"
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Decree_94_2014_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(csv_path, xlsx_path, docx_path):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Decree 94/2014 — 4P Index Deliverables</title>
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
  <p class="meta">Mozambique — Regulation on Urban Solid Waste Management (Decree No. 94/2014)</p>
  <p>Coded analysis and English translation generated on {date.today().isoformat()}.</p>

  <div class="card">
    <h2>4P Index Coding Table (Excel)</h2>
    <p>Structured coding table with {len(INSTRUMENTS)} instrument rows (columns A–W).</p>
    <a class="download" href="4P_Index_Mozambique_Decree_94_2014.xlsx" download>Download Excel (.xlsx)</a>
  </div>

  <div class="card">
    <h2>4P Index Coding Table (CSV)</h2>
    <p>Same data in comma-separated format for import into databases or R/Stata.</p>
    <a class="download" href="4P_Index_Mozambique_Decree_94_2014.csv" download>Download CSV</a>
  </div>

  <div class="card">
    <h2>English Translation (Word)</h2>
    <p>Full English translation of Decree No. 94/2014 and the Regulation on Urban Solid Waste Management.</p>
    <a class="download" href="Mozambique_Decree_94_2014_English_Translation.docx" download>Download Word (.docx)</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    df = build_dataframe()
    csv_path = export_csv(df)
    xlsx_path = export_xlsx(df)
    docx_path = create_translation_doc()
    html_path = create_index_html(csv_path, xlsx_path, docx_path)
    print("Generated files:")
    for p in [csv_path, xlsx_path, docx_path, html_path]:
        print(f"  {p}")


if __name__ == "__main__":
    main()
