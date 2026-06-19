#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Decreto 16/2015 only.

Source: user-uploaded Boletim da República (Regulation_on_Management_and_Control_of_Plastic_Bags_1994.pdf).
Regulamento sobre a Gestão e Controlo do Saco de Plástico — Decreto n.º 16/2015, de 5 de Agosto.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_Decreto_16_2015.xlsx")
WORD_OUT = Path("/workspace/Decreto_16_2015_English_Translation.docx")

COLUMNS = [
    "A_policy_name",
    "B_policy_url",
    "C_policy_year",
    "D_policy_objective",
    "E_policy_target",
    "F_policy_target_text",
    "G_policy_type",
    "H_policy_type_justification",
    "I_policy_integration",
    "J_policy_sectors_list",
    "K_policy_circularity",
    "L_policy_lifecycle_phases_list",
    "M_policy_budget",
    "N_policy_budget_text",
    "O_policy_score",
    "P_instrument_type",
    "Q_instrument_lifecycle_stage",
    "R_instrument_description",
    "S_instrument_in_force",
    "T_instrument_implementation",
    "U_instrument_implementation_text",
    "V_instrument_score",
    "W_comments",
]

POLICY = {
    "policy_name": (
        "Regulamento sobre a Gestão e Controlo do Saco de Plástico "
        "(Decreto n.º 16/2015, de 5 de Agosto)"
    ),
    "policy_url": "not available",
    "policy_year": 2015,
    "policy_objective": (
        "Establish norms and procedures for the management and control of plastic bags "
        "regarding their production, import, marketing and use, to reduce negative impacts "
        "on human health, infrastructure, biodiversity and the environment, principally "
        "due to non-biodegradability."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree (Decreto n.º 16/2015) approved by the Council of Ministers "
        "on 30 June 2015 and published 5 August 2015, under Lei 20/97 Art. 33 and "
        "Constitution Art. 204. Sub-legislative instrument, not parliamentary legislation."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "production, consumption, retail, waste management, industry, municipalities, "
        "packaging, fisheries"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        "Art. 8: Fines from violations are allocated 40% to State Budget, 30% to the "
        "Environment Fund (Fundo do Ambiente), and 30% to the inspecting entity. "
        "Art. 8(2): Environment Minister approves percentage of Environment Fund "
        "proceeds channelled to improving fiscalization services."
    ),
    "instruments": [
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 4(1)(a): Prohibits production, import, and wholesale or retail "
                "marketing of plastic bags with thickness below 30 micrometres. "
                "Exceptions: bags for weighing food products; bags for conditioning "
                "municipal solid waste; bags produced in export free zones (Art. 4(2)–(3))."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'é proibida... saco de plástico cuja espessura seja inferior a 30 micrómetros.' "
                "Fines: production 40, import 80, retail/wholesale 50 minimum wages (Art. 7). "
                "Authorities: MAAP, Finance, Industry, INAE, municipalities (Art. 6). "
                "Exemptions for food-weighing and municipal waste bags reduce unconditional score."
            ),
            "comments": (
                "Core plastic-bag regulatory instrument. Decreto Art. 3(2): production/import "
                "prohibitions for in-process applications enter force 180 days after publication "
                "(February 2016). Explicit plastics mention (filter criterion a)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 4(1)(b): Prohibits free distribution of plastic bags at all premises "
                "where commercial activity is exercised."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'é proibida... A distribuição gratuita de saco de plástico em todos os locais "
                "onde se exerça a actividade comercial.' Fine: 25 minimum wages (Art. 7(d)). "
                "INAE and municipalities responsible for fiscalization (Art. 6)."
            ),
            "comments": "Mandatory prohibition language; directly reduces single-use bag consumption.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 4(1)(c) & Art. 5(2): Prohibits marketing or distribution of plastic bags "
                "containing more than 40% recycled material in establishments selling food "
                "products; food retailers must ensure bags do not exceed 40% recycled content."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'é proibida... saco de plástico que contenham acima de 40% de material reciclado "
                "em estabelecimentos que comercializem produtos alimentares.' Fine: 60 minimum "
                "wages (Art. 7(f)). Labelling must indicate recycled percentage (Art. 5(4))."
            ),
            "comments": (
                "Recycled-content standard with food-safety rationale. Arts. 4 and 5 coded as "
                "one instrument (prohibition + compliance duty)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 5(1): Production and import of plastic bags must comply with Mozambican "
                "Standard NM 596."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'A produção e importação do saco de plástico deve estar em conformidade com a "
                "Norma Moçambicana NM 596.' Industry Ministry licenses production (Art. 6(3)); "
                "INAE fiscalizes (Art. 6(4)); MAAP monitors (Art. 6(1)(b))."
            ),
            "comments": "Mandatory technology/performance standard for plastic bag production.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 5(4): Producers must label plastic bags with company name/logo, physical "
                "address, product characteristics (volume, material, plastic symbol, thickness, "
                "and recycled percentage if applicable)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'o produtor deve rotular o saco de plástico produzido' with mandatory "
                "indications listed. INAE fiscalizes production (Art. 6(4)); MAAP monitors "
                "compliance (Art. 6(1)(b)). No specific fine listed in Art. 7 for labelling "
                "violations."
            ),
            "comments": (
                "Labelling scheme; also supports Information instrument type but Regulatory "
                "scored as mandatory standard (Rule 10: highest type)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 5(4) (final paragraph): Mandatory separate display of the price of plastic "
                "bags relative to product prices at all commercial establishments."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'É de carácter obrigatório a indicação, em separado, o preço do saco de plástico "
                "relativamente ao preço dos produtos.' Fine: 30 minimum wages (Art. 7(e)). "
                "INAE inspection (Art. 6(4))."
            ),
            "comments": "Economic disincentive through price transparency at point of sale.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 6(1): Environment Ministry (MAAP) shall disseminate compliance rules, "
                "monitor regulation compliance, adopt measures to reduce plastic bag use and "
                "identify sustainable alternatives, and ensure environmental norms in production."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Compete ao Ministério que superintende a área do Ambiente' with monitoring "
                "(+0.25), authority designated (+0.25). Dissemination of compliance rules "
                "(+0.25 monitoring mechanism). No penalties in Art. 6 itself."
            ),
            "comments": "National governance/coordination instrument for plastic bag management.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 6(2): Finance Ministry shall fiscalize the import process of plastic bags "
                "through competent bodies."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Compete ao Ministério que Superintende a área das Finanças, fiscalizar através "
                "de órgãos competentes o processo de importação do saco de plástico.' "
                "Authority (+0.25) and import inspection (+0.25 monitoring)."
            ),
            "comments": "Border/import enforcement coordination for plastic bags.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 6(3): Industry and Commerce Ministry shall license production and "
                "marketing activities of plastic bags and register entities that produce, "
                "market and import plastic bags."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Licenciar as actividades de produção e comercialização do saco plástico'; "
                "'Cadastrar entidades que produzem, comercializam e importam sacos plásticos.' "
                "Responsible authority designated (+0.25)."
            ),
            "comments": "Producer/importer registration and licensing framework.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 6(4): National Economic Activities Inspection (INAE) shall inspect "
                "production, marketing and use of plastic bags."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'Compete a Inspecção Nacional de Actividades Económicas fiscalizar a produção, "
                "comercialização e uso do saco plástico.' Authority (+0.25), enforcement via "
                "Art. 7 fines (+0.25), monitoring/inspection (+0.25). Receives 30% of fines "
                "(Art. 8)."
            ),
            "comments": "Primary enforcement body for retail and production compliance.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 6(5): Local State bodies and Municipal Councils shall, within their "
                "competences, ensure compliance with this Regulation."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Compete aos Órgãos Locais do Estado e Conselhos Municipais... velar pelo "
                "cumprimento do presente Regulamento.' Municipal authority designated (+0.25)."
            ),
            "comments": (
                "Multi-level governance coordination (Rule 11); municipal waste-bag exception "
                "in Art. 4(2) links instrument to municipal waste streams."
            ),
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 7: Fine schedule for violations — production below 30 µm (40 min. wages), "
                "import (80), retail/wholesale (50), free distribution (25), failure to display "
                "separate price (30), >40% recycled in food retail (60). Repeat offences tripled. "
                "Payment within 20 days or coercive collection."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'As transgressões às disposições deste Regulamento ficam sujeitas às seguintes "
                "multas.' Explicit fines (+0.25 enforcement). Payment to State Treasury within "
                "20 days with coercive collection (+0.25). Triple repeat offences (+0.25). "
                "Inspecting entities designated via Art. 6 (+0.25)."
            ),
            "comments": "Economic enforcement instrument; fine levels are quantified but are penalties not policy targets (Col E = 0).",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 8: Fine revenue allocated 40% to State Budget, 30% to Environment Fund, "
                "30% to inspecting entity; Environment Minister determines share of Fund proceeds "
                "for improving fiscalization services."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'40% para o Orçamento do Estado; 30% para o Fundo do Ambiente; 30% para a "
                "Entidade Fiscalizadora.' Ring-fenced revenue allocation (+0.25). Environment "
                "Minister authority over Fund disbursement (+0.25)."
            ),
            "comments": "Supports policy_budget = 1 (dedicated Environment Fund allocation from fines).",
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 6(1)(c): Environment Ministry shall adopt coordinated measures necessary "
                "to reduce plastic bag use and identify sustainable alternatives."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Adoptar em coordenação com os outros sectores, medidas necessárias para a "
                "redução do uso de saco de plástico e identificar alternativas sustentáveis.' "
                "Authority designated (+0.25); measures not specified — discretionary programme."
            ),
            "comments": (
                "Information/voluntary-type programme instrument; complements regulatory bans. "
                "No quantifiable reduction target stated (hence Col E = 0)."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Decreto Art. 2: Environment Minister, heard Industry Minister, may approve "
                "complementary norms for implementation of the Regulation."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Compete ao Ministro que superintende a área do Ambiente... aprovar normas "
                "complementares para a implementação do presente Regulamento.' Enabling power "
                "for subordinate norms (+0.25 authority)."
            ),
            "comments": (
                "In-force = 0 per Rule 12 unless complementary norms evidenced as adopted "
                "(not in this document)."
            ),
        },
    ],
}

ENGLISH_TRANSLATION = """DECREE NO. 16/2015 OF 5 AUGUST 2015
Approving the Regulation on the Management and Control of Plastic Bags

Source document: User-uploaded Boletim da República, Series I, No. 62, 5 August 2015.

Approved by the Council of Ministers on 30 June 2015.
Published 5 August 2015.
The Prime Minister, Carlos Agostinho do Rosário.

---

DECREE OF THE COUNCIL OF MINISTERS

There being a need to establish norms and procedures concerning the production, import, marketing and use of plastic bags in order to reduce their negative impacts on human health, infrastructure, biodiversity and the environment in general, due principally to their non-biodegradability characteristic, under the provisions of Article 33 of Law No. 20/97 of 1 October (Environmental Law), together with Article 204 of the Constitution of the Republic, the Council of Ministers decrees:

Article 1. The Regulation on the Management and Control of Plastic Bags is approved, annexed to this Decree and forming an integral part thereof.

Article 2. It is the responsibility of the Minister overseeing the Environment area, heard the Minister overseeing Industry and Commerce, to approve complementary norms for implementation of this Regulation.

Article 3.
1. This Decree enters into force on the date of its publication.
2. The following provisions enter into force 180 days after publication of this Decree:
   a) Prohibition of import of plastic bags whose processes are underway;
   b) Prohibition of production of plastic bags with the characteristics described in paragraphs 1 and 2 of Article 4 of the Regulation annexed to this Decree;
   c) Prohibition of resale.

Published.

---

REGULATION ON THE MANAGEMENT AND CONTROL OF PLASTIC BAGS

CHAPTER I — GENERAL PROVISIONS

Article 1 (Definitions)

For the purposes of this Regulation:

a) Biodegradable — a substance that decomposes through the action of a biological agent.

b) Management of plastic bags — refers to the control cycle of production, import, marketing, distribution, use and final disposal of plastic bags.

c) Virgin raw material — material used for production of plastic bags not obtained from recycling processes.

d) Recycled material — raw material or material obtained after physical-chemical recycling processes of waste.

e) Micrometre — unit of measure equal to 10⁻⁶ metre (0.000001 metre), i.e. one thousandth of a millimetre, used to measure very thin thicknesses of various materials.

f) Mozambican Standard (NM) — document established by consensus and approved by the National Institute for Standardisation and Quality (INNOQ), providing for common and repeated use rules, guidelines or characteristics for activities or their results.

g) Plastic — solid organic polymers of high molecular weight, synthetic or semi-synthetic, mouldable, produced mainly from petrochemicals or partially from natural products.

h) Plastic bag — a type of plastic bag used to transport or preserve any product.

i) Plastic bag for weighing food products — handleless bag with variable thickness between 5 and 12 micrometres used to package products for specific weighing purposes.

Article 2 (Purpose)

This Regulation has as its object to establish norms and procedures concerning the management and control of plastic bags with respect to their production, import, marketing and use, with a view to reducing negative impacts on human health and the environment in general.

Article 3 (Scope)

This Regulation applies to all public and private entities, natural and legal persons involved in the production, import, marketing and use of plastic bags on national territory.

Article 4 (Prohibitions)

1. Under this Regulation it is prohibited:
   a) Production, import, retail or wholesale marketing of plastic bags with thickness below 30 micrometres;
   b) Free distribution of plastic bags at all premises where commercial activity is exercised;
   c) Marketing or distribution of plastic bags containing more than 40% recycled material in establishments that sell food products.

2. Exception to paragraph 1(a): plastic bags used for weighing food products and those specifically used for conditioning municipal solid waste.

3. Further exception to paragraph 1(a): plastic bags produced in the free export zone.

Article 5 (Production, use and marketing of plastic bags)

1. Production and import of plastic bags must comply with Mozambican Standard NM 596.

2. Establishments or premises dedicated to marketing food products must respect the thickness provided in Article 4(1) and ensure that marketed plastic bags do not exceed 40% recycled material in their composition.

3. Distribution and use of plastic bags with recycled material is authorised at establishments not involving trade in food products, provided thickness requirements in Article 4(1) are respected.

4. Without prejudice to the above Standard, the producer must label produced plastic bags with:
   a) Company name and/or logo;
   b) Physical address;
   c) Product characteristics including volume, material used, plastic symbol, thickness and, if containing recycled material, its percentage.

5. It is mandatory to indicate separately the price of the plastic bag relative to the price of products at all establishments practising commercial activity.

Article 6 (Competencies)

1. The Ministry overseeing the Environment shall:
   a) Disseminate mandatory compliance rules on procedures to observe in plastic bag management;
   b) Monitor compliance with this Regulation;
   c) Adopt, in coordination with other sectors, measures necessary to reduce plastic bag use and identify sustainable alternatives;
   d) Ensure compliance with environmental norms and procedures in the plastic bag production process.

2. The Ministry overseeing Finance shall, through competent bodies, inspect the import process of plastic bags.

3. The Ministry overseeing Industry and Commerce shall:
   a) License production and marketing activities of plastic bags;
   b) Register entities that produce, market and import plastic bags.

4. The National Economic Activities Inspection (INAE) shall inspect production, marketing and use of plastic bags.

5. Local State bodies and Municipal Councils shall, within their competences, ensure compliance with this Regulation.

CHAPTER II — INFRACTIONS AND PENALTIES

Article 7 (Infractions and penalties)

1. Violations of this Regulation are subject to the following fines:
   a) Production of plastic bags with thickness below 30 micrometres — fine of 40 minimum wages;
   b) Import of plastic bags with thickness below 30 micrometres — fine of 80 minimum wages;
   c) Retail or wholesale marketing of plastic bags under 30 micrometres — fine of 50 minimum wages;
   d) Free distribution of plastic bags — fine of 25 minimum wages;
   e) Failure to indicate separately the price of the plastic bag relative to product prices — fine of 30 minimum wages;
   f) Distribution of plastic bags containing more than 40% recycled material in establishments selling food products — fine of 60 minimum wages.

2. Fines must be paid at the Treasury receiving office in the jurisdiction of the establishment within a maximum of 20 days from notification, after which the offender is subject to coercive collection.

3. Repeated commission of acts in the preceding paragraphs is subject to payment of triple the respective fine.

Article 8 (Use of fines)

1. Fines established in this Regulation are allocated as follows:
   a) 40% to the State Budget;
   b) 30% to the Environment Fund;
   c) 30% to the inspecting entity.

2. The Minister overseeing the Environment shall approve the percentage of amounts allocated to the Environment Fund to be channelled to improving fiscalization services.

---

Note: The uploaded filename references '1994' but the document is Decreto 16/2015 published 5 August 2015. Thickness standard (30 µm), recycled-content limit (40% in food retail), and phased entry (180 days) for certain prohibitions are operative regulatory provisions."""


def build_rows() -> list[dict]:
    rows = []
    for inst in POLICY["instruments"]:
        rows.append({
            "A_policy_name": POLICY["policy_name"],
            "B_policy_url": POLICY["policy_url"],
            "C_policy_year": POLICY["policy_year"],
            "D_policy_objective": POLICY["policy_objective"],
            "E_policy_target": POLICY["policy_target"],
            "F_policy_target_text": POLICY["policy_target_text"],
            "G_policy_type": POLICY["policy_type"],
            "H_policy_type_justification": POLICY["policy_type_justification"],
            "I_policy_integration": POLICY["policy_integration"],
            "J_policy_sectors_list": POLICY["policy_sectors_list"],
            "K_policy_circularity": POLICY["policy_circularity"],
            "L_policy_lifecycle_phases_list": POLICY["policy_lifecycle_phases_list"],
            "M_policy_budget": POLICY["policy_budget"],
            "N_policy_budget_text": POLICY["policy_budget_text"],
            "O_policy_score": "[auto]",
            "P_instrument_type": inst["instrument_type"],
            "Q_instrument_lifecycle_stage": inst["instrument_lifecycle_stage"],
            "R_instrument_description": inst["instrument_description"],
            "S_instrument_in_force": inst["instrument_in_force"],
            "T_instrument_implementation": inst["instrument_implementation"],
            "U_instrument_implementation_text": inst["instrument_implementation_text"],
            "V_instrument_score": "[auto]",
            "W_comments": inst.get("comments", ""),
        })
    return rows


def write_excel(rows: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=9)

    ws.append(COLUMNS)
    for col in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for i, row in enumerate(rows, start=2):
        ws.append([row[c] for c in COLUMNS])
        ws[f"O{i}"] = f"=AVERAGE(G{i},I{i},K{i},M{i},P{i})"
        ws[f"V{i}"] = f"=AVERAGE(P{i},T{i})"

    widths = {
        "A": 45, "B": 18, "C": 8, "D": 45, "E": 8, "F": 12, "G": 8, "H": 35,
        "I": 8, "J": 35, "K": 8, "L": 28, "M": 8, "N": 40, "O": 10,
        "P": 8, "Q": 20, "R": 50, "S": 8, "T": 8, "U": 40, "V": 10, "W": 40,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.column in (7, 9, 11, 13, 15, 16, 19, 20, 22):
                cell.number_format = "0.00"

    meta = wb.create_sheet("Metadata")
    meta.append(["Field", "Value"])
    meta.append(["Country", "Mozambique"])
    meta.append(["Index", "4P Index v2"])
    meta.append(["Policy", POLICY["policy_name"]])
    meta.append(["Source", "User upload: Regulation_on_Management_and_Control_of_Plastic_Bags_1994.pdf"])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Total instrument rows", len(rows)])
    meta.append([
        "Note",
        "Columns O and V use Excel formulas. "
        "policy_score = AVERAGE(G,I,K,M,P); instrument_score = AVERAGE(P,T).",
    ])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_translation_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading(
        "Regulamento sobre a Gestão e Controlo do Saco de Plástico (Decreto 16/2015) — English Translation",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English translation of Mozambique's Regulation on the Management and Control of Plastic Bags, "
        "approved by Council of Ministers Decree No. 16/2015 of 5 August 2015, based on the "
        "user-uploaded Boletim da República. This is the sole policy covered by the accompanying "
        "4P Index coding table."
    )

    p = doc.add_paragraph()
    p.add_run("Portuguese title: ").bold = True
    p.add_run("Regulamento sobre a Gestão e Controlo do Saco de Plástico (Decreto n.º 16/2015, de 5 de Agosto)")

    p2 = doc.add_paragraph()
    p2.add_run("Source: ").bold = True
    p2.add_run(
        "User-uploaded document (Regulation_on_Management_and_Control_of_Plastic_Bags_1994.pdf); "
        "policy URL not available."
    )

    doc.add_heading("English translation", level=2)

    for block in ENGLISH_TRANSLATION.strip().split("\n\n"):
        text = block.strip()
        if not text:
            continue
        if text.startswith("---"):
            continue
        if text.startswith(("DECREE", "REGULATION", "CHAPTER")):
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
        elif text.startswith("Article "):
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
        else:
            para = doc.add_paragraph(text)
            para.paragraph_format.space_after = Pt(6)

    doc.save(WORD_OUT)
    print(f"Wrote {WORD_OUT}")


def main() -> None:
    rows = build_rows()
    write_excel(rows)
    write_translation_doc()


if __name__ == "__main__":
    main()
