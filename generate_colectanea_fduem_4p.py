#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for FDUEM Colectânea de Legislação do Ambiente.

Source: user-uploaded compilation (482841231-Colectanea-de-Legislacao-do-Ambiente-FDUEM-pdf_compressed_7cd5.pdf)
Organizer: Carlos Manuel Serra, FDUEM, July 2020. Contains 16 environmental laws/decrees.
"""

from __future__ import annotations

import copy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

from generate_4p_index import CODING as INDEX_CODING
from generate_decreto_16_2015_4p import POLICY as DECRETO_16_2015
from generate_lei_20_97_4p import POLICY as LEI_20_97

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_Colectanea_FDUEM.xlsx")
WORD_OUT = Path("/workspace/Colectanea_FDUEM_English_Translation.docx")

COLUMNS = [
    "A_policy_name", "B_policy_url", "C_policy_year", "D_policy_objective",
    "E_policy_target", "F_policy_target_text", "G_policy_type", "H_policy_type_justification",
    "I_policy_integration", "J_policy_sectors_list", "K_policy_circularity",
    "L_policy_lifecycle_phases_list", "M_policy_budget", "N_policy_budget_text",
    "O_policy_score", "P_instrument_type", "Q_instrument_lifecycle_stage",
    "R_instrument_description", "S_instrument_in_force", "T_instrument_implementation",
    "U_instrument_implementation_text", "V_instrument_score", "W_comments",
]

SOURCE_NOTE = (
    "Source: Colectânea de Legislação do Ambiente (FDUEM, Maputo 2020), "
    "user-uploaded PDF; policy URL not available."
)

# Policies in compilation with no instruments passing plastics relevance filter
EXCLUDED_NO_PLASTIC = [
    "Decreto n.º 19/2007 — Acesso e Partilha de Benefícios de Recursos Genéticos",
    "Decreto n.º 24/2008 — Gestão das Substâncias que Destroem a Camada de Ozono",
    "Decreto n.º 25/2008 — Controlo de Espécies Exóticas Invasivas",
    "Decreto n.º 55/2010 — Banimento do Amianto e seus Derivados",
    "Decreto n.º 71/2014 — Biossegurança de Organismos Geneticamente Modificados",
    "Decreto n.º 23/2018 — REDD+ / Redução de Emissões por Desmatamento",
]

DEC_8_2003 = {
    "policy_name": "Regulamento sobre a Gestão de Resíduos Biomédicos (Decreto n.º 8/2003, de 18 de Fevereiro)",
    "policy_url": "not available",
    "policy_year": 2003,
    "policy_objective": (
        "Establish rules for biomedical waste management to safeguard worker and public health "
        "and minimise environmental impacts, including requirements for plastic bags and "
        "containers used in segregation, storage and disposal of infectious and other clinical waste."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": "Executive decree approved by Council of Ministers, 18 February 2003, under Lei 20/97 Art. 33.",
    "policy_integration": 0.50,
    "policy_sectors_list": "waste management, health, municipalities, industry, environment",
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "consumption, disposal",
    "policy_budget": 0,
    "policy_budget_text": "",
    "instruments": [
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Art. 9: Infectious waste must be segregated in yellow plastic bags or other "
                "plastic bags/impermeable containers stamped with yellow 'Lixo Infeccioso' labels; "
                "containers identified with international infectious-waste symbol."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'O lixo infeccioso deverá ser segregado em sacos plásticos amarelos.' "
                "Colour-coding and labelling mandatory (+0.25 unconditional). "
                "Occupational health monitor designated per unit (+0.25 authority). "
                "Transport and disposal rules with penalties in later articles (+0.25 enforcement)."
            ),
            "comments": "Explicit plastic bag requirements for clinical waste streams.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Art. 12: Common (non-infectious) waste shall be placed in clear transparent "
                "plastic bags, or alternative containers where transparent bags unavailable."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'O lixo comum deverá ser colocado em sacos plásticos claros e transparentes.' "
                "Mandatory segregation standard (+0.25). Health units responsible (+0.25 authority)."
            ),
            "comments": "Plastic bag specification for non-hazardous clinical waste.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Art. 10–11: Sharps/cutting waste stored in rigid containers; pharmaceutical waste "
                "in rigid or recycled plastic pharmaceutical containers with specific colour coding."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Rigid containers and recycled plastic pharmaceutical containers specified. "
                "Segregation and storage standards (+0.25 monitoring via identification system)."
            ),
            "comments": "Plastic containers for sharps and pharmaceutical waste.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Arts. 5–6: Assigns responsibilities for biomedical waste management from generation "
                "to final disposal, including occupational health monitors in health units."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Health Ministry and unit-level monitors designated (+0.25). "
                "Management plan and training requirements (+0.25 monitoring)."
            ),
            "comments": "Governance framework for clinical waste including plastic packaging.",
        },
    ],
}

DEC_25_2011 = {
    "policy_name": "Regulamento sobre o Processo de Auditoria Ambiental (Decreto n.º 25/2011, de 15 de Junho)",
    "policy_url": "not available",
    "policy_year": 2011,
    "policy_objective": (
        "Regulate environmental audit processes as systematic management instruments for "
        "evaluating environmental control and protection systems, applicable to activities "
        "including industrial production and waste-generating facilities such as plastic manufacturing."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": "Executive decree approved by Council of Ministers, 15 June 2011, under Lei 20/97 Arts. 18 and 33.",
    "policy_integration": 0.50,
    "policy_sectors_list": "industry, environment, waste management, municipalities, production",
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "production, disposal",
    "policy_budget": 0,
    "policy_budget_text": "",
    "instruments": [
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 4: Environmental audit object includes pollution control equipment, conformity "
                "with environmental norms, effective/potential pollution levels, and corrective measures "
                "for activities during implementation, deactivation and restoration."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Mandatory audit scope defined (+0.25). Environment Ministry conducts public audits; "
                "entities may conduct private audits (+0.25 authority). "
                "Corrective action and reporting requirements in subsequent articles (+0.25 monitoring)."
            ),
            "comments": "Applies to plastic production facilities and waste management operations.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Decreto Art. 2: Environment Minister shall approve general and specific directives "
                "on environmental auditing and implementation norms."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Compete ao Ministro... aprovar as directivas gerais e específicas sobre a auditoria "
                "ambiental.' Enabling power for subordinate directives."
            ),
            "comments": "Enabling instrument; directives may operationalise audit requirements.",
        },
    ],
}


def _from_index(name_fragment: str) -> dict | None:
    for p in INDEX_CODING:
        if name_fragment in p["policy_name"]:
            return copy.deepcopy(p)
    return None


def _normalize(policy: dict) -> dict:
    p = copy.deepcopy(policy)
    p["policy_url"] = "not available"
    for inst in p["instruments"]:
        inst["comments"] = (inst.get("comments", "") + f" {SOURCE_NOTE}").strip()
    return p


def build_policies() -> list[dict]:
    policies = [
        _normalize(LEI_20_97),
        _normalize(DEC_8_2003),
        _normalize(_from_index("18/2004")),
        _normalize(_from_index("45/2006")),
        _normalize(DEC_25_2011),
        _normalize(_from_index("83/2014")),
        _normalize(_from_index("94/2014")),
        _normalize(DECRETO_16_2015),
        _normalize(_from_index("54/2015")),
        _normalize(_from_index("79/2017")),
    ]
    return [p for p in policies if p is not None]


ENGLISH_TRANSLATION = """COLECTÂNEA DE LEGISLAÇÃO DO AMBIENTE
Environmental Legislation Collection — English Summaries

Source: Colectânea de Legislação sobre o Ambiente, organized by Carlos Manuel Serra, Centro de Direito do Ambiente, da Biodiversidade e da Qualidade de Vida, Faculdade de Direito da Universidade Eduardo Mondlane (FDUEM), Maputo, July 2020.

User-uploaded PDF compilation. This document provides English summaries of all 16 laws included in the collection. The accompanying 4P Index Excel codes instruments for the 10 laws with plastics-relevant provisions (per 4P Index v2 plastics relevance filter). Six laws without plastics-specific instruments are summarised here but excluded from the coding table.

---

1. LEI N.º 20/97, DE 1 DE OUTUBRO — ENVIRONMENTAL LAW

Enacted by the Assembly of the Republic. Defines legal bases for correct environmental management and sustainable development. Key provisions: pollution prohibition (Art. 9), environmental quality standards (Art. 10), protected areas, infrastructure/waste dump restrictions (Art. 14), environmental licensing and EIA (Arts. 15–17), environmental audits (Art. 18), citizen rights, strict liability and insurance (Arts. 25–26), environmental inspection (Arts. 28–30), and complementary regulation by Government (Art. 33). Creates CONDES and PNGA framework.

[4P Index: coded — see instrument rows in Excel]

---

2. DECRETO N.º 8/2003 — BIOMEDICAL WASTE MANAGEMENT REGULATION

Executive decree under Lei 20/97. Establishes rules for management of biomedical waste from health units. Defines waste categories (infectious, anatomical, sharps, pharmaceutical, radioactive, common). Requires segregation: infectious waste in yellow plastic bags; common waste in clear transparent plastic bags; sharps in rigid containers. Assigns occupational health monitors and transport/disposal requirements.

[4P Index: coded — plastic bag/container requirements for clinical waste]

---

3. DECRETO N.º 18/2004 (ALTERADO PELO 67/2010) — ENVIRONMENTAL QUALITY AND EFFLUENT STANDARDS

Sets national parameters for air quality, water quality, effluent emissions, and noise. Requires five-year review cycle. Establishes enforceable concentration limits for industrial emissions applicable to manufacturing including plastic production facilities.

[4P Index: coded — industrial emission standards]

---

4. DECRETO N.º 45/2006 — MARINE AND COASTAL POLLUTION PREVENTION REGULATION

Comprehensive regulation preventing marine pollution from ships, ports, and land-based sources. Includes investigation of violations, sanctions, compensation, contingency planning for spills, and restrictions on discharge of wastes to sea. Directly relevant to marine plastic litter pathways.

[4P Index: coded — marine pollution prevention]

---

5. DECRETO N.º 19/2007 — ACCESS AND BENEFIT-SHARING OF GENETIC RESOURCES

Regulates access to genetic resources and fair benefit-sharing. Not plastics-specific.

[4P Index: not coded — no plastics-relevant instruments per filter]

---

6. DECRETO N.º 24/2008 — OZONE-DEPLETING SUBSTANCES MANAGEMENT

Controls substances that deplete the ozone layer. Not plastics-specific.

[4P Index: not coded]

---

7. DECRETO N.º 25/2008 — INVASIVE ALIEN SPECIES CONTROL

Regulates control of invasive exotic species. Not plastics-specific.

[4P Index: not coded]

---

8. DECRETO N.º 55/2010 — ASBESTOS BAN

Prohibits production, use, import, export and marketing of asbestos and derivatives (with research exceptions). Addresses mineral fibres, not synthetic plastics.

[4P Index: not coded]

---

9. DECRETO N.º 25/2011 — ENVIRONMENTAL AUDIT PROCESS REGULATION

Replaces Decreto 32/2003. Defines environmental audit as systematic documented evaluation of environmental management and control systems. Applies to public and private activities that may influence environmental components during implementation, deactivation and restoration. Public audits by Environment Ministry; private audits by entities.

[4P Index: coded — applies to industrial/waste activities including plastics]

---

10. DECRETO N.º 71/2014 — BIOSAFETY OF GENETICALLY MODIFIED ORGANISMS

Biosafety regulation for GMOs. Not plastics-specific.

[4P Index: not coded]

---

11. DECRETO N.º 83/2014 — HAZARDOUS WASTE MANAGEMENT REGULATION

Regulates hazardous waste generation, transport, treatment and disposal. Includes restrictions on recycling plastic packaging contaminated by agro-toxins; take-back obligations for hazardous packaging; landfill requirements for hazardous residues.

[4P Index: coded — hazardous plastic packaging waste]

---

12. DECRETO N.º 94/2014 — URBAN SOLID WASTE MANAGEMENT REGULATION

National framework for urban solid waste: waste hierarchy, municipal responsibilities, prohibition of dumping waste in beaches/sea/watercourses, integrated waste management plans (PGIRSU), segregation and selective collection including plastics.

[4P Index: coded — municipal solid waste including plastic fractions]

---

13. DECRETO N.º 16/2015 — PLASTIC BAG MANAGEMENT AND CONTROL REGULATION

Prohibits bags below 30 micrometres thickness; bans free distribution; limits recycled content in food retail to 40%; requires NM 596 compliance, labelling, and separate pricing. Fines allocated to State Budget, Environment Fund, and inspecting entity.

[4P Index: coded — primary plastic bag regulation]

---

14. DECRETO N.º 54/2015 — ENVIRONMENTAL IMPACT ASSESSMENT PROCESS REGULATION

Replaces Decreto 45/2004. Categorises projects (A+, A, B, C) with EIA/EAS requirements; annexes list activities requiring assessment. Applies to industrial installations including plastic production.

[4P Index: coded — EIA for plastic-related projects]

---

15. DECRETO N.º 79/2017 — EXTENDED PRODUCER RESPONSIBILITY FOR PACKAGING

Establishes EPR for packaging producers and importers: internal management systems, Packaging Environmental Fee (TAE), packaging normalisation, take-back and valorisation of packaging waste. Hierarchy: prevention, reuse, recycling, recovery, disposal.

[4P Index: coded — packaging/plastic EPR]

---

16. DECRETO N.º 23/2018 — REDD+ IMPLEMENTATION REGULATION

Regulates projects for reducing emissions from deforestation and forest degradation, carbon conservation and enhancement. Forest/climate focus; not plastics-specific.

[4P Index: not coded]

---

Note: For full English translations of individual laws previously uploaded separately, see Lei_20_97_English_Translation.docx and Decreto_16_2015_English_Translation.docx in this repository."""


def build_rows(policies: list[dict]) -> list[dict]:
    rows = []
    for policy in policies:
        for inst in policy["instruments"]:
            rows.append({
                "A_policy_name": policy["policy_name"],
                "B_policy_url": policy["policy_url"],
                "C_policy_year": policy["policy_year"],
                "D_policy_objective": policy["policy_objective"],
                "E_policy_target": policy["policy_target"],
                "F_policy_target_text": policy["policy_target_text"],
                "G_policy_type": policy["policy_type"],
                "H_policy_type_justification": policy["policy_type_justification"],
                "I_policy_integration": policy["policy_integration"],
                "J_policy_sectors_list": policy["policy_sectors_list"],
                "K_policy_circularity": policy["policy_circularity"],
                "L_policy_lifecycle_phases_list": policy["policy_lifecycle_phases_list"],
                "M_policy_budget": policy["policy_budget"],
                "N_policy_budget_text": policy["policy_budget_text"],
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


def write_excel(rows: list[dict], policies: list[dict]) -> None:
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
        "A": 48, "B": 18, "C": 8, "D": 42, "E": 8, "F": 12, "G": 8, "H": 32,
        "I": 8, "J": 32, "K": 8, "L": 28, "M": 8, "N": 35, "O": 10,
        "P": 8, "Q": 18, "R": 48, "S": 8, "T": 8, "U": 38, "V": 10, "W": 38,
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
    meta.append(["Source", "Colectânea de Legislação do Ambiente (FDUEM 2020), user upload"])
    meta.append(["Compilation policies (total)", "16"])
    meta.append(["Policies coded (plastics filter)", len(policies)])
    meta.append(["Instrument rows", len(rows)])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Excluded from coding (no plastics instruments)", "; ".join(EXCLUDED_NO_PLASTIC)])
    meta.append([
        "Note",
        "Columns O and V use formulas. One row per instrument; columns A–O repeated per policy.",
    ])

    toc = wb.create_sheet("Compilation contents")
    toc.append(["#", "Law", "Coded in 4P Index"])
    contents = [
        ("1", "Lei 20/97 — Lei do Ambiente", "Yes"),
        ("2", "Decreto 8/2003 — Resíduos Biomédicos", "Yes"),
        ("3", "Decreto 18/2004 — Padrões de Qualidade Ambiental", "Yes"),
        ("4", "Decreto 45/2006 — Poluição Marinho-Costeira", "Yes"),
        ("5", "Decreto 19/2007 — Recursos Genéticos", "No"),
        ("6", "Decreto 24/2008 — Camada de Ozono", "No"),
        ("7", "Decreto 25/2008 — Espécies Invasivas", "No"),
        ("8", "Decreto 55/2010 — Banimento do Amianto", "No"),
        ("9", "Decreto 25/2011 — Auditoria Ambiental", "Yes"),
        ("10", "Decreto 71/2014 — Biossegurança OGM", "No"),
        ("11", "Decreto 83/2014 — Resíduos Perigosos", "Yes"),
        ("12", "Decreto 94/2014 — Resíduos Sólidos Urbanos", "Yes"),
        ("13", "Decreto 16/2015 — Saco de Plástico", "Yes"),
        ("14", "Decreto 54/2015 — Avaliação de Impacto Ambiental", "Yes"),
        ("15", "Decreto 79/2017 — Responsabilidade Alargada Embalagens", "Yes"),
        ("16", "Decreto 23/2018 — REDD+", "No"),
    ]
    for row in contents:
        toc.append(row)

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(policies)} policies, {len(rows)} instrument rows)")


def write_translation_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading(
        "Colectânea de Legislação do Ambiente (FDUEM 2020) — English Summaries",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English summaries of all 16 environmental laws in the FDUEM legislation collection, "
        "based on the user-uploaded PDF. The accompanying 4P Index codes plastics-relevant "
        "instruments for 10 of the 16 laws."
    )

    p = doc.add_paragraph()
    p.add_run("Portuguese title: ").bold = True
    p.add_run("Colectânea de Legislação sobre o Ambiente (FDUEM, Maputo, 2020)")

    p2 = doc.add_paragraph()
    p2.add_run("Source: ").bold = True
    p2.add_run("User-uploaded compilation PDF; policy URL not available.")

    doc.add_heading("English summaries", level=2)

    for block in ENGLISH_TRANSLATION.strip().split("\n\n"):
        text = block.strip()
        if not text or text.startswith("---"):
            continue
        if text.startswith("COLECTÂNEA") or text[0:2].isdigit():
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
        else:
            doc.add_paragraph(text)

    doc.save(WORD_OUT)
    print(f"Wrote {WORD_OUT}")


def main() -> None:
    policies = build_policies()
    rows = build_rows(policies)
    write_excel(rows, policies)
    write_translation_doc()


if __name__ == "__main__":
    main()
