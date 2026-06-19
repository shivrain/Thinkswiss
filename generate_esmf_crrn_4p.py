#!/usr/bin/env python3
"""Generate 4P Index Excel and English document for CRRN Project ESMF.

Source: https://fe.gov.mz/storage/app/uploads/public/65f/2da/871/65f2da8718c30622471525.pdf
Climate Resilient Roads for the North (P500488) — Environmental and Social Management Framework.
Prepared for ANE, March 2024.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_ESMF_CRRN.xlsx")
WORD_OUT = Path("/workspace/ESMF_CRRN_English_Document.docx")

POLICY_URL = (
    "https://fe.gov.mz/storage/app/uploads/public/65f/2da/871/65f2da8718c30622471525.pdf"
)

COLUMNS = [
    "A_policy_name", "B_policy_url", "C_policy_year", "D_policy_objective",
    "E_policy_target", "F_policy_target_text", "G_policy_type", "H_policy_type_justification",
    "I_policy_integration", "J_policy_sectors_list", "K_policy_circularity",
    "L_policy_lifecycle_phases_list", "M_policy_budget", "N_policy_budget_text",
    "O_policy_score", "P_instrument_type", "Q_instrument_lifecycle_stage",
    "R_instrument_description", "S_instrument_in_force", "T_instrument_implementation",
    "U_instrument_implementation_text", "V_instrument_score", "W_comments",
]

POLICY = {
    "policy_name": (
        "Environmental and Social Management Framework (ESMF) — "
        "Climate Resilient Roads for the North Project (P500488)"
    ),
    "policy_url": POLICY_URL,
    "policy_year": 2024,
    "policy_objective": (
        "Establish procedures for environmental and social assessment, management, and monitoring "
        "of the World Bank–financed Climate Resilient Roads for the North project (Cabo Delgado, "
        "Nampula, Niassa). Includes contractor waste-management obligations, compliance with "
        "Mozambique waste regulations (Decretos 94/2014 and 83/2014), and mitigation of construction "
        "solid waste and litter including plastic packaging debris."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "Project-level Environmental and Social Management Framework prepared for ANE under "
        "World Bank Environmental and Social Framework (ESF). Safeguard programme document, "
        "not national legislation or executive decree."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "production, consumption, waste management, recycling, municipalities, industry, "
        "transport, water, environment, labour, tourism"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "ESMF states total budget for implementation of the ESMF and associated instruments "
        "is US$3,100,000; World Bank project envelope US$125 million. Funding mentioned for "
        "E&S management but not ring-fenced solely to waste/plastic measures."
    ),
    "instruments": [
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§5.2 / ESMP requirements: Contractors must prepare site-specific Contractor "
                "Environmental and Social Management Plans (CESMP) including Waste Management Plan, "
                "approved by ANE, IP and World Bank before construction."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'The ESMP requires the Contractors to prepare site specific Contractor's ESMP (CESMP).' "
                "ANE, IP and WB approval required (+0.25 authority). WMP annex mandatory (+0.25 monitoring "
                "via plan implementation)."
            ),
            "comments": (
                "Governance/planning instrument for project subprojects. Applies to construction "
                "waste streams including plastic packaging (filter criterion c)."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§7.2: E&S screening, scoping and ESIA process per Decreto 54/2015 for project "
                "activities; MTA review and approval before works commence."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Six-step E&S assessment process with MTA approval (Step 5). ANE/IP monitoring during "
                "implementation (Step 6). National EIA regulation referenced (+0.25 authority, +0.25 monitoring)."
            ),
            "comments": "Applies to road/bridge construction with potential plastic-related impacts.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "§5.2.2 (Resources and Waste): Preparation of Waste Management Plan following the "
                "waste hierarchy, supported by staff training, for construction and maintenance phases."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Preparation of Waste Management Plan following the waste hierarchy, supported by "
                "staff training.' Mandatory mitigation measure for inefficient waste management (+0.25). "
                "Staff training (+0.25 monitoring mechanism)."
            ),
            "comments": "Waste hierarchy includes recycling; applies to municipal-style construction waste.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Contractor specifications (Waste Management Plan): Contractor shall submit solid waste "
                "control method statement (bins, clean-up schedule) to Supervision Engineer for approval; "
                "site kept free of litter with litter bins at all workplaces."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'The Contractor shall submit a method statement detailing a solid waste control system.' "
                "'The site shall be kept free of litter.' Supervision Engineer approval (+0.25 authority). "
                "Litter bins and refuse collection (+0.25 monitoring). Mandatory language (+0.25 unconditional)."
            ),
            "comments": "Directly addresses litter and packaging waste on construction sites.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Contractor specifications: Prohibition on burning, on-site burying or dumping of waste; "
                "all solid waste disposed offsite at approved landfill with disposal certificates supplied "
                "to Supervision Engineer."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'No burning, on-site burying or dumping of waste shall occur.' "
                "'All solid waste shall be disposed of offsite at an approved landfill site.' "
                "Certificates of disposal required (+0.25 enforcement). Licensed waste collector (+0.25 authority)."
            ),
            "comments": "Regulatory disposal standard for construction solid waste including plastics.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "Contractor specifications: Random disposal of solid waste in scenery areas strictly "
                "prohibited; waste storage area fenced to prevent wind-blown litter; containers covered, "
                "tip-proof, weatherproof and scavenger-proof."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'Random disposal of solid waste in scenery areas shall be strictly prohibited.' "
                "'The waste storage area shall be fenced off to prevent wind-blown litter.' "
                "Mandatory prohibition (+0.25). Engineering controls (+0.25 monitoring). Supervision (+0.25 authority)."
            ),
            "comments": "Directly prevents plastic litter leakage to environment during construction.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Contractor specifications: Recyclable materials (wooden plates, steel, scaffolding, "
                "packaging material, etc.) shall be collected and separated on-site and re-used or sold "
                "to waste collector for recycling."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Recyclable materials such as... packaging material, etc. shall be collected and "
                "separated on-site.' 'Collected recyclable material shall be re-used for other projects "
                "or sold to waste collector for recycling.' Mandatory separation (+0.25). Licensed collectors (+0.25)."
            ),
            "comments": (
                "Explicit packaging material recycling requirement. Economic/recycling instrument; "
                "highest type scored as Economic 0.60 (recycling incentive via sale to collectors)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Contractor specifications: Hazardous and chemical waste (including bitumen) disposed at "
                "approved hazardous landfill per local legislation; used oil/grease sold to approved recycling "
                "company with disposal certificates."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'All hazardous and chemical waste... shall be disposed of at an approved hazardous landfill "
                "site and in accordance with local legislative requirements.' Disposal certificates (+0.25 "
                "enforcement). Approved recycling companies (+0.25 authority). Spill reporting (+0.25 monitoring)."
            ),
            "comments": "Links to Decreto 83/2014 hazardous waste framework.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.2 legal framework table: Project must comply with Regulation on urban solid waste "
                "management (Decreto 94/2014) and Hazardous Waste Management (Decreto 83/2014), administered "
                "by Ministry of Land and Environment (MTA)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Both regulations listed as 'Yes' applicable with MTA as authority. Incorporation by reference "
                "into project safeguard framework (+0.25 authority)."
            ),
            "comments": (
                "Governance coordination with national waste law; Decreto 94/2014 covers municipal solid "
                "waste including plastic fractions."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§5.2.2: Use of authorized contractors for hazardous and other wastes the project cannot "
                "dispose of safely."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Use of authorized contractors for hazardous and any other wastes which the project "
                "cannot dispose of safely.' Contractor authorization requirement (+0.25 authority)."
            ),
            "comments": "Waste contractor licensing/coordination instrument.",
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§6 / §8.3: Stakeholder engagement, public disclosure, and Grievance Redress Mechanism "
                "(GRM) for project-related complaints including environmental damage; contractor GRMs at "
                "ancillary facilities."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'GRMs to be put in place to receive, evaluate, and resolve all grievances related to the "
                "CRRN Project.' ANE/IP monitoring (+0.25). Contractor GRM at facilities (+0.25 authority). "
                "Public consultation requirements (+0.25 participation mechanism)."
            ),
            "comments": "Information/participation instrument; can receive waste/litter complaints.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§11: ESMF implementation monitoring and reporting by ANE/IP in collaboration with "
                "provincial authorities; monitoring indicators for environmental and social performance "
                "including ESMP/CESMP compliance."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'The primary responsibility for monitoring rests with ANE, IP.' Site visits, reporting, "
                "third-party input via GRM (+0.25 monitoring). World Bank oversight (+0.25 authority)."
            ),
            "comments": "Multi-level monitoring coordination for contractor waste compliance.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§5.2.2: Earthworks designed to balance cut and fill to minimise excess material "
                "consumption and waste generation."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Earthworks to be designed to achieve a balance between cut and fill wherever possible.' "
                "Waste prevention measure (+0.25 unconditional design requirement)."
            ),
            "comments": "Waste prevention at production/construction phase; reduces overall waste including packaging.",
        },
    ],
}

ENGLISH_DOCUMENT = """ENVIRONMENTAL AND SOCIAL MANAGEMENT FRAMEWORK (ESMF)
Climate Resilient Roads for the North Project (P500488)
Republic of Mozambique

Source URL: https://fe.gov.mz/storage/app/uploads/public/65f/2da/871/65f2da8718c30622471525.pdf
Prepared for: Administração Nacional de Estradas (ANE)
Prepared by: JBN Consult / EA Consultoria (Joint Venture)
Date: 06 March 2024
World Bank Project P500488

Note: This document is already written in English. The following is a structured summary of key provisions for reference alongside the 4P Index coding table. The full 209-page ESMF should be consulted for complete text.

---

EXECUTIVE SUMMARY

The Climate Resilient Roads for the North (CRRN) project aims to improve climate-resilient connectivity in Cabo Delgado, Nampula and Niassa provinces through rehabilitation and construction of roads, bridges and drainage structures. The World Bank has agreed to fund a US$125 million envelope.

This ESMF establishes the framework for assessing, managing and monitoring environmental and social impacts of CRRN subprojects in compliance with Mozambique national legislation and World Bank Environmental and Social Standards (ESS). Contractors must prepare site-specific Contractor ESMPs (CESMP) including Waste Management Plans. A Grievance Redress Mechanism (GRM) will receive and resolve project-related complaints.

Total budget for ESMF implementation and associated instruments: US$3,100,000.

---

1. PROJECT COMPONENTS

- Sub-component 1.1: Rehabilitation of priority secondary roads
- Sub-component 1.2: Improvement of bridges and drainage structures (US$38.1 million)
- Component 2: Institutional strengthening and project management
- Contingent Emergency Response Component (CERC) for emergency recovery activities

---

2. APPLICABLE MOZAMBIQUE LEGISLATION (WASTE-RELEVANT)

The ESMF confirms applicability of:
- Environmental Law (Law No. 20/97)
- Regulation on urban solid waste management (Decree No. 94/2014) — MTA
- Hazardous Waste Management (Decree No. 83/2014) — MTA
- Environmental Impact Assessment Regulation (Decree No. 54/2015) — MTA
- Regulation on Environmental Standards and Effluent Emission (Decree No. 18/2004, as amended)
- Regulation on Environmental Audit Process (Decree No. 25/2011)

---

3. WASTE MANAGEMENT REQUIREMENTS (KEY PROVISIONS)

Generic mitigation (Section 5.2.2):
- Preparation of Waste Management Plan following the waste hierarchy, supported by staff training
- Use of authorized contractors for hazardous and other wastes
- Earthworks balanced to minimise material consumption and waste

Contractor specifications (Annex / CESMP requirements):

Solid waste:
- Submit solid waste control method statement for Supervision Engineer approval
- Keep site free of litter; provide litter bins and refuse collection at all workplaces
- No burning, on-site burying or dumping
- Dispose all solid waste at approved landfill with certificates of disposal
- Random disposal in scenery areas strictly prohibited
- Fence waste storage areas to prevent wind-blown litter
- Separate and recycle packaging materials, steel, wood, scaffolding — re-use or sell to waste collectors

Hazardous waste:
- Dispose at approved hazardous landfill per Decreto 83/2014
- Used oil/grease to approved recycling company
- No spoiling of bitumen on site; return unused products to supplier
- Spill reporting and remedial action required

Wastewater:
- Conservancy tanks for kitchen/shower/laboratory wastewater
- Licensed waste collectors for sewage
- Treated discharge must comply with legislation

---

4. ENVIRONMENTAL AND SOCIAL ASSESSMENT PROCESS

Six steps: (1) Screening and scoping; (2) Risk classification; (3) ESIA/assessment; (4) Stakeholder consultation and disclosure; (5) MTA review and approval; (6) Implementation and monitoring.

ESIA annexes include: OHS plan, Community Health and Safety Plan, Biodiversity Management Plan (if required), Traffic Safety Plan, Chance Find Procedures, and Waste Management Plan.

---

5. GRIEVANCE REDRESS MECHANISM

Project-level GRM established under Stakeholder Engagement Plan (SEP) to receive, evaluate and resolve grievances. Contractors must establish GRMs at ancillary facilities. Budget for SEP: US$390,000.

---

6. MONITORING AND REPORTING

ANE and Implementing Partner (IP) responsible for ESMF monitoring in collaboration with provincial authorities, site visits, reporting, and GRM inputs. World Bank supervision throughout project cycle.

---

Plastics relevance note for 4P Index: The ESMF does not mention plastics or plastic bags explicitly. It is coded because construction solid waste management, litter prevention, packaging material recycling, and compliance with Decreto 94/2014 (urban solid waste) apply to plastic waste streams generated during road construction (filter criteria b and c)."""


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
        "A": 50, "B": 55, "C": 8, "D": 45, "E": 8, "F": 12, "G": 8, "H": 35,
        "I": 8, "J": 35, "K": 8, "L": 30, "M": 8, "N": 40, "O": 10,
        "P": 8, "Q": 18, "R": 50, "S": 8, "T": 8, "U": 40, "V": 10, "W": 40,
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
    meta.append(["Source URL", POLICY_URL])
    meta.append(["Project", "World Bank P500488 — Climate Resilient Roads for the North"])
    meta.append(["Client", "ANE (Administração Nacional de Estradas)"])
    meta.append(["Document date", "06 March 2024"])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Instrument rows", len(rows)])
    meta.append([
        "Note",
        "Project safeguard framework (not national legislation). In-force = operative for "
        "project contractors during CRRN implementation. No explicit plastics mention; coded "
        "via construction waste/litter/packaging recycling and Decreto 94/2014 compliance.",
    ])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_word_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading(
        "ESMF — Climate Resilient Roads for the North (P500488)",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "The source document is already in English. This Word file provides a structured summary "
        "of key environmental and social provisions, particularly waste management requirements "
        "relevant to the 4P Index coding."
    )

    p = doc.add_paragraph()
    p.add_run("Official title: ").bold = True
    p.add_run("Environmental and Social Management Framework (ESMF) — Climate Resilient Roads for the North")

    p2 = doc.add_paragraph()
    p2.add_run("Source URL: ").bold = True
    p2.add_run(POLICY_URL)

    doc.add_heading("Document summary", level=2)

    for block in ENGLISH_DOCUMENT.strip().split("\n\n"):
        text = block.strip()
        if not text or text.startswith("---"):
            continue
        if text.isupper() and len(text) < 80:
            doc.add_heading(text.title(), level=3)
        elif text[0].isdigit() and "." in text[:4]:
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
    write_word_doc()


if __name__ == "__main__":
    main()
