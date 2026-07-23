#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for PROMOGED (English only)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_PROMOGED_Dechets_Solides.xlsx"

COLUMNS = [
    ("A", "policy_name"), ("B", "policy_url"), ("C", "policy_year"),
    ("D", "policy_objective"), ("E", "policy_target"), ("F", "policy_target_text"),
    ("G", "policy_type"), ("H", "policy_type_justification"),
    ("I", "policy_integration"), ("J", "policy_sectors_list"),
    ("K", "policy_circularity"), ("L", "policy_lifecycle_phases_list"),
    ("M", "policy_budget"), ("N", "policy_budget_text"),
    ("O", "policy_score"), ("P", "instrument_type"),
    ("Q", "instrument_lifecycle_stage"), ("R", "instrument_description"),
    ("S", "instrument_in_force"), ("T", "instrument_implementation"),
    ("U", "instrument_implementation_text"), ("V", "instrument_score"),
    ("W", "comments"),
]

POLICY = {
    "policy_name": "Solid Waste Integrated Management and Economy Promotion Project (PROMOGED)",
    "policy_url": (
        "https://www.urbanisme.gouv.sn/actualites/r%C3%A9union-du-comit%C3%A9-de-pilotage-du-promoged"
    ),
    "policy_year": 2021,
    "policy_objective": (
        "World Bank–supported national project to modernise municipal solid-waste collection, sorting, transfer "
        "and treatment infrastructure and develop valorisation value chains — including plastics (~9% of Mbeubeuss "
        "waste stream) — across selected Senegalese municipalities. Legal framework: Ministerial Order 027932/2020 "
        "(Official Gazette No. 7392 of 16 January 2021)."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "World Bank P161477 PDO: strengthen solid-waste governance and improve services in selected municipalities; "
        "Phase I targets: 350 infrastructures, 15 landfills rehabilitated, 7 regions / 148 communes; "
        "WB indicators: 10 collection/transfer facilities; sanitary landfills operational."
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "National development programme/project financed by donors (World Bank P161477, AFD, EIB). "
        "Ministerial Order 027932/2020 creates steering committee (≠1.0 law). Programme with quantifiable targets."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "waste management, municipalities, industry, environment, recycling, agriculture"
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "recycling, disposal",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Total ~FCFA 206 billion Phase I; World Bank credit US$175M (P161477): "
        "Component 1 US$20M governance; Component 2 US$256.1M infrastructure; Component 3 US$18.6M implementation."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Ministerial Order 027932/2020 Art. 1: creates PROMOGED steering committee (COPIL) within the "
            "Ministry of Urban Planning, Housing and Public Hygiene; published Official Gazette No. 7392 "
            "of 16 January 2021."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 3: chaired by Urban Planning Minister; inter-ministerial membership (+0.25 authority); "
            "Art. 2: supervises implementation and reviews disbursement-linked indicators (+0.25 monitoring)."
        ),
        "comments": "Governance instrument for municipal solid-waste management including plastic fractions.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "World Bank Component 2: construction of collection points, sorting/transfer centres, collection "
            "and marketing centres, and sanitary landfills in Greater Dakar and secondary urban poles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "WB targets: 10 collection/transfer facilities; ~350 infrastructures across 148 communes (+0.25 monitoring); "
            "UCG/SONAGED designated implementing agency (+0.25 authority). No enforcement provisions in project documents."
        ),
        "comments": "Infrastructure for collection/sorting of municipal solid waste including plastics.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "PROMOGED Component 2: Collection and Marketing Centres (CCV) and Integrated Waste Recovery Centres "
            "for sorted materials including plastics; Mbeubeuss rehabilitation with sorting before landfilling."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "ESIA notes ~9% plastic in Mbeubeuss stream; 4 CCV per pole planned (+0.25 monitoring via project indicators); "
            "UCG coordination (+0.25 authority). No enforcement mechanism in project text."
        ),
        "comments": "Most direct plastic-relevant instrument — valorisation of sorted plastic waste.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Component 1: regulatory/institutional reform and public-private partnership framework for municipal "
            "solid-waste management; inter-municipal coordination for waste services."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Results-based financing disbursed on reform indicators (+0.25 monitoring); "
            "COPIL/Technical Committee designated (+0.25 authority). No fines or penalties in project framework."
        ),
        "comments": "Multi-level waste governance coordination applying to plastic waste streams in practice.",
    },
]


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index Coding"
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    auto_fill = PatternFill("solid", fgColor="E2EFDA")

    for col_idx, (col_letter, col_name) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=f"{col_letter} — {col_name}")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row_idx, instrument in enumerate(INSTRUMENTS, start=2):
        row_data = {**POLICY, **instrument}
        for col_idx, (_, key) in enumerate(COLUMNS, start=1):
            if key in ("policy_score", "instrument_score"):
                continue
            ws.cell(row=row_idx, column=col_idx, value=row_data.get(key, ""))
        g_col, i_col, k_col, m_col = (get_column_letter(n) for n in (7, 9, 11, 13))
        p_col, t_col, o_col, v_col = (get_column_letter(n) for n in (16, 20, 15, 22))
        ws[f"{o_col}{row_idx}"] = f"=AVERAGE({g_col}{row_idx},{i_col}{row_idx},{k_col}{row_idx},{m_col}{row_idx},{p_col}{row_idx})"
        ws[f"{v_col}{row_idx}"] = f"=AVERAGE({p_col}{row_idx},{t_col}{row_idx})"
        ws[f"{o_col}{row_idx}"].fill = auto_fill
        ws[f"{v_col}{row_idx}"].fill = auto_fill

    widths = {"A": 48, "B": 52, "C": 10, "D": 52, "E": 12, "F": 60, "G": 10, "H": 52,
              "I": 14, "J": 40, "K": 14, "L": 36, "M": 12, "N": 52, "O": 12, "P": 14,
              "Q": 22, "R": 60, "S": 14, "T": 18, "U": 52, "V": 14, "W": 52}
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(OUTPUT)
    print(f"Saved {OUTPUT} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    build_workbook()
