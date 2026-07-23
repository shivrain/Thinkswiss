#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2022-18 SONAGED (English only)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2022-18_SONAGED.xlsx"

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
    "policy_name": (
        "Law No. 2022-18 of 23 May 2022 authorising creation of the National Integrated "
        "Waste Management Company (SONAGED S.A.)"
    ),
    "policy_url": "https://www.fao.org/faolex/results/details/en/c/LEX-FAOC216236",
    "policy_year": 2022,
    "policy_objective": (
        "Authorise creation of SONAGED S.A., a State-owned company replacing UCG, to coordinate integrated "
        "solid-waste collection, treatment and valorisation nationally — integrating PROMOGED and other public "
        "programmes. Applies to municipal solid waste including plastics; no plastic-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Law adopted by the National Assembly on 5 May 2022 and promulgated 23 May 2022 as Law No. 2022-18. "
        "Parliamentary legislation authorising creation of a State-owned company (≠0.75)."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": "waste management, municipalities, industry, environment, recycling",
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "recycling, disposal",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. 4: resources from State endowment, other authorised resources and operating revenues "
        "to fund nationwide collection and operating costs."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 1: authorises creation of Société nationale de Gestion intégrée des Déchets (SONAGED S.A.), "
            "a State-owned company under Public Hygiene Minister technical supervision."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 1: SONAGED and supervising ministers explicitly designated (+0.25 authority). "
            "Organisation and statutes require implementing decree (Art. 6); no enforcement in authorising law."
        ),
        "comments": "Institutional creation for national solid-waste management including plastic fractions.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 2: SONAGED charged to ensure collection, transport, landfilling, treatment and valorisation "
            "of solid waste nationwide; manage treatment/valorisation infrastructure; act as sector regulator; "
            "promote inter-municipal integrated waste management and circular economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 2: SONAGED designated with nationwide mandate (+0.25 authority). "
            "No fines, penalties or monitoring requirements in the authorising law itself."
        ),
        "comments": "Core operational mandate — full MSW chain including plastics. Single consolidated Art. 2 row.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5: UCG and all public integrated solid-waste management projects and programmes "
            "(including PROMOGED) transferred to SONAGED."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 5: mandatory transfer of UCG and PROMOGED to SONAGED (+0.25 authority). "
            "No enforcement or monitoring provisions in authorising law."
        ),
        "comments": "Integrates PROMOGED infrastructure and programmes for waste including plastics.",
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
