#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2019-1855 MEDD attributions (English only).

Source: FAOLEX sen204304.pdf. Superseded by later government-reorganisation decrees.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2019-1855_MEDD_Attributions.xlsx"

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
        "Decree No. 2019-1855 of 7 November 2019 on the powers of the Minister of "
        "Environment and Sustainable Development"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/sen204304.pdf",
    "policy_year": 2019,
    "policy_objective": (
        "Define ministerial powers for environmental monitoring, pollution control and support to territorial "
        "collectivities for waste collection and treatment. Institutional mandate only — does not regulate plastics directly."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Presidential decree of 7 November 2019 setting ministerial attributions (replaces Decree 2019-975). "
        "Executive organisational regulation (≠1.0 law). Superseded by later government-reorganisation decrees."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": "waste management, municipalities, industry, environment, water",
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "disposal, environmental leakage",
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 1: Minister responsible for environmental protection; takes measures to prevent and combat "
            "pollution of all kinds; ensures safety of potentially polluting installations."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 1: MEDD designated as responsible authority (+0.25). Subject to local-government competences. "
            "No enforcement in attributions decree. Superseded by later cabinet-reform decrees."
        ),
        "comments": "Indirect plastic-pollution framework. Superseded — instrument_in_force=0.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 1: Minister helps territorial collectivities cope with waste collection and ensures treatment."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 1: MEDD designated to support municipal waste collection and treatment (+0.25 authority). "
            "Multi-level waste governance; no enforcement provisions. Superseded."
        ),
        "comments": "Most plastics-relevant provision — municipal waste coordination. Passes relevance filter (c).",
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
