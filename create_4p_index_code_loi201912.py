#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2019-12 CGCT amendment (English only)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2019-12_CGCT.xlsx"

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
        "Law No. 2019-12 of 8 July 2019 amending the General Code of Local Authorities"
    ),
    "policy_url": "https://www.dri.gouv.sn/sites/default/files/LOI/LOI%202019/L-2019-12.pdf",
    "policy_year": 2019,
    "policy_objective": (
        "Amend CGCT fiscal framework (CEL replacing business patent contribution) while retaining "
        "local household-waste collection and street-sweeping taxes funding municipal waste services "
        "including plastic fractions. No plastic-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Amending law adopted by National Assembly 29 June 2019, promulgated 8 July 2019. "
        "Parliamentary legislation (≠0.75)."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": "municipalities, waste management, environment",
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": "disposal",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. 195 bis: CEL value-added redistribution to communes; Arts 185/195 retain "
        "household waste collection and street-sweeping taxes as local revenue sources."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 195(1)(b): commune direct local taxes include household waste collection tax "
            "and street-sweeping tax, created by municipal council deliberation within legal maxima."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Municipal council creates taxes by deliberation (+0.25 authority); "
            "assessment/collection modalities set by law (+0.25 monitoring). No penalties in this article."
        ),
        "comments": "Most plastics-relevant — local taxes funding household waste collection.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 185(1)(b): city operating revenues include proceeds of the household waste collection tax, "
            "created by city council deliberation within legal maxima."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "City council deliberation required (+0.25 authority); State grants 25% quarterly advance "
            "on prior-year direct-tax collections (+0.25 monitoring)."
        ),
        "comments": "City-level waste collection tax — funds collection including plastic fractions.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 195 bis(1): minimum CEL allocation guichet — not below FCFA 12,000,000 per commune "
            "per year for authorised personnel expenditure, supporting local service delivery capacity."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "National redistribution mechanism (+0.25 authority). Indirect support to commune waste "
            "service capacity; no direct waste mandate or enforcement in this provision."
        ),
        "comments": "Indirect fiscal support to communes managing waste; borderline plastics relevance.",
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
