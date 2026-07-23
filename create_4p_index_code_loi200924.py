#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2009-24 Sanitation Code (English only).

Explicit plastic/sewer waste provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2009-24_Code_Assainissement.xlsx"

COLUMNS = [
    ("A", "policy_name"),
    ("B", "policy_url"),
    ("C", "policy_year"),
    ("D", "policy_objective"),
    ("E", "policy_target"),
    ("F", "policy_target_text"),
    ("G", "policy_type"),
    ("H", "policy_type_justification"),
    ("I", "policy_integration"),
    ("J", "policy_sectors_list"),
    ("K", "policy_circularity"),
    ("L", "policy_lifecycle_phases_list"),
    ("M", "policy_budget"),
    ("N", "policy_budget_text"),
    ("O", "policy_score"),
    ("P", "instrument_type"),
    ("Q", "instrument_lifecycle_stage"),
    ("R", "instrument_description"),
    ("S", "instrument_in_force"),
    ("T", "instrument_implementation"),
    ("U", "instrument_implementation_text"),
    ("V", "instrument_score"),
    ("W", "comments"),
]

POLICY = {
    "policy_name": "Law No. 2009-24 of 8 July 2009 establishing the Sanitation Code",
    "policy_url": "https://www.sante.gouv.sn/sites/default/files/3.%20Code_assainissement.pdf",
    "policy_year": 2009,
    "policy_objective": (
        "Unify the legal framework for liquid sanitation (wastewater, excreta, stormwater) by setting planning "
        "obligations for communes and rural communities, prohibitions on discharge into public sewers "
        "(including plastic waste), regimes for domestic, industrial and hospital effluents, and sanctions — "
        "indirect governance of plastic waste through explicit prohibition of discharge into public collectors."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Sanitation Code adopted by the National Assembly on 17 June 2009 and by the Senate on "
        "29 June 2009, promulgated by the President of the Republic on 8 July 2009 as Law No. 2009-24. "
        "Parliamentary legislation (Code), not executive implementing regulation (≠0.75)."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "sanitation, urban planning, local government, environment, health, industry"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "disposal, environmental leakage",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. L 60: industrial effluent fee calculated on discharged volume and effluent quality; "
        "Art. L 73: taxes and fees for discharge into natural environment; "
        "Art. L 88: septage fee per tonne or cubic metre at managed deposit sites."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L 29: it is strictly prohibited to discharge into public wastewater collectors household "
            "waste and plastic waste, along with hydrocarbons, used oils, radioactive substances, paint "
            "residues, septic tank contents, sludge, sand, rubble and other solids or harmful substances "
            "that could compromise sewer or treatment-plant operation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. L 30: prohibited-discharge list is not exhaustive (+0.25 monitoring); "
            "sworn sanitation agents may take control samples at user premises (+0.25 authority); "
            "Arts. L 104–L 108: penal sanctions — 2 months to 2 years imprisonment and "
            "FCFA 20,000–2,000,000 fines (+0.25 enforcement)."
        ),
        "comments": (
            "Explicit mention of plastic waste — direct prohibition on discharge into public sewers."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Arts. L 79–L 80: discharge of septic tank sludge is prohibited except in sealed tanks, "
            "treatment plants or authorised deposit sites; deposit sites may not accept household waste "
            "(even after prior shredding), household refuse (even after prior shredding), industrial waste "
            "or healthcare activity waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. L 79: septage transport only by trucks approved by the Sanitation Minister (+0.25 authority); "
            "Art. L 82: emptying companies approved by sanitation service (+0.25 monitoring); "
            "Arts. L 104–L 108: penal sanctions for violations (+0.25 enforcement)."
        ),
        "comments": (
            "Prohibits household/plastic waste (even shredded) in sludge deposit sites; complements Art. L 29."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L 13: discharge of untreated domestic effluent, excreta and septage into gutters, open "
            "stormwater channels, soil surfaces, rivers, lakes, ponds and the sea is prohibited throughout "
            "the national territory."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. L 3: all liquid waste discharges into the natural environment require prior depollution (+0.25 authority); "
            "Art. L 5: pollution sources subject to prior authorisation and administrative inquiry (+0.25 monitoring); "
            "Arts. L 98–L 103: sworn agents from sanitation, health and environment ministries may inspect (+0.25 enforcement)."
        ),
        "comments": (
            "Limits leakage of solid/plastic waste carried by untreated wastewater into surface waters."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. L 39–L 46: every urbanised public or private place must have a stormwater collection and "
            "evacuation system preventing stagnation; public stormwater works must receive only rainwater "
            "(except unitary networks); private land within 30 m of a public storm network must connect."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. L 50: natural runoff servitude; prohibited to discharge polluted stormwater to lower plots (+0.25 authority); "
            "Art. L 28: storm drains/gutters prohibited on separate domestic wastewater systems (+0.25 monitoring). "
            "No article-specific plastic-waste penalty."
        ),
        "comments": (
            "Stormwater infrastructure rules reduce transport of solid/plastic litter to gutters and receiving waters."
        ),
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

    widths = {
        "A": 48, "B": 52, "C": 10, "D": 52, "E": 12, "F": 60,
        "G": 10, "H": 52, "I": 14, "J": 40, "K": 14, "L": 36,
        "M": 12, "N": 52, "O": 12, "P": 14, "Q": 22, "R": 60,
        "S": 14, "T": 18, "U": 52, "V": 14, "W": 52,
    }
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
