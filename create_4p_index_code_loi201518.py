#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Law 2015-18 Maritime Fisheries Code (English only).

Marine litter and plastic fishing-gear provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2015-18_Code_Peche_Maritime.xlsx"

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
        "Law No. 2015-18 of 13 July 2015 establishing the Maritime Fisheries Code"
    ),
    "policy_url": "https://www.ditp.gouv.sn/download/file/fid/53",
    "policy_year": 2015,
    "policy_objective": (
        "Refound the maritime fisheries management framework, including marine environment protection "
        "and fishing-gear control — notably nylon monofilament/multimonofilament nets — relevant to "
        "plastic fishing gear and ghost-gear reduction, without plastic marine litter-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Maritime Fisheries Code adopted by the National Assembly on 30 June 2015 and promulgated on "
        "13 July 2015 as Law No. 2015-18. Parliamentary legislation (Code), not executive implementing "
        "regulation (≠0.75)."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "maritime fisheries, artisanal fishing, industrial fishing, marine environment, maritime trade"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "production, consumption, disposal, environmental leakage",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. 64: fee for permits and authorisations (amount set by joint ministerial order); "
        "Arts. 123-127: criminal fines up to FCFA 30,000,000 for serious industrial fishing offences."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 66: import, sale, purchase, possession and use of gillnets made from nylon monofilament "
            "or multimonofilament elements are prohibited except by special derogation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Exposé des motifs: prohibition raised to legislative level to combat ghost gear (+0.25 enforcement); "
            "Arts. 125-126: use of prohibited gear is a very serious offence with confiscation and destruction "
            "(+0.25 enforcement); Arts. 84-85: surveillance agents may inspect vessels and gear at sea "
            "(+0.25 authority)."
        ),
        "comments": "Direct ban on nylon (plastic polymer) nets — main explicit link to marine plastics/ghost gear.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 68: all artisanal fishing vessels belonging to nationals are registered and marked under "
            "ministerial rules; same for foreign residents' vessels. Art. 33(k): regulatory measures on "
            "signalling of artisanal fishing gear."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 83: Minister responsible for fisheries surveillance and protection (+0.25 authority); "
            "Arts. 84-85: agents may inspect fishing gear on board or have gear removed from water "
            "(+0.25 monitoring); Decree 2016-1804 Arts. 56-58: vessel marking specifications (+0.25 enforcement)."
        ),
        "comments": (
            "Mandatory gear registration and marking — regulatory traceability for identifying and recovering "
            "abandoned gear."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 125(a),(e): very serious offences include use of prohibited fishing gear or methods and "
            "use of devices reducing mesh opening below the minimum or reducing selective action; "
            "gear is confiscated and destroyed."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Penalty: FCFA 20,000,000-30,000,000 for industrial fishing (+0.25 enforcement); "
            "Art. 126: artisanal fishing FCFA 150,000-300,000 with gear confiscation and destruction "
            "(+0.25 enforcement); Art. 85(a): agents may have fishing gear removed from water (+0.25 authority)."
        ),
        "comments": "Confiscation and destruction of illegal gear — elimination mechanism for prohibited plastic fishing gear.",
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
