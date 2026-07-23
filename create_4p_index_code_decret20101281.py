#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Decree 2010-1281 (English only).

Lead/mercury battery provisions only; plastic casings borderline — strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2010-1281_Batteries_Usees.xlsx"

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
        "Decree No. 2010-1281 of 16 September 2010 regulating recovery of lead from used batteries"
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/19090/"
        "decret-num-3-2010-1281-du-16-septembre-2010-reglementant-les-conditions-dexploitation-du-plomb-issu-des-batteries-usagees.pdf"
    ),
    "policy_year": 2010,
    "policy_objective": (
        "Regulate import, collection, transport, recycling, storage, treatment and disposal of lead from "
        "used batteries and mercury use, by imposing Environment Minister authorisation, ICPE/EIA requirements "
        "and NS 05-062 compliance — addressing hazardous components (lead, mercury) often combined with plastic "
        "casings in waste electrical equipment, without plastic-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree adopted by the President of the Republic on 16 September 2010, implementing the "
        "Environmental Code and Decree 2001-282. Sub-legislative implementing regulation, not parliamentary "
        "legislation (≠1.0)."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": "industry, waste management, environment, mining, health",
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "consumption, recycling, disposal",
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 1: it is prohibited for any natural or legal person to import, collect, transport, recycle, "
            "store, handle, treat or eliminate lead from used batteries and other sources, as well as mercury "
            "and its compounds, without authorisation from the Minister of the Environment."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 1: authorisation delivery conditions set by ministerial order (+0.25 authority); "
            "Art. 5: violations punished per Environmental Code sanctions (+0.25 enforcement); "
            "Art. 3: ICPE permit and EIA required (+0.25 monitoring)."
        ),
        "comments": (
            "Core prohibition without authorisation; used batteries include plastic casings "
            "(component not explicitly regulated)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 2: companies specialised in recovery and recycling of lead from used accumulators are "
            "eligible for authorisation. Art. 3: operators must demonstrate process mastery from arrival "
            "through treatment to finished product output."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 3: compliance with hazardous chemical and waste management legislation (+0.25 authority); "
            "Art. 1: only duly authorised persons may carry out activities (+0.25 enforcement); "
            "Presentation report targets uncontrolled informal battery recycling (+0.25 monitoring)."
        ),
        "comments": (
            "Lead recycling instrument; plastic casings/co-products not explicitly addressed in decree."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 3: operators must manage waste in accordance with Article L 30 of the Environmental Code "
            "(ecologically rational hazardous waste management)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Presentation report regulates storage, treatment and disposal of lead from used batteries "
            "(+0.25 authority); Code L 30: producers must eliminate/recycle or use licensed enterprises "
            "(+0.25 enforcement); Art. 5: Environmental Code sanctions (+0.25 monitoring)."
        ),
        "comments": (
            "Residual waste management (incl. plastic battery fractions) via Environmental Code L 30."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 3: operators must hold an authorisation to operate a classified installation, having "
            "previously undergone an environmental assessment, and comply with Chapter 2 of NS 05-062 on "
            "atmospheric pollution."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "References Decree 2001-282 ICPE framework and EIA orders 9468-9472 (+0.25 authority); "
            "NS 05-062 Ch. II sets emission limits for stationary installations (+0.25 monitoring); "
            "Art. 5: Environmental Code sanctions (+0.25 enforcement)."
        ),
        "comments": (
            "ICPE + EIA mandatory for battery recycling plants; relevant for plastic/battery facilities."
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
