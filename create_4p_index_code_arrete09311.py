#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Order 09311/2007 (English only).

Used-oil provisions with indirect plastic relevance only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Arrete_09311_Huiles_Usees.xlsx"

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
        "Interministerial Order No. 09311 of 5 October 2007 on used-oil management"
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/73/arretes-cadre-juridique/"
        "19096/arrete-interministeriel-n-09311-du-5-octobre-2007-portant-gestion-des-huiles-usees.pdf"
    ),
    "policy_year": 2007,
    "policy_objective": (
        "Regulate collection, storage, transport, treatment and disposal of used oils by imposing state "
        "approval, the green register and prohibition of environmental discharge — reduce hazardous waste "
        "streams often co-occurring with plastic waste in industrial and municipal settings, without "
        "plastic-specific measures."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Five hundred litres annual production threshold for green register; "
        "fifteen days maximum pickup delay for lots exceeding 600 litres; "
        "1,400°C minimum for incineration/co-incineration; "
        "five years renewable approval validity."
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Interministerial order adopted on 5 October 2007 jointly by the Ministers of Environment, "
        "Mines and Industry, and Energy, published in the Official Journal. Executive implementing "
        "regulation operationalising the Environmental Code and Decree 2001-282 (≠1.0 law)."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": "industry, transport, environment, energy, waste management",
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "consumption, disposal, environmental leakage",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Annex 1: deposit required for collection and elimination approval applications; "
        "§19: annual statistics reporting to DEEC."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 3: prohibitions including depositing or allowing used oils to flow on or into soil, "
            "surface or groundwater, sewers, pipes or collectors; burning used oils except under Article 2 "
            "conditions; disposing of used oils except by handing them to approved companies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 15: approved companies controlled by competent authority (+0.25 monitoring); "
            "Art. 20: violations expose offender to sanctions under applicable regulations (+0.25 enforcement); "
            "Art. 18: approval withdrawal for non-compliance (+0.25 authority)."
        ),
        "comments": (
            "Core environmental prohibitions; reduces hydrocarbon discharges co-located with plastic/municipal waste."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 3 §3-5: prohibited mixing of used oils with PCBs or hazardous waste, or adding water or "
            "foreign matter such as solvents, cleaning products, detergents, antifreeze or other fuels before "
            "or during collection/storage."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Annex 1 (eliminator): material accounting including PCB content (+0.25 monitoring); "
            "Annex 1 (collector): separation between stored oils and all other waste (+0.25 authority); "
            "Art. 20: sanctions under applicable regulations (+0.25 enforcement)."
        ),
        "comments": (
            "Prevents co-mingling of oils/hazardous waste; relevant for mixed oil-plastic streams at industrial sites."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 1 (PCB): polychlorinated biphenyls, used in industry as additive in paints, carbonless "
            "papers and in plastics. Art. 3 §4: prohibited to mix used oils with PCBs."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 3 §4: mixing prohibition (+0.25 enforcement); "
            "Annex 1: eliminators must record PCB content in material accounting (+0.25 monitoring); "
            "Art. 5: only state-approved companies may collect or eliminate used oils (+0.25 authority)."
        ),
        "comments": (
            "Only explicit mention of plastics (PCB as additive); indirect relevance for chlorinated plastic waste."
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
