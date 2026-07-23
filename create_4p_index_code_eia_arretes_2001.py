#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for EIA ministerial orders 9468-9472 (English only).

EIA provisions for plastic/waste facilities only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Arretes_EIA_9468-9472_2001.xlsx"

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
        "Ministerial orders of 28 November 2001 on environmental impact assessment "
        "(Nos. 9468-9472 MJEHP-DEEC)"
    ),
    "policy_url": "https://www.denv.gouv.sn/decrets/",
    "policy_year": 2001,
    "policy_objective": (
        "Operationalise the environmental assessment procedure under the Environmental Code and Decree "
        "2001-282 by regulating public participation, technical committee organisation, consultant "
        "accreditation, terms of reference and EIA report content — ensuring environmental review of "
        "industrial projects including plastics plants and waste-treatment facilities, without "
        "plastic-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Package of five ministerial orders (Nos. 9468-9472) adopted on 28 November 2001 by the Minister "
        "of the Environment (MJEHP-DEEC), published in Official Journal No. 6025 of 12 January 2002. "
        "Executive implementing regulations, not acts enacted by Parliament."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "environment, industry, planning, waste management, urban planning, health, water, municipalities"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "production, disposal, environmental leakage",
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Order No. 9471, Art. 1: terms of reference for any environmental impact study must include "
            "14 mandatory elements, including item 4: an assessment of measures envisaged for wastewater "
            "evacuation, solid waste elimination and emission reduction. Art. 2: DEEC may develop "
            "sector-specific ToR."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 1: mandatory ToR elements (+0.25 authority); "
            "Art. 3: Director of Environment and Classified Establishments responsible for implementation "
            "(+0.25 monitoring); "
            "Enforcement via EIA irreceivability (Order 9472 Art. 3) and committee review (9469) "
            "(+0.25 enforcement)."
        ),
        "comments": (
            "Direct entry point for assessing solid waste and emissions in EIAs for plastic projects. "
            "No quantitative plastic thresholds."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Order No. 9472, Art. 1: environmental impact study report shall comprise 16 mandatory sections "
            "including project description, baseline analysis, impact evaluation, accident risks, "
            "preventive/mitigation measures, and an environmental monitoring and follow-up plan framework (EMP). "
            "Art. 2: report must be in French, 10 copies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 1 §13: promoter must submit detailed EMP at pre-construction with cost, schedule and "
            "responsible structures (+0.25 authority); "
            "Art. 3: non-compliant reports declared inadmissible (+0.25 enforcement); "
            "Review by Technical Committee (Order 9469) (+0.25 monitoring)."
        ),
        "comments": (
            "Mandatory EIA report framework for industrial projects; includes post-project monitoring "
            "relevant to operational plastic impacts."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Order No. 9470, Arts. 1-7: accreditation to carry out environmental impact study activities may "
            "be granted to qualified physical or moral persons meeting diploma, experience and logistics "
            "requirements. Art. 6: accreditation may be withdrawn for serious professional breaches or loss "
            "of required qualifications."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Arts. 1-2: accreditation commission chaired by Environment Minister (+0.25 authority); "
            "Art. 4: minimum requirements — higher degree, EIA experience, logistics, or legal entity with "
            "five senior experts (+0.25 monitoring); "
            "Art. 6: withdrawal for serious breach of professional obligations (+0.25 enforcement). "
            "Accreditation valid 5 years."
        ),
        "comments": (
            "Qualifies consultancy firms conducting EIAs for plastics plants and waste-treatment facilities."
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
        "A": 52, "B": 48, "C": 10, "D": 52, "E": 12, "F": 20,
        "G": 10, "H": 52, "I": 14, "J": 44, "K": 14, "L": 36,
        "M": 12, "N": 36, "O": 12, "P": 14, "Q": 22, "R": 60,
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
