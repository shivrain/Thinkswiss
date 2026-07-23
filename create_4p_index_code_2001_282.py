#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret n° 2001-282 (English only).

Plastic-relevant ICPE, EIA and pollution provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2001-282_Code_Environnement.xlsx"

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
    "policy_name": (
        "Decree No. 2001-282 of 12 April 2001 implementing the Environmental Code"
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/sen37192.pdf",
    "policy_year": 2001,
    "policy_objective": (
        "Implement the Environmental Code by regulating classified installations, environmental impact "
        "assessment, water pollution, water police, air pollution and noise pollution — providing an "
        "indirect regulatory framework for plastic pollution through ICPE rules, discharges, "
        "incineration and EIA for waste-treatment projects, without plastic-specific measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree adopted by the President of the Republic on 12 April 2001, after opinion of "
        "the Council of State (13 October 2000), on the report of the Minister of the Environment, "
        "implementing Law No. 2001-01. Sub-legislative implementing regulation (regulatory part "
        "of the Code), not an act enacted by Parliament."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "environment, industry, transport, municipalities, fisheries, water, health, urban planning, "
        "agriculture, waste management, packaging"
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. R26: classified-installation rights and taxes must be paid within 45 days after assessment notice. "
        "Art. R32: surface tax payable by every classified installation. "
        "Art. R54: effluent pollution tax set according to degree of effluent pollution."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Arts. R3–R4, R16–R17: classified installations — including plastic processing plants listed "
            "in the ICPE nomenclature — must obtain operating authorisation or file a declaration; "
            "Class II declarations must specify modes and conditions for wastewater treatment, emissions "
            "and elimination of waste and operational residues."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Art. R22: administrative and technical surveillance and control of installations (+0.25 authority); "
            "Arts. R24–R25: criminal penalties and official reports by sworn agents (+0.25 enforcement); "
            "Arts. R20–R21: authorised sworn inspectors under Environment Minister (+0.25 monitoring); "
            "Art. R17: declaration receipt issued within two months (+0.25 monitoring)."
        ),
        "comments": (
            "Cross-cutting ICPE regime for plastic production/processing and waste-elimination disclosure."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. R38–R44: environmental impact studies required prior to administrative authorisation "
            "for planned activities; EIA must assess effects of recycling and elimination of residues and "
            "waste; waste treatment and storage is an accredited consultant category; projects classified "
            "Category 1 or 2 in annexes include recycling plants and landfills."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Arts. R43–R44: technical committee reviews EIA reports monthly; Minister decides within 15 days (+0.25 authority); "
            "Arts. R6–R7: 15-day public inquiry (+0.25 monitoring); "
            "Operationalised via Ministerial Orders 9468–9472 (28 November 2001) (+0.25 enforcement)."
        ),
        "comments": (
            "EIA framework applicable to plastic recycling plants, landfills and recovery units listed in annexes."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Arts. R49, R51 and R56: discharged effluent must not deteriorate the receiving environment; "
            "effluent discharge authorisation conditional on EIA results and compliance with physical, "
            "chemical, biological and bacteriological standards; all discharges, flows, deposits or acts "
            "likely to pollute continental or marine waters — including industrial waste from the coast — "
            "are prohibited under water police."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. R50: impact study required for discharges into natural receiving environments (+0.25 authority); "
            "Arts. R59–R66: judicial police officers and sworn agents may record violations (+0.25 monitoring); "
            "Art. R69: vessel detention for pollution offences (+0.25 enforcement)."
        ),
        "comments": (
            "Direct prohibition applicable to plastic waste discharges into waters; complements Art. L41 of the Code."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Arts. R75–R78: fixed incineration, combustion or heating installations must meet technical "
            "specifications for equipment placed on the Senegalese market; operating conditions may be "
            "set by ministerial order; installations subject to periodic inspection by an accredited expert "
            "or body."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. R76–R77: ministerial orders may set technical specifications and operating conditions (+0.25 authority); "
            "Art. R78: periodic inspection framework explicit (+0.25 monitoring). "
            "Detailed mandatory standards deferred to subordinate orders — in_force=0 for standards themselves."
        ),
        "comments": (
            "Incineration framework relevant to plastic waste end-of-life; detailed technical norms via arrêtés."
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
        "A": 48, "B": 52, "C": 10, "D": 52, "E": 12, "F": 20,
        "G": 10, "H": 48, "I": 14, "J": 40, "K": 14, "L": 36,
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
