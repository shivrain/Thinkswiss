#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for LPSEDD 2016-2020 (English only).

Plastic-relevant strategic provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Lettre_Politique_Environnement_2016-2020.xlsx"

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
        "Policy letter for the environment and sustainable development sector 2016-2020 (LPSEDD)"
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/61/documentation/18323/"
        "lettre-de-politique-du-secteur-de-lenvironnement-et-du-developpement-durable-2016-2020.pdf"
    ),
    "policy_year": 2016,
    "policy_objective": (
        "Define strategic orientation for the environment and sustainable development sector for 2016-2020, "
        "covering living-environment management, pollution control and SDG alignment — with explicit references "
        "to plastic-waste proliferation, the 0-30 micron bag law and opportunities for plastic-waste "
        "valorisation/recycling."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "Sector policy letter validated December 2015 for the 2016-2020 period. Strategic orientation and "
        "ministerial planning document with aspirational commitments but no quantifiable plastic targets "
        "(≠0.50 plan with targets, ≠0.75 order/decree, ≠1.0 law)."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "environment, sustainable development, living environment, waste, pollution, climate, "
        "local government, industry"
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "production, consumption, recycling, disposal",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "§II.2.3.1(c): MEDD budget fell from FCFA 31.25 billion (2011) to approximately FCFA 22 billion "
        "(2015); implementation via DPPD budget-programme framework; no dedicated plastic-waste budget line."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§II.2.2.1: plastic-waste proliferation is symptomatic of persistent bad consumption practices, "
            "low preparedness of responsible structures and absence of a real strategy; its eradication "
            "remains a priority."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Policy expects Law banning 0-30 micron plastic bags to help (+0.25 authority via ministerial planning). "
            "Aspirational priority only — no binding obligation. Period expired 2020."
        ),
        "comments": "Most explicit plastic mention — aspirational priority to eradicate plastic waste.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§II.2.2.1: entry into force of the law prohibiting manufacture, distribution and use of plastic "
            "bags between 0 and 30 microns should help address plastic-waste proliferation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "§II.2.3.1(b): plastic-bag law passed as part of legal reform (+0.25 authority). "
            "Informational reference only — not an operative instrument in LPSEDD. Period expired 2020."
        ),
        "comments": "Reference to Law 2015-09 (thin bags) — informational, not a regulatory instrument.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§II.2.3.1(c): opportunities in valorisation and recycling of plastic waste, alongside ecotourism, "
            "forestry and renewables, to complement corporate social responsibility."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0,
        "instrument_implementation_text": (
            "Identified as economic opportunity only; no quantified recycling targets, EPR framework or "
            "binding obligation in LPSEDD. Period expired 2020."
        ),
        "comments": "Only explicit mention of plastic-waste recycling — voluntary/aspirational.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§II.2.2.1: living-environment management requires salubrity, adequate pollution and nuisance "
            "management, disaster-risk management and support for waste collection and treatment; diagnostic "
            "notes weak sorting, collection, transport and valorisation performance."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "§III Programme 3 OS2: improve living-environment quality through rational pollution management "
            "(+0.25 authority). Planning orientation only — period expired 2020."
        ),
        "comments": "Living-environment governance orientation — waste collection/treatment including plastics.",
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
