#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for NS 05-062 (English only).

Plastic production/disposal emission provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_NS_05-062_Pollution_Atmosphérique.xlsx"

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
    "policy_name": "Senegalese Standard NS 05-062 — Atmospheric pollution (October 2003)",
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/74/normes-cadre-juridique/"
        "19106/norme-rejets-ns05-062-pollution-atmospherique-octobre-2003.pdf"
    ),
    "policy_year": 2003,
    "policy_objective": (
        "Set atmospheric emission and immission limit values, open-burning prohibitions, waste incineration "
        "requirements and monitoring modalities for gaseous effluents from stationary installations, "
        "complementing Decree No. 2001-282 — indirectly controlling industrial emissions related to plastic "
        "production, plastic surface treatment and thermal disposal of plastic waste."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Total dust: 100 mg/m3 (D ≤ 1 kg/h); 50 mg/m3 (D > 1 kg/h); "
        "total organic compounds: 150 mg/m3 (D > 2 kg/h); "
        "waste incineration — dust: 10 mg/m3; organic matter (TOC): 50 mg/m3."
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Official technical standard adopted in October 2003 by the ASN/CT5 standardisation committee, published "
        "by the Senegalese Standardisation Association, complementing Decree No. 2001-282. Operative technical "
        "regulation of reference for atmospheric discharges (≠1.0 law)."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": "industry, environment, energy, transport, waste management",
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "production, disposal, environmental leakage",
    "policy_budget": 0,
    "policy_budget_text": (
        "No explicit budget, fee or economic instrument in the standard. Monitoring and measurement "
        "costs are borne by the operator (Ch. V, §5)."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "§8.2: open burning of tyres, plastics and any other compound containing chemical products is prohibited."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Ch. V §5: competent authority may request sampling and analysis of gaseous effluents at any time "
            "(+0.25 monitoring); costs borne by operator (+0.25 authority); "
            "§4.1: authority ensures emission limit values are complied with (+0.25 enforcement)."
        ),
        "comments": (
            "Explicit mention of plastics — direct prohibition of open burning (end-of-life, waste)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Annex II §K: provisions apply to installations for surface treatment of objects and products in "
            "metal, glass, ceramic, plastic materials, rubber or other materials by halogenated hydrocarbons. "
            "Mass flow of halogenated hydrocarbon emissions must not exceed 100 g/h (Annex I) and 25 g/h "
            "(Annex III); objects must be treated in a closed enclosure with vapour recovery."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "§K §2a: closed-enclosure requirement (+0.25 authority); "
            "§K §2c: vapour recovery system or equivalent (+0.25 monitoring); "
            "Ch. V §4: quarterly transmission of measurement results (+0.25 enforcement)."
        ),
        "comments": (
            "Explicit mention of plastic materials — industrial surface treatment (degreasing, coating)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "§8.1: incineration or thermal decomposition of waste is authorised only in installations "
            "technologically designed for that purpose; Annex II §J applies. "
            "Annex II §J §2: waste incineration emission limits — dust 10 mg/m3, organic matter (TOC) 50 mg/m3, "
            "SO2 50 mg/m3, NOx 80 mg/m3."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Annex II §J §3: continuous monitoring of temperature, oxygen and CO (+0.25 monitoring); "
            "§J §5.1: prohibition on incinerating urban/special waste in installations <350 kW (+0.25 enforcement); "
            "§J §6.1: trial burns required before incinerating particularly dangerous waste (+0.25 authority); "
            "Ch. V §4: quarterly reporting (+0.25 monitoring)."
        ),
        "comments": (
            "Strict thresholds for waste incinerators; direct relevance for thermal disposal of plastic waste."
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
