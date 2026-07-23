#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for PGDBM May 2019 (English only)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Plan_Gestion_Dechets_Biomedicaux_2019.xlsx"

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
    "policy_name": "Biomedical waste management plan — updated, May 2019 (REDISSE / ISMEA)",
    "policy_url": "https://www.sante.gouv.sn/sites/default/files/plan_gestion_dechets_biom%C3%A9dicaux_0.pdf",
    "policy_year": 2019,
    "policy_objective": (
        "Update the national biomedical waste management plan to require source segregation, colour-coded plastic "
        "collection receptacles (bags, bins, safety boxes) in health facilities, treatment of plastic-packaged "
        "clinical waste, and prohibition of incinerating halogenated plastics (PVC) in biomedical waste streams."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "National operational plan updated May 2019 under the Ministry of Health (MSAS), funded by World Bank "
        "REDISSE and ISMEA. Planning document without binding quantifiable national targets (≠1.0 law, ≠0.75 decree)."
    ),
    "policy_integration": 0.50,
    "policy_sectors_list": "health, waste management, environment, municipalities, industry",
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": "consumption, disposal",
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Section V.1 cost table: e.g. 2,210 pre-collection units (FCFA 22.1M REDISSE + 17.5M ISMEA); "
        "2,000 plastic waste bags (FCFA 400,000); incinerators and sterilisers budgeted under donor projects."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Section V.B / IV.E: mandatory source segregation — waste must be sorted at production and placed "
            "immediately in the correct container; separates contaminated (~15% WHO benchmark) from non-hazardous waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "MSAS/PRONALIN designated to lead implementation; CLIN/CHSCT facility-level supervision planned. "
            "No fines or penalties in the plan itself (+0.25 authority, +0.25 monitoring via supervision)."
        ),
        "comments": "Mandatory sorting procedure for clinical waste streams including plastic containers.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Section IV.E: colour-coded plastic receptacles — black plastic bags/bins for general waste "
            "(packaging, plastic bottles); yellow plastic bags for infectious waste; red bins for pathological "
            "waste; yellow safety boxes for sharps (OPCT)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Receptacles must be non-transparent, moisture-resistant and closed; plan budgets 2,000 plastic bags. "
            "PRONALIN designated (+0.25 authority); facility assessments and CLIN monitoring (+0.25 monitoring)."
        ),
        "comments": "Most explicit plastic mention — colour-coded plastic bags, bins and safety boxes.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Section VI.1: biomedical waste that cannot be incinerated includes 'Halogenated Plastics (PVC)'; "
            "pre-sorting required before incineration."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "MSAS-led plan sets prohibition on PVC incineration (+0.25 authority designated). "
            "No enforcement penalties specified in the plan itself."
        ),
        "comments": "Direct regulatory prohibition on incinerating halogenated clinical plastics (PVC).",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Section V.A.2 / IV.2.3: provision of treatment infrastructure (incinerators, autoclaves/shredders, "
            "controlled landfill cells) for biomedical waste including plastic-packaged clinical waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Infrastructure investment budgeted under REDISSE/ISMEA (+0.25 authority). "
            "No enforcement or monitoring mechanism in plan text for infrastructure delivery."
        ),
        "comments": "Infrastructure for disposal of plastic-packaged biomedical waste after segregation/treatment.",
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
