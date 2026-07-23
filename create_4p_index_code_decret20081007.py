#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2008-1007 (English only).

Plastic clinical/biomedical waste provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2008-1007_Dechets_Biomedicaux.xlsx"

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
        "Decree No. 2008-1007 of 18 August 2008 regulating biomedical waste management"
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/"
        "19089/decret-portant-reglementation-de-la-gestion-des-dechets-biomedicaux.pdf"
    ),
    "policy_year": 2008,
    "policy_objective": (
        "Establish classification, source segregation, packaging, storage, transport, treatment and disposal "
        "rules for biomedical waste (infectious, anatomical, sharps, pharmaceutical, recyclable, special) "
        "by imposing operator approval and traceability — explicitly including healthcare plastic waste "
        "(bottles, packaging) without measures specific to non-biomedical plastics."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Art. 8–9: storage duration must not exceed 48 hours; storage room decontaminated at least once per week; "
        "Art. 12: minimum required incineration temperature is 800°C."
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree adopted by the President of the Republic on 18 August 2008, on the report of "
        "the Minister of Health and Prevention, implementing the Hygiene Code, Environmental Code and "
        "hospital laws. Sub-legislative implementing regulation, not an act enacted by Parliament (≠1)."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "health, environment, local government, industry, research, waste management"
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": "consumption, recycling, disposal",
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Arts. 4–5: waste classification includes recyclable plastic waste such as serum bottles, "
            "bleach bottles, glass equipment and pressurised containers; any person producing or holding "
            "biomedical waste must ensure elimination or recycling itself or via Health Ministry-approved companies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 6: all waste sorted at place of production into category-specific circuits with pictograms (+0.25 monitoring); "
            "Art. 14: every biomedical waste operator must obtain Health Ministry approval (+0.25 authority); "
            "Art. 20: producers and operators must record waste and ensure traceability (+0.25 enforcement)."
        ),
        "comments": (
            "Explicit mention of recyclable healthcare plastic waste (bottles, containers)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. 6–7: all medical, pharmaceutical, veterinary or research waste is sorted at the place "
            "of production and placed in the dedicated category circuit; packaging is carried out at production "
            "to avoid health and environmental risk; reusable infectious-waste containers must be cleaned and "
            "disinfected after each use; container colours and labelling fixed by Health Ministry order."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 7: packaging colours and labelling set by ministerial order (+0.25 authority); "
            "Art. 16: producers/operators must have appropriate packaging equipment (+0.25 monitoring). "
            "No article-specific penalty for non-segregation in these articles."
        ),
        "comments": (
            "Mandatory source segregation separating recyclable plastics from infectious healthcare waste."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Arts. 11–12: infectious and sharps waste must be incinerated, landfilled or undergo "
            "sterilisation/shredding treatment; recyclable waste follows the municipal waste stream; "
            "incinerator installation and operation must comply with ICPE regulations and achieve a "
            "minimum incineration temperature of 800°C; incineration residues must undergo sanitary landfilling."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 12: ICPE conformity required for incinerators (+0.25 authority); "
            "Art. 13: treatment activity must undergo environmental and social impact study (+0.25 monitoring); "
            "Art. 21: approval withdrawal or 1–12 month suspension for dangerous operation (+0.25 enforcement)."
        ),
        "comments": (
            "Differentiated treatment pathways; shredding/sterilisation and 800°C incineration relevant for "
            "infectious medical plastic waste; recyclables routed to municipal stream."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Arts. 8–10 and preamble: biomedical waste stored in ventilated secure rooms decontaminated at "
            "least weekly, for no more than 48 hours; external transport in hermetic tamper-proof containers "
            "during low-traffic periods; depositing healthcare waste in uncontrolled areas can contaminate "
            "surface waters; poor incineration can emit dioxins and furans."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 10: hermetic GRV-type container or specially equipped vehicle for external transport (+0.25 authority); "
            "Art. 16: appropriate transport equipment required (+0.25 monitoring); "
            "Art. 21: administrative sanctions — approval withdrawal or 1–12 month suspension (+0.25 enforcement)."
        ),
        "comments": (
            "Quantitative storage/transport rules and leakage prevention for healthcare waste including plastics."
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
