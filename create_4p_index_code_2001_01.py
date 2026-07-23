#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi n° 2001-01 (Environmental Code).

English only; plastic-relevant provisions only; strict plastics-relevance filter.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2001-01_Code_Environnement.xlsx"

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
    "policy_name": "Law No. 2001-01 of 12 April 2001 establishing the Environmental Code",
    "policy_url": "https://www.denv.gouv.sn/telechargement/69/codes/18921/code-de-lenvironnement-2001.pdf",
    "policy_year": 2001,
    "policy_objective": (
        "Establish foundational environmental protection principles and obligations covering "
        "pollution control, waste management, product regulation, environmental assessment, "
        "and decentralised State/local responsibilities — providing the overarching legal "
        "framework through which plastic waste and pollution are indirectly governed, without "
        "plastic-specific targets or measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1.0,
    "policy_type_justification": (
        "Foundational environmental code enacted by the National Assembly (29 December 2000) "
        "and Senate (4 January 2001) and promulgated by the President as Law No. 2001-01; "
        "legislation adopted by parliament, not an executive regulation."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "environment, industry, agriculture, municipalities, fisheries, water, chemicals, "
        "waste management, urban planning, packaging"
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. L26: works charged to the environmental protection fund when polluter is unidentified. "
        "Art. L27: annual pollution and facility taxes on classified installations, including pollution "
        "taxes calculated according to set FCFA schedules."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. L30–L31: all waste categories must be eliminated or recycled in an ecologically rational "
            "manner; producers/holders must ensure elimination or recycling themselves or via "
            "Ministry-licensed enterprises, or deliver waste to local authorities or State-licensed operators; "
            "recycling must comply with national standards."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L31: licensed-enterprise requirement and local-authority delivery option (+0.25 authority); "
            "L38: removal at polluter's cost after formal notice (+0.25 enforcement); "
            "Title IV criminal/administrative sanctions (+0.25 monitoring). "
            "No explicit plastic mention; applies to all solid waste including plastics."
        ),
        "comments": (
            "Foundational producer/holder waste obligation. Law 2023-15 replaced Law 2001-01 on 2 August 2023; "
            "substantive waste obligations largely carried forward."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. L32 and L34: local authorities and constituted groupings must ensure elimination of "
            "household waste, potentially with regional/national State services, and may collect/treat "
            "non-household waste; conditions for waste collection, sorting, storage, transport, recovery, "
            "reuse, recycling, treatment and final disposal are set by ministerial order."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "L32: responsible authority designated; optional special fee for non-household waste (+0.25 authority); "
            "L34: enabling provision operationalised through Decree 2001-282 and subsequent waste regulations (+0.25 monitoring). "
            "No article-specific penalty in L32."
        ),
        "comments": (
            "Key governance instrument for municipal plastic waste collection; complements decentralised environment competences."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Arts. L35–L36: absolute prohibition on depositing waste on the public domain, including "
            "maritime public domain; public-domain concessionaires must eliminate or recycle waste found there; "
            "local authorities must stop illegal dumping and eliminate abandoned waste of unidentified ownership."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L35: mandatory ban language ('absolutely prohibited') (+0.25 authority); "
            "L36: local authorities must monitor and remediate illegal dumps (+0.25 monitoring); "
            "L38 and Title IV sanctions for violations (+0.25 enforcement)."
        ),
        "comments": (
            "Directly relevant to littering and marine plastic leakage."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. L38: when waste is abandoned, deposited or treated contrary to the Code, the police "
            "authority must — after formal notice — remove it at the responsible party's expense and may "
            "require an escrow deposit with a public accountant."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "L38: mandatory, unconditional enforcement — police authority must remove waste at polluter's expense "
            "after formal notice (+0.25 authority); must require escrow deposit for works cost (+0.25 enforcement); "
            "cross-cutting application to all illegal waste including plastics (+0.25 monitoring); "
            "Title IV penal sanctions (+0.25 monitoring)."
        ),
        "comments": (
            "Cross-cutting enforcement instrument for illegal waste disposal including plastic litter."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Arts. L40–L41: manufacture, import, possession for sale and supply of products or materials "
            "generating waste must be regulated by joint ministerial order to facilitate waste elimination "
            "or, if necessary, prohibit them; immersion, incineration or any disposal of waste in continental, "
            "marine or fluvio-maritime waters under Senegalese jurisdiction is prohibited."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L40: enabling power exercised for plastics through Law No. 2020-04 (+0.25 authority); "
            "L41: mandatory aquatic-waste ban (+0.25 enforcement); "
            "L97: marine/continental water pollution penalties — FCFA 500,000–2,000,000 and 6 months to 1 year "
            "imprisonment, doubled on recurrence (+0.25 monitoring)."
        ),
        "comments": (
            "Product-regulation enabling power (operationalised via Law 2020-04) and direct marine plastic "
            "pollution prohibition."
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

        g_col = get_column_letter(7)
        i_col = get_column_letter(9)
        k_col = get_column_letter(11)
        m_col = get_column_letter(13)
        p_col = get_column_letter(16)
        t_col = get_column_letter(20)
        o_col = get_column_letter(15)
        v_col = get_column_letter(22)

        ws[f"{o_col}{row_idx}"] = f"=AVERAGE({g_col}{row_idx},{i_col}{row_idx},{k_col}{row_idx},{m_col}{row_idx},{p_col}{row_idx})"
        ws[f"{v_col}{row_idx}"] = f"=AVERAGE({p_col}{row_idx},{t_col}{row_idx})"
        ws[f"{o_col}{row_idx}"].fill = auto_fill
        ws[f"{v_col}{row_idx}"].fill = auto_fill

    widths = {
        "A": 42, "B": 48, "C": 10, "D": 44, "E": 12, "F": 20,
        "G": 10, "H": 36, "I": 14, "J": 36, "K": 14, "L": 28,
        "M": 12, "N": 36, "O": 12, "P": 14, "Q": 22, "R": 52,
        "S": 14, "T": 18, "U": 44, "V": 14, "W": 44,
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
