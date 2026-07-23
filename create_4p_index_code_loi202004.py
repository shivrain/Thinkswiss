#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2020-04 (English only)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2020-04_Produits_Plastiques.xlsx"

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
        "Law No. 2020-04 of 8 January 2020 on prevention and reduction of "
        "environmental impacts of plastic products"
    ),
    "policy_url": "https://www.au-senegal.com/IMG/pdf/loi-plastique-senegal-2020-04.pdf",
    "policy_year": 2020,
    "policy_objective": (
        "Flagship plastics law: ban single-use/disposable plastic products and checkout bags; "
        "establish plastic-bottle deposit-return and extended producer responsibility; set recycled-content "
        "targets; ban plastic-waste imports; institute recycling floor price and plastic tax on non-recyclables."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Art. 4: prohibits single-use or disposable plastic products (cups, cutlery, straws, sachets); "
        "Art. 5: prohibits checkout plastic bags regardless of thickness; "
        "Art. 16: producers must integrate a share of recycled plastic in new products (targets set by decree); "
        "Art. 19: import of plastic waste into national territory is prohibited."
    ),
    "policy_type": 1.0,
    "policy_type_justification": (
        "Law adopted by the National Assembly on 30 December 2019 and promulgated 8 January 2020 as Law No. 2020-04. "
        "Parliamentary legislation (≠0.75 executive regulation)."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "production, consumption, retail, waste management, recycling, industry, packaging, environment"
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. 21: floor price for recyclers buying plastic waste (set by decree); "
        "Art. 22: plastic tax on non-recyclable plastic products (rates by decree); "
        "Arts. 26–39: criminal fines up to FCFA 100,000,000 and imprisonment up to 5 years."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 4: prohibits production, import, detention for sale, sale, supply and use of single-use or "
            "disposable plastic products, including cups/glasses/lids, cutlery/plates, straws/stirrers, "
            "and water/beverage sachets."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 26: manufacture/import — 1–3 years imprisonment and FCFA 5–10 million fine (+0.25 enforcement); "
            "Art. 25: control agents from Environment, Health, Industry, Commerce, Finance ministries (+0.25 authority); "
            "Art. 23: seizure of prohibited products (+0.25 monitoring). Art. 4 exception for primary food packaging."
        ),
        "comments": "Central single-use plastic ban. Exception for transparent recyclable primary food packaging reduces unconditionality.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 5: checkout plastic bags are prohibited regardless of thickness "
            "(including biodegradable/oxo-degradable bags)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 27: sale/use offence — 1–3 months imprisonment and FCFA 50,000–100,000 (+0.25 enforcement); "
            "Art. 25: designated control ministries (+0.25 authority); Art. 5(2) exception for transparent "
            "recyclable bags at point of sale for food protection (+0.25 monitoring via controls)."
        ),
        "comments": "Checkout bag ban with narrow food-packaging exception.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Arts. 6–7: deposit required on purchase of products in plastic bottles; amount set by decree; "
            "sellers must accept returned bottles and deliver to nearest collection point."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 33: refusal penalty — 15 days to 1 month imprisonment and FCFA 50,000–100,000 (+0.25 enforcement); "
            "deposit amount and modalities set by subordinate decree (+0.25 authority). No monitoring mechanism in law text."
        ),
        "comments": "Deposit-return scheme; operative but deposit amount requires implementing decree.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Arts. 8–10: producers must establish collection points for plastic bottles and recover collected "
            "bottles prioritising reuse, recycling then other recovery; biannual sectoral report to Environment Minister."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 34: insufficient collection points — 3–6 months and FCFA 5–10 million (+0.25 enforcement); "
            "Art. 9–10: biannual reporting on market vs collected volumes (+0.25 monitoring); "
            "producer obligations mandatory (+0.25 authority, +0.25 unconditional)."
        ),
        "comments": "Bottle-producer collection and recovery obligations.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 11: producers placing plastic products on the market are responsible for managing waste "
            "generated by these products (extended producer responsibility); compliance via individual programmes "
            "or accredited eco-organisms (Arts. 12–14)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 35: EPR non-compliance — 1–3 years and FCFA 10–20 million (+0.25 enforcement); "
            "Arts. 12–14: ministerial approval, periodic controls, annual reporting (+0.25 monitoring, +0.25 authority)."
        ),
        "comments": "EPR framework including operational mechanisms in Arts. 12–14.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Arts. 15–16: producers must reduce waste at source and, when technically and economically viable, "
            "integrate a share of recycled plastic in new plastic products; decree sets national targets and deadlines."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 32: non-compliance — 1–3 months and FCFA 5–10 million when viable (+0.25 enforcement); "
            "specific targets set by subordinate decree (+0.25 authority). Viability condition reduces unconditionality."
        ),
        "comments": "Recycled-content targets; operative framework but numeric targets in implementing decree.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 19: import of plastic waste into national territory is prohibited; illegal imports seized and "
            "re-exported. Art. 20: exports only with Environment Minister authorisation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 28–29: illegal import/unauthorised export — 3–5 years and FCFA 50–100 million (+0.25 enforcement); "
            "Art. 25: control agents designated (+0.25 authority); seizure mechanism (+0.25 monitoring)."
        ),
        "comments": "Cross-border plastic-waste controls.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 21: floor price at which recycling enterprises must buy plastic waste per kg (set by decree). "
            "Art. 22: plastic tax on non-recyclable plastic products (product list and rates by decree)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Art. 31: buying below floor price — FCFA 2–5 million fine (+0.25 enforcement); "
            "rates and product lists set by decree (+0.25 authority). No monitoring mechanism in law text."
        ),
        "comments": "Economic instruments; operative framework pending implementing decrees for rates.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Arts. 23–25, 37: seizure of prohibited products; consumers must deliver plastic waste to designated "
            "collection points; abandoning waste outside collection points penalised."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Art. 37: abandoning plastic waste — 15 days to 1 month and FCFA 20,000–50,000 (+0.25 enforcement); "
            "Art. 39: corporate penalties including closure and confiscation (+0.25 authority, +0.25 monitoring)."
        ),
        "comments": "Enforcement and consumer delivery obligations combined as one enforcement cluster.",
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
