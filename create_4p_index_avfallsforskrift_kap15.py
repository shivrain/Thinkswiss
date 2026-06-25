#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften Kap. 15."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": "Avfallsforskriften – Kap. 15: Kommunalt avfallsgebyr",
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930/KAPITTEL_15",
    "policy_year": 2024,
    "policy_objective": (
        "To ensure municipal waste fees for statutory household waste handling are determined in "
        "accordance with Forurensningsloven §34, achieve full cost recovery without profit, prevent "
        "illegal cross-subsidisation between statutory household waste services and commercial waste "
        "services, and structure local financing of collection and treatment—including plastic household waste."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) issued by Miljøverndepartementet (1 June 2004); Chapter 15 "
        "reintroduced 8 September 2014 (in force 1 January 2015). Parent Avfallsforskriften last "
        "amended October 2024; Kap. 15 substantive provisions unchanged since §15-6 repeal (2021)."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "municipalities, households, waste management, recycling, consumption, packaging, retail"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "consumption, recycling, disposal, environmental leakage",
    "policy_budget": 1,
    "policy_budget_text": (
        '"The waste fee is determined to correspond to the total costs incurred by statutory handling of '
        'household waste. Full cost coverage must be ensured." (§15-3); '
        '"All costs of statutory handling of household waste shall be covered by the waste fee." (§15-4)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-1: Chapter purpose—to ensure waste charges for statutory household waste handling are "
            "determined in accordance with Forurensningsloven §34 first paragraph, and to prevent illegal "
            "cross-subsidisation between statutory household waste services and commercial waste services "
            "sold in the market (including industrial waste and inter-municipal services)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"The purpose of this chapter is to ensure that waste charges for statutory handling of '
            'household waste are determined in accordance with Section 34 first paragraph of the Pollution Act." '
            "(§15-1); cross-subsidisation prevention purpose (§15-1)"
        ),
        "comments": (
            "Governance/policy framing instrument; implements national fee rules at municipal level. "
            "Differentiated fees for recycling incentives are in Forurensningsloven §34(2), not Kap. 15."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-2: Scope—applies to determination of waste fees for statutory household waste handling "
            "under Forurensningsloven §29(3), §30 and §31(1). Excludes municipal waste services sold "
            "commercially in the market."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"This chapter applies to the determination of waste fees for statutory handling of household '
            'waste in the municipality, cf. Section 29 third paragraph … Section 30 and Section 31 first '
            'paragraph of the Pollution Act." (§15-2); market services excluded (§15-2)'
        ),
        "comments": "Defines which municipal fee-setting is covered; includes household waste streams with plastics.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-3: Municipal council (kommunestyret) establishes the waste fee. Fee must correspond to "
            "total costs of statutory household waste handling; full cost coverage required; municipality "
            "shall not profit; only costs and income from statutory household waste handling may be included."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"The municipal council establishes the waste fee." (§15-3); "The waste fee is determined to '
            'correspond to the total costs … Full cost coverage must be ensured. The municipality shall '
            'not have profits on such waste management." (§15-3)'
        ),
        "comments": "Core economic instrument; operative municipal fee-setting duty (skal/full cost coverage).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-4: Municipality shall maintain separate financial statements for statutory household waste "
            "handling (annual results and balance accounts), distinguishing statutory services from "
            "commercial market waste services."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"The municipality shall have separate financial statements for statutory handling of household '
            'waste. This means that the municipality for each financial year shall prepare separate financial '
            'statements for results and balance" (§15-4)'
        ),
        "comments": "Monitoring/accountability mechanism for fee revenue use; mandatory (skal).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-4: All costs of statutory household waste handling shall be covered by the waste fee. "
            "Where the municipality also sells waste services commercially, shared costs shall be "
            "distributed proportionately between statutory and market activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"All costs of statutory handling of household waste shall be covered by the waste fee." (§15-4); '
            '"costs common to the statutory handling … and for the waste services sold in the market shall '
            'be distributed proportionately." (§15-4)'
        ),
        "comments": "Cost-allocation rule prevents underfunding of household waste (incl. plastics) by market income.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§15-5: If the municipality has used waste fees to illegally subsidise competitive waste "
            "management activities in violation of this chapter, the municipality shall ensure that "
            "the subsidy is repaid (tilbakebetaling)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"If, through the waste fee, the municipality has subsidised the municipality\'s activities '
            'related to competitive waste management in violation of the rules here, the municipality '
            'shall ensure that this is returned." (§15-5)'
        ),
        "comments": "Corrective economic enforcement; repayment duty (skal sørge for) when violation found.",
    },
]

HEADERS = [
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


def main() -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index - Avfallsforskrift Kap15"

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)

    for col_idx, (col_letter, header) in enumerate(HEADERS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=f"{col_letter} — {header}")
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, inst in enumerate(INSTRUMENTS, start=2):
        policy_score = (
            POLICY["policy_type"]
            + POLICY["policy_integration"]
            + POLICY["policy_circularity"]
            + POLICY["policy_budget"]
            + inst["instrument_type"]
        ) / 5
        instrument_score = (inst["instrument_type"] + inst["instrument_implementation"]) / 2

        row_values = [
            POLICY["policy_name"],
            POLICY["policy_url"],
            POLICY["policy_year"],
            POLICY["policy_objective"],
            POLICY["policy_target"],
            POLICY["policy_target_text"],
            POLICY["policy_type"],
            POLICY["policy_type_justification"],
            POLICY["policy_integration"],
            POLICY["policy_sectors_list"],
            POLICY["policy_circularity"],
            POLICY["policy_lifecycle_phases_list"],
            POLICY["policy_budget"],
            POLICY["policy_budget_text"],
            round(policy_score, 3),
            inst["instrument_type"],
            inst["instrument_lifecycle_stage"],
            inst["instrument_description"],
            inst["instrument_in_force"],
            inst["instrument_implementation"],
            inst["instrument_implementation_text"],
            round(instrument_score, 3),
            inst["comments"],
        ]

        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 30

    widths = {
        "A": 48, "B": 42, "C": 12, "D": 45, "E": 12, "F": 20, "G": 12, "H": 40,
        "I": 16, "J": 42, "K": 16, "L": 35, "M": 14, "N": 45, "O": 12, "P": 14,
        "Q": 20, "R": 55, "S": 14, "T": 18, "U": 45, "V": 14, "W": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path = "/workspace/4P_Index_Avfallsforskriften_Kap15.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
