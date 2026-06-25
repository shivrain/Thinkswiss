#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften Kap. 4 (end-of-life vehicles)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Avfallsforskriften – Kap. 4: Kasserte kjøretøy "
        "(Waste Regulation – Ch. 4: End-of-life vehicles)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930/KAPITTEL_4",
    "policy_year": 2024,
    "policy_objective": (
        "To prevent and reduce environmental problems from end-of-life vehicles (ELVs), including plastic "
        "components such as bumpers and interior parts, through extended producer responsibility, free return "
        "systems, depollution and dismantling at licensed treatment facilities, reuse, material recovery and "
        "energy recovery, supported by a national scrap-deposit (vrakpant) scheme."
    ),
    "policy_target": 1,
    "policy_target_text": (
        '"Producers shall by 1 January 2006 ensure that 85 %, measured by weight, of their proportional share '
        'of ELVs is recovered, of which at least 80 percentage points shall be materially recovered." (§4-4); '
        '"Producers shall by 1 January 2015 ensure that in total 95 % … is recovered, of which at least 85 '
        'percentage points shall be materially recovered." (§4-4)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) issued 1 June 2004 under forurensningsloven; Chapter 4 implements "
        "EU End-of-Life Vehicles Directive 2000/53/EC. Amended July 2014 (treatment requirements), May 2021 "
        "(vrakmelding), and June 2024 (approval/annual fees for return systems, §4-5a–§4-5c). Parent "
        "Avfallsforskriften last substantively amended October 2024."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "automotive, waste management, producers, recycling, manufacturing, retail, public administration"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": "consumption, end-of-life, recycling, disposal, environmental leakage",
    "policy_budget": 1,
    "policy_budget_text": (
        '"Producer obligations shall be fulfilled through participation in an approved return system." (§4-4); '
        '"Anyone shall be able to deliver free of charge to the return system any scrapped vehicle." (§4-5); '
        'scrap deposit (vrakpant) paid by the tax office when ELVs are delivered to approved plants (§4-8, §4-12)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-1: Purpose—to prevent and reduce environmental problems caused by vehicles when they become "
            "waste, framing national ELV policy including recovery of plastic and other material fractions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"The purpose of the provisions in this chapter is to prevent and reduce environmental problems '
            'that vehicles cause when they end up as waste." (§4-1)'
        ),
        "comments": "Policy objective instrument for the ELV chapter.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-2: Scope—return-system and scrap-deposit provisions apply to M1/N1 vehicles, certain buses, "
            "campers and caravans; treatment requirements in Section III apply to all vehicles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Section II Return system for scrapped vehicles … and Section IV Completion of scrap notices and '
            'payment of scrap deposit … apply to vehicles in group M1 and N1 …" (§4-2); '
            '"Section III Treatment of scrapped vehicles applies to all vehicles." (§4-2)'
        ),
        "comments": "Defines which vehicle categories are covered by EPR and depollution rules.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-4: Producer EPR—producers must ensure environmentally sound collection and treatment of ELVs "
            "in proportion to market share; meet recovery targets of 85 % by 2006 (≥80 % material recovery) and "
            "95 % by 2015 (≥85 % material recovery); provide public information; fulfil duties via approved "
            "return system."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Producers are obliged to ensure environmentally sound collection and treatment of scrapped '
            'vehicles." (§4-4); "Producers shall by 1 January 2015 ensure that in total 95 %, measured by weight, '
            'of their proportional share … is recovered, of which at least 85 percentage points shall be materially '
            'recovered." (§4-4)'
        ),
        "comments": "Core EPR instrument with binding quantitative recovery targets (plikt/plikter, skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-5: Approved return system—must be pre-approved by Miljødirektoratet; anyone may deliver any "
            "scrapped in-scope vehicle free of charge; system must have good geographic coverage, open and "
            "non-discriminatory producer access, and documented financial security."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Return systems for scrapped vehicles shall be pre-approved by the Norwegian Environment Agency." '
            "(§4-5); "
            '"Anyone shall be able to deliver free of charge to the return system any scrapped vehicle." (§4-5); '
            '"All producers shall have access to participate in the return system." (§4-5)'
        ),
        "comments": "Nationwide free take-back infrastructure; Miljødirektoratet kan trekke godkjenning.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-6: Producer documentation—vehicle producers and importers must be able to document that "
            "vehicles they place on the market are linked to an approved return system, and present documentation "
            "to Miljødirektoratet on request."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Anyone who produces or imports vehicles shall be able to document that the vehicles produced or '
            'imported are linked to an approved return system." (§4-6)'
        ),
        "comments": "Accountability mechanism for EPR participation.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§4-7 + Annex 1: Licensed ELV treatment plants must hold a pollution permit and meet minimum "
            "technical requirements for depollution, fluid removal, component dismantling and storage; marked "
            "hazardous components must be removed; larger plastic components shall be dismantled unless recovered "
            "in subsequent shredding."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Anyone operating a treatment plant for scrapped vehicles must have a separate permit from the '
            'pollution authority." (§4-7); Annex 1: "dismantling of glass and larger plastic components unless '
            'these materials are separated out in subsequent fragmentation." (Annex 1, item 4)'
        ),
        "comments": "Direct plastics-relevant depollution/dismantling rule; mandatory minimum treatment standard.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-8: Scrap deposit (vrakpant)—deposit payable for scrapped vehicles registered after 1 January 1977 "
            "(or still registered on that date); minimum requirement is delivery of intact frame/chassis/body with "
            "readable VIN to an approved treatment plant."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Scrap deposit is paid for scrapped vehicles that were registered (licensed) after 1 January 1977." '
            "(§4-8); minimum requirement is delivery to approved treatment plant with intact frame/chassis/body "
            "where the chassis number is stamped (§4-8)"
        ),
        "comments": "Economic return incentive driving ELVs—including plastic-rich fractions—to licensed treatment.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-10: Scrap notice (vrakmelding)—on receipt of an ELV, treatment plants must complete the official "
            "scrap notice with VIN verification, owner identification and registration documentation; foreign "
            "scrap notices accepted for ELVs scrapped in other EEA states."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"On receipt of a scrapped vehicle at a treatment plant, a notice about the vehicle shall be completed '
            'on a form prescribed by the Norwegian Tax Administration." (§4-10)'
        ),
        "comments": "Administrative traceability instrument linking vehicles to treatment and deposit payout.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-11–§4-12: Scrap-notice administration and deposit payout—plants must journal and transmit "
            "notices to the tax office; tax office pays scrap deposit when conditions are met at rates set by "
            "annual parliamentary budget decisions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Scrap notices shall be registered at treatment plants in ascending serial-number order." (§4-11); '
            '"The tax office pays scrap deposit when the conditions are met." (§4-12); deposit rate set by annual '
            "Storting budget decisions (§4-12)"
        ),
        "comments": "Operative deposit-payment chain; Skattedirektoratet determines payment method and timing.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-14: Reporting—vehicle producers must be able to document how chapter obligations are fulfilled "
            "and present documentation to Miljødirektoratet on request; Miljødirektoratet may issue reporting guidelines."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Vehicle producers shall be able to document how the obligations under this chapter are fulfilled." '
            "(§4-14); documentation shall be submitted to Miljødirektoratet on request (§4-14)"
        ),
        "comments": "Monitoring instrument; Miljødirektoratet kan fastsette retningslinjer (enabling).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4-15: Design-for-environment link—provisions on phase-out of components, marking and information "
            "duties are set out in produktforskriften §2-19 (restrictions on hazardous substances in vehicles, "
            "including plastics-related substances)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Provisions on phase-out of components, marking and information duties are set out in Regulation '
            '1 June 2004 No. 922 on restriction of hazardous chemicals and other products §2-19." (§4-15)'
        ),
        "comments": "Cross-reference to upstream vehicle design and marking rules affecting plastic components.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-16: Supervision—Miljødirektoratet supervises producer/return-system duties (§4-4–§4-6); "
            "County Governor supervises treatment plants (§4-7); tax authorities supervise scrap deposit provisions "
            "(§4-8–§4-13)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"The Norwegian Environment Agency supervises compliance with §4-4 to §4-6." (§4-16); '
            '"The County Governor supervises compliance with §4-7." (§4-16)'
        ),
        "comments": "Multi-agency enforcement framework.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-17: Coercive fines—the tax office may impose coercive fines under produktkontrolloven §13 or "
            "forurensningsloven §73 within its supervision area; otherwise Chapter 17 rules apply."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"The tax office may decide on a coercive fine under the Product Control Act §13 or the Pollution '
            'Act §73 within its area of supervision." (§4-17)'
        ),
        "comments": "Enforcement instrument (kan); supports compliance with ELV recovery chain.",
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
    ws.title = "4P Index - Avfallsforskr Kap4"

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
        "A": 52, "B": 42, "C": 12, "D": 45, "E": 12, "F": 40, "G": 12, "H": 40,
        "I": 16, "J": 42, "K": 16, "L": 35, "M": 14, "N": 45, "O": 12, "P": 14,
        "Q": 20, "R": 55, "S": 14, "T": 18, "U": 45, "V": 14, "W": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path = "/workspace/4P_Index_Avfallsforskriften_Kap4.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
