#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften Kap. 5 (discarded tyres)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Avfallsforskriften – Kap. 5: Kasserte dekk "
        "(Waste Regulation – Ch. 5: Discarded tyres)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930/KAPITTEL_5",
    "policy_year": 2024,
    "policy_objective": (
        "To reduce environmental problems from discarded tyres—including synthetic rubber and plastic-polymer "
        "fractions—when they end up in landfill or the environment, by banning tyre landfilling, requiring free "
        "retailer take-back and producer/importer collection, ensuring recovery through reuse, material recycling "
        "or energy recovery, and annual reporting to reduce tyre waste and related litter/microplastic pollution."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) issued 1 June 2004 under produktkontrolloven §4. Chapter 5 establishes "
        "national EPR for motor-vehicle and trailer tyres. Amended June 2010 and March 2013 (reporting). "
        "§5-7a–§5-7b on annual and data-system fees added June 2024 (in force 1 July 2024). Parent "
        "Avfallsforskriften last substantively amended October 2024."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "automotive, transport, waste management, retail, producers, recycling, manufacturing, environmental leakage"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "consumption, end-of-life, recycling, disposal, environmental leakage, litter/pollution"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        '"Producers/importers have a duty to collect discarded tyres free of charge in geographic areas where '
        'their tyres are sold." (§5-5); producers/importers or third parties shall pay annual fee to the Treasury '
        'for Miljødirektoratet reporting oversight (§5-7a) and cover data-system costs (§5-7b)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-1: Purpose—to reduce environmental problems tyres cause when ending as landfill waste by "
            "ensuring a high degree of recovery of discarded tyres, including synthetic rubber/plastic polymer fractions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"The purpose of the provisions in this chapter is to reduce the environmental problems tyres cause '
            'when they end up as waste in landfill etc., by ensuring a high degree of recovery of discarded tyres." '
            "(§5-1)"
        ),
        "comments": "Qualitative recovery objective; no quantified percentage target in Kap. 5.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-2: Scope—chapter regulates collection and recovery of discarded tyres for motor vehicles and "
            "trailers/trailer equipment under kjøretøyforskriften Chapter 2."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"The provisions in this chapter regulate collection and recovery of discarded tyres." (§5-2)'
        ),
        "comments": "Defines national tyre EPR scope.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§5-3: Definitions—tyres means tyres for motor vehicles and trailers; recovery means reuse, material "
            "recovery and energy recovery of discarded tyres."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Recovery means in this chapter utilisation of discarded tyres in the form of reuse, material recovery '
            'and energy recovery." (§5-3)'
        ),
        "comments": "Defines eligible recovery routes for polymer-rich tyre waste.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "§5-4: Landfill ban—it is prohibited to landfill tyres at landfill sites (fyllplass), diverting tyre "
            "waste—including rubber and plastic-polymer fractions—from disposal to recovery streams."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"It is prohibited to landfill tyres at landfill sites." (§5-4)'
        ),
        "comments": "Binding prohibition (forbudt); reduces tyre dumping and related pollution.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-5: Retailer take-back—tyre retailers must accept discarded tyres free of charge in return, limited "
            "to a reasonable quantity of the tyre category the retailer sells."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"A retailer of tyres has a duty to accept discarded tyres free of charge in return. The duty is limited '
            'to a reasonable quantity of the category of tyres the retailer sells." (§5-5)'
        ),
        "comments": "Point-of-sale collection instrument; mandatory plikt.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-5: Producer/importer collection—producers and importers must collect discarded tyres free of charge "
            "in geographic areas where their tyres are sold, limited to reasonable quantities; duty may be fulfilled "
            "by a third party."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"A producer/importer has a duty to collect discarded tyres free of charge in corresponding geographic '
            'areas of the country where the producer\'s/importer\'s tyres are sold." (§5-5); collection duty may be '
            "fulfilled by third party (§5-5)"
        ),
        "comments": "Core EPR collection obligation for nationwide tyre take-back.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§5-6: Recovery obligation—tyre producers and importers must ensure that discarded tyres they are "
            "responsible for under §5-5 are recovered (reuse, material recovery or energy recovery); may be "
            "fulfilled by third party."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Producers/importers of tyres have a duty to ensure that the discarded tyres they are responsible for '
            'under §5-5 are recovered." (§5-6)'
        ),
        "comments": "Binding recovery duty; key operative EPR instrument for tyre recycling.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-7: Annual reporting—producers/importers shall report annually to Miljødirektoratet on tyre "
            "production, import, collection and recovery; Miljødirektoratet may issue reporting guidelines; may be "
            "fulfilled by third party."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Producers/importers shall annually report to the Norwegian Environment Agency on production and '
            'import of tyres and on collection and recovery of discarded tyres." (§5-7)'
        ),
        "comments": "Monitoring and transparency; Miljødirektoratet kan fastsette retningslinjer (enabling).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-7a: Annual fee—producers/importers or third parties shall pay an annual fee (NOK 20,400) to the "
            "Treasury for Miljødirektoratet's work on annual reporting oversight."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Producers/importers or third parties shall pay an annual fee to the Treasury for the Norwegian '
            'Environment Agency\'s work on annual reporting." (§5-7a); annual fee is NOK 20,400 (§5-7a)'
        ),
        "comments": "Economic instrument financing EPR administration; added 2024.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-7b: Data-system costs—producers/importers or third parties shall cover Miljødirektoratet's costs "
            "for developing and operating data systems necessary to ensure compliance with chapter EPR requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Producers/importers or third parties shall cover the Norwegian Environment Agency\'s costs for '
            'development and operation of data systems necessary to ensure that requirements linked to producer '
            'responsibility in this chapter are complied with." (§5-7b)'
        ),
        "comments": "Producer-financed EPR data infrastructure; Miljødirektoratet fastsetter gebyr (enabling).",
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
    ws.title = "4P Index - Avfallsforskr Kap5"

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
        "A": 52, "B": 42, "C": 12, "D": 45, "E": 12, "F": 20, "G": 12, "H": 40,
        "I": 16, "J": 42, "K": 16, "L": 35, "M": 14, "N": 45, "O": 12, "P": 14,
        "Q": 20, "R": 55, "S": 14, "T": 18, "U": 45, "V": 14, "W": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path = "/workspace/4P_Index_Avfallsforskriften_Kap5.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
