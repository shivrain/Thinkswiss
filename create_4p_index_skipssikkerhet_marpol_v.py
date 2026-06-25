#!/usr/bin/env python3
"""Generate 4P Index Excel for Skipssikkerhetsforskriften §11 (MARPOL Annex V)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om miljømessig sikkerhet for skip – § 11: MARPOL vedlegg V "
        "(Regulation on environmental safety for ships – § 11: MARPOL Annex V (garbage))"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2012-05-30-488/%C2%A711",
    "policy_year": 2025,
    "policy_objective": (
        "Incorporates MARPOL consolidated edition 2022 Annex V on prevention of garbage pollution from "
        "ships and mobile offshore units into Norwegian law, prohibiting discharge of garbage including "
        "all plastics at sea, requiring onboard storage, garbage management plans, placards and record-keeping, "
        "and delivery to port reception facilities—with national supplements for Antarctic operations and "
        "offshore petroleum units."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Utslipp i sjøen av all plast, herunder men ikke begrenset til syntetiske taustumper, syntetiske '
        'fiskegarn, plastsøppelposer og forbrenningsrester fra plastprodukter, er forbudt." (MARPOL vedlegg V '
        'regel 3 nr. 2, via §11) / '
        'EN: "Discharge into the sea of all plastics, including but not limited to synthetic ropes, synthetic '
        'fishing nets, plastic garbage bags and incinerator ashes from plastic products is prohibited." '
        "(MARPOL Annex V regulation 3.2, via §11)"
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under skipssikkerhetsloven, småbåtloven and produktkontrolloven; "
        "implements MARPOL 73/78 Annex V as amended by MEPC.360(79) and MEPC.382(80). Section originally "
        "in force July 2012; MARPOL Annex V updated November 2023 and December 2024 (FOR-2024-12-19-3397)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "maritime, offshore, fisheries, shipping, ports, aquaculture, leisure craft, mobile offshore units, "
        "petroleum, cruise tourism"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "use/consumption, litter/pollution, environmental leakage, end-of-life"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Et skip som befinner seg i en havn eller offshoreterminal i en annen parts jurisdiksjon, er '
        'underlagt inspeksjon … når det foreligger klare grunner til å tro at skipsføreren eller mannskapet '
        'ikke er kjent med essensielle prosedyrer om bord … for å hindre forurensning fra søppel." '
        "(MARPOL vedlegg V regel 9 nr. 1, via §11) / "
        'EN: "A ship when in a port or an offshore terminal of another Party is subject to inspection … where '
        'there are clear grounds for believing that the master or crew are not familiar with essential shipboard '
        'procedures relating to the prevention of pollution by garbage." (MARPOL Annex V regulation 9.1, via §11)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§11: MARPOL Annex V incorporation—MARPOL consolidated edition 2022 Annex V on prevention of "
            "garbage pollution, as amended by MEPC.360(79) and MEPC.382(80), applies as Norwegian regulation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "MARPOL konsolidert utgave 2022 vedlegg V om hindring av søppelforurensning som endret ved '
            'MEPC.360(79) og MEPC.382(80) gjelder som forskrift." (§11) / '
            'EN: "MARPOL consolidated edition 2022 Annex V on the prevention of garbage pollution as amended by '
            'MEPC.360(79) and MEPC.382(80) applies as regulation." (§11)'
        ),
        "comments": "Parent incorporation instrument; links to full MARPOL Annex V regime.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§11: Scope of 'ship'—in MARPOL, 'ship' also means mobile offshore units (flyttbare innretninger), "
            "extending garbage rules to offshore petroleum and installation activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Med «skip» menes i MARPOL også flyttbare innretninger." (§11) / '
            'EN: "In MARPOL, ship also means mobile offshore units." (§11)'
        ),
        "comments": "Broadens Annex V to offshore sector beyond conventional vessels.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 2 (via §11): Application—unless expressly provided otherwise, Annex V "
            "provisions apply to all ships, including merchant vessels, platforms, pleasure craft and yachts."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Med mindre noe annet uttrykkelig er bestemt, skal bestemmelsene i dette vedlegget gjelde for '
            'alle skip." (MARPOL vedlegg V regel 2, via §11) / '
            'EN: "Unless expressly provided otherwise, the provisions of this Annex shall apply to all ships." '
            "(MARPOL Annex V regulation 2, via §11)"
        ),
        "comments": "Universal maritime scope including fisheries and leisure boats.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 3.1 (via §11): General prohibition—all garbage discharge into the sea is "
            "prohibited except as permitted in regulations 4, 5, 6 and 7; establishes default no-discharge rule."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Utslipp av alt søppel i sjøen er forbudt, med mindre noe annet følger av reglene 4, 5, 6 og 7 '
            'i dette vedlegget." (MARPOL vedlegg V regel 3 nr. 1, via §11) / '
            'EN: "Discharge of all garbage into the sea is prohibited, except as provided otherwise in '
            'regulations 4, 5, 6 and 7 of this Annex." (MARPOL Annex V regulation 3.1, via §11)'
        ),
        "comments": "Overarching discharge ban; mandatory (forbudt).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 3.2 (via §11): Plastics prohibition—discharge of all plastics at sea is "
            "prohibited, including synthetic ropes, fishing nets, plastic bags and incinerator ashes from "
            "plastic products; no distance-from-shore exceptions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Utslipp i sjøen av all plast, herunder men ikke begrenset til syntetiske taustumper, syntetiske '
            'fiskegarn, plastsøppelposer og forbrenningsrester fra plastprodukter, er forbudt." (MARPOL vedlegg V '
            'regel 3 nr. 2, via §11) / '
            'EN: "Discharge into the sea of all plastics, including but not limited to synthetic ropes, synthetic '
            'fishing nets, plastic garbage bags and incinerator ashes from plastic products is prohibited." '
            "(MARPOL Annex V regulation 3.2, via §11)"
        ),
        "comments": "Core anti-plastic-marine-litter rule; zero-discharge for all plastics.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 4.3 (via §11): Mixed garbage—when garbage is mixed with or contaminated by "
            "substances prohibited from discharge or with different discharge requirements, the more stringent "
            "requirements apply (plastic-mixed waste treated as plastic)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Når søppel er blandet med eller forurenset av andre stoffer som det er forbudt å slippe ut, '
            'eller som har ulike utslippskrav, skal de strengeste kravene gjelde." (MARPOL vedlegg V regel 4 '
            'nr. 3, via §11) / '
            'EN: "When garbage is mixed with or contaminated by other substances prohibited from discharge or '
            'having different discharge requirements, the more stringent requirements shall apply." '
            "(MARPOL Annex V regulation 4.3, via §11)"
        ),
        "comments": "Prevents circumvention by mixing plastics with dischargeable waste.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 5.1 (via §11): Offshore platforms—discharge of any garbage is prohibited "
            "from fixed or floating platforms and from all ships alongside or within 500 m of such platforms."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Utslipp i sjøen av søppel er forbudt fra faste eller flyttbare plattformer og fra alle andre '
            'skip når de ligger ved siden av eller innenfor 500 m fra slike plattformer." (MARPOL vedlegg V '
            'regel 5 nr. 1, via §11) / '
            'EN: "The discharge into the sea of any garbage is prohibited from fixed or floating platforms and '
            'from all other ships when alongside or within 500 m of such platforms." '
            "(MARPOL Annex V regulation 5.1, via §11)"
        ),
        "comments": "Offshore zero-discharge buffer zone; directly relevant to petroleum sector.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 6 (via §11): Special areas—stricter discharge rules within MARPOL special "
            "areas including the North Sea area (covering Norwegian waters); plastics remain fully prohibited."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Utslipp i sjøen av følgende søppel innenfor spesielle områder skal bare være tillatt mens '
            'skipet er underveis" (MARPOL vedlegg V regel 6 nr. 1, via §11); North Sea special area defined in '
            'regel 1 nr. 1.6 / '
            'EN: "Discharge of the following garbage into the sea within special areas shall only be permitted '
            'while the ship is en route" (MARPOL Annex V regulation 6.1, via §11); North Sea special area per '
            "regulation 1.6"
        ),
        "comments": "Enhanced controls in Norwegian sea areas; plastics still zero-discharge.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "MARPOL Annex V Reg. 8.1 (via §11): Port reception facilities—each Party must ensure adequate "
            "facilities at ports and terminals for garbage reception without undue delay, supporting shore "
            "delivery of plastics and other retained garbage."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Hver part forplikter seg til å sikre tilstrekkelige fasiliteter i havner og terminaler for '
            'mottak av søppel uten å forårsake unødig forsinkelse for skipene" (MARPOL vedlegg V regel 8 nr. 1, '
            'via §11) / '
            'EN: "Each Party undertakes to ensure the provision of adequate facilities at ports and terminals '
            'for the reception of garbage without causing undue delay to ships" (MARPOL Annex V regulation 8.1, '
            "via §11)"
        ),
        "comments": "Port infrastructure duty; complements Forurensningsforskriften Kap. 20 delivery rules.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "MARPOL Annex V Reg. 10.1 (via §11): Placards—ships ≥12 m and platforms must display placards "
            "notifying crew and passengers of Annex V discharge requirements in working language and English, "
            "French or Spanish."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Alle skip på 12 m eller mer i lengde over alt og faste eller flyttbare plattformer skal vise '
            'plakater som informerer mannskapet og passasjerene om utslippskravene i reglene 3, 4, 5 og 6 i '
            'dette vedlegget" (MARPOL vedlegg V regel 10 nr. 1, via §11) / '
            'EN: "Every ship of 12 m or more in length overall and fixed or floating platforms shall display '
            'placards which notify the crew and passengers of the discharge requirements of regulations 3, 4, 5 '
            'and 6 of this Annex" (MARPOL Annex V regulation 10.1, via §11)'
        ),
        "comments": "Onboard awareness and compliance signage; mandatory (skal vise).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "MARPOL Annex V Reg. 10.2 (via §11): Garbage management plan—ships ≥100 GT or certified for "
            "≥15 persons and platforms must carry a written plan for minimizing, collecting, storing, processing "
            "and disposing of garbage including onboard equipment use."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Alle skip med bruttotonnasje 100 eller mer, og alle skip som er sertifisert for 15 personer '
            'eller mer, og faste eller flyttbare plattformer skal ha en plan for håndtering av søppel som '
            'mannskapet skal følge" (MARPOL vedlegg V regel 10 nr. 2, via §11) / '
            'EN: "Every ship of 100 gross tonnage and above, and every ship which is certified to carry 15 or more '
            'persons, and fixed or floating platforms shall carry a garbage management plan which the crew shall '
            'follow" (MARPOL Annex V regulation 10.2, via §11)'
        ),
        "comments": "Operational waste-management procedures aboard ship.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "MARPOL Annex V Reg. 10.3 (via §11): Garbage record book—ships ≥400 GT or ≥15 persons on "
            "international voyages and platforms must maintain a Garbage Record Book logging all discharges, "
            "incinerations and deliveries to reception facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Alle skip med bruttotonnasje 400 eller mer … skal være utstyrt med en avfallsdagbok" (MARPOL '
            'vedlegg V regel 10 nr. 3, via §11); each discharge "skal snarest mulig noteres i avfallsdagboken" '
            '(regel 10 nr. 3.1) / '
            'EN: "Every ship of 400 gross tonnage and above … shall be provided with a Garbage Record Book" '
            '(MARPOL Annex V regulation 10.3, via §11); each discharge "shall be promptly recorded in the '
            'Garbage Record Book" (regulation 10.3.1)'
        ),
        "comments": "Traceability for enforcement; mandatory record-keeping (skal være utstyrt).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§11: Antarctic storage—ships in the Antarctic area south of 60°S must have sufficient onboard "
            "capacity to store all garbage produced while in the area and facilities for transfer to reception "
            "facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Skip i Antarktisområdet sør for 60° S skal ha tilstrekkelig kapasitet til å kunne oppbevare '
            'søppel som produseres om bord mens skipet er i området, og ha innretninger for overføring av søppel '
            'til mottaksanlegg." (§11) / '
            'EN: "Ships in the Antarctic area south of 60°S shall have sufficient capacity to store garbage '
            'produced on board while the ship is in the area, and have facilities for transfer of garbage to '
            'reception facilities." (§11)'
        ),
        "comments": "National supplement to MARPOL Reg. 6.3 Antarctic requirements; mandatory (skal ha).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§11: Offshore petroleum exemption—garbage record book requirement (MARPOL Reg. V/10.3) does not "
            "apply when a mobile unit is in petroleum activity on the Norwegian continental shelf and has an "
            "onboard waste-handling logging system."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Kravet om avfallsdagbok i MARPOL regel V/10.3 gjelder ikke når den flyttbare innretningen er '
            'i petroleumsvirksomhet på norsk kontinentalsokkel og innretningen har et system for å loggføre '
            'avfallshåndteringen om bord." (§11) / '
            'EN: "The requirement for a garbage record book in MARPOL regulation V/10.3 does not apply when the '
            'mobile unit is in petroleum activity on the Norwegian continental shelf and the unit has a system for '
            'logging waste management on board." (§11)'
        ),
        "comments": "Enabling exemption (gjelder ikke); alternative logging still required.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 7 (via §11): Exceptions—Regs 3–6 do not apply to discharges necessary for "
            "ship/person safety, accidental loss with reasonable precautions, or accidental fishing-gear loss; "
            "does not create general plastic discharge permission."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Reglene 3, 4, 5 og 6 i dette vedlegget skal ikke gjelde for utslipp av søppel fra et skip som '
            'er nødvendig for å sikre skipets og personers sikkerhet om bord eller for å redde liv til sjøs" '
            '(MARPOL vedlegg V regel 7 nr. 1.1, via §11) / '
            'EN: "Regulations 3, 4, 5 and 6 of this Annex shall not apply to the discharge of garbage from a ship '
            'necessary for the purpose of securing the safety of a ship and those on board or saving life at sea" '
            "(MARPOL Annex V regulation 7.1.1, via §11)"
        ),
        "comments": "Narrow safety exceptions only; instrument_in_force=0 (kan unntas).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "MARPOL Annex V Reg. 9 (via §11): Port State control—foreign ships in Norwegian ports may be "
            "inspected where crew are not familiar with garbage-prevention procedures; ship may be detained until "
            "compliance is restored."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Et skip som befinner seg i en havn eller offshoreterminal i en annen parts jurisdiksjon, er '
            'underlagt inspeksjon … når det foreligger klare grunner til å tro at skipsføreren eller mannskapet '
            'ikke er kjent med essensielle prosedyrer om bord … for å hindre forurensning fra søppel" (MARPOL '
            'vedlegg V regel 9 nr. 1, via §11) / '
            'EN: "A ship when in a port or an offshore terminal of another Party is subject to inspection … where '
            'there are clear grounds for believing that the master or crew are not familiar with essential '
            'shipboard procedures relating to the prevention of pollution by garbage." (MARPOL Annex V regulation '
            "9.1, via §11)"
        ),
        "comments": "Enforcement/detention instrument; applies to foreign ships in Norway.",
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
    ws.title = "4P Index - MARPOL V"

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
        "A": 52, "B": 42, "C": 12, "D": 45, "E": 12, "F": 42, "G": 12, "H": 40,
        "I": 16, "J": 42, "K": 16, "L": 38, "M": 14, "N": 45, "O": 12, "P": 14,
        "Q": 20, "R": 55, "S": 14, "T": 18, "U": 55, "V": 14, "W": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path = "/workspace/4P_Index_Skipssikkerhet_MARPOL_V.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
