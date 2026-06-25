#!/usr/bin/env python3
"""Generate 4P Index Excel for Forurensningsforskriften Kap. 20 (ship waste reception)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forurensningsforskriften – Kap. 20: Levering og mottak av avfall fra skip "
        "(Pollution Regulation – Ch. 20: Delivery and reception of waste from ships)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-931/KAPITTEL_20",
    "policy_year": 2025,
    "policy_objective": (
        "To protect the external environment by ensuring satisfactory port reception facilities for ship waste "
        "and mandatory delivery of waste—including garbage and plastics under MARPOL Annex V—to port "
        "reception arrangements, with sorting for reuse and material recovery, advance notification, waste "
        "receipts, indirect fee financing, and monitoring of passively fished marine litter."
    ),
    "policy_target": 0.5,
    "policy_target_text": (
        'NO: "Mottaksordningene skal legge til rette for at avfall fra skip blir samlet inn og utsortert på en '
        'måte som tilrettelegger for ombruk og materialgjenvinning." (§20-5) / '
        'EN: "The reception arrangements shall facilitate that waste from ships is collected and sorted in a '
        'way that enables reuse and material recovery." (§20-5)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under forurensningsloven and skipssikkerhetsloven; implements EU Port "
        "Reception Facilities Directive 2019/883 (MARPOL Annex V garbage/plastics). Chapter substantially "
        "revised October 2023 (FOR-2023-06-02-1213, in force 1 October 2023)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "maritime, ports, fisheries, shipping, waste management, coastal municipalities, aquaculture, "
        "recreation (leisure craft)"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "end-of-life, litter/pollution, environmental leakage, recycling"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        'NO: "Omkostningene forbundet med mottak og videre håndtering av avfall fra skip … skal dekkes ved '
        'innkreving av generelt avfallsgebyr fra de skip som anløper havnen." (§20-9) / '
        'EN: "The costs associated with receiving and further handling of waste from ships … shall be covered '
        'by levying a general waste fee on the ships calling at the port." (§20-9)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-1: Purpose—to protect the external environment by ensuring establishment and operation of "
            "satisfactory reception arrangements for ship waste and that ship waste is delivered to port "
            "reception facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Dette kapitlets formål er å verne det ytre miljø ved å sikre etablering og drift av '
            'tilfredsstillende mottaksordninger for avfall fra skip, og å sørge for at avfall fra skip blir levert '
            'til mottaksordning i havn." (§20-1) / '
            'EN: "The purpose of this chapter is to protect the external environment by ensuring the '
            'establishment and operation of satisfactory reception arrangements for waste from ships, and to '
            'ensure that waste from ships is delivered to a reception arrangement in port." (§20-1)'
        ),
        "comments": "Frames EU PRF Directive / MARPOL ship-waste regime.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-2: Scope—applies to all Norwegian and foreign ships (including fishing vessels and leisure craft) "
            "calling at Norwegian ports, and to all Norwegian ports normally visited; notification/delivery rules "
            "also apply to Norwegian ships in EEA ports; excludes Svalbard."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Dette kapitlet gjelder a. alle norske og utenlandske skip … som anløper norsk havn" and '
            '"b. alle norske havner som normalt anløpes av skip som omfattes av bokstav a." (§20-2) / '
            'EN: "This chapter applies to a. all Norwegian and foreign ships … calling at a Norwegian port" and '
            '"b. all Norwegian ports normally visited by ships covered by letter a." (§20-2)'
        ),
        "comments": "Broad maritime scope including fisheries and leisure boats.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-3: Definitions—ship waste includes all waste under MARPOL Annexes I, II, IV, V (garbage/plastics) "
            "and VI, plus passively fished waste; defines port, port authority, reception arrangement, and "
            "fishing vessel."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "avfall fra skip: enhver form for avfall … som omfattes av vedleggene I (olje), II (skadelige '
            'flytende stoffer i bulk), IV (kloakk), V (søppel) og VI (utslipp til luft) til … MARPOL 73/78 … '
            'samt oppfisket avfall" (§20-3 nr. 1) / '
            'EN: "waste from ships: any form of waste … covered by Annexes I (oil), II (noxious liquid substances '
            'in bulk), IV (sewage), V (garbage) and VI (air emissions) to … MARPOL 73/78 … as well as passively '
            'fished waste" (§20-3 no. 1)'
        ),
        "comments": "MARPOL Annex V explicitly covers ship garbage including plastics.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-5: Port reception facilities—port authorities shall establish and operate reception arrangements "
            "facilitating collection and sorting for reuse and material recovery; arrangements must cover normal "
            "delivery needs without undue ship delay."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Havneansvarlig skal ut fra behovet for levering sørge for etablering og drift av mottaksordninger '
            'for avfall fra skip i havnen. Mottaksordningene skal legge til rette for at avfall fra skip blir samlet '
            'inn og utsortert på en måte som tilrettelegger for ombruk og materialgjenvinning." (§20-5) / '
            'EN: "The port authority shall, based on delivery needs, ensure the establishment and operation of '
            'reception arrangements for waste from ships in the port. The reception arrangements shall facilitate '
            'that waste from ships is collected and sorted in a way that enables reuse and material recovery." '
            "(§20-5)"
        ),
        "comments": "Core infrastructure duty; mandatory (skal). Includes MARPOL Annex V garbage.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-6: Port waste plans—port authorities must prepare and implement waste plans (Annex I) approved by "
            "county governor for five years, published in Norwegian and English via SafeSeaNet Norway; "
            "municipalities maintain harbour overview."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Havneansvarlig plikter å utarbeide og gjennomføre en avfallsplan i samråd med berørte parter" '
            'and "Avfallsplanen skal godkjennes av statsforvalteren" (§20-6); plan info shall be "lett '
            'tilgjengelig på norsk og engelsk" (§20-6) / '
            'EN: "The port authority is obliged to prepare and implement a waste plan in consultation with '
            'affected parties" and "The waste plan shall be approved by the county governor" (§20-6); plan '
            'information shall be "easily available in Norwegian and English" (§20-6)'
        ),
        "comments": "Planning and transparency; exemption for small leisure-craft harbours.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-7: Advance waste notification—vessels ≥300 GT (with exceptions) bound for EEA ports must notify "
            "waste delivery via SafeSeaNet Norway at least 24 hours before arrival (or as soon as port known)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Fartøy med bruttotonnasje 300 eller mer … skal gi melding om levering av avfall i havnen a. '
            'minst 24 timer før anløp, dersom anløpshavnen er kjent" (§20-7); notification via SafeSeaNet '
            'Norway (§20-7) / '
            'EN: "Vessels with gross tonnage of 300 or more … shall give notification of waste delivery in the '
            'port a. at least 24 hours before arrival, if the port of call is known" (§20-7); notification via '
            "SafeSeaNet Norway (§20-7)"
        ),
        "comments": "Enables port preparedness for garbage including plastics; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-8: Mandatory waste delivery—ship masters shall ensure waste is delivered to port reception before "
            "departure in accordance with MARPOL discharge norms; sufficient on-board storage may defer delivery "
            "only where next port has adequate facilities; Sjøfartsdirektoratet may require delivery."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Skipsfører skal sørge for at avfall leveres mottaksordning i havn før avgang, i samsvar med de '
            'relevante utslippsnormene fastsatt i MARPOL-konvensjonen." (§20-8) / '
            'EN: "The ship master shall ensure that waste is delivered to a reception arrangement in port before '
            'departure, in accordance with the relevant discharge standards set out in the MARPOL Convention." '
            "(§20-8)"
        ),
        "comments": "Direct anti-dumping rule for ship garbage; mandatory (skal).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-8a: Waste receipts—port authorities must issue standardised waste receipts (Annex IV) to masters "
            "without undue delay; vessels ≥300 GT report receipt data in SafeSeaNet Norway and retain records on "
            "board for two years."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Når et skip har levert avfall i mottaksordning i havn, skal havneansvarlig … sørge for at en '
            'avfallskvittering … blir fylt ut. Avfallskvitteringen skal utstedes til skipsfører uten unødig '
            'forsinkelse." (§20-8a) / '
            'EN: "When a ship has delivered waste to a reception arrangement in port, the port authority … shall '
            'ensure that a waste receipt … is completed. The waste receipt shall be issued to the ship master '
            'without undue delay." (§20-8a)'
        ),
        "comments": "Traceability and enforcement documentation; added October 2023.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-8b: Exemption for regular-line ships—scheduled vessels with regular port calls may be exempt from "
            "§20-7 notification and §20-8 delivery if documented waste-fee arrangement exists on route, subject "
            "to Sjøfartsdirektoratet oversight."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Rutegående skip med regelmessige havneanløp kan fritas fra meldeplikten etter § 20-7 og plikten '
            'til å levere avfall etter § 20-8 dersom det foreligger en ordning for å levere avfall og betale '
            'avfallsgebyrer i en havn på skipets rute" (§20-8b) / '
            'EN: "Scheduled ships with regular port calls may be exempted from the notification duty under §20-7 '
            'and the duty to deliver waste under §20-8 if there is an arrangement for delivering waste and paying '
            'waste fees at a port on the ship\'s route" (§20-8b)'
        ),
        "comments": "Conditional exemption (kan fritas); must not compromise storage capacity.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-9: Port waste fees—costs of receiving and handling ship waste (except cargo residues and scrubber "
            "waste) funded by general port waste fee levied on all calling ships regardless of delivery; "
            "differentiated by ship type/size and hazardous waste; MARPOL Annex V surcharge for excess garbage."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Omkostningene forbundet med mottak og videre håndtering av avfall fra skip … skal dekkes ved '
            'innkreving av generelt avfallsgebyr fra de skip som anløper havnen. Gebyr som nevnt skal innkreves '
            'uavhengig av om det leveres avfall fra skipet til mottaksordningen." (§20-9); '
            '"Tilleggsgebyr for avfall omfattet av MARPOL vedlegg V (søppel) kan dessuten innkreves" (§20-9) / '
            'EN: "The costs associated with receiving and further handling of waste from ships … shall be covered '
            'by levying a general waste fee on the ships calling at the port. The fee mentioned shall be levied '
            'regardless of whether waste from the ship is delivered to the reception arrangement." (§20-9); '
            '"A surcharge for waste covered by MARPOL Annex V (garbage) may furthermore be levied" (§20-9)'
        ),
        "comments": "Polluter-pays / no-special-fee incentive for garbage delivery.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-10: Fee calculation—fees must cover indirect administration and ≥30% of direct operating costs; "
            "MARPOL Annex V garbage fees based on persons on board and sailing days since last delivery port."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Gebyret for de direkte driftskostnadene skal dekke minst 30 % av de samlede direkte kostnadene '
            'til levering av avfall foregående år" (§20-10); '
            '"Gebyr for avfall omfattet av MARPOL vedlegg V (søppel), inkludert oppfisket avfall, skal beregnes på '
            'grunnlag av antall personer som skipet har tillatelse til å transportere eller antall besetningsmedlemmer, '
            'samt antall seilingsdøgn siden forrige leveringshavn." (§20-10) / '
            'EN: "The fee for direct operating costs shall cover at least 30% of the total direct costs of waste '
            'delivery in the preceding year" (§20-10); '
            '"The fee for waste covered by MARPOL Annex V (garbage), including passively fished waste, shall be '
            'calculated on the basis of the number of persons the ship is permitted to carry or the number of crew '
            'members, as well as the number of sailing days since the previous delivery port." (§20-10)'
        ),
        "comments": "Methodological fee rules for garbage/plastics stream.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-11: Fee discounts—reductions where ship design/operation generates less waste and handles waste "
            "sustainably per EU Regulation 2022/91; also discounts for intra-European traffic."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Det skal gis fradrag i gebyret dersom skipets konstruksjon, utstyr eller drift bidrar til at '
            'a. skipet genererer reduserte mengder avfall, og b. skipet håndterer avfallet på en bærekraftig og '
            'miljøvennlig måte." (§20-11) / '
            'EN: "A reduction in the fee shall be granted if the ship\'s construction, equipment or operation '
            'contributes to a. the ship generating reduced quantities of waste, and b. the ship handling waste in '
            'a sustainable and environmentally friendly manner." (§20-11)'
        ),
        "comments": "Incentive for waste reduction aboard ship.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-12: Supervision—Sjøfartsdirektoratet supervises ship compliance and may issue orders; county "
            "governor supervises ports; EU Regulation 2022/90 on risk-based inspection selection applies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Sjøfartsdirektoratet skal føre tilsyn med at skipenes plikter etter dette kapitlet overholdes, '
            'og kan gi pålegg til rederiet. Statsforvalteren er forurensningsmyndighet for havner og fører tilsyn '
            'med at bestemmelsene i dette kapitlet blir overholdt." (§20-12) / '
            'EN: "The Norwegian Maritime Authority shall supervise compliance with ships\' duties under this '
            'chapter, and may issue orders to the shipowner. The county governor is the pollution authority for '
            'ports and supervises compliance with the provisions of this chapter." (§20-12)'
        ),
        "comments": "Dual enforcement: ships (NMA) and ports (county governor).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-14: Passive-fished-waste monitoring—EU Regulation 2022/92 (as amended by 2024/917) on collection "
            "and reporting of passively fished waste applies as regulation, implementing PRF Directive monitoring "
            "for marine litter including plastics retrieved by fishing gear."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forordning (EU) 2022/92 om fastsettelse av regler for anvendelsen av direktiv (EU) 2019/883 med '
            'hensyn til metoder for innsamling av overvåkingsdata og formatet for å rapportere passivt oppfisket '
            'avfall … gjelder som forskrift." (§20-14) / '
            'EN: "Regulation (EU) 2022/92 laying down rules for the application of Directive (EU) 2019/883 as '
            'regards methods for collecting monitoring data and the format for reporting passively fished waste … '
            'applies as regulation." (§20-14)'
        ),
        "comments": "Marine litter data-collection; EEA-incorporated EU regulation.",
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
    ws.title = "4P Index - Forurens Kap20"

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

    output_path = "/workspace/4P_Index_Forurensningsforskriften_Kap20.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
