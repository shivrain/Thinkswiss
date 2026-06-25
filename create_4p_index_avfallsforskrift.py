#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften (Waste Regulation)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om gjenvinning og behandling av avfall (Avfallsforskriften) "
        "(Regulation on recovery and treatment of waste – Waste Regulation)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930",
    "policy_year": 2025,
    "policy_objective": (
        "Central waste regulation under forurensningsloven and produktkontrolloven implementing EU waste "
        "directives—establishing extended producer responsibility (EPR) schemes, municipal collection and "
        "sorting duties, landfill and incineration rules, and treatment standards for plastic-containing waste "
        "streams including packaging, WEEE, tyres, beverage containers, SUP products, and fisheries gear."
    ),
    "policy_target": 1,
    "policy_target_text": (
        '"Fra og med 2025; 47 prosent plastemballasje" material recovery (§7-10(b)); '
        '"Fra og med 2030; 52 prosent plastemballasje" (§7-10(c)); '
        '"minst følgende andel … plastavfall … utsorteres ved kildesortering: 50 prosent fra og med 2028, '
        '60 prosent fra og med 2030 og 70 prosent fra og med 2035" (§10a-4(b))'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under forurensningsloven; principal Norwegian waste code since 2004. "
        "Major Kap. 7 packaging EPR restructure June 2025 (FOR-2025-06-18-1129); Kap. 7A SUP EPR November "
        "2025; Kap. 7B fisheries-gear EPR December 2025; Kap. 10a municipal sorting targets from 2025."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "waste management, municipalities, manufacturing, retail, packaging, beverages, electronics, "
        "automotive, agriculture, fisheries, aquaculture, all sectors generating plastic waste"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "consumption, end-of-life, recycling, litter/pollution, environmental leakage, disposal"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        '"Produsent skal dekke nødvendige kostnader til separat innsamling og etterfølgende transport og '
        'behandling av avfall fra emballasjen" (§7-8); '
        '"Avfallsgebyret fastsettes slik at det svarer til de totale kostnadene … ved lovpålagt håndtering av '
        'husholdningsavfall" (§15-3)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§1-2: WEEE chapter purpose—to prevent and reduce environmental and health problems from waste "
            "electrical and electronic equipment through separate collection, hazardous-substance sorting, and "
            "high recovery of remaining fractions including plastic housings and components."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med bestemmelsene i dette kapitlet er å forebygge og redusere miljø- og helseproblemer '
            'som EE-avfall forårsaker … gjennom separat innsamling, utsortering og behandling" (§1-2)'
        ),
        "comments": "EU WEEE Directive framework chapter; many EEE products contain plastics.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§1-10: WEEE producer EPR—producers shall finance collection, sorting and treatment of WEEE through "
            "membership in a Miljødirektoratet-approved collective or individual return company covering their "
            "imported/produced EE products."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Produsentene skal finansiere innsamling, sortering og behandling av EE-avfall gjennom medlemskap '
            'i et kollektivt eller individuelt finansiert returselskap som er godkjent av Miljødirektoratet" (§1-10)'
        ),
        "comments": "Core WEEE EPR financing; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§3-10: Battery EPR collection—producers must establish or join approved collection systems ensuring "
            "free take-back and proper treatment/recycling of waste portable, industrial and automotive batteries "
            "(many in plastic casings)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Producers shall ensure collection systems for waste batteries (Kap. 3); reporting of recovery rates "
            "per §3-13 and §3-18a"
        ),
        "comments": "EU Battery Directive implementation; mandatory producer duties.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4-4: End-of-life vehicle producer responsibility—vehicle producers must ensure environmentally sound "
            "collection and treatment of scrapped vehicles proportional to market share, with binding recovery "
            "targets (95% by 2015, ≥85% material recovery)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Produsenter plikter å sørge for miljømessig forsvarlig innsamling og behandling av kasserte kjøretøy" '
            "(§4-4); 95% recovery by 2015 including 85% material recovery (§4-4)"
        ),
        "comments": "ELV Directive; vehicles contain substantial plastic fractions.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§5-4: Tyre landfill ban—prohibition on landfilling scrap tyres; complemented by free take-back duties "
            "for dealers/producers (§5-5) and producer recycling obligation (§5-6)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å deponere dekk på fyllplass." (§5-4)'
        ),
        "comments": "Rubber/plastic composite waste stream; absolute landfill prohibition.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§6-2 + §6-4: Beverage-container deposit-return systems—purpose to achieve high return rates preventing "
            "litter; producers may establish/join Miljødirektoratet-approved return systems for inner beverage "
            "packaging (plastic bottles/cans) with minimum 25% return threshold."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Formålet … er å bidra til effektive retursystemer med høy returandel … slik at retursystemene bidrar '
            'til å hindre forsøpling" (§6-2); approved systems must achieve minimum 25% return (§6-4)'
        ),
        "comments": "DRS/pant for plastic beverage packaging; links to særavgifter Kap. 3–5.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§7-10: Packaging material-recovery targets—producers must ensure material recovery of packaging waste "
            "at minimum rates: from 2025 ≥47% plastic packaging, ≥60% carton, ≥80% paper, ≥70% ferrous metal, "
            "≥50% aluminium, ≥70% glass, ≥25% wood; from 2030 ≥52% plastic."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Produsent skal sørge for materialgjenvinning av emballasjeavfall" (§7-10); '
            '"Fra og med 2025; 47 prosent plastemballasje" (§7-10(b)); '
            '"Fra og med 2030; 52 prosent plastemballasje" (§7-10(c))'
        ),
        "comments": "Binding quantitative plastic packaging recycling targets; key plastics instrument.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§7-5a: EPR for non-packaging plastic drink cups—anyone placing single-use plastic drink cups/lids on "
            "the market must ensure separate collection, prevent energy recovery/landfill of recyclables, cover "
            "collection/treatment costs, and join approved PRO."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Enhver som ervervsmessig bringer i omsetning engangs drikkebegre … av plast … skal a. sørge for '
            'separat innsamling … c. dekke nødvendige kostnader" (§7-5a); via approved PRO (§7-5a)'
        ),
        "comments": "Extends plastic EPR beyond packaging definition; added November 2025.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§7A-1: SUP EPR chapter purpose—to prevent and reduce environmental impacts of certain single-use "
            "plastic products, reduce littering and microplastic spread, and contribute to a circular economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med dette kapittelet er å forebygge og redusere miljøpåvirkningen fra enkelte '
            'engangsprodukter av plast, herunder redusere og forebygge forsøpling og spredning av mikroplast" '
            "(§7A-1)"
        ),
        "comments": "New SUP producer-responsibility chapter; in force from 2025/2027 payments.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§7A-4: SUP litter-cost coverage—producers placing listed SUP products on the market must cover a "
            "share of municipalities' litter cleanup and collection/treatment costs; separate public-bin collection "
            "costs for food containers, flexibles, beverage containers, cups, carrier bags and tobacco filters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Produsent … skal dekke en andel av kommunenes nødvendige kostnader til opprydning av forsøpling '
            'fra slike produkter" (§7A-4); payments from 1 October 2028 for costs from 2027 (§7A-4)'
        ),
        "comments": "EU SUP Directive EPR litter financing; mandatory (skal dekke).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§7B-1: Fisheries-gear EPR purpose—to prevent and reduce environmental impacts of fishing equipment "
            "containing plastic, reduce littering and microplastic spread, and contribute to circular economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med dette kapittelet er å forebygge og redusere miljøpåvirkningen fra fiskeutstyr som '
            'inneholder plast, herunder redusere og forebygge forsøpling og spredning av mikroplast" (§7B-1)'
        ),
        "comments": "New chapter December 2025 (FOR-2025-12-18-2839); marine plastic gear EPR.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§9-4: Landfill restrictions—waste separately collected for reuse or material recovery may not be "
            "landfilled (letter h); also bans landfilling whole scrap tyres; implements EU Landfill Directive hierarchy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Avfall som er separat innsamlet for forberedelse til ombruk eller materialgjenvinning" er ikke '
            'tillatt å deponere (§9-4(h)); whole scrap tyres banned from landfill (§9-4(e))'
        ),
        "comments": "Prevents plastic recyclables reaching landfill; mandatory prohibition.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§10a-4: Municipal source-sorting targets—municipalities must ensure household plastic waste suitable "
            "for recovery is source-sorted at minimum rates: 50% from 2028, 60% from 2030, 70% from 2035; parallel "
            "targets for paper, glass/metal packaging, food waste and textiles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Kommunen skal sørge for … at plastavfall fra husholdninger … utsorteres ved kildesortering, og at '
            'minst følgende andel … 50 prosent fra og med 2028, 60 prosent fra og med 2030 og 70 prosent fra og '
            'med 2035" (§10a-4(b))'
        ),
        "comments": "EU Waste Framework Directive separate-collection implementation; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§10a-11: Agricultural plastic waste—users of agricultural plastics must source-sort and separately "
            "collect agricultural plastic waste (bale wrap, PP bags, fleece, silage film etc.) and deliver for "
            "reuse or material recovery."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Virksomheter som bruker landbruksplast skal sørge for at landbruksplastavfall … utsorteres ved '
            'kildesortering" and "sørge for separat innsamling … og at … landbruksplastavfall leveres til '
            'forberedelse til ombruk eller materialgjenvinning" (§10a-11)'
        ),
        "comments": "Farm plastic film/silage wrap stream; mandatory producer-user duties.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§15-3: Municipal household-waste fee full cost recovery—municipal councils set waste fees covering "
            "total costs of statutory household-waste management without profit, implementing polluter-pays "
            "for municipal plastic and other household waste collection."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Avfallsgebyret fastsettes slik at det svarer til de totale kostnadene kommunene påføres ved lovpålagt '
            'håndtering av husholdningsavfall. Det skal sikres full kostnadsdekning." (§15-3)'
        ),
        "comments": "Municipal financing instrument for collection infrastructure including plastics.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§18a-1–§18a-3: End-of-waste criteria—EEA-incorporated EU regulations (333/2011, 1179/2012, 715/2013) "
            "on when metal scrap and recovered glass cease to be waste, supporting secondary raw-material markets "
            "relevant to packaging and WEEE metal/plastic fractions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Forordning (EU) nr. 333/2011 on metal scrap end-of-waste criteria applies as regulation (§18a-1); "
            "similar for recovered glass (§18a-2) and copper scrap (§18a-3)"
        ),
        "comments": "EU end-of-waste framework; enables counting toward recovery targets.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§7-9(c): Packaging EPR landfill/energy ban—producers must ensure packaging waste sorted for material "
            "recovery is not energy-recovered or landfilled unless justified by environmental/resource/BAT assessment."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Produsent skal sørge for … c. At emballasjeavfall utsortert for materialgjenvinning ikke '
            'energiutnyttes eller deponeres, såfremt dette ikke ut ifra en avveiing … er berettiget." (§7-9(c))'
        ),
        "comments": "Hierarchy enforcement for plastic packaging fractions; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§18B-1: Infringement fines—Miljødirektoratet and county governors may impose infringement fees for "
            "breaches of forurensningsloven §32 on commercial waste handling and other listed provisions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Miljødirektoratet og statsforvalteren kan ved overtredelse … ilegge den ansvarlige for overtredelsen '
            'et overtredelsesgebyr" (§18B-1)'
        ),
        "comments": "Enforcement instrument (kan ilegge); caps in §18B-3.",
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
    ws.title = "4P Index - Avfallsforskr"

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
        "Q": 20, "R": 55, "S": 14, "T": 18, "U": 45, "V": 14, "W": 42,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    output_path = "/workspace/4P_Index_Avfallsforskriften.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
