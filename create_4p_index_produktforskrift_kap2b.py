#!/usr/bin/env python3
"""Generate 4P Index Excel for Produktforskriften Kap. 2b (certain plastic products / SUP)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Produktforskriften – Kap. 2b: Regulering av enkelte produkter av plast "
        "(Product Regulation – Ch. 2b: Regulation of certain plastic products)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-922/KAPITTEL_2b",
    "policy_year": 2026,
    "policy_objective": (
        "Chapter 2b of the Norwegian Product Regulation (FOR-2004-06-01-922) under produktkontrolloven, "
        "implementing EU Single-Use Plastics Directive (2019/904) and related EEA acts. Governs SUP product "
        "bans, oxo-degradable plastic ban, harmonised marking (EU 2020/2151), tethered caps/lids (EN 17665), "
        "recycled-content requirements for PET beverage bottles (25% from 2025, 30% from 2030), and "
        "infringement fees. Chapter added July 2021 (FOR-2020-12-18-3200); marking from October 2021 "
        "(FOR-2021-06-28-2278); tethered caps, recycled content and §2b-8 from December 2024 (FOR-2024-12-20-3499)."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Det er forbudt å bringe i omsetning følgende produkter … laget helt eller delvis av plast og er '
        'beregnet for engangsbruk" (§2b-3); "Fra 1. januar 2025 … minst 25 % materialgjenvunnet plast" (§2b-7); '
        '"Fra 1. januar 2030 … minst 30 % materialgjenvunnet plast" (§2b-7) / '
        'EN: "It is prohibited to place on the market … made wholly or partly of plastic and intended for single use" '
        '(§2b-3); "From 1 January 2025 … at least 25% material-recycled plastic" (§2b-7); '
        '"From 1 January 2030 … at least 30% material recycled plastic" (§2b-7)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under produktkontrolloven §4, as consolidated Chapter 2b of "
        "Produktforskriften (FOR-2004-06-01-922). Norwegian EEA implementation of EU Directive 2019/904 "
        "(SUP), EU Regulation 2020/2151 (marking), and subsequent SUP design/content requirements. "
        "Chapter inserted 3 July 2021; latest substantive amendments December 2024; EEA references updated "
        "March 2026 (FOR-2026-03-23-509)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "retail, food service, cosmetics, manufacturing, hospitality, tobacco, personal care, "
        "beverage production, import/export, packaging, catering, events"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, design, use/consumption, litter/pollution, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Miljødirektoratet kan ved overtredelse av § 2b-3, § 2b-4, § 2b-5 første ledd, § 2b-6 eller § 2b-7 '
        'ilegge den ansvarlige for overtredelsen et overtredelsesgebyr" (§2b-8); kap. 6b: inntil 15 ganger '
        'folketrygdens grunnbeløp for foretak (§6b-2) / '
        'EN: "The Norwegian Environment Agency may … impose … an infringement fee" (§2b-8); '
        'Chapter 6b: up to 15× folketrygd basic amount for enterprises (§6b-2)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§2b-1: Chapter purpose—prevent and reduce environmental impacts of certain plastic products, "
            "including reducing and preventing marine littering and spread of microplastics."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Formålet med dette kapitlet er å forebygge og redusere miljøpåvirkningen fra enkelte '
            'plastprodukter, herunder redusere og forebygge marin forsøpling og spredning av mikroplast." '
            '(§2b-1) / '
            'EN: "The purpose of this chapter is to prevent and reduce the environmental impact of certain plastic '
            'products, including reducing and preventing marine littering and spreading microplastics." (§2b-1)'
        ),
        "comments": "Frames EU SUP Directive (2019/904) national implementation objective; added FOR-2020-12-18-3200.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-2: Definitions—'bringing into turnover' (first placement on Norwegian market), 'plastic' "
            "(polymer per REACH Art. 3(5), excluding unmodified natural polymers), and 'oxo-degradable plastic' "
            "(oxidation-induced fragmentation to microfragments)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Med plast menes et materiale som består av en polymer som definert i artikkel 3 nr. 5 i '
            'REACH-forordningen"; "Med okso-nedbrytbar plast menes plastmaterialer som inneholder tilsetningsstoffer '
            'som, ved oksidasjon, fører til fragmentering … til mikrofragmenter" (§2b-2) / '
            'EN: "Plastic means a material consisting of a polymer as defined in Article 3(5) of the REACH '
            'Regulation"; "Oxo-degradable plastics means plastic materials containing additives which, by oxidation, '
            'lead to fragmentation … to microfragments" (§2b-2)'
        ),
        "comments": "Legal scope definitions linking to REACH-forskriften; amended FOR-2021-06-28-2278.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-3: SUP market ban (items a–f)—prohibition on placing on market single-use plastic cotton buds, "
            "cutlery, plates, straws, beverage stirrers, and balloon sticks/attachments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning følgende produkter dersom de er laget helt eller delvis av '
            'plast og er beregnet for engangsbruk" including "bomullspinner", "bestikk", "tallerkener", "sugerør", '
            '"rørepinner til drikkevarer" and "ballongpinner" (§2b-3 a–f) / '
            'EN: "It is forbidden to bring into turnover … made in whole or in part of plastic and intended for '
            'one-time use" including cotton sticks, cutlery, plates, straws, stirring sticks and balloon sticks '
            "(§2b-3 a–f)"
        ),
        "comments": "Core EU SUP Directive Art. 5 product ban list; mandatory (forbudt); in force 3 July 2021.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-3: SUP market ban (items g–i)—prohibition on single-use EPS food containers, EPS beverage "
            "packaging with caps/lids, and EPS beverage cups with lids."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning … g. matbeholdere av ekspandert polystyren (EPS) … h. '
            'drikkevareemballasje av ekspandert polystyren (EPS) og korker og lokk til disse i. drikkebegre av '
            'ekspandert polystyren (EPS) og lokk til disse" (§2b-3) / '
            'EN: "It is forbidden to bring into turnover … g. food containers of expanded polystyrene (EPS) … h. '
            'drinking goods packaging of expanded polystyrene (EPS) and corks and lids for these i. drinking cups '
            'of expanded polystyrene (EPS) and lids for these" (§2b-3)'
        ),
        "comments": "Take-away/food-service EPS prohibition; high beach-litter items.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-3(2–3): Medical and industrial exemptions—cotton buds (a) and straws (d) exempt when used as "
            "medical equipment; balloon sticks (f) exempt for industrial/professional use not distributed to consumers."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forbudene i første ledd bokstav a og d gjelder ikke for bomullspinner og sugerør til bruk som '
            'medisinsk utstyr"; "Forbudet i første ledd bokstav f gjelder ikke for ballongpinner … til industriell '
            'eller annen ervervsmessig bruk, som ikke distribueres til forbrukere" (§2b-3) / '
            'EN: "The prohibitions … letters a and d do not apply to cotton swabs and straws for use as medical '
            'equipment"; "The prohibition … letter f does not apply to balloon sticks … for industrial or other '
            'acquisitional use, which are not distributed to consumers" (§2b-3)'
        ),
        "comments": "Conditional carve-outs (gjelder ikke); healthcare and B2B sectors.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-4: Oxo-degradable plastic ban—prohibition on placing on the market any product made of "
            "oxo-degradable plastic."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning produkter laget av okso-nedbrytbar plast." (§2b-4) / '
            'EN: "It is forbidden to bring into turnover products made of oxo-degradable plastic." (§2b-4)'
        ),
        "comments": "EU SUP oxo-plastic ban; prevents intentional microplastic generation; in force 3 July 2021.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-5(1): SUP marking obligation—operators placing sanitary items, wet wipes, tobacco filters and "
            "drinking cups on the market must ensure harmonised marking when products are wholly/partly plastic "
            "and single-use."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Den som bringer i omsetning følgende produkter, skal sørge for at produktene er merket i '
            'henhold til merkekravene i andre ledd … dersom de er laget helt eller delvis av plast og er beregnet '
            'for engangsbruk" including "bind, tamponger", "våtservietter", "tobakksprodukt med filter" and '
            '"drikkebegre" (§2b-5) / '
            'EN: "The person who handles the following products shall ensure that the products are marked … if they '
            'are made in whole or in part of plastic and are intended for single use" including sanitary items, wet '
            'wipes, tobacco filters and drinking cups (§2b-5)'
        ),
        "comments": "Mandatory marking (skal sørge for); behaviour-change/litter-awareness; added FOR-2021-06-28-2278.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-5(2): Direct incorporation of EU harmonised marking—EU Regulation 2020/2151 applies as Norwegian "
            "regulation with EEA adaptations (Annex II Ch. XVII No. 9dd)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Forordning (EU) 2020/2151 om harmoniserte merkekrav for enkeltstående plastprodukter … gjelder '
            'som forskrift" med EØS-tilpasninger (§2b-5) / '
            'EN: "Regulation (EU) 2020/2151 on harmonised labelling requirements for single-use products in '
            'plastics … applies as regulations" with EEA adaptations (§2b-5)'
        ),
        "comments": "EEA dynamic incorporation; pictograms warn against littering/toilet disposal.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§2b-6(1): Tethered caps and lids—single-use plastic beverage containers ≤3 L may only be placed on "
            "market if plastic cap/lid remains attached throughout intended use, meeting EN 17665:2022+A1:2023."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "En drikkevarebeholder til engangsbruk … kan bare bringes i omsetning hvis korken eller lokket '
            'sitter fast i beholderen i hele perioden produktet er ment å brukes"; standard EN 17665:2022+A1:2023 '
            '(§2b-6) / '
            'EN: "A single-use beverage container … can only be brought into turnover if the cap or lid is stuck '
            'in the container throughout the period the product is intended to be used"; standard EN 17665:2022+A1:2023 '
            "(§2b-6)"
        ),
        "comments": "EU SUP design requirement for cap collection; added FOR-2024-12-20-3499.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§2b-6(2–3): Tethered-cap exemptions and scope—glass/metal containers with plastic caps exempt; "
            "special medical-purpose beverages exempt; metal caps with plastic seal not considered plastic."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Første ledd gjelder ikke for … drikkevarebeholdere av glass eller metall med korker eller lokk av '
            'plast"; "Korker og lokk av metall med plastforsegling skal ikke anses som laget av plast" (§2b-6) / '
            'EN: "The first paragraph does not apply to … beverage containers of glass or metal with corks or lids '
            'of plastic"; "The corks and lids of metal with plastic seal should not be considered to be made of '
            'plastic" (§2b-6)'
        ),
        "comments": "Scope delimitation (gjelder ikke); reduces burden on non-plastic primary containers.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§2b-6(4): Written compliance documentation—operator placing tethered-cap beverage containers on "
            "market must document in writing that EN 17665 standard is met."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Den som bringer i omsetning drikkevarebeholdere omfattet av denne bestemmelsen, skal kunne '
            'dokumentere skriftlig at den europeiske standarden … er oppfylt" (§2b-6) / '
            'EN: "The person who brings in the turnover beverage containers covered by this provision shall be '
            'able to document in writing that the European standard referred to in the first paragraph is met" '
            "(§2b-6)"
        ),
        "comments": "Documentation duty (skal kunne dokumentere); supports market surveillance.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-7(1): Recycled PET content 2025—from 1 January 2025, single-use PET beverage bottles ≤3 L must "
            "contain ≥25% material-recycled plastic (de minimis: <1,000 kg placed on market per calendar year)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Fra 1. januar 2025 … skal … flaskene inneholde minst 25 % materialgjenvunnet plast"; unntak for '
            'under 1 000 kg per kalenderår (§2b-7) / '
            'EN: "From 1 January 2025 … ensuring that the bottles contain at least 25% material-recycled plastic"; '
            'exemption for under 1,000 kg per calendar year (§2b-7)'
        ),
        "comments": "Binding recycled-content target; EU SUP Art. 6; small-producer threshold.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-7(2): Recycled PET content 2030—from 1 January 2030, single-use PET beverage bottles ≤3 L must "
            "contain ≥30% material-recycled plastic."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Fra 1. januar 2030 … skal … flaskene inneholde minst 30 % materialgjenvunnet plast" (§2b-7) / '
            'EN: "From 1 January 2030 … the bottles contain at least 30% material recycled plastic" (§2b-7)'
        ),
        "comments": "Escalating circularity target; same PET bottle scope as 2025 requirement.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-7(3–5): Recycled-content compliance mechanics—cap/lid counted as part of bottle; target may be met "
            "per bottle or as portfolio average for calendar year; glass/metal and special medical bottles exempt."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Korken eller lokket på flasken regnes som en del av flasken"; kravene kan oppfylles "enten ved at '
            'hver enkelt drikkeflaske, eller det totale antallet drikkeflasker … inneholder minst den angitte andelen '
            'materialgjenvunnet plast" (§2b-7) / '
            'EN: "The cork or lid of the bottle is considered part of the bottle"; requirements may be met "either by '
            'the fact that each individual drinking bottle, or the total amount of drinking bottles … contains at least '
            'the indicated proportion of material recovered plastics" (§2b-7)'
        ),
        "comments": "Portfolio averaging allows producer flexibility; exemptions mirror §2b-6.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§2b-7(6): Annual recycled-content reporting—operators must prepare annual report on compliance via "
            "approved beverage DRS (Avfallsforskriften Kap. 6) or packaging PRO (Kap. 7)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Den som skal oppfylle kravene … skal utarbeide en årlig rapport om oppfyllelse av kravene i sitt '
            'godkjente retursystem for drikkevareemballasje, jf. avfallsforskriften kapittel 6, eller returselskap '
            'for emballasje, jf. kapittel 7" (§2b-7) / '
            'EN: "The person who shall meet requirements … must prepare an annual report on the fulfilment of the '
            'requirements … approved return system for beverage packaging, cf. Chapter 6 of the Waste Regulations, '
            'or the return company for packaging, cf. Chapter 7" (§2b-7)'
        ),
        "comments": "Cross-links to Infinitum DRS and packaging PRO reporting chains.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§2b-8: Infringement fees—Miljødirektoratet may impose overtredelsesgebyr for breaches of §§2b-3–2b-7; "
            "governed by Chapter 6b (up to 15× folketrygd grunnbeløp for enterprises)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet kan ved overtredelse av § 2b-3, § 2b-4, § 2b-5 første ledd, § 2b-6 eller § 2b-7 '
            'ilegge den ansvarlige for overtredelsen et overtredelsesgebyr"; "For ileggelse av overtredelsesgebyr '
            'gjelder bestemmelsene i kapittel 6b" (§2b-8) / '
            'EN: "The Norwegian Environment Agency may, upon infringement of section 2b-3, § 2b-4, section 2b-5 first '
            'paragraph, section 2b-6 or section 2b-7, impose … an infringement fee"; "Chapter 6b applies" (§2b-8)'
        ),
        "comments": "Enforcement instrument (kan ilegge); added FOR-2024-12-20-3499.",
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
    ws.title = "4P Index - Kap 2b"

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

    output_path = "/workspace/4P_Index_Produktforskrift_Kap2b.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
