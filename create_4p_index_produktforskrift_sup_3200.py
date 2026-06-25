#!/usr/bin/env python3
"""Generate 4P Index Excel for FOR-2020-12-18-3200 (Produktforskriften Kap. 2b SUP bans)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om endring i produktforskriften (forbud mot enkelte plastprodukter) – "
        "18.12.2020 nr. 3200 (Amendment to the Product Regulation (ban on certain plastic products) – "
        "18 December 2020 No. 3200)"
    ),
    "policy_url": "https://lovdata.no/dokument/LTI/forskrift/2020-12-18-3200",
    "policy_year": 2021,
    "policy_objective": (
        "Amending regulation under produktkontrolloven inserting new Chapter 2b into "
        "Produktforskriften (FOR-2004-06-01-922), implementing Norway's initial EU Single-Use Plastics "
        "Directive (2019/904) measures: market bans on listed single-use plastic products and oxo-degradable "
        "plastic, with purpose, definitions and entry into force 3 July 2021."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Det er forbudt å bringe i omsetning følgende produkter dersom de er laget helt eller delvis av plast '
        'og er beregnet for engangsbruk" (§2b-3); '
        '"Det er forbudt å bringe i omsetning produkter laget av okso-nedbrytbar plast." (§2b-4) / '
        'EN: "It is prohibited to place on the market the following products if they are made wholly or partly of '
        'plastic and are intended for single use" (§2b-3); '
        '"It is prohibited to place on the market products made of oxo-degradable plastic." (§2b-4)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive amending regulation (forskrift) under produktkontrolloven §4; inserts Kap. 2b into "
        "Produktforskriften as Norwegian implementation of EU SUP Directive Article 5 product bans. "
        "Adopted 18 December 2020 (FOR-2020-12-18-3200), in force 3 July 2021; parent chapter now consolidated "
        "in Produktforskriften kap. 2b."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "retail, food service, manufacturing, hospitality, events, healthcare (exemptions), import/export, "
        "packaging, catering"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, use/consumption, litter/pollution, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Forskriften trer i kraft 3. juli 2021." (II); enforcement via produktkontrolloven and subsequent '
        'Produktforskriften §2b-8 (overtredelsesgebyr, added 2024) / '
        'EN: "The regulation enters into force on 3 July 2021." (Part II); enforcement via Product Control Act '
        'and subsequent Produktforskriften §2b-8 (infringement fees, added 2024)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Amending regulation Part I: Inserts new Chapter 2b into Produktforskriften (FOR-2004-06-01-922) "
            "regulating certain plastic products; repeals former §1-1 of produktforskriften."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "I forskrift 1. juni 2004 nr. 922 … (produktforskriften) gjøres følgende endringer: § 1-1 '
            'oppheves. Nytt kapittel 2b skal lyde: Kapittel 2b. Regulering av enkelte produkter av plast." '
            '(FOR-2020-12-18-3200 I) / '
            'EN: "In Regulation 1 June 2004 No. 922 … (the Product Regulation) the following amendments are made: '
            '§1-1 is repealed. New Chapter 2b shall read: Chapter 2b. Regulation of certain plastic products." '
            "(FOR-2020-12-18-3200 Part I)"
        ),
        "comments": "Structural amendment; Kap. 2b is the operative parent chapter in consolidated produktforskriften.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Part II: Entry into force—amendment and new Chapter 2b provisions apply from 3 July 2021."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Forskriften trer i kraft 3. juli 2021." (FOR-2020-12-18-3200 II) / '
            'EN: "The regulation enters into force on 3 July 2021." (FOR-2020-12-18-3200 Part II)'
        ),
        "comments": "Transitional deadline for SUP ban compliance.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§2b-1: Purpose—prevent and reduce environmental impacts of certain plastic products, including "
            "reducing and preventing marine littering and spread of microplastics."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Formålet med dette kapitlet er å forebygge og redusere miljøpåvirkningen fra enkelte '
            'plastprodukter, herunder redusere og forebygge marin forsøpling og spredning av mikroplast." '
            '(§2b-1) / '
            'EN: "The purpose of this chapter is to prevent and reduce the environmental impacts of certain plastic '
            'products, including reducing and preventing marine littering and spread of microplastics." (§2b-1)'
        ),
        "comments": "Frames EU SUP Directive (2019/904) national implementation objective.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-2: Definitions—'placing on market' (bring i omsetning), 'plastic' (polymer per REACH Art. 3(5), "
            "excluding unmodified natural polymers), and 'oxo-degradable plastic' (oxidation-induced fragmentation)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Med plast menes et materiale som består av en polymer som definert i artikkel 3 nr. 5) i '
            'REACH-forordningen" (§2b-2); "Med okso-nedbrytbar plast menes plastmaterialer som inneholder '
            'tilsetningsstoffer som, ved oksidasjon, fører til fragmentering … til mikrofragmenter" (§2b-2) / '
            'EN: "Plastic means a material consisting of a polymer as defined in Article 3(5) of the REACH '
            'Regulation" (§2b-2); "Oxo-degradable plastic means plastic materials containing additives which, '
            'through oxidation, lead to fragmentation … into microfragments" (§2b-2)'
        ),
        "comments": "Legal scope definitions; links to REACH-forskriften polymer definition.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3: SUP market ban—prohibition on placing on market single-use plastic cotton buds, cutlery, "
            "plates, straws, beverage stirrers, balloon sticks, EPS food containers, EPS beverage packaging "
            "and EPS cups with lids."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning følgende produkter dersom de er laget helt eller delvis av '
            'plast og er beregnet for engangsbruk" including "bomullspinner", "bestikk", "sugerør", '
            '"matbeholdere av ekspandert polystyren (EPS)" and "drikkevareemballasje av ekspandert polystyren '
            '(EPS)" (§2b-3) / '
            'EN: "It is prohibited to place on the market the following products if they are made wholly or partly '
            'of plastic and are intended for single use" including "cotton bud sticks", "cutlery", "straws", '
            '"food containers made of expanded polystyrene (EPS)" and "beverage packaging made of expanded '
            'polystyrene (EPS)" (§2b-3)'
        ),
        "comments": "Core EU SUP Directive Article 5 product ban list; mandatory (forbudt).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3: Medical exemptions—ban on cotton buds (a) and straws (d) does not apply to products "
            "for use as medical devices."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forbudene i første ledd bokstav a og d gjelder ikke for bomullspinner og sugerør til bruk som '
            'medisinsk utstyr." (§2b-3) / '
            'EN: "The prohibitions in the first paragraph letters a and d do not apply to cotton bud sticks and '
            'straws for use as medical devices." (§2b-3)'
        ),
        "comments": "Conditional exemption (gjelder ikke); healthcare sector carve-out.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3: Industrial balloon-stick exemption—ban on balloon sticks and attachments (f) does not apply "
            "to industrial or other professional use not distributed to consumers."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forbudet i første ledd bokstav f gjelder ikke for ballongpinner og festemekanismer til ballonger '
            'til industriell eller annen ervervsmessig bruk, som ikke distribueres til forbrukere." (§2b-3) / '
            'EN: "The prohibition in the first paragraph letter f does not apply to balloon sticks and attachments '
            'for balloons for industrial or other professional use that are not distributed to consumers." (§2b-3)'
        ),
        "comments": "B2B exemption (gjelder ikke); events/retail still covered.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-4: Oxo-degradable plastic ban—prohibition on placing on the market any product made of "
            "oxo-degradable plastic that fragments into microplastics through oxidation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning produkter laget av okso-nedbrytbar plast." (§2b-4) / '
            'EN: "It is prohibited to place on the market products made of oxo-degradable plastic." (§2b-4)'
        ),
        "comments": "EU SUP Directive oxo-plastic ban; prevents intentional microplastic generation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3(b–c): Cutlery and plates ban—single-use plastic forks, knives, spoons, chopsticks and plates "
            "prohibited from EEA market placement."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning … b. bestikk (gafler, kniver, skjeer og spisepinner) c. '
            'tallerkener" dersom de er laget helt eller delvis av plast og er beregnet for engangsbruk (§2b-3) / '
            'EN: "It is prohibited to place on the market … b. cutlery (forks, knives, spoons and chopsticks) c. '
            'plates" if made wholly or partly of plastic and intended for single use (§2b-3)'
        ),
        "comments": "Food-service sector ban; mandatory (forbudt).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3(g–i): EPS ban—single-use expanded polystyrene food containers, beverage packaging with caps/lids, "
            "and cups with lids prohibited from market."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning … g. matbeholdere av ekspandert polystyren (EPS) … h. '
            'drikkevareemballasje av ekspandert polystyren (EPS) og korker og lokk til disse i. drikkebegre av '
            'ekspandert polystyren (EPS) og lokk til disse" (§2b-3) / '
            'EN: "It is prohibited to place on the market … g. food containers made of expanded polystyrene (EPS) … '
            'h. beverage packaging made of expanded polystyrene (EPS) and caps and lids thereof i. beverage cups '
            'made of expanded polystyrene (EPS) and lids thereof" (§2b-3)'
        ),
        "comments": "Take-away/food-service EPS prohibition; high litter-impact items.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3(d, e): Straws and stirrers ban—single-use plastic straws and beverage stirrers prohibited "
            "(subject to medical-device exemption for straws)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Det er forbudt å bringe i omsetning … d. sugerør e. rørepinner til drikkevarer" dersom de er '
            'laget helt eller delvis av plast og er beregnet for engangsbruk (§2b-3) / '
            'EN: "It is prohibited to place on the market … d. straws e. beverage stirrers" if made wholly or partly '
            'of plastic and intended for single use (§2b-3)'
        ),
        "comments": "Hospitality/retail ban; iconic marine litter items.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Legal basis: Adopted under produktkontrolloven §4 by Klima- og miljødepartementet, implementing "
            "EEA-relevant EU Single-Use Plastics Directive (2019/904) product restrictions in Norwegian law."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Fastsatt av Klima- og miljødepartementet 18. desember 2020 med hjemmel i lov 11. juni 1976 nr. 79 '
            'om kontroll med produkter og forbrukertjenester (produktkontrolloven) § 4" (hjemmel) / '
            'EN: "Established by the Ministry of Climate and Environment on 18 December 2020 pursuant to the Act '
            'of 11 June 1976 No. 79 relating to the control of products and consumer services (Product Control Act) '
            '§4" (legal basis)'
        ),
        "comments": "Enabling act and EEA SUP Directive linkage.",
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
    ws.title = "4P Index - SUP 3200"

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

    output_path = "/workspace/4P_Index_Produktforskrift_SUP_3200.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
