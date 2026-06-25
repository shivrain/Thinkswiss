#!/usr/bin/env python3
"""Generate 4P Index Excel for Produktforskriften (Product Regulation)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om begrensning i bruk av helse- og miljøfarlige kjemikalier og andre produkter "
        "(Produktforskriften) (Regulation on restriction of the use of hazardous chemicals and other "
        "products – Product Regulation)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-922",
    "policy_year": 2026,
    "policy_objective": (
        "Umbrella product regulation under produktkontrolloven and forurensningsloven implementing EU/EEA "
        "chemical and product rules—restricting hazardous substances in articles and mixtures, with "
        "Chapter 2b as the parent instrument for single-use plastic (SUP) measures, recycled-content "
        "requirements, tethered caps, and other plastic-product restrictions relevant to production and consumption."
    ),
    "policy_target": 1,
    "policy_target_text": (
        '"Fra 1. januar 2025 … skal … flaskene inneholder minst 25 % materialgjenvunnet plast" (§2b-7); '
        '"Fra 1. januar 2030 … minst 30 % materialgjenvunnet plast" (§2b-7); '
        '"Formålet med dette kapitlet er å forebygge og redusere miljøpåvirkningen fra enkelte plastprodukter" '
        "(§2b-1)"
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under produktkontrolloven and forurensningsloven; principal Norwegian "
        "implementation vehicle for EU RoHS, Packaging Directive heavy-metal limits, REACH Annex XVII, POPs "
        "Regulation, and SUP Directive (2019/904). Chapter 2b added July 2021; recycled-content and tethered-cap "
        "rules December 2024; EEA references updated March 2026 (FOR-2026-03-23-509)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "manufacturing, retail, chemicals, packaging, beverages, food service, tobacco, personal care, "
        "electronics, automotive, waste management, import/export"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, design, consumption, use, end-of-life, recycling, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        '"Miljødirektoratet kan ved overtredelse av § 2b-3, § 2b-4, § 2b-5 første ledd, § 2b-6 eller § 2b-7 '
        'ilegge den ansvarlige … et overtredelsesgebyr" (§2b-8); Chapter 6b caps infringement fines at '
        "15× folketrygdens grunnbeløp for enterprises (§6b-2)"
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§1-2: Scope—regulation applies to production, import, export, turnover and use of substances, "
            "mixtures and products except for analysis/research purposes unless individual provisions state otherwise; "
            "defines the umbrella reach of all product restrictions including plastic measures."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Forskriften gjelder ikke for produksjon, import, eksport, omsetning og bruk av stoff, stoffblandinger '
            'og produkter til analyse- og forskningsformål, med mindre annet fremgår av de enkelte bestemmelsene." '
            "(§1-2)"
        ),
        "comments": "Framing scope instrument; R&D exemption reduces unconditional score.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2-15: Heavy metals in packaging—prohibition on producing, importing, exporting or placing on the "
            "market packaging where total lead, cadmium, mercury and hexavalent chromium exceeds 100 mg/kg; "
            "exemptions for recycled glass/plastic crates and pallets."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å produsere, importere, eksportere og omsette emballasje der det samlede innhold av '
            'bly, kadmium, kvikksølv og seksverdig krom overstiger 100 mg/kg." (§2-15); exemption for '
            'plastic crates/pallets with recycled content (§2-15)'
        ),
        "comments": "EU Packaging Directive heavy-metal limit; directly affects plastic packaging streams.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2-30: REACH Annex XVII restrictions—export prohibition on substances/mixtures/products that may not "
            "be placed on the market under listed Annex XVII entries, including cadmium, phthalates (DEHP, DBP, BBP), "
            "nickel and other substances used in plastic articles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å eksportere stoff, stoffblandinger eller produkter som i henhold til følgende poster '
            'i REACH-forordningen vedlegg XVII ikke er tillatt å bringe i omsetning" (§2-30); includes Post 51 '
            '"Di(2-ethylheksyl)ftalat (DEHP), dibutylftalat (DBP) og butylbenzylftalat (BBP)" and Post 52 "Ftalater"'
        ),
        "comments": "National export mirror of REACH restrictions on plasticisers and additives in plastics.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2-32: SCIP/SVHC notification—suppliers must submit information on Substances of Very High Concern "
            "(SVHC) in articles to ECHA per REACH Article 33, supporting traceability of hazardous substances in "
            "plastic and other consumer products."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Enhver leverandør av et produkt … skal gi opplysninger om stoffer med svært uønskede egenskaper '
            '(SVHC) i produkter i henhold til REACH-forordningen artikkel 33 nr. 1 til Det europeiske '
            'kjemikaliebyrået (ECHA)." (§2-32)'
        ),
        "comments": "Reporting instrument (skal gi opplysninger); consumer-direct supply exempt.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§2a-10: EEE waste marking and information—electrical/electronic products must bear crossed-bin "
            "wheeled-bucket symbol and producers must provide material/hazardous-substance/disposal information "
            "to WEEE treatment facilities, supporting proper plastic-fraction recovery from EEE waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"EE-produkter skal merkes med et piktogram som består av en overkrysset avfallsbeholder på hjul" '
            "(§2a-10); producers shall provide information on materials/components and hazardous substances "
            '"for miljømessig forsvarlig behandling av kasserte EE-produkter" (§2a-10)'
        ),
        "comments": "RoHS/WEEE interface; many EEE housings are plastic.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§2b-1: Purpose of Chapter 2b—to prevent and reduce environmental impacts of certain plastic products, "
            "including reducing and preventing marine littering and spread of microplastics."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med dette kapitlet er å forebygge og redusere miljøpåvirkningen fra enkelte plastprodukter, '
            'herunder redusere og forebygge marin forsøpling og spredning av mikroplast." (§2b-1)'
        ),
        "comments": "Parent chapter for EU Single-Use Plastics Directive (2019/904) implementation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-3: SUP prohibition—ban on placing on the market single-use plastic cotton buds, cutlery, plates, "
            "straws, stirrers, balloon sticks, EPS food containers, EPS beverage packaging/cups and lids (with "
            "medical/industrial exemptions)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å bringe i omsetning følgende produkter dersom de er laget helt eller delvis av plast '
            'og er beregnet for engangsbruk" including "sugerør", "matbeholdere av ekspandert polystyren (EPS)" '
            'and "drikkevareemballasje av ekspandert polystyren (EPS)" (§2b-3)'
        ),
        "comments": "Core SUP ban list; absolute prohibition (forbudt).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-4: Oxo-degradable plastic ban—prohibition on placing on the market products made of oxo-degradable "
            "plastic that fragment into microplastics through oxidation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å bringe i omsetning produkter laget av okso-nedbrytbar plast." (§2b-4)'
        ),
        "comments": "Prevents microplastic-generating plastic types; mandatory market ban.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-5: SUP marking requirements—operators placing sanitary items, wet wipes, tobacco filters and "
            "drink cups on the market must ensure harmonised marking per EU Regulation 2020/2151 (EEA-incorporated)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Den som bringer i omsetning følgende produkter, skal sørge for at produktene er merket i henhold til '
            'merkekravene … dersom de er laget helt eller delvis av plast og er beregnet for engangsbruk" (§2b-5); '
            "EU 2020/2151 applies as regulation (§2b-5)"
        ),
        "comments": "Behaviour-change and litter-awareness labelling; mandatory (skal sørge for).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§2b-6: Tethered caps and lids—single-use plastic beverage containers up to 3 L may only be placed on "
            "the market if plastic caps/lids remain attached throughout intended use, meeting EN 17665:2022+A1:2023; "
            "written compliance documentation required."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"En drikkevarebeholder til engangsbruk … kan bare bringes i omsetning hvis korken eller lokket sitter '
            'fast i beholderen i hele perioden produktet er ment å brukes" (§2b-6); standard EN 17665:2022+A1:2023 '
            "(§2b-6)"
        ),
        "comments": "EU SUP design requirement for cap collection; added December 2024.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2b-7: Recycled plastic content in PET drink bottles—from 2025 ≥25% recycled plastic in single-use "
            "PET bottles ≤3 L (de minimis <1,000 kg/year); from 2030 ≥30%; annual reporting via approved DRS (Kap. 6) "
            "or packaging PRO (Kap. 7)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Fra 1. januar 2025 … skal … flaskene inneholder minst 25 % materialgjenvunnet plast" (§2b-7); '
            '"Fra 1. januar 2030 … minst 30 % materialgjenvunnet plast" (§2b-7); annual report to approved return '
            "system or PRO (§2b-7)"
        ),
        "comments": "Binding recycled-content targets; cross-linked to Avfallsforskriften Kap. 6/7 reporting.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "§2b-8: Infringement fines for Chapter 2b—Miljødirektoratet may impose infringement fees for breaches "
            "of §§2b-3–2b-7; governed by Chapter 6b (up to 15× folketrygd basic amount for enterprises)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Miljødirektoratet kan ved overtredelse av § 2b-3, § 2b-4, § 2b-5 første ledd, § 2b-6 eller § 2b-7 '
            'ilegge den ansvarlige for overtredelsen et overtredelsesgebyr." (§2b-8)'
        ),
        "comments": "Enforcement instrument (kan ilegge); added December 2024.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4-1: POPs Regulation implementation—EU Regulation 2019/1021 on persistent organic pollutants applies "
            "as Norwegian law with EEA adaptations, restricting production, placement on market and use of POP "
            "substances including those used as plastic additives (e.g. deca-BDE, HBCDD)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Forordning (EU) 2019/1021 om persistente organiske forbindelser (POPs-forordningen) … gjelder som "
            "forskrift (§4-1); includes amendments on PBDE, HBCDD, UV-328 etc. through 2025/1930"
        ),
        "comments": "Broad chemical restriction umbrella affecting plastic formulations and waste.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4-2: POPs export ban—prohibition on exporting listed POP substances (including deca-BDE, HBCDD, SCCPs) "
            "to non-EEA countries to the extent they are banned from being placed on the market under §4-1."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å eksportere følgende stoff … til land utenfor EØS-området i den utstrekning stoffene er '
            'forbudt å omsette i henhold til § 4-1" including "Dekabromdifenyleter" (§4-2)'
        ),
        "comments": "Export control on flame retardants used in plastics; mandatory (forbudt).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2-18: Battery marking and design—producers/importers must mark batteries with crossed-bin symbol and "
            "Hg/Cd/Pb chemical symbols; battery-powered products must be designed for easy battery removal to "
            "facilitate separate collection of plastic-cased batteries."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Produsenter og importører plikter å sørge for at batterier er merket med … symbol med overkrysset '
            'avfallsbeholder" (§2-18); "batteridrevne produkter plikter å sørge for at produktene er konstruert '
            'slik at batteriene lett kan fjernes" (§2-18)'
        ),
        "comments": "EU Battery Directive provisions; many portable batteries in plastic housings.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7-2: Supervision—Miljødirektoratet (or delegated authority) supervises compliance with the regulation; "
            "Sjøfartsdirektoratet supervises Chapter 6 (ozone substances) on board Norwegian civil ships."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Miljødirektoratet eller den Klima- og miljødepartementet bemyndiger, fører tilsyn med at bestemmelsene '
            'i denne forskriften og vedtak truffet i medhold av forskriften overholdes." (§7-2)'
        ),
        "comments": "Governance/enforcement framework for all chapters including 2b.",
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
    ws.title = "4P Index - Produktforskr"

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

    output_path = "/workspace/4P_Index_Produktforskrift.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
