#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften Kap. 7A (SUP litter EPR)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Avfallsforskriften – Kap. 7A: Utvidet produsentansvar for enkelte engangsprodukter av plast "
        "(Waste Regulation – Ch. 7A: Extended producer responsibility for single-use plastic products)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930/KAPITTEL_8",
    "policy_year": 2025,
    "policy_objective": (
        "Chapter 7A of the Norwegian Waste Regulation (FOR-2004-06-01-930) under forurensningsloven, "
        "added 12 November 2025 (forskrift 12 nov 2025 nr. 2278). Implements EU Single-Use Plastics "
        "Directive (2019/904) extended producer responsibility for litter: producers of listed SUP items "
        "must cover municipalities' costs for litter clean-up and public-bin collection of discarded "
        "products, plus information campaigns, producer responsibility organisation membership, cost "
        "surveys, and annual payments from 1 January 2027 (first payment 1 October 2028). Covers food "
        "containers, flexible food packaging, beverage containers, cups, thin carrier bags, wet wipes, "
        "tobacco filters, and balloons."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Plikten til kostnadsdekning gjelder kostnader påløpt fra og med 1. januar 2027"; '
        '"Utbetaling til kommunene skal skje innen 1. oktober hvert år, første gang 1. oktober 2028" '
        '(§7A-4) / '
        'EN: "The cost-coverage obligation applies to costs incurred from 1 January 2027"; '
        '"Payment to municipalities shall be made by 1 October each year, for the first time on '
        '1 October 2028" (§7A-4)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under forurensningsloven and lov om bærekraftige produkter "
        "og verdikjeder §4(i). Chapter 7A added 12 November 2025 (forskrift 12 nov 2025 nr. 2278); "
        "fee provisions amended 17 December 2025 (forskrift 17 des 2025 nr. 2622, in force 1 January 2026). "
        "Implements EEA obligations under Directive (EU) 2019/904 on SUP litter EPR; complements "
        "Produktforskriften Kap. 2b (product bans/marking) and Kap. 7 packaging EPR."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "retail; food service; municipalities; waste management; tobacco; packaging; "
        "producer responsibility organisations"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "use/consumption, litter/pollution, end-of-life"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Produsent … skal dekke en andel av kommunenes nødvendige kostnader til opprydning av '
        'forsøpling" (§7A-4); "Gebyr … Kr 203 900" (godkjenning, §7A-27); "Årsgebyret utgjør kr 20 400" '
        '(§7A-28) / '
        'EN: "The producer … shall cover a share of municipalities\' necessary litter clean-up costs" (§7A-4); '
        '"Fee … NOK 203,900" (approval, §7A-27); "Annual fee is NOK 20,400" (§7A-28)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-1: Chapter purpose—prevent and reduce environmental impact from certain single-use plastic "
            "products, including litter and microplastic spread, and contribute to a circular economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Formålet med dette kapittelet er å forebygge og redusere miljøpåvirkningen fra enkelte '
            'engangsprodukter av plast, herunder redusere og forebygge forsøpling og spredning av mikroplast, '
            'samt bidra til en sirkulær økonomi." (§7A-1) / '
            'EN: "The purpose of this chapter is to prevent and reduce the environmental impact of certain '
            'single-use plastic products, including reducing and preventing litter and the spread of microplastics, '
            'and to contribute to a circular economy." (§7A-1)'
        ),
        "comments": "Frames SUP litter EPR under Avfallsforskriften; added FOR-2025-11-12-2278.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§7A-2: Scope—eight listed SUP product categories (food containers, flexible food packaging, "
            "beverage containers ≤3 L, cups, thin carrier bags <50 µm, wet wipes, tobacco filters, balloons); "
            "applies alongside Kap. 6 DRS and Kap. 7 packaging EPR where relevant; excludes Svalbard/Jan Mayen."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "matbeholdere … fleksibel innpakning … drikkevarebeholdere som rommer 3 liter eller mindre … '
            'drikkebegre … bæreposer laget av plast som er tynnere enn 50 mikrometer … våtservietter … '
            'tobakksprodukter med filter … ballonger" (§7A-2) / '
            'EN: "food containers … flexible packaging … beverage containers of 3 litres or less … beverage cups … '
            'plastic carrier bags thinner than 50 micrometres … wet wipes … tobacco products with filters … '
            'balloons" (§7A-2)'
        ),
        "comments": "Product list mirrors EU SUP Directive Article 8 litter EPR scope.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-3: Definitions—plastic (REACH polymer), single-use plastic product, producer (commercial "
            "placer on market), placing on market, and tobacco products."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "engangsprodukt av plast; et produkt som helt eller delvis er fremstilt av plast, og som ikke '
            'er utformet … for å bli brukt om igjen flere ganger"; "produsent; enhver som ervervsmessig bringer '
            'i omsetning engangsprodukter av plast" (§7A-3) / '
            'EN: "single-use plastic product; a product made wholly or partly from plastic and not designed … '
            'to be used again multiple times"; "producer; any person who places single-use plastic products on '
            'the market on a commercial basis" (§7A-3)'
        ),
        "comments": "Legal definitions aligning with SUP Directive terminology.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-4(1): Core litter EPR—producer shall cover share of municipalities' necessary costs for "
            "litter clean-up of listed SUP products and subsequent transport and treatment of that waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Produsent … skal dekke en andel av kommunenes nødvendige kostnader til opprydning av '
            'forsøpling fra slike produkter og etterfølgende transport og behandling av dette avfallet"; '
            '"gjelder kostnader påløpt fra og med 1. januar 2027" (§7A-4) / '
            'EN: "The producer … shall cover a share of municipalities\' necessary costs for cleaning up litter '
            'from such products and subsequent transport and treatment of that waste"; '
            '"applies to costs incurred from 1 January 2027" (§7A-4)'
        ),
        "comments": "Primary SUP litter-cost EPR obligation; operative from 1 January 2027.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§7A-4(2): Public-bin collection EPR—producer of products a–e and g shall cover share of "
            "municipalities' costs for collection from public waste bins, including infrastructure and operation, "
            "transport and treatment."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Produsent … bokstav a–e og g skal dekke en andel av kommunenes nødvendige kostnader i '
            'forbindelse med innsamling av slike produkter som kasseres i avfallsbeholdere på offentlig sted, '
            'herunder kostnader til infrastruktur og driften av innsamlingssystemet" (§7A-4) / '
            'EN: "The producer … letters a–e and g shall cover a share of municipalities\' necessary costs for '
            'collection of such products discarded in waste bins in public places, including infrastructure and '
            'operation of the collection system" (§7A-4)'
        ),
        "comments": "Public-space bin collection cost coverage; excludes balloons if waste analyses show negligible findings.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-4(3)–(4): Payment mechanics—costs calculated per Part IV; producer share based on per-product "
            "rates (§7A-20) and units placed on market; annual payment to municipalities by 1 October "
            "(first payment 1 October 2028); balloon exemption if waste analyses show no/significant findings."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Andelen hver enkelt produsent skal dekke beregnes basert på satsene per produkt fastsatt etter '
            '§ 7A-20 og antall produkter"; "Utbetaling til kommunene skal skje innen 1. oktober hvert år, '
            'første gang 1. oktober 2028" (§7A-4) / '
            'EN: "Each producer\'s share shall be calculated based on per-product rates set under §7A-20 and the '
            'number of products placed on market"; "Payment to municipalities shall be made by 1 October each year, '
            'for the first time on 1 October 2028" (§7A-4)'
        ),
        "comments": "Quantified payment schedule; links to cost-calculation methodology in Part IV.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§7A-5: Producer information and awareness—consumer information on reuse alternatives, proper "
            "waste handling, sewer impacts, and litter/environmental effects; annual nationwide campaigns."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "sørge for at det gis informasjon til forbrukere og fremme ansvarlig forbrukeradferd for å '
            'redusere forsøpling"; "gjennomføre årlige landsdekkende informasjonskampanjer" (§7A-5) / '
            'EN: "ensure that information is provided to consumers and promote responsible consumer behaviour '
            'to reduce litter"; "conduct annual nationwide information campaigns" (§7A-5)'
        ),
        "comments": "SUP Directive consumer-awareness obligation; applies to products a–f and h.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-6: Authorised representative—Norwegian producers exporting to other EEA states must appoint "
            "representative; non-Norwegian producers may appoint Norwegian representative with written mandate "
            "to fulfil chapter obligations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent … etablert i Norge … som ervervsmessig gjør engangsprodukter av plast tilgjengelig '
            'for første gang i et annet EU/EØS-land … skal utpeke en representant"; '
            '"Produsent som ikke er etablert i Norge kan utpeke en representant som er etablert i Norge" (§7A-6) / '
            'EN: "A producer established in Norway … who makes single-use plastic products available for the '
            'first time in another EU/EEA country … shall appoint a representative"; '
            '"A producer not established in Norway may appoint a representative established in Norway" (§7A-6)'
        ),
        "comments": "Cross-border producer registration aligned with SUP Directive Article 8a.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-7: Mandatory membership in approved producer responsibility organisation (PRO)—producer "
            "must join approved PRO to fulfil §§7A-4, 7A-5, 7A-19–21, 7A-23 obligations; fallback self-compliance "
            "if PRO fails."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Produsent skal oppfylle sine plikter … ved å etablere eller melde seg inn i et godkjent '
            'produsentansvarsselskap"; "Dersom produsentansvarsselskapet ikke ivaretar pliktene … skal produsenten '
            'selv sørge for at pliktene oppfylles" (§7A-7) / '
            'EN: "The producer shall fulfil its obligations … by establishing or joining an approved producer '
            'responsibility organisation"; "If the PRO does not fulfil the obligations … the producer shall '
            'ensure compliance itself" (§7A-7)'
        ),
        "comments": "Binding PRO membership; infringement fee under §7A-29 for breach of first paragraph.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-8: Producer annual reporting to PRO—information necessary for PRO to fulfil §7A-14 reporting "
            "obligations; at least once yearly or more often if required."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal rapportere til produsentansvarsselskapet … den informasjon som er nødvendig"; '
            '"Rapporteringen skal skje en gang i året eller oftere" (§7A-8) / '
            'EN: "The producer shall report to the PRO … the information necessary"; '
            '"Reporting shall take place once a year or more frequently" (§7A-8)'
        ),
        "comments": "Data-flow obligation linking producers to national SUP reporting.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-9–§7A-10: PRO requirements and Miljødirektoratet approval—separate Enhetsregisteret entity, "
            "non-profit, no dividend; must be approved for all §7A-2 products; approval may be withdrawn."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal være uten økonomisk formål (non-profit)"; '
            '"Produsentansvarsselskap skal være godkjent av Miljødirektoratet"; '
            '"må godkjennes for alle produktene nevnt i § 7A-2" (§7A-9–§7A-10) / '
            'EN: "The PRO shall be non-profit"; "The PRO shall be approved by the Norwegian Environment Agency"; '
            '"must be approved for all products listed in §7A-2" (§§7A-9–10)'
        ),
        "comments": "Governance framework for collective SUP litter EPR schemes.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-11–§7A-13: PRO operational duties—fulfil members' obligations; levy fees not exceeding "
            "necessary costs; equal access for all obliged producers."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal ivareta medlemmenes plikter"; '
            '"Vederlaget skal ikke overstige det som er nødvendig"; '
            '"Alle produsenter med plikt til medlemskap … skal ha adgang … på like vilkår" (§7A-11–§7A-13) / '
            'EN: "The PRO shall fulfil members\' obligations"; "Fees shall not exceed what is necessary"; '
            '"All producers with a membership obligation … shall have access on equal terms" (§§7A-11–13)'
        ),
        "comments": "PRO cost-recovery and non-discrimination rules.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-14–§7A-16: PRO reporting and municipal payment system—annual reporting to Miljødirektoratet "
            "on quantities placed on market and municipal payments; triennial reporting on tobacco waste, "
            "rates, and municipal involvement; establish annual payment system to municipalities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal årlig rapportere til Miljødirektoratet om … mengde produkter … '
            'brakt i omsetning … utbetaling til kommuner"; '
            '"sørge for å etablere et system for utbetaling til kommuner for kostnader nevnt i § 7A-4" '
            '(§7A-14–§7A-16) / '
            'EN: "The PRO shall annually report to the Norwegian Environment Agency on … quantity of products '
            'placed on market … payments to municipalities"; '
            '"establish a system for payment to municipalities for costs under §7A-4" (§§7A-14–16)'
        ),
        "comments": "Transparency and payment infrastructure for litter EPR.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-15, §7A-17–§7A-18: PRO transparency and financial safeguards—public information on owners, "
            "members, and fees; minimum six-month financial reserves; cover Miljødirektoratet data-system costs."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "sørge for at følgende opplysninger er offentlig tilgjengelig … vederlag … i kroner per solgte '
            'produkt"; "finansielle reserver for minimum seks måneders drift"; '
            '"dekke Miljødirektoratets kostnader til utvikling og drift av datasystemer" (§7A-15, §7A-17–§7A-18) / '
            'EN: "ensure the following information is publicly available … fees … in NOK per product sold"; '
            '"financial reserves for at least six months\' operation"; '
            '"cover the Norwegian Environment Agency\'s costs for development and operation of data systems" '
            "(§§7A-15, 17–18)"
        ),
        "comments": "Financial and transparency requirements for PRO solvency.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-19: Cost surveys and waste analyses—producer shall commission triennial representative municipal "
            "cost surveys and waste analyses to determine SUP product shares in litter and collected waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal sørge for de undersøkelser som er nødvendige for å beregne kommunenes kostnader"; '
            '"avfallsanalyser for å finne andel engangsprodukter av plast … i oppryddet og innsamlet avfall"; '
            '"skal gjennomføres hvert tredje år" (§7A-19) / '
            'EN: "The producer shall ensure the investigations necessary to calculate municipalities\' costs"; '
            '"waste analyses to determine the share of single-use plastic products … in cleaned-up and collected waste"; '
            '"shall be conducted every three years" (§7A-19)'
        ),
        "comments": "Evidence base for per-product litter-cost rates under §7A-20.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-20: Annual municipal payment calculation—uniform per-capita rates per product type; triennial "
            "per-product rate from average municipal handling cost scaled nationally; multiplied by units "
            "placed on market; divided by SSB population."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Beløpet som skal utbetales til kommunene … beregnes ut fra satser per innbygger"; '
            '"fastsettes en sats per produkt brakt i omsetning for hver produkttype basert på en beregnet '
            'gjennomsnittskostnad for kommunenes håndtering"; "Samme satser skal gjelde for alle kommuner" '
            '(§7A-20) / '
            'EN: "The amount payable to municipalities … calculated from per-capita rates"; '
            '"a per-product rate shall be set for each product type based on calculated average municipal '
            'handling cost"; "The same rates shall apply to all municipalities" (§7A-20)'
        ),
        "comments": "Methodology for allocating litter and public-bin costs to producers.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§7A-21–§7A-24: Verification and governance of cost calculations—independent consultant and "
            "third-party verification; PRO cooperation on shared consultant; municipal and industry consultation "
            "before rate-setting; Miljødirektoratet may set rates by individual decision if obligations not met."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "skal utføres av en uavhengig konsulent"; "verifiseres av en tredjepart"; '
            '"interesseorganisasjon for kommunene høres … før satsene bestemmes"; '
            '"Miljødirektoratet kan fastsette satsene ved enkeltvedtak" (§7A-21–§7A-24) / '
            'EN: "shall be carried out by an independent consultant"; "verified by a third party"; '
            '"municipal interest organisation shall be consulted … before rates are determined"; '
            '"The Norwegian Environment Agency may set rates by individual decision" (§§7A-21–24)'
        ),
        "comments": "Quality assurance and fallback rate-setting for litter-cost methodology.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§7A-25: Municipal cooperation duty—municipalities must provide cost and population data and waste "
            "for analyses under §7A-19 to producers, PROs, and Miljødirektoratet."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kommunene plikter å stille til rådighet nødvendige opplysninger om kostnader og innbyggertall, '
            'samt avfall til avfallsanalyser" (§7A-25) / '
            'EN: "Municipalities shall make available necessary information on costs and population, as well as '
            'waste for waste analyses" (§7A-25)'
        ),
        "comments": "Municipal data-sharing obligation enabling litter-cost calculations.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-26: Supervision—Miljødirektoratet supervises compliance with chapter provisions and decisions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet fører tilsyn med at bestemmelsene i dette kapittelet og vedtak truffet i '
            'medhold av bestemmelsene i dette kapittelet overholdes." (§7A-26) / '
            'EN: "The Norwegian Environment Agency shall supervise compliance with the provisions of this chapter '
            'and decisions made pursuant to them." (§7A-26)'
        ),
        "comments": "Centralised enforcement authority for SUP litter EPR.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-27–§7A-28: PRO fees to state—approval processing fees (NOK 25,500–203,900 scale) and annual "
            "reporting fee (NOK 20,400); CPI-indexed from 1 January; amended FOR-2025-12-17-2622."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Gebyr skal betales etter følgende satser … Kr 203 900 … Kr 25 500"; '
            '"Årsgebyret utgjør kr 20 400 per godkjente produsentansvarsselskap" (§7A-27–§7A-28) / '
            'EN: "Fees shall be paid according to the following rates … NOK 203,900 … NOK 25,500"; '
            '"The annual fee is NOK 20,400 per approved PRO" (§§7A-27–28)'
        ),
        "comments": "Administrative fee schedule; updated December 2025 regulation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7A-29: Infringement fees—Miljødirektoratet may impose administrative fines for breach of "
            "§7A-7(1) PRO membership obligation; applies Kap. 18B infringement-fee framework."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet kan ved overtredelse av § 7A-7 første ledd ilegge den ansvarlige for '
            'overtredelsen et overtredelsesgebyr"; "gjelder bestemmelsene i denne forskriftens kapittel 18B" '
            '(§7A-29) / '
            'EN: "The Norwegian Environment Agency may impose an infringement fee on the person responsible for '
            'a breach of §7A-7(1)"; "the provisions of Chapter 18B of this regulation apply" (§7A-29)'
        ),
        "comments": "Enforcement sanction for failure to join approved PRO.",
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
    ws.title = "4P Index - Kap 7A"

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

    output_path = "/workspace/4P_Index_Avfallsforskrift_Kap7a.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
