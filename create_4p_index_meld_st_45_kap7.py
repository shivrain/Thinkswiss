#!/usr/bin/env python3
"""Generate 4P Index Excel for Meld. St. 45 (2016–2017) Kap. 7 (marine litter and microplastics)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Meld. St. 45 (2016–2017) – Kap. 7: Marin forsøpling og mikroplast "
        "(White Paper to the Storting No. 45 (2016–2017) – Ch. 7: Marine litter and microplastics)"
    ),
    "policy_url": "https://www.regjeringen.no/no/dokumenter/meld.-st.-45-20162017/id2558274/?ch=7",
    "policy_year": 2017,
    "policy_objective": (
        "National government strategy ('plaststrategi') on marine litter and microplastic spread, presented "
        "in the White Paper on waste as a resource and circular economy. Sets ambitions, priority measures "
        "and international engagement, and serves as the government's follow-up to Parliament's request "
        "for a microplastics action plan—covering prevention, cleanup, land-based and sea-based sources, "
        "and Nordic/EU/global cooperation."
    ),
    "policy_target": 0.5,
    "policy_target_text": (
        'NO: "Regjeringen har som ambisjon at havområder og ferskvann ikke skal tilføres plast og mikroplast." '
        '(Kap. 7.1) / '
        'EN: "The Government ambition is that marine areas and freshwater shall not receive plastic and '
        'microplastics." (Ch. 7.1)'
    ),
    "policy_type": 0.20,
    "policy_type_justification": (
        "Government white paper / national strategy chapter (stortingsmelding) under Klima- og "
        "miljødepartementet; not directly binding law but frames parliamentary debate and subsequent "
        "regulations, subsidies and international initiatives. Presented June 2017 as part of Meld. St. 45 "
        "on waste policy and circular economy."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "marine, fisheries, aquaculture, municipalities, shipping, ports, tourism, cosmetics, "
        "road transport, sports facilities, waste management, retail, construction"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "litter/pollution, end-of-life, use/consumption, environmental leakage, recycling"
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        'NO: "Tilskuddsmidlene ble i 2017 økt fra 10 til 35 millioner kroner" (Miljødirektoratets '
        'tilskuddsordning for opprydding); "Det ble enighet om å tildele 300 millioner kroner, blant annet '
        'til økt kassering av fritidsbåter" (Kap. 7.2.2.1) / '
        'EN: "Subsidy funds were increased from 10 to 35 million kroner in 2017" (Miljødirektoratet cleanup '
        'scheme); "Agreement was reached to allocate 300 million kroner, inter alia for increased scrapping '
        'of leisure boats" (Ch. 7.2.2.1)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7: National plastic strategy—chapter constitutes the Government's strategy on marine litter "
            "and microplastic spread and serves as follow-up to Parliament's request (anmodningsvedtak) for "
            "a microplastics action plan."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Dette kapitlet utgjør regjeringens strategi for arbeidet med marin forsøpling og spredning av '
            'mikroplast. Omtalen her vil også tjene som regjeringens oppfølgning av Stortingets '
            'anmodningsvedtak om en handlingsplan for mikroplast." (Kap. 7 innledning) / '
            'EN: "This chapter constitutes the Government strategy for work on marine litter and spread of '
            'microplastics. The account here will also serve as the Government follow-up to Parliament request for a '
            'request for a microplastics action plan." (Ch. 7 introduction)'
        ),
        "comments": "Parent strategy instrument; frames all subsequent measures in the chapter.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.1: Overarching ambition—marine areas and freshwater shall not receive plastic or "
            "microplastics; aligned with SDG 14.1 and OSPAR marine litter action plan goals."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'NO: "Regjeringen har som ambisjon at havområder og ferskvann ikke skal tilføres plast og '
            'mikroplast." (Kap. 7.1) / '
            'EN: "The Government ambition is that marine areas and freshwater shall not receive plastic and '
            'microplastics." (Ch. 7.1)'
        ),
        "comments": "Strategic ambition (ambisjon), not a binding legal target.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.2.2.1: Legal baseline—littering prohibited under forurensningsloven on land and at sea; "
            "municipalities must provide waste bins at departure points and heavily visited public areas."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Ifølge forurensningsloven er det forbudt å forsøple. Forbudet gjelder både på land og i sjø." '
            '(Kap. 7.2.2.1); "kommunene skal sørge for oppsetting og tømming av avfallsbeholdere på '
            'utfartssteder og andre sterkt besøkte offentlige steder" (Kap. 7.2.2.1) / '
            'EN: "According to the Pollution Control Act, littering is prohibited. The prohibition applies both '
            'on land and at sea." (Ch. 7.2.2.1); "municipalities shall ensure placement and emptying of waste '
            'bins at departure points and other heavily visited public places" (Ch. 7.2.2.1)'
        ),
        "comments": "References existing binding law (forbudt å forsøple); strategy reinforces.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 7.2.2.1: Extended producer responsibility for fishery and aquaculture equipment—Government "
            "intends EPR scheme where producers/importers take responsibility for gear throughout lifecycle "
            "including end-of-life and prevention of gear loss."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen tar sikte på å innføre en produsentansvarsordning for henholdsvis fiskeri- og '
            'oppdrettsnæringen der produsenter og importører av utstyr … får et ansvar for produktene '
            'gjennom hele livsløpet, også når de har blitt avfall (utvidet produsentansvar)." (Kap. 7.2.2.1) / '
            'EN: "The Government aims to introduce an extended producer responsibility scheme for the fishery '
            'and aquaculture sectors whereby producers and importers of equipment … take responsibility for '
            'products throughout their lifecycle, including when they become waste." (Ch. 7.2.2.1)'
        ),
        "comments": "Planned EPR (tar sikte på); Miljødirektoratet tasked with study.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.2.2.1 / Boks 7.1: Fishing for Litter—free port delivery of marine litter collected by "
            "fishers; Government to establish permanent scheme based on pilot (28 vessels, 4 ports); "
            "NOK 1 million budget increase 2017."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet skal derfor utrede muligheten for å etablere et system som sikrer at fiskere '
            'og andre uten merkostnad kan levere til havn avfall som de har plukket fra havet (marint avfall)." '
            '(Kap. 7.2.2.1); "Regjeringen har … foreslått å styrke ordningen med 1 million kroner" (Boks 7.1) / '
            'EN: "The Norwegian Environment Agency shall therefore examine the possibility of establishing a '
            'system ensuring that fishers and others can deliver waste collected from the sea to port without '
            'extra cost." (Ch. 7.2.2.1); "The Government has … proposed strengthening the scheme by 1 million '
            'kroner" (Box 7.1)'
        ),
        "comments": "Pilot operational; permanent free-delivery scheme under study.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 7.2.2.1: Light plastic carrier bags—Norway to implement EU rules requiring sustained "
            "reduction (90 bags/person by 2019, 40 by 2025); Government supports industry-led charging/fund "
            "solution."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Norge skal gjennomføre EUs regelverk om reduksjon av bruk av lette plastbæreposer. '
            'Direktivet pålegger medlemslandene å oppnå en vedvarende reduksjon … til 90 innen utgangen av '
            '2019 og til 40 innen utgangen av 2025." (Kap. 7.2.2.1) / '
            'EN: "Norway shall implement EU rules on reducing use of lightweight plastic carrier bags. The '
            'Directive requires member states to achieve a sustained reduction … to 90 by end-2019 and to 40 '
            'by end-2025." (Ch. 7.2.2.1)'
        ),
        "comments": "EEA implementation commitment; industry fund under dialogue.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 7.2.2.1: Abandoned leisure boats—NOK 300 million subsidy scheme for scrapping abandoned "
            "recreational boats to prevent illegal dumping and microplastic release from composite hulls; "
            "municipal reception for smaller boats."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Det ble enighet om å tildele 300 millioner kroner, blant annet til økt kassering av '
            'fritidsbåter." (Kap. 7.2.2.1); "Dumping av båter i havet vil føre til nedbryting av båtens '
            'polyesterlag … og over tid gi tilførsel av mikroplast til havet." (Kap. 7.2.2.1) / '
            'EN: "Agreement was reached to allocate 300 million kroner, inter alia for increased scrapping '
            'of leisure boats." (Ch. 7.2.2.1); "Dumping boats at sea will lead to breakdown of the polyester layer '
            '… and over time supply microplastics to the sea." (Ch. 7.2.2.1)'
        ),
        "comments": "Budget agreement 2017; forskrift on municipal responsibility under consultation.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.2.2.1: Municipal infringement penalties—Government proposes new forurensningsloven "
            "authority for municipalities to impose administrative fines for illegal littering as a more "
            "preventive sanction."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen forslår videre å innta en ny hjemmel i forurensningsloven som åpner for bruk av '
            'overtredelsesgebyr blant annet i saker om ulovlig forsøpling. Ved å gi kommunen en mulighet til å '
            'ilegge overtredelsesgebyr … vil en ha et fleksibelt og effektivt sanksjonssystem." (Kap. 7.2.2.1) / '
            'EN: "The Government further proposes a new provision in the Pollution Control Act enabling use of '
            'administrative fines, inter alia in cases of illegal littering. By giving municipalities the power '
            'to impose fines … a flexible and effective sanction system will be created." (Ch. 7.2.2.1)'
        ),
        "comments": "Proposed legal amendment (forslår); not yet in force at publication.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.2.2.2 / Boks 7.5: Marine litter cleanup subsidies—Miljødirektoratet grant scheme increased "
            "from NOK 10 to 35 million (2017) plus NOK 5 million in revised budget; funds beach, seabed and "
            "preventive projects."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Tilskuddsmidlene ble i 2017 økt fra 10 til 35 millioner kroner." (Kap. 7.2.2.2); '
            '"20 millioner kroner er tildelt prosjekter for opprydding av marint avfall på sjøbunn og langs '
            'strender" (Boks 7.5) / '
            'EN: "Subsidy funds were increased from 10 to 35 million kroner in 2017." (Ch. 7.2.2.2); '
            '"20 million kroner allocated to projects for cleanup of marine waste on the seabed and along '
            'beaches" (Box 7.5)'
        ),
        "comments": "Budget instrument; operational grant scheme via Miljødirektoratet.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.2.2.2: Oil protection and environment centre—new centre in Lofoten/Vesterålen as hub for "
            "knowledge, coordination and cost-effective marine plastic litter cleanup logistics."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen har besluttet at det skal opprettes et oljevern- og miljøsenter i Lofoten/Vesterålen, '
            'med kompetanse og tiltak knyttet til opprydding av marin forsøpling som en sentral del av senterets '
            'arbeid." (Kap. 7.2.2.2) / '
            'EN: "The Government has decided to establish an oil protection and environment centre in '
            'Lofoten/Vesterålen, with expertise and measures on marine litter cleanup as a central part of the '
            'centre work." (Ch. 7.2.2.2)'
        ),
        "comments": "Institutional measure (besluttet); coordination and knowledge hub.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Kap. 7.2.3: Municipal microplastics subsidy—Government will introduce municipal grant scheme for "
            "local measures against marine litter and microplastic pollution."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen vil innføre en kommunal tilskuddsordning til gjennomføring av aktuelle tiltak på '
            'dette området." (Kap. 7.2.3) / '
            'EN: "The Government will introduce a municipal subsidy scheme for implementation of relevant '
            'measures in this area." (Ch. 7.2.3)'
        ),
        "comments": "Planned municipal grants (vil innføre); key local implementation instrument.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Kap. 7.2.3: Tyre wear microplastics—largest land-based microplastic source (~5,000 t/year); "
            "Miljødirektoratet to study road washing and drainage treatment to capture particles from roads "
            "and impervious surfaces."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Det er anslått at slitasje fra bildekk er den største kjente enkeltkilden til mikroplast etter '
            'marin forsøpling, og utgjør rundt 5 000 tonn årlig i Norge." (Kap. 7.2.3); '
            '"Miljødirektoratet er bedt om å utrede virkemidler for økt veivasking … og … renseløsninger" '
            '(Kap. 7.2.3) / '
            'EN: "Tyre wear is estimated to be the largest known single source of microplastics after marine '
            'litter, at around 5,000 tonnes annually in Norway." (Ch. 7.2.3); "Miljødirektoratet asked to '
            'examine instruments for increased road washing … and … treatment solutions" (Ch. 7.2.3)'
        ),
        "comments": "Study-phase measure; major microplastics source identified.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Kap. 7.2.3 / Boks 7.6: Artificial turf rubber granulate—~1,500 t/year microplastic loss; "
            "Miljødirektoratet to study alternatives and possible regulation; dialogue with Norges fotballforbund."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Det er videre anslått at tap av gummigranulat fra kunstgressbaner kan stå for omtrent 1500 '
            'tonn mikroplast årlig." (Kap. 7.2.3); "Miljødirektoratet skal nå utrede muligheter og kostnader '
            'ved å erstatte gummigranulat i kunstgressbaner med mer miljøvennlig alternativer." (Kap. 7.2.3) / '
            'EN: "Loss of rubber granulate from artificial turf pitches is estimated at approximately 1,500 '
            'tonnes of microplastics annually." (Ch. 7.2.3); "Miljødirektoratet shall examine possibilities and '
            'costs of replacing rubber granulate in artificial turf with more environmentally friendly '
            'alternatives." (Ch. 7.2.3)'
        ),
        "comments": "Regulatory study; sports-facility microplastics source.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 7.2.3: Microplastics in cosmetics—Parliament requested ban; Government supports EU-level "
            "regulation over national ban; will review Swedish proposal and pursue Nordic/EU influence."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Stortinget har anmodet regjeringen om å fremme forslag med sikte på å forby mikroplast i '
            'kroppspleieprodukter." (Kap. 7.2.3); "Miljødirektoratet anbefaler derfor at norske myndigheter '
            'bør støtte opp om en internasjonal regulering … framfor å innføre et nasjonalt forbud." (Kap. 7.2.3) / '
            'EN: "Parliament has requested the Government to propose a ban on microplastics in personal care '
            'products." (Ch. 7.2.3); "Miljødirektoratet therefore recommends that Norwegian authorities should '
            'support international regulation … rather than introducing a national ban." (Ch. 7.2.3)'
        ),
        "comments": "Parliament anmodningsvedtak; Government prefers EU harmonised approach.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.3.1: Nordic plastics programme—under Norwegian 2017 Council presidency, programme covering "
            "prevention, waste systems, marine litter cleanup, microplastics knowledge and bioplastic impacts."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Under Norges ledelse er det vedtatt et eget nordisk program for å redusere miljøkonsekvensene '
            'av plast." (Kap. 7.3.1); programme areas include "Samarbeid om tiltak for å stanse marin '
            'plastforsøpling" and "Øke kunnskapen om mikroplast" (Kap. 7.3.1) / '
            'EN: "Under Norway leadership a dedicated Nordic programme to reduce environmental impacts of '
            'plastic has been adopted." (Ch. 7.3.1); areas include "Cooperation on measures to stop marine '
            'plastic pollution" and "Increase knowledge on microplastics" (Ch. 7.3.1)'
        ),
        "comments": "Regional cooperation framework; Norway as 2017 presidency driver.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.3 / Boks 7.7: Clean Seas campaign—Norway joined UN #CleanSeas global campaign (Feb 2017) "
            "for five years to raise awareness and mobilise action against marine litter."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'NO: "Ti land sluttet seg da til kampanjen, inkludert Norge." (Boks 7.7); "Formålet med kampanjen '
            'er å spre kunnskap om kildene, men også konsekvensene av marin forsøpling og fremme tiltak." '
            '(Boks 7.7) / '
            'EN: "Ten countries joined the campaign then, including Norway." (Box 7.7); "The purpose of the '
            'campaign is to spread knowledge about sources and consequences of marine litter and promote '
            'measures." (Box 7.7)'
        ),
        "comments": "Awareness/soft instrument; UN Environment campaign 2017–2021.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.3.3 / Boks 7.9: International development aid—NOK 100 million aid programme to reduce marine "
            "litter and microplastics in developing countries with weak waste management (announced at UN "
            "Ocean Conference 2017)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen har i Meld. St. 22 (2016–2017) … lansert et nytt bistandsprogram for å redusere '
            'marin forsøpling og mikroplast i havet." (Kap. 7.3.3); "bistandsprogrammet på 100 millioner kroner '
            'for å bekjempe marin forsøpling" (Boks 7.9) / '
            'EN: "The Government … launched a new aid programme to reduce marine litter and microplastics in '
            'the sea." (Ch. 7.3.3); "aid programme of 100 million kroner to combat marine litter" (Box 7.9)'
        ),
        "comments": "Global source-reduction instrument; addresses 80% land-based marine litter estimate.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 7.4 / Boks 7.10: Monitoring and knowledge—strengthen monitoring of marine litter in Norwegian "
            "sea areas including seabed; OSPAR guidelines for beach litter, fulmar stomach plastics and seabed "
            "monitoring."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringen vil: styrke overvåkingen av marin forsøpling i norske havområder, blant annet '
            'gjennom å styrke kunnskapen om marin forsøpling på havbunnen" (Kap. 7.4); OSPAR guidelines for '
            '"overvåking av strandsøppel" and "overvåking av plastpartikler i magen hos havhest" (Boks 7.10) / '
            'EN: "The Government will: strengthen monitoring of marine litter in Norwegian sea areas, including '
            'knowledge of seabed marine litter" (Ch. 7.4); OSPAR guidelines for "monitoring of beach litter" '
            'and "monitoring of plastic particles in fulmar stomachs" (Box 7.10)'
        ),
        "comments": "Evidence base for future instruments; OSPAR reporting obligations.",
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
    ws.title = "4P Index - Meld.St.45 Kap7"

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

    output_path = "/workspace/4P_Index_Meld_St_45_Kap7.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
