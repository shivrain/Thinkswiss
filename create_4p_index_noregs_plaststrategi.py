#!/usr/bin/env python3
"""Generate 4P Index Excel for Noregs plaststrategi (Norwegian Plastics Strategy, 2021)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": "Noregs plaststrategi (Norway's Plastics Strategy)",
    "policy_url": "https://www.regjeringen.no/no/dokumenter/noregs-plaststrategi/id2867004/",
    "policy_year": 2021,
    "policy_objective": (
        "Comprehensive national government plastics strategy (published 2021) coordinating Norway's "
        "plastics policy across the full lifecycle—from sustainable production and consumption, through "
        "waste management and recycling, to litter/pollution and microplastics. It implements UNEA's "
        "2017 zero vision to end ocean plastic litter discharge, serves as Norway's national plastics "
        "management plan example for a future global agreement, and integrates EEA/EU product and waste "
        "rules with national measures for fisheries, aquaculture, municipalities, agriculture, transport, "
        "and clean-up."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Regjeringas visjon er ei meir berekraftig verdikjede for plast globalt, regionalt og nasjonalt"; '
        '"50 % av emballasje skal vinnast att som materiale i 2025 … 55 % i 2030" (emballasje); '
        '"minst 70 % av plastavfall frå hushald skal sorterast ut innan 2035" (kap. 6) / '
        'EN: "The Norwegian Government\'s vision is a more sustainable plastics value chain globally, '
        'regionally, and nationally"; "50 % of packaging recycled by 2025 and 55 % by 2030" (packaging); '
        '"municipalities must sort at least 70 % of plastic waste from households by 2035" (Ch. 6)'
    ),
    "policy_type": 0.20,
    "policy_type_justification": (
        "National government strategy / action plan (plaststrategi) adopted by the Norwegian Ministries "
        "(Klima- og miljødepartementet et al.), published 2021. Non-binding policy framework that "
        "coordinates and describes implemented and planned measures; operationalised through EEA/EU "
        "regulations (Produktforskriften Kap. 2b, Avfallsforskriften, REACH), national regulations "
        "(Forurensningsforskriften Kap. 23A), EPR schemes, development programmes, and sector action plans."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "all sectors; fisheries; aquaculture; shipping; municipalities; agriculture; transport; "
        "waste management; retail; packaging; construction; textiles; electronics; sports facilities"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, use/consumption, end-of-life, recycling, litter/pollution, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Regjeringa vil bruke 1,6 milliardar kroner i perioden 2019–2024" (Utviklingsprogrammet mot '
        'marin forsøpling); "1 milliard kroner over tre år" (Grønn plattform, kap. 13) / '
        'EN: "The Norwegian Government will spend NOK 1.6 billion in the period 2019–2024" (Development '
        'Programme to Combat Marine Litter); "NOK 1 billion over three years" (Green Platform Initiative, Ch. 13)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 1 Vision: Sustainable plastics value chain globally, regionally and nationally; strategy "
            "implements UNEA 2017 zero vision to end all discharge of plastic litter into the ocean in the "
            "long term and takes a comprehensive lifecycle approach to plastic pollution."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringas visjon er ei meir berekraftig verdikjede for plast globalt, regionalt og '
            'nasjonalt"; "strategien skal bidra til å førege UNEAs nullvisjon frå 2017 om å avslutte alt utslepp '
            'av plastforsøpling i havet på lang sikt" (kap. 1) / '
            'EN: "The Norwegian Government\'s vision is a more sustainable plastics value chain globally, '
            'regionally, and nationally"; "this strategy will assist in implementing the UN Environment '
            'Assembly\'s 2017 «zero vision» to end all discharge of plastic litter into the ocean in the long term" '
            "(Ch. 1)"
        ),
        "comments": "Strategic framing aligned with UNEA zero vision; not a standalone legal obligation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 1/3: Strategy as Norway's national plastics management plan—example of the national plan "
            "obligation Norway advocates in a future legally binding global agreement on plastic pollution."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Alle land bør ha ei plikt til å utarbeide nasjonale plastforvaltningsplanar"; '
            '"Denne strategien er Noregs eksempel på ein slik plastforvaltningsplan" (kap. 1) / '
            'EN: "all countries should prepare national plastic management plans"; '
            '"This strategy is Norway\'s example of such a plastics management plan" (Ch. 1)'
        ),
        "comments": "National plan model for global agreement advocacy; strategy itself is the instrument.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 1/5.4: Non-toxic material cycles—same limit values for hazardous substances in primary and "
            "secondary plastic raw materials to enable safe recycling and prevent spreading of SVHCs through "
            "recycled products."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Ein meir sirkulær plastøkonomi må vere giftfri"; "same grenseverdi … for farlege stoff i '
            'plastprodukt produsert frå primære og sekundære råmaterial" (kap. 5.4) / '
            'EN: "A more circular plastics economy must be non-toxic"; "the same limit value must be set for '
            'hazardous substances in plastic products produced from primary and secondary raw materials" (Ch. 5.4)'
        ),
        "comments": "Core policy principle (skal/antar); implementation via EU chemicals/product rules and REACH.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 1: Polluter-pays principle—those responsible for pollution bear costs of preventing and "
            "reducing plastic pollution; measures must be placed early in the plastic product lifecycle."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Prinsippet om at forurensar betaler … dei som er ansvarlege for forureininga, må bere '
            'kostnadene ved tiltak for å hindre og redusere plastforureining" (kap. 1) / '
            'EN: "the principle that those responsible for pollution must bear the costs of measures to prevent '
            'and reduce plastic pollution" (Ch. 1)'
        ),
        "comments": "Established environmental policy principle underpinning EPR, taxes, and clean-up liability.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 1: Knowledge forum under Miljødirektoratet—contact forum for cooperation between plastic "
            "industry, importers, distributors, and Norwegian waste actors to strengthen knowledge on plastics "
            "in circulation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "etablere eit kontaktforum for samarbeid mellom aktørar i plastnæringa … under Miljødirektoratet" '
            '(kap. 1) / '
            'EN: "establish a contact forum for cooperation between actors in plastic industries … under the '
            'auspices of the Norwegian Environment Agency" (Ch. 1)'
        ),
        "comments": "Governance/coordination commitment (vil etablere); supports data and value-chain cooperation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3: Advocacy for a new legally binding global agreement on plastic pollution covering full "
            "lifecycle, land- and sea-based sources, national action plans, reporting, knowledge sharing, and "
            "coordination with existing agreements."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Noreg arbeider for ei ny globalt bindande avtale mot marin plastforsøpling og plastforureining"; '
            '"alle land … plikt til å iverksetje tiltak for både landbaserte og sjøbaserte kjelder" (kap. 3.1) / '
            'EN: "Norway is working for a new legally binding global agreement against marine plastic litter and '
            'plastic pollution"; "obligations for all countries to implement measures for both land-based and '
            'sea-based sources" (Ch. 3.1)'
        ),
        "comments": "International diplomacy instrument (arbeider for); led to UNEA 5.2 mandate in 2022.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3.2: Development Programme to Combat Marine Litter and Microplastics—NOK 1.6 billion "
            "2019–2024 for SDG 14.1 and UNEP zero vision; funds waste management, research and capacity "
            "building mainly in Asia, Africa and small island states."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Utviklingsprogrammet mot marin forsøpling og mikroplast … 1,6 milliardar kroner i perioden '
            '2019–2024" (kap. 3.2) / '
            'EN: "Development Programme to Combat Marine Litter and Microplastics … NOK 1.6 billion in the '
            'period 2019–2024" (Ch. 3.2); NOK 238 million granted to 32 projects in 2020'
        ),
        "comments": "Funded international development programme; operative budget allocation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3: Nordic initiative 2021–2024 'The Nordic Countries as a driving force'—three-year cooperation "
            "on knowledge, clean-up, EU/OSPAR/UNEA engagement, and leading call for global agreement; led by Norway."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Dei nordiske landa som pådrivar i arbeidet mot marin forsøpling og plastforureining" '
            '(2021–2024); "initiativet er leidd av Noreg" (kap. 3.2) / '
            'EN: "The Nordic Countries as a driving force in the work against marine litter and plastic pollution" '
            '(2021–2024); "the initiative is led by Norway" (Ch. 3.2)'
        ),
        "comments": "Regional cooperation platform; supports acute pollution, pellets, and clean-up projects.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 5.5/5.6: Ban on oxo-degradable plastic products in Norway from 3 July 2021—prevents "
            "microplastic formation from additives that break plastics into small pieces."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Alle okso-nedbrytande plastprodukt er forbode i Noreg frå 3. juli 2021" (kap. 5.5); '
            'Produktforskriften § 2b-3 / '
            'EN: "All oxo-degradable plastic products are banned in Norway as of 3rd July 2021" (Ch. 5.5); '
            'Norwegian Product Regulations § 2b-3'
        ),
        "comments": "Binding national/EØS product ban via Produktforskriften; directly reduces microplastic risk.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Ch. 5.6: SUP Directive implementation—ban on certain single-use plastic products (cotton swabs, "
            "straws, cutlery, etc.) and labelling for sanitary products, tobacco filters and wet wipes from "
            "3 July 2021 under Norwegian Product Regulations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Visse engangsplastprodukt … forbode frå 3. juli 2021"; "merkeplikt … at bind, tampongar, '
            'tobakksprodukt med filter og våtserviettar inneheld plast" (kap. 5.6); Produktforskriften § 2b-3 / '
            'EN: "certain single-use plastic products will be banned … from 3rd July 2021"; "labelling requirements '
            '… sanitary towels, tampons, tobacco products with filters and wet wipes all contain plastic" (Ch. 5.6)'
        ),
        "comments": "Operative EØS ban and labelling; estimated 1.9 billion fewer SUP items/year in Norway.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 5.6/6: Plastic packaging recycling targets—50 % by 2025 and 55 % by 2030 under EU Packaging "
            "Directive; Norway proposes minimum requirements in revised packaging regulations and packaging register."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "50 % av emballasje … vinnast att i 2025 og 55 % i 2030"; "Miljødirektoratet har starta '
            'etableringa av eit emballasjeregister" (kap. 5.6/6) / '
            'EN: "50 % of packaging recycled by 2025 and 55 % by 2030"; "The Norwegian Environment Agency have '
            'started establishing a packaging register" (Ch. 5.6/6)'
        ),
        "comments": "Binding EEA waste targets; 33 % plastic packaging recycled in 2019 vs. rising targets.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 5.8/9: Extended producer responsibility review—Miljødirektoratet to examine EPR for SUP litter "
            "(fast-food containers, wet wipes, tobacco filters) and fishing gear/aquaculture equipment; fishing gear "
            "EPR required by 1 January 2025."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "utvida produsentansvar … for følgjande tre produktgrupper" (kap. 5.8); "Frå 1. januar 2025 skal '
            'utvida produsentansvar for fiskeutstyr … vere på plass" (kap. 9) / '
            'EN: "extended producer responsibility for the following three product groups" (Ch. 5.8); '
            '"By 1st January 2025, extended producer responsibility schemes for fishing gear … should be in place" '
            "(Ch. 9)"
        ),
        "comments": "SUP litter EPR under review 2021; fishing gear EPR deadline 2025 per EU SUP Directive.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Ch. 5.10: Norwegian 'plastic pact'—Ministry will encourage private sector voluntary agreement to reduce "
            "consumption and environmental impacts of certain plastic products; builds on 2020 industry working group "
            "report and Plastløftet initiative."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "oppmuntre næringslivet til å utvikle ein norsk «plastpakt»"; "redusere forbruket av visse '
            'engangsplastprodukt" (kap. 5.10) / '
            'EN: "encourage the private sector to develop a Norwegian «plastic pact»"; "reducing the consumption '
            'of certain single-use plastic products" (Ch. 5.10)'
        ),
        "comments": "Voluntary industry commitment (vil oppmuntre); Plastløftet already reports to Grønt Punkt Norge.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 6: Mandatory municipal sorting of plastic waste—including proposal requiring municipalities to sort "
            "at least 70 % of household plastic waste by 2035; part of broader municipal waste recycling targets "
            "(55 %/60 %/65 % by 2025/2030/2035)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "kommunane … sortere minst 70 % av plastavfall frå hushald innan 2035"; '
            '"55 % i 2025, 60 % i 2030 og 65 % i 2035" (kommunalt avfall, kap. 6) / '
            'EN: "municipalities to sort at least 70 % of plastic waste from households by 2035"; '
            '"55 % in 2025, 60 % in 2030, and 65 % in 2035" (municipal waste, Ch. 6)'
        ),
        "comments": "Proposed Avfallsforskriften amendment under public consultation 2021; EEA Waste Framework targets.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 6.3: Basel Convention stricter plastic waste trade rules (Norwegian proposal adopted 2019, "
            "effective 1 January 2021)—prior informed consent for mixed/unsorted exports; Norway follows EU ban "
            "on export to non-OECD countries."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Noreg har implementert desse reglane … EU har innført forbod mot eksport til land utanfor OECD" '
            '(kap. 6.3) / '
            'EN: "Norway has implemented these rules … The EU has introduced a ban on export to non-OECD countries"; '
            '"new rules came into effect on 1st January 2021" (Ch. 6.3)'
        ),
        "comments": "Binding international/EU waste shipment rules; makes sorted recycling exports more profitable.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 6.3/5.7: Measures against illegal e-waste export—six St.meld. 19 (2019–2020) measures including "
            "clarified criminal liability for waste removal from return schemes and strengthened customs inspection."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "spesifisere regelverket … at fjerning av avfall levert i returordningar … er ulovleg"; '
            '"styrkje kontrolltiltak for å hindre ulovleg eksport av EE-avfall" (kap. 6.3) / '
            'EN: "specify regulations … that removal of waste delivered in return schemes … is an illegal or '
            'criminal act"; "strengthen inspection measures to prevent the illegal export of e-waste" (Ch. 6.3)'
        ),
        "comments": "E-waste contains substantial plastics; measures proposed in Meld. St. 19 environmental crime.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 7: Agricultural plastics sorting and recycling—new Avfallsforskriften chapter requiring farmers to "
            "sort/return agricultural plastics (silage wrap, fertiliser sacks, nets); builds on Grønt Punkt voluntary "
            "scheme (~86 % collection rate)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "nye krav til sortering og retur av landbruksplast for gjenvinning" (kap. 7); '
            '"13 875 tonn landbruksplast … 86 % innsamlingsgrad" (kap. 7) / '
            'EN: "new requirements for sorting and recycling of agricultural plastics" (Ch. 7); '
            '"13,875 tonnes of agricultural plastic … approximately 86 % collection rate" (Ch. 7)'
        ),
        "comments": "Proposed mandatory sorting regulation; voluntary Grønt Punkt scheme already operational.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Ch. 8: Transport-sector microplastics—tyre wear largest land-based source (~40 % of emissions); "
            "measures include road cleaning, catch pits for runoff, EU tyre labelling for abrasion when test method "
            "available, and strengthened preventive measures by transport operating companies."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Slitasje frå bildekk … største kjelde til spreiing av mikroplast"; "vegrensning, oppsamling av '
            'overvann i sandfang" (kap. 8) / '
            'EN: "Wear and tear from car tyres is considered the major source of the spread of microplastics"; '
            '"road cleaning, capturing runoff water in catch pits" (Ch. 8); ~40 % of land-based microplastic emissions'
        ),
        "comments": "Infrastructure and EU product-design measures; national road maintenance where appropriate.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 9: Fisheries ghost gear framework—Marine Resources Act prohibits dumping/abandoning gear at sea; "
            "fishers must attempt retrieval and report losses to Kystvakten; BarentsWatch real-time gear placement "
            "portal; annual Fiskeridirektoratet seabed clean-ups."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "ulovleg å dumpe eller forlate fiskeutstyr på havet"; "plikt til å forsøke å hente opp tapt utstyr … '
            'melde til Kystvakten" (kap. 9); Fritidsfiskeappen for fritidsfiske / '
            'EN: "illegal to dump or abandon fishing gear at sea"; "obligated to attempt to retrieve … report to the '
            'Norwegian Coast Guard" (Ch. 9); recreational Fritidsfiske app'
        ),
        "comments": "Binding fisheries law + MARPOL Annex V; Fiskeridirektoratet action plan 2021–2026.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 9: Fishing gear EPR from 1 January 2025—producers/importers of plastic commercial fishing and "
            "aquaculture gear responsible for lifecycle costs; Miljødirektoratet assessing scheme design."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "utvida produsentansvar for fiskeutstyr som inneheld plast … frå 1. januar 2025" (kap. 9) / '
            'EN: "extended producer responsibility schemes for fishing gear that contain plastic … should be in place '
            'by 1st January 2025" (Ch. 9)'
        ),
        "comments": "EU SUP Directive Art. 8 obligation; scheme design under Miljødirektoratet review 2021.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 9: Fishing for Litter (FFL) scheme—fishing vessels collect marine litter during normal operations "
            "with port reception and disposal funded; expanded from pilot 2016; linked to GloLitter and GGGI."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Fishing for Litter-ordninga … fiskarar samlar inn marin forsøpling under normale fiskeoperasjonar" '
            '(kap. 9) / '
            'EN: "The Fishing for Litter (FFL) scheme … fishing vessels collect marine litter during normal fishing '
            'operations" (Ch. 9); Norway joined Global Ghost Gear Initiative (GGGI) 2019'
        ),
        "comments": "Operative clean-up programme; Norway contributes ~NOK 40 million to GloLitter Partnerships.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Ch. 10.1: Artificial turf pitch microplastic reduction—Forurensningsforskriften Kap. 23A requirements "
            "for design/operation of pitches with loose infill; in force 1 July 2021; up to 90 % emission reduction "
            "when fully implemented."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "krav til utforming og drift av idrettsbaner som bruker løst fyllmateriale av mikroplast"; '
            '"redusere utslepp … med opptil 90 %" (kap. 10.1); Forurensningsforskriften kap. 23A / '
            'EN: "requirements for the design and operation of sports pitches that use loose microplastic as infill '
            'materials"; "microplastic emissions … reduced by up to 90 %" (Ch. 10.1); entered force 1 July 2021'
        ),
        "comments": "Binding pollution regulation; ~1,700 pitches; second-largest land-based microplastic source.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Ch. 10: National microplastic assessment—~19,000 tonnes/year land-based emissions (3.5 kg/person); "
            "prioritises tyre abrasion, artificial turf, paint, textiles, intentional microplastics, pellets; Norway "
            "active in EU REACH restriction on intentional microplastics."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "årleg utslepp av mikroplast frå landbaserte kjelder … 19 000 tonn"; "følgje opp EU/EØS-regelverk '
            'for tilsatt mikroplast" (kap. 10) / '
            'EN: "annual emission of microplastics from land-based sources … 19,000 tonnes"; '
            '"regulate the use of intentionally added microplastics … in accordance with forthcoming EU/EEA regulations" '
            "(Ch. 10)"
        ),
        "comments": "Knowledge base + EU chemicals restriction pathway; Mepex/NIVA 2020–2021 estimates.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Ch. 10.4: Operation Clean Sweep and OSPAR pellet guidelines—industry initiative to prevent loss of "
            "plastic pellets/flakes from production, transport and recycling; Norway to follow OSPAR and EU pellet "
            "measures under Nordic 2021–2024 initiative."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "vidare samarbeid om tiltak for å redusere utslepp av plastpellets" (kap. 10.4); '
            'Operation Clean Sweep / '
            'EN: "further cooperation on measures to reduce the emissions of plastic pellets" (Ch. 10.4); '
            '"Operation Clean Sweep … avoid and limit the emission of pellets from raw material production plants" '
            "(Ch. 10.4)"
        ),
        "comments": "Industry voluntary standard + OSPAR guidance; Trans Carrier spill (13.2 t) raised awareness.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 2/11: Pollution Control Act framework for plastic litter and pollution—discarded plastic in nature "
            "can be both littering and pollution; acute pollution chapter requires emergency response systems for "
            "sudden discharges including plastic pellets (Trans Carrier 2020 case)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forsøpling og forureining … kan vere både forsøpling og forureining"; '
            '"Den som er ansvarleg for akutt forureining … plikt til å iverksetje tiltak" (kap. 11); '
            'Forurensningsloven kap. 7 / '
            'EN: "discarded plastic in nature can be both littering and pollution"; '
            '"Those responsible for causing acute pollution are obligated to implement measures" (Ch. 11)'
        ),
        "comments": "Existing binding statute; strategy applies to plastic pellets, spills, and municipal/state response.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 12: National Centre for Marine Litter—from 1 January 2022 SOMM becomes dedicated centre under "
            "Klima- og miljødepartementet; coordinates clean-up, Rydde/Rent Hav digital tools, regional support, "
            "and volunteer facilitation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Frå 1. januar 2022 … eit eige nasjonalt senter for marin forsøpling"; '
            '"vidareutvikle og bruke digitale verktøy Rydde og Rent hav" (kap. 12) / '
            'EN: "From 1 January 2022 … a dedicated National Centre for Marine Litter"; '
            '"continue to further develop and to use the digital tools Rydde and Rent Hav" (Ch. 12)'
        ),
        "comments": "Institutional governance reform; SOMM subordinate to Klima- og miljødepartementet from 2022.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 12: Clean-up funding and polluter-pays—Miljødirektoratet marine litter grants, Norwegian Retailers' "
            "Environment Fund (NOK 105 million to 123 projects in 2020), wreck deposit scheme for recreational "
            "craft (NOK 1,000/boat), and volunteer coordination via Keep Norway Beautiful."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet si tilskotsordning for marin forsøpling"; "Handelens Miljøfond … 105 millionar '
            'kroner til 123 prosjekt" (kap. 12); "vrakpant … 1 000 kroner per båt" (kap. 12) / '
            'EN: "Norwegian Environment Agency grant schemes on marine litter"; "Norwegian Retailers\' Environment Fund '
            'distributed NOK 105 million to 123 projects" (Ch. 12); "wreck deposit (NOK 1,000 per boat)" (Ch. 12)'
        ),
        "comments": "Multiple funded schemes; clean-up prioritised near source per UNEA principles.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 13.2: Dsolve R&D centre—research-driven innovation on biodegradable fishing gear and aquaculture "
            "materials to reduce ghost fishing and microplastics; led by UiT with Norges forskningsråd support."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Dsolve … senter for forskingsdriven innovasjon … biologisk nedbrytbart fiskeutstyr" (kap. 13.2) / '
            'EN: "Dsolve is a centre for research driven innovation, which studies biodegradable plastics for fishing '
            'gear"; "put Norway at the forefront of research" on sea-based microplastic sources (Ch. 13.2)'
        ),
        "comments": "R&D instrument; addresses dolly rope, escape vents, and biodegradable gear alternatives.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 13.2: Green Platform Initiative—NOK 1 billion over three years (NOK 333 million in 2021) for "
            "coordinated green industrial transformation including plastics innovation projects for industry and "
            "research institutes."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Grønn plattform … 1 milliard kroner over tre år"; "333 millionar kroner … i 2021" (kap. 13.2) / '
            'EN: "The Green Platform Initiative … NOK 1 billion over three years"; '
            '"In 2021, NOK 333 million will be granted to the Initiative" (Ch. 13.2)'
        ),
        "comments": "National innovation funding; complements development programme and sector R&D (FuturePack, etc.).",
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
    ws.title = "4P Index - Plaststrategi"

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

    output_path = "/workspace/4P_Index_Noregs_Plaststrategi.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
