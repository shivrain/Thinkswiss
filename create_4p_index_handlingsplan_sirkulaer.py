#!/usr/bin/env python3
"""Generate 4P Index Excel for Handlingsplan for en sirkulær økonomi 2024–2025."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Handlingsplan for en sirkulær økonomi 2024–2025 "
        "(Action Plan for a Circular Economy 2024–2025)"
    ),
    "policy_url": (
        "https://www.regjeringen.no/no/dokumenter/handlingsplan-for-en-sirkulaer-okonomi/id3029477/"
    ),
    "policy_year": 2024,
    "policy_objective": (
        "Government follow-up action plan (2024–2025) complementing Norway's 2021 national circular "
        "economy strategy. It sets concrete measures across seven priority value chains (batteries, vehicles, "
        "electronics, packaging, textiles, plastics, food/nutrients, construction) and coordinates EØS/EU "
        "regulatory implementation—including SUP EPR, fishing-gear EPR, REACH microplastics restriction, "
        "packaging regulation preparation, pellet-loss regulation, and global plastics treaty advocacy—with "
        "national instruments on procurement, waste, chemicals, finance, partnerships, and monitoring."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Norge skal være et foregangsland i utviklingen av en grønn, sirkulær økonomi"; '
        '"innføre nye produsentansvarsordninger for enkelte engangsprodukter av plast og for utstyr i plast '
        'fra fiskeri, akvakultur og fritidsfiske" (kap. 3 handlingspunkter); '
        '"redusere matsvinn med 50 prosent innen 2030" (kap. 3) / '
        'EN: "Norway shall be a pioneering country in the development of a green, circular economy"; '
        '"introduce new EPR schemes for certain single-use plastic products and for plastic equipment from '
        'fisheries, aquaculture and recreational fishing" (Ch. 3 action points); '
        '"reduce food waste by 50 per cent by 2030" (Ch. 3)'
    ),
    "policy_type": 0.20,
    "policy_type_justification": (
        "Government action plan (handlingsplan) adopted by Klima- og miljødepartementet and "
        "Nærings- og fiskeridepartementet, published 2024. Non-binding policy framework with concrete "
        "regulatory commitments and follow-up measures; operationalised through EØS/EU regulations "
        "(SUP Directive, REACH Annex XVII, emballasjeforordningen), national legislation "
        "(lov om bærekraftige produkter og verdikjeder), EPR schemes, procurement rules, and voluntary partnerships."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "all sectors; fisheries; aquaculture; packaging; plastics; waste management; municipalities; "
        "construction; textiles; electronics; food; retail; public procurement; finance"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, use/consumption, end-of-life, recycling, litter/pollution, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Regjeringen har besluttet å opprette en ekspertgruppe … rapport innen april 2025"; '
        '"2,5 milliarder kroner gjennom Grønn plattform (2020–2023)"; '
        '"Stortinget bevilget i 2024 et tilskudd til Standard Norge på fire millioner kroner" (kap. 2) / '
        'EN: "The Government has decided to establish an expert group … report by April 2025"; '
        '"NOK 2.5 billion through the Green Platform (2020–2023)"; '
        '"In 2024 the Storting allocated NOK 4 million to Standard Norway" (Ch. 2)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 1 Vision: Norway as pioneering country for green circular economy; resources used efficiently "
            "in toxic-free material cycles replacing virgin extraction; complements 2021 national CE strategy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Regjeringens visjon er at Norge skal være et foregangsland i utviklingen av en grønn, '
            'sirkulær økonomi"; "ressurser må brukes og ombrukes mer effektivt i giftfrie kretsløp" (kap. 1) / '
            'EN: "The Government\'s vision is for Norway to be a pioneering country in the development of a green, '
            'circular economy"; "resources must be used and reused more efficiently in toxic-free material cycles" '
            "(Ch. 1)"
        ),
        "comments": "Strategic framing; action plan integrates CE into climate, nature, and industrial policy.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 1/3: Seven priority value chains—including packaging and plastics—as focal areas for CE transition; "
            "EU strengthened product framework to be implemented through EØS cooperation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "syv prioriterte verdikjeder som blir viktige for å lykkes med overgangen til en sirkulær økonomi. '
            'Fra batterier og kjøretøy, til emballasje og tekstiler" (Forord); '
            '"forsterket rammeverk for bærekraftige produkter og verdikjeder" (kap. 3) / '
            'EN: "seven priority value chains … from batteries and vehicles to packaging and textiles" (Preface); '
            '"strengthened framework for sustainable products and value chains" (Ch. 3)'
        ),
        "comments": "Value-chain prioritisation; plastics explicitly among seven chains (kap. 3.1).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.1: New Sustainable Products and Value Chains Act—to enable effective EØS implementation of "
            "EU product framework; Government to present to Storting spring 2024."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Forslaget til en helhetlig ny lov om bærekraftige produkter og verdikjeder"; '
            '"Regjeringen vil fremme loven til behandling i Stortinget i 2024" (kap. 3.1) / '
            'EN: "proposal for a comprehensive new act on sustainable products and value chains"; '
            '"The Government will present the act to the Storting in 2024" (Ch. 3.1)'
        ),
        "comments": "Enabling national legislation (vil fremme); key for SUP, ecodesign, and packaging rules.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.1: Ecodesign Regulation—minimum sustainability standards for nearly all products; digital "
            "product passport requirements; Norway participates in delegated acts for priority areas."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "økodesignforordningen … nye minimumsstandard for produkter i Europa"; '
            '"krav om digitalt produktpass" (kap. 3.1, boks 2.2) / '
            'EN: "Ecodesign Regulation … new minimum standard for products in Europe"; '
            '"requirement for a digital product passport" (Ch. 3.1, Box 2.2)'
        ),
        "comments": "EU/EØS regulation under implementation; affects plastic product design indirectly.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.1 Plastics value chain: SUP Directive as central plastics regulation; packaging currently under "
            "packaging directive; new packaging regulation with plastic-specific requirements under negotiation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Et sentralt regelverk som regulerer plast er direktivet om plastprodukter"; '
            '"Plastemballasje er i dag regulert av direktivet om emballasje og emballasjeavfall" (kap. 3.1) / '
            'EN: "A key regulatory framework for plastics is the Single-Use Plastics Directive"; '
            '"Plastic packaging is currently regulated by the Packaging and Packaging Waste Directive" (Ch. 3.1)'
        ),
        "comments": "Existing EØS framework; basis for SUP EPR and Produktforskriften Kap. 2b.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3.1/Handlingspunkter: SUP EPR—producers of certain single-use plastic products shall cover "
            "municipalities' costs for litter clean-up in public spaces."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Norge skal også innføre regler om produsentansvar for enkelte plastprodukter som innebærer at '
            'produsenter skal dekke kommunenes utgifter til opprydding av disse produktene i offentlige rom" '
            '(kap. 3.1); "Innføre nye produsentansvarsordninger for enkelte engangsprodukter av plast" '
            '(kap. 3 handlingspunkter) / '
            'EN: "Norway shall also introduce EPR rules for certain plastic products requiring producers to cover '
            'municipalities\' costs for cleaning up these products in public spaces" (Ch. 3.1); '
            '"Introduce new EPR schemes for certain single-use plastic products" (Ch. 3 action points)'
        ),
        "comments": "Concrete regulatory commitment (skal innføre); SUP litter-cost EPR.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 3.1/Handlingspunkter: Fishing-gear EPR—extended producer responsibility for equipment containing "
            "plastic used in fisheries, aquaculture, and recreational fishing."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Norge skal også innføre produsentansvar for utstyr som inneholder plast til fiskeri, akvakultur '
            'og fritidsfiske" (kap. 3.1); "for utstyr i plast fra fiskeri, akvakultur og fritidsfiske" '
            '(kap. 3 handlingspunkter) / '
            'EN: "Norway shall also introduce EPR for equipment containing plastic for fisheries, aquaculture and '
            'recreational fishing" (Ch. 3.1); "for plastic equipment from fisheries, aquaculture and recreational '
            'fishing" (Ch. 3 action points)'
        ),
        "comments": "Concrete regulatory commitment (skal innføre); addresses ghost gear and marine litter.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.1: REACH Annex XVII restriction on intentionally added microplastics—agreement reached; "
            "limits microplastic content in products placed on EEA market."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Det er oppnådd enighet om en restriksjon om tilsatt mikroplast under kjemikalieregelverket REACH" '
            '(kap. 3.1); "REACH-forordning vedlegg XVII om tilsatt mikroplast" (fig. 3.1) / '
            'EN: "Agreement has been reached on a restriction on intentionally added microplastics under the REACH '
            'chemicals regulation" (Ch. 3.1); "REACH Regulation Annex XVII on intentionally added microplastics" '
            "(Fig. 3.1)"
        ),
        "comments": "EU/EØS chemicals restriction; addresses microplastic pollution at source.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3.1: Proposed EU regulation on unintentional release of plastic pellets—prevents pellet losses "
            "along plastics supply chain."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Europakommisjonen har også foreslått en ny forordning om å forhindre utilsiktede utslipp av '
            'plastpellets" (kap. 3.1); "Forordning om utilsiktede utslipp av plastpellets" (fig. 3.1) / '
            'EN: "The European Commission has also proposed a new regulation to prevent unintentional release of '
            'plastic pellets" (Ch. 3.1); "Regulation on unintentional release of plastic pellets" (Fig. 3.1)'
        ),
        "comments": "EU proposal under negotiation; addresses pellet spill/leakage pollution.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 3.1 Packaging/Handlingspunkter: New Packaging and Packaging Waste Regulation—EU agreement March "
            "2024; Norway to strengthen packaging EPR for circular economy; preparation for EØS implementation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "EU kom til enighet om ny emballasjeforordning i mars 2024"; '
            '"Styrke produsentansvarsordningen for emballasje for å sikre at den er robust, effektiv og tilpasset '
            'en sirkulær økonomi" (kap. 3 handlingspunkter) / '
            'EN: "The EU reached agreement on a new packaging regulation in March 2024"; '
            '"Strengthen the packaging EPR scheme to ensure it is robust, effective and adapted to a circular '
            'economy" (Ch. 3 action points)'
        ),
        "comments": "Packaging regulation preparation; includes plastic packaging recycling targets.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 3.1 Packaging: Avfallsforskriften sorting requirements and Miljødirektoratet proposal to "
            "strengthen packaging EPR—separate collection of glass/metal, paper/cardboard, textiles under consultation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "nye krav om utsortering og materialgjenvinning i avfallsforskriften som ble vedtatt våren 2022"; '
            '"Miljødirektoratet har foreslått endringer i dagens produsentansvarsordning for emballasje" (kap. 3.1) / '
            'EN: "new sorting and recycling requirements in Avfallsforskriften adopted spring 2022"; '
            '"The Norwegian Environment Agency has proposed changes to the current packaging EPR scheme" (Ch. 3.1)'
        ),
        "comments": "National waste regulation; complements EU packaging rules.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Ch. 3.3/Box 3.3: Global plastics treaty advocacy—Norway leads High Ambition Coalition with Rwanda; "
            "promotes circular plastics economy, sustainable products, and increased collection/recycling."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Norge har tatt en ledende rolle … som leder av en høyambisjonskoalisjon av land sammen med Rwanda"; '
            '"overgang til sirkulær økonomi for plast" (boks 3.3) / '
            'EN: "Norway has taken a leading role … as leader of a High Ambition Coalition of countries together with '
            'Rwanda"; "transition to a circular economy for plastics" (Box 3.3)'
        ),
        "comments": "International diplomacy (arbeider for); handlingspunkt on ambitious global agreement.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.3: Chemicals policy for CE—same requirements for primary and secondary raw materials; "
            "hazardous substances removed from cycles for safe recycling."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "det må stilles samme krav til produkter produsert av primær og sekundær råvare"; '
            '"de farligste stoffene tas ut av kretsløpet" (kap. 3.3) / '
            'EN: "the same requirements must be set for products produced from primary and secondary raw materials"; '
            '"the most hazardous substances are removed from the cycles" (Ch. 3.3)'
        ),
        "comments": "Core CE chemicals principle; supports non-toxic plastic recycling.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 3.4: Transboundary waste regulation—export ban on plastic waste to non-OECD countries; "
            "stricter controls on waste exports outside EU."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "eksportforbud av plastavfall til ikke-OECD land"; '
            '"strengere kontroll på eksport av avfall til land utenfor EU" (kap. 3.4) / '
            'EN: "export ban on plastic waste to non-OECD countries"; '
            '"stricter controls on exports of waste to countries outside the EU" (Ch. 3.4)'
        ),
        "comments": "EU Waste Shipment Regulation revision; negotiated December 2023.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.2: EU taxonomy criteria for plastic packaging production—incorporated into EEA Agreement "
            "and Norwegian regulation February 2024; channels finance to circular activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kriteriene omfatter aktiviteter innenfor produksjon av plastemballasje"; '
            '"innlemmet i EØS-avtalen og gjennomført i norsk rett i forskrift i februar 2024" (kap. 3.2) / '
            'EN: "The criteria include activities within the production of plastic packaging"; '
            '"incorporated into the EEA Agreement and implemented in Norwegian law by regulation in February 2024" '
            "(Ch. 3.2)"
        ),
        "comments": "Sustainable finance instrument; affects plastic packaging investment classification.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 2.6: Standardisation for plastic packaging recycling chains—ISO circular economy standards "
            "2024; Standard Norway NOK 4 million allocation; prerequisite for recycled plastic in new products."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "EUs plaststrategi og forordning om emballasje … omfattende standardiseringsarbeid"; '
            '"verdikjede for plastprodukter og plastholdig emballasje egnet for materialgjenvinning" (kap. 2.6) / '
            'EN: "The EU plastics strategy and packaging regulation … extensive standardisation work"; '
            '"value chain for plastic products and plastic packaging suitable for material recycling" (Ch. 2.6)'
        ),
        "comments": "Technical enabler; supports packaging regulation implementation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 2.3: Public procurement—main rule from 1 January 2024 requiring climate/environment criteria "
            "(30% weighting or requirement specification) for all public procurements including circular solutions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "hovedregel om at det skal stilles klima- og miljøkrav ved alle offentlige anskaffelser"; '
            '"minimum 30 prosent" vekting (kap. 2.3, boks 2.1) / '
            'EN: "main rule that climate and environmental requirements must be set for all public procurements"; '
            '"minimum 30 per cent" weighting (Ch. 2.3, Box 2.1)'
        ),
        "comments": "Binding procurement regulation (skal); demand-side lever for circular plastics.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 2.2: Expert group on economic instruments—comprehensive study of regulatory, economic, and "
            "informational measures to promote circular economy; report due April 2025."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "ekspertgruppe som skal gjøre en helhetlig utredning av hvilke virkemidler … som er effektive '
            'for å fremme en mer sirkulær økonomi"; "rapport innen april 2025" (kap. 2.2) / '
            'EN: "expert group tasked with conducting a comprehensive assessment of which instruments are effective '
            'in promoting a more circular economy"; "report by April 2025" (Ch. 2.2)'
        ),
        "comments": "Policy development instrument (besluttet å opprette); may inform future plastic taxes/EPR.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Ch. 4.2 Plastpartnerskap: Voluntary partnership (KLD + NHO, Virke, Emballasjeforeningen et al.) to "
            "reduce consumption of single-use food containers and drink cups with lids made wholly or partly of plastic."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "plastpartnerskap om forbruksreduksjon av enkelte engangsprodukter av plast"; '
            '"reduksjon i forbruk av engangs matbeholdere for hurtigmat og ta-med-mat og drikkebegre med lokk" '
            '(kap. 4.2) / '
            'EN: "partnership on plastic … reduction in consumption of certain single-use plastic products"; '
            '"reduction in consumption of single-use food containers and take-away cups with lids" (Ch. 4.2)'
        ),
        "comments": "Voluntary industry partnership; complements SUP Directive; reporting on reductions.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 4.3 Statistics: Strengthened plastic statistics development—budget increases; official waste "
            "statistics insufficient; government to strengthen CE statistics for policy monitoring."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Bevilgningene til å utvikle plaststatistikk har økt"; '
            '"regjeringen vil styrke utviklingsarbeidet på statistikk for sirkulær økonomi" (kap. 4.3) / '
            'EN: "Funding to develop plastic statistics has increased"; '
            '"the Government will strengthen development work on statistics for the circular economy" (Ch. 4.3)'
        ),
        "comments": "Monitoring/enabling instrument (vil styrke); supports plastics policy evidence base.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 4.3 Monitoring: EU circular economy indicator framework—Miljødirektoratet recommendations; "
            "government to update Norwegian data by end 2024 and consider national indicators."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Oppdatere norske data til indikatorsettet til EU i løpet av 2024"; '
            '"vurdere nasjonale indikatorer på bakgrunn av EUs rammeverk" (kap. 4 handlingspunkter) / '
            'EN: "Update Norwegian data for the EU indicator set by the end of 2024"; '
            '"consider national indicators based on the EU framework" (Ch. 4 action points)'
        ),
        "comments": "Results monitoring; includes production/consumption and waste-handling indicators.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3 Handlingspunkter: Stricter regulations for sustainable and toxic-free products and value chains—"
            "implement EU product rules for level playing field with European industry."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Innføre strengere regelverk for bærekraftige og giftfrie produkter og verdikjeder for å bidra til '
            'en grønn sirkulær økonomi og like rammevilkår for norsk næringsliv som i hele Europa" '
            '(kap. 3 handlingspunkter) / '
            'EN: "Introduce stricter regulations for sustainable and toxic-free products and value chains to '
            'contribute to a green circular economy and equal competitive conditions for Norwegian industry as in '
            'the whole of Europe" (Ch. 3 action points)'
        ),
        "comments": "Overarching regulatory commitment; covers REACH, ecodesign, SUP, and packaging.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Ch. 3.4: Waste Framework Directive recycling targets—55%/60%/65% municipal waste recycling "
            "2025/2030/2035; government to propose further measures via Miljødirektoratet action analysis."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "mål om 55 prosent materialgjenvinning … i 2025, 60 prosent i 2030 og 65 prosent i 2035"; '
            '"Foreslå nye virkemidler for å øke andelen … som forberedes til ombruk eller materialgjenvinnes" '
            '(kap. 3.4) / '
            'EN: "targets of 55 per cent material recycling … in 2025, 60 per cent in 2030 and 65 per cent in 2035"; '
            '"Propose new instruments to increase the share … prepared for reuse or recycling" (Ch. 3.4)'
        ),
        "comments": "Binding EU targets; includes plastic fraction in municipal waste streams.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Ch. 3.1 Consumer information: Green transition consumer protection directive adopted Feb 2024; "
            "green claims and right-to-repair directives—enable informed choices and reduce greenwashing."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Direktivet om styrket forbrukervern i det grønne skiftet ble vedtatt 20. februar 2024"; '
            '"forhindre «grønnvasking» og førtidig produktsvikt" (kap. 3.1) / '
            'EN: "The directive on strengthening consumer protection in the green transition was adopted on '
            '20 February 2024"; "prevent greenwashing and premature product failure" (Ch. 3.1)'
        ),
        "comments": "Consumer-facing EU rules; supports sustainable product choices including packaging.",
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
    ws.title = "4P Index - Handlingsplan"

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

    output_path = "/workspace/4P_Index_Handlingsplan_Sirkulaer.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
