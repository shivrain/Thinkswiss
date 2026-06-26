#!/usr/bin/env python3
"""Generate 4P Index Excel for Nasjonal strategi for ein grøn, sirkulær økonomi (T-1573 N, 2021)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Nasjonal strategi for ein grøn, sirkulær økonomi (National Strategy for a Green, Circular Economy)"
    ),
    "policy_url": (
        "https://www.regjeringen.no/no/dokumenter/nasjonal-strategi-for-ein-gron-sirkular-okonomi/id2861253/"
    ),
    "policy_year": 2021,
    "policy_objective": (
        "Government national circular economy strategy (T-1573 N, published 16 June 2021) setting Norway's "
        "ambition to be a pioneering country for a green circular economy. It aligns national policy with the "
        "EU Circular Economy Action Plan (2020) across four pillars—sustainable production/design, sustainable "
        "consumption, toxic-free cycles, and value creation—and includes sector action points for bioeconomy, "
        "process industry, construction, and retail/services, with extensive plastics measures to reduce marine "
        "litter, strengthen recycling, and implement EU plastics strategy elements."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Omstillinga skal bidra til å redusere tap av naturressursar … og doble bruken av sekundære '
        'råmateriale i løpet av dei neste ti åra." (kap. 1.1); "65 prosent av all emballasje skal vinnast att '
        'som materiale i 2025 … 70 prosent i 2035" (kap. 3.3/4.3) / '
        'EN: "The transition shall contribute to reducing loss of natural resources … and double the use of '
        'secondary raw materials over the next ten years." (Ch. 1.1); "65 per cent of all packaging shall be '
        'recovered as material in 2025 … 70 per cent in 2035" (Ch. 3.3/4.3)'
    ),
    "policy_type": 0.20,
    "policy_type_justification": (
        "National government strategy (strategi, publikasjonskode T-1573 N) adopted by multiple ministries "
        "(Klima- og miljødepartementet et al.), published 16 June 2021. Non-binding policy framework coordinating "
        "EØS/EU circular economy implementation and national sector measures; operationalised through subsequent "
        "regulations (e.g. Avfallsforskriften, Produktforskriften Kap. 2b, emballasjeforskrifter)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "all sectors; bioeconomy; process industry; construction and buildings; retail and wholesale trade; "
        "waste management; municipalities; fisheries/aquaculture; packaging; plastics; textiles; electronics; "
        "food; finance; public procurement"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": "production, use/consumption, end-of-life, recycling, litter/pollution, environmental leakage",
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Regjeringa vil som del av ein større gjennomgang av skatte- og avgiftssystemet sjå på korleis '
        'riktigare miljøprising og andre økonomiske verkemiddel kan bidra til betre ressursutnytting og sirkulær '
        'økonomi" (kap. 15) / '
        'EN: "As part of a wider review of the tax system, the Government will consider how more appropriate '
        'environmental pricing and other economic instruments can promote better resource use and a circular economy" '
        "(Ch. 15)"
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Vision and overarching goal (kap. 1.1): Society where resources circulate efficiently in toxic-free "
            "loops replacing virgin resource extraction; transition to circular economy shall reach climate/environment "
            "targets, SDGs, value creation, and double use of secondary raw materials within ten years."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Eit samfunn der ressursar blir brukte og brukte om att på effektivt vis i giftfrie krinsløp der '
            'dei erstattar uttak av, og produksjon med, nye ressursar" (VISJON, kap. 1.1); "doble bruken av '
            'sekundære råmateriale i løpet av dei neste ti åra" (kap. 1.1) / '
            'EN: "A society where resources are used and re-used efficiently in toxic-free cycles that replace '
            'extraction of, and production with, new resources" (VISION, Ch. 1.1); "double the use of secondary '
            'raw materials over the next ten years" (Ch. 1.1)'
        ),
        "comments": "Strategic framing; not legally binding but sets government ambition aligned with EU CEAP.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Pioneering-country ambition (kap. 1.1, Forord): Norway shall be a leading country for green circular "
            "economy through national and EU policy development; strategy implements Granavolden platform goal."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Det er ambisjonen til regjeringa at Noreg skal vere eit føregangsland i utviklinga av ein grøn, '
            'sirkulær økonomi som utnyttar ressursane betre" (kap. 1.1); "Noreg skal vere eit føregangsland '
            'gjennom aktiv vidareutvikling av politikk og verkemiddel både nasjonalt og i samarbeidet med EU" '
            '(Forord) / '
            'EN: "It is the Government\'s ambition that Norway shall be a pioneering country in developing a green, '
            'circular economy that makes better use of resources" (Ch. 1.1); "Norway shall be a pioneering country '
            'through active further development of policy and instruments both nationally and in cooperation with the EU" '
            "(Foreword)"
        ),
        "comments": "High-level political commitment (vil/skal); shares EU ambition rather than separate national targets.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Four-pillar EU alignment (kap. 1, Samandrag): Strategy structures policy on (I) sustainable production/"
            "design, (II) sustainable consumption, (III) toxic-free cycles, (IV) circular economy and value creation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Strategien gjer greie for samanhengane i EUs og Noregs politikk på fire aspekt: I sirkulær '
            'økonomi gjennom berekraftig produksjon og produktdesign, II … berekraftige måtar å forbruke og bruke '
            'materiale, produkt og tenester, III … giftfrie sirkulære krinsløp, og IV sirkulær økonomi og verdiskaping" '
            '(Samandrag) / '
            'EN: "The strategy describes how EU and Norwegian policy fit together under four aspects: I circular '
            'economy through sustainable production and product design, II … sustainable ways to consume and use '
            'materials, products and services, III … toxic-free material cycles, and IV circular economy and value '
            'creation" (Summary)'
        ),
        "comments": "Integrates EU Circular Economy Action Plan (COM/2020/98) into national policy architecture.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Microplastics from imported products (kap. 1.1): EU measures on microplastics in hygiene products and "
            "washing machines, and more recyclable plastic packaging, can reduce microplastic spread in Norway."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "For eksempel kan eit EU-forbod mot å tilsetje mikroplast i hygieneprodukt eller krav om å samle '
            'opp mikroplast i vaskemaskiner redusere spreiinga av mikroplast i Noreg frå importerte produkt" (kap. 1.1) / '
            'EN: "For example, an EU ban on adding microplastics to hygiene products or requirements to capture '
            'microplastics in washing machines can reduce the spread of microplastics in Norway from imported products" '
            "(Ch. 1.1)"
        ),
        "comments": "Anticipated EEA/EU product measures; indirect domestic effect via imports.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 2.1: Strengthened product framework—ecodesign expansion and seven EU key value chains (electronics, "
            "batteries/vehicles, packaging, plastics, textiles, buildings, food/water/nutrients) with circular design "
            "criteria (durability, repairability, recycled content, absence of hazardous chemicals)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Sirkulær økonomi føreset betre produkt med eigenskapar som er tilpassa ein sirkulær økonomi" '
            'including "levetid", "potensial for ombruk", "innhald av attvunne materiale" (kap. 2.1); '
            '"elektronikk og IKT-utstyr, batteri og køyretøy, emballasje, plast, tekstilar, byggjevarer og mat, '
            'vatn og næringsstoff" (kap. 2.1) / '
            'EN: "A circular economy requires better products with characteristics adapted to a circular economy" '
            'including "lifetime", "potential for reuse", "content of recovered material" (Ch. 2.1); '
            '"electronics and ICT equipment, batteries and vehicles, packaging, plastics, textiles, building products '
            'and food, water and nutrients" (Ch. 2.1)'
        ),
        "comments": "Plastics explicitly listed as EU key value chain; 80% of environmental burden set at design phase.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 2.1 REGJERINGA VIL: Work for expanded EU ecodesign rules; products designed for durability, "
            "repair, upgrade, reuse and material recovery; sustainability requirements across whole value chains."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "arbeide for at regelverket til EU for økodesign blir utvida til fleire typeprodukt"; '
            '"arbeide for at produkt blir designa med formål om energieffektivitet, haldbarheit og moglegheit for '
            'reparasjon, oppgradering, ombruk og attvinning" (kap. 2.1) / '
            'EN: "work to ensure that EU ecodesign rules are expanded to more product types"; '
            '"work to ensure that products are designed for energy efficiency, durability and possibilities for '
            'repair, upgrade, reuse and material recovery" (Ch. 2.1)'
        ),
        "comments": "Policy commitment (vil arbeide for); implementation via EEA adoption of EU product legislation.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 2.2: Extended producer responsibility (EPR) review—all Norwegian EPR schemes to be reformed "
            "to promote circular value chains, including packaging, single-use articles and plastic fishing gear "
            "under revised EU waste framework."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "styrkje utvida produsentansvar i Noreg for å fremje høg grad av ressursutnytting av materiale i '
            'avfall og for å få fram produkt som er tilpassa ein sirkulær økonomi" (kap. 2.2); '
            '"på områda emballasje, eingongsartiklar og fiskeutstyr av plast" (kap. 2.2) / '
            'EN: "strengthen extended producer responsibility in Norway to promote high resource utilisation of '
            'materials in waste and to deliver products adapted to a circular economy" (Ch. 2.2); '
            '"in the areas of packaging, single-use articles and plastic fishing gear" (Ch. 2.2)'
        ),
        "comments": "EPR reform in progress 2021; directly covers plastic fishing gear and SUP-related products.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 3: Waste policy transformation—waste policy shall drive circular economy through preparation for "
            "reuse/recycling, secondary raw materials, and integrated product-waste policy; national target from 2021 "
            "prioritises material recovery over energy recovery."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Det nye målet … slår fast at materialattvinninga skal auke, og erstattar eit tidlegare mål der '
            'energiutnytting var likestilt med materialattvinning" (kap. 3.3); "materialattvinning av for eksempel '
            'plast gjer at vi har eit stort potensial til å redusere avfallsmengda" (kap. 3.1) / '
            'EN: "The new target … establishes that material recovery shall increase, replacing a previous target '
            'where energy recovery was on a par with material recovery" (Ch. 3.3); "material recovery of for example '
            'plastic means we have large potential to reduce waste volumes" (Ch. 3.1)'
        ),
        "comments": "National waste policy pivot; plastic recycling highlighted as underperforming (~40% packaging).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 3.3: Time-bound national/EØS recycling targets—household waste prep for reuse/recycling 65% by "
            "2035; all packaging material recovery 65% (2025) rising to 70% (2035); construction waste 70%; food "
            "waste reduction 15%/30%/50%."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Materialattvinning av all emballasje … 65 % … 70 % … 70 %" (tabell kap. 3.3); '
            '"Førebuing til ombruk og materialattvinning av hushaldsavfall … 65 % … 70 %" (kap. 3.3) / '
            'EN: "Material recovery of all packaging … 65% … 70% … 70%" (table Ch. 3.3); '
            '"Preparation for reuse and material recovery of household waste … 65% … 70%" (Ch. 3.3)'
        ),
        "comments": "EØS-incorporated EU waste targets; packaging/plastics central to achievement.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 3.4 REGJERINGA VIL: Introduce source-sorting of food waste and plastic waste from households and "
            "business; work toward 65% prep for reuse/recycling of household waste by 2035; support EU waste "
            "shipment regulation revision."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "allereie no innføre krav til kjeldesortering av matavfall og utsortering av plastavfall frå '
            'hushald og næringsliv" (kap. 3.4); "gradvis innføre verkemiddel som skal sikre 65 prosent førebuing '
            'til ombruk og materialattvinning av avfall i 2035" (kap. 3.4) / '
            'EN: "already now introduce requirements for source sorting of food waste and sorting out of plastic '
            'waste from households and business" (Ch. 3.4); "gradually introduce instruments to ensure 65 per cent '
            'preparation for reuse and material recovery of waste in 2035" (Ch. 3.4)'
        ),
        "comments": "Led to Avfallsforskriften consultation Jan 2021 on bio/plastic waste sorting.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 3.4: Strengthened efforts to reduce food waste and single-use plastic articles; advanced sorting "
            "technology for plastic waste quality; government supports EU rules incentivising digitalisation in waste."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "regjeringa retta ein særleg innsats mot å redusere matsvinnet og redusere bruken av '
            'eingongsartiklar av plast, og dette arbeidet skal no styrkjast ytterlegare" (kap. 3.4) / '
            'EN: "the Government directed a particular effort towards reducing food waste and reducing use of '
            'single-use plastic articles, and this work shall now be strengthened further" (Ch. 3.4)'
        ),
        "comments": "Cross-reference to Kap. 4.5 plastics and Kap. 4.7 food waste.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 4.3: Packaging policy—producer responsibility since 1995; deposit system for beverage packaging; "
            "2021 consultation on new packaging rules; Circular Packaging cluster; unified Scandinavian waste labelling."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Vi har hatt produsentansvar for emballasje sidan 1995, og det norske avgiftssystemet for '
            'drikkevareemballasje og pantesystemet har sørga for høg innsamling og attvinning"; '
            '"arbeide for EØS-krav som fremjar design for sirkulær økonomi og minimerer forsøpling, særleg for '
            'plastemballasje" (kap. 4.3) / '
            'EN: "We have had producer responsibility for packaging since 1995, and the Norwegian fee system for '
            'beverage packaging and the deposit system have ensured high collection and material recovery"; '
            '"work for EEA requirements promoting design for circular economy and minimising littering, especially '
            'for plastic packaging" (Ch. 4.3)'
        ),
        "comments": "Plastic packaging litter prevention explicitly prioritised in EU packaging directive revision.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 4.4: Textiles—microplastic spread through production, use and washing; government will support "
            "EU textile strategy including microplastic measures, separate collection by 2025, and value-chain cooperation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Tekstilar bidreg òg til å spreie mikroplast gjennom produksjon, bruk og vask" (kap. 4.4); '
            '"støtte opp om EUs arbeid med ein tekstilstrategi, med … tiltak mot mikroplast" (kap. 4.4) / '
            'EN: "Textiles also contribute to spreading microplastics through production, use and washing" (Ch. 4.4); '
            '"support the EU\'s work on a textile strategy, including … measures against microplastics" (Ch. 4.4)'
        ),
        "comments": "Microplastics pathway beyond primary plastics chapter.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 4.5.1: Plastics significance—marine plastic litter and microplastic spread as growing global problem; "
            "<10% of produced plastic reused in new products; ~3 million tonnes plastic in use in Norway, ~0.5 Mt "
            "plastic waste/year; ~240,000 t plastic packaging placed on market in 2019."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Marin plastforsøpling og spreiing av mikroplast er eit aukande globalt miljøproblem"; '
            '"Mindre enn 10 prosent av all plasten som er produsert, er utnytta i nye produkt" (kap. 4.5.1) / '
            'EN: "Marine plastic littering and spread of microplastics is a growing global environmental problem"; '
            '"Less than 10 per cent of all plastic that has been produced is used in new products" (Ch. 4.5.1)'
        ),
        "comments": "Problem diagnosis framing EU plastics strategy and national action.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 4.5.1: National plastic policy measures adopted/incoming—stricter Basel/EU rules on plastic waste "
            "trade from 1 Jan 2021; SUP ban from 3 July 2021 with marking requirements; consultation on bio/plastic "
            "waste sorting regulation from Jan 2021; plastic packaging EPR review."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Strengare reglar for eksport og import av plastavfall er i kraft frå 1. januar 2021"; '
            '"Forbod mot enkelte eingongsprodukt av plast vil tre i kraft frå 3. juli 2021"; '
            '"Forslag til forskrift med krav til … utsortering og materialattvinning av … plastavfall er sendt på '
            'høyring frå januar 2021" (kap. 4.5.1) / '
            'EN: "Stricter rules for export and import of plastic waste are in force from 1 January 2021"; '
            '"Ban on certain single-use plastic products will enter into force from 3 July 2021"; '
            '"Proposed regulation on requirements for … sorting and material recovery of … plastic waste was sent '
            'for consultation from January 2021" (Ch. 4.5.1)'
        ),
        "comments": "Links strategy to Produktforskriften Kap. 2b and forthcoming Avfallsforskriften amendments.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 4.5.2: Sustainable plastic products—focus on reducing unnecessary single-use/takeaway plastic "
            "packaging; new marking requirements beyond SUP ban; dialogue with social partners on voluntary measures; "
            "review of bio-based and biodegradable plastics."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Eit fokusområde for regjeringa er redusert ikkje-naudsynt forbruk av eingongsprodukt med '
            'forsøplingsrisiko, slik som plastemballasje for «takeaway» mat og drikke"; '
            '"I tillegg til forbodet mot visse eingongsprodukt av plast vil regjeringa innføre nye merkekrav" '
            '(kap. 4.5.2) / '
            'EN: "A focus area for the Government is reduced unnecessary consumption of single-use products with '
            'littering risk, such as plastic packaging for takeaway food and drink"; '
            '"In addition to the ban on certain single-use plastic products, the Government will introduce new '
            'marking requirements" (Ch. 4.5.2)'
        ),
        "comments": "Consumption-side litter prevention complementing SUP Directive implementation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 4.5.3: Plastic waste material recovery—EU targets 50% plastic packaging recovery by 2025 and 55% "
            "by 2030; Norway 36% plastic packaging recovered (2018); municipal source-sorting 70% of household plastic "
            "by 2035; agricultural plastic regulation consultation; land-based sorting technology."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "materialattvinning av plastemballasjeavfall må auke til 50 prosent innan 2025 og til 55 prosent '
            'innan 2030"; "kommunane skal kjeldesortere minimum 70 prosent av plastavfall frå hushalda innan 2035" '
            '(kap. 4.5.3) / '
            'EN: "material recovery of plastic packaging waste must increase to 50 per cent by 2025 and to 55 per cent '
            'by 2030"; "municipalities shall source-sort at least 70 per cent of plastic waste from households by 2035" '
            "(Ch. 4.5.3)"
        ),
        "comments": "Quantified recycling pathway; gap between current 36% and EU targets drives further instruments.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Kap. 4.5.4: Global plastic pollution treaty—Norway leadership since 2014 on marine litter/microplastics; "
            "Nordic ministers' 2019 call; Norway advocates binding global agreement with national plastic plans, "
            "sustainable products, reduced waste, higher plastic waste utilisation; UNEA 2022 mandate target."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Noreg har sidan 2014 teke ei leiarrolle for sterkare og meir forpliktande global innsats mot marin '
            'plastforsøpling og mikroplast i FN-regi"; "arbeide for at FNs miljøforsamling i 2022 fattar vedtak om '
            'forhandlingsmandat for ein internasjonal rettsleg bindande avtale mot plastforureining" (kap. 4.5.4) / '
            'EN: "Since 2014 Norway has taken a leading role in stronger and more binding global action against marine '
            'plastic littering and microplastics under the UN"; "work for UNEA in 2022 to adopt a negotiating mandate '
            'for an international legally binding agreement on plastic pollution" (Ch. 4.5.4)'
        ),
        "comments": "International diplomacy instrument; precursor to future UN plastics treaty (adopted 2022).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 4.5.4 REGJERINGA VIL: Require sustainability and circular-economy design in value chains where "
            "plastic is used; develop statistics/analysis base for plastic material recovery; pursue global treaty."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "arbeide for at det blir stilt krav til berekraft og design for sirkulær økonomi i verdikjeder for '
            'produkt der plast inngår"; "utvikle vidare det samla statistikk- og analysegrunnlaget for materialattvinning '
            'av plast" (kap. 4.5.4) / '
            'EN: "work to ensure requirements are set for sustainability and design for circular economy in value '
            'chains for products containing plastic"; "further develop the consolidated statistics and analysis base for '
            'material recovery of plastic" (Ch. 4.5.4)'
        ),
        "comments": "Consolidated plastics action commitments from dedicated strategy chapter.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 5: Stronger consumer rights and tools—right to repair, five-year Norwegian reclamation period "
            "for durable goods, sharing economy, anti-greenwashing, product information on environmental properties; "
            "consumer contact point for circular consumption."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Gode forbrukarrettar i form av lengre reklamasjonsfristar, som den norske femårsfristen for '
            'reklamasjon på varer som bilar og møbel"; "Regjeringa vil vurdere å opprette eit kontaktpunkt for '
            'informasjon og inspirasjon" (kap. 5.3) / '
            'EN: "Good consumer rights in the form of longer complaint periods, such as the Norwegian five-year period '
            'for complaints on goods such as cars and furniture"; "The Government will consider establishing a contact '
            'point for information and inspiration" (Ch. 5.3)'
        ),
        "comments": "Demand-side circular economy; extends product lifetimes reducing plastic/product waste.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Kap. 6: Circular public procurement—green public procurement as demand-side driver for circular products "
            "and solutions; EU mandatory green procurement criteria; government will promote sustainable public "
            "consumption and innovation through procurement."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "fremja berekraftig offentleg forbruk og grøn innovasjon gjennom offentlege anskaffingar" '
            '(Samandrag); styrking av grøne offentlege innkjøp under EUs handlingsplan (kap. 6) / '
            'EN: "promote sustainable public consumption and green innovation through public procurement" '
            '(Summary); strengthening of green public procurement under the EU action plan (Ch. 6)'
        ),
        "comments": "Public sector demand pull for recycled-content and circular plastic products.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 7: Toxic-free circular cycles—strengthened EU chemicals policy; hazardous substances must not "
            "enter recycled material loops; essential for safe increase of waste-based and secondary raw materials "
            "in economy including plastic recyclate."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "arbeider for giftfrie sirkulære krinsløp gjennom ein kontinuerleg ambisiøs kjemikaliepolitikk i '
            'samarbeid med EU" (Samandrag); "auka omløp av avfallsråstoff … skjer på berekraftige premissar" (kap. 1) / '
            'EN: "work towards toxic-free material cycles through a continuously ambitious chemicals policy in '
            'cooperation with the EU" (Summary); "increased circulation of waste-based raw materials … on sustainable '
            'premises" (Ch. 1)'
        ),
        "comments": "Chemicals-plastics nexus; additives in plastic hinder recycling (Kap. 4.5.2).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 1.4 / SINTEF analysis: Circular economy value creation potential—including plastic packaging "
            "scenarios (emballasjereduksjon, increased R&D) showing significant GDP and employment effects by 2030."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Verdiskaping og sysselsetjingseffektar i 2030 av ei omstilling til ein meir sirkulær økonomi" '
            'inkludert "Plastemballasje" med "Emballasjereduksjon, økt FOU" (kap. 1.4, SINTEF 2020) / '
            'EN: "Value creation and employment effects in 2030 of a transition to a more circular economy" '
            'including "plastic packaging" with "packaging reduction, increased R&D" (Ch. 1.4, SINTEF 2020)'
        ),
        "comments": "Economic rationale for plastics circularity investments.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Samandrag REGJERINGA VIL (consolidated): Use whole country for circular economy; strengthen recycling "
            "industry as secondary raw material supplier; exploit digitalisation for product info and secondary raw "
            "material markets; research/innovation cross-cutting priority; economic instruments review."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "styrkjer rolla avfallssektoren har som den som tek i vare materialressursar, og er leverandør av '
            'sekundære råmateriale til sirkulære krinsløp"; "utnyttar moglegheitene i digitalisering for tilgang til '
            'informasjon om produkteigenskapar og marknader for sekundære råvarer" (Samandrag) / '
            'EN: "enhance the role of the waste sector in managing material resources and as a supplier of secondary '
            'raw materials to circular cycles"; "use the potential of digitalisation for access to information on product '
            'properties and markets for secondary raw materials" (Summary)'
        ),
        "comments": "Top-level government commitment box; spans all four strategy pillars.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Bioeconomy and plastics (Samandrag kap. 9): Digital marketplace for reuse of resources, plastics and "
            "end-of-life equipment in bio-based sectors; common standards for biological resource reuse; regulatory "
            "changes needed for circular resource use in food chain."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "støtte kartlegging, dokumentasjon og utgreiing av aktuelle løysingar for ein digital marknadsplass for '
            'auka ombruk og meir lønnsam sirkulær utnytting av ressursar, plast og utrangert utstyr i bionæringane" '
            '(Samandrag) / '
            'EN: "support mapping, documentation and review of solutions for a digital marketplace to promote reuse '
            'and more profitable circular use of resources, plastics and end-of-life equipment in the bio-based sectors" '
            "(Summary)"
        ),
        "comments": "Cross-sector plastics reuse in agriculture/aquaculture contexts.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "Kap. 4.7 / wastewater: EU sewage directive revision to address pollutants of emerging concern including "
            "microplastics and pharmaceutical residues; upstream prevention needed to protect sludge reuse in cycles."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "teknologi for å redusere «pollutants of emerging concern», som inkluderer mellom anna legemiddelrestar, '
            'mikroplast og antibiotikaresistente bakteriar" (kap. 4.7.2) / '
            'EN: "technology to reduce pollutants of emerging concern, including among others pharmaceutical residues, '
            'microplastics and antibiotic-resistant bacteria" (Ch. 4.7.2)'
        ),
        "comments": "Microplastic capture in wastewater treatment; links water sector to plastics strategy.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Kap. 14–17 cross-cutting: Municipalities as circular economy enablers; economic instruments and "
            "environmental pricing review; digitalisation policy (digital product passports, secondary raw material "
            "marketplaces); Green Platform research funding for circular economy."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Digitale produktinformasjon, for eksempel gjennom digitale produktpass, vil effektivisere eit styrkt '
            'produktrammeverk"; "satsar på forskingsbasert kunnskap og innovasjon" (kap. 17/Samandrag) / '
            'EN: "Digital product information, for example through digital product passports, will increase the '
            'effectiveness of a strengthened product policy framework"; "invest in research-based knowledge and '
            'innovation" (Ch. 17/Summary)'
        ),
        "comments": "Enabling infrastructure for plastics traceability and circular markets.",
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
    ws.title = "4P Index - CE Strategy"

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

    output_path = "/workspace/4P_Index_Sirkulaerokonomistrategi.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
