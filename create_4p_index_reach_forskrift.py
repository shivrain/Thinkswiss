#!/usr/bin/env python3
"""Generate 4P Index Excel for REACH-forskriften (REACH Regulation)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om registrering, vurdering, godkjenning og begrensning av kjemikalier "
        "(REACH-forskriften) (Regulation on registration, evaluation, authorisation and "
        "restriction of chemicals – REACH Regulation)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2008-05-30-516",
    "policy_year": 2026,
    "policy_objective": (
        "Norwegian EEA implementation of EU REACH (EC 1907/2006) governing registration, evaluation, "
        "authorisation and restriction of chemicals, with national provisions on scope, competent "
        "authorities, supervision and penalties. Includes restriction on intentionally added synthetic "
        "polymer microparticles (microplastics) via EU Regulation 2023/2055 (Annex XVII entry 78), "
        "effective in Norway from 17 March 2025, with phased market bans for cosmetics, detergents "
        "and other product categories."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Skal ikke bringes i omsetning som stoffer alene eller, der de syntetiske polymermikropartiklene '
        'er til stede for å gi en ettertraktet egenskap, i stoffblandinger i en konsentrasjon lik eller større '
        'enn 0,01 vektprosent." (REACH vedlegg XVII post 78 pkt. 1, via forordning (EU) 2023/2055) / '
        'EN: "Shall not be placed on the market as substances on their own or, where the synthetic polymer '
        'microparticles are present to confer a sought-after characteristic, in mixtures in a concentration '
        'equal to or greater than 0.01% by weight." (REACH Annex XVII entry 78 para. 1, via Regulation (EU) '
        "2023/2055)"
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under produktkontrolloven, forurensningsloven, arbeidsmiljøloven and "
        "brann- og eksplosjonsvernloven; principal Norwegian vehicle for EEA-incorporated REACH Regulation "
        "(EC) 1907/2006 and associated EU acts. Microplastics restriction added via forordning (EU) 2023/2055 "
        "incorporated 17 March 2025 (FOR-2025-03-17-487); last amended April 2026 (FOR-2026-04-08-559)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "cosmetics, detergents, chemicals, manufacturing, personal care, agriculture, horticulture, "
        "food additives, pharmaceuticals, import/export, retail"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, use/consumption, litter/pollution, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Ved overtredelse av denne forskrift eller vedtak fattet i medhold av den, kommer bestemmelsene '
        'om straff i hjemmelslovene til anvendelse" (§8) / '
        'EN: "In the event of a breach of this regulation or decisions made pursuant to it, the provisions on '
        'penalties in the enabling acts shall apply" (§8)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§1: REACH Regulation incorporation—Regulation (EC) No 1907/2006 as EEA Annex II Ch. XV No 12zc "
            "applies as Norwegian regulation with EEA adaptations and listed amendments, including "
            "forordning (EU) 2023/2055 on synthetic polymer microparticles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Forordning (EF) nr. 1907/2006 om registrering, vurdering og godkjenning av samt begrensninger '
            'for kjemikalier (REACH) … gjelder som forskrift" (§1); includes "forordning (EU) 2023/2055" (§1) / '
            'EN: "Regulation (EC) No 1907/2006 on the registration, evaluation, authorisation and restriction of '
            'chemicals (REACH) … applies as regulation" (§1); includes "Regulation (EU) 2023/2055" (§1)'
        ),
        "comments": "Umbrella EEA incorporation; microplastics amendment effective Norway 17 March 2025.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§3: Scope—regulation applies in Norway and to petroleum activities on the continental shelf; on "
            "Svalbard limited to SDS (Art. 31), restrictions (Art. 67/Annex XVII) and national §§4–9."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Forskriften gjelder i Norge og for petroleumsvirksomhet på norsk kontinentalsokkel." (§3) / '
            'EN: "The regulation applies in Norway and to petroleum activities on the Norwegian continental shelf." '
            "(§3)"
        ),
        "comments": "Territorial scope including offshore petroleum; Svalbard partial application.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4: Competent authority—Miljødirektoratet is responsible authority under REACH and makes decisions "
            "including authorisations; Arbeidstilsynet and Miljødirektoratet share responsibility for REACH Title IV."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet er ansvarlig myndighet etter REACH-forordningen og fatter vedtak etter '
            'forordningen, herunder godkjenninger" (§4) / '
            'EN: "The Norwegian Environment Agency is the responsible authority pursuant to the REACH Regulation '
            'and makes decisions pursuant to the Regulation, including authorisations" (§4)'
        ),
        "comments": "National authority allocation; enabling (er ansvarlig myndighet).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§6: Supervision—Miljødirektoratet, Arbeidstilsynet, Havindustritilsynet, DSBB and relevant ministries "
            "supervise compliance within their respective areas of authority."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet, Arbeidstilsynet, Havindustritilsynet, Direktoratet for samfunnssikkerhet og '
            'beredskap … fører tilsyn med at denne forskriften overholdes innenfor sine respektive '
            'myndighetsområder." (§6) / '
            'EN: "The Norwegian Environment Agency, the Labour Inspection Authority, the Norwegian Maritime '
            'Authority, the Directorate of Public Safety and Emergency Preparedness … supervise that this '
            'regulation is complied with within their respective areas of authority." (§6)'
        ),
        "comments": "Multi-agency enforcement framework.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "REACH Art. 6–7 (via §1): Registration—manufacturers and importers must register substances "
            "manufactured/imported ≥1 tonne/year in ECHA database before placing on market; downstream users "
            "must apply registration information in risk management."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Stoffer som fremstilles eller importeres i mengder på 1 tonn eller mer per år … skal registreres" '
            '(REACH art. 6) / '
            'EN: "Substances manufactured or imported in quantities of 1 tonne or more per year … shall be '
            "registered\" (REACH Art. 6)"
        ),
        "comments": "Core REACH registration duty; mandatory (skal registreres).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "REACH Art. 33 (via §1): SVHC communication—suppliers of articles containing Substances of Very High "
            "Concern >0.1% w/w must provide sufficient information to recipients and consumers on safe use."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Enhver leverandør av et produkt som inneholder et stoff … på kandidatlisten … i en '
            'konsentrasjon på mer enn 0,1 vektprosent … skal gi tilstrekkelig informasjon" (REACH art. 33 nr. 1) / '
            'EN: "Any supplier of an article containing a substance … on the candidate list … in a concentration '
            'above 0.1% weight by weight … shall provide sufficient information" (REACH Art. 33(1))'
        ),
        "comments": "Information duty along supply chain; includes plastic articles with SVHC additives.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "REACH Art. 55–60 (via §1): Authorisation—substances on Annex XIV require authorisation for specific "
            "uses after sunset date; aims to ensure risks from SVHCs are adequately controlled and substituted."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Stoffer som er oppført i vedlegg XIV skal ikke bringes i omsetning eller brukes etter datoen som '
            'er angitt i vedlegg XIV kolonne 2, med mindre en godkjenning er gitt" (REACH art. 56 nr. 1) / '
            'EN: "Substances listed in Annex XIV shall not be placed on the market or used after the date specified '
            'in column 2 of Annex XIV, unless an authorisation has been granted" (REACH Art. 56(1))'
        ),
        "comments": "Authorisation regime for highest-concern substances; mandatory prohibition with exceptions.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "REACH Art. 67–69 (via §1): Restriction procedure—manufacture, placing on market and use of substances "
            "may be restricted via Annex XVII when risks are not adequately controlled; basis for all Annex XVII "
            "entries including microplastics."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Stoffer, stoffblandinger og produkter kan være underlagt begrensninger i fremstilling, '
            'omsetning eller bruk … i vedlegg XVII" (REACH art. 67 nr. 1) / '
            'EN: "Substances, mixtures and articles may be subject to restrictions on manufacture, placing on the '
            'market or use … in Annex XVII" (REACH Art. 67(1))'
        ),
        "comments": "Legal basis for Annex XVII restrictions including entry 78 microplastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Annex XVII entry 78 para. 1 (via §1 / EU 2023/2055): Microplastics market ban—prohibition on placing "
            "on market of synthetic polymer microparticles on their own or in mixtures ≥0.01% w/w when "
            "intentionally added to confer a sought-after characteristic."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Skal ikke bringes i omsetning som stoffer alene eller, der de syntetiske polymermikropartiklene '
            'er til stede for å gi en ettertraktet egenskap, i stoffblandinger i en konsentrasjon lik eller større '
            'enn 0,01 vektprosent." (vedlegg XVII post 78 pkt. 1) / '
            'EN: "Shall not be placed on the market as substances on their own or, where the synthetic polymer '
            'microparticles are present to confer a sought-after characteristic, in mixtures in a concentration '
            'equal to or greater than 0.01% by weight." (Annex XVII entry 78 para. 1)'
        ),
        "comments": "Core microplastics restriction; applies Norway from 17 March 2025 (EEA).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Annex XVII entry 78 para. 6(b): Immediate microbeads ban—no transitional period for synthetic polymer "
            "microparticles used as abrasives (exfoliate, polish, clean) in rinse-off cosmetics or detergents."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Fra 17. oktober 2027 for … skyllerensprodukter … med mindre … de inneholder syntetiske '
            'polymermikropartikler for bruk som et skuremiddel, dvs. for å eksfoliere, polere eller rengjøre '
            '("mikroperler")" — mikroperler uten overgangsperiode (vedlegg XVII post 78 pkt. 6 bokstav b) / '
            'EN: "From 17 October 2027 for rinse-off products … unless … they contain synthetic polymer '
            'microparticles for use as an abrasive, i.e. namely to exfoliate, polish or clean ("microbeads")" — '
            'microbeads with no transitional period (Annex XVII entry 78 para. 6(b))'
        ),
        "comments": "Microbeads banned immediately from general restriction entry into force; no phase-in.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Annex XVII entry 78 para. 6: Phased market bans—transitional dates for cosmetics (rinse-off 2027, "
            "leave-on 2029, lip/nail/makeup 2035), detergents/waxes/polishes/air care 2028, fragrance "
            "encapsulation 2029, agricultural/horticultural products 2028, sports-surface infill 2031."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Fra 17. oktober 2027 for skyllerensprodukter … Fra 17. oktober 2028 for vaskemidler … Fra '
            '17. oktober 2029 for produkter til etterbehandling … Fra 17. oktober 2035 for leppeprodukter, '
            'negleprodukter og sminkeprodukter" (vedlegg XVII post 78 pkt. 6) / '
            'EN: "From 17 October 2027 for rinse-off products … From 17 October 2028 for detergents … From '
            '17 October 2029 for leave-on products … From 17 October 2035 for lip products, nail products, and '
            'make-up products" (Annex XVII entry 78 para. 6)'
        ),
        "comments": "Sector-specific phased bans; key sectors: cosmetics, detergents, manufacturing.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "Annex XVII entry 78 paras 4–5: Derogations—exemptions for industrial-site use, medicinal/veterinary "
            "products, EU fertilisers, food additives, IVD devices, food/feed; and for microparticles contained by "
            "technical means, permanently modified during use, or incorporated in solid matrix."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Punkt 1 gjelder ikke for … syntetiske polymermikropartikler … for bruk på industrielle steder" '
            '(vedlegg XVII post 78 pkt. 4 a); "syntetiske polymermikropartikler som er permanent innlemmet i en '
            'fast matrise under tiltenkt sluttbruk" (pkt. 5 c) / '
            'EN: "Paragraph 1 shall not apply to … synthetic polymer microparticles … for use at industrial sites" '
            '(Annex XVII entry 78 para. 4(a)); "synthetic polymer microparticles which are permanently incorporated '
            'into a solid matrix during intended end use" (para. 5(c))'
        ),
        "comments": "Defines scope boundaries; trapped glitter in solid matrix exempt.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Annex XVII entry 78 paras 7–8: Instructions for use and disposal—suppliers of industrial-site "
            "microparticles must provide use/disposal instructions from October 2025; professional and consumer "
            "products with derogated microparticles from October 2025–2026."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Fra 17. oktober 2025 skal leverandører … gi følgende informasjon: … instruksjoner for bruk og '
            'avhending som forklarer … hvordan utslipp … kan forhindres" (vedlegg XVII post 78 pkt. 7) / '
            'EN: "From 17 October 2025 suppliers … shall provide the following information: … instructions for use '
            'and disposal explaining … how to prevent releases of synthetic polymer microparticles to the '
            'environment" (Annex XVII entry 78 para. 7)'
        ),
        "comments": "Behaviour-change instrument for non-banned uses; mandatory (skal gi).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Annex XVII entry 78 para. 9: Consumer labelling—suppliers of lip, nail and make-up products containing "
            "microplastics must bear statement \"This product contains microplastics\" from October 2031 until "
            "October 2035."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Fra 17. oktober 2031 til 16. oktober 2035 skal leverandører … gi følgende utsagn: "Dette '
            'produktet inneholder mikroplast."" (vedlegg XVII post 78 pkt. 9) / '
            'EN: "From 17 October 2031 until 16 October 2035 suppliers … shall provide the following statement: '
            '"This product contains microplastics."" (Annex XVII entry 78 para. 9)'
        ),
        "comments": "Interim consumer-information requirement before 2035 cosmetics ban.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Annex XVII entry 78 paras 11–12: ECHA reporting—manufacturers and industrial users of pellet/flake/powder "
            "feedstock report uses and release estimates from 2026; other suppliers report end uses from 2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Fra og med 2026 skal produsenter og industrielle nedstrømsbrukere … sende følgende informasjon '
            'til byrået innen 31. mai hvert år" (vedlegg XVII post 78 pkt. 11) / '
            'EN: "Starting from 2026 manufacturers and industrial downstream users … shall submit the following '
            'information to the Agency by 31 May of each year" (Annex XVII entry 78 para. 11)'
        ),
        "comments": "Monitoring and transparency instrument; mandatory reporting (skal sende).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§8: Penalties—breaches of the regulation or decisions made under it are subject to penalties under "
            "the enabling acts unless stricter criminal provisions apply."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Ved overtredelse av denne forskrift eller vedtak fattet i medhold av den, kommer bestemmelsene '
            'om straff i hjemmelslovene til anvendelse" (§8) / '
            'EN: "In the event of a breach of this regulation or decisions made pursuant to it, the provisions on '
            'penalties in the enabling acts shall apply" (§8)'
        ),
        "comments": "Backstop enforcement via forurensningsloven and other parent acts.",
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
    ws.title = "4P Index - REACH"

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

    output_path = "/workspace/4P_Index_REACH_forskrift.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
