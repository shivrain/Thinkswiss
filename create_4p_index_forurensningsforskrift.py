#!/usr/bin/env python3
"""Generate 4P Index Excel for Forurensningsforskriften (Pollution Regulation)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forskrift om begrensning av forurensning (Forurensningsforskriften) "
        "(Regulation on limitation of pollution – Pollution Regulation)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-931",
    "policy_year": 2026,
    "policy_objective": (
        "Principal executive pollution regulation under forurensningsloven, setting detailed rules to limit "
        "pollution to air, water and soil—including ship waste reception (MARPOL Annex V garbage/plastics), "
        "artificial-turf microplastic controls, wastewater discharge standards, and prohibitions on at-sea "
        "burning and dumping—thereby preventing plastic and other waste from entering the marine environment."
    ),
    "policy_target": 0.5,
    "policy_target_text": (
        '"Mottaksordningene skal legge til rette for at avfall fra skip blir samlet inn og utsortert på en måte '
        'som tilrettelegger for ombruk og materialgjenvinning." (§20-5); '
        '"Formålet med dette kapitlet er å hindre utslipp og spredning av plastholdig løst fyllmateriale …" '
        "(§23A-1)"
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) under forurensningsloven; principal technical pollution code since "
        "2004 with successive amendments. Major Chapter 20 overhaul October 2023 (EU Port Reception "
        "Facilities Directive 2019/883); Chapter 23A on artificial turf July 2021; EEA references updated "
        "February 2026 (FOR-2026-02-03-191)."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "industry, maritime, ports, municipalities, wastewater utilities, sports facilities, fisheries, "
        "recreation, coastal management, waste management"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, end-of-life, litter/pollution, environmental leakage, recycling"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        '"Omkostningene forbundet med mottak og videre håndtering av avfall fra skip … skal dekkes ved '
        'innkreving av generelt avfallsgebyr fra de skip som anløper havnen." (§20-9); '
        '"Kommunale saksbehandlings- og kontrollgebyrer" for wastewater permits (§11-4)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§11-1: Purpose of Chapters 11–15B—to protect the environment from adverse effects of wastewater "
            "discharge, establishing the framework for municipal and industrial effluent controls that indirectly "
            "prevent microplastics and plastic-associated pollutants reaching water bodies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med kapittel 11 til 15B er å beskytte miljøet mot uheldige virkninger av utslipp av '
            'avløpsvann." (§11-1)'
        ),
        "comments": "Overarching wastewater chapter objective; operative through Ch. 14 discharge rules.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§14-5: Sewer network requirements—networks shall be dimensioned, built, operated and maintained "
            "using best available technology, with particular regard to preventing leaks and limiting recipient "
            "pollution from overflows; responsible party must maintain overview of all overflows and significant leaks."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Avløpsnettet skal … dimensjoneres, bygges, drives og vedlikeholdes … særlig med hensyn til … '
            'b. forebygging av lekkasjer og c. begrensning av forurensning av resipienten som følge av overløp." '
            "(§14-5)"
        ),
        "comments": "Indirect plastic-pollution prevention via sewer integrity; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§14-6–§14-8: Municipal wastewater treatment tiers—binding secondary treatment, phosphorus removal "
            "and (in sensitive areas) nitrogen removal for discharges from larger settlements to freshwater, "
            "estuaries and sea, reducing particulate and microplastic-associated loads to recipients."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Kommunalt avløpsvann med utslipp til følsomt område … skal gjennomgå fosforfjerning" and '
            '"sekundærrensing" (§14-6); "Kommunalt avløpsvann med utslipp til normalt område … skal '
            'gjennomgå fosforfjerning" (§14-7); "mindre følsomt område … skal gjennomgå sekundærrensing" '
            "(§14-8)"
        ),
        "comments": "EU Urban Wastewater Directive implementation; mandatory treatment thresholds.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-1: Purpose of Chapter 20—to protect the external environment by ensuring satisfactory port "
            "reception facilities for ship waste and delivery of ship waste to port reception arrangements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Dette kapitlets formål er å verne det ytre miljø ved å sikre etablering og drift av '
            'tilfredsstillende mottaksordninger for avfall fra skip, og å sørge for at avfall fra skip blir levert '
            'til mottaksordning i havn." (§20-1)'
        ),
        "comments": "Frames EU Port Reception Facilities Directive / MARPOL garbage (Annex V) regime.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-5: Port reception facilities—port authorities shall establish and operate reception arrangements "
            "for ship waste, facilitating collection and sorting for reuse and material recovery; arrangements must "
            "cover normal delivery needs without undue ship delay."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Havneansvarlig skal ut fra behovet for levering sørge for etablering og drift av mottaksordninger '
            'for avfall fra skip i havnen. Mottaksordningene skal legge til rette for at avfall fra skip blir samlet '
            'inn og utsortert på en måte som tilrettelegger for ombruk og materialgjenvinning." (§20-5)'
        ),
        "comments": "Core ship-waste infrastructure duty; includes MARPOL Annex V (garbage/plastics).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-6: Port waste plans—port authorities must prepare and implement waste plans (Annex I), "
            "approved by county governor for five years, published in Norwegian and English via SafeSeaNet Norway; "
            "municipalities maintain harbour overview."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Havneansvarlig plikter å utarbeide og gjennomføre en avfallsplan i samråd med berørte parter" '
            "(§20-6); plan shall include waste types/quantities received (Annex I §1(g)); "
            '"Avfallsplanen skal godkjennes av statsforvalteren" (§20-6)'
        ),
        "comments": "Planning and transparency instrument; mandatory (plikter/skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-7: Advance waste notification—vessels ≥300 GT (with exceptions) bound for EEA ports must "
            "notify waste delivery via SafeSeaNet Norway at least 24 hours before arrival (or as soon as port known)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Fartøy med bruttotonnasje 300 eller mer … skal gi melding om levering av avfall i havnen a. minst '
            '24 timer før anløp" (§20-7); notification via SafeSeaNet Norway (§20-7)'
        ),
        "comments": "Enables port preparedness for garbage including plastics; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-8: Mandatory waste delivery—ship masters shall ensure waste is delivered to port reception "
            "before departure in accordance with MARPOL discharge norms; sufficient on-board storage may defer "
            "delivery only where next port has adequate facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Skipsfører skal sørge for at avfall leveres mottaksordning i havn før avgang, i samsvar med de '
            'relevante utslippsnormene fastsatt i MARPOL-konvensjonen." (§20-8)'
        ),
        "comments": "Direct anti-dumping rule for ship garbage; Sjøfartsdirektoratet kan kreve levering.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-8a: Waste receipts—port authorities must issue standardised waste receipts (Annex IV) to masters "
            "without undue delay; vessels ≥300 GT report receipt data in SafeSeaNet Norway and retain records "
            "on board for two years."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Når et skip har levert avfall i mottaksordning i havn, skal havneansvarlig … sørge for at en '
            'avfallskvittering … blir fylt ut" (§20-8a); vessels shall report receipt information in SafeSeaNet '
            "Norway (§20-8a)"
        ),
        "comments": "Traceability and enforcement documentation; added October 2023.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§20-9–§20-10: Port waste fees—costs of receiving and handling ship waste (except cargo residues "
            "and scrubber waste) funded by general port waste fee levied on all calling ships regardless of "
            "delivery; differentiated by ship category/size and hazardous waste; MARPOL Annex V surcharge "
            "for excess garbage volumes."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Omkostningene forbundet med mottak og videre håndtering av avfall fra skip … skal dekkes ved '
            'innkreving av generelt avfallsgebyr fra de skip som anløper havnen" (§20-9); '
            '"Tilleggsgebyr for avfall omfattet av MARPOL vedlegg V (søppel) kan dessuten innkreves" (§20-9)'
        ),
        "comments": "Polluter-pays / no-special-fee incentive structure for garbage delivery.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§20-14: Passive-fished-waste monitoring—EU Regulation 2022/92 (as amended by 2024/917) on "
            "collection and reporting of passively fished waste applies as regulation, implementing PRF Directive "
            "monitoring for marine litter including plastics retrieved by fishing gear."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Forordning (EU) 2022/92 … om … formatet for å rapportere passivt oppfisket avfall … gjelder som "
            "forskrift (§20-14); amended by forordning (EU) 2024/917 (FOR-2026-02-03-191)"
        ),
        "comments": "Marine litter data-collection instrument; EEA-incorporated EU regulation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§21-2: Prohibition on at-sea incineration—burning waste or other material on board ships and "
            "installations in Norwegian waters is prohibited, including Norwegian ships worldwide and within "
            "Norwegian EEZ/continental shelf."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Det er forbudt å forbrenne avfall eller annet materiale ombord på skip og innretninger til sjøs her '
            'i riket." (§21-2)'
        ),
        "comments": "Absolute prohibition (forbudt); prevents plastic waste destruction at sea.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§23-1: Purpose of Chapter 23—to prevent discharge of sewage and grey water causing "
            "environmental, hygienic and aesthetic problems in watercourses and coastal areas (complemented by "
            "separate ship environmental safety regulation)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Kapitlets formål er å hindre utslipp av kloakk og gråvann som bidrar til miljømessige, hygieniske '
            'og estetiske problemer i vassdrag og sjøområder." (§23-1)'
        ),
        "comments": "Municipalities may set stricter local sewage rules (§23-2).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§23A-1: Purpose of Chapter 23A—to prevent discharge and spread of plastic-containing loose infill "
            "material used on sports and recreation fields (artificial turf / kunstgress microplastic controls)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            '"Formålet med dette kapitlet er å hindre utslipp og spredning av plastholdig løst fyllmateriale som '
            'brukes på baner der det utøves idretts- eller fritidsaktiviteter." (§23A-1)'
        ),
        "comments": "Direct microplastic-leakage policy objective; in force since 1 July 2021.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§23A-4: Artificial-turf field design—outdoor fields require physical barriers (≥20 cm sealed height), "
            "drainage/stormwater capture for loose plastic infill, and user/maintenance controls; indoor fields "
            "require measures preventing infill entering wastewater and spreading beyond the field."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Utendørs idrettsbaner … skal ha a. en fysisk barriere … som hindrer at plastholdig løst '
            'fyllmateriale spres utenfor banen" and "b. løsninger for håndtering av drensvann og overvann som '
            'sikrer oppsamling" (§23A-4); responsible party shall ensure compliance (§23A-4)'
        ),
        "comments": "Core technical microplastic containment rules; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§23A-7: Loose infill handling—responsible party must prevent stored infill spreading beyond the "
            "field; collected infill shall be reused on the same field or delivered to lawful waste facility; snow "
            "cleared from fields deposited in designated snow areas with containment (§23A-6)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            '"Den ansvarlige … skal sørge for at oppsamlet plastholdig løst fyllmateriale ombrukes på den samme '
            'idrettsbanen eller leveres til lovlig avfallsanlegg." (§23A-7)'
        ),
        "comments": "End-of-life / circularity rule for artificial-turf granules; mandatory (skal).",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§23A-9: Substitution duty—operators of fields using plastic loose infill must assess lower-risk "
            "alternatives (including non-plastic infill) and choose them if feasible without unreasonable cost or "
            "inconvenience; assessments documented under §23A-8."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Den ansvarlige … skal vurdere om det finnes alternativ som medfører mindre risiko for '
            'miljøforstyrrelse … Den ansvarlige skal i så fall velge dette alternativet hvis det kan skje uten '
            'urimelig kostnad eller ulempe." (§23A-9)'
        ),
        "comments": "Chemicals-style substitution (skal vurdere/velge); transitional rules in §23A-12.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§23A-8: Knowledge and documentation—operators must document annual infill added/removed, "
            "handling routes, compliance measures and substitution assessments; records kept ≥5 years for "
            "inspection by pollution authorities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            '"Den ansvarlige … skal sørge for å ha kunnskap om og dokumentasjon på a. hvor mye plastholdig '
            'løst fyllmateriale som årlig er fylt på idrettsbanen" and "d. hvilke vurderinger og tiltak som er '
            'gjennomført for å overholde substitusjonsplikten i § 23A-9" (§23A-8)'
        ),
        "comments": "Monitoring and compliance documentation; statsforvalteren supervises (§23A-10).",
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
    ws.title = "4P Index - Forurensningsf"

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

    output_path = "/workspace/4P_Index_Forurensningsforskrift.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
