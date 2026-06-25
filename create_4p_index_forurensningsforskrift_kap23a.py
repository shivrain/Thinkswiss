#!/usr/bin/env python3
"""Generate 4P Index Excel for Forurensningsforskriften Kap. 23A (sports surfaces with plastic infill)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Forurensningsforskriften – Kap. 23A: Idrettsbaner med plastholdig fyllmateriale "
        "(Pollution Regulation – Ch. 23A: Sports surfaces using plastic-containing loose infill material)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-931/KAPITTEL_23A",
    "policy_year": 2021,
    "policy_objective": (
        "Chapter 23A of the Pollution Regulation (Forurensningsforskriften, FOR-2004-06-01-931), added by "
        "FOR-2021-04-07-1096, sets design and operation requirements for outdoor and indoor sports pitches using "
        "plastic-containing loose infill (e.g. artificial turf granulate). It mandates physical barriers, drainage "
        "and stormwater controls, user and maintenance measures, snow management, granulate handling and reuse, "
        "documentation, and substitution assessments to prevent release and spread of microplastics. In force "
        "1 July 2021."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Formålet med dette kapitlet er å hindre utslipp og spredning av plastholdig løst fyllmateriale som '
        'brukes på baner der det utøves idretts- eller fritidsaktiviteter." (§23A-1) / '
        'EN: "The purpose of this chapter is to prevent release and spread of plastic-containing loose infill '
        'material used on pitches where sports or leisure activities are practised." (§23A-1)'
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation chapter (forskrift) under forurensningsloven §9 and §33, inserted by "
        "FOR-2021-04-07-1096 into the consolidated Forurensningsforskriften (FOR-2004-06-01-931). "
        "Adopted 7 April 2021, in force 1 July 2021; supervised by county governors (statsforvalteren) under Kap. 41."
    ),
    "policy_integration": 1,
    "policy_sectors_list": "sports, municipalities, construction, leisure, facility management, waste management",
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "use/consumption, litter/pollution, environmental leakage, waste management",
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Kapitlet trer i kraft 1. juli 2021." (§23A-11); tilsyn via forurensningsforskriften kapittel 41 '
        '(§23A-10) / '
        'EN: "The chapter enters into force on 1 July 2021." (§23A-11); supervision via Pollution Regulation '
        "Chapter 41 (§23A-10)"
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-1: Purpose—prevent release and spread of plastic-containing loose infill used on sports and "
            "leisure activity pitches."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Formålet med dette kapitlet er å hindre utslipp og spredning av plastholdig løst fyllmateriale '
            'som brukes på baner der det utøves idretts- eller fritidsaktiviteter." (§23A-1) / '
            'EN: "The purpose of this chapter is to prevent release and spread of plastic-containing loose infill '
            'material used on pitches where sports or leisure activities are practised." (§23A-1)'
        ),
        "comments": "Frames microplastic leakage prevention objective for artificial turf and similar surfaces.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-2: Scope—chapter applies to sports pitches where plastic-containing loose infill is used."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Dette kapitlet gjelder for idrettsbaner der det brukes plastholdig løst fyllmateriale." (§23A-2) / '
            'EN: "This chapter applies to sports pitches where plastic-containing loose infill material is used." '
            "(§23A-2)"
        ),
        "comments": "Defines regulated facility type; covers outdoor and indoor pitches per §23A-3.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-2: Territorial exclusion—chapter does not apply to Svalbard."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Dette kapitlet gjelder ikke for Svalbard." (§23A-2) / '
            'EN: "This chapter does not apply to Svalbard." (§23A-2)'
        ),
        "comments": "Geographic carve-out (gjelder ikke).",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-3: Definitions—plastic loose infill (solid particulate, non-water-soluble, containing synthetic "
            "polymer material); sports pitches (outdoor and indoor); responsible person (forurensningsloven §7, "
            "owner and operator); reuse (non-waste re-use for same purpose)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Med plastholdig løst fyllmateriale menes faste partikler i kornet eller revet form, som ikke '
            'løses opp i vann, som inneholder syntetisk materiale med innhold av polymerer." (§23A-3) / '
            'EN: "Plastic-containing loose infill material means solid particles in granular or shredded form that '
            'do not dissolve in water and contain synthetic material with polymer content." (§23A-3)'
        ),
        "comments": "Legal scope definitions; links responsible person to forurensningsloven §7 polluter-pays duty.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-4(a): Outdoor pitches—mandatory physical barrier around pitch preventing spread of plastic "
            "loose infill; at least 20 cm of barrier height measured from ground must be tight."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "en fysisk barriere rundt idrettsbanen som hindrer at plastholdig løst fyllmateriale spres '
            'utenfor banen. Minst 20 centimeter av barrierens høyde, målt fra bakken, skal være tett" (§23A-4(a)) / '
            'EN: "a physical barrier around the sports pitch that prevents plastic-containing loose infill material '
            'from spreading outside the pitch. At least 20 centimetres of the barrier height, measured from the '
            'ground, must be tight" (§23A-4(a))'
        ),
        "comments": "Core design requirement for artificial turf containment; transitional delay for pre-2021 pitches (§23A-12).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-4(b): Outdoor pitches—drainage and stormwater solutions must collect loose plastic infill so "
            "it does not spread outside the pitch."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "løsninger for håndtering av drensvann og overvann som sikrer oppsamling av løst plastholdig '
            'fyllmateriale slik at dette ikke spres utenfor banen" (§23A-4(b)) / '
            'EN: "solutions for handling drainage water and stormwater that ensure collection of loose plastic-containing '
            'infill material so that it is not spread outside the pitch" (§23A-4(b))'
        ),
        "comments": "Stormwater pathway control; prevents granulate wash-off to environment.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-4(c): Outdoor pitches—measures preventing spread via pitch users, construction machinery and "
            "other equipment used for maintenance and snow clearing."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "tiltak som hindrer at plastholdig løst fyllmateriale spres utenfor banen via brukere av banen '
            'eller via anleggsmaskiner og annet utstyr som brukes ved vedlikehold og snørydding av banen" '
            '(§23A-4(c)) / '
            'EN: "measures that prevent plastic-containing loose infill material from spreading outside the pitch '
            'via users of the pitch or via construction machinery and other equipment used for maintenance and '
            'snow clearing of the pitch" (§23A-4(c))'
        ),
        "comments": "Operational controls during use and maintenance.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-4(a): Indoor pitches—solutions preventing plastic loose infill from entering wastewater."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "løsninger for å hindre at plastholdig løst fyllmateriale havner i avløpsvannet" (§23A-4(a), '
            'innendørs) / '
            'EN: "solutions to prevent plastic-containing loose infill material from entering wastewater" '
            "(§23A-4(a), indoor)"
        ),
        "comments": "Indoor microplastic discharge prevention to sewer systems.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-4(b): Indoor pitches—measures preventing spread outside pitch via users."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "tiltak som hindrer at plastholdig løst fyllmateriale spres utenfor banen via brukere av banen" '
            '(§23A-4(b), innendørs) / '
            'EN: "measures that prevent plastic-containing loose infill material from spreading outside the pitch '
            'via users of the pitch" (§23A-4(b), indoor)'
        ),
        "comments": "User behaviour controls for indoor facilities.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-4 final paragraph: Responsible person for the sports pitch must ensure compliance with "
            "outdoor and indoor design/operation requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige for idrettsbanen skal sørge for at kravene i denne bestemmelsen oppfylles." '
            '(§23A-4) / '
            'EN: "The person responsible for the sports pitch shall ensure that the requirements in this provision '
            'are fulfilled." (§23A-4)'
        ),
        "comments": "Overarching compliance duty on owner/operator under forurensningsloven §7.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-5: Information duty—responsible person must inform pitch users about spread of plastic loose "
            "infill and measures to reduce spread risk."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige for idrettsbaner der plastholdig løst fyllmateriale brukes skal sørge for å '
            'informere brukerne av idrettsbanen om spredning av plastholdig løst fyllmateriale og tiltak for å '
            'redusere risikoen for slik spredning." (§23A-5) / '
            'EN: "The person responsible for sports pitches where plastic-containing loose infill material is used '
            'shall ensure that users of the sports pitch are informed about spread of plastic-containing loose '
            'infill material and measures to reduce the risk of such spread." (§23A-5)'
        ),
        "comments": "Soft behavioural instrument; complements physical barriers.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-6: Snow management—cleared snow deposited in dedicated snow depot on or off pitch; measures "
            "to keep plastic infill within pitch or snow depot until handled per §23A-7."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige … skal sørge for at snø som ryddes av idrettsbanen, deponeres på eget område '
            'for snødeponering på eller utenfor banen. Den ansvarlige skal sørge for nødvendige tiltak for å '
            'sikre at plastholdig løst fyllmateriale forblir innenfor idrettsbanen eller snødeponiet frem til det '
            'håndteres i tråd med § 23A-7 annet ledd." (§23A-6) / '
            'EN: "The person responsible … shall ensure that snow cleared from the sports pitch is deposited in a '
            'dedicated snow depot area on or off the pitch. The person responsible shall ensure necessary measures '
            'so that plastic-containing loose infill material remains within the sports pitch or snow depot until '
            'handled in accordance with §23A-7 second paragraph." (§23A-6)'
        ),
        "comments": "Seasonal pathway control; relevant for Nordic climate artificial turf management.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Litter/Pollution",
        "instrument_description": (
            "§23A-7(1): Granulate storage—responsible person must take measures preventing stored plastic loose "
            "infill on or off pitch from spreading to surroundings."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige … skal sørge for nødvendige tiltak for å hindre at plastholdig løst fyllmateriale '
            'som oppbevares på eller utenfor banen spres til omgivelsene utenfor banen." (§23A-7(1)) / '
            'EN: "The person responsible … shall ensure necessary measures to prevent plastic-containing loose '
            'infill material stored on or off the pitch from spreading to the surroundings outside the pitch." '
            "(§23A-7(1))"
        ),
        "comments": "Storage and stockpile containment requirement.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§23A-7(2): Collected plastic loose infill must be reused on the same pitch (if suitable) or delivered "
            "to lawful waste facility."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige … skal sørge for at oppsamlet plastholdig løst fyllmateriale ombrukes på den '
            'samme idrettsbanen eller leveres til lovlig avfallsanlegg. Oppsamlet fyllmateriale kan kun ombrukes '
            'på den samme idrettsbanen hvis fyllmaterialet er egnet for ombruk." (§23A-7(2)) / '
            'EN: "The person responsible … shall ensure that collected plastic-containing loose infill material is '
            'reused on the same sports pitch or delivered to a lawful waste facility. Collected infill material may '
            'only be reused on the same sports pitch if the infill material is suitable for reuse." (§23A-7(2))'
        ),
        "comments": "Circular economy pathway; links to Avfallsforskriften lawful disposal.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-8: Knowledge and documentation duty—annual records of infill added and removed (quantities, "
            "composition, handling), compliance measures implemented, and substitution assessments under §23A-9; "
            "retain documentation at least 5 years and make available to pollution authority."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige … skal sørge for å ha kunnskap om og dokumentasjon på … hvor mye plastholdig løst '
            'fyllmateriale som årlig er fylt på idrettsbanen … hvor mye … er fjernet … hvilke tiltak som er '
            'gjennomført … hvilke vurderinger og tiltak som er gjennomført for å overholde substitusjonsplikten '
            'i § 23A-9. Dokumentasjonen skal tas vare på i minst 5 år og være tilgjengelig ved kontroll eller '
            'på forespørsel fra forurensningsmyndigheten." (§23A-8) / '
            'EN: "The person responsible … shall ensure knowledge of and documentation on … how much plastic-containing '
            'loose infill material is added to the sports pitch annually … how much is removed annually … which '
            'measures have been implemented … which assessments and measures have been carried out to comply with '
            'the substitution duty in §23A-9. Documentation shall be kept for at least 5 years and be available on '
            'inspection or upon request from the pollution authority." (§23A-8)'
        ),
        "comments": "Monitoring and traceability instrument for enforcement and mass-balance accounting.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-9: Substitution duty—responsible person must assess whether less environmentally risky "
            "alternatives exist (including reduced infill spread risk) and choose them if not unreasonable cost "
            "or inconvenience."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Den ansvarlige … skal vurdere om det finnes alternativ som medfører mindre risiko for '
            'miljøforstyrrelse, herunder spredning av plastholdig løst fyllmateriale til omgivelsene. Den ansvarlige '
            'skal i så fall velge dette alternativet hvis det kan skje uten urimelig kostnad eller ulempe." '
            '(§23A-9) / '
            'EN: "The person responsible … shall assess whether alternatives exist that entail less risk of '
            'environmental disturbance, including spread of plastic-containing loose infill material to the '
            'surroundings. The person responsible shall in that case choose this alternative if it can be done '
            'without unreasonable cost or inconvenience." (§23A-9)'
        ),
        "comments": "Substitution pathway to non-infill or lower-risk surfaces; cost/inconvenience limit applies.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-10: Supervision—exceptions, supervision, appeals, penalties etc. regulated in Kap. 41; "
            "county governor (statsforvalteren) supervises and may grant exemptions from chapter provisions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Forhold som gjelder unntak, tilsyn, klage, straff mv. er regulert i forurensningsforskriften '
            'kapittel 41. Statsforvalteren fører tilsyn med og kan gjøre unntak fra bestemmelsene i dette '
            'kapittelet." (§23A-10) / '
            'EN: "Matters concerning exceptions, supervision, appeals, penalties etc. are regulated in Pollution '
            'Regulation Chapter 41. The county governor supervises and may grant exemptions from the provisions of '
            'this chapter." (§23A-10)'
        ),
        "comments": "Cross-reference to enforcement framework; exemption power (kan gjøre unntak) for statsforvalteren.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-11: Entry into force—chapter applies from 1 July 2021."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Kapitlet trer i kraft 1. juli 2021." (§23A-11) / '
            'EN: "The chapter enters into force on 1 July 2021." (§23A-11)'
        ),
        "comments": "Added by FOR-2021-04-07-1096.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "§23A-12: Transitional rules—for pitches built before entry into force (and permitted planned pitches), "
            "barrier requirement (§23A-4(a)) and substitution duty (§23A-9) apply only at rehabilitation including "
            "artificial turf mat or base replacement."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "For idrettsbaner som er anlagt før dette kapitlets ikrafttredelse gjelder bestemmelsene i § 23A-4 '
            'første ledd bokstav a og § 23A-9 først når banen skal rehabiliteres, herunder når kunstgressmatter '
            'eller andre baneunderlag skal skiftes ut." (§23A-12) / '
            'EN: "For sports pitches established before this chapter entered into force, the provisions in §23A-4 '
            'first paragraph letter a and §23A-9 apply only when the pitch is to be rehabilitated, including when '
            'artificial turf mats or other pitch bases are to be replaced." (§23A-12)'
        ),
        "comments": "Phased compliance for legacy pitches; other §23A-4 requirements apply immediately from 2021.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Use/Consumption",
        "instrument_description": (
            "Legal basis: Chapter adopted under forurensningsloven §9 and §33 by FOR-2021-04-07-1096, also citing "
            "produktkontrolloven §4 for product-related aspects."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Fastsatt med hjemmel i lov 13. mars 1981 nr. 6 om vern mot forurensninger og om avfall '
            '(forurensningsloven) § 9 første ledd nr. 3 og § 33 første ledd" (kapittelhjemmel, FOR-2021-04-07-1096) / '
            'EN: "Established pursuant to the Act of 13 March 1981 No. 6 relating to protection against pollution '
            'and relating to waste (Pollution Control Act) §9 first paragraph no. 3 and §33 first paragraph" '
            "(chapter legal basis, FOR-2021-04-07-1096)"
        ),
        "comments": "Parent enabling legislation linkage.",
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
    ws.title = "4P Index - Kap 23A"

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

    output_path = "/workspace/4P_Index_Forurensningsforskriften_Kap23A.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
