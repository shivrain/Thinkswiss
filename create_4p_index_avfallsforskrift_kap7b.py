#!/usr/bin/env python3
"""Generate 4P Index Excel for Avfallsforskriften Kap. 7B (fishing-gear plastic EPR)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Avfallsforskriften – Kap. 7B: Utvidet produsentansvar for utstyr fra fiskeri, "
        "fritidsfiske og akvakultur, som inneholder plast "
        "(Waste Regulation – Ch. 7B: Extended producer responsibility for fishing gear containing plastic)"
    ),
    "policy_url": "https://lovdata.no/dokument/SF/forskrift/2004-06-01-930/KAPITTEL_9",
    "policy_year": 2025,
    "policy_objective": (
        "Chapter 7B of the Norwegian Waste Regulation (FOR-2004-06-01-930), added 18 December 2025 "
        "(forskrift 18 des 2025 nr. 2839, in force 1 January 2026). Implements EU Single-Use Plastics "
        "Directive (2019/904) extended producer responsibility for fishing gear containing plastic. "
        "Producers of gear used in commercial fisheries, aquaculture, and recreational fishing must ensure "
        "separate collection, harbour/municipal reception pick-up, preparation for reuse or material "
        "recycling, and full cost coverage, with mandatory PRO membership, reporting, and user information "
        "from 1 March 2027."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive regulation (forskrift) issued by Klima- og miljødepartementet under forurensningsloven "
        "§§31–33, 49, 52a, 63, 80–81 and produktkontrolloven §4. Chapter 7B added 18 December 2025 "
        "(forskrift nr. 2839); EEA reference: Directive (EU) 2019/904. Sub-legislative instrument, not "
        "parliamentary legislation."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "fisheries, aquaculture, recreational fishing, manufacturing, waste management, municipalities, "
        "harbours, producer responsibility organisations"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Produsent skal dekke nødvendige kostnader til separat innsamling og etterfølgende transport '
        'og behandling" (§7B-7); "Gebyr … kr 203 900" (godkjenning, §7B-23); "Årsgebyret utgjør kr 20 400" '
        '(§7B-24) / '
        'EN: "The producer shall cover necessary costs of separate collection and subsequent transport and '
        'treatment" (§7B-7); "Fee … NOK 203,900" (approval, §7B-23); "Annual fee is NOK 20,400" (§7B-24)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "§7B-4: Producer take-back EPR—producer shall ensure separate collection of a reasonable quantity "
            "of discarded fishing gear containing plastic for gear types placed on market; collection systems "
            "sufficient in geographic areas where products are used; transport to lawful waste treatment. "
            "Effective 1 March 2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal sørge for separat innsamling av en rimelig mengde kassert fiskeutstyr som '
            'inneholder plast"; "sørge for videre transport … til lovlig avfallsbehandling" (§7B-4) / '
            'EN: "The producer shall ensure separate collection of a reasonable quantity of discarded fishing '
            'gear containing plastic"; "ensure further transport … to lawful waste treatment" (§7B-4) / '
            'Implementation: +0.25 authority (Miljødirektoratet tilsyn §7B-22); +0.25 enforcement '
            '(overtredelsesgebyr §7B-25, kap. 18B); +0.25 monitoring (rapportering §7B-11, §7B-18); '
            'no +0.25 unconditional ("rimelig mengde" is qualitative, not a fixed collection target).'
        ),
        "comments": (
            "Core fishing-gear EPR collection obligation; implements SUP Directive Art. 8(7). "
            "Binding skal-language; operative from 1 March 2027."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§7B-5: Producer shall ensure discarded plastic fishing gear separately collected at harbour "
            "reception facilities (forurensningsforskriften §20-3) and municipal reception points is picked up "
            "and treated per §7B-6. Effective 1 March 2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal sørge for at kassert fiskeutstyr … som er samlet inn separat ved '
            'mottaksordning i havn … eller ved kommunale mottak, hentes på mottakene og behandles i tråd med '
            'kravene i § 7B-6" (§7B-5) / '
            'EN: "The producer shall ensure that discarded fishing gear … separately collected at harbour '
            'reception facilities … or municipal reception points is collected and treated in accordance with '
            '§7B-6" (§7B-5) / '
            'Implementation: +0.25 authority (§7B-22); +0.25 enforcement (§7B-25); +0.25 monitoring '
            '(§7B-18 collection quantities); no +0.25 unconditional (no exemptions stated in §7B-5 itself, '
            'but treatment subject to §7B-6 proportionality test).'
        ),
        "comments": (
            "Links harbour/municipal reception infrastructure to producer collection duty; "
            "multi-level coordination (harbour, municipality, producer)."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§7B-6: Producer shall ensure separately collected discarded fishing gear is prepared for reuse "
            "or materially recycled unless disproportionate on environmental, resource, BAT, and economic "
            "grounds. Effective 1 March 2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Produsent skal sørge for at kassert fiskeutstyr som er samlet inn separat … forberedes til '
            'ombruk eller materialgjenvinnes, med mindre dette ut fra en avveiing av miljøhensyn, ressurshensyn, '
            'beste tilgjengelige teknikk og økonomi ikke er forholdsmessig" (§7B-6) / '
            'EN: "The producer shall ensure that separately collected discarded fishing gear … is prepared '
            'for reuse or materially recycled, unless disproportionate on environmental, resource, BAT and '
            'economic grounds" (§7B-6) / '
            'Implementation: +0.25 authority (§7B-22); +0.25 enforcement (§7B-25); +0.25 monitoring '
            '(§7B-18(c) treatment reporting); no +0.25 unconditional (explicit proportionality exemption).'
        ),
        "comments": (
            "Waste-hierarchy treatment obligation for ghost-gear/plastic fishing equipment; "
            "proportionality clause reduces implementation strength."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "§7B-7: Producer shall cover necessary costs of separate collection, transport, and treatment "
            "of plastic fishing gear waste from products placed on market; income from sale of discarded gear "
            "or recycled material credited against costs. Effective 1 March 2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal dekke nødvendige kostnader til separat innsamling og etterfølgende transport '
            'og behandling av avfall fra fiskeutstyr som inneholder plast"; "Inntekter fra salg … skal '
            'hensyntas ved beregning av kostnadene" (§7B-7) / '
            'EN: "The producer shall cover necessary costs of separate collection and subsequent transport and '
            'treatment of waste from fishing gear containing plastic"; "Income from sale … shall be taken into '
            'account when calculating costs" (§7B-7) / '
            'Implementation: +0.25 authority (§7B-22); +0.25 enforcement (§7B-25); +0.25 monitoring '
            '(§7B-18 dokumentasjon of cost coverage); +0.25 unconditional (no cost-coverage exemption stated).'
        ),
        "comments": "Primary EPR financing obligation for fishing-gear plastic waste streams.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "§7B-8: Producer shall provide user information and promote responsible behaviour to reduce litter "
            "from fishing gear; written annual action plan required with documented follow-up available to "
            "Miljødirektoratet on request or inspection."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal sørge for informasjon til brukere av fiskeutstyr og fremme ansvarlig '
            'brukeradferd for å redusere forsøpling"; "skal ha en skriftlig plan for årlige tiltak"; '
            '"Plan og dokumentasjon … skal gjøres tilgjengelig for Miljødirektoratet ved kontroll" (§7B-8) / '
            'EN: "The producer shall provide information to users and promote responsible behaviour to reduce '
            'litter"; "shall have a written plan for annual measures"; "Plan and documentation … shall be made '
            'available to the Norwegian Environment Agency on inspection" (§7B-8) / '
            'Implementation: +0.25 authority (Miljødirektoratet); +0.25 monitoring (written plan, dokumentasjon, '
            'kontroll); +0.25 enforcement (via general tilsyn §7B-22 and kap. 18B); +0.25 unconditional.'
        ),
        "comments": (
            "Information and awareness instrument targeting marine litter from fishing gear; "
            "SUP Directive consumer-awareness parallel for gear users."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-9: Authorised representative—Norwegian producers placing gear on other EEA markets must "
            "appoint representative there; non-Norwegian producers may appoint Norwegian representative with "
            "written mandate; disclosure obligations to Miljødirektoratet on request."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "skal utpeke en representant i det landet"; "Representanten skal oppfylle forpliktelsene '
            'produsenten har etter dette kapittelet"; "skal … gi nærmere opplysninger om utpekingen" '
            '(§7B-9) / '
            'EN: "shall appoint a representative in that country"; "The representative shall fulfil the '
            'producer\'s obligations under this chapter"; "shall … provide further information on the '
            'appointment" (§7B-9) / '
            'Implementation: +0.25 authority (Miljødirektoratet); +0.25 monitoring (opplysningsplikt on '
            'request); +0.25 enforcement (§7B-25, §7B-22); +0.25 unconditional (mandatory where applicable).'
        ),
        "comments": "Cross-border producer registration aligned with SUP Directive representative rules.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-10: Mandatory membership in Miljødirektoratet-approved producer responsibility organisation "
            "(PRO) to fulfil §§7B-4–7B-8 obligations; applies to producers of components and finished gear "
            "(component producers exempt when incorporated into new gear); fallback self-compliance if PRO fails."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Produsenter skal oppfylle sine plikter etter §§ 7B-4 til 7B-8 ved å etablere eller melde seg '
            'inn i et godkjent produsentansvarsselskap"; "Dersom produsentansvarsselskapet ikke ivaretar '
            'pliktene … skal produsenten selv sørge for at pliktene oppfylles" (§7B-10) / '
            'EN: "Producers shall fulfil their obligations under §§7B-4 to 7B-8 by establishing or joining an '
            'approved PRO"; "If the PRO does not fulfil the obligations … the producer shall ensure compliance '
            'itself" (§7B-10) / '
            'Implementation: +0.25 authority (Miljødirektoratet godkjenning §7B-13); +0.25 enforcement '
            '(overtredelsesgebyr §7B-25); +0.25 monitoring (§7B-11, §7B-18); +0.25 unconditional.'
        ),
        "comments": "Central collective EPR membership obligation; direct infringement sanction in §7B-25.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-11: Producer annual reporting to PRO on tonnes of plastic fishing gear placed on market, "
            "split by commercial fisheries, aquaculture, and recreational fishing; based on total product weight."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsent skal rapportere til produsentansvarsselskapet … om mengden fiskeutstyr som '
            'inneholder plast den bringer i omsetning"; "angis mengden fiskeutstyr i tonn, fordelt på utstyr '
            'benyttet i fiskeri, i akvakultur og i fritidsfiske" (§7B-11) / '
            'EN: "The producer shall report to the PRO … the quantity of fishing gear containing plastic placed '
            'on market"; "quantity in tonnes, distributed across gear used in fisheries, aquaculture and '
            'recreational fishing" (§7B-11) / '
            'Implementation: +0.25 authority (PRO/Miljødirektoratet chain); +0.25 monitoring (annual reporting); '
            '+0.25 enforcement (§7B-25 via PRO membership); +0.25 unconditional.'
        ),
        "comments": "Data-flow obligation enabling sector-split EPR accounting.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§7B-13: PRO approval framework—PRO must be approved by Miljødirektoratet; fisheries-gear PRO must "
            "also cover recreational fishing gear; individual approval possible; conditions and withdrawal with "
            "residual treatment duty for already-collected gear."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal være godkjent av Miljødirektoratet"; "Godkjenningen kan trekkes '
            'tilbake dersom pliktene … ikke overholdes"; "skal sørge for at kassert fiskeutstyr som allerede er '
            'separat innsamlet … behandles i tråd med kravene i § 7B-6" (§7B-13) / '
            'EN: "The PRO shall be approved by the Norwegian Environment Agency"; "Approval may be withdrawn if '
            'obligations are not met"; "shall ensure already collected gear is treated per §7B-6" (§7B-13) / '
            'Implementation: +0.25 authority (Miljødirektoratet designated); +0.25 enforcement (withdrawal); '
            '+0.25 monitoring (approval conditions); no +0.25 unconditional (vilkår may be set).'
        ),
        "comments": (
            "Governance/coordination instrument establishing PRO approval regime; "
            "fisheries PRO must include recreational fishing coverage."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "§7B-15: Enabling power—Miljødirektoratet may by regulation or individual decision set minimum "
            "requirements for degree of separate collection, preparation for reuse, or material recycling of "
            "discarded plastic fishing gear that PROs must achieve on behalf of members."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet kan ved forskrift eller i enkeltvedtak sette minstekrav til grad av separat '
            'innsamling, forberedelse til ombruk eller materialgjenvinning" (§7B-15) / '
            'EN: "The Norwegian Environment Agency may by regulation or individual decision set minimum '
            'requirements for the degree of separate collection, preparation for reuse or material recycling" '
            '(§7B-15) / '
            'Implementation: +0.25 authority (Miljødirektoratet) only; enabling kan-language—not yet exercised; '
            'no enforcement, monitoring, or operative standard in this provision itself.'
        ),
        "comments": (
            "Enabling power (kan); in_force=0 per Rule 12 until subordinate minstekrav forskrift is adopted. "
            "Would become regulatory performance standard if exercised."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§7B-17: Enabling coordination power—Miljødirektoratet may order how §7B-4/7B-5 collection shall "
            "be fulfilled by PROs or require PROs to coordinate collection for continuous adequate coverage."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet kan gi pålegg om hvordan plikten til innsamling … skal oppfylles"; '
            '"kan … pålegge produsentansvarsselskap å koordinere innsamling" (§7B-17) / '
            'EN: "The Norwegian Environment Agency may order how the collection obligation … shall be fulfilled"; '
            '"may … order the PRO to coordinate collection" (§7B-17) / '
            'Implementation: +0.25 authority (Miljødirektoratet) only; enabling kan-power not yet exercised.'
        ),
        "comments": (
            "Multi-level governance/coordination enabling instrument; in_force=0 until pålegg issued. "
            "Note coordination across PROs and collection actors."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "§7B-18: PRO annual reporting to Miljødirektoratet on tonnes placed on market, separately collected, "
            "and treatment by sector (fisheries/aquaculture/recreational); biennial third-party verification; "
            "5-year documentation retention; cost-coverage documentation for §7B-7."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal årlig rapportere til Miljødirektoratet om … mengde fiskeutstyr … '
            'brakt i omsetning … samlet inn separat … behandling"; "Hvert andre år skal … en nøytral tredjepart '
            'verifisere"; "Dokumentasjon … skal tas vare på i minst 5 år" (§7B-18) / '
            'EN: "The PRO shall annually report … quantities placed on market … separately collected … treatment"; '
            '"Every two years a neutral third party shall verify"; "Documentation … kept for at least 5 years" '
            '(§7B-18) / '
            'Implementation: +0.25 authority (Miljødirektoratet); +0.25 enforcement (§7B-22, §7B-25); '
            '+0.25 monitoring (annual reporting, tredjepartsverifisering); +0.25 unconditional.'
        ),
        "comments": "Comprehensive monitoring and verification framework for fishing-gear EPR performance.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-19: PRO public transparency—owners, members, fee rates (NOK per unit or tonne), and waste-handler "
            "selection procedures must be publicly available."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskaper skal sørge for at følgende opplysninger er offentlig tilgjengelige: '
            '… eiere … medlemmer … vederlag … informasjon om fremgangsmåten for valg av aktører" (§7B-19) / '
            'EN: "PROs shall ensure the following information is publicly available: … owners … members … fees … '
            'information on procedures for selecting waste handlers" (§7B-19) / '
            'Implementation: +0.25 authority (PRO obligation); +0.25 monitoring (public disclosure); '
            'no explicit penalty in §7B-19; no unconditional enforcement sub-score.'
        ),
        "comments": "Information/transparency instrument supporting market accountability of PRO fees.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "§7B-22: Supervision—Miljødirektoratet supervises compliance with chapter provisions and decisions; "
            "may charge supervision fees under forurensningsforskriften kap. 39."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet fører tilsyn med at bestemmelsene i dette kapittelet … overholdes"; '
            '"Miljødirektoratet kan kreve gebyr ved gjennomføring av tilsyn" (§7B-22) / '
            'EN: "The Norwegian Environment Agency shall supervise compliance with this chapter"; '
            '"may charge a fee for conducting supervision" (§7B-22) / '
            'Implementation: +0.25 authority (Miljødirektoratet designated); +0.25 monitoring (tilsyn); '
            'enforcement via §7B-25 and kap. 18B referenced elsewhere, not in §7B-22 itself.'
        ),
        "comments": "Governance/coordination—central enforcement authority for fishing-gear EPR chapter.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-23–§7B-24: PRO administrative fees to state—approval processing fee scale (NOK 25,500–203,900) "
            "and annual reporting fee (NOK 20,400 per approved PRO); CPI-indexed."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Produsentansvarsselskap skal betale gebyr til statskassen"; "kr 203 900 … kr 25 500"; '
            '"Årsgebyret utgjør kr 20 400" (§7B-23–§7B-24) / '
            'EN: "The PRO shall pay fees to the Treasury"; "NOK 203,900 … NOK 25,500"; '
            '"Annual fee is NOK 20,400" (§§7B-23–24) / '
            'Implementation: +0.25 authority (Miljødirektoratet sets sats); +0.25 enforcement (mandatory gebyr); '
            'no monitoring/unconditional sub-scores for fee schedule itself.'
        ),
        "comments": "Economic instrument—cost-recovery fees funding Miljødirektoratet PRO administration.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7B-25: Infringement fee—Miljødirektoratet may impose administrative fine for breach of §7B-10(1) "
            "PRO membership obligation; applies Avfallsforskriften kap. 18B framework."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Miljødirektoratet kan ved overtredelse av § 7B-10 første ledd ilegge … overtredelsesgebyr"; '
            '"gjelder bestemmelsene i denne forskriftens kapittel 18B" (§7B-25) / '
            'EN: "The Norwegian Environment Agency may impose an infringement fee for breach of §7B-10(1)"; '
            '"the provisions of Chapter 18B apply" (§7B-25) / '
            'Implementation: +0.25 authority (Miljødirektoratet); +0.25 enforcement (overtredelsesgebyr); '
            '+0.25 monitoring (breach detection via tilsyn/rapportering); kan-language for fee imposition '
            'reduces unconditional score.'
        ),
        "comments": (
            "Direct enforcement sanction for failure to join approved PRO; cross-ref. Kap. 18B penalty framework."
        ),
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
    ws.title = "4P Index - Kap 7B"

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

    output_path = "/workspace/4P_Index_Avfallsforskrift_Kap7b.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
