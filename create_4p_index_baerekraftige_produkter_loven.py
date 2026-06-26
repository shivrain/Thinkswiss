#!/usr/bin/env python3
"""Generate 4P Index Excel for Lov om bærekraftige produkter og verdikjeder (2024)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

POLICY = {
    "policy_name": (
        "Lov om bærekraftige produkter og verdikjeder "
        "(Act relating to sustainable products and value chains / Sustainable Products Act)"
    ),
    "policy_url": "https://lovdata.no/dokument/NL/lov/2024-06-25-69",
    "policy_year": 2024,
    "policy_objective": (
        "Enabling framework statute (bærekraftige produkter-loven, Lov 2024-06-25-69) adopted by the "
        "Storting 17 June 2024 and in force from 1 July 2024. It empowers the Government to adopt regulations "
        "and conduct supervision to implement Norwegian and EEA/EU sustainability requirements across product "
        "value chains—including batteries, vehicles, packaging, plastics, electrical/electronic products, and "
        "textiles. It provides the national legal basis for ecodesign rules, extended producer responsibility "
        "(EPR), product bans, recycling targets, and information requirements aligned with the EU Green Deal "
        "and circular economy acquis; operationalised through delegated regulations (e.g. packaging, SUP, plastics)."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'NO: "Formålet med loven er å fremme bærekraftige produkter og verdikjeder for produkter som bidrar til '
        'et ressurseffektivt og bærekraftig produksjons- og forbruksmønster for produkter i en sirkulær økonomi" '
        '(§1); "bindende mål for gjenvinning av avfall inkludert mål for forberedelse til ombruk og '
        'materialgjenvinning" (§4 j) / '
        'EN: "The purpose of the Act is to promote sustainable products and value chains that contribute to a '
        'resource-efficient and sustainable production and consumption pattern for products in a circular economy" '
        '(§1); "binding targets for waste recycling including targets for preparation for reuse and material '
        'recycling" (§4(j))'
    ),
    "policy_type": 1,
    "policy_type_justification": (
        "Primary statute (lov) enacted by the Storting (Lovvedtak 93, 2023–2024; Prop. 69 LS), sanctioned "
        "25 June 2024, in force 1 July 2024 (res. 25 June 2024 nr. 1267). Framework enabling law delegating "
        "detailed product requirements to regulations; amended 5 June 2026 (§4a on fuels). Implements EEA/EU "
        "sustainable product and circular economy acquis including future plastic packaging and EPR rules."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "all sectors; manufacturing; packaging; plastics; batteries; vehicles; textiles; electronics; "
        "retail; waste management; public procurement"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, design, use/consumption, end-of-life, recycling, litter/pollution"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        'NO: "Kongen kan gi forskrift om gebyrer for behandling av søknader … og for kontrolltiltak" (§12); '
        '"kostnadene ved framleggelse av prøver eller undersøkelser … kan … dekkes av det offentlige" (§6) / '
        'EN: "The King may issue regulations on fees for processing applications … and for control measures" '
        '(§12); "costs of submitting samples or conducting investigations … may … be covered by the public '
        'authorities" (§6)'
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§1: Statutory purpose—promote sustainable products and value chains contributing to resource-efficient "
            "production and consumption patterns in a circular economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Formålet med loven er å fremme bærekraftige produkter og verdikjeder for produkter som bidrar '
            'til et ressurseffektivt og bærekraftig produksjons- og forbruksmønster for produkter i en sirkulær '
            'økonomi" (§1) / '
            'EN: "The purpose of the Act is to promote sustainable products and value chains that contribute to a '
            'resource-efficient and sustainable production and consumption pattern for products in a circular economy" '
            "(§1)"
        ),
        "comments": "Binding purpose provision; operationalised through §§3–4 regulations.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§2: Enabling power for application to Svalbard with location-specific rules."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift om lovens anvendelse for Svalbard og fastsette særlige regler" (§2) / '
            'EN: "The King may issue regulations on the application of the Act to Svalbard and establish special '
            'rules" (§2)'
        ),
        "comments": "Territorial enabling provision; no Svalbard plastics rules adopted yet.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§3: Ecodesign enabling power—regulations to implement EEA ecodesign directives including sustainability "
            "requirements on product design, composition, recycled content, durability, repairability, recyclability, "
            "and documentation/information obligations."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift for å gjennomføre rettsakter … om bærekraftskrav knyttet til produkters '
            'design … innhold av materialgjenvunnet råvare … holdbarhet, ombrukbarhet, reparerbarhet … '
            'materialgjenvinnbarhet" (§3 a); "krav til dokumentasjon av og informasjon om oppfyllelse av krav" '
            '(§3 c) / '
            'EN: "The King may issue regulations to implement acts … on sustainability requirements relating to '
            'product design … content of recycled raw materials … durability, reusability, repairability … '
            'recyclability" (§3(a)); "requirements for documentation of and information on compliance" (§3(c))'
        ),
        "comments": "Key enabling basis for EU Ecodesign Regulation and plastic product design rules.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4: Core enabling power for sustainable products and value chains—regulations for EEA obligations on "
            "batteries, vehicles, packaging, plastics, electrical/electronic products, and textiles across the "
            "full value chain."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift for å gjennomføre EØS-rettslige forpliktelser om krav til bærekraft i '
            'hele verdikjeden for batterier, kjøretøy, emballasje, plast, elektriske og elektroniske produkter '
            'og tekstiler" (§4) / '
            'EN: "The King may issue regulations to implement EEA obligations on sustainability requirements '
            'throughout the value chain for batteries, vehicles, packaging, plastics, electrical and electronic '
            'products and textiles" (§4)'
        ),
        "comments": "Central plastics/packaging enabling provision; basis for emballasje- and SUP-forskrifter.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4(a)–(b): Enabling regulations on market-actor duties and due-diligence requirements for sustainability "
            "in value chains, including social sustainability and due-diligence declarations."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "plikter og ansvar for bærekraft i verdikjeden hos markedsaktører" (§4 a); '
            '"aktsomhetskrav og aktsomhetserklæringer om bærekraft, inkludert sosial bærekraft" (§4 b) / '
            'EN: "duties and responsibilities for sustainability in the value chain of market actors" (§4(a)); '
            '"due diligence requirements and due diligence declarations on sustainability, including social '
            'sustainability" (§4(b))'
        ),
        "comments": "Value-chain governance enabling power; aligns with EU sustainable products framework.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Design",
        "instrument_description": (
            "§4(c): Enabling product sustainability and design requirements—durability, quality, SVHC content, "
            "recycled raw material content, environmental footprint, repairability, separability, reusability, "
            "and recyclability."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "bærekraftskrav til produktets egenskaper, herunder krav til design og utforming av produkter, '
            'slik som holdbarhet og kvalitet … innhold av materialgjenvunnet råvare … reparerbarhet, fradelbarhet, '
            'ombrukbarhet eller gjenvinnbarhet" (§4 c) / '
            'EN: "sustainability requirements for product properties, including requirements for product design, '
            'such as durability and quality … content of recycled raw materials … repairability, separability, '
            'reusability or recyclability" (§4(c))'
        ),
        "comments": "Design-phase enabling power; directly relevant to plastic packaging recycled-content rules.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4(g)–(h): Enabling restrictions, bans, and consumption-reduction measures for products—including "
            "import, sale, and use restrictions and measures to promote sustainable consumption patterns."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "restriksjoner på og forbud mot innførsel, utførsel, omsetning og bruk av produkter" (§4 g); '
            '"forbruksreduksjon og tiltak for å fremme bærekraftige forbruksmønstre" (§4 h) / '
            'EN: "restrictions on and prohibitions against import, export, sale and use of products" (§4(g)); '
            '"consumption reduction and measures to promote sustainable consumption patterns" (§4(h))'
        ),
        "comments": "Legal basis for SUP bans and single-use plastic consumption reduction (cf. Produktforskriften).",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4(i): Extended producer responsibility enabling power—including producer registers—for products "
            "placed on the market; national basis for packaging EPR, SUP litter EPR, and fishing-gear EPR."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "utvidet produsentansvar, inkludert registre over produsenter" (§4 i) / '
            'EN: "extended producer responsibility, including registers of producers" (§4(i))'
        ),
        "comments": "Primary EPR enabling provision for plastic packaging, SUP products, and fishing gear.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End-of-life",
        "instrument_description": (
            "§4(j): Waste prevention and management enabling power—binding recycling targets, preparation for "
            "reuse and material recycling targets, and waste plans."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "forebygging av at avfall oppstår, håndtering av avfall, bindende mål for gjenvinning av avfall '
            'inkludert mål for forberedelse til ombruk og materialgjenvinning, og avfallsplaner" (§4 j) / '
            'EN: "prevention of waste generation, waste management, binding targets for waste recycling including '
            'targets for preparation for reuse and material recycling, and waste plans" (§4(j))'
        ),
        "comments": "Enabling basis for packaging/plastic waste recycling targets under emballasjeforordningen.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§4(e)–(f), (k)–(l): Enabling powers on sustainability documentation, market information, green public "
            "procurement, and reporting/control procedures—including digital product passport information."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "dokumentasjon av bærekraftsaspekter ved produkter" (§4 e); '
            '"tilgjengeliggjøring av informasjon til markedet og brukere om bærekraftsaspekter" (§4 f); '
            '"grønne offentlige anskaffelser" (§4 k); "rapporterings- og kontrollprosedyrer" (§4 l) / '
            'EN: "documentation of sustainability aspects of products" (§4(e)); '
            '"making information available to the market and users on sustainability aspects" (§4(f)); '
            '"green public procurement" (§4(k)); "reporting and control procedures" (§4(l))'
        ),
        "comments": "Information and procurement enabling powers; supports packaging labelling and traceability.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§5: Supervision—King designates supervisory authority; powers of access to premises, digital supervision, "
            "document inspection, product sampling, test purchases, and written control reports."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kongen fastsetter hvem som er tilsynsmyndighet"; "Tilsynsmyndigheten skal ha fri adgang til '
            'bygninger … hvor produkter … befinner seg"; "kan anskaffe produkter under skjult identitet" (§5) / '
            'EN: "The King shall designate the supervisory authority"; "The supervisory authority shall have free '
            'access to buildings … where regulated products are located"; "may purchase products under concealed '
            'identity" (§5)'
        ),
        "comments": "Direct supervisory powers in force; authority designation via regulation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§6: Information and investigation duties—supervisory authority may order information (including from "
            "public bodies despite confidentiality), require free product samples, and order or allocate costs of "
            "product testing and supply-chain investigations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Tilsynsmyndigheten kan pålegge enhver å gi opplysninger"; "Offentlige myndigheter plikter å gi '
            'opplysninger uten hinder av taushetsplikt"; "framlegge prøver av produktet vederlagsfritt" (§6) / '
            'EN: "The supervisory authority may order anyone to provide information"; "Public authorities shall '
            'provide information notwithstanding confidentiality obligations"; "submit samples of the product free '
            'of charge" (§6)'
        ),
        "comments": "Binding investigative powers; Kongen may issue further regulations on documentation duties.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§7: Necessary enforcement decisions—supervisory authority may ban production/import/sale/use, order "
            "recall or withdrawal, and require compliance measures for products under §§3–4 regulations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "forby produksjon, innførsel, omsetning, eksport, bruk eller annen behandling av produkter"; '
            '"pålegge tilbakekall eller tilbaketrekking av produkter"; "pålegge gjennomføring av tiltak for å '
            'bringe et produkt … i samsvar med bestemmelser" (§7) / '
            'EN: "prohibit production, import, sale, export, use or other handling of products"; '
            '"order recall or withdrawal of products"; "order measures to bring a product into compliance" (§7)'
        ),
        "comments": "Direct market-removal powers; enforceable once product regulations are adopted.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§8: Temporary prohibition—authority may temporarily ban products until sufficient information is "
            "provided; extendable up to 12 months when needed for assessment or regulation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "midlertidig forby produksjon, innførsel, omsetning, bruk eller annen behandling av produkter '
            'inntil tilstrekkelige opplysninger er framlagt"; "forlenges i inntil 6 måneder … ytterligere 6 måneder" '
            '(§8) / '
            'EN: "temporarily prohibit production, import, sale, use or other handling of products until sufficient '
            'information is submitted"; "extended for up to 6 months … a further 6 months" (§8)'
        ),
        "comments": "Precautionary market-access instrument pending product compliance documentation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§9: Risk-reducing measures—where health/environment risk exists, authority may order warnings, recall, "
            "withdrawal, neutralisation, export ban, and recover costs from responsible parties."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "pålegge den som produserer, innfører … å treffe tiltak … for å redusere risikoen"; '
            '"tilbakekalle eller tilbaketrekke produktet"; "uskadeliggjøre produktet"; "forbud mot eksport" (§9) / '
            'EN: "order the producer, importer … to take measures to reduce the risk"; "recall or withdraw the product"; '
            '"render the product harmless"; "prohibition on export" (§9)'
        ),
        "comments": "Risk-based remediation; applicable to non-compliant plastic products on market.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§10–§11: Enabling reporting duty and internal-control requirements for regulated products and "
            "product groups."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift om plikt til å informere tilsynsmyndigheten"; '
            '"Kongen kan gi forskrift om internkontroll og internkontrollsystemer" (§10–§11) / '
            'EN: "The King may issue regulations on the duty to inform the supervisory authority"; '
            '"The King may issue regulations on internal control and internal control systems" (§§10–11)'
        ),
        "comments": "Compliance-system enabling provisions; to be specified per product regulation.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§12: Fee regulations—King may set fees for permit processing and control measures, capped at supervisory "
            "authority costs; fees enforceable by execution."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift om gebyrer … Gebyrene settes slik at de samlet ikke overstiger '
            'tilsynsmyndighetens kostnader"; "Gebyret er tvangsgrunnlag for utlegg" (§12) / '
            'EN: "The King may issue regulations on fees … Fees shall not collectively exceed the supervisory '
            'authority\'s costs"; "The fee is enforceable by execution" (§12)'
        ),
        "comments": "Cost-recovery enabling provision; fee regulations not yet adopted under this Act.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§13: Coercive fines (tvangsmulkt)—supervisory authority may impose running or advance fines to secure "
            "compliance with decisions under the Act."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Tilsynsmyndigheten kan fatte vedtak om tvangsmulkt til staten"; "Tvangsmulkten begynner å løpe '
            'dersom den ansvarlige oversitter fristen for retting" (§13) / '
            'EN: "The supervisory authority may issue decisions on coercive fines payable to the State"; '
            '"The coercive fine begins to accrue if the responsible party exceeds the deadline for rectification" '
            "(§13)"
        ),
        "comments": "Direct enforcement instrument; operative from entry into force.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§14–§15: Infringement fees and criminal fines—King may provide regulations on administrative fines for "
            "breaches of product regulations and decisions; 2-year limitation period; regulations may set fines for "
            "intentional/negligent breaches."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            'NO: "Kongen kan gi forskrift om ileggelse og utmåling av overtredelsesgebyr"; '
            '"Adgangen til å ilegge overtredelsesgebyr foreldes 2 år"; '
            '"forsettlig eller uaktsom overtredelse av forskriften straffes med bøter" (§14–§15) / '
            'EN: "The King may issue regulations on imposition and calculation of infringement fees"; '
            '"The right to impose infringement fees is time-barred after 2 years"; '
            '"intentional or negligent breach of the regulations is punishable by fines" (§§14–15)'
        ),
        "comments": "Penalty enabling provisions; infringement fee regulations to be adopted with product rules.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "§16: Entry into force—Act applies from date determined by the King; in force from 1 July 2024 "
            "(res. 25 June 2024 nr. 1267)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            'NO: "Loven gjelder fra den tiden Kongen bestemmer"; "Fra 1 juli 2024 iflg. res. 25 juni 2024 nr. 1267" '
            '(§16) / '
            'EN: "The Act applies from the date determined by the King"; "From 1 July 2024 pursuant to res. '
            '25 June 2024 no. 1267" (§16)'
        ),
        "comments": "Transitional provision; law operative since 1 July 2024.",
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
    ws.title = "4P Index - Produktloven"

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

    output_path = "/workspace/4P_Index_Baerekraftige_Produkter_Loven.xlsx"
    wb.save(output_path)
    print(f"Saved {output_path} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    main()
