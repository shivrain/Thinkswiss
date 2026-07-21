#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Arrêté 09311/2007 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Arrete_09311_Huiles_Usees.xlsx"

COLUMNS = [
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


def bi(fr: str, en: str) -> str:
    return f"{fr} / {en}"


def q(fr: str, en: str) -> str:
    return f"'{fr}' / '{en}'"


POLICY = {
    "policy_name": bi(
        "Arrêté interministériel n° 09311 du 5 octobre 2007 portant gestion des huiles usagées",
        "Interministerial Order No. 09311 of 5 October 2007 on used-oil management",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/73/arretes-cadre-juridique/"
        "19096/arrete-interministeriel-n-09311-du-5-octobre-2007-portant-gestion-des-huiles-usees.pdf"
    ),
    "policy_year": 2007,
    "policy_objective": bi(
        "Réglementer la collecte, le stockage, le transport, le traitement et l'élimination des huiles "
        "usagées (minérales, végétales, animales, synthétiques) en imposant l'agrément d'État, le "
        "registre vert et l'interdiction des rejets dans l'environnement — réduire les flux de déchets "
        "dangereux souvent co-présents avec les déchets plastiques en milieu industriel et municipal, "
        "sans mesures spécifiques aux plastiques.",
        "Regulate collection, storage, transport, treatment and disposal of used oils (mineral, vegetable, "
        "animal, synthetic) by imposing state approval, the green register and prohibition of environmental "
        "discharge — reduce hazardous waste streams often co-occurring with plastic waste in industrial "
        "and municipal settings, without plastic-specific measures.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('cinq cents litres', 'five hundred litres')} annual production threshold for green register; "
        f"{q('quinze jours', 'fifteen days')} maximum pickup delay for lots >{q('600 litres', '600 litres')}; "
        f"{q('1 400°C au minimum', '1,400°C minimum')} for incineration/co-incineration; "
        f"{q('capacité de stockage de 100 m3 et d au minimum 25 m3', 'storage capacity of 100 m3 and at least 25 m3')} for collectors; "
        f"{q('douzième de la capacité annuelle d élimination', 'one twelfth of annual elimination capacity')} minimum eliminator storage; "
        f"{q('cinq ans renouvelables', 'five years renewable')} agrément validity."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Arrêté interministériel adopté le 5 octobre 2007 conjointement par les Ministres de "
        "l'Environnement, des Mines et de l'Industrie, et de l'Énergie, publié au Journal officiel. "
        "Il s'agit d'un texte réglementaire d'application exécutive opérationnalisant le Code de "
        "l'environnement (Loi 2001-01) et le Décret 2001-282, et non d'une loi parlementaire (≠1).",
        "Interministerial order adopted on 5 October 2007 jointly by the Ministers of Environment, "
        "Mines and Industry, and Energy, published in the Official Journal. It is an executive "
        "implementing regulation operationalising the Environmental Code (Law 2001-01) and Decree "
        "2001-282, not parliamentary legislation (≠1).",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "industrie, transport, environnement, énergie, gestion des déchets",
        "industry, transport, environment, energy, waste management",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "utilisation, collecte, élimination",
        "use, collection, disposal",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Annexe 1: {q('Elle dépose une consignation de... FCFA au bureau de gestion du MEPN', 'A deposit of... FCFA must be paid at the MEPN management office')} "
        f"for collection and elimination agrément applications (amount set administratively; blank in published text). "
        f"§19: annual statistics reporting to DEEC."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. 3: Prohibitions including "
            f"{q('de déposer ou de laisser couler des huiles usagées... dans ou sur le sol, dans les eaux de surface ou les eaux souterraines, dans les égouts, les canalisations ou les collecteurs', 'depositing or allowing used oils to flow... on or into soil, surface or groundwater, sewers, pipes or collectors')}; "
            f"{q('d effectuer la combustion des huiles usagées, sauf si elle est réalisée dans les conditions prévues à l article 2', 'burning used oils except under Article 2 conditions')}; "
            f"{q('de se débarrasser des huiles usagées, sauf à les remettre à des entreprises agréées', 'disposing of used oils except by handing them to approved companies')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 15: {q('Les entreprises agréées sont contrôlées par l autorité compétente', 'Approved companies are controlled by the competent authority')}. "
            f"Art. 20: {q('Toute violation... expose le contrevenant aux sanctions prévues par la réglementation en vigueur', 'Any violation exposes the offender to sanctions under applicable regulations')}."
        ),
        "comments": bi(
            "Interdictions environnementales centrales; réduit les rejets d'hydrocarbures co-localisés avec déchets plastiques/municipaux.",
            "Core environmental prohibitions; reduces hydrocarbon discharges co-located with plastic/municipal waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 3 §3–5: Prohibited mixing — "
            f"{q('de mélanger les huiles usagées avec des PCB ou avec des déchets dangereux', 'mixing used oils with PCBs or hazardous waste')}; "
            f"{q('de mélanger... de l eau ou tout corps étranger, tel que solvants, produits de nettoyage, détergents, antigel, autres combustibles', 'adding water or foreign matter such as solvents, cleaning products, detergents, antifreeze, other fuels')} "
            f"before or during collection/storage."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 1 (éliminateur): {q('tenir une comptabilité matière... notamment la teneur en PCB', 'keep material accounting... including PCB content')}. "
            f"Annexe 1 (ramasseur): {q('séparation entre les huiles stockées et tous autres déchets et substances d une autre nature', 'separation between stored oils and all other waste and substances of a different nature')}."
        ),
        "comments": bi(
            "Empêche la co-mélange huiles/déchets dangereux; pertinent pour flux mixtes huiles-plastiques en sites industriels.",
            "Prevents co-mingling of oils/hazardous waste; relevant for mixed oil-plastic streams at industrial sites.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 2: {q('les seules utilisations... sont... la régénération et rutilisation industrielle comme combustible', 'the only uses... are regeneration and industrial reuse as fuel')}; "
            f"{q('Cette dernière utilisation ne peut être autorisée que dans des établissements agréés', 'The latter use may only be authorised in approved establishments')}. "
            f"{q('L administration fixera annuellement les quantités destinées à la régénération et à l utilisation en combustible', 'The administration shall set annual quantities for regeneration and fuel use')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 1 (définitions): {q('Valorisation... régénération, le recyclage, la valorisation énergétique, avec neutralisation effective de toute émanation susceptible de polluer l atmosphère', 'Valorisation... regeneration, recycling, energy recovery, with effective neutralisation of emissions likely to pollute the atmosphere')}. "
            f"Art. 19: annual statistics on final destination of oils."
        ),
        "comments": bi(
            "Hiérarchie de valorisation (régénération > combustible agréé); instrument d'économie circulaire pour huiles usagées.",
            "Valorisation hierarchy (regeneration > approved fuel use); circular economy instrument for used oils.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 5: {q('Seules les Sociétés agréées par l Etat sont autorisées à effectuer l élimination et/ou la collecte des huiles usagées en vue de leur traitement', 'Only companies approved by the State are authorised to carry out elimination and/or collection of used oils for treatment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 8–9: agrément required from Environment Minister for collectors and eliminators. "
            f"Art. 10–12: commission examines applications; Minister decides. "
            f"Art. 13: {q('L agrément ne dispense pas... de déposer... un dossier de demande d autorisation d exploiter... installations classées', 'Approval does not exempt... from filing an ICPE operating authorisation application')}."
        ),
        "comments": bi(
            "Monopole d'agrément d'État pour la filière huiles usagées; gouvernance des déchets dangereux industriels.",
            "State approval monopoly for used-oil chain; governance of industrial hazardous waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 6: Holder obligations — "
            f"{q('remettre leurs huiles usagées aux ramasseurs agréés', 'hand used oils to approved collectors')}; "
            f"{q('assurer eux mêmes le transport... aux éliminateurs agréés', 'themselves transport... to approved eliminators')}; "
            f"or {q('assurer eux-mêmes l élimination... après avoir obtenu un agrément', 'themselves eliminate... after obtaining approval')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 4: defines détenteurs, ramasseurs, éliminateurs. "
            f"Annexe 2 (registre vert): traceability from producer to collector to eliminator via removal slips (bordereaux)."
        ),
        "comments": bi(
            "Responsabilité élargie du producteur pour huiles usagées (garages, industries, ateliers plasturgie).",
            "Extended producer responsibility for used oils (garages, industries, plastics workshops).",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 7 & Annexe 2: {q('Toute entreprise qui produit une quantité annuelle minimale de cinq cents litres d huiles usagées tient un registre appelé registre vert', 'Any company producing a minimum annual quantity of five hundred litres of used oils must keep a register called the green register')} "
            f"with model established by DEEC; must allow DEEC inspection at any time."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 2: register records new lubricant purchases, used oil quantities/characteristics, "
            f"generator process, transfer dates, transporter identity, destination, elimination/valorisation method. "
            f"Three bordereau templates (producer, collector-transporter, eliminator)."
        ),
        "comments": bi(
            "Seuil quantitatif (500 L/an) pour traçabilité; pertinent pour PME industrielles incl. maintenance d'équipements plastiques.",
            "Quantitative threshold (500 L/year) for traceability; relevant for industrial SMEs incl. plastics equipment maintenance.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 1 (Transport): {q('Ensemble des opérations de chargement, d acheminement et de déchargement des huiles usagées, au moyen de véhicule spécialement aménagé... citerne ou camionnette disposant de cuve étanche et solidement fixée à la carrosserie', 'Loading, transport and unloading operations by specially equipped vehicle... tanker or van with watertight tank firmly fixed to body')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 2 (bordereau): transport modes {q('BENN', 'BENN')}, {q('Citerne', 'Tanker')}; "
            f"{q('admises au transport selon les dispositions du règlement pour le transport des matières dangereuses', 'admitted for transport under dangerous goods transport regulations')}. "
            f"Art. 10: Transport Ministry represented on agrément commission."
        ),
        "comments": bi(
            "Exigences de transport sécurisé des matières dangereuses; complément à la gestion des flux de déchets mixtes.",
            "Secure dangerous goods transport requirements; complement to mixed waste stream management.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Annexe 1 (ramasseur): {q('Le ramasseur agréé doit procéder dans un délai de quinze jours à l enlèvement de tout lot d huiles usagées supérieur à 600 litres', 'The approved collector must remove any lot of used oils exceeding 600 litres within fifteen days')}. "
            f"Storage: {q('capacité de stockage de 100 m3 et d au minimum 25 m3', 'storage capacity of 100 m3 and at least 25 m3')} with separation from other wastes."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 1: removal slip (bordereau d'enlèvement) mandatory for each collection; copy to holder. "
            f"Contracts between collectors and eliminators must be filed with DEEC within 15 days. "
            f"Annual tonnage reporting to DEEC."
        ),
        "comments": bi(
            "Délais et capacités quantitatifs contraignants pour la collecte; séparation obligatoire des autres déchets.",
            "Binding quantitative collection deadlines and capacities; mandatory separation from other waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 4: Eliminator includes {q('régénération des huiles... valorisation énergétique au moyen de la technique dite de l incinération ou de la co-incinération, à une température de 1 400°C au minimum', 'oil regeneration... energy recovery by incineration or co-incineration at a minimum temperature of 1,400°C')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 1 (éliminateur): technical note on recycling, regeneration, incineration, co-incineration processes and capacities. "
            f"Art. 1 (Valorisation énergétique): {q('neutralisation des effets polluants et récupération adéquate de la chaleur produite', 'neutralisation of polluting effects and adequate heat recovery')}."
        ),
        "comments": bi(
            "Seuil technique de 1 400°C pour incinération/co-incinération; lien avec NS 05-062 (pollution atmosphérique).",
            "1,400°C technical threshold for incineration/co-incineration; link to NS 05-062 (atmospheric pollution).",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Annexe 1 (éliminateur): obligations including "
            f"{q('comptabilité matière... teneur en PCB et le pourcentage d eau', 'material accounting... PCB content and water percentage')}; "
            f"{q('capacité minimale de stockage... égale au douzième de la capacité annuelle d élimination', 'minimum storage capacity equal to one twelfth of annual elimination capacity')}; "
            f"{q('obligation de reprise des huiles usagées proposées dans la limite de la capacité de traitement', 'obligation to take back proposed used oils within treatment capacity')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe 1: material accounting presented on first request of ICPE control service. "
            f"Monthly December statistics on reception and treatment tonnages to DEEC. "
            f"Art. 14: operations must not cause avoidable harm to water, air and/or soil."
        ),
        "comments": bi(
            "Traçabilité PCB pertinente pour additifs plastiques/chlorés; comptabilité matière pour filière d'élimination.",
            "PCB traceability relevant for plastic/chlorinated additives; material accounting for elimination chain.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. 14: {q('Toute entreprise agréée doit effectuer toutes les opérations sans qu il en résulte des préjudices évitables pour l eau, l air et/ou le sol', 'Every approved company must carry out all operations without causing avoidable harm to water, air and/or soil')}. "
            f"Agréments valid {q('cinq ans renouvelables', 'five years renewable')}, subject to withdrawal or suspension."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 15: control by competent authority on agrément conditions. "
            f"Art. 18: one-month compliance deadline before agrément withdrawal; "
            f"eliminated oils must be transferred to approved company within one month."
        ),
        "comments": bi(
            "Protection intégrée eau/air/sol; complète NS 05-061 (eaux) et NS 05-062 (air) pour sites de traitement.",
            "Integrated water/air/soil protection; complements NS 05-061 (water) and NS 05-062 (air) for treatment sites.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 19: {q('Durant le premier trimestre de chaque année, les entreprises agréées doivent transmettre leurs statistiques annuelles de collecte, de stockage et/ou de transformation des huiles usagées à la DEEC', 'During the first quarter of each year, approved companies must transmit annual statistics on collection, storage and/or transformation of used oils to DEEC')}, "
            f"including percentage on use and final destination."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Annexe 1: collectors and eliminators also report annual tonnages collected/delivered/treated. "
            f"Art. 10: interministerial commission (Environment, Industry, Sanitation, Energy, Transport, Finance) examines agrément dossiers."
        ),
        "comments": bi(
            "Reporting annuel pour suivi de la filière; données utiles pour co-occurrence huiles/déchets en milieu urbain.",
            "Annual reporting for chain monitoring; data useful for oil/waste co-occurrence in urban settings.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 1 (PCB): {q('Biphényles polychlorurés, utilisés dans l industrie... comme additif dans les peintures, les papiers autocopiants et dans les plastiques', 'Polychlorinated biphenyls, used in industry... as additive in paints, carbonless papers and in plastics')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3 §4: {q('interdit... de mélanger les huiles usagées avec des PCB', 'prohibited to mix used oils with PCBs')}. "
            f"Annexe 1: eliminators must record PCB content in material accounting."
        ),
        "comments": bi(
            "Seule mention explicite des plastiques (PCB comme additif); pertinence indirecte pour déchets plastiques chlorés.",
            "Only explicit mention of plastics (PCB as additive); indirect relevance for chlorinated plastic waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 18 & 20: {q('En cas de non respect... proposition de suspension ou de retrait de l agrément', 'In case of non-compliance... proposal to suspend or withdraw approval')}; "
            f"{q('Toute violation des dispositions du présent arrêté, expose le contrevenant aux sanctions prévues par la réglementation en vigueur', 'Any violation exposes the offender to sanctions under applicable regulations')}. "
            f"Art. 21: abrogates Arrêté interministériel n° 003032 of 24 March 1982."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 18: one-month corrective deadline; upon withdrawal, holder must ensure stored oils cause no nuisance "
            f"and arrange elimination by approved company within one month. "
            f"Art. 17: motivated refusal of agrément notified by Environment Minister."
        ),
        "comments": bi(
            "Sanctions et retrait d'agrément; renforce l'effectivité de la filière huiles usagées.",
            "Sanctions and agrément withdrawal; strengthens used-oil chain enforceability.",
        ),
    },
]


def build_workbook():
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index Coding"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    auto_fill = PatternFill("solid", fgColor="E2EFDA")

    for col_idx, (col_letter, col_name) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=1, column=col_idx, value=f"{col_letter} — {col_name}")
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for row_idx, instrument in enumerate(INSTRUMENTS, start=2):
        row_data = {**POLICY, **instrument}
        for col_idx, (_, key) in enumerate(COLUMNS, start=1):
            if key in ("policy_score", "instrument_score"):
                continue
            ws.cell(row=row_idx, column=col_idx, value=row_data.get(key, ""))

        g_col, i_col, k_col, m_col = (get_column_letter(n) for n in (7, 9, 11, 13))
        p_col, t_col, o_col, v_col = (get_column_letter(n) for n in (16, 20, 15, 22))
        ws[f"{o_col}{row_idx}"] = f"=AVERAGE({g_col}{row_idx},{i_col}{row_idx},{k_col}{row_idx},{m_col}{row_idx},{p_col}{row_idx})"
        ws[f"{v_col}{row_idx}"] = f"=AVERAGE({p_col}{row_idx},{t_col}{row_idx})"
        ws[f"{o_col}{row_idx}"].fill = auto_fill
        ws[f"{v_col}{row_idx}"].fill = auto_fill

    widths = {
        "A": 48, "B": 52, "C": 10, "D": 52, "E": 12, "F": 60,
        "G": 10, "H": 52, "I": 14, "J": 40, "K": 14, "L": 36,
        "M": 12, "N": 52, "O": 12, "P": 14, "Q": 22, "R": 60,
        "S": 14, "T": 18, "U": 52, "V": 14, "W": 52,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")

    ws.freeze_panes = "A2"
    wb.save(OUTPUT)
    print(f"Saved {OUTPUT} ({len(INSTRUMENTS)} instrument rows)")


if __name__ == "__main__":
    build_workbook()
