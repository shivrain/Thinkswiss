#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2009-24 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2009-24_Code_Assainissement.xlsx"

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
        "Loi n° 2009-24 du 8 juillet 2009 portant Code de l'Assainissement",
        "Law No. 2009-24 of 8 July 2009 establishing the Sanitation Code",
    ),
    "policy_url": "https://www.sante.gouv.sn/sites/default/files/3.%20Code_assainissement.pdf",
    "policy_year": 2009,
    "policy_objective": bi(
        "Unifier le cadre juridique de l'assainissement liquide (eaux usées, excrétas, eaux pluviales) en "
        "fixant les obligations de planification des communes et communautés rurales, les interdictions de "
        "déversement dans les égouts publics (dont déchets plastiques), les régimes des effluents domestiques, "
        "industriels et hospitaliers, et les sanctions — gouvernance indirecte des déchets plastiques via "
        "l'interdiction explicite de leur rejet dans les collecteurs publics.",
        "Unify the legal framework for liquid sanitation (wastewater, excreta, stormwater) by setting planning "
        "obligations for communes and rural communities, prohibitions on discharge into public sewers "
        "(including plastic waste), regimes for domestic, industrial and hospital effluents, and sanctions — "
        "indirect governance of plastic waste through explicit prohibition of discharge into public collectors.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('soixante mètres', 'sixty metres')} mandatory sewer connection distance; "
        f"{q('trente-cinq mètres', 'thirty-five metres')} / {q('quinze mètres', 'fifteen metres')} infiltration setback from wells/water bodies; "
        f"{q('eaux de température supérieure à 30°C', 'water temperature above 30°C')} prohibited in sewers; "
        f"{q('pH<5.5 et >8.5', 'pH <5.5 and >8.5')} prohibited; "
        f"{q('coliformes fécaux... inférieure ou égale à 1.000 UFC/100 ml', 'faecal coliforms... less than or equal to 1,000 CFU/100 ml')} for restricted irrigation reuse; "
        f"{q('détergents biodégradables à 90% au moins', 'detergents biodegradable at least 90%')} for hospital wastewater."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Code de l'assainissement adopté par l'Assemblée nationale le 17 juin 2009 et par le Sénat le "
        "29 juin 2009, promulgué par le Président de la République le 8 juillet 2009 en tant que "
        "Loi n° 2009-24. Il s'agit d'une loi parlementaire (Code), et non d'un texte réglementaire "
        "d'application exécutive (≠0.75).",
        "Sanitation Code adopted by the National Assembly on 17 June 2009 and by the Senate on "
        "29 June 2009, promulgated by the President of the Republic on 8 July 2009 as Law No. 2009-24. "
        "It is parliamentary legislation (Code), not executive implementing regulation (≠0.75).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "assainissement, urbanisme, collectivités locales, environnement, santé, industrie",
        "sanitation, urban planning, local government, environment, health, industry",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "utilisation, collecte, élimination",
        "use, collection, disposal",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Art. L 60: {q('redevance... calculé sur la base du volume d eau rejeté et de la qualité des effluents', 'fee... calculated on volume of water discharged and effluent quality')} for industrial effluents; "
        f"Art. L 73: {q('taxes et redevances de rejet d eau en milieu naturel', 'taxes and fees for discharge into natural environment')}; "
        f"Art. L 88: {q('redevance calculées à la tonne ou au mètre cube', 'fee calculated per tonne or cubic metre')} for septage at déposantes; "
        f"Art. L 97: fees for autonomous sanitation control."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. L 29: {q('Il est formellement interdit de déverser dans les collecteurs publics d eaux usées : ... les ordures ménagères, les déchets plastiques', 'It is strictly prohibited to discharge into public wastewater collectors: ... household waste, plastic waste')}, "
            f"along with hydrocarbons, used oils, radioactive substances, paint residues, septic tank contents, etc."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 30: {q('La liste de ces déversements interdits n est pas limitative', 'The list of prohibited discharges is not exhaustive')}; "
            f"sworn sanitation agents may take control samples at any user premises. "
            f"Arts. L 104–L 108: penal sanctions (2 months–2 years imprisonment, FCFA 20,000–2,000,000 fines)."
        ),
        "comments": bi(
            "Mention explicite des déchets plastiques — interdiction directe de déversement dans les égouts publics.",
            "Explicit mention of plastic waste — direct prohibition on discharge into public sewers.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. L 13: {q('Le rejet d effluents non épurés d origine domestique, d excrétas et de boues de vidange dans les caniveaux, canaux d eaux pluviales... ainsi que sur la surface des sols... est interdit sur toute l étendue du territoire national', 'Discharge of untreated domestic effluent, excreta and septage into gutters, open stormwater channels... and onto soil surfaces... is prohibited throughout the national territory')}. "
            f"Same prohibition for rivers, lakes, ponds and sea."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 3: all liquid waste discharges into natural environment require prior depollution. "
            f"Art. L 5: pollution sources subject to prior authorisation and administrative inquiry. "
            f"Arts. L 198–L 103: sworn agents from sanitation, health, environment ministries may inspect."
        ),
        "comments": bi(
            "Interdiction des rejets non épurés; limite la fuite de déchets solides/plastiques via les eaux usées.",
            "Prohibition of untreated discharges; limits leakage of solid/plastic waste via wastewater.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Arts. L 8–L 12: {q('Toute commune doit être dotée d un plan directeur d assainissement des eaux usées et eaux pluviales', 'Every commune must have a wastewater and stormwater sanitation master plan')}; "
            f"{q('Toute communauté rurale doit également être dotée d un plan local d hydraulique et d assainissement', 'Every rural community must also have a local water and sanitation plan')} including investment programming."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. L 9: plan requires prior urban master plan; sanitation zoning opposable to third parties. "
            f"Art. L 10: adopted by commune/rural community deliberation; approved by competent authority. "
            f"Art. L 12: plan costs borne by local authorities; State may contribute."
        ),
        "comments": bi(
            "Obligation de planification locale pour l'assainissement; cadre pour gestion intégrée des déchets urbains incl. plastiques.",
            "Local sanitation planning obligation; framework for integrated urban waste management incl. plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. L 18: {q('Lorsqu un égout public est accessible à moins de soixante mètres... le dispositif d évacuation... doit être raccordé à l égout public', 'When a public sewer is accessible within sixty metres... the evacuation system... must be connected to the public sewer')}. "
            f"Art. L 20: if no sewer within 60 m, {q('installation d assainissement autonome', 'autonomous sanitation installation')} required per Senegalese standards."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 19: owners failing to connect despite obligation liable for sewer taxes/redevances. "
            f"Art. L 92: upon new public network, buildings must connect within {q('six mois', 'six months')}. "
            f"Art. L 93: owners without public network must install compliant autonomous sanitation."
        ),
        "comments": bi(
            "Seuil quantitatif de 60 m pour raccordement; réduit les rejets sauvages de déchets dans les caniveaux.",
            "60 m quantitative threshold for connection; reduces uncontrolled waste discharge into gutters.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. L 29 (suite): additional sewer prohibitions — "
            f"{q('eaux de température supérieure à 30°C', 'water temperature above 30°C')}, "
            f"{q('eaux de pH<5.5 et >8.5', 'water pH <5.5 and >8.5')}, "
            f"{q('corps et matières solides les liquides ou produits gazeux nocifs ou inflammables', 'solid bodies and harmful or flammable liquids or gases')}, "
            f"{q('boues, les sables, les gravats, les colles, les goudrons, les huiles', 'sludge, sand, rubble, glues, tars, oils')} compromising sewer/station operation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 37: effluent exceeding pollution limits must be depolluted before sewer discharge. "
            f"Art. L 36: dangerous sewer discharges assimilated to Art. L 35 offences. "
            f"Art. L 104: 2 months–2 years imprisonment, FCFA 20,000–2,000,000 fine."
        ),
        "comments": bi(
            "Paramètres quantitatifs (température, pH) et interdictions de matières solides; complète l'interdiction des plastiques (L 29).",
            "Quantitative parameters (temperature, pH) and solid matter prohibitions; complements plastic prohibition (L 29).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arts. L 53–L 56: ICPE installations discharging polluted water must submit treatment dossier with authorisation request. "
            f"{q('Les teneurs en substances polluantes... sont fixées sur la base des valeurs retenues par les textes en vigueur, notamment le code de l environnement et la norme sénégalaise NS 05-061', 'Pollutant concentrations... are set based on applicable texts, notably the Environmental Code and Senegalese Standard NS 05-061')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. L 54–L 55: special discharge convention with sanitation service before connection; "
            f"sets maximum flow, pollutant content and pollution load. "
            f"Art. 57: industrial effluents may require pre-treatment. "
            f"Arts. L 61–L 62: flow metering required."
        ),
        "comments": bi(
            "Régime ICPE pour effluents industriels (incl. plasturgie); renvoie à NS 05-061 pour limites de rejet.",
            "ICPE regime for industrial effluents (incl. plastics processing); refers to NS 05-061 for discharge limits.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arts. L 58–L 59: {q('En cas de rejets non conformes... le Service chargé de l assainissement... peut obturer d office les branchements concernés et suspendre immédiatement les autorisations de déversement', 'In case of non-compliant discharges... the sanitation service... may seal off connections and immediately suspend discharge authorisations')}. "
            f"{q('Le branchement... est obturé sur constat, par un agent assermenté', 'The connection is sealed on report by a sworn agent')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Art. L 66: sampling and controls at any time to verify compliance with special convention. "
            f"Environment Ministry must be informed of suspension (Art. L 58). "
            f"Art. L 68: technical interventions due to ICPE fault billed to operator."
        ),
        "comments": bi(
            "Sanction opérationnelle (obturation) pour rejets industriels non conformes; pertinent pour effluents plastiques.",
            "Operational sanction (sealing) for non-compliant industrial discharges; relevant for plastic-related effluents.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Arts. L 69–L 70: Hospital wastewater discharge authorisation by Sanitation Minister. "
            f"Precautions include {q('le stockage et la récupération des produits chimiques de laboratoire', 'storage and recovery of laboratory chemical products')}, "
            f"{q('l élimination des huiles et hydrocarbures... Les huiles usagées doivent être stockées et récupérées par une entreprise agréée', 'elimination of oils and hydrocarbons... Used oils must be stored and recovered by an approved company')}, "
            f"{q('détergents biodégradables à 90% au moins', 'detergents biodegradable at least 90%')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 70: separate network if sewer is separative; amalgam separator for dental clinics; "
            f"buffer basin for laundry water >30°C; radioactive/radiology products recovered by approved company. "
            f"Art. L 71: Environment Minister consults Health and Sanitation Ministers before setting norms."
        ),
        "comments": bi(
            "Eaux hospitalières (déchets plastiques/contenants médicaux); lien avec Arrêté 09311/2007 (huiles usagées).",
            "Hospital wastewater (plastic/medical containers); link to Order 09311/2007 (used oils).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Arts. L 72–L 73: {q('Les conditions de rejet des eaux usées épurées en milieu naturel, obéissent aux normes en vigueur... notamment... la norme NS 05-061', 'Conditions for discharge of treated wastewater into natural environment comply with applicable standards... notably NS 05-061')}. "
            f"{q('Des taxes et redevances de rejet d eau en milieu naturel sont perçues', 'Taxes and fees for discharge into natural environment are collected')} per Environmental Code."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 74: decrees set conditions for use of treated water and treatment residues. "
            f"Art. L 78: treatment plants must meet ICPE standards (no nuisance odour, fumes, unauthorised soil infiltration)."
        ),
        "comments": bi(
            "Normes de rejet en milieu naturel via NS 05-061; MES peut capturer particules plastiques en suspension.",
            "Natural environment discharge standards via NS 05-061; TSS may capture suspended plastic particles.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Arts. L 75–L 76: {q('Les eaux d origine domestique peuvent, après traitement, être utilisées à des fins agricoles et maraîchères', 'Domestic-origin water may, after treatment, be used for agricultural and market gardening purposes')}. "
            f"WHO quality criteria: {q('coliformes fécaux... inférieure ou égale à 1.000 UFC/100 ml', 'faecal coliforms... less than or equal to 1,000 CFU/100 ml')} (restricted irrigation); "
            f"{q('œufs de nématodes... inférieure ou égale à un œuf viable/litre', 'nematode eggs... less than or equal to one viable egg/litre')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. L 77: other reuse types require purification rates set by ministerial order. "
            f"Joint order of Sanitation, Health, Agriculture and Environment Ministers for irrigation characteristics (Art. L 75)."
        ),
        "comments": bi(
            "Réutilisation des eaux épurées; instrument de circularité de l'eau (pas directement plastique).",
            "Reuse of treated water; water circularity instrument (not directly plastic-specific).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Arts. L 79–L 80: {q('Les déchargements et déversements de matières issues de vidange de fosses septiques... sont interdits', 'Discharge of septic tank sludge... is prohibited')} except in sealed tanks, treatment plants or déposantes. "
            f"Art. L 80: prohibited in déposantes — {q('de déchets ménagers, même après broyage préalable', 'household waste, even after prior shredding')}, "
            f"{q('d ordures ménagères, même après broyage préalable', 'household refuse, even after prior shredding')}, "
            f"{q('de déchets industriels', 'industrial waste')}, "
            f"{q('de déchets d activités de soins', 'healthcare activity waste')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 79: septage transport by trucks approved by Sanitation Minister. "
            f"Art. L 82: emptying companies approved by sanitation service. "
            f"Art. L 88: fee per tonne/m³ at managed déposantes."
        ),
        "comments": bi(
            "Interdit déchets ménagers/plastiques (même broyés) dans déposantes; complète L 29 pour filière boues.",
            "Prohibits household/plastic waste (even shredded) in déposantes; complements L 29 for sludge chain.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Arts. L 39–L 46: Stormwater management — {q('Tout lieu public ou privé urbanisé doit disposer d un système de collecte et d évacuation des eaux pluviales', 'Every urbanised public or private place must have a stormwater collection and evacuation system')} preventing stagnation. "
            f"Art. L 46: public stormwater works receive only rainwater (except unitary networks)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. L 41: private land within 30 m of public storm network must connect. "
            f"Art. L 50: natural runoff servitude; prohibited to discharge polluted stormwater to lower plots. "
            f"Art. L 28: storm drains/gutters prohibited on separate domestic wastewater systems."
        ),
        "comments": bi(
            "Gestion des eaux pluviales; réduit transport de déchets solides/plastiques vers caniveaux et milieux récepteurs.",
            "Stormwater management; reduces transport of solid/plastic waste to gutters and receiving environments.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Arts. L 92–L 96: Autonomous sanitation — {q('Les rejets d eaux usées des ouvrages d assainissement autonome... doivent faire l objet d un traitement préalable', 'Wastewater from autonomous sanitation works... must undergo prior treatment')}; "
            f"septic tank alone insufficient; discharge prohibited. "
            f"Owners must maintain installations; sanitation service may intervene at owner expense."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 92: 6-month connection deadline when public network becomes available; "
            f"autonomous works decommissioned at owner expense. "
            f"Art. L 97: fees for design, construction and operation control of autonomous installations."
        ),
        "comments": bi(
            "Assainissement autonome en zones non collectées; pertinent pour zones périurbaines à forte pollution plastique.",
            "Autonomous sanitation in non-networked areas; relevant for peri-urban zones with high plastic pollution.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. L 60: {q('Les prestations... pour assurer la collecte et l épuration des eaux usées industrielles sont payées... au moyen d une redevance dont le montant est calculé sur la base du volume d eau rejeté et de la qualité des effluents', 'Services for collection and treatment of industrial wastewater are paid via a fee calculated on discharged volume and effluent quality')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. L 63: special financial contribution if discharge requires additional network/treatment equipment. "
            f"Fee set by interministerial order (Sanitation, Environment Ministers). "
            f"Art. L 19: non-connected owners liable for sewer taxes."
        ),
        "comments": bi(
            "Instrument économique (redevance pollution-volume); incite industries (incl. plasturgie) à réduire rejets.",
            "Economic instrument (pollution-volume fee); incentivises industries (incl. plastics) to reduce discharges.",
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
