#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2013-10 CGCT (bilingual FR/EN cells).

Source text extracted via OCR from the official dri.gouv.sn PDF (image-based, 68 pages).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2013-10_CGCT.xlsx"

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
        "Loi n° 2013-10 du 28 décembre 2013 portant Code Général des Collectivités Territoriales (CGCT)",
        "Law No. 2013-10 of 28 December 2013 establishing the General Code of Local Authorities (CGCT)",
    ),
    "policy_url": (
        "https://www.dri.gouv.sn/sites/default/files/an-documents/"
        "LOI%20N%202013%2010%20DU%2028%20DECEMBRE%202013.pdf"
    ),
    "policy_year": 2013,
    "policy_objective": bi(
        "Refondre le cadre de décentralisation en définissant les compétences des départements et communes "
        "(collectivités territoriales), incluant la propreté municipale, la gestion des déchets ménagers, "
        "l'hygiène et la salubrité publique — cadre subnational clé pour les déchets solides, dont les "
        "plastiques, sans mesures spécifiques aux matières plastiques.",
        "Refound the decentralisation framework by defining competences of departments and communes "
        "(local authorities), including municipal cleanliness, household waste management, hygiene and public "
        "health — key subnational framework for solid waste, including plastics, without plastic-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Commune competence: {q('la gestion des déchets et la lutte contre l insalubrité', 'waste management and the fight against insalubrity')}; "
        f"municipal police: {q('nettoiement', 'cleaning')}, {q('salubrité publics', 'public health/salubrity')}; "
        f"communal taxes: {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')}, {q('taxe de balayage', 'street-sweeping tax')}."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Code général adopté par l'Assemblée nationale le 19 décembre 2013 (procédure d'urgence) et promulgué "
        "le 28 décembre 2013 en tant que Loi n° 2013-10. Il s'agit d'une loi parlementaire (Code), et non d'un "
        "texte réglementaire d'application exécutive (≠0.75).",
        "General Code adopted by the National Assembly on 19 December 2013 (urgency procedure) and promulgated "
        "on 28 December 2013 as Law No. 2013-10. It is parliamentary legislation (Code), not executive "
        "implementing regulation (≠0.75).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "collectivités territoriales, assainissement, environnement, gouvernance locale, finances publiques locales",
        "local government, sanitation, environment, local governance, local public finance",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "collecte, élimination, gouvernance",
        "collection, disposal, governance",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Art. 195: {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')}, "
        f"{q('taxe de balayage', 'street-sweeping tax')}; "
        f"Art. 202: {q('dépenses d entretien et néttoiement des rues', 'expenditure on maintenance and cleaning of streets')} "
        f"and {q('dépenses des services locaux de désinfection et d hygiène', 'expenditure on local disinfection and hygiene services')} "
        f"are obligatory commune operating expenses."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 3: {q('Les collectivités locales ont pour mission la conception, la programmation et la mise en œuvre des actions de développement économique, social et environnemental d intérêt local', 'Local authorities are tasked with designing, programming and implementing local economic, social and environmental development actions')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 4: {q('La loi détermine les compétences des collectivités locales', 'The law determines local authority competences')}; "
            f"any competence transfer must be accompanied by concomitant transfer of resources and means by the State."
        ),
        "comments": bi(
            "Cadre de gouvernance locale intégrant l'environnement; base pour compétences déchets sans mention des plastiques.",
            "Local governance framework integrating environment; basis for waste competences without plastic mention.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 106 (11): {q('de veiller à la protection de l environnement, de prendre en conséquence les mesures propres, d une part, à empêcher ou à supprimer la pollution et les nuisances, d autre part, à assurer la protection des espaces verts et, enfin, à contribuer à l embellissement de la commune', 'to ensure environmental protection and accordingly take measures to prevent or eliminate pollution and nuisances, protect green spaces and contribute to beautifying the commune')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Mayor duties under council control (Art. 106); municipal police under Arts. 118–119; "
            f"State representative may intervene if mayor fails (Arts. 128–129)."
        ),
        "comments": bi(
            "Pouvoir exécutif communal sur pollution/nuisances — pertinent pour déchets solides et plastiques en espace public.",
            "Commune executive power over pollution/nuisances — relevant for solid waste and plastics in public space.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 119: {q('La police municipale a... pour objet d assurer le bon ordre, la sûreté, la tranquillité, la sécurité et la salubrité publics', 'Municipal police aims to ensure public order, safety, tranquillity, security and salubrity')}. "
            f"Mission 1 includes {q('le nettoiement, l éclairage, l enlèvement des encombrements', 'cleaning, lighting, removal of obstructions')} on public streets and places, and "
            f"{q('les modalités de mise en œuvre des missions relatives au nettoiement et à la salubrité', 'modalities for implementing cleaning and salubrity missions')} in capital-region local authorities by decree where needed."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 118: mayor responsible for municipal police under State representative control; "
            f"municipal police service created by decree. Art. 124: State may act if municipal authorities fail."
        ),
        "comments": bi(
            "Police municipale — nettoiement et salubrité publique; cadre direct pour collecte/enlèvement des déchets en voirie.",
            "Municipal police — public cleaning and salubrity; direct framework for waste collection/removal on roads.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 122: {q('Le maire peut prescrire... d entourer d une clôture suffisante les puits et les excavations présentant un danger pour la sécurité publique, ainsi que les terrains insalubres présentant un danger pour la santé publique', 'The mayor may require owners or occupiers to fence wells and excavations dangerous to public safety, as well as insalubrious land dangerous to public health')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 128–129: Minister may execute measures d'office if mayor refuses or neglects prescribed acts."
        ),
        "comments": bi(
            "Lutte contre l'insalubrité des terrains; complète la compétence déchets (Art. 205 Livre IV).",
            "Fight against insalubrious land; complements waste competence (Art. 205 Livre IV).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 126: {q('Les maires peuvent nommer des agents assermentés, chargés, sous le contrôle du service d hygiène, de fonctions relatives à la police sanitaire de la commune', 'Mayors may appoint sworn agents, under hygiene service supervision, for functions relating to municipal sanitary police')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Linked to Art. 202 (14): obligatory expenditure on {q('services locaux de désinfection et d hygiène', 'local disinfection and hygiene services')}."
        ),
        "comments": bi(
            "Mise en œuvre institutionnelle de l'hygiène communale.",
            "Institutional implementation of municipal hygiene.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 170 (ville): {q('la gestion des déchets et la lutte contre l insalubrité', 'waste management and the fight against insalubrity')} transferred to the city (ville) level alongside other urban competences."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Applies to cities (Dakar, Pikine, Guédiawaye, etc.) where waste competence sits above arrondissement communes; "
            f"Art. 171 (10): city mayor must ensure environmental protection and prevent pollution/nuisances."
        ),
        "comments": bi(
            "Compétence déchets au niveau ville dans les agglomérations urbaines; central pour gestion des déchets plastiques urbains.",
            "Waste competence at city level in urban agglomerations; central for urban plastic waste management.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 205 (Livre IV, Titre II, Ch. II, S. 2 — commune): {q('la gestion des déchets et la lutte contre l insalubrité', 'waste management and the fight against insalubrity')} among transferred commune competences on environment and natural resources "
            f"(alongside {q('l élaboration des plans communaux d action pour l environnement', 'development of communal environmental action plans')})."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 203: obligatory expenses from competence transfers; Art. 202 (11)–(14): obligatory cleaning and hygiene service costs. "
            f"Note: official PDF OCR reads {q('Article 305', 'Article 305')} for this provision; legal numbering is Art. 205 in Livre IV."
        ),
        "comments": bi(
            "Compétence explicite de gestion des déchets au niveau communal — pilier subnational pour déchets solides incluant plastiques.",
            "Explicit waste management competence at commune level — key subnational pillar for solid waste including plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 307 (commune, santé): {q('la mise en œuvre des mesures de prévention et d hygiène', 'implementation of prevention and hygiene measures')} among commune health competences."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Complements Art. 205 waste/insalubrity competence and Art. 119 municipal police salubrity missions."
        ),
        "comments": bi(
            "Hygiène communale liée à la salubrité publique et aux déchets.",
            "Communal hygiene linked to public salubrity and waste.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 195 (5): commune direct local taxes include {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')} and {q('taxe de balayage', 'street-sweeping tax')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Assessment and collection modalities set by law; centimes additionnels may be created by municipal council deliberation within legal maxima."
        ),
        "comments": bi(
            "Instrument fiscal local finançant la collecte des ordures ménagères (incluant fractions plastiques).",
            "Local fiscal instrument financing household waste collection (including plastic fractions).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 202 (11): obligatory commune operating expenditure on {q('les dépenses d entretien et néttoiement des rues, chemins de voirie et places publiques situés sur le territoire de la collectivité locale', 'maintenance and cleaning of streets, roads and public squares on local authority territory')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 201–202: obligatory operating expenses must appear in budget; Art. 202 lists 18 mandatory expenditure categories."
        ),
        "comments": bi(
            "Obligation budgétaire de nettoyage des voies publiques — enlèvement des déchets abandonnés incluant plastiques.",
            "Budgetary obligation to clean public roads — removal of litter including plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 202 (14): obligatory expenditure on {q('Les dépenses des services locaux de désinfection et d hygiène dans les conditions déterminées par la réglementation en vigueur', 'expenditure on local disinfection and hygiene services under applicable regulations')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Linked to Art. 126 sworn hygiene agents and Art. 307 commune prevention/hygiene measures."
        ),
        "comments": bi(
            "Financement obligatoire des services d'hygiène locale.",
            "Mandatory funding of local hygiene services.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 204 (Livre IV, département, environnement): department receives competences including "
            f"{q('l élaboration et mise en œuvre de plans départementaux d actions de l environnement', 'development and implementation of departmental environmental action plans')} and "
            f"{q('l élaboration et mise en œuvre des plans d action locale pour l environnement', 'development and implementation of local environmental action plans')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Department-level environmental planning complements commune waste management (Art. 205) and commune environmental action plans."
        ),
        "comments": bi(
            "Planification environnementale départementale encadrant les politiques locales de déchets.",
            "Departmental environmental planning framing local waste policies.",
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
