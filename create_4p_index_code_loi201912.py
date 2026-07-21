#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2019-12 CGCT amendment (bilingual FR/EN).

Source: official dri.gouv.sn PDF / user upload (image-based, OCR via ocrmypdf).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2019-12_CGCT.xlsx"

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
        "Loi n° 2019-12 du 8 juillet 2019 portant modification du Code Général des Collectivités Territoriales",
        "Law No. 2019-12 of 8 July 2019 amending the General Code of Local Authorities",
    ),
    "policy_url": "https://www.dri.gouv.sn/sites/default/files/LOI/LOI%202019/L-2019-12.pdf",
    "policy_year": 2019,
    "policy_objective": bi(
        "Harmoniser le CGCT (Loi 2013-10) avec la réforme fiscale de la Loi 2018-10 en remplaçant la "
        "« contribution des patentes » par la « contribution économique locale » (CEL), en refondant les "
        "recettes de fonctionnement des villes et communes (Arts 185 et 195) et en instaurant un mécanisme "
        "de répartition nationale de la CEL sur la valeur ajoutée — avec maintien des taxes locales "
        "d'enlèvement des ordures ménagères et de balayage finançant les services de collecte/assainissement, "
        "sans mesures spécifiques aux plastiques.",
        "Harmonise the CGCT (Law 2013-10) with the Law 2018-10 tax reform by replacing the business "
        "patent contribution with the local economic contribution (CEL), refounding city and commune "
        "operating revenues (Arts 185 and 195) and establishing a national redistribution mechanism for "
        "CEL on value added — while retaining local household-waste collection and street-sweeping taxes "
        "funding collection/sanitation services, without plastic-specific measures.",
    ),
    "policy_target": 0.50,
    "policy_target_text": (
        f"Art. 195(b): {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')}, "
        f"{q('taxe de balayage', 'street-sweeping tax')}; "
        f"Art. 185(b): city {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')}; "
        f"Art. 195 bis: CEL redistribution with minimum allocation of FCFA 12,000,000/commune/year."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Loi d'amendement adoptée par l'Assemblée nationale le 29 juin 2019 et promulguée le 8 juillet 2019. "
        "Il s'agit d'une loi parlementaire modifiant le Code (≠0.75 arrêté/décret).",
        "Amending law adopted by the National Assembly on 29 June 2019 and promulgated on 8 July 2019. "
        "It is parliamentary legislation amending the Code (≠0.75 order/decree).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "collectivités territoriales, finances locales, fiscalité, assainissement, propreté urbaine, gouvernance locale",
        "local government, local finance, taxation, sanitation, urban cleanliness, local governance",
    ),
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": bi(
        "collecte, élimination, gouvernance",
        "collection, disposal, governance",
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        f"Art. 195 bis: CEL valeur ajoutée split into three redistribution guichets (allocation minimale ≥FCFA 12M/commune, "
        f"stabilisation ≤70%, équité territoriale ≥30%); Arts 185/195: State treasury advances of 25% of prior-year "
        f"direct-tax collections per quarter; waste-collection and sweeping taxes retained as commune revenue sources."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. premier: {q('Les articles 185 et 195 de la loi n°2013-10... sont abrogés et remplacés', 'Articles 185 and 195 of Law No. 2013-10... are repealed and replaced')} "
            f"to align CGCT with Law 2018-10 replacing {q('contribution des patentes', 'business patent contribution')} with "
            f"{q('contribution économique locale', 'local economic contribution')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Exposé des motifs: corrects omissions on contribution globale unique and contribution globale foncière; "
            f"adopted 29 June 2019, promulgated 8 July 2019."
        ),
        "comments": bi(
            "Loi d'harmonisation fiscale CGCT — pas de nouvelle compétence déchets, refonte des recettes locales.",
            "CGCT fiscal harmonisation law — no new waste competence, refounds local revenue framework.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 185(1)(b): city operating revenues include {q('Les produits de la taxe d enlèvement des ordures ménagères', 'proceeds of the household waste collection tax')}; "
            f"tax created by city council deliberation within legal maxima (Title V, Book I CGCT)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 185(1)(a): city also receives CEL and built-property land tax; State grants 25% quarterly advance on prior-year direct-tax collections."
        ),
        "comments": bi(
            "Taxe ordures ménagères au niveau ville — finance la collecte des déchets (incluant fractions plastiques).",
            "City-level household waste tax — funds waste collection (including plastic fractions).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 195(1)(b): commune direct local taxes include {q('taxe d enlèvement des ordures ménagères', 'household waste collection tax')}, "
            f"{q('taxe de balayage', 'street-sweeping tax')} and {q('taxe de déversement à l égout', 'sewer discharge tax')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Taxes created by municipal council deliberation; assessment/collection modalities and maxima set by law."
        ),
        "comments": bi(
            "Instruments fiscaux locaux finançant collecte et propreté — pertinents pour déchets solides/plastiques urbains.",
            "Local fiscal instruments funding collection and cleanliness — relevant for urban solid/plastic waste.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195(1)(a): commune direct taxes now include {q('contribution économique locale', 'local economic contribution')} "
            f"(replacing contribution des patentes), minimum fiscal tax, built/unbuilt land taxes, licence contribution, "
            f"CGU and CGF commune shares."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"State grants 25% quarterly advance on prior-year direct-tax collections for commune treasury."
        ),
        "comments": bi(
            "Réforme fiscale locale — ressources générales pouvant financer les services communaux dont les déchets.",
            "Local tax reform — general resources that may fund communal services including waste.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 185 bis: {q('Le produit de la contribution sur la valeur locative des locaux professionnels, de la contribution foncière sur les propriétés bâties et de la taxe d enlèvement des ordures ménagères visées à l article 185 n est perçu au profit de la ville que si les biens imposables sont situés en dehors des infrastructures et équipements marchands', 'Proceeds of professional-premises rental-value contribution, built-property land tax and household waste collection tax under Art. 185 are collected for the city only if taxable property is located outside commercial infrastructure and facilities')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Commercial infrastructure defined as markets, fairs, shopping centres and any place for commerce "
            f"whose administrative/financial management falls to the commune."
        ),
        "comments": bi(
            "Exception fiscale pour infrastructures marchandes — clarifie le périmètre de la taxe ordures en ville.",
            "Fiscal exception for commercial infrastructure — clarifies city waste-tax perimeter.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195 bis: {q('La contribution sur la valeur locative des locaux professionnels est perçue au profit de la commune sur le territoire de laquelle les locaux d exploitation sont situés', 'Professional-premises rental-value contribution is collected for the commune where the business premises are located')}, "
            f"subject to Art. 185 bis."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Territoriality rule links business premises to commune revenue; complements CEL valeur ajoutée national redistribution."
        ),
        "comments": bi(
            "Affectation territoriale de la CEL — renforce les ressources des communes gestionnaires des déchets.",
            "Territorial allocation of CEL — strengthens resources of waste-managing communes.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195 bis: CEL valeur ajoutée proceeds deposited in Treasury account {q('contribution économique locale / valeur ajoutée', 'local economic contribution / value added')} "
            f"and redistributed annually to all communes via three guichets: allocation minimale, stabilisation, équité territoriale."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Distribution modalities specified by decree; national equitable sharing per Art. 339 Loi 2018-10."
        ),
        "comments": bi(
            "Mécanisme de péréquation financière intercommunale — finance indirecte des services publics locaux.",
            "Inter-communal financial equalisation mechanism — indirect funding of local public services.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195 bis(1): guichet {q('allocation minimale', 'minimum allocation')} — fixed by joint Finance/Local Government Ministerial order; "
            f"not below FCFA 12,000,000 per commune per year for authorised personnel expenditure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Funded first before stabilisation and équité territoriale guichets."
        ),
        "comments": bi(
            "Plancher de ressources communales — soutient la capacité administrative d'exécution des compétences déchets.",
            "Commune resource floor — supports administrative capacity to deliver waste competences.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195 bis(2): guichet {q('stabilisation', 'stabilisation')} — restores difference between former patent-contribution receipts and new "
            f"professional-premises rental-value contribution; resources ≤70% of distributable CEL valeur ajoutée (after minimum allocations); 4-year period (extendable)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"If insufficient resources, proportional stabilisation by regulation; prorogation possible by Local Government Minister."
        ),
        "comments": bi(
            "Transition fiscale sur 4 ans — évite la chute brutale des recettes communales post-réforme.",
            "4-year fiscal transition — avoids abrupt post-reform drop in commune revenues.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 195 bis(3): guichet {q('équité territoriale', 'territorial equity')} — distributed among communes based on population and poverty index; "
            f"resources ≥30% of distributable CEL valeur ajoutée (after minimum allocations)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Funded after allocation minimale guichet; targets poorer/less-resourced communes."
        ),
        "comments": bi(
            "Péréquation selon pauvreté — peut renforcer les communes déficitaires en services d'assainissement.",
            "Poverty-based equalisation — may strengthen communes lacking sanitation services.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 195(1)(c): commune indirect local taxes include electricity, water, advertising, night establishments, "
            f"abattoir/sanitary inspection, entertainment, furnished lodgings and fuel-dispenser taxes — general local revenue base."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Created by municipal council deliberation within legal maxima; complements direct taxes including waste-collection tax."
        ),
        "comments": bi(
            "Recettes indirectes communales — financement général des services locaux sans lien plastique direct.",
            "Commune indirect revenues — general local-service funding without direct plastic link.",
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
