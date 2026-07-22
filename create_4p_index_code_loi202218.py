#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2022-18 SONAGED (bilingual FR/EN cells).

Source: user upload (JO extract) / FAOLEX sen216236.pdf (text-extractable).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2022-18_SONAGED.xlsx"

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
        "Loi n° 2022-18 du 23 mai 2022 autorisant la création de la Société nationale de gestion intégrée des déchets (SONAGED S.A.)",
        "Law No. 2022-18 of 23 May 2022 authorising creation of the National Integrated Waste Management Company (SONAGED S.A.)",
    ),
    "policy_url": "https://www.fao.org/faolex/results/details/en/c/LEX-FAOC216236",
    "policy_year": 2022,
    "policy_objective": bi(
        "Autoriser la création de SONAGED S.A., société anonyme d'État remplaçant l'UCG pour coordonner la gestion intégrée "
        "des déchets solides à l'échelle nationale — collecte, transport, traitement, mise en décharge et valorisation — "
        "intégrant PROMOGED et autres programmes publics, en tant que régulateur du secteur et promoteur de l'économie "
        "circulaire et de l'intercommunalité, sans mesures spécifiques aux plastiques.",
        "Authorise creation of SONAGED S.A., a State-owned company replacing UCG to coordinate integrated solid-waste "
        "management nationally — collection, transport, treatment, landfilling and valorisation — integrating PROMOGED "
        "and other public programmes, as sector regulator and promoter of circular economy and inter-municipal cooperation, "
        "without plastic-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Art. 2: nationwide {q('collecte, transport, mise en décharge, traitement et valorisation des déchets solides', 'collection, transport, landfilling, treatment and valorisation of solid waste')}; "
        f"Art. 5: transfer of UCG and all public integrated solid-waste programmes (incl. PROMOGED); "
        f"Art. 2: {q('développer une économie circulaire et valoriser les déchets', 'develop a circular economy and valorise waste')}."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Loi adoptée par l'Assemblée nationale le 5 mai 2022 et promulguée le 23 mai 2022 en tant que Loi n° 2022-18. "
        "Il s'agit d'une loi parlementaire autorisant la création d'une société d'État (≠0.75).",
        "Law adopted by the National Assembly on 5 May 2022 and promulgated on 23 May 2022 as Law No. 2022-18. "
        "It is parliamentary legislation authorising creation of a State-owned company (≠0.75).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "gestion des déchets solides, hygiène publique, collectivités territoriales, secteur privé, environnement, économie circulaire",
        "solid waste management, public hygiene, local government, private sector, environment, circular economy",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "collecte, transport, traitement, valorisation, élimination, gouvernance",
        "collection, transport, treatment, valorisation, disposal, governance",
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        f"Art. 4: resources from State endowment, other authorised resources, and operating revenues to fund nationwide "
        f"collection and operating costs; Art. 3: State holds 100% of shares (capital may open to other public shareholders)."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('Il est autorisé la création d une société anonyme dénommée Société nationale de Gestion intégrée des Déchets (SONAGED S.A.)', 'Creation of a public limited company named National Integrated Waste Management Company (SONAGED S.A.) is authorised')}; "
            f"technical supervision by Public Hygiene Minister, financial supervision by Finance Minister."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Exposé des motifs: replaces institutional instability (7 successive structures since 1990s); published JO 4 June 2022."
        ),
        "comments": bi(
            "Création société d'État — cadre institutionnel national déchets solides.",
            "State-company creation — national solid-waste institutional framework.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 2: SONAGED mission — {q('assurer la coordination de la gestion intégrée des déchets solides sur l ensemble du territoire national', 'ensure coordination of integrated solid-waste management across national territory')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Autonomous entity with State and other public bodies as shareholders; replaces UCG coordination role."
        ),
        "comments": bi(
            "Mission centrale de coordination — pilier de gouvernance post-UCG.",
            "Central coordination mission — governance pillar post-UCG.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 2: SONAGED notably charged to {q('assurer la collecte, le transport, la mise en décharge, le traitement et la valorisation des déchets solides sur l ensemble du territoire national', 'ensure collection, transport, landfilling, treatment and valorisation of solid waste across national territory')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Covers full municipal solid-waste chain including plastics as part of solid-waste stream; nationwide scope."
        ),
        "comments": bi(
            "Chaîne complète déchets solides — pertinence directe pour plastiques dans les ordures ménagères.",
            "Full solid-waste chain — direct relevance for plastics in household waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Valorisation",
        "instrument_description": (
            f"Art. 2: manage {q('équipements et infrastructures de traitement et de valorisation des déchets', 'waste treatment and valorisation equipment and infrastructure')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Inherits PROMOGED infrastructure (sorting centres, landfills, CCV) per Art. 5 transfer."
        ),
        "comments": bi(
            "Gestion des infrastructures de valorisation — tri/recyclage incluant plastiques.",
            "Valorisation infrastructure management — sorting/recycling including plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 2: exercise {q('son autorité sur le secteur en qualité de régulateur', 'authority over the sector as regulator')} and contribute to visibility of State and partner actions in waste-sector development."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Also proposes governance reforms (institutional, regulatory, financial) to improve sector performance."
        ),
        "comments": bi(
            "Rôle régulateur sectoriel — cadre pour politiques déchets/plastiques.",
            "Sector regulator role — framework for waste/plastics policies.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 2: {q('améliorer la gestion des déchets solides sur toute la chaîne de valeur et de promouvoir une gestion intégrée dans toutes les communes en favorisant l intercommunalité', 'improve solid-waste management along the entire value chain and promote integrated management in all communes by fostering inter-municipal cooperation')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Builds on CGCT commune waste competences and PROMOGED inter-municipal coordination."
        ),
        "comments": bi(
            "Chaîne de valeur et intercommunalité — coordination locale des filières déchets.",
            "Value chain and inter-municipal cooperation — local waste-stream coordination.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Valorisation",
        "instrument_description": (
            f"Art. 2: {q('développer une économie circulaire et valoriser les déchets en tenant compte de la dimension socio-économique', 'develop a circular economy and valorise waste taking account of socio-economic dimension')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Aligns with Loi 2020-04 circular-economy orientation and PROMOGED valorisation centres."
        ),
        "comments": bi(
            "Économie circulaire explicite — valorisation des déchets plastiques dans le flux solide.",
            "Explicit circular economy — valorisation of plastic waste in solid stream.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 2: supervise {q('programmes et projets de l Etat en matière de gestion des déchets solides', 'State programmes and projects on solid-waste management')}; "
            f"create contextualised sustainable management systems and maximise private-sector involvement."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 5: UCG and all public integrated solid-waste programmes transferred to SONAGED (incl. PROMOGED, Zero Waste programme)."
        ),
        "comments": bi(
            "Intégration PROMOGED/UCG — continuité opérationnelle nationale.",
            "PROMOGED/UCG integration — national operational continuity.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 5: {q('Sont transférés à la SONAGED, l Unité de Coordination de la Gestion des déchets solides et les autres projets et programmes publics de gestion intégrée des déchets solides', 'UCG and other public integrated solid-waste management projects and programmes are transferred to SONAGED')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Exposé des motifs: UCG (arrêté 01048/2018) and PROMOGED (décret 2021-831) absorbed into single State company."
        ),
        "comments": bi(
            "Transfert UCG + PROMOGED — fusion des structures de gestion déchets.",
            "UCG + PROMOGED transfer — merger of waste-management structures.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 3: {q('L Etat détient la totalité des actions de la Société', 'The State holds all shares in the Company')} and may later open capital to other shareholders per statutes."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Other public bodies may join capital; private opening possible per future statutes."
        ),
        "comments": bi(
            "Société d'État à capital public — gouvernance actionnariale.",
            "State company with public shareholding — corporate governance.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 4: SONAGED resources include State endowment, other authorised resources and operating revenues "
            f"to {q('assurer la collecte des déchets sur tout le territoire national et de couvrir les charges de fonctionnement', 'ensure waste collection nationwide and cover operating costs')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Funding modalities detailed in statutes (Art. 6) approved by decree."
        ),
        "comments": bi(
            "Financement de la collecte nationale — ressources État et exploitation.",
            "Nationwide collection funding — State endowment and operating revenue.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 6: {q('L organisation et le fonctionnement de la SONAGED sont fixés par les statuts approuvés par décret', 'SONAGED organisation and operation are set by statutes approved by decree')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Implementing decree required; law executed as law of the State (promulgated 23 May 2022)."
        ),
        "comments": bi(
            "Modalités d'application par décret et statuts.",
            "Implementation modalities by decree and statutes.",
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
