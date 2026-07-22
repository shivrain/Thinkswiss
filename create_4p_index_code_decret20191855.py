#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2019-1855 MEDD attributions (bilingual FR/EN).

Source: FAOLEX sen204304.pdf (text-extractable, 2 pages).
Note: Superseded by later government-reorganisation attribution decrees (post-2019).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2019-1855_MEDD_Attributions.xlsx"

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
        "Décret n° 2019-1855 du 7 novembre 2019 relatif aux attributions du Ministre de l'Environnement et du Développement durable",
        "Decree No. 2019-1855 of 7 November 2019 on the powers of the Minister of Environment and Sustainable Development",
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/sen204304.pdf",
    "policy_year": 2019,
    "policy_objective": bi(
        "Définir les attributions du Ministre de l'Environnement et du Développement durable (MEDD) pour préparer et mettre en œuvre "
        "la politique nationale en matière de veille environnementale, lutte contre les pollutions, protection de la nature, faune et flore, "
        "établissements classés, aires protégées, lutte contre la désertification et appui aux collectivités territoriales pour la collecte "
        "et le traitement des déchets — mandat institutionnel transversal sans mesures spécifiques aux plastiques.",
        "Define powers of the Minister of Environment and Sustainable Development (MEDD) to prepare and implement national policy on "
        "environmental monitoring, pollution control, nature/fauna/flora protection, classified facilities, protected areas, desertification "
        "control and support to territorial collectivities for waste collection and treatment — cross-cutting institutional mandate without "
        "plastic-specific measures.",
    ),
    "policy_target": 0.50,
    "policy_target_text": (
        f"Art. 1: {q('veille environnementale, de lutte contre les pollutions et de protection de la nature', 'environmental monitoring, pollution control and nature protection')}; "
        f"Art. 1: {q('aide les collectivités territoriales à faire face à la collecte des déchets et il en assure le traitement', 'helps territorial collectivities with waste collection and ensures its treatment')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret présidentiel du 7 novembre 2019 fixant les attributions ministérielles (remplace le décret n° 2019-975). "
        "Texte réglementaire d'organisation exécutive (≠1.0 loi). Abrogé/remplacé par des décrets de réorganisation gouvernementale ultérieurs.",
        "Presidential decree of 7 November 2019 setting ministerial powers (replaces Decree No. 2019-975). "
        "Executive organisational regulation (≠1.0 law). Repealed/replaced by later government-reorganisation decrees.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "environnement, développement durable, pollution, biodiversité, forêts, chasse, climat, collectivités territoriales, déchets, écotourisme",
        "environment, sustainable development, pollution, biodiversity, forests, hunting, climate, local government, waste, ecotourism",
    ),
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": bi(
        "collecte, traitement, gouvernance",
        "collection, treatment, governance",
    ),
    "policy_budget": 0,
    "policy_budget_text": (
        f"No budget allocation in attributions decree; ministerial powers exercised through MEDD budget and programmes "
        f"(DEEC, DREEC, SNDD secretariat, etc.)."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: MEDD {q('prépare et met en œuvre ladite politique en matière de veille environnementale, de lutte contre les pollutions et de protection de la nature, de la faune et de la flore', 'prepares and implements national policy on environmental monitoring, pollution control and protection of nature, fauna and flora')} "
            f"under Presidential authority."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 2: repeals Décret n° 2019-975 of 29 May 2019; Art. 3: MEDD executes decree; published JO. Superseded by later cabinet-reform attribution decrees."
        ),
        "comments": bi(
            "Cadre attributions MEDD 2019 — abrogé par réorganisations gouvernementales ultérieures.",
            "MEDD 2019 attributions framework — superseded by later government reorganisations.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('responsable, sous réserve des compétences dévolues aux collectivités territoriales, de la protection de l environnement', 'responsible, subject to powers of territorial collectivities, for environmental protection')}; "
            f"takes measures to {q('prévenir et lutter contre les pollutions de toute nature', 'prevent and combat pollution of all kinds')}; "
            f"ensures safety of potentially polluting installations."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ensures polluting activities do not compromise living environment and environmental quality; complements CGCT local waste competences."
        ),
        "comments": bi(
            "Protection environnement et pollutions — cadre indirect pour pollution plastique.",
            "Environmental protection and pollution — indirect framework for plastic pollution.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: for {q('établissements classés', 'classified facilities')}, MEDD conducts dossier review and signs individual acts in coordination with concerned ministries."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Links to ICPE/regulatory framework (e.g. Décret 2001-282, NS standards); DEEC/DREEC implementation."
        ),
        "comments": bi(
            "Établissements classés — industries plasturgie/recyclage soumises à autorisation.",
            "Classified facilities — plastics manufacturing/recycling industries subject to permitting.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 1: {q('aide les collectivités territoriales à faire face à la collecte des déchets et il en assure le traitement', 'helps territorial collectivities cope with waste collection and ensures its treatment')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Support role to communes/cities under CGCT Art. 205 waste competence; does not replace local authority responsibilities."
        ),
        "comments": bi(
            "Mandat déchets le plus pertinent pour plastiques — appui collecte/traitement aux communes.",
            "Most waste-relevant mandate for plastics — support to communes for collection/treatment.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('appuie les initiatives des collectivités territoriales et des mouvements associatifs en matière d environnement', 'supports initiatives of territorial collectivities and civil-society movements on environment')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Enables local environmental/waste programmes and NGO partnerships."
        ),
        "comments": bi(
            "Appui initiatives locales — cadre pour programmes déchets/plastiques communaux.",
            "Local initiative support — framework for communal waste/plastic programmes.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: charged with {q('développement de l éducation environnementale', 'development of environmental education')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Supports awareness on pollution, waste and sustainable consumption."
        ),
        "comments": bi(
            "Éducation environnementale — sensibilisation déchets plastiques.",
            "Environmental education — plastic-waste awareness.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: manages {q('mécanisme de veille et de suivi des tendances de changement de climat et de modification de l état de l environnement', 'monitoring mechanism for climate-change trends and environmental state changes')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Supports national environmental reporting and climate policy (CDN, SNDD)."
        ),
        "comments": bi(
            "Veille climat/environnement — suivi état de l'environnement.",
            "Climate/environment monitoring — environmental state tracking.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: preservation of fauna and flora; protects waterways from aquatic plant invasion; authority over "
            f"{q('parcs nationaux et autres aires protégées', 'national parks and other protected areas')} with public access and high protection."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Also: marine flora protection, coastal/estuary erosion; hunting legislation; endangered species and anti-poaching."
        ),
        "comments": bi(
            "Biodiversité et aires protégées — cadre indirect milieux marins/littoraux.",
            "Biodiversity and protected areas — indirect marine/coastal framework.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('lutte contre la désertification et celle contre les feux de brousse', 'combat desertification and bush fires')}; "
            f"applies soil protection and regeneration policy; promotes forestry economy with territorial collectivities, reforestation and rational forest use."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Water/soil conservation via retention basins and artificial lakes; aquaculture development with Fisheries/Agriculture ministries."
        ),
        "comments": bi(
            "Forêts et sols — hors périmètre plastique direct.",
            "Forests and soils — outside direct plastic scope.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: represents Senegal at international technical meetings on {q('protection de l environnement, au développement durable, au climat et à la biodiversité', 'environment protection, sustainable development, climate and biodiversity')}; "
            f"supervises Permanent Secretariat of National Commission on Sustainable Development (SNDD)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Also: ecotourism development with concerned ministers; chairs Superior Council of Hunting and Fishing."
        ),
        "comments": bi(
            "Représentation internationale et SNDD — gouvernance DD.",
            "International representation and SNDD — SD governance.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 2: {q('Le décret n° 2019-975 du 29 mai 2019 relatif aux attributions du Ministre de l Environnement et du Développement durable est abrogé', 'Decree No. 2019-975 of 29 May 2019 on MEDD powers is repealed')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"2019-1855 was the operative MEDD attributions decree from 7 November 2019 until superseded by later reorganisation."
        ),
        "comments": bi(
            "Remplace décret 2019-975 — chaîne d'abrogation attributions MEDD.",
            "Replaces Decree 2019-975 — MEDD attributions repeal chain.",
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
