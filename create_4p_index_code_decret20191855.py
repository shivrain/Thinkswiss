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
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret présidentiel du 7 novembre 2019 fixant les attributions ministérielles (remplace le décret n° 2019-975). "
        "Texte réglementaire d'organisation exécutive (≠1.0 loi). Abrogé/remplacé par des décrets de réorganisation gouvernementale ultérieurs.",
        "Presidential decree of 7 November 2019 setting ministerial powers (replaces Decree No. 2019-975). "
        "Executive organisational regulation (≠1.0 law). Repealed/replaced by later government-reorganisation decrees.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "waste management, municipalities, industry, fisheries, tourism, water, environment",
        "waste management, municipalities, industry, fisheries, tourism, water, environment",
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": bi(
        "disposal, environmental leakage",
        "disposal, environmental leakage",
    ),
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. 1: MEDD {q('prépare et met en œuvre ladite politique en matière de veille environnementale, de lutte contre les pollutions et de protection de la nature, de la faune et de la flore', 'prepares and implements national policy on environmental monitoring, pollution control and protection of nature, fauna and flora')} "
            f"under Presidential authority."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD explicitly designated to prepare and implement national environmental policy. "
            f"No fines/penalties or monitoring requirements in this attributions decree. Superseded by later cabinet-reform attribution decrees."
        ),
        "comments": bi(
            "Gouvernance ministérielle — P=0.20; abrogé par réorganisations gouvernementales ultérieures.",
            "Ministerial governance — P=0.20; superseded by later government reorganisations.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. 1: {q('responsable, sous réserve des compétences dévolues aux collectivités territoriales, de la protection de l environnement', 'responsible, subject to powers of territorial collectivities, for environmental protection')}; "
            f"takes measures to {q('prévenir et lutter contre les pollutions de toute nature', 'prevent and combat pollution of all kinds')}; "
            f"ensures safety of potentially polluting installations."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD designated as responsible authority for environmental protection and pollution prevention. "
            f"Subject to local-government competences (reservation reduces unconditionality). No enforcement provisions in decree."
        ),
        "comments": bi(
            "Prévention pollutions — cadre indirect pollution plastique; T=0.25 (autorité désignée seulement).",
            "Pollution prevention — indirect plastic-pollution framework; T=0.25 (designated authority only).",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 1: for {q('établissements classés', 'classified facilities')}, MEDD conducts dossier review and signs individual acts in coordination with concerned ministries."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD designated to instruct dossiers and sign individual acts for classified facilities. "
            f"Operational permitting via subordinate ICPE framework (e.g. Décret 2001-282); no enforcement in this decree."
        ),
        "comments": bi(
            "Établissements classés — plasturgie/recyclage indirectement concernés; P=0.20 coordination, pas réglementation.",
            "Classified facilities — plastics manufacturing/recycling indirectly covered; P=0.20 coordination, not regulation.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Art. 1: {q('aide les collectivités territoriales à faire face à la collecte des déchets et il en assure le traitement', 'helps territorial collectivities cope with waste collection and ensures its treatment')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD designated to support municipalities on waste collection and ensure treatment. "
            f"Multi-level waste governance (national–municipal); support role under CGCT Art. 205. No enforcement provisions."
        ),
        "comments": bi(
            "Instrument le plus pertinent plastiques — coordination déchets collectivités; passe filtre pertinence (c).",
            "Most plastics-relevant instrument — municipal waste coordination; passes relevance filter (c).",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Art. 1: {q('appuie les initiatives des collectivités territoriales et des mouvements associatifs en matière d environnement', 'supports initiatives of territorial collectivities and civil-society movements on environment')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD designated to support local-government and civil-society environmental initiatives. "
            f"Enables communal waste/environment programmes; no enforcement or monitoring in decree."
        ),
        "comments": bi(
            "Appui initiatives locales — cadre programmes déchets/plastiques communaux; pertinence borderline.",
            "Local initiative support — framework for communal waste/plastic programmes; borderline relevance.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            f"Art. 1: charged with {q('développement de l éducation environnementale', 'development of environmental education')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. 1: MEDD designated to develop environmental education. "
            f"Information/awareness mandate; no enforcement, monitoring or unconditional binding requirements in decree."
        ),
        "comments": bi(
            "Éducation environnementale — P=0.40 information & voluntary; sensibilisation déchets plastiques borderline.",
            "Environmental education — P=0.40 information & voluntary; plastic-waste awareness borderline.",
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
