#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2016-1804 (bilingual FR/EN cells).

Source: official ditp.gouv.sn PDF / uploaded copy (text-extractable).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2016-1804_Code_Peche_Maritime.xlsx"

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
        "Décret n° 2016-1804 du 22 novembre 2016 portant application du Code de la Pêche maritime",
        "Decree No. 2016-1804 of 22 November 2016 implementing the Maritime Fisheries Code",
    ),
    "policy_url": "https://www.ditp.gouv.sn/download/file/fid/76",
    "policy_year": 2016,
    "policy_objective": bi(
        "Fixer les modalités d'application de la Loi n° 2015-18 (Code de la Pêche maritime), notamment "
        "les mesures de conservation des ressources halieutiques, le contrôle des engins et maillages "
        "des filets, le marquage des navires, la surveillance des opérations de pêche et la protection "
        "de l'environnement marin — pertinent pour les engins de pêche (souvent en polymères synthétiques) "
        "et la pollution marine, sans mesures spécifiques aux déchets plastiques.",
        "Set modalities for applying Law No. 2015-18 (Maritime Fisheries Code), notably resource "
        "conservation measures, fishing-gear and net-mesh control, vessel marking, fisheries-operation "
        "surveillance and marine environment protection — relevant to fishing gear (often synthetic polymers) "
        "and marine pollution, without plastic-waste-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Art. 24: authorised artisanal {q('filets maillants', 'gillnets')} with minimum mesh sizes; "
        f"Art. 35: ban on devices {q('permettant d obstruer les mailles du filet ou ayant pour effet de réduire leur action sélective', 'obstructing mesh or reducing selective action')}; "
        f"Art. 37: Minister may require selective gear to {q('garantir la préservation des ressources et de l environnement marins', 'ensure preservation of resources and the marine environment')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret d'application adopté le 22 novembre 2016 pour mettre en œuvre la Loi n° 2015-18. "
        "Il s'agit d'un texte réglementaire d'application exécutive (≠1.0), et non d'une loi parlementaire.",
        "Implementing decree adopted on 22 November 2016 to apply Law No. 2015-18. "
        "It is sub-legislative executive implementing regulation (≠1.0), not parliamentary legislation.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "pêche maritime, pêche artisanale, pêche industrielle, environnement marin, surveillance maritime",
        "maritime fisheries, artisanal fishing, industrial fishing, marine environment, maritime surveillance",
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": bi(
        "utilisation, élimination, environnement marin",
        "use, disposal, marine environment",
    ),
    "policy_budget": 0,
    "policy_budget_text": (
        f"No dedicated budget allocation in decree; licensing fees governed separately under Loi 2015-18 Art. 64. "
        f"Enforcement via fisheries surveillance agents and observer programme (Ch. 6)."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('Le présent décret a pour objet de fixer les modalités d application de la loi n° 2015-18 du 13 juillet 2015 portant Code de la Pêche maritime', 'This decree sets modalities for applying Law No. 2015-18 of 13 July 2015 establishing the Maritime Fisheries Code')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Art. 70: repeals and replaces Décret n° 98-498; Art. 71: execution by Fisheries, Environment, "
            f"Defence, Interior, Justice, Foreign Affairs and Finance Ministers."
        ),
        "comments": bi(
            "Décret d'application opérationnel du Code 2015-18.",
            "Operational implementing decree for the 2015-18 Code.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 10: fishing may be prohibited when {q('l embarcation n est pas immatriculée et marquée conformément aux règles prescrites', 'the vessel is not registered and marked as prescribed')} or "
            f"{q('les engins de pêche ne sont pas recensés selon les règles prescrites', 'fishing gear is not inventoried under prescribed rules')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Also applies when vessel fails safety standards or sustainable resource management requires; "
            f"modalities defined by ministerial order."
        ),
        "comments": bi(
            "Contrôle administratif des engins — traçabilité des filets et équipements.",
            "Administrative gear control — traceability of nets and equipment.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 24: lists authorised artisanal gear in Senegalese waters, including multiple "
            f"{q('filets maillants', 'gillnets')} types with minimum {q('maille étirée', 'stretched mesh')} and "
            f"{q('maille de côté', 'side mesh')} dimensions, plus lines and longlines."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 27–28: mesh measured wet with graduated rule; special conditions for coastal longlines, "
            f"shrimp set-nets, trammel nets and beach seines by ministerial order."
        ),
        "comments": bi(
            "Réglementation détaillée des filets (souvent en nylon/polyester) — complète l'interdiction des filets monofilaments de la Loi 2015-18 Art. 66.",
            "Detailed net regulation (often nylon/polyester) — complements Law 2015-18 Art. 66 monofilament ban.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 25: prohibits use of bottom longlines in estuaries/mangroves; beach seines in specified zones; "
            f"{q('les filets maillants dérivants pour la pêche à la crevette', 'drift gillnets for shrimp fishing')} in all Senegalese waters; "
            f"nets encircling rock fish habitats."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 26: any gear not listed in Art. 24 requires prior ministerial authorisation after advisory opinion."
        ),
        "comments": bi(
            "Interdiction de filets dérivants à crevette — réduit les engins fantômes et la capture non sélective.",
            "Drift gillnet ban for shrimp — reduces ghost gear and non-selective catch.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 35: {q('Il est interdit, pour tous types d engins de pêche, d employer des moyens ou dispositifs permettant d obstruer les mailles du filet ou ayant pour effet de réduire leur action sélective', 'For all fishing gear types, it is prohibited to use means or devices obstructing mesh or reducing selective action')} "
            f"(with limited trawl protection panels exception)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 31–34: detailed mesh-measurement procedures for enforcement agents (ICES gauge, 25-mesh series)."
        ),
        "comments": bi(
            "Interdiction des dispositifs de réduction de sélectivité — implémente Loi 2015-18 Art. 125(e).",
            "Ban on selectivity-reducing devices — implements Law 2015-18 Art. 125(e).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 36: prohibits in Senegalese waters: {q('le chalutage en bœuf', 'pair trawling')}, "
            f"{q('les filets maillants dérivants à thons', 'drift gillnets for tuna')}, "
            f"oversized shrimp trawlers and certain double-wire trawls."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 29–30: industrial minimum mesh sizes for trawls and surrounding nets also specified."
        ),
        "comments": bi(
            "Interdictions d'engins industriels destructeurs, incluant filets dérivants à thons.",
            "Prohibition of destructive industrial gear, including tuna drift gillnets.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 37: {q('Le Ministre chargé de la Pêche maritime est habilité à prendre les mesures nécessaires concernant l utilisation de tout dispositif ou gréement de nature à détruire les habitats naturels des espèces en vue de garantir la préservation des ressources et de l environnement marins', 'The Fisheries Minister may take necessary measures on use of any device or rigging that destroys natural habitats to ensure preservation of resources and the marine environment')}; "
            f"may promote or require selective gear for biodiversity and stock management."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 23: Minister may supplement or tighten conservation measures by order; "
            f"Art. 42: temporary zone closures possible."
        ),
        "comments": bi(
            "Protection explicite de l'environnement marin et des habitats — cadre pour mesures anti-pollution/engins.",
            "Explicit marine environment and habitat protection — framework for anti-pollution/gear measures.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 55: industrial and artisanal operators must report catch data including "
            f"{q('les engins et méthodes de pêche utilisés', 'fishing gear and methods used')}, zones, species and vessel characteristics "
            f"for resource management and effective surveillance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Reporting modalities set by ministerial order; supports monitoring of illegal or abandoned gear use."
        ),
        "comments": bi(
            "Déclaration des engins utilisés — outil de suivi des pratiques de pêche.",
            "Declaration of gear used — tool for monitoring fishing practices.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Arts. 56–58: fishing vessels must permanently display identification marks (radio call sign or "
            f"flag-state code + registry number); technical specifications for letter height, contrast and maintenance; "
            f"marks must not be {q('masquées par les engins de pêche', 'masked by fishing gear')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 58: owner/operator must maintain marks in good condition; detailed specifications by ministerial order."
        ),
        "comments": bi(
            "Marquage des navires — complète l'immatriculation des engins (Loi 2015-18 Art. 68); identification en mer.",
            "Vessel marking — complements gear registration (Law 2015-18 Art. 68); at-sea identification.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 4: {q('un représentant du Ministère chargé de l Environnement', 'a representative of the Ministry of Environment')} sits on the "
            f"Conseil national consultatif des Pêches maritimes, which advises on fisheries management and resource questions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: Council gives opinions on fishery management plans and major resource/sector questions."
        ),
        "comments": bi(
            "Intégration interministérielle environnement-pêche dans la gouvernance.",
            "Inter-ministerial environment-fisheries integration in governance.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 6 (conseils locaux): missions include {q('promouvoir les bonnes pratiques d hygiène, de salubrité et de conservation des produits de la pêche', 'promoting good hygiene, salubrity and conservation practices for fishery products')} "
            f"and {q('participer à l élaboration et à l exécution des plans d aménagement locaux des pêcheries', 'participating in local fishery management plans')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Local councils also prevent conflicts between fishing communities and support local MCS (monitoring, control, surveillance)."
        ),
        "comments": bi(
            "Gouvernance locale — bonnes pratiques et cogestion des pêcheries.",
            "Local governance — good practices and fisheries co-management.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 59–62 (observateurs): observers monitor fishing operations and collect data on "
            f"{q('engins, zones de pêche, nature des espèces capturées', 'gear, fishing zones, species caught')}; "
            f"commandants must allow inspection of {q('les engins de pêche à bord', 'fishing gear on board')} and filming of fishing activities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 66: prohibited to obstruct observers; industrial vessels must embark observers per Law 2015-18 Art. 72."
        ),
        "comments": bi(
            "Surveillance embarquée des engins et méthodes de pêche en mer.",
            "On-board surveillance of gear and fishing methods at sea.",
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
