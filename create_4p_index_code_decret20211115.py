#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2021-1115 CENPOLMAR (bilingual FR/EN).

Source: JO n° 7470 du 13 novembre 2021 (text-extractable); official PDF on hassmar.gouv.sn is image-based.
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2021-1115_CENPOLMAR.xlsx"

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
        "Décret n° 2021-1115 du 25 août 2021 portant création du Centre national de Coordination de la Lutte contre la Pollution marine (CENPOLMAR)",
        "Decree No. 2021-1115 of 25 August 2021 establishing the National Marine Pollution Response Coordination Centre (CENPOLMAR)",
    ),
    "policy_url": "https://www.hassmar.gouv.sn/sites/default/files/reglementations/Decret%20CENPOLMAR_0.pdf",
    "policy_year": 2021,
    "policy_objective": bi(
        "Créer le CENPOLMAR au sein de la HASSMAR comme organe opérationnel de coordination de la lutte contre la pollution marine "
        "par hydrocarbures, avec veille environnementale permanente, appui aux plans POLMAR et POLMAR-TERRE, suivi du traitement des "
        "déchets issus de pollutions, formation, équipements, R&D et coopération internationale — renforce indirectement la gouvernance "
        "de la pollution côtière et marine (débris marins) sans mesures spécifiques aux plastiques.",
        "Establish CENPOLMAR within HASSMAR as the operational coordination body for marine hydrocarbon pollution response, with "
        "permanent environmental watch, support to POLMAR and POLMAR-TERRE plans, oversight of pollution-related waste treatment, "
        "training, equipment, R&D and international cooperation — indirectly strengthening coastal/marine pollution governance "
        "(marine debris) without plastic-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Art. 2: {q('organe opérationnel de la HASSMAR en matière de coordination de la lutte contre la pollution marine par hydrocarbures', 'HASSMAR operational body for coordinating marine hydrocarbon pollution response')}; "
        f"Art. 5: {q('participation à la surveillance, au contrôle et suivi du traitement des déchets résultant d une pollution marine par hydrocarbures', 'participation in surveillance, control and follow-up of waste treatment from hydrocarbon marine pollution')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret présidentiel du 25 août 2021 créant un centre opérationnel national au sein de la HASSMAR (texte réglementaire "
        "institutionnel/organisationnel, ≠1.0 loi). Publié au JO n° 7470 du 13 novembre 2021. Complète Arrêté 07022/2009 (POLMAR).",
        "Presidential decree of 25 August 2021 establishing a national operational centre within HASSMAR (institutional/organisational "
        "regulation, ≠1.0 law). Published in Official Gazette No. 7470 of 13 November 2021. Complements Order 07022/2009 (POLMAR).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "transport maritime, pêche, environnement, littoral, pétrole et énergies, sécurité maritime",
        "maritime transport, fisheries, environment, coast, oil and gas, maritime security",
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": bi(
        "utilisation, élimination, milieu marin",
        "use, disposal, marine environment",
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        f"Art. 12: CENPOLMAR resources from {q('budget de la HASSMAR', 'HASSMAR budget')}, "
        f"{q('subventions de partenaires nationaux ou internationaux', 'grants from national or international partners')}, "
        f"{q('prestations d acteurs du secteur privé national', 'services from national private-sector actors')}, "
        f"{q('financements, contributions ou subventions exceptionnelles', 'exceptional financing, contributions or grants')}, "
        f"{q('dons et legs', 'donations and bequests')} and {q('coopération internationale', 'international cooperation')}."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 1: creates {q('Centre national de Coordination de la Lutte contre la Pollution marine', 'National Marine Pollution Response Coordination Centre')} "
            f"({q('CENPOLMAR', 'CENPOLMAR')}) within HASSMAR. "
            f"Art. 2: operational body for {q('coordination de la lutte contre la pollution marine par hydrocarbures', 'coordinating marine hydrocarbon pollution response')}. "
            f"Art. 3: seat in Dakar (transferable by HASSMAR Secretary-General)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Preamble cites CNUDM, MARPOL, OPRC, Abidjan Convention, Décrets 2006-322 (HASSMAR), 2006-323 (PNIUM), 2001-282; "
            f"complements Arrêté 07022/2009 POLMAR. Published JO 7470."
        ),
        "comments": bi(
            "Institution opérationnelle HASSMAR — chaîne POLMAR/CENPOLMAR pour pollution marine.",
            "HASSMAR operational institution — POLMAR/CENPOLMAR chain for marine pollution.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 4: general mission of {q('veille, de préparation et d appui technique et opérationnel à la lutte contre la pollution marine par hydrocarbures', 'watch, preparedness and technical/operational support for marine hydrocarbon pollution response')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Addresses operational gap identified in preamble: national expertise to assess hydrocarbon pollution, recommend measures and support operations."
        ),
        "comments": bi(
            "Mission centrale veille/préparation/intervention — cadre hydrocarbures, pas plastiques explicites.",
            "Core watch/preparedness/response mission — hydrocarbon framework, no explicit plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('appui à la mise en œuvre du plan national de lutte contre la pollution marine (POLMAR)', 'support for implementation of the National Marine Pollution Response Plan (POLMAR)')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Operationalises Arrêté 07022/2009 POLMAR Tier system, MRCC/RSC coordination and sectoral plans."
        ),
        "comments": bi(
            "Lien direct avec POLMAR 2009 — déchets récupérés et MARPOL 73/78 en amont.",
            "Direct link to POLMAR 2009 — recovered waste and upstream MARPOL 73/78.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('appui à la mise en œuvre du plan national de lutte contre la pollution marine à terre (POLMAR-TERRE)', 'support for implementation of the National Shoreline Marine Pollution Response Plan (POLMAR-TERRE)')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Extends coordination to shoreline/land-based marine pollution response; relevant to coastal litter and debris cleanup operations."
        ),
        "comments": bi(
            "POLMAR-TERRE — littoral et débris côtiers (plastiques marins indirectement).",
            "POLMAR-TERRE — shoreline and coastal debris (marine plastics indirectly).",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('veille environnementale permanente en relation avec les diverses administrations compétentes', 'permanent environmental watch in liaison with relevant administrations')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Cross-ministerial monitoring mechanism; supports early detection of marine environmental incidents."
        ),
        "comments": bi(
            "Veille permanente — détection incidents pollution marine.",
            "Permanent watch — marine pollution incident detection.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 5: {q('participation à la surveillance, au contrôle et suivi du traitement des déchets résultant d une pollution marine par hydrocarbures', 'participation in surveillance, control and follow-up of waste treatment from hydrocarbon marine pollution')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Complements POLMAR Art. 37 (Environment Ministry waste treatment/restoration); covers recovered absorbent materials and polluted waste from spill operations."
        ),
        "comments": bi(
            "Provision la plus pertinente pour déchets — inclut matériaux absorbants/plastiques pollués post-intervention.",
            "Most waste-relevant provision — includes absorbent/polluted plastic materials post-intervention.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('formation, renforcement de capacités et tenue d une base de données des ressources humaines qualifiées POLMAR', 'training, capacity building and maintenance of a database of qualified POLMAR human resources')}; "
            f"{q('entrainement des personnels des administrations compétentes en mer à la lutte contre la pollution marine par hydrocarbures', 'training of competent administration personnel at sea for marine hydrocarbon pollution response')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Builds national POLMAR-qualified workforce; complements HASSMAR training plan under Arrêté 07022 Arts. 74–76."
        ),
        "comments": bi(
            "Renforcement capacités nationales — base pour prévention pollution marine.",
            "National capacity building — basis for marine pollution prevention.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('acquisition, entretien et suivi du matériel de lutte contre la pollution marine par hydrocarbures', 'acquisition, maintenance and monitoring of marine hydrocarbon pollution response equipment')}; "
            f"{q('stockage et positionnement du matériel de lutte contre la pollution marine par hydrocarbures', 'storage and positioning of marine hydrocarbon pollution response equipment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 9: secondary centres may be established in different maritime zones for equipment positioning."
        ),
        "comments": bi(
            "Équipements d'intervention — barrières, absorbants, matériel de confinement.",
            "Response equipment — booms, absorbents, containment materials.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 5: {q('recherche et développement en matière de pollution marine par hydrocarbures', 'research and development on marine hydrocarbon pollution')}; "
            f"{q('prévision et calcul des dérives des nappes d hydrocarbures', 'forecasting and calculation of oil-slick drift')}; "
            f"{q('fourniture de prestations payantes en rapport avec ses missions', 'provision of paid services related to its missions')}; "
            f"{q('coopération internationale en matière de lutte contre la pollution marine par hydrocarbures', 'international cooperation on marine hydrocarbon pollution response')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Preamble references international conventions (CNUDM, MARPOL, OPRC, Abidjan Convention); enables technical cooperation and modelling capacity."
        ),
        "comments": bi(
            "R&D et coopération internationale — alignement normes SSEC et conventions maritimes.",
            "R&D and international cooperation — SSEC standards and maritime conventions alignment.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Arts. 6–11: Director (naval officer or MEDD hierarchy-A specialist); technical adviser to HASSMAR SG; organigramme by HASSMAR SG; "
            f"secondary maritime centres (Art. 9); appointments and hierarchical authority (Art. 10); HR from detached civil/military/paramilitary staff or suspended agents; contractual recruitment possible (Art. 11)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Inter-ministerial staffing (Armed Forces, Interior, Fisheries, Environment, etc. per Art. 13); operational structure for 24/7 coordination."
        ),
        "comments": bi(
            "Gouvernance opérationnelle interministérielle — MEDD impliqué à la direction.",
            "Inter-ministerial operational governance — MEDD involved in leadership.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 12: CENPOLMAR funding from HASSMAR budget, national/international grants, private-sector services, exceptional contributions, donations/bequests and international cooperation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Diversified financing beyond State budget; paid services (Art. 5) provide additional revenue stream."
        ),
        "comments": bi(
            "Financement diversifié — complète fonds POLMAR (Arrêté 07022 Art. 80).",
            "Diversified financing — complements POLMAR fund (Order 07022 Art. 80).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 13: execution by Ministers of Presidency, Government Secretariat, Armed Forces, Interior, Finance, Fisheries and Maritime Economy, "
            f"Justice, Petroleum and Energy, Transport, and Environment and Sustainable Development; decree to be published in Official Gazette."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Signed Dakar 25 August 2021; published JO 7470 of 13 November 2021. CENPOLMAR operational within HASSMAR."
        ),
        "comments": bi(
            "Texte en vigueur — centre opérationnel actif au sein HASSMAR.",
            "Instrument in force — operational centre active within HASSMAR.",
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
