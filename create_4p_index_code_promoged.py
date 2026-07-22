#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for PROMOGED (bilingual FR/EN cells).

Sources:
- JO n° 7392 du 16 janvier 2021 / Arrêté ministériel n° 027932 du 11 décembre 2020 (user upload)
- World Bank P161477 Senegal Municipal Solid Waste Management Project documents
- PROMOGED project website (promoged.sn)
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_PROMOGED_Dechets_Solides.xlsx"

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
        "Projet de Promotion de la Gestion intégrée et de l'Économie des Déchets solides (PROMOGED)",
        "Solid Waste Integrated Management and Economy Promotion Project (PROMOGED)",
    ),
    "policy_url": (
        "https://www.urbanisme.gouv.sn/actualites/r%C3%A9union-du-comit%C3%A9-de-pilotage-du-promoged"
    ),
    "policy_year": 2021,
    "policy_objective": bi(
        "Projet national (Banque mondiale P161477, AFD, BEI, coopération espagnole) pour renforcer la gouvernance "
        "des déchets solides et moderniser les services de collecte, tri, transfert, traitement et élimination dans "
        "les agglomérations sélectionnées — développer les filières de valorisation (dont plastiques ~9 % du flux "
        "à Mbeubeuss), réhabiliter les décharges (Mbeubeuss, Mbao) et instaurer des PPP et financement basé sur "
        "les résultats (REP). Cadre juridique : Arrêté 027932/2020 (JO 7392 du 16/01/2021).",
        "National project (World Bank P161477, AFD, EIB, Spanish cooperation) to strengthen solid-waste governance "
        "and modernise collection, sorting, transfer, treatment and disposal services in selected agglomerations — "
        "develop valorisation value chains (including plastics ~9% of Mbeubeuss stream), rehabilitate dumpsites "
        "(Mbeubeuss, Mbao) and establish PPPs and results-based financing (RBF). Legal framework: Order 027932/2020 "
        "(Official Gazette 7392 of 16/01/2021).",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"PDO (WB P161477): {q('strengthen the governance of solid waste management in Senegal and improve solid waste management services in selected municipalities', 'strengthen SWM governance and improve SWM services in selected municipalities')}; "
        f"Phase I: 350 infrastructures, 15 landfills rehabilitated, 7 regions / 148 communes (promoged.sn); "
        f"WB indicators: 10 SW collection/transfer facilities; sanitary landfills operational."
    ),
    "policy_type": 0.50,
    "policy_type_justification": bi(
        "Programme/projet national de développement financé par bailleurs (≠1.0 loi). Arrêté ministériel n° 027932/2020 "
        "crée le comité de pilotage (≠1.0); instrument de gouvernance projet.",
        "National development programme/project financed by donors (≠1.0 law). Ministerial Order No. 027932/2020 "
        "creates steering committee (≠1.0); project governance instrument.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "hygiène urbaine, gestion des déchets solides, collectivités territoriales, secteur privé, environnement, "
        "économie circulaire, agriculture (compost)",
        "urban hygiene, solid waste management, local government, private sector, environment, circular economy, "
        "agriculture (compost)",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "collecte, tri, transfert, valorisation, recyclage, élimination, gouvernance",
        "collection, sorting, transfer, valorisation, recycling, disposal, governance",
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        f"Total ~FCFA 206 billion Phase I (WB, AFD, Spanish cooperation, EIB per promoged.sn); "
        f"WB credit US$175M (P161477): Component 1 US$20M governance; Component 2 US$256.1M infrastructure; "
        f"Component 3 US$18.6M implementation; RBF/REP disbursement linked to reform indicators."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"WB P161477 PDO: {q('strengthen the governance of solid waste management in Senegal and improve solid waste management services in selected municipalities', 'strengthen SWM governance and improve SWM services in selected municipalities')}; "
            f"implemented by Ministry of Urban Planning/Housing/Public Hygiene via UCG (later SONAGED)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"WB Board approval March 2020; Phase I 2020–2026; 7 regions (Dakar, Thiès, Saint-Louis, Matam, Kolda, Sédhiou, Ziguinchor)."
        ),
        "comments": bi(
            "Projet phare déchets solides — successeur opérationnel de l'UCG, intégré à SONAGED (2022).",
            "Flagship solid-waste project — operational successor to UCG, integrated into SONAGED (2022).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Arrêté 027932 Art. 1: creates {q('Comité de pilotage du Projet de Promotion de la Gestion intégrée et de l Économie des Déchets solides (PROMOGED)', 'PROMOGED steering committee')} "
            f"within Ministry of Urban Planning, Housing and Public Hygiene; published JO n° 7392 du 16 janvier 2021."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 9: effective from signature date; COPIL dissolved automatically at project closure; secretariat by UCG Coordinator."
        ),
        "comments": bi(
            "Instrument juridique officiel — JO 7392 / arrêté 027932.",
            "Official legal instrument — Official Gazette 7392 / Order 027932.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Arrêté 027932 Art. 2: COPIL missions — orient/supervise implementation; approve annual budgeted work plan (PTBA) and progress reports; "
            f"review Disbursement-Linked Indicators (ILDs); examine RBF fund allocation (REP); review performance evaluations, "
            f"eligible-expenditure audits and annual account audits."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: chaired by Urban Planning Minister; members include Finance, Economy, Health, Local Government, Environment, AMS, CONGAD; meets twice yearly."
        ),
        "comments": bi(
            "Gouvernance interministérielle — financement basé sur les résultats.",
            "Inter-ministerial governance — results-based financing.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Arrêté 027932 Arts. 4–6: Technical Committee assists COPIL; examines strategic studies on "
            f"{q('développement du secteur des déchets solides ménagers', 'household solid-waste sector development')}; "
            f"chaired by UCG Coordinator with representatives from Finance, Interior, Economy, Health, Local Government, Environment, MULHP."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Secretariat by Project Director; reports actions to COPIL."
        ),
        "comments": bi(
            "Comité technique — planification sectorielle déchets ménagers.",
            "Technical committee — household waste sector planning.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Arrêté 027932 Art. 7: four thematic technical groups on institutional aspects, financing mechanisms, "
            f"{q('Partenariat public privé (PPP)', 'public-private partnership (PPP)')} and technical/environmental/social aspects; "
            f"composition set by Technical Committee decision."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"WB Component 1: regulatory/institutional reform, PPP framework, inter-municipal coordination for SWM."
        ),
        "comments": bi(
            "Groupes thématiques PPP/financement — cadre pour secteur privé.",
            "PPP/financing thematic groups — framework for private sector.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"WB Component 1 (US$20M): strengthen sector governance and institutional capacity — regulatory/institutional "
            f"framework reform, financing mechanisms, PPP framework, inter-municipal SWM coordination; "
            f"RBF disbursed on achievement of reform indicators (PBCs)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"WB indicator: inter-municipal SWM coordination framework improved; 10% of annual O&M budget identified in inter-municipal plans spent."
        ),
        "comments": bi(
            "Composante 1 — réforme institutionnelle et REP.",
            "Component 1 — institutional reform and RBF.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"WB Component 2 (US$256.1M): improve solid-waste infrastructure and services in Greater Dakar and secondary poles "
            f"(Thiès/Mbour/Tivaouane; Saint-Louis/Matam; Ziguinchor/Kolda/Sédhiou) — collection points, standardized collection, "
            f"sorting/transfer centres, collection and marketing centres, sanitary landfills."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"WB targets: 10 collection/transfer facilities; sanitary landfills constructed and operational; "
            f"Phase I: ~350 infrastructures across 148 communes (promoged.sn)."
        ),
        "comments": bi(
            "Infrastructures collecte/tri/transfert — chaîne complète déchets solides incluant plastiques.",
            "Collection/sorting/transfer infrastructure — full solid-waste chain including plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Valorisation",
        "instrument_description": (
            f"PROMOGED Component 2: integrated waste management along value chain "
            f"(pre-collection, collection, recycling, transfer, treatment, disposal); "
            f"Collection and Marketing Centres (CCV) and Integrated Waste Recovery Centres for sorted materials."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"promoged.sn: 4 CCV per pole; sorting enables recovery channels; objective to make waste use profitable."
        ),
        "comments": bi(
            "Filières de valorisation — pertinence directe pour déchets plastiques triés.",
            "Valorisation value chains — direct relevance for sorted plastic waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Mbeubeuss/Mbao rehabilitation (Component 2): Sorting and Transfer Centre (CTT) and composting platform; "
            f"gradual dumpsite resorption; sorted products to recovery channels, sort refusals to landfill."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"ESIA: waste stream ~9% plastic at Mbeubeuss; existing plastic-press activities; 15 landfills to be rehabilitated Phase I."
        ),
        "comments": bi(
            "Réhabilitation Mbeubeuss — tri des plastiques et autres fractions avant enfouissement.",
            "Mbeubeuss rehabilitation — sorting plastics and other fractions before landfilling.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Component 1 PPP reform: establish sector-specific PPP framework with guidance and standard bidding documents; "
            f"clarify roles/responsibilities of actors in municipal waste management to enable private investment."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Objective: create enabling environment for private sector in solid-waste sector; maximum feasible private-sector involvement."
        ),
        "comments": bi(
            "PPP déchets solides — délégation collecte/traitement au privé.",
            "Solid-waste PPPs — private collection/treatment delegation.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"WB Component 3 (US$18.6M): project implementation support — coordination, communication, monitoring and evaluation, "
            f"fiduciary management; UCG provides COPIL secretariat per Arrêté 027932."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"WB indicator: 800 waste pickers supported (target 1,500 by 2026); M&E of intermediate results indicators."
        ),
        "comments": bi(
            "Appui à la mise en œuvre — coordination UCG/PROMOGED.",
            "Implementation support — UCG/PROMOGED coordination.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Social inclusion: support for waste pickers (récupérateurs) — WB target 800 (2024) rising to 1,500 (2026); "
            f"Mbeubeuss ESIA notes pickers using plastic bags and plastic-press activities in informal recovery economy."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Component 2 aims to improve socio-economic performance of sector through profitable waste use."
        ),
        "comments": bi(
            "Inclusion des récupérateurs — lien avec économie informelle du plastique.",
            "Waste-picker inclusion — link to informal plastic recovery economy.",
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
