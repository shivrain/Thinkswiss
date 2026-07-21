#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Arrêté 07022/2009 POLMAR (bilingual FR/EN)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Arrete_07022_POLMAR.xlsx"

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
        "Arrêté n° 07022 du 16 juillet 2009 portant organisation et fonctionnement du Plan national de lutte contre la pollution marine (POLMAR)",
        "Order No. 07022 of 16 July 2009 on organisation and operation of the National Marine Pollution Response Plan (POLMAR)",
    ),
    "policy_url": "https://www.hassmar.gouv.sn/sites/default/files/reglementations/PLAN%20POLMAR.pdf",
    "policy_year": 2009,
    "policy_objective": bi(
        "Organiser le système national de préparation et d'intervention contre la pollution marine (hydrocarbures, "
        "produits chimiques/HNS, déchets récupérés) via HASSMAR, MRCC/RSC, plans sectoriels et classification Tier — "
        "couvrir indirectement les déchets plastiques marins et la détritus côtiers via MARPOL 73/78, la réception des "
        "déchets d'exploitation des navires et le nettoyage/restauration du littoral, sans mesures spécifiques aux plastiques.",
        "Organise the national marine pollution preparedness and response system (oil, chemicals/HNS, recovered waste) "
        "through HASSMAR, MRCC/RSC, sectoral plans and Tier classification — indirectly covering marine plastic waste "
        "and coastal litter through MARPOL 73/78, ship operational waste reception and shoreline cleaning/restoration, "
        "without plastic-specific measures.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('Tier 1 : déversement... inférieur à 7 tonnes', 'Tier 1: spill... less than 7 tonnes')}; "
        f"{q('Tier 2 :... compris entre 7 et 700 tonnes', 'Tier 2:... between 7 and 700 tonnes')}; "
        f"{q('Tier 3 :... supérieur à 700 tonnes', 'Tier 3:... greater than 700 tonnes')} of pollutant products in national waters."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Arrêté du Premier Ministre adopté le 16 juillet 2009, sur proposition du Ministre d'État, Ministre des "
        "Forces armées, opérationnalisant le décret n° 2006-322 (HASSMAR) et le PNIUM. Il s'agit d'un plan "
        "national et d'un texte réglementaire d'application exécutive, et non d'une loi parlementaire (≠1).",
        "Prime Minister's order adopted on 16 July 2009, on proposal of the Minister of State, Minister of the Armed "
        "Forces, operationalising Decree No. 2006-322 (HASSMAR) and PNIUM. It is a national plan and executive "
        "implementing regulation, not parliamentary legislation (≠1).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "pêche, transport maritime, environnement, littoral, industrie pétrolière, collectivités locales",
        "fisheries, maritime transport, environment, coast, oil industry, local government",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "utilisation, élimination, milieu marin",
        "use, disposal, marine environment",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Art. 80: {q('Le financement... est assuré par un fonds POLMAR', 'Financing... is provided by a POLMAR fund')}; "
        f"Art. 77: {q('Les frais de cette assistance sont entièrement à la charge du navire et/ou de son armateur', 'Assistance costs are entirely borne by the ship and/or shipowner')}; "
        f"Art. 77: State preserves indemnification rights for pollution damage."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 2: {q('La Pollution marine : l introduction physique ou chimique par l homme... de substances ou d énergies dans l environnement', 'Marine pollution: physical or chemical introduction by man... of substances or energy into the environment')} causing harm to health, biological resources, ecosystems or legitimate uses. "
            f"{q('Les Produits polluants : les hydrocarbures et les produits chimiques', 'Pollutant products: hydrocarbons and chemical products')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Preamble cites MARPOL 73/78, OPRC 90, OPRC-HNS 2000, Abidjan Convention, UNCLOS. "
            f"Art. 4: POLMAR is reference framework for hydrocarbon and chemical pollution in national waters."
        ),
        "comments": bi(
            "Définition large incluant introduction physique de substances; MARPOL 73/78 (Annexe V déchets/plastiques navires) cité en préambule.",
            "Broad definition including physical introduction of substances; MARPOL 73/78 (Annex V ship garbage/plastics) cited in preamble.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 3–6: POLMAR objectives — {q('réduire les risques de pollution marine à un niveau aussi faible que possible', 'reduce marine pollution risks to as low a level as possible')}; "
            f"{q('identifier les risques, l impact probable de la pollution et les priorités de protection', 'identify risks, probable pollution impact and protection priorities')}; "
            f"{q('Le Plan POLMAR privilégie la prévention', 'The POLMAR Plan prioritises prevention')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 72: {q('La prévention des incidents de pollution marine constitue l objectif majeur du Plan POLMAR', 'Prevention of marine pollution incidents is the major objective of Plan POLMAR')}. "
            f"Art. 73: prevention measures include ratifying conventions, ship/installation inspections, sensitivity atlas."
        ),
        "comments": bi(
            "Stratégie prévention-lutte; pertinent pour déchets plastiques marins en tant que pollution physique.",
            "Prevention-response strategy; relevant for marine plastic waste as physical pollution.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 7: Mandatory reporting — captains, shipowners, port authorities, industrial/responsible parties must report without delay to MRCC/RSC: "
            f"{q('tout évènement... qui entraîne ou risque d entraîner un rejet de produits polluants', 'any event... causing or risking release of pollutant products')}; "
            f"{q('toute présence de produits polluants dans les eaux sous juridiction nationale', 'any presence of pollutant products in national waters')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 81–82: POLMAR watch via MRCC, RSC, associated centres and alert posts; immediate reporting required. "
            f"Art. 8: Plan POLMAR triggered by Prime Minister's order in emergency."
        ),
        "comments": bi(
            "Obligation de signalement; base pour détection déversements huile/chimique/déchets en mer.",
            "Reporting obligation; basis for detecting oil/chemical/waste spills at sea.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 11–16: Mandatory sectoral plans — "
            f"{q('Les autorités portuaires doivent développer pour chaque port, un plan de prévention et de lutte', 'Port authorities must develop a prevention and response plan for each port')} with adequate decontamination means; "
            f"{q('Les Collectivités locales sur le littoral élaborent et mettent en œuvre des mesures pour prévenir les pollutions marines', 'Coastal local authorities develop and implement measures to prevent marine pollution')}; "
            f"industry, ships, marine protected areas must have sectoral POLMAR plans (Arts. 13–15)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 16: sectoral plans (Arts. 11, 12, 13, 15) submitted for approval to National Coordinator and annexed to POLMAR. "
            f"Art. 12: local authorities acquire Tier 1 or 2 response means according to sensitivity."
        ),
        "comments": bi(
            "Plans locaux/portuaires pour pollution marine; cadre pour gestion détritus côtiers et déchets portuaires.",
            "Local/port plans for marine pollution; framework for coastal litter and port waste management.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 17–28: Institutional framework — HASSMAR Secretary-General as National Coordinator; "
            f"{q('le Centre national de coordination de la lutte contre la pollution marine', 'national marine pollution response coordination centre')} (MRCC Dakar); "
            f"RSC secondary centres in three maritime zones (North, Centre, South); associated coastal centres and alert posts."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 29: National Coordinator approves decontamination companies, maintains early warning system. "
            f"Arts. 31–32: National and local coordination committees activated in emergencies. "
            f"Art. 26: intervention personnel from military, public, private, volunteer and local authorities."
        ),
        "comments": bi(
            "Architecture institutionnelle HASSMAR/MRCC; complète CENPOLMAR (2021) pour gouvernance pollution marine.",
            "HASSMAR/MRCC institutional architecture; complements CENPOLMAR (2021) for marine pollution governance.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 29: National Coordinator responsibilities include {q('l agrément des sociétés de dépollution', 'approval of decontamination companies')} and "
            f"{q('l élaboration et la mise en place d un système national d information et d alerte précoce', 'establishment of a national early warning information system')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 29: also coordinates national policy, training programmes, communication/sensitisation plan, international cooperation. "
            f"Art. 36 (Transport/Pêche): ship technical inspections, dangerous goods transport standards."
        ),
        "comments": bi(
            "Agrément des sociétés de dépollution; pertinent pour opérateurs de nettoyage littoral/déchets marins.",
            "Approval of decontamination companies; relevant for shoreline/marine waste cleanup operators.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 69–71: Three-tier spill classification — "
            f"{q('Tier 1... inférieur à 7 tonnes', 'Tier 1... less than 7 tonnes')}; "
            f"{q('Tier 2... compris entre 7 et 700 tonnes', 'Tier 2... between 7 and 700 tonnes')}; "
            f"{q('Tier 3... supérieur à 700 tonnes', 'Tier 3... greater than 700 tonnes')} of pollutant products. "
            f"Tier 3 triggers national POLMAR by National Coordinator (Art. 94)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Art. 68: graduated operational posture at three Tier levels. "
            f"Art. 94: national trigger when zonal response insufficient, pollution affects ≥2 zones, or Tier 3 magnitude. "
            f"Art. 8: Prime Minister's order for emergency trigger/lift."
        ),
        "comments": bi(
            "Seuils quantitatifs (tonnes) pour escalade opérationnelle; applicables aux hydrocarbures/chimiques, pas spécifiquement plastiques.",
            "Quantitative thresholds (tonnes) for operational escalation; applicable to oil/chemicals, not plastic-specific.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 73: Prevention measures including "
            f"{q('ratifier les conventions maritimes pertinentes', 'ratify relevant maritime conventions')}; "
            f"{q('veiller à l application par les navires... de la réglementation en vigueur relative à la protection de l environnement marin', 'ensure ships apply environmental protection regulations')}; "
            f"{q('inspecter les installations et les industries du secteur maritime', 'inspect maritime sector installations and industries')}; "
            f"{q('constituer et prépositionner des stocks de niveau 1 ou 2 à proximité des sites sensibles', 'pre-position Tier 1 or 2 stocks near sensitive sites')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 36: port ballast water management and {q('réception des déchets d exploitation', 'reception of operational waste')} from ships. "
            f"Art. 37 (Environment): control of maritime installations, marine protected area stocks, waste treatment coordination."
        ),
        "comments": bi(
            "Réception déchets d'exploitation navires (MARPOL V); mesure indirecte contre déchets plastiques en mer.",
            "Ship operational waste reception (MARPOL V); indirect measure against plastic waste at sea.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 84: Response modes include "
            f"{q('récupérer les produits polluants déversés', 'recover spilled pollutant products')}; "
            f"{q('nettoyer et restaurer le littoral', 'clean and restore the shoreline')}; "
            f"{q('traiter les déchets', 'treat waste')}; "
            f"protect sensitive marine/coastal resources; use dispersants per national legislation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 83: maritime response (MRCC) from high sea to low-water mark; terrestrial response (Regional Governor) on beaches and rivers. "
            f"Arts. 85–95: escalation from sectoral → initial → reinforced → national POLMAR."
        ),
        "comments": bi(
            "Nettoyage/restauration du littoral et traitement des déchets; pertinent pour détritus plastiques côtiers.",
            "Shoreline cleaning/restoration and waste treatment; relevant for coastal plastic litter.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 79: {q('Le Ministère chargé de l Environnement identifie... les lieux de stockage provisoire de produits polluants et de déchets pollués', 'Environment Ministry identifies... provisional storage sites for pollutant products and polluted waste')}; "
            f"final elimination and environmental impact monitoring subject to regulatory text."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 37: Environment Ministry coordinates {q('le traitement des déchets et à la restauration des sites pollués', 'waste treatment and restoration of polluted sites')}. "
            f"Arts. 40, 46, 57, 61: coastal storage of recovered pollutants and waste at ports/industry."
        ),
        "comments": bi(
            "Gestion des déchets récupérés après opération POLMAR (incl. matériaux pollués/plastiques absorbants).",
            "Management of waste recovered after POLMAR operations (incl. polluted/absorbent plastic materials).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 86–95: Graduated response — sectoral intervention by responsible structure; initial intervention by MRCC/RSC; "
            f"reinforced intervention by Delegate; national POLMAR by National Coordinator. "
            f"Art. 86: polluter must lead fight per sectoral plan and inform MRCC/RSC immediately."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 88–89: Operations Coordinator and On-Scene Coordinator (OSC) roles defined. "
            f"Art. 88: may requisition reinforcement means; operational control of all public/private means. "
            f"Art. 91–92: reinforced intervention if initial response insufficient."
        ),
        "comments": bi(
            "Procédure d'escalade opérationnelle; pollueur responsable de la première intervention.",
            "Operational escalation procedure; polluter responsible for first response.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 77: {q('les services compétents de l Etat peuvent exiger l assistance du navire... Les frais... sont entièrement à la charge du navire et/ou de son armateur', 'competent State services may require ship assistance... costs entirely borne by ship and/or shipowner')}; "
            f"State preserves indemnification rights; detailed register of actions, personnel and materials kept."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 77: injunction measures from confirmation of pollution danger. "
            f"Arts. 93, 96: local/national coordination committees handle indemnification dossiers. "
            f"Art. 44 (Justice Ministry): State litigation on marine pollution."
        ),
        "comments": bi(
            "Principe pollueur-payeur; sanctions civiles/pénales complétées par Code environnement et Marine marchande.",
            "Polluter-pays principle; civil/criminal sanctions complemented by Environmental and Merchant Marine Codes.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 80: {q('Le financement des dépenses afférentes à la mise en œuvre du Plan POLMAR est assuré par un fonds POLMAR', 'Financing of POLMAR implementation costs is provided by a POLMAR fund')}; "
            f"modalities defined by regulation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 29: National Coordinator mobilises resources for plan operation. "
            f"Art. 63 (indirect): special financial participation for ICPE discharges affecting sanitation network. "
            f"Art. 88: requisition of public/private reinforcement means."
        ),
        "comments": bi(
            "Fonds POLMAR pour financement interventions; complète redevances pollueur (Art. 77).",
            "POLMAR fund for intervention financing; complements polluter charges (Art. 77).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Arts. 74–76: Support activities — HASSMAR training plan for all stakeholders; "
            f"regular exercises for notification/alert/crisis networks (periodicity set by Prime Minister instruction); "
            f"{q('plan de communication destiné... à sensibiliser les populations sur la pollution marine', 'communication plan to raise public awareness on marine pollution')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 101: HASSMAR centralises pollution incident statistics; exercises and quality reviews enable corrective updates. "
            f"Art. 78: HASSMAR revises POLMAR based on exercises and incidents. "
            f"Arts. 97–100: end-of-operations reporting and plan lift by Prime Minister."
        ),
        "comments": bi(
            "Formation, exercices et sensibilisation; base pour prévention déchets marins plastiques.",
            "Training, exercises and awareness; basis for marine plastic waste prevention.",
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
