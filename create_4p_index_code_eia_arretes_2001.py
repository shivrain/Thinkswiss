#!/usr/bin/env python3
"""Generate 4P Index coding for EIA ministerial orders 9468-9472 (bilingual FR/EN)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Arretes_EIA_9468-9472_2001.xlsx"

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
        "Arrêtés ministériels du 28 novembre 2001 relatifs aux études d'impact environnemental "
        "(n° 9468–9472 MJEHP-DEEC)",
        "Ministerial orders of 28 November 2001 on environmental impact assessment "
        "(Nos. 9468–9472 MJEHP-DEEC)",
    ),
    "policy_url": "https://www.denv.gouv.sn/decrets/",
    "policy_year": 2001,
    "policy_objective": bi(
        "Opérationnaliser la procédure d'évaluation environnementale du Code de l'environnement et du "
        "Décret 2001-282 en réglementant la participation du public, l'organisation du comité technique, "
        "l'agrément des consultants, le contenu des termes de référence et des rapports d'EIE — assurant "
        "l'examen environnemental des projets industriels (dont usines plastiques et installations de "
        "traitement des déchets) sans mesures spécifiques aux plastiques.",
        "Operationalise the environmental assessment procedure under the Environmental Code and Decree "
        "2001-282 by regulating public participation, technical committee organisation, consultant "
        "accreditation, terms of reference and EIA report content — ensuring environmental review of "
        "industrial projects (including plastics plants and waste-treatment facilities) without "
        "plastic-specific measures.",
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Ensemble de cinq arrêtés ministériels (n° 9468–9472) adoptés le 28 novembre 2001 par le Ministre "
        "chargé de l'environnement (MJEHP-DEEC), publiés au Journal officiel n° 6025 du 12 janvier 2002. "
        "Il s'agit de textes réglementaires d'application exécutive, et non de lois adoptées par le "
        "Parlement.",
        "Package of five ministerial orders (Nos. 9468–9472) adopted on 28 November 2001 by the Minister "
        "of the Environment (MJEHP-DEEC), published in Official Journal No. 6025 of 12 January 2002. "
        "These are executive implementing regulations, not acts enacted by Parliament.",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "environnement, industrie, planification, agriculture, mines, énergie, élevage, urbanisme, commerce, "
        "forêts, travaux publics, santé, hydraulique, tourisme, municipalités, gestion des déchets",
        "environment, industry, planning, agriculture, mines, energy, livestock, urban planning, commerce, "
        "forestry, public works, health, water, tourism, municipalities, waste management",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "production, élimination, fuite environnementale",
        "production, disposal, environmental leakage",
    ),
    "policy_budget": 0,
    "policy_budget_text": "",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9468, Art. 1–8: {q('La participation publique est un élément constitutif de l étude d impact environnementale', 'Public participation is a constitutive element of the environmental impact study')} "
            f"with mandatory steps: announcement, document deposit at municipality, information meeting, "
            f"written/oral comments, negotiations, report. "
            f"Art. 6–7: {q('L audience publique se fera sur site au plus tard quinze jours après la validation interne du rapport', 'The public hearing shall take place on site no later than fifteen days after internal validation of the report')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 4: {q('L information du public est à la charge du promoteur', 'Public information is at the promoter s expense')} involving technical committee, host local authority and promoter. "
            f"Art. 5: local authority has {q('dix jours', 'ten days')} for written comments. "
            f"Art. 8: Environment Minister decision notified within {q('quinze jours', 'fifteen days')}. "
            f"Authorities: technical committee, DEEC, decentralised collectivity."
        ),
        "comments": bi(
            "Gouvernance multi-niveaux (État–collectivités–promoteur). Pertinent pour projets plastiques soumis à EIE; pas de mention explicite des plastiques.",
            "Multi-level governance (State–local authorities–promoter). Relevant for plastic projects subject to EIA; no explicit mention of plastics.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9469, Art. 1–6: Establishes the {q('Comité technique', 'Technical Committee')} as the administrative unit for EIA management under Décret 2001-282 Art. R43, with functions including "
            f"{q('administrer le processus d évaluation environnementale', 'administer the environmental assessment process')}, "
            f"{q('évaluer la qualité des rapports d étude d impact', 'assess the quality of impact study reports')}, and "
            f"{q('formuler un avis sur tous les projets assujettis à l étude d impact', 'issue an opinion on all projects subject to impact study')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 2: interministerial membership (Environment, Industry, Mines, Agriculture, Urban Planning, Health, etc.). "
            f"Art. 3: {q('validation interne des rapports', 'internal validation of reports')}; {q('dix jours', 'ten days')} to respond on study type. "
            f"Art. 5: secretariat may {q('inspecter périodiquement les sites des projets', 'periodically inspect project sites')}."
        ),
        "comments": bi(
            "Instrument de coordination intersectorielle pour l'examen des EIE de projets industriels et de gestion des déchets (Annexe Décret 2001-282 R42).",
            "Cross-sector coordination instrument for review of EIAs for industrial and waste-management projects (Decree 2001-282 Annex R42).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9470, Art. 1–7: {q('L agrément à l exercice des activités relatives aux études d impact environnemental peut être accordé', 'Accreditation to carry out environmental impact study activities may be granted')} to qualified physical or moral persons meeting diploma, experience and logistics requirements (Art. 4). "
            f"Art. 6: accreditation {q('peut être retiré', 'may be withdrawn')} for serious professional breaches or loss of required qualifications."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 1–2: accreditation commission chaired by Environment Minister examines dossiers. "
            f"Art. 4: minimum requirements — {q('diplôme supérieur (DEA-Doctorat)', 'higher degree (DEA-Doctorate)')}, EIA experience, logistics; or legal entity with {q('cinq experts de haut niveau', 'five senior experts')}. "
            f"Art. 6: withdrawal for {q('Manquement grave aux obligations professionnelles (qualité des travaux)', 'serious breach of professional obligations (quality of work)')}."
        ),
        "comments": bi(
            "Qualifie les bureaux d'études réalisant les EIE pour usines plastiques et installations de traitement des déchets. Agrément valable 5 ans (Art. 6).",
            "Qualifies consultancy firms conducting EIAs for plastics plants and waste-treatment facilities. Accreditation valid 5 years (Art. 6).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9471, Art. 1: {q('Les termes de référence de toute étude d impact sur l environnement doivent comprendre', 'The terms of reference for any environmental impact study must include')} 14 mandatory elements, including "
            f"item 4: {q('une évaluation des mesures envisagées pour l évacuation des eaux usées, l élimination des déchets solides et la réduction des émissions', 'an assessment of measures envisaged for wastewater evacuation, solid waste elimination and emission reduction')}. "
            f"Art. 2: DEEC may develop sector-specific ToR."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 1: mandatory ToR elements use {q('doivent comprendre', 'must include')} language. "
            f"Art. 3: {q('Le Directeur de l Environnement et des Etablissements classés... est chargé de l exécution', 'The Director of the Environment and Classified Establishments... is responsible for implementation')}. "
            f"Enforcement via EIA irreceivability (Arrêté 9472 Art. 3) and committee review (9469)."
        ),
        "comments": bi(
            "Point d'entrée direct pour l'évaluation des déchets solides et émissions dans les EIE de projets plastiques. Pas de seuils quantitatifs plastique.",
            "Direct entry point for assessing solid waste and emissions in EIAs for plastic projects. No quantitative plastic thresholds.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9472, Art. 1: {q('Le rapport d étude d impact environnement (REIE) est constitué de', 'The environmental impact study report (EIS report) shall comprise')} 16 mandatory sections including project description, baseline analysis, impact evaluation, accident risks, "
            f"preventive/mitigation measures, and {q('un cadre de plan de surveillance et de suivi de l environnement (PSE)', 'an environmental monitoring and follow-up plan framework (EMP)')}. "
            f"Art. 2: report must be in French, 10 copies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 1 §13: promoter must submit detailed EMP at pre-construction with cost, schedule and responsible structures. "
            f"Art. 3: {q('Tout rapport... qui ne satisfait pas aux dispositions... sera déclaré irrecevable', 'Any report... that does not comply... shall be declared inadmissible')}. "
            f"Review by Technical Committee (9469)."
        ),
        "comments": bi(
            "Cadre obligatoire du rapport EIE pour projets industriels; inclut suivi post-projet (PSE) pertinent pour impacts plastiques en exploitation.",
            "Mandatory EIA report framework for industrial projects; includes post-project monitoring (EMP) relevant to operational plastic impacts.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Arrêté n° 9472, Art. 3: {q('Tout rapport d une étude d impact environnemental, qui ne satisfait pas aux dispositions des articles précédemment cités sera déclaré irrecevable et la décision sera notifiée au promoteur pour qu il se conforme aux positions prévues par le présent arrêté', 'Any environmental impact study report that does not comply with the preceding provisions shall be declared inadmissible and the decision shall be notified to the promoter to comply with the positions set out in this order')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: binding non-compliance consequence ({q('déclaré irrecevable', 'declared inadmissible')}). "
            f"Art. 8 (9468): Minister decision on finalised EIA within 15 days. "
            f"Technical Committee validation prerequisite (9469 Art. 3)."
        ),
        "comments": bi(
            "Mécanisme d'application procédurale; empêche l'approbation de rapports EIE incomplets pour projets plastiques ou de traitement des déchets.",
            "Procedural enforcement mechanism; prevents approval of incomplete EIA reports for plastic or waste-treatment projects.",
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
        "A": 52, "B": 48, "C": 10, "D": 52, "E": 12, "F": 20,
        "G": 10, "H": 52, "I": 14, "J": 44, "K": 14, "L": 36,
        "M": 12, "N": 36, "O": 12, "P": 14, "Q": 22, "R": 60,
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
