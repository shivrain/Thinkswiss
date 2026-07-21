#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for LPSEDD 2016-2020 (bilingual FR/EN cells).

Source: official denv.gouv.sn PDF (OCR via ocrmypdf; pages 12-24 blank in official file;
Section III strategic framework recovered from FAOLEX mirror sen179622.pdf for coding).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Lettre_Politique_Environnement_2016-2020.xlsx"

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
        "Lettre de politique du secteur de l'environnement et du développement durable 2016–2020 (LPSEDD)",
        "Policy letter for the environment and sustainable development sector 2016–2020 (LPSEDD)",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/61/documentation/18323/"
        "lettre-de-politique-du-secteur-de-lenvironnement-et-du-developpement-durable-2016-2020.pdf"
    ),
    "policy_year": 2016,
    "policy_objective": bi(
        "Définir l'orientation stratégique du secteur environnement et développement durable pour 2016–2020 "
        "(vision, axes, programmes et lignes d'action), couvrant la gestion du cadre de vie, la lutte contre "
        "les pollutions et nuisances, la gouvernance environnementale verte, l'intégration transversale du DD "
        "dans le PSE et l'alignement ODD — avec mentions explicites de la prolifération des déchets plastiques, "
        "de la loi sur les sachets 0–30 microns et des opportunités de valorisation/recyclage des déchets plastiques.",
        "Define strategic orientation for the environment and sustainable development sector for 2016–2020 "
        "(vision, axes, programmes and action lines), covering living-environment management, pollution and "
        "nuisance control, green environmental governance, cross-cutting SD integration in the PSE and SDG "
        "alignment — with explicit references to plastic-waste proliferation, the 0–30 micron bag law and "
        "opportunities for plastic-waste valorisation/recycling.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"§II.2.2.1: {q('la prolifération des déchets plastiques est symptomatique... Son éradication demeure une priorité', 'plastic-waste proliferation is symptomatic... its eradication remains a priority')}; "
        f"§III.3.5.2 OS2: {q('améliorer la qualité du cadre de vie par une gestion rationnelle et concertée des pollutions et nuisances', 'improve living-environment quality through rational, concerted pollution and nuisance management')}; "
        f"§III Programme 3: {q('Lutte contre les pollutions, les nuisances et les effets néfastes des changements climatiques', 'Combat pollution, nuisances and adverse climate-change effects')}."
    ),
    "policy_type": 0.50,
    "policy_type_justification": bi(
        "Lettre de politique sectorielle validée en décembre 2015 pour la période 2016–2020. Document "
        "d'orientation stratégique et de planification ministérielle (≠1.0 loi, ≠0.75 arrêté/décret).",
        "Sector policy letter validated December 2015 for the 2016–2020 period. Strategic orientation and "
        "ministerial planning document (≠1.0 law, ≠0.75 order/decree).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "environnement, développement durable, cadre de vie, déchets, pollutions, climat, biodiversité, "
        "gouvernance, agriculture, pêche, industrie, mines, santé, collectivités locales",
        "environment, sustainable development, living environment, waste, pollution, climate, biodiversity, "
        "governance, agriculture, fisheries, industry, mining, health, local government",
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": bi(
        "production, consommation, collecte, traitement, valorisation, élimination, gouvernance",
        "production, consumption, collection, treatment, valorisation, disposal, governance",
    ),
    "policy_budget": 0.25,
    "policy_budget_text": (
        f"§II.2.3.1(c): MEDD budget fell from FCFA 31.25 billion (2011) to ~FCFA 22 billion (2015); "
        f"implementation via DPPD budget-programme framework and mobilisation of innovative financing; "
        f"no dedicated plastic-waste budget line."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Intro: LPSEDD 2016–2020 succeeds LPSERN 2009–2015; {q('formulée de façon consensuelle et participative', 'formulated consensually and participatively')} "
            f"with shared vision, common values, strategic axes and programmes linked to global and specific objectives."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Adopted December 2015 for 2016–2020 period; operationalised through DPPD and CSPE monitoring framework (§V)."
        ),
        "comments": bi(
            "Cadre stratégique sectoriel 2016–2020 — période expirée.",
            "Sector strategic framework 2016–2020 — period expired.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Intro/PSE: Government commits to {q('une trajectoire de développement sobre en carbone', 'a low-carbon development trajectory')} "
            f"and integration of sustainable-development principles into national policies to reverse natural-resource degradation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Aligned with PSE (2014–2035), national CDN and SNDD validated at 2015 National Conference on Sustainable Development."
        ),
        "comments": bi(
            "Trajectoire bas-carbone — cadre indirect pour réduction des émissions liées aux plastiques.",
            "Low-carbon trajectory — indirect framework for emissions linked to plastics.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"§II.2.2.1: Living-environment management requires {q('salubrité, de gestion adéquate des pollutions, des nuisances, des risques de catastrophes et d appui à la collecte et au traitement des déchets', 'salubrity, adequate pollution and nuisance management, disaster-risk management and support for waste collection and treatment')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Diagnostic notes weak sorting, collection, transport and valorisation performance and absence of filière approach."
        ),
        "comments": bi(
            "Orientation cadre de vie — collecte/traitement des déchets (incluant plastiques).",
            "Living-environment orientation — waste collection/treatment (including plastics).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§II.2.2.1: {q('la prolifération des déchets plastiques est symptomatique de la persistance des mauvaises pratiques de consommation, du faible niveau de préparation des structures en charge de la question, et de l absence d une véritable stratégie. Son éradication demeure une priorité', 'plastic-waste proliferation is symptomatic of persistent bad consumption practices, low preparedness of responsible structures and absence of a real strategy. Its eradication remains a priority')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Policy expects Loi banning 0–30 micron plastic bags to help; no standalone plastic-waste strategy in LPSEDD."
        ),
        "comments": bi(
            "Mention plastique la plus explicite — priorité d'éradication des déchets plastiques.",
            "Most explicit plastic mention — priority to eradicate plastic waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"§II.2.2.1: {q('l entrée en vigueur de la loi interdisant la fabrication, la distribution et l utilisation des sachets plastiques compris entre 0 et 30 microns devrait y aider', 'entry into force of the law prohibiting manufacture, distribution and use of plastic bags between 0 and 30 microns should help')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§II.2.3.1(b): {q('la loi sur les sachets plastiques a été votée', 'the plastic-bag law has been passed')} as part of legal reform; coastal law in adoption."
        ),
        "comments": bi(
            "Référence à la Loi 2015-09 (sachets fins) — lien direct production/distribution plastique.",
            "Reference to Law 2015-09 (thin bags) — direct plastic production/distribution link.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§II.2.2.1: Wild dumps, household, industrial and chemical waste and liquid discharges suffer from "
            f"{q('déficit d infrastructures de traitement performantes', 'lack of effective treatment infrastructure')}; "
            f"natural environments and communities exposed to poor air quality and health/disaster risks."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Notes weak national response capacity; no specific landfill or dumpsite closure targets in extracted text."
        ),
        "comments": bi(
            "Diagnostic dépôts sauvages — pertinent pour déchets plastiques non collectés.",
            "Wild-dump diagnostic — relevant for uncollected plastic waste.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Valorisation",
        "instrument_description": (
            f"§II.2.3.1(c): Opportunities in {q('la valorisation et du recyclage des déchets plastiques', 'valorisation and recycling of plastic waste')}, "
            f"ecotourism, forestry, non-timber products, renewables and sustainable agriculture to complement CSR."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Identified as economic opportunity; no quantified recycling targets or EPR framework in LPSEDD."
        ),
        "comments": bi(
            "Seule mention explicite de recyclage des déchets plastiques.",
            "Only explicit mention of plastic-waste recycling.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§II.2.3.1(b): Regulatory gaps remain for {q('déchets biomédicaux, déchets d équipements électriques et électroniques', 'biomedical waste, waste electrical and electronic equipment')} "
            f"and modern biotechnology products."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Notes need for harmonisation of environment, forestry and mining codes with decentralisation Act III."
        ),
        "comments": bi(
            "Lacunes DEEE/biomédicaux — pertinence indirecte (plastiques dans déchets médicaux et DEEE).",
            "WEEE/biomedical gaps — indirect relevance (plastics in medical and e-waste).",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"§II.2.2.1: 718 km coastline and riverbanks face erosion; {q('érosion côtière', 'coastal erosion')} major problem for "
            f"maritime/island populations (Langue de Barbarie, Saloum); socio-economic coastal infrastructure at risk."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§III Programme 3 LA31: combat coastal erosion; littoral concentrates 70% of national GDP (fisheries, tourism, ports)."
        ),
        "comments": bi(
            "Littoral — cadre indirect pour déchets plastiques marins et pollution côtière.",
            "Coast — indirect framework for marine plastic litter and coastal pollution.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§III.3.1 Vision: {q('À l horizon 2025, la gestion de l environnement et la gouvernance verte soient le socle d un Sénégal émergent, pour un développement socio-économique inclusif et durable', 'By 2025, environmental management and green governance should underpin an emerging Senegal for inclusive, sustainable socio-economic development')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§III.3.3 Global objective: create national momentum for environmental/resource management, SD integration and climate resilience."
        ),
        "comments": bi(
            "Vision et objectif global LPSEDD.",
            "LPSEDD vision and global objective.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§III.3.5.2 OS2 / Programme 3: {q('Lutte contre les pollutions, les nuisances et les effets néfastes des changements climatiques', 'Combat pollution, nuisances and adverse climate-change effects')} "
            f"including LA32 on hazardous chemicals/mercury waste, LA33 on EIA and air/water quality monitoring, LA34 on adaptation/mitigation."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Annex 2 indicators: CO2 emissions, PGES follow-up, % population satisfied with living environment."
        ),
        "comments": bi(
            "Programme 3 pollutions — cadre pour gestion des déchets dangereux (incl. plastiques chimiques).",
            "Programme 3 pollution — framework for hazardous waste (incl. chemical plastics).",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            f"§III.3.5.2 OS2: Integrate SD principles in public policies, {q('la gestion du cadre de vie, la promotion de moyens d existences... et les modes de production et de consommation', 'living-environment management, livelihood promotion... and production and consumption patterns')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§II.2.2.2: notes non-sustainable consumption/production patterns depleting resources; Programme 4 promotes green economy and environmental education."
        ),
        "comments": bi(
            "Modes de consommation/production — pertinent pour réduction des plastiques à usage unique.",
            "Consumption/production patterns — relevant for single-use plastic reduction.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§III Programme 4: institutional/legal reform (LA41), strengthen MEDD role in SD (LA42), environmental information and early-warning systems (LA43–44), "
            f"transversal multi-actor approach (LA45), green economy and renewable energy (LA48), PPP for natural-resource/environment management (LA411)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§IV: roles of local authorities (decentralised environmental competences), private sector, NGOs, parliamentarians and CSPE coordination."
        ),
        "comments": bi(
            "Gouvernance verte et coordination multi-acteurs — cadre pour politiques déchets sectorielles.",
            "Green governance and multi-actor coordination — framework for sectoral waste policies.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§V: MRV-based monitoring/evaluation with semi-annual indicators; mid-term and final independent evaluations; "
            f"CSPE coordinates ministries, agencies, local authorities, private sector, civil society and donors."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Annex 2 results framework tracks environmental state, biodiversity, CO2/habitant and PGES implementation."
        ),
        "comments": bi(
            "Suivi-évaluation — pas d'indicateur plastique spécifique identifié.",
            "M&E framework — no plastic-specific indicator identified.",
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
