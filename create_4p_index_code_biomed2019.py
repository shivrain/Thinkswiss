#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for PGDBM May 2019 (bilingual FR/EN cells).

Source: official sante.gouv.sn PDF (text-extractable, 80 pages; updated May 2019).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Plan_Gestion_Dechets_Biomedicaux_2019.xlsx"

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
        "Plan de gestion des déchets biomédicaux — réactualisation, mai 2019 (REDISSE / ISMEA)",
        "Biomedical waste management plan — updated, May 2019 (REDISSE / ISMEA)",
    ),
    "policy_url": "https://www.sante.gouv.sn/sites/default/files/plan_gestion_dechets_biom%C3%A9dicaux_0.pdf",
    "policy_year": 2019,
    "policy_objective": bi(
        "Réactualiser le PGDBM (2016) pour les projets Banque mondiale REDISSE et ISMEA, en définissant "
        "ségrégation à la source, équipements de pré-collecte/collecte (poubelles, sachets plastiques codés, "
        "boîtes de sécurité), infrastructures de traitement (incinérateurs, stérilisateurs, autoclaves/broyeurs) "
        "et cadre institutionnel de gestion des DBM dans les formations sanitaires — incluant explicitement "
        "les déchets cliniques en contenants plastiques (sacs noirs/jaunes/rouges, safety boxes) et l'interdiction "
        "d'incinérer les plastiques halogénés (PVC).",
        "Update the 2016 biomedical waste management plan (PGDBM) for World Bank REDISSE and ISMEA projects, "
        "defining source segregation, pre-collection/collection equipment (bins, colour-coded plastic bags, "
        "safety boxes), treatment infrastructure (incinerators, sterilisers, autoclaves/shredders) and institutional "
        "framework for biomedical waste in health facilities — explicitly covering clinical waste in plastic "
        "containers (black/yellow/red bags, safety boxes) and the ban on incinerating halogenated plastics (PVC).",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"§V.B: {q('tri correctement dès leur production afin de les mettre immédiatement dans le bon contenant', 'sort correctly at production to place immediately in the right container')}; "
        f"§IV.E: colour-coded receptacles — {q('sacs noirs', 'black bags')} (general), {q('sacs rouges', 'red bags')} (pathological), {q('boîtes de sécurité', 'safety boxes')} (OPCT); "
        f"WHO benchmark: ~15% contaminated vs 85% non-hazardous waste requiring source separation."
    ),
    "policy_type": 0.50,
    "policy_type_justification": bi(
        "Plan national/programme opérationnel réactualisé en mai 2019 sous MSAS, financé par REDISSE et ISMEA. "
        "Document de planification et d'appui projet (≠1.0 loi, ≠0.75 arrêté/décret).",
        "National plan/operational programme updated May 2019 under MSAS, funded by REDISSE and ISMEA. "
        "Planning and project-support document (≠1.0 law, ≠0.75 order/decree).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "santé, environnement, assainissement, secteur privé, collectivités locales, formation sanitaire",
        "health, environment, sanitation, private sector, local government, health facilities",
    ),
    "policy_circularity": 0.25,
    "policy_lifecycle_phases_list": bi(
        "utilisation, collecte, stockage, transport, traitement, élimination",
        "use, collection, storage, transport, treatment, disposal",
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        f"§V.1: detailed cost table — e.g. 2,210 units pre-collection equipment (FCFA 22.1M REDISSE + 17.5M ISMEA); "
        f"2,000 plastic waste bags (FCFA 400,000); needle shredders, DBM sterilisers, Montfort/electromechanical "
        f"incinerators; PRONALIN/CLIN/CHSCT training budgets. Plan funds priority activities only, not full national strategy."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Executive summary: May 2019 update of 2016 PGDBM harmonising REDISSE and ISMEA World Bank projects; "
            f"national DBM production ~124.2 m³/day across 3,084 health structures; plan covers investment needs, "
            f"capacity building and coordination mechanisms."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Led by MSAS/PRONALIN; implementation in project-supported health facilities; government to continue "
            f"activities beyond project end via national programmes or development partners."
        ),
        "comments": bi(
            "Plan opérationnel projet — réactualisation 2016 pour REDISSE/ISMEA.",
            "Project operational plan — 2016 update for REDISSE/ISMEA.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§V.A.1: global objective — {q('contribuer au bien-être de tous les Sénégalais par une gestion durable des DBM dans les formations sanitaires', 'contribute to well-being of all Senegalese through sustainable DBM management in health facilities')} "
            f"with environmentally viable, technically feasible, socially acceptable systems."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Four specific objectives: legal/institutional strengthening; equipment; training/awareness; monitoring."
        ),
        "comments": bi(
            "Objectif global PGDBM 2019.",
            "PGDBM 2019 global objective.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"§V.A.2 Objectif 2: provide health facilities with {q('matériels de pré collecte, collecte et stockage (poubelles)', 'pre-collection, collection and storage materials (bins)')}, "
            f"PPE (gloves, boots, masks, coveralls) and treatment infrastructure (DBM sterilisers/modern incinerators for regional hospitals; "
            f"De Montfort incinerators for health centres/posts)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§V.1 cost table: 2,210 pre-collection units; 2,000 plastic waste bags; needle shredders, sterilisers, incinerators budgeted under REDISSE/ISMEA."
        ),
        "comments": bi(
            "Axe équipement — poubelles et sachets plastiques de collecte.",
            "Equipment axis — bins and plastic collection bags.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"§V.B / §IV.E: mandatory {q('tri à la source', 'source segregation')}; producer responsible for placing waste "
            f"immediately in correct container; contaminated waste (~15%) must be separated from non-hazardous (~85%) per WHO."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Diagnostic notes tri not generalised; mixing of DBM with household waste common; training and CLIN/CHSCT supervision planned."
        ),
        "comments": bi(
            "Ségrégation à la source — réduit volumes infectieux et besoins en contenants plastiques.",
            "Source segregation — reduces infectious volumes and plastic-container needs.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"§IV.E / Annex tri table: colour-coded plastic receptacles — {q('Sachets plastiques noirs et poubelles noirs', 'black plastic bags and black bins')} "
            f"for general waste (packaging, plastic bottles, sterile material packaging); "
            f"{q('sachets plastiques jaune', 'yellow plastic bags')} for infectious waste; "
            f"{q('Sachets plastiques dans poubelles rouge', 'plastic bags in red bins')} for pathological/soiled waste; "
            f"{q('Boîtes safety boxes jaunes', 'yellow safety boxes')} for sharps (OPCT)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Receptacles must be non-transparent, moisture-resistant, manipulation-resistant and closed; "
            f"deficit of plastic bags and safety boxes noted in facility assessments."
        ),
        "comments": bi(
            "Mention plastique la plus explicite — codage couleur des sachets/poubelles plastiques.",
            "Most explicit plastic mention — colour-coded plastic bags/bins.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"§IV.E: sharps placed in {q('boîtes de sécurité safety boxes ou des contenants en plastique', 'safety boxes or plastic containers')}; "
            f"OPCT must be segregated immediately at point of care in compliant puncture-resistant, liquid-tight containers "
            f"bearing biological-risk symbol; do not recap needles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Plan recommends plastic (not cardboard) safety boxes; notes improvised containers (e.g. water bottles) in some facilities."
        ),
        "comments": bi(
            "OPCT en contenants plastiques/boîtes de sécurité — seringues et matériel à usage unique.",
            "Sharps in plastic containers/safety boxes — syringes and single-use items.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§IV.2.3 / §VI: treatment options — modern incinerators (HPD, Dantec, Fann, Grand Yoff); De Montfort single-chamber incinerators "
            f"for health centres; autoclave/shredder at Grand Yoff (non-incineration WHO/UNDP/GEF pilot); chemical disinfection; "
            f"controlled landfill cells for final disposal."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Most facilities lack compliant incinerators; open burning and uncontrolled dumps common; Dakar controlled landfill to include DBM cells."
        ),
        "comments": bi(
            "Infrastructures d'élimination — traite déchets conditionnés en plastique après incinération/stérilisation.",
            "Disposal infrastructure — treats plastic-packaged waste after incineration/sterilisation.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§VI.1: DBM that {q('ne pouvant pas être incinérés', 'cannot be incinerated')} include {q('Plastiques Halogénés (PVC)', 'Halogenated Plastics (PVC)')}, "
            f"pressurised gas containers, large chemical quantities, radioactive waste, mercury/cadmium waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Pre-sorting required before incineration; only combustible DBM (>60% combustible, <30% moisture) may enter incinerators."
        ),
        "comments": bi(
            "Interdiction d'incinération du PVC — lien direct plastique halogéné clinique.",
            "PVC incineration ban — direct halogenated clinical plastic link.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§IV.D.2: references Décret n° 2008-1007 — producers must eliminate or have eliminated by {q('entreprises agréées par le Ministre chargé de la santé', 'enterprises approved by the Health Minister')}; "
            f"sets disinfection of infectious containers, pre-treatment, sorting, storage, transport and disposal modalities; "
            f"operators require health ministry approval."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Plan notes lack of standardised technical guides and unclear MSAS/MEDD/commune responsibilities."
        ),
        "comments": bi(
            "Cadre juridique DBM — complète le plan opérationnel.",
            "DBM legal framework — complements operational plan.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"§V.C: public-private partnership model — public facilities with incinerators to {q('polariser des cabinets privés', 'serve as hub for private clinics')} "
            f"for DBM treatment; private clinics to contract collection/transport to zone incinerator; "
            f"polluter-pays and producer-responsibility principles for private facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"No specialised private DBM collection companies; cleaning firms and municipalities mix DBM with household waste."
        ),
        "comments": bi(
            "PPP secteur privé — collecte/transport vers incinérateurs.",
            "Private-sector PPP — collection/transport to incinerators.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§V.A.2 Objectif 1: strengthen legal/institutional framework — workshops, clarify actor roles, support internal DBM plans "
            f"at health facilities, public-private partnerships, advocacy for dedicated DBM budgets."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"PRONALIN leads priority action plan; CLIN/CHSCT for facility-level proximity monitoring."
        ),
        "comments": bi(
            "Gouvernance institutionnelle — plans internes de gestion DBM.",
            "Institutional governance — internal DBM management plans.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"§V.A.2 Objectif 3: training programmes for all DBM-chain operators; trainer-of-trainers; public awareness on "
            f"dangers of improper DBM management; PIC (infection prevention and control) strategy support."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"PRONALIN national training history; plan budgets CLIN/CHSCT and PRONALIN training units under REDISSE/ISMEA."
        ),
        "comments": bi(
            "Formation/sensibilisation — bonnes pratiques de tri et manipulation des contenants plastiques.",
            "Training/awareness — good practices for sorting and handling plastic containers.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"§V.A.2 Objectif 4 / §V.D: monitoring by CLIN/CHSCT (proximity), PRONALIN supervision, external oversight by "
            f"DEEC/DREEC (environment directorates); waste pickers at dumps risk reusing contaminated recycled objects."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Diagnostic: external supervision insufficient; recovery of reusables/recyclables at dumps poses infection risk."
        ),
        "comments": bi(
            "Suivi-évaluation — risque de récupération d'objets plastiques contaminés en décharge.",
            "M&E — risk of recovering contaminated plastic objects at dumps.",
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
