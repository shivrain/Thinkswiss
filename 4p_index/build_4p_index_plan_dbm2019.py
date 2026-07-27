"""
Builds the 4P Index (Plastic Pollution Policy Index) coding table for:

Reactualisation du Plan de Gestion des Dechets Biomedicaux (Senegal), Mai 2019 -- the Ministry of
Health's (MSAS) updated national biomedical-waste management plan, prepared with World Bank support
under the REDISSE and ISMEA health-system-strengthening projects.

Source: fetched from the task-brief URL (https://www.sante.gouv.sn/sites/default/files/plan_gestion_
dechets_biom%C3%A9dicaux_0.pdf); a 1957-line, 77-page national plan document (originally elaborated
in March 2014, updated in 2016 under REDISSE, and again in May 2019 to integrate ISMEA).

KEY FINDING: unlike most documents in this series, this plan repeatedly and explicitly discusses
plastics -- colour-coded plastic bags/containers ("sachets/sacs en plastique") are the primary
packaging medium for every biomedical-waste category, plastic bottles/packaging are explicitly listed
within the general/household ("ordures menageres") waste stream, and the plan's technical-guidance
section on elimination technologies explicitly identifies "Plastiques Halogenes (PVC)" as unsuitable
for incineration due to toxic emissions (dioxins), directing such waste toward alternative treatment.
It also proposes (but does not yet mandate) a polluter-pays/EPR-style regulatory reform for private
healthcare-waste producers, and recommends installing new incinerators/treatment infrastructure with
a dedicated, quantified budget (1.6 billion FCFA).

Produces:
  - senegal_plan_dbm2019_4p_index.csv
  - senegal_plan_dbm2019_4p_index.xlsx (formatted, wrapped, one sheet)
"""
import csv
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

COLUMNS = [
    "A_policy_name",
    "B_policy_url",
    "C_policy_year",
    "D_policy_objective",
    "E_policy_target",
    "F_policy_target_text",
    "G_policy_type",
    "H_policy_type_justification",
    "I_policy_integration",
    "J_policy_sectors_list",
    "K_policy_circularity",
    "L_policy_lifecycle_phases_list",
    "M_policy_budget",
    "N_policy_budget_text",
    "O_policy_score",
    "P_instrument_type",
    "Q_instrument_lifecycle_stage",
    "R_instrument_description",
    "S_instrument_in_force",
    "T_instrument_implementation",
    "U_instrument_implementation_text",
    "V_instrument_score",
    "W_comments",
]

# ---------------------------------------------------------------------------
# Section A - Policy-level fields (identical across every row)
# ---------------------------------------------------------------------------

POLICY_NAME = "Réactualisation Plan de Gestion des Déchets Biomédicaux (Mai 2019)"
POLICY_URL = "https://www.sante.gouv.sn/sites/default/files/plan_gestion_dechets_biom%C3%A9dicaux_0.pdf"
POLICY_YEAR = 2019
POLICY_OBJECTIVE = (
    "Updates Senegal's national biomedical-waste management plan (originally prepared by the "
    "Ministère de la Santé et de l'Action Sociale, MSAS, in March 2014; revised in 2016 under the "
    "REDISSE project; and updated again in May 2019 to integrate the World Bank-financed ISMEA "
    "project) to strengthen segregation, collection, storage, transport, treatment infrastructure "
    "(including private-sector incinerators) and staff/population training and awareness across "
    "health facilities, in compliance with World Bank environmental/social safeguards and national "
    "legislation (Code de l'Environnement; Décret n° 2008-1007 on biomedical waste, already coded "
    "elsewhere in this series). It is not a plastics-specific policy, but plastics feature "
    "extensively and explicitly throughout: colour-coded plastic bags and containers ('sachets/sacs "
    "en plastique', 'boîtes de sécurité en plastique') are the prescribed packaging medium for every "
    "biomedical-waste category; plastic bottles and packaging ('sacs et bouteilles en plastiques') "
    "are explicitly listed within the general/household waste stream; and the plan's technical-"
    "guidance section on elimination technologies explicitly identifies 'Plastiques Halogénés (PVC)' "
    "as a waste category that cannot be incinerated (due to toxic dioxin emissions), directing such "
    "plastic waste toward alternative treatment. The plan also proposes -- without yet mandating -- a "
    "polluter-pays/extended-producer-responsibility-style regulatory reform obliging private "
    "healthcare-waste producers to ensure collection and destruction of their waste."
)
POLICY_TARGET = 0
POLICY_TARGET_TEXT = ""
POLICY_TYPE = 0.50
POLICY_TYPE_JUSTIFICATION = (
    "A national sectoral plan/programme prepared by the Ministry of Health (MSAS) with World Bank "
    "technical and financial support (REDISSE, ISMEA projects) -- not a law or executive decree. "
    "Coded 0.50 rather than 0.25 because, unlike a purely aspirational strategy, this plan "
    "incorporates a quantified, costed and time-bound action-plan matrix (Section V.B: 'Coûts des "
    "actions', 'Calendrier de mise en œuvre, coûts et responsabilités') with a specific overall "
    "budget of 1,600,000,000 FCFA and defined institutional responsibilities and monitoring-"
    "evaluation indicators, rather than only general commitments."
)
POLICY_INTEGRATION = 1
POLICY_SECTORS_LIST = (
    "health, environment, private sector, local government/municipalities, education & training, "
    "international development cooperation, waste management, veterinary/animal health"
)
POLICY_CIRCULARITY = 0.75
POLICY_LIFECYCLE_PHASES_LIST = "consumption, disposal, environmental leakage"
POLICY_BUDGET = 1
POLICY_BUDGET_TEXT = (
    "'Il faut préciser que le Plan de Gestion élaboré en 2016 dans le cadre du REDISSE avait prévu un "
    "budget de 600 000 000FCFA. Aussi, le présent plan maintien cette même proposition en plus des "
    "1 000 000 000 F CFA prévus pour la contribution du projet ISMEA ce qui fait un budget global de "
    "1 600 000 000 F CFA centré essentiellement sur les structures sanitaires que le REDISSE et "
    "l'ISMEA vont appuyer.' A specific, quantified budget entirely dedicated to this plan's own "
    "biomedical-waste-management actions (not a general health-sector budget line) -> coded as "
    "ring-fenced for the plan's purposes (M=1)."
)
POLICY_SCORE_FORMULA = "[auto] O = (G + I + K + M) / 4"

policy_fields = dict(
    A_policy_name=POLICY_NAME,
    B_policy_url=POLICY_URL,
    C_policy_year=POLICY_YEAR,
    D_policy_objective=POLICY_OBJECTIVE,
    E_policy_target=POLICY_TARGET,
    F_policy_target_text=POLICY_TARGET_TEXT,
    G_policy_type=POLICY_TYPE,
    H_policy_type_justification=POLICY_TYPE_JUSTIFICATION,
    I_policy_integration=POLICY_INTEGRATION,
    J_policy_sectors_list=POLICY_SECTORS_LIST,
    K_policy_circularity=POLICY_CIRCULARITY,
    L_policy_lifecycle_phases_list=POLICY_LIFECYCLE_PHASES_LIST,
    M_policy_budget=POLICY_BUDGET,
    N_policy_budget_text=POLICY_BUDGET_TEXT,
    O_policy_score=POLICY_SCORE_FORMULA,
)

# ---------------------------------------------------------------------------
# Section B - Instrument-level fields (one dict per row)
# ---------------------------------------------------------------------------

instruments = [
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Recycling",
        R_instrument_description=(
            "Section IV.D.4 ('Tri, conditionnement, transport et élimination'): a mandatory colour-"
            "coded segregation and packaging standard using plastic bags/containers by waste category "
            "-- 'Ordures ménagères...Sachets plastiques noirs et poubelles noirs'; infectious/soiled "
            "waste in 'sachets plastiques jaune et poubelles jaunes...Sacs en plastique ou conteneurs "
            "résistants, étanches et autoclavables portant le symbole de risque biologique'; chemical/"
            "pharmaceutical waste in 'Sachets plastiques dans poubelles rouge...portant la mention "
            "\"toxique\"'; sharps in yellow 'safety boxes'. 'Les récipients de collecte de DBM doivent "
            "être : Non transparents et Résistants à l'humidité...'"
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: CLIN/CHSCT committees at each health facility are explicitly tasked to "
            "'Mettre en œuvre un système durable (Plan) de gestion des DBM dans leur formation "
            "sanitaire' (Section IV.D) (+0.25). Monitoring: 'Nombre et coût des poubelles, sachets et "
            "chariots' is an explicit monitoring-evaluation indicator tracked under this plan (Section "
            "V) (+0.25). Enforcement: no fine or penalty is stated within this plan document itself for "
            "non-compliant packaging -- not awarded (the plan's own site assessments extensively "
            "document real-world equipment shortages/non-compliance, e.g. at Louga, St-Louis, "
            "Ziguinchor and Kaspar Kamara facilities). Unconditional: not awarded given this documented, "
            "widespread pattern of equipment shortfalls acknowledged throughout the plan's own "
            "situation analysis."
        ),
        W_comments=(
            "The plan's most directly plastics-relevant operative instrument: virtually every "
            "biomedical-waste category is packaged in colour-coded plastic bags/containers by "
            "prescription. Legal force ultimately derives from the cross-referenced Décret n° 2008-1007 "
            "(already coded elsewhere in this series), which this plan operationalises and reinforces "
            "with detailed technical specifications rather than creating independent legal obligations "
            "of its own."
        ),
    ),
    dict(
        P_instrument_type=0.80,
        Q_instrument_lifecycle_stage="End of life",
        R_instrument_description=(
            "Section V.B (costed action-plan matrix) and Section V.C ('Mécanisme d'implantation des "
            "incinérateurs du secteur privé'): a budgeted, time-bound programme to install/upgrade "
            "incineration and treatment infrastructure at public health facilities, with private-"
            "sector cost-sharing/co-management arrangements ('cogestion à déterminer' with a 'grille de "
            "répartition des coûts de fonctionnement et d'entretien') recommended as the most reliable "
            "financing model given weak payment capacity in the public sub-sector."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: implementation responsibilities are explicitly assigned by institution in "
            "Section V.B.3 ('Responsabilités institutionnelles de mise en œuvre et de suivi') (+0.25). "
            "Monitoring: Section V.B.4 ('Arrangements institutionnels de suivi-évaluation') and the "
            "dedicated Section VI ('Suivi Evaluation du Plan') apply directly to this action (+0.25). "
            "Enforcement: no penalty is stated for non-implementation -- not awarded. Unconditional: not "
            "awarded -- implementation is explicitly scoped to 'les structures sanitaires que le "
            "REDISSE et l'ISMEA vont appuyer', a defined subset rather than all facilities nationally."
        ),
        W_comments=(
            "Matches the instrument-type table's own example for 0.80 ('Construction of a new waste "
            "management facility'). This is the primary planned-investment instrument addressing "
            "end-of-life treatment of plastic-containing biomedical waste (PPE, tubing, syringes, "
            "packaging) via incineration/treatment infrastructure."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Section V.C: 'il est nécessaire de prévoir des mesures incitatives, dont la plus "
            "essentielle porte sur le renforcement et l'application de la réglementation basée sur les "
            "principes «pollueur-payeur» et «obligation au producteur de déchets d'assurer leur "
            "collecte et leur destruction», notamment par l'élaboration d'une réglementation appropriée "
            "qui oblige les formations sanitaires (notamment privées) soit à traiter leurs DBM, soit à "
            "évacuer ou contracter un service de collecte des DBM vers les incinérateurs placés dans "
            "les zones de référence.' A proposed polluter-pays/extended-producer-responsibility-style "
            "regulatory reform for private healthcare-waste producers."
        ),
        S_instrument_in_force=0,
        T_instrument_implementation=0,
        U_instrument_implementation_text=(
            "No sub-score is awarded: this is explicitly framed as a future recommendation -- 'il est "
            "nécessaire de prévoir...notamment par l'élaboration d'une réglementation appropriée qui "
            "oblige...' -- rather than an operative rule. No specific implementing authority, "
            "enforcement mechanism, monitoring mechanism or unconditional scope is evidenced for this "
            "specific proposal within the plan's own text."
        ),
        W_comments=(
            "In-force test (Rule 12): textbook 'not yet exercised' enabling/aspirational recommendation "
            "-- the plan calls for future regulation to be drafted and applied, rather than stating a "
            "binding EPR obligation itself, so S=0. Cross-reference: this proposal echoes the EPR "
            "principle already codified as a general environmental-law principle in Loi n° 2023-15 "
            "(Art. 5), and the producer/holder waste-elimination duty already in force under Décret "
            "n° 2008-1007 (Art. 5) -- both coded elsewhere in this series -- suggesting the general "
            "legal hook already exists even though this plan's own specific private-sector EPR "
            "proposal has not yet been enacted as sector-specific regulation."
        ),
    ),
    dict(
        P_instrument_type=0.20,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Section V.A ('Cadrage', objectives and activities): 'Plaidoyers auprès des formations "
            "sanitaires pour l'allocation de budget spécifiques aux DBM' -- an advocacy action item "
            "encouraging individual health facilities to allocate their own dedicated internal budget "
            "lines for biomedical-waste management."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Monitoring: 'existence d'un budget pour la gestion des DBM et niveau de suffisance' is an "
            "explicit monitoring-evaluation indicator tracking the outcome of this advocacy action "
            "(Section V, indicator list) (+0.25). Authority/Enforcement/Unconditional: not awarded -- no "
            "specific responsible authority, penalty, or universal scope is evidenced for this "
            "particular advocacy item beyond the general plan-implementation framework."
        ),
        W_comments=(
            "A voluntary/persuasion-type instrument (advocacy) rather than a binding budget mandate; "
            "included because it is one of the plan's explicitly listed action items with its own "
            "tracked monitoring indicator."
        ),
    ),
    dict(
        P_instrument_type=0.20,
        Q_instrument_lifecycle_stage="End of life",
        R_instrument_description=(
            "Section V.F ('Technologies d'élimination des DBM'): technical guidance listing "
            "'Caractéristiques des DBM ne pouvant pas être incinérés', explicitly including "
            "'Plastiques Halogénés (PVC)' alongside pressurised-gas containers, large chemical "
            "quantities, radioactive waste, and mercury/cadmium-containing waste -- directing that such "
            "waste not be placed in incinerators and that alternative treatment/disposal be used, and "
            "instructing operators to 'Faire un triage des déchets au préalable' and 'Ne mettre dans "
            "l'incinérateur que des DBM incinérables'."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Unconditional: the technical exclusion of halogenated plastics/PVC from incineration is "
            "stated as a general technical fact/rule with no exception identified (+0.25). Authority/"
            "Enforcement/Monitoring: not awarded -- no specific responsible authority, penalty or "
            "tracking mechanism is tied to this specific technical-guidance item within the plan's text."
        ),
        W_comments=(
            "Satisfies Plastics Relevance Filter criterion (a) directly -- 'Plastiques Halogénés (PVC)' "
            "is an explicit synthetic-polymer-material reference. This is technical/scientific guidance "
            "embedded in the plan rather than a legally binding prohibition; it reflects the well-"
            "established environmental-health rationale (PVC incineration generates dioxins/HCl) "
            "already reflected in the general incineration standards coded elsewhere in this series "
            "(e.g. Décret n° 2008-1007, Art. 12; Norme NS 05-062)."
        ),
    ),
    dict(
        P_instrument_type=0.20,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Section V.E ('Stratégie de formation et de sensibilisation'): a costed training strategy "
            "for DBM-management actors (health workers, waste handlers, CLIN/CHSCT members) and a "
            "population/decision-maker awareness-raising strategy, both included in the budgeted "
            "action-plan matrix."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Authority: CLIN/CHSCT committees are explicitly tasked to 'Planifier et mettre en œuvre "
            "des programmes de formation, d'information et d'éducation pour le personnel, les malades "
            "et leur accompagnant' (Section IV.D) (+0.25). Enforcement/Monitoring/Unconditional: not "
            "awarded -- no penalty, dedicated tracking metric, or universal (non-resource-dependent) "
            "scope is evidenced specifically for this training/awareness action."
        ),
        W_comments=(
            "Coded as a Voluntary/information-type instrument (0.20), consistent with the type table's "
            "'awareness campaigns' example, even though the underlying institutional duty to plan "
            "training is stated in mandatory terms."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Sections IV.D.1/3 and VI: a multi-sectoral institutional-coordination framework for plan "
            "implementation and monitoring, assigning explicit responsibilities to MSAS, the Ministry "
            "of Environment (MEDD/DEEC), facility-level CLIN/CHSCT committees, a communal "
            "representative, and private-sector actors, together with a dedicated 'Suivi Evaluation du "
            "Plan' chapter (Section VI) setting out monitoring arrangements and indicators (e.g. "
            "equipment counts/costs, budget existence)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: detailed, named institutional responsibilities are set out in Section IV.D.3 "
            "('Responsabilités institutionnelles de mise en œuvre et de suivi') (+0.25). Monitoring: "
            "Section VI is an entire dedicated monitoring-evaluation chapter with specific indicators "
            "(+0.25). Enforcement: not awarded -- no penalty is attached to institutional non-"
            "performance. Unconditional: not awarded -- implementation remains scoped to REDISSE/"
            "ISMEA-supported facilities rather than the whole health system."
        ),
        W_comments=(
            "Coded per Rule 11 as a multi-level/multi-sectoral governance-coordination instrument, "
            "spanning national ministries (health, environment), facility-level committees, local "
            "government and the private sector -- the institutional backbone through which the "
            "plastics-relevant instruments above (segregation packaging, incinerator investment, PVC "
            "exclusion guidance) are meant to be delivered."
        ),
    ),
]

# ---------------------------------------------------------------------------
# Assemble rows, computing auto-fields O and V
# ---------------------------------------------------------------------------

O_score = round(
    (policy_fields["G_policy_type"] + policy_fields["I_policy_integration"]
     + policy_fields["K_policy_circularity"] + policy_fields["M_policy_budget"]) / 4,
    3,
)
policy_fields["O_policy_score"] = f"{O_score} [auto = (G+I+K+M)/4]"

rows = []
for inst in instruments:
    row = dict(policy_fields)
    row.update(inst)
    s = row["S_instrument_in_force"]
    p = row["P_instrument_type"]
    t = row["T_instrument_implementation"]
    v = round(s * (p + t) / 2, 3)
    row["V_instrument_score"] = f"{v} [auto = S*(P+T)/2]"
    rows.append(row)

with open("4p_index/senegal_plan_dbm2019_4p_index.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLUMNS)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)

df = pd.DataFrame(rows, columns=COLUMNS)
xlsx_path = "4p_index/senegal_plan_dbm2019_4p_index.xlsx"
df.to_excel(xlsx_path, sheet_name="4P_Index_PlanDBM2019", index=False)

wb = load_workbook(xlsx_path)
ws = wb["4P_Index_PlanDBM2019"]

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF")
wrap_top = Alignment(wrap_text=True, vertical="top")

col_widths = {
    "A": 34, "B": 30, "C": 10, "D": 46, "E": 10, "F": 30, "G": 10, "H": 34,
    "I": 10, "J": 30, "K": 10, "L": 26, "M": 10, "N": 30, "O": 16,
    "P": 12, "Q": 16, "R": 46, "S": 10, "T": 12, "U": 40, "V": 16, "W": 40,
}
for letter in [get_column_letter(c) for c in range(1, len(COLUMNS) + 1)]:
    ws.column_dimensions[letter].width = col_widths.get(letter, 20)

for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")

for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for cell in row:
        cell.alignment = wrap_top

ws.freeze_panes = "A2"
ws.row_dimensions[1].height = 30
for r in range(2, ws.max_row + 1):
    ws.row_dimensions[r].height = 120

wb.save(xlsx_path)
print("Wrote", xlsx_path, "and CSV with", len(rows), "instrument rows.")
print("Policy score O =", O_score)
for row in rows:
    print(row["Q_instrument_lifecycle_stage"], "->", row["V_instrument_score"])
