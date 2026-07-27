"""
Builds the 4P Index (Plastic Pollution Policy Index) coding table for:

Loi n. 2013-10 du 28 decembre 2013 portant Code General des Collectivites Territoriales (CGCT)
(Senegal), as amended.

Sources consulted:
  - Original text: primature.sn (task-brief-matching mirror at dri.gouv.sn / primature.sn)
  - Consolidated text (incorporating Loi n. 2014-19 du 24 avril 2014; Loi n. 2018-15 du 8 juin 2018;
    Loi n. 2018-16 du 8 juin 2018; Loi n. 2021-38 du 3 decembre 2021), via ceracle.com,
    "texte consolide 03 decembre 2021" -- the most recent amendment identified, used for Col C per
    Coding Rule 2.
  - User-uploaded PDF of Loi n. 2019-12 du 8 juillet 2019 (OCR'd), one of the amending laws, used to
    verify/cross-check Art. 185/195/185bis/195bis on municipal and "ville" revenue sources, including
    the "taxe d'enlevement des ordures menageres" and the "fonds de dotation de la decentralisation".

KEY FINDING: this Code contains no plastics-specific provision, but Art. 170 and Art. 305 explicitly
transfer "la gestion des dechets et la lutte contre l'insalubrite" (waste management and anti-
insalubrity) to the "ville" and the "commune" respectively -- the foundational subnational-competence
basis for all municipal solid/plastic waste collection in Senegal, exactly as flagged in the task
brief. Two dedicated, VAT-funded decentralisation funds (Art. 324, 328) help finance these transferred
competences.

Produces:
  - senegal_loi_2013-10_4p_index.csv
  - senegal_loi_2013-10_4p_index.xlsx (formatted, wrapped, one sheet)
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

POLICY_NAME = "Loi n° 2013-10 du 28 décembre 2013 portant Code Général des Collectivités Territoriales"
POLICY_URL = (
    "www.dri.gouv.sn/sites/default/files/an-documents/LOI%20N%202013%201... (as provided in task "
    "brief) ; full/consolidated text at https://primature.sn/sites/default/files/2022-04/Loi%20n%C2"
    "%B0%202013-10%20du%2028%20d%C3%A9cembre%202013%20portant%20Code%20g%C3%A9n%C3%A9ral%20des%20"
    "Collectivit%C3%A9s%20locales.pdf and https://ceracle.com/wp-content/uploads/2020/08/Code-"
    "g%C3%A9n%C3%A9ral-des-Collectivit%C3%A9s-territoriales-texte-consolid%C3%A9.pdf (consolidated "
    "text, 3 Dec 2021) ; amending Loi n° 2019-12 du 8 juillet 2019 (user-uploaded PDF, OCR-verified)"
)
POLICY_YEAR = 2021
POLICY_OBJECTIVE = (
    "Establishes the 'Acte III de la décentralisation' framework, defining the organisational and "
    "financial architecture of Senegal's local authorities (department, commune, and the "
    "'ville'/city created to pool multiple communes' competences) and allocating nine transferred "
    "competence domains between the State and these authorities. It contains no plastics-specific "
    "provision, but Art. 170 and Art. 305 explicitly transfer 'la gestion des déchets et la lutte "
    "contre l'insalubrité' (waste management and anti-insalubrity control) to the ville and the "
    "commune respectively -- the foundational subnational-competence basis for municipal solid/"
    "plastic waste collection in Senegal. It also assigns binding municipal 'police' powers over "
    "street cleanliness and nuisance removal to mayors (Art. 118-126), creates a municipal waste-"
    "removal tax ('taxe d'enlèvement des ordures ménagères', Art. 195, 185), and establishes two "
    "dedicated, VAT-funded decentralisation funds (Fonds de Dotation de la Décentralisation, Art. "
    "324; Fonds d'Equipement des Collectivités Territoriales, Art. 328) that help finance these "
    "transferred competences, including waste management. Per the task brief, this is the 'key "
    "governance law underpinning plastic waste management' at the subnational level, essential for "
    "understanding how national environmental/waste rules coded elsewhere in this series are "
    "actually implemented by communes and villes on the ground."
)
POLICY_TARGET = 0
POLICY_TARGET_TEXT = ""
POLICY_TYPE = 1
POLICY_TYPE_JUSTIFICATION = (
    "A 'Loi' (Act) adopted by the National Assembly and promulgated by the President of the Republic, "
    "consistent with all other 'Loi n°...' instruments in this coding series; it is not an executive "
    "decree or ministerial order. It has since been amended by further Acts of Parliament (Loi "
    "n° 2014-19; Loi n° 2018-15; Loi n° 2018-16; Loi n° 2019-12, itself confirmed by an uploaded, "
    "OCR-verified copy signed by President Macky Sall; and Loi n° 2021-38), each also full Acts of "
    "Parliament -> legislation, coded 1 per Rule 6."
)
POLICY_INTEGRATION = 1
POLICY_SECTORS_LIST = (
    "local government/municipalities, environment, waste management, health, education, culture, "
    "sports & youth, agriculture/forestry, water resources, urban planning & construction, "
    "land administration, regional planning"
)
POLICY_CIRCULARITY = 0.50
POLICY_LIFECYCLE_PHASES_LIST = "disposal, environmental leakage"
POLICY_BUDGET = 1
POLICY_BUDGET_TEXT = (
    "Art. 324: 'Le Fonds de Dotation de la Décentralisation, créé par la loi des finances, reçoit une "
    "dotation équivalant à 3,5% de la taxe sur la valeur ajoutée perçue au profit du budget de l'Etat "
    "de la dernière gestion connue.' Art. 328: 'Il est créé le Fonds d'Equipement des Collectivités "
    "territoriales. Le Fonds d'Equipement des Collectivités territoriales reçoit une dotation "
    "équivalant à 2% de la taxe sur la valeur ajoutée au profit du budget de l'Etat de la dernière "
    "gestion connue.' Two dedicated funds, each generating its own earmarked revenue via a fixed "
    "percentage of national VAT receipts, used (among other purposes) to finance the transferred "
    "waste-management/anti-insalubrity competence -> ring-fenced, coded M=1."
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
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 170 and Art. 305: 'Les compétences suivantes sont transférées à la ville : - la "
            "gestion des déchets et la lutte contre l'insalubrité...' (Art. 170); 'La commune reçoit "
            "les compétences suivantes : ... - la gestion des déchets et la lutte contre "
            "l'insalubrité...' (Art. 305, within the 'Environnement et Gestion des Ressources "
            "Naturelles' competence domain). The explicit statutory assignment of waste-management "
            "competence to both tiers of municipal government."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: the specific level of government (ville or commune) is explicitly designated "
            "as the competent authority for this function (Art. 170, 305) (+0.25). Unconditional: the "
            "competence-transfer lists themselves state no exemption (+0.25). Enforcement: no penalty "
            "is tied to a commune's/ville's failure to exercise this competence -- not awarded. "
            "Monitoring: no dedicated audit/inspection mechanism specific to waste-competence exercise "
            "was identified in these articles -- not awarded."
        ),
        W_comments=(
            "The single most important instrument in this Code for the 4P Index: it is the statutory "
            "basis on which every other national waste/plastics rule in this series (Code de "
            "l'Environnement, Code de l'Assainissement, biomedical-waste decree) is actually delivered "
            "at street level. Coded per Rule 11 as a multi-level governance/competence-assignment "
            "instrument (0.40); note per Art. 4 that 'Tout transfert de compétence à une collectivité "
            "doit être accompagné du transfert concomitant par l'Etat...des ressources et moyens "
            "nécessaires' -- the financial counterpart to this transfer is coded separately below "
            "(Fonds de Dotation de la Décentralisation)."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 118-120, 122, 126: the mayor is charged with 'police municipale', whose object "
            "includes ensuring 'la sûreté, la tranquillité, la sécurité et la salubrité publics', "
            "explicitly covering 'le nettoiement, l'éclairage, l'enlèvement des encombrements' (street "
            "cleaning, lighting, removal of obstructions/refuse) (Art. 119); the mayor issues binding "
            "'arrêtés' to order local measures on matters entrusted to his vigilance, and 'est tenu "
            "d'assurer le respect des prescriptions de police qu'il édicte' (Art. 115); mayors may "
            "appoint sworn agents for sanitary-police functions under the oversight of the hygiene "
            "service (Art. 126)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: 'le maire', acting 'sous le contrôle du représentant de l'Etat' (Art. 118) "
            "(+0.25). Monitoring: sworn agents appointed by mayors for sanitary-police duties act "
            "'sous le contrôle du service d'hygiène' (Art. 126) (+0.25). Unconditional: the police-"
            "municipale duty over cleanliness/nuisance removal is stated generally (Art. 119) with no "
            "exemption identified (+0.25). Enforcement: no specific fine schedule for littering/"
            "insalubrity violations is set out within these general Code articles themselves (such "
            "penalties would typically appear in the mayor's own binding 'arrêtés', not in the Code) -- "
            "not awarded."
        ),
        W_comments=(
            "Mixed instrument per Rule 10/11: combines a multi-level assignment of enforcement "
            "responsibility to the municipal level (Rule 11) with a genuinely binding regulatory power "
            "(the mayor's 'arrêtés' have direct legal effect and are backed by State oversight) -- "
            "scored at the higher Regulatory value (1.0) with the coordination aspect noted here, per "
            "Rule 11's explicit carve-out for coordination mechanisms 'combined with a higher-scoring "
            "instrument type'."
        ),
    ),
    dict(
        P_instrument_type=0.60,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 195(1)(b) and Art. 185(1)(b) (as most recently restated by Loi n° 2019-12, Art. 1): "
            "'Les produits des taxes communales directes suivantes : taxe d'enlèvement des ordures "
            "ménagères...' is listed among the commune's operating tax revenues, and 'Les produits de "
            "la taxe d'enlèvement des ordures ménagères' among the ville's. These direct taxes 'sont "
            "créées par délibération du conseil municipal [or 'de la ville']' within the limits fixed "
            "by law."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Authority: 'délibération du conseil municipal' or 'du conseil de la ville' is explicitly "
            "designated as the mechanism for setting/adopting the tax (Art. 195, 185) (+0.25). "
            "Enforcement/Monitoring: no penalty or data-collection requirement specific to this tax is "
            "stated within these articles -- not awarded. Unconditional: not awarded -- the tax is only "
            "actually levied once the relevant municipal council has deliberated to create it, so its "
            "practical application is conditional on local council action rather than automatic."
        ),
        W_comments=(
            "In-force test (Rule 12) uncertainty flagged per Rule 8: Art. 195/185 list this tax "
            "declaratively ('sont les suivantes') as a recognised category of commune/ville revenue "
            "rather than using permissive 'peut' language, so S=1 is applied; however, since 'ces "
            "taxes...sont créées par délibération du conseil municipal', a reasonable alternative view "
            "would treat the tax as an enabling power not yet exercised (S=0) until a given commune's "
            "council actually adopts it by deliberation -- best estimate here treats the statutory "
            "revenue category itself as in force, with local adoption as an implementing detail."
        ),
    ),
    dict(
        P_instrument_type=0.60,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 324 and Art. 328: two dedicated decentralisation funds, each generating its own "
            "earmarked revenue from national VAT receipts, used to finance transferred competences "
            "(including 'la gestion des déchets et la lutte contre l'insalubrité') and local "
            "equipment/investment needs -- 'Le Fonds de Dotation de la Décentralisation...reçoit une "
            "dotation équivalant à 3,5% de la taxe sur la valeur ajoutée' (Art. 324); 'Il est créé le "
            "Fonds d'Equipement des Collectivités territoriales...reçoit une dotation équivalant à 2% "
            "de la taxe sur la valeur ajoutée' (Art. 328), distributed annually between departments, "
            "villes and communes according to objective criteria (Art. 325-327)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: distribution is fixed 'par arrêté conjoint du Ministre chargé des "
            "Collectivités territoriales et du Ministre chargé des Finances' (Art. 325, 328) (+0.25). "
            "Monitoring: 'Le Conseil national de Développement des Collectivités territoriales est "
            "consulté chaque année pour avis' on both the FDD rate and its distribution criteria (Art. "
            "324-325) (+0.25). Unconditional: the dedicated VAT-percentage funding mechanism itself "
            "applies without exemption, though the percentage 'est modifié dans le sens d'une hausse "
            "progressive...compte tenu des compétences' (+0.25). Enforcement: no penalty is attached to "
            "non-payment/misuse of these funds within these articles -- not awarded."
        ),
        W_comments=(
            "Directly underpins Column M (ring-fenced budget): these are the general financing "
            "instruments for the waste-management competence transferred in the first row above, not a "
            "plastics- or waste-specific levy in themselves, but coded here because they are the fiscal "
            "counterpart Art. 4 requires to accompany any competence transfer."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 320-323: 'Les charges financières résultant pour chaque département ou commune des "
            "transferts de compétences définies par le présent code font l'objet d'une attribution par "
            "l'Etat de ressources d'un montant au moins équivalent auxdites charges' (Art. 320); the "
            "amount of compensation is assessed jointly by the Ministers of Local Government and "
            "Finance at each stage of competence transfer (Art. 322), funded via the VAT-based "
            "dotation fund mechanism (Art. 323)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: 'arrêté conjoint du Ministre chargé des Collectivités territoriales et du "
            "Ministre chargé des Finances' assesses transfer-related charges (Art. 322) (+0.25). "
            "Monitoring: assessment occurs 'après avis du Conseil national de Développement des "
            "Collectivités territoriales' (Art. 322) (+0.25). Unconditional: 'Les ressources attribuées "
            "sont au moins équivalentes aux dépenses effectuées par l'Etat' -- a guaranteed minimum with "
            "no exemption identified (Art. 320) (+0.25). Enforcement: no penalty is attached to a "
            "failure by the State to provide adequate compensation -- not awarded."
        ),
        W_comments=(
            "A State-local government fiscal-coordination instrument under Rule 11, ensuring that "
            "transferred competences such as waste management are not unfunded mandates -- the general "
            "principle that the two dedicated funds above (Art. 324, 328) operationalise."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Environmental leakage",
        R_instrument_description=(
            "Art. 304: 'Le département reçoit les compétences suivantes : - la création et la gestion "
            "des forêts, zones protégées et sites naturels d'intérêt départemental...- l'élaboration et "
            "mise en œuvre de plans départementaux d'actions de l'environnement...- la protection des "
            "eaux souterraines et de surface...' -- environmental planning and natural-resource "
            "management competences assigned to the intermediate (department) tier of local "
            "government, complementing the commune/ville-level waste-management competence."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: 'le département' explicitly designated (Art. 304) (+0.25). Unconditional: the "
            "competence list is stated without exemption (+0.25). Enforcement/Monitoring: no penalty "
            "or dedicated audit mechanism was identified specific to the department's exercise of "
            "these competences -- not awarded."
        ),
        W_comments=(
            "Completes the three-tier competence picture requested by the task brief ('communes, "
            "departments, regions and other local authorities'); the department does not itself receive "
            "an explicit waste-management competence (that sits with commune/ville per Rule of "
            "subsidiarity), but its environmental-action-plan and water-protection competences are "
            "coordination-relevant to the broader pollution/leakage-prevention framework."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Environmental leakage",
        R_instrument_description=(
            "Art. 171(10): the mayor of the ville is charged, under the city council's oversight, "
            "'de veiller à la protection de l'environnement, de prendre en conséquence les mesures "
            "propres, d'une part, à empêcher ou à supprimer la pollution et les nuisances, d'autre "
            "part, à assurer la protection des espaces verts et, enfin, à contribuer à l'embellissement "
            "de la ville.' A binding duty on the mayor to take measures preventing or suppressing "
            "pollution and nuisances within the city."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: 'le maire de la ville', acting 'sous le contrôle du conseil de la ville' (Art. "
            "171) (+0.25). Unconditional: the duty is stated generally with no exemption identified "
            "(+0.25). Enforcement/Monitoring: no specific fine or inspection/reporting mechanism is "
            "attached to this general duty within the article itself -- not awarded."
        ),
        W_comments=(
            "A direct, binding pollution/nuisance-prevention duty on city mayors -- relevant to plastic "
            "litter and illegal dumping as a form of urban 'nuisance' the mayor must act to suppress, "
            "even though 'plastic' is not named."
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

with open("4p_index/senegal_loi_2013-10_4p_index.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLUMNS)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)

df = pd.DataFrame(rows, columns=COLUMNS)
xlsx_path = "4p_index/senegal_loi_2013-10_4p_index.xlsx"
df.to_excel(xlsx_path, sheet_name="4P_Index_Loi2013-10", index=False)

wb = load_workbook(xlsx_path)
ws = wb["4P_Index_Loi2013-10"]

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
