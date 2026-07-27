"""
Builds the 4P Index (Plastic Pollution Policy Index) coding table for:

Loi n° 2001-01 (Senegal) portant Code de l'Environnement
[official texts consulted show promulgation date 15 January 2001; see comments
for the 12 April 2001 date-discrepancy note]

Produces:
  - senegal_loi_2001-01_4p_index.csv
  - senegal_loi_2001-01_4p_index.xlsx (formatted, wrapped, one sheet)
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

POLICY_NAME = "Loi n° 2001-01 du 15 janvier 2001 portant Code de l'Environnement"
POLICY_URL = (
    "https://faolex.fao.org/docs/pdf/sen34608.pdf ; "
    "https://acbep.gouv.sn/sites/default/files/2021-12/Loi%20n%C2%B0%202001-01%20du%2012%20avril%202001"
    "%20portant%20code%20de%20l%27environnement.pdf (government mirror; source database URL "
    "www.denv.gouv.sn/telechargement/69/codes/18921/code-de-lenvironneme... was truncated/unverifiable)"
)
POLICY_YEAR = 2001
POLICY_OBJECTIVE = (
    "Foundational, cross-sectoral environmental framework law. It does not name 'plastics' or "
    "'plastic waste' anywhere in its text, but it establishes the general legal architecture that "
    "governs plastics indirectly: (i) a classified-installations (ICPE) permitting regime for any "
    "industrial/artisanal/commercial activity (incl. plastics manufacturing/processing); (ii) a "
    "general duty to eliminate or recycle all waste 'in an ecologically rational manner' (Art. L30-L38), "
    "which applies to plastic waste streams in practice; (iii) an enabling power (Art. L40) allowing the "
    "State to regulate, and if necessary prohibit, the manufacture, import, possession-for-sale or "
    "consumer supply of any 'waste-generating product or material' - the legal hook later used by "
    "Senegal to ban single-use plastics; (iv) mandatory environmental impact assessment (EIA) for "
    "projects likely to harm the environment; and (v) a decentralisation framework assigning waste "
    "and environmental-protection duties across State and local government (collectivités locales). "
    "There is no plastic-specific reduction, recycling or EPR target in this text."
)
POLICY_TARGET = 0
POLICY_TARGET_TEXT = ""
POLICY_TYPE = 1
POLICY_TYPE_JUSTIFICATION = (
    "Full Act of Parliament: 'L'Assemblée Nationale a délibéré et adopté, en sa séance du vendredi 29 "
    "décembre 2000 ; Le Sénat a délibéré et adopté, en sa séance du jeudi 4 janvier 2001 ; Le Président de "
    "la République promulgue la loi...'. Adopted by both chambers of the (then-bicameral) Senegalese "
    "legislature and promulgated by the President -> legislation, not an executive/ministerial instrument."
)
POLICY_INTEGRATION = 1
POLICY_SECTORS_LIST = (
    "environment, industry, agriculture, health, fisheries, water resources, mining, energy, "
    "maritime transport, urban planning & construction, commerce/trade, local government/municipalities, "
    "waste management"
)
POLICY_CIRCULARITY = 1
POLICY_LIFECYCLE_PHASES_LIST = "production, consumption, recycling, disposal, environmental leakage"
POLICY_BUDGET = 1
POLICY_BUDGET_TEXT = (
    "Art. L25: installations classées 'sont assujetties aux droits et taxes prévus à l'article L27'. "
    "Art. L26/L27: taxes (superficiaires, sur appareils à pression, à la pollution) are collected by the "
    "Ministère chargé de l'environnement. Art. L25 al.2: 'les travaux sont réglés sur le fonds pour la "
    "protection de l'environnement' -- a dedicated Environmental Protection Fund financed by the Code's "
    "own pollution levies (ring-fenced)."
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
        Q_instrument_lifecycle_stage="Production",
        R_instrument_description=(
            "Art. L9-L24: 'Installations classées pour la protection de l'environnement' (ICPE). Any "
            "factory, workshop, depot, worksite or quarry (incl. plastics manufacturing/processing/"
            "recycling plants) presenting a danger or nuisance is subject to mandatory prior authorisation "
            "or declaration depending on classification (Art. L10-L13); the Minister of Environment can "
            "order closure of non-compliant facilities (Art. L23) and existing facilities may be exempted "
            "under a grandfather clause (Art. L24)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: 'Le Ministre chargé de l'environnement' issues authorisations and closure orders "
            "(Art. L20-L23). Enforcement: fines for unauthorised modification (Art. L87: 500,000-1,500,000 "
            "FCFA) and suspension power (Art. L105). Monitoring: 'L'inspection des installations classées "
            "est assurée par des agents assermentés' (Art. L22). Not unconditional: existing installations "
            "are grandfathered (Art. L24), so the +0.25 unconditional sub-score is not awarded."
        ),
        W_comments=(
            "This is the primary permitting gateway through which any plastics production/processing "
            "facility in Senegal would be regulated; the Code itself does not name plastics. DATE NOTE: "
            "the task brief cited 'Loi n° 2001-01 du 12 avril 2001'; the primary legislative texts "
            "(FAOLEX, ACBEP, Ministry of Health mirror) all show the law itself is dated 15 January 2001 "
            "(adopted by the National Assembly on 29 Dec 2000 and the Senate on 4 Jan 2001). 12 April 2001 "
            "is the date of the separate implementing Décret n° 2001-282 portant application du Code de "
            "l'environnement. SUPERSESSION NOTE: this 2001 Code was repealed and replaced by Loi n° "
            "2023-15 du 2 août 2023 portant Code de l'Environnement (adopted 7 June 2023); this table "
            "codes the 2001 text as its own distinct policy (year = 2001), consistent with it being a "
            "separately-numbered replacement rather than an amendment."
        ),
    ),
    dict(
        P_instrument_type=0.60,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. L25-L27 and L73: annual pollution-linked charges/taxes on classified installations "
            "('taxes superficiaires', taxes on pressurised equipment, and a pollution charge/'taxe à la "
            "pollution' set by measured pollutant load), whose proceeds fund remediation and are payable "
            "into the ring-fenced 'fonds pour la protection de l'environnement' (Art. L25 al.2)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=1.0,
        U_instrument_implementation_text=(
            "Authority: taxes 'sont perçus par le Ministère chargé de l'environnement' (Art. L27). "
            "Enforcement: 'Le montant des taxes est majoré de 10% lorsque le paiement n'est pas effectué "
            "dans les délais' (Art. L25), continuing monthly. Monitoring: pollution charge is based on "
            "'la moyenne des résultats des prélèvements effectués lors d'une ou de plusieurs campagnes de "
            "mesures' (Art. L73). Unconditional: charge applies to all classified installations with no "
            "exemption identified."
        ),
        W_comments=(
            "Directly underpins Col M (ring-fenced budget). Note Art. L15 offers a separate 3-year tax "
            "exoneration to firms newly complying with anti-pollution rules -- a targeted incentive, not "
            "a carve-out from this charge, so it is not treated as an exemption/loophole here."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. L30-L38: general duty that 'les déchets doivent être éliminés ou recyclés de manière "
            "écologiquement rationnelle' (Art. L30), applicable to all waste categories (Art. L30 al.2); "
            "waste holders must dispose of/recycle waste themselves or via approved operators or hand it "
            "to the local authority (Art. L31); municipalities ('collectivités locales') are assigned "
            "responsibility for household waste and may levy a special charge for non-household waste "
            "(Art. L32); industrial waste disposal requires Ministry authorisation and supervision "
            "(Art. L37)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: 'Les collectivités locales...assurent l'élimination de déchets des ménages' "
            "(Art. L32); Ministry authorises/supervises industrial disposal (Art. L37). Enforcement: "
            "'l'autorité détentrice du pouvoir de police doit, après mise en demeure, assurer d'office "
            "l'élimination...aux frais du responsable' and may require a financial deposit (Art. L38). "
            "Monitoring: 'sur autorisation et surveillance du Ministère' (Art. L37). Unconditional: not "
            "awarded -- local implementation is conditioned on 'leurs caractéristiques et...quantités "
            "produites' (Art. L32) and depends on further implementing arrêtés (Art. L34)."
        ),
        W_comments=(
            "This is also a multi-level governance/coordination instrument under Rule 11 (State-"
            "municipality division of waste-management duties, consistent with the general decentralisation "
            "framework of Art. L4-L6); scored here at the higher applicable type (1.0, regulatory duty) "
            "per Rule 10 rather than at 0.40, since the core obligation is a binding regulatory duty, not "
            "merely a coordination mechanism."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="End of life",
        R_instrument_description=(
            "Absolute prohibitions on uncontrolled disposal/leakage: Art. L35 al.2 'Est interdit de façon "
            "absolue le dépôt des déchets sur le domaine public y compris le domaine public maritime'; "
            "Art. L39 'Il est formellement interdit d'importer des déchets dangereux sur le territoire "
            "sénégalais'; Art. L41 bans immersion/incineration/any disposal of waste in inland, marine or "
            "estuarine waters under Senegalese jurisdiction; Art. L92 criminalises clandestine import of "
            "hazardous waste (10-50 million FCFA fine + 1-5 years' imprisonment)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=1.0,
        U_instrument_implementation_text=(
            "Authority: Ministère chargé de l'environnement / collectivités locales enforce (Art. L36, "
            "L38). Enforcement: Art. L92 sets a criminal penalty (10,000,000-50,000,000 FCFA and 1-5 "
            "years' imprisonment) for illegal hazardous-waste import; general penal Titre IV applies to "
            "dumping offences. Monitoring: sworn inspectors/agents (Art. L22) and 'agent assermenté' powers "
            "under the water-police regime (Art. L70) cover illegal dumping into water. Unconditional: "
            "the prohibitions are stated in absolute terms ('de façon absolue', 'formellement interdit') "
            "with no exemption identified."
        ),
        W_comments=(
            "Directly relevant to plastic litter/marine leakage even though 'plastic' is not named: "
            "single-use plastic bags and packaging are the paradigmatic waste dumped on public land or "
            "into water bodies. Cross-reference: Senegal later adopted plastics-specific bans (Loi n° "
            "2015-09 of 4 May 2015, then Loi n° 2020-04 of 8 January 2020) that build on this general "
            "prohibition regime."
        ),
    ),
    dict(
        P_instrument_type=1.0,
        Q_instrument_lifecycle_stage="Consumption",
        R_instrument_description=(
            "Art. L40: 'La fabrication, l'importation, la détention en vue de la vente, la mise à la "
            "disposition du consommateur de produits ou matériaux générateurs de déchets doivent être "
            "réglementées par arrêté conjoint des Ministres chargés du commerce, de l'environnement, et de "
            "la santé publique, en vue de faciliter l'élimination desdits déchets ou, en cas de nécessité, "
            "les interdire.' This is the general enabling clause under which waste-generating products "
            "(including single-use plastics) can be regulated or banned by joint ministerial order."
        ),
        S_instrument_in_force=0,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Authority: joint order of three named ministries is specified (+0.25). No enforcement or "
            "monitoring mechanism is specified within Art. L40 itself. Not unconditional: the ban option "
            "is explicitly contingent -- 'en cas de nécessité' -- and depends on a subordinate arrêté that "
            "had not yet been adopted under this Code at the time of enactment."
        ),
        W_comments=(
            "In-force test (Rule 12): this is an enabling power ('doivent être réglementées par arrêté "
            "conjoint... en cas de nécessité, les interdire'), not a self-executing ban, so S=0 under this "
            "Code. IMPORTANT cross-reference: Senegal subsequently operationalised comparable powers "
            "through dedicated legislation rather than an Art. L40 arrêté -- Loi n° 2015-09 du 4 mai 2015 "
            "(banned thin plastic bags below a micron threshold) and, replacing it, Loi n° 2020-04 du 8 "
            "janvier 2020 (bans single-use plastic tableware/straws/sachets and all point-of-sale plastic "
            "bags, plus EPR and deposit-refund provisions). Those laws should be coded as separate policy "
            "entries in the 4P Index rather than as an 'exercised' instance of this Art. L40 power."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Production",
        R_instrument_description=(
            "Art. L48-L52: mandatory environmental impact assessment (EIA) / 'évaluation environnementale' "
            "for any development project or activity likely to harm the environment, plus sectoral "
            "policies, plans and programmes; a decree lists the categories requiring an EIA (Art. L50); "
            "public authorities 'ne pourront décider, approuver ou autoriser des travaux' on listed "
            "projects without an EIA; a public hearing ('audience publique') is an integral part of the "
            "procedure (Art. L52)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: EIA is submitted to and reviewed by the 'Ministère chargé de l'environnement' "
            "(Art. L49). Enforcement: carrying out a listed project without an EIA is a separately listed "
            "penal offence ('réalisé un projet visé à l'article L50 sans étude d'impact'). Monitoring: "
            "mandatory public hearing procedure (Art. L52). Unconditional: not awarded -- the EIA "
            "requirement only attaches to activities listed in the implementing decree under Art. L50, "
            "not universally."
        ),
        W_comments=(
            "Coded as a Procedural measure (0.40, 'planning requirements' per the instrument-type table) "
            "rather than as Regulatory, since its function is a mandatory ex-ante planning/appraisal "
            "gate rather than a substantive standard; a reasonable alternative view would score it higher "
            "given its binding, approval-blocking effect -- flagged here as a borderline case."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Environmental leakage",
        R_instrument_description=(
            "Art. L55-L56: national/sectoral emergency response planning for serious pollution incidents. "
            "A 'Plan National d'Intervention d'Urgence' is prepared by the Ministry of Environment in "
            "coordination with other ministries via a specialised technical committee under the "
            "'Secrétariat Permanent du Conseil Supérieur des Ressources Naturelles et de l'Environnement' "
            "(Art. L55); a marine/coastal pollution plan implements the 1981 Abidjan Convention; operators "
            "of Class-1 (authorisation-level) classified installations must establish an internal "
            "emergency/operations plan (Art. L56 al.1); the Minister may extend this duty to Class-2 "
            "(declaration-level) installations by order (Art. L56 al.2, an enabling power)."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: Ministry of Environment + specialised technical committee under the Conseil "
            "Supérieur (+0.25). Unconditional: the Class-1 internal-plan duty in Art. L56 al.1 is stated "
            "as mandatory ('est tenu d'établir') with no exemption identified (+0.25). No explicit "
            "fine/penalty is tied to this specific provision, and no dedicated inspection/audit mechanism "
            "for the plan itself is specified in the text, so enforcement and monitoring sub-scores are "
            "not awarded."
        ),
        W_comments=(
            "The Class-2 extension in Art. L56 al.2 ('peut...être tenu') is a separate enabling power and "
            "would be coded 0 (not in force) unless activated by ministerial order; only the mandatory "
            "Class-1 duty is scored as in force here (S=1)."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. L4-L6: multi-level governance/coordination framework. Environmental protection is "
            "declared part of national socio-economic policy (Art. L4); implementation is led by the "
            "'Ministère chargé de l'environnement' in close collaboration with all other line ministries "
            "(Art. L5); pursuant to the Code des collectivités locales and the decentralisation laws, "
            "local governments receive a transfer of competencies for environment and natural-resource "
            "management and must comply with the Code's provisions in exercising them (Art. L6), subject "
            "to the State's reserved civil-defence powers."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.25,
        U_instrument_implementation_text=(
            "Authority: 'Ministère chargé de l'environnement' explicitly designated as coordinator "
            "(Art. L5) (+0.25). No specific enforcement mechanism or monitoring/reporting requirement is "
            "attached to this coordination provision itself, and the competency transfer is explicitly "
            "conditioned ('ce transfert de compétences ne fait pas obstacle au droit pour l'Etat de "
            "prendre...les mesures nécessaires...en matière de défense civile ou militaire', Art. L6), so "
            "the unconditional sub-score is not awarded."
        ),
        W_comments=(
            "Coded per Rule 11 as a procedural/coordination instrument assigning responsibilities across "
            "national and municipal levels of government; the more concrete municipal waste-management "
            "duties that flow from this framework are scored separately in the waste-elimination row above "
            "to avoid double counting the same obligation twice."
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

with open("4p_index/senegal_loi_2001-01_4p_index.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLUMNS)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)

df = pd.DataFrame(rows, columns=COLUMNS)
xlsx_path = "4p_index/senegal_loi_2001-01_4p_index.xlsx"
df.to_excel(xlsx_path, sheet_name="4P_Index_Loi2001-01", index=False)

wb = load_workbook(xlsx_path)
ws = wb["4P_Index_Loi2001-01"]

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF")
wrap_top = Alignment(wrap_text=True, vertical="top")

col_widths = {
    "A": 34, "B": 30, "C": 10, "D": 46, "E": 10, "F": 30, "G": 10, "H": 34,
    "I": 10, "J": 30, "K": 10, "L": 26, "M": 10, "N": 30, "O": 16,
    "P": 12, "Q": 16, "R": 46, "S": 10, "T": 12, "U": 40, "V": 16, "W": 40,
}
for i, letter in enumerate([get_column_letter(c) for c in range(1, len(COLUMNS) + 1)]):
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
    ws.row_dimensions[r].height = 110

wb.save(xlsx_path)
print("Wrote", xlsx_path, "and CSV with", len(rows), "instrument rows.")
print("Policy score O =", O_score)
for row in rows:
    print(row["Q_instrument_lifecycle_stage"], "->", row["V_instrument_score"])
