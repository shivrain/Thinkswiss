#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi n° 2001-01 (Environmental Code)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2001-01_Code_Environnement.xlsx"

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

POLICY = {
    "policy_name": "Loi n° 2001-01 du 12 avril 2001 portant Code de l'Environnement",
    "policy_url": "https://www.denv.gouv.sn/telechargement/69/codes/18921/code-de-lenvironnement-2001.pdf",
    "policy_year": 2001,
    "policy_objective": (
        "Establish foundational environmental protection principles and obligations covering "
        "pollution control, waste management, product regulation, environmental assessment, "
        "and decentralised State/local responsibilities—providing the overarching legal "
        "framework through which plastic waste and pollution are indirectly governed, without "
        "plastic-specific targets or measures."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1,
    "policy_type_justification": (
        "Foundational environmental code enacted by the National Assembly (29 December 2000) "
        "and Senate (4 January 2001) and promulgated by the President as Law No. 2001-01; "
        "legislation adopted by parliament, not an executive regulation."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "environment, industry, agriculture, municipalities, fisheries, water, chemicals, "
        "waste management, urban planning, packaging"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "Art. L26: 'les travaux sont réglés sur le fonds pour la protection de l'environnement' "
        "when the polluter is unidentified. Art. L27: annual pollution and facility taxes on "
        "classified installations, including 'taxes à la pollution' calculated according to "
        "set FCFA schedules."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. L30–L31: All waste categories (including biomedical) must be eliminated or "
            "recycled ecologically rationally (L30). Producers/holders must ensure elimination "
            "or recycling themselves or via Ministry-licensed enterprises, or deliver waste to "
            "local authorities or State-licensed operators; recycling must comply with national "
            "standards (L31)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "L31: 'doit en assurer elle-même l'élimination ou le recyclage... entreprises agréées "
            "par le Ministre chargé de l'environnement... collectivité locale'. Enforcement via "
            "L38 (removal at polluter's cost) and criminal/administrative sanctions (Title IV). "
            "Monitoring via L37 authorisation and surveillance."
        ),
        "comments": (
            "No explicit mention of plastics; applies to all solid waste streams including plastic "
            "waste. Law 2023-15 replaced Law 2001-01 on 2 August 2023; substantive waste "
            "obligations largely carried forward."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. L32: Local authorities and constituted groupings must ensure elimination of "
            "household waste, potentially with regional/national State services; may also "
            "collect and treat non-household waste and create a special fee (redevance) for this."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "L32: 'Les collectivités locales... assurent l'élimination de déchets des ménages' "
            "and 'peuvent... créer une redevance spéciale'. Responsible authority designated; "
            "fee creation is optional; enforcement mechanisms not detailed in this article."
        ),
        "comments": (
            "Key governance instrument for municipal plastic waste collection. Complements Art. L6 "
            "decentralisation of environment competences to collectivités locales."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. L34: Conditions for waste collection, sorting, storage, transport, recovery, "
            "reuse, recycling, other treatment and final disposal—to prevent overproduction, "
            "recoverable waste loss and environmental pollution—are set by ministerial arrêté "
            "with other concerned ministers."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "L34: 'sont fixées par arrêté du Ministre chargé de l'environnement en collaboration "
            "avec les autres Ministres concernés'. Enabling provision operationalised through "
            "Décret n° 2001-282 and subsequent waste regulations."
        ),
        "comments": (
            "Enabling power scored in_force=1 because subordinate regulations (Décret 2001-282) "
            "were adopted. Enables technical standards applicable to plastic waste operations."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L35: Absolute prohibition on depositing waste on the public domain, including "
            "maritime public domain as defined by the Merchant Marine Code; public-domain "
            "concessionaires must eliminate or recycle waste found there."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L35: 'Est interdit de façon absolue le dépôt des déchets sur le domaine public y "
            "compris le domaine public maritime'. Authorities: concessionaires and collectivités "
            "(L36). Enforcement via L38 and general penal/administrative sanctions; no "
            "article-specific fine in legislative part."
        ),
        "comments": (
            "Directly relevant to littering and marine plastic leakage. Mandatory ban language "
            "('interdit de façon absolue')."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L36: Local authorities must stop all illegal dumping (dépôts sauvages) and "
            "eliminate abandoned waste of unidentified ownership with State services or licensed "
            "enterprises."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "L36: 'Les collectivités locales veillent à enrayer tous les dépôts sauvages. Elles "
            "assurent l'élimination... des déchets abandonnés'. Authority designated; monitoring "
            "duty ('veillent'); enforcement mechanisms indirect."
        ),
        "comments": "Governance/coordination instrument for illegal dumpsites, common pathway for plastic waste leakage.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. L37: Industrial waste elimination by producing and/or treatment structures must "
            "be carried out under Ministry of Environment authorisation and supervision with "
            "prescribed conditions; consumers and consumer associations must help ensure compliance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L37: 'doit être faite sur autorisation et surveillance du Ministère chargé de "
            "l'environnement qui fixe des prescriptions'. Authority and monitoring explicit; "
            "enforcement via L38 and Title IV sanctions."
        ),
        "comments": "Applies to industrial plastic waste streams and plastic processing facilities.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Art. L38: When waste is abandoned, deposited or treated contrary to the Code, the "
            "police authority must—after formal notice—remove it at the responsible party's "
            "expense and may require an escrow deposit with a public accountant."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "L38: 'l'autorité détentrice du pouvoir de police doit, après mise en demeure, "
            "assurer d'office l'élimination desdits déchets aux frais du responsable' and must "
            "'obliger le responsable à consigner... une somme correspondant au montant des "
            "travaux'. Mandatory, unconditional enforcement mechanism."
        ),
        "comments": "Cross-cutting enforcement instrument for illegal waste disposal including plastic litter.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. L39: Formal prohibition on importing hazardous waste onto Senegalese territory; "
            "criminalised under Art. L92."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L39: 'Il est formellement interdit d'importer des déchets dangereux sur le territoire "
            "sénégalais'. L92: imprisonment 1–5 years and fine 10–50 million FCFA for clandestine "
            "import of toxic hazardous waste."
        ),
        "comments": "Relevant to transboundary plastic/hazardous waste flows; implements Bamako/Bâle Convention principles.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. L40: Manufacture, import, possession for sale and supply to consumers of "
            "products or materials generating waste must be regulated by joint ministerial arrêté "
            "to facilitate waste elimination or, if necessary, prohibit them."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "L40: 'doivent être réglementées par arrêté conjoint des Ministres chargés du "
            "commerce, de l'environnement, et de la santé publique... ou... les interdire'. "
            "Responsible ministers designated; operational detail left to subordinate instruments."
        ),
        "comments": (
            "Enabling power exercised for plastics through Loi n° 2020-04. Scored governance "
            "(0.20) as the Code provision itself delegates action to arrêtés; in_force=1 because "
            "subsequent plastic legislation operationalised the power."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L41: Prohibition of immersion, incineration or any disposal of waste in "
            "continental, marine or fluvio-maritime waters under Senegalese jurisdiction."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L41: 'sont interdites'. L97 penalises marine/continental water pollution in violation "
            "of corresponding provisions (fine 500,000–2,000,000 FCFA; imprisonment 6 months–1 "
            "year; doubled on recurrence)."
        ),
        "comments": "Directly relevant to marine plastic pollution and aquatic waste disposal.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "Art. L42: Underground disposal (enfouissement) may only occur after Environment "
            "Ministry authorisation setting technical prescriptions and special rules."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L42: 'ne peut être opéré qu'après autorisation du Ministre chargé de l'environnement "
            "qui fixe des prescriptions techniques'. Authority and conditional permitting explicit."
        ),
        "comments": "Governs landfill/disposal infrastructure relevant to plastic waste end-of-life.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. L48–L51: Projects and activities likely to harm the environment, and relevant "
            "policies/plans/programmes, must undergo environmental assessment (EIA, SEA, audits); "
            "EIA content and public participation requirements specified; activity list set by decree "
            "(L50)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L48: 'devront faire l'objet d'une évaluation environnementale'. L49–L51: promoter "
            "funds EIA; Ministry issues certificate; public hearing integral. L94 penalises "
            "projects without EIA. Operationalised via Décret 2001-282 and Arrêtés 9468–9472 (2001)."
        ),
        "comments": (
            "Applies to plastic manufacturing, recycling plants and waste facilities listed in EIA "
            "decree. L50 list-setting is enabling but exercised through subordinate regulations."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. L55: Environment Ministry prepares emergency pollution plans with other ministries; "
            "includes sea and littoral pollution response plan under the Abidjan Convention; "
            "establishes intervention committees for small/medium emergencies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "L55: 'Des plans d'urgence... sont préparés par le Ministre chargé de l'environnement' "
            "and 'Le plan de lutte contre la pollution de la mer et du littoral est élaboré et "
            "adopté'. Authority designated; operational detail via PNIUM/POLMAR subordinate plans."
        ),
        "comments": (
            "Institutional framework later complemented by POLMAR (2009) and CENPOLMAR (2021). "
            "Hydrocarbon-focused in practice but covers littoral pollution governance."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. L13: Class I classified installations require prior exploitation authorisation; "
            "Class II require declaration and receipt before construction or operation—applies to "
            "plastic manufacturing/processing facilities in the ICPE nomenclature."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "L13: 'doivent faire l'objet... d'une autorisation' / 'd'une déclaration'. L22: "
            "sworn inspectors may visit installations at any time. L104–L105: administrative "
            "sanctions including suspension and closure for non-compliance."
        ),
        "comments": "Regulates industrial plastic production/processing installations under ICPE regime.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. L27: Annual rights and taxes on classified installations including surface taxes, "
            "pressure-equipment taxes and pollution taxes calculated according to existing pollution "
            "levels; quantified FCFA schedules in the Code."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "L27 sets fixed FCFA amounts (e.g. 30,000 F class I; 150 F/m² equipped surface). "
            "L26: 10% monthly surcharge for late payment. Revenue collected by Environment Ministry."
        ),
        "comments": (
            "Economic instrument applying to polluting industrial installations including plastics "
            "sector; not plastic-specific."
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

        # Auto-calculated scores
        g_col = get_column_letter(7)
        i_col = get_column_letter(9)
        k_col = get_column_letter(11)
        m_col = get_column_letter(13)
        p_col = get_column_letter(16)
        t_col = get_column_letter(20)
        o_col = get_column_letter(15)
        v_col = get_column_letter(22)

        ws[f"{o_col}{row_idx}"] = f"=AVERAGE({g_col}{row_idx},{i_col}{row_idx},{k_col}{row_idx},{m_col}{row_idx},{p_col}{row_idx})"
        ws[f"{v_col}{row_idx}"] = f"=AVERAGE({p_col}{row_idx},{t_col}{row_idx})"
        ws[f"{o_col}{row_idx}"].fill = auto_fill
        ws[f"{v_col}{row_idx}"].fill = auto_fill

    widths = {
        "A": 42, "B": 48, "C": 10, "D": 44, "E": 12, "F": 20,
        "G": 10, "H": 36, "I": 14, "J": 36, "K": 14, "L": 28,
        "M": 12, "N": 36, "O": 12, "P": 14, "Q": 22, "R": 52,
        "S": 14, "T": 18, "U": 44, "V": 14, "W": 44,
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
