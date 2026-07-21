#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret n° 2001-282 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2001-282_Code_Environnement.xlsx"

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
        "Décret n° 2001-282 du 12 avril 2001 portant application du Code de l'Environnement",
        "Decree No. 2001-282 of 12 April 2001 implementing the Environmental Code",
    ),
    "policy_url": "https://faolex.fao.org/docs/pdf/sen37192.pdf",
    "policy_year": 2001,
    "policy_objective": bi(
        "Mettre en œuvre le Code de l'environnement par la réglementation des installations classées, "
        "des études d'impact, de la pollution des eaux, de la police de l'eau, de la pollution "
        "atmosphérique et de la pollution sonore — fournissant le cadre réglementaire indirect de "
        "lutte contre la pollution plastique via les ICPE, les rejets, l'incinération et l'évaluation "
        "environnementale des projets de traitement des déchets, sans mesures spécifiques aux plastiques.",
        "Implement the Environmental Code by regulating classified installations, environmental impact "
        "assessment, water pollution, water police, air pollution and noise pollution — providing an "
        "indirect regulatory framework for plastic pollution through ICPE rules, discharges, "
        "incineration and EIA for waste-treatment projects, without plastic-specific measures.",
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret exécutif adopté par le Président de la République le 12 avril 2001, après avis du "
        "Conseil d'État (13 octobre 2000), sur rapport du Ministre de l'environnement, en application "
        "de la Loi n° 2001-01. Il s'agit d'un texte réglementaire d'application (partie réglementaire "
        "du Code), et non d'une loi adoptée par le Parlement.",
        "Executive decree adopted by the President of the Republic on 12 April 2001, after opinion of "
        "the Council of State (13 October 2000), on the report of the Minister of the Environment, "
        "implementing Law No. 2001-01. It is sub-legislative implementing regulation (regulatory part "
        "of the Code), not an act enacted by Parliament.",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "environnement, industrie, transport, municipalités, pêche, eau, santé, urbanisme, agriculture, "
        "gestion des déchets, emballage",
        "environment, industry, transport, municipalities, fisheries, water, health, urban planning, "
        "agriculture, waste management, packaging",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "production, consommation, recyclage, élimination, fuite environnementale",
        "production, consumption, recycling, disposal, environmental leakage",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Art. R26: {q('les droits et taxes prévus pour les installations classées doivent être acquittés dans un délai de quarante cinq (45) jours après l émission du bulletin de liquidation', 'the rights and taxes provided for classified installations must be paid within forty-five (45) days after issue of the assessment notice')}. "
        f"Art. R32: {q('La taxe superficiaire est due par toute installation classée', 'The surface tax is payable by every classified installation')}. "
        f"Art. R54: {q('la taxe à payer par l exploitant est fixé', 'the tax payable by the operator is set')} sur la base du degré de pollution des effluents / based on the degree of effluent pollution."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. R3–R4, R5, R16–R17: {q('Les installations classées... doivent selon le cas faire l objet d une demande d autorisation... ou faire l objet d une déclaration', 'Classified installations must, as the case may be, be subject to an authorisation request... or to a declaration')} "
            f"({bi('Art. R3', 'Art. R3')}). "
            f"{q('Le Ministre chargé de l environnement délivre au déclarant un récépissé', 'The Minister of the Environment issues the declarant a receipt')} ({bi('Art. R17', 'Art. R17')}). "
            f"{bi('S applique aux usines de transformation du plastique inscrites à la nomenclature ICPE', 'Applies to plastic processing plants listed in the ICPE nomenclature')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Art. R22: {q('veillent à l application des présentes dispositions... surveillance et le contrôle administratif et technique', 'shall ensure application of these provisions... administrative and technical surveillance and control')}. "
            f"Art. R24–R25: {q('sanctions pénales', 'criminal penalties')} and {q('procès-verbaux', 'official reports')} by sworn agents. "
            f"Art. R20–R21: {q('habilitées et assermentées', 'authorised and sworn')} inspectors under the Environment Minister."
        ),
        "comments": bi(
            "Instrument transversal pour les installations de production et de traitement des matières plastiques. "
            "La nomenclature ICPE est fixée par arrêté ministériel (R4).",
            "Cross-cutting instrument for plastic production and treatment installations. ICPE nomenclature is set by ministerial order (R4).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. R9–R11: {q('doit faire l objet d une étude d impact préalable permettant d évaluer les incidences directes ou indirectes... sur l équilibre écologique', 'must be subject to a prior environmental impact study assessing direct or indirect effects... on the ecological balance')} "
            f"for Class I installations with significant environmental incidence. "
            f"Art. R11: {q('La décision sur l étude d impact fait l objet d arrêté ministériel qui est publié au Journal Officiel', 'The decision on the impact study is made by ministerial order published in the Official Journal')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R6–R7: {q('enquête publique... durée de 15 jours', 'public inquiry... duration of 15 days')}. "
            f"Art. R10: mandatory EIA content including emissions and waste effects. "
            f"Art. R11: {q('délai de deux semaines maximum', 'maximum period of two weeks')} for Ministry opinion."
        ),
        "comments": bi(
            "Couvre les grandes installations plastiques de classe I. L EIE préalable est distincte du régime EIE général du Titre II (R38+).",
            "Covers major Class I plastic installations. Prior EIA is distinct from the general EIA regime in Title II (R38+).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Art. R16: Class II installation declarations must specify "
            f"{q('le mode et les conditions d utilisation, d épuration et d évacuation des eaux résiduaires et des émanations de toute nature ainsi que d élimination des déchets et résidus de l exploitation', 'the mode and conditions of use, treatment and discharge of wastewater and emissions of any kind, as well as elimination of waste and residues from the operation')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R17–R18: {q('délivre au déclarant un récépissé', 'issues the declarant a receipt')} within two months; "
            f"{q('prescriptions générales', 'general requirements')} by ministerial order. "
            f"Art. R22: inspection and control of declared installations."
        ),
        "comments": bi(
            "Obligation d information sur l élimination des déchets/résidus pour les ICPE de 2e classe, incluant ateliers de transformation plastique.",
            "Disclosure obligation on waste/residue elimination for Class II ICPE, including plastic processing workshops.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. R26, R32, R54: {q('les droits et taxes prévus pour les installations classées doivent être acquittés', 'the rights and taxes provided for classified installations must be paid')} (R26); "
            f"{q('La taxe superficiaire est due par toute installation classée', 'The surface tax is payable by every classified installation')} (R32); "
            f"effluent pollution tax set according to pollution degree (R54)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R26: {q('délai de quarante cinq (45) jours', 'period of forty-five (45) days')} for payment after assessment notice; "
            f"{q('pénalités pécuniaires', 'financial penalties')} for violations. "
            f"Art. R32: tax amount specified in authorisation or declaration receipt."
        ),
        "comments": bi(
            "Instrument économique général pour les ICPE et les rejets; s applique aux installations plastiques classées et aux effluents industriels.",
            "General economic instrument for ICPE and discharges; applies to classified plastic installations and industrial effluents.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. R38–R44: {q('Les études d impact régies par le présent décret sont réalisées préalablement à toute autorisation administrative exigée pour la réalisation de l activité envisagée', 'Impact studies governed by this decree are carried out prior to any administrative authorisation required for the planned activity')} (R38). "
            f"Art. R39 lists {q('les effets du recyclage et de l élimination des résidus et des déchets', 'the effects of recycling and elimination of residues and waste')} as EIA aspects. "
            f"Art. R42 includes {q('Traitement et stockage des déchets', 'Waste treatment and storage')} as an accredited consultant category."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R40: projects classified Category 1 (significant impacts) or Category 2 (limited impacts) in annexes. "
            f"Art. R43–R44: {q('comité technique', 'technical committee')} reviews EIA reports monthly; Minister decides within 15 days. "
            f"Operationalised via Arrêtés ministériels n° 9468–9472 (28 November 2001)."
        ),
        "comments": bi(
            "Cadre EIE applicable aux usines de recyclage plastique, décharges et unités de valorisation listées en annexe. Pas de cible quantitative plastique.",
            "EIA framework applicable to plastic recycling plants, landfills and recovery units listed in annexes. No quantitative plastic target.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. R49, R51: {q('L effluent rejeté ne doit en aucun cas entraîner la détérioration du milieu récepteur', 'Discharged effluent must in no case cause deterioration of the receiving environment')} (R49). "
            f"{q('L autorisation de rejeter des effluents est conditionnée par les résultats de l étude d impact... et par le respect des normes physiques, chimiques, biologiques et bactériologiques', 'Authorisation to discharge effluents is conditional on EIA results... and compliance with physical, chemical, biological and bacteriological standards')} (R51)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R50: {q('Une étude d impact est exigée de tout exploitant voulant utiliser les milieux récepteurs naturels pour effectuer des rejets d effluents', 'An impact study is required of any operator wishing to use natural receiving environments to discharge effluents')}. "
            f"Art. R52–R53: sworn agents, sampling and NS compliance. Standards operationalised via NS 05-061 and joint ministerial orders."
        ),
        "comments": bi(
            "Pertinent pour les effluents industriels plastiques et les microparticules dans les eaux usées. Conditions de rejet fixées par arrêté conjoint (R49).",
            "Relevant for plastic industrial effluents and microplastics in wastewater. Discharge conditions set by joint ministerial order (R49).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. R56(a): {q('Sont interdits au titre de la police de l eau: a) tous déversements, écoulements, dépôts directs ou indirects, tout fait en général susceptible de polluer les eaux continentales ou marines', 'Prohibited under water police: a) all discharges, flows, direct or indirect deposits, or any act likely to pollute continental or marine waters')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R59–R66: {q('officiers de police judiciaire et les agents assermentés', 'judicial police officers and sworn agents')} may record violations; "
            f"{q('procès-verbal', 'official report')} procedure and transaction fines. "
            f"Art. R69: vessel detention for pollution offences."
        ),
        "comments": bi(
            "Interdiction directe applicable aux rejets de déchets plastiques dans les eaux. Complète l Art. L41 du Code (partie législative).",
            "Direct prohibition applicable to plastic waste discharges into waters. Complements Art. L41 of the Code (legislative part).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. R56(b): {q('tous rejets à partir de la côte d eaux et de toutes substances usées, de déchets industriels, de toutes substances solides ou liquides toxiques pouvant entraîner la pollution des plages et des zones littorales', 'all discharges from the coast of water and all used substances, industrial waste, and all toxic solid or liquid substances that may cause pollution of beaches and coastal zones')} are prohibited."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R58: {q('contrôle trimestriel des zones de baignade', 'quarterly monitoring of bathing areas')}; "
            f"{q('interdisent purement et simplement la baignade', 'bathing is strictly prohibited')} if pollution found. "
            f"Marine enforcement agents listed in Art. R60–R61."
        ),
        "comments": bi(
            "Pertinent pour la pollution plastique littorale et les déchets industriels en zone côtière.",
            "Relevant for coastal plastic pollution and industrial waste in coastal zones.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. R57: {q('les rejets ou immersions à partir des navires de déchets industriels, de substances liquides ou de mélanges contenant de telles substances peuvent être autorisés dans des cas limitativement prévus par arrêté conjoint', 'discharges or disposals from ships of industrial waste, liquid substances or mixtures containing such substances may be authorised in cases limitatively provided for by joint ministerial order')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. R57: {q('peuvent être autorisés... par arrêté conjoint des Ministres chargés respectivement de l environnement et de la Marine marchande', 'may be authorised... by joint order of the Ministers of the Environment and Merchant Marine')}. "
            f"Enabling power only; operative conditions depend on subordinate arrêté and MARPOL compliance."
        ),
        "comments": bi(
            "Pouvoir habilitant, non une interdiction directe. Score in_force=0 sauf si un arrêté conjoint spécifique a été adopté et est en vigueur.",
            "Enabling power, not a direct ban. in_force=0 unless a specific joint order has been adopted and is operative.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. R71–R73: {q('les exploitants de ces installations doivent mettre en œuvre toutes les dispositions utiles pour supprimer ou réduire leurs émissions polluantes', 'operators of these installations must implement all useful measures to eliminate or reduce their polluting emissions')} when meteorological conditions threaten air quality (R72). "
            f"Art. R73: interministerial orders may impose {q('l interdiction de l usage de certains produits chimiques, le ralentissement ou l arrêt du fonctionnement de certains appareils', 'prohibition on use of certain chemicals, slowing or stopping of certain equipment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. R71–R73: applies to {q('installations fixes pouvant engendrer des émissions polluantes', 'fixed installations that may generate polluting emissions')}. "
            f"Art. R79: sworn agents may access combustion/incineration equipment. Enforcement via R82 contraventions."
        ),
        "comments": bi(
            "S applique aux émissions des installations de production et de traitement plastique. Art. R73 renvoie à des arrêtés pour les prescriptions détaillées.",
            "Applies to emissions from plastic production and treatment installations. Art. R73 refers to orders for detailed prescriptions.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. R74: {q('peuvent prescrire toutes mesures utiles en vue de limiter la pollution atmosphérique résultant de la combustion de certaines matières en dehors de toute installation appropriée', 'may prescribe all useful measures to limit atmospheric pollution resulting from combustion of certain materials outside any appropriate installation')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. R74: {q('Des arrêtés pris conjointement par les Ministres chargés respectivement de l environnement, de la santé, de l agriculture et de l industrie peuvent prescrire', 'Joint orders of the Ministers of the Environment, Health, Agriculture and Industry may prescribe')} measures. Enabling language ('peuvent'); binding prescriptions require subordinate arrêtés."
        ),
        "comments": bi(
            "Pertinent pour la combustion à l air libre de déchets plastiques, mais pouvoir habilitant non encore opérationnalisé dans le texte même.",
            "Relevant to open burning of plastic waste, but enabling power not self-operational in the decree text.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. R75–R78: applies to {q('installations fixes d incinération, de combustion ou de chauffage', 'fixed incineration, combustion or heating installations')}. "
            f"Art. R76: orders may set {q('des spécifications techniques auxquelles doivent répondre, pour pouvoir être fabriqués, importés ou mis en vente sur le marché sénégalais, des matériels d incinération, de combustion ou de chauffage', 'technical specifications that incineration, combustion or heating equipment must meet to be manufactured, imported or placed on the Senegalese market')}. "
            f"Art. R78: {q('visite périodique par un expert ou un organisme agréé', 'periodic inspection by an accredited expert or body')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. R76–R77: {q('peuvent fixer', 'may set')} technical specifications and operating conditions; "
            f"Art. R78: {q('soumises à une visite périodique', 'subject to periodic inspection')}. "
            f"Authority and monitoring framework explicit; detailed standards via arrêtés (enabling elements scored in_force=0 for mandatory standards themselves)."
        ),
        "comments": bi(
            "Instrument mixte: cadre d incinération pertinent pour les déchets plastiques, mais la plupart des prescriptions techniques sont renvoyées à des arrêtés ('peuvent'). in_force=0 pour les normes détaillées; cadre d inspection (R78) partiellement opérationnel si agrément en place.",
            "Mixed instrument: incineration framework relevant to plastic waste, but most technical standards are deferred to orders ('may'). in_force=0 for detailed standards; inspection framework (R78) partially operative if accreditation exists.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. R80–R82: {q('Des zones de protection spéciale peuvent être créées et délimitées par des arrêtés pris conjointement', 'Special protection zones may be created and delimited by joint ministerial orders')} based on air-quality parameters including particulate concentrations (R80–R81)."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            f"Art. R80: {q('peuvent être créées', 'may be created')} — enabling power. "
            f"Art. R82: contravention penalties apply once zones and measures are designated by order."
        ),
        "comments": bi(
            "Zones de protection spéciale pour la qualité de l air; pertinence indirecte pour les émissions de combustion de plastiques. in_force=0 tant qu aucune zone n est désignée par arrêté.",
            "Special air-quality protection zones; indirect relevance to plastic combustion emissions. in_force=0 until a zone is designated by order.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Art. R27: {q('Le Ministre chargé de l environnement peut déléguer son pouvoir d octroi de l autorisation d exploitation ou du récépissé de déclaration au Gouverneur de la Région', 'The Minister of the Environment may delegate the power to grant operating authorisation or declaration receipts to the Regional Governor')} where regional environment services exist."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. R7: regional coordination via {q('Comité Régional de Développement', 'Regional Development Committee')} and municipal/rural council opinions during Class I inquiries. "
            f"Art. R27: delegation mechanism with reporting obligation to the Minister."
        ),
        "comments": bi(
            "Instrument de gouvernance et de coordination national-régional-municipal pour les ICPE, incluant les installations de gestion des déchets.",
            "National–regional–municipal governance/coordination instrument for ICPE, including waste-management installations.",
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
        "A": 48, "B": 52, "C": 10, "D": 52, "E": 12, "F": 20,
        "G": 10, "H": 48, "I": 14, "J": 40, "K": 14, "L": 36,
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
