#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for NS 05-061 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_NS_05-061_Eaux_Usees.xlsx"

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
        "Norme sénégalaise NS 05-061 — Rejets d'eaux usées (juillet 2001)",
        "Senegalese Standard NS 05-061 — Wastewater discharges (July 2001)",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/74/normes-cadre-juridique/"
        "19105/norme-rejets-ns-05-061-eaux-usees-juillet-2001.pdf"
    ),
    "policy_year": 2001,
    "policy_objective": bi(
        "Fixer les valeurs limites, interdictions, modalités de surveillance et critères de rejet des eaux "
        "usées industrielles et domestiques dans les milieux récepteurs (surface, souterrain, marin), en "
        "complément du Décret n° 2001-282 — contrôler indirectement les effluents industriels plastiques "
        "via les paramètres physico-chimiques (MES, DCO, DBO5, COT, micropolluants) sans mesures "
        "spécifiques aux plastiques.",
        "Set limit values, prohibitions, monitoring modalities and discharge criteria for industrial and "
        "domestic wastewater into receiving environments (surface, groundwater, marine), complementing "
        "Decree No. 2001-282 — indirectly controlling plastic-related industrial effluents through "
        "physico-chemical parameters (TSS, COD, BOD5, TOC, micropollutants) without plastic-specific measures.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('Matières en suspension totales: 50 mg/l', 'Total suspended solids: 50 mg/l')}; "
        f"{q('DBO5 (sur effluent non décanté) : 80 mg/l si le flux journalier maximal autorisé n excède pas 30 kg/j; 40 mg/l au-delà', 'BOD5 (on non-settled effluent): 80 mg/l if maximum authorised daily load does not exceed 30 kg/d; 40 mg/l beyond that')}; "
        f"{q('DCO (sur effluent non décanté) : 200 mg/l si le flux journalier maximal autorisé n excède pas 100 kg/j; 100 mg/l au-delà', 'COD (on non-settled effluent): 200 mg/l if maximum authorised daily load does not exceed 100 kg/d; 100 mg/l beyond that')}; "
        f"{q('Azote... 30 mg/l en concentration moyenne mensuelle lorsque le flux journalier maximal est égal ou supérieur à 50 kg/jour', 'Nitrogen... 30 mg/l as monthly average concentration when maximum daily load is equal to or greater than 50 kg/day')}; "
        f"{q('Phosphore (phosphore total) : 10 mg/l en concentration moyenne mensuelle lorsque le flux journalier maximal autorisé est égal ou supérieur à 15 kg/jour', 'Phosphorus (total phosphorus): 10 mg/l as monthly average concentration when maximum authorised daily load is equal to or greater than 15 kg/day')}; "
        f"{q('hydrocarbures totaux 15 mg/l si le rejet dépasse 150 g/j', 'total hydrocarbons 15 mg/l if discharge exceeds 150 g/d')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Norme technique officielle adoptée en juillet 2001 par le Comité technique de normalisation "
        "ISN/CT5 (Environnement et Ressources naturelles), éditée par l'Institut sénégalais de "
        "Normalisation (ISN), et explicitement présentée comme complétant le Décret n° 2001-282 "
        "d'application du Code de l'environnement. Elle n'est pas une loi parlementaire (≠1), mais "
        "constitue la réglementation technique opérationnelle de référence pour les rejets d'effluents, "
        "avec des valeurs limites contraignantes lorsqu'elle est invoquée par le cadre juridique "
        "environnemental.",
        "Official technical standard adopted in July 2001 by the ISN/CT5 standardisation technical "
        "committee (Environment and Natural Resources), published by the Senegalese Standardisation "
        "Institute (ISN), and explicitly presented as complementing Decree No. 2001-282 implementing "
        "the Environmental Code. It is not parliamentary legislation (≠1), but constitutes the operative "
        "technical regulation of reference for effluent discharges, with binding limit values when invoked "
        "by the environmental legal framework.",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "industrie, assainissement, environnement, eau, pêche, agriculture, gestion des déchets",
        "industry, sanitation, environment, water, fisheries, agriculture, waste management",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "production, élimination, fuite environnementale",
        "production, disposal, environmental leakage",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        f"Ch. 3, §1.4–1.5: {q('détermination du degré de pollution des effluents, et du taux de la taxe à payer par l exploitant', 'determination of the degree of effluent pollution, and of the tax rate payable by the operator')}; "
        f"{q('Le calcul de la redevance est donné dans l arrêté interministériel relatif à l application de la présente norme', 'The fee calculation is set out in the interministerial order on application of this standard')}. "
        f"Annexe I (fin): {q('(MES – 50) +[(DCO - 200) +2 (DBO5 – 80) ]/3= X mg/l', '(TSS – 50) +[(COD - 200) +2 (BOD5 – 80) ]/3= X mg/l')} for pollution load/redevance."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"§5.1: {q('Tout rejet d effluents liquides entraînant des stagnations, des incommodités pour le voisinage, ou des pollutions des eaux de surface, souterraines ou marines est interdit sur toute l étendue du territoire national', 'Any discharge of liquid effluents causing stagnation, neighbourhood nuisance, or pollution of surface, groundwater or marine waters is prohibited throughout the national territory')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. 3 §1.1–1.2: {q('prélèvements et analyses... effectués par des techniciens de la Direction de l Environnement ou par toute personne ou entité désignée', 'sampling and analyses... carried out by technicians of the Environment Directorate or any person or entity designated')} by the Environment Ministry. "
            f"§1.7: enhanced monitoring when Annex II thresholds exceeded."
        ),
        "comments": bi(
            "Interdiction générale applicable aux rejets d effluents industriels plastiques. Pas de mention explicite des plastiques.",
            "General prohibition applicable to plastic-industry effluent discharges. No explicit mention of plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"§5.2 & Annexe II §2: {q('Tout effluent traité, pour pouvoir être rejeté dans un milieu récepteur, doit respecter les valeurs indiquées à l annexe II', 'Any treated effluent, to be discharged into a receiving environment, must comply with the values set out in Annex II')}. "
            f"Quantitative limits include {q('Matières en suspension totales: 50 mg/l', 'Total suspended solids: 50 mg/l')}, "
            f"{q('DBO5... 80 mg/l... 40 mg/l', 'BOD5... 80 mg/l... 40 mg/l')}, "
            f"{q('DCO... 200 mg/l... 100 mg/l', 'COD... 200 mg/l... 100 mg/l')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Ch. 3 §1.3–1.5: mandatory sampling before discharge; minimum {q('deux fois par an', 'twice per year')}; "
            f"§1.7: continuous flow measurement when daily flow >100 m³. "
            f"§1.4: results determine {q('le taux de la taxe à payer par l exploitant', 'the tax rate payable by the operator')}."
        ),
        "comments": bi(
            "Instrument central. Les MES peuvent capturer des particules plastiques en suspension dans les effluents industriels, sans ciblage spécifique.",
            "Core instrument. TSS may capture suspended plastic particles in industrial effluents, without specific targeting.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"§5.3: Prohibited discharges include "
            f"{q('tous déversements de composés cycliques hydroxylés et de leurs dérivés halogénés', 'all discharges of hydroxylated cyclic compounds and their halogenated derivatives')}, "
            f"{q('tous déversements d hydrocarbures ou autres produits chimiques, toxiques', 'all discharges of hydrocarbons or other toxic chemical products')}, "
            f"{q('tout déversement dans les lacs, étangs et mares', 'any discharge into lakes, ponds and pools')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. 3 §1.1–1.2: Environment Directorate technicians responsible for resolving water pollution problems; "
            f"must have {q('matériels et moyens nécessaires pour les prélèvements et analyses', 'equipment and means necessary for sampling and laboratory analyses')}."
        ),
        "comments": bi(
            "Interdictions de substances chimiques; pertinence indirecte pour additifs plastiques et solvants industriels.",
            "Chemical substance prohibitions; indirect relevance for plastic additives and industrial solvents.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Ch. 2 §II: {q('l effluent doit être débarrassé de tous produits susceptibles de nuire à la conservation des ouvrages, ainsi que des matières flottantes, déposables ou précipitables', 'the effluent must be cleared of all products likely to harm the integrity of infrastructure, as well as floating, settleable or precipitable matter')} before discharge via public evacuation channels without treatment plant."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. 2 §II: applies when discharging through {q('canal public d évacuation sans station d épuration', 'public evacuation channel without treatment plant')}. "
            f"Ch. 3 monitoring regime applies. Connection protocols required for non-domestic effluent (Ch. 2 §II)."
        ),
        "comments": bi(
            "Pertinent pour les particules/matières plastiques flottantes ou précipitables dans les effluents industriels.",
            "Relevant for floating or precipitable plastic particles/matter in industrial effluents.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            f"Ch. 2 §II: {q('Tout branchement d un réseau d effluent autre que domestique, au réseau public (municipal...) muni de station d épuration, doit faire l objet d un protocole d accord entre le générateur... et le gestionnaire de la station', 'Any connection of a non-domestic effluent network to a public (municipal...) network with a treatment plant must be subject to an agreement protocol between the generator... and the plant operator')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Ch. 2 §II: {q('Au cas où le rejet renferme des produits chimiques toxiques, des valeurs plus contraignantes seront appliquées au rejet', 'Where the discharge contains toxic chemical products, more stringent values will apply to the discharge')}. "
            f"Annexe II §3: convention sets maximum effluent characteristics and self-monitoring obligations."
        ),
        "comments": bi(
            "Gouvernance des branchements d effluents industriels (incl. plasturgie) aux réseaux d assainissement.",
            "Governance of industrial effluent connections (incl. plastics processing) to sanitation networks.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"§4.1.2 & Annexe I: {q('Tous les émissaires d évacuations des eaux usées traitées... doivent être équipés de dispositifs pour permettre un échantillonnage adéquat et une mesure de débit normalisée', 'All outfalls for treated wastewater... must be equipped with devices allowing adequate sampling and standardised flow measurement')}. "
            f"Annexe I §4: {q('Sur chaque canalisation de rejet d effluents, doivent être prévus des points de prélèvement d échantillons, de mesure de débit', 'On each effluent discharge pipeline, sampling points and flow measurement points must be provided')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe I §6: {q('dispositifs normalisés de mesure de débit', 'standardised flow measurement devices')}; "
            f"{q('faciliter l intervention d organismes extérieurs habilités', 'facilitate intervention by authorised external bodies')}. "
            f"Ch. 3 §1.7.1: continuous flow measurement when daily flow >100 m³."
        ),
        "comments": bi(
            "Infrastructure de contrôle des rejets; condition préalable à la vérification de conformité des effluents plastiques.",
            "Effluent monitoring infrastructure; prerequisite for verifying compliance of plastic-related effluents.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Annexe II §2 iii: Micropollutant limits including "
            f"{q('hydrocarbures totaux 15 mg/l si le rejet dépasse 150 g/j', 'total hydrocarbons 15 mg/l if discharge exceeds 150 g/d')}, "
            f"{q('chrome hexavalent 0,2 mg/l si le rejet dépasse 5 g/j', 'hexavalent chromium 0.2 mg/l if discharge exceeds 5 g/d')}, "
            f"{q('cyanures 0,2 mg/l si le rejet dépasse 3 g/j', 'cyanides 0.2 mg/l if discharge exceeds 3 g/d')}, etc."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe II §2 iii: authorised ICPE operators must submit {q('chaque année... un dossier faisant le bilan des rejets', 'each year... a report on discharges')} including flows, concentrations and reduction possibilities. "
            f"Ch. 3 §1.7.2: daily measurement when authorised daily flux thresholds exceeded."
        ),
        "comments": bi(
            "Seuils quantitatifs pour micropolluants industriels; pas de seuil spécifique microplastiques.",
            "Quantitative thresholds for industrial micropollutants; no microplastic-specific threshold.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Tableau 1 (Annexe II): Stricter limits for {q('Milieux spécialement protégés', 'Specially protected environments')} — e.g. "
            f"DCO {q('90 mg/l', '90 mg/l')} vs 200 mg/l; MES {q('30 mg/l', '30 mg/l')} vs 40 mg/l; DBO5 {q('20 mg/l', '20 mg/l')} vs 50 mg/l. "
            f"§2.1: {q('les rejets d eau sur les milieux suivants sont interdits : lacs, étangs, mares et réserves d eau', 'water discharges into the following environments are prohibited: lakes, ponds, pools and water reserves')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. 3 §1.6: {q('milieux récepteurs, sous protection spéciale, font l objet d une surveillance et d un contrôle des eaux plus réguliers', 'receiving environments under special protection are subject to more regular water surveillance and control')}. "
            f"§2.3: quarterly analysis reports over four years for sensitive bays (e.g. Baie de Hann)."
        ),
        "comments": bi(
            "Seuils renforcés en zones sensibles (littoral, eau potable, conchyliculture). Pertinent pour rejets plastiques en milieu marin.",
            "Strengthened thresholds in sensitive zones (coast, drinking water, shellfish farming). Relevant for plastic discharges in marine environments.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Annexe II §3: Pre-treatment limits before connection to collective treatment when flux exceeds thresholds — "
            f"{q('MEST : 600 mg/l', 'TSS: 600 mg/l')}, {q('DB05 : 800 mg/l', 'BOD5: 800 mg/l')}, {q('DCO : 2 000 mg/l', 'COD: 2,000 mg/l')}, "
            f"{q('Azote total... 150 mg/l', 'Total nitrogen... 150 mg/l')}, {q('Phosphore total... 50 mg/l', 'Total phosphorus... 50 mg/l')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe II §3: applies when flux may exceed {q('15 kg/j de MEST ou 15 kg/j de DBO ou 45 kg/j de DCO', '15 kg/d TSS or 15 kg/d BOD or 45 kg/d COD')}. "
            f"{q('Tout raccordement doit faire l objet d une convention préalable', 'Any connection must be subject to a prior agreement')} with treatment plant operator."
        ),
        "comments": bi(
            "Limites de prétraitement pour effluents industriels avant station d épuration collective (usines plastiques raccordées).",
            "Pre-treatment limits for industrial effluents before collective treatment plants (connected plastics factories).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Annexe II (Tableau 1 note): {q('Lorsque le rejet maximal de DCO dépasse 2 t/j, la mesure en continu du COT (carbone organique total) doit être réalisée', 'When maximum COD discharge exceeds 2 t/d, continuous measurement of TOC (total organic carbon) must be carried out')}; "
            f"{q('0,5 t/j', '0.5 t/d')} threshold for specially protected environments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe II: {q('mesure en continu du COT', 'continuous TOC measurement')} required above flux thresholds; "
            f"{q('mesures journalières de DCO poursuivies parallèlement... durée minimale d un an', 'daily COD measurements pursued in parallel... minimum duration of one year')} to establish COT/COD correlation."
        ),
        "comments": bi(
            "Exigence de mesure continue du carbone organique total pour grands rejets industriels; pertinent pour pollution organique liée aux procédés plastiques.",
            "Continuous total organic carbon measurement requirement for large industrial discharges; relevant to organic pollution from plastic processes.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Ch. 3 §1.4–1.5 & Annexe I (fin): Pollution fee/redevance linked to effluent quality. Formula: "
            f"{q('(MES – 50) +[(DCO - 200) +2 (DBO5 – 80) ]/3= X mg/l', '(TSS – 50) +[(COD - 200) +2 (BOD5 – 80) ]/3= X mg/l')} multiplied by water volume gives daily pollution load in kg/day."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. 3 §1.5: {q('Le calcul de la redevance peut se faire par la moyenne des deux prélèvements', 'The fee may be calculated from the average of the two samples')}; "
            f"{q('donné dans l arrêté interministériel relatif à l application de la présente norme', 'set out in the interministerial order on application of this standard')}."
        ),
        "comments": bi(
            "Instrument économique: redevance proportionnelle à la charge polluante (MES/DCO/DBO5). Détails dans arrêté interministériel séparé.",
            "Economic instrument: fee proportional to pollution load (TSS/COD/BOD5). Details in separate interministerial order.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Annexe III §1: {q('L épandage d effluents ou de boues contenant des substances qui, du fait de leur toxicité, de leur persistance ou de leur bio-accumulation, sont susceptibles d être dangereuses pour l environnement, est interdit', 'Spreading of effluents or sludge containing substances that, due to their toxicity, persistence or bio-accumulation, may be dangerous for the environment, is prohibited')}. "
            f"{q('le déversement dans le milieu naturel des trop-pleins des ouvrages de stockage est interdit', 'discharge into the natural environment from storage facility overflows is prohibited')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe III: storage capacity for {q('production de pointe de 15 jours', '15-day peak production')}; "
            f"{q('ouvrages de stockage doivent être étanches', 'storage facilities must be watertight')}; "
            f"spreading conditions set by ICPE authorisation order."
        ),
        "comments": bi(
            "Interdit l épandage de substances persistantes/bioaccumulables; pertinence indirecte pour certains additifs plastiques et boues de traitement.",
            "Prohibits spreading of persistent/bioaccumulative substances; indirect relevance for certain plastic additives and treatment sludge.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Tableau 2 (Annexe II): Microbiological effluent quality limits — e.g. "
            f"{q('Coliformes Fécaux... 2000 Par 100 ml', 'Faecal coliforms... 2000 per 100 ml')}, "
            f"{q('Salmonelles Par 5 000 ml Absence', 'Salmonella per 5,000 ml Absence')} for public maritime and hydraulic domains."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Ch. 3 §1.3: sampling and analysis conditions in Annexe IV. "
            f"§1.4: analyses must cover {q('caractéristiques physiques, chimiques, bactériologiques', 'physical, chemical and bacteriological characteristics')} as applicable."
        ),
        "comments": bi(
            "Normes microbiologiques pour rejets en domaine public maritime/hydraulique; complément aux paramètres physico-chimiques.",
            "Microbiological standards for discharges in public maritime/hydraulic domains; complement to physico-chemical parameters.",
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
