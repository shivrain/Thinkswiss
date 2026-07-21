#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for NS 05-062 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_NS_05-062_Pollution_Atmosphérique.xlsx"

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
        "Norme sénégalaise NS 05-062 — Pollution atmosphérique (octobre 2003)",
        "Senegalese Standard NS 05-062 — Atmospheric pollution (October 2003)",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/74/normes-cadre-juridique/"
        "19106/norme-rejets-ns05-062-pollution-atmospherique-octobre-2003.pdf"
    ),
    "policy_year": 2003,
    "policy_objective": bi(
        "Fixer les valeurs limites d'émissions et d'immissions atmosphériques, les interdictions de brûlage "
        "à l'air libre, les exigences d'incinération des déchets et les modalités de surveillance des effluents "
        "gazeux des installations stationnaires et des véhicules, en complément du Décret n° 2001-282 — "
        "contrôler indirectement les émissions industrielles liées à la production plastique, au traitement de "
        "surface des matières plastiques et à l'élimination thermique des déchets plastiques.",
        "Set atmospheric emission and immission limit values, open-burning prohibitions, waste incineration "
        "requirements and monitoring modalities for gaseous effluents from stationary installations and vehicles, "
        "complementing Decree No. 2001-282 — indirectly controlling industrial emissions related to plastic "
        "production, plastic surface treatment and thermal disposal of plastic waste.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('Poussières totales : 100 mg/m3 (D <= 1 kg/h); 50 mg/m3 (D > 1 kg/h)', 'Total dust: 100 mg/m3 (D <= 1 kg/h); 50 mg/m3 (D > 1 kg/h)')}; "
        f"{q('Oxydes de soufre (exprimés en SO2) : 500 mg/m3 (D > 25 kg/h)', 'Sulphur oxides (as SO2): 500 mg/m3 (D > 25 kg/h)')}; "
        f"{q('Oxydes d azote (exprimés en NO2) : 500 mg/m3 (D > 25 kg/h)', 'Nitrogen oxides (as NO2): 500 mg/m3 (D > 25 kg/h)')}; "
        f"{q('Composés organiques totaux : 150 mg/m3 (D > 2 kg/h)', 'Total organic compounds: 150 mg/m3 (D > 2 kg/h)')}; "
        f"{q('Incinération de déchets — poussières : 10 mg/m3; matières organiques (COT) : 50 mg/m3', 'Waste incineration — dust: 10 mg/m3; organic matter (TOC): 50 mg/m3')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Norme technique officielle adoptée en octobre 2003 par le Comité technique de normalisation "
        "ASN/CT5 (Environnement et Ressources naturelles), éditée par l'Association sénégalaise de "
        "Normalisation (ASN), et explicitement présentée comme complétant le Décret n° 2001-282 "
        "d'application du Code de l'environnement. Elle n'est pas une loi parlementaire (≠1), mais "
        "constitue la réglementation technique opérationnelle de référence pour les rejets atmosphériques, "
        "avec des valeurs limites contraignantes lorsqu'elle est invoquée par le cadre juridique "
        "environnemental.",
        "Official technical standard adopted in October 2003 by the ASN/CT5 standardisation technical "
        "committee (Environment and Natural Resources), published by the Senegalese Standardisation "
        "Association (ASN), and explicitly presented as complementing Decree No. 2001-282 implementing "
        "the Environmental Code. It is not parliamentary legislation (≠1), but constitutes the operative "
        "technical regulation of reference for atmospheric discharges, with binding limit values when "
        "invoked by the environmental legal framework.",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "industrie, environnement, énergie, transport, gestion des déchets",
        "industry, environment, energy, transport, waste management",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "production, élimination, fuite environnementale",
        "production, disposal, environmental leakage",
    ),
    "policy_budget": 0,
    "policy_budget_text": bi(
        "Aucune disposition budgétaire, redevance ou instrument économique explicite dans la norme. "
        "Les frais de surveillance et de mesure sont à la charge de l'exploitant (Ch. V, §5).",
        "No explicit budget, fee or economic instrument in the standard. Monitoring and measurement "
        "costs are borne by the operator (Ch. V, §5).",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Ch. II §1: {q('Les installations existantes et nouvelles stationnaires doivent être équipées et exploitées de manière à respecter la limitation maximale des émissions fixée aux annexes I, II, III', 'Existing and new stationary installations must be equipped and operated so as to comply with the maximum emission limits set out in Annexes I, II and III')}. "
            f"§1.1.1: {q('Les émissions sont captées aussi complètement et aussi près que possible de leur source', 'Emissions must be captured as completely and as close to their source as possible')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"§4.1: {q('L autorité compétente s assure que les valeurs limites maximales des émissions sont respectées', 'The competent authority ensures that maximum emission limit values are complied with')}; "
            f"§4.2: continuous measurement ordered for installations with significant emissions. "
            f"Ch. V: operator monitoring programme with quarterly reporting."
        ),
        "comments": bi(
            "Cadre général applicable aux émissions gazeuses des installations industrielles (incl. plasturgie).",
            "General framework applicable to gaseous emissions from industrial installations (incl. plastics processing).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Annexe I, tableau général: quantitative emission limits including "
            f"{q('Poussières totales... 100 mg/m3... 50 mg/m3', 'Total dust... 100 mg/m3... 50 mg/m3')}, "
            f"{q('Oxydes de soufre... 500 mg/m3', 'Sulphur oxides... 500 mg/m3')}, "
            f"{q('Oxydes d Azote... 500 mg/m3', 'Nitrogen oxides... 500 mg/m3')}, "
            f"{q('Rejet total en composés organiques... 150 mg/m3', 'Total organic compound emissions... 150 mg/m3')}, "
            f"{q('HAP... 20 mg/m3', 'PAHs... 20 mg/m3')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"§6.4: {q('aucune moyenne journalière n est supérieure à la valeur limite; aucune moyenne horaire ne dépasse le double de la valeur limite', 'no daily average exceeds the limit value; no hourly average exceeds twice the limit value')} for continuous monitoring. "
            f"§5.1: measurements per Senegalese atmospheric pollution analysis standards (Annexe V)."
        ),
        "comments": bi(
            "Limites quantitatives centrales pour polluants atmosphériques industriels; COT/HAP pertinents pour procédés plastiques.",
            "Core quantitative limits for industrial air pollutants; TOC/PAHs relevant to plastic processes.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§8.2: {q('Le brûlage à l air libre des pneumatiques, plastiques et tout autre composé renfermant des produits chimiques est interdit', 'Open burning of tyres, plastics and any other compound containing chemical products is prohibited')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. V §5: {q('l autorité administrative compétente peut demander à tout moment la réalisation... de prélèvements et analyses d effluents gazeux. Les frais occasionnés sont à la charge de l exploitant', 'the competent administrative authority may at any time request sampling and analysis of gaseous effluents. Costs are borne by the operator')}."
        ),
        "comments": bi(
            "Mention explicite des plastiques — interdiction directe du brûlage à l air libre (fin de vie, déchets).",
            "Explicit mention of plastics — direct prohibition of open burning (end-of-life, waste).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"§8.1: {q('L incinération ou la décomposition thermique des déchets n est autorisée que dans des installations technologiquement destinées à cet effet', 'Incineration or thermal decomposition of waste is authorised only in installations technologically designed for that purpose')}. "
            f"{q('Les dispositions de l annexe II, lettre J sont applicables', 'The provisions of Annex II, section J apply')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Annexe II §J §3: {q('On doit mesurer et on doit enregistrer en permanence', 'Temperature, oxygen and CO must be measured and recorded continuously')} (temperature, O2, CO). "
            f"§J §5.1: prohibition on incinerating urban/special waste in installations <350 kW thermal power."
        ),
        "comments": bi(
            "Gouvernance de l incinération des déchets (incl. déchets plastiques urbains/spéciaux). Complété par l arrêté interministériel n° 7358 du 5 novembre 2003.",
            "Governance of waste incineration (incl. urban/special plastic waste). Complemented by Interministerial Order No. 7358 of 5 November 2003.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Annexe II §J §2: Waste incineration emission limits — "
            f"{q('Poussières : 10 mg/m3', 'Dust: 10 mg/m3')}, "
            f"{q('Matières organiques sous forme de gaz, exprimées en carbone total : 50 mg/m3', 'Organic matter in gaseous form, expressed as total carbon: 50 mg/m3')}, "
            f"{q('Oxydes de soufre : 50 mg/m3', 'Sulphur oxides: 50 mg/m3')}, "
            f"{q('Oxydes d azote : 80 mg/m3', 'Nitrogen oxides: 80 mg/m3')}, "
            f"{q('Mercure et cadmium : 0,1 mg/m3', 'Mercury and cadmium: 0.1 mg/m3')} per substance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Annexe II §J §3: {q('On doit surveiller en permanence le fonctionnement de l installation d épuration des gaz', 'Continuous monitoring of gas treatment plant operation')} (temperature, pressure drop or scrubber flow). "
            f"§J §6.1: trial burns required before incinerating particularly dangerous waste."
        ),
        "comments": bi(
            "Seuils stricts pour incinérateurs de déchets; pertinence directe pour élimination thermique des déchets plastiques.",
            "Strict thresholds for waste incinerators; direct relevance for thermal disposal of plastic waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Annexe II §J §5.1: {q('Il est interdit d incinérer des déchets urbains et des déchets spéciaux dans des installations d une puissance calorifique inférieure à 350 kW', 'It is prohibited to incinerate urban waste and special waste in installations with thermal power below 350 kW')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"§3.1: operators must declare {q('la nature et la quantité des émissions', 'the nature and quantity of emissions')} to the competent authority. "
            f"Ch. V §4: quarterly transmission of measurement results."
        ),
        "comments": bi(
            "Interdit les petits brûleurs sauvages pour déchets urbains/spéciaux (incl. plastiques). Exception pour déchets hospitaliers spécifiques (§5.2).",
            "Prohibits small uncontrolled burners for urban/special waste (incl. plastics). Exception for specific hospital waste (§5.2).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Annexe II §K: {q('Les dispositions s appliquent aux installations destinées au traitement des surfaces d objets et de produits en métal, verre, céramique, matières plastiques, caoutchouc, ou autres matières', 'The provisions apply to installations for surface treatment of objects and products in metal, glass, ceramic, plastic materials, rubber or other materials')} "
            f"by halogenated hydrocarbons with boiling point <1013 mbar. "
            f"§K §2b: {q('le débit massique des émissions d hydrocarbures halogénés... ne doit pas dépasser 100 g/h', 'mass flow of halogenated hydrocarbon emissions must not exceed 100 g/h')} (Annexe I) "
            f"and {q('25 g/h', '25 g/h')} (Annexe III)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"§K §2a: {q('les objets et les produits doivent être traités dans une enceinte fermée', 'objects and products must be treated in a closed enclosure')}. "
            f"§K §2c: {q('les émissions seront réduites au moyen d un système de récupération des vapeurs', 'emissions must be reduced by means of a vapour recovery system')} or equivalent measure."
        ),
        "comments": bi(
            "Mention explicite des matières plastiques — traitement de surface industriel (dégraissage, revêtement).",
            "Explicit mention of plastic materials — industrial surface treatment (degreasing, coating).",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Ch. III & Annexe I (tableau immissions): ambient air quality limits including "
            f"{q('Anhydride sulfureux (SO2) : 50 mg/m3 (moyenne annuelle); 125 ug/m3 (moyenne journalière)', 'Sulphur dioxide (SO2): 50 mg/m3 (annual average); 125 ug/m3 (daily average)')}, "
            f"{q('Dioxyde d azote (NO2) : 200 mg/m3 (moyenne horaire); 40 ug/m3 (moyenne annuelle)', 'Nitrogen dioxide (NO2): 200 mg/m3 (hourly average); 40 ug/m3 (annual average)')}, "
            f"{q('Poussières en suspension (PM10) : 80 ug/m3 (moyenne annuelle)', 'Suspended dust (PM10): 80 ug/m3 (annual average)')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Ch. III §1.1: {q('L autorité compétente surveille l état et l évolution de la pollution de l air sur le territoire national', 'The competent authority monitors the state and evolution of air pollution on the national territory')}. "
            f"§2.1: immission forecasts may be required before construction of high-emission installations."
        ),
        "comments": bi(
            "Normes de qualité de l air ambiant; protection indirecte contre les émissions industrielles (particules, NOx, SO2).",
            "Ambient air quality standards; indirect protection against industrial emissions (particles, NOx, SO2).",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"§9.1.1: {q('Les dispositions de la norme sénégalaise NS 05-060 sont applicables', 'The provisions of Senegalese Standard NS 05-060 apply')} for CO, volatile hydrocarbons (HC) and smoke opacity from motor vehicles."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"§5.1: {q('Pour les véhicules les dispositions relatives aux méthodes de mesure contenues dans la norme sénégalaise NS 05-060 sont applicables', 'For vehicles, the measurement method provisions in Senegalese Standard NS 05-060 apply')}. "
            f"§9.1.2: new automotive industries must also comply with international construction standards."
        ),
        "comments": bi(
            "Référence croisée à NS 05-060 pour émissions véhicules; pertinence indirecte pour transport de déchets plastiques.",
            "Cross-reference to NS 05-060 for vehicle emissions; indirect relevance for transport of plastic waste.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Ch. V: {q('L exploitant doit mettre en place un programme de surveillance de ses rejets', 'The operator must establish a monitoring programme for its discharges')}. "
            f"§4: {q('Les résultats des mesures sont transmis au moins trimestriellement à l autorité administrative compétente', 'Measurement results are transmitted to the competent administrative authority at least quarterly')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Ch. V §3: when pollutant quantities exceed limits, {q('l arrêté d autorisation doit fixer la liste des paramètres à mesurer et la fréquence des mesures', 'the authorisation order must set the list of parameters to measure and measurement frequency')}; "
            f"at least once per year by an approved body. "
            f"§2: continuous monitoring of representative parameters may replace specific pollutant measurements."
        ),
        "comments": bi(
            "Obligation d auto-surveillance et reporting trimestriel pour toutes installations soumises à la norme.",
            "Self-monitoring and quarterly reporting obligation for all installations subject to the standard.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"§4.1–4.2: {q('L autorité compétente... procède elle-même à des mesures ou à des contrôles des émissions ou les fait exécuter par des services ou organismes agréés', 'The competent authority... itself carries out emission measurements or controls, or has them carried out by approved services or bodies')}. "
            f"§4.2: {q('ces émissions... soient mesurées et enregistrées en permanence', 'these emissions... must be measured and recorded continuously')} for installations with significant emissions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"§5.3: {q('Les valeurs mesurées et les valeurs calculées... sont consignées dans un rapport tenu par le détenteur de l installation, visé par les services agréés', 'Measured and calculated values... are recorded in a report kept by the installation holder, endorsed by approved services')}. "
            f"Ch. IV §4: mandatory sampling and measurement points on discharge pipelines."
        ),
        "comments": bi(
            "Contrôle autorité + mesure continue pour grandes installations industrielles (raffineries, incinérateurs, usines chimiques/plastiques).",
            "Authority control + continuous measurement for large industrial installations (refineries, incinerators, chemical/plastics plants).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"§6.4: For continuous emission monitoring, limits are met if during a calendar year "
            f"{q('aucune moyenne journalière n est supérieure à la valeur limite', 'no daily average exceeds the limit value')} and "
            f"{q('aucune moyenne horaire ne dépasse le double de la valeur limite', 'no hourly average exceeds twice the limit value')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"§6.2: {q('les valeurs calculées... sont converties en moyenne horaire', 'calculated values... are converted to hourly averages')}. "
            f"§6.5: when the same pollutant is discharged through multiple outlets, {q('le flux total de l ensemble des rejets est rapporté aux valeurs limites', 'the total flow of all discharges is compared to the limit values')}."
        ),
        "comments": bi(
            "Critères de conformité quantitatifs et opérationnels pour l appréciation des émissions continues.",
            "Quantitative and operational compliance criteria for assessing continuous emissions.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Annexe IV: {q('La hauteur de cheminée... ne peut être inférieure à 10 m', 'Chimney height... may not be less than 10 m')}. "
            f"Height calculated as S = kq/c_m where k=340 (gases) or 680 (dust), based on pollutant emission rate and admissible ground-level concentration."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Annexe IV: formula-based chimney height calculation using reference concentrations for "
            f"SO2, NOx, dust, HCl, organic compounds, Pb, Cd. "
            f"Ch. IV §3: {q('Les rejets dans l atmosphère sont... évacués... par l intermédiaire de cheminées', 'Discharges into the atmosphere must... be evacuated... via chimneys')}."
        ),
        "comments": bi(
            "Exigences techniques de dispersion atmosphérique pour installations industrielles; complément aux valeurs limites d émission.",
            "Technical atmospheric dispersion requirements for industrial installations; complement to emission limit values.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Annexe III: Carcinogenic substance emission limits including "
            f"{q('Benzidine, benzo(a)pyrène... 0,1 mg/m3', 'Benzidine, benzo(a)pyrene... 0.1 mg/m3')}, "
            f"{q('Acrylonitrile ; épichlorhydrine ; chlorure de vinyle... 1 mg/m3', 'Acrylonitrile; epichlorohydrin; vinyl chloride... 1 mg/m3')}, "
            f"{q('Benzène ; 1-3 butadiène... 5 mg/m3', 'Benzene; 1-3 butadiene... 5 mg/m3')}. "
            f"§11.1: authorisation order sets limits for other carcinogens per national/international recommendations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"§11.1: {q('l arrêté d autorisation fixe la limitation maximale en considération des recommandations de l autorité compétente', 'the authorisation order sets the maximum limit considering recommendations of the competent authority')}. "
            f"Annexe II §J §6.2: particularly dangerous emissions include {q('hydrocarbures aromatiques polyhalogénés', 'polyhalogenated aromatic hydrocarbons')}."
        ),
        "comments": bi(
            "Limites pour substances cancérigènes (VCM, acrylonitrile, benzène) — pertinents pour industrie chimique/plastique.",
            "Limits for carcinogenic substances (VCM, acrylonitrile, benzene) — relevant to chemical/plastics industry.",
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
