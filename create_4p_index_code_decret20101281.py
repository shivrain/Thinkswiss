#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2010-1281 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2010-1281_Batteries_Usees.xlsx"

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
        "Décret n° 2010-1281 du 16 septembre 2010 réglementant l'exploitation du plomb issu des batteries usagées",
        "Decree No. 2010-1281 of 16 September 2010 regulating recovery of lead from used batteries",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/19090/"
        "decret-num-3-2010-1281-du-16-septembre-2010-reglementant-les-conditions-dexploitation-du-plomb-issu-des-batteries-usagees.pdf"
    ),
    "policy_year": 2010,
    "policy_objective": bi(
        "Réglementer l'importation, la collecte, le transport, le recyclage, le stockage, le traitement et "
        "l'élimination du plomb issu des batteries usagées et l'utilisation du mercure, en imposant "
        "l'autorisation du Ministre chargé de l'Environnement, les exigences ICPE/EIE et la conformité à "
        "NS 05-062 — traiter les composants dangereux (plomb, mercure) souvent combinés à des boîtiers "
        "plastiques dans les équipements électriques usagés, sans mesures spécifiques aux plastiques.",
        "Regulate import, collection, transport, recycling, storage, treatment and disposal of lead from "
        "used batteries and mercury use, by imposing Environment Minister authorisation, ICPE/EIA requirements "
        "and NS 05-062 compliance — addressing hazardous components (lead, mercury) often combined with plastic "
        "casings in waste electrical equipment, without plastic-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Authorization conditions include ICPE operating permit with prior {q('évaluation environnementale', 'environmental assessment')}; "
        f"compliance with {q('chapitre 2 de la norme NS 05-062 sur la pollution atmosphère', 'Chapter 2 of NS 05-062 on atmospheric pollution')}; "
        f"waste management per {q('article L 30 du Code de l Environnement', 'Article L 30 of the Environmental Code')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret exécutif adopté par le Président de la République le 16 septembre 2010, sur proposition "
        "du Ministre d'État, Ministre de l'Environnement et de la Protection de la Nature, en application "
        "du Code de l'environnement et du Décret 2001-282. Il s'agit d'un texte réglementaire d'application "
        "exécutive, et non d'une loi parlementaire (≠1).",
        "Executive decree adopted by the President of the Republic on 16 September 2010, on proposal of the "
        "Minister of State, Minister of Environment and Nature Protection, implementing the Environmental "
        "Code and Decree 2001-282. It is sub-legislative implementing regulation, not parliamentary "
        "legislation (≠1).",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "industrie, gestion des déchets, environnement, mines, santé",
        "industry, waste management, environment, mining, health",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "utilisation, collecte, recyclage, élimination",
        "use, collection, recycling, disposal",
    ),
    "policy_budget": 0,
    "policy_budget_text": bi(
        "Aucune disposition budgétaire, redevance ou instrument économique explicite dans le décret.",
        "No explicit budget, fee or economic instrument in the decree.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 1: {q('Il est interdit à toute personne physique ou morale, d importer, de collecter, de transporter, de recycler, de stocker, de manipuler, de traiter ou d éliminer le plomb issu des batteries usagées et d autres sources, ainsi que le mercure et ses composés, sans l autorisation du Ministre chargé de l Environnement', 'It is prohibited for any natural or legal person to import, collect, transport, recycle, store, handle, treat or eliminate lead from used batteries and other sources, as well as mercury and its compounds, without authorisation from the Minister of the Environment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 1: {q('Les conditions de délivrance de cette autorisation sont fixées par arrêté du Ministre chargé de l Environnement', 'Authorisation delivery conditions are set by order of the Minister of the Environment')}. "
            f"Art. 5: violations punished per Environmental Code sanctions."
        ),
        "comments": bi(
            "Interdiction centrale sans autorisation; batteries usagées incluent boîtiers plastiques (composant non réglementé explicitement).",
            "Core prohibition without authorisation; used batteries include plastic casings (component not explicitly regulated).",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 2 (détenteurs): {q('des détenteurs définis comme des personnes physiques ou morales qui accumulent dans leurs propres établissements, des batteries usagées et d autres sources de plomb, ainsi que le mercure et ses composés, en raison de leurs activités professionnelles', 'holders defined as natural or legal persons who accumulate in their own establishments used batteries and other lead sources, as well as mercury and its compounds, due to their professional activities')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: holders must meet authorization conditions (hygiene/safety installations, ICPE permit, EIA, PPE, NS 05-062, waste management L 30). "
            f"Rapport: targets uncontrolled informal battery recycling sector."
        ),
        "comments": bi(
            "Responsabilité des détenteurs professionnels (garages, industries, DEEE); batteries à boîtier plastique.",
            "Professional holder responsibility (garages, industries, WEEE); plastic-cased batteries.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 2 (collecteurs): {q('des collecteurs définis comme des personnes physiques ou morales qui assurent la collecte de batteries usagées et d autres sources de plomb... et qui en assurent le transport en l état sans aucune forme de traitement jusqu au point d élimination', 'collectors defined as natural or legal persons who collect used batteries and other lead sources... and transport them as-is without any form of treatment to the elimination point')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: collectors require Environment Minister authorisation. "
            f"Art. 4: import/export subject to Basel and Rotterdam Convention provisions."
        ),
        "comments": bi(
            "Filière de collecte/transport sans traitement intermédiaire; pertinent pour batteries automobiles/plastique.",
            "Collection/transport chain without intermediate treatment; relevant for automotive/plastic batteries.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            f"Art. 2 (entreprises): {q('Des entreprises spécialisées dans la récupération et le recyclage du plomb issu des accumulateurs usagées', 'Companies specialised in recovery and recycling of lead from used accumulators')} eligible for authorisation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Rapport: {q('Seules les personnes physiques ou morales dûment autorisées... seront habilitées à exercer ces activités suivant des pratiques et des technologies appropriées', 'Only duly authorised persons... will be permitted to carry out these activities using appropriate practices and technologies')}. "
            f"Art. 3: process mastery from arrival through treatment to finished product output."
        ),
        "comments": bi(
            "Instrument de recyclage du plomb; co-produits/boîtiers plastiques non traités explicitement dans le décret.",
            "Lead recycling instrument; plastic casings/co-products not explicitly addressed in decree.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 3: {q('disposer d une autorisation d exploiter une installation classée, ayant fait au préalable l objet d une évaluation environnementale', 'hold an authorisation to operate a classified installation, having previously undergone an environmental assessment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3: compliance with hazardous chemical and waste management legislation. "
            f"References Décret 2001-282 ICPE framework and EIA arrêtés 9468-9472."
        ),
        "comments": bi(
            "ICPE + EIE obligatoires pour usines de recyclage batteries; pertinent pour installations plastique/batterie.",
            "ICPE + EIA mandatory for battery recycling plants; relevant for plastic/battery facilities.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            f"Art. 3: {q('disposer d installations conformes aux normes d hygiène et de sécurité en vigueur, compte tenu de la spécificité des produits à manipuler', 'have installations compliant with applicable hygiene and safety standards, given the specific products handled')}; "
            f"{q('assurer la surveillance médicale de son personnel... et le doter d équipements de protection individuelle répondant aux normes', 'ensure medical surveillance of personnel... and provide PPE meeting standards')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Rapport: addresses health risks to informal sector workers and neighbouring populations from uncontrolled battery recycling. "
            f"Art. 3: heavy metal exposure monitoring required."
        ),
        "comments": bi(
            "Protection sanitaire des travailleurs du recyclage; secteur informel ciblé par le rapport de présentation.",
            "Health protection for recycling workers; informal sector targeted by presentation report.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Art. 3: {q('respecter les dispositions du chapitre 2 de la norme NS 05-062 sur la pollution atmosphère', 'comply with the provisions of Chapter 2 of NS 05-062 on atmospheric pollution')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Preamble cites NS 05-062. NS 05-062 Ch. II sets emission limits for stationary installations (dust, SO2, NOx, etc.) applicable to battery smelting/recycling. "
            f"Art. 3: process mastery required throughout treatment."
        ),
        "comments": bi(
            "Renvoi croisé à NS 05-062 pour émissions atmosphériques du recyclage plomb; fumées de fusion batteries.",
            "Cross-reference to NS 05-062 for lead recycling air emissions; battery smelting fumes.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 3: {q('gérer les déchets conformément aux dispositions de l article L 30 du Code de l Environnement', 'manage waste in accordance with Article L 30 of the Environmental Code')} (ecologically rational hazardous waste management)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Rapport: regulates {q('le stockage, le traitement et l élimination du plomb issu des batteries usagées', 'storage, treatment and disposal of lead from used batteries')}. "
            f"Code L 30: producers must eliminate/recycle or use licensed enterprises."
        ),
        "comments": bi(
            "Gestion des déchets résiduels (incl. fractions plastiques des batteries) via Code environnement L 30.",
            "Residual waste management (incl. plastic battery fractions) via Environmental Code L 30.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Rapport & Art. 1: regulates {q('L utilisation du mercure ou des équipements en contenant par le secteur formel et informel', 'use of mercury or equipment containing it by the formal and informal sector')}, notably artisanal gold mining (orpaillage)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 1: mercury and compounds subject to same authorization prohibition as battery lead. "
            f"Rapport: supports Minamata Convention adherence (international mercury initiative)."
        ),
        "comments": bi(
            "Mercure en orpaillage; composé dangereux distinct du plomb mais dans le même décret.",
            "Mercury in artisanal gold mining; hazardous compound distinct from lead but in same decree.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 4: {q('En cas d importation ou d exportation, les dispositions pertinentes des Conventions de Bâle et de Rotterdam... devront être respectées', 'In case of import or export, relevant provisions of the Basel and Rotterdam Conventions... must be complied with')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Preamble cites Basel (1992), Bamako (1996), Rotterdam (2004). "
            f"Rapport: strengthens Senegal legal implementation of transboundary hazardous waste control."
        ),
        "comments": bi(
            "Contrôle des mouvements transfrontières de déchets batteries/plomb; Bamako Convention (Africa).",
            "Transboundary battery/lead waste movement control; Bamako Convention (Africa).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            f"Art. 3: {q('justifier d une maîtrise des processus et procédés lié à l exploitation du plomb ou du mercure, depuis l arrivée au niveau de l installation, pendant le traitement et la sortie des produits finis', 'demonstrate mastery of processes and procedures related to lead or mercury operation, from arrival at the installation, during treatment and to finished product output')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 1: authorization conditions detailed in ministerial arrêté (not coded separately). "
            f"Art. 3: compliance with chemical and hazardous waste legislation required for authorization."
        ),
        "comments": bi(
            "Exigence technique de maîtrise des procédés; condition d agrément pour recycleurs de batteries.",
            "Technical process mastery requirement; authorisation condition for battery recyclers.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 5: {q('Toute infraction aux dispositions du présent décret sera punie conformément aux sanctions prévues par le Code de l Environnement', 'Any violation of this decree shall be punished in accordance with sanctions provided by the Environmental Code')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Loi 2001-01 / Décret 2001-282: criminal and administrative sanctions for environmental violations. "
            f"Rapport: targets uncontrolled informal recycling exposing workers and communities."
        ),
        "comments": bi(
            "Sanctions via Code environnement; renforce interdiction du recyclage informel de batteries.",
            "Sanctions via Environmental Code; reinforces prohibition on informal battery recycling.",
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
