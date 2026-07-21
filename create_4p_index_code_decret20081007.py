#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Décret 2008-1007 (bilingual FR/EN cells)."""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Decret_2008-1007_Dechets_Biomedicaux.xlsx"

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
        "Décret n° 2008-1007 du 18 août 2008 portant réglementation de la gestion des déchets biomédicaux",
        "Decree No. 2008-1007 of 18 August 2008 regulating biomedical waste management",
    ),
    "policy_url": (
        "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/"
        "19089/decret-portant-reglementation-de-la-gestion-des-dechets-biomedicaux.pdf"
    ),
    "policy_year": 2008,
    "policy_objective": bi(
        "Établir les règles de classification, tri à la source, conditionnement, stockage, transport, "
        "traitement et élimination des déchets biomédicaux (DASRI, anatomiques, piquants-tranchants, "
        "pharmaceutiques, recyclables, spéciaux) en imposant l'agrément des opérateurs et la traçabilité — "
        "incluant explicitement les déchets plastiques de santé (flacons, emballages) sans mesures "
        "spécifiques aux plastiques non biomédicaux.",
        "Establish classification, source segregation, packaging, storage, transport, treatment and disposal "
        "rules for biomedical waste (infectious, anatomical, sharps, pharmaceutical, recyclable, special) "
        "by imposing operator approval and traceability — explicitly including healthcare plastic waste "
        "(bottles, packaging) without measures specific to non-biomedical plastics.",
    ),
    "policy_target": 1,
    "policy_target_text": (
        f"{q('La durée de stockage ne doit pas dépasser 48 heures', 'Storage duration must not exceed 48 hours')}; "
        f"{q('régulièrement décontaminé (une fois par semaine au moins)', 'regularly decontaminated (at least once per week)')}; "
        f"{q('La température minimale d incinération requise est de 800° C', 'The minimum required incineration temperature is 800°C')}; "
        f"{q('suspension temporaire de 1 à 12 mois', 'temporary suspension of 1 to 12 months')}; "
        f"{q('prise en charge précoce en cas d AES dans les 48 heures', 'early care in case of blood exposure within 48 hours')}."
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Décret exécutif adopté par le Président de la République le 18 août 2008, sur rapport du "
        "Ministre de la Santé et de la Prévention, en application du Code de l'hygiène, du Code de "
        "l'environnement et des lois hospitalières. Il s'agit d'un texte réglementaire d'application "
        "exécutive, et non d'une loi adoptée par le Parlement (≠1).",
        "Executive decree adopted by the President of the Republic on 18 August 2008, on the report of "
        "the Minister of Health and Prevention, implementing the Hygiene Code, Environmental Code and "
        "hospital laws. It is sub-legislative implementing regulation, not an act enacted by Parliament (≠1).",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "santé, environnement, collectivités locales, industrie, recherche, gestion des déchets",
        "health, environment, local government, industry, research, waste management",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "utilisation, collecte, élimination",
        "use, collection, disposal",
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
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 4: Waste classification including "
            f"{q('Déchets recyclables : Il s agit des déchets plastiques tels que les flacons de sérum, d eau de javel, le matériel en verre, les contenants sous pression', 'Recyclable waste: plastic waste such as serum bottles, bleach bottles, glass equipment, pressurised containers')}; "
            f"plus infectious, anatomical, sharps, pharmaceutical and special waste categories."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 6: {q('tout déchet... est trié au niveau du lieu même de production', 'all waste... is sorted at the place of production')} into category-specific circuits with pictograms displayed. "
            f"Art. 7: colour-coded containers fixed by Health Ministry order."
        ),
        "comments": bi(
            "Mention explicite des déchets plastiques recyclables (flacons, contenants) — pertinence directe pour déchets plastiques de santé.",
            "Explicit mention of recyclable plastic waste (bottles, containers) — direct relevance for healthcare plastics.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 3 (définitions): {q('Recyclage : Réintroduction directe d un matériau dans son propre cycle de production... (papier, plastique)', 'Recycling: direct reintroduction of a material into its production cycle... (paper, plastic)')}; "
            f"{q('Valorisation des déchets : Elle recouvre le réemploi, la réutilisation, le recyclage ou la régénération', 'Waste valorisation: covers reuse, reutilisation, recycling or regeneration')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 11: {q('les déchets assimilés aux ordures ménagères, de même que les déchets recyclables, suivent la filière des déchets ménagers', 'household-type and recyclable waste follow the municipal waste stream')}. "
            f"Art. 5: producers must eliminate/recycle or use Health Ministry-approved companies."
        ),
        "comments": bi(
            "Cadre de valorisation/recyclage incluant le plastique; recyclables intégrés à la filière déchets ménagers.",
            "Valorisation/recycling framework including plastic; recyclables integrated into municipal waste stream.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 5: {q('Toute personne physique ou morale, qui produit ou détient des déchets biomédicaux, en assure elle-même l élimination ou le recyclage ou les fait éliminer ou recycler auprès des entreprises agréées par le Ministre chargé de la santé', 'Any person producing or holding biomedical waste must itself ensure elimination or recycling, or have it eliminated or recycled by companies approved by the Minister of Health')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 14: {q('Tout opérateur de déchets biomédicaux doit obtenir l agrément du Ministère chargé de la santé', 'Every biomedical waste operator must obtain approval from the Ministry of Health')}. "
            f"Art. 20: {q('tenus d enregistrer leurs déchets et d en assurer la traçabilité', 'must record their waste and ensure traceability')}."
        ),
        "comments": bi(
            "Responsabilité élargie du producteur pour tous déchets biomédicaux (incl. plastiques de santé).",
            "Extended producer responsibility for all biomedical waste (incl. healthcare plastics).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 6: {q('Tout déchet issu des activités médicales, pharmaceutiques, vétérinaires ou de recherche, est trié au niveau du lieu même de production et mis dans le circuit spécifique dédié à cette catégorie', 'All waste from medical, pharmaceutical, veterinary or research activities is sorted at the place of production and placed in the specific circuit dedicated to that category')}. "
            f"{q('Des pictogrammes d indication des catégories de déchets par type de contenant sont affichés', 'Pictograms indicating waste categories by container type are displayed')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3 (Tri à la source): {q('Séparation effectuée au niveau même du lieu où le déchet est produit et au moment de la production', 'Separation at the very place and moment of waste production')}. "
            f"Art. 7: packaging colours and labelling fixed by ministerial order."
        ),
        "comments": bi(
            "Tri à la source obligatoire; sépare déchets recyclables plastiques des DASRI infectieux.",
            "Mandatory source segregation; separates recyclable plastics from infectious healthcare waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 7: {q('Le conditionnement est effectué dès la production, pour éviter tout risque sanitaire et environnemental', 'Packaging is carried out at production to avoid any health and environmental risk')}. "
            f"{q('Les contenants à usages multiples pour déchets à risques infectieux sont obligatoirement nettoyés et désinfectés après chaque usage avant réutilisation', 'Reusable containers for infectious waste must be cleaned and disinfected after each use before reuse')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 7: {q('La couleur, la nature des différents types d emballages et d étiquetages en fonction des déchets sont fixées par arrêté du Ministère chargé de la Santé', 'Colour, nature of packaging and labelling by waste type are set by Health Ministry order')}. "
            f"Art. 16: producers/operators must have appropriate packaging equipment."
        ),
        "comments": bi(
            "Conditionnement immédiat; pertinent pour emballages plastiques de déchets infectieux et recyclables.",
            "Immediate packaging; relevant for plastic packaging of infectious and recyclable waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 8–9: {q('Sont soumis à un prétraitement sur place... les cultures de laboratoires', 'Laboratory cultures subject to on-site pre-treatment')}. "
            f"Storage in {q('local aéré et sécurisé... facilement décontaminable, régulièrement décontaminé (une fois par semaine au moins)', 'ventilated and secure room... easily decontaminable, regularly decontaminated (at least once per week)')}. "
            f"{q('La durée de stockage ne doit pas dépasser 48 heures', 'Storage duration must not exceed 48 hours')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 9: storage room must be of sufficient capacity and accessible for collection. "
            f"Art. 16: appropriate storage equipment required for producers/operators."
        ),
        "comments": bi(
            "Délais quantitatifs (48 h, hebdomadaire) pour stockage; limite les risques de fuite de déchets plastiques/médicaux.",
            "Quantitative deadlines (48 h, weekly) for storage; limits leakage risk of plastic/medical waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 10: {q('Le transport des déchets biomédicaux... se fait dans des conditions telles que la protection de l environnement et des personnes soit assurée', 'Transport of biomedical waste... must ensure protection of the environment and persons')}. "
            f"{q('Un contenant hermétique et inviolable est utilisé pour le transport externe', 'A hermetic and tamper-proof container is used for external transport')}; "
            f"{q('contenant de type GRV (Grand Récipient pour Vrac)', 'GRV (Large Bulk Container) type container')} or specially equipped vehicle conforming to dangerous goods transport."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 10: transport during {q('périodes de faible circulation', 'low-traffic periods')}. "
            f"Art. 16: appropriate transport equipment required. "
            f"Art. 3: transport defined as transfer from production to storage pending treatment."
        ),
        "comments": bi(
            "Transport sécurisé des déchets biomédicaux (incl. contenants plastiques) vers centres de traitement.",
            "Secure transport of biomedical waste (incl. plastic containers) to treatment centres.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 11: Treatment pathways — "
            f"{q('les déchets infectieux... sont incinérés, enfouis ou bien subissent un traitement de type stérilisation/broyage', 'infectious waste is incinerated, landfilled or undergoes sterilisation/shredding treatment')}; "
            f"{q('les déchets piquants et tranchants sont incinérés, enfouis ou bien subissent un traitement de type stérilisation/broyage', 'sharps are incinerated, landfilled or undergo sterilisation/shredding')}; "
            f"{q('les déchets recyclables... suivent la filière des déchets ménagers', 'recyclable waste... follows the municipal waste stream')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 11: pharmaceutical waste per Health Ministry procedure; special/chemical waste per specific treatment; "
            f"radioactive waste per ionising radiation law. "
            f"Art. 12: incinerator must comply with ICPE regulations."
        ),
        "comments": bi(
            "Voies de traitement différenciées; broyage/stérilisation pertinent pour déchets plastiques médicaux infectieux.",
            "Differentiated treatment pathways; shredding/sterilisation relevant for infectious medical plastic waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 12: {q('L installation et le fonctionnement d un incinérateur doivent être conformes à la réglementation en vigueur, notamment aux prescriptions édictées dans le dossier des installations classées', 'Installation and operation of an incinerator must comply with applicable regulations, notably ICPE requirements')}. "
            f"{q('La température minimale d incinération requise est de 800° C', 'The minimum required incineration temperature is 800°C')}. "
            f"{q('Les résidus de l incinération doivent faire l objet d un enfouissement hygiénique', 'Incineration residues must undergo sanitary landfilling')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 13: {q('Quel que soit le type de traitement choisi, l activité doit faire l objet d une étude d impact environnemental et social', 'Whatever treatment type is chosen, the activity must be subject to an environmental and social impact study')}. "
            f"Rapport de présentation: inadequate incineration emits dioxins/furans."
        ),
        "comments": bi(
            "Seuil 800°C pour incinération; lien avec NS 05-062 et filière plastique médicale non recyclable.",
            "800°C incineration threshold; link to NS 05-062 and non-recyclable medical plastic stream.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 13: {q('Quel que soit le type de traitement choisi, l activité doit faire l objet d une étude d impact environnemental et social', 'Whatever treatment type is chosen, the activity must be subject to an environmental and social impact study')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Links to EIA arrêtés 9468–9472 and Décret 2001-282 ICPE framework. "
            f"Art. 12: incinerator installation subject to ICPE authorisation dossier."
        ),
        "comments": bi(
            "EIE obligatoire pour installations de traitement; pertinent pour usines de traitement de déchets plastiques médicaux.",
            "Mandatory EIA for treatment facilities; relevant for medical plastic waste treatment plants.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Arts. 14–15: {q('Tout opérateur de déchets biomédicaux doit obtenir l agrément du Ministère chargé de la santé', 'Every biomedical waste operator must obtain approval from the Ministry of Health')}; "
            f"{q('La délivrance de l agrément est assujettie à l avis consultatif d une commission', 'Approval is subject to consultative opinion of a commission')} whose composition is set by ministerial order."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 21: {q('L agrément est retiré lorsque... l opérateur fonctionne dans des conditions dangereuses pour la santé publique', 'Approval is withdrawn when... the operator operates in conditions dangerous to public health')}; "
            f"{q('suspension temporaire de 1 à 12 mois', 'temporary suspension of 1 to 12 months')} in emergencies."
        ),
        "comments": bi(
            "Agrément d'État pour toute la filière (collecte, transport, traitement) des déchets biomédicaux.",
            "State approval for the entire biomedical waste chain (collection, transport, treatment).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Arts. 16–19: Producer/operator obligations — "
            f"{q('Équipements de protection pour le personnel', 'protective equipment for staff')}, "
            f"{q('Le respect des règles d hygiène tout le long de la filière est obligatoire', 'compliance with hygiene rules throughout the chain is mandatory')}, "
            f"{q('Les personnels... sont vaccinés contre l hépatite B, le tétanos', 'staff... vaccinated against hepatitis B, tetanus')}, "
            f"{q('prise en charge précoce en cas d AES dans les 48 heures', 'early care for blood exposure within 48 hours')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 16: required equipment for packaging, transport, treatment and disposal. "
            f"Art. 19: regular awareness for all staff and users; prevention measures for blood exposure accidents."
        ),
        "comments": bi(
            "Mesures de protection du personnel manipulant déchets plastiques/médicaux infectieux.",
            "Protection measures for staff handling infectious plastic/medical waste.",
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 20: {q('Les producteurs et les opérateurs sont tenus d enregistrer leurs déchets et d en assurer la traçabilité', 'Producers and operators must record their waste and ensure traceability')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 3 (Gestion des déchets): includes {q('tri, le conditionnement, la collecte, le transport, le stockage, le recyclage, le traitement et l élimination des déchets, y compris la surveillance des sites d élimination', 'sorting, packaging, collection, transport, storage, recycling, treatment and disposal, including monitoring of disposal sites')}."
        ),
        "comments": bi(
            "Traçabilité de bout en bout pour déchets biomédicaux incluant flacons et emballages plastiques.",
            "End-to-end traceability for biomedical waste including bottles and plastic packaging.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            f"Rapport de présentation & Art. 1: protection against dangers from biomedical waste and management modalities; "
            f"{q('le dépôt des déchets d activités de soins dans des zones non contrôlées... peut... contaminer... les eaux de surface', 'depositing healthcare waste in uncontrolled areas... can contaminate surface waters')}; "
            f"{q('Une incinération dans de mauvaises conditions techniques peut polluer l atmosphère... (dioxines, furanes)', 'Incineration under poor technical conditions can pollute the atmosphere... (dioxins, furanes)')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 21: administrative sanctions — agrément withdrawal or 1–12 month suspension; "
            f"notification by registered letter. "
            f"Preamble cites Basel, Rotterdam, Stockholm Conventions."
        ),
        "comments": bi(
            "Prévention des fuites environnementales (eau, air, sol) liées aux déchets de soins incluant plastiques.",
            "Prevention of environmental leakage (water, air, soil) from healthcare waste including plastics.",
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
