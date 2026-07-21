#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2015-18 Maritime Fisheries Code (bilingual FR/EN).

Source text extracted via OCR from official ditp.gouv.sn PDF (image-based, 36 pages).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2015-18_Code_Peche_Maritime.xlsx"

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
        "Loi n° 2015-18 du 13 juillet 2015 portant Code de la Pêche maritime",
        "Law No. 2015-18 of 13 July 2015 establishing the Maritime Fisheries Code",
    ),
    "policy_url": "https://www.ditp.gouv.sn/download/file/fid/53",
    "policy_year": 2015,
    "policy_objective": bi(
        "Refondre le cadre de gestion et d'aménagement des pêches maritimes, incluant la conservation "
        "des écosystèmes marins, la protection de l'environnement marin, le contrôle des engins de pêche "
        "(notamment filets en nylon monofilament/multimonofilament) et la lutte contre la pêche INN — "
        "pertinent pour les engins de pêche en plastique et les déchets marins, sans mesures spécifiques "
        "aux déchets plastiques marins.",
        "Refound the maritime fisheries management framework, including marine ecosystem conservation, "
        "marine environment protection, fishing-gear control (notably nylon monofilament/multimonofilament "
        "nets) and IUU fishing control — relevant to plastic fishing gear and marine litter, without "
        "plastic marine litter-specific measures.",
    ),
    "policy_target": 0.75,
    "policy_target_text": (
        f"Art. 66: ban on {q('nappes et filets maillants fabriqués à partir d éléments monofilaments ou multimonofilaments en nylon', 'gillnets made from nylon monofilament or multimonofilament elements')}; "
        f"Art. 33(j): regulatory measures for {q('la protection de l environnement marin', 'protection of the marine environment')}; "
        f"Art. 68: mandatory {q('immatriculation et marquage', 'registration and marking')} of artisanal fishing vessels and gear."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Code de la pêche maritime adopté par l'Assemblée nationale le 30 juin 2015 et promulgué le "
        "13 juillet 2015 en tant que Loi n° 2015-18. Il s'agit d'une loi parlementaire (Code), et non d'un "
        "texte réglementaire d'application exécutive (≠0.75).",
        "Maritime Fisheries Code adopted by the National Assembly on 30 June 2015 and promulgated on "
        "13 July 2015 as Law No. 2015-18. It is parliamentary legislation (Code), not executive "
        "implementing regulation (≠0.75).",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "pêche maritime, pêche artisanale, pêche industrielle, environnement marin, commerce maritime",
        "maritime fisheries, artisanal fishing, industrial fishing, marine environment, maritime trade",
    ),
    "policy_circularity": 0.50,
    "policy_lifecycle_phases_list": bi(
        "utilisation, élimination, environnement marin",
        "use, disposal, marine environment",
    ),
    "policy_budget": 0.25,
    "policy_budget_text": (
        f"Art. 64: {q('redevance', 'fee')} for permits and authorisations (amount set by joint ministerial order); "
        f"Arts. 123–127: criminal fines up to FCFA 30,000,000 for serious industrial fishing offences; "
        f"no dedicated budget allocation for marine litter or gear recovery."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 14: {q('L Etat adopte une approche de gestion intégrée fondée sur l écosystème, incorporant des objectifs de conservation en vue d assurer la viabilité des espèces et des habitats', 'The State adopts an ecosystem-based integrated management approach incorporating conservation objectives to ensure viability of species and habitats')} "
            f"in Senegalese maritime waters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 15–18: measures based on scientific/technical advice; Art. 16: Minister may create "
            f"marine protected areas, fish-aggregating devices and artificial reefs."
        ),
        "comments": bi(
            "Gestion écosystémique des pêches — cadre indirect pour la protection des milieux marins face aux déchets/engins.",
            "Ecosystem-based fisheries management — indirect framework for protecting marine environments from waste/gear.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 17: {q('Lorsqu il est nécessaire de prendre en considération des mesures de conservation intégrant des facteurs environnementaux ou anthropiques autres que la pêche, un arrêté interministériel est pris... afin de mieux assurer la protection des ressources et de la biodiversité marine', 'Where conservation measures must integrate environmental or anthropogenic factors other than fishing, an interministerial order is issued... to better ensure protection of resources and marine biodiversity')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Joint order by Fisheries Minister and other concerned ministers; enables cross-sector marine pollution measures."
        ),
        "comments": bi(
            "Ouverture à des mesures environnementales anthropiques (pollution, déchets) au-delà de la pêche.",
            "Opens scope for anthropogenic environmental measures (pollution, waste) beyond fishing per se.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 21: {q('Les récifs artificiels désignent les aménagements... par la mise en place de substrats durs d origines diverses, notamment, blocs rocheux, divers matériels industriels usagés et autres ensembles spécialement manufacturés. Ils sont immergés sur le sédiment', 'Artificial reefs are installations placing hard substrates of various origins, notably rock blocks, various used industrial materials and specially manufactured assemblies, submerged on the seabed')} "
            f"to increase biological productivity, favour biodiversity and protect natural habitats."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 16: creation modalities fixed by Fisheries Minister order; repurposing of industrial materials at sea."
        ),
        "comments": bi(
            "Réutilisation/immersion de matériaux industriels usagés — lien indirect avec gestion de déchets solides en mer.",
            "Reuse/submersion of used industrial materials — indirect link to solid-waste management at sea.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Marine environment",
        "instrument_description": (
            f"Art. 33(j): regulatory measures may include {q('la limitation du volume de capture de certaines espèces... favorisant la conservation des ressources et la protection de l environnement marin', 'limitation of catch volume of certain species... favouring resource conservation and protection of the marine environment')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 33 also covers mesh sizes, closed zones, gear types/methods (h), and artisanal gear signalling (k)."
        ),
        "comments": bi(
            "Protection explicite de l'environnement marin dans les mesures réglementaires d'application.",
            "Explicit marine environment protection in implementing regulatory measures.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 66: {q('Sont interdits l importation, la mise en vente, l achat, la détention et l utilisation des nappes et filets maillants fabriqués à partir d éléments monofilaments ou multimonofilaments en nylon sauf dérogation spéciale', 'Import, sale, purchase, possession and use of gillnets made from nylon monofilament or multimonofilament elements are prohibited except by special derogation')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Exposé des motifs: prohibition raised to legislative level to combat ghost gear and destructive fishing; "
            f"Arts. 125–126: use of prohibited gear is a very serious offence with confiscation and destruction."
        ),
        "comments": bi(
            "Interdiction directe des filets en nylon (polymère plastique) — principal lien explicite avec les plastiques marins/engins fantômes.",
            "Direct ban on nylon (plastic polymer) nets — main explicit link to marine plastics/ghost gear.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 65: prohibits use of {q('matières explosives ou de substances ou appâts toxiques', 'explosives or toxic substances or baits')} in fishing and possession on board; "
            f"prohibits certain underwater-breathing equipment for spearfishing."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 125(h): use/transport of explosives or toxic substances for fishing is a very serious industrial offence."
        ),
        "comments": bi(
            "Contrôle des méthodes de pêche destructrices; complète le régime des engins (Art. 66).",
            "Control of destructive fishing methods; complements gear regime (Art. 66).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Use",
        "instrument_description": (
            f"Art. 68: {q('Toutes les embarcations de pêche artisanale appartenant aux nationaux... sont immatriculées et marquées conformément aux règles fixées par arrêté', 'All artisanal fishing vessels belonging to nationals... are registered and marked under ministerial rules')}; "
            f"same for foreign residents' vessels."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 33(k): regulatory measures on {q('signalisation des engins de pêche artisanale', 'signalling of artisanal fishing gear')} and vessel safety norms."
        ),
        "comments": bi(
            "Traçabilité et marquage des engins — pertinent pour identifier et récupérer les engins abandonnés.",
            "Gear traceability and marking — relevant for identifying and recovering abandoned gear.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 83: {q('Le Ministre chargé de la pêche maritime est l autorité responsable de la supervision et de la coordination de l ensemble des activités et opérations de surveillance et de protection des pêcheries', 'The Minister of Maritime Fisheries is responsible for supervising and coordinating all fisheries surveillance and protection operations')} in Senegalese waters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Arts. 84–85: surveillance agents (fisheries, navy, gendarmerie, customs, etc.) may inspect vessels and gear at sea."
        ),
        "comments": bi(
            "Autorité de surveillance des pêcheries et des engins en mer.",
            "Authority for fisheries and at-sea gear surveillance.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 85(a): surveillance agents may {q('inspecter les engins de pêche utilisés à bord ou à partir du navire et, à cette fin, faire retirer de l eau les engins de pêche', 'inspect fishing gear used on board or from the vessel and, for this purpose, have fishing gear removed from the water')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 87: vessels may be ordered to stop; Art. 100: offending vessels must proceed to designated Senegalese port."
        ),
        "comments": bi(
            "Pouvoir de retirer les engins de l'eau — pertinent pour engins abandonnés ou illégaux.",
            "Power to remove gear from water — relevant for abandoned or illegal gear.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 125(a),(e): very serious offences include {q('l usage d engins ou de méthodes de pêche interdits', 'use of prohibited fishing gear or methods')} and "
            f"{q('l utilisation... de tous moyens ou dispositifs ayant pour effet de rendre l ouverture de la maille inférieure à l ouverture minimale autorisée ou de réduire l action sélective des engins de pêche', 'use of any means or devices reducing mesh opening below the minimum or reducing selective action of fishing gear')}; "
            f"{q('les engins... sont confisqués et détruits', 'gear... is confiscated and destroyed')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Penalty: FCFA 20,000,000–30,000,000 for industrial fishing; Art. 126 applies to artisanal fishing "
            f"(FCFA 150,000–300,000) with gear confiscation and destruction."
        ),
        "comments": bi(
            "Confiscation et destruction des engins illégaux — mécanisme d'élimination des engins de pêche en plastique prohibés.",
            "Confiscation and destruction of illegal gear — elimination mechanism for prohibited plastic fishing gear.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 106: {q('La conclusion de la transaction peut être subordonnée à l abandon des engins et produits saisis ou du montant de la vente des produits saisis au profit de l Etat', 'Settlement may be conditional on abandonment of seized gear and products or payment of proceeds of sale of seized products to the State')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 113: Minister or court decides final destination of seized/confiscated goods per applicable regulations."
        ),
        "comments": bi(
            "Destination des engins saisis — gestion des déchets/engins confisqués.",
            "Destination of seized gear — management of confiscated waste/gear.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Titre IV (Arts. 74–76): port State measures — foreign fishing vessels must announce port entry; "
            f"{q('Tout débarquement de capture par un navire étranger dans un port sénégalais doit faire l objet d une autorisation', 'Any landing of catch by a foreign vessel in a Senegalese port requires authorisation')} "
            f"after surveillance inspections; results of offences communicated to flag State — IUU fishing control."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Integrates international port State measures against IUU fishing; supports control of illegal gear and waste at port."
        ),
        "comments": bi(
            "Mesures de l'État du port — contrôle des navires étrangers et réduction de la pêche illégale (source d'engins abandonnés).",
            "Port State measures — control of foreign vessels and reduction of IUU fishing (source of abandoned gear).",
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
