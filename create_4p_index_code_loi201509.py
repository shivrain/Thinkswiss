#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2015-09 (bilingual FR/EN cells).

Source: uploaded official PDF / dri.gouv.sn (abrogated 2020, replaced by Loi 2020-04).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2015-09_Sachets_Plastiques.xlsx"

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
        "Loi n° 2015-09 du 4 mai 2015 relative à l'interdiction des sachets plastiques de faible micronnage (abrogée)",
        "Law No. 2015-09 of 4 May 2015 on prohibition of thin plastic bags (repealed)",
    ),
    "policy_url": (
        "https://www.dri.gouv.sn/sites/default/files/an-documents/"
        "LOI%20N%202015%2009%20DU%204%20MAI%202015.pdf"
    ),
    "policy_year": 2015,
    "policy_objective": bi(
        "Première loi nationale dédiée aux sachets plastiques : interdire la production, l'importation, la détention, "
        "la distribution et l'utilisation des sachets de faible micronnage (<30 microns), mettre fin à leur gratuité "
        "pour les sachets ≥30 microns, standardiser les sachets autorisés et imposer une gestion rationnelle des "
        "déchets plastiques (collecte, valorisation, recyclage). Abrogée et remplacée par la Loi n° 2020-04.",
        "First dedicated national plastic-bag law: ban production, import, possession, distribution and use of "
        "thin bags (<30 microns), end free distribution of bags ≥30 microns, standardise permitted bags and require "
        "rational plastic-waste management (collection, recovery, recycling). Repealed and replaced by Law No. 2020-04.",
    ),
    "policy_target": 1.0,
    "policy_target_text": (
        f"Threshold: {q('épaisseur inférieure à 30 microns', 'thickness below 30 microns')} prohibited; "
        f"{q('épaisseur supérieure ou égale à 30 microns', 'thickness greater than or equal to 30 microns')} "
        f"may not be {q('distribués ou proposés gratuitement', 'distributed or offered free of charge')}; "
        f"Art. 10 penalties: {q('10 000 000 à 20 000 000 francs CFA', 'FCFA 10,000,000 to 20,000,000')} and "
        f"{q('trois (3) mois à six (6) mois', 'three (3) to six (6) months')} imprisonment for production offences."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Loi adoptée par l'Assemblée nationale le 21 avril 2015 et promulguée le 4 mai 2015 en tant que "
        "Loi n° 2015-09. Il s'agit d'une loi parlementaire (≠0.75), première législation nationale spécifique "
        "aux sachets plastiques.",
        "Law adopted by the National Assembly on 21 April 2015 and promulgated on 4 May 2015 as Law No. 2015-09. "
        "It is parliamentary legislation (≠0.75), the first national law specifically targeting plastic bags.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "commerce de détail, plasturgie, importation, environnement, collectivités locales, santé publique",
        "retail, plastics manufacturing, import, environment, local government, public health",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "production, distribution, utilisation, fin de vie",
        "production, distribution, use, end-of-life",
    ),
    "policy_budget": 0,
    "policy_budget_text": (
        f"No dedicated budget allocation; enforcement via criminal/administrative fines "
        f"(e.g. Art. 10: {q('10 000 000 à 20 000 000 francs CFA', 'FCFA 10,000,000 to 20,000,000')}; "
        f"Art. 12: {q('20 000 à 50 000 francs CFA', 'FCFA 20,000 to 50,000')}; "
        f"Art. 14: {q('10 000 à 30 000 francs CFA', 'FCFA 10,000 to 30,000')})."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 2: {q('Sont interdites, sur toute l étendue du territoire national, la production, l importation, l utilisation, la détention en vue de la mise en vente, la mise en vente et la vente ou la distribution à titre gratuit de sachets plastiques d une épaisseur inférieure à 30 microns', 'Production, import, use, possession for sale, sale and free distribution of plastic bags under 30 microns thickness are prohibited throughout national territory')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 17: entered into force six months after JO publication. "
            f"Arts. 10–12: criminal/customs penalties. Abrogated 2020; replaced by Loi n° 2020-04."
        ),
        "comments": bi(
            "Interdiction centrale — première loi nationale dédiée aux sachets plastiques fins. Abrogée (2020).",
            "Core ban — first dedicated national thin plastic-bag law. Repealed (2020).",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Distribution",
        "instrument_description": (
            f"Art. 3: {q('Les sachets plastiques d une épaisseur supérieure ou égale à 30 microns, quel que soit l usage auquel ils sont destinés, ne peuvent être distribués ou proposés gratuitement', 'Plastic bags of thickness greater than or equal to 30 microns, whatever their intended use, may not be distributed or offered free of charge')}. "
            f"{q('Un arrêté conjoint des ministres respectivement en charge du commerce et de l environnement fixe le prix de leur cession aux utilisateurs', 'A joint order of the Ministers of Commerce and Environment sets the price at which they are transferred to users')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Price-setting via joint ministerial order; complements Art. 2 ban on thin free bags."
        ),
        "comments": bi(
            "Fin de la gratuité des sachets ≥30 microns — instrument économique de régulation des flux.",
            "End of free distribution for bags ≥30 microns — economic instrument to regulate flows.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 4: {q('Les sachets plastiques d une épaisseur supérieure ou égale à 30 microns doivent respecter les normes techniques concernant la fabrication, la composition des matériaux, l étiquetage et l écotoxicité fixées par un décret pris sur proposition du ministre en charge de l environnement', 'Plastic bags ≥30 microns must comply with technical standards on manufacture, material composition, labelling and ecotoxicity set by decree on proposal of the Environment Minister')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 1 defines {q('Ecotoxicité', 'ecotoxicity')} and {q('Sachet plastique', 'plastic bag')} (polyethylene film with handle cut-out). Standards via implementing decree."
        ),
        "comments": bi(
            "Standardisation des sachets autorisés — qualité, étiquetage et écotoxicité.",
            "Standardisation of permitted bags — quality, labelling and ecotoxicity.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 5: {q('Les industriels du plastique sont tenus de réduire les quantités de déchets plastiques qui peuvent résulter de leurs activités en développant, le cas échéant, des activités de valorisation des déchets issus de leur process ou procédés de production', 'Plastics industrialists must reduce plastic waste quantities from their activities by developing, where appropriate, recovery activities for waste from their production processes')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 13: failure to comply with Arts. 5–6 when revealed by register inspection — "
            f"{q('amende de 5 000 000 à 10 000 000 francs CFA', 'fine of FCFA 5,000,000 to 10,000,000')} and "
            f"{q('emprisonnement de un (1) mois à trois (3) mois', 'imprisonment of one (1) to three (3) months')}."
        ),
        "comments": bi(
            "Responsabilité élargie des producteurs (précurseur) — réduction et valorisation des déchets en amont.",
            "Extended producer responsibility (precursor) — upstream waste reduction and recovery.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 6: {q('Les opérateurs du secteur du plastique sont tenus de proposer aux ménages et autres utilisateurs, un système de collecte ou de reprise des déchets plastiques en vue de leur valorisation, recyclage ou élimination', 'Plastics-sector operators must offer households and other users a collection or take-back system for plastic waste for recovery, recycling or disposal')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Environment Minister sets by order, after consultative opinion, conditions for "
            f"{q('collecte ou reprise, de stockage, de tri et de transport des déchets plastiques', 'collection or take-back, storage, sorting and transport of plastic waste')} "
            f"and for {q('valorisation, recyclage ou élimination', 'recovery, recycling or disposal')}."
        ),
        "comments": bi(
            "Système de collecte/reprise obligatoire — pilier de la gestion rationnelle des déchets plastiques.",
            "Mandatory collection/take-back system — pillar of rational plastic-waste management.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 7: {q('Les opérateurs du secteur du plastique tiennent un registre dans lequel ils consignent les mesures qu ils ont l obligation de prendre en application des articles 5 et 6 de la présente loi', 'Plastics-sector operators keep a register recording measures they must take under Articles 5 and 6')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Register presented on first demand to control agents (Art. 9). "
            f"Art. 13: omission — {q('amende de 2 000 000 à 5 000 000 francs CFA', 'fine of FCFA 2,000,000 to 5,000,000')}."
        ),
        "comments": bi(
            "Obligation de traçabilité et de contrôle des mesures de gestion des déchets.",
            "Traceability and control obligation for waste-management measures.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 8: {q('Toute personne qui détient ou utilise des produits en matière plastique est tenue, lorsque ces produits deviennent des déchets, de les déposer ou de les faire acheminer vers les points de collecte ou de reprise aménagés à cet effet', 'Anyone holding or using plastic products must, when they become waste, deposit them or have them delivered to designated collection or take-back points')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 14: littering offence — {q('amende de 10 000 à 30 000 francs CFA', 'fine of FCFA 10,000 to 30,000')}."
        ),
        "comments": bi(
            "Obligation citoyenne de dépôt des déchets plastiques — complète le système de collecte Art. 6.",
            "Citizen obligation to deposit plastic waste — complements Art. 6 collection system.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: defines {q('Déchet plastique', 'plastic waste')}, {q('Gestion rationnelle des déchets plastiques', 'rational management of plastic waste')}, "
            f"{q('Recyclage', 'recycling')}, {q('Valorisation', 'recovery')} and {q('Sachet plastique', 'plastic bag')} "
            f"({q('contenant plastique ayant, dans sa partie supérieure, une découpe sous forme de bretelle (sachet bretelle) fabriqué à partir d un film polyéthylène', 'plastic container with upper handle cut-out (handle bag) made from polyethylene film')})."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Definitions frame scope: thin handle-bags (primarily imported) targeted; small retail packaging bags "
            f"excluded per parliamentary debate (rapport commission)."
        ),
        "comments": bi(
            "Définitions opérationnelles — cible les sachets bretelle fins; petits sachets d emballage du commerce de détail exclus.",
            "Operational definitions — targets thin handle-bags; small retail packaging bags excluded.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 10: {q('La production ou fabrication de sachets plastiques en infraction aux dispositions de l article 2 de la présente loi est punie d une amende de 10 000 000 à 20 000 000 francs CFA et d un emprisonnement de trois (3) mois à six (6) mois ou de l une de ces deux peines seulement', 'Production or manufacture of plastic bags in breach of Article 2 is punishable by a fine of FCFA 10,000,000 to 20,000,000 and imprisonment of three (3) to six (6) months, or one of these penalties')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 11: import of bags <30 microns is a {q('infraction douanière', 'customs offence')} under the Customs Code. "
            f"Arts. 15–16: corporate criminal liability (up to 5× fine, closure, confiscation, publication)."
        ),
        "comments": bi(
            "Sanctions pénales fortes pour production illégale de sachets fins.",
            "Strong criminal penalties for illegal production of thin bags.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Distribution",
        "instrument_description": (
            f"Art. 12: {q('L utilisation, la détention en vue de la mise en vente, la mise en vente et la vente ou la distribution à titre gratuit de sachets plastiques d une épaisseur inférieure à 30 microns sont punies d une amende de 20 000 à 50 000 francs CFA', 'Use, possession for sale, sale or free distribution of plastic bags under 30 microns is punishable by a fine of FCFA 20,000 to 50,000')}. "
            f"{q('Si l auteur de l infraction est un commerçant ou un distributeur, le montant maximal de la peine est prononcé', 'If the offender is a trader or distributor, the maximum penalty applies')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 9: offences recorded by police, Environment, Health, Industry, Commerce and Economy agents."
        ),
        "comments": bi(
            "Sanctions commerciales et de distribution — cible détaillants et distributeurs.",
            "Commercial and distribution penalties — targets retailers and distributors.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 14: {q('Est puni d une amende de 10 000 à 30 000 francs CFA quiconque abandonne ou jette des déchets plastiques ailleurs que dans les points de collecte ou de reprise prévus à cet effet', 'Anyone who abandons or throws away plastic waste other than at designated collection or take-back points is liable to a fine of FCFA 10,000 to 30,000')}. "
            f"{q('En cas de récidive, le montant maximal de la peine est prononcé', 'On recurrence, the maximum penalty applies')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Links to Art. 8 user deposit obligation and Art. 6 operator collection infrastructure."
        ),
        "comments": bi(
            "Sanction du jet de déchets plastiques — lutte contre l'abandon dans la nature.",
            "Penalty for littering plastic waste — combats abandonment in the environment.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 17: {q('La présente loi entre en vigueur à la fin du sixième mois à compter de sa publication au Journal Officiel', 'This law enters into force at the end of the sixth month following its publication in the Official Journal')}."
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Six-month transition for plastics-sector adjustment (per exposé des motifs). "
            f"Law repealed 2020; replaced by Loi n° 2020-04 on plastic bags."
        ),
        "comments": bi(
            "Loi abrogée en 2020 et remplacée par la Loi n° 2020-04 — instrument_in_force=0 pour toutes les lignes.",
            "Law repealed in 2020 and replaced by Law No. 2020-04 — instrument_in_force=0 for all rows.",
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
