#!/usr/bin/env python3
"""Generate 4P Index coding spreadsheet for Loi 2020-04 (bilingual FR/EN cells).

Source: au-senegal.com PDF (OCR) cross-checked with FAOLEX sen200382.pdf text extraction.
Official legislative portal: dri.gouv.sn/les-actes-législatifs (general index).
"""

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT = "/workspace/4P_Index_Loi_2020-04_Produits_Plastiques.xlsx"

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
        "Loi n° 2020-04 du 8 janvier 2020 relative à la prévention et à la réduction de l'incidence "
        "sur l'environnement des produits plastiques",
        "Law No. 2020-04 of 8 January 2020 on prevention and reduction of environmental impacts of plastic products",
    ),
    "policy_url": "https://www.au-senegal.com/IMG/pdf/loi-plastique-senegal-2020-04.pdf",
    "policy_year": 2020,
    "policy_objective": bi(
        "Loi phare sur les plastiques remplaçant la Loi 2015-09 : interdire les produits plastiques à usage unique/jetables "
        "et les sacs sortie de caisse ; instaurer la consigne sur les bouteilles en plastique ; mettre en place la "
        "responsabilité élargie des producteurs (REP/EPR) ; fixer des objectifs de contenu recyclé ; interdire "
        "l'importation de déchets plastiques ; instituer un prix plancher pour le recyclage et une taxe plastique "
        "sur les produits non recyclables.",
        "Flagship plastics law replacing Law 2015-09: ban single-use/disposable plastic products and checkout bags; "
        "establish plastic-bottle deposit-return; implement extended producer responsibility (EPR); set recycled-content "
        "targets; ban plastic-waste imports; institute recycling floor price and plastic tax on non-recyclable products.",
    ),
    "policy_target": 1.0,
    "policy_target_text": (
        f"Art. 4: ban on {q('produits plastiques à usage unique ou produits plastiques jetables', 'single-use or disposable plastic products')} "
        f"(cups, cutlery, straws, water/beverage sachets); "
        f"Art. 5: ban on {q('sacs plastiques sortie de caisse', 'checkout plastic bags')} regardless of thickness; "
        f"Art. 16: recycled-content targets set by decree; Art. 19: {q('importation de déchets plastiques... interdite', 'import of plastic waste... prohibited')}."
    ),
    "policy_type": 1.0,
    "policy_type_justification": bi(
        "Loi adoptée par l'Assemblée nationale le 30 décembre 2019 et promulguée le 8 janvier 2020 en tant que "
        "Loi n° 2020-04. Il s'agit d'une loi parlementaire (≠0.75), loi phare nationale sur les produits plastiques.",
        "Law adopted by the National Assembly on 30 December 2019 and promulgated on 8 January 2020 as Law No. 2020-04. "
        "It is parliamentary legislation (≠0.75), the national flagship law on plastic products.",
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": bi(
        "plasturgie, fabrication, importation, commerce de détail, distribution, consommation, gestion des déchets, "
        "recyclage, environnement, santé publique",
        "plastics manufacturing, production, import, retail, distribution, consumption, waste management, "
        "recycling, environment, public health",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "production, distribution, utilisation, collecte, recyclage, élimination, gouvernance",
        "production, distribution, use, collection, recycling, disposal, governance",
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        f"Art. 21: {q('prix plancher', 'floor price')} for recyclers buying plastic waste (set by decree); "
        f"Art. 22: {q('taxe plastique', 'plastic tax')} on non-recyclable plastic products (list and rates by decree); "
        f"Arts. 26–39: criminal fines up to FCFA 100,000,000 and imprisonment up to 5 years; "
        f"Art. 24: financial settlement option for certain offences."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 1: {q('La présente loi fixe les règles relatives à la prévention et la réduction de l impact sur l environnement et la santé humaine des produits en plastique et à la gestion écologique rationnelle des déchets plastiques', 'This law sets rules on preventing and reducing environmental and human-health impacts of plastic products and on rational ecological management of plastic waste')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 2: applies to products made from plastic materials (single-use or not) and resulting waste; "
            f"Art. 42: enters into force 3 months after JO publication; Art. 41: implementing decrees."
        ),
        "comments": bi(
            "Loi-cadre phare — remplace et abroge la Loi 2015-09 (Art. 40).",
            "Flagship framework law — replaces and repeals Law 2015-09 (Art. 40).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 4: prohibits production, import, detention for sale, sale, supply and use of "
            f"{q('produits plastiques à usage unique ou produits plastiques jetables', 'single-use or disposable plastic products')}, "
            f"including cups/glasses/lids, cutlery/plates, straws/stirrers, and sachets for water/beverages."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 26: manufacture/import offence — 1–3 years imprisonment and FCFA 5–10 million fine; "
            f"Art. 27: sale/use offence — 1–3 months and FCFA 50,000–100,000; Art. 23: seizure of prohibited products."
        ),
        "comments": bi(
            "Interdiction centrale des plastiques à usage unique — bien au-delà des sachets fins de 2015-09.",
            "Central single-use plastic ban — far beyond 2015-09 thin-bag scope.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Distribution",
        "instrument_description": (
            f"Art. 5: {q('Les sacs plastiques sortie de caisse... sont interdits, quelle que soit leur épaisseur', 'Checkout plastic bags... are prohibited regardless of thickness')} "
            f"(with or without handles/straps; includes biodegradable/oxo-degradable bags per exposé des motifs)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 5(2): exception for transparent recyclable bags used at point of sale to package food for protection/handling "
            f"(producer to consumer); import subject to prior Environment Minister authorisation."
        ),
        "comments": bi(
            "Interdiction des sacs caisse — exception emballage alimentaire primaire transparent recyclable.",
            "Checkout bag ban — exception for transparent recyclable primary food packaging.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 6: {q('Une consigne est exigée à l achat de tout produit contenu dans des bouteilles en plastique', 'A deposit is required on purchase of any product in plastic bottles')}; "
            f"amount set by decree, collected by seller and refunded on return of empty bottle."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 7: sellers must accept returned bottles and deliver to nearest collection point; "
            f"Art. 33: refusal penalty — 15 days to 1 month imprisonment and FCFA 50,000–100,000."
        ),
        "comments": bi(
            "Système de consigne bouteilles PET — instrument de collecte direct.",
            "Plastic-bottle deposit-return system — direct collection instrument.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 8: producers must establish collection points for plastic bottles and "
            f"{q('valoriser ou faire valoriser', 'recover or have recovered')} collected bottles prioritising reuse, recycling then other recovery."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Arts. 9–10: biannual sectoral report to Environment Minister (market vs collected volumes, collection points, gap measures); "
            f"Art. 34: insufficient collection points — 3–6 months and FCFA 5–10 million."
        ),
        "comments": bi(
            "Obligations producteurs bouteilles — points de collecte et valorisation.",
            "Bottle-producer obligations — collection points and recovery.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 11: producers placing plastic products on the market are {q('responsables de la gestion des déchets générés par ces produits', 'responsible for managing waste generated by these products')} "
            f"(extended producer responsibility); may comply via individual programmes or eco-organisms."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 35: EPR non-compliance — 1–3 years imprisonment and FCFA 10–20 million; "
            f"Arts. 12–14: approved individual programmes (3-year renewable) and accredited eco-organisms (max 10 years)."
        ),
        "comments": bi(
            "REP/EPR — pilier de la gouvernance déchets plastiques post-consommation.",
            "EPR — pillar of post-consumer plastic-waste governance.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": (
            f"Art. 12: individual collection/treatment programmes approved by Environment Ministerial order for 3 years (renewable); "
            f"minimum requirements by order; periodic sworn-agent controls; suspension up to 3 months or permanent cessation for non-compliance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 13: eco-organisms improve selective collection and treatment; 10-year max accreditation; charter with assigned objectives; "
            f"Art. 14: annual activity report by 30 April."
        ),
        "comments": bi(
            "Mécanismes opérationnels REP — programmes individuels et éco-organismes.",
            "Operational EPR mechanisms — individual programmes and eco-organisms.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 15: producers must reduce waste at source and market products that, once waste, can be "
            f"{q('recyclés ou valorisés', 'recycled or recovered')} in environmentally sound conditions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Exposé des motifs: systemic approach — production reduction, resource efficiency, circular economy transition."
        ),
        "comments": bi(
            "Prévention à la source — réduction des déchets plastiques en amont.",
            "Source prevention — upstream plastic-waste reduction.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            f"Art. 16: when technically feasible and economically viable, producers must integrate "
            f"{q('une part de plastique recyclé dans les produits plastiques neufs', 'a share of recycled plastic in new plastic products')}; "
            f"decree sets national recycled-content targets and deadlines."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 32: non-compliance — 1–3 months imprisonment and FCFA 5–10 million fine when technically/economically viable."
        ),
        "comments": bi(
            "Objectifs de contenu recyclé — levier économie circulaire.",
            "Recycled-content targets — circular-economy lever.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Distribution",
        "instrument_description": (
            f"Art. 17: plastic products placed on the market must bear visible, legible, indelible marking on packaging or product "
            f"indicating producer identity/company name and address."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 30: marking breach — 3–6 months imprisonment and FCFA 2–5 million fine."
        ),
        "comments": bi(
            "Traçabilité producteur — soutient le contrôle et la REP.",
            "Producer traceability — supports enforcement and EPR.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            f"Art. 18: consumers and end-users must, when plastic products become waste, "
            f"{q('les acheminer vers les points de collectes aménagés à cet effet', 'deliver them to designated collection points')}."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 37: abandoning plastic waste outside collection points — 15 days to 1 month imprisonment and FCFA 20,000–50,000."
        ),
        "comments": bi(
            "Obligation consommateur — acheminement vers points de collecte.",
            "Consumer obligation — deliver to collection points.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 19: {q('L importation de déchets plastiques sur le territoire national est interdite', 'Import of plastic waste into national territory is prohibited')}; "
            f"illegal imports seized and re-exported at importer's cost. "
            f"Art. 20: exports only with Environment Minister authorisation to countries with adequate treatment facilities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 28: illegal import — 3–5 years and FCFA 50–100 million; Art. 29: unauthorised export — same penalties."
        ),
        "comments": bi(
            "Contrôle transfrontalier des déchets plastiques — alignement Convention de Bâle.",
            "Cross-border plastic-waste control — Basel Convention alignment.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Valorisation",
        "instrument_description": (
            f"Art. 21: establishes a {q('prix plancher', 'floor price')} at which recycling enterprises must buy plastic waste per kg (set by decree). "
            f"Art. 22: institutes {q('taxe plastique', 'plastic tax')} on non-recyclable plastic products (product list, rates and collection by decree)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            f"Art. 31: buying below floor price — FCFA 2–5 million fine; tax incentivises recyclable materials."
        ),
        "comments": bi(
            "Instruments financiers/fiscaux — soutien au recyclage et pénalisation du non-recyclable.",
            "Financial/fiscal instruments — support recycling and penalise non-recyclables.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 23: prohibited products held or placed on the market are seized by control agents (Art. 25: Environment, Health, "
            f"Industry, Commerce, Finance ministries). Art. 24: financial settlement for offences under Arts. 26, 27, 30, 31, 33, 37."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            f"Art. 39: corporate penalties include fines up to 5× individual max, closure up to 5 years, confiscation, publication."
        ),
        "comments": bi(
            "Mécanismes d'application — saisie, transaction et sanctions personnes morales.",
            "Enforcement mechanisms — seizure, settlement and corporate sanctions.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Governance",
        "instrument_description": (
            f"Art. 40: {q('La loi n° 2015-09... est abrogée', 'Law No. 2015-09... is repealed')} "
            f"(thin plastic-bag ban replaced by comprehensive plastics framework)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            f"Promulgated 8 January 2020; effective 3 months after JO publication; current flagship plastics law in Senegal."
        ),
        "comments": bi(
            "Abrogation Loi 2015-09 — instrument_in_force=1 pour toutes les lignes de la Loi 2020-04.",
            "Repeal of Law 2015-09 — instrument_in_force=1 for all Law 2020-04 rows.",
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
