"""
Builds the 4P Index (Plastic Pollution Policy Index) coding table for:

Loi n. 2015-18 portant Code de la Peche maritime (Senegal).

Source: fetched directly from the task-brief URL (https://www.ditp.gouv.sn/download/file/fid/53),
a scanned 36-page PDF; OCR'd with tesseract (French model, 300dpi) since the PDF had no extractable
text layer. Confirmed authentic: "L'Assemblee nationale a adopte en sa seance du mardi 30 juin 2015,
Le President de la Republique promulgue la loi..."; exposé des motifs explicitly references
replacement of Loi n. 98-32 du 14 avril 1998.

KEY FINDING: Art. 66 explicitly bans nylon monofilament/multimonofilament fishing nets -- nylon is a
synthetic polymer (plastic) material -- directly satisfying Plastics Relevance Filter criterion (a).
This type of gear is a major contributor to "ghost gear" marine plastic pollution and wildlife
entanglement worldwide. No other explicit plastics/synthetic-polymer reference was found in the
36-page text (confirmed by full-text search for "plastique", "nylon", "dechet", "immersion",
"rejet", "abandon", "epave", "detritus", "residu", "ordure").

Produces:
  - senegal_loi_2015-18_4p_index.csv
  - senegal_loi_2015-18_4p_index.xlsx (formatted, wrapped, one sheet)
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

POLICY_NAME = "Loi n° 2015-18 portant Code de la Pêche maritime"
POLICY_URL = "https://www.ditp.gouv.sn/download/file/fid/53"
POLICY_YEAR = 2015
POLICY_OBJECTIVE = (
    "Modernises Senegal's maritime fisheries management framework (replacing Loi n° 98-32 du 14 "
    "avril 1998), introducing measures against illegal/unreported/unregulated (IUU) fishing, "
    "fisheries co-management ('cogestion'), ecosystem-based conservation tools (marine protected "
    "areas, fish-aggregating devices, artificial reefs), port-State measures, a vessel monitoring "
    "system, and strengthened sanctions. It is not a plastics policy, and this Code's only explicit "
    "plastics/synthetic-polymer reference is Art. 66, which bans 'l'importation, la mise en vente, "
    "l'achat, la détention et l'utilisation des nappes et filets maillants fabriqués à partir "
    "d'éléments monofilaments ou multimonofilaments en nylon' -- nylon monofilament/multimonofilament "
    "gillnets. Nylon is a synthetic polymer (plastic) material, and monofilament nylon nets are a "
    "well-documented major contributor to 'ghost gear' -- lost, abandoned or discarded fishing gear "
    "that persists in the marine environment, entangles wildlife, and constitutes a significant "
    "share of ocean plastic pollution by mass. This targeted, technology-based ban is the Code's "
    "principal plastics-relevant instrument, consistent with the task brief's framing ('relevant to "
    "fishing gear and marine plastic litter')."
)
POLICY_TARGET = 0
POLICY_TARGET_TEXT = ""
POLICY_TYPE = 1
POLICY_TYPE_JUSTIFICATION = (
    "Full Act of Parliament: 'L'Assemblée nationale a adopté en sa séance du mardi 30 juin 2015, Le "
    "Président de la République promulgue la loi dont la teneur suit.' Adopted by the National "
    "Assembly and promulgated by the President -> legislation, not an executive/ministerial "
    "instrument."
)
POLICY_INTEGRATION = 1
POLICY_SECTORS_LIST = (
    "fisheries, maritime/marine transport, environment, aquaculture, industry, trade, tourism, "
    "research"
)
POLICY_CIRCULARITY = 0.75
POLICY_LIFECYCLE_PHASES_LIST = "production, consumption, environmental leakage"
POLICY_BUDGET = 0
POLICY_BUDGET_TEXT = ""
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
        Q_instrument_lifecycle_stage="Environmental leakage",
        R_instrument_description=(
            "Art. 66: 'Sont interdits l'importation, la mise en vente, l'achat, la détention et "
            "l'utilisation des nappes et filets maillants fabriqués à partir d'éléments monofilaments "
            "ou multimonofilaments en nylon sauf dérogation spéciale.' A comprehensive, supply-chain-"
            "wide ban -- covering import, sale, purchase, possession and use -- on nylon monofilament/"
            "multimonofilament gillnets, subject to a special derogation."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.75,
        U_instrument_implementation_text=(
            "Authority: enforced under the Ministre chargé de la pêche maritime's general fisheries-"
            "control and surveillance apparatus (Art. 13, 33(d)) (+0.25). Enforcement: Art. 125(a) "
            "qualifies 'l'usage d'engins...interdits' as a 'very grave' industrial-fishing offence, "
            "punished by a fine of 20,000,000 to 30,000,000 FCFA, with the prohibited gear "
            "'confisqués et détruits'; Art. 126 extends the same offence to artisanal fishing (fine of "
            "150,000 to 300,000 FCFA), also with confiscation/destruction of the gear (+0.25). "
            "Monitoring: the Code establishes a general 'système de contrôle et de surveillance des "
            "pêches maritimes' including a vessel monitoring system and onboard observers (Art. 33(d)-"
            "(f)) (+0.25). Unconditional: not awarded -- Art. 66 itself carves out an exception "
            "('sauf dérogation spéciale')."
        ),
        W_comments=(
            "The Code's only explicit plastics/synthetic-polymer provision (Plastics Relevance Filter "
            "criterion (a)). Full-text search confirms no other mention of 'plastique', 'déchet', "
            "'immersion', 'rejet', 'abandon', 'épave', 'détritus', 'résidu' or 'ordure' anywhere in "
            "the 36-page Code; the general marine-ecosystem-management title (Art. 14-21: marine "
            "protected areas, fish-aggregating devices, artificial reefs) was reviewed and excluded "
            "from this table as a fisheries/biodiversity conservation tool rather than a waste-"
            "management or plastics-relevant governance instrument under any of the three filter "
            "criteria."
        ),
    ),
    dict(
        P_instrument_type=0.40,
        Q_instrument_lifecycle_stage="Waste management",
        R_instrument_description=(
            "Art. 5-6, 22-23 and 33(g): the State promotes fisheries co-management ('cogestion des "
            "pêcheries') and participatory decision-making with professional organisations and fishing "
            "communities (Art. 5-6); a national 'Conseil national consultatif des Pêches maritimes' is "
            "created, chaired by the Director of Maritime Fisheries (Art. 22); regional local artisanal-"
            "fishing councils may be established (Art. 23); and implementing regulations are to fix the "
            "conditions for exercising 'la surveillance participative' -- community-based monitoring of "
            "compliance with fishing rules, including gear restrictions such as the nylon-net ban above "
            "(Art. 33(g))."
        ),
        S_instrument_in_force=1,
        T_instrument_implementation=0.50,
        U_instrument_implementation_text=(
            "Authority: the national Conseil is explicitly created and chaired by 'le Directeur des "
            "Pêches maritimes' (Art. 22) (+0.25). Monitoring: 'la surveillance participative' is an "
            "explicit, named community-monitoring mechanism anticipated by Art. 33(g), consistent with "
            "the exposé des motifs' description of a new 'surveillance participative' regime (+0.25). "
            "Enforcement: no penalty is tied to non-participation or non-establishment of local "
            "councils -- not awarded. Unconditional: not awarded -- Art. 23 expressly states local "
            "councils 'peuvent être institués', an optional/enabling formulation."
        ),
        W_comments=(
            "Coded per Rule 11 as a multi-level governance/coordination instrument (national Conseil + "
            "optional regional councils + professional organisations), included because participatory "
            "surveillance directly supports enforcement of gear-type restrictions such as the plastics-"
            "relevant nylon-net ban above. In-force uncertainty (Rule 8): the national Conseil's "
            "creation (Art. 22, 'il est créé') is mandatory, but the regional local councils (Art. 23) "
            "and the specific 'surveillance participative' operating rules (Art. 33(g)) remain "
            "conditional on further implementing texts; S=1 is applied on the strength of the "
            "mandatory national-Conseil creation, with the conditional elements reflected in the T=0.50 "
            "score rather than in S."
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

with open("4p_index/senegal_loi_2015-18_4p_index.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=COLUMNS)
    writer.writeheader()
    for row in rows:
        writer.writerow(row)

df = pd.DataFrame(rows, columns=COLUMNS)
xlsx_path = "4p_index/senegal_loi_2015-18_4p_index.xlsx"
df.to_excel(xlsx_path, sheet_name="4P_Index_Loi2015-18", index=False)

wb = load_workbook(xlsx_path)
ws = wb["4P_Index_Loi2015-18"]

header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
header_font = Font(bold=True, color="FFFFFF")
wrap_top = Alignment(wrap_text=True, vertical="top")

col_widths = {
    "A": 34, "B": 30, "C": 10, "D": 46, "E": 10, "F": 30, "G": 10, "H": 34,
    "I": 10, "J": 30, "K": 10, "L": 26, "M": 10, "N": 30, "O": 16,
    "P": 12, "Q": 16, "R": 46, "S": 10, "T": 12, "U": 40, "V": 16, "W": 40,
}
for letter in [get_column_letter(c) for c in range(1, len(COLUMNS) + 1)]:
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
    ws.row_dimensions[r].height = 120

wb.save(xlsx_path)
print("Wrote", xlsx_path, "and CSV with", len(rows), "instrument rows.")
print("Policy score O =", O_score)
for row in rows:
    print(row["Q_instrument_lifecycle_stage"], "->", row["V_instrument_score"])
