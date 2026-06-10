import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "4P Index Coding"

# ── Colour palette ─────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"   # header background
MID_BLUE    = "2E75B6"   # section-A sub-header
TEAL        = "17375E"   # section-B sub-header
LIGHT_BLUE  = "D9E2F3"   # policy-level fill
LIGHT_GREEN = "E2EFDA"   # instrument fill
ORANGE      = "FCE4D6"   # auto-calculated field fill
YELLOW      = "FFFF99"   # highlight row 1
WHITE       = "FFFFFF"
LIGHT_GREY  = "F2F2F2"

def hfill(hex_):
    return PatternFill("solid", fgColor=hex_)

def hfont(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

thin = Side(style="thin", color="BFBFBF")
medium = Side(style="medium", color="808080")

def thin_border():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def medium_border():
    return Border(left=medium, right=medium, top=medium, bottom=medium)

def wrap(ws, row, col, value, fill=None, font=None, align=None, border=None):
    cell = ws.cell(row=row, column=col, value=value)
    if fill:   cell.fill   = fill
    if font:   cell.font   = font
    if align:  cell.alignment = align
    if border: cell.border = border
    return cell

WRAP_ALIGN    = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER_ALIGN  = Alignment(wrap_text=True, vertical="center", horizontal="center")
CENTER_TOP    = Alignment(wrap_text=True, vertical="top",    horizontal="center")

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 1  –  Main title
# ═══════════════════════════════════════════════════════════════════════════════
ws.merge_cells("A1:X1")
c = ws["A1"]
c.value = "Plastic Pollution Policy Index (4P Index) — Coding Table"
c.fill  = hfill(DARK_BLUE)
c.font  = hfont(bold=True, size=13, color=WHITE)
c.alignment = CENTER_ALIGN
ws.row_dimensions[1].height = 22

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 2  –  Document metadata
# ═══════════════════════════════════════════════════════════════════════════════
ws.merge_cells("A2:X2")
c = ws["A2"]
c.value = ("Policy: Karnali Province Tourism Master Plan 2076/77–2085/86 BS (2020/21–2029/30)  |  "
           "Country: Nepal (Karnali Province)  |  Year: 2020  |  "
           "Source: https://faolex.fao.org/docs/pdf/nep220153.pdf")
c.fill  = hfill(MID_BLUE)
c.font  = hfont(bold=False, size=9, color=WHITE)
c.alignment = CENTER_ALIGN
ws.row_dimensions[2].height = 16

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 3  –  Section labels (Section A / Section B)
# ═══════════════════════════════════════════════════════════════════════════════
ws.merge_cells("A3:O3")
c = ws["A3"]
c.value = "SECTION A — Policy-Level Fields  (repeated identically across all rows)"
c.fill  = hfill(MID_BLUE)
c.font  = hfont(bold=True, size=10, color=WHITE)
c.alignment = CENTER_ALIGN

ws.merge_cells("P3:X3")
c = ws["P3"]
c.value = "SECTION B — Instrument-Level Fields  (one row per instrument)"
c.fill  = hfill(TEAL)
c.font  = hfont(bold=True, size=10, color=WHITE)
c.alignment = CENTER_ALIGN
ws.row_dimensions[3].height = 18

# ═══════════════════════════════════════════════════════════════════════════════
# ROW 4  –  Column headers
# ═══════════════════════════════════════════════════════════════════════════════
headers = [
    # Section A (1–15)
    ("A", "policy_name"),
    ("B", "policy_url"),
    ("C", "policy_year"),
    ("D", "policy_objective"),
    ("E", "policy_target\n(0/1)"),
    ("F", "policy_target_text"),
    ("G", "policy_type\n(score)"),
    ("H", "policy_type_justification"),
    ("I", "policy_integration\n(score)"),
    ("J", "policy_sectors_list"),
    ("K", "policy_circularity\n(score)"),
    ("L", "policy_lifecycle_phases_list"),
    ("M", "policy_budget\n(score)"),
    ("N", "policy_budget_text"),
    ("O", "policy_score\n[auto]"),
    # Section B (16–24)
    ("P", "instrument_type\n(score)"),
    ("Q", "instrument_lifecycle_stage"),
    ("R", "instrument_description"),
    ("S", "instrument_in_force\n(0/1)"),
    ("T", "instrument_implementation\n(score)"),
    ("U", "instrument_implementation_text"),
    ("V", "instrument_score\n[auto]"),
    ("W", "comments"),
    ("X", "—"),  # spare
]

for col_idx, (col_letter, label) in enumerate(headers, start=1):
    cell = ws.cell(row=4, column=col_idx, value=f"Col {col_letter}\n{label}")
    if col_idx <= 15:
        cell.fill = hfill("C5D9F1")   # light blue header
        cell.font = hfont(bold=True, size=9, color=DARK_BLUE)
    else:
        cell.fill = hfill("C4E1C0")   # light green header
        cell.font = hfont(bold=True, size=9, color="1F4E17")
    cell.alignment = CENTER_TOP
    cell.border = thin_border()

ws.row_dimensions[4].height = 36

# ═══════════════════════════════════════════════════════════════════════════════
# DATA — shared policy fields
# ═══════════════════════════════════════════════════════════════════════════════
policy_name  = "Karnali Province Tourism Master Plan 2076/77 – 2085/86 BS (2020/21–2029/30)"
policy_url   = "https://faolex.fao.org/docs/pdf/nep220153.pdf"
policy_year  = 2020
policy_obj   = ("To develop sustainable tourism in Karnali Province (Nepal) as an engine for "
                "economic and social transformation while preserving natural and cultural heritages; "
                "includes aspirational measures to ban plastic use in protected areas, establish solid "
                "waste management facilities and guidelines in tourism areas, and prevent plastic "
                "pollution at specific lake and wetland eco-tourism sites.")
policy_tgt   = 0
policy_tgt_t = ""
policy_type  = 0.25
policy_type_j= ("Sub-national provincial tourism master plan published January 2020 by the Ministry "
                "of Industry, Tourism, Forest and Environment (MoITFE) of Karnali Province, with WWF "
                "Nepal technical assistance. Issued as a guiding planning document through a consultative "
                "process — not enacted by any legislature. Contains aspirational commitments on plastic "
                "bans and waste management without quantifiable plastic-specific targets.")
policy_intg  = 1
policy_sect  = "tourism, waste management, conservation, water, energy, transport, agriculture, municipalities, infrastructure"
policy_circ  = 0.75
policy_lc    = "consumption, disposal, environmental leakage"
policy_budg  = 0.5
policy_budg_t= ("Total plan budget NRs. 10,300,000,000. Programme 7 (Cultural and Natural Heritage "
                "Conservation), covering plastic ban and solid waste measures = NRs. 1,030,000,000 (10%). "
                "Programme 2 (Tourism Infrastructure Development and Upgrading), including waste "
                "management infrastructure = NRs. 3,605,000,000 (35%). Funded by three tiers of "
                "government plus development partners/donors. Not ring-fenced for plastic. "
                "(Executive Summary, p. iv–vii)")
policy_score_formula = "=AVERAGE(G{row},I{row},K{row},M{row},P{row})"

# ── Instrument rows ─────────────────────────────────────────────────────────
instruments = [
    {
        "P": 1.0,
        "Q": "Consumption",
        "R": ("Section 5.7.5 (p. 99), Programme 7 — Cultural and Natural Heritage Conservation: "
              "The plan commits to banning the use of plastic in Protected Areas as a measure to "
              "minimise and mitigate potential negative impacts of tourism on natural heritage. "
              "Text: '– Ban on use of plastic in Protected Area.' "
              "Relevant PAs in Karnali Province: Rara National Park and Shey-Phoksundo National Park."),
        "S": 0,
        "T": 0.50,
        "U": ("Sec. 5.7.5 (p. 99): '– Ban on use of plastic in Protected Area. Ensure safe and "
              "adequate waste and sewage management facilities in touristic areas.'\n"
              "Action Plan (p. 127): '7.5 … Implement solid waste management guidelines / "
              "Leading: MoITFE, M/RM / Partner: DFO, BZUC, PAs.'\n"
              "→ +0.25 responsible authority (MoITFE, M/RM designated at programme level).\n"
              "→ +0 enforcement (no penalties stated).\n"
              "→ +0 monitoring (no mechanism stated).\n"
              "→ +0.25 unconditional (no exemptions stated)."),
        "V_formula": True,
        "W": ("Coded as Regulatory (ban) — explicit prohibition language. instrument_in_force = 0: "
              "ban appears in a planning document, not an enacted regulation; no subordinate binding "
              "instrument found. Responsible authority (+0.25) designated at programme level, not "
              "instrument-specifically — borderline case. If a subsequent provincial/local government "
              "ordinance operationalises this ban, update S to 1.\n"
              "Cross-reference: Nepal's national Plastic Bag Control Directive 2082 applies nationwide; "
              "this plan's PA-level ban is narrower but potentially stricter for PA settings."),
    },
    {
        "P": 0.80,
        "Q": "Waste management",
        "R": ("Sections 5.7.5 and Chapter 7 Action Plan item 7.5: The plan commits to "
              "(a) ensuring safe and adequate waste and sewage management facilities in touristic areas; "
              "(b) controlling and managing solid waste and sludge in high-altitude Yarsagumba collection "
              "areas, including a waste collection and management system; "
              "(c) implementing solid waste management guidelines. "
              "MoITFE and Municipalities/Rural Municipalities (M/RM) are designated lead institutions; "
              "DFO, BZUC, and PA offices are partners."),
        "S": 0,
        "T": 0.50,
        "U": ("Sec. 5.7.5 (p. 99): '– Ensure safe and adequate waste and sewage management facilities "
              "in touristic areas.'\n"
              "'– Control and manage solid waste, sludge … waste collection and management system …'\n"
              "Action Plan (p. 127): 'Implement solid waste management guidelines / "
              "Leading: MoITFE, M/RM / Partner: DFO, BZUC, PAs.'\n"
              "→ +0.25 responsible authority.\n"
              "→ +0 enforcement. → +0 monitoring. → +0.25 unconditional."),
        "V_formula": True,
        "W": ("Combines infrastructure (waste/sewage facilities) and governance (guidelines) elements; "
              "highest type — Infrastructure (0.80) — applied per coding rules; governance aspect noted. "
              "No explicit mention of plastic in this instrument; relevance via Criterion (c) "
              "(waste management instrument commonly applying to plastic waste streams). "
              "BORDERLINE: P = 0.20 (governance/planning requirement) defensible if facility "
              "commitment is viewed as purely aspirational rather than a specific construction project.\n"
              "Contextual plastic evidence — Sec. 2.8.1 (p. 27): "
              "'Plastic wrapper, mineral plastic bottles, glass bottles and tin that are not "
              "biodegradable pose threats to environment.'"),
    },
    {
        "P": 1.0,
        "Q": "Consumption",
        "R": ("Appendix 1 — Strategic Tourism Projects (p. 138): Development plans for two lake "
              "eco-tourism hubs include explicit bans on plastic and bottles as conservation measures.\n"
              "(1) Syarpu Tal Eco-tourism Hub (West Rukum, Banfikot Rural Municipality): "
              "'ban on use of plastic and bottles.'\n"
              "(2) Kubhinde Daha Eco-tourism Hub (Salyan): 'ban on use of plastic and bottles.'\n"
              "Both bans appear alongside siltation control, drainage management, invasive species "
              "removal, and habitat restoration measures."),
        "S": 0,
        "T": 0.25,
        "U": ("Appendix 1, Syarpu Tal and Kubhinde Daha project descriptions (p. 138):\n"
              "'To conserve the lake, develop siltation controlling mechanism at the mouth (head), "
              "control draining house-hold sewage into lake water (develop outer drainage), re-store "
              "endemic fish by removing invasive imported fish species; plant local but ornamental trees "
              "and flowers around the lake road; ban on use of plastic and bottles.'\n"
              "(Identical language at both lake sites.)\n"
              "→ +0 responsible authority (not designated for site-specific ban).\n"
              "→ +0 enforcement. → +0 monitoring.\n"
              "→ +0.25 unconditional (no exemptions stated)."),
        "V_formula": True,
        "W": ("Site-specific bans at two lakes that are NOT designated national protected areas "
              "(making this distinct from Instrument 1). Identical wording at both sites suggests "
              "template language. If Syarpu Tal or Kubhinde Daha are subsequently declared provincial "
              "conservation areas, this instrument could be subsumed under Instrument 1.\n"
              "T = 0.25: only 'unconditional' sub-score met; no authority, enforcement, or monitoring "
              "explicitly evidenced at the project level."),
    },
]

# ═══════════════════════════════════════════════════════════════════════════════
# Write data rows  (rows 5, 6, 7)
# ═══════════════════════════════════════════════════════════════════════════════
instr_labels = [
    "Instrument 1 — Ban on plastic in Protected Areas",
    "Instrument 2 — Solid waste management facilities & guidelines for tourism areas",
    "Instrument 3 — Plastic/bottle bans at Syarpu Tal & Kubhinde Daha lake sites",
]

row_fills = [LIGHT_BLUE, LIGHT_BLUE, LIGHT_BLUE]   # policy columns
instr_fills = ["E2EFDA", "E2EFDA", "FFF2CC"]        # instrument columns (3rd slightly different)

for i, (instr, label) in enumerate(zip(instruments, instr_labels), start=1):
    r = 4 + i  # rows 5, 6, 7

    # ── Row label (hidden in col X for reference) ─────────────────────
    ws.cell(row=r, column=24, value=label).font = hfont(italic=True, size=8, color="808080")

    # ── Policy-level columns (A–O) ────────────────────────────────────
    pol_data = [
        policy_name, policy_url, policy_year, policy_obj, policy_tgt,
        policy_tgt_t, policy_type, policy_type_j, policy_intg, policy_sect,
        policy_circ, policy_lc, policy_budg, policy_budg_t,
        None,   # policy_score — auto formula
    ]

    pol_fill = hfill(LIGHT_BLUE)
    auto_fill = hfill(ORANGE)

    for col_idx, val in enumerate(pol_data, start=1):
        is_auto = (col_idx == 15)  # Col O
        cell = ws.cell(row=r, column=col_idx)
        if is_auto:
            cell.value   = f"=AVERAGE(G{r},I{r},K{r},M{r},P{r})"
            cell.fill    = auto_fill
            cell.font    = hfont(italic=True, size=9, color="7F3F00")
            cell.number_format = "0.00"
        else:
            cell.value   = val
            cell.fill    = pol_fill
            cell.font    = hfont(size=9)
            if col_idx == 3:   # year — center
                cell.alignment = CENTER_ALIGN
                cell.number_format = "0"
            elif col_idx in (5, 7, 9, 11, 13):  # numeric scores
                cell.alignment = CENTER_ALIGN
                cell.number_format = "0.00"
            else:
                cell.alignment = WRAP_ALIGN
        cell.border = thin_border()

    # ── Instrument-level columns (P–W) ────────────────────────────────
    instr_fill = hfill(instr_fills[i - 1])

    instr_data = [
        instr["P"], instr["Q"], instr["R"], instr["S"],
        instr["T"], instr["U"],
        None,   # V — auto formula
        instr["W"],
    ]

    for j, val in enumerate(instr_data):
        col_idx = 16 + j   # cols 16–23 = P–W
        is_auto = (col_idx == 22)   # Col V
        cell = ws.cell(row=r, column=col_idx)
        if is_auto:
            cell.value   = f"=AVERAGE(P{r},T{r})"
            cell.fill    = auto_fill
            cell.font    = hfont(italic=True, size=9, color="7F3F00")
            cell.number_format = "0.000"
        else:
            cell.value   = val
            cell.fill    = instr_fill
            cell.font    = hfont(size=9)
            if col_idx in (16, 18, 19, 20):  # numeric scores + lifecycle stage
                cell.alignment = CENTER_ALIGN
                cell.number_format = "0.00" if col_idx in (16, 20) else "0"
            else:
                cell.alignment = WRAP_ALIGN
        cell.border = thin_border()

    # ── Row height ─────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 200

# ═══════════════════════════════════════════════════════════════════════════════
# Column widths
# ═══════════════════════════════════════════════════════════════════════════════
col_widths = {
    1:  28,   # A policy_name
    2:  32,   # B policy_url
    3:   8,   # C policy_year
    4:  35,   # D policy_objective
    5:   8,   # E policy_target
    6:  20,   # F policy_target_text
    7:   8,   # G policy_type
    8:  35,   # H policy_type_justification
    9:   9,   # I policy_integration
    10: 35,   # J policy_sectors_list
    11:  8,   # K policy_circularity
    12: 28,   # L policy_lifecycle_phases_list
    13:  8,   # M policy_budget
    14: 40,   # N policy_budget_text
    15:  9,   # O policy_score [auto]
    16: 10,   # P instrument_type
    17: 16,   # Q instrument_lifecycle_stage
    18: 42,   # R instrument_description
    19: 10,   # S instrument_in_force
    20: 10,   # T instrument_implementation
    21: 42,   # U instrument_implementation_text
    22:  9,   # V instrument_score [auto]
    23: 42,   # W comments
    24: 20,   # X label
}

for col, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = width

# ═══════════════════════════════════════════════════════════════════════════════
# Freeze panes — freeze rows 1-4 and columns A-C
# ═══════════════════════════════════════════════════════════════════════════════
ws.freeze_panes = "D5"

# ═══════════════════════════════════════════════════════════════════════════════
# Add a second sheet: Score Reference
# ═══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Score Reference")

ref_rows = [
    ("POLICY TYPE (Col G)", None),
    ("Score", "Description"),
    (0.25, "Strategy, plan, or programme — aspirational commitments without concrete targets"),
    (0.50, "Strategy, plan, or programme — incorporates quantifiable targets"),
    (0.75, "Regulation or executive decree (sub-legislative, issued by executive branch)"),
    (1.00, "Legislation enacted by parliament or equivalent legislative body"),
    ("", ""),
    ("POLICY INTEGRATION (Col I)", None),
    ("Score", "Description"),
    (0,    "0 sectors"),
    (0.25, "1–2 sectors"),
    (0.50, "3–4 sectors"),
    (0.75, "5–6 sectors"),
    (1.00, "7 or more sectors"),
    ("", ""),
    ("POLICY CIRCULARITY (Col K)", None),
    ("Score", "Description"),
    (0.25, "1 life-cycle phase"),
    (0.50, "2 life-cycle phases"),
    (0.75, "3–4 life-cycle phases"),
    (1.00, "All 5 phases (production, consumption, recycling, disposal, environmental leakage)"),
    ("", ""),
    ("POLICY BUDGET (Col M)", None),
    ("Score", "Description"),
    (0,    "No budget or funding source mentioned"),
    (0.50, "Budget/funding source mentioned (not ring-fenced)"),
    (1.00, "Ring-fenced budget or self-generating fund (dedicated fund, levy, tax)"),
    ("", ""),
    ("INSTRUMENT TYPE (Col P)", None),
    ("Score", "Description"),
    (0,    "No instrument"),
    (0.20, "Governance & coordination (coordinating body, planning requirement, role assignment)"),
    (0.40, "Information & voluntary (awareness campaigns, labelling, voluntary agreements)"),
    (0.60, "Economic (taxes, levies, deposit-refund schemes, subsidies)"),
    (0.80, "Infrastructure (waste management facility construction, recycling infrastructure)"),
    (1.00, "Regulatory (bans, mandatory standards, mandatory take-back, EPR obligations)"),
    ("", ""),
    ("INSTRUMENT IN FORCE (Col S)", None),
    ("Score", "Description"),
    (0, "Not in force — enabling power not yet exercised, or aspirational/contingency measure"),
    (1, "In force — operative and binding (shall/must/prohibited, or operationalised by subordinate instrument)"),
    ("", ""),
    ("INSTRUMENT IMPLEMENTATION (Col T — additive, max 1.0)", None),
    ("Sub-score", "Criterion (each must be explicitly evidenced)"),
    ("+0.25", "Responsible authority explicitly established or designated"),
    ("+0.25", "Enforcement (fines, penalties, enforcement measures)"),
    ("+0.25", "Monitoring mechanism (inspections, audits, data collection)"),
    ("+0.25", "Unconditional (no exemptions or loopholes)"),
    ("", ""),
    ("AUTO-CALCULATED FIELDS", None),
    ("Field", "Formula"),
    ("policy_score (O)", "= AVERAGE(G, I, K, M, P)  — averaged per instrument row"),
    ("instrument_score (V)", "= AVERAGE(P, T)"),
]

for r_idx, (a, b) in enumerate(ref_rows, start=1):
    ca = ws2.cell(row=r_idx, column=1, value=a)
    if b is None and a:   # section header
        ca.fill  = hfill(MID_BLUE)
        ca.font  = hfont(bold=True, size=10, color=WHITE)
        ws2.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)
        ca.alignment = CENTER_ALIGN
    elif b == "Description" or b == "Criterion (each must be explicitly evidenced)" or b == "Formula":
        ca.fill = hfill("C5D9F1")
        ca.font = hfont(bold=True, size=9)
        ws2.cell(row=r_idx, column=2, value=b).fill = hfill("C5D9F1")
        ws2.cell(row=r_idx, column=2).font = hfont(bold=True, size=9)
    else:
        ca.font = hfont(size=9)
        if b:
            cb = ws2.cell(row=r_idx, column=2, value=b)
            cb.font = hfont(size=9)
            cb.alignment = Alignment(wrap_text=True)
        if isinstance(a, (int, float)):
            ca.alignment = CENTER_ALIGN
            ca.number_format = "0.00"
    ca.border = thin_border()
    if b and not isinstance(b, type(None)):
        ws2.cell(row=r_idx, column=2).border = thin_border()

ws2.column_dimensions["A"].width = 14
ws2.column_dimensions["B"].width = 75
for r_idx in range(1, len(ref_rows) + 1):
    ws2.row_dimensions[r_idx].height = 20

# ═══════════════════════════════════════════════════════════════════════════════
# Save
# ═══════════════════════════════════════════════════════════════════════════════
path = "/workspace/4p_index_karnali_tourism_masterplan.xlsx"
wb.save(path)
print(f"Saved: {path}")
