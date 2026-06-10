import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "4P Index Coding"

# ── Colour palette ─────────────────────────────────────────────────────────────
DARK_BLUE   = "1F3864"
MID_BLUE    = "2E75B6"
TEAL        = "17375E"
LIGHT_BLUE  = "D9E2F3"
ORANGE      = "FCE4D6"
WHITE       = "FFFFFF"

def hfill(hex_): return PatternFill("solid", fgColor=hex_)
def hfont(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)
thin   = Side(style="thin",   color="BFBFBF")
medium = Side(style="medium", color="808080")
def thin_border():   return Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP_ALIGN   = Alignment(wrap_text=True, vertical="top",    horizontal="left")
CENTER_ALIGN = Alignment(wrap_text=True, vertical="center", horizontal="center")
CENTER_TOP   = Alignment(wrap_text=True, vertical="top",    horizontal="center")

# ── Row 1 title ────────────────────────────────────────────────────────────────
ws.merge_cells("A1:X1")
c = ws["A1"]
c.value = "Plastic Pollution Policy Index (4P Index) — Coding Table"
c.fill = hfill(DARK_BLUE); c.font = hfont(bold=True, size=13, color=WHITE)
c.alignment = CENTER_ALIGN; ws.row_dimensions[1].height = 22

# ── Row 2 metadata ─────────────────────────────────────────────────────────────
ws.merge_cells("A2:X2")
c = ws["A2"]
c.value = ("Policy: Health Care Waste Management Guideline (2071 BS / 2014)  |  "
           "Country: Nepal  |  Year: 2014  |  "
           "Source: http://climate.mohp.gov.np/downloads/Health_Care_Waste_Management_Guideline_2071.pdf")
c.fill = hfill(MID_BLUE); c.font = hfont(size=9, color=WHITE)
c.alignment = CENTER_ALIGN; ws.row_dimensions[2].height = 16

# ── Row 3 section labels ───────────────────────────────────────────────────────
ws.merge_cells("A3:O3")
c = ws["A3"]; c.value = "SECTION A — Policy-Level Fields  (repeated identically across all rows)"
c.fill = hfill(MID_BLUE); c.font = hfont(bold=True, size=10, color=WHITE); c.alignment = CENTER_ALIGN

ws.merge_cells("P3:X3")
c = ws["P3"]; c.value = "SECTION B — Instrument-Level Fields  (one row per instrument)"
c.fill = hfill(TEAL); c.font = hfont(bold=True, size=10, color=WHITE); c.alignment = CENTER_ALIGN
ws.row_dimensions[3].height = 18

# ── Row 4 column headers ───────────────────────────────────────────────────────
headers = [
    ("A","policy_name"),("B","policy_url"),("C","policy_year"),("D","policy_objective"),
    ("E","policy_target\n(0/1)"),("F","policy_target_text"),("G","policy_type\n(score)"),
    ("H","policy_type_justification"),("I","policy_integration\n(score)"),
    ("J","policy_sectors_list"),("K","policy_circularity\n(score)"),
    ("L","policy_lifecycle_phases_list"),("M","policy_budget\n(score)"),
    ("N","policy_budget_text"),("O","policy_score\n[auto]"),
    ("P","instrument_type\n(score)"),("Q","instrument_lifecycle_stage"),
    ("R","instrument_description"),("S","instrument_in_force\n(0/1)"),
    ("T","instrument_implementation\n(score)"),("U","instrument_implementation_text"),
    ("V","instrument_score\n[auto]"),("W","comments"),("X","—"),
]
for col_idx, (ltr, label) in enumerate(headers, start=1):
    cell = ws.cell(row=4, column=col_idx, value=f"Col {ltr}\n{label}")
    cell.fill = hfill("C5D9F1") if col_idx <= 15 else hfill("C4E1C0")
    cell.font = hfont(bold=True, size=9, color=DARK_BLUE if col_idx<=15 else "1F4E17")
    cell.alignment = CENTER_TOP; cell.border = thin_border()
ws.row_dimensions[4].height = 36

# ── Shared policy-level values ─────────────────────────────────────────────────
POL = dict(
    name   = "Health Care Waste Management Guideline",
    url    = "http://climate.mohp.gov.np/downloads/Health_Care_Waste_Management_Guideline_2071.pdf [URL partially visible in source image; full path inferred]",
    year   = 2014,
    obj    = ("To provide minimum standards for the safe and efficient management of health care "
              "waste (HCW) in Nepal's healthcare facilities; includes a prohibition on the incineration "
              "of PVC/halogenated plastics to prevent dioxin/furan emissions, mandatory waste segregation "
              "standards specifying plastic bottles and plastics as recyclable materials (dark blue "
              "containers), governance requirements for HCW Management Committees and plans, take-back "
              "policies for pharmaceutical/cytotoxic waste containing plastic items, and training for "
              "all HCF personnel."),
    tgt    = 0,
    tgt_t  = "",
    gtype  = 0.75,
    gtype_j= ("Administrative guideline issued in 2014 (2071 BS) by the Department of Health Services "
              "(DoHS), Ministry of Health and Population (MoHP), Government of Nepal. Functions as a "
              "sub-legislative standard-setting instrument operationalizing SWM Act 2011 (Section 43 on "
              "medical waste management licensing) and Environmental Protection Act 1997. Not enacted by "
              "Parliament. The document explicitly describes its provisions as 'minimum standards' for "
              "all HCFs and uses both mandatory ('shall', 'must') and advisory ('should') language."),
    intg   = 1,
    sects  = "healthcare, waste management, recycling, chemicals, municipalities, occupational health, pharmaceuticals",
    circ   = 0.75,
    lc     = "consumption, recycling, disposal, environmental leakage",
    budg   = 0.5,
    budg_t = ("Section 5.1.2 (p. 22): 'Ensure adequate financial and human resources for the "
              "implementation of HCWM plan (to support this, the authorized body can recommend the "
              "formulation of strategy to allocate certain percentage of total budget for HCWM).' "
              "No specific budget amount is ring-fenced for plastic waste management."),
)

# ── Instrument rows ────────────────────────────────────────────────────────────
INSTR_FILLS = ["E2EFDA","E2EFDA","FFF2CC","E2EFDA","D9F0FF"]

instruments = [
    {
        "P": 1.0,
        "Q": "Waste management",
        "R": ("Section 6.5h (pp. 38–39), Incineration: The guideline prohibits incineration of "
              "halogenated plastics including PVC, stating 'following wastes should never be "
              "incinerated: ... Halogenated plastics (e.g. PVC).' This applies to PVC-containing items "
              "such as blood bags and fluid bags used in HCFs. The prohibition is grounded in evidence "
              "that PVC incineration at low temperatures (<800°C) generates dioxins, furans and other "
              "toxic air pollutants (Section 4.2, p. 19)."),
        "S": 1,
        "T": 1.00,
        "U": ("Section 6.5h (p. 39): 'following wastes should never be incinerated: ... Halogenated "
              "plastics (e.g. PVC).' \n"
              "Section 4.2 (p. 19): 'when plastics that contain polyvinyl chloride (some plastics, "
              "some blood bags and fluid bags) are incinerated, dioxins, furans and other toxic air "
              "pollutants may be produced as emissions.'\n"
              "Responsible authority (+0.25): Section 5.1.2 designates HCWMC and Waste Management "
              "Officer; tables designate 'Chief of HCF, concerned health workers and the authorized "
              "person.'\n"
              "Enforcement (+0.25): Section 2.2.1 quotes SWM Act 2011 Section 39: 'the local body may "
              "impose a fine from fifty thousand to one hundred thousand rupees on anyone who commits "
              "offence as mentioned below: To throw, keep, discharge or cause to discharge chemical "
              "waste, industrial waste, medical waste or hazardous waste haphazardly.'\n"
              "Monitoring (+0.25): Section 5.3: 'Three types of monitoring mechanism need to be "
              "enforced... Baseline monitoring, Compliance monitoring, Impact monitoring.'\n"
              "Unconditional (+0.25): No exemptions stated for halogenated plastic incineration ban."),
        "W": ("Language note: 'should never be incinerated' uses advisory 'should' rather than "
              "mandatory 'shall/must.' However, in the context of this minimum-standards guideline, "
              "'should never' is treated as a prohibitory standard (S = 1). If the stricter language "
              "test is applied (only 'shall/must/prohibited' = in force), then S = 0. Enforcement "
              "(+0.25) is based on SWM Act 2011 Section 39 penalties explicitly quoted in Section 2.2.1 "
              "of this guideline — not a provision of the guideline itself, which is a borderline case. "
              "Cross-reference: Stockholm Convention on POPs (referenced in Section 6, p. 26) "
              "similarly targets dioxin/furan emissions from incineration."),
    },
    {
        "P": 1.0,
        "Q": "Waste management",
        "R": ("Section 6.2 (pp. 28–30) and Table 5 (pp. 29–30), Waste Segregation: Establishes a "
              "mandatory color-coded waste segregation system for all HCFs. Recyclable non-risk HCW "
              "(dark blue containers) specifically includes 'plastic bottles, cans, metals, glass, "
              "plastics, papers, rubber etc.' Segregation must occur at the point of waste generation "
              "(ward, bedside, laboratory). Containers must be rigid and leak-proof; sharps containers "
              "must be 'puncture proof made of either metal or high density plastic.' Segregation must "
              "be 'regularly monitored to ensure that the procedures are followed strictly.'"),
        "S": 1,
        "T": 0.75,
        "U": ("Section 6.2 (p. 28–29): 'Segregation should: Always take place at the source... Be "
              "regularly monitored to ensure that the procedures are followed strictly.'\n"
              "Table 5 (p. 29): 'Non-risk HCW Recyclable [Dark Blue]: Non-biodegradable, which can be "
              "recycled: plastic bottles, cans, metals, glass, plastics, papers, rubber etc.'\n"
              "Section 6.2 (p. 28): container 'must be impermeable to fluids,' 'must be leak-proof.'\n"
              "Responsible authority (+0.25): Section 5.1.2 designates HCWMC, Waste Management Officer.\n"
              "Enforcement (+0.25): SWM Act 2011 Section 39 penalties quoted in Section 2.2.1.\n"
              "Monitoring (+0.25): Section 5.3 compliance monitoring framework.\n"
              "Unconditional (−0): Table 5 note: 'If the container with the recommended color is not "
              "available, any colored container can be used to segregate wastes with proper labeling.' "
              "Section 3.3 allows minimum categorization into only two categories (non-risk vs. risk) "
              "rather than full multi-category system — explicit flexibility/exemption exists."),
        "W": ("The 'unconditional' sub-score (−0.25) is not awarded because Table 5 explicitly "
              "provides an alternative when recommended color containers are not available, and "
              "Section 3.3 allows HCFs to use a simplified two-category minimum system instead of the "
              "full classification. These constitute explicit exemptions that reduce T to 0.75. "
              "The guideline uses mixed mandatory language: some provisions use 'must' (containers must "
              "be leak-proof) while segregation procedures predominantly use 'should.' Both warrant "
              "S = 1 given the minimum-standards framing of the document."),
    },
    {
        "P": 0.20,
        "Q": "Waste management",
        "R": ("Sections 5.1.1–5.1.3 (pp. 21–23), Organizational Issues: Every HCF must establish a "
              "Health Care Waste Management Committee (HCWMC) and develop a written HCWM Plan. "
              "Section 5.1.3: 'Every HCF must develop its HCWM Plan.' HCWMC must include the Chief/"
              "Director, Department Heads, Matron, Waste Management Officer and support staff "
              "representative. HCWMC must implement, review, and update the plan annually. A designated "
              "Waste Management Officer must be appointed to supervise and coordinate the plan."),
        "S": 1,
        "T": 1.00,
        "U": ("Section 5.1.3 (p. 21): 'Every HCF must develop its HCWM Plan.'\n"
              "Section 5.1.1 (p. 22): 'Designate a waste management officer to supervise and "
              "coordinate the waste management plan.'\n"
              "Section 5.1.2 (p. 22): membership requirements of HCWMC explicitly specified.\n"
              "Responsible authority (+0.25): HCWMC and Waste Management Officer explicitly designated.\n"
              "Enforcement (+0.25): Section 2.2.1 quotes SWM Act 2011 Section 43: 'the authority that "
              "grants license to establish a health institution as per the prevalent law shall, before "
              "granting license for establishment and operation of the health institution, confirm "
              "whether appropriate management has been made for solid waste management or not and it "
              "shall have to grant license only if appropriate arrangement is made.'\n"
              "Monitoring (+0.25): Section 5.3: 'Regular monitoring and evaluation of the plan in "
              "each HCF should be performed.' Three-level monitoring framework (baseline, compliance, "
              "impact) mandated.\n"
              "Unconditional (+0.25): Requirement for HCWMC and HCWM Plan applies to all HCFs; no "
              "blanket exemptions stated (simplified committee composition for smaller HCFs does not "
              "exempt them from the requirement)."),
        "W": ("Mixed language: 'Every HCF must develop its HCWM Plan' is mandatory, while the "
              "committee requirement uses 'should have HCWMC.' The licensing enforcement (+0.25) is "
              "from SWM Act 2011 Section 43 quoted in Section 2.2.1 of this guideline, constituting "
              "a robust enforcement mechanism — healthcare licenses can be withheld or conditioned "
              "on compliance with waste management standards. Section 5.3 monitoring uses 'should be "
              "performed' rather than 'shall' — monitoring sub-score (+0.25) is awarded as it is "
              "explicitly described as a required activity within the guideline's framework. "
              "The committee requirement covers all HCW including plastic waste (criterion c)."),
    },
    {
        "P": 1.0,
        "Q": "End of life",
        "R": ("Section 7.9 (p. 43) and Tables 6–13 (pp. 45–54), Take-back/Return Policy for "
              "Pharmaceutical and Cytotoxic Waste: The guideline mandates a return-back policy for "
              "date-expired, contaminated and cytotoxic pharmaceutical waste. Tables 6–13 state "
              "'Apply return back policy; return the waste to the store and from the store to the "
              "supplier.' This applies to cytotoxic pharmaceuticals (including plastic-containing "
              "delivery items: syringes, vials) and to radioactive waste items including 'paper cups, "
              "straws, needles, syringes, test tubes' (many of which are plastic)."),
        "S": 1,
        "T": 0.75,
        "U": ("Tables 6–13 (pp. 45–54), repeated for each facility level: 'Pharmaceutical waste such "
              "as waste comprising of date expired, contaminated and discarded medicines — Apply "
              "return back policy; return the waste to the store and from the store to the supplier.'\n"
              "'Cytotoxic pharmaceutical waste — Apply return back policy; return the waste to the "
              "store and from the store to the supplier.'\n"
              "Section 7.9 (p. 43): 'Return of date expired drugs to suppliers. Take back policy "
              "should be applied for these kinds of materials. Agreement should be signed while "
              "purchasing the cytotoxic and radioactive materials and these materials should be "
              "collected back by the suppliers after usage.'\n"
              "Responsible authority (+0.25): Chief of HCF, Waste Management Officer designated.\n"
              "Enforcement (+0.25): SWM Act 2011 Section 39 penalties for improper disposal.\n"
              "Monitoring (+0.25): Section 5.3 compliance monitoring applies.\n"
              "Unconditional (−0): Each table entry provides an 'OR' alternative (secure landfill "
              "after encapsulation), indicating the return-back policy is not the sole mandated "
              "approach — alternatives exist."),
        "W": ("The take-back policy passes the plastics relevance filter under criterion (b) as an "
              "EPR-like instrument, and partially under (a) since it applies to plastic-containing "
              "items (syringes, test tubes, cytotoxic waste containers). Tables use direct imperative "
              "language ('Apply return back policy') for both pharmaceutical and cytotoxic waste at all "
              "facility levels, supporting S = 1. However, the 'OR' alternative disposal option in "
              "every table entry means the instrument is conditional, reducing T to 0.75 (no "
              "'unconditional' sub-score). This is an informal EPR mechanism not formally designated "
              "as Extended Producer Responsibility; a case could be made for P = 0.60 (economic) "
              "if the take-back arrangement is viewed as a voluntary contractual mechanism rather than "
              "a binding regulatory obligation."),
    },
    {
        "P": 0.40,
        "Q": "Waste management",
        "R": ("Chapter 9 (pp. 56–57), Training and Raising Awareness: The guideline requires "
              "mandatory training programs for all HCF personnel on HCWM, including waste management "
              "of plastics (segregation into recyclable category, no incineration of PVC). Training "
              "must cover: concept of HCW and basic management steps, legal provisions, impact of HCW "
              "on health and environment, and practical waste management at different HCF levels. "
              "Section 7: 'no training/no hiring policy should be instituted.' "
              "Training responsibility designated to the administration/management section of each HCF."),
        "S": 0,
        "T": 0.75,
        "U": ("Section 9.3 (p. 57): 'The administration/management section should be given "
              "responsibility for all training related to HCWM.'\n"
              "Section 7 (p. 40): 'no training/no hiring policy should be instituted; immunization "
              "at the first day of the work.'\n"
              "Section 9.3 (p. 57): 'The record of all training sessions and the content of training "
              "programs should be documented.'\n"
              "Responsible authority (+0.25): administration/management section designated.\n"
              "Enforcement (−0): No explicit penalties for failing to conduct training stated in "
              "guideline.\n"
              "Monitoring (+0.25): Documentation requirement: 'The record of all training sessions "
              "and the content of training programs should be documented.'\n"
              "Unconditional (+0.25): Training requirement applies to all HCF personnel with no stated "
              "exemptions."),
        "W": ("instrument_in_force = 0: Training obligations use advisory language ('should be "
              "convinced', 'should be given responsibility', 'should be instituted') rather than "
              "mandatory 'shall/must.' 'No training/no hiring policy' is a strong recommendation but "
              "uses 'should.' If training provisions are viewed as mandatory minimum standards (given "
              "the guideline's overall framing), S = 1 could be justified — noted as borderline. "
              "The training program covers plastic waste management as part of general HCWM curriculum "
              "(criterion c), including awareness of PVC incineration risks, plastic segregation in "
              "color-coded containers, and recycling of non-contaminated plastics."),
    },
]

instr_labels = [
    "Instrument 1 — Prohibition on incineration of PVC/halogenated plastics",
    "Instrument 2 — Mandatory waste segregation standard with color-coded containers (including plastic waste)",
    "Instrument 3 — HCWMC and HCWM Plan requirement (governance)",
    "Instrument 4 — Take-back/return policy for pharmaceutical and cytotoxic waste",
    "Instrument 5 — Training and awareness program for HCWM",
]

for i, (instr, label) in enumerate(zip(instruments, instr_labels), start=1):
    r = 4 + i   # rows 5-9

    # Col X — row label
    c = ws.cell(row=r, column=24, value=label)
    c.font = hfont(italic=True, size=8, color="808080")

    # Policy columns A–O
    pol_fill = hfill(LIGHT_BLUE)
    auto_fill = hfill(ORANGE)
    pol_data = [
        POL["name"], POL["url"], POL["year"], POL["obj"], POL["tgt"],
        POL["tgt_t"], POL["gtype"], POL["gtype_j"], POL["intg"], POL["sects"],
        POL["circ"], POL["lc"], POL["budg"], POL["budg_t"], None,
    ]
    for col_idx, val in enumerate(pol_data, start=1):
        is_auto = (col_idx == 15)
        cell = ws.cell(row=r, column=col_idx)
        if is_auto:
            cell.value  = f"=AVERAGE(G{r},I{r},K{r},M{r},P{r})"
            cell.fill   = auto_fill
            cell.font   = hfont(italic=True, size=9, color="7F3F00")
            cell.number_format = "0.00"
        else:
            cell.value  = val
            cell.fill   = pol_fill
            cell.font   = hfont(size=9)
            if col_idx == 3:
                cell.alignment  = CENTER_ALIGN
                cell.number_format = "0"
            elif col_idx in (5, 7, 9, 11, 13):
                cell.alignment  = CENTER_ALIGN
                cell.number_format = "0.00"
            else:
                cell.alignment  = WRAP_ALIGN
        cell.border = thin_border()

    # Instrument columns P–W
    instr_fill = hfill(INSTR_FILLS[i - 1])
    instr_data = [
        instr["P"], instr["Q"], instr["R"], instr["S"],
        instr["T"], instr["U"], None, instr["W"],
    ]
    for j, val in enumerate(instr_data):
        col_idx = 16 + j
        is_auto = (col_idx == 22)
        cell = ws.cell(row=r, column=col_idx)
        if is_auto:
            cell.value  = f"=AVERAGE(P{r},T{r})"
            cell.fill   = auto_fill
            cell.font   = hfont(italic=True, size=9, color="7F3F00")
            cell.number_format = "0.000"
        else:
            cell.value  = val
            cell.fill   = instr_fill
            cell.font   = hfont(size=9)
            if col_idx in (16, 18, 19, 20):
                cell.alignment  = CENTER_ALIGN
                cell.number_format = "0.00" if col_idx in (16, 20) else "0"
            else:
                cell.alignment  = WRAP_ALIGN
        cell.border = thin_border()

    ws.row_dimensions[r].height = 220

# ── Column widths ──────────────────────────────────────────────────────────────
col_widths = {
    1:28, 2:32, 3:8, 4:35, 5:8, 6:20, 7:8, 8:35, 9:9,
    10:35, 11:8, 12:28, 13:8, 14:40, 15:9,
    16:10, 17:16, 18:42, 19:10, 20:10, 21:42, 22:9, 23:42, 24:22,
}
for col, width in col_widths.items():
    ws.column_dimensions[get_column_letter(col)].width = width

ws.freeze_panes = "D5"

# ══════════════════════════════════════════════════════════════════════════════
# Sheet 2 — Key Evidence
# ══════════════════════════════════════════════════════════════════════════════
ws2 = wb.create_sheet("Key Evidence & Score Summary")

evidence = [
    ("SCORE SUMMARY", None),
    ("Instrument", "P (type)", "S (in force)", "T (impl.)", "V (score [auto])"),
    ("1 — PVC/halogenated plastic incineration prohibition",         "1.00", "1", "1.00", "1.000"),
    ("2 — Waste segregation standard (incl. plastic waste)",         "1.00", "1", "0.75", "0.875"),
    ("3 — HCWMC and HCWM Plan requirement",                          "0.20", "1", "1.00", "0.600"),
    ("4 — Take-back policy for pharmaceutical/cytotoxic waste",      "1.00", "1", "0.75", "0.875"),
    ("5 — Training and awareness program",                           "0.40", "0", "0.75", "0.575"),
    ("", "", "", "", ""),
    ("POLICY-LEVEL SCORES", None),
    ("Field", "Value"),
    ("policy_type (G)",        "0.75"),
    ("policy_integration (I)", "1.00  (7 sectors: healthcare, waste management, recycling, chemicals, municipalities, occupational health, pharmaceuticals)"),
    ("policy_circularity (K)", "0.75  (4 lifecycle phases: consumption, recycling, disposal, environmental leakage)"),
    ("policy_budget (M)",      "0.50  (budget for HCWM mentioned but not ring-fenced for plastics)"),
    ("policy_score (O)",       "[auto] = AVERAGE(G, I, K, M, P per row)"),
    ("", ""),
    ("KEY VERBATIM EVIDENCE — PLASTIC-RELEVANT PASSAGES", None),
    ("Location", "Verbatim text"),
    ("Sec 6.5h (p.39) — PVC ban",
     "'following wastes should never be incinerated: ... Halogenated plastics (e.g. PVC)'"),
    ("Sec 4.2 (p.19) — PVC context",
     "'when plastics that contain polyvinyl chloride (some plastics, some blood bags and fluid bags) "
     "are incinerated, dioxins, furans and other toxic air pollutants may be produced as emissions'"),
    ("Table 5 (p.29) — Recyclable plastics",
     "Dark Blue container: 'Non-biodegradable, which can be recycled: plastic bottles, cans, metals, "
     "glass, plastics, papers, rubber etc.'"),
    ("Sec 6.1.2 (p.27) — Recyclable materials",
     "'Some of the materials which can be recycled are given below: Glass, Plastics, Aluminium cans, "
     "Paper and card board, Iron'"),
    ("Sec 6.2 (p.28) — Sharps containers",
     "'Container should be puncture proof made of either metal or high density plastic and fitted with covers'"),
    ("Tables 6-13 — Take-back (imperative)",
     "'Cytotoxic pharmaceutical waste — Apply return back policy; return the waste to the store and "
     "from the store to the supplier.'"),
    ("Sec 5.1.3 (p.21) — Plan mandatory",
     "'Every HCF must develop its HCWM Plan.'"),
    ("SWM Act 2011 Sec.43 (quoted p.7-8)",
     "'the authority that grants license to establish a health institution ... shall ... confirm whether "
     "appropriate management has been made for solid waste management or not and it shall have to grant "
     "license only if appropriate arrangement is made.'"),
    ("Sec 5.1.2 (p.22) — Budget",
     "'Ensure adequate financial and human resources for the implementation of HCWM plan (to support "
     "this, the authorized body can recommend the formulation of strategy to allocate certain percentage "
     "of total budget for HCWM).'"),
]

for r_idx, row in enumerate(evidence, start=1):
    if row[1] is None and row[0]:   # section header
        ws2.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=5)
        c = ws2.cell(row=r_idx, column=1, value=row[0])
        c.fill = hfill(MID_BLUE); c.font = hfont(bold=True, size=10, color=WHITE)
        c.alignment = CENTER_ALIGN
    elif row[1] in ("P (type)", "Value", "Verbatim text"):  # sub-header row
        for ci, v in enumerate(row, start=1):
            if v:
                c = ws2.cell(row=r_idx, column=ci, value=v)
                c.fill = hfill("C5D9F1"); c.font = hfont(bold=True, size=9)
                c.alignment = Alignment(wrap_text=True)
                c.border = thin_border()
    elif row[0] == "":
        pass
    else:
        for ci, v in enumerate(row, start=1):
            if v is not None:
                c = ws2.cell(row=r_idx, column=ci, value=v)
                c.font = hfont(size=9)
                c.alignment = Alignment(wrap_text=True)
                c.border = thin_border()
    ws2.row_dimensions[r_idx].height = 30

ws2.column_dimensions["A"].width = 45
ws2.column_dimensions["B"].width = 12
ws2.column_dimensions["C"].width = 12
ws2.column_dimensions["D"].width = 12
ws2.column_dimensions["E"].width = 12

# ══════════════════════════════════════════════════════════════════════════════
# Sheet 3 — Score Reference (reuse from previous script pattern)
# ══════════════════════════════════════════════════════════════════════════════
ws3 = wb.create_sheet("Score Reference")
ref_rows = [
    ("POLICY TYPE (Col G)", None),("Score","Description"),
    (0.25,"Strategy/plan — aspirational commitments without concrete targets"),
    (0.50,"Strategy/plan — incorporates quantifiable targets"),
    (0.75,"Regulation or executive decree (sub-legislative, executive branch)"),
    (1.00,"Legislation enacted by parliament or equivalent legislative body"),
    ("",""),("POLICY INTEGRATION (Col I)", None),("Score","Description"),
    (0,"0 sectors"),(0.25,"1–2 sectors"),(0.50,"3–4 sectors"),
    (0.75,"5–6 sectors"),(1.00,"7 or more sectors"),
    ("",""),("POLICY CIRCULARITY (Col K)", None),("Score","Description"),
    (0.25,"1 lifecycle phase"),(0.50,"2 lifecycle phases"),
    (0.75,"3–4 lifecycle phases"),(1.00,"All 5 phases"),
    ("",""),("POLICY BUDGET (Col M)", None),("Score","Description"),
    (0,"No budget"),(0.50,"Budget/funding mentioned (not ring-fenced)"),
    (1.00,"Ring-fenced or self-generating fund"),
    ("",""),("INSTRUMENT TYPE (Col P)", None),("Score","Description"),
    (0,"No instrument"),(0.20,"Governance & coordination"),
    (0.40,"Information & voluntary"),(0.60,"Economic"),
    (0.80,"Infrastructure"),(1.00,"Regulatory"),
    ("",""),("INSTRUMENT IN FORCE (Col S)", None),("Score","Description"),
    (0,"Not in force — enabling power not yet exercised or aspirational"),
    (1,"In force — mandatory language (shall/must/prohibited) or operationalised"),
    ("",""),("INSTRUMENT IMPLEMENTATION (Col T — additive, max 1.0)", None),
    ("Sub-score","Criterion (each must be explicitly evidenced)"),
    ("+0.25","Responsible authority explicitly established or designated"),
    ("+0.25","Enforcement (fines, penalties, enforcement measures)"),
    ("+0.25","Monitoring mechanism (inspections, audits, data collection)"),
    ("+0.25","Unconditional (no exemptions or loopholes)"),
]
for r_idx, row_data in enumerate(ref_rows, start=1):
    a, b = row_data[0], row_data[1] if len(row_data) > 1 else None
    ca = ws3.cell(row=r_idx, column=1, value=a)
    if b is None and a:
        ws3.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=2)
        ca.fill = hfill(MID_BLUE); ca.font = hfont(bold=True, size=10, color=WHITE)
        ca.alignment = CENTER_ALIGN
    elif b in ("Description", "Criterion (each must be explicitly evidenced)"):
        ca.fill = hfill("C5D9F1"); ca.font = hfont(bold=True, size=9)
        ws3.cell(row=r_idx, column=2, value=b).fill = hfill("C5D9F1")
        ws3.cell(row=r_idx, column=2).font = hfont(bold=True, size=9)
    else:
        ca.font = hfont(size=9)
        if b: ws3.cell(row=r_idx, column=2, value=b).font = hfont(size=9)
        if isinstance(a, (int, float)): ca.alignment = CENTER_ALIGN; ca.number_format = "0.00"
    ca.border = thin_border()
    if b: ws3.cell(row=r_idx, column=2).border = thin_border()
    ws3.row_dimensions[r_idx].height = 18
ws3.column_dimensions["A"].width = 14
ws3.column_dimensions["B"].width = 70

path = "/workspace/4p_index_hcwm_guideline_2014.xlsx"
wb.save(path)
print(f"Saved: {path}")
