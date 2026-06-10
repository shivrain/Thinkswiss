import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "4P Index Coding"

DARK_BLUE="1F3864"; MID_BLUE="2E75B6"; TEAL="17375E"
LIGHT_BLUE="D9E2F3"; ORANGE="FCE4D6"; WHITE="FFFFFF"

def hfill(h): return PatternFill("solid", fgColor=h)
def hfont(bold=False,size=10,color="000000",italic=False):
    return Font(bold=bold,size=size,color=color,italic=italic)
thin=Side(style="thin",color="BFBFBF")
def tb(): return Border(left=thin,right=thin,top=thin,bottom=thin)
WA=Alignment(wrap_text=True,vertical="top",horizontal="left")
CA=Alignment(wrap_text=True,vertical="center",horizontal="center")
CT=Alignment(wrap_text=True,vertical="top",horizontal="center")

# Row 1 title
ws.merge_cells("A1:X1")
c=ws["A1"]; c.value="Plastic Pollution Policy Index (4P Index) — Coding Table"
c.fill=hfill(DARK_BLUE); c.font=hfont(bold=True,size=13,color=WHITE); c.alignment=CA
ws.row_dimensions[1].height=22

# Row 2 metadata
ws.merge_cells("A2:X2")
c=ws["A2"]
c.value=("Policy: The Sixteenth Plan (Fiscal Year 2024/25–2028/29)  |  Country: Nepal  |  Year: 2024  |  "
         "Source: http://elibrary.moest.gov.np/bitstream/... [URL truncated in source image]")
c.fill=hfill(MID_BLUE); c.font=hfont(size=9,color=WHITE); c.alignment=CA
ws.row_dimensions[2].height=16

# Row 3 section labels
ws.merge_cells("A3:O3")
c=ws["A3"]; c.value="SECTION A — Policy-Level Fields  (repeated identically across all rows)"
c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
ws.merge_cells("P3:X3")
c=ws["P3"]; c.value="SECTION B — Instrument-Level Fields  (one row per instrument)"
c.fill=hfill(TEAL); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
ws.row_dimensions[3].height=18

# Row 4 headers
headers=[("A","policy_name"),("B","policy_url"),("C","policy_year"),("D","policy_objective"),
         ("E","policy_target\n(0/1)"),("F","policy_target_text"),("G","policy_type\n(score)"),
         ("H","policy_type_justification"),("I","policy_integration\n(score)"),("J","policy_sectors_list"),
         ("K","policy_circularity\n(score)"),("L","policy_lifecycle_phases_list"),("M","policy_budget\n(score)"),
         ("N","policy_budget_text"),("O","policy_score\n[auto]"),
         ("P","instrument_type\n(score)"),("Q","instrument_lifecycle_stage"),("R","instrument_description"),
         ("S","instrument_in_force\n(0/1)"),("T","instrument_implementation\n(score)"),
         ("U","instrument_implementation_text"),("V","instrument_score\n[auto]"),("W","comments"),("X","—")]
for ci,(l,lab) in enumerate(headers,1):
    cell=ws.cell(row=4,column=ci,value=f"Col {l}\n{lab}")
    cell.fill=hfill("C5D9F1") if ci<=15 else hfill("C4E1C0")
    cell.font=hfont(bold=True,size=9,color=DARK_BLUE if ci<=15 else "1F4E17")
    cell.alignment=CT; cell.border=tb()
ws.row_dimensions[4].height=36

# Shared policy data
POL=dict(
 name="The Sixteenth Plan (Fiscal Year 2024/25–2028/29)",
 url=("http://elibrary.moest.gov.np/bitstream/... [URL truncated in source image; "
      "document published by Government of Nepal, National Planning Commission, May 2024]"),
 year=2024,
 obj=("Nepal's national five-year development plan for FY 2024/25–2028/29, setting targets for: "
      "discouraging the use of plastic materials and prohibiting burning plastics; controlling plastic items; "
      "displacing plastic products with green raw materials; imposing a landfill waste tax; constructing 17 "
      "integrated solid waste management centers and sanitary landfill sites in cities; implementing extended "
      "producer responsibility for producers and importers of industrial, hospital and hazardous waste; and "
      "prohibiting dumping of wastes in rivers and open burning in forests."),
 tgt=1,
 tgt_t=("Section 7.6 Quantitative Targets (p. 134): 'Modern integrated solid waste management centers and "
        "sanitary landfill sites — FY 2022/23 actual: 1; FY 2028/29 target: 17.'"),
 gtype=0.50,
 gtype_j=("National five-year development plan formulated by the National Planning Commission of Nepal and "
          "published in May 2024. Issued under executive authority as a planning document — not enacted by "
          "Parliament. Contains quantifiable targets including waste management infrastructure targets (17 "
          "integrated solid waste management centers and landfill sites by FY 2028/29). Scored 0.50 "
          "(strategy/plan incorporating quantifiable targets), not 0.75 (regulation/decree) as it is a "
          "programmatic plan, not a binding sub-legislative regulation."),
 intg=1,
 sects="waste management, urbanization, municipalities, industry, environment, water, energy, agriculture, transport",
 circ=0.75,
 lc="consumption, recycling, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Table 2.6 (p. 44): Water supply, sewerage, waste management and recycling activities — total plan "
         "investment NPR 3,991 million (Public: NPR 2,091 million / 52.4%; Private: NPR 1,852 million / 46.4%; "
         "Cooperative: NPR 48 million / 1.2%). Total plan investment across all sectors (at FY 2023/24 prices): "
         "approx. NPR 11,481 billion, divided public sector 30.2%, private 67.2%, cooperative 2.6%. Budget "
         "exists for waste management and recycling sector but is not ring-fenced specifically for plastic."),
)

FILLS=["E2EFDA","FFF2CC","E2EFDA","D9F0FF","E2EFDA"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Chapter 7, Section 7.5, Item 10 (p. 133) — Integrated Urban Waste Management Program: "
     "'Discourage the use of plastic materials and prohibit burning plastics.' "
     "Chapter 13, Section 13.5, Item 4 (p. 213) — Pollution Control Program: 'Control plastic items.' "
     "Chapter 13, Section 13.5, Item 1 (p. 212) — Green Economy Program: 'Increase the use of products "
     "based on green raw materials, displacing plastic products.' "
     "Three distinct plastic-specific commitments: (a) discourage use; (b) prohibit burning; "
     "(c) control/displace plastic items."),
  S=0,
  T=0.50,
  U=("Ch. 7, Sec. 7.5.10 (p. 133): 'Discourage the use of plastic materials and prohibit burning plastics.'\n"
     "Ch. 13, Sec. 13.5.4 (p. 213): 'Control plastic items.'\n"
     "Ch. 13, Sec. 13.5.1 (p. 212): 'Increase the use of products based on green raw materials, "
     "displacing plastic products.'\n"
     "Responsible authority (+0.25): Appendix 1 designates Ministry of Urban Development (Ch.7 measures) "
     "and Ministry of Forests and Environment (Ch.13 measures) as implementing agencies.\n"
     "Enforcement (−0): No explicit penalties for plastic use or burning stated in the plan text.\n"
     "Monitoring (+0.25): Section 13.5.8: 'Expand measurement centers for climate, pollution, etc.' and "
     "Appendix 1 monitoring framework (Federal Parliament, NPC, OPMCM).\n"
     "Unconditional (−0): 'Discourage' is aspirational; 'control plastic items' is open-ended; "
     "implementation flexibility exists."),
  W=("Three plastic-specific commitments combined: (a) regulatory prohibition on burning (strongest), "
     "(b) administrative control of plastic items, and (c) information/economic displacement of plastic "
     "products. Highest type (regulatory, 1.0) applied per coding rules; dual nature noted. "
     "instrument_in_force = 0: plan-level commitments require subsequent implementing regulation — "
     "'prohibit' in a national plan is not the same as an operative legal prohibition. "
     "The SWM Act 2011 and Environmental Protection Act 2019 provide existing legal backing for some "
     "of these measures, but the plan commits to new/strengthened measures. "
     "Q = Consumption selected as primary lifecycle stage since 'discourage use' and 'displace plastic "
     "products' target consumption; burning prohibition targets Waste management/End of life — noted here."),
 ),
 dict(
  P=0.80, Q="Waste management",
  R=("Chapter 7, Section 7.5, Items 9–10 (pp. 132–133) and Section 7.6 Target 8 (p. 134): "
     "Commits to developing waste treatment systems including waste treatment centers and sanitary "
     "landfill sites in all cities, with a target to increase from 1 to 17 modern integrated solid "
     "waste management centers by FY 2028/29. Also requires making 'arrangements for collecting "
     "biodegradable, non-biodegradable and recyclable wastes separately through local levels by "
     "developing separate standards for management of these wastes.' "
     "Section 13.5.4: 'Make waste management cost effective and sustainable by involving the private "
     "sector; utilize waste reduction, reuse and recycling.'"),
  S=0,
  T=0.50,
  U=("Sec. 7.6 Target 8 (p. 134): 'Modern integrated solid waste management centers and sanitary "
     "landfill sites — FY 2022/23 actual: 1; FY 2028/29 target: 17.'\n"
     "Sec. 7.5 Item 10 (p. 133): 'Develop waste treatment system, including management of waste "
     "treatment centers and landfill sites, in all cities. Make arrangements for collecting "
     "biodegradable, non-biodegradable and recyclable wastes separately...'\n"
     "Responsible authority (+0.25): Appendix 1 designates Ministry of Urban Development and "
     "Ministry of Physical Infrastructure and Transport as implementing agencies.\n"
     "Enforcement (−0): No explicit penalties for non-compliance with waste facility targets.\n"
     "Monitoring (+0.25): Appendix 1 monitoring framework; semi-annual and annual progress reviews.\n"
     "Unconditional (−0): 'In all cities' but implementation dependent on local capacity and resource "
     "availability — flexibility implied."),
  W=("This instrument combines infrastructure (construction of 17 centers) and governance (separate "
     "waste collection standards). Highest type (Infrastructure, 0.80) applied. The quantitative target "
     "(1 → 17 centers) is the basis for policy_target = 1 at the policy level. The target directly "
     "covers plastic waste streams through 'non-biodegradable and recyclable waste' collection. "
     "Cross-reference: Chapter 13.4 strategy: 'Make waste management cost effective through "
     "public-private partnerships' — PPP modality for waste facilities is an additional governance element. "
     "S = 0: The plan commits to building these centers but the commitment is not an operative regulation; "
     "annual budgets and implementing regulations are needed."),
 ),
 dict(
  P=1.0, Q="End of life",
  R=("Chapter 13, Section 13.4, Transformative Strategy 3 (p. 224) — Controlling Pollution for a "
     "Healthy Society: 'Make producers or importers and polluters responsible for industrial, hospital "
     "and hazardous wastes management, as well as implement the concept of extended producer "
     "responsibility and cut down industrial emissions.' Also: 'Enforce the polluter pays principle.' "
     "This establishes EPR as a national policy priority applicable to producers and importers of "
     "plastic-containing industrial, hospital and hazardous waste products."),
  S=0,
  T=0.50,
  U=("Sec. 13.4, Strategy 3 (p. 224): 'Make producers or importers and polluters responsible for "
     "industrial, hospital and hazardous wastes management, as well as implement the concept of "
     "extended producer responsibility and cut down industrial emissions.'\n"
     "'Enforce the polluter pays principle.'\n"
     "Responsible authority (+0.25): Appendix 1 designates Ministry of Forests and Environment and "
     "Ministry of Finance as implementing agencies for Ch.13 programs.\n"
     "Enforcement (−0): No specific penalties for EPR non-compliance stated in the plan.\n"
     "Monitoring (+0.25): Appendix 1 quarterly/semi-annual/annual progress review; NPC oversight.\n"
     "Unconditional (−0): EPR concept to be 'implemented' — future action without defined scope or "
     "sector-specific exemptions; breadth of coverage is unclear."),
  W=("EPR is committed to as a 'concept' to be implemented, indicating this is an aspirational strategy "
     "rather than an operative EPR scheme. S = 0: no subordinate EPR regulation has been enacted under "
     "this plan yet. Passes plastics relevance filter under criterion (b) as it 'directly enables plastic-"
     "relevant policy actions — extended producer responsibility.' If a subsequent EPR regulation is "
     "enacted, S should be updated to 1. The 'polluter pays principle' enforcement commitment provides "
     "additional framing. Type coded as 1.0 (Regulatory/EPR obligation) because EPR when operationalised "
     "is a regulatory instrument; however, a case can be made for 0.20 (governance, planning requirement) "
     "since this is only a planning commitment at present."),
 ),
 dict(
  P=0.60, Q="Waste management",
  R=("Chapter 13, Section 13.5, Item 4 (p. 213) — Pollution Control Program for Healthy Society: "
     "'Impose tax on wastes disposed of in landfills.' This is an economic instrument (landfill charge/"
     "tax) that would apply to all solid waste deposited in landfills, including plastic waste, creating "
     "a financial incentive to reduce waste generation and increase recycling/diversion from landfill."),
  S=0,
  T=0.75,
  U=("Sec. 13.5, Item 4 (p. 213): 'Impose tax on wastes disposed of in landfills.'\n"
     "Responsible authority (+0.25): Ministry of Finance and Ministry of Forests and Environment "
     "(Appendix 1 for Ch.13 programs) designated as implementing agencies.\n"
     "Enforcement (−0): No explicit enforcement mechanism or penalty for non-payment of landfill tax "
     "stated in the plan; tax collection details not specified.\n"
     "Monitoring (+0.25): Appendix 1 monitoring framework; revenue collection would be monitored through "
     "fiscal management systems.\n"
     "Unconditional (+0.25): No exemptions or loopholes explicitly stated for the landfill tax in the "
     "plan text."),
  W=("The landfill tax is a plan-level commitment — 'Impose tax' uses future-oriented imperative "
     "language but has not been enacted as a law/regulation yet. S = 0: requires legislative or "
     "regulatory action to implement. T = 0.75 because responsible authority, monitoring, and "
     "unconditional sub-scores are met; enforcement is not explicitly detailed. "
     "The landfill tax directly creates incentives to divert plastic waste from landfills and thereby "
     "passes the plastics relevance filter under criterion (c) (landfill charges). "
     "Revenue from this tax could fund waste management infrastructure, adding a self-financing element."),
 ),
 dict(
  P=1.0, Q="Environmental leakage",
  R=("Chapter 13, Section 13.5, Item 4 (p. 213) — Pollution Control Program: 'Prohibit dumping "
     "wastes in streams and rivers or burning in forests and open areas.' This directly targets "
     "environmental leakage of waste, including plastic waste, into aquatic and terrestrial environments. "
     "It also reinforces the plastic burning prohibition from Chapter 7 in the context of environmental "
     "protection and pollution control."),
  S=0,
  T=0.75,
  U=("Sec. 13.5, Item 4 (p. 213): 'Prohibit dumping wastes in streams and rivers or burning in forests "
     "and open areas.'\n"
     "Sec. 13.2 (p. 207): 'open disposal of urban and industrial wastes or dumping in rivers have caused "
     "environmental pollution and affected biodiversity' — contextual evidence for this prohibition.\n"
     "Responsible authority (+0.25): Ministry of Forests and Environment (Appendix 1, Ch.13).\n"
     "Enforcement (−0): No explicit penalties or enforcement mechanism stated for this prohibition in "
     "the plan text.\n"
     "Monitoring (+0.25): Sec. 13.5.8: 'Expand measurement centers for climate, pollution, etc.' and "
     "Appendix 1 monitoring.\n"
     "Unconditional (+0.25): No exemptions stated for the dumping/burning prohibition."),
  W=("This is the strongest environmental leakage instrument in the plan. 'Prohibit' uses mandatory "
     "language but within a planning document — S = 0 as the plan commits to enacting/enforcing this "
     "prohibition but has not yet operationalised it through a specific binding regulation. "
     "The SWM Act 2011 and Environmental Protection Act 2019 contain related provisions; this plan "
     "commits to strengthening/reinforcing those existing measures. "
     "T = 0.75 (no enforcement sub-score) because no specific penalties are stated for violations of "
     "this commitment in the plan text. Cross-reference: Ch.7 Sec.7.5.9 mandates 'cleaning of river "
     "sources and conservation program' including mandatory sewerage treatment — a related instrument."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Plastic use discouragement + prohibition on burning plastics + control of plastic items",
 "Instrument 2 — Urban waste management infrastructure (17 centers target) + separate collection standards",
 "Instrument 3 — Extended Producer Responsibility (EPR) for producers/importers of industrial & hazardous waste",
 "Instrument 4 — Landfill waste tax",
 "Instrument 5 — Waste disposal prohibition (dumping in rivers; burning in open areas)",
]),start=1):
 r=4+i
 ws.cell(row=r,column=24,value=label).font=hfont(italic=True,size=8,color="808080")

 pol_data=[POL["name"],POL["url"],POL["year"],POL["obj"],POL["tgt"],POL["tgt_t"],
           POL["gtype"],POL["gtype_j"],POL["intg"],POL["sects"],POL["circ"],POL["lc"],
           POL["budg"],POL["budg_t"],None]
 af=hfill(ORANGE); pf=hfill(LIGHT_BLUE)
 for ci,val in enumerate(pol_data,1):
  is_auto=(ci==15)
  cell=ws.cell(row=r,column=ci)
  if is_auto:
   cell.value=f"=AVERAGE(G{r},I{r},K{r},M{r},P{r})"
   cell.fill=af; cell.font=hfont(italic=True,size=9,color="7F3F00"); cell.number_format="0.00"
  else:
   cell.value=val; cell.fill=pf; cell.font=hfont(size=9)
   if ci==3: cell.alignment=CA; cell.number_format="0"
   elif ci in(5,7,9,11,13): cell.alignment=CA; cell.number_format="0.00"
   else: cell.alignment=WA
  cell.border=tb()

 ifl=hfill(FILLS[i-1])
 idata=[instr["P"],instr["Q"],instr["R"],instr["S"],instr["T"],instr["U"],None,instr["W"]]
 for j,val in enumerate(idata):
  ci=16+j; is_auto=(ci==22)
  cell=ws.cell(row=r,column=ci)
  if is_auto:
   cell.value=f"=AVERAGE(P{r},T{r})"; cell.fill=af
   cell.font=hfont(italic=True,size=9,color="7F3F00"); cell.number_format="0.000"
  else:
   cell.value=val; cell.fill=ifl; cell.font=hfont(size=9)
   if ci in(16,18,19,20): cell.alignment=CA; cell.number_format="0.00" if ci in(16,20) else "0"
   else: cell.alignment=WA
  cell.border=tb()
 ws.row_dimensions[r].height=220

# Column widths
for c,w in {1:28,2:32,3:8,4:35,5:8,6:30,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:42,19:10,20:10,21:42,22:9,23:42,24:25}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Score Summary")
evidence=[
 ("SCORE SUMMARY",None),
 ("Instrument","P (type)","S (in force)","T (impl.)","V (score [auto])"),
 ("1 — Plastic use discouragement + prohibition on burning plastics + control of plastic items","1.00","0","0.50","0.750"),
 ("2 — Urban waste management infrastructure (17 centers target) + separate collection standards","0.80","0","0.50","0.650"),
 ("3 — EPR for producers/importers of industrial and hazardous waste","1.00","0","0.50","0.750"),
 ("4 — Landfill waste tax","0.60","0","0.75","0.675"),
 ("5 — Waste disposal prohibition (dumping in rivers; burning in open areas)","1.00","0","0.75","0.875"),
 ("","","","",""),
 ("POLICY-LEVEL SCORES",None),
 ("Field","Value"),
 ("policy_type (G)","0.50  (strategy/plan with quantifiable targets)"),
 ("policy_target (E)","1  — Target 8: waste management centers 1 → 17 by FY 2028/29"),
 ("policy_integration (I)","1.00  (9 sectors)"),
 ("policy_circularity (K)","0.75  (4 lifecycle phases: consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","0.50  (NPR 3,991 million allocated to waste management & recycling sector; not ring-fenced for plastic)"),
 ("policy_score (O)","[auto] = AVERAGE(G, I, K, M, P per row)"),
 ("",""),
 ("KEY VERBATIM EVIDENCE — PLASTIC-RELEVANT PASSAGES",None),
 ("Location","Verbatim text"),
 ("Ch.7, Sec.7.5.10 (p.133) — Plastic ban",
  "'Discourage the use of plastic materials and prohibit burning plastics.'"),
 ("Ch.7, Sec.7.6, Target 8 (p.134) — Quantitative target",
  "'Modern integrated solid waste management centers and sanitary landfill sites: "
  "FY 2022/23 actual: 1; FY 2028/29 target: 17.'"),
 ("Ch.13, Sec.13.4, Strategy 3 (p.224) — EPR",
  "'Make producers or importers and polluters responsible for industrial, hospital and hazardous "
  "wastes management, as well as implement the concept of extended producer responsibility and "
  "cut down industrial emissions.'"),
 ("Ch.13, Sec.13.5, Item 1 (p.212) — Displace plastics",
  "'Increase the use of products based on green raw materials, displacing plastic products.'"),
 ("Ch.13, Sec.13.5, Item 4 (p.213) — Waste tax + plastic control + river ban",
  "'Make waste management cost effective and sustainable by involving the private sector; "
  "utilize waste reduction, reuse and recycling. Prohibit dumping wastes in streams and rivers "
  "or burning in forests and open areas. Impose tax on wastes disposed of in landfills. "
  "Control plastic items.'"),
 ("Ch.13, Sec.13.4, Strategy 3 (p.224) — Polluter pays",
  "'Enforce the polluter pays principle. Make waste management cost effective through "
  "public-private partnerships.'"),
]
for ri,row in enumerate(evidence,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=5)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif len(row)>=2 and row[1] in("P (type)","Value","Verbatim text"):
  for ci2,v in enumerate(row,1):
   if v:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
    c.alignment=Alignment(wrap_text=True); c.border=tb()
 elif row[0]=="": pass
 else:
  for ci2,v in enumerate(row,1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
 ws2.row_dimensions[ri].height=32
ws2.column_dimensions["A"].width=55
for col in "BCDE": ws2.column_dimensions[col].width=12

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational, no concrete targets"),
     (0.50,"Strategy/plan — incorporates quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("POLICY INTEGRATION (Col I)",None),("Score","Description"),
     (0,"0 sectors"),(0.25,"1–2 sectors"),(0.50,"3–4 sectors"),(0.75,"5–6 sectors"),(1.00,"7+ sectors"),
     ("",""),("POLICY CIRCULARITY (Col K)",None),("Score","Description"),
     (0.25,"1 lifecycle phase"),(0.50,"2 lifecycle phases"),(0.75,"3–4 lifecycle phases"),(1.00,"All 5 phases"),
     ("",""),("POLICY BUDGET (Col M)",None),("Score","Description"),
     (0,"No budget"),(0.50,"Budget mentioned, not ring-fenced"),(1.00,"Ring-fenced / self-generating fund"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic (tax, levy, subsidy)"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, mandatory standard)"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — plan-level commitment, enabling power not yet operationalised"),
     (1,"In force — mandatory language (shall/must/prohibited) or operationalised by subordinate instrument"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T — additive)",None),
     ("Sub-score","Criterion (must be explicitly evidenced)"),
     ("+0.25","Responsible authority designated"),("+0.25","Enforcement (fines/penalties)"),
     ("+0.25","Monitoring mechanism"),("+0.25","Unconditional (no exemptions)")]
for ri,row in enumerate(ref,1):
 a,b=row[0],row[1] if len(row)>1 else None
 ca=ws3.cell(row=ri,column=1,value=a)
 if b is None and a:
  ws3.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  ca.fill=hfill(MID_BLUE); ca.font=hfont(bold=True,size=10,color=WHITE); ca.alignment=CA
 elif b in("Description","Criterion (must be explicitly evidenced)"):
  ca.fill=hfill("C5D9F1"); ca.font=hfont(bold=True,size=9)
  ws3.cell(row=ri,column=2,value=b).fill=hfill("C5D9F1")
  ws3.cell(row=ri,column=2).font=hfont(bold=True,size=9)
 else:
  ca.font=hfont(size=9)
  if b: ws3.cell(row=ri,column=2,value=b).font=hfont(size=9)
  if isinstance(a,(int,float)): ca.alignment=CA; ca.number_format="0.00"
 ca.border=tb()
 if b: ws3.cell(row=ri,column=2).border=tb()
 ws3.row_dimensions[ri].height=18
ws3.column_dimensions["A"].width=14; ws3.column_dimensions["B"].width=70

path="/workspace/4p_index_16th_plan_nepal_2024.xlsx"
wb.save(path)
print(f"Saved: {path}")
