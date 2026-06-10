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

ws.merge_cells("A1:X1")
c=ws["A1"]; c.value="Plastic Pollution Policy Index (4P Index) — Coding Table"
c.fill=hfill(DARK_BLUE); c.font=hfont(bold=True,size=13,color=WHITE); c.alignment=CA
ws.row_dimensions[1].height=22

ws.merge_cells("A2:X2")
c=ws["A2"]
c.value=("Policy: Management Plan of Bardia National Park and its Buffer Zone (FY 2079/80–2083/84 BS / 2022/23–2026/27)  |  "
         "Country: Nepal  |  Year: 2022  |  "
         "Source: https://giwmscdnone.gov.np/media/pdf_upload/Management%20Plan%20of%20BardiaNPand%20BZ%20(2023_2027)_zedb9hi.pdf")
c.fill=hfill(MID_BLUE); c.font=hfont(size=9,color=WHITE); c.alignment=CA
ws.row_dimensions[2].height=16

ws.merge_cells("A3:O3")
c=ws["A3"]; c.value="SECTION A — Policy-Level Fields  (repeated identically across all rows)"
c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
ws.merge_cells("P3:X3")
c=ws["P3"]; c.value="SECTION B — Instrument-Level Fields  (one row per instrument)"
c.fill=hfill(TEAL); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
ws.row_dimensions[3].height=18

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

POL=dict(
 name="Management Plan of Bardia National Park and its Buffer Zone (FY 2079/80–2083/84 BS / 2022/23–2026/27)",
 url=("https://giwmscdnone.gov.np/media/pdf_upload/"
      "Management%20Plan%20of%20BardiaNPand%20BZ%20(2023_2027)_zedb9hi.pdf"),
 year=2022,
 obj=("Five-year management plan (2022/23–2026/27) for Bardia National Park (968 km²) and its Buffer "
      "Zone (507 km²) in Bardia, Banke and Surkhet districts, Madhesh and Lumbini Provinces, Nepal; "
      "issued by the Department of National Parks and Wildlife Conservation. In relation to plastics: "
      "prohibits use of polluting items such as plastic bags; promotes 5R (Reduce, Reuse, Recycle, "
      "Remove, Reject) approach; installs waste collection infrastructure (dustbins, waste disposal "
      "pits) at tourist entry points; prepares sanitation guidelines for accommodation operators "
      "(hotels/lodges/homestays/restaurants) and water/sanitation/hygiene guidelines for local "
      "communities; and monitors upstream water quality from industries."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Administrative management plan (FY 2022/23–2026/27) prepared by Bardia National Park Office "
          "under DNPWC, Ministry of Forests and Environment, Government of Nepal; published 2022. "
          "Not enacted by Parliament and not a sub-legislative regulation or executive decree. Issued "
          "as an operational management plan for the park's five-year period. Contains activity-level "
          "operational targets with budget allocations but no quantifiable plastic-specific reduction "
          "targets. Scored 0.25 (strategy/plan with aspirational commitments)."),
 intg=1,
 sects="tourism, waste management, conservation, water, agriculture, municipalities, industry",
 circ=0.75,
 lc="consumption, recycling, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Table 8 (p. 90, Section 11.1): 'Climate change and Solid waste management: NRs. 28,262,000 "
         "(1.42% of total plan budget of NRs. 1,989,765,000).' Total plan budget: NRs. 1,989,765,000 "
         "(approximately NPR 2 billion). Solid waste management activities budgeted at activity level "
         "(Annex V). Budget is an administrative plan allocation, not a legally ring-fenced dedicated "
         "fund or self-generating mechanism."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC","E2EFDA"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Section 9.2.4 (p. 79–80) — Solid Waste Management Activities: 'Provide support to manage "
     "garbage with special focus on reducing production, recycling, and destruction by prohibiting "
     "the use of polluting items such as plastic bags.' "
     "Context (Section 9.2.1, p. 79): 'Water sources along the major trails are being contaminated "
     "with human waste, plastics and water bottles.' "
     "Section 9.2.3 (5R Strategy): 'Promote recycle, reuse, reduce, remove, and reject (5R) approach "
     "to manage wastes in the Park.' Both core park and Buffer Zone are targeted."),
  S=0,
  T=0.75,
  U=("Sec. 9.2.1 (p. 79): 'Water sources along the major trails are being contaminated with human "
     "waste, plastics and water bottles.'\n"
     "Sec. 9.2.4 (p. 80): 'Provide support to manage garbage with special focus on reducing "
     "production, recycling, and destruction by prohibiting the use of polluting items such as "
     "plastic bags.'\n"
     "Sec. 9.2.3 (p. 79): '...Promote recycle, reuse, reduce, remove, and reject (5R) approach...'\n"
     "Responsible authority (+0.25): Bardia National Park Office and BZ institutions (BZUC/BZCF) are "
     "designated implementing bodies throughout the plan; Section 11.1 specifies implementing agencies.\n"
     "Enforcement (−0): No fines or penalties for plastic bag use explicitly stated in this plan.\n"
     "Monitoring (+0.25): Section 9.2.4 includes 'Monitor water quality nearby industries' — "
     "monitoring mechanism covers pollution aspects; annual progress reports required.\n"
     "Unconditional (+0.25): Plastic bag prohibition stated without exemptions in plan activities."),
  W=("instrument_in_force = 0: Plan commits to 'provide support to manage garbage by prohibiting '— "
     "this is a future implementation commitment, not a currently operative legal prohibition. The "
     "NPWC Act 2029 and its amendments give the park authority power to regulate activities within "
     "and around the park, but a specific plastic ban ordinance is needed. "
     "P = 1.0 (Regulatory) based on explicit prohibition language ('prohibiting the use of polluting "
     "items such as plastic bags'). "
     "The 5R approach is folded into this instrument as it combines regulatory prohibition with "
     "voluntary promotion; highest type (Regulatory) applied per coding rules. "
     "Plastic bag prohibition + 5R strategy combination: both directly target plastic at consumption "
     "stage. Cross-reference: Nepal's national Plastic Bag Control Directive 2082 may already "
     "operationalise elements of this plan commitment at national level."),
 ),
 dict(
  P=0.80, Q="Waste management",
  R=("Section 8.2.5 (p. 76) — Tourism Management Activities: 'Install dustbin in the BZ in proper "
     "numbers as per tourist pressure; Organize Clean-up campaign to manage waste in the highways "
     "(waste collection and disposal).' "
     "Section 9.2.4 (pp. 79–80): 'Construct waste disposal pits or put waste collection pots near "
     "entry point, ticket counter and Hattisar; Provide water supply, toilet, drainage, collection and "
     "recycling centre to schools, public buildings, and household with the support from conservation "
     "partners; Support eco-clubs to organize clean-up campaign regularly.'"),
  S=0,
  T=0.75,
  U=("Sec. 8.2.5 (p. 76): 'Install dustbin in the BZ in proper numbers as per tourist pressure.'\n"
     "'Organize Clean-up campaign to manage waste in the highways (waste collection and disposal).'\n"
     "Sec. 9.2.4 (p. 80): 'Construct waste disposal pits or put waste collection pots near entry "
     "point, ticket counter and Hattisar.'\n"
     "'Provide water supply, toilet, drainage, collection and recycling centre to schools...'\n"
     "'Support eco-clubs to organize clean-up campaign regularly.'\n"
     "Responsible authority (+0.25): BNP office, Nepal Army, BZUCs and eco-clubs designated.\n"
     "Enforcement (−0): No enforcement mechanism for waste infrastructure compliance.\n"
     "Monitoring (+0.25): Annual progress report tracking; eco-club campaign monitoring implicit.\n"
     "Unconditional (+0.25): Dustbin installation and clean-up campaigns stated as specific targets "
     "without stated exemptions."),
  W=("This instrument combines physical infrastructure (waste disposal pits, dustbins, recycling "
     "centres) and community campaigns (clean-up campaigns via eco-clubs). Highest type "
     "(Infrastructure, 0.80) applied per coding rules; the community campaign element (0.40) noted "
     "here. The 'collection and recycling centre' activity also includes recycling infrastructure, "
     "strengthening the 0.80 (Infrastructure) classification. All activities pass the plastics "
     "relevance filter under criterion (c) — waste management instruments that commonly apply to "
     "plastic waste streams."),
 ),
 dict(
  P=1.0, Q="Consumption",
  R=("Section 9.2.4 (pp. 79–80) — Solid Waste Management Activities: 'Prepare a common sanitation "
     "guideline to make hotel, lodge, homestay and restaurant adopt minimum standards.' "
     "Section 9.2.3 — Strategies: 'Develop water, sanitation and hygiene guideline for local "
     "communities in BNP.' These mandatory-minimum standards for accommodation operators and a "
     "community-level WASH guideline together constitute a standard-setting instrument for plastic "
     "and other waste management in the park and buffer zone."),
  S=0,
  T=0.50,
  U=("Sec. 9.2.4 (p. 80): 'Prepare a common sanitation guideline to make hotel, lodge, homestay "
     "and restaurant adopt minimum standards.'\n"
     "Sec. 9.2.3 (p. 79): 'Develop water, sanitation and hygiene guideline for local communities "
     "in BNP.'\n"
     "Responsible authority (+0.25): BNP park authority designated as the preparing body.\n"
     "Enforcement (−0): No explicit enforcement mechanism or penalties for non-compliance with "
     "sanitation guidelines stated in the plan.\n"
     "Monitoring (−0): No specific monitoring mechanism stated for accommodation compliance with "
     "the guideline.\n"
     "Unconditional (+0.25): Sanitation guideline to apply to 'hotel, lodge, homestay and restaurant' "
     "universally — no stated exemptions within the accommodation sector."),
  W=("instrument_in_force = 0: The plan commits to PREPARING both guidelines — these are future "
     "actions. Once issued, they would function as mandatory minimum standards (P = 1.0 Regulatory). "
     "P = 1.0 (Regulatory/mandatory standard) coded because the guideline is described as making "
     "accommodation operators 'adopt minimum standards' — mandatory language. "
     "Borderline case: P = 0.20 (governance) is defensible since the plan only commits to preparing "
     "the guidelines, not operationalising them. "
     "The combination of accommodation-sector sanitation guideline + community WASH guideline "
     "is treated as one standard-setting instrument since both address waste management practices "
     "at the consumption/use stage and are complementary in scope."),
 ),
 dict(
  P=0.20, Q="Environmental leakage",
  R=("Section 9.2.4 (pp. 79–80) — Solid Waste Management Activities: 'Monitor water quality nearby "
     "industries to check whether they treat their waste before discharging into the river.' "
     "'Pilot work with upstream local government – Tulshipur and Ghorahi sub-metropolitan cities, "
     "Babai river catchment.' "
     "Section 9.2.3: 'Engaging local government and private sector including those in the upstream.' "
     "This governance and monitoring instrument addresses the environmental leakage of plastic and "
     "other pollutants from upstream industries into waterways flowing through Bardia National Park."),
  S=0,
  T=0.75,
  U=("Sec. 9.2.4 (p. 80): 'Monitor water quality nearby industries to check whether they treat their "
     "waste before discharging into the river.'\n"
     "'Pilot work with upstream local government – Tulshipur and Ghorahi sub-metropolitan cities, "
     "Babai river catchment.'\n"
     "Sec. 9.2.3 (p. 79): 'Engaging local government and private sector including those in the "
     "upstream.'\n"
     "Responsible authority (+0.25): BNP office designated as monitoring body; coordination with "
     "upstream municipalities is explicitly assigned.\n"
     "Enforcement (−0): No explicit penalties for upstream industrial non-compliance stated in the "
     "plan.\n"
     "Monitoring (+0.25): Water quality monitoring explicitly stated as an activity — this is an "
     "explicit monitoring mechanism.\n"
     "Unconditional (+0.25): Both monitoring and coordination activities are stated without "
     "exemptions."),
  W=("This instrument is distinct from the Khaptad NP plan and unique to Bardia — it targets "
     "upstream pollution sources (industries and municipalities in Tulshipur and Ghorahi) whose "
     "waste including plastic enters the Babai river and affects the park ecosystem. "
     "P = 0.20 (Governance & coordination) because the instrument coordinates and monitors across "
     "levels of government (national park + sub-metropolitan municipalities) and assigns monitoring "
     "responsibilities. "
     "Passes plastics relevance filter under criterion (c): waste management governance instrument "
     "that commonly applies to plastic waste streams in practice, specifically controlling "
     "environmental leakage of plastic into river systems. "
     "S = 0: piloting upstream coordination is an aspirational plan activity; no binding agreement "
     "with Tulshipur/Ghorahi has been concluded within this plan."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Plastic bag prohibition + 5R waste reduction approach",
 "Instrument 2 — Waste collection infrastructure (dustbins/pits) + clean-up campaigns + recycling centres",
 "Instrument 3 — Sanitation guideline for accommodation operators + WASH guideline for communities",
 "Instrument 4 — Upstream water quality monitoring + governance coordination with municipalities",
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

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:42,19:10,20:10,21:42,22:9,23:42,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2
ws2=wb.create_sheet("Key Evidence & Score Summary")
evidence=[
 ("SCORE SUMMARY",None),
 ("Instrument","P (type)","S (in force)","T (impl.)","V (score [auto])"),
 ("1 — Plastic bag prohibition + 5R approach","1.00","0","0.75","0.875"),
 ("2 — Waste collection infrastructure (dustbins/pits) + clean-up campaigns","0.80","0","0.75","0.775"),
 ("3 — Sanitation guideline (accommodation) + WASH guideline (communities)","1.00","0","0.50","0.750"),
 ("4 — Upstream water quality monitoring + governance coordination with municipalities","0.20","0","0.75","0.475"),
 ("","","","",""),
 ("POLICY-LEVEL SCORES",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (management plan — aspirational, no quantifiable plastic targets)"),
 ("policy_target (E)","0"),
 ("policy_integration (I)","1.00  (7 sectors: tourism, waste management, conservation, water, agriculture, municipalities, industry)"),
 ("policy_circularity (K)","0.75  (4 lifecycle phases: consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","0.50  (NRs. 28,262,000 allocated to climate change + solid waste management; not ring-fenced)"),
 ("policy_score (O)","[auto] = AVERAGE(G, I, K, M, P per row)"),
 ("",""),
 ("KEY VERBATIM EVIDENCE — PLASTIC-RELEVANT PASSAGES",None),
 ("Location","Verbatim text"),
 ("Sec. 9.2.1, p.79 — Context (plastic in water sources)",
  "'Water sources along the major trails are being contaminated with human waste, plastics and "
  "water bottles.'"),
 ("Sec. 9.2.4, p.80 — Plastic bag prohibition (with budget)",
  "'Provide support to manage garbage with special focus on reducing production, recycling, and "
  "destruction by prohibiting the use of polluting items such as plastic bags.'"),
 ("Sec. 9.2.3, p.79 — 5R strategy",
  "'Promote recycle, reuse, reduce, remove, and reject (5R) approach to manage wastes in the Park.'"),
 ("Sec. 9.2.4, p.80 — Sanitation guideline",
  "'Prepare a common sanitation guideline to make hotel, lodge, homestay and restaurant adopt "
  "minimum standards.'"),
 ("Sec. 9.2.3, p.79 — WASH guideline",
  "'Develop water, sanitation and hygiene guideline for local communities in BNP.'"),
 ("Sec. 9.2.4, p.80 — Water quality monitoring (upstream)",
  "'Monitor water quality nearby industries to check whether they treat their waste before "
  "discharging into the river.'"),
 ("Sec. 9.2.4, p.80 — Upstream coordination",
  "'Pilot work with upstream local government – Tulshipur and Ghorahi sub-metropolitan cities, "
  "Babai river catchment.'"),
 ("Sec. 8.2.5, p.76 — Waste infrastructure",
  "'Install dustbin in the BZ in proper numbers as per tourist pressure; Organize Clean-up "
  "campaign to manage waste in the highways (waste collection and disposal).'"),
 ("Table 8, p.90 — Budget",
  "'Climate change and Solid waste management: NRs. 28,262,000 (1.42% of total plan budget "
  "of NRs. 1,989,765,000).'"),
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
 ws2.row_dimensions[ri].height=35
ws2.column_dimensions["A"].width=52
for col in "BCDE": ws2.column_dimensions[col].width=12

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational, no concrete plastic targets"),
     (0.50,"Strategy/plan — incorporates quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("POLICY INTEGRATION (Col I)",None),("Score","Description"),
     (0,"0 sectors"),(0.25,"1–2 sectors"),(0.50,"3–4 sectors"),(0.75,"5–6 sectors"),(1.00,"7+ sectors"),
     ("",""),("POLICY CIRCULARITY (Col K)",None),("Score","Description"),
     (0.25,"1 lifecycle phase"),(0.50,"2 lifecycle phases"),
     (0.75,"3–4 lifecycle phases"),(1.00,"All 5 phases"),
     ("",""),("POLICY BUDGET (Col M)",None),("Score","Description"),
     (0,"No budget"),(0.50,"Budget mentioned, not ring-fenced"),(1.00,"Ring-fenced / self-generating fund"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, mandatory standard)"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — plan-level commitment not yet operationalised"),
     (1,"In force — mandatory language or operationalised by subordinate instrument"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T)",None),
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

path="/workspace/4p_index_bnp_management_plan_2022.xlsx"
wb.save(path)
print(f"Saved: {path}")
