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
c.value=("Policy: Management Plan of Khaptad National Park and its Buffer Zone (FY 2081/82–2085/86 BS / 2024/25–2028/29)  |  "
         "Country: Nepal  |  Year: 2024  |  "
         "Source: https://giwmscdntwo.gov.np/media/pdf_upload/Khaptad%20NP%20Management%20Plan%20(2082-2086)_vshs1ki.pdf")
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
 name="Management Plan of Khaptad National Park and its Buffer Zone (FY 2081/82–2085/86 BS / 2024/25–2028/29)",
 url=("https://giwmscdntwo.gov.np/media/pdf_upload/"
      "Khaptad%20NP%20Management%20Plan%20(2082-2086)_vshs1ki.pdf"),
 year=2024,
 obj=("Five-year management plan (2024/25–2028/29) for Khaptad National Park (225 km²) and its Buffer "
      "Zone (216 km²) in Sudurpaschim Province, Nepal, issued by the Department of National Parks and "
      "Wildlife Conservation. Addresses biodiversity conservation, habitat management, species protection, "
      "tourism, human-wildlife conflict, and community livelihoods. In relation to plastics: includes "
      "banning of plastic bags and prohibition on polluting plastic items; waste collection infrastructure "
      "(70 dustbins); organising clean-up campaigns to collect plastics; preparing a common sanitation "
      "guideline for tourism accommodation operators; training on waste segregation and recycling; "
      "and a 5R (Reduce, Reuse, Recycle, Remove, Reject) approach to solid waste in the park and BZ."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Five-year management plan for Khaptad National Park and its Buffer Zone, prepared by the Khaptad "
          "National Park Office under the Department of National Parks and Wildlife Conservation (DNPWC), "
          "Ministry of Forests and Environment, Government of Nepal; published 2024. Issued as an "
          "administrative/operational management plan — not enacted by Parliament and not a sub-legislative "
          "regulation or executive decree. The plan contains activity-level operational targets with budget "
          "allocations but no quantifiable plastic-specific reduction targets. Scored 0.25 "
          "(strategy/plan with aspirational commitments)."),
 intg=1,
 sects="tourism, waste management, conservation, water, agriculture, energy, municipalities",
 circ=0.75,
 lc="consumption, recycling, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Table 7, Section 11.1 (p. 64): 'Climate change adaptation and Solid waste management — "
         "total planned budget: NRs. 24,275,000 (2.70% of total plan budget of NRs. 899,797,000).' "
         "Specific solid waste activity budgets in Section 10.4 budget tables (p. 131): "
         "garbage disposal demonstrations NRs. 512,500; prohibiting plastic bags NRs. 820,000; "
         "organising clean-up campaigns to collect plastics NRs. 3,063,570; preparing sanitation "
         "guideline NRs. 300,000. Budget is an administrative allocation within the management plan "
         "framework, subject to annual government appropriations; not legally ring-fenced as a "
         "dedicated fund or self-generating mechanism."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC","E2EFDA"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Section 9.2.3.3 (p. 52) — Solid Waste Management Strategies: 'Promote waste reduction practices "
     "by banning plastics, promoting reusable items, and providing water refilling station.' "
     "Section 9.2.3.4 / Activity 10.4.2 (pp. 53, 131) — Solid Waste Management Activities: "
     "'Provide support to manage garbage with special focus on reducing production, recycling, by "
     "prohibiting use of polluting items such as plastic bags.' "
     "Context (Sec. 9.2.3.1, p. 52): 'In the BZ, water sources along the major trails are being "
     "contaminated due to huge amount of plastics.' Both the core park and Buffer Zone are targeted."),
  S=0,
  T=0.75,
  U=("Sec. 9.2.3.3 (p. 52): 'Promote waste reduction practices by banning plastics, promoting "
     "reusable items, and providing water refilling station.'\n"
     "Activity 10.4.2 (p. 131): 'Provide support to manage garbage with special focus on reducing "
     "production, recycling, by prohibiting use of polluting items such as plastic bags.' "
     "Budget: NRs. 820,000 over 2 years.\n"
     "Responsible authority (+0.25): KNP Senior Conservation Officer (SCO), Nepal Army, BZUCs "
     "designated as implementing bodies throughout plan; Section 11.1 designates implementing "
     "agencies per activity.\n"
     "Enforcement (−0): No fines or penalties for plastic bag use explicitly stated in the plan.\n"
     "Monitoring (+0.25): Section 3.6.2 (p. 17): 'Monitoring activities in KNP primarily focus on "
     "assessing...solid waste management during festivals.' Annual progress reports required.\n"
     "Unconditional (+0.25): The plastic bag prohibition is stated without exemptions in the "
     "plan activities."),
  W=("instrument_in_force = 0: The plan 'promotes banning' and 'provides support to prohibit' plastic "
     "bags — these are future implementation commitments, not existing operative regulations. The "
     "National Parks and Wildlife Conservation Act 2029 (1972) gives the park authority to regulate "
     "activities within the park; a specific plastic ban regulation would need to be issued under "
     "DNPWC authority to be in force. The plastic bag ban is explicitly stated as a strategy (S.9.2.3.3) "
     "and an activity (10.4.2), which is the strongest plastic-specific commitment in this plan. "
     "Cross-reference: Nepal's national Plastic Bag Directive 2082 (national-level) may already "
     "operationalise some elements of this commitment; if so, S could be updated to 1."),
 ),
 dict(
  P=0.80, Q="Waste management",
  R=("Section 8.4 / Activities 9.21 (pp. 48, 127) — Tourism Management Activities: 'Install 70 dustbin "
     "in the BZ in proper numbers as per tourist pressure.' 'Organize Clean-up campaign to manage waste "
     "in the highway (waste collection and disposal).' "
     "Section 9.2.3.3 (p. 52) — Strategy: 'Establish waste collection bins labelled with different "
     "types of waste, including recyclables, organic waste, and non-recyclable waste.' "
     "Activity 10.4.3 (p. 131): 'Organize clean up campaigns together with security units and BZUC to "
     "collect plastics, glass, paper, and metal.' Budget: NRs. 3,063,570 over 5 years. "
     "Section 9.2.3.3: 'Promote recycle, reuse, reduce, remove, and reject (5R) approach.'"),
  S=0,
  T=0.75,
  U=("Activity 8.4 (p. 48): 'Install 70 dustbin in the BZ in proper numbers as per tourist pressure.'\n"
     "Activity 9.21 (p. 127): 'Organize Clean-up campaign to manage waste in the highway (waste "
     "collection and disposal).' Budget: NRs. 2,200 (thousand) over 5 years.\n"
     "Activity 10.4.3 (p. 131): 'Organize clean up campaigns together with security units and BZUC "
     "to collect plastics, glass, paper, and metal.' Budget: NRs. 3,063.57 (thousand) over 5 years.\n"
     "Responsible authority (+0.25): KNP office, Nepal Army, BZUCs designated.\n"
     "Enforcement (−0): No enforcement mechanism for waste infrastructure compliance.\n"
     "Monitoring (+0.25): Annual progress report system; waste management monitoring during festivals "
     "(Sec. 3.6.2).\n"
     "Unconditional (+0.25): 70 dustbins and 5 clean-up campaigns stated as specific targets without "
     "stated exemptions."),
  W=("This instrument combines infrastructure (70 dustbins with labelled waste types for recycling/ "
     "organic/non-recyclable) and governance (5R approach and clean-up campaigns). Highest type "
     "(Infrastructure, 0.80) applied per coding rules; the 5R promotional aspect (which would be 0.40) "
     "is noted here. The 70-dustbin target and 5 annual clean-up campaigns are operationally specific. "
     "Activity 10.4.3 explicitly targets plastics: 'collect plastics, glass, paper, and metal' — "
     "passes criterion (a). Instrument applies to both core park and Buffer Zone areas."),
 ),
 dict(
  P=1.0, Q="Consumption",
  R=("Section 8.4 / Activity 9.22 (pp. 48, 127) — Tourism Management Activities: 'Prepare operational "
     "and management guidelines for home stay in KNP.' "
     "Activity 10.4.4 (pp. 53, 131) — Solid Waste Management Activities: 'Prepare a common sanitation "
     "guideline to make hotel, lodge, homestay and restaurant to make them adopt minimum standards.' "
     "Budget: NRs. 300,000. This mandatory minimum standard for accommodation operators would govern "
     "plastic waste management in tourist-facing establishments within the park and BZ."),
  S=0,
  T=0.50,
  U=("Activity 10.4.4 (p. 131): 'Prepare a common sanitation guideline to make hotel, lodge, homestay "
     "and restaurant to make them adopt minimum standards.' Budget: NRs. 300,000 in Year 1.\n"
     "Activity 9.22 (p. 127): 'Prepare operational and management guidelines for home stay in KNP.' "
     "Budget: NRs. 825,000 over 5 years.\n"
     "Responsible authority (+0.25): KNP park authority designated as the preparing body.\n"
     "Enforcement (−0): No explicit enforcement mechanism or penalties for non-compliance with the "
     "sanitation guideline stated in the plan.\n"
     "Monitoring (−0): No specific monitoring mechanism stated for accommodation compliance.\n"
     "Unconditional (+0.25): The guideline is to apply to 'hotel, lodge, homestay and restaurant' "
     "— universal within its scope with no stated exemptions."),
  W=("The plan commits to PREPARING a common sanitation guideline — this is a future planning action "
     "(S = 0). The sanitation guideline is the most governance-oriented plastic-relevant instrument "
     "in this plan. It would function as a mandatory performance standard once issued, targeting "
     "consumption practices (plastic waste generation) at accommodation operators. "
     "P = 1.0 (Regulatory - mandatory performance standard) because the guideline is described as "
     "requiring hotels/lodges to 'adopt minimum standards' — mandatory language. "
     "However, P = 0.20 (governance) is defensible since the plan merely commits to preparing the "
     "guideline, not enacting it. Borderline case noted. Cross-reference: Activity 9.22 (homestay "
     "operational guidelines) is a related but separate governance instrument for homestay operations "
     "specifically; both are combined here as a single standard-setting instrument."),
 ),
 dict(
  P=0.40, Q="Waste management",
  R=("Section 6.3 (p. 36) — Capacity Building for Frontline Staff: 'Providing training on waste "
     "management practices, including waste segregation, recycling procedures, and proper disposal "
     "techniques.' "
     "Section 9.2.3.3 (p. 52) — Strategies: 'Develop educational materials and signage to inform "
     "visitors about proper waste management practices; Mobilize eco-clubs to raise awareness about "
     "importance of solid waste management.' "
     "Activity 10.4.1 (p. 131): 'Provide support to demonstrate proper techniques of garbage disposal "
     "and recycling techniques.' Budget: NRs. 512,500."),
  S=0,
  T=0.75,
  U=("Sec. 6.3 (p. 36): 'Providing training on waste management practices, including waste segregation, "
     "recycling procedures, and proper disposal techniques.'\n"
     "Activity budget table (p. 121): 'Providing training on waste management practices, including "
     "waste segregation, recycling procedures, and proper disposal techniques — NRs. 1,000 per session, "
     "5 sessions = NRs. 1,100 (thousands) over 5 years.'\n"
     "Sec. 9.2.3.3: 'Develop educational materials and signage to inform visitors about proper waste "
     "management practices; Mobilize eco-clubs to raise awareness about importance of solid waste "
     "management.'\n"
     "Responsible authority (+0.25): KNP administration and BZ institutions responsible for capacity "
     "building (Section 6.3 explicitly designates this).\n"
     "Enforcement (−0): No penalties for non-participation in training.\n"
     "Monitoring (+0.25): Activity budget table and annual progress reports track training activities.\n"
     "Unconditional (+0.25): Training applies to 'frontline staff and security units' universally; "
     "eco-club mobilization and educational materials for visitors have no stated exemptions."),
  W=("Training and awareness are critical enablers of the plastic ban and 5R approach commitments. "
     "S = 0: training activities are future commitments. P = 0.40 (Information & voluntary) reflects "
     "the educational nature of this instrument; the waste management training explicitly covers "
     "'waste segregation, recycling procedures, and proper disposal techniques' — directly relevant "
     "to plastic waste. The eco-club mobilization and visitor signage elements also address plastic "
     "waste awareness (criterion a — mentions plastics in context). T = 0.75 because responsible "
     "authority, monitoring, and unconditional sub-scores are met."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Plastic bag ban / prohibition on polluting plastic items",
 "Instrument 2 — Waste collection infrastructure (70 dustbins + labelled bins) and clean-up campaigns",
 "Instrument 3 — Common sanitation guideline for hotels/lodges/homestays/restaurants",
 "Instrument 4 — Training on waste management (waste segregation, recycling, disposal)",
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
for c,w in {1:30,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:42,19:10,20:10,21:42,22:9,23:42,24:25}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Score Summary")
evidence=[
 ("SCORE SUMMARY",None),
 ("Instrument","P (type)","S (in force)","T (impl.)","V (score [auto])"),
 ("1 — Plastic bag ban / prohibition on polluting plastic items","1.00","0","0.75","0.875"),
 ("2 — Waste collection infrastructure (70 dustbins + labelled bins) + clean-up campaigns","0.80","0","0.75","0.775"),
 ("3 — Common sanitation guideline for hotels/lodges/homestays/restaurants","1.00","0","0.50","0.750"),
 ("4 — Training on waste management (segregation, recycling, disposal)","0.40","0","0.75","0.575"),
 ("","","","",""),
 ("POLICY-LEVEL SCORES",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (strategy/plan — aspirational commitments, no quantifiable plastic targets)"),
 ("policy_target (E)","0"),
 ("policy_integration (I)","1.00  (7 sectors: tourism, waste management, conservation, water, agriculture, energy, municipalities)"),
 ("policy_circularity (K)","0.75  (4 lifecycle phases: consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","0.50  (budget allocated to solid waste management activities; not legally ring-fenced)"),
 ("policy_score (O)","[auto] = AVERAGE(G, I, K, M, P per row)"),
 ("",""),
 ("KEY VERBATIM EVIDENCE — PLASTIC-RELEVANT PASSAGES",None),
 ("Location","Verbatim text"),
 ("Sec. 9.2.3.1, p.52 — Context",
  "'On the other hand, improper management is adopted to dispose of other non-biodegradable waste "
  "types, including glass, metal, and plastic materials. In the BZ, water sources along the major "
  "trails are being contaminated due to huge amount of plastics.'"),
 ("Sec. 9.2.3.3, p.52 — Strategy (plastic ban)",
  "'Promote waste reduction practices by banning plastics, promoting reusable items, and providing "
  "water refilling station.'"),
 ("Sec. 9.2.3.3, p.52 — Strategy (5R)",
  "'Promote recycle, reuse, reduce, remove, and reject (5R) approach to manage wastes in the Park.'"),
 ("Activity 10.4.2, p.131 — Plastic bag prohibition (with budget NRs 820,000)",
  "'Provide support to manage garbage with special focus on reducing production, recycling, by "
  "prohibiting use of polluting items such as plastic bags.'"),
 ("Activity 10.4.3, p.131 — Clean-up for plastics (budget NRs 3,063,570)",
  "'Organize clean up campaigns together with security units and BZUC to collect plastics, glass, "
  "paper, and metal.'"),
 ("Activity 10.4.4, p.131 — Sanitation guideline (budget NRs 300,000)",
  "'Prepare a common sanitation guideline to make hotel, lodge, homestay and restaurant to make "
  "them adopt minimum standards.'"),
 ("Sec. 8.4, p.48 — Dustbins",
  "'Install 70 dustbin in the BZ in proper numbers as per tourist pressure.'"),
 ("Sec. 6.3, p.36 — Training",
  "'Providing training on waste management practices, including waste segregation, recycling "
  "procedures, and proper disposal techniques.'"),
 ("Table 7, p.64 — Budget",
  "'Climate change adaptation and Solid waste management: NRs. 24,275,000 (2.70% of total "
  "plan budget of NRs. 899,797,000).'"),
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
ws2.column_dimensions["A"].width=50
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
     (0.25,"1 lifecycle phase"),(0.50,"2 lifecycle phases"),(0.75,"3–4 lifecycle phases"),(1.00,"All 5 phases"),
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

path="/workspace/4p_index_knp_management_plan_2024.xlsx"
wb.save(path)
print(f"Saved: {path}")
