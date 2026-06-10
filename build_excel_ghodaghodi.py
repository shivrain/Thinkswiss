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
c.value=("Policy: Comprehensive Master Plan of Ghodaghodi Lake Area (2077 BS / 2020)  |  Country: Nepal  |  Year: 2020  |  "
         "Source: https://ghodaghodimun.gov.np/sites/ghodaghodimun.gov.np/files/Comprehensive%20Master%20Plan%20of%20GLA.pdf  "
         "| FAOLEX: https://faolex.fao.org/docs/pdf/nep220194.pdf")
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
 name="Comprehensive Master Plan of Ghodaghodi Lake Area (2077 BS)",
 url=("https://ghodaghodimun.gov.np/sites/ghodaghodimun.gov.np/files/"
      "Comprehensive%20Master%20Plan%20of%20GLA.pdf\n"
      "[Also available at FAOLEX: https://faolex.fao.org/docs/pdf/nep220194.pdf]"),
 year=2020,
 obj=("Long-term conservation and tourism development master plan for Ghodaghodi Lake Area (GLA), "
      "a Ramsar Wetland site in Ghodaghodi Municipality, Kailali, Sudurpaschim Province, Nepal. "
      "Prepared by the Comprehensive Ghodaghodi Lake and Tourism Development Board (CGLTDB) under "
      "Ghodaghodi Municipality. In relation to plastics: explicitly identifies 'increasing use of "
      "plastic waste and rising water pollution' as a key threat to the lake area (SWOT Threats, "
      "p.29); includes a 'Pragmatic Garbage Management System' with two-type waste segregation "
      "(degradable/non-degradable dustbins), recycling priority, and annual cleaning campaigns; "
      "mandates campaigns to 'collect and pick out wastes like plastic bottles, bags and other sorts "
      "of wastes' from the surface and underwater of Ghodaghodi Lake."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Comprehensive municipal master plan prepared by Ghodaghodi Municipality (through the "
          "Comprehensive Ghodaghodi Lake and Tourism Development Board, CGLTDB) in 2077 BS "
          "(approximately 2020 CE). Formulated under the Local Government Operation Act, 2074 "
          "(LGOA 2017). Not enacted by Parliament and not a national regulation/executive decree. "
          "It is a municipal-level tourism and conservation strategy/plan document. The plastic-"
          "relevant provisions are embedded in broader tourism and environmental sustainability "
          "strategies with planning/aspiration language ('should', 'shall be', 'organize campaigns'). "
          "No quantifiable plastic-specific targets. G = 0.25 (strategy/plan with aspirational "
          "commitments) — the lowest appropriate category given the plan's overall advisory character "
          "at the municipal level."),
 intg=1,
 sects="tourism, conservation/environment, waste management, water/wetlands, agriculture, municipalities, fisheries, infrastructure",
 circ=0.75,
 lc="consumption, disposal, recycling, environmental leakage",
 budg=0.5,
 budg_t=("The master plan includes detailed physical infrastructure cost estimates in Annex 2 "
         "(Proposed Components: Master Plan of Tourism Physical Infrastructure in GLA) and logical "
         "framework analyses (LFAs) for each strategic plan area including environment-friendly "
         "infrastructure. The CGLTDB under Ghodaghodi Municipality is the responsible body for "
         "implementation. Budget allocations are specified per component in the LFA tables but "
         "no separate budget is ring-fenced specifically for plastic waste management. M = 0.5 "
         "(budget/funding source mentioned for overall plan; not ring-fenced for plastic waste)."),
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=0.80, Q="Waste management",
  R=("Section 7.2.1 — Infrastructure Development (pp.51–52): 'Establish a Pragmatic Garbage "
     "Management System' as a key infrastructure activity:\n\n"
     "'b. Garbage management procedures at GLA complex system should be activated through "
     "formulating a main committee or sub-committee for making responsible.\n"
     "c. Must link with GLA garbage management system with the garbage management holistic "
     "system of GM.\n"
     "d. Arrange hardcore items [put two types of dustbins (one each for degradable and "
     "non-degradable)] at key junctions including frontal entry gate, religious complex areas, "
     "hiking trails of insidious area of GLA etc.\n"
     "e. Put in and apply the system of timely delivery of collected garbage to the landfilled site "
     "or recycling site or composting site. Prioritize to process the garbage (reuse, recycle etc.). "
     "Dumping at landfilled site to be as last option only.\n"
     "f. Agree on soft-core matters (e.g. formulating an annual calendar for cleaning campaign, "
     "collection, segregation and garbage processing schedule) through the volunteer participations "
     "of civil society and communicate it for collective actions.'\n\n"
     "This infrastructure instrument combines physical waste collection infrastructure (dustbins, "
     "waste delivery system to recycling/landfill sites) with governance elements (management "
     "committee, annual calendar). Non-biodegradable (plastic) waste streams are specifically "
     "addressed through the two-type segregation system."),
  S=0,
  T=0.75,
  U=("Responsible authority (+0.25): 'formulating a main committee or sub-committee for making "
     "responsible' — a designated waste management committee at GLA complex level is explicitly "
     "required. CGLTDB is the overall implementing agency. 'Must link with GLA garbage management "
     "system with the garbage management holistic system of GM' designates Ghodaghodi Municipality "
     "(GM) as the broader authority.\n\n"
     "Enforcement (−0): No explicit fines or penalties for non-compliance with waste management "
     "provisions stated in the master plan.\n\n"
     "Monitoring (+0.25): 'formulating an annual calendar for cleaning campaign, collection, "
     "segregation and garbage processing schedule' — an explicit monitoring/reporting mechanism "
     "through annual planning cycles; 'communicate it for collective actions' implies regular "
     "progress reviews.\n\n"
     "Unconditional (+0.25): The garbage management system applies to all of GLA including "
     "'frontal entry gate, religious complex areas, hiking trails' — comprehensive coverage with "
     "no stated exemptions. 'Must link' uses mandatory language for the municipal connection."),
  W=("P = 0.80 (Infrastructure — waste management facility investment): The instrument primarily "
     "involves physical infrastructure construction (two-type dustbins at key junctions, waste "
     "delivery system to recycling/landfill sites) combined with governance (management committee). "
     "Highest applicable type = Infrastructure (0.80).\n\n"
     "S = 0: Plan-level commitment. Language is predominantly planning/advisory ('should be "
     "activated', 'Arrange', 'Put in and apply', 'Agree on'), though 'Must link' uses mandatory "
     "language for one specific sub-provision. Overall, this is a future implementation commitment "
     "in a master plan document, not a currently operative regulatory standard.\n\n"
     "T = 0.75: Authority (+0.25), monitoring (+0.25), unconditional (+0.25) credited. "
     "Enforcement not credited — no penalties stated for non-compliance.\n\n"
     "PLASTIC WASTE RELEVANCE: The two-type dustbin system (degradable/non-degradable) directly "
     "addresses plastic waste streams in the 'non-biodegradable' category. The recycling priority "
     "('Prioritize to process the garbage (reuse, recycle etc.)') is especially relevant for "
     "plastic waste. The master plan explicitly identifies 'increasing use of plastic waste and "
     "rising water pollution' as a key threat (SWOT Threats, p.29).\n\n"
     "Cross-reference: The garbage management plan is to be linked with GM's holistic waste "
     "management system, which is governed by Nepal's Solid Waste Management Act 2011 and "
     "the 2068 (2011) SWM Rules."),
 ),
 dict(
  P=0.40, Q="Environmental leakage",
  R=("Section 7.2.6 — Wetland Conservation and Management (p.66): Among the measures to "
     "conserve Ghodaghodi Lake:\n\n"
     "'e. Organize campaigns to collect and pick out wastes like plastic bottles, bags and other "
     "sorts of wastes following over surface and under surface of water of Ghodaghodi Lake.'\n\n"
     "Chapter 2 — Key Expectations (p.24):\n"
     "'Conservation from agro-chemicals, pesticides, fertilizer, plastic waste and zoning of "
     "core area.'\n\n"
     "SWOT Analysis — Threats (p.29):\n"
     "'Increasing use of plastic waste and rising water pollution'\n\n"
     "These provisions target plastic waste that has already entered or threatens to enter the "
     "lake environment (environmental leakage prevention and remediation) — plastic bottles "
     "and bags in lake water are specifically identified as requiring organized clean-up campaigns."),
  S=0,
  T=0.25,
  U=("Responsible authority (+0.25): CGLTDB is implicitly responsible for organizing campaigns "
     "as the GLA management body. The broader LFA on 'Social Mobilization and Awareness Raising "
     "Strategic Plan' in the master plan assigns responsibilities for awareness activities to CGLTDB "
     "and community organizations.\n\n"
     "Enforcement (−0): No enforcement provisions for lake plastic pollution.\n\n"
     "Monitoring (−0): No specific monitoring mechanism for lake plastic waste stated in the "
     "relevant provisions.\n\n"
     "Unconditional (−0): Cleaning campaigns are described as activities to be organized — "
     "their frequency, scope, and conditions are not specified."),
  W=("P = 0.40 (Information & voluntary — awareness campaigns and clean-up events): The provision "
     "to 'organize campaigns to collect and pick out wastes like plastic bottles, bags' is an "
     "awareness/clean-up campaign instrument, typical of information & voluntary type. It directly "
     "names plastic materials (criterion a).\n\n"
     "S = 0: Planning language — 'Organize campaigns' is a future commitment in a master plan, "
     "not an operative regulatory standard.\n\n"
     "T = 0.25: Only responsible authority sub-score credited (CGLTDB implicit). No enforcement, "
     "monitoring, or unconditional sub-scores met at the level of this specific provision.\n\n"
     "SIGNIFICANCE: This is one of the most geographically specific plastic pollution provisions "
     "in Nepal's policy landscape — explicitly targeting plastic bottles and bags in the water of "
     "Ghodaghodi Lake, a Ramsar Wetland site. The provision for both surface AND underwater "
     "clean-up campaigns acknowledges the depth of plastic pollution already in this ecologically "
     "important lake.\n\n"
     "The three plastic-relevant passages (key expectations p.24, SWOT threat p.29, lake clean-up "
     "p.66) are coded as one instrument since they all relate to the same preventive/remediation "
     "objective of protecting the lake from plastic pollution."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Pragmatic Garbage Management System (two-type dustbins + recycling priority + annual campaigns) Section 7.2.1",
 "Instrument 2 — Clean-up campaigns for plastic bottles/bags in Ghodaghodi Lake (surface + underwater) Section 7.2.6",
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
 ws.row_dimensions[r].height=240

for c,w in {1:30,2:40,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:35,15:9,
            16:10,17:18,18:50,19:10,20:10,21:50,22:9,23:50,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Context")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Comprehensive Master Plan of Ghodaghodi Lake Area (2077 BS)"),
 ("Year","2077 BS ≈ 2020/21 CE (posted on municipal website August 2022)"),
 ("Prepared by","Comprehensive Ghodaghodi Lake and Tourism Development Board (CGLTDB) under Ghodaghodi Municipality"),
 ("Area","Ghodaghodi Lake Area (GLA), Ghodaghodi Municipality, Kailali, Sudurpaschim Province"),
 ("Lake significance",
  "Ramsar Wetland Site (designated 13 August 2003); IBBA (Important Bird & Biodiversity Area); "
  "KBA (Key Biodiversity Area); part of Basanta transboundary wildlife corridor; cluster of 24 lakes (147.17 ha)"),
 ("Municipal URL","https://ghodaghodimun.gov.np/sites/ghodaghodimun.gov.np/files/Comprehensive%20Master%20Plan%20of%20GLA.pdf"),
 ("FAOLEX URL","https://faolex.fao.org/docs/pdf/nep220194.pdf"),
 ("Legal basis","Local Government Operation Act, 2074 (LGOA 2017)"),
 ("",""),
 ("KEY PLASTIC-RELEVANT TEXT (verbatim)",None),
 ("Location","Text"),
 ("Chapter 2 — Key Expectations (p.24)",
  "'Conservation from agro-chemicals, pesticides, fertilizer, plastic waste and zoning of core area'"),
 ("SWOT Analysis — Threats (p.29)",
  "'Increasing use of plastic waste and rising water pollution'"),
 ("Section 7.2.1 — Garbage Management (p.51-52)",
  "'Establish a Pragmatic Garbage Management System'\n"
  "'Arrange hardcore items [put two types of dustbins (one each for degradable and non-degradable)] "
  "at key junctions including frontal entry gate, religious complex areas, hiking trails'\n"
  "'Put in and apply the system of timely delivery of collected garbage to the landfilled site or "
  "recycling site or composting site. Prioritize to process the garbage (reuse, recycle etc.). "
  "Dumping at landfilled site to be as last option only.'\n"
  "'Agree on soft-core matters (e.g. formulating an annual calendar for cleaning campaign, "
  "collection, segregation and garbage processing schedule)'"),
 ("Section 7.2.6 — Wetland Conservation (p.66) — KEY",
  "'Organize campaigns to collect and pick out wastes like plastic bottles, bags and other sorts "
  "of wastes following over surface and under surface of water of Ghodaghodi Lake.'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (municipal tourism/conservation master plan — aspirational, plan-level)"),
 ("policy_target (E)","0  (no quantifiable plastic-specific targets)"),
 ("policy_integration (I)","1.00  (8 sectors: tourism, conservation/environment, waste management, water/wetlands, agriculture, municipalities, fisheries, infrastructure)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, recycling, environmental leakage — 4 lifecycle phases)"),
 ("policy_budget (M)","0.50  (LFA budget tables + infrastructure cost estimates in Annex 2; not ring-fenced for plastic)"),
 ("",""),
 ("Instr. 1: Garbage Management System (Section 7.2.1)","P=0.80 | S=0 | T=0.75 | V=[auto 0.775]"),
 ("Instr. 2: Lake plastic clean-up campaigns (Section 7.2.6)","P=0.40 | S=0 | T=0.25 | V=[auto 0.325]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Text","Value"):
  for ci2,v in enumerate(row,1):
   if v:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
    c.alignment=Alignment(wrap_text=True); c.border=tb()
 else:
  for ci2,v in enumerate(row,1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
 ws2.row_dimensions[ri].height=55
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — plan commitment"),(1,"In force — mandatory/operative"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T)",None),
     ("Sub-score","Criterion"),
     ("+0.25","Responsible authority"),("+0.25","Enforcement"),
     ("+0.25","Monitoring mechanism"),("+0.25","Unconditional")]
for ri,row in enumerate(ref,1):
 a,b=row[0],row[1] if len(row)>1 else None
 ca=ws3.cell(row=ri,column=1,value=a)
 if b is None and a:
  ws3.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  ca.fill=hfill(MID_BLUE); ca.font=hfont(bold=True,size=10,color=WHITE); ca.alignment=CA
 elif b in("Description","Criterion"):
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

path="/workspace/4p_index_ghodaghodi_lake_masterplan_2020.xlsx"
wb.save(path)
print(f"Saved: {path}")
