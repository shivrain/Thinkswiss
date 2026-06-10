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
c.value=("Policy: BBIN MPA Regional Transport and Trade Facilitation Program — Nepal Phase 1 ESMF (2022)  |  "
         "Country: Nepal  |  Year: 2022  |  "
         "Source: https://dor.gov.np/uploads/publicatio... [URL truncated in source image]")
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
 name=("Bangladesh-Bhutan-India-Nepal (BBIN) Multi-phase Programmatic Approach (MPA) "
       "Regional Transport and Trade Facilitation Program - Nepal Phase 1\n"
       "Environmental and Social Management Framework (ESMF)"),
 url=("https://dor.gov.np/uploads/publicatio... [URL truncated in source image — "
      "document published by Government of Nepal, Ministry of Public Infrastructure and "
      "Transport, Department of Roads and Ministry of Industries, Commerce and Supplies]"),
 year=2022,
 obj=("Environmental and Social Management Framework (ESMF) for the World Bank-financed "
      "BBIN MPA Nepal Phase 1 project, which involves: (1) digital/automated trade and "
      "customs facilitation systems; (2) upgrading of the Butwal-Gorusinghe-Chanauta section "
      "of the East-West Highway from 2 to 4 lanes; (3) construction of a bridge; and (4) "
      "capacity building for trade policy reforms. In relation to plastics: the ESMF mandates "
      "contractors to develop and implement Solid Waste Management Plans (SWMPs) covering "
      "plastic and other construction waste; specifies that contaminated and worn plastic "
      "sheeting used on construction sites must be packed into drums and disposed of off-site; "
      "and requires compliance with World Bank ESS3 Resource Efficiency and Pollution "
      "Prevention standards to prevent plastic waste from contaminating water bodies."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Environmental and Social Management Framework (ESMF) prepared in January 2022 by the "
          "Government of Nepal (Ministry of Public Infrastructure and Transport, Department of "
          "Roads, Ministry of Industries, Commerce and Supplies) as a condition for World Bank "
          "financing. Not enacted by Parliament and not a national regulation/decree — it is a "
          "project-level environmental governance/safeguard framework binding on project "
          "implementers and contractors within the project context only.\n\n"
          "NOTE: This is a PROJECT-LEVEL environmental safeguard document (ESMF) prepared for "
          "a specific infrastructure investment, not a standalone national plastic pollution "
          "policy. It creates binding obligations through contract conditions (ESHS conditions "
          "in bidding documents) rather than through national legislation. Scored 0.25 as the "
          "lowest available category, equivalent to a strategy/plan document.\n\n"
          "Components: Component 1 (trade facilitation — ASYCUDA, border management systems), "
          "Component 2 (East-West Highway upgrade + bridge + green highway concept), "
          "Component 3 (capacity building — trade/transport reforms). Implementing agencies: "
          "Department of Roads (DOR) for road works; Ministry of Industries, Commerce and "
          "Supplies (MOICS) for trade components."),
 intg=0.75,
 sects="transport/infrastructure, waste management, environment, water, occupational health, trade/customs",
 circ=0.75,
 lc="consumption, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Section 8.6 (p.107): 'Budgets for Implementation of ESMF' — the ESMF includes a "
         "dedicated budget section for environmental and social management. The project is "
         "financed by the World Bank (IDA financing). ESMF implementation costs are to be "
         "included in overall project budgets. Section 8.4: 'Capacity Building of DOR and MOICS' "
         "includes budgets for ESMF training and capacity building. Budget exists for overall "
         "ESMF implementation but is not ring-fenced specifically for plastic waste management."),
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=0.20, Q="Waste management",
  R=("Section 5.3 Impacts and Risk Mitigation Measures Table (p.20) — Construction Stage, "
     "Solid Waste Management:\n"
     "'Develop a Solid Waste Management Plan to manage solid wastes during road construction "
     "and operations. Ensure appropriate disposal sites for muck and rock cuttings as well as "
     "proper management of solid and hazardous waste.'\n\n"
     "Responsible agencies: Contractors (implementation), PMUs and CSCs (supervision/monitoring).\n\n"
     "Section 2.6 and Table 2.5 (p.4/20): Environmental, Health and Safety (ESHS) conditions "
     "are made mandatory in all bidding documents, requiring contractors to comply with ESMF "
     "provisions including the SWMP during project execution.\n\n"
     "Also covered by ESS3: Resource Efficiency and Pollution Prevention (Table 3.1, p.39): "
     "'Resource efficiency and pollution prevention in any project activity will be captured "
     "in ESIA/ESMP preparation.' This includes solid waste management to prevent plastic "
     "and other construction waste from entering water bodies (Section 6.3.2, p.76: 'The "
     "construction/maintenance contractor will require developing a waste management plan "
     "which details the use, storage and disposal of toxic, solid and sanitary waste and "
     "materials. No direct spillage of petrochemical or toxic materials should be allowed "
     "into the aquatic ecosystem.')"),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): 'Contractors' designated as implementing agency; "
     "'PMUs and CSCs' designated as supervision/monitoring authorities (Impacts Table, p.20). "
     "Section 8.5 (p.106): 'Roles and Responsibilities of Implementing Agencies' explicitly "
     "assigns ESMF oversight to DOR and MOICS.\n\n"
     "Enforcement (+0.25): Section 2.6 and Table 2.5 state that ESHS conditions are "
     "included in bidding documents — contractual enforcement mechanism. Table 3.2 (p.41): "
     "'The ESMP to be prepared shall be made integral part of bidding document so that the "
     "Contractor shall adhere to the provisions prescribed in the ESMP during execution of "
     "the project.' Non-compliance triggers contractual remedies.\n\n"
     "Monitoring (+0.25): p.20 designates 'PMUs and CSCs' for supervision of SWMP "
     "implementation. Section 7 (Monitoring) and Section 8.2-8.3 describe monitoring "
     "responsibilities of DOR and MOICS. Table 2.4 (p.4): Sample Monitoring Plan included "
     "in ESMF provides monitoring framework for all environmental requirements.\n\n"
     "Unconditional (+0.25): The SWMP requirement applies to ALL construction subprojects "
     "under the BBIN 1 project with no stated exemptions — all contractors must comply."),
  W=("P = 0.20 (Governance & coordination — planning requirement): The ESMF mandates that "
     "contractors develop and implement Solid Waste Management Plans. This is a planning "
     "requirement creating a governance obligation (SWMP = a coordination and management "
     "planning instrument), not a direct regulatory ban or economic incentive.\n\n"
     "This instrument passes the plastics relevance filter under criterion (c): 'waste "
     "management or governance instrument that commonly applies to plastic waste streams "
     "in practice' — specifically waste hierarchy obligations and construction waste "
     "management planning.\n\n"
     "S = 1: The ESMF is in force as a project safeguard document — all contractors are "
     "contractually bound to comply through ESHS conditions in bidding documents. The ESMF "
     "uses mandatory language throughout ('shall', 'will be required to', 'must').\n\n"
     "T = 1.0: All four sub-scores explicitly evidenced (authority, enforcement, monitoring, "
     "unconditional). High T despite low P (0.20) because the ESMF's governance mechanisms "
     "are very well-specified with clear responsibilities, contractual enforcement, and "
     "monitoring frameworks."),
 ),
 dict(
  P=1.0, Q="Waste management",
  R=("Section 6.3.2.13 — Avoiding Hazards Caused by Storage of Hazardous Materials "
     "(p.74): Mitigative Measures include:\n"
     "'Hazardous materials should be stored only on impervious (concrete or plastic sheeting "
     "as approved by Engineer) floor with drainage and collection sump so as to retain leaks "
     "and spills.'\n\n"
     "'Contaminated and worn plastic sheeting shall be packed into drums and disposed off "
     "site.'\n\n"
     "This provision explicitly names plastic materials (plastic sheeting used as impervious "
     "flooring for hazardous materials storage) and mandates specific waste disposal actions: "
     "all contaminated and worn plastic sheeting used on construction sites must be collected, "
     "packed into drums, and disposed of at an approved off-site location — preventing plastic "
     "pollution of construction site environments and surrounding water bodies."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Contractors responsible for implementation (implied by "
     "the context of Section 6.3 Mitigative Measures for contractors); CSCs as supervision "
     "authority (general ESMF monitoring framework — p.20 designates PMUs/CSCs for all "
     "mitigation measure oversight).\n\n"
     "Enforcement (+0.25): ESHS conditions in bidding documents contractually bind "
     "contractors to all mitigation measures in the ESMF, including plastic sheeting "
     "disposal (Section 2.6). Non-compliance constitutes a breach of contract.\n\n"
     "Monitoring (+0.25): PMUs and CSCs monitor implementation of all Section 6 mitigative "
     "measures, including hazardous materials management (p.20-21). Compliance with "
     "hazardous materials management plans is explicitly monitored.\n\n"
     "Unconditional (+0.25): 'Contaminated and worn plastic sheeting shall be packed into "
     "drums and disposed off site' — no exceptions stated; mandatory language ('shall'); "
     "applies to all plastic sheeting used on all construction sites under the project."),
  W=("P = 1.0 (Regulatory — mandatory disposal standard for plastic construction materials): "
     "The provision uses mandatory language ('shall be packed into drums and disposed off "
     "site') for a specific plastic material (plastic sheeting). This constitutes a mandatory "
     "performance/disposal standard for plastic waste from construction activities.\n\n"
     "This instrument passes the plastics relevance filter under criterion (a): explicitly "
     "mentions 'plastic sheeting' as a construction material requiring specific waste "
     "management treatment.\n\n"
     "S = 1: Mandatory language 'shall be packed into drums and disposed off site' — "
     "language test confirms in-force status.\n\n"
     "T = 1.0: All four sub-scores met (authority via overall ESMF structure, enforcement "
     "via bidding document conditions, monitoring via PMU/CSC oversight, unconditional since "
     "no exemptions stated for this specific provision).\n\n"
     "CONTEXT: The provision is in the context of preventing environmental contamination "
     "from hazardous materials (e.g., chemicals, bitumen, oils) used in road construction. "
     "Plastic sheeting is used as an impervious lining under hazardous material storage "
     "areas; when contaminated, it becomes a plastic waste item requiring controlled disposal. "
     "This prevents contaminated plastic waste from being discarded in the natural environment "
     "around the construction site — preventing plastic pollution of roadside areas and "
     "waterways (environmental leakage prevention)."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Mandatory Solid Waste Management Plan (SWMP) + ESS3 pollution prevention for all construction subprojects",
 "Instrument 2 — Mandatory proper disposal of contaminated/worn plastic sheeting from construction sites (Section 6.3.2.13)",
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
 ws.row_dimensions[r].height=260

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:35,15:9,
            16:10,17:18,18:50,19:10,20:10,21:50,22:9,23:50,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Project Overview")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title",
  "Bangladesh-Bhutan-India-Nepal (BBIN) Multi-phase Programmatic Approach (MPA) Regional "
  "Transport and Trade Facilitation Program - Nepal Phase 1: Environmental and Social "
  "Management Framework (ESMF)"),
 ("Published","January 2022"),
 ("Prepared by",
  "Government of Nepal: Ministry of Public Infrastructure and Transport; "
  "Department of Roads (DOR); Ministry of Industries, Commerce and Supplies (MOICS)"),
 ("Financed by","World Bank (IDA financing)"),
 ("Document type","Environmental and Social Management Framework (ESMF) — project safeguard document"),
 ("",""),
 ("PROJECT COMPONENTS",None),
 ("Component","Description"),
 ("Component 1: Trade facilitation",
  "(a) Development of ASYCUDA World new modules; "
  "(b) Development of Automated Border Management System (ABMS)"),
 ("Component 2: Transport infrastructure",
  "(a) Upgrading Butwal-Gorusinghe-Chanauta section of East-West Highway from 2 to 4 lanes; "
  "(b) Construction of signature bridge; "
  "(c) Green resilient highway concept; "
  "(d) Upgradation of trade facilities"),
 ("Component 3: Capacity building",
  "(a) Transport and trade reforms; (b) Local community support; "
  "(c) Training programs; (d) Project preparation studies"),
 ("",""),
 ("KEY PLASTIC-RELEVANT TEXT (verbatim)",None),
 ("Location","Text"),
 ("Section 5.3, p.20 — SWMP requirement",
  "'Solid waste management: Develop a Solid Waste Management Plan to manage solid wastes "
  "during road construction and operations. Ensure appropriate disposal sites for muck and "
  "rock cuttings as well as proper management of solid and hazardous waste.' "
  "[Contractors: implementation; PMUs and CSCs: supervision]"),
 ("Section 6.3.2, p.76 — No spillage",
  "'The construction/maintenance contractor will require developing a waste management plan "
  "which details the use, storage and disposal of toxic, solid and sanitary waste and "
  "materials. No direct spillage of petrochemical or toxic materials should be allowed into "
  "the aquatic ecosystem.'"),
 ("Section 6.3.2.13, p.74 — Plastic sheeting disposal (KEY)",
  "'Contaminated and worn plastic sheeting shall be packed into drums and disposed off site.'\n"
  "'Hazardous materials should be stored only on impervious (concrete or plastic sheeting "
  "as approved by Engineer) floor with drainage and collection sump so as to retain leaks "
  "and spills.'"),
 ("Table 3.2 / ESS3, p.41 — Pollution prevention",
  "'Resource efficiency and pollution prevention in any project activity will be captured in "
  "ESIA/ESMP preparation. WBG EHS guidelines or/national standards (depending on which one "
  "is stricter) related to environmental protection and resource efficiency will be complied "
  "with by the project.'"),
 ("Table 2.1 / ESHS in bidding docs, p.4",
  "'ESHS Conditions in the Bidding Documents — The ESMP to be prepared shall be made "
  "integral part of bidding document so that the Contractor shall adhere to the provisions "
  "prescribed in the ESMP during execution of the project.'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("CRITICAL NOTE",
  "This is a PROJECT-LEVEL ESMF (environmental safeguard document for a World Bank-funded "
  "infrastructure project), not a standalone national plastic pollution policy. Its plastic-"
  "relevant instruments operate through contractual enforcement mechanisms (bidding document "
  "conditions), not through national legislation or regulation."),
 ("policy_type (G)","0.25  (project-level safeguard/governance framework — not national legislation)"),
 ("policy_target (E)","0  (no quantifiable plastic-specific targets)"),
 ("policy_integration (I)","0.75  (6 sectors: transport/infrastructure, waste management, environment, water, occupational health, trade/customs)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, environmental leakage — 3 lifecycle phases)"),
 ("policy_budget (M)","0.50  (World Bank IDA financing + ESMF implementation budget in Section 8.6; not ring-fenced for plastic)"),
 ("",""),
 ("Instr. 1: Mandatory SWMP + ESS3 pollution prevention","P=0.20 | S=1 | T=1.00 | V=[auto 0.600]"),
 ("Instr. 2: Mandatory disposal of contaminated plastic sheeting (Sec.6.3.2.13)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Description","Text","Value"):
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
 ws2.row_dimensions[ri].height=50
ws2.column_dimensions["A"].width=38
ws2.column_dimensions["B"].width=82

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational [used here for project-level ESMF]"),
     (0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination (planning requirement, SWMP)"),
     (0.40,"Information & voluntary"),(0.60,"Economic"),
     (0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, mandatory disposal standard)"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — enabling power not exercised"),
     (1,"In force — mandatory language (shall/must) or operationalised via contract conditions"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T)",None),
     ("Sub-score","Criterion (must be explicitly evidenced)"),
     ("+0.25","Responsible authority designated"),("+0.25","Enforcement (contract conditions/penalties)"),
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

path="/workspace/4p_index_bbin_mpa_esmf_2022.xlsx"
wb.save(path)
print(f"Saved: {path}")
