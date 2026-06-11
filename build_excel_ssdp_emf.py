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
c.value=("Policy: Environmental Management Framework — Nepal: School Sector Development Plan (SSDP)  |  "
         "Country: Nepal  |  Year: 2017  |  "
         "Source: https://www.doe.gov.np/assets/uploads/files/bf1832e3ce7c66e703b31bfc77c0e363.pdf")
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
 name=("Environmental Management Framework\n"
       "Nepal: School Sector Development Plan (SSDP)\n"
       "[Ministry of Education, Science and Technology / CEHRD, Government of Nepal]\n"
       "SSDP Program Period: 2016–2023"),
 url=("https://www.doe.gov.np/assets/uploads/files/bf1832e3ce7c66e703b31bfc77c0e363.pdf\n"
      "[Department of Education / Center for Education and Human Resource Development website]"),
 year=2017,
 obj=("Environmental Management Framework (EMF) for the School Sector Development Plan "
      "(SSDP, 2016–2023), financed by the Government of Nepal and nine Joint Financing "
      "Partners (World Bank, ADB, UNICEF, EU, GPE, Australia, JICA, Finland, Norway). "
      "Prepared by the Ministry of Education, Science and Technology (MOEST) and Centre "
      "for Education and Human Resource Development (CEHRD). In relation to plastics: "
      "Table IV.A (Mitigation Measures) includes an EXPLICIT BAN on the use of plastic "
      "products in schools ('Ban on the use of plastic products in schools'); mandates "
      "proper solid waste management with waste segregation, composting, and recycling "
      "in all schools; requires daily monitoring of solid waste segregation during "
      "construction; and annual monitoring of solid waste management systems during "
      "school operation. References the Solid Waste Management Act 2068 as applicable law."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Environmental Management Framework (EMF) prepared by the Government of Nepal "
          "(MOEST/CEHRD) as a safeguard document for the World Bank/ADB/multi-partner "
          "financed School Sector Development Plan (SSDP, 2016–2023). Not enacted by "
          "Parliament and not a national regulation/executive decree — it is a project-level "
          "environmental governance/safeguard framework binding on all SSDP implementing "
          "schools and agencies through contractual conditions in bidding documents.\n\n"
          "NOTE: This is a DIFFERENT document from the SSRP EMF (2009) already coded in "
          "this dataset. The SSRP EMF (2009) used advisory language ('discouraging use of "
          "plastic products'). This SSDP EMF uses MANDATORY language: 'Ban on the use of "
          "plastic products in schools' — significantly stronger instrument (P=1.0 vs 0.40)."),
 intg=0.75,
 sects="education, waste management, water/sanitation, health, construction/infrastructure",
 circ=0.75,
 lc="consumption, disposal, recycling, environmental leakage",
 budg=0.5,
 budg_t=("Page 24: 'environmental management cost shall be part of the detailed design "
         "and the project cost.' Page 28 (Table F): 'EMP requirements clearly defined in "
         "BOQ and contract agreement.' The EMF implementation budget is included in "
         "project/school construction costs — not ring-fenced specifically for plastic, "
         "but dedicated environmental management budget exists. M = 0.5."),
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Table IV.A — Anticipated Environmental Impacts: Mitigation Measures (p.19):\n\n"
     "Solid waste management category:\n"
     "Potential impact: 'Spreading of waste, pungent smell, deterioration of aesthetics; "
     "Use batteries, laboratory chemicals disposed haphazardly; Leachate of hazardous "
     "waste in soil and water.'\n\n"
     "Mitigation Measures:\n"
     "1. 'Proper solid waste management system shall be introduced in schools with "
     "segregation of waste, and its proper disposal'\n"
     "2. 'Encourage composting to use in school garden'\n"
     "3. 'Awareness raising on solid waste management with waste minimization, recovery "
     "and recycling system established in the school'\n"
     "4. **'BAN ON THE USE OF PLASTIC PRODUCTS IN SCHOOLS.'** — explicit prohibition\n"
     "5. 'Safe disposal of hazardous waste.'\n\n"
     "This is a mandatory performance standard embedded in the school construction and "
     "operation framework — all schools under SSDP must implement a plastic ban as a "
     "mitigation measure for solid waste management impacts."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): 'School Management Committees and Parent Teacher "
     "Association are fully responsible to implement the work under technical supervision "
     "of the Implementing Agency (IA).' (p.27) MOEST as Executing Agency (EA); CEHRD "
     "at central level; Education Development Coordination Unit (EDCU) at district level. "
     "Each level has designated safeguard focal persons.\n\n"
     "Enforcement (+0.25): Page 27: 'Consequences for failing to comply with safeguard "
     "requirements will be clearly stated in the contract documents.' EMP provisions are "
     "included in bidding documents and BOQ — non-compliance has contractual consequences.\n\n"
     "Monitoring (+0.25): Construction phase (p.31): 'Solid waste segregation: Direct "
     "Observation, Every Day, Contractor/Environment consultant/SMC.' "
     "Operation phase (p.33): 'Solid waste management system: Records of waste collected "
     "and managed, Annual, SMC/IA.' Semi-annual compliance monitoring reports to JFPs.\n\n"
     "Unconditional (+0.25): The ban and waste management requirements apply to ALL "
     "schools under SSDP without stated exemptions. 'Ban on the use of plastic products "
     "in schools' uses unqualified prohibitory language."),
  W=("S = 1 (IN FORCE): 'Ban on the use of plastic products in schools' uses present-tense "
     "prohibitory language as a MANDATORY mitigation measure within the EMF. All schools "
     "under SSDP are contractually bound by this through bidding documents. Compare with "
     "SSRP EMF (2009) which used 'discouraging use of plastic products' (advisory, S=0, "
     "P=0.40) — this SSDP EMF is significantly stronger.\n\n"
     "T = 1.0: All four sub-scores met — responsible authority (MOEST/CEHRD/SMCs), "
     "enforcement (contractual consequences in bidding documents), monitoring (daily "
     "during construction + annual during operation), unconditional (no exemptions).\n\n"
     "SIGNIFICANCE: This is arguably the STRONGEST plastic instrument in Nepal's education "
     "sector policy landscape — an explicit ban on ALL plastic products in ALL SSDP schools, "
     "with daily monitoring during construction. Covers 365 integrated schools, 38,000+ "
     "classrooms, and 21,000+ school blocks across Nepal.\n\n"
     "COMPARISON WITH SSRP EMF (already coded):\n"
     "• SSRP EMF 2009: 'discouraging use of plastic products' → P=0.40 (information), S=0, T=0.50\n"
     "• SSDP EMF 2017: 'Ban on the use of plastic products in schools' → P=1.0 (regulatory), S=1, T=1.0\n"
     "This represents a significant strengthening of the education sector's plastic policy."),
 ),
 dict(
  P=0.20, Q="Waste management",
  R=("SECTION F — Institutional Arrangement and Mechanism (p.27-28):\n\n"
     "GOVERNANCE STRUCTURE:\n"
     "EA (MOEST), IA (CEHRD/EDCU), SMCs, and management support consultants are "
     "designated with specific EMP roles. Each level has a safeguard focal person. "
     "EMP provisions are incorporated into school improvement plans (SIPs), bidding "
     "documents, and BOQs.\n\n"
     "MONITORING FRAMEWORK (pp.31, 33):\n"
     "Construction phase: 'Solid waste segregation' — monitored Every Day by "
     "Contractor/Environment consultant/SMC.\n"
     "Operation phase: 'Solid waste management system: Records of waste collected and "
     "managed' — monitored Annually by SMC/IA.\n\n"
     "REPORTING:\n"
     "Semi-annual environmental compliance monitoring reports submitted to JFPs during "
     "Joint Review Meeting (JRM) and Budget Review Meeting (BRM). EA conducts annual "
     "review of environmental safeguards performance."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): EA (MOEST), IA (CEHRD/EDCU), and SMCs explicitly "
     "designated throughout the institutional arrangement section (pp.27-29).\n\n"
     "Enforcement (+0.25): Contractual enforcement through bidding documents and BOQ; "
     "'Consequences for failing to comply with safeguard requirements will be clearly "
     "stated in the contract documents' (p.27). JFPs can withhold financing for non-"
     "compliance.\n\n"
     "Monitoring (+0.25): Explicit daily and annual monitoring schedules for solid waste "
     "management in Tables 8.1 and 8.2 (pp.31, 33). Semi-annual compliance monitoring "
     "reports submitted to JFPs.\n\n"
     "Unconditional (−0): The governance framework applies to all SSDP schools — but "
     "it is a framework requiring further implementation by each school (some discretion "
     "in how SMCs implement). Not fully unconditional at the individual school level."),
  W=("P = 0.20 (Governance & coordination — planning requirement, institutional setup): "
     "The governance and monitoring framework creates coordination obligations across EA, "
     "IA, SMC, management consultants, and JFPs for environmental safeguards implementation "
     "including waste (plastic) management.\n\n"
     "S = 1: The institutional arrangement is operative — MOEST/CEHRD are designated and "
     "functioning; monitoring requirements are in force through contractual conditions.\n\n"
     "T = 0.75: Authority, enforcement, and monitoring credited. Unconditional not fully "
     "credited because school-level implementation has some discretion in execution details."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — EXPLICIT BAN on plastic products in schools + solid waste management with segregation/recycling/composting (Mitigation Measures Table, p.19)",
 "Instrument 2 — EMF governance + daily solid waste monitoring framework (daily construction + annual operation monitoring, pp.31/33)",
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

for c,w in {1:32,2:42,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:30,11:8,12:28,13:8,14:35,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Comparison")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Title","Environmental Management Framework — Nepal: School Sector Development Plan (SSDP)"),
 ("Program","School Sector Development Plan (SSDP), 2016–2023"),
 ("Executing Agency","Ministry of Education, Science and Technology (MOEST)"),
 ("Implementing Agency","Centre for Education and Human Resource Development (CEHRD)"),
 ("Joint Financing Partners","World Bank, ADB, UNICEF, European Union, GPE, Australia, JICA, Finland, Norway"),
 ("URL","https://www.doe.gov.np/assets/uploads/files/bf1832e3ce7c66e703b31bfc77c0e363.pdf"),
 ("Year","~2017 (SSDP program period 2016–2023; references Three Year Plan 2017-2020)"),
 ("",""),
 ("KEY PLASTIC PROVISION — VERBATIM (p.19)",None),
 ("Section","Text"),
 ("Solid waste management mitigation — plastic ban",
  "'Ban on the use of plastic products in schools.' [EXACT QUOTE, p.19, Mitigation Measures Table]"),
 ("Full mitigation package",
  "1. 'Proper solid waste management system shall be introduced in schools with segregation of waste, and its proper disposal'\n"
  "2. 'Encourage composting to use in school garden'\n"
  "3. 'Awareness raising on solid waste management with waste minimization, recovery and recycling system established in the school'\n"
  "4. 'Ban on the use of plastic products in schools.'\n"
  "5. 'Safe disposal of hazardous waste.'"),
 ("Monitoring — construction (p.31)",
  "'Solid waste segregation: Direct Observation, Every Day, Contractor/Environment consultant/SMC'"),
 ("Monitoring — operation (p.33)",
  "'Solid waste management system: Records of waste collected and managed, Annual, SMC/IA'"),
 ("Enforcement (p.27)",
  "'Consequences for failing to comply with safeguard requirements will be clearly stated in the contract documents.'"),
 ("",""),
 ("COMPARISON WITH SSRP EMF 2009 (already coded)",None),
 ("Aspect","SSRP EMF (2009)","SSDP EMF (this document, ~2017)"),
 ("Plastic provision","'discouraging use of plastic products'","'**Ban on the use of plastic products in schools**'"),
 ("Language","Advisory ('should be promoted')","Mandatory ban (present tense prohibition)"),
 ("P (instrument_type)","0.40 (Information & voluntary)","**1.00 (Regulatory — ban)**"),
 ("S (in_force)","0 (advisory)","**1 (mandatory)**"),
 ("T (implementation)","0.50","**1.00 (all 4 sub-scores met)**"),
 ("V (instrument_score)","0.450","**1.000**"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (project-level EMF — not national legislation; similar to SSRP EMF and BBIN ESMF already coded)"),
 ("policy_target (E)","0  (no quantifiable plastic-specific targets)"),
 ("policy_integration (I)","0.75  (5 sectors: education, waste management, water/sanitation, health, construction/infrastructure)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, recycling, environmental leakage — 4 phases)"),
 ("policy_budget (M)","0.50  (EMP costs in BOQ + project budget; not ring-fenced for plastic)"),
 ("",""),
 ("Instr. 1: PLASTIC BAN + solid waste mgmt + segregation/recycling/composting","P=1.00 | S=1 | T=1.00 | V=[auto **1.000**]"),
 ("Instr. 2: EMF governance + daily solid waste monitoring framework","P=0.20 | S=1 | T=0.75 | V=[auto 0.475]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=3)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif len(row)==3 and row[1]=="SSRP EMF (2009)":
  for ci2,v in enumerate(row,1):
   if v:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
    c.alignment=Alignment(wrap_text=True); c.border=tb()
 elif len(row)==3 and row[2] is not None:
  for ci2,v in enumerate(row,1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
 elif row[1] in("Details","Text","Section","Value"):
  for ci2,v in enumerate(row,1):
   if v:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
    c.alignment=Alignment(wrap_text=True); c.border=tb()
  if len(row)==2: ws2.merge_cells(start_row=ri,start_column=2,end_row=ri,end_column=3)
 else:
  for ci2,v in enumerate(row,1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
  if len(row)==2: ws2.merge_cells(start_row=ri,start_column=2,end_row=ri,end_column=3)
 ws2.row_dimensions[ri].height=55
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=45
ws2.column_dimensions["C"].width=45

ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational [used for project-level EMF documents]"),
     (0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, standard)"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T)",None),
     ("Sub-score","Criterion"),
     ("+0.25","Responsible authority"),("+0.25","Enforcement"),("+0.25","Monitoring"),("+0.25","Unconditional")]
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

path="/workspace/4p_index_ssdp_emf_2017.xlsx"
wb.save(path)
print(f"Saved: {path}")
