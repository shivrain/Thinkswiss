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
c.value=("Policy: Environmental Management Framework for School Sector Reform Plan (SSRP) Nepal  |  "
         "Country: Nepal  |  Year: 2009  |  Source: https://www.doe.gov.np/a... [URL truncated]")
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
 name="Environmental Management Framework for School Sector Reform Plan (SSRP) Nepal",
 url=("https://www.doe.gov.np/a... [URL truncated in source image — "
      "published by Government of Nepal, Department of Education (DOE), May 2009]"),
 year=2009,
 obj=("Environmental Management Framework (EMF) for the School Sector Reform Plan (SSRP) "
      "prepared by Nepal's Department of Education (DOE) in May 2009, with financing from "
      "development partners. Establishes environmental safeguard procedures for school "
      "physical infrastructure development (construction, rehabilitation). In relation to "
      "plastics: Chapter III Mitigation Measures mandate that 'awareness regarding the "
      "management of solid waste, discouraging use of plastic products etc should be promoted "
      "in schools' and 'use of recycled materials should be promoted'; the Monitoring "
      "Framework (Table p.24) requires daily SMC monitoring of solid waste segregation during "
      "school construction; and the design guidelines note that students throwing 'solid waste "
      "(paper, plastics etc.) in toilet resulting in blocking' is a common problem requiring "
      "management attention."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Environmental Management Framework (EMF) prepared in May 2009 by the Department of "
          "Education (DOE), Government of Nepal, as a safeguard document for the School Sector "
          "Reform Plan (SSRP), with financing from development partners (World Bank and others "
          "referenced as DPs/development partners). The EMF is a project-level environmental "
          "governance/safeguard document — not national legislation and not a stand-alone "
          "regulation or executive decree. It creates environmental management obligations "
          "through the program implementation framework, with the DOE designated as the "
          "primary institutional authority. G = 0.25 (strategy/plan with aspirational "
          "commitments) — the plastic-relevant provisions use predominantly advisory language "
          "('should be promoted', 'should be introduced') within an environmental management "
          "planning framework. No quantifiable plastic-specific targets."),
 intg=0.75,
 sects="education, waste management, water/sanitation, construction/infrastructure, health",
 circ=0.75,
 lc="consumption, disposal, recycling",
 budg=0.5,
 budg_t=("'GON/MOE will have a full-time Environmental Officer at DOE to look after the EMF "
         "activities for SSRP. Additional human resources or agency will be hired, if necessary, "
         "in order to effectively implement the EMF.' (Executive Summary, p.3). Development "
         "partners (DPs) finance the SSRP program which includes the EMF implementation costs. "
         "Budget exists for overall program implementation including EMF but not ring-fenced "
         "specifically for plastic waste management. M = 0.5 (funding source mentioned; not "
         "ring-fenced for plastic)."),
)

# Single row (one instrument)
r = 5
ws.cell(row=r,column=24,value=(
    "Instrument 1 — Solid waste management guidance (discouraging plastic product use + recycling promotion) "
    "+ mandatory daily monitoring of solid waste segregation during construction (Chapter III, Tables p.21 and p.24)"
)).font=hfont(italic=True,size=8,color="808080")

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

# Instrument
ifl=hfill("E2EFDA")
idata=[
 0.40,  # P
 "Waste management",  # Q
 ("Chapter III — Mitigation Measures Table (p.21): Solid Waste category:\n"
  "'Proper solid waste management system should be introduced in schools which includes "
  "segregation of waste, and its proper disposal. The environmentally friendly management "
  "measures like composting should be encouraged. Awareness regarding the management of "
  "solid waste, **discouraging use of plastic products** etc should be promoted in schools. "
  "The use of recycled materials should be promoted.'\n\n"
  "Chapter III — Table of Environmental Issues (p.13): S.N.5 Wastes:\n"
  "'Solid waste should be disposed of properly. Key strategies for improving solid waste "
  "management and disposal are to minimize the waste by schools and to recycle waste "
  "whenever possible.'\n\n"
  "Monitoring Framework (p.24): Item 6 — 'Solid waste segregation disposal': monitored "
  "by SMC every day during construction using CM/Direct Observation method.\n\n"
  "Design Guidelines (p.54): 'students throw carelessly all the solid waste (paper, "
  "plastics etc) in toilet resulting in blocking the trap' — identified as a design "
  "and management problem.\n\n"
  "Environmental Screening Checklist (p.39): 'Material specification — practices of "
  "reduce, reuse, recycle of materials during construction.'"),
 0,  # S
 0.50,  # T
 ("Responsible authority (+0.25): DOE designated as primary implementing authority; "
  "School Management Committee (SMC) designated as the responsible body for monitoring "
  "solid waste segregation at school level (Monitoring Table, p.24, Item 6).\n\n"
  "Enforcement (−0): No explicit fines or penalties for non-compliance with solid waste "
  "management guidance in schools stated in the EMF.\n\n"
  "Monitoring (+0.25): Monitoring Table (p.24) Item 6: 'Solid waste segregation disposal' "
  "— monitored by SMC every day during construction phase using 'CM/Direct Observation' "
  "method. This constitutes an explicitly mandated, regular monitoring mechanism.\n\n"
  "Unconditional (−0): Mitigation measures are advisory ('should be introduced', 'should "
  "be promoted') and the discouragement of plastic products is framed as an awareness "
  "campaign rather than a mandatory ban — exemptions exist by implication since the "
  "guidance is not binding."),
 None,  # V (auto)
 ("S = 0: Mitigation measures use advisory language throughout — 'should be introduced', "
  "'should be promoted', 'should be encouraged.' The plastic discouragement element is "
  "explicitly framed as an awareness/information campaign ('Awareness regarding the management "
  "of solid waste, discouraging use of plastic products etc should be promoted in schools'). "
  "The EMF is a project-level safeguard document, not national legislation.\n\n"
  "T = 0.50: Responsible authority (+0.25: SMC and DOE explicitly designated) and monitoring "
  "(+0.25: daily SMC monitoring of solid waste segregation in Monitoring Table) credited. "
  "Enforcement not credited (no penalties). Unconditional not credited (advisory language).\n\n"
  "P = 0.40 (Information & voluntary): The primary plastic-relevant instrument is an "
  "awareness campaign + guidance to discourage plastic products in schools and promote "
  "recycling. The mandatory daily monitoring of solid waste segregation strengthens the "
  "implementation but does not change the primary instrument type.\n\n"
  "PLASTIC RELEVANCE: This document is primarily about school construction safeguards, "
  "not plastic pollution policy. The plastic-relevant provisions are limited to: (a) advisory "
  "guidance to discourage plastic product use in schools (criterion a — mentions plastics); "
  "(b) solid waste segregation monitoring (criterion c — waste management instrument). "
  "Note: the EMF was prepared in 2009, well before Nepal's plastic bag bans.\n\n"
  "FIELD FINDING (p.32): '80% of the school visited doesn't have any kind of solid waste "
  "management system' — confirming the critical baseline gap this instrument addresses."),
]
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
ws.row_dimensions[r].height=280

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:30,11:8,12:25,13:8,14:35,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Context")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Environmental Management Framework for School Sector Reform Plan (SSRP) Nepal"),
 ("Published","May 2009"),
 ("Prepared by","Government of Nepal, Department of Education (DOE)"),
 ("Program","School Sector Reform Plan (SSRP), financed by development partners including World Bank"),
 ("Document type","Environmental Management Framework (EMF) — project-level environmental safeguard document"),
 ("URL","https://www.doe.gov.np/a... [URL truncated in source image]"),
 ("Total pages","67 pages"),
 ("",""),
 ("KEY PLASTIC/WASTE-RELEVANT TEXT (verbatim)",None),
 ("Location","Text"),
 ("Chapter III, p.21 — KEY (Mitigation Measures)",
  "'Proper solid waste management system should be introduced in schools which includes "
  "segregation of waste, and its proper disposal. The environmentally friendly management "
  "measures like composting should be encouraged. Awareness regarding the management of "
  "solid waste, **discouraging use of plastic products** etc should be promoted in schools. "
  "The use of recycled materials should be promoted.'"),
 ("Chapter II, p.13 — Table row 5 (Wastes)",
  "'Solid waste should be disposed of properly. Key strategies for improving solid waste "
  "management and disposal are to minimize the waste by schools and to recycle waste "
  "whenever possible.'"),
 ("Monitoring Table, p.24 — Item 6 (MANDATORY)",
  "Indicator: 'Solid waste segregation disposal' | Monitoring: 'CM/Direct Observation' "
  "| Frequency: 'Every Day' | Responsibility: 'SMC'"),
 ("Field Visit (p.32)",
  "'Solid waste management: 80% of the school visited doesn't have any kind of solid waste "
  "management system.'"),
 ("Design Guidelines, p.54 (plastic in toilets)",
  "'students throw carelessly all the solid waste (paper, **plastics** etc) in toilet "
  "resulting in blocking the trap'"),
 ("Environmental Screening (p.39)",
  "'Material specification: Use of locally available suitable material, practices of reduce, "
  "reuse, recycle of materials during construction.'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (project-level EMF — not national legislation; advisory guidance language)"),
 ("policy_target (E)","0  (no quantifiable plastic targets)"),
 ("policy_integration (I)","0.75  (5 sectors: education, waste management, water/sanitation, construction/infrastructure, health)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, recycling — 3 lifecycle phases)"),
 ("policy_budget (M)","0.50  (DOE Environmental Officer funded + development partner financing; not ring-fenced for plastic)"),
 ("",""),
 ("Instr. 1: Solid waste management + plastic discouragement (advisory) + daily monitoring (mandatory)",
  "P=0.40 | S=0 | T=0.50 | V=[auto 0.450]"),
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
 ws2.row_dimensions[ri].height=50
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

# Sheet 3 score reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational [used for EMF/project safeguard documents]"),
     (0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory"),
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

path="/workspace/4p_index_school_sector_emf_2009.xlsx"
wb.save(path)
print(f"Saved: {path}")
