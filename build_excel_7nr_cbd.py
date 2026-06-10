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
c.value=("Policy: Nepal's Seventh National Report (7NR) to the Convention on Biological Diversity (CBD), 2026  |  "
         "Country: Nepal  |  Year: 2026  |  "
         "Source: https://giwmscdntwo.gov.np/media/p... [URL truncated in source image]")
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
 name="Nepal's Seventh National Report (7NR) to the Convention on Biological Diversity (CBD), 2026",
 url=("https://giwmscdntwo.gov.np/media/p... [URL truncated in source image — "
      "document hosted on giwmscdntwo.gov.np (DNPWC/Government of Nepal portal)]"),
 year=2026,
 obj=("Nepal's seventh biennial national report (February 2026) submitted to the Convention on "
      "Biological Diversity (CBD), documenting progress against Nepal's NBSAP 2025-2030 and "
      "the Kunming-Montreal Global Biodiversity Framework (KM-GBF). In relation to plastics: "
      "reports progress against National Biodiversity Target 9 (reduce plastic pollution to "
      "levels not harmful to biodiversity by 2030, with Indicator 9.5 targeting reduction of "
      "plastic use from 2.7 g/day/capita [2015] to 0 g/day/capita [2030]), and National "
      "Biodiversity Target 15 (develop a supportive framework for sustainable consumption "
      "including circular economy approaches by 2028). The report rates plastic-related "
      "progress as 'progress made but at an insufficient rate' due to weak enforcement. "
      "Microplastics have been detected in Nepal's rivers and human blood."),
 tgt=1,
 tgt_t=("National Biodiversity Target 9, Indicator 9.5 (Annex 3.9, p.76): "
        "'By 2030, the extent of plastics pollution is reduced.' "
        "Use of plastics indicator: 2.7 g/day/capita (2015 baseline) → 0.9 g/day/capita "
        "(2028 milestone) → 0 g/day/capita (2030 target). Lead agency: DoE/MoFE. "
        "[Quoted directly from Monitoring Framework table, p.76]"),
 gtype=0.50,
 gtype_j=("National reporting document submitted to the Convention on Biological Diversity "
          "(CBD) in February 2026. Prepared by the Environment and Biodiversity Division, "
          "Ministry of Forests and Environment (MoFE), with technical and financial support "
          "from GEF (EAS Project, Umbrella Project) and BIOFIN/UNDP. Not enacted by "
          "Parliament; constitutes a strategy/plan document. The national biodiversity targets "
          "reported herein were formally endorsed through a ministerial-level decision and "
          "multi-stakeholder mechanism during the NBSAP process. Scored 0.50 (not 0.25) "
          "because the document contains quantifiable targets with defined baselines, milestones, "
          "and lead agencies — specifically Indicator 9.5 (plastic use: 2.7 → 0 g/day/capita "
          "by 2030) in the NBSAP Monitoring Framework (Annex 3.9)."),
 intg=1,
 sects="agriculture, waste management, conservation, water, forestry, fisheries, municipalities, tourism, industry, chemicals",
 circ=0.75,
 lc="consumption, recycling, disposal, environmental leakage",
 budg=0.5,
 budg_t=("'This report was prepared with technical and financial assistance from the GEF Early "
         "Action Support (EAS) Project and the GEF Umbrella Project, together with additional "
         "technical support from the Biodiversity Finance Initiative (BIOFIN) Project of UNDP.' "
         "(p.ii) International climate finance access also referenced. Budget is for overall NBSAP "
         "preparation and biodiversity management — not ring-fenced specifically for plastic "
         "pollution reduction. M = 0.5 (budget/funding source mentioned, not ring-fenced)."),
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=0.20, Q="Environmental leakage",
  R=("Annex 3.9 (pp. 73–76): National Biodiversity Target 9 (NBT 9), aligned with KM-GBF "
     "Target 7: 'By 2030, reduce impacts of pollution from all sources, especially from "
     "plastics, pesticides, wastewater, and nutrients, to levels that are not harmful to "
     "biodiversity, especially in areas of high importance for biodiversity.'\n\n"
     "Actions for plastics under NBT 9: '(c) reducing, reusing, and recycling plastics'\n\n"
     "Monitoring Indicator 9.5 (NBSAP Monitoring Framework, p.76):\n"
     "'By 2030, the extent of plastics pollution is reduced.'\n"
     "• 2015 baseline: 2.7 g/day/capita\n"
     "• 2024 actual: Not available (last monitoring: 2015)\n"
     "• 2028 milestone: 0.9 g/day/capita\n"
     "• 2030 target: 0 g/day/capita\n"
     "• Lead agency: DoE/MoFE\n\n"
     "Current progress assessment: 'Progress made but at an insufficient rate.'\n"
     "Evidence: 'Nepal enacted new legislation on...plastics' but 'plastic regulation "
     "enforcement is weak.' Microplastics detected in Nepal's rivers and human blood."),
  S=0,
  T=0.50,
  U=("Responsible authority (+0.25): DoE/MoFE explicitly designated as lead agency for "
     "Indicator 9.5 (p.76): 'DoE/MoFE' listed in monitoring framework table. The NBSAP "
     "monitoring framework designates responsible agencies for each target indicator.\n\n"
     "Enforcement (−0): National reports to the CBD are strategic/planning documents — "
     "no binding enforcement provisions within the 7NR itself. Progress is monitored "
     "and reported but there are no legal penalties for missing targets set in this document.\n\n"
     "Monitoring (+0.25): Indicator 9.5 provides an explicit monitoring framework with "
     "a defined baseline (2.7 g/day/capita), milestone (0.9 in 2028), and target (0 in 2030). "
     "The NBSAP Monitoring Framework document ('Computation of Indicators for National "
     "Reporting on NBSAP 2025-2030') provides methodology. Monitoring uses SDG framework "
     "data. The indicator is tracked annually by DoE/MoFE through the SDG Status and "
     "Roadmap document.\n\n"
     "Unconditional (−0): Progress rated 'at an insufficient rate'; enforcement acknowledged "
     "as weak; data for 2024 is 'NA' (not monitored since 2015). These conditions suggest "
     "implementation is conditional on enforcement improvements. The target itself has no "
     "stated exemptions, but the de facto conditional nature is noted."),
  W=("P = 0.20 (Governance & coordination — planning requirement): NBT 9 constitutes a "
     "national planning commitment to reduce plastic pollution, integrated into the CBD/KM-GBF "
     "framework. It is not a regulatory ban (P=1.0) or economic instrument (P=0.60) — it is "
     "a planning target that assigns responsibilities (DoE/MoFE) and establishes monitoring.\n\n"
     "S = 0: The national biodiversity targets are endorsed through ministerial decision and "
     "reported to an international body, but do not carry direct legal force as binding "
     "regulations. The 7NR is a reporting/strategy document, not legislation or a binding "
     "executive decree. Individual policy instruments cited in the 7NR (e.g., plastic bag ban) "
     "may be in force, but the 7NR itself is not an operative instrument.\n\n"
     "KEY CONTEXT: This instrument is important because it represents Nepal's first "
     "internationally-reported, quantified plastic pollution reduction target (0 g/day/capita "
     "by 2030). Reaching this target would require reducing plastic use from 2.7 to 0 g/per "
     "person per day — a very ambitious goal. The 2015 baseline has not been updated since.\n\n"
     "Cross-reference: The 7NR explicitly references the following existing plastic instruments: "
     "Action Plan for Ban on Plastic Bags 2022; plastic bag ban since August 2021 (<40 microns). "
     "NBT 9 also directly supports Basel, Rotterdam and Stockholm Conventions and Nepal's NDC 3.0."),
 ),
 dict(
  P=0.20, Q="Consumption",
  R=("Annex 3.15 (pp. 99–103): National Biodiversity Target 15 (NBT 15), aligned with "
     "KM-GBF Target 16: 'By 2028, develop a supportive, legal or regulatory framework to "
     "encourage people towards sustainable consumption, including sensitization and education.'\n\n"
     "Actions under NBT 15 include: (a) integrating circular economy approaches for waste "
     "reduction and resource efficiency, and (b) promoting sustainable lifestyles.\n\n"
     "This target directly addresses the consumption of plastic and other materials through "
     "a circular economy approach. It commits Nepal to developing a legal/regulatory framework "
     "for sustainable consumption by 2028, which encompasses plastic consumption. Current "
     "assessment: 'There are no specific policies or strategies targeting sustainable "
     "consumption' (7NR, p.106), meaning the framework is yet to be developed.\n\n"
     "Progress: 'Progress made but at an insufficient rate.' Nepal's domestic material "
     "consumption per capita increased from 6.563 in 2020 to 6.885 in 2024, indicating "
     "rising consumption rather than a decline."),
  S=0,
  T=0.25,
  U=("Responsible authority (+0.25): National Planning Commission / MoFE and relevant "
     "ministries would lead the development of a sustainable consumption framework. The NBSAP "
     "Monitoring Framework designates lead agencies for each national target.\n\n"
     "Enforcement (−0): No enforcement provisions — this is a commitment to develop a future "
     "framework, not an operative instrument.\n\n"
     "Monitoring (−0): No specific monitoring indicator for plastic consumption under NBT 15 "
     "identified in the available text. NBT 15 monitoring relies on SDG-linked sustainability "
     "indicators (ecological footprint, domestic material consumption), which are available but "
     "not plastic-specific.\n\n"
     "Unconditional (−0): The 7NR itself acknowledges 'There are no specific policies or "
     "strategies targeting sustainable consumption' — confirming that NBT 15 is an aspirational "
     "commitment, not an operative standard. The target is to develop a framework by 2028 — "
     "it is conditional on that future development."),
  W=("P = 0.20 (Governance & coordination — planning requirement): NBT 15 commits to "
     "developing a legal/regulatory framework for sustainable consumption by 2028, which "
     "is a governance/coordination instrument (it creates future regulatory requirements "
     "rather than being a regulation itself).\n\n"
     "S = 0: No operative legal framework for sustainable consumption yet exists. The target "
     "is to CREATE one by 2028 — it is a future planning commitment.\n\n"
     "T = 0.25: Only responsible authority is credited (+0.25). Enforcement, monitoring, "
     "and unconditional sub-scores not credited because the instrument is a future commitment "
     "without specific monitoring mechanism or enforcement provision for plastic consumption "
     "in particular.\n\n"
     "SIGNIFICANCE: NBT 15 is broader than plastic consumption and addresses all unsustainable "
     "consumption patterns. It passes the plastics relevance filter under criterion (c) — "
     "'waste hierarchy obligations, municipal waste management planning' — because circular "
     "economy approaches for waste reduction directly apply to plastic waste streams.\n\n"
     "BORDERLINE: A case could be made for P = 0.40 (Information & voluntary) if the "
     "'sensitization and education' element is seen as the primary instrument rather than "
     "the regulatory framework development. Coded as P = 0.20 (governance) because the "
     "stated target is to develop a legal/regulatory framework."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — NBT 9: Reduce plastic pollution to 0 g/day/capita by 2030 (Indicator 9.5); actions: reduce, reuse, recycle plastics",
 "Instrument 2 — NBT 15: Develop legal/regulatory framework for sustainable consumption (incl. circular economy/plastic waste reduction) by 2028",
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

for c,w in {1:32,2:38,3:8,4:35,5:8,6:42,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:35,15:9,
            16:10,17:18,18:48,19:10,20:10,21:48,22:9,23:48,24:30}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Provisions + Monitoring Framework
ws2=wb.create_sheet("Key Provisions & Plastic Targets")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Nepal's Seventh National Report (7NR) to the Convention on Biological Diversity (CBD)"),
 ("Date","February 2026 (Final Draft, Modified 16 February 2026)"),
 ("Prepared by","Environment and Biodiversity Division, Ministry of Forests and Environment (MoFE)"),
 ("Support","GEF Early Action Support (EAS) Project, GEF Umbrella Project, BIOFIN/UNDP"),
 ("Based on","Nepal NBSAP 2025-2030; NBSAP Vision Document 2050; National Biodiversity Targets 2030"),
 ("Status","Final Draft — For Discussion, not for circulation or reference"),
 ("CBD alignment","Kunming-Montreal Global Biodiversity Framework (KM-GBF)"),
 ("",""),
 ("NATIONAL BIODIVERSITY TARGET 9 — POLLUTION CONTROL (PLASTICS)",None),
 ("Field","Content"),
 ("Full text",
  "'By 2030, reduce impacts of pollution from all sources, especially from plastics, "
  "pesticides, wastewater, and nutrients, to levels that are not harmful to biodiversity, "
  "especially in areas of high importance for biodiversity.'"),
 ("KM-GBF alignment","Target 7: Reduce Pollution to Levels That Are Not Harmful to Biodiversity"),
 ("Actions for plastics",
  "'(c) reducing, reusing, and recycling plastics' [Annex 3.9, p.73]"),
 ("Progress status","☒ Progress made but at an insufficient rate"),
 ("Challenges",
  "'plastic regulation enforcement is weak'; plastic use in 2024 not monitored; "
  "microplastics detected in Nepal's rivers and human blood"),
 ("SDG linkage",
  "SDG 11 (Sustainable Cities), SDG 12 (Responsible Consumption), SDG 14 (Life Below Water), "
  "SDG 15 (Life on Land)"),
 ("Convention linkage","Basel, Rotterdam, Stockholm Conventions; Nepal's NDC 3.0"),
 ("",""),
 ("INDICATOR 9.5 — PLASTIC POLLUTION (from Monitoring Framework, p.76)",None),
 ("Field","Value"),
 ("Full indicator text","By 2030, the extent of plastics pollution is reduced"),
 ("Lead agency","DoE/MoFE (Department of Environment / Ministry of Forests and Environment)"),
 ("Data source","Secondary sources (SDG framework)"),
 ("2015 BASELINE","2.7 g/day/capita (last available data point — NOT updated since 2015)"),
 ("2020 actual","NA (not monitored)"),
 ("2024 actual","NA (not monitored)"),
 ("2028 MILESTONE","0.9 g/day/capita"),
 ("2030 TARGET","0 g/day/capita"),
 ("",""),
 ("NATIONAL BIODIVERSITY TARGET 15 — SUSTAINABLE CONSUMPTION",None),
 ("Field","Content"),
 ("Full text",
  "'By 2028, develop a supportive, legal or regulatory framework to encourage people towards "
  "sustainable consumption, including sensitization and education.'"),
 ("Actions",
  "(a) Integrating circular economy approaches for waste reduction and resource efficiency; "
  "(b) promoting sustainable lifestyles"),
 ("Progress status","☒ Progress made but at an insufficient rate"),
 ("Current gap",
  "'There are no specific policies or strategies targeting sustainable consumption' (7NR, p.106). "
  "Nepal's DMC per capita increased from 6.563 (2020) to 6.885 (2024) — rising consumption."),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.50  (strategy/plan with quantifiable targets — Indicator 9.5 has specific milestones)"),
 ("policy_target (E)","1  — Indicator 9.5: plastic use → 0 g/day/capita by 2030"),
 ("policy_integration (I)","1.00  (10+ sectors: agriculture, waste management, conservation, water, forestry, fisheries, municipalities, tourism, industry, chemicals)"),
 ("policy_circularity (K)","0.75  (consumption, recycling, disposal, environmental leakage — 4 lifecycle phases)"),
 ("policy_budget (M)","0.50  (GEF + BIOFIN/UNDP funding mentioned; not plastic-specific ring-fenced)"),
 ("",""),
 ("Instr. 1: NBT 9 — plastic pollution reduction (Indicator 9.5)","P=0.20 | S=0 | T=0.50 | V=[auto 0.350]"),
 ("Instr. 2: NBT 15 — sustainable consumption framework","P=0.20 | S=0 | T=0.25 | V=[auto 0.225]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Content","Value") and row[0]=="Field":
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
 ws2.row_dimensions[ri].height=45
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational, no concrete targets"),
     (0.50,"Strategy/plan — incorporates quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination (planning requirement, target-setting)"),
     (0.40,"Information & voluntary"),(0.60,"Economic"),
     (0.80,"Infrastructure"),(1.00,"Regulatory"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — plan/aspiration; no operative legal provision"),
     (1,"In force — mandatory/prohibitory language or operationalised"),
     ("",""),("INSTRUMENT IMPLEMENTATION (Col T)",None),
     ("Sub-score","Criterion (must be explicitly evidenced)"),
     ("+0.25","Responsible authority designated"),("+0.25","Enforcement"),
     ("+0.25","Monitoring mechanism"),("+0.25","Unconditional")]
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

path="/workspace/4p_index_nepal_7nr_cbd_2026.xlsx"
wb.save(path)
print(f"Saved: {path}")
