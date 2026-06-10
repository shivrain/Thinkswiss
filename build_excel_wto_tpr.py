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
c.value=("Policy: WTO Trade Policy Review — Report by the Secretariat — Nepal (WT/TPR/S/478, 2025)  |  "
         "Country: Nepal  |  Year: 2025  |  "
         "Source: https://www.wto.org/english/news_e/... [URL truncated]")
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
 name=("WTO Trade Policy Review — Report by the Secretariat — Nepal\n"
       "(Document WT/TPR/S/478, 22 September 2025)\n"
       "[Third Trade Policy Review of Nepal; review period 2018–2025]"),
 url=("https://www.wto.org/english/news_e/... [URL truncated in source image — "
      "WTO document reference: WT/TPR/S/478, dated 22 September 2025]"),
 year=2025,
 obj=("WTO Secretariat analytical review of Nepal's trade policies and practices for 2018–2025 "
      "(Nepal's third Trade Policy Review). In relation to plastics, the report documents "
      "Nepal's trade measures affecting plastic products: (1) import prohibition on plastic bags "
      "and sheets thinner than 30 microns, plastic scrap, and recycled granules (Table 3.6); "
      "(2) conditional duty-free tariff exemption for PET chips imported by POY manufacturing "
      "industries (Table 3.3); and (3) high applied MFN tariffs on the plastics and rubber "
      "HS section. The document also notes Nepal's commitments to environmental goals through "
      "its NDC and participation in international climate agreements."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("IMPORTANT NOTE: This document is a WTO Secretariat review report (WT/TPR/S/478), not a "
          "national policy instrument. It is prepared by the WTO Secretariat on its own responsibility "
          "to inform the Trade Policy Review Body. It is NOT: legislation, a regulation/decree, or a "
          "strategy/plan adopted by the Government of Nepal. It DESCRIBES Nepal's existing trade "
          "policies but does not create or modify them.\n\n"
          "For 4P Index purposes, this document is coded as an INFORMATION/REVIEW document "
          "equivalent to a strategy/plan with aspirational commitments (G = 0.25) — the lowest "
          "available category, which best approximates the nature of a review/reporting document. "
          "The policy 'instruments' coded (Cols P–W) are the plastic-relevant trade measures "
          "DOCUMENTED in the report as being in force in Nepal, not instruments OF the report itself.\n\n"
          "This TPR documents Nepal's existing plastic-related trade measures as part of its "
          "systematic review of all trade policies; it is valuable for the 4P Index as it provides "
          "independent verification and a current snapshot of Nepal's plastic trade measures from "
          "an international perspective."),
 intg=1,
 sects="imports/trade, industry, waste management, environment, agriculture, energy, tourism",
 circ=0.75,
 lc="production, consumption, recycling, disposal",
 budg=0,
 budg_t="",
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Table 3.6 (p.53) — Products subject to import prohibitions, 2018 and 2025:\n\n"
     "Row 4 [existing, updated]: 'Plastics scrap and bags and sheets of plastics below "
     "30-micron thickness' — prohibited in BOTH 2018 AND 2025.\n"
     "Footnote a: 'In 2018, the thickness threshold was below 20 microns.' → threshold "
     "increased from 20 to 30 microns during the review period.\n\n"
     "Row 7 [NEW in 2025 only]: 'Scrap plastic and recycled granules' — newly added to "
     "import prohibition list by 2025.\n\n"
     "Row 8 [NEW]: 'Used second-hand goods, except for those permitted by law (However, "
     "used goods can be imported on recommendation of Ministry of Forest and Environment "
     "during manmade or natural disasters)' — relevant for plastic-containing used goods.\n\n"
     "Legal basis: Export and Import (Control) Act, 1957 and its Regulation, 1978; new "
     "Import (Control) Bill, 2025 submitted to the House of Representatives to replace "
     "the 1957 Act. Enforcement: Department of Commerce, Supplies and Consumer Protection "
     "(DOCSCP) under the Ministry of Industry, Commerce and Supplies."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): DOCSCP under Ministry of Industry, Commerce and Supplies "
     "designated as the competent authority administering Nepal's import licensing and prohibition "
     "regime (WT/TPR/S/478, p.53). Customs enforces at the border.\n\n"
     "Enforcement (+0.25): Import prohibitions are legally binding — any import of prohibited "
     "items violates Nepal's trade law (Export and Import Control Act, 1957). The TPR confirms "
     "enforcement through Nepal's ASYCUDA World customs clearance system (99.8% of Nepal's "
     "cross-border trade now operates under ASYCUDA World, p.45/para.15). Import prohibitions "
     "are policed at all major customs offices.\n\n"
     "Monitoring (+0.25): Nepal's ASYCUDA World system provides tracking of all imports; the "
     "Nepal National Single Window for customs clearance provides transparency; the DOCSCP "
     "administers the licensing regime in coordination with the Department of Customs (DOC).\n\n"
     "Unconditional (−0): The prohibition on bags <30 microns and plastic scrap has no "
     "stated exemptions within the prohibition list itself. HOWEVER, the threshold changed "
     "from 20 to 30 microns between 2018 and 2025, indicating the standard is subject to "
     "policy modification. Additionally, 'used goods can be imported on recommendation of "
     "Ministry of Forest and Environment during manmade or natural disasters' (Row 8), "
     "suggesting executive discretion in import control. These conditional elements prevent "
     "the 'unconditional' sub-score from being awarded."),
  W=("This instrument directly passes the plastics relevance filter under criterion (a): "
     "both Row 4 ('Plastics scrap and bags and sheets of plastics below 30-micron thickness') "
     "and Row 7 ('Scrap plastic and recycled granules') explicitly mention plastics.\n\n"
     "S = 1: Import prohibitions are in force as confirmed by the WTO TPR (both 2018 and "
     "2025 columns marked X for Row 4; 2025 column marked X for Row 7).\n\n"
     "T = 0.75: Authority, enforcement, and monitoring sub-scores credited. Unconditional "
     "sub-score not credited because: (1) threshold changed from 20 to 30 microns (not "
     "truly unconditional — threshold is modifiable); (2) emergency import exceptions exist "
     "for MoFE-recommended used goods.\n\n"
     "THRESHOLD CHANGE SIGNIFICANCE: The increase in prohibition threshold from 20 to "
     "30 microns is an EXPANSION of the ban (more bags are now prohibited). This is a "
     "strengthening of Nepal's plastic import controls.\n\n"
     "NEW 2025 MEASURE: 'Scrap plastic and recycled granules' import prohibition "
     "(Row 7) is new since the 2018 review. This closes a potential gap where "
     "plastic waste could be imported as recycled feedstock — a common challenge in "
     "global plastic waste trade under the Basel Convention.\n\n"
     "Cross-reference: The import prohibition on thin plastic bags aligns with Nepal's "
     "domestic plastic bag ban (gazette notice August 2021; Action Plan 2022; "
     "Directive 2082/2026). The WTO TPR provides independent confirmation that these "
     "import controls are in force at the border."),
 ),
 dict(
  P=0.60, Q="Production",
  R=("Table 3.3 (p.51) — Customs duty exemptions, 2025, Item 1:\n"
     "'PET chips under Subheadings 3907.61.00 and 3907.69.00 may be imported by the "
     "Partially Oriented Yarn (POY) manufacturing industry, based on limits set by the "
     "Department of Industry (DOI) based on the raw material utilization ratio and the "
     "ratio of domestic consumption to the export of finished goods.'\n\n"
     "This is a conditional duty-free tariff exemption for plastic raw material imports "
     "(polyethylene terephthalate chips used in textile manufacturing). It provides "
     "economic incentive to domestic plastics-consuming industries by reducing their input "
     "costs for PET-based plastic materials.\n\n"
     "Context: Para. 3.27 (p.48) notes that 'Plastics and rubber' is among the HS "
     "sections with the highest average applied MFN tariffs in Nepal, confirming that the "
     "PET chip exemption provides significant cost advantage to eligible industries."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): Department of Industry (DOI) explicitly designated as "
     "the authority that 'sets limits' for the exemption 'based on the raw material "
     "utilization ratio and the ratio of domestic consumption to the export of finished "
     "goods.' (Table 3.3, Item 1)\n\n"
     "Enforcement (+0.25): The exemption is conditional on DOI approval and compliance "
     "with utilization ratios. Non-compliance would result in loss of exemption and "
     "standard duty liability. Nepal's customs ASYCUDA World system tracks import flows "
     "against DOI-approved limits.\n\n"
     "Monitoring (+0.25): DOI tracks utilization of the exemption against production "
     "ratios; customs monitors import quantities. The conditional nature ('based on limits "
     "set by DOI') implies a monitoring and approval system.\n\n"
     "Unconditional (−0): Explicitly conditional — 'based on limits set by the DOI based "
     "on the raw material utilization ratio and the ratio of domestic consumption to "
     "the export of finished goods.' The exemption only applies if: (i) imported by "
     "POY industry specifically; (ii) within limits set by DOI; (iii) based on documented "
     "production ratios."),
  W=("P = 0.60 (Economic — tariff exemption/subsidy): The duty-free import of PET chips "
     "reduces production costs for plastic-using industries, constituting an economic "
     "incentive. This passes the plastics relevance filter under criterion (a) — "
     "explicitly mentions 'PET chips' (HS 3907.61.00, 3907.69.00), a plastic raw material.\n\n"
     "NOTE ON POLICY AMBIVALENCE: This instrument creates a POSITIVE economic incentive "
     "for plastic raw material imports — it reduces the cost of importing PET chips, "
     "potentially increasing plastic production in Nepal's manufacturing sector. This is "
     "the OPPOSITE of a plastic pollution control measure. It is coded here as an "
     "economic instrument because it directly enables plastic-related production decisions.\n\n"
     "S = 1: The exemption is currently listed in Nepal's 2025 customs duty exemption "
     "table (Table 3.3), confirming it is in operative force.\n\n"
     "T = 0.75: Authority (DOI), enforcement (customs + DOI), and monitoring credited. "
     "Unconditional not credited — explicitly conditional on DOI limits and utilization ratios.\n\n"
     "Cross-reference: The plastics and rubber HS section (Section 07 in Nepal's tariff "
     "schedule) has one of the highest average MFN applied rates (Chart 3.3, p.48), "
     "meaning the PET chip exemption provides substantial tariff relief relative to "
     "the general tariff on plastics."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Import prohibition: plastic bags/sheets <30 microns + plastic scrap + recycled granules (Table 3.6)",
 "Instrument 2 — Conditional duty-free exemption for PET chips imported by POY/textile manufacturing (Table 3.3)",
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

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:32,11:8,12:28,13:8,14:25,15:9,
            16:10,17:18,18:50,19:10,20:10,21:50,22:9,23:50,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Plastic Measures")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Document title","WTO Trade Policy Review — Report by the Secretariat — Nepal"),
 ("WTO document number","WT/TPR/S/478"),
 ("Date","22 September 2025"),
 ("Prepared by","WTO Secretariat (on its own responsibility)"),
 ("Review body","Trade Policy Review Body"),
 ("Review period","2018–2025 (Nepal's third Trade Policy Review)"),
 ("Companion document","WT/TPR/G/478 (policy statement submitted by Nepal)"),
 ("Status","RESTRICTED — subject to press embargo until end of first session"),
 ("CRITICAL NOTE",
  "This is a WTO SECRETARIAT REVIEW REPORT, not a national policy instrument. It DESCRIBES "
  "Nepal's existing trade policies but does not create new policies. The 4P Index instruments "
  "coded below are Nepal's plastic-relevant trade measures DOCUMENTED in this report."),
 ("",""),
 ("TABLE 3.6 — PRODUCTS SUBJECT TO IMPORT PROHIBITIONS, 2018 AND 2025 (p.53)",None),
 ("Row","Description","2018","2025"),
 ("4",
  "Plastics scrap and bags and sheets of plastics below 30-micron thickness\n"
  "[Note: In 2018, the threshold was below 20 microns — threshold INCREASED to 30 microns]",
  "X (prohibited)","X (prohibited — threshold raised 20→30 microns)"),
 ("7","Scrap plastic and recycled granules [NEW in 2025]","—","X (NEW prohibition)"),
 ("8",
  "Used second-hand goods (with MoFE emergency exception) [NEW in 2025]","—","X"),
 ("",""),
 ("TABLE 3.3 — CUSTOMS DUTY EXEMPTIONS, 2025 (p.51)",None),
 ("Item","Description"),
 ("1",
  "PET chips (HS 3907.61.00, 3907.69.00): Duty-free imports allowed for "
  "Partially Oriented Yarn (POY) manufacturing industry. Limits set by DOI "
  "based on raw material utilization ratio and domestic consumption-to-export ratio."),
 ("",""),
 ("OTHER PLASTIC-RELEVANT CONTENT IN TPR",None),
 ("Section","Content"),
 ("Para.21 / Para.3.65 (p.9, 58)",
  "SEZ industries include 'plastics and polymers' — facilitation of plastic manufacturing in SEZs"),
 ("Para.3.27 / Chart 3.3 (p.48)",
  "'Plastics and rubber' (HS Section 07) among highest average MFN tariffs in Nepal's tariff schedule"),
 ("Para.3.34 (p.50)",
  "Some plastic product tariff lines have applied MFN rates EXCEEDING bound rates — "
  "potential WTO compliance issue"),
 ("1.3.1.6 Trade and Environment (p.20)",
  "Nepal's NDC: 'Nationally Determined Contribution Implementation Plan focuses on key sectors such "
  "as energy, agriculture, forestry, waste management, industrial processes, urban settlements, and "
  "tourism.' No direct plastic mention in trade-environment section."),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("CRITICAL NOTE on policy type","G=0.25 is the LOWEST available code. This review document is less than a strategy/plan but the scale has no lower value. The instruments described are IN FORCE as Nepal's national measures."),
 ("policy_type (G)","0.25  (review/reporting document — lowest available code; NOT a national policy instrument)"),
 ("policy_target (E)","0  (no quantifiable plastic targets set in the review document itself)"),
 ("policy_integration (I)","1.00  (7+ sectors documented: imports/trade, industry, waste management, environment, agriculture, energy, tourism)"),
 ("policy_circularity (K)","0.75  (production, consumption, recycling, disposal — 4 lifecycle phases touched on)"),
 ("policy_budget (M)","0  (no budget mentioned for plastic measures)"),
 ("",""),
 ("Instr. 1: Import prohibition (plastic bags <30μm, scrap, recycled granules)","P=1.00 | S=1 | T=0.75 | V=[auto 0.875]"),
 ("Instr. 2: PET chip duty-free exemption (economic incentive for plastic production)","P=0.60 | S=1 | T=0.75 | V=[auto 0.675]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=4)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif len(row)==4 and row[1]=="Description":
  for ci2,v in enumerate(row,1):
   if v:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
    c.alignment=Alignment(wrap_text=True); c.border=tb()
 elif len(row)==4 and row[0].isdigit():
  for ci2,v in enumerate(row,1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
 elif row[1] in("Details","Content","Value"):
  for ci2,v in enumerate(row[:2],1):
   c=ws2.cell(row=ri,column=ci2,value=v)
   c.fill=hfill("C5D9F1"); c.font=hfont(bold=True,size=9)
   c.alignment=Alignment(wrap_text=True); c.border=tb()
  ws2.merge_cells(start_row=ri,start_column=2,end_row=ri,end_column=4)
 else:
  for ci2,v in enumerate(row[:2],1):
   if v is not None:
    c=ws2.cell(row=ri,column=ci2,value=v)
    c.font=hfont(size=9); c.alignment=Alignment(wrap_text=True); c.border=tb()
  if len(row)==2: ws2.merge_cells(start_row=ri,start_column=2,end_row=ri,end_column=4)
 ws2.row_dimensions[ri].height=45
ws2.column_dimensions["A"].width=30
ws2.column_dimensions["B"].width=65
ws2.column_dimensions["C"].width=20
ws2.column_dimensions["D"].width=25

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational [used here for review/reporting document]"),
     (0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic (taxes, levies, subsidies, tariff exemptions)"),
     (0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, mandatory standard)"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — enabling power not exercised"),
     (1,"In force — mandatory/prohibitory language or operationalised"),
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

path="/workspace/4p_index_wto_tpr_nepal_2025.xlsx"
wb.save(path)
print(f"Saved: {path}")
