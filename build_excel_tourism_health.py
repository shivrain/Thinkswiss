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
c.value=("Policy: Operational Guideline with Health Protocol for Tourism Sector  |  "
         "Country: Nepal  |  Year: 2020  |  "
         "Source: https://giwmscdntwo.gov.np/media/pdf_v... [URL truncated in source image]")
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
 name="Operational Guideline with Health Protocol for Tourism Sector",
 url=("https://giwmscdntwo.gov.np/media/pdf_v... [URL truncated in source image — "
      "document hosted at giwmscdntwo.gov.np (Nepal government portal)]"),
 year=2020,
 obj=("COVID-19 pandemic reopening guideline for Nepal's tourism sector, providing health, "
      "hygiene and safety (H.H.S) protocols across hotels, restaurants, tourism offices, "
      "transportation, tours, and adventure activities. In relation to plastics: Section VII "
      "(Environment Friendly Recommendations) provides sector-specific guidance to eliminate "
      "single-use plastics (bottles, straws, plates, bags, packaging) from tourism operations; "
      "Section VI requires trekking groups to obtain a clearance letter confirming responsible "
      "disposal/return of non-biodegradable garbage from local government; and scattered 'Green "
      "Tips' throughout sections encourage burn-paper-not-plastic, carry-out-all-plastic, and "
      "Leave No Trace principles for outdoor activities."),
 tgt=0,
 tgt_t="",
 gtype=0.25,
 gtype_j=("Operational guideline issued by the Government of Nepal (hosted on government portal "
          "giwmscdntwo.gov.np) as a COVID-19 pandemic reopening framework for the tourism sector. "
          "Published approximately mid-2020 in response to COVID-19 lockdown (Nepal lockdown "
          "began 24 March 2020). Section VII is explicitly titled 'Environmental Friendly "
          "RECOMMENDATIONS' — aspirational/voluntary throughout. Contains one mandatory provision "
          "(clearance letter for non-biodegradable garbage in trekking, Section VI). Not enacted "
          "by Parliament; not a regulation or executive decree. Overall character is a guidance/"
          "standards document. Scored 0.25 (strategy/plan with aspirational commitments) — the "
          "environmental sections are advisory 'green tips' and 'recommendations', not binding "
          "regulations on the tourism industry."),
 intg=1,
 sects="tourism, hotels/accommodation, restaurants/food & beverage, transport, waste management, environment, retail/packaging, adventure/trekking",
 circ=0.75,
 lc="consumption, disposal, recycling, environmental leakage",
 budg=0,
 budg_t="",
)

FILLS=["FFF2CC","E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=1.0, Q="Waste management",
  R=("Section VI: Protocol for Tours and Adventure Activities (p.21) — Trekking/Hiking/"
     "Mountaineering:\n"
     "'A clearance letter confirming responsible disposal/return of non-biodegradable garbage "
     "has to be obtained from the local government body or authorized entity in the trekking "
     "area.'\n\n"
     "This administrative requirement obligates trekking agencies to obtain formal government "
     "verification that all non-biodegradable waste (including plastic waste) generated during "
     "a trek has been properly disposed of or removed from the trekking area. The clearance "
     "letter must be obtained from the local government or an authorized entity in each "
     "trekking region."),
  S=1,
  T=0.50,
  U=("Section VI, p.21 (verbatim): 'A clearance letter confirming responsible disposal/return "
     "of non-biodegradable garbage has to be obtained from the local government body or "
     "authorized entity in the trekking area.'\n\n"
     "Responsible authority (+0.25): 'Local government body or authorized entity in the "
     "trekking area' is explicitly designated as the authority issuing the clearance letter. "
     "This creates a clear institutional accountability mechanism.\n\n"
     "Enforcement (−0): No explicit fines or penalties stated in the guideline for failure "
     "to obtain the clearance letter.\n\n"
     "Monitoring (−0): The clearance letter itself IS a monitoring mechanism (it verifies "
     "waste was properly managed), but no explicit ongoing monitoring or audit system is "
     "described within the guideline.\n\n"
     "Unconditional (+0.25): The requirement applies to all trekking/hiking/mountaineering "
     "groups with no stated exemptions — 'has to be obtained' uses mandatory language."),
  W=("S = 1 (IN FORCE): Mandatory language 'has to be obtained' — language test confirms "
     "operative status. This is the only mandatory plastic-relevant provision in the "
     "guideline; all other plastic-related measures are voluntary/advisory.\n\n"
     "T = 0.50: Responsible authority (+0.25: local government/authorized entity) and "
     "unconditional (+0.25: no exemptions) sub-scores credited. Enforcement not credited "
     "(no penalties stated). Monitoring not credited (clearance letter is itself a one-time "
     "verification, not a monitoring system with regular inspections/audits).\n\n"
     "P = 1.0 (Regulatory — mandatory administrative requirement): While this guideline "
     "is overall advisory, this specific provision uses mandatory language ('has to be "
     "obtained') and constitutes a regulatory standard for trekking operations — agencies "
     "must demonstrate compliance with non-biodegradable waste management to local "
     "government authorities. A case could be made for P = 0.20 (governance) since it "
     "creates an administrative requirement, but P = 1.0 is applied as the clearance letter "
     "requirement functions as a mandatory performance standard for waste management.\n\n"
     "Cross-reference: This instrument is consistent with existing Nepal conservation "
     "requirements in national parks (DNPWC garbage deposit system) and the Sagarmatha "
     "region's waste management requirements."),
 ),
 dict(
  P=0.40, Q="Consumption",
  R=("Section VII: Environment Friendly Recommendations (pp.27–32) — Hotels, Homestays, "
     "Restaurants:\n\n"
     "HOTELS/HOMESTAYS:\n"
     "1. 'Replace plastic water bottles with glass bottles.' (p.30)\n"
     "2. 'avoid using plastic wrapped soaps and mini plastic bottled shampoos.' (p.30)\n"
     "3. 'Reusable bag for laundry... No more plastic bags.' (p.30)\n"
     "4. 'snacks wrapped in plastic can be switched out with fresh fruits or bakery treats.' (p.30)\n"
     "RECOMMENDED GREEN TIPS for Smaller Hotels/Homestays: 'Do not use plastic mineral "
     "water bottles; Do not use plastic Straw; Do not use disposable plastic plates and "
     "other cutleries.' (p.31)\n\n"
     "RESTAURANTS/FOOD & BEVERAGE:\n"
     "5. 'NO straws. Stop purchasing single-use straws, especially plastic ones.' (p.31)\n"
     "6. 'No single-serve, individually wrapped products at breakfast.' (p.31)\n"
     "7. 'Avoid using plastic plates, cups and cutlery.' (p.31)\n"
     "8. 'Provide Eco-friendly take-away. Use only paper cups for beverages and "
     "biodegradable food containers.' (p.31)\n\n"
     "TOUR/TRAVEL OFFICES:\n"
     "9. 'Do not use plastic mineral water bottles. Use reusable water bottles and mugs.' (p.32)\n\n"
     "TOURIST VEHICLES:\n"
     "10. 'Replace plastic bags with paper bags.' (p.32-33)\n"
     "11. 'Use paper sickness bags.' (p.33)\n\n"
     "These recommendations target single-use plastic substitution at the consumption stage "
     "across all tourism sector operators."),
  S=0,
  T=0.25,
  U=("Language assessment: All provisions in Section VII use advisory language — 'Replace', "
     "'Avoid', 'Recommended Green Tips', 'Do not use', 'Stop purchasing', 'Use only', "
     "'Provide'. Section VII is explicitly titled 'Environmental Friendly RECOMMENDATIONS.'\n\n"
     "Responsible authority (+0.25): The guideline is issued by the Government of Nepal and "
     "addresses specific categories of operators (hotels, restaurants, offices, vehicles). "
     "While the guideline does not designate a specific enforcement authority, the Department "
     "of Tourism (DOT) is referenced elsewhere in the document as the oversight body for "
     "tourism sector compliance.\n\n"
     "Enforcement (−0): No fines or penalties for non-compliance with environmental "
     "recommendations.\n\n"
     "Monitoring (−0): No monitoring or audit mechanism for environmental green tips.\n\n"
     "Unconditional (−0): All advisory recommendations — operators may choose to implement "
     "or not implement them at their discretion."),
  W=("S = 0: All Section VII environmental recommendations use advisory/voluntary language. "
     "'Recommended Green Tips' explicitly signal voluntary nature. The document makes clear "
     "throughout that environmental sustainability measures are recommendations, not mandates.\n\n"
     "T = 0.25: Only responsible authority sub-score credited (DOT implied throughout document; "
     "Section VI specifies group health records 'should be submitted to DOT'). No enforcement, "
     "monitoring, or unconditional sub-scores met for advisory recommendations.\n\n"
     "P = 0.40 (Information & voluntary): These are awareness/guidance recommendations for "
     "plastic substitution across the tourism sector. They provide information on best "
     "practices but lack enforcement or economic incentives. The guideline functions as an "
     "awareness-raising and capacity-building tool for the tourism industry's transition "
     "away from single-use plastics.\n\n"
     "RICHNESS OF PLASTIC CONTENT: Despite being advisory, this section contains arguably "
     "the most comprehensive and sector-specific plastic substitution recommendations of "
     "any Nepal tourism policy document coded in this dataset — covering hotels, restaurants, "
     "offices, and vehicles with specific product-by-product replacement guidance.\n\n"
     "Instruments 2 and 3 could be combined as one 'information & voluntary' instrument "
     "covering all plastic reduction recommendations. They are kept separate because they "
     "address distinctly different lifecycle stages (consumption vs. environmental leakage) "
     "and target different activities (indoor operations vs. outdoor adventure tourism)."),
 ),
 dict(
  P=0.40, Q="Environmental leakage",
  R=("Section VI (p.27-28) and Section VII (pp.33-34) — Leave No Trace Principles for "
     "Adventure Activities (Trekking, Hiking, Mountaineering, Rafting/Kayaking):\n\n"
     "SECTION VI (Adventure Activities, p.27): 'burn paper but not plastic, leave the "
     "campsite the way you found it.'\n\n"
     "SECTION VI (p.28, Recommended Green Tip): 'Carry out all cans, bottles and plastic. "
     "Separate trash in the dining and kitchen tents, burn dry paper trash but not plastics, "
     "bury organic waste.'\n\n"
     "'Decant as much as possible from glass and plastic into reusable plastic containers.'\n\n"
     "SECTION VII (p.33-34, Trekking, Leave No Trace):\n"
     "1. 'Carry out bring back all cans, bottles and plastic.'\n"
     "2. 'Separate trash... burn dry paper trash but not plastics.'\n"
     "3. 'Encourage trekkers NOT to bring plastic gifts for children which end up broken "
     "on the trail or in the village.'\n"
     "4. 'Camping groups register at the checkposts with a list of all cans and bottles "
     "brought into the park and have to bring the same number back out when they leave.'"),
  S=0,
  T=0.25,
  U=("Language assessment: Advisory throughout — 'encourage', 'ensure', 'make sure', "
     "green tips format. These are recommendations/best practices for guiding companies "
     "and trekkers, not enforceable regulations.\n\n"
     "Responsible authority (+0.25): Section VI references 'national park or trekking "
     "region' checkposts as the verification point for carry-in/carry-out compliance. "
     "This designates an implicit responsible authority structure (park/checkpost officials).\n\n"
     "Enforcement (−0): Advisory/voluntary — no penalties for not following Leave No "
     "Trace plastic principles.\n\n"
     "Monitoring (−0): 'Camping groups register at checkposts with a list of all cans "
     "and bottles brought in' mentions a monitoring-like system (checkpost registration), "
     "but this is presented as a recommendation ('Work with each national park... to have "
     "a system like Upper Mustang'), not as a currently operating mandatory mechanism.\n\n"
     "Unconditional (−0): All recommendations — tour companies and trekkers choose "
     "whether to follow them."),
  W=("S = 0: Leave No Trace principles are clearly framed as recommendations and aspirational "
     "goals for responsible tourism, not as binding legal obligations.\n\n"
     "T = 0.25: Responsible authority sub-score credited (checkpost/park authority referenced "
     "as a monitoring-like mechanism). Monitoring not credited as the checkpost registration "
     "system is presented as an aspirational model ('like Upper Mustang'), not as a currently "
     "operative mandatory system.\n\n"
     "P = 0.40 (Information & voluntary): The Leave No Trace principles provide environmental "
     "awareness and guidance targeting plastic waste that would otherwise enter Nepal's mountain "
     "and river environments (environmental leakage prevention).\n\n"
     "NOTE: The checkpost carry-in/carry-out system mentioned (register cans and bottles brought "
     "into the park, bring same number out) is actually a monitoring mechanism that exists in "
     "Upper Mustang and similar regions. If this system has been formally operationalised for "
     "all Nepal trekking areas, the monitoring sub-score (+0.25) could apply and S might be 1. "
     "However, the guideline presents it as a recommended model, not as a currently universal "
     "system. Coded as advisory (S=0) with a note about the borderline."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Mandatory clearance letter for non-biodegradable garbage disposal/return (trekking)",
 "Instrument 2 — Information/voluntary: Replace/avoid single-use plastics in accommodation, restaurants, offices, vehicles (Section VII green tips)",
 "Instrument 3 — Information/voluntary: Leave No Trace + carry-out-all-plastic principles for outdoor adventure activities (Sections VI-VII)",
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
 ws.row_dimensions[r].height=250

for c,w in {1:30,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:25,15:9,
            16:10,17:18,18:50,19:10,20:10,21:50,22:9,23:50,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Provisions
ws2=wb.create_sheet("Key Provisions & Evidence")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Operational Guideline with Health Protocol for Tourism Sector"),
 ("Year","2020 (COVID-19 pandemic reopening protocols; Nepal lockdown began 24 March 2020)"),
 ("Hosted at","https://giwmscdntwo.gov.np/media/pdf_v... [Nepal government portal]"),
 ("Covers",
  "8 sections: Re-Start Protocol; Hotels; Restaurants; Tourism Office Management; "
  "Transportation Providers; Tours and Adventure Activities; Environment Friendly "
  "Recommendations; Cleaning and Disinfecting Guidelines"),
 ("Primary purpose","COVID-19 health, hygiene and safety protocols for tourism sector reopening"),
 ("Plastic relevance","Section VII: Environment Friendly Recommendations + Green Tips throughout"),
 ("",""),
 ("INSTRUMENT 1: CLEARANCE LETTER FOR NON-BIODEGRADABLE GARBAGE (Section VI, p.21)",None),
 ("Text","'A clearance letter confirming responsible disposal/return of non-biodegradable garbage has to be obtained from the local government body or authorized entity in the trekking area.'"),
 ("Applies to","Trekking, Hiking, Mountaineering groups"),
 ("Authority","Local government body or authorized entity in the trekking area"),
 ("",""),
 ("INSTRUMENT 2: SINGLE-USE PLASTIC SUBSTITUTION — KEY PROVISIONS (Section VII)",None),
 ("Location","Provision"),
 ("Hotels p.30","'Replace plastic water bottles with glass bottles.'"),
 ("Hotels p.30","'avoid using plastic wrapped soaps and mini plastic bottled shampoos'"),
 ("Hotels p.30","'Reusable bag for laundry... No more plastic bags.'"),
 ("Hotels p.30","'snacks wrapped in plastic can be switched out with fresh fruits'"),
 ("Restaurants p.31","'NO straws. Stop purchasing single-use straws, especially plastic ones.'"),
 ("Restaurants p.31","'No single-serve, individually wrapped products at breakfast'"),
 ("Restaurants p.31","'Avoid using plastic plates, cups and cutlery.'"),
 ("Restaurants p.31","'Provide Eco-friendly take-away. Use only paper cups for beverages and biodegradable food containers.'"),
 ("Green Tips p.31",
  "'Do not use plastic mineral water bottles; Do not use plastic Straw; "
  "Do not use disposable plastic plates and other cutleries'"),
 ("Tour/Travel Offices p.32","'Do not use plastic mineral water bottles. Use reusable water bottles and mugs at work.'"),
 ("Vehicles p.32-33","'Replace plastic bags with paper bags.' 'Use paper sickness bags.'"),
 ("",""),
 ("INSTRUMENT 3: LEAVE NO TRACE / PLASTIC CARRY-OUT PRINCIPLES (Sections VI-VII)",None),
 ("Location","Provision"),
 ("Sec.VI p.27","'burn paper but not plastic, leave the campsite the way you found it'"),
 ("Sec.VI p.28","'Carry out all cans, bottles and plastic.'"),
 ("Sec.VI p.28","'burn dry paper trash but not plastics, bury organic waste'"),
 ("Sec.VII p.33","'Carry out bring back all cans, bottles and plastic.'"),
 ("Sec.VII p.34","'Encourage trekkers NOT to bring plastic gifts for children'"),
 ("Checkpost system","'register at the checkposts with a list of all cans and bottles brought into the park and have to bring the same number back out'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.25  (operational guideline — advisory/voluntary; Section VII explicitly titled 'RECOMMENDATIONS')"),
 ("policy_target (E)","0  (no quantifiable plastic-specific targets)"),
 ("policy_integration (I)","1.00  (8 sectors: tourism, hotels/accommodation, restaurants/F&B, transport, waste management, environment, retail/packaging, adventure/trekking)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, recycling, environmental leakage — 4 lifecycle phases)"),
 ("policy_budget (M)","0  (no budget or funding mentioned)"),
 ("",""),
 ("Instr. 1: Clearance letter (non-biodegradable waste, trekking)","P=1.00 | S=1 | T=0.50 | V=[auto 0.750]"),
 ("Instr. 2: Single-use plastic substitution (hotels/restaurants/offices/vehicles)","P=0.40 | S=0 | T=0.25 | V=[auto 0.325]"),
 ("Instr. 3: Leave No Trace + carry-out plastic (outdoor adventure)","P=0.40 | S=0 | T=0.25 | V=[auto 0.325]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Provision"):
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
 ws2.row_dimensions[ri].height=42
ws2.column_dimensions["A"].width=32
ws2.column_dimensions["B"].width=88

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — advisory language or enabling power not exercised"),
     (1,"In force — mandatory language or operationalised"),
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

path="/workspace/4p_index_tourism_health_protocol_2020.xlsx"
wb.save(path)
print(f"Saved: {path}")
