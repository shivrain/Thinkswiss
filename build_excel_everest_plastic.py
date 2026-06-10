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
c.value=("Policy: Directive on the Ban of Single-Use Plastics in the Sagarmatha (Everest) Region  |  "
         "Country: Nepal  |  Year: 2020  |  "
         "Source: https://mofe.gov.np/content/79/plastic... [URL truncated in source image]")
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
 name="Directive on the Ban of Single-Use Plastics in the Sagarmatha (Everest) Region\n[Nepali: सगरमाथा क्षेत्रमा एकल प्रयोगका प्लाष्टिकजन्य सामग्री प्रतिबन्ध सम्बन्धी निर्देशिका]",
 url=("https://mofe.gov.np/content/79/plastic... [URL truncated in source image — "
      "document hosted at Ministry of Forests and Environment (mofe.gov.np); "
      "last page references Office of the Prime Minister and Council of Ministers "
      "(opmcm.gov.np), suggesting Council of Ministers-level authority]"),
 year=2020,
 obj=("To ban single-use plastic products in the Sagarmatha (Everest) region of Nepal — "
      "including plastic bags thinner than 30 microns, all plastic beverage bottles (water, "
      "soda), straws, and single-use food/beverage packaging — in order to reduce plastic "
      "pollution from trekking and mountaineering tourism in the Khumbu/Sagarmatha National "
      "Park area. The directive also establishes multi-stakeholder coordination for enforcement "
      "(involving local government, trekking companies, airlines, and the Nepal Mountaineering "
      "Association) and promotes biodegradable/starch-based alternatives."),
 tgt=0,
 tgt_t="",
 gtype=0.75,
 gtype_j=("Directive/administrative order hosted on the Ministry of Forests and Environment "
          "(mofe.gov.np) website, effective from 1 January 2020. The document is 12 pages and "
          "contains detailed provisions, numbered sections, and a table of prohibited items and "
          "alternatives. The last page of the document references the Office of the Prime Minister "
          "and Council of Ministers (opmcm.gov.np), indicating Cabinet-level authority. This "
          "constitutes a sub-legislative executive directive — not enacted by Parliament.\n\n"
          "TRANSLATION NOTE: The original document is in Nepali. OCR quality of the uploaded PDF "
          "is severely degraded. Coding is based on: (a) partial Nepali text extraction; (b) "
          "supplementary research from news reports (BBC, The Hindu, myRepublica, UPI, Himalayan "
          "Times); and (c) the SWITCH-Asia Nepal plastic policy overview. Key terms visible in "
          "partial OCR: 'एकल प्रयोग' (single use), listed prohibition sections, table of "
          "alternatives (including 'starch based biodegradable plastic'), monitoring provisions."),
 intg=0.75,
 sects="tourism, waste management, conservation, water, retail, packaging",
 circ=0.75,
 lc="consumption, disposal, environmental leakage",
 budg=0,
 budg_t="",
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("Directive effective 1 January 2020: Bans all single-use plastic products in the "
     "Sagarmatha/Everest region (Khumbu Pasang Lhamu Rural Municipality and Sagarmatha "
     "National Park area). Specifically prohibited items include:\n"
     "• All plastic bags thinner than 30 microns\n"
     "• All plastic beverage bottles (water, soda: Coke, Fanta, Sprite, Mirinda, etc.)\n"
     "• Plastic straws\n"
     "• Single-use food and beverage packaging\n\n"
     "Items PERMITTED: beverages in metal cans; plastic items thicker than 30 microns; "
     "local households provided with 5 plastic bags for daily use.\n\n"
     "The directive includes a detailed table (visible on p.5 of document) comparing "
     "prohibited plastic items with required biodegradable/starch-based alternatives. "
     "All stores in the region prohibited from selling banned items.\n\n"
     "[TRANSLATION NOTE: The document is in Nepali with severely corrupted OCR. The above "
     "is based on: (a) partial Nepali text; (b) news reports from BBC, myRepublica, The Hindu, "
     "UPI; (c) SWITCH-Asia Nepal policy overview. The document visible on mofe.gov.np is "
     "significantly more comprehensive than the local municipality's two-line decision — "
     "12 pages including definitions, prohibited item lists, alternatives table, monitoring "
     "provisions, and implementation responsibilities.]"),
  S=1,
  T=0.50,
  U=("Responsible authority (+0.25): The directive is issued by/hosted on MoFE with "
     "Council of Ministers authority (opmcm.gov.np referenced on last page). Multiple "
     "implementation bodies named: local government (Khumbu Pasang Lhamu Rural Municipality), "
     "DNPWC/Sagarmatha National Park office, Nepal Mountaineering Association, trekking "
     "companies, airlines. Chief Administrative Officer Ganesh Ghimire quoted: 'All types "
     "of plastic bags, bottles and items not meeting the given standard will be banned here. "
     "Anyone using banned plastic items will be fined.'\n\n"
     "Enforcement (−0): At time of announcement (August 2019), 'No penalty has yet been "
     "agreed on for people violating the rule' (BBC, The Hindu). While the official stated "
     "fines would apply, no specific penalty amounts were codified in the available document "
     "text. [Cannot verify explicit penalty provisions from corrupted OCR.]\n\n"
     "Monitoring (+0.25): Sagarmatha Pollution Control Committee (SPCC) already operational "
     "for $4,000 per-expedition waste deposits; coordination with trekking companies and "
     "airlines for compliance monitoring explicitly mandated. Pages 9-10 of document appear "
     "to contain monitoring and reporting provisions (partially readable in Nepali).\n\n"
     "Unconditional (−0): Explicit exemptions stated — plastic items >30 microns permitted; "
     "beverages in metal cans permitted; local households provided 5 plastic bags (use "
     "exemption for daily activities). These constitue explicit loopholes."),
  W=("S = 1 (IN FORCE): The ban came into effect on 1 January 2020 with mandatory prohibitory "
     "language ('banned', 'will be fined'). Published as a government directive with immediate "
     "operative effect.\n\n"
     "T = 0.50: Responsible authority (+0.25) and monitoring (+0.25) credited. Enforcement "
     "not credited (+0): at time of issuance, specific penalty amounts had not been determined "
     "(BBC: 'No penalty has yet been announced'). Unconditional not credited (+0): explicit "
     "exemptions for >30 micron plastics, metal cans, and household plastic bag provision.\n\n"
     "TRANSLATION UNCERTAINTY: The full 12-page document in Nepali contains more detailed "
     "provisions than can be extracted from the corrupted OCR. The document includes a table "
     "of specific prohibited items and their biodegradable alternatives. If the document "
     "contains mandatory penalty provisions (possible under EPA 2076 or another enabling Act), "
     "T would increase to 0.75. If enforcement provisions are explicit, S = 1 is confirmed.\n\n"
     "GEOGRAPHIC SCOPE: This ban applies specifically to the Sagarmatha (Everest) region — "
     "Khumbu Pasang Lhamu Rural Municipality and Sagarmatha National Park, Solukhumbu "
     "district. It is NOT a nationwide ban (unlike the plastic bag directive or plastic "
     "flowers ban, which are nationwide).\n\n"
     "Cross-reference: Builds on the earlier 1999 Sagarmatha plastic ban (never enforced). "
     "Related to: DNPWC's $4,000 per-expedition waste deposit system requiring 8 kg of "
     "waste removal per climber (operational since 2014). The 2020 directive represents "
     "a significant escalation to single-use plastic prohibition at the consumption stage."),
 ),
 dict(
  P=0.20, Q="Waste management",
  R=("Multi-stakeholder coordination framework for enforcement and monitoring of the "
     "single-use plastic ban in the Sagarmatha region. The directive establishes "
     "coordination responsibilities across: (a) local government (Khumbu Pasang Lhamu "
     "Rural Municipality); (b) national park authority (DNPWC/Sagarmatha NP); (c) Nepal "
     "Mountaineering Association; (d) trekking companies and agencies; (e) airlines serving "
     "the Khumbu region; (f) Sagarmatha Pollution Control Committee (SPCC). Also establishes "
     "awareness obligations targeting visitors, trekkers, and local shopkeepers. Pages 9-11 "
     "of the document contain monitoring, reporting, and implementation provisions "
     "(partially readable in Nepali OCR)."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): Multiple designated implementing bodies: local "
     "government, DNPWC/park office, NMA, trekking companies, airlines, SPCC. 'We will "
     "be working with the local body, trekking companies and the Mountaineering Association "
     "of Nepal to enforce the ban' (Chief Administrative Officer Ghimire, quoted by The Hindu, "
     "BBC, myRepublica).\n\n"
     "Enforcement (−0): No specific penalties for coordination failures stated in available "
     "document text.\n\n"
     "Monitoring (+0.25): SPCC already has operational waste monitoring (expedition waste "
     "deposits); trekking company compliance monitoring; awareness campaigns targeting "
     "visitors. Pages 9-10 of document contain monitoring provisions.\n\n"
     "Unconditional (+0.25): The coordination obligation applies to all stakeholders in "
     "the Sagarmatha region without stated exemptions."),
  W=("This governance instrument is distinct from Instrument 1 (the product ban): it "
     "specifically creates the multi-stakeholder coordination framework needed to enforce "
     "the ban. Without effective coordination across airlines, trekking companies, local "
     "government, and park authority, the ban would replicate the failed 1999 attempt.\n\n"
     "P = 0.20 (Governance & coordination): The instrument assigns enforcement "
     "responsibilities across multiple levels/sectors of government and private actors. "
     "This is the defining feature of the Sagarmatha ban relative to national bans: the "
     "explicit multi-actor coordination mechanism involving tourism operators.\n\n"
     "IMPLEMENTATION REALITY: Despite the strong coordination framework, news reports "
     "(2023, Online Khabar) noted ongoing enforcement challenges and continued imports "
     "of banned items. The 2019/1999 ban comparison suggests coordination remains "
     "difficult in practice.\n\n"
     "NOTE: Pages 9-11 of the Nepali document appear to contain detailed "
     "monitoring and responsibility provisions that could further support higher T scoring "
     "if fully translated."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Comprehensive ban on single-use plastic items (bags <30μm, all plastic bottles, straws, packaging)",
 "Instrument 2 — Multi-stakeholder coordination framework (local govt + NMA + trekking companies + airlines + SPCC)",
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
            16:10,17:18,18:48,19:10,20:10,21:48,22:9,23:48,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Translation + evidence
ws2=wb.create_sheet("Translation & Context")
rows=[
 ("DOCUMENT OVERVIEW & TRANSLATION NOTES",None),
 ("Field","Details"),
 ("Official title (Nepali, approx.)",
  "सगरमाथा क्षेत्रमा एकल प्रयोगका प्लाष्टिकजन्य सामग्री प्रतिबन्ध सम्बन्धी निर्देशिका\n"
  "(Directive on the Prohibition of Single-Use Plastic Products in the Sagarmatha Region)"),
 ("Policy type","Directive / Administrative Order"),
 ("Effective date","1 January 2020 (announced August 2019)"),
 ("Geographic scope",
  "Khumbu Pasang Lhamu Rural Municipality, Solukhumbu district + Sagarmatha National Park "
  "area (the Everest region on Nepal's side of Mount Everest)"),
 ("Issuing authority",
  "Ministry of Forests and Environment (mofe.gov.np), with Council of Ministers authority "
  "(opmcm.gov.np referenced in document). Chief Administrative Officer: Ganesh Ghimire."),
 ("Document length","12 pages in Nepali; OCR quality severely degraded"),
 ("",""),
 ("PROHIBITED ITEMS (from partial OCR + news reports)",None),
 ("Category","Specific items banned"),
 ("Plastic bags","All plastic bags thinner than 30 microns"),
 ("Plastic bottles","ALL plastic beverage bottles (water, soda) — Coke, Fanta, Sprite, Mirinda prohibited"),
 ("Straws","Plastic straws"),
 ("Food packaging","Single-use plastic food and beverage packaging"),
 ("Other","All single-use plastic items ('एकल प्रयोग' = single use)"),
 ("",""),
 ("EXEMPTIONS",None),
 ("Permitted","Beverages in metal cans; plastic items >30 microns thick"),
 ("Household provision","Local households given 5 plastic bags for daily use (household exemption)"),
 ("",""),
 ("ALTERNATIVES REFERENCED (from partial OCR of p.5 table)",None),
 ("Alternative","Description"),
 ("Starch-based biodegradable","'starch based biodegradable plastic' — visible in partial English OCR"),
 ("Metal containers","Beverage cans explicitly permitted as alternative to plastic bottles"),
 ("",""),
 ("KEY QUOTED TEXT (English translations from news sources)",None),
 ("Source","Quote"),
 ("Chief Admin Officer Ghimire (myRepublica)",
  "'All types of plastic bags, bottles and items not meeting the given standard will be "
  "banned here. Anyone using banned plastic items will be fined.'"),
 ("Chief Admin Officer Ghimire (The Hindu)",
  "'Popular soft drink items like Coke, Fanta, Sprite, Mirinda and other beverages in "
  "plastic bottles will not be allowed. But beverages in metal cans will be allowed.'"),
 ("Enforcement statement (The Hindu/BBC)",
  "'We will be working with the local body, trekking companies and the Mountaineering "
  "Association of Nepal to enforce the ban.'"),
 ("Penalty status at announcement (BBC)",
  "'No penalty has yet been announced for violators.' [Note: fines pledged later]"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (executive directive — not Parliament-enacted)"),
 ("policy_target (E)","0  (ban is regulatory, not a % reduction target)"),
 ("policy_integration (I)","0.75  (6 sectors: tourism, waste management, conservation, water, retail, packaging)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, environmental leakage — 3 lifecycle phases)"),
 ("policy_budget (M)","0  (no budget stated in available document text)"),
 ("Instr.1: Ban on single-use plastics","P=1.00 | S=1 | T=0.50 | V=0.750"),
 ("Instr.2: Multi-stakeholder coordination","P=0.20 | S=1 | T=0.75 | V=0.475"),
 ("",""),
 ("TRANSLATION CAVEAT",None),
 ("Important note",
  "The Nepali source document (12 pages) has severely corrupted OCR output. Coding is based "
  "on: (1) partial Nepali text extraction; (2) English news sources (BBC, The Hindu, "
  "myRepublica, UPI, Himalayan Times — all August 2019); (3) SWITCH-Asia Nepal policy "
  "overview (2025). If a clean translation of the full 12-page document becomes available, "
  "the coding should be reviewed — particularly T scores for both instruments and any "
  "explicit penalty provisions (which would add +0.25 enforcement sub-score)."),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Specific items banned","Description","Quote","Value"):
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
ws2.column_dimensions["A"].width=38
ws2.column_dimensions["B"].width=80

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("POLICY INTEGRATION (Col I)",None),("Score","Description"),
     (0,"0 sectors"),(0.25,"1–2"),(0.50,"3–4"),(0.75,"5–6"),(1.00,"7+"),
     ("",""),("POLICY CIRCULARITY (Col K)",None),("Score","Description"),
     (0.25,"1 phase"),(0.50,"2 phases"),(0.75,"3–4 phases"),(1.00,"All 5 phases"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, standard)"),
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

path="/workspace/4p_index_everest_singleuse_plastic_ban_2020.xlsx"
wb.save(path)
print(f"Saved: {path}")
