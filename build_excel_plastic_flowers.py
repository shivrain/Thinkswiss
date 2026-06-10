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
c.value=("Policy: Notice Imposing Complete Ban on Plastic Flower Bouquets  |  Country: Nepal  |  Year: 2022  |  "
         "Source: https://mofe.gov.np/content/127/production-of-plastic-flowers---inflictinginformation/")
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

# Single row — one instrument
r = 5

ws.cell(row=r,column=24,value="Instrument 1 — Complete ban on plastic flower bouquets (production, import, sale, distribution, storage)").font=hfont(italic=True,size=8,color="808080")

pol_data = [
 # A
 ("Notice Imposing Complete Ban on Plastic Flower Bouquets\n"
  "(Nepal Gazette, Volume 72, No. 24, 2079/04/12 BS / 29 July 2022)\n"
  "[Nepali original title: सूचना — प्लाष्टिकजन्य फूलगुच्छाको पूर्णरुपमा रोक]"),
 # B
 "https://mofe.gov.np/content/127/production-of-plastic-flowers---inflictinginformation/",
 # C
 2022,
 # D
 ("To impose a complete ban on the production, import, sale, distribution, and storage of plastic "
  "flower bouquets (including plastic artificial flowers and garlands) throughout Nepal, in order "
  "to control plastic pollution and reduce environmental degradation caused by non-biodegradable "
  "plastic decorative items. The ban was also intended to stop approximately NRs 100 million in "
  "annual plastic decoration imports from China and to promote domestic floriculture."),
 # E
 0,
 # F
 "",
 # G
 0.75,
 # H
 ("Ministerial executive notice published in the Nepal Official Gazette (Volume 72, No.24, dated "
  "2079/04/12 BS / 29 July 2022) by the Ministry of Forests and Environment (MoFE). Issued under "
  "the authority granted by Sub-section (6) of Section 15 of the Environment Protection Act, 2076 "
  "(2019), which was enacted by Parliament. Signed by Dr. Pem Narayan Kandel, Secretary of the "
  "Government of Nepal. Constitutes a sub-legislative executive notice — not itself enacted by "
  "Parliament, but issued under delegated authority from an Act of Parliament (EPA 2076). "
  "Gazette publication in Nepal confers immediate binding legal force."),
 # I
 0.75,
 # J
 "industry, retail, waste management, imports/trade, packaging",
 # K
 0.75,
 # L
 "production, consumption, disposal, environmental leakage",
 # M
 0,
 # N
 "",
 # O (auto)
 None,
]

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

# Instrument data
idata = [
 # P
 1.0,
 # Q
 "Production",
 # R
 ("Nepal Gazette, Volume 72, No. 24, dated 2079/04/12 BS (29 July 2022): "
  "The Ministry of Forests and Environment, exercising the authority granted by Sub-section (6) "
  "of Section 15 of the Environment Protection Act, 2076 (2019), has imposed a COMPLETE BAN "
  "(पूर्णरुपमा रोक) on the PRODUCTION, IMPORT, SALE OR DISTRIBUTION, AND STORAGE of "
  "plastic flower bouquets throughout Nepal.\n\n"
  "[FULL TRANSLATION OF NEPALI ORIGINAL]\n"
  "Nepal Gazette, Volume 72, No.24 | Date: 2079/04/12\n"
  "Part 5 — Government of Nepal, Ministry of Forests and Environment\n"
  "NOTICE\n"
  "The Government of Nepal, Ministry of Forests and Environment, exercising the authority "
  "granted by sub-section (6) of Section 15 of the Environment Protection Act, 2076, has "
  "decided to impose a complete ban on the production, import, sale or distribution, and "
  "storage of plastic flower bouquets throughout Nepal, and this notice has been published.\n"
  "By order, Dr. Pem Narayan Kandel, Secretary of the Government of Nepal.\n"
  "Printed at the Printing Department, Singha Durbar, Kathmandu. Price: Rs. 5."),
 # S
 1,
 # T
 0.50,
 # U
 ("FULL TRANSLATED TEXT (Nepali original in document):\n"
  "'नेपाल सरकार, वन तथा वातावरण मन्त्रालयले वातावरण संरक्षण ऐन, २०७६ को दफा १५ को "
  "उपदफा (६) ले दिएको अधिकार प्रयोग गरी नेपालभर प्लाष्टिकजन्य फूलगुच्छाको उत्पादन, आयात, "
  "बिक्री वितरण वा भण्डारण गर्न पूर्णरुपमा रोक लगाउने निर्णय गरेकोले यो सूचना प्रकाशन गरिएको छ ।'\n\n"
  "TRANSLATION: 'The Government of Nepal, Ministry of Forests and Environment, exercising the "
  "authority granted by sub-section (6) of Section 15 of the Environment Protection Act, 2076, "
  "has decided to impose a complete ban on the production, import, sale or distribution, and "
  "storage of plastic flower bouquets throughout Nepal, and this notice has been published.'\n\n"
  "Responsible authority (+0.25): Ministry of Forests and Environment explicitly named as the "
  "issuing authority; Secretary of Government of Nepal (Dr. Pem Narayan Kandel) signed.\n"
  "Enforcement (−0): No explicit fines or penalties stated within this notice text. [NOTE: "
  "EPA 2076 provides fines of Rs.15,000–20,000 by environmental inspectors and up to "
  "Rs.100,000 by the Department of Environment for violations, but these are NOT stated in "
  "this specific notice — referenced in subsequent MoFE enforcement warnings.]\n"
  "Monitoring (−0): No monitoring mechanism specified in the notice text.\n"
  "Unconditional (+0.25): 'Throughout Nepal' — complete nationwide ban with no stated "
  "exemptions, phase-in period, or product-size threshold."),
 # V (auto)
 None,
 # W
 ("S = 1 (IN FORCE): The gazette notice is published in the Nepal Official Gazette with "
  "mandatory prohibitory language ('complete ban imposed', 'पूर्णरुपमा रोक'). Gazette "
  "publication in Nepal confers immediate binding legal force. This is an operative "
  "prohibition, not an enabling power. Language test: 'has decided to impose a complete ban' "
  "= mandatory/prohibitory → S = 1.\n\n"
  "T = 0.50: Responsible authority (+0.25: MoFE explicitly named) and unconditional (+0.25: "
  "nationwide, no exemptions). Enforcement: no penalties stated in the notice text — "
  "penalty provisions exist in EPA 2076 (Rs.15,000–100,000 fines per MoFE enforcement "
  "warnings) but are not quoted in this notice, so not credited per strict evidence rule. "
  "Monitoring: not stated in notice.\n\n"
  "G = 0.75: This is a ministerial executive notice under delegated authority from the EPA "
  "2076 (passed by Parliament). The EPA itself would score G = 1.0; this specific notice "
  "is G = 0.75 as a sub-legislative executive instrument.\n\n"
  "SCOPE: The ban covers 'plastic flower bouquets' (प्लाष्टिकजन्य फूलगुच्छा) — "
  "interpreted by MoFE as covering all plastic artificial flowers, bouquets, and garlands. "
  "MoFE stated the ban would stop NRs ~100 million/year in plastic decoration imports from "
  "China and reduce non-biodegradable plastic pollution.\n\n"
  "NOTE on '2021' label: The user image labels this as 2021, but the gazette date is "
  "2079/04/12 BS = 29 July 2022 Gregorian. An earlier related action may have been taken "
  "in 2021 (2078 BS), but the gazette notice in this document is dated 2022. Policy_year "
  "coded as 2022 per the document date.\n\n"
  "Cross-reference: Issued under EPA 2076 Section 15(6). Related: Nepal Plastic Bag "
  "Directive 2082 (2025) which replaced Directive 2068. This flowers ban is a standalone "
  "product-specific ban using the same EPA 2076 authority."),
]

ifl=hfill("E2EFDA")
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
ws.row_dimensions[r].height=300

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:30,11:8,12:28,13:8,14:25,15:9,
            16:10,17:18,18:55,19:10,20:10,21:55,22:9,23:55,24:30}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Translation + evidence
ws2=wb.create_sheet("Translation & Key Evidence")
rows=[
 ("FULL DOCUMENT TRANSLATION (Nepali → English)",None),
 ("Field","Content"),
 ("Document type","Nepal Official Gazette Notice (सूचना)"),
 ("Gazette reference","Volume 72, Number 24 (खण्ड ७२, संख्या २४)"),
 ("Date published","2079/04/12 BS = 29 July 2022 (Gregorian)"),
 ("Issuing authority","Government of Nepal, Ministry of Forests and Environment (वन तथा वातावरण मन्त्रालय)"),
 ("Legal basis","Environment Protection Act, 2076 (2019), Section 15, Sub-section (6)"),
 ("Signatory","Dr. Pem Narayan Kandel, Secretary of the Government of Nepal"),
 ("",""),
 ("NEPALI ORIGINAL TEXT",""),
 ("नेपाल राजपत्र",
  "नेपाल सरकार, वन तथा वातावरण मन्त्रालयले वातावरण संरक्षण ऐन, २०७६ को दफा १५ को उपदफा (६) ले "
  "दिएको अधिकार प्रयोग गरी नेपालभर प्लाष्टिकजन्य फूलगुच्छाको उत्पादन, आयात, बिक्री वितरण वा "
  "भण्डारण गर्न पूर्णरुपमा रोक लगाउने निर्णय गरेकोले यो सूचना प्रकाशन गरिएको छ ।"),
 ("",""),
 ("ENGLISH TRANSLATION",""),
 ("Nepal Gazette",
  "The Government of Nepal, Ministry of Forests and Environment, exercising the authority granted by "
  "sub-section (6) of Section 15 of the Environment Protection Act, 2076 (2019), has decided to impose "
  "a COMPLETE BAN on the production, import, sale or distribution, and storage of PLASTIC FLOWER "
  "BOUQUETS throughout Nepal, and this notice has been published."),
 ("",""),
 ("KEY TERMS TRANSLATED",""),
 ("प्लाष्टिकजन्य फूलगुच्छा","Plastic flower bouquets (plastic artificial flowers, garlands, and bouquets made of plastic)"),
 ("पूर्णरुपमा रोक","Complete/total ban (100% prohibition)"),
 ("उत्पादन","Production / Manufacturing"),
 ("आयात","Import"),
 ("बिक्री वितरण","Sale or distribution"),
 ("भण्डारण","Storage"),
 ("नेपालभर","Throughout Nepal (nationwide)"),
 ("",""),
 ("ENFORCEMENT FRAMEWORK (from EPA 2076 and subsequent MoFE notices)",None),
 ("Penalty 1","Rs. 15,000–20,000 fine imposed by Environmental Inspector (from EPA 2076 penalty framework)"),
 ("Penalty 2","Up to Rs. 100,000 fine imposed by Department of Environment (from EPA 2076 penalty framework)"),
 ("Note","These penalty amounts are NOT stated in the gazette notice itself; they derive from the EPA 2076 and were confirmed in subsequent MoFE enforcement warnings issued before festivals (Tihar/Chhath)."),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (ministerial executive notice under EPA 2076 Section 15(6))"),
 ("policy_target (E)","0  (complete ban is a regulatory instrument, not a percentage reduction target)"),
 ("policy_integration (I)","0.75  (5 sectors: industry, retail, waste management, imports/trade, packaging)"),
 ("policy_circularity (K)","0.75  (4 lifecycle phases: production, consumption, disposal, environmental leakage)"),
 ("policy_budget (M)","0  (no budget or funding source mentioned in the notice)"),
 ("instrument_type (P)","1.00  (Regulatory — comprehensive ban)"),
 ("instrument_in_force (S)","1  (mandatory prohibitory language; gazette publication = immediate binding force)"),
 ("instrument_implementation (T)","0.50  (+0.25 responsible authority; +0.25 unconditional)"),
 ("instrument_score (V)","[auto] = 0.750"),
 ("policy_score (O)","[auto] = AVERAGE(G=0.75, I=0.75, K=0.75, M=0, P=1.0) = 0.65 per row 5"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[1] in("Content","Value") and row[0] in("Field",):
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
    c.font=hfont(size=9,bold=(row[1] is None and ci2==1))
    c.alignment=Alignment(wrap_text=True); c.border=tb()
 ws2.row_dimensions[ri].height=35
ws2.column_dimensions["A"].width=40
ws2.column_dimensions["B"].width=80

# Sheet 3 — Score Reference (abbreviated)
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, standard)"),
     ("",""),("INSTRUMENT IN FORCE (Col S)",None),("Score","Description"),
     (0,"Not in force — enabling power not exercised"),
     (1,"In force — mandatory language (shall/must/prohibited) or operationalised"),
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

path="/workspace/4p_index_plastic_flowers_ban_2022.xlsx"
wb.save(path)
print(f"Saved: {path}")
