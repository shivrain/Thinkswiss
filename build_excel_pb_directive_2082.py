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
c.value=("Policy: Plastic Bag (Regulation and Control) Directive, 2082 BS (2026)  |  Country: Nepal  |  Year: 2026  |  "
         "Source: https://mofe.gov.np/content/128/thin... [URL truncated in source image]")
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
 name=("Plastic Bag (Regulation and Control) Directive, 2082 BS (2026)\n"
       "[Nepali: प्लास्टिक झोला (नियमन तथा नियन्त्रण) निर्देशिका, २०८२]"),
 url=("https://mofe.gov.np/content/128/thin... [URL truncated in source image — "
      "Ministry of Forests and Environment, Government of Nepal]"),
 year=2026,
 obj=("Comprehensive directive replacing the earlier Plastic Bag (Regulation and Control) "
      "Directive, 2068 (2011). Prohibits the production, import, storage, sale, distribution "
      "and use of plastic bags thinner than 40 microns throughout Nepal. Introduces mandatory "
      "size standards (general bags: ≥7×14 inches; garbage bags: ≥14×26 inches), mandatory "
      "colour-coding for food safety (natural/uncoloured for food/medicine; black for garbage; "
      "other colours for non-food only), mandatory labelling (producer name/address, recycling "
      "symbol, thickness, size, virgin granule indicator — permanent ink, 15–20 pt font), and "
      "mandatory material standard (High Molecular High Density Polyethylene / HMHDPE). "
      "Introduces Nepal's first statutory EPR obligation: producers and importers must collect "
      "and properly manage at least 10 percent of the plastic waste from their annual "
      "production/import volumes, with annual reporting to the Department of Environment."),
 tgt=1,
 tgt_t=("'Under the Extended Producer Responsibility (EPR) framework, producers must ensure "
        "proper disposal of at least ten percent of the plastic waste from their total production "
        "annually. Importers must similarly comply with EPR obligations for at least ten percent "
        "of imported volumes annually.' (Himalayan Times, Ratopati — reporting the Directive 2082)"),
 gtype=0.75,
 gtype_j=("Ministerial directive issued by the Ministry of Forests and Environment under the "
          "authority of Section 45 of the Environment Protection Act, 2076 (2019). Comes into "
          "immediate effect. Applies to all producers, importers, storers, sellers, and users of "
          "plastic bags throughout Nepal. Not enacted by Parliament; constitutes a sub-legislative "
          "executive directive with binding legal force backed by the EPA 2076 penalty framework. "
          "Replaces the Plastic Bag (Regulation and Control) Directive, 2068 (all actions under "
          "the previous directive remain valid under the new one).\n\n"
          "NOTE: The source document (2-page image-based PDF) could not be OCR'd. Coding is based "
          "on English news reports from The Himalayan Times and Ratopati (both June 2026) which "
          "provide detailed descriptions of the directive's provisions."),
 intg=1,
 sects="industry, imports/trade, retail, waste management, packaging, recycling, food & beverage, chemicals",
 circ=1.0,
 lc="production, consumption, recycling, disposal, environmental leakage",
 budg=0,
 budg_t="",
)

FILLS=["E2EFDA","D9F0FF"]
instruments=[
 dict(
  P=1.0, Q="Production",
  R=("Directive 2082: Comprehensive regulatory package covering the entire plastic bag supply "
     "chain with FIVE interlinked mandatory standards:\n\n"
     "(1) THICKNESS BAN: Complete prohibition on production, import, storage, sale, distribution "
     "and use of plastic bags thinner than 40 microns throughout Nepal (replaces/strengthens "
     "earlier bans of 2011 and 2015).\n\n"
     "(2) SIZE STANDARDS: Only bags of minimum 7 inches × 14 inches (general bags) or minimum "
     "14 inches × 26 inches (garbage bags) may be produced, imported, and sold.\n\n"
     "(3) COLOUR-CODING: Natural or uncoloured bags: food, water, pharmaceutical use only. "
     "Black bags: exclusively for garbage. Other colours: non-food items only.\n\n"
     "(4) MANDATORY LABELLING: Every bag must permanently display — in ink at minimum 15–20 "
     "font size — the producer's name and address, recycling symbol, thickness in microns, "
     "size, and whether made from virgin granules.\n\n"
     "(5) MATERIAL STANDARD: Bags must meet High Molecular High Density Polyethylene (HMHDPE) "
     "specifications; producers must maintain calibrated measuring equipment; production "
     "records (volumes, sizes, batch numbers) must be maintained and submitted annually to "
     "the Environment Department."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Ministry of Forests and Environment (issuing authority); "
     "Department of Environment (DoEnv) as enforcement and compliance body — 'The Environment "
     "Department is empowered to regulate compliance and deploy inspectors under the Environment "
     "Protection Act, 2076. Inspectors may submit findings in the prescribed format, and the "
     "department may take action against non-compliant individuals or entities under Section "
     "35(3) of the Act. The Ministry may also issue directions to the department for effective "
     "implementation.' (Himalayan Times)\n\n"
     "Enforcement (+0.25): EPA 2076 penalty framework applies; 'Inspectors may submit findings "
     "in the prescribed format, and the department may take action against non-compliant "
     "individuals or entities under Section 35(3) of the Act.' (Himalayan Times). MoFE stated "
     "this directive 'will now be strictly enforced' (Ratopati).\n\n"
     "Monitoring (+0.25): Producers must maintain calibrated measuring equipment and submit "
     "annual production details to the Environment Department. Inspector deployment explicitly "
     "mandated. Colour-coding and labelling standards enable point-of-sale monitoring.\n\n"
     "Unconditional (+0.25): No exemptions stated for the thickness ban, size standards, colour "
     "standards, or labelling requirements — all apply to ALL producers, importers, storers, "
     "sellers, and users of plastic bags throughout Nepal."),
  W=("S = 1 (IN FORCE): Directive comes into 'immediate effect' upon issuance (Himalayan Times: "
     "'the directive comes into immediate effect'). Uses mandatory language ('must', 'shall', "
     "'may only') throughout. Backed by EPA 2076 Section 45 authority.\n\n"
     "T = 1.0: All four sub-scores explicitly evidenced — authority (DoEnv + MoFE), enforcement "
     "(EPA Section 35(3) actions + inspector deployment), monitoring (calibrated equipment + "
     "annual production records + inspector system), unconditional (no stated exemptions for any "
     "of the five standards).\n\n"
     "FIVE COMBINED STANDARDS: Thickness ban, size standards, colour-coding, labelling, and "
     "material specifications are combined into one regulatory package. Per coding rule 10 "
     "(mixed instruments), the highest score (1.0 = Regulatory) is applied. The colour-coding "
     "and labelling elements include a food safety dimension, making this directive also relevant "
     "to the food & beverage sector.\n\n"
     "IMPROVEMENT OVER 2068 DIRECTIVE: The 2082 directive adds mandatory size standards, "
     "colour-coding system, mandatory labelling with permanent ink, and HMHDPE material "
     "specification — significantly more comprehensive than the 2068 predecessor which only "
     "banned bags <40 microns.\n\n"
     "Cross-reference: The ban on thin plastic bags was previously gazette-published (August "
     "2021) and operationalised through the Action Plan 2022. This directive supersedes all "
     "previous plastic bag standards and regulations."),
 ),
 dict(
  P=1.0, Q="Recycling",
  R=("Directive 2082, EPR provisions: Nepal's FIRST statutory Extended Producer Responsibility "
     "obligation for plastic bags. Producers and importers must:\n\n"
     "• Collect and properly manage at least 10 percent of the plastic waste generated from "
     "their total annual production volume (for producers).\n"
     "• Collect and properly manage at least 10 percent of the plastic waste from their total "
     "annual import volume (for importers).\n"
     "• Submit annual reports to the Department of Environment documenting EPR compliance.\n\n"
     "This requires producers and importers to establish collection systems, contribute to "
     "recycling infrastructure, and demonstrate end-of-life plastic management for a defined "
     "fraction of the products they place on the market each year."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Producers and importers designated as EPR-responsible "
     "parties; Department of Environment designated as the reporting/oversight body. "
     "'annual reports required by the Department of Environment' (Ratopati).\n\n"
     "Enforcement (+0.25): EPA 2076 penalty framework applies to EPR non-compliance; 'department "
     "may take action against non-compliant individuals or entities under Section 35(3)' "
     "(Himalayan Times). Annual reporting requirement creates an audit trail for enforcement.\n\n"
     "Monitoring (+0.25): 'annual reports required by the Department of Environment' (Ratopati) "
     "— explicit monitoring mechanism through mandatory annual EPR compliance reporting. "
     "Producers must 'keep records of production volumes, sizes, and batch numbers' as the "
     "basis for calculating the 10% EPR obligation.\n\n"
     "Unconditional (+0.25): EPR obligation applies to ALL producers and ALL importers without "
     "stated exemptions — 'producers must ensure proper disposal of at least ten percent... "
     "Importers must similarly comply' (Himalayan Times). No de minimis threshold or phase-in "
     "mentioned."),
  W=("S = 1 (IN FORCE): Mandatory language — 'must ensure proper disposal of at least ten "
     "percent' and 'must similarly comply.' Directive comes into immediate effect.\n\n"
     "T = 1.0: All four sub-scores met. This is Nepal's first statutory EPR for plastic bags — "
     "a significant policy leap. The 10% collection rate is a quantifiable minimum standard "
     "(hence policy_target E = 1, col F).\n\n"
     "SIGNIFICANCE: The EPR provision in Directive 2082 makes this the most comprehensive and "
     "ambitious plastic bag regulation Nepal has issued. Previous bans (2011, 2015, 2021) had "
     "no EPR component. This requires producers/importers to actively participate in waste "
     "management rather than simply complying with product standards.\n\n"
     "IMPLEMENTATION NOTE: While the 10% EPR rate is modest (compared to international "
     "standards that often require 50–75% collection), it establishes the legal precedent "
     "and infrastructure for Nepal's first plastic EPR scheme. Annual reporting to DoEnv "
     "creates a compliance audit trail that was absent from all previous Nepal plastic bans.\n\n"
     "This instrument passes the plastics relevance filter under criterion (b): 'directly "
     "enables plastic-relevant policy actions — extended producer responsibility.'"),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Comprehensive regulatory package: ban <40μm + size/colour/labelling/material standards",
 "Instrument 2 — EPR: producers and importers must collect/manage ≥10% annual plastic waste (Nepal's first statutory EPR)",
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

for c,w in {1:32,2:38,3:8,4:35,5:8,6:40,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:25,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:30}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key provisions
ws2=wb.create_sheet("Key Provisions & Evidence")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Official title (Nepali)","प्लास्टिक झोला (नियमन तथा नियन्त्रण) निर्देशिका, २०८२"),
 ("Official title (English)","Plastic Bag (Regulation and Control) Directive, 2082 BS (2026)"),
 ("Issued by","Ministry of Forests and Environment (MoFE), Government of Nepal"),
 ("Legal authority","Section 45 of the Environment Protection Act, 2076 (2019)"),
 ("Effective date","Immediate effect upon issuance (2026)"),
 ("Replaces","Plastic Bag (Regulation and Control) Directive, 2068 (2011)"),
 ("PDF note","Source PDF is image-based (2 pages, no extractable text). Coding from English news reports."),
 ("",""),
 ("KEY PROVISIONS — INSTRUMENT 1 (from Himalayan Times, Ratopati)",None),
 ("Provision","Details"),
 ("1. Thickness ban","Complete ban: production, import, storage, sale, distribution, use of bags <40 microns throughout Nepal"),
 ("2. Size standard (general bags)","Minimum 7 inches × 14 inches — only bags of this size or larger may be produced, imported, and sold"),
 ("3. Size standard (garbage bags)","Minimum 14 inches × 26 inches (some sources: 14×28) — minimum size for garbage bags"),
 ("4. Colour-coding",
  "• Natural/uncoloured: food, water, pharmaceutical use ONLY\n"
  "• Black: garbage ONLY (must state 'for garbage collection only')\n"
  "• Other colours: non-food items only"),
 ("5. Mandatory labelling",
  "Every bag must permanently display (minimum 15–20 pt font, permanent ink):\n"
  "• Producer's name and address\n"
  "• Recycling symbol/logo\n"
  "• Thickness in microns\n"
  "• Size\n"
  "• Whether made from virgin granules"),
 ("6. Material standard","High Molecular High Density Polyethylene (HMHDPE) specifications"),
 ("7. Producer obligations",
  "• Maintain calibrated measuring equipment\n"
  "• Keep records of production volumes, sizes, batch numbers\n"
  "• Manage production waste responsibly\n"
  "• Submit annual production details to Environment Department"),
 ("8. Importer obligations",
  "• Only import bags meeting prescribed standards (≥40 microns, minimum sizes)\n"
  "• Maintain records of quantities, sizes, and sales\n"
  "• Possess calibrated measuring instruments"),
 ("",""),
 ("KEY PROVISIONS — INSTRUMENT 2: EPR",None),
 ("Provision","Details"),
 ("EPR rate","At least 10 percent of annual production/import volume"),
 ("EPR obligation (producers)",
  "'producers must ensure proper disposal of at least ten percent of the plastic waste "
  "from their total production annually' (Himalayan Times)"),
 ("EPR obligation (importers)",
  "'importers must similarly comply with EPR obligations for at least ten percent of "
  "imported volumes annually' (Himalayan Times)"),
 ("Reporting","Annual reports to the Department of Environment"),
 ("Significance","Nepal's FIRST statutory EPR obligation for plastic bags"),
 ("",""),
 ("ENFORCEMENT",None),
 ("Provision","Details"),
 ("Enforcing body","Department of Environment (DoEnv) — empowered to deploy inspectors"),
 ("Legal action","Section 35(3) of EPA 2076 — 'department may take action against non-compliant entities'"),
 ("Inspectors","'The Environment Department is empowered to regulate compliance and deploy inspectors'"),
 ("Ministry powers","'Ministry may also issue directions to the department for effective implementation'"),
 ("Local bodies","'local bodies may adopt the [directive] as needed'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (ministerial directive under EPA 2076 §45; not Parliament-enacted)"),
 ("policy_target (E)","1  — EPR: ≥10% annual waste collection/management for producers and importers"),
 ("policy_integration (I)","1.00  (8+ sectors: industry, imports/trade, retail, waste management, packaging, recycling, food & beverage, chemicals)"),
 ("policy_circularity (K)","1.00  (ALL 5 phases: production, consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","0  (no budget or funding source stated in directive)"),
 ("",""),
 ("Instr. 1: Regulatory package (ban + 5 standards)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 2: EPR (≥10% annual waste collection)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("",""),
 ("COMPARISON WITH PREVIOUS DIRECTIVES",None),
 ("Aspect","2068 Directive (2011)  →  2082 Directive (2026)"),
 ("Thickness ban","<40 microns  →  <40 microns (same)"),
 ("Size standards","None  →  ≥7×14 in (general), ≥14×26 in (garbage) [NEW]"),
 ("Colour-coding","None  →  Mandatory food/garbage/non-food colour system [NEW]"),
 ("Labelling","Basic  →  Mandatory name, address, recycling symbol, microns, size, virgin granule indicator, 15-20pt font [ENHANCED]"),
 ("Material standard","None  →  HMHDPE specification [NEW]"),
 ("EPR","None  →  ≥10% annual collection obligation [NEW — first statutory EPR]"),
 ("Annual reporting","None  →  Annual production/import reports to DoEnv [NEW]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Provision","Composition","Value","Aspect") and row[0]=="Field":
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

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory (ban, EPR, mandatory standard)"),
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

path="/workspace/4p_index_pb_directive_2082_2026.xlsx"
wb.save(path)
print(f"Saved: {path}")
