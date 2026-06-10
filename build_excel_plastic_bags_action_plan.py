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
c.value=("Policy: Action Plan for Ban on Plastic Bags, 2022 (प्लाष्टिक झोला प्रतिबन्ध कार्ययोजना, २०७८ BS)  |  "
         "Country: Nepal  |  Year: 2022  |  "
         "Source: https://doenv.gov.np/content/25/action-plan-related-to-plastic-bag-prevention")
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
 name=("Action Plan for Ban on Plastic Bags, 2022\n"
       "[Nepali: प्लाष्टिक झोला प्रतिबन्ध कार्ययोजना, २०७८ BS]"),
 url="https://doenv.gov.np/content/25/action-plan-related-to-plastic-bag-prevention",
 year=2022,
 obj=("Implementation framework for the nationwide ban on plastic bags thinner than 40 microns "
      "throughout Nepal. Approved by the Council of Ministers and published by the Ministry of "
      "Forests and Environment on the Department of Environment (DoEnv) website. Objective: to "
      "ensure citizens' right to live in a clean and healthy environment through prevention and "
      "control of the use of plastic bags below 40 microns. Four strategies: (1) stop import of "
      "plastic bags/products thinner than 40 microns; (2) ban single-use plastic bags below 40 "
      "microns; (3) provide capital grants and technical support to industries transitioning to "
      "bags thicker than 40 microns or eco-friendly alternatives; (4) encourage consumers to use "
      "eco-friendly bags. Creates three-tier monitoring structure with enforcement and seizure powers."),
 tgt=0,
 tgt_t="",
 gtype=0.75,
 gtype_j=("Action plan approved by the Government of Nepal Council of Ministers (Cabinet level) in "
          "January 2022 (Magh 2078 BS). Published by the Ministry of Forests and Environment on "
          "the Department of Environment website (doenv.gov.np). Not enacted by Parliament, but "
          "approved at Cabinet level, conferring executive decree-level authority. "
          "Creates mandatory institutional structures (21-member Central Monitoring Committee "
          "led by Chief Secretary; provincial and local monitoring committees), seizure powers, "
          "and is backed by the Environment Protection Act 2076 (2019) penalty framework "
          "(fines up to Rs 300,000). Operationalises the Nepal Gazette notice of August 2021 "
          "banning plastic bags <40 microns nationwide.\n\n"
          "TRANSLATION NOTE: Source document is in Nepali with severely corrupted OCR. Coding "
          "based on: (a) partial Nepali text extraction; (b) English reports from The Himalayan "
          "Times, myRepublica; (c) SWITCH-Asia Nepal plastic policy overview (2025)."),
 intg=1,
 sects="industry, imports/trade, retail, waste management, municipalities, packaging, consumer goods",
 circ=0.75,
 lc="production, consumption, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Action plan provides 'grant to industries on the purchase of new machines for the production "
         "of plastic bags above 40 microns and eco-friendly alternative bags' and 'technical support to "
         "related industries' (The Himalayan Times, myRepublica). Specific budget amounts not stated "
         "in available English sources. Budget/funding mechanism exists (industry grants + technical "
         "support) but amounts are not ring-fenced in a dedicated fund."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC","E2EFDA"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("The action plan implements a COMPLETE BAN on the production, import, storage, sale or "
     "distribution, and use of plastic bags thinner than 40 microns throughout Nepal. "
     "This operationalises the Nepal Gazette notice of August 2021 (published by the "
     "Department of Environment under MoFE) which banned the production, import, sale, "
     "distribution and use of plastic bags <40 microns. The 40 micron threshold applies to "
     "ALL plastic bags nationwide. Violations subject to fines of up to Rs 300,000 under "
     "the Environment Protection Act 2076 (2019) and seizure of non-compliant products. "
     "The Department of Environment (DoEnv) designated as the primary enforcement body."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Department of Environment (DoEnv) under MoFE designated "
     "as enforcement body. Shankar Prasad Poudel, DoEnv spokesperson, explicitly stated: "
     "'This ban will be strictly enforced.' Central Monitoring Committee led by Chief Secretary "
     "of Government empowered to mobilise monitoring teams.\n\n"
     "Enforcement (+0.25): 'Violators may face legal action, including fines of up to Rs 300,000 "
     "under the Environment Protection Act, 2019' (multiple English news sources). 'Seizure of "
     "non-compliant plastic products' empowered. 'DoE has warned of stringent actions against "
     "those flouting the government's decision as per the Environment Protection Act, 2019 and "
     "the Environmental Protection Rule, 2020' (myRepublica, August 2022).\n\n"
     "Monitoring (+0.25): Three-tier monitoring committee structure — Central Monitoring Committee "
     "(21 members, led by Chief Secretary) empowered to 'mobilise monitoring teams to conduct "
     "monitoring as to whether or not plastic bags above 40 microns have been produced, collected, "
     "sold or distributed, stored and used.' (Himalayan Times). DoEnv stated 'from August 1, it "
     "will strictly monitor the production, import, sale, distribution and use of plastic bags "
     "thinner than 40 microns.'\n\n"
     "Unconditional (+0.25): Ban applies to ALL plastic bags <40 microns throughout Nepal with "
     "no stated product exemptions (unlike the Everest ban which had >30 micron exemption). "
     "'Provided for an immediate ban on the use of plastic thinner than 40 microns.' "
     "(myRepublica)"),
  W=("S = 1 (IN FORCE): The underlying ban was gazette-published in August 2021; this action plan "
     "provides the implementation framework and was approved by Council of Ministers in January 2022. "
     "Both the ban and the implementation framework are operative.\n\n"
     "T = 1.0: All four sub-scores met. Enforcement (+0.25) confirmed by EPA 2019 penalties "
     "(Rs up to 300,000 fines) and seizure powers explicitly referenced in multiple news sources. "
     "Unconditional (+0.25): No product-category exemptions mentioned — applies to ALL plastic bags "
     "<40 microns.\n\n"
     "IMPLEMENTATION REALITY: Despite strong institutional provisions, multiple news sources report "
     "weak enforcement in practice. SWITCH-Asia (2025): 'from conversations with policymakers it "
     "can be presumed that the ban is not being enforced.' This discrepancy between de jure strength "
     "(T=1.0) and de facto implementation does not affect the coding but is noted here.\n\n"
     "DISTINCTION from previous bans: Nepal had plastic bag bans in 2011 and 2015 (both weakly "
     "enforced). This 2022 action plan is the most institutionally robust attempt, with the three-"
     "tier monitoring committee structure being a significant new feature."),
 ),
 dict(
  P=0.20, Q="Waste management",
  R=("The action plan establishes a three-tier monitoring and coordination structure:\n"
     "(1) CENTRAL MONITORING COMMITTEE: 21 members, led by the Chief Secretary of the "
     "Government of Nepal. Functions: mobilise monitoring teams; suggest policy/legal/structural "
     "reforms; oversee provincial and local committees' progress.\n"
     "(2) PROVINCIAL MONITORING COMMITTEE: Coordinated by the Chief Secretary of the Office of "
     "the Chief Minister and Council of Ministers at the provincial level.\n"
     "(3) LOCAL MONITORING COMMITTEE: 5 members, coordinated by the Deputy Mayor/Vice President "
     "at the local level.\n"
     "All three committees have powers to seize non-compliant plastic products and report to the "
     "Department of Environment for further action under EPA 2019."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Central committee led by Chief Secretary of Government; "
     "provincial committee by Provincial Chief Secretary; local committee by Deputy Mayor/Vice "
     "President. Explicit three-tier structure covering all levels of Nepal's federal system.\n\n"
     "Enforcement (+0.25): Committees empowered to 'seize non-compliant plastic products' and "
     "refer violators to DoEnv for EPA 2019 fines (up to Rs 300,000). Enforcement action "
     "directly linked to committee monitoring activities.\n\n"
     "Monitoring (+0.25): The committees' core function IS monitoring — 'mobilise monitoring "
     "teams to conduct monitoring as to whether or not plastic bags above 40 microns have been "
     "produced, collected, sold or distributed, stored and used' (Himalayan Times). 'Take stock "
     "of progress of provincial monitoring committee and local monitoring committee.'\n\n"
     "Unconditional (+0.25): Three-tier structure covers central, provincial AND local levels — "
     "comprehensive coordination mandate with no stated exemptions."),
  W=("This governance instrument is the institutional innovation in the 2022 action plan relative "
     "to previous Nepal plastic bag bans. It creates Nepal's first three-tier plastic bag "
     "enforcement committee structure under the 2015 federal governance system, with distinct "
     "roles for federal, provincial, and local governments.\n\n"
     "P = 0.20 (Governance & coordination) — creates coordinating bodies and assigns enforcement "
     "responsibilities across levels of government (federal–provincial–local coordination "
     "framework). The committee has enforcement and seizure powers, but these are governance "
     "mechanisms rather than direct regulatory instruments.\n\n"
     "T = 1.0: All four sub-scores met — authority (explicitly designated), enforcement (seizure "
     "and referral powers), monitoring (core function), unconditional (covers all three tiers "
     "with no exemptions)."),
 ),
 dict(
  P=0.60, Q="Production",
  R=("The action plan provides capital grants and technical support to plastic bag manufacturing "
     "industries to facilitate transition from thin (<40 micron) plastic bag production to: "
     "(a) bags thicker than 40 microns, or (b) eco-friendly alternative bags. Specifically: "
     "'providing grant to industries on the purchase of new machines for the production of plastic "
     "bags above 40 microns and eco-friendly alternative bags' (Himalayan Times). Also includes "
     "'technical support to related industries' (myRepublica). This creates a positive economic "
     "incentive to support the production-side transition away from thin plastic bags."),
  S=1,
  T=0.50,
  U=("Responsible authority (+0.25): MoFE/DoEnv designated as administering body for grants "
     "and technical support; specific budget not stated in available sources.\n\n"
     "Enforcement (−0): No penalties for industries that do NOT apply for/accept grants; "
     "the incentive is voluntary — industries choose whether to participate.\n\n"
     "Monitoring (+0.25): Progress of industry transitions tracked through the three-tier "
     "monitoring committee (Central Committee functions include tracking implementation progress).\n\n"
     "Unconditional (−0): Grant is conditional — only for industries that apply and meet "
     "eligibility criteria for transition to >40 micron or eco-friendly production."),
  W=("P = 0.60 (Economic — subsidy/grant): Capital grants for machine purchase and technical "
     "support constitute an economic instrument (positive incentive). This is distinct from "
     "the regulatory ban (Instrument 1) and the governance structure (Instrument 2).\n\n"
     "S = 1: Council of Ministers approved this economic incentive alongside the ban; it is "
     "an operative provision of the action plan (not an aspiration to be considered later).\n\n"
     "T = 0.50: Responsible authority and monitoring sub-scores met. Enforcement (−0) because "
     "the grant is voluntary — there is no penalty for not transitioning. Unconditional (−0) "
     "because the grant is conditional on applying and being eligible.\n\n"
     "COMBINED INSTRUMENT NOTE: This instruments combines (i) production substitution standard "
     "(encouraging >40 micron bags) and (ii) economic subsidy. The highest applicable score "
     "for a tax/subsidy is 0.60, applied here. The production standard element (regulatory, "
     "1.0) is already captured in Instrument 1 (the ban) which implicitly mandates that any "
     "bags produced must be ≥40 microns."),
 ),
 dict(
  P=0.40, Q="Consumption",
  R=("The action plan includes a consumer awareness strategy: 'encouraging the general public "
     "to carry their own eco-friendly bags for shopping' as one of the four core strategies. "
     "This involves awareness campaigns, information dissemination, and behavioural change "
     "programmes targeting consumers to voluntarily adopt eco-friendly bag alternatives. "
     "The Central Monitoring Committee is also mandated to 'provide suggestions... for "
     "necessary policy, legal and structural reforms' — which includes awareness raising "
     "as a policy lever."),
  S=0,
  T=0.50,
  U=("Responsible authority (+0.25): DoEnv/MoFE and three-tier committees tasked with "
     "awareness activities as part of the action plan mandate.\n\n"
     "Enforcement (−0): Consumer awareness is voluntary — no penalties for consumers who "
     "do not adopt eco-friendly bags (though they are prohibited from using/purchasing banned "
     "<40 micron bags, the awareness campaign is the information element, not enforcement).\n\n"
     "Monitoring (+0.25): Monitoring committees track overall implementation progress, which "
     "implicitly includes consumer behaviour change indicators.\n\n"
     "Unconditional (+0.25): Awareness campaign applies to all consumers, general public, "
     "and stakeholders without stated exemptions."),
  W=("S = 0: The consumer awareness strategy uses advisory/aspirational language "
     "('encouraging the general public'). It is an enabling measure, not a mandatory "
     "obligation with legal consequences for non-compliance. The actual ban (Instrument 1) "
     "covers consumer use, but the awareness campaign element is informational.\n\n"
     "P = 0.40 (Information & voluntary — awareness campaign): One of the four stated "
     "strategies is consumer awareness. This is the weakest instrument in the action plan "
     "and functions as an enabling measure for the regulatory ban.\n\n"
     "T = 0.50: Responsible authority and monitoring sub-scores credited. Enforcement not "
     "credited (no penalties for low awareness). Unconditional credited (no stated "
     "exemptions to the campaign scope)."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Nationwide ban on plastic bags <40 microns (production/import/storage/sale/distribution/use)",
 "Instrument 2 — Three-tier monitoring committee (Central 21-member + Provincial + Local 5-member)",
 "Instrument 3 — Industry transition grants and technical support (>40 micron / eco-friendly bags)",
 "Instrument 4 — Consumer awareness campaign for eco-friendly bag adoption",
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
 ws.row_dimensions[r].height=230

for c,w in {1:32,2:38,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:35,15:9,
            16:10,17:18,18:48,19:10,20:10,21:48,22:9,23:48,24:30}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence & Translation Notes")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Official Nepali title","प्लाष्टिक झोला प्रतिबन्ध कार्ययोजना, २०७८ BS (Action Plan for Ban on Plastic Bags, 2078 BS)"),
 ("Policy type","Action Plan (कार्ययोजना)"),
 ("Released","January 2022 (Magh 2078 BS)"),
 ("Approved by","Government of Nepal Council of Ministers (Cabinet level)"),
 ("Published on","Ministry of Forests and Environment — Department of Environment (doenv.gov.np)"),
 ("Implements","Nepal Gazette notice (August 2021) banning plastic bags <40 microns under EPA 2076"),
 ("Issuing ministry","Ministry of Forests and Environment (MoFE)"),
 ("Primary enforcement body","Department of Environment (DoEnv)"),
 ("",""),
 ("FOUR CORE STRATEGIES (from Himalayan Times English report)",None),
 ("Strategy","Description"),
 ("1. Stop imports",
  "Stopping the import of plastic bags or other plastic products thinner than 40 microns"),
 ("2. Ban single-use bags",
  "Banning the import of single-use plastic bags below 40 microns"),
 ("3. Industry grants",
  "Providing grants to industries on purchase of new machines for production of plastic bags "
  "above 40 microns and eco-friendly alternative bags"),
 ("4. Consumer campaign",
  "Encouraging the general public to carry their own eco-friendly bags for shopping"),
 ("",""),
 ("THREE-TIER MONITORING COMMITTEE STRUCTURE",None),
 ("Level","Composition and Role"),
 ("Central Committee",
  "21 members, led by Chief Secretary of the Government of Nepal. "
  "Functions: mobilise monitoring teams, suggest policy/legal reforms, oversee "
  "provincial and local committee progress, and report to GoN."),
 ("Provincial Committee",
  "Coordinated by Chief Secretary of the Office of the Chief Minister and Council of "
  "Ministers at provincial level. Monitors production, sale, distribution at province."),
 ("Local Committee",
  "5 members, coordinated by Deputy Mayor / Vice President at local government level. "
  "First-line monitoring of shops, markets, and consumers."),
 ("Enforcement powers","ALL committees: empowered to seize non-compliant plastic products"),
 ("",""),
 ("ENFORCEMENT PROVISIONS",None),
 ("Provision","Details"),
 ("Fine (max)","Rs 300,000 under Environment Protection Act 2076 (2019)"),
 ("Seizure","Non-compliant plastic products may be seized by monitoring committees"),
 ("EPA authority","Actions taken per Environment Protection Act 2076 and Environmental Protection Rules 2077"),
 ("Spokesperson",
  "Shankar Prasad Poudel (DoEnv): 'This ban will be strictly enforced.' (myRepublica)"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (Cabinet-approved action plan — executive decree level)"),
 ("policy_target (E)","0  (40-micron threshold is a product standard, not a % reduction target)"),
 ("policy_integration (I)","1.00  (7+ sectors: industry, imports/trade, retail, waste management, municipalities, packaging, consumer goods)"),
 ("policy_circularity (K)","0.75  (production, consumption, disposal, environmental leakage — 4 phases)"),
 ("policy_budget (M)","0.50  (industry grants mentioned; specific budget amounts not stated)"),
 ("",""),
 ("Instr.1: Ban <40 micron bags",          "P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr.2: Three-tier monitoring committee","P=0.20 | S=1 | T=1.00 | V=[auto 0.600]"),
 ("Instr.3: Industry transition grants",    "P=0.60 | S=1 | T=0.50 | V=[auto 0.550]"),
 ("Instr.4: Consumer awareness campaign",   "P=0.40 | S=0 | T=0.50 | V=[auto 0.450]"),
 ("",""),
 ("TRANSLATION NOTE",None),
 ("Note",
  "Source document is in Nepali with severely corrupted OCR. Coding based on: (1) partial "
  "Nepali text extraction; (2) English reports from The Himalayan Times and myRepublica; "
  "(3) SWITCH-Asia Nepal plastic policy overview (2025). SWITCH-Asia (2025) notes: "
  "'from conversations with policymakers it can be presumed that the ban is not being "
  "enforced' — significant gap between de jure strength (T=1.0 for ban) and de facto "
  "implementation."),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Description","Composition and Role","Details","Value") and row[0]=="Field":
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
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

# Sheet 3 — Score Reference
ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic (taxes, levies, subsidies)"),(0.80,"Infrastructure"),
     (1.00,"Regulatory (ban, EPR, mandatory standard)"),
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

path="/workspace/4p_index_plastic_bags_action_plan_2022.xlsx"
wb.save(path)
print(f"Saved: {path}")
