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
c.value=("Policy: Environment-friendly Local Governance Framework, 2013  |  Country: Nepal  |  Year: 2013  |  "
         "Source: https://www.google.com/u... [URL truncated in source image]")
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
 name="Environment-friendly Local Governance Framework, 2013",
 url=("https://www.google.com/u... [URL truncated in source image — "
      "document approved by Ministerial Council, Government of Nepal, September 2013; "
      "published by Ministry of Federal Affairs and Local Development, Kathmandu]"),
 year=2013,
 obj=("Comprehensive indicator-based framework for environment-friendly local governance "
      "at household, tole/settlement, ward, municipality, VDC, and district levels throughout "
      "Nepal. Approved by the Ministerial Council (Cabinet) in September 2013. "
      "In relation to plastics:\n"
      "• HOUSEHOLD: Mandatory plastic waste segregation at source; plastic bag regulation "
      "(use jute/cloth/paper instead)\n"
      "• WARD: Ward office must purchase collected plastics from households\n"
      "• MUNICIPALITY: Plastic ban by municipal council decision; separate landfill after "
      "plastic pre-separation; plastic management training (≥2 per ward); "
      "environment protection special fund\n"
      "• VDC: Household plastic segregation; plastic use discouraged; VDC plastic management "
      "training; eco-friendly bag production encouraged\n"
      "• DISTRICT: Declaration of plastic-bag-free district by District Council; "
      "district-level mechanism to control import of prohibited plastic bags; "
      "environment-friendly bag production encouraged"),
 tgt=1,
 tgt_t=("'At least 2 people in the ward trained for plastic collection, re-use and "
        "management; Trained person deployed for the plastic collection.' "
        "(Advanced Indicators for Municipality, Section e — Plastic Management, p.16)\n"
        "'At least two people provided with the training on plastic collection and reuse "
        "in each ward; Trained human resource employed in related activity.' "
        "(VDC Advanced Indicators, Section d — Proper Management of Plastic, p.22)"),
 gtype=0.75,
 gtype_j=("Approved by the Ministerial Council (Cabinet) of the Government of Nepal in "
          "September 2013. Issued by the Ministry of Federal Affairs and Local Development. "
          "Not enacted by Parliament — constitutes a Cabinet-approved framework with binding "
          "force for local governments through the performance certification system. "
          "Creates mandatory indicators for local governance certification across six levels "
          "(household, tole, ward, municipality, VDC, district). Has a full multi-tier "
          "institutional structure from Central Directive Committee (chaired by Vice "
          "Chairperson of National Planning Commission) to household level.\n\n"
          "Scored 0.75 (executive decree level): Cabinet approval gives this framework "
          "binding executive authority over local government performance evaluation. "
          "Local governments must comply to receive certification and associated rewards/recognition."),
 intg=1,
 sects="waste management, municipalities, water/sanitation, agriculture, energy, disaster management, forest/biodiversity, education, industry, retail/packaging",
 circ=1.0,
 lc="production, consumption, recycling, disposal, environmental leakage",
 budg=1.0,
 budg_t=("Advanced Indicators for Municipality, Section c (p.16): 'Environment protection "
         "special fund established' — explicitly requires municipalities to establish a "
         "ring-fenced environment protection special fund as a condition for advanced-level "
         "certification. Also: Section 6.5 'Financial Resource Management' in Chapter 6 "
         "implementation arrangements (p.38 area). The environment protection special fund "
         "is explicitly ring-fenced for environmental protection purposes within municipalities."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("COMPREHENSIVE MULTI-LEVEL PLASTIC BAG RESTRICTION AND BAN SYSTEM:\n\n"
     "HOUSEHOLD Level (Municipal + VDC, Basic & Advanced):\n"
     "Basic: 'Among the inorganic (non-decomposing) wastes, plastics collected separately "
     "and sold or gathered in a fixed or specified location' (p.11, p.19)\n"
     "Advanced: 'Plastic Bag Regulation: Instead of all sorts of plastic bags; jute, cloth or "
     "paper bags used and written commitment towards it submitted to the Tole Development "
     "Organization.' (p.12, p.20) AND 'Discouraging the use of plastic: Plastic bag "
     "regulation: Jute, cloth or paper bags used instead of all sorts of plastic bags.' (p.20)\n\n"
     "TOLE Level (Advanced): 'Cleaning of public places: No non-decomposing wastes, "
     "plastics and open defecation within the public places of one's Tole such as road, "
     "drainage, alleys, footpaths, streams, river, fields, forest, land etc.' (p.13)\n\n"
     "MUNICIPALITY Level (Basic): 'Ban on plastic use: Through the decision of municipal "
     "counsel, use of plastic restricted as per the national standard approved by the "
     "Government of Nepal. As an alternative to the plastic bags, environment-friendly "
     "bags encouraged.' (p.15)\n\n"
     "VDC Level (Basic): 'Plastic use discouraged' + eco-friendly bags encouraged (p.21)\n"
     "VDC Advanced: 'Environment-friendly bags encouraged to manufacture as an alternative "
     "to the plastic bags' (p.22)\n\n"
     "DISTRICT Level (Basic): 'Control on Plastic Bags: Declaration of plastic bag free "
     "district: Decision from the District Council issued on banning plastic bags in district "
     "as prescribed by the Government of Nepal; District level mechanism established to "
     "control the import of prohibited plastic bags; Environment-friendly bag production "
     "encouraged as an alternative to the plastic bags.' (p.25)"),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Six-level institutional structure explicitly designated — "
     "Central Directive Committee (VP of NPC as president) → Central Implementation "
     "Coordination Committee → District Committee → Municipality/VDC Committee → Ward "
     "→ Tole/Household. Municipal Council and District Council explicitly responsible for "
     "plastic ban decisions. Ministry of Federal Affairs and Local Development is the "
     "nodal ministry.\n\n"
     "Enforcement (+0.25): Certification system creates binding incentive — local governments "
     "must achieve indicators to receive certification, rewards, and recognition. "
     "District Committee: 'Give recommendation to concerned institution to provide cash "
     "reward and honor the individual, organizations/institutions contributing considerably "
     "towards environmental protection, waste management' (p.40). Compliance monitoring "
     "is built into the performance measurement system.\n\n"
     "Monitoring (+0.25): 'Provision a system to send the numbers of environment-friendly "
     "household, tole and ward...to the Ministry of Federal Affairs and Local Development in "
     "a quarterly basis' (p.42). Tables 18 and 19 (in ToC p.5) describe monitoring and "
     "measurement methods. Annual and quarterly reporting to central level.\n\n"
     "Unconditional (+0.25): Applies to ALL levels of governance throughout Nepal — "
     "household, tole, ward, municipality, VDC, district — without geographic or "
     "population-based exceptions. The restriction applies 'as per the national standard "
     "approved by the Government of Nepal' at all levels."),
  W=("S = 1 (IN FORCE): Cabinet-approved framework with binding force over local "
     "governments. The municipal council and district council decisions on plastic bans "
     "use mandatory language ('through the decision of municipal counsel, use of plastic "
     "restricted'). The district-level indicator requires 'Decision from the District "
     "Council issued on banning plastic bags' — this IS a binding requirement within the "
     "certification framework.\n\n"
     "T = 1.0: All four sub-scores met with explicit evidence throughout the document. "
     "The quarterly reporting system to MoFALD provides a comprehensive monitoring "
     "mechanism. Cash rewards and recognition for compliance constitute enforcement.\n\n"
     "SIGNIFICANCE: This is the most comprehensive plastic reduction framework across all "
     "levels of Nepal's governance structure in this dataset. It creates a cascading system "
     "where household plastic bag regulation, tole-level plastic-free public spaces, ward "
     "plastic purchasing, municipal bans, and district-level plastic-free declarations form "
     "a vertically integrated plastic governance architecture.\n\n"
     "INSTRUMENT SCOPE: Combines aspects of: information (awareness campaigns), regulatory "
     "(bans at municipal/district level), and governance (multi-level coordination). "
     "Highest applicable type = Regulatory (1.0) applied; governance aspect noted here."),
 ),
 dict(
  P=1.0, Q="Recycling",
  R=("MULTI-LEVEL PLASTIC COLLECTION, PURCHASE AND RECYCLING SYSTEM:\n\n"
     "WARD Level — Municipal (Basic): 'Waste management: Plastics collected from home "
     "purchased by the ward office itself or made provisions to purchase by any other "
     "organization or person.' (p.14) — Ward offices must purchase plastics from households, "
     "creating a reverse logistics system.\n\n"
     "WARD Level — VDC (Basic): 'Plastics collected from home purchased by some organization "
     "or person, or collected in the ward and made provisions to send to the VDC.' (p.20)\n\n"
     "MUNICIPALITY Level — Advanced Indicators, Section e: 'Plastic Management: At least "
     "2 people in the ward trained for plastic collection, re-use and management; Trained "
     "person deployed for the plastic collection; Re-using of the plastics and its produced "
     "materials supported by the municipality.' (p.16)\n\n"
     "VDC Level — Advanced Indicators, Section d: 'Proper Management of Plastic: At least "
     "two people provided with the training on plastic collection and reuse in each ward; "
     "Trained human resource employed in related activity.' (p.22)\n\n"
     "MUNICIPALITY Level — Basic: 'Provision of sanitary landfill site for the proper and "
     "appropriate management of remaining wastes after the separation of decomposing wastes "
     "and plastics.' (p.15) — Plastic pre-separation from general waste before landfill.\n\n"
     "MUNICIPALITY Level — Advanced: 'Scientific management of sanitary landfill site: All "
     "the process of waste collection, transportation, and discharge done in an environment-"
     "friendly way; incineration, recycling, bio-gas production etc. adopted.' (p.16)"),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Ward offices explicitly responsible for plastic "
     "purchasing. Municipality designated as the supporting body for recycling materials "
     "and plastic management training. VDC responsible for VDC-level plastic management.\n\n"
     "Enforcement (+0.25): Certification system enforces compliance — ward offices MUST "
     "purchase or arrange purchase of collected plastics as a BASIC indicator requirement "
     "(non-optional for any ward seeking environment-friendly certification). The 'at least "
     "2 people per ward trained' quantified target creates measurable accountability.\n\n"
     "Monitoring (+0.25): Quarterly reporting to Ministry of Federal Affairs and Local "
     "Development (p.42). Central Implementation Coordination Committee monitors "
     "'Programme coordination, facilitation and monitoring regionally' (p.40). "
     "The specific number of trained people ('at least 2') provides a measurable indicator.\n\n"
     "Unconditional (+0.25): Ward plastic purchasing is a BASIC indicator (mandatory "
     "minimum, not optional advanced level). Applies to all wards throughout Nepal without "
     "exceptions. The 'at least 2 people per ward' standard applies uniformly."),
  W=("P = 1.0 (Regulatory — mandatory EPR-like take-back obligation): The requirement for "
     "ward offices to PURCHASE collected plastics from households constitutes a mandatory "
     "take-back/reverse logistics obligation — functionally equivalent to EPR at the local "
     "government level. Ward offices act as the 'producer responsibility organization'.\n\n"
     "S = 1: Mandatory indicator language throughout — ward purchasing is a BASIC (mandatory) "
     "indicator, not just an advanced option. 'Must' and 'shall' language applies through "
     "the certification framework.\n\n"
     "T = 1.0: All four sub-scores met — authority (ward offices + municipality), enforcement "
     "(certification system), monitoring (quarterly reporting + quantified targets), "
     "unconditional (basic indicator for all wards).\n\n"
     "This instrument is particularly innovative for Nepal's policy landscape — creating "
     "a reverse logistics system where ward offices financially incentivize household plastic "
     "segregation by purchasing the collected plastics. This predates Nepal's later EPR "
     "framework discussions by almost a decade.\n\n"
     "POLICY_TARGET: The quantified target ('at least 2 people per ward') in col F refers "
     "to this instrument. It is the most clearly quantifiable plastic-specific target in "
     "this framework."),
 ),
 dict(
  P=0.60, Q="Waste management",
  R=("ENVIRONMENT PROTECTION SPECIAL FUND:\n\n"
     "Municipal Advanced Indicators, Section c — Renewable Energy (p.16-17):\n"
     "'Environment protection special fund established' — municipalities seeking advanced "
     "certification must establish a dedicated/ring-fenced environment protection fund.\n\n"
     "This fund is specifically designated for environmental protection activities including "
     "waste management, plastic control, and environmental monitoring — providing a "
     "dedicated financial mechanism to support the plastic management activities required "
     "under this framework.\n\n"
     "Section 6.5 (Chapter 6 — Implementation): 'Financial Resource Management' covers "
     "the broader financial architecture of the framework, including central, provincial, "
     "and local-level funding."),
  S=1,
  T=0.50,
  U=("Responsible authority (+0.25): Municipality designated as the authority to establish "
     "and manage the fund. Ministry of Federal Affairs and Local Development (MoFALD) "
     "provides central guidance on financial resource management (Section 6.5).\n\n"
     "Enforcement (−0): No specific penalties for municipalities that fail to establish "
     "the fund, other than not achieving advanced certification. The fund is tied to the "
     "incentive-based certification system.\n\n"
     "Monitoring (+0.25): The certification monitoring system checks whether the fund has "
     "been established as part of advanced-level assessment. Quarterly reporting to MoFALD "
     "includes financial management indicators.\n\n"
     "Unconditional (−0): The fund is required only for municipalities seeking "
     "ADVANCED certification — it is not a BASIC indicator (not required at the minimum "
     "certification level). This conditionality prevents the 'unconditional' sub-score."),
  W=("P = 0.60 (Economic — dedicated fund): The environment protection special fund is an "
     "economic instrument providing ring-fenced financial resources for environmental "
     "protection activities including plastic management. It generates dedicated funding "
     "through the municipal budget allocation process.\n\n"
     "S = 1: The fund requirement is mandatory for advanced-level certification. Once a "
     "municipality seeks and achieves advanced certification, the fund is operative and "
     "binding. Coded as S=1 since the framework is in force and the indicator creates "
     "a binding obligation for compliant municipalities.\n\n"
     "T = 0.50: Authority (+0.25) and monitoring (+0.25) credited. Enforcement not "
     "credited (no penalties; incentive-based). Unconditional not credited (only required "
     "at advanced certification level, not basic).\n\n"
     "SIGNIFICANCE: The environment protection special fund provides the financial "
     "backbone for all other environmental instruments in the framework, including plastic "
     "management. It is listed separately from the general municipal budget, indicating "
     "intent for ring-fencing."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Multi-level plastic bag restriction/ban (Household → Tole → Ward → Municipality → VDC → District)",
 "Instrument 2 — Plastic collection purchase + recycling system (Ward purchase + ≥2 trained persons per ward + landfill pre-separation)",
 "Instrument 3 — Environment protection special fund (Municipal Advanced Indicator)",
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
 ws.row_dimensions[r].height=280

for c,w in {1:32,2:38,3:8,4:35,5:8,6:40,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Evidence
ws2=wb.create_sheet("Key Evidence — Full Plastic Provisions")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Environment-friendly Local Governance Framework, 2013"),
 ("Approval","Approved by the Ministerial Council (Cabinet), Government of Nepal, September 2013"),
 ("Published by","Ministry of Federal Affairs and Local Development, Kathmandu"),
 ("Coverage","6 governance levels: Household → Tole → Ward → Municipality → VDC → District"),
 ("Structure","Basic Indicators (mandatory minimum) + Advanced Indicators (higher certification)"),
 ("",""),
 ("ALL PLASTIC-SPECIFIC PROVISIONS (verbatim)",None),
 ("Level / Section","Text"),
 ("Municipal Household — Basic p.11",
  "'Among the inorganic (non-decomposing) wastes, plastics collected separately and sold or "
  "gathered in a fixed or specified location'"),
 ("Municipal Household — Advanced p.12",
  "'Plastic Bag Regulation: Instead of all sorts of plastic bags; jute, cloth or paper bags "
  "used and written commitment towards it submitted to the Tole Development Organization.'"),
 ("Tole — Advanced p.13",
  "'Cleaning of public places: No non-decomposing wastes, plastics and open defecation "
  "within the public places of one's Tole such as road, drainage, alleys, footpaths, streams, "
  "river, fields, forest, land etc.'"),
 ("Ward — Basic p.14",
  "'Plastics collected from home purchased by the ward office itself or made provisions to "
  "purchase by any other organization or person'"),
 ("Municipality — Basic p.15",
  "'Provision of sanitary landfill site for the proper and appropriate management of remaining "
  "wastes after the separation of decomposing wastes and plastics.'\n"
  "'Ban on plastic use: Through the decision of municipal counsel, use of plastic restricted "
  "as per the national standard approved by the Government of Nepal. As an alternative to "
  "the plastic bags, environment-friendly bags encouraged.'"),
 ("Municipality — Advanced p.16 (KEY: QUANTIFIED)",
  "'Plastic Management (Section e): At least 2 people in the ward trained for plastic "
  "collection, re-use and management; Trained person deployed for the plastic collection; "
  "Re-using of the plastics and its produced materials supported by the municipality.'"),
 ("Municipality — Advanced p.16",
  "'Environment protection special fund established'"),
 ("VDC Household — Basic p.19",
  "'Among the non-decomposing wastes, plastics collected separately and sold or gathered "
  "in a specified location'"),
 ("VDC Household — Advanced p.20",
  "'Plastic bag regulation: Jute, cloth or paper bags used instead of all sorts of plastic bags, "
  "and a written commitment of the same submitted to the Tole Development Organization.'"),
 ("VDC Ward — Basic p.20",
  "'Plastics collected from home purchased by some organization or person, or collected in "
  "the ward and made provisions to send to the VDC'"),
 ("VDC — Advanced p.22",
  "'Sanitary landfill site management: Waste, except which is decomposable, recyclable or "
  "reusable, managed properly... Environment-friendly bags encouraged to manufacture as an "
  "alternative to the plastic bags'\n"
  "'Proper Management of Plastic (Section d): At least two people provided with the training "
  "on plastic collection and reuse in each ward; Trained human resource employed in "
  "related activity'"),
 ("District — Basic p.25 (KEY: BINDING BAN)",
  "'Control on Plastic Bags: Declaration of plastic bag free district: Decision from the "
  "District Council issued on banning plastic bags in district as prescribed by the Government "
  "of Nepal; District level mechanism established to control the import of prohibited plastic "
  "bags; Environment-friendly bag production encouraged as an alternative to the plastic bags.'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (Cabinet/Ministerial Council approved; binding on local governments)"),
 ("policy_target (E)","1  — 'At least 2 people in the ward trained for plastic collection, re-use and management'"),
 ("policy_integration (I)","1.00  (10+ sectors)"),
 ("policy_circularity (K)","1.00  (ALL 5 lifecycle phases: production, consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","1.00  (environment protection special fund explicitly required — ring-fenced)"),
 ("",""),
 ("Instr. 1: Multi-level plastic ban/restriction system","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 2: Plastic collection purchase + recycling (EPR-like)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 3: Environment protection special fund","P=0.60 | S=1 | T=0.50 | V=[auto 0.550]"),
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
 ws2.row_dimensions[ri].height=55
ws2.column_dimensions["A"].width=32
ws2.column_dimensions["B"].width=90

ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree (sub-legislative)"),(1.00,"Legislation (parliament)"),
     ("",""),("INSTRUMENT TYPE (Col P)",None),("Score","Description"),
     (0,"No instrument"),(0.20,"Governance & coordination"),(0.40,"Information & voluntary"),
     (0.60,"Economic"),(0.80,"Infrastructure"),(1.00,"Regulatory"),
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

path="/workspace/4p_index_eflg_framework_2013.xlsx"
wb.save(path)
print(f"Saved: {path}")
