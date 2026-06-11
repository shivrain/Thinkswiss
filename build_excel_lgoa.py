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
c.value=("Policy: Local Government Operation Act, 2074 BS (2017), as amended in 2081 BS (2025)  |  "
         "Country: Nepal  |  Year: 2025  |  Source: https://lawcommission.gov.np")
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
 name=("Local Government Operation Act, 2074 BS (2017 CE)\n"
       "[स्थानीय सरकार सञ्चालन ऐन, २०७४]\n"
       "Act No. 26 of 2074 BS — as amended in 2081 BS (2025 CE)\n"
       "[Most recent amendment: Good Governance and Public Service Act, 2081/12/18 BS ≈ April 2025]"),
 url=("https://lawcommission.gov.np [full Nepali text from Nepal Law Commission]\n"
      "Image shows URL: https://lawcommission.gov.np/conte..."),
 year=2025,
 obj=("Comprehensive legislation governing the operation of local governments (municipalities "
      "and rural municipalities) in Nepal under the federal structure. In relation to plastics "
      "and waste: the Act mandates that local governments perform solid waste management "
      "functions (Schedule, Municipal Functions झ.7 and ञ.15); control the sale and "
      "consumption of consumer goods having adverse effects on public health including "
      "environmental pollution and hazardous substances at the local level (ञ.14) — providing "
      "the statutory basis for local plastic bans; adopt environment-friendly and low-carbon "
      "development (ञ.16); collect service fees for waste management services (Section 62); "
      "and manage health waste collection, reuse, processing and disposal with fee regulation "
      "(झ.7). The Act also requires municipalities to have a solid waste processing and "
      "management system as a condition for classification as a municipality (Section 5)."),
 tgt=0,
 tgt_t="",
 gtype=1.0,
 gtype_j=("Legislation enacted by the Legislative Parliament (व्यवस्थापिका-संसद) of Nepal "
          "under Article 296(1) of the Constitution of Nepal. Act No. 26 of 2074 BS (2017 CE). "
          "Originally enacted in 2074 BS (2017 CE). Amended four times:\n"
          "1. Finance Act, 2075 BS (2018/19 CE) — 2075/03/32\n"
          "2. Some Nepal Acts Amendment Act, 2075 BS — 2075/11/19\n"
          "3. Federal, Provincial and Local Coordination Act, 2077 BS — 2077/04/13\n"
          "4. Good Governance and Public Service Flow Act, 2081 BS — 2081/12/18 (≈ April 2025) [MOST RECENT]\n\n"
          "G = 1.0 (Legislation): Enacted by the full Legislative Parliament as a major "
          "governance law under the new federal constitution. Highest possible policy type."),
 intg=1,
 sects="waste management, municipalities, water/sanitation, environment, agriculture, industry, healthcare, education, tourism, urban development",
 circ=0.75,
 lc="consumption, disposal, recycling, environmental leakage",
 budg=0.5,
 budg_t=("Section 62(2)(ख): 'Service charges may be levied on service users for: (ख) "
         "Service facilities such as waste management, sanitation, drainage, street lighting.' "
         "[Translation: गाउँपालिका तथा नगरपालिकाले...फोहरमैला व्यवस्थापन, सरसफाई, ढल "
         "निकास, सडक बत्ती जस्ता सेवा सुविधा...सेवा शुल्क लगाउन सक्नेछ] "
         "Local governments have mandatory authority to levy service fees for waste "
         "management but this is not a ring-fenced dedicated plastic waste fund. M = 0.5."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC"]
instruments=[
 dict(
  P=1.0, Q="Waste management",
  R=("SCHEDULE — MUNICIPAL FUNCTIONS (Section झ: Basic Health and Sanitation):\n"
     "(6) Raising sanitation awareness and management of health-related waste [स्वास्थ्यजन्य "
     "फोहोरमैलाको व्यवस्थापन]\n"
     "(7) Collection, reuse, processing, disposal of health waste and determination and "
     "regulation of service charges therefor [स्वास्थ्यजन्य फोहरमैला सङ्कलन, पुनः उपयोग, "
     "प्रशोधन, विसर्जन र सोको सेवा शुल्क निर्धारण र नियमन]\n"
     "(10) Coordination, cooperation and partnership with private and non-governmental "
     "sector for management of waste from sanitation and health sector [सरसफाई तथा "
     "स्वास्थ्य क्षेत्रबाट निष्कासित फोहोरमैला व्यवस्थापनमा निजी तथा गैरसरकारी क्षेत्रसँग "
     "समन्वय, सहकार्य र साझेदारी]\n\n"
     "SCHEDULE — MUNICIPAL FUNCTIONS (Section ञ: Local Market, Environment Conservation "
     "and Biodiversity):\n"
     "(15) Sanitation and waste management at the local level [स्थानीयस्तरमा सरसफाई "
     "तथा फोहरमैला व्यवस्थापन]\n"
     "(16) Adopting low-carbon and environment-friendly development at the local level\n\n"
     "WARD FUNCTIONS (Section 33 area):\n"
     "(11) Collection and management of waste discharged from homes, cleaning of squares "
     "and alleys, drainage management [घरबाट निकास हुने फोहरमैलाको सङ्कलन र व्यवस्थापन]\n\n"
     "MUNICIPALITY CLASSIFICATION (Section 5): Having a 'solid waste processing and "
     "management system' is listed as one of the criteria for a locality to be classified "
     "as a municipality."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Local governments (municipalities, rural municipalities, "
     "ward committees) explicitly designated as responsible for waste management functions "
     "throughout the Schedule. The Act creates a three-tier governance system: municipal, "
     "ward, and household level responsibilities for waste management.\n\n"
     "Enforcement (+0.25): The LGOA gives local governments enforcement powers including "
     "the ability to impose penalties for violations of local laws and regulations. "
     "Section 62 authorizes service fee collection; non-payment can result in service "
     "suspension. The national government (Ministry of Federal Affairs and Local "
     "Development) provides oversight and can direct local governments.\n\n"
     "Monitoring (+0.25): The Act requires local governments to prepare annual plans and "
     "budgets, submit progress reports to provincial and federal governments, and undergo "
     "auditing. The waste management functions are included in the mandatory performance "
     "measurement system for local governments.\n\n"
     "Unconditional (+0.25): The waste management functions are listed as MANDATORY "
     "functions in the Schedule — all municipalities and rural municipalities are required "
     "to perform these functions without exemption. The Act uses 'गर्ने' (shall do) "
     "language throughout the function schedule."),
  W=("S = 1: The Schedule of the LGOA lists waste management as mandatory functions "
     "using 'गर्ने, गराउने' (shall perform, shall cause to perform) language. These are "
     "legally binding obligations on all local governments.\n\n"
     "T = 1.0: All four sub-scores met — local governments explicitly designated, "
     "enforcement powers through local laws and national oversight, mandatory reporting/monitoring, "
     "and universal applicability to all local governments.\n\n"
     "PLASTIC RELEVANCE: Solid waste management functions (criterion c) cover plastic waste "
     "streams. Item झ.7 specifically includes 'reuse' (पुनः उपयोग) and 'processing' of "
     "health waste. Item ञ.15 covers all local-level waste management without restriction. "
     "The Act creates the institutional framework that makes the other plastic-specific "
     "policies (plastic bag bans, EFLG indicators etc.) operable at the local level."),
 ),
 dict(
  P=1.0, Q="Consumption",
  R=("SCHEDULE — MUNICIPAL FUNCTIONS (Section ञ: Local Market, Environment Conservation "
     "and Biodiversity), Item 14:\n\n"
     "'Control, monitoring and regulation of the sale and consumption of consumer goods "
     "having adverse effects on public health at the local level, as well as environmental "
     "pollution and hazardous substances.'\n\n"
     "[Nepali: स्थानीयस्तरमा जनस्वास्थ्यमा प्रतिकूल असर पर्ने वकिर्समका उपभोग्य वस्तुको "
     "बेचविखन र उपभोग तथा वातावरणीय प्रदूषण र हानिकारक पदार्थहरूको नियन्त्रण, अनुगमन "
     "तथा नियमन]\n\n"
     "This provision grants local governments the statutory authority to:\n"
     "• Control, monitor and regulate the SALE and USE of consumer goods harmful to public "
     "health — including plastic bags and other plastic products that pollute the environment\n"
     "• Control environmental pollution and hazardous substances at the local level\n\n"
     "This is the primary statutory basis that enables municipalities (like Khumbu Pasang "
     "Lhamu Rural Municipality for the Everest region) to enact local plastic bans. It is "
     "also used by municipalities implementing the Environment-friendly Local Governance "
     "Framework plastic bag bans."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Local governments (municipalities and rural "
     "municipalities) explicitly designated with this regulatory power in the Schedule. "
     "The LGOA Section 11 empowers local governments to enact local laws (स्थानीय कानून) "
     "to regulate any matter within their jurisdiction.\n\n"
     "Enforcement (+0.25): Local governments have power to enact local laws with penalty "
     "provisions (Section 11+). The LGOA empowers municipalities to fine, penalize and "
     "take regulatory action against those selling/using items harmful to public health, "
     "including plastic products. Enforcement through local law enforcement and inspection.\n\n"
     "Monitoring (+0.25): 'नियन्त्रण, अनुगमन तथा नियमन' explicitly includes monitoring "
     "(अनुगमन) as part of the mandatory function. Local governments must monitor compliance "
     "with regulations on harmful consumer goods.\n\n"
     "Unconditional (+0.25): The function is listed in the mandatory schedule applying to "
     "ALL municipalities and rural municipalities throughout Nepal. No geographic or "
     "institutional exemptions. The word 'नियमन' (regulation) implies ongoing mandatory duty."),
  W=("S = 1: Item ञ.14 is part of the mandatory Municipal Functions Schedule. The language "
     "'नियन्त्रण, अनुगमन तथा नियमन' = 'control, monitoring and regulation' is mandatory "
     "and operative — local governments have this duty immediately upon the Act's commencement.\n\n"
     "T = 1.0: All four sub-scores met — authority (local governments), enforcement (local "
     "law-making + penalty powers), monitoring (explicit 'अनुगमन'), unconditional (universal "
     "mandatory schedule).\n\n"
     "SIGNIFICANCE: This is the single most important statutory basis for local plastic "
     "bans in Nepal's policy landscape. When any municipality in Nepal bans plastic bags "
     "or plastic products, it exercises this power under ञ.14 of the LGOA Schedule. "
     "This provision transforms the LGOA into Nepal's primary enabling legislation for "
     "local-level plastic governance.\n\n"
     "Cross-reference: This provision was the legal basis for:\n"
     "• Khumbu Pasang Lhamu Rural Municipality's single-use plastic ban (2020)\n"
     "• Local implementations of the Environment-friendly Local Governance Framework (2013)\n"
     "• Sagarmatha and Ghodaghodi plastic bans\n"
     "The LGOA ञ.14 is the umbrella enabling statute for all these local instruments."),
 ),
 dict(
  P=0.60, Q="Waste management",
  R=("SECTION 62(2) — Service Charges:\n"
     "Local governments shall levy service charges on service users for services provided, "
     "including:\n"
     "'(ख) Service facilities such as waste management, sanitation, drainage, street "
     "lighting and similar services.'\n\n"
     "[Nepali: नगरपालिकाले निर्माण, सञ्चालन वा व्यवस्थापन गरेको...फोहरमैला व्यवस्थापन, "
     "सरसफाई, ढल निकास, सडक बत्ती जस्ता सेवा सुविधा...सेवा शुल्क लगाउन सक्नेछ]\n\n"
     "This creates a legally authorized self-financing mechanism for local waste management. "
     "The service fee system enables local governments to generate dedicated revenue for "
     "waste management operations including plastic waste collection, transport, and processing."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): Local governments (गाउँपालिका तथा नगरपालिका) are "
     "explicitly designated as the authority to levy and administer waste management "
     "service fees under Section 62.\n\n"
     "Enforcement (+0.25): Section 19 of the Solid Waste Management Act 2068 (which this "
     "Act operates alongside) authorizes suspension of waste management services for non-"
     "payment. The LGOA's broader enforcement powers apply to fee collection.\n\n"
     "Monitoring (+0.25): Annual financial auditing required for all local government "
     "revenues including service fees. Revenue collection must be reported to provincial "
     "and federal governments. The LGOA requires annual budget and financial reports.\n\n"
     "Unconditional (−0): Section 62(2) uses 'सक्नेछ' (may/can) — this is a permissive "
     "provision, not a mandatory obligation. Local governments are authorized but not "
     "required to levy fees for waste management. This permissive language means the fee "
     "is not unconditional."),
  W=("P = 0.60 (Economic — service fee authorization): The waste management service fee "
     "is an economic instrument enabling cost-recovery for local waste management services.\n\n"
     "S = 1: The fee-charging provision is operative once a municipality decides to levy "
     "fees (many municipalities have done so). The authorization is in force immediately "
     "upon the Act's commencement.\n\n"
     "T = 0.75: Authority, enforcement, and monitoring sub-scores credited. Unconditional "
     "sub-score not credited because Section 62(2) uses permissive 'सक्नेछ' (may) language "
     "rather than mandatory 'गर्नुपर्नेछ' (shall) — local governments are authorized but "
     "not required to levy waste management service fees.\n\n"
     "SIGNIFICANCE: In practice, most Nepali municipalities do levy waste management "
     "service fees under this provision, creating the financial foundation for local "
     "waste management including plastic waste collection."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Mandatory waste management functions for all local governments (Schedule झ.7 + ञ.15 + Ward functions)",
 "Instrument 2 — Mandatory local-level control of goods harmful to health and environment — ENABLES PLASTIC BANS (Schedule ञ.14)",
 "Instrument 3 — Service fee authorization for local waste management (Section 62)",
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
 ws.row_dimensions[r].height=270

for c,w in {1:32,2:40,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Translation + Key Provisions
ws2=wb.create_sheet("Translation & Key Provisions")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Title in Nepali","स्थानीय सरकार सञ्चालन ऐन, २०७४ [Local Government Operation Act, 2074 BS]"),
 ("Act number","Act No. 26 of 2074 BS"),
 ("Enactment","Enacted by Legislative Parliament under Article 296(1) of Nepal's Constitution"),
 ("Original year","2074 BS = 2017 CE"),
 ("Most recent amendment","Good Governance and Public Service Flow Act, 2081 BS — 2081/12/18 BS (≈ April 2025)"),
 ("All amendments",
  "1. Finance Act, 2075 BS (2075/03/32) | 2. Some Nepal Acts Amendment Act, 2075 BS (2075/11/19) | "
  "3. Federal, Provincial and Local Coordination Act, 2077 BS (2077/04/13) | "
  "4. Good Governance Act, 2081 BS (2081/12/18) [MOST RECENT]"),
 ("Source URL","https://lawcommission.gov.np (Nepal Law Commission)"),
 ("",""),
 ("KEY PLASTIC/WASTE PROVISIONS (verbatim Nepali + translation)",None),
 ("Location","Nepali (key terms) / Translation"),
 ("Section ञ.14 — KEY (enables plastic bans)",
  "Nepali: 'स्थानीयस्तरमा जनस्वास्थ्यमा प्रतिकूल असर पर्ने वकिर्समका उपभोग्य वस्तुको "
  "बेचविखन र उपभोग तथा वातावरणीय प्रदूषण र हानिकारक पदार्थहरूको नियन्त्रण, अनुगमन तथा नियमन'\n"
  "Translation: 'Control, monitoring and regulation of the sale and consumption of consumer "
  "goods having adverse effects on public health at the local level, as well as environmental "
  "pollution and hazardous substances.'"),
 ("Schedule झ.7 — Health waste",
  "Nepali: 'स्वास्थ्यजन्य फोहरमैला सङ्कलन, पुनः उपयोग, प्रशोधन, विसर्जन र सोको सेवा शुल्क "
  "निर्धारण र नियमन'\n"
  "Translation: 'Collection, reuse, processing, disposal of health waste and determination "
  "and regulation of service charges therefor.'"),
 ("Schedule ञ.15 — General waste management",
  "Nepali: 'स्थानीयस्तरमा सरसफाई तथा फोहरमैला व्यवस्थापन'\n"
  "Translation: 'Sanitation and waste management at the local level.'"),
 ("Section 62(2)(ख) — Service fees",
  "Nepali: 'फोहरमैला व्यवस्थापन, सरसफाई, ढल निकास, सडक बत्ती जस्ता सेवा सुविधा'\n"
  "Translation: 'Service facilities such as waste management, sanitation, drainage, street "
  "lighting and similar services.' [for which service charges may be levied]"),
 ("Municipality classification (p.9)",
  "Nepali: 'फोहरमैला प्रशोधन तथा व्यवस्थापन प्रणाली भएको'\n"
  "Translation: 'Having a solid waste processing and management system' [required for "
  "municipality classification]"),
 ("",""),
 ("SIGNIFICANCE OF SECTION ञ.14",None),
 ("Note",
  "Section ञ.14 is the primary statutory basis for local plastic bans in Nepal. "
  "When any municipality bans plastic bags, it exercises this power. Known examples:\n"
  "• Khumbu Pasang Lhamu Rural Municipality (Everest region): banned single-use plastics "
  "from 1 January 2020 under this provision\n"
  "• Ghodaghodi Municipality: plastic reduction measures\n"
  "• Various municipalities under the EFLG Framework 2013\n"
  "This makes the LGOA the foundational enabling statute for local plastic governance."),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","1.00  (Legislation — enacted by Legislative Parliament under Constitution Art.296(1))"),
 ("policy_target (E)","0  (no quantifiable plastic targets in the Act itself)"),
 ("policy_integration (I)","1.00  (10+ sectors: waste management, municipalities, water/sanitation, environment, agriculture, industry, healthcare, education, tourism, urban development)"),
 ("policy_circularity (K)","0.75  (consumption, disposal, recycling, environmental leakage — 4 lifecycle phases)"),
 ("policy_budget (M)","0.50  (Section 62 authorizes waste management service fees; not ring-fenced specifically for plastic)"),
 ("",""),
 ("Instr. 1: Mandatory waste management functions (Schedule झ.7 + ञ.15)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 2: Control of harmful goods + env. pollution (Schedule ञ.14) — ENABLES PLASTIC BANS","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 3: Service fee for waste management (Section 62)","P=0.60 | S=1 | T=0.75 | V=[auto 0.675]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Nepali (key terms) / Translation","Value","Note"):
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
 ws2.row_dimensions[ri].height=60
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree"),(1.00,"Legislation (parliament) ← THIS POLICY"),
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

path="/workspace/4p_index_lgoa_2074_2025.xlsx"
wb.save(path)
print(f"Saved: {path}")
