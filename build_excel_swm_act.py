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
c.value=("Policy: Solid Waste Management Act, 2068 BS (2011), as amended in 2075 BS (2019) | Nepal | Year: 2019 | "
         "Source: https://faolex.fao.org/docs/ [URL truncated] / https://lawcommission.gov.np")
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
 name=("Solid Waste Management Act, 2068 BS (2011)\n"
       "[फोहरमैला व्यवस्थापन ऐन, २०६८]\n"
       "As amended by 'Some Nepal Acts Amendment Act, 2075 BS' (2018/19 CE)\n"
       "Act No. 5 of 2068 BS — most recent version"),
 url=("https://faolex.fao.org/docs/ [URL truncated in source image] / "
      "https://www.lawcommission.gov.np [full Nepali text]\n"
      "NOTE: User label says 'National Solid Waste Management Policy 2022' but the "
      "uploaded document (Solid_Waste_Management_Act_7618.pdf) is the SWM Act 2068 "
      "(legislation from lawcommission.gov.np), not the National SWM Policy 2079 (2022). "
      "The SWM Act 2068 is the primary legislation; the National SWM Policy 2022 is a "
      "separate policy document."),
 year=2019,
 obj=("Primary legislation on solid waste management in Nepal, enacted by the Constituent "
      "Assembly acting as Parliament. In relation to plastics: establishes mandatory duty "
      "for all persons/entities to reduce waste at source (Section 5); mandates waste "
      "segregation at source into organic/inorganic categories (Section 6); requires "
      "industries to reuse packaging materials (Section 10); places responsibility for "
      "hazardous/chemical/industrial/medical waste management on the generator under the "
      "polluter-pays principle (Section 4); provides ring-fenced service fee mechanism for "
      "waste management (Section 18); prescribes fines of NPR 50,000–100,000 for haphazard "
      "disposal of hazardous/chemical/industrial/medical waste (Section 39); and grants "
      "the Government an enabling power to ban production/sale of items generating "
      "excessive waste by publishing notification in the Nepal Gazette (Section 43/44 area). "
      "The Act covers plastic waste implicitly as part of 'non-decomposable waste' and "
      "'industrial waste' categories."),
 tgt=0,
 tgt_t="",
 gtype=1.0,
 gtype_j=("Legislation enacted by the Constituent Assembly of Nepal acting as Parliament "
          "under Article 83 of the Interim Constitution 2063 (2007). Act No. 5 of 2068 BS "
          "(2011 CE). Authentication date: 2068/03/01 BS (approximately July 15, 2011 CE). "
          "Came into force immediately. Most recently amended by 'Some Nepal Acts Amendment "
          "Act, 2075 BS' on 2075/11/19 BS (approximately February/March 2019 CE) — this is "
          "the most recent version of the Act.\n\n"
          "G = 1.0 (Legislation): The Constituent Assembly enacted this Act in its capacity "
          "as Parliament — the highest legislative authority in Nepal. It is primary "
          "legislation, not an executive decree or sub-legislative instrument."),
 intg=0.75,
 sects="waste management, municipalities, industry, healthcare, chemicals, private sector",
 circ=1.0,
 lc="production, consumption, recycling, disposal, environmental leakage",
 budg=1.0,
 budg_t=("Section 18(5): 'The revenue collected from service fees as per this Section shall "
         "be kept by the local body in a separate heading and shall be expended on waste "
         "management, environmental protection and development of the landfill site affected "
         "area.' [Translation from Nepali: यस दफा बमोजिम उठाइएको सेवा शुल्कबाट प्राप्त "
         "आम्दानी स्थानीय तहले एउटा छुट्टै शीर्षकमा राखी सो रकम फोहरमैलाको व्यवस्थापन, "
         "वातावरणीय संरक्षण तथा फोहरमैला व्यवस्थापन स्थल प्रभावित क्षेत्रको विकासमा "
         "खर्च गर्नुपर्नेछ।] — this creates a legally ring-fenced fund for waste management "
         "purposes, self-generated through mandatory service fees."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC"]
instruments=[
 dict(
  P=1.0, Q="Consumption",
  R=("SECTION 5 — Mandatory waste reduction duty (फोहरमैलाको उत्पादन कम गर्ने):\n"
     "(1) Every person, institution or entity shall minimize the waste generated during "
     "their activities (as much as possible).\n"
     "(2) It is the duty of every person, institution or entity to reduce the quantum of "
     "solid waste by making arrangements to dispose of the disposable solid waste within "
     "their own area or making arrangement for the reuse and discharging the remaining "
     "solid waste thereafter.\n\n"
     "SECTION 6 — Mandatory waste segregation (फोहरमैलाको पृथकीकरण):\n"
     "(1) The local body shall prescribe for segregation of solid waste at source into at "
     "least organic and inorganic categories.\n"
     "(2) The responsibility to segregate solid waste at source as prescribed by the local "
     "body and carrying them into the collection center shall rest with the person, "
     "institution or entity who produces the solid waste.\n\n"
     "SECTION 10 — Waste minimization, reuse and recycling (फोहरमैलाको न्यूनीकरण, "
     "पुनः प्रयोग तथा पुनः चक्रीय प्रयोग):\n"
     "(1) The local body shall take necessary action to encourage waste reduction, reuse "
     "and recycling.\n"
     "(2) In coordination with relevant industries, the local body may encourage "
     "industries to reuse the packaging materials used to pack their products, thereby "
     "reducing the quantity of waste."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Local bodies (municipalities, rural municipalities) "
     "explicitly designated as the authority to prescribe segregation and promote "
     "recycling. Section 4(1): Responsibility for waste management rests with local bodies. "
     "Section 3(1): Responsibility to build and operate waste infrastructure rests with local bodies.\n\n"
     "Enforcement (+0.25): Section 39 penalties apply: S.39(1) — haphazard dumping in "
     "unauthorized places incurs fines of NPR 5,000 (first offense), 5,000–10,000 (second), "
     "15,000/occurrence (subsequent). S.39(8) — haphazard disposal of hazardous/chemical/"
     "industrial/medical waste: NPR 50,000–100,000; double for repeat offense and license "
     "cancellation recommendation.\n\n"
     "Monitoring (+0.25): Section 21 establishes monitoring framework; Section 28-30 "
     "establish the Solid Waste Management Technical Assistance Centre and Board with "
     "monitoring functions; local body monitoring is implicit in its designated "
     "implementation responsibility.\n\n"
     "Unconditional (+0.25): Sections 5 and 6 apply to ALL persons, institutions and "
     "entities throughout Nepal with no stated exemptions. 'प्रत्येक व्यक्ति, संस्था वा "
     "निकायको कर्तव्य' = 'duty of every person, institution or entity.'"),
  W=("S = 1 (IN FORCE): Primary legislation with mandatory language — 'shall be the "
     "duty' (कर्तव्य हुनेछ), 'shall prescribe' (तोक्नु पर्नेछ). The Act came into force "
     "immediately upon authentication (2068/03/01 BS = July 2011) and was last amended "
     "in 2075 BS (2018/19 CE).\n\n"
     "T = 1.0: All four sub-scores met — local bodies as designated authority, Section 39 "
     "penalties as explicit enforcement, monitoring framework established, duty applies "
     "to all persons without exemptions.\n\n"
     "PLASTIC RELEVANCE: Sections 5, 6, and 10 apply to all solid waste including plastic "
     "waste. Section 6(1) requires segregation into at least 'organic and inorganic' — "
     "plastic is non-organic (inorganic/non-decomposable) waste. Section 10(2) specifically "
     "references 'packaging materials' used by industries for reuse, which directly targets "
     "plastic packaging waste streams."),
 ),
 dict(
  P=1.0, Q="End of life",
  R=("SECTION 4 — Polluter pays / generator responsibility (फोहरमैला व्यवस्थापन गर्ने "
     "दायित्व):\n"
     "(2) Notwithstanding anything in sub-section (1), the responsibility for processing "
     "and management of hazardous waste, medical waste, chemical waste or industrial waste "
     "under the prescribed standards shall rest with the person or institution that has "
     "generated the solid waste. [This creates a mandatory EPR-like obligation on industrial "
     "and commercial waste generators.]\n\n"
     "SECTION 39(8) — Penalties for haphazard disposal of hazardous/chemical/industrial/medical waste:\n"
     "Fine of NPR 50,000 to 100,000 for:\n"
     "• Haphazard disposal/keeping of chemical waste, industrial waste, medical waste or "
     "hazardous waste;\n"
     "• Haphazard disposal of hazardous waste from any industrial enterprise or health "
     "institution;\n"
     "Double fine for repeat offense + recommendation to cancel license.\n\n"
     "SECTION 43 — Licensing conditions for health institutions:\n"
     "License-granting authority shall confirm whether appropriate waste management "
     "arrangements exist before granting/renewing license to health institutions.\n\n"
     "GOVERNMENT ENABLING POWER (Section 39(8)/44 area):\n"
     "Government of Nepal may, by publishing notification in the Nepal Gazette, prohibit "
     "or restrict the production and sale of items that generate excessive waste. [This is "
     "the statutory basis for potential plastic product bans under this Act.]"),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Hazardous/industrial/chemical/medical waste "
     "generators explicitly designated as responsible parties (Section 4(2)). Local "
     "bodies, licensing authorities, and Department of Environment designated as "
     "enforcement bodies.\n\n"
     "Enforcement (+0.25): Section 39(8): NPR 50,000–100,000 fines explicitly stated for "
     "haphazard disposal of chemical/industrial/medical/hazardous waste. Double fine for "
     "repeat offense. Recommendation for license cancellation. These explicit penalty "
     "amounts are in the Act itself.\n\n"
     "Monitoring (+0.25): Section 43 requires licensing authorities to verify waste "
     "management compliance before issuing/renewing health institution licenses. "
     "Section 39 penalties create a monitoring incentive. Local body monitoring "
     "responsibility established throughout the Act.\n\n"
     "Unconditional (+0.25): Section 4(2) — generator responsibility applies to ALL "
     "industrial enterprises and health institutions producing hazardous/chemical/medical "
     "waste with no stated exemptions. Penalty provisions apply to all violators."),
  W=("S = 1: Section 4(2) uses mandatory language — 'shall rest with the person or "
     "institution that has generated the waste' (त्यस्तो फोहरमैला उत्पादन गर्ने व्यक्ति वा "
     "निकायको हुनेछ). This is currently operative.\n\n"
     "T = 1.0: All four sub-scores met — generator responsibility explicitly designated, "
     "explicit fines of 50,000–100,000 in Section 39(8), licensing monitoring in Section 43, "
     "unconditional applicability to all generators.\n\n"
     "PLASTIC RELEVANCE: Industrial packaging waste (plastic packaging from manufacturing) "
     "falls under 'industrial waste' (औद्योगिक फोहरमैला) defined in Section 2(क) as "
     "'hazardous and polluting waste discharged from industrial establishments.' This "
     "makes Section 4(2) directly applicable to plastic manufacturing and packaging "
     "waste under criterion (c).\n\n"
     "The government enabling power to ban high-waste items (Section 39/44 area) was "
     "NOT directly invoked through this Act for plastic bans — those bans were issued "
     "under EPA 2076 Section 15(6). However, this Act provides the broader statutory "
     "framework supporting waste management regulations."),
 ),
 dict(
  P=0.60, Q="Waste management",
  R=("SECTION 18 — Mandatory service fee collection with ring-fenced fund:\n"
     "(1) The local body may levy and collect service charges from the concerned person, "
     "institution or entity for waste management services.\n"
     "(2) Service charge rates shall be determined by the local body based on waste "
     "quantity, weight, and nature.\n"
     "(3) Revenue from service fees collected through waste management contractors may "
     "be retained by those contractors with local body agreement.\n"
     "(4) Concessions may be provided to underprivileged groups.\n"
     "(5) Service fee revenues shall be kept in a SEPARATE HEADING by the local body "
     "and expended on waste management, environmental protection, and development of "
     "the landfill site affected area.\n\n"
     "This creates a self-financing mechanism for Nepal's local solid waste management "
     "system, providing ring-fenced funding for waste (including plastic) management "
     "infrastructure and services."),
  S=1,
  T=1.00,
  U=("Responsible authority (+0.25): Local bodies (municipalities, rural municipalities) "
     "are explicitly authorized and responsible for levying service fees and managing "
     "the ring-fenced fund. Section 18(1) designates local bodies as the fee-levying authority.\n\n"
     "Enforcement (+0.25): Section 19 authorizes suspension or termination of waste "
     "management services for non-payment of fees. Service fee collection is backed by "
     "the legal authority of the Act. Non-payment triggers service suspension — an "
     "effective enforcement mechanism.\n\n"
     "Monitoring (+0.25): Section 18(5) requires keeping fees in a SEPARATE account "
     "heading — this implies auditable ring-fencing. The Act's general monitoring "
     "framework (Section 21, Solid Waste Management Technical Assistance Centre) "
     "monitors implementation.\n\n"
     "Unconditional (+0.25): Section 18(2) — fees shall be determined on the basis of "
     "waste quantity, weight and nature, applicable to all waste generators with "
     "only a concession (not exemption) for underprivileged groups. The fee system "
     "applies universally."),
  W=("P = 0.60 (Economic — mandatory service fee with ring-fenced fund): Section 18 "
     "establishes the economic financing mechanism for Nepal's local waste management "
     "system. The ring-fenced separate heading (Section 18(5)) for waste management, "
     "environmental protection, and landfill area development constitutes a dedicated "
     "fund requirement — hence M = 1.0 at the policy level.\n\n"
     "S = 1: The service fee provision is operative legislation — local bodies 'may levy' "
     "(सेवा शुल्क लगाई उठाउन सक्नेछ) — this is an enabling/permissive provision, not "
     "a mandatory one. Language test: 'may' = S = 0? But the ring-fenced spending in "
     "Section 18(5) uses mandatory language 'shall be expended' (खर्च गर्नुपर्नेछ). "
     "Coded as S = 1 because the ring-fenced fund mechanism is operative and binding "
     "once fees are collected, even though the fee levy itself is permissive.\n\n"
     "T = 1.0: All sub-scores met — local body designated, service suspension enforcement, "
     "separate account monitoring, universal applicability with minor concession only.\n\n"
     "This instrument is relevant to plastic waste management as it provides the financial "
     "mechanism for all solid waste management activities, including plastic collection, "
     "transport, and disposal."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Mandatory waste reduction (S.5) + segregation at source (S.6) + recycling promotion (S.10)",
 "Instrument 2 — Generator responsibility for hazardous/chemical/industrial/medical waste (S.4) + penalties (S.39) + licensing conditions (S.43)",
 "Instrument 3 — Ring-fenced service fee mechanism for local waste management (S.18)",
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

for c,w in {1:32,2:40,3:8,4:35,5:8,6:20,7:8,8:35,9:9,10:30,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Translation + Key Provisions
ws2=wb.create_sheet("Translation & Key Provisions")
rows=[
 ("DOCUMENT IDENTIFICATION — IMPORTANT NOTE",None),
 ("User label","National Solid Waste Management Policy 2022"),
 ("Actual document","Solid Waste Management Act, 2068 BS (2011 CE), as amended in 2075 BS (2018/19 CE)"),
 ("Title in Nepali","फोहरमैला व्यवस्थापन ऐन, २०६८ [Solid Waste Management Act, 2068]"),
 ("Document source","https://www.lawcommission.gov.np (Nepal Law Commission)"),
 ("Act number","Act No. 5 of 2068 BS"),
 ("Authentication date","2068/03/01 BS = approximately 15 July 2011"),
 ("Amendment 1","Some Nepal Acts Amendment Act, 2072 BS — 2072/11/13 BS (≈ February 2016)"),
 ("Amendment 2","Some Nepal Acts Amendment Act, 2075 BS — 2075/11/19 BS (≈ March 2019) [MOST RECENT]"),
 ("",""),
 ("NOTE on document mismatch",
  "The uploaded file (Solid_Waste_Management_Act_7618.pdf) contains the SWM ACT 2068 "
  "(primary legislation), not the National SWM POLICY 2022 (2079 BS). These are two "
  "different documents. The National SWM Policy 2022 (2079 BS) is a separate policy "
  "document providing strategic framework. The SWM Act 2068 is the primary legislation "
  "that actually IMPLEMENTS waste management — higher relevance for 4P Index (G=1.0)."),
 ("",""),
 ("KEY TRANSLATED PROVISIONS",None),
 ("Section","Translation"),
 ("Section 2 — Definitions",
  "Defines: Industrial waste (औद्योगिक फोहरमैला), chemical waste (रासायनिक फोहरमैला), "
  "hazardous waste (हानिकारक फोहरमैला), medical waste (स्वास्थ्य संस्थाजन्य फोहरमैला), "
  "processing (प्रशोधन), recycling (पुनः चक्रीय प्रयोग), reduction (न्यूनीकरण), "
  "sanitary landfill site (फोहरमैला व्यवस्थापन स्थल / Sanitary Landfill Site)"),
 ("Section 3","Responsibility of local bodies to build and operate waste infrastructure (transfer stations, landfill sites, processing plants, compost plants, biogas plants)"),
 ("Section 4(2)",
  "'The responsibility for processing and management of hazardous waste, medical waste, "
  "chemical waste or industrial waste under the prescribed standards shall rest with the "
  "person or institution that has generated the solid waste.' [POLLUTER PAYS / EPR-like]"),
 ("Section 5",
  "'Every person, institution or entity shall minimize waste generated during their "
  "activities.' 'It is the duty of every person, institution or entity to reduce waste "
  "quantity by making arrangements for disposal/reuse within their own area.'"),
 ("Section 6",
  "'The local body shall prescribe for segregation of solid waste at source into at least "
  "organic and inorganic categories.' 'The responsibility to segregate at source rests "
  "with the waste producer.'"),
 ("Section 10",
  "'Local body shall take necessary action to encourage waste reduction, reuse and recycling. "
  "In coordination with relevant industries, local body may encourage reuse of packaging "
  "materials, thereby reducing waste quantity.'"),
 ("Section 18(5) — KEY BUDGET",
  "'Revenue from service fees shall be kept in a SEPARATE HEADING by the local body and "
  "expended on waste management, environmental protection and development of the landfill "
  "site affected area.' [Ring-fenced fund requirement]"),
 ("Section 39(8) — KEY PENALTY",
  "Fine of NPR 50,000–100,000 for: haphazard disposal of chemical/industrial/medical/"
  "hazardous waste; double fine for repeat offense; recommendation for license cancellation."),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","1.00  (Legislation — enacted by Constituent Assembly acting as Parliament)"),
 ("policy_target (E)","0  (no quantifiable plastic-specific targets in the Act)"),
 ("policy_integration (I)","0.75  (6 sectors: waste management, municipalities, industry, healthcare, chemicals, private sector)"),
 ("policy_circularity (K)","1.00  (ALL 5 phases: production, consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","1.00  (Section 18(5): ring-fenced service fee fund for waste management)"),
 ("",""),
 ("Instr. 1: Waste reduction + segregation + recycling (S.5, 6, 10)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 2: Generator responsibility + penalties + licensing conditions (S.4, 39, 43)","P=1.00 | S=1 | T=1.00 | V=[auto 1.000]"),
 ("Instr. 3: Ring-fenced service fee mechanism (S.18)","P=0.60 | S=1 | T=1.00 | V=[auto 0.800]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Translation","Value","Actual document","User label"):
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
ws2.column_dimensions["A"].width=30
ws2.column_dimensions["B"].width=90

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

path="/workspace/4p_index_swm_act_2068_2019.xlsx"
wb.save(path)
print(f"Saved: {path}")
