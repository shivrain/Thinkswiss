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
c.value=("Policy: Solid Waste Management National Policy, 2079 BS (2022 CE)  |  Country: Nepal  |  Year: 2022  |  "
         "Source: https://dpnet.org.np/resource-detail/1781 [original Nepali PDF — image-based]")
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
 name=("Solid Waste Management National Policy, 2079 BS (2022 CE)\n"
       "[फोहरमैला व्यवस्थापन राष्ट्रीय निति, २०७९]\n"
       "Approved by the Council of Ministers (Cabinet), October 2022"),
 url=("https://dpnet.org.np/resource-detail/1781 [original Nepali PDF via DPNet]\n"
      "NOTE: Source document is image-based Nepali PDF (10 pages, no OCR). "
      "Coding based on English translations from: The Himalayan Times (Oct 21, 2022), "
      "JICA Data Collection Survey on Waste Management in Nepal (Table 2-6 / NSWMP2079), "
      "and DPNet summary."),
 year=2022,
 obj=("National strategic policy framework for solid waste management in Nepal, approved "
      "by the Council of Ministers in October 2022. Sets 4 objectives and 8 strategies "
      "(Articles 9 and 10). In relation to plastics/waste: targets zero-waste-to-landfill "
      "by 2030; designates plastic-free zones; mandates complete ban on waste disposal in "
      "rivers, lakes, wetlands, wildlife habitat, religious sites, heritage sites, protected "
      "areas, roads, and other public places; applies polluter pays / waste generator pays "
      "principle with service charges; makes local governments responsible for household "
      "waste and industry/health institutions responsible for their hazardous/chemical/"
      "industrial/medical waste (EPR-like); promotes 3R (reduce, reuse, recycle) and waste "
      "segregation at source; requires waste management in school curricula; establishes "
      "three-tier governance with National Waste Management Coordination Committee (NWMCC)."),
 tgt=1,
 tgt_t=("'The Waste Management National Policy 2079 targets zero-waste-to-landfill by 2030 "
        "through enhanced recycling and waste-to-energy initiatives.' (DPNet summary) AND "
        "'designating plastic-free zones' (JICA Survey Report, NSWMP2079 strategy context)."),
 gtype=0.75,
 gtype_j=("Approved by the Council of Ministers (Cabinet) of the Government of Nepal in "
          "October 2022 (Kartik 2079 BS). Not enacted by Parliament — constitutes a Cabinet-"
          "approved national strategic policy document. The 4 objectives and 8 strategies "
          "(Articles 9 and 10) are defined in the policy, with working policies for each "
          "strategy. Scored 0.75 (executive decree/regulation level — Cabinet approved). "
          "Contains quantifiable target (zero-waste-to-landfill by 2030), so could score "
          "0.50 if viewed purely as a strategy/plan, but Cabinet-level approval elevates "
          "it to 0.75."),
 intg=1,
 sects="waste management, municipalities, industry, healthcare, chemicals, water/environment, education, private sector/PPP, agriculture/protected areas",
 circ=1.0,
 lc="production, consumption, recycling, disposal, environmental leakage",
 budg=0.5,
 budg_t=("Strategy 9.3: 'Service charges will be arranged according to the nature and "
         "quantity of waste and the Waste generators pays principle will be adopted.' "
         "(JICA Table 2-6, NSWMP2079 Strategies)\n"
         "Strategy 9.6: 'In addition to co-financing of all three tiers of the government, "
         "mobilisation of foreign aid in waste management and building partnership and "
         "collaboration with national and international NGOs.' (Himalayan Times)\n"
         "Budget/funding mechanisms exist (service charges + three-tier co-financing + "
         "foreign aid) but are not ring-fenced specifically for plastic waste management. "
         "M = 0.5."),
)

FILLS=["E2EFDA","D9F0FF","FFF2CC"]
instruments=[
 dict(
  P=1.0, Q="End of life",
  R=("Working Policies — Article 10 (per Himalayan Times and JICA Table 2-6):\n\n"
     "GENERATOR RESPONSIBILITY / POLLUTER PAYS:\n"
     "'The local levels will be made responsible for management of household waste, whereas "
     "the concerned organisation or institution will have to be accountable and responsible "
     "in managing hazardous, chemical, industrial and medical waste produced by them.'\n\n"
     "EPR-LIKE PROVISION:\n"
     "Strategy 9.1 directional policy: 'The local level will be made responsible for "
     "household waste management and the related producers or organizations for the management "
     "of hazardous, chemical, industrial and health institutional waste.'\n\n"
     "This creates an EPR-like mandatory responsibility framework for ALL waste generators. "
     "For plastic-containing industrial/chemical/health waste, the generating entity must "
     "manage it responsibly. The policy also commits to formulating 'an integrated act related "
     "to waste management' to operationalise these responsibilities legally.\n\n"
     "WASTE SEGREGATION AT SOURCE + 3R:\n"
     "'The policy has also encouraged segregation of waste at source, and reduction, "
     "recycling, and reuse of waste.' (Himalayan Times)\n\n"
     "SCHOOL CURRICULUM:\n"
     "'Subject related to waste management will be incorporated in school curricula, besides "
     "launching awareness campaign about civic duty in waste management.'"),
  S=0,
  T=0.50,
  U=("Responsible authority (+0.25): Ministry of Urban Development (MoUD) — responsibility "
     "transferred from MoFAGA to MoUD in November 2023 per JICA report. National Waste "
     "Management Coordination Committee (NWMCC) mandated to coordinate waste management "
     "across agencies (JICA Table 2-6, Strategy 9.5).\n\n"
     "Enforcement (−0): The policy says it 'will formulate and implement an integrated act "
     "related to waste management' — current enforcement relies on existing SWM Act 2068. "
     "No specific penalties stated within the policy itself.\n\n"
     "Monitoring (+0.25): Strategy 9.7: 'The development of the information system related "
     "to waste management and a strong statistical base will be prepared.' NWMCC provides "
     "coordination and monitoring function across all three government levels.\n\n"
     "Unconditional (−0): Working policy language uses 'will be made responsible' (future "
     "tense) — requires subsequent integrated act for full operationalisation. Currently "
     "aspirational rather than unconditionally binding."),
  W=("S = 0: The generator responsibility provisions use future-tense language ('will be "
     "made responsible'; 'will formulate and implement an integrated act') — indicating "
     "these are planning commitments rather than currently operative obligations. The "
     "existing SWM Act 2068 already has some of these provisions but this policy expands "
     "the framework pending new legislation.\n\n"
     "P = 1.0 (Regulatory — EPR-like mandatory responsibility standard): The principle "
     "that waste generators are responsible for their waste constitutes a mandatory performance "
     "standard when operationalised. For industrial/chemical/plastic packaging waste, this "
     "directly creates generator liability for plastic waste management.\n\n"
     "The 3R + segregation provisions use 'encouraged' language → more information/voluntary "
     "(P=0.40). However, they are combined with the mandatory generator responsibility "
     "provisions, so highest applicable type = 1.0 (Regulatory) is applied per coding rules."),
 ),
 dict(
  P=1.0, Q="Environmental leakage",
  R=("Working Policies — Article 10 (per Himalayan Times, JICA context):\n\n"
     "COMPLETE WASTE DISPOSAL BAN:\n"
     "'The government will completely impose ban on disposal of waste in river systems, "
     "lakes, wetlands, wildlife habitat, religious sites, heritage sites, protected areas, "
     "roads and other public places.'\n\n"
     "PLASTIC-FREE ZONES:\n"
     "From JICA Survey Report (NSWMP2079 context): the policy includes provision for "
     "'requiring proper treatment of hazardous wastes by generators, and designating "
     "plastic-free zones.' (JICA p.10, Section on NSWMP2079 BKM municipality)\n\n"
     "These provisions directly address the environmental leakage of solid waste — "
     "including plastic waste — into water bodies, natural habitats, and public spaces. "
     "The combined ban on open disposal and designation of plastic-free zones creates "
     "a regulatory framework preventing plastic from entering the environment."),
  S=0,
  T=0.50,
  U=("Responsible authority (+0.25): Government of Nepal (central level) responsible for "
     "imposing the ban; local governments responsible for enforcement in their jurisdiction. "
     "MoUD as coordinating ministry per November 2023 transfer.\n\n"
     "Enforcement (−0): Ban is stated as future action ('will completely impose ban') — "
     "no specific enforcement mechanism or penalties stated within the policy text. "
     "The integrated waste management act to be formulated will presumably include "
     "enforcement provisions.\n\n"
     "Monitoring (+0.25): Regulatory bodies provisioned for waste stream management; "
     "NWMCC coordination. Strategy 9.7 information system will enable monitoring of "
     "whether waste reaches water bodies/public places.\n\n"
     "Unconditional (−0): Future tense ('will impose') — not currently operative. "
     "No timeline for ban's commencement stated in the policy text."),
  W=("S = 0: The waste disposal ban uses future tense — 'will completely impose ban' — "
     "indicating this is a policy commitment, not a currently operative legal prohibition. "
     "Compare with the plastic flowers ban (gazette notice, July 2022) which used mandatory "
     "language and was immediately operative — that scores S=1. This policy's ban provisions "
     "require subsequent regulation to become operative.\n\n"
     "PLASTIC-FREE ZONES: This provision directly targets plastic waste at specific "
     "locations — passes criterion (a) (explicitly enables plastic-relevant actions: "
     "zone-based plastic restrictions).\n\n"
     "SIGNIFICANCE: When operationalised, this ban on waste disposal in rivers/lakes/"
     "wetlands/public places will be one of the most comprehensive environmental leakage "
     "prevention instruments in Nepal's plastic governance landscape."),
 ),
 dict(
  P=0.60, Q="Waste management",
  R=("Strategy 9.3 — Minimizing waste at source and making landfill sites sustainable "
     "(JICA Table 2-6 directional excerpts):\n\n"
     "'Service charges will be arranged according to the nature and quantity of waste "
     "and the Waste generators pays principle will be adopted.'\n\n"
     "Strategy 9.5 — Clarifying roles of federal, provincial and local levels:\n"
     "'The role of the provincial government will be established for coordination between "
     "local levels in the construction of infrastructure related to waste management. "
     "The local level will be made responsible for the construction and operation of "
     "the infrastructure and structures required for the collection, disposal and processing "
     "of waste.'\n\n"
     "Strategy 9.6 — Prioritising co-financing and partnership:\n"
     "'In the construction of large infrastructures, cooperation between federations, "
     "provinces and local levels and partnership and cooperation with national and "
     "international NGOs will be encouraged. The private sector will be mobilised "
     "according to the concept of PPP in waste management.'\n\n"
     "Government co-financing + foreign aid + PPP + service charges = comprehensive "
     "financial architecture for waste management."),
  S=1,
  T=0.75,
  U=("Responsible authority (+0.25): NWMCC (National Waste Management Coordination "
     "Committee) mandated by the policy; MoUD as coordinator; all three government levels "
     "designated with specific roles in Strategy 9.5.\n\n"
     "Enforcement (+0.25): The 'Waste generators pays' principle creates a binding "
     "economic obligation — generators must pay service charges. This is backed by the "
     "existing SWM Act 2068's service fee provisions (Section 18) which this policy "
     "reinforces. The principle is stated as a binding commitment ('will be adopted').\n\n"
     "Monitoring (+0.25): Strategy 9.7 explicitly mandates development of information "
     "system and strong statistical database for waste management — monitoring mechanism.\n\n"
     "Unconditional (−0): Co-financing and PPP elements use 'will be encouraged' language — "
     "some conditional/aspirational elements exist alongside the binding service charge "
     "principle."),
  W=("P = 0.60 (Economic — service charge/polluter pays + co-financing mechanism): "
     "The 'Waste generators pays' service charge constitutes an economic instrument. "
     "The three-tier co-financing and PPP elements supplement this.\n\n"
     "S = 1: The 'Waste generators pays principle will be adopted' uses binding language "
     "within a Cabinet-approved policy, making this an operative commitment. It reinforces "
     "the existing SWM Act 2068 service fee provisions.\n\n"
     "T = 0.75: Three of four sub-scores met — NWMCC authority designated, enforcement "
     "through service charge mechanism, monitoring through Strategy 9.7 information system. "
     "Unconditional not fully credited due to some aspirational language in co-financing/"
     "PPP elements."),
 ),
]

for i,(instr,label) in enumerate(zip(instruments,[
 "Instrument 1 — Generator responsibility/EPR-like + waste segregation/3R + school curriculum (Articles 9.1-9.4)",
 "Instrument 2 — Complete ban on waste disposal in rivers/lakes/wetlands/public places + designating plastic-free zones (Article 10)",
 "Instrument 3 — Waste generator pays principle + three-tier co-financing + NWMCC governance (Strategies 9.3, 9.5, 9.6)",
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

for c,w in {1:32,2:42,3:8,4:35,5:8,6:38,7:8,8:35,9:9,10:35,11:8,12:28,13:8,14:40,15:9,
            16:10,17:18,18:52,19:10,20:10,21:52,22:9,23:52,24:28}.items():
 ws.column_dimensions[get_column_letter(c)].width=w
ws.freeze_panes="D5"

# Sheet 2 — Key Provisions
ws2=wb.create_sheet("Key Provisions (English Translation)")
rows=[
 ("DOCUMENT OVERVIEW",None),
 ("Field","Details"),
 ("Full title","Solid Waste Management National Policy, 2079 BS [फोहरमैला व्यवस्थापन राष्ट्रीय निति, २०७९]"),
 ("Year","2079 BS = 2022 CE"),
 ("Approval","Council of Ministers (Cabinet), approximately October 14, 2022"),
 ("Announced","October 20-21, 2022 (Himalayan Times, myRepublica)"),
 ("Document format","Image-based Nepali PDF (10 pages, no OCR). Coding from English translations."),
 ("Download","https://dpnet.org.np/resource-detail/1781"),
 ("Translation sources",
  "1. The Himalayan Times, Oct 21, 2022 (English reporting)\n"
  "2. JICA Data Collection Survey on Waste Management in Nepal (Table 2-6 NSWMP2079 strategies)\n"
  "3. DPNet summary"),
 ("",""),
 ("4 OBJECTIVES",None),
 ("Objective","Description"),
 ("1","Providing guidance to law and standards related to management of waste from houses, industrial and tertiary sectors"),
 ("2","Mitigating the adverse impacts of waste on the environment and public health"),
 ("3","Clarifying the roles of federal units in waste management"),
 ("4","Contributing to the national economy through mobilisation of waste as resource using innovative technologies"),
 ("",""),
 ("8 STRATEGIES (Article 9) + KEY WORKING POLICIES (Article 10)",None),
 ("Strategy","Directional Policy Excerpt"),
 ("9.1 Legal basis for waste classification",
  "'The local level will be made responsible for household waste management and the related "
  "producers or organizations for the management of hazardous, chemical, industrial and "
  "health institutional waste.' — EPR/polluter pays framework"),
 ("9.2 Separate standards by waste classification",
  "'Adherence and coordination of standards will be made effective by arranging appropriate "
  "institutional structures at the federal, provincial and local levels.'"),
 ("9.3 Minimize waste at source + sustainable landfill",
  "'Service charges will be arranged according to the nature and quantity of waste and "
  "the Waste generators pays principle will be adopted.' + 'Recycling, processing and "
  "composting of household waste will be encouraged.'"),
 ("9.4 Citizen responsibility",
  "'Public awareness raising will be broadened and institutionalised by including topics "
  "related to waste management in the curriculum of school education.'"),
 ("9.5 Clarify federal/provincial/local roles",
  "'The local level will be made responsible for the construction and operation of the "
  "infrastructure and structures required for the collection, disposal and processing of waste.'"),
 ("9.6 Co-financing + PPP",
  "'The private sector will be mobilised according to the concept of PPP in waste management.' "
  "+ Co-financing by all three tiers + foreign aid"),
 ("9.7 Research and database","Information system for waste management and strong statistical base"),
 ("9.8 Capacity building","Training, study and observation visits"),
 ("",""),
 ("KEY WORKING POLICIES (Article 10)",None),
 ("Policy","Text"),
 ("Integrated Act commitment",
  "'The government will...formulate and implement an integrated act related to waste management.'"),
 ("COMPLETE WASTE DISPOSAL BAN (KEY)",
  "'The government will completely impose ban on disposal of waste in river systems, lakes, "
  "wetlands, wildlife habitat, religious sites, heritage sites, protected areas, roads and "
  "other public places.'"),
 ("Plastic-free zones",
  "'designating plastic-free zones' (JICA Survey Report, NSWMP2079 strategy context)"),
 ("Generator responsibility",
  "'The local levels will be made responsible for management of household waste whereas the "
  "concerned organisation or institution will have to be accountable and responsible in "
  "managing hazardous, chemical, industrial and medical waste produced by them.'"),
 ("Regulatory bodies",
  "'Provision of regulatory bodies will be made for management of hazardous, chemical, "
  "industrial, medical and household waste.'"),
 ("",""),
 ("SCORE SUMMARY",None),
 ("Field","Value"),
 ("policy_type (G)","0.75  (Cabinet/Council of Ministers approved — not Parliament-enacted)"),
 ("policy_target (E)","1  — 'zero-waste-to-landfill by 2030' + 'designating plastic-free zones'"),
 ("policy_integration (I)","1.00  (9 sectors: waste management, municipalities, industry, healthcare, chemicals, water/environment, education, private sector/PPP, agriculture/protected areas)"),
 ("policy_circularity (K)","1.00  (ALL 5 phases: production, consumption, recycling, disposal, environmental leakage)"),
 ("policy_budget (M)","0.50  ('Waste generators pays' service charges + three-tier co-financing + foreign aid; not ring-fenced)"),
 ("",""),
 ("Instr. 1: Generator responsibility/EPR + 3R + school curriculum","P=1.00 | S=0 | T=0.50 | V=[auto 0.750]"),
 ("Instr. 2: Complete waste disposal ban + plastic-free zones","P=1.00 | S=0 | T=0.50 | V=[auto 0.750]"),
 ("Instr. 3: Waste generator pays + three-tier co-financing + NWMCC","P=0.60 | S=1 | T=0.75 | V=[auto 0.675]"),
]
for ri,row in enumerate(rows,1):
 if len(row)>=2 and row[1] is None and row[0]:
  ws2.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=2)
  c=ws2.cell(row=ri,column=1,value=row[0])
  c.fill=hfill(MID_BLUE); c.font=hfont(bold=True,size=10,color=WHITE); c.alignment=CA
 elif row[0]=="": pass
 elif row[1] in("Details","Description","Directional Policy Excerpt","Text","Value","Objective","Strategy","Policy"):
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
ws2.column_dimensions["A"].width=35
ws2.column_dimensions["B"].width=85

ws3=wb.create_sheet("Score Reference")
ref=[("POLICY TYPE (Col G)",None),("Score","Description"),
     (0.25,"Strategy/plan — aspirational"),(0.50,"Strategy/plan with quantifiable targets"),
     (0.75,"Regulation or executive decree ← THIS POLICY"),(1.00,"Legislation (parliament)"),
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

path="/workspace/4p_index_nswmp_2079_2022.xlsx"
wb.save(path)
print(f"Saved: {path}")
