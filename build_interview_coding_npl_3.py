"""Generate interview coding spreadsheet for Nepal interview NPL_3.

Coding follows the Overview All Interviews codebook and is based on the
Kathmandu Metropolitan City municipality officer interview (23 June 2025) —
transcript, interview notes, and guideline.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_3.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "KTM Municipal Office — 23 June 2025 (English)"),
  ("Interview notes", "Municipality Officer guideline — bullet notes"),
  ("Interview guideline", "Plastic Pollution Governance in Nepal — Municipality Officer"),
  ("Codebook", "Overview All Interviews — NEW Codebook"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_3",
  "Country": "NEPAL",
  "Interview date": "23 June 2025",
  "Affiliation / role": "Kathmandu Metropolitan City — municipal officer (solid/dry waste management)",
  "Work focus": (
    "Dry waste management since 2017; sorting by grade; processing and passing to "
    "recyclers; UNDP consulting; multilayer plastics research; policy framework "
    "research with UK university"
  ),
  "Coded actor type": "governmental",
  "Actor classification rationale": (
    "Local-government official implementing Kathmandu Metropolitan City's solid waste "
    "and plastic management responsibilities under the Solid Waste Management Act 2068."
  ),
  "Role in plastic governance": (
    "Municipal implementation: seven ward clusters, private-sector MoUs, pollution "
    "standards, enforcement of local bans, transfer to Bancharedanda landfill"
  ),
  "Perspective": (
    "Implementation insider — strong on ground-level constraints (land, enforcement, "
    "informal sector) and municipal–federal–provincial coordination gaps"
  ),
  "Key partners mentioned": (
    "DoCoRecyclers (Cluster 7), UNDP (mechanized plastic recovery MoU), "
    "Dakshinkali Municipality (debris only), 18-municipality mayor forum"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_3",
  "Actortype": "governmental",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "NA",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "yes",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "yes",
  "Problem_import": "yes",
  "Impacts": "water, cities, biodiversity, health",
  "Governance": "yes",
  "Relevance_international_pol": "no",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "yes",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "yes",
  "Local government": "yes",
  "Students": "yes",
  "Private_sector": "yes",
  "Civil_society": "yes",
  "Science": "yes",
  "Households": "yes",
  "Education institutions": "yes",
  "Private_companies(hotels, shops, etc.)": "yes",
  "Implementation issues": "yes",
  "Monitoring": "yes",
  "Financial_resources": "yes",
  "Research": "yes",
  "Infrastructure": "yes",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "no",
  "Pol_import": "yes",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
  "Pol_capacity": "no",
  "Pol_RD": "yes",
  "Pol_tax": "no",
  "Pol_ban": "yes",
  "Pol_subsitutes": "yes",
  "Pol_clean_up": "yes",
  "Pol_upcycling": "yes",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "yes",
  "Sol_responsibilities": "yes",
  "Sol_epr": "yes",
  "Sol_awarness": "yes",
  "Sol_segregation": "yes",
  "Sol_upcycling": "yes",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "yes",
  "Sol_RD": "yes",
  "Sol_tax": "no",
  "Sol_ban": "no",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "yes",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Lobti (leaf) plates at festivals in some Kathmandu communities; municipal "
    "cloth-bag distribution; department stores no longer providing free plastic bags"
  ),
}

CODEBOOK = [
  ("Country", "Country ISO code"),
  ("Country name", "Country name"),
  ("ID", "InterviewID: ISO_Nr"),
  ("Actortype", "governmental, pol_party, science, cso, igo, private_sector, shop_owner, hotel_owner, household"),
  ("Problem_awareness_pop", "high, medium, low, NA — lack of awareness among population mentioned"),
  ("Problem_awareness_pol", "high, medium, low, NA — lack of awareness among policy makers mentioned"),
  ("Problem_severity", "high, medium, low, NA"),
  ("Problem_littering", "Littering is a problem; was mentioned: yes/no"),
  ("Problem_consumption", "High consumption is a problem; was mentioned: yes/no"),
  ("Problem_recycling", "No or too little recycling is a problem; was mentioned: yes/no"),
  ("Problem_waste_mgmt", "Ill waste management is a problem; was mentioned: yes/no"),
  ("Problem_production", "High production rates; was mentioned: yes/no"),
  ("Problem_alternatives", "No good alternatives; was mentioned: yes/no"),
  ("Problem_waste_segregation", "Lack of segregation was mentioned: yes/no"),
  ("Problem_import", "Plastic imports are a problem; was mentioned: yes/no"),
  ("Impacts", "Impacts mentioned on water, cities, biodiversity, health, etc. or NA"),
  ("Governance", "Mentioned: yes/no"),
  ("Relevance_international_pol", "International treaties are relevant for national level policy"),
  ("Coordination_sectoral", "Lack of coordination across sectors"),
  ("Coordination_levels", "Lack of coordination across levels"),
  ("Unclear_responsibilities", "Unclear responsibilities among the involved actors"),
  ("Actors", "Mentioned: yes/no"),
  ("Federal government", "Mentioned: yes/no"),
  ("Provincial government", "Mentioned: yes/no"),
  ("Local government", "Mentioned: yes/no"),
  ("Students", "Mentioned: yes/no"),
  ("Private_sector", "Mentioned: yes/no"),
  ("Civil_society", "Mentioned: yes/no"),
  ("Science", "Mentioned: yes/no"),
  ("Households", "Mentioned: yes/no"),
  ("Education institutions", "Mentioned: yes/no"),
  ("Private_companies(hotels, shops, etc.)", "Mentioned: yes/no"),
  ("Implementation issues", "Mentioned: yes/no"),
  ("Monitoring", "Lack of monitoring"),
  ("Financial_resources", "Lack of financial resources"),
  ("Research", "Lack of research"),
  ("Infrastructure", "Lack of infrastructure"),
  ("Enforcement", "Lack of enforcement"),
  ("Policies in place", "Mentioned: yes/no"),
  ("Pol_epr", "Extended producer responsibility"),
  ("Pol_import", "Import regulations"),
  ("Pol_awarness", "Awareness campaigns"),
  ("Pol_education", "Education programs"),
  ("Pol_capacity", "Capacity building"),
  ("Pol_RD", "Research & development"),
  ("Pol_tax", "Levies or taxes on specific products (often bags)"),
  ("Pol_ban", "Bans of specific products"),
  ("Pol_subsitutes", "Alternatives like reusable bags, or banana leaf plates"),
  ("Pol_clean_up", "Clean-up campaigns"),
  ("Pol_upcycling", "Making a new product, like blacktop"),
  ("Pol_recycling", "Making a new raw material"),
  ("Pol_waste_collection", "Waste collection"),
  ("Solutions", "Mentioned: yes/no"),
  ("Sol_lead_agency", "Procedural measures to establish lead agency"),
  ("Sol_responsibilities", "Clear responsibilities"),
  ("Sol_epr", "Extended producer responsibility"),
  ("Sol_awarness", "Awareness raising measures"),
  ("Sol_segregation", "Segregation of waste"),
  ("Sol_upcycling", "Upcycling of plastics"),
  ("Sol_recycling", "New raw materials"),
  ("Sol_education", "Education programs at schools"),
  ("Sol_capacity", "Capacity-building programs"),
  ("Sol_RD", "Research programs"),
  ("Sol_tax", "Taxes or levies on products"),
  ("Sol_ban", "Ban of products"),
  ("Sol_finance", "Financial measures"),
  ("Sol_infrastructure", "Infrastructural measures"),
  ("Sol_subsitutes", "Alternatives"),
  ("Sol_clean_up", "Clean-up campaigns"),
  ("Sol_enforcement", "Enforcement procedures"),
  ("Sol_monitoring", "Monitoring procedures"),
  ("Traditions to build on (free-hand)", "Freehand or NA"),
]

NOTES = {
  "Actortype": (
    "Kathmandu Metropolitan City municipal officer responsible for dry waste management "
    "and local implementation of plastic/solid waste policy."
  ),
  "Problem_awareness_pop": (
    "People know microplastic harms (even in wildlife excreta) but still use plastic "
    "for convenience; no habit of carrying cloth bags. Notes: people not taking it "
    "seriously enough; education needed to reduce ignorance."
  ),
  "Problem_severity": (
    "Plastic share of waste rose from 8% (1998) to 15%; blocks drainage; major visual "
    "pollution; multi-layer plastic with no scavenger value left on riverbanks."
  ),
  "Problem_littering": (
    "Waste dumped riverside/roadside; visual pollution from tourism and lifestyle; "
    "no city space to store sorted waste."
  ),
  "Problem_consumption": (
    "Multi-layer and single-use plastics; lifestyle changes; department stores shifted "
    "but convenience consumption persists."
  ),
  "Problem_recycling": (
    "Multi-layer plastic has no value and is not collected; lack of recycling "
    "technology and factory capacity within the valley."
  ),
  "Problem_waste_mgmt": (
    "No space to separate, collect, or store within KTM; all waste transported to "
    "Bancharedanda in another district."
  ),
  "Problem_production": "Growing plastic share of municipal waste (8% to 15%).",
  "Problem_alternatives": (
    "Lobti leaf plates cost ~20 NPR vs cheap plastic; cloth bags must be purchased; "
    "alternatives unaffordable for most."
  ),
  "Problem_waste_segregation": (
    "Sorting into grades is done but land constraints limit on-site separation and "
    "storage; public opposition to neighbourhood waste storage."
  ),
  "Problem_import": (
    "Plastic <40 microns and plastic flowers imported from neighbouring countries "
    "despite bans."
  ),
  "Impacts": (
    "Cities: drainage blockage, visual pollution. Water: riverside dumping. Health: "
    "microplastics. Biodiversity: plastics found in wildlife excreta (research cited)."
  ),
  "Coordination_sectoral": (
    "Lack of stakeholder coordination; informal scrap collectors outside formal system; "
    "private sector only recently formalized."
  ),
  "Coordination_levels": (
    "Fragmented policies; no synchronization; national land control vs local waste "
    "duty; 23 municipalities use Bancharedanda but only KTM manages it."
  ),
  "Unclear_responsibilities": (
    "Uncertain which federal department oversees plastics; implementing/guiding body "
    "missing; valley municipalities lack visible landfill contribution."
  ),
  "Federal government": (
    "Ministry of Forest and Environment national policy (under revision); national "
    "land jurisdiction; Department of Environment."
  ),
  "Provincial government": (
    "Provincial help needed to identify land for MRF/recovery facilities within valley."
  ),
  "Local government": (
    "KTM responsible under SWM Act 2068; seven clusters; 18-municipality mayor forum; "
    "Environment and Natural Resource Conservation Act 2077 BS."
  ),
  "Students": "Waste management should enter education system as early as possible.",
  "Private_sector": (
    "Registered private collectors in seven clusters; DoCoRecyclers MoU (NPR 6.2 million/year); "
    "plastic companies still producing banned thin plastics."
  ),
  "Civil_society": "DoCoRecyclers (local NGO); UNDP partnership.",
  "Science": (
    "UNDP multilayer plastics project; UK university end-of-life research; research "
    "articles on microplastics in wildlife."
  ),
  "Households": (
    "Consumers targeted to reduce plastic use; convenience behaviour despite awareness."
  ),
  "Education institutions": (
    "Waste management should be taught from young age in the education system."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Hotels historically discarded bottles; department stores; plastic producers; "
    "scrap collectors when material has value."
  ),
  "Monitoring": (
    "Weak field inspection; no full assessment of micron-ban effectiveness; "
    "municipality should be accountable monitoring partner."
  ),
  "Financial_resources": (
    "Private investors lack confidence; sector lacks incentivization; need financing "
    "for collection chains and recycling facilities."
  ),
  "Research": (
    "No assessment of 20/40/75 micron policy effectiveness; end-of-life management "
    "research with UK university."
  ),
  "Infrastructure": (
    "Land is biggest constraint — no transfer stations or MRFs within city; lack of "
    "recycling factory and technology; Bansari Dada soil-cover problems."
  ),
  "Enforcement": (
    "Bans well-written on paper but not enforced on ground; companies still produce "
    "<40 micron bags; imported banned products widely available."
  ),
  "Policies in place": (
    "Environment & Natural Resource Conservation Act 2077 BS; Pollution Standard 2081; "
    "SWM regulation (cluster private-sector MoUs); national <40 micron and flower bans."
  ),
  "Pol_epr": "EPR recommended as first future policy but not yet implemented.",
  "Pol_import": (
    "National/municipal bans include import restrictions but neighbouring-country "
    "imports undermine enforcement."
  ),
  "Pol_awarness": (
    "Bhaktapur awareness campaigns; municipal cloth-bag production/distribution."
  ),
  "Pol_education": (
    "Advocacy for waste-management curriculum in schools from early age."
  ),
  "Pol_RD": (
    "UNDP multilayer plastics project; UK university end-of-life household products "
    "research; Pokhara skill-development training (400 participants)."
  ),
  "Pol_ban": (
    "Ban on plastic <40 microns (planned upgrade to 75 microns); plastic flowers banned; "
    "attempted SUP ban."
  ),
  "Pol_subsitutes": (
    "Municipal cloth bags; lobti leaf plates in some festival communities."
  ),
  "Pol_clean_up": (
    "Regular riverside debris extraction with excavators; transport to Dakshinkali (23 km)."
  ),
  "Pol_upcycling": (
    "UNDP recycling/up-cycling and paper up-cycling projects; multilayer plastics processing."
  ),
  "Pol_recycling": (
    "Dry waste sorting by grade; Deku transfer station; recyclables to processors via "
    "DoCoRecyclers."
  ),
  "Pol_waste_collection": (
    "Seven ward clusters with formalized private-sector collection under 2024 SWM regulation."
  ),
  "Sol_lead_agency": (
    "Need guiding/implementing body at national level (Ministry of Environment cited)."
  ),
  "Sol_responsibilities": (
    "Municipalities accountable for monitoring; federal EPR and import regulation; "
    "provincial coordination."
  ),
  "Sol_epr": "First recommended policy: nationwide EPR for producers.",
  "Sol_awarness": "Reduce public ignorance; sustained awareness on alternatives and habits.",
  "Sol_segregation": "Sorting into grades; mechanized recovery and storage (UNDP MoU).",
  "Sol_upcycling": "Expand UNDP up-cycling and multilayer processing pilots.",
  "Sol_recycling": (
    "End-of-life management focus over temporary bans; processing/treatment/recycling "
    "infrastructure; export policies for recycled products."
  ),
  "Sol_education": "Integrate waste management into education from youngest age.",
  "Sol_capacity": "Technology transfer — do not reinvent the wheel; UNDP mechanized systems.",
  "Sol_RD": "Research on low-value plastic destinations and end-of-life pathways.",
  "Sol_ban": (
    "Banning/replacing plastics is temporary; end-of-life management is the core problem."
  ),
  "Sol_finance": (
    "Incentivize sector; subsidies to increase plastic thickness (20→40→75 microns); "
    "finance collection chains and recycling facilities."
  ),
  "Sol_infrastructure": (
    "Federal/provincial help to secure valley land for MRF and plastic recovery facilities."
  ),
  "Sol_subsitutes": (
    "Subsidize producers of leaf plates and other alternatives; expand cloth-bag use."
  ),
  "Sol_clean_up": "Continue riverside debris removal; valley-wide landfill burden sharing.",
  "Sol_enforcement": "Regular field inspections; hold producers and importers accountable.",
  "Sol_monitoring": (
    "Municipality as implementing/monitoring partner; assess micron-standard effectiveness."
  ),
  "Traditions to build on (free-hand)": (
    "Lobti leaf plates at cultural festivals (wealthier communities); cloth bags at "
    "department stores; municipal cloth-bag distribution programme."
  ),
}


def hfill(hex_):
  return PatternFill("solid", fgColor=hex_)


def hfont(bold=False, size=10, color="000000", italic=False):
  return Font(bold=bold, size=size, color=color, italic=italic)


thin = Side(style="thin", color="BFBFBF")


def thin_border():
  return Border(left=thin, right=thin, top=thin, bottom=thin)


WRAP_ALIGN = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER_ALIGN = Alignment(wrap_text=True, vertical="center", horizontal="center")


def build_workbook():
  wb = openpyxl.Workbook()

  ws = wb.active
  ws.title = "Interview Coding"

  ws.merge_cells("A1:D1")
  title = ws["A1"]
  title.value = "Plastic Pollution Governance — Interview Coding (Codebook)"
  title.fill = hfill(DARK_BLUE)
  title.font = hfont(bold=True, size=13, color=WHITE)
  title.alignment = CENTER_ALIGN

  ws.merge_cells("A2:D2")
  meta = ws["A2"]
  meta.value = (
    "Interview: NPL_3  |  Country: NEPAL  |  Kathmandu Metropolitan City Officer  |  "
    "Actor: governmental  |  23 June 2025"
  )
  meta.fill = hfill(MID_BLUE)
  meta.font = hfont(size=9, color=WHITE)
  meta.alignment = CENTER_ALIGN

  headers = ["Variable", "Code", "Codebook definition", "Coding notes"]
  for col, header in enumerate(headers, start=1):
    cell = ws.cell(row=4, column=col, value=header)
    cell.fill = hfill(LIGHT_BLUE)
    cell.font = hfont(bold=True, size=10, color=DARK_BLUE)
    cell.alignment = CENTER_ALIGN
    cell.border = thin_border()

  row = 5
  for variable, definition in CODEBOOK:
    values = [
      variable,
      CODING[variable],
      definition,
      NOTES.get(variable, ""),
    ]
    fill = hfill(LIGHT_GREEN) if row % 2 == 0 else hfill(WHITE)
    for col, value in enumerate(values, start=1):
      cell = ws.cell(row=row, column=col, value=value)
      cell.fill = fill
      cell.font = hfont(size=10)
      cell.alignment = WRAP_ALIGN
      cell.border = thin_border()
    row += 1

  ws.column_dimensions["A"].width = 34
  ws.column_dimensions["B"].width = 28
  ws.column_dimensions["C"].width = 52
  ws.column_dimensions["D"].width = 58
  ws.freeze_panes = "A5"

  wide = wb.create_sheet("Wide Format")
  wide_headers = [variable for variable, _ in CODEBOOK]
  for col, header in enumerate(wide_headers, start=1):
    cell = wide.cell(row=1, column=col, value=header)
    cell.fill = hfill(LIGHT_BLUE)
    cell.font = hfont(bold=True, size=9, color=DARK_BLUE)
    cell.alignment = CENTER_ALIGN
    cell.border = thin_border()
  for col, header in enumerate(wide_headers, start=1):
    cell = wide.cell(row=2, column=col, value=CODING[header])
    cell.fill = hfill(ORANGE)
    cell.alignment = WRAP_ALIGN
    cell.border = thin_border()
    wide.column_dimensions[get_column_letter(col)].width = max(14, min(28, len(header) + 2))

  ref = wb.create_sheet("Codebook Reference")
  ref.merge_cells("A1:B1")
  ref["A1"].value = "Codebook Variable Definitions"
  ref["A1"].fill = hfill(DARK_BLUE)
  ref["A1"].font = hfont(bold=True, size=12, color=WHITE)
  ref["A1"].alignment = CENTER_ALIGN
  for idx, (variable, definition) in enumerate(CODEBOOK, start=3):
    ref.cell(row=idx, column=1, value=variable).font = hfont(bold=True)
    ref.cell(row=idx, column=2, value=definition).alignment = WRAP_ALIGN
  ref.column_dimensions["A"].width = 34
  ref.column_dimensions["B"].width = 70

  stakeholder = wb.create_sheet("Stakeholder Analysis")
  stakeholder.merge_cells("A1:B1")
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_3"
  stakeholder["A1"].fill = hfill(DARK_BLUE)
  stakeholder["A1"].font = hfont(bold=True, size=12, color=WHITE)
  stakeholder["A1"].alignment = CENTER_ALIGN
  for idx, (key, value) in enumerate(STAKEHOLDER.items(), start=3):
    stakeholder.cell(row=idx, column=1, value=key).font = hfont(bold=True)
    cell = stakeholder.cell(row=idx, column=2, value=value)
    cell.alignment = WRAP_ALIGN
  stakeholder.column_dimensions["A"].width = 28
  stakeholder.column_dimensions["B"].width = 90

  sources = wb.create_sheet("Sources")
  sources.merge_cells("A1:B1")
  sources["A1"].value = "Source Documents Used for Verification"
  sources["A1"].fill = hfill(DARK_BLUE)
  sources["A1"].font = hfont(bold=True, size=12, color=WHITE)
  sources["A1"].alignment = CENTER_ALIGN
  for idx, (doc, detail) in enumerate(SOURCES, start=3):
    sources.cell(row=idx, column=1, value=doc).font = hfont(bold=True)
    sources.cell(row=idx, column=2, value=detail).alignment = WRAP_ALIGN
  sources.column_dimensions["A"].width = 42
  sources.column_dimensions["B"].width = 60

  wb.save(OUTPUT_PATH)
  print(f"Saved {OUTPUT_PATH}")


if __name__ == "__main__":
  build_workbook()
