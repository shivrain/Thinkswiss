"""Generate interview coding spreadsheet for Nepal interview NPL_17.

Coding follows the Overview All Interviews codebook and is based on the
Romy Prasad Shrestha (elected ward chairman, five terms) interview —
interview guideline. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_17.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview guideline", "Romy Prasad Shrestha — ward chairman structured guideline"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_17",
  "Country": "NEPAL",
  "Interview number": "Ward chairman / local governance series",
  "Interview date": "June 2025 (fieldwork; exact date not recorded in guideline)",
  "Interviewee": "Mr. Romy Prasad Shrestha",
  "Affiliation / role": (
    "Elected ward chairman (five terms) — local governance and waste management"
  ),
  "Organization": (
    "Municipality ward office and Nagarpalika (municipal central); waste "
    "dumped in Ward 8 affecting Panchkhal downstream"
  ),
  "Stakeholder list mapping": "local gov",
  "Coded actor type (codebook)": "governmental",
  "Actor classification rationale": (
    "Five-term elected ward chairman exercising grassroots governance but "
    "without ward-level budget, land, or formal authority for plastic policy — "
    "local government actor."
  ),
  "Perspective": (
    "Experienced ward leader frustrated by limited power — deeply concerned "
    "about dump smell near home; advocates central delegation of budget/authority, "
    "private-company collection model, and traditional thokodi bags"
  ),
  "Project context": (
    "Land bought for wet/dry segregation and plastic refining but not approved; "
    "dustbin distribution and household collection rounds failed; micron ban "
    "unenforced; processing companies collect recyclables without formal ward MoU"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_17",
  "Actortype": "governmental",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "high",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "yes",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "yes",
  "Problem_import": "no",
  "Impacts": (
    "rivers, foul smell, visual pollution, health, air pollution, dumping sites"
  ),
  "Governance": "yes",
  "Relevance_international_pol": "no",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "yes",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "no",
  "Local government": "yes",
  "Students": "no",
  "Private_sector": "yes",
  "Civil_society": "no",
  "Science": "yes",
  "Households": "yes",
  "Education institutions": "no",
  "Private_companies(hotels, shops, etc.)": "yes",
  "Implementation issues": "yes",
  "Monitoring": "yes",
  "Financial_resources": "yes",
  "Research": "yes",
  "Infrastructure": "yes",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "no",
  "Pol_import": "no",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
  "Pol_capacity": "yes",
  "Pol_RD": "yes",
  "Pol_tax": "no",
  "Pol_ban": "yes",
  "Pol_subsitutes": "yes",
  "Pol_clean_up": "yes",
  "Pol_upcycling": "no",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "yes",
  "Sol_responsibilities": "yes",
  "Sol_epr": "no",
  "Sol_awarness": "yes",
  "Sol_segregation": "yes",
  "Sol_upcycling": "no",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "yes",
  "Sol_RD": "yes",
  "Sol_tax": "no",
  "Sol_ban": "yes",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "yes",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Traditional paper bags (thokodi) as culturally rooted alternative to "
    "plastic bags"
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
    "Romy Prasad Shrestha, five-term ward chairman. Stakeholder list: local gov. "
    "Codebook: governmental."
  ),
  "Problem_awareness_pop": (
    "Residents absent during collection rounds; plastic bags left outside "
    "uncollected; highway littering requires 24-hour monitoring; junk food "
    "packaging ubiquitous."
  ),
  "Problem_awareness_pol": (
    "Micron ban issued but no enforcement/monitoring; national waste policy "
    "draft on website but interviewee unaware; ward has no real power despite "
    "five terms."
  ),
  "Problem_severity": (
    "Deeply concerned — plastic destroying entire ecological system; foul smell "
    "and pollution from dump near home; no river free of plastics; visually "
    "polluting and spread by wind."
  ),
  "Problem_littering": (
    "Plastic everywhere including highways; sweepers dump in Ward 8 affecting "
    "Panchkhal downstream; wind spreads lightweight plastic."
  ),
  "Problem_consumption": (
    "All junk food comes in plastic packaging; finding alternatives is a "
    "challenge researchers should address."
  ),
  "Problem_recycling": (
    "Processing companies collect recyclables for resale but no formal ward "
    "arrangement; planned segregation/refining initiative failed — no collector–"
    "processor link."
  ),
  "Problem_waste_mgmt": (
    "Despite five terms as chairman, little achieved; sustainable mechanisms "
    "not developed; keeping own surroundings clean insufficient."
  ),
  "Problem_production": (
    "Ubiquitous plastic packaging in food products; ecological system destruction "
    "cited."
  ),
  "Problem_alternatives": (
    "Ban or viable alternatives welcome; researchers should address junk-food "
    "packaging substitutes; thokodi paper bags as cultural alternative."
  ),
  "Problem_waste_segregation": (
    "Planned wet/dry segregation and manure processing never materialized; "
    "household collection rounds failed when residents away."
  ),
  "Impacts": (
    "Rivers: no river free of plastics. Foul smell: dump near chairman's home "
    "affects residents. Visual pollution: harmful and wind-spread. Health: "
    "pollution and smell as direct personal impact. Air pollution: burning/"
    "dump emissions implied. Dumping sites: Ward 8 dump affects Panchkhal "
    "downstream."
  ),
  "Coordination_sectoral": (
    "No mechanism connecting waste collectors with processors; processing "
    "companies visit informally without ward MoU."
  ),
  "Coordination_levels": (
    "Central government must delegate budget and authority to LGs; ward lacks "
    "power while Nagarpalika responsible at municipal level; progress slow from "
    "municipality."
  ),
  "Unclear_responsibilities": (
    "No clear accountability at ward level; chairman takes primary responsibility "
    "but lacks authority and budget; private-company model proposed not "
    "formalized."
  ),
  "Federal government": (
    "Micron-plastic ban issued; national waste policy draft on website; must "
    "allocate ward budgets and delegate waste-management authority."
  ),
  "Local government": (
    "Ward chairman and Nagarpalika (municipal central) both responsible; "
    "municipal progress gradual."
  ),
  "Private_sector": (
    "Proposed model: private company given budget to collect, segregate, "
    "recycle valuable waste, landfill remainder — not formalized."
  ),
  "Science": (
    "Researchers should address alternatives to junk-food plastic packaging."
  ),
  "Households": (
    "Absent during collection rounds; must be present for dustbin/plastic "
    "collection to work."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Processing companies collect recyclable plastics for resale without "
    "formal ward agreement."
  ),
  "Monitoring": (
    "No dedicated ward monitoring personnel; micron ban not monitored; highway "
    "littering needs 24-hour monitoring."
  ),
  "Financial_resources": (
    "No central-government budget allocation for ward waste management; land "
    "bought for initiative but approval/funding blocked."
  ),
  "Research": (
    "Researchers needed for junk-food packaging alternatives and ecological "
    "impact understanding."
  ),
  "Infrastructure": (
    "Critical lack of land for disposal and segregation; land purchased but not "
    "approved for wet/dry processing facility."
  ),
  "Enforcement": (
    "Micron ban not enforced; dustbin distribution without follow-through; "
    "collection rounds not sustained."
  ),
  "Policies in place": (
    "Government ban on thin/micron plastics; national waste policy draft "
    "(online); ward dustbin distribution and collection rounds attempted."
  ),
  "Pol_awarness": "Awareness programs essential — core solution cited.",
  "Pol_education": (
    "Public awareness to change behaviour — be present during collection, "
    "reduce plastic use."
  ),
  "Pol_capacity": (
    "No ward-level waste-management staff; need designated monitoring personnel."
  ),
  "Pol_RD": "Research for packaging alternatives to junk-food plastics.",
  "Pol_ban": (
    "Micron/thin plastic ban issued nationally but not monitored or enforced."
  ),
  "Pol_subsitutes": "Traditional thokodi paper bags as viable cultural alternative.",
  "Pol_clean_up": (
    "Sweepers collect and dump at Ward 8; highway littering problem."
  ),
  "Pol_recycling": (
    "Processing companies collect recyclables; planned refine-and-sell "
    "initiative failed to launch."
  ),
  "Pol_waste_collection": (
    "Ward dustbin distribution and household plastic collection rounds "
    "(largely unsuccessful)."
  ),
  "Sol_lead_agency": (
    "Central government must take responsibility and formally delegate to LGs."
  ),
  "Sol_responsibilities": (
    "Formalize private-company collection/segregation/recycling/landfill model; "
    "clarify ward vs Nagarpalika accountability."
  ),
  "Sol_awarness": (
    "Essential awareness programmes; optimistic reduction possible within "
    "few years with right steps."
  ),
  "Sol_segregation": (
    "Wet/dry segregation with manure processing and plastic refining — land "
    "bought but initiative stalled."
  ),
  "Sol_recycling": (
    "Formal mechanism linking collectors to processors; private company "
    "recycles valuable fraction."
  ),
  "Sol_education": (
    "Educate households to be present during collection rounds."
  ),
  "Sol_capacity": (
    "Designated ward monitoring personnel; 24-hour highway monitoring capacity."
  ),
  "Sol_RD": (
    "Researcher-led alternatives for junk-food plastic packaging."
  ),
  "Sol_ban": (
    "Enforce micron ban with monitoring; ban viable if alternatives exist."
  ),
  "Sol_finance": (
    "Central budget allocation to wards for waste management; budget for "
    "private collection company model."
  ),
  "Sol_infrastructure": (
    "Approve land for segregation and disposal; connect collectors to processors."
  ),
  "Sol_subsitutes": (
    "Thokodi traditional paper bags; researcher-developed packaging alternatives."
  ),
  "Sol_clean_up": (
    "Sustained collection rounds; highway litter monitoring."
  ),
  "Sol_enforcement": (
    "Monitor and enforce micron ban; follow-through on dustbin programme."
  ),
  "Sol_monitoring": (
    "Dedicated ward waste-monitoring staff; 24-hour highway surveillance."
  ),
  "Traditions to build on (free-hand)": (
    "Traditional paper bags (thokodi) — culturally rooted alternative to "
    "plastic bags."
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
    "Interview: NPL_17  |  Country: NEPAL  |  Romy Prasad Shrestha  |  "
    "Actor: governmental  |  June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_17"
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
