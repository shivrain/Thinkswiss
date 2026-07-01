"""Generate interview coding spreadsheet for Nepal interview NPL_16.

Coding follows the Overview All Interviews codebook and is based on the
Kathmandu Ward No. 6 interview of 22 June 2025 — interview guideline and
field notes. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_16.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview notes", "Kathmandu Ward No. 6 — 22 June 2025 (11:30am – 12:30pm)"),
  ("Interview guideline", "Ward-level plastic governance guideline (Themes A–D)"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_16",
  "Country": "NEPAL",
  "Interview number": "1 (Kathmandu Ward No. 6 series)",
  "Interview date": "22 June 2025 (Sunday, 11:30am – 12:30pm)",
  "Interviewee": "Ward No. 6 official — Kathmandu municipality (name/title not recorded)",
  "Affiliation / role": (
    "Ward-level local government representative — solid waste management "
    "at Ward No. 6"
  ),
  "Organization": (
    "Kathmandu municipality — Ward No. 6; waste from all 12 wards goes to "
    "landfill at Ward 8; private company contracted for collection; "
    "Panchkhal open sewerage noted nearby"
  ),
  "Stakeholder list mapping": "local gov",
  "Coded actor type (codebook)": "governmental",
  "Actor classification rationale": (
    "Elected/appointed ward-level official responsible for local waste "
    "management advocacy but without authority to formulate plastic rules — "
    "local government actor."
  ),
  "Perspective": (
    "Ward implementer with limited mandate — emphasizes source-level action, "
    "alternatives accessibility, pilot household committees, and gap between "
    "policies on paper and ward-level implementation capacity"
  ),
  "Project context": (
    "Waste collected twice weekly to single dump; rains wash waste downstream; "
    "1,000 paper-bag hospital advocacy (unsustainable); 40-micron ban 3–4 years "
    "ago ineffective; no ward budget — private company manages waste"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_16",
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
    "water, rivers, agriculture, soil, health"
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
  "Civil_society": "yes",
  "Science": "no",
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
  "Sol_lead_agency": "no",
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
  "Sol_clean_up": "no",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "no",
  "Traditions to build on (free-hand)": (
    "Paper thungas (bags), leaves, and natural materials; pre-plastic era "
    "practices to revive; old ways without plastics"
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
    "Kathmandu Ward No. 6 official. Stakeholder list: local gov. Codebook: "
    "governmental."
  ),
  "Problem_awareness_pop": (
    "People not comfortable with alternatives — not available or accessible; "
    "plastic affordable and easily available; less knowledge of health/"
    "environment impacts."
  ),
  "Problem_awareness_pol": (
    "No regulatory mechanism in place; many policies exist but implementation "
    "lacking; ward has no adhikar (authority) to formulate plastic rules."
  ),
  "Problem_severity": (
    "Problem everywhere; big issue at ward level; plastics all over water "
    "bodies and rivers; affects areas below when rain washes dumped waste "
    "downstream."
  ),
  "Problem_littering": (
    "Plastics all over water bodies and rivers; open-space deposit at hospital "
    "agreement; Panchkhal open sewerage noted."
  ),
  "Problem_consumption": (
    "Everything literally comes in plastics; plastic affordable and easily "
    "available; need to start replacing drinking water bottles."
  ),
  "Problem_recycling": (
    "Only one NGO worked on recycling once, now none; ten years ago punching/"
    "hooking collection from household level; recycling options need exploration "
    "with enforcement beyond single NGO capacity."
  ),
  "Problem_waste_mgmt": (
    "Waste gathered twice weekly and dumped in one place; hospital waste "
    "management unit exists but no dumping site; burn suggestion difficult "
    "due to volume."
  ),
  "Problem_production": (
    "Everything packaged in plastic; problem at source — budgetary constraints "
    "and lack of alternatives."
  ),
  "Problem_alternatives": (
    "Alternatives talked about but nothing happening; people not comfortable "
    "without accessible options; 1,000 paper bags to hospitals not sustainable."
  ),
  "Problem_waste_segregation": (
    "No segregation at source; segregation only at dumping site; pilot proposed "
    "for 100–200 household committee covering segregation."
  ),
  "Impacts": (
    "Water/rivers: plastics all over water bodies and rivers; rain washes "
    "dumped waste downstream affecting lower areas. Agriculture: plastics in "
    "agricultural fields. Soil: affects soil fertility. Health: less current "
    "knowledge — impacts and effects research needed."
  ),
  "Coordination_sectoral": (
    "Tried hospital collaboration — did not work; hospital has waste unit and "
    "dumpsite agreement but open-space deposit problems."
  ),
  "Coordination_levels": (
    "Ward has no authority to formulate rules; private company manages waste "
    "without ward budget; all 12 municipal wards dump at Ward 8 landfill."
  ),
  "Unclear_responsibilities": (
    "'Everyone's responsibility' but ward lacks adhikar for plastic rules; "
    "private company should be given clearer responsibility per respondent."
  ),
  "Federal government": (
    "40-micron plastics ban notification issued 3–4 years ago but did not work; "
    "regulatory mechanism absent at implementation level."
  ),
  "Local government": (
    "Ward-level management responsibility acknowledged; advocacy at every level; "
    "no ward budget for waste — municipality/private contract."
  ),
  "Private_sector": (
    "Private company managing ward waste collection; should be given "
    "responsibility for better management."
  ),
  "Civil_society": (
    "Past NGO recycling initiative (only once); 1,000 paper-bag advocacy with "
    "community and hospitals — not sustainable."
  ),
  "Households": (
    "Household-level action needed; pilot 100–200 household committee for "
    "segregation, recycling, processing, and networking."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Hospitals received paper-bag advocacy; hospital waste management unit "
    "without proper dumpsite."
  ),
  "Monitoring": (
    "40-micron ban not enforced effectively; no regulatory mechanism at ward level."
  ),
  "Financial_resources": (
    "No budget at ward level; problem at source includes budgetary constraints; "
    "private company funded through municipal contract."
  ),
  "Research": (
    "Less knowledge about plastic impacts on environment and health — research "
    "needed to inform action."
  ),
  "Infrastructure": (
    "No hospital dumping site; open-space deposit; landfill at Ward 8 for all "
    "12 wards; Panchkhal open sewerage."
  ),
  "Enforcement": (
    "40-micron ban did not work; policies in place but implementation lacking; "
    "needs enforcement beyond NGO capacity."
  ),
  "Policies in place": (
    "40-micron ban notification (~3–4 years ago); advocacy/awareness campaigns "
    "at every level; many policies exist on paper."
  ),
  "Pol_awarness": (
    "Advocacy carried out at every level; 1,000 paper bags distributed to "
    "hospitals with community."
  ),
  "Pol_education": "Awareness necessary — core solution cited.",
  "Pol_capacity": (
    "People not able to do management; capacity gap at ward and household level."
  ),
  "Pol_RD": (
    "Research on plastic impacts on environment and health needed."
  ),
  "Pol_ban": (
    "40-micron plastics ban initiated; government ban 3–4 years ago ineffective."
  ),
  "Pol_subsitutes": (
    "Paper thungas, leaves, natural materials; paper bags to hospitals; "
    "replace drinking water bottles."
  ),
  "Pol_clean_up": (
    "Waste collected twice weekly; dumping-site segregation (not at source)."
  ),
  "Pol_recycling": (
    "Past NGO household collection (punching/hooking plastics); recycling "
    "options need exploration with enforcement."
  ),
  "Pol_waste_collection": (
    "Twice-weekly collection to single dump site; private company manages."
  ),
  "Sol_responsibilities": (
    "Give private waste company clearer responsibility; everyone's "
    "responsibility but ward needs mandate."
  ),
  "Sol_awarness": (
    "Awareness necessary at all levels; advocacy campaigns continued."
  ),
  "Sol_segregation": (
    "Source segregation via pilot 100–200 household committee covering "
    "segregation, recycling, processing, and networking."
  ),
  "Sol_recycling": (
    "Explore recycling beyond single NGO; household-level collection model "
    "from ten years ago as reference."
  ),
  "Sol_education": (
    "Educate on health/environment impacts; household-level behaviour change."
  ),
  "Sol_capacity": (
    "Build household and ward capacity; pilot committee model; empower "
    "private company with responsibility."
  ),
  "Sol_RD": (
    "Research plastic impacts on environment and health to fill knowledge gap."
  ),
  "Sol_ban": (
    "Enforce 40-micron regulation effectively — previous ban failed."
  ),
  "Sol_finance": (
    "Ward budget needed; address budgetary constraints at source; municipal "
    "funding for private contractor accountability."
  ),
  "Sol_infrastructure": (
    "Proper hospital dumpsite; reduce open-space deposit and burning; "
    "landfill management at Ward 8 for all 12 wards."
  ),
  "Sol_subsitutes": (
    "Provide accessible affordable alternatives; revive paper thungas, leaves, "
    "natural materials; replace drinking water bottles first."
  ),
  "Sol_enforcement": (
    "Regulatory mechanism needed; enforce existing policies; beyond NGO-only "
    "recycling efforts."
  ),
  "Traditions to build on (free-hand)": (
    "Paper thungas, leaves, and natural materials; pre-plastic era practices; "
    "revive old ways without plastics."
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
    "Interview: NPL_16  |  Country: NEPAL  |  Kathmandu Ward No. 6  |  "
    "Actor: governmental  |  22 June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_16"
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
