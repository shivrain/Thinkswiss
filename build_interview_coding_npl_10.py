"""Generate interview coding spreadsheet for Nepal interview NPL_10.

Coding follows the Overview All Interviews codebook and is based on the
Dhulikhel municipality veteran ward-chairman interview (Ward No. 3 /
Panchkhaal–Thakuri area) — transcript and interview guideline. Primary
verification against ward-head transcription. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_10.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "Dhulikhel ward chairman — Ram Devi / Deep field interview"),
  ("Interview guideline", "Ward-level plastic governance guideline (Themes A–D)"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
  (
    "Coding note",
    "Thematic guideline notes include policy-science content (Bio-Camp, Plast "
    "Foundation) not spoken in this transcript; codes verified primarily against "
    "the ward-chairman transcription.",
  ),
]

STAKEHOLDER = {
  "Interview ID": "NPL_10",
  "Country": "NEPAL",
  "Interview number": "Dhulikhel ward-level (veteran chairman)",
  "Interview date": "June 2025 (Dhulikhel fieldwork; exact time not recorded)",
  "Interviewee": (
    "Ward chairman — Dhulikhel municipality (32 years in office, five terms); "
    "Ward No. 3 (Thakuri village / Panchkhaal area)"
  ),
  "Affiliation / role": (
    "Elected ward-level local government representative — long-serving chairman "
    "with direct responsibility for waste dumping sites and transfer-station planning"
  ),
  "Organization": "Dhulikhel municipality — Ward No. 3",
  "Stakeholder list mapping": "local gov",
  "Coded actor type (codebook)": "governmental",
  "Actor classification rationale": (
    "Elected ward chairman exercising grassroots governance over waste dumping, "
    "segregation pilots, and municipal coordination — not national ministry or CSO."
  ),
  "Perspective": (
    "Experienced local implementer-advocate — personal zero-waste practice; "
    "emphasizes enforcement, penalties, municipality empowerment, transfer station, "
    "and international models (Bhutan, Germany, Switzerland, US)"
  ),
  "Project context": (
    "20–30 years open dumping at Panchkhaal riverbank; transfer station and "
    "refining area in Thakuri village (land purchased); waste from Banepa and "
    "Okharpauwa; Sunday plastic collection; Rs 10–15/kg recyclables; landfill "
    "plan 1–2 years"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_10",
  "Actortype": "governmental",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "medium",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "no",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "yes",
  "Problem_import": "no",
  "Impacts": (
    "rivers, health, visual pollution, foul smell, dumping sites, air pollution, "
    "forests, soil, water"
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
  "Financial_resources": "no",
  "Research": "yes",
  "Infrastructure": "yes",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "no",
  "Pol_import": "no",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
  "Pol_capacity": "no",
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
  "Sol_tax": "yes",
  "Sol_ban": "yes",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "yes",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Baskets from natural materials; paper bags for vegetables; cloth bags "
    "encouraged; pre-plastic era when plastic was unnecessary"
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
    "Dhulikhel Ward No. 3 chairman (32 years, five terms). Stakeholder list: "
    "local gov. Codebook: governmental."
  ),
  "Problem_awareness_pop": (
    "Awareness improved through door-to-door programs and Bhutan/Germany examples; "
    "households still leave plastic outside containers or throw bags; people won't "
    "change without penalties."
  ),
  "Problem_awareness_pol": (
    "Big gap between local and federal government — plans and talk but very slow "
    "progress; no master plans; central government must push hard (Bagmati)."
  ),
  "Problem_severity": (
    "Ecosystem already destroyed; no good rivers in Nepal; Panchkhaal riverbank "
    "20–30 years dumping; daily smell near chairman's house; dust everywhere."
  ),
  "Problem_littering": (
    "Plastic spread everywhere; waste in forests and along highways; riverbank "
    "covered; wind carries lightweight plastic."
  ),
  "Problem_consumption": (
    "Plastic used everywhere — food packaging, drinking glasses, fast-food "
    "packaging; would be great if banned with alternatives."
  ),
  "Problem_recycling": (
    "2–4 people separate paper/plastic at dump; Rs 10–15/kg paid for refining; "
    "processing-company agreements terminated."
  ),
  "Problem_waste_mgmt": (
    "Sweepers collect and dump; some open dumping sites; waste still ends up in "
    "river despite collection efforts."
  ),
  "Problem_alternatives": (
    "Alternatives needed for fast-food packaging; researchers should find "
    "substitutes; cloth bags and natural baskets promoted."
  ),
  "Problem_waste_segregation": (
    "Household containers distributed but people left plastic outside; "
    "biodegradable/non-biodegradable separation promoted; Sunday plastic-only "
    "collection proposed."
  ),
  "Impacts": (
    "Rivers: Panchkhaal riverbank pollution, no good rivers in Nepal. Health: "
    "plastic harmful when burned, affects environment. Visual pollution: dusty "
    "landscapes, embarrassing vs Switzerland image. Foul smell: daily near dump "
    "near chairman's home. Dumping sites: 20–30 years open dumping. Air pollution: "
    "burning plastic pollutes environment. Forests: waste thrown in forests. "
    "Soil/water: ecosystem destroyed; river impacts on lower areas."
  ),
  "Coordination_sectoral": (
    "Coordination with big company on categorized waste at dump; land disputes "
    "with local people; contractor system for collection."
  ),
  "Coordination_levels": (
    "Big gap local–federal; municipality main authority but central must act; "
    "draft SWM Act on webpage but ward hasn't reached that stage."
  ),
  "Unclear_responsibilities": (
    "Central government should take charge vs municipality given full power to "
    "make laws and penalties; ward chairman takes personal responsibility as "
    "elected representative."
  ),
  "Federal government": (
    "Restricts plastic beyond certain size; slow progress on plans; Bagmati "
    "World Bank/ADB funding; land ownership request to Nepal government for Ward 3."
  ),
  "Local government": (
    "Dhulikhel municipality contracts cleanliness; frequent municipal meetings; "
    "comparing to Dharan municipality model; landfill plan passed 1–2 years."
  ),
  "Private_sector": (
    "Private companies visit monthly for collection; contractors given waste "
    "contracts; big company discussed for refining area."
  ),
  "Science": (
    "Researchers will know how to solve problem by finding alternatives; "
    "Ram Devi/KU research team conducting study."
  ),
  "Households": (
    "Door-to-door container program; households separate biodegradable and "
    "non-biodegradable; swap system for designated waste areas."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Processing/refining companies (agreements terminated); waste trucks with "
    "whistle signal for collection."
  ),
  "Monitoring": (
    "Cannot monitor highways and forests 24/7; people throw waste when unwatched."
  ),
  "Research": (
    "Researchers needed for plastic alternatives; good plan and data needed like "
    "Switzerland."
  ),
  "Infrastructure": (
    "Transfer station and refining area in Thakuri village (land purchased); "
    "landfill plan 1–2 years; land acquisition time-consuming; access roads built."
  ),
  "Enforcement": (
    "Enforcement is key; lawmakers must make strict laws; penalties needed — "
    "Germany model of double charge for unsegregated waste (Rs 200 vs 500)."
  ),
  "Policies in place": (
    "Government plastic size restriction; municipal waste contracts; Environment "
    "Day source-collection campaigns; draft SWM Act (not yet at ward level)."
  ),
  "Pol_awarness": (
    "Door-to-door programs; Environment Day campaigns; Bhutan morning public "
    "announcements cited as model."
  ),
  "Pol_education": (
    "Awareness programs; teaching alone insufficient without penalties."
  ),
  "Pol_RD": "Researchers to find alternatives to fast-food and packaging plastics.",
  "Pol_ban": (
    "Would be great if plastic banned; government restricts use beyond certain size."
  ),
  "Pol_subsitutes": (
    "Cloth bags encouraged; natural-material baskets and paper bags before plastic."
  ),
  "Pol_clean_up": (
    "Highway cleaning; source waste collection on Environment Day; swap system."
  ),
  "Pol_recycling": (
    "Plastic/paper/glass separated at dump; sent for refining at Rs 10–15/kg."
  ),
  "Pol_waste_collection": (
    "Waste trucks whistle for collection; Sunday plastic collection; monthly "
    "private company visits; contractors under municipality."
  ),
  "Sol_lead_agency": (
    "Central government must take charge; municipality given full power to make "
    "laws, set penalties, and enforce."
  ),
  "Sol_responsibilities": (
    "Elected ward chairman takes responsibility; municipality as main authority; "
    "local self-reliance to prevent pollution."
  ),
  "Sol_awarness": (
    "Door-to-door awareness; Bhutan/Germany international examples; public "
    "announcements for segregation."
  ),
  "Sol_segregation": (
    "Biodegradable/non-biodegradable separation; Sunday plastic-only collection; "
    "household containers; separate trucks for segregated vs unsegregated."
  ),
  "Sol_recycling": (
    "Refining at Rs 10–15/kg; 2–4 people separate recyclables at transfer area."
  ),
  "Sol_education": "Awareness improved but penalties required for behavior change.",
  "Sol_capacity": (
    "No designated ward-level waste staff; follow municipal contract system."
  ),
  "Sol_RD": "Research-based alternatives for packaging plastics.",
  "Sol_tax": (
    "Germany model — lower fee (Rs 200) for segregated vs higher (Rs 500) for "
    "unsegregated waste at tax collection."
  ),
  "Sol_ban": "Ban plastic with viable alternatives in place.",
  "Sol_finance": (
    "Municipality provides funds through contracts; differential charging incentivizes "
    "segregation."
  ),
  "Sol_infrastructure": (
    "Transfer station in Thakuri village; landfill 1–2 years; waste from Banepa "
    "and Okharpauwa processed locally before transport."
  ),
  "Sol_subsitutes": (
    "Natural baskets, paper bags, cloth bags; pre-plastic practices revived."
  ),
  "Sol_clean_up": (
    "Source-level action accumulating over time; chairman personally picks up litter."
  ),
  "Sol_enforcement": (
    "Strict laws and penalties essential; people won't change without being penalized."
  ),
  "Sol_monitoring": (
    "24/7 monitoring impossible; municipal and ward oversight of contractors."
  ),
  "Traditions to build on (free-hand)": (
    "Baskets from natural materials; paper bags for vegetables; cloth bags; era "
    "before plastic when alternatives sufficed."
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
    "Interview: NPL_10  |  Country: NEPAL  |  Dhulikhel Ward 3 chairman  |  "
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_10"
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
