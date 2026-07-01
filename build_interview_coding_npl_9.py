"""Generate interview coding spreadsheet for Nepal interview NPL_9.

Coding follows the Overview All Interviews codebook and is based on the
Dhulikhel municipality ward-head interview — transcript and interview
guideline. Ward-level local government perspective. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_9.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "Dhulikhel ward head panel — field interview (Ram Devi / Deep)"),
  ("Interview guideline", "Ward-level plastic governance guideline (Themes A–D)"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_9",
  "Country": "NEPAL",
  "Interview number": "Dhulikhel ward-level (fieldwork series)",
  "Interview date": "June 2025 (Dhulikhel fieldwork; exact time not recorded)",
  "Interviewee": (
    "Ward head and ward officials (Dhulikhel municipality); panel includes "
    "ward chairperson and ward office representatives"
  ),
  "Affiliation / role": (
    "Ward-level elected local government — grassroots administration under "
    "Dhulikhel municipality (Ward No. 2 agriculture impacts; dumping site Ward 8)"
  ),
  "Organization": (
    "Dhulikhel municipality ward office; Tol Sudhar Samiti (community "
    "improvement committees, ~100–150 households each)"
  ),
  "Stakeholder list mapping": "local gov",
  "Coded actor type (codebook)": "governmental",
  "Actor classification rationale": (
    "Elected ward-level local government official responsible for grassroots "
    "governance but without separate waste-management budget or authority — "
    "waste contracted at municipal level."
  ),
  "Perspective": (
    "Local implementer with limited mandate — emphasizes awareness, resources, "
    "and enforcement as three essentials; frustrated by municipal fee collection "
    "without ward-level fund allocation or segregation"
  ),
  "Project context": (
    "Municipal waste collection twice weekly by contracted company; dumping site "
    "in Ward 8 affecting Wards 2–3; Dhulikhel Hospital medical-waste issues; "
    "1,000 cloth-bag campaign with nursing college; ETPC NGO past collection"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_9",
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
    "agriculture, rivers, health, drainage blockage, soil, water, "
    "visual pollution, dumping sites"
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
  "Pol_import": "no",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
  "Pol_capacity": "no",
  "Pol_RD": "no",
  "Pol_tax": "yes",
  "Pol_ban": "yes",
  "Pol_subsitutes": "yes",
  "Pol_clean_up": "yes",
  "Pol_upcycling": "no",
  "Pol_recycling": "no",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "no",
  "Sol_responsibilities": "yes",
  "Sol_epr": "no",
  "Sol_awarness": "yes",
  "Sol_segregation": "yes",
  "Sol_upcycling": "no",
  "Sol_recycling": "no",
  "Sol_education": "yes",
  "Sol_capacity": "yes",
  "Sol_RD": "yes",
  "Sol_tax": "no",
  "Sol_ban": "no",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "no",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Bamboo and paper baskets; meat wrapped in straw bundles; large reusable "
    "cloth bags for carrying goods; pre-plastic era when people carried "
    "everything in reusable bags"
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
    "Dhulikhel ward head and ward officials. Stakeholder list: local gov. "
    "Codebook: governmental."
  ),
  "Problem_awareness_pop": (
    "People don't understand consequences; collect and store waste unaware of "
    "harm; businesses prioritize quick profits over plastic reduction. Ward head "
    "personally stopped plastic bottles after awareness program."
  ),
  "Problem_awareness_pol": (
    "Ward knows micron ban and municipal rules but enforcement lacking; policies "
    "exist on paper only; ward lacks authority to act alone on some matters."
  ),
  "Problem_severity": (
    "Major problem everywhere; severe in Ward 2 farmlands and rivers; plastic "
    "creating severe problems; farmland fertility affected."
  ),
  "Problem_littering": (
    "Plastic visible in farmlands and rivers; open syringes along Panchkhal stream; "
    "dumping-site runoff affects lower settlements."
  ),
  "Problem_consumption": (
    "Everything packaged in plastic (noodles, food); plastic convenient and cheap "
    "(1 rupee); even without bringing plastic home, products arrive in packaging."
  ),
  "Problem_recycling": (
    "No source segregation — all waste collected together in trucks; some "
    "segregation later but not properly."
  ),
  "Problem_waste_mgmt": (
    "Waste collected twice weekly and dumped at one site (Ward 8); annual cleaning "
    "taxes collected but no effective execution; implementation aspect weak."
  ),
  "Problem_alternatives": (
    "No affordable alternatives; plastic cheap and convenient; awareness alone "
    "insufficient without substitutes."
  ),
  "Problem_waste_segregation": (
    "No segregation at collection; municipality handles all waste together; ward "
    "has no manpower or materials to manage separately."
  ),
  "Impacts": (
    "Agriculture: plastic in farmlands, fertility severely affected. Rivers: plastic "
    "in rivers, Panchkhal stream syringes. Health: plastic bottles, syringe "
    "scavenging, medical waste mixed at dump. Drainage blockage: canal blockages. "
    "Soil: farmland fertility loss. Water: dumping-site runoff to lower "
    "settlements (topography). Visual pollution: open waste. Dumping sites: Ward 8 "
    "dump affects Wards 2–3."
  ),
  "Coordination_sectoral": (
    "Ward–hospital–nursing college–NGO coordination attempted but not sustained; "
    "businesses resisted plastic-container campaign."
  ),
  "Coordination_levels": (
    "Ward reports to municipality; no ward budget for waste; some matters under "
    "higher authorities; unclear fund distribution across 12 wards."
  ),
  "Unclear_responsibilities": (
    "Waste management contracted to company at municipal level; ward has no "
    "authority, economic independence, or separate budget; ward is first contact "
    "but cannot manage dumping."
  ),
  "Federal government": (
    "National micron-thickness ban (~3–4 years ago); factories resisted, not enforced."
  ),
  "Local government": (
    "Dhulikhel municipality collects cleaning fees and manages contracted waste "
    "collection; ward office for awareness only."
  ),
  "Private_sector": (
    "Private company contracted for municipal waste collection; private companies "
    "could collect from households for fee (proposed)."
  ),
  "Civil_society": (
    "ETPC (Environment Tourism Committee) NGO past household collection; Tol Sudhar "
    "Samiti community committees under ward guidance."
  ),
  "Science": (
    "Tribhuvan University coordination for market-area development; Kathmandu "
    "University researchers conducting study; researcher support welcomed."
  ),
  "Households": (
    "Source of plastic waste; pay annual cleaning fees; Tol Sudhar Samiti covers "
    "100–150 households each."
  ),
  "Education institutions": (
    "Nursing college collaborated on 1,000 cloth-bag distribution; TU and KU "
    "research engagement."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Hotels and shops received cloth bags; restaurants resisted stopping plastic "
    "containers; businesses prioritize profit."
  ),
  "Monitoring": (
    "No monitoring of what kind of plastic is produced or used; ban not enforced."
  ),
  "Financial_resources": (
    "Ward lacks budget and manpower; annual cleaning fees collected by municipality "
    "but ward fund allocation unclear; programs need resources for sustainability."
  ),
  "Research": (
    "Researchers' support would help continue ward efforts; KU plastic governance "
    "study welcomed."
  ),
  "Infrastructure": (
    "No ward-level dumping site; municipal dump in Ward 8; hospital lacks own "
    "dumping site; medical waste still dumped at same place."
  ),
  "Enforcement": (
    "Micron ban and municipal rules exist on paper only; government couldn't "
    "enforce; factories resisted."
  ),
  "Policies in place": (
    "Micron-thickness ban (~40 microns, 3–4 years ago); municipal-level rules; "
    "annual cleaning tax; Environment Day rallies."
  ),
  "Pol_awarness": (
    "Environment Day rallies; cloth-bag campaign with nursing college; awareness "
    "programs changed ward head's bottle use."
  ),
  "Pol_education": (
    "Education identified as biggest challenge; awareness essential before any "
    "program works."
  ),
  "Pol_tax": "Annual cleaning taxes collected by municipality from residents.",
  "Pol_ban": "Plastics below certain micron thickness banned nationally (~3–4 years).",
  "Pol_subsitutes": "~1,000 cloth bags distributed to hotels and shops.",
  "Pol_clean_up": "Annual cleaning tax collected but execution ineffective.",
  "Pol_waste_collection": (
    "Municipality-contracted company collects twice weekly from households."
  ),
  "Sol_responsibilities": (
    "Ward should implement rules like road building; coordinate with private "
    "companies and Tol Sudhar Samiti for pilot projects."
  ),
  "Sol_awarness": (
    "One of three essentials (with resources and enforcement); people will act "
    "once they understand impacts."
  ),
  "Sol_segregation": (
    "Household separation at source proposed; separate bins or cloth bags; ward "
    "lacks authority to implement alone."
  ),
  "Sol_education": (
    "Biggest challenge is community education; awareness campaigns and bin "
    "distribution needed."
  ),
  "Sol_capacity": (
    "Ward lacks manpower and materials; municipality technical backing needed."
  ),
  "Sol_RD": (
    "Researcher support and TU collaboration for market-area development; "
    "post-research recommendations requested."
  ),
  "Sol_finance": (
    "Resources required for awareness and sanitation programs; ward initial "
    "budget for TU pilot."
  ),
  "Sol_infrastructure": (
    "Need proper dumping/segregation infrastructure; hospital needs separate "
    "medical-waste disposal."
  ),
  "Sol_subsitutes": (
    "Bamboo baskets, straw meat wrapping, cloth bags as affordable alternatives "
    "needed."
  ),
  "Sol_enforcement": (
    "Strict policy enforcement as third essential alongside awareness and resources."
  ),
  "Sol_monitoring": (
    "Monitoring of plastic types and ban compliance needed at local level."
  ),
  "Traditions to build on (free-hand)": (
    "Bamboo and paper baskets; straw bundles for meat; large reusable cloth bags; "
    "era before plastic when reusable bags were standard."
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
    "Interview: NPL_9  |  Country: NEPAL  |  Dhulikhel ward head  |  "
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_9"
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
