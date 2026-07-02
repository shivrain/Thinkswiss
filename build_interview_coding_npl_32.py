"""Generate interview coding spreadsheet for Nepal interview NPL_32.

Coding follows the Overview All Interviews codebook and is based on the
Furinji Sherpa (trekking guide, Khumbu Pasang Lhamu) interview of
19 February 2026 — interview guideline (recorded). Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_32.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview guideline", "Furinji Sherpa trekking guide — guideline, 19 Feb 2026"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
  (
    "Coding note",
    "Consent given with recording. Questions 5, 6, 9, and 10 not asked. "
    "Complements NPL_19 (TAAN/trekking-sector perspective) from guide-level "
    "Khumbu municipal system view.",
  ),
]

STAKEHOLDER = {
  "Interview ID": "NPL_32",
  "Country": "NEPAL",
  "Interview number": "Trekking guide / Khumbu fieldwork series",
  "Interview date": "19 February 2026",
  "Interviewee": "Furinji Sherpa",
  "Affiliation / role": "Trekking guide",
  "Organization": (
    "Khumbu Pasang Lhamu Rural Municipality home village; Kathmandu/mountain "
    "trekking routes"
  ),
  "Stakeholder list mapping": "private_sector",
  "Coded actor type (codebook)": "private_sector",
  "Actor classification rationale": (
    "Trekking guide in tourism/private trekking sector — operational perspective "
    "on mountain litter, tourist behaviour, and rural municipal waste systems."
  ),
  "Perspective": (
    "Concerned trekking practitioner — plastics biggest problem in Kathmandu and "
    "mountains; tourist litter on high peaks costly to remove; Khumbu weekly "
    "segregated collection flown to Kathmandu; advocates awareness, recycling, "
    "cloth baskets, government policy, and littering fines for locals and tourists"
  ),
  "Project context": (
    "Khumbu Pasang Lhamu RM: weekly pickup, household segregation, plane to "
    "Kathmandu, partial burning/recycling; municipality distributing 4–5 cloth "
    "bags per house; national government sends cans/glass to India for recycling"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_32",
  "Actortype": "private_sector",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "medium",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "no",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "no",
  "Problem_import": "no",
  "Impacts": (
    "wildlife, health, tourism, forests, cities"
  ),
  "Governance": "yes",
  "Relevance_international_pol": "yes",
  "Coordination_sectoral": "no",
  "Coordination_levels": "no",
  "Unclear_responsibilities": "no",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "no",
  "Local government": "yes",
  "Students": "no",
  "Private_sector": "yes",
  "Civil_society": "no",
  "Science": "no",
  "Households": "yes",
  "Education institutions": "no",
  "Private_companies(hotels, shops, etc.)": "no",
  "Implementation issues": "yes",
  "Monitoring": "no",
  "Financial_resources": "yes",
  "Research": "no",
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
  "Pol_ban": "no",
  "Pol_subsitutes": "yes",
  "Pol_clean_up": "no",
  "Pol_upcycling": "yes",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "no",
  "Sol_responsibilities": "no",
  "Sol_epr": "no",
  "Sol_awarness": "yes",
  "Sol_segregation": "no",
  "Sol_upcycling": "yes",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "no",
  "Sol_RD": "no",
  "Sol_tax": "yes",
  "Sol_ban": "no",
  "Sol_finance": "no",
  "Sol_infrastructure": "no",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "no",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "no",
  "Traditions to build on (free-hand)": (
    "Clothing/cloth bags; handmade handy baskets for dry food (biscuits); "
    "Khumbu municipality distributing 4–5 cloth bags per household"
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
    "Furinji Sherpa — trekking guide (Khumbu). Stakeholder list: private_sector. "
    "Codebook: private_sector."
  ),
  "Problem_awareness_pop": (
    "Mixed awareness — ~60–70% of friends concerned, others not; tourists sometimes "
    "educate locals on trash management."
  ),
  "Problem_awareness_pol": (
    "Respondent doesn't really know national policies; aware of municipal Khumbu "
    "system and national recycling exports."
  ),
  "Problem_severity": (
    "Plastics are the biggest problem in Kathmandu and the mountains."
  ),
  "Problem_littering": (
    "Tourists throw waste in mountains; plastics on high peaks — hard and expensive "
    "to clean."
  ),
  "Problem_consumption": (
    "Difficult to find ways to re-use plastics in trekking/mountain context."
  ),
  "Problem_recycling": (
    "Need more recycling/reuse options; national government sends separated cans/"
    "glass to India; village partly recycles."
  ),
  "Problem_waste_mgmt": (
    "High-altitude cleanup costly; Khumbu system flies waste to Kathmandu weekly; "
    "partly burning remains."
  ),
  "Problem_alternatives": (
    "Reuse difficult; advocates handy baskets for dry food and cloth bags distributed "
    "by municipality."
  ),
  "Impacts": (
    "Wildlife: harmful for animals. Health: diseases from plastics. Environment/"
    "forests: mountain environment degraded. Tourism/cities: Kathmandu and peak "
    "trekking areas affected."
  ),
  "Relevance_international_pol": (
    "Recyclables (cans, glass bottles) sent to India for processing; tourists "
    "bring foreign waste-management ideas to locals."
  ),
  "Federal government": (
    "National government collects, separates, and exports recyclables to India; "
    "more government policies advocated."
  ),
  "Local government": (
    "Khumbu Pasang Lhamu Rural Municipality: weekly pickup, segregation, airlift "
    "to Kathmandu; distributing 4–5 cloth bags per household; focus on non-plastic "
    "packaging."
  ),
  "Private_sector": (
    "Trekking guide/tourism sector — tourist litter source and awareness channel."
  ),
  "Households": (
    "Khumbu households separate waste; village-level compliance model."
  ),
  "Financial_resources": (
    "Cleaning plastics from high peaks hard and expensive — resource barrier."
  ),
  "Infrastructure": (
    "Weekly collection plus aircraft transport to Kathmandu — remote mountain "
    "logistics."
  ),
  "Enforcement": (
    "Fines for littering advocated for both locals and tourists."
  ),
  "Policies in place": (
    "Khumbu weekly segregated collection and airlift; partial burning/recycling; "
    "cloth-bag distribution; national export of recyclables to India; municipal "
    "non-plastic packaging focus."
  ),
  "Pol_awarness": (
    "Tourists share trash-management ideas with locals; more public information "
    "needed."
  ),
  "Pol_education": (
    "More information for public to increase awareness proposed."
  ),
  "Pol_tax": (
    "Fines for littering for everyone — locals and tourists."
  ),
  "Pol_subsitutes": (
    "Cloth/clothing bags; handy baskets for dry biscuits; municipality distributing "
    "cloth bags."
  ),
  "Pol_upcycling": (
    "Villages can make handy baskets for dry food — local reuse craft."
  ),
  "Pol_recycling": (
    "National separation and export to India; partial village recycling."
  ),
  "Pol_waste_collection": (
    "Khumbu municipality weekly pickup, segregation, and plane transport to "
    "Kathmandu."
  ),
  "Sol_awarness": (
    "More public information to increase awareness."
  ),
  "Sol_upcycling": (
    "Handmade handy baskets for dry food in villages."
  ),
  "Sol_recycling": (
    "More recycling and reuse options needed beyond current partial system."
  ),
  "Sol_education": (
    "Public information campaigns; tourist-to-local knowledge transfer."
  ),
  "Sol_tax": (
    "Littering fines for locals and tourists."
  ),
  "Sol_subsitutes": (
    "Cloth/clothing bags and handy baskets for packaged dry foods."
  ),
  "Sol_enforcement": (
    "Universal littering fines — locals and tourists equally."
  ),
  "Traditions to build on (free-hand)": (
    "Clothing/cloth bags; village-made handy baskets for biscuits/dry food; "
    "Khumbu municipality cloth-bag distribution (4–5 per house)."
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
    "Interview: NPL_32  |  Country: NEPAL  |  Furinji Sherpa (trekking guide)  |  "
    "Actor: private_sector  |  19 Feb 2026"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_32"
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
