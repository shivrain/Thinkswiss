"""Generate interview coding spreadsheet for Nepal interview NPL_14.

Coding follows the Overview All Interviews codebook and is based on the
Rajesh Aryal (Silent Park / Chitwan tourism) interview of 26 June 2025 —
transcript, field notes, and interview guideline. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_14.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "Silent Park — Rajesh Aryal, Ram Devi / Deep, 26 June 2025"),
  ("Interview notes", "Interview Silent Park field notes — Rajesh Aryal"),
  ("Interview guideline", "Plastic governance guideline (Themes A–D; policy-science sections)"),
  (
    "Coding note",
    "Guideline includes national-policy content (EPR, Bio-Camp, Plast Foundation) "
    "not spoken in this transcript; codes verified primarily against transcript and "
    "Silent Park field notes.",
  ),
]

STAKEHOLDER = {
  "Interview ID": "NPL_14",
  "Country": "NEPAL",
  "Interview number": "Silent Park / Chitwan tourism series",
  "Interview date": "26 June 2025",
  "Interviewee": "Rajesh Aryal",
  "Affiliation / role": (
    "Tourism/hospitality sector representative — Chitwan hotel area near "
    "national park; familiar with safari operations, nature guides, and "
    "municipal waste collection"
  ),
  "Organization": (
    "Silent Park area; Chitwan hotel/tourism cluster (~100 hotels); "
    "hotel association and nature-guide training networks"
  ),
  "Stakeholder list mapping": "hotel_owner",
  "Coded actor type (codebook)": "hotel_owner",
  "Actor classification rationale": (
    "Hotel/tourism-sector respondent discussing hotel waste, safari dustbins, "
    "hotel-association programmes, and tourist-area cleanliness — not "
    "government, recycler, or household."
  ),
  "Perspective": (
    "Local tourism practitioner — plastic pollution acknowledged but aluminum "
    "foil seen as greater disposal problem; emphasizes irregular collection, "
    "nature-guide training, and ward-level awareness campaigns"
  ),
  "Project context": (
    "Eco-Green collects PET bottles every 15–20 days (irregular); municipal "
    "trucks for wrappers/bags; bottles NPR 10–15/kg; safari vehicles require "
    "in-vehicle dustbins; elephant-dung cleaning twice daily as model"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_14",
  "Actortype": "hotel_owner",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "medium",
  "Problem_severity": "medium",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "yes",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "yes",
  "Problem_import": "no",
  "Impacts": (
    "wildlife, air pollution, visual pollution, health"
  ),
  "Governance": "yes",
  "Relevance_international_pol": "no",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "yes",
  "Actors": "yes",
  "Federal government": "no",
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
  "Research": "no",
  "Infrastructure": "yes",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "no",
  "Pol_import": "no",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
  "Pol_capacity": "yes",
  "Pol_RD": "no",
  "Pol_tax": "no",
  "Pol_ban": "no",
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
  "Sol_RD": "no",
  "Sol_tax": "no",
  "Sol_ban": "no",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "yes",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "no",
  "Traditions to build on (free-hand)": (
    "Cloth or fiber bags instead of plastic; clay or cement pots instead of "
    "plastic flower pots; safari in-vehicle waste buckets; nature-guide no-litter "
    "training in national park"
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
    "Rajesh Aryal — Chitwan/Silent Park tourism-hotel context. Stakeholder list: "
    "hotel_owner. Codebook: hotel_owner."
  ),
  "Problem_awareness_pop": (
    "Ignorance and laziness; free plastic bags at shops; convenience over cloth "
    "bags; people sometimes follow if one person models behaviour."
  ),
  "Problem_awareness_pol": (
    "No municipal notices on plastic removal; awareness/clean-up programmes "
    "irregular — World Environment Day only, no continuity."
  ),
  "Problem_severity": (
    "Plastic pollution everywhere; respondent more concerned about aluminum foil "
    "food wrapping than plastic; burning has serious air impact; production and "
    "usage problems acknowledged."
  ),
  "Problem_littering": (
    "Bins overflow when collection delayed; wind and animals (dogs, wildlife) "
    "scatter waste; salty snack wrappers ingested by animals."
  ),
  "Problem_consumption": (
    "Plastic bags given free at shops; easy and fast; thousands of hotels "
    "producing waste in Chitwan."
  ),
  "Problem_recycling": (
    "Eco-Green bottle collection irregular (15–20 days, sometimes much longer); "
    "aluminum foil not taken by recyclers; collection system discontinuous."
  ),
  "Problem_waste_mgmt": (
    "Difficult to eradicate plastic; management through organizations but not "
    "continuous; hotel collective management discusses rates/security not waste."
  ),
  "Problem_production": (
    "Production happening with usage problems; pollution increases if not addressed."
  ),
  "Problem_alternatives": (
    "Cloth/fiber bags available but inconsistently used; clay/cement pots as "
    "alternative to plastic flower pots."
  ),
  "Problem_waste_segregation": (
    "Bottles stored separately for Eco-Green; wrappers/bags to municipal trucks; "
    "bottles must be emptied completely before disposal."
  ),
  "Impacts": (
    "Wildlife: domestic and wild animals ingest plastic (salty snack wrappers). "
    "Air pollution: burning plastic has serious impact. Visual pollution: "
    "overflowing dustbins spread litter. Health: smoke from burning plastics "
    "(winter home burning implied in guideline; burning confirmed in transcript)."
  ),
  "Coordination_sectoral": (
    "Hotel association, NGOs, park rangers train nature guides; but hotel "
    "business community has not formed waste-collection committee."
  ),
  "Coordination_levels": (
    "Ward should launch awareness programmes; municipal trucks vs Eco-Green "
    "private collection — fragmented system."
  ),
  "Unclear_responsibilities": (
    "No designated street cleaners except elephant-dung teams; hotel owners clean "
    "near own homes; municipal collection irregular."
  ),
  "Local government": (
    "Municipal trucks collect wrappers/bags; ward should enforce rules and "
    "provide dustbins; no municipal notices seen."
  ),
  "Private_sector": (
    "Eco-Green collects PET bottles; individuals buy bottles at NPR 10–15/kg; "
    "small organizations collect valuable materials."
  ),
  "Civil_society": (
    "NGOs, hotel association, park rangers provide nature-guide awareness training."
  ),
  "Households": (
    "Store bottles for recycling; discard wrappers; people clean near own homes."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "~100 hotels in area; hotel association training; shops give free plastic bags; "
    "hotel committee for tourist-area cleaning proposed (NPR 50,000/month rickshaw)."
  ),
  "Monitoring": (
    "No regular street cleaning monitoring; collection timing not enforced — "
    "bins overflow."
  ),
  "Financial_resources": (
    "Monetary incentive drives bottle collection; hotel committee could fund "
    "rickshaw collection (~NPR 50,000/month) but not formed."
  ),
  "Infrastructure": (
    "Dustbins provided but overflow; safari vehicles require in-vehicle buckets; "
    "roadside dustbins for guides; no designated street-cleaning staff."
  ),
  "Enforcement": (
    "No municipal notices; ward strict enforcement of rules advocated; safari "
    "dustbin rule compulsory but arrival briefing needed."
  ),
  "Policies in place": (
    "Eco-Green recycling collection; municipal truck collection; Environment Day "
    "awareness (one day only); nature-guide no-litter training."
  ),
  "Pol_awarness": (
    "Awareness programmes by NGOs, hotel association, government officials — "
    "irregular, lack continuity."
  ),
  "Pol_education": (
    "Nature guides trained by park rangers and hotel association not to litter; "
    "safari briefing on in-vehicle waste bucket."
  ),
  "Pol_capacity": (
    "Nature-guide and hotel-sector training on waste handling during safari and "
    "park walks."
  ),
  "Pol_subsitutes": "Cloth/fiber bags; clay/cement pots instead of plastic.",
  "Pol_clean_up": (
    "Environment Day clean-up; elephant-dung cleaning twice daily as operational "
    "model; irregular ward street cleaning."
  ),
  "Pol_recycling": (
    "Eco-Green PET bottle recycling; bottles sold at NPR 10–15/kg to collectors."
  ),
  "Pol_waste_collection": (
    "Municipal trucks for bags/wrappers; Eco-Green for bottles; monetary-value "
    "materials collected by individuals morning/evening."
  ),
  "Sol_responsibilities": (
    "Ward should launch awareness and designate street cleaners; hotel business "
    "committee for tourist-area waste; elephant-dung model extended to other waste."
  ),
  "Sol_awarness": (
    "Spread awareness to reduce negligence; ward-enforced rules with dustbins; "
    "inform tourists on arrival about safari waste bucket."
  ),
  "Sol_segregation": (
    "Store bottles separately for Eco-Green; empty bottles completely; wrappers "
    "via municipal trucks."
  ),
  "Sol_recycling": (
    "Regular Eco-Green collection; monetize bottles to incentivize collection."
  ),
  "Sol_education": (
    "Continuous awareness (not just Environment Day); nature-guide training "
    "scaled; hotel-owner education on waste management."
  ),
  "Sol_capacity": (
    "Designated people for street cleaning; nature guides as waste stewards "
    "during safari and park walks."
  ),
  "Sol_finance": (
    "Hotel committee NPR ~50,000/month for rickshaw waste collection; NPR "
    "10–15/kg bottle buy-back incentivizes informal collectors."
  ),
  "Sol_infrastructure": (
    "Ward-provided dustbins; in-vehicle safari buckets compulsory; daily "
    "collection at least once (vs current irregular schedule)."
  ),
  "Sol_subsitutes": (
    "Cloth/fiber bags replace plastic bags; clay/cement pots for planting."
  ),
  "Sol_clean_up": (
    "Daily street collection; hotel-community rickshaw rounds; extend elephant-dung "
    "cleaning frequency model to plastic waste."
  ),
  "Sol_enforcement": (
    "Ward strict rule enforcement; compulsory safari waste-bucket use with "
    "tourist briefing on arrival."
  ),
  "Traditions to build on (free-hand)": (
    "Cloth bags for shopping; clay/cement pots; safari in-vehicle waste buckets; "
    "nature-guide stewardship keeping park and streets clean."
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
    "Interview: NPL_14  |  Country: NEPAL  |  Rajesh Aryal (Silent Park)  |  "
    "Actor: hotel_owner  |  26 June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_14"
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
