"""Generate interview coding spreadsheet for Nepal interview NPL_1.

Coding follows the Overview All Interviews codebook and is based on the
Environmental Department interview guideline (June 2023).
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_1.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

STAKEHOLDER = {
  "Interview ID": "NPL_1",
  "Country": "NEPAL",
  "Coded actor type": "governmental",
  "Likely affiliation": "Department of Environment / Ministry of Forests and Environment (federal)",
  "Role in plastic governance": "National policy design, monitoring of producers/markets, awareness campaigns, SWM Act revision",
  "Perspective": "Insider — describes measures as 'we' (government actor implementing national plastic policy)",
  "Confidence": "High for governmental coding; see RSCT note below",
  "RSCT note": (
    "Page 1 contains RSCT cooperative profile notes (est. 1991, 200+ cooperatives, "
    "grassroots loans, recently urban plastics). This does not match the interview "
    "voice in Themes A–D, which is clearly a federal ministry/department official. "
    "RSCT notes are treated as misplaced interviewer notes, not the coded stakeholder."
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_1",
  "Actortype": "governmental",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "low",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "yes",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "no",
  "Problem_import": "yes",
  "Impacts": "water, cities, biodiversity, health, agriculture",
  "Governance": "yes",
  "Relevance_international_pol": "no",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "no",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "no",
  "Local government": "yes",
  "Students": "no",
  "Private_sector": "yes",
  "Civil_society": "yes",
  "Science": "yes",
  "Households": "yes",
  "Education institutions": "no",
  "Private_companies(hotels, shops, etc.)": "yes",
  "Implementation issues": "yes",
  "Monitoring": "yes",
  "Financial_resources": "yes",
  "Research": "yes",
  "Infrastructure": "no",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "no",
  "Pol_import": "yes",
  "Pol_awarness": "yes",
  "Pol_education": "no",
  "Pol_capacity": "no",
  "Pol_RD": "no",
  "Pol_tax": "no",
  "Pol_ban": "yes",
  "Pol_subsitutes": "yes",
  "Pol_clean_up": "no",
  "Pol_upcycling": "no",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "yes",
  "Sol_responsibilities": "yes",
  "Sol_epr": "yes",
  "Sol_awarness": "yes",
  "Sol_segregation": "no",
  "Sol_upcycling": "no",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "no",
  "Sol_RD": "yes",
  "Sol_tax": "no",
  "Sol_ban": "no",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "no",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": "NA",
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
    "Federal governmental actor (Department of Environment). Uses first-person "
    "'we' for policy implementation, ministry-only questions (4.a/4.b), and "
    "describes Chief Secretary committee oversight."
  ),
  "Problem_awareness_pop": (
    "Public is ignorant; people know plastic is a problem but not the full extent."
  ),
  "Problem_awareness_pol": (
    "Low political priority among decision-makers: 'no political will', "
    "plastics/waste 'not a priority' and 'stops at major' level."
  ),
  "Problem_severity": (
    "Serious issue across rivers, lakes, land, forests; microplastics in Himalayan snow."
  ),
  "Problem_littering": "Blocking sewer systems; leakage into waterways; landfill pressure.",
  "Problem_consumption": "Lots of plastics are produced.",
  "Problem_recycling": "Inadequate waste management; limited economically viable recycling for single-use items.",
  "Problem_waste_mgmt": "Waste management for plastics is inadequate.",
  "Problem_production": "High production rates mentioned.",
  "Problem_alternatives": "Biodegradable alternatives limited and costly (~100 NPR/kg more).",
  "Problem_import": (
    "Import is part of the problem space: ban on production, import, and use of "
    "bags under 40 microns explicitly targets imported plastics."
  ),
  "Impacts": (
    "Water (rivers/lakes), cities (sewers/urban), biodiversity (forests), "
    "health (burning/indoor air), agriculture (research gap noted for Chitwan)."
  ),
  "Coordination_sectoral": "Stakeholders are not collaborating or coordinating.",
  "Coordination_levels": (
    "Challenges differ across national, district, and municipal levels; "
    "SWM Act revision needed for federal governance alignment."
  ),
  "Federal government": (
    "Department of Environment, Chief Secretary committee, Department of Commerce and Supplies."
  ),
  "Local government": "Local governments lead waste management and collection.",
  "Private_sector": "Private companies and contractors in collection/recycling.",
  "Civil_society": "NGOs operate recovery facilities and awareness campaigns.",
  "Science": (
    "Cites studies on microplastics; notes limited research on health/environmental effects."
  ),
  "Households": (
    "Household waste reduction discussed; 300 NPR/month collector incentive mentioned "
    "(household incentive, not a product tax)."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Private contractors, plastic producers, and markets monitored by Department of Environment."
  ),
  "Monitoring": "Implementation is always the problem; monitoring cited as barrier.",
  "Financial_resources": (
    "Funds exist but are not released; earmarked funds may be frozen if unused."
  ),
  "Research": "Limited knowledge and research on health/environmental effects.",
  "Infrastructure": (
    "Facilities exist in at least some contexts ('facility is there'); not coded as a "
    "primary lack-of-infrastructure barrier."
  ),
  "Enforcement": "Weak enforcement; cheaper plastics remain in use despite bans.",
  "Pol_import": "Ban on production, import, and use of plastic bags under 40 microns.",
  "Pol_awarness": "Government documentaries, traditional/social media, website materials.",
  "Pol_tax": (
    "No levy/tax on specific products coded. The 300 NPR/month measure is a "
    "household waste-reduction incentive for collectors, not a bag tax/levy."
  ),
  "Pol_ban": "Ban on bags under 40 microns and plastic decorative flowers.",
  "Pol_subsitutes": "Biodegradable starch bags discussed as existing but costly alternative.",
  "Pol_recycling": "Material recovery facilities send recyclables to processing units.",
  "Pol_waste_collection": "Collection by local governments or private contractors.",
  "Sol_lead_agency": "High-level committee chaired by Chief Secretary oversees measures.",
  "Sol_responsibilities": "EPR framed as key missing responsibility mechanism.",
  "Sol_epr": "EPR identified as key gap; no producer responsibility in Nepal.",
  "Sol_awarness": "Need for communication and awareness on how to act.",
  "Sol_segregation": "Collection/recycling encouraged, but segregation not proposed as a forward solution.",
  "Sol_RD": "R&D needed for affordable alternatives.",
  "Sol_ban": (
    "Bans described as existing policy; forward-looking solutions emphasize EPR, "
    "R&D, communication, and mitigation/adaptation rather than new bans."
  ),
  "Sol_finance": "Financial support identified in Q10.",
  "Sol_infrastructure": "Infrastructural support identified in Q10.",
  "Sol_subsitutes": "Affordable eco-friendly substitutes needed before behaviour change.",
  "Sol_enforcement": "Enforcement procedures needed alongside bans.",
  "Sol_monitoring": "Monitoring improvements implied from implementation challenges.",
  "Traditions to build on (free-hand)": "Follow-up asked (Q8) but no specific tradition recorded.",
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
    "Interview: NPL_1  |  Country: NEPAL  |  Actor: Federal governmental "
    "(Department of Environment / Ministry)  |  "
    "Source: Interview Guideline — Environmental Department (June 2023)"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_1"
  stakeholder["A1"].fill = hfill(DARK_BLUE)
  stakeholder["A1"].font = hfont(bold=True, size=12, color=WHITE)
  stakeholder["A1"].alignment = CENTER_ALIGN
  for idx, (key, value) in enumerate(STAKEHOLDER.items(), start=3):
    stakeholder.cell(row=idx, column=1, value=key).font = hfont(bold=True)
    cell = stakeholder.cell(row=idx, column=2, value=value)
    cell.alignment = WRAP_ALIGN
  stakeholder.column_dimensions["A"].width = 28
  stakeholder.column_dimensions["B"].width = 90

  wb.save(OUTPUT_PATH)
  print(f"Saved {OUTPUT_PATH}")


if __name__ == "__main__":
  build_workbook()
