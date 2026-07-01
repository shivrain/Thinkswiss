"""Generate interview coding spreadsheet for Nepal interview NPL_7.

Coding follows the Overview All Interviews codebook and is based on the
Narayan Nirola (RSDC — Rural Self-Reliance Development Centre) interview of
23 June 2025 — transcript, interview notes, and guideline. Uses expanded
Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_7.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "RSDC — Narayan Nirola, 23 June 2025"),
  ("Interview notes", "Interview no. 4 — RSDC structured guideline notes"),
  ("Interview guideline", "Rural Self-Reliance Development Center — 23 June 2025"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_7",
  "Country": "NEPAL",
  "Interview number": "4 (RSDC / Kathmandu series)",
  "Interview date": "23 June 2025 (12:52pm – 1:30pm)",
  "Interviewee": "Narayan Nirola",
  "Affiliation / role": (
    "Project coordinator — Rural Self-Reliance Development Centre (RSDC); "
    "WASH specialist; former environmental-science lecturer"
  ),
  "Organization": (
    "RSDC (est. 1991) — NGO focused on poverty alleviation, cooperatives, "
    "WASH, and pilot plastic/waste-management project in Budhanilkantha"
  ),
  "Stakeholder list mapping": "cso",
  "Coded actor type (codebook)": "cso",
  "Actor classification rationale": (
    "Civil society organization (NGO) implementing community-based waste "
    "management through cooperatives and household segregation in partnership "
    "with local government — not government, private company, or household."
  ),
  "Perspective": (
    "CSO implementer-advocate — emphasizes multi-stakeholder coordination, "
    "evidence-based localized research, EPR, and political will at municipal level"
  ),
  "Project context": (
    "Budhanilkantha municipality; 100+ households; cooperative waste collection; "
    "300 NPR/month private collection fee; segregation incentive policy not implemented"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_7",
  "Actortype": "cso",
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
    "water, rivers, cities, health, air pollution, microplastics, dumping sites, "
    "aquatic life, drainage blockage"
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
  "Infrastructure": "no",
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
  "Pol_clean_up": "no",
  "Pol_upcycling": "no",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "yes",
  "Sol_responsibilities": "yes",
  "Sol_epr": "yes",
  "Sol_awarness": "yes",
  "Sol_segregation": "yes",
  "Sol_upcycling": "no",
  "Sol_recycling": "yes",
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
    "Rural practice of reusing plastic bags many times; indigenous practice of "
    "bringing own cup to gatherings (long-term revival goal); cooperative "
    "self-reliance (Swabalamban) community fund model"
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
    "Narayan Nirola, RSDC project coordinator (CSO/NGO). Stakeholder list: cso. "
    "Organization: Rural Self-Reliance Development Centre."
  ),
  "Problem_awareness_pop": (
    "Superficial awareness — know plastic is bad but not extent; reuse PET bottles; "
    "burn plastic if far from house; no consistent behaviour change."
  ),
  "Problem_awareness_pol": (
    "Mayor lacks priority/internalization; plastic not priority vs other issues; "
    "ward officials supportive but municipal head not engaged."
  ),
  "Problem_severity": (
    "High per capita plastic production (ADB 2013); landfill half-filled in 3 years "
    "(designed for 20); huge problem in growing cities."
  ),
  "Problem_littering": (
    "Rivers in Kathmandu and other cities used as dumping sites; plastic everywhere "
    "in urban areas."
  ),
  "Problem_consumption": (
    "High per capita production; plastic pervasive in metropolitan areas; urban sprawl."
  ),
  "Problem_recycling": (
    "Households recover/sell recyclables via cooperative but earnings minimal; most "
    "plastic still goes to dump."
  ),
  "Problem_waste_mgmt": (
    "Stakeholders not working in coordinated integrated manner; implementation always "
    "the gap despite policies existing."
  ),
  "Problem_production": "ADB 2013: high per capita plastic production in Nepal.",
  "Problem_alternatives": (
    "R&D needed for plant-based plastic bags (~20–30 NPR/kg more); cannot eliminate "
    "plastic from daily life."
  ),
  "Problem_waste_segregation": (
    "People mix all waste for morning garbage truck; LG won't reduce 300 NPR fee for "
    "segregating households despite own policy."
  ),
  "Impacts": (
    "Water/rivers: freshwater bodies polluted, rivers as dumping sites. Cities: sewer "
    "systems clogged, urban pervasive pollution. Health: indoor air pollution from "
    "burning MLP, PET bottle misuse. Air pollution: burning plastic in rural/urban areas. "
    "Microplastics: Chitwan lake, salt (research cited). Dumping sites: landfill "
    "Sisdol half-filled in 3 years. Aquatic life: softeners, congestion, fertility "
    "impacts expected. Drainage blockage: plastic clogs sewer systems."
  ),
  "Coordination_sectoral": (
    "All stakeholders (government, private collectors, cooperatives, producers, public) "
    "not working in coordinated way."
  ),
  "Coordination_levels": (
    "Federal budget frozen when LG doesn't spend; ward level supportive but mayor "
    "disengaged; earmarked vs general fund issues."
  ),
  "Unclear_responsibilities": (
    "LG doesn't have list of waste-management organizations; unlicensed private "
    "collectors operate with LG knowledge but no action."
  ),
  "Federal government": (
    "National policy frameworks; budget allocation to LG; SWM Act 2011; DoE occasional "
    "monitoring."
  ),
  "Local government": (
    "Budhanilkantha municipality; ward chairperson vs mayor disconnect; 300 NPR "
    "collection policy not implemented."
  ),
  "Private_sector": (
    "Private waste companies dominate collection (Sawman/Neximac umbrella); influence "
    "LG through money; compete for contracts."
  ),
  "Civil_society": (
    "RSDC, cooperatives (200+ promoted), community groups, segregation points."
  ),
  "Science": (
    "Respondent former environmental-science lecturer; advocates localized "
    "evidence-based research; few Nepal-specific health/environment studies."
  ),
  "Households": (
    "100+ households in project area; segregation and cooperative collection model."
  ),
  "Education institutions": (
    "Respondent academic background; education campaigns for all stakeholders advocated."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Plastic producers and brand owners influence policy; unregistered waste collectors; "
    "bottled water/pasta producers under EPR."
  ),
  "Monitoring": (
    "20–40 micron ban not regularly monitored; DoE monitors only occasionally via Facebook "
    "posts."
  ),
  "Financial_resources": (
    "Environment budget earmarked but frozen yearly; LG has money but doesn't use it; "
    "not insufficient but unspent."
  ),
  "Research": (
    "Very limited Nepal-specific research on plastic health/environment impacts; need "
    "localized quantified studies."
  ),
  "Infrastructure": (
    "Segregation points and vendors exist locally; facilities available but NIMBY and "
    "lack of motivation — not primary infrastructure gap per respondents."
  ),
  "Enforcement": (
    "SWM Act requires registered/licensed collectors but none registered in project area; "
    "micron ban unenforced."
  ),
  "Policies in place": (
    "20–40 micron bag regulation; household segregation fee-reduction policy; SWM Act "
    "2011; EPR draft (not enacted)."
  ),
  "Pol_epr": "Draft EPR policy exists but not yet enacted.",
  "Pol_awarness": "NGO/project-based awareness; exposure visits to change official mindsets.",
  "Pol_education": "Awareness and education for government, private sector, and communities.",
  "Pol_capacity": "Exposure visits, capacity building for enforcement and regulation.",
  "Pol_RD": "R&D for plant-based plastic alternatives advocated.",
  "Pol_ban": "Regulation on plastic bag thickness (minimum 20–40 microns).",
  "Pol_subsitutes": "Plant-based plastic bags as alternative (20–30 NPR/kg premium).",
  "Pol_recycling": "Cooperative collects segregated waste and sells to vendors.",
  "Pol_waste_collection": (
    "Private companies collect at 300 NPR/month/household; cooperative parallel collection."
  ),
  "Sol_lead_agency": (
    "Multi-stakeholder engagement with localized evidence base — no single actor alone."
  ),
  "Sol_responsibilities": (
    "EPR for producers; register/license private collectors; LG implement own policies."
  ),
  "Sol_epr": "EPR essential — hold producers/importers/brand owners accountable.",
  "Sol_awarness": (
    "Awareness, education, and exposure visits; communication with evidence for decision-makers."
  ),
  "Sol_segregation": (
    "Household source segregation with fee-reduction incentives; cooperative collection."
  ),
  "Sol_recycling": (
    "Mitigation (reduce/minimize) + adaptation (proper waste management); waste-to-energy."
  ),
  "Sol_education": "Educate all stakeholders including bureaucrats and communities.",
  "Sol_capacity": (
    "Enforcement capacity; regulate private collectors; motivate environmental staff."
  ),
  "Sol_RD": (
    "Localized research; plant-based alternatives; cost-benefit analysis for substitutes."
  ),
  "Sol_ban": (
    "Cannot aim for zero plastic production; minimize use and manage end-of-life instead."
  ),
  "Sol_finance": (
    "Use existing frozen environment budgets; reduced waste fees for segregating households; "
    "support cooperatives."
  ),
  "Sol_infrastructure": (
    "Waste-to-energy with minimal pollution; overcome NIMBY for collection points."
  ),
  "Sol_subsitutes": "Promote plant-based plastic bags and affordable alternatives.",
  "Sol_enforcement": (
    "License/register private waste collectors; regular micron-ban monitoring."
  ),
  "Sol_monitoring": "Regular government monitoring of plastic bag thickness and collectors.",
  "Traditions to build on (free-hand)": (
    "Rural reuse of plastic bags multiple times; cooperative self-reliance (Swabalamban) "
    "model; reviving indigenous practice of bringing own cup to gatherings."
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
    "Interview: NPL_7  |  Country: NEPAL  |  RSDC — Narayan Nirola  |  "
    "Actor: cso  |  23 June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_7"
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
