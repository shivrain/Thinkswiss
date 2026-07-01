"""Generate interview coding spreadsheet for Nepal interview NPL_8.

Coding follows the Overview All Interviews codebook and is based on the
Ministry of Urban Development (MOUD) interview of 23 June 2025 — transcript,
interview notes, and guideline. Participants: Kamal Adhikar (Senior Sociologist),
Nawrag (Joint Secretary), Senior Division Engineer. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_8.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "MOUD — 23 June 2025 (10:55am – 12:00pm)"),
  ("Interview notes", "Interview no. 2 — Ministry of Urban Development guideline notes"),
  ("Interview guideline", "Ministry of Urban Development — 23 June 2025"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_8",
  "Country": "NEPAL",
  "Interview number": "2 (MOUD / Kathmandu series)",
  "Interview date": "23 June 2025 (10:55am – 12:00pm)",
  "Interviewee": (
    "Panel: Kamal Adhikar (Senior Sociologist); Nawrag (Joint Secretary, "
    "29 years in service); Senior Division Engineer (DUDBC)"
  ),
  "Affiliation / role": (
    "Ministry of Urban Development (MOUD) — national policy formulation for "
    "solid waste management; draft SWM Act revision; PPP and EPR provisions"
  ),
  "Organization": (
    "Ministry of Urban Development; Department of Urban Development and "
    "Building Construction (DUDBC) training centre"
  ),
  "Stakeholder list mapping": "national gov",
  "Coded actor type (codebook)": "governmental",
  "Actor classification rationale": (
    "Central/national government ministry with cabinet-delegated authority to "
    "develop policies, acts, and regulations for solid waste management; "
    "supports but does not directly implement municipal waste operations."
  ),
  "Perspective": (
    "National policy architect — confident in draft SWM Act; emphasizes "
    "source segregation, circular economy, three-tier coordination, PPP/FDI, "
    "and future separate plastics regulatory framework"
  ),
  "Project context": (
    "Draft Solid Waste Management Act (post-2015 federal constitution); "
    "Banchare Danda (Bonsai centre) landfill filled in 3–4 years (designed 20); "
    "Bagmati cleaning campaign ~10 years; 10-micron ban, plan for 100 microns"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_8",
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
    "water, rivers, cities, health, tourism, agriculture, microplastics, "
    "visual pollution, wildlife, dumping sites, foul smell, drainage blockage"
  ),
  "Governance": "yes",
  "Relevance_international_pol": "yes",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "yes",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "yes",
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
  "Financial_resources": "no",
  "Research": "yes",
  "Infrastructure": "no",
  "Enforcement": "yes",
  "Policies in place": "yes",
  "Pol_epr": "yes",
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
  "Sol_epr": "yes",
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
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Leaf plates instead of plastic; pet-to-khet / khet-to-pet sustainable "
    "consumption–agriculture cycle (broken by plastic); customary organic "
    "waste practices when population density was low; emerging home "
    "dustbin segregation (decomposable vs non-decomposable)"
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
    "MOUD panel (Joint Secretary, Senior Sociologist, Senior Division Engineer). "
    "Stakeholder list: national gov. Codebook: governmental."
  ),
  "Problem_awareness_pop": (
    "Public perception critical — people find plastic convenient, don't care about "
    "environmental/public health implications; throw mixed waste in streets. New "
    "generation more concerned; home segregation emerging but habits take 20–25 years."
  ),
  "Problem_awareness_pol": (
    "MOUD actively drafting new SWM Act with stakeholder consultation; confident it "
    "will address challenges; plans separate plastics regulatory framework."
  ),
  "Problem_severity": (
    "Plastic majority of urban solid waste; disturbs rivers, canals, sewer lines; "
    "nuisance to old cities and emerging towns; Bagmati campaign running ~10 years."
  ),
  "Problem_littering": (
    "Plastic thrown everywhere indiscriminately; waste thrown in streets, river banks, "
    "from doors/windows/rooftops."
  ),
  "Problem_consumption": (
    "People's lives closely attached to plastic for generations; convenient carrying "
    "of goods in plastic bags."
  ),
  "Problem_recycling": (
    "Circular economy promoted but blocked without source segregation; recycling "
    "provisions in draft act."
  ),
  "Problem_waste_mgmt": (
    "Solid waste management in totality is the issue; sanitary landfills operated "
    "as dumping sites; Banchare Danda filled in 3–4 years (designed for 20)."
  ),
  "Problem_production": (
    "Plastic is majority volume in urban waste composition; whole manufacturing "
    "sector (packaging + plastic producers) implicated."
  ),
  "Problem_alternatives": (
    "No specific plastic-waste policy yet; separate framework planned; leaf plates "
    "and traditional practices cited as past alternatives."
  ),
  "Problem_waste_segregation": (
    "Main challenge — no culture of source segregation; mixed composite waste "
    "makes municipal management very difficult."
  ),
  "Impacts": (
    "Water/rivers: disturbs river and canal flow, blocks water lines. Cities: "
    "nuisance in old and emerging urban areas. Health: meat/blood/yogurt in black "
    "plastic without quality awareness; leachate and flies at dumpsites. Tourism: "
    "aesthetic/visual pollution. Agriculture: productivity decreasing. "
    "Microplastics: indirect via livestock eating plastic. Wildlife: cows and "
    "animals eat plastics. Dumping sites: landfills converted to open dumps. "
    "Foul smell: stinking at dumping sites. Drainage blockage: sewer and water "
    "line blockages."
  ),
  "Relevance_international_pol": (
    "SDG framework and human rights framework integrated into act formulation; "
    "Human Rights Commission reporting requirements cited."
  ),
  "Coordination_sectoral": (
    "Wide stakeholder consultation (universities, private sector, researchers, "
    "mayors); war between people and policymakers on responsibility."
  ),
  "Coordination_levels": (
    "New act defines central, provincial, and local roles; provincial/central "
    "to help LG acquire landfill land; inter- and intra-municipal waste conflicts."
  ),
  "Unclear_responsibilities": (
    "People think municipality responsible, government thinks producer "
    "responsible; LG confused on private-sector onboarding, land procurement, "
    "donor coordination."
  ),
  "Federal government": (
    "MOUD develops policies, acts, regulations; coordinates international donors; "
    "cabinet-delegated authority."
  ),
  "Provincial government": (
    "Will help local government acquire central/regional landfill land when LG "
    "faces difficulties."
  ),
  "Local government": (
    "Full authority under Local Self-Governance Act 2016 to manage community "
    "solid waste; must follow central acts."
  ),
  "Private_sector": (
    "PPP model, FDI, sample municipality–private-sector agreement documents; "
    "domestic and international private waste operators."
  ),
  "Civil_society": (
    "NGOs interested in plastic waste management policy; welcomed for expertise "
    "during policy formulation."
  ),
  "Science": (
    "University people and researchers consulted in act drafting; lack of expert "
    "ideas domestically — welcome international expertise."
  ),
  "Households": (
    "Source segregation habits must develop; emerging culture of separate home "
    "dustbins (decomposable vs non-decomposable)."
  ),
  "Education institutions": (
    "Universities consulted in stakeholder discussions for act formulation."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Manufacturing industries (biscuits, noodles, packaging) and plastic producers; "
    "commercial activities generating byproduct waste."
  ),
  "Monitoring": (
    "No monitoring technology to analyze segregation patterns; planned municipal "
    "software to track waste generation, segregation, transport, disposal."
  ),
  "Research": (
    "Lack of expert ideas domestically; welcome international expertise for "
    "circular economy and plastic policy."
  ),
  "Infrastructure": (
    "Transfer stations and sanitary landfill sites exist but are misoperated as "
    "dumping sites — operational failure, not absence of facilities."
  ),
  "Enforcement": (
    "Polluter's pay principle weak in implementation; micron ban impact not "
    "assessed; strict burning prohibition proposed in new act."
  ),
  "Policies in place": (
    "General SWM policy (not plastic-specific); plastic classified as hazardous "
    "waste; 10-micron ban; draft SWM Act with EPR, PPP, burning ban."
  ),
  "Pol_epr": (
    "EPR (EPI) proposed in draft act — producers responsible for disposal; "
    "polluter's pay principle referenced."
  ),
  "Pol_awarness": "Bagmati/Bhagwati river cleaning campaign (~10 years).",
  "Pol_education": (
    "Province/district orientation and dissemination sessions planned; DUDBC "
    "training centre to develop orientation packages."
  ),
  "Pol_capacity": (
    "DUDBC training centre; technical input to municipalities; capacity "
    "development for elected leaders lacking technical know-how."
  ),
  "Pol_RD": (
    "Circular economy promotion; foreign technology investment; separate "
    "plastic policy and guidelines planned."
  ),
  "Pol_ban": (
    "10-micron plastic ban in force; plan to increase to 100 microns; strict "
    "burning prohibition in draft act."
  ),
  "Pol_subsitutes": "Traditional leaf plates cited as pre-plastic alternative.",
  "Pol_clean_up": "Bagmati cleaning campaign ongoing for ~10 years.",
  "Pol_recycling": "Segregation and recycling provisions in draft act.",
  "Pol_waste_collection": (
    "Act covers collection, transportation, and disposal; private sector entry "
    "via PPP."
  ),
  "Sol_lead_agency": (
    "MOUD as central coordinating ministry; cabinet submission and donor "
    "coordination role."
  ),
  "Sol_responsibilities": (
    "Clarify central, provincial, local roles in new act; regional sanitary "
    "landfill sites covering multiple municipalities."
  ),
  "Sol_epr": (
    "Extended producer responsibility in draft act; whole business chain "
    "(packaging + plastic producers) to co-create with ministry."
  ),
  "Sol_awarness": (
    "Province/district consultations, dissemination, orientation; not just "
    "website posting or hard copies."
  ),
  "Sol_segregation": (
    "Source segregation as primary solution — enables circular economy and "
    "reduces waste volume."
  ),
  "Sol_recycling": "Circular economy as core approach once segregation achieved.",
  "Sol_education": (
    "Training/orientation packages via DUDBC; municipal official engagement."
  ),
  "Sol_capacity": (
    "Technical backing for elected municipal leaders; DUDBC training centre; "
    "guidelines and regulations post-act."
  ),
  "Sol_RD": (
    "Separate plastic waste policy and guidelines planned after SWM Act; "
    "welcome NGO/researcher input."
  ),
  "Sol_ban": (
    "Strict burning prohibition with punishment; micron ban increase to 100."
  ),
  "Sol_finance": (
    "PPP model; FDI for waste-management technology; sample agreement documents."
  ),
  "Sol_infrastructure": (
    "Regional sanitary landfill sites; transfer stations; provincial/central "
    "support for land acquisition."
  ),
  "Sol_subsitutes": (
    "Revive pet-to-khet cycle; leaf plates; reduce waste through segregation."
  ),
  "Sol_enforcement": (
    "Burning prohibition with penalties; EPR fees if producer cannot manage waste."
  ),
  "Sol_monitoring": (
    "Planned municipal software tracking population, production rate, segregation, "
    "transport, disposal, and landfill status."
  ),
  "Traditions to build on (free-hand)": (
    "Leaf plates instead of plastic; pet-to-khet / khet-to-pet sustainable cycle; "
    "customary organic waste disposal (not 'bad habits'); emerging home dustbin "
    "segregation culture."
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
    "Interview: NPL_8  |  Country: NEPAL  |  MOUD panel  |  "
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_8"
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
