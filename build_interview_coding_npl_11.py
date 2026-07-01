"""Generate interview coding spreadsheet for Nepal interview NPL_11.

Coding follows the Overview All Interviews codebook and is based on the
Safa Urja private waste-management company interview — transcript and
interview guideline. Operates in Ratnanagar, Khaireni, and Kalika
municipalities. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_11.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "Safa Urja — Ram Devi / Deep field interview"),
  ("Interview guideline", "Safa Urja structured guideline (Themes A–D)"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_11",
  "Country": "NEPAL",
  "Interview number": "Safa Urja / Chitwan–Khaireni fieldwork series",
  "Interview date": "June 2025 (fieldwork; exact time not recorded)",
  "Interviewee": "Safa Urja representative — private waste-management operator",
  "Affiliation / role": (
    "Private-sector waste collector, segregator, and processor — contracted "
    "operator across Ratnanagar, Khaireni, and Kalika municipalities"
  ),
  "Organization": (
    "Safa Urja — Khaireni factory; 70–80 staff; GPS-tracked fleet; German "
    "operations manager; international stakeholders (UK, US); first initiative "
    "of its kind in Nepal per respondent"
  ),
  "Stakeholder list mapping": "private_sector",
  "Coded actor type (codebook)": "private_sector",
  "Actor classification rationale": (
    "Private waste-management company collecting, segregating, and processing "
    "municipal waste under contract — pays municipality NPR 6.6M+ annually; "
    "not government, CSO, or household."
  ),
  "Perspective": (
    "Ground-level private implementer — very concerned; operates at annual loss; "
    "advocates PPP across three municipalities, coordinated regulation, fines, "
    "school-based awareness, and technology (GPS, barcodes)"
  ),
  "Project context": (
    "Ratnanagar ~21 tractor-loads/day; landfills at Udaypur community forest and "
    "Khaireni riverside; Sauraha tourism zone; fines NPR 5,000–11,000; PPP TOR "
    "planned in 5–6 months; barcode payment tracking in 3 months"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_11",
  "Actortype": "private_sector",
  "Problem_awareness_pop": "low",
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
    "agriculture, health, soil, air pollution, water, visual pollution"
  ),
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
  "Civil_society": "no",
  "Science": "no",
  "Households": "yes",
  "Education institutions": "yes",
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
  "Pol_tax": "yes",
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
  "Sol_RD": "no",
  "Sol_tax": "yes",
  "Sol_ban": "yes",
  "Sol_finance": "yes",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "yes",
  "Sol_clean_up": "no",
  "Sol_enforcement": "yes",
  "Sol_monitoring": "yes",
  "Traditions to build on (free-hand)": (
    "Paper plates and cups instead of plastic at feasts; reusable items and "
    "jute sacks/bags for purchased plastics; minimize and reuse before disposal"
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
    "Safa Urja private waste-management company. Stakeholder list: private_sector. "
    "Codebook: private_sector."
  ),
  "Problem_awareness_pop": (
    "Awareness very low — people discard carelessly, don't reuse, don't follow "
    "segregation; mentality that waste management is company's job; night dumping "
    "in bazaar/Sauraha."
  ),
  "Problem_awareness_pol": (
    "No improvements seen from state; coordinated regulation across municipality, "
    "provincial, and central government needed; government hasn't formalized system."
  ),
  "Problem_severity": (
    "Very concerned (core business); Ratnanagar very difficult to manage at current "
    "pace; 21 tractor-loads/day; monsoon Jestha–Bhadra (~4 months) critical."
  ),
  "Problem_littering": (
    "Improper disposal everywhere; night dumping outside bazaar; scavengers, "
    "jackals, dogs scatter waste."
  ),
  "Problem_consumption": (
    "Promote using less plastic in house and business; purchase wisely; minimize "
    "even if cannot stop completely."
  ),
  "Problem_recycling": (
    "Company segregates at facility — plastic, textiles, glass, food waste; "
    "1–2 quintals condensed to 8–10 kg; only small portion to landfill."
  ),
  "Problem_waste_mgmt": (
    "Long-term management needed; landfill operations challenging; terrain and "
    "weather delay collection; contract operates at annual loss."
  ),
  "Problem_alternatives": (
    "Jute sacks/bags for purchased plastics; paper plates/cups for feasts; "
    "LP plastic instead of PP; no policy to reduce plastic yet."
  ),
  "Problem_waste_segregation": (
    "Household segregation tried but inconsistent; people give mixed waste; "
    "only ~500 of 2,000 NPR households actually segregate when pushed."
  ),
  "Impacts": (
    "Agriculture: monsoon chemicals reach crops; soil pH and fertility affected; "
    "tilling brings plastic to surface. Health: winter home burning — smoke and "
    "carbon emissions. Soil: long-term chemical effects. Air pollution: burning "
    "plastics at home. Water: Khaireni landfill on riverside land between channels. "
    "Visual pollution: chaotic scattered waste in bazaar areas."
  ),
  "Coordination_sectoral": (
    "Company–municipality–community communication essential; temple/school "
    "representatives don't understand systematic waste management."
  ),
  "Coordination_levels": (
    "Regulation needed from municipality, provincial, and central government "
    "together; PPP across three municipalities planned."
  ),
  "Unclear_responsibilities": (
    "Public assumes company responsibility and free collection; leadership vs "
    "household participation both needed."
  ),
  "Federal government": (
    "Central regulation with provincial and municipal coordination advocated; "
    "respondent unaware of new federal waste act."
  ),
  "Provincial government": "Mentioned as needed partner in coordinated regulation.",
  "Local government": (
    "Ratnanagar, Khaireni, Kalika municipalities; mayor/deputy mayor meetings; "
    "company pays NPR 6.6M+ annual tax; fines NPR 5,000–11,000 in Ratnanagar."
  ),
  "Students": (
    "School awareness every 3 months — children influence families; easier to "
    "convince students than parents directly."
  ),
  "Private_sector": (
    "Safa Urja itself; CREASION plastic pellets mentioned by interviewer; "
    "comparison with Doko Recyclers Kathmandu."
  ),
  "Households": (
    "NPR 25/month household fee; late payment up to 7 months; segregation "
    "compliance inconsistent."
  ),
  "Education institutions": (
    "School programs every 3 months; 4–5 schools per ward; 150+ public spots "
    "across 16 wards needing waste management."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Sauraha tourism hotels (NPR 300/month); hardware, hostels, barbers, labs, "
    "colleges tiered fees; businesses resist fee increases."
  ),
  "Monitoring": (
    "GPS-enabled vehicles; barcode payment/household-visit tracking in 3 months; "
    "staff attendance from home; partial fine enforcement only in some areas."
  ),
  "Financial_resources": (
    "Operating at loss every year under contract; rates frozen 7 years (only 10% "
    "annual increase); fee-increase proposals not passing municipal council."
  ),
  "Infrastructure": (
    "Proper landfill site in good location needed; Udaypur community forest and "
    "Khaireni riverside sites; landfill management challenging."
  ),
  "Enforcement": (
    "Fines NPR 5,000–11,000 exist but cannot enforce everywhere — strong "
    "resistance would halt work; penalties must come officially from municipality."
  ),
  "Policies in place": (
    "Municipal waste-collection fee schedule; progressive littering fines; "
    "40-micron bag rule (thicker bags cost more); no comprehensive plastic policy."
  ),
  "Pol_awarness": (
    "Community awareness programs; Sauraha tourism zone every 5–6 months with "
    "women's involvement focus."
  ),
  "Pol_education": "Quarterly school awareness programs discussed with chief officer.",
  "Pol_capacity": (
    "70–80 staff; German operations manager; technology systems (GPS, barcodes); "
    "worker responsibility training needed."
  ),
  "Pol_tax": (
    "Tiered household/business fees (NPR 25–300); company pays municipality "
    "NPR 6.6M rising 10%/year (NPR 7.266M current year)."
  ),
  "Pol_ban": (
    "Below 40-micron bags not allowed; thicker bags higher cost discourages use."
  ),
  "Pol_subsitutes": (
    "Jute sacks for purchased plastics; plan to replace feast plastic plates/cups "
    "with paper."
  ),
  "Pol_clean_up": (
    "Daily collection in bazaar/Sauraha; company cleans night-dumped waste."
  ),
  "Pol_recycling": (
    "Facility segregation — hazardous (red), food (green), less harmful (yellow); "
    "compostable/recyclable processing."
  ),
  "Pol_waste_collection": (
    "Daily vehicles in tourism/bazaar areas; monsoon delays; 21 tractor-loads/day "
    "Ratnanagar."
  ),
  "Sol_lead_agency": (
    "PPP model across Ratnanagar, Khaireni, Kalika — TOR and consultant study "
    "in 5–6 months."
  ),
  "Sol_responsibilities": (
    "Municipality must take responsibility all areas; company follows; community "
    "supports; leadership and households both critical."
  ),
  "Sol_awarness": (
    "Public awareness programs; positive promotion of waste management; school "
    "and tourism-zone campaigns."
  ),
  "Sol_segregation": (
    "Household source segregation with color-coded baskets; company facility "
    "sorting; 1–2 quintals to 8–10 kg volume reduction."
  ),
  "Sol_recycling": (
    "Reusable items recovered; segregated plastics, textiles, glass processed; "
    "minimal landfill."
  ),
  "Sol_education": (
    "Quarterly school programs — children convince parents; tourism-zone business "
    "education."
  ),
  "Sol_capacity": (
    "Workers must understand duties; 70–80 staff; international operations expertise."
  ),
  "Sol_tax": (
    "Late-payment penalties (must come from municipality); differential fees for "
    "segregated vs unsegregated proposed."
  ),
  "Sol_ban": (
    "Law banning plastic plates/cups at feasts; coordinated top-down regulation."
  ),
  "Sol_finance": (
    "PPP to address unsustainable loss-making contract; municipality proposals "
    "for support."
  ),
  "Sol_infrastructure": (
    "Proper landfill in good location; three municipalities sharing PPP "
    "infrastructure."
  ),
  "Sol_subsitutes": (
    "Paper plates/cups for feasts; jute bags at purchase; LP instead of PP plastic."
  ),
  "Sol_enforcement": (
    "Progressive fines NPR 5,000–11,000; penalties for unsegregated waste; "
    "official municipal enforcement needed."
  ),
  "Sol_monitoring": (
    "GPS vehicle tracking; barcode household visit/payment verification in "
    "3 months; staff attendance monitoring."
  ),
  "Traditions to build on (free-hand)": (
    "Paper plates/cups instead of plastic at feasts; jute sacks for purchased "
    "goods; reuse items multiple times before disposal."
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
    "Interview: NPL_11  |  Country: NEPAL  |  Safa Urja  |  "
    "Actor: private_sector  |  June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_11"
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
