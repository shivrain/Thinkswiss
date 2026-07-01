"""Generate interview coding spreadsheet for Nepal interview NPL_13.

Coding follows the Overview All Interviews codebook and is based on the
Creasion recycling/waste-management organization interview of 23 June 2025
— transcript and interview guideline. Uses expanded Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_13.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview transcript", "Creasion — 23 June 2025 (5:15pm)"),
  ("Interview guideline", "Creasion structured guideline (Themes A–D)"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_13",
  "Country": "NEPAL",
  "Interview number": "Creasion / Chitwan recycling series",
  "Interview date": "23 June 2025 (5:15pm)",
  "Interviewee": "Creasion representative — recycling and waste-management organization",
  "Affiliation / role": (
    "Private recycling operator and CSO-adjacent organization — PET recovery, "
    "granule production, informal-worker formalization, entrepreneur capacity-building"
  ),
  "Organization": (
    "Creasion — recycling facility (~70 tons PET/month capacity); EU-funded "
    "recovery project; CAP project (World Bank, UNOPS, SACEP); Plus Nepal "
    "association member"
  ),
  "Stakeholder list mapping": "private_sector",
  "Coded actor type (codebook)": "private_sector",
  "Actor classification rationale": (
    "Industrial recycler registered with Department of Industry producing granules "
    "for market — also runs CSO-type programmes (informal-worker support, "
    "municipal collaboration) but core identity is private recycling enterprise."
  ),
  "Perspective": (
    "Recycler-advocate — emphasizes end-of-life plastics, food-contact safety, "
    "microplastics from coloured PET, scrap-tax burden, bureaucratic barriers, "
    "and exploitation of informal women waste pickers"
  ),
  "Project context": (
    "70 tons PET/month to facility; NPR 25/kg formal rate vs NPR 10 informal; "
    "scrap tax at municipal level; 60% export condition after year 3; only "
    "~2 recycling centres nationally; WB estimate 60,000 tons plastic/year (27% PET)"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_13",
  "Actortype": "private_sector",
  "Problem_awareness_pop": "medium",
  "Problem_awareness_pol": "high",
  "Problem_severity": "high",
  "Problem_littering": "yes",
  "Problem_consumption": "yes",
  "Problem_recycling": "yes",
  "Problem_waste_mgmt": "yes",
  "Problem_production": "yes",
  "Problem_alternatives": "yes",
  "Problem_waste_segregation": "no",
  "Problem_import": "no",
  "Impacts": "health, microplastics, air pollution, water",
  "Governance": "yes",
  "Relevance_international_pol": "yes",
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
  "Pol_awarness": "no",
  "Pol_education": "yes",
  "Pol_capacity": "yes",
  "Pol_RD": "yes",
  "Pol_tax": "yes",
  "Pol_ban": "no",
  "Pol_subsitutes": "no",
  "Pol_clean_up": "no",
  "Pol_upcycling": "yes",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "no",
  "Sol_responsibilities": "yes",
  "Sol_epr": "no",
  "Sol_awarness": "no",
  "Sol_segregation": "no",
  "Sol_upcycling": "yes",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "yes",
  "Sol_RD": "yes",
  "Sol_tax": "yes",
  "Sol_finance": "yes",
  "Sol_ban": "no",
  "Sol_infrastructure": "yes",
  "Sol_subsitutes": "no",
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
    "Creasion — industrial recycler with CSO programmes. Stakeholder list: "
    "private_sector. Codebook: private_sector."
  ),
  "Problem_awareness_pop": (
    "Most Nepali consumers and businesses prefer virgin over recycled plastic; "
    "food-contact uses virgin; recycled granules for non-food tools only."
  ),
  "Problem_awareness_pol": (
    "Not consulted in solid waste management policy development (gradually "
    "being considered now); consultation gap between recyclers and government."
  ),
  "Problem_severity": (
    "End-of-life plastics critical; food-packaging safety from recycled materials; "
    "coloured PET emits significant microplastics; only ~2 recycling centres nationally."
  ),
  "Problem_littering": (
    "Informal pickers collect PET from landfills and dumpsites; exploitation "
    "by intermediary mafias."
  ),
  "Problem_consumption": (
    "Virgin plastic preferred over recycled; market demand structure unclear "
    "for granule buyers."
  ),
  "Problem_recycling": (
    "Only 2 recycling centres identified; sector social/environmental standards "
    "'pathetic' — processing at 100–200°C with carcinogenic emissions."
  ),
  "Problem_waste_mgmt": (
    "60,000 tons plastic/year estimated (WB: 27% PET) but measurement "
    "methodology questioned."
  ),
  "Problem_production": (
    "Nepal generates ~60,000 tons plastic annually per World Bank data; "
    "70 tons PET/month to Creasion facility alone."
  ),
  "Problem_alternatives": (
    "Virgin plastic replacing recycled products; recycled could be cheaper "
    "without scrap tax but isn't price-competitive."
  ),
  "Impacts": (
    "Health: carcinogenic gas emissions from low-standard recycling at "
    "100–200°C; food-contact safety risks from recycled packaging. "
    "Microplastics: coloured PET bottles cannot be reused — only transparent; "
    "coloured PET emits significant microplastics when processed. Air pollution: "
    "processing emissions. Water: Creasion's ATP recycles ~10,000 litres with "
    "no drainage outflow (positive practice contrasted with sector norms)."
  ),
  "Relevance_international_pol": (
    "EU-funded recovery project; CAP project (World Bank, UNOPS, SACEP) — "
    "South Asia regional recycling initiative."
  ),
  "Coordination_sectoral": (
    "Need platforms linking private recyclers, CSOs, formal (Plus Nepal) and "
    "informal (NRR) associations; limited cross-sector consultation."
  ),
  "Coordination_levels": (
    "Department of Industry + multiple LG approvals for land, water, electricity; "
    "repeated bureaucratic process when relocating operations."
  ),
  "Unclear_responsibilities": (
    "Unclear granule buyers and market structure; who sets/enforces recycling "
    "standards across informal sector."
  ),
  "Federal government": (
    "Department of Industry — land registration, water, electricity; 60% export "
    "condition after year 3; environmental standards for output testing."
  ),
  "Local government": (
    "LG approvals for industry setup; municipal scrap tax on PET bottles collected."
  ),
  "Private_sector": (
    "Creasion, Plus Nepal (formal association), NRR (informal/unregistered); "
    "only ~2 recycling centres nationally."
  ),
  "Civil_society": (
    "Creasion CSO-type programmes; EU project; entrepreneur capacity-building; "
    "CSO trainings via associations."
  ),
  "Households": (
    "Informal waste pickers (especially women) selling PET at NPR 10 from "
    "landfills vs NPR 25/kg formal rate."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Virgin plastic preferred by businesses for food contact; granule market "
    "buyers unclear."
  ),
  "Monitoring": (
    "National plastic generation data methodology unclear (60,000 tons/year, "
    "27% PET from WB); GPS/barcode systems not discussed — standards monitoring weak."
  ),
  "Financial_resources": (
    "Scrap tax on PET raises raw material and granule costs; 60% export "
    "requirement after year 3 constrains business model; recycled products "
    "could be cheaper without tax."
  ),
  "Research": (
    "Methodology behind 60,000 tons/year plastic estimate questioned; need "
    "transparent national waste data for policymaking."
  ),
  "Infrastructure": (
    "Bureaucratic burden for land, water, electricity at setup and relocation; "
    "High-Affluent Treatment Plant (ATP) enables responsible water recycling."
  ),
  "Enforcement": (
    "Weak social/environmental standards enforcement across recycling sector; "
    "Creasion tests output to government environmental standards but sector "
    "norms described as pathetic."
  ),
  "Policies in place": (
    "Municipal scrap tax on PET; Department of Industry export condition (60% "
    "after year 3); government environmental standards for recycled output testing."
  ),
  "Pol_education": "CSO trainings via formal/informal associations (Plus Nepal, NRR).",
  "Pol_capacity": (
    "EU project enhancing plastic recovery capacity; CAP project building "
    "entrepreneur capacity; sophisticated recycling technology (1 ton/hour input)."
  ),
  "Pol_RD": (
    "Creasion developing high-quality granule standardization; safety testing "
    "before market; food-contact recycled plastic standards needed."
  ),
  "Pol_tax": "Scrap tax on PET bottles paid at municipal level.",
  "Pol_upcycling": "Convert plastic to new products — dual livelihood and environment benefit.",
  "Pol_recycling": (
    "Recycle and downsize plastics; PET to granules; 70 tons/month facility capacity."
  ),
  "Pol_waste_collection": (
    "Collaborate with municipalities to identify and formalize informal collectors; "
    "informal pickers from landfills/dumpsites."
  ),
  "Sol_responsibilities": (
    "Municipalities as collaborators in formalizing informal collection networks; "
    "industry standards for recyclers."
  ),
  "Sol_upcycling": "Convert recovered plastic into new products supporting waste-worker livelihoods.",
  "Sol_recycling": (
    "Expand high-quality recycling capacity; CAP/World Bank technology; granule "
    "standardization for market confidence."
  ),
  "Sol_education": "CSO trainings through Plus Nepal and NRR associations.",
  "Sol_capacity": (
    "Build recycling entrepreneur capacity; EU and CAP projects; formalize "
    "informal waste workers through municipal partnerships."
  ),
  "Sol_RD": (
    "Standardization of recycled granule quality; food-contact safety research; "
    "transparent national plastic-waste measurement methodology."
  ),
  "Sol_tax": (
    "Reconsider scrap tax structure — raises costs and undermines recycled "
    "product competitiveness vs virgin plastic."
  ),
  "Sol_finance": (
    "Reconsider 60% export condition after year 3; fair pricing for informal "
    "pickers (NPR 25/kg vs NPR 10 exploitation)."
  ),
  "Sol_infrastructure": (
    "Streamline Department of Industry and LG approval processes for land, "
    "water, electricity — especially on relocation."
  ),
  "Sol_enforcement": (
    "Enforce social and environmental standards across recycling sector; stop "
    "100–200°C carcinogenic-emission processing."
  ),
  "Sol_monitoring": (
    "Clear transparent methodology for national plastic waste generation estimates; "
    "sector quality and safety monitoring."
  ),
  "Traditions to build on (free-hand)": (
    "NA — no traditional practices identified; focus on formalizing informal "
    "sector and industrial recycling standards."
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
    "Interview: NPL_13  |  Country: NEPAL  |  Creasion  |  "
    "Actor: private_sector  |  23 June 2025"
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_13"
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
