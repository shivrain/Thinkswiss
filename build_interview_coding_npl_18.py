"""Generate interview coding spreadsheet for Nepal interview NPL_18.

Coding follows the Overview All Interviews codebook and is based on the
dry waste management practitioner / consultant / researcher interview
(Kathmandu & Pokhara, active since 2017) — interview guideline. Uses expanded
Impacts taxonomy.
"""

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = "/workspace/interview_coding_NPL_18.xlsx"

DARK_BLUE = "1F3864"
MID_BLUE = "2E75B6"
LIGHT_BLUE = "D9E2F3"
LIGHT_GREEN = "E2EFDA"
ORANGE = "FCE4D6"
WHITE = "FFFFFF"

SOURCES = [
  ("Interview guideline", "Dry waste practitioner / consultant — structured guideline"),
  ("Codebook", "Overview All Interviews — NEW Codebook (expanded Impacts)"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_18",
  "Country": "NEPAL",
  "Interview number": "Dry waste practitioner / consultant series",
  "Interview date": "June 2025 (fieldwork; exact date not recorded in guideline)",
  "Interviewee": "Dry waste management practitioner — name not recorded",
  "Affiliation / role": (
    "Dry waste management operator, UNDP consultant, and policy researcher "
    "(UK university collaboration) — Kathmandu base with Pokhara activities"
  ),
  "Organization": (
    "Independent practitioner since 2017; UNDP skill-development training "
    "(~400 participants, Pokhara); Bhaktapur awareness campaigns; multilayer "
    "plastic management focus"
  ),
  "Stakeholder list mapping": "private_sector",
  "Coded actor type (codebook)": "private_sector",
  "Actor classification rationale": (
    "Private-sector dry waste operator sorting plastics by grade for downstream "
    "recyclers — also consults for UNDP and conducts academic policy research, "
    "but core identity is waste-management practitioner/entrepreneur."
  ),
  "Perspective": (
    "Sector insider-advocate — emphasizes missing value chain, cherry-picking, "
    "EPR as top priority, end-of-life management over bans, and coherent "
    "national framework with municipal implementation"
  ),
  "Project context": (
    "UNDP-funded Pokhara training (~400) on recycling/upcycling/multilayer "
    "plastics; Bhaktapur plastic-reduction campaigns; UK university end-of-life "
    "household-products research; SUP >40 micron ban discussed but not implemented"
  ),
}

CODING = {
  "Country": "NPL",
  "Country name": "NEPAL",
  "ID": "NPL_18",
  "Actortype": "private_sector",
  "Problem_awareness_pop": "high",
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
    "rivers, dumping sites, visual pollution"
  ),
  "Governance": "yes",
  "Relevance_international_pol": "yes",
  "Coordination_sectoral": "yes",
  "Coordination_levels": "yes",
  "Unclear_responsibilities": "yes",
  "Actors": "yes",
  "Federal government": "yes",
  "Provincial government": "no",
  "Local government": "yes",
  "Students": "yes",
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
  "Pol_capacity": "yes",
  "Pol_RD": "yes",
  "Pol_tax": "no",
  "Pol_ban": "yes",
  "Pol_subsitutes": "no",
  "Pol_clean_up": "no",
  "Pol_upcycling": "yes",
  "Pol_recycling": "yes",
  "Pol_waste_collection": "yes",
  "Solutions": "yes",
  "Sol_lead_agency": "yes",
  "Sol_responsibilities": "yes",
  "Sol_epr": "yes",
  "Sol_awarness": "yes",
  "Sol_segregation": "yes",
  "Sol_upcycling": "yes",
  "Sol_recycling": "yes",
  "Sol_education": "yes",
  "Sol_capacity": "yes",
  "Sol_RD": "yes",
  "Sol_tax": "no",
  "Sol_ban": "no",
  "Sol_finance": "yes",
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
    "Dry waste management practitioner/consultant (since 2017). Stakeholder list: "
    "private_sector. Codebook: private_sector."
  ),
  "Problem_awareness_pop": (
    "Plastic pollution not taken seriously enough by public; need school "
    "curriculum to reduce ignorance from childhood."
  ),
  "Problem_awareness_pol": (
    "Lack of coherent policy; no dedicated plastic mandate under MSW Act; "
    "unclear which department leads; decision-makers don't take issue seriously."
  ),
  "Problem_severity": (
    "Low-value plastics discarded at riverbanks and dump sites; multilayer "
    "plastics among hardest to manage; no established value chain."
  ),
  "Problem_littering": (
    "Low-value plastics with no market value end up at riverbanks and dump sites."
  ),
  "Problem_consumption": (
    "Consumers must reduce plastic consumption at household/individual level."
  ),
  "Problem_recycling": (
    "Cherry-picking — only high-value plastics collected; low-value and multilayer "
    "plastics neglected; inadequate processing/recycling infrastructure."
  ),
  "Problem_waste_mgmt": (
    "No established value chain from collection through processing; fragmented "
    "municipal policies without shared framework."
  ),
  "Problem_production": (
    "Producers face no EPR obligation for end-of-life of plastics they market."
  ),
  "Problem_alternatives": (
    "Focus should be end-of-life management not banning/replacing plastics "
    "(temporary fixes); sustainable solutions for low-value plastics hardest gap."
  ),
  "Problem_waste_segregation": (
    "Core practitioner work is sorting collected plastics into grades for "
    "downstream recyclers — implies segregation needed but not systematic."
  ),
  "Impacts": (
    "Rivers: low-value plastics discarded at riverbanks. Dumping sites: plastics "
    "with no market value dumped. Visual pollution: implied through littering "
    "of uncollected low-value fractions."
  ),
  "Relevance_international_pol": (
    "Technology transfer from countries with solved models; UK university policy "
    "research collaboration; learn from existing international models."
  ),
  "Coordination_sectoral": (
    "Poor alignment among government, private sector, communities, and donors; "
    "no coordination structure connecting actors and donors."
  ),
  "Coordination_levels": (
    "Fragmented municipal policies; national MSW Act without plastic-specific "
    "mandate; DoE role without community coordination."
  ),
  "Unclear_responsibilities": (
    "No dedicated governing body; unclear which national department leads; "
    "Ministry of Environment has not taken role effectively."
  ),
  "Federal government": (
    "Government of Nepal responsible nationally but subsumed under MSW Act; "
    "planned SUP >40 micron ban not fully implemented; DoE involved."
  ),
  "Local government": (
    "Municipalities key actors with authority for own waste policies; should be "
    "primary implementing and monitoring body — formally accountable for outcomes."
  ),
  "Students": (
    "Waste management education in school curriculum — 'catch them as young as "
    "you can.'"
  ),
  "Private_sector": (
    "Informal scrap collectors recover high-value plastics; formal private sector "
    "hesitant to invest; waste managers need formal support; practitioner sorts "
    "by grade since 2017."
  ),
  "Civil_society": (
    "Community groups worked with on multilayer plastic management; Bhaktapur "
    "awareness campaigns."
  ),
  "Science": (
    "Policy research with UK university on end-of-life household products; "
    "UNDP consulting on waste management."
  ),
  "Households": (
    "Consumer reduction at household level; responsible consumption and disposal "
    "culture needed."
  ),
  "Education institutions": (
    "Formal school curriculum for waste management education advocated."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Producers accountable via EPR for waste their products generate."
  ),
  "Monitoring": (
    "Municipalities should monitor plastic waste outcomes; SUP ban not enforced."
  ),
  "Financial_resources": (
    "No sector incentives; private sector lacks investment confidence; funding "
    "needed for collection chains and recycling facilities."
  ),
  "Research": (
    "UK university end-of-life policy research; multilayer plastic project; "
    "need applied research on low-value plastic end-of-life."
  ),
  "Infrastructure": (
    "Inadequate collection, processing, and recycling infrastructure at scale; "
    "invest in facilities especially for low-value plastics."
  ),
  "Enforcement": (
    "SUP ban not fully enforced; absence of implementing mechanisms and "
    "practical guidelines."
  ),
  "Policies in place": (
    "Planned SUP ban above 40 microns (not fully implemented); MSW Act (no "
    "dedicated plastic mandate); fragmented municipal policies."
  ),
  "Pol_awarness": (
    "Bhaktapur plastic-reduction awareness campaigns; UNDP Pokhara training "
    "reaching ~400 participants."
  ),
  "Pol_education": (
    "UNDP skill development on recycling/upcycling/paper/multilayer plastics; "
    "school curriculum advocacy."
  ),
  "Pol_capacity": (
    "UNDP-funded training ~400 participants in Pokhara; capacity for waste "
    "managers and sorters."
  ),
  "Pol_RD": (
    "UK university end-of-life household-products research; multilayer plastic "
    "management project."
  ),
  "Pol_ban": (
    "Planned national SUP ban above 40 microns — discussed but not fully "
    "implemented."
  ),
  "Pol_upcycling": (
    "UNDP training covered upcycling and paper upcycling alongside recycling."
  ),
  "Pol_recycling": (
    "Practitioner sorts plastics by grade for downstream recyclers; recycling "
    "facilities investment needed."
  ),
  "Pol_waste_collection": (
    "Informal collectors cherry-pick high-value plastics; collection chain "
    "funding needed."
  ),
  "Sol_lead_agency": (
    "Dedicated national governing/implementing body with clear plastic mandate — "
    "Ministry of Environment as natural candidate."
  ),
  "Sol_responsibilities": (
    "Municipalities formally accountable for plastic outcomes; producers via EPR; "
    "waste managers formally recognized."
  ),
  "Sol_epr": (
    "Single most important first step — producers accountable for end-of-life "
    "of plastic products."
  ),
  "Sol_awarness": (
    "Bhaktapur campaigns; public education to build responsible consumption culture."
  ),
  "Sol_segregation": (
    "Grade-sorting at collection points as practitioner model; systematic "
    "value chain from collection to processing."
  ),
  "Sol_upcycling": (
    "Upcycling training in UNDP programme; part of broader processing options."
  ),
  "Sol_recycling": (
    "Invest in processing/treatment/recycling facilities; export pathways for "
    "recycled products to improve viability."
  ),
  "Sol_education": (
    "Waste management in formal school curriculum from early age."
  ),
  "Sol_capacity": (
    "UNDP training model scaled; technology transfer from countries with solved "
    "systems; formal support for waste managers."
  ),
  "Sol_RD": (
    "End-of-life solutions for low-value and multilayer plastics; UK university "
    "research collaboration."
  ),
  "Sol_ban": (
    "Bans/replacements are temporary fixes — focus on end-of-life management "
    "instead."
  ),
  "Sol_finance": (
    "Sector incentives to attract private investment; funding for collection "
    "chains and facilities; export incentives for recycled products."
  ),
  "Sol_infrastructure": (
    "Collection, processing, and recycling facilities at scale; technology "
    "transfer rather than reinventing wheel."
  ),
  "Sol_enforcement": (
    "Coherent national framework with implementing mechanisms; municipal "
    "accountability for outcomes."
  ),
  "Sol_monitoring": (
    "Municipalities as primary monitoring body for plastic waste outcomes."
  ),
  "Traditions to build on (free-hand)": (
    "NA — emphasis on modern value chain, EPR, and technology transfer rather "
    "than traditional practices."
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
    "Interview: NPL_18  |  Country: NEPAL  |  Dry waste practitioner  |  "
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
  stakeholder["A1"].value = "Stakeholder Analysis — NPL_18"
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
