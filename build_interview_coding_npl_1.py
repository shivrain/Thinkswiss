"""Generate interview coding spreadsheet for Nepal interview NPL_1.

Coding follows the Overview All Interviews codebook and is verified against:
- Notes1_Department_of_environment (primary interview notes, 23 June 2025)
- Notes2_Department_of_environment / interview guideline (supplementary notes)
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

SOURCES = [
  ("Notes1 — Department of Environment (primary)", "Interview no. 5, 23 June 2025, 3:50–4:40pm"),
  ("Notes2 — Department of Environment (supplementary)", "Interview guideline with interviewer bullet notes"),
  ("Codebook", "Overview All Interviews — NEW Codebook"),
]

STAKEHOLDER = {
  "Interview ID": "NPL_1",
  "Country": "NEPAL",
  "Interview number": "5",
  "Interview date": "23 June 2025 (3:50pm – 4:40pm)",
  "Affiliation": "Department of Environment, Ministry of Forests and Environment",
  "Interviewee": "Deepak Diwali — Deputy Director, Pollution Control (air and plastics)",
  "Coded actor type": "governmental",
  "Role in plastic governance": (
    "Federal monitoring of plastic producers/markets; policy instruments (import "
    "restrictions, bans); SWM Act revision; coordination with local governments "
    "and Department of Commerce and Supplies"
  ),
  "Perspective": "Insider — describes DoE monitoring role, planned grants, and national policy instruments",
  "Verification status": "Confirmed Environmental Department / Ministry interview",
  "Source documents": "Notes1 (primary), Notes2 (supplementary)",
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
  "Problem_waste_segregation": "yes",
  "Problem_import": "yes",
  "Impacts": "water, health",
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
  "Pol_import": "yes",
  "Pol_awarness": "yes",
  "Pol_education": "yes",
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
    "Deepak Diwali, Deputy Director (Pollution Control — air and plastics), "
    "Department of Environment. Interview no. 5, 23 June 2025."
  ),
  "Problem_awareness_pop": (
    "Notes1 Q10: public needs awareness on health impacts; behavioural/mindset change "
    "needed. Notes2: people know but not the full extent."
  ),
  "Problem_awareness_pol": (
    "Notes2: no political will; plastics/waste not a priority at decision-maker level."
  ),
  "Problem_severity": (
    "Notes1: problem everywhere; rivers; high leakage in water bodies; microplastics pollution."
  ),
  "Problem_littering": (
    "Notes2: landfills full, blocking sewer systems. Notes1: leakage into water bodies."
  ),
  "Problem_consumption": "Notes2: lots of plastics produced.",
  "Problem_recycling": (
    "Notes1: single-use plastics not economically viable to recycle; MRFs to recycling units."
  ),
  "Problem_waste_mgmt": (
    "Notes1: lack of plastic waste management; high leakage in water bodies."
  ),
  "Problem_production": "Notes2: high production rates.",
  "Problem_alternatives": (
    "Notes1: non-availability of alternative options; biodegradable plastics costly."
  ),
  "Problem_waste_segregation": (
    "Notes1 Q10: need for separate collection and segregation of plastics implied as gap."
  ),
  "Problem_import": (
    "Notes1: restriction on import and use of plastic less than 40 microns in place."
  ),
  "Impacts": (
    "Notes1: rivers/water bodies, microplastics. Notes1 Q10: health impacts. "
    "Notes2: burning/indoor health, agriculture research gap (Chitwan)."
  ),
  "Coordination_sectoral": "Notes2: stakeholders not collaborating or coordinating.",
  "Coordination_levels": (
    "Notes1 Q10: central govt policy vs LG implementation; municipals manage landfills, "
    "DoE monitors markets — multi-level split."
  ),
  "Unclear_responsibilities": (
    "Notes1 Q5: LG responsible on ground but DoE role questioned; municipals manage "
    "landfills but DoE does not monitor them."
  ),
  "Federal government": (
    "Department of Environment, Department of Commerce and Supplies, central government policy."
  ),
  "Local government": (
    "Notes1: LG responsible on ground; municipals manage landfills; LG monitors markets."
  ),
  "Private_sector": "Notes1: recycling industry; formally established industries monitored.",
  "Science": "Notes2: not much research on effects of plastics; more research needed.",
  "Households": (
    "Notes1 Q9–10: waste producers; implementation at household level; "
    "Notes2: 300 NPR household incentive."
  ),
  "Education institutions": (
    "Notes1 Q10: schools mentioned for implementation; kids already taught about plastics."
  ),
  "Private_companies(hotels, shops, etc.)": (
    "Notes1: plastic producers monitored; informal industries harder to surveil."
  ),
  "Monitoring": (
    "Notes1: lack of surveillance; monitors only 2–3 times/year; too few staff. "
    "Notes2: implementation/monitoring is the problem."
  ),
  "Financial_resources": (
    "Notes2: funds exist but not released; earmarked funds frozen if unused. "
    "Notes1: grant/subsidy provisions being developed for industry shift."
  ),
  "Research": "Notes2: not much research/knowledge on plastic effects.",
  "Infrastructure": (
    "Notes2: facilities exist in some contexts. Notes1: MRFs in use — not coded as "
    "primary infrastructure gap."
  ),
  "Enforcement": (
    "Notes1: no penalty in place; lack of surveillance. Notes1/2: weak enforcement."
  ),
  "Pol_import": "Notes1: restriction on import and use of plastic <40 microns.",
  "Pol_awarness": "Notes1 Q10: awareness campaigns and advocacy through documentaries.",
  "Pol_education": "Notes1 Q10: education in schools; kids taught about plastics.",
  "Pol_tax": (
    "No product levy/tax coded. Notes1: planned grant/subsidy for industry shift "
    "(financial incentive, not a bag tax)."
  ),
  "Pol_ban": "Notes1: ban on plastic flowers and bags (<40 microns).",
  "Pol_subsitutes": (
    "Notes1: biodegradable plastics produced; govt exploring subsidies for "
    "environment-friendly alternatives."
  ),
  "Pol_recycling": (
    "Notes1: MRFs and recycling units; circular economy discussed."
  ),
  "Pol_waste_collection": (
    "Notes1 Q10: collection system with separate collection needed/implied as measure."
  ),
  "Sol_lead_agency": (
    "Notes1: need for a specific plastics management guideline (national framework)."
  ),
  "Sol_responsibilities": (
    "Notes1: EPR / polluters-pay principle; clarify DoE monitoring role."
  ),
  "Sol_epr": "Notes1: no EPR in Nepal; polluters-pay principle recommended.",
  "Sol_awarness": (
    "Notes1 Q10: awareness on health impacts; behavioural/mindset change needed."
  ),
  "Sol_segregation": (
    "Notes1 Theme D & Q10: segregation of plastics; separate collection; "
    "systematic circulation/circular economy."
  ),
  "Sol_recycling": "Notes1: circular economy; MRFs to recycling units.",
  "Sol_education": (
    "Notes1 Q9–10: education for plastic producers; schools; public learning on health impacts."
  ),
  "Sol_capacity": (
    "Notes1: inadequate monitoring due to too few staff (only 2–3 inspections/year)."
  ),
  "Sol_RD": (
    "Notes2: R&D for alternatives needed. Notes1: need environment-friendly plastics."
  ),
  "Sol_finance": (
    "Notes1: incentives for alternatives; planned grants to shift industry from 40-micron plastics."
  ),
  "Sol_infrastructure": (
    "Notes1 Q10: collection system, separate collection, biodegradable waste to manure."
  ),
  "Sol_subsitutes": (
    "Notes1: need affordable eco-friendly alternatives; subsidy for biodegradable products."
  ),
  "Sol_enforcement": "Notes1: no penalty in place; need stronger enforcement.",
  "Sol_monitoring": (
    "Notes1: strengthen DoE monitoring; increase frequency beyond 2–3 times/year."
  ),
  "Traditions to build on (free-hand)": (
    "Notes2 follow-up asked; Notes1 cites Bhaktapur as example only — no tradition coded."
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
    "Interview: NPL_1 (no. 5)  |  Country: NEPAL  |  "
    "Department of Environment — Deepak Diwali, Deputy Director  |  "
    "23 June 2025  |  Verified against Notes1 & Notes2"
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
