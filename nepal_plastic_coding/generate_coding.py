"""
Generates a codebook-based qualitative coding of three Nepal plastic-pollution
governance interviews into a single Excel workbook, strictly following the
exact column list supplied by the research team (Country, ID, Actortype, all
Problem_*, Impacts, NPF victims/villains/hero, Governance/coordination
variables, Res_/Cul_/Tar_ actor grids, Actor_role, Discretion,
implementation-issue variables, Pol_* / Sol_* variables, Traditions to build
on, and Notes).

Sheets produced:
  1. Read_Me            - scope, sources, coding conventions
  2. Codebook            - variable dictionary matching the Coded_Data columns
  3. Coded_Data          - wide-format matrix: one row per interview, one
                           column per COLUMNS entry (exact order/spelling as
                           supplied)
  4. Coding_Explanations - long-format table (ID | Variable | Value |
                           Evidence / Explanation) documenting the
                           quote/observation behind every coded (non-NA) value

Interviews coded (Nepal):
  - NPL_1: Department of Environment (DoE), Government of Nepal - Deepak
    Diwali, Deputy Director, Pollution Control (air & plastics), Kathmandu.
    Labelled "Interview no. 5" in the project's internal running order, but
    ID'd here as NPL_1 per the research team's instruction. Coded from the
    interview notes AND the full verbatim transcript (PEGO / ENV DEP dialogue).
  - NPL_2: Former Minister, Government of Nepal - Ganesh Shah, Kathmandu.
    Coded from the interview guideline notes AND the full verbatim transcript
    (21 June 2025).
  - NPL_3: RSCT, a Nepali savings-and-credit cooperative network (est. 1991)
    that recently expanded into urban plastics work. Coded from the
    hand-written interview notes only (no transcript available). (Originally
    ID'd NPL_2 before the Ganesh Shah interview was added and assigned NPL_2
    by the research team; RSCT was renumbered to NPL_3 to keep IDs unique.)

Coding conventions:
  - yes / no / NA -> NA means the topic was not addressed in that interview
    (absence of evidence); "no" is only used when the interviewee explicitly
    indicated the item is not an issue / not currently in place.
  - Free-hand fields contain short descriptive text, or "NA" if not discussed.
  - Actor_role: 1 = formulation/policy actor, 2 = managerial/organisational
    implementer, 3 = street-level implementer, 4 = target group (comma
    separated if more than one applies).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT_PATH = "/workspace/nepal_plastic_coding/Nepal_Plastic_Governance_Coding.xlsx"

# ---------------------------------------------------------------------------
# 1. EXACT COLUMN LIST (as supplied by the research team) + descriptions
#    used both for Coded_Data headers and the Codebook reference sheet.
# ---------------------------------------------------------------------------
NPF_VICTIMS_COL = "NPF_victims"

COLUMNS = [
    "Country", "ID", "Actortype",
    "Problem_awareness_pop", "Problem_awareness_pol", "Problem_concerndness",
    "Problem_littering", "Problem_consumption", "Problem_recycling",
    "Problem_waste_mgmt", "Problem_production", "Problem_alternatives",
    "Problem_waste_segregation", "Problem_import",
    "Impacts", NPF_VICTIMS_COL, "NPF_villains", "NPF_hero",
    "Relevance_international_pol", "Coordination_sectoral", "Coordination_levels",
    "Unclear_responsibilities",
    "Res_nat_government", "Res_prov_government", "Res_loc_government",
    "Res_students", "Res_private_sector", "Res_civil_society", "Res_science",
    "Res_households", "Res_edu_institutions", "Res_private_companies",
    "Cul_nat_government", "Cul_prov_government", "Cul_loc_government",
    "Cul_students", "Cul_private_sector", "Cul_civil_society", "Cul_science",
    "Cul_households", "Cul_edu_institutions", "Cul_private_companies",
    "Tar_nat_government", "Tar_prov_government", "Tar_loc_government",
    "Tar_students", "Tar_private_sector", "Tar_civil_society", "Tar_science",
    "Tar_households", "Tar_edu_institutions", "Tar_private_companies",
    "Actor_role", "Discretion",
    "Monitoring", "Financial_resources", "Research", "Infrastructure",
    "Capacity", "Enforcement",
    "Pol_epr", "Pol_import", "Pol_awareness", "Pol_education", "Pol_capacity",
    "Pol_RD", "Pol_tax", "Pol_ban", "Pol_subsitutes", "Pol_clean_up",
    "Pol_upcycling", "Pol_recycling", "Pol_waste_collection", "Pol_effectiveness",
    "Pol_effectiveness_example",
    "Sol_lead_agency", "Sol_responsibilities", "Sol_epr", "Sol_awareness",
    "Sol_segregation", "Sol_upcycling", "Sol_recycling", "Sol_education",
    "Sol_capacity", "Sol_RD", "Sol_tax", "Sol_ban", "Sol_finance",
    "Sol_infrastructure", "Sol_subsitutes", "Sol_clean_up", "Sol_enforcement",
    "Sol_monitoring",
    "Traditions to build on (free-hand)", "Notes",
]

DESCRIPTIONS = {
    "Country": "Country name",
    "ID": "InterviewID: ISO_Nr",
    "Actortype": "nat_government, prov_government, loc_government, students, private_sector, "
                 "civil_society, science, households, edu_institutions, private_companies",
    "Problem_awareness_pop": "Lack of awareness among population was mentioned: yes/no",
    "Problem_awareness_pol": "Lack of awareness among policy makers was mentioned: yes/no",
    "Problem_concerndness": "How concerned is the interview partner: high, medium, low, NA",
    "Problem_littering": "Littering is a problem; was mentioned: yes/no",
    "Problem_consumption": "High consumption is a problem (midstream); was mentioned: yes/no",
    "Problem_recycling": "No or too little recycling is a problem; was mentioned: yes/no",
    "Problem_waste_mgmt": "Ill waste management is a problem; was mentioned: yes/no",
    "Problem_production": "High production rates (too much production of plastic products - "
                          "upstream); was mentioned: yes/no",
    "Problem_alternatives": "No good alternatives; was mentioned: yes/no",
    "Problem_waste_segregation": "Lack of segregation was mentioned: yes/no",
    "Problem_import": "Plastic imports are a problem; was mentioned: yes/no",
    "Impacts": "Free hand - impacts mentioned on water, cities, biodiversity, health, etc. or NA "
               "(examples: water, cities, biodiversity, health, visual pollution, rivers, drainage "
               "blockage, forests, aquatic life, foul smell, dumping sites, air pollution, "
               "microplastics, marine pollution, tourism, coral damage, agriculture, wildlife, "
               "workforce, climate, soil)",
    NPF_VICTIMS_COL: "NPF_victims - free hand - who are the main victims mentioned "
                     "(e.g. human, wildlife, aquatic life, etc.)",
    "NPF_villains": "Free hand - who are the bad guys/culprits contributing most to plastic "
                    "pollution mentioned by interview partners (e.g. Chinese tourists, "
                    "households, sea nomads, etc.)",
    "NPF_hero": "Free hand - who is perceived as the actor committed to solving the problem "
               "(e.g. the children will educate their parents)",
    "Relevance_international_pol": "International treaties are relevant for national level policy",
    "Coordination_sectoral": "Lack of coordination across sectors",
    "Coordination_levels": "Lack of coordination across levels",
    "Unclear_responsibilities": "Unclear responsibilities among the involved actors",
    "Res_nat_government": "Responsible: National",
    "Res_prov_government": "Responsible: State/province (e.g. Sabah)",
    "Res_loc_government": "Responsible: Municipality and below (e.g. Semporna)",
    "Res_students": "Responsible: University and school level",
    "Res_private_sector": "Responsible: Waste collectors, recyclers (connected to waste)",
    "Res_civil_society": "Responsible: NGOs, IGOs, monks, etc.",
    "Res_science": "Responsible: Researchers",
    "Res_households": "Responsible: Individuals/wider population",
    "Res_edu_institutions": "Responsible: Schools, universities",
    "Res_private_companies": "Responsible: Hotels, shops, tour guides, cafes, restaurants, etc.",
    "Cul_nat_government": "Culprit: National",
    "Cul_prov_government": "Culprit: State/province (e.g. Sabah)",
    "Cul_loc_government": "Culprit: Municipality and below (e.g. Semporna)",
    "Cul_students": "Culprit: University and school level",
    "Cul_private_sector": "Culprit: Waste collectors, recyclers (connected to waste)",
    "Cul_civil_society": "Culprit: NGOs, IGOs, monks, etc.",
    "Cul_science": "Culprit: Researchers",
    "Cul_households": "Culprit: Individuals/wider population",
    "Cul_edu_institutions": "Culprit: Schools, universities",
    "Cul_private_companies": "Culprit: Hotels, shops, tour guides, cafes, etc.",
    "Tar_nat_government": "Target group: National",
    "Tar_prov_government": "Target group: State/province (e.g. Sabah)",
    "Tar_loc_government": "Target group: Municipality and below (e.g. Semporna)",
    "Tar_students": "Target group: University and school level",
    "Tar_private_sector": "Target group: Waste collectors, recyclers (connected to waste)",
    "Tar_civil_society": "Target group: NGOs, IGOs, monks, etc.",
    "Tar_science": "Target group: Researchers",
    "Tar_households": "Target group: Individuals/wider population",
    "Tar_edu_institutions": "Target group: Schools, universities",
    "Tar_private_companies": "Target group: Hotels, shops, tour guides, cafes, etc.",
    "Actor_role": "Can be more than one selection, comma separated. 1 = formulation/policy actor "
                 "- not an implementer (e.g. ministry, governmental agency); 2 = managerial/"
                 "organisational implementer (not street-level, e.g. district office); "
                 "3 = street level implementer (e.g. village heads, local NGO staff, heads of "
                 "youth organisations, etc.); 4 = target group (e.g. shops, households, etc.). "
                 "Must be derived from the interview, not the actor group.",
    "Discretion": "If Actor_role = 2 or 3 (implementer) - free hand - does the interview partner "
                 "indicate a certain degree of discretion, e.g. private sector waste collectors "
                 "developing their own schedule/collection plan.",
    "Monitoring": "Lack of monitoring",
    "Financial_resources": "Lack of financial resources",
    "Research": "Lack of research",
    "Infrastructure": "Lack of infrastructure",
    "Capacity": "Lack of capacity of implementers",
    "Enforcement": "Lack of enforcement",
    "Pol_epr": "Policy in place: extended producer responsibility",
    "Pol_import": "Policy in place: import regulations",
    "Pol_awareness": "Policy in place: awareness campaigns",
    "Pol_education": "Policy in place: education programs",
    "Pol_capacity": "Policy in place: capacity building",
    "Pol_RD": "Policy in place: research & development",
    "Pol_tax": "Policy in place: levies or taxes on specific products (often bags)",
    "Pol_ban": "Policy in place: bans of specific products",
    "Pol_subsitutes": "Policy in place: alternatives like reusable bags, or banana leaf plates",
    "Pol_clean_up": "Policy in place: clean-up campaigns",
    "Pol_upcycling": "Policy in place: making a new product, like blacktop",
    "Pol_recycling": "Policy in place: making a new raw material",
    "Pol_waste_collection": "Policy in place: waste collection",
    "Pol_effectiveness": "Lack of policy effectiveness",
    "Pol_effectiveness_example": "Freehand, provide example or copy/paste from the interview what "
                                 "they said.",
    "Sol_lead_agency": "Solution: procedural measures to establish lead agency",
    "Sol_responsibilities": "Solution: clear responsibilities",
    "Sol_epr": "Solution: extended producer responsibility",
    "Sol_awareness": "Solution: awareness raising measures",
    "Sol_segregation": "Solution: segregation of waste",
    "Sol_upcycling": "Solution: upcycling of plastics",
    "Sol_recycling": "Solution: new raw materials",
    "Sol_education": "Solution: education programs at schools",
    "Sol_capacity": "Solution: capacity-building programs",
    "Sol_RD": "Solution: research programs",
    "Sol_tax": "Solution: taxes or levies on products",
    "Sol_ban": "Solution: ban of products",
    "Sol_finance": "Solution: financial measures",
    "Sol_infrastructure": "Solution: infrastructural measures",
    "Sol_subsitutes": "Solution: alternatives",
    "Sol_clean_up": "Solution: clean-up campaigns",
    "Sol_enforcement": "Solution: enforcement procedures",
    "Sol_monitoring": "Solution: monitoring procedures",
    "Traditions to build on (free-hand)": "Freehand or NA",
    "Notes": "Free-hand notes: interviewee identity/affiliation/location, data sources used, "
             "and any coding caveats.",
}

# ---------------------------------------------------------------------------
# 2. CODED DATA - one dict per interview, keyed exactly by COLUMNS entries.
#    Any column not explicitly set below defaults to "NA".
# ---------------------------------------------------------------------------


def blank_row():
    return {c: "NA" for c in COLUMNS}


# --- NPL_1: Department of Environment - Deepak Diwali -----------------------
doe = blank_row()
doe.update({
    "Country": "Nepal",
    "ID": "NPL_1",
    "Actortype": "nat_government",

    "Problem_awareness_pop": "yes",
    "Problem_concerndness": "high",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "rivers, lakes, land, forests, drainage/water leakage, microplastics "
               "(including residues found in Himalayan snow), health (indirectly, via "
               "call for public health awareness)",
    NPF_VICTIMS_COL: "rivers/lakes/water bodies, forests and land, the Himalayan "
                     "ecosystem (microplastics in snow), and by extension public health",
    "NPF_villains": "plastic-producing industries without extended producer responsibility "
                    "(\"there is no producer responsibility in Nepal\"), informal/unregistered "
                    "industries operating without government notice, and households/consumers "
                    "whose habits and price-sensitivity keep demand for cheap conventional "
                    "plastic bags high",
    "NPF_hero": "Department of Environment and the high-level committee (chaired by the Chief "
               "Secretary) driving bans/subsidies; local governments and NGOs operating "
               "material recovery facilities and awareness campaigns",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_edu_institutions": "yes",

    "Cul_private_sector": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "1,2",
    "Discretion": "DoE has discretion in which industries it assesses/monitors (currently "
                 "only formally established/registered plastic industries and markets - "
                 "informal producers fall outside its monitoring) and is independently "
                 "designing a new subsidy scheme for the next fiscal year to help industries "
                 "shift to compliant (>40 micron) plastic production.",

    "Monitoring": "yes",
    "Research": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_import": "yes",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"There is a production of biodegradable plastics, but "
                                 "price is high... due to this price gap the market also uses "
                                 "conventional plastic\"; \"lack of a strong surveillance "
                                 "mechanism\"; \"there is inadequate monitoring due to lack of "
                                 "[a] few people. 2-3 times [we] monitor in a year\"; \"no "
                                 "penalty in place\" - together showing the 40-micron bag ban "
                                 "and plastic-flower ban are weakly enforced.",

    "Sol_epr": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Bhaktapur municipality was cited as a positive "
                                          "local example (\"eg. Of Bhaktapur\") of waste "
                                          "management practice that could be built on/"
                                          "replicated elsewhere; no other traditional "
                                          "practices were mentioned.",
    "Notes": "Affiliation: Department of Environment (Government of Nepal). Interviewee: "
             "Deepak Diwali, Deputy Director, Pollution Control (air & plastics). Location: "
             "Kathmandu. Labelled \"Interview no. 5\" in the project's internal running order "
             "(23 June 2025, 3:50-4:40pm); coded here as NPL_1. Coded from both the interview "
             "notes and the full verbatim transcript (PEGO / ENV DEP dialogue).",
})

# --- NPL_2: Former Minister - Ganesh Shah ------------------------------------
ganesh = blank_row()
ganesh.update({
    "Country": "Nepal",
    "ID": "NPL_2",
    "Actortype": "nat_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_waste_segregation": "yes",
    "Problem_import": "yes",

    "Impacts": "soil contamination (microplastics, reduced moisture, crop yields), water "
               "pollution (rivers, drinking water, irrigation water, organic fertiliser), "
               "air pollution (burning of plastic, carcinogenic smoke), drainage/sewer "
               "blockage, aquatic life/fisheries, visible litter in urban and rural areas, "
               "agriculture, health (unsafe practices such as hot food in thin plastic bags, "
               "unlabelled/low-grade bottled water), microplastics generally",
    NPF_VICTIMS_COL: "the general public/consumers (unlabelled, low-quality plastics and "
                     "unsafe water practices), poor and rural communities (burning plastic "
                     "for fuel, historically child \"kathe\" waste-pickers), farmers and "
                     "agricultural land (microplastics/moisture loss reducing crop yields), "
                     "rivers and aquatic life (Bagmati and Chitwan-area rivers, fish, people "
                     "who fish/fetch water there)",
    "NPF_villains": "manufacturers/importers who do not label or disclose the type of "
                    "plastic used (\"many bottles do not even display the company name\"; "
                    "no clarity on PET vs. other polymers); the absence of Extended Producer "
                    "Responsibility, which lets producers externalise disposal costs; "
                    "widespread public habits/overuse (single-use bags per item, plastic "
                    "khada/folders handed out at events, hot food in plastic, single-use "
                    "tableware); and successive governments' political instability and weak "
                    "enforcement of the existing 40-micron ban",
    "NPF_hero": "informal and semi-formalising waste pickers/collectors and the ~78 "
               "registered private waste companies (Solid Waste Management Association); "
               "grassroots/people-led initiatives such as the Bagmati Cleaning Movement; "
               "innovation projects like PLEASE (Institute of Engineering) and Bio-Camp "
               "(upcycling multilayer plastic into boards/flowerpots); and the interviewee "
               "himself as a former Minister/scientist pushing for EPR, research and "
               "science diplomacy",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_science": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_private_sector": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_prov_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_science": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "1",
    "Discretion": "NA (Actor_role = 1, formulation/policy actor - the interviewee speaks as "
                 "a former Minister and science-diplomacy advocate shaping policy direction, "
                 "not as a day-to-day implementer, so the discretion field - which applies to "
                 "roles 2/3 - is not applicable).",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_import": "no",
    "Pol_ban": "yes",
    "Pol_clean_up": "yes",
    "Pol_upcycling": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Nepal banned plastics thinner than 40 microns ten years "
                                 "ago, but the ban has never been enforced\"; \"Weak "
                                 "enforcement (e.g., ban on plastics <40 microns exist on "
                                 "paper but is ignored)\"; \"there is no act, because officers "
                                 "cannot act without strict laws, bylaws, and guidelines. So, "
                                 "there is none.\"",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_upcycling": "yes",
    "Sol_recycling": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Use of leaf plates (pat/banana leaves) instead of "
                                          "plastic at large community feasts (bhoj); carrying "
                                          "a cotton bag or plain towel to market instead of "
                                          "taking a new plastic bag per item; reusing a "
                                          "plastic bag two or three times before sending it "
                                          "for recycling rather than discarding it after a "
                                          "single use.",
    "Notes": "Affiliation: Former Minister, Government of Nepal (created Nepal's Department "
             "of Environment ~15 years ago); also active internationally in Global Plastics "
             "Treaty negotiations/science diplomacy (UN meeting in Korea, regional meeting in "
             "Colombo) and in the UNOPS-supported PLEASE project. Interviewee: Ganesh Shah. "
             "Location: Kathmandu. Interview date: 21 June 2025. Two associations named in "
             "the interview: Nepal Plastic Foundation (also referred to as \"Nepal Plus "
             "Foundation\" in the transcript - likely the same organisation, a manufacturers' "
             "CSR consortium) and the Solid Waste Management Association of Nepal (~78-79 "
             "registered private waste companies). Coded from both the interview guideline "
             "notes and the full verbatim transcript.",
})

# --- NPL_3: RSCT --------------------------------------------------------------
rsct = blank_row()
rsct.update({
    "Country": "Nepal",
    "ID": "NPL_3",
    "Actortype": "civil_society",

    "Problem_awareness_pop": "yes",
    "Problem_concerndness": "medium",
    "Problem_littering": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",

    "Impacts": "air pollution (open burning), drainage/sewer blockage, dumping sites "
               "(landfills), health (indoor air pollution from burning)",
    NPF_VICTIMS_COL: "general public/urban residents (health effects of burning), local "
                     "environment in Chitwan district",
    "NPF_villains": "NA (no specific actor named; general public ignorance and lack of "
                    "political will/fund disbursement referenced instead)",
    "NPF_hero": "NA (not explicitly named)",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",

    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",

    "Cul_nat_government": "yes",
    "Cul_households": "yes",

    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",

    "Actor_role": "2,3",
    "Discretion": "RSCT designs its own bottom-up cooperative loan/fund mechanisms across "
                 "its 200+ member cooperatives and recently redirected this model to urban "
                 "plastics work, indicating discretion in programme design.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "no",
    "Capacity": "yes",

    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Fund is there; they don't release the funds\" / \"Two "
                                 "funds: ear-marked; free to use (if you don't use it it will "
                                 "be frozen)\" - illustrates how earmarked funds go unused, "
                                 "undermining policy effectiveness; also \"Implementation is "
                                 "always the problem: monitoring\" and \"Policies are there\" "
                                 "(but not enforced).",

    "Sol_awareness": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_finance": "yes",
    "Sol_subsitutes": "yes",

    "Traditions to build on (free-hand)": "RSCT's existing bottom-up, grassroots cooperative "
                                          "fund/loan model (built since 1991 through a "
                                          "two-tier structure of 200+ member cooperatives "
                                          "governed by a general assembly) could be adapted "
                                          "to finance and organise community-level plastic "
                                          "waste management, as RSCT itself has already begun "
                                          "doing in urban areas.",
    "Notes": "Affiliation: RSCT, a savings-and-credit cooperative network (est. 1991), "
             "originally focused on poverty-alleviation microfinance for rural communities in "
             "Nepal, now also working on urban plastics. Interviewee name and precise location "
             "not specified in the source notes (RSCT's own remarks reference Chitwan "
             "district). Coded from interview notes only (short bullet-point answers); no "
             "full transcript available for this interview.",
})

INTERVIEWS = [doe, ganesh, rsct]

# ---------------------------------------------------------------------------
# 3. EXPLANATIONS (long format) - quote/reasoning behind each coded value.
#    Only variables with a substantive (non-NA) value are documented here.
# ---------------------------------------------------------------------------

doe_expl = {
    "Actortype": "Deepak Diwali is Deputy Director for Pollution Control (air and plastics) "
                "at the Department of Environment, a national government body.",
    "Problem_awareness_pop": "The interviewee states people need more awareness \"on impacts "
                             "of plastic on human health\" and that \"behavioural change - "
                             "mindset change\" is still needed, implying currently "
                             "insufficient public awareness.",
    "Problem_concerndness": "The DoE describes a wide-ranging institutional response (policy "
                            "instruments, a high-level committee chaired by the Chief "
                            "Secretary, planned subsidies, monitoring) reflecting a high "
                            "level of institutional concern about the issue.",
    "Problem_recycling": "\"Single-use plastics... are not economically viable for "
                         "recycling\" identifies a structural limitation in current "
                         "recycling capacity/viability.",
    "Problem_waste_mgmt": "\"It is lacking to manage the waste, including the plastic "
                          "waste\" and \"high level of leak[age] of plastic in our different "
                          "river systems\" directly state inadequate waste management.",
    "Problem_alternatives": "\"The option is quite expensive\" and the ~100 rupee/kg price "
                            "gap between conventional and compostable bags show viable "
                            "alternatives are limited/unaffordable.",
    "Problem_waste_segregation": "The solutions raised later (\"systematic circulation of "
                                 "plastics waste (circular economy), segregation\"; "
                                 "\"Collection system - segregation of plastics\") imply that "
                                 "segregation is currently insufficient.",
    "Impacts": "\"In our river system, in our lake system, in our land, forest, everywhere "
               "our plastic pollution exists\"; \"high level of leak[age] of plastic in our "
               "different river systems\"; microplastic \"residue... in our snow\" in the "
               "Himalayan range.",
    NPF_VICTIMS_COL: "Rivers, lakes, land and forests are described as polluted, and "
                     "microplastics were found even \"in our snow\" in the Himalayan range - "
                     "framing the natural environment (and, by extension, the people relying "
                     "on it) as the victims.",
    "NPF_villains": "\"There is no [producer] responsibility in Nepal\" (extended producer "
                    "responsibility absent) points to plastic producers; \"informal without "
                    "govt notice these are the problems\" points to unregistered industries; "
                    "and \"it is our habits and needs\" plus the price-sensitivity of "
                    "consumers points to households/the public.",
    "NPF_hero": "The Department of Environment presents itself and the high-level committee "
               "(chaired by the Chief Secretary) as driving bans, monitoring and planned "
               "subsidies, while crediting local governments and NGOs for running material "
               "recovery facilities and awareness campaigns.",
    "Coordination_sectoral": "The interviewee is unsure of DoE's own boundary versus the "
                             "Department of Urban Development on the Solid Waste Management "
                             "Act (\"these things look after our urban development\"; \"DoE "
                             "role in plastic or waste mgmt.?? Department of Urban "
                             "Development??\"), indicating unclear cross-sector coordination.",
    "Coordination_levels": "\"Whose responsibility is it to manage the landfills? "
                           "Municipals. Do you monitor this as well? - No.\" shows a gap "
                           "between what national (DoE) and local government each monitor.",
    "Unclear_responsibilities": "The interviewee explicitly could not clearly state whether "
                                "solid-waste-act ownership sits with DoE or the Department of "
                                "Urban Development (\"I also heard that the solid waste "
                                "management act is being revised... actually these things "
                                "look after our urban development\").",
    "Res_nat_government": "DoE issues directives/gazette notices, chairs (via the Chief "
                          "Secretary) the high-level committee, and regularly monitors "
                          "plastic-producing industries and markets.",
    "Res_loc_government": "\"Local government[s] lead\" waste management; \"our local "
                          "government is the governing body for waste management.\"",
    "Res_private_sector": "\"Private companies... participate in recycling\"; the recycling "
                          "industry and private waste-collection contractors (via tenders) "
                          "are named as active responsible actors.",
    "Res_civil_society": "\"NGOs also operate recovery facilities and run awareness "
                         "campaigns\"; \"Some NGOs are involved in conducting the material "
                         "recovery facility in Kathmandu... and getting awareness.\"",
    "Res_edu_institutions": "\"In education kids are taught about it\" indicates schools "
                            "already play some role in addressing the issue.",
    "Cul_private_sector": "\"There is no [extended] producer responsibility in Nepal\" and "
                          "continued production/sale of banned items due to weak enforcement "
                          "implicate plastic-producing/selling businesses.",
    "Cul_households": "\"It is our habits and needs\" and price-sensitivity keeping demand "
                      "for cheap conventional bags high attribute part of the problem to "
                      "household/consumer behaviour.",
    "Cul_private_companies": "\"Informal without govt notice these are the problems\" - "
                             "unregistered/informal industries selling non-compliant "
                             "plastics outside of regulatory reach.",
    "Tar_nat_government": "Q9 answer: \"all stakeholders. Central govt and LG.\"",
    "Tar_loc_government": "Same Q9 answer explicitly includes local government as a focus of "
                          "future policy.",
    "Tar_private_sector": "\"Waste producers\" and \"plastic producers - how to manage "
                          "better education\" are named as groups future policy should "
                          "target.",
    "Tar_households": "Q10 answer: support is needed \"at every household level\".",
    "Tar_edu_institutions": "Q10 answer explicitly lists \"schools\" alongside households as "
                            "a focus for support/education measures.",
    "Tar_private_companies": "\"Plastic producers... they need to invest as well\" frames "
                             "producing companies as a target group that must change "
                             "behaviour/investment.",
    "Actor_role": "As Deputy Director at DoE, the interviewee both drafts/oversees policy "
                 "instruments (gazette notices, directives - role 1) and personally "
                 "describes DoE's operational monitoring of industries and markets (role 2).",
    "Discretion": "DoE \"only assess[es] formally established industries\" (informal ones "
                 "are excluded from its monitoring scope) and is independently developing a "
                 "subsidy scheme (\"next fiscal year... we will provide grant... to shift the "
                 "industry\") - both reflect discretionary decisions by DoE.",
    "Monitoring": "\"There is inadequate monitoring due to lack of [a] few people. 2-3 times "
                 "[we] monitor in a year\"; assessments only cover \"formally established "
                 "industries\", leaving informal ones unchecked.",
    "Research": "\"Role of DoE - issues to monitor, how to strengthen it? - types of "
               "plastics to be known\" and the call for \"a specific plastics guideline\" "
               "point to a knowledge/research gap.",
    "Capacity": "Inadequate monitoring is attributed to \"lack of [a] few people\" - a "
               "staffing/capacity constraint.",
    "Enforcement": "\"Due to the lack of a strong surveillance mechanism\" products \"are "
                  "sold even without a ban\"; \"no penalty in place\" for non-compliance.",
    "Pol_epr": "\"There is no [extended] producer responsibility in Nepal\" - explicitly "
              "stated as not currently in place.",
    "Pol_import": "\"Restriction of import and use of plastic less than 40 microns\" is an "
                 "active import-related regulation.",
    "Pol_awareness": "\"We also launched... different types of awareness campaigns\", "
                     "documentaries and social/traditional media use.",
    "Pol_education": "\"In education kids are taught about it\" indicates an existing "
                     "school-based education element.",
    "Pol_ban": "\"Ban on production, import, and use of plastic bags under 40 microns\" and "
              "\"restrictions on plastic decorative flowers\".",
    "Pol_subsitutes": "Biodegradable/starch-based bags already exist on the market as "
                      "substitutes, and a subsidy to support the shift to compliant plastic "
                      "is being planned.",
    "Pol_recycling": "\"Private companies and NGOs participate in recycling\"; material "
                     "recovery facilities feed recyclables \"to processing/recycling units\".",
    "Pol_waste_collection": "\"Waste collection is done either by local governments directly "
                            "or via private contractors\" (tender system).",
    "Pol_effectiveness": "Weak surveillance, only 2-3 monitoring visits/year, no penalties, "
                         "and the persistent price gap for alternatives all point to limited "
                         "effectiveness of the existing ban/import restriction.",
    "Sol_epr": "\"Extended producer responsibility is key - there is no producer "
              "responsibility in Nepal, while in other countries there is\" is offered "
              "directly as a needed solution.",
    "Sol_awareness": "Q10: \"awareness campaigns, advocacy through documentary\" listed as "
                     "helpful support going forward.",
    "Sol_segregation": "\"Systematic circulation of plastics waste (circular economy), "
                       "segregation\"; \"Collection system - segregation of plastics. "
                       "Separate collection and managed.\"",
    "Sol_recycling": "Circular-economy framing and \"manure out of biodegradable waste\" "
                     "point to expanded recycling/resource recovery as a way forward.",
    "Sol_education": "Q10: support needed \"at every household level., schools, monitoring "
                     "of public places\" and better education for plastic producers.",
    "Sol_capacity": "\"Plastic producers - how to manage better education. Circular economy "
                    "- make sustainability. They need to invest as well\" implies "
                    "capacity-building among producers.",
    "Sol_RD": "\"Need for a specific plastics guideline for its management\" and the need to "
             "have \"types of plastics to be known\" both call for further "
             "research/knowledge development.",
    "Sol_finance": "\"Need for incentives - alternatives\"; the planned fiscal-year grant "
                  "programme to help industries shift to compliant plastic.",
    "Sol_infrastructure": "\"Make transport department for wastes along roads. Need [for] "
                          "continuity. Sustainability.\" calls for infrastructural/logistics "
                          "investment in collection.",
    "Sol_subsitutes": "\"Need for incentives - alternatives\" - cheaper/accessible "
                      "substitute products are called for.",
    "Sol_enforcement": "\"No penalty in place\" is implicitly paired with the need to "
                       "introduce enforcement/penalty mechanisms as part of strengthening "
                       "policy.",
    "Sol_monitoring": "\"Role of DoE - issues to monitor how to strengthen it?\" and Q10's "
                      "\"monitoring of public places\" call for strengthened monitoring "
                      "going forward.",
    "Traditions to build on (free-hand)": "\"Eg. Of Bhaktapur\" is raised as a positive "
                                          "existing local example under the solutions "
                                          "discussion, suggesting an existing municipal "
                                          "practice that could be replicated elsewhere.",
}

rsct_expl = {
    "Actortype": "RSCT is described as a savings-and-credit cooperative network (\"Based on "
                "that cooperative; two tier cooperative; More than 200 cooperative\") "
                "originally focused on poverty alleviation, now also working on urban "
                "plastics - classified as civil society/cooperative sector rather than "
                "government or private business.",
    "Problem_awareness_pop": "\"In general they know; [but] they don't know the extent of "
                             "it\" and \"Public is ignorant\" - population has only partial "
                             "awareness of the plastic pollution problem.",
    "Problem_concerndness": "The interviewee acknowledges the problem exists and is "
                            "generally known (\"In general they know\") but stresses the "
                            "extent is under-appreciated, and no strong emotional or urgent "
                            "language is used - rated as medium concern.",
    "Problem_littering": "\"Blocking sewer systems\" implies discarded/littered plastic "
                         "accumulating in drainage systems.",
    "Problem_waste_mgmt": "\"Land-fills are already [full]\"; \"Blocking sewer systems\"; "
                          "\"It's burned - air pollution\" collectively describe inadequate "
                          "waste management.",
    "Problem_production": "\"Lots of plastics are produced\" - explicit mention of high "
                          "upstream production volumes.",
    "Problem_alternatives": "\"R&D - alternatives are needed\" implies current alternatives "
                            "are insufficient.",
    "Impacts": "Derived from: landfills (dumping sites), blocked sewers (drainage blockage), "
               "burning of plastic causing air pollution, and indoor health effects of "
               "burning (\"can impact health by burning; indoor\").",
    "Coordination_sectoral": "\"Stakeholders are not collaborating; they do not coordinate\" "
                             "- explicit statement of poor cross-actor collaboration.",
    "Coordination_levels": "Same quote (\"they do not coordinate\") combined with \"so it "
                           "stops at major[levels]\" suggests breakdowns between government "
                           "levels, not only sectors.",
    "Res_loc_government": "The described funding/facility system (earmarked funds, "
                          "facilities) is administered at a sub-national level, and \"it "
                          "stops at major\" implies responsibility sits with government "
                          "below the national tier.",
    "Res_private_sector": "The 300 rupee/month incentive scheme is paid to \"waste "
                          "collectors\", who are private-sector actors central to "
                          "implementation.",
    "Res_civil_society": "RSCT itself, a cooperative/NGO-type organisation, has taken on "
                         "plastics work in urban areas, i.e. it sees itself as (partly) "
                         "responsible.",
    "Cul_nat_government": "\"Fund is there; they don't release the funds\" implicates a "
                          "governing authority (fund holder) in stalling implementation.",
    "Cul_households": "\"Public is ignorant\" attributes part of the problem to household/ "
                      "population behaviour and lack of awareness.",
    "Tar_loc_government": "Local-level bodies administer the waste-collector incentive "
                          "scheme and facilities described, making them a natural target "
                          "for future measures.",
    "Tar_private_sector": "Waste collectors are directly targeted by the existing 300 "
                          "rupee/month incentive and would remain a target of future policy.",
    "Tar_households": "The 300 rupee/month scheme is conditional on \"the household reduce "
                      "[waste]\", explicitly targeting households.",
    "Actor_role": "RSCT coordinates a network of 200+ member cooperatives (organisational/ "
                 "managerial role = 2) while also directly running grassroots urban "
                 "plastics activities (street-level role = 3).",
    "Discretion": "RSCT's cooperative model was self-designed bottom-up (\"Bottom-up and "
                 "grassroot fund to give loan\") and it independently chose to redirect this "
                 "model toward plastics work, indicating discretion.",
    "Monitoring": "\"Implementation is always the problem: monitoring\" - monitoring is "
                 "explicitly named as a recurring implementation failure.",
    "Financial_resources": "\"Fund is there; they don't release the funds\"; \"Two funds: "
                           "ear-marked; free to use (if you don't use it will be frozen)\" - "
                           "funds exist on paper but are not effectively available/used.",
    "Research": "\"Not much research on effects of plastics\" and the closing remark \"More "
               "research is needed\" both point to a research gap.",
    "Infrastructure": "\"Facility is there (at least in this [case])\" - explicitly states "
                      "infrastructure/facilities are available, i.e. not lacking.",
    "Capacity": "\"Lack of how to act\" describes implementers/communities not knowing how "
               "to act - a capacity gap.",
    "Pol_waste_collection": "\"300 Rupees per month for waste collectors, if the household "
                            "reduce[s]\" describes an active local waste-collection/"
                            "incentive scheme.",
    "Pol_effectiveness": "\"Implementation is always the problem\" combined with the "
                         "earmarked-funds example shows existing policies are not "
                         "translating into effective action.",
    "Sol_awareness": "\"There is a need of communication\" is offered as a way forward, i.e. "
                     "more awareness/communication work.",
    "Sol_capacity": "The same \"lack of how to act\" problem is paired with an implicit call "
                    "to build the capacity/know-how to act.",
    "Sol_RD": "\"R&D - alternatives are needed\" is explicitly proposed as a way forward.",
    "Sol_finance": "The funds/disbursement problem identified above implies a needed "
                  "solution of ensuring earmarked funds are actually released and used.",
    "Sol_subsitutes": "Tied to the R&D need for \"alternatives\", i.e. viable substitute "
                      "products are called for.",
    "Traditions to build on (free-hand)": "RSCT's own background (\"Established in 1991\"; "
                                          "\"Bottom-up and grassroot fund to give loan\"; "
                                          "\"More than 200 cooperative[s]\"; \"Governed by "
                                          "general assembly\") describes a long-standing, "
                                          "locally rooted cooperative financing tradition "
                                          "that has already been repurposed for plastics "
                                          "work and could be scaled further.",
}

ganesh_expl = {
    "Actortype": "Ganesh Shah is a Former Minister of the Government of Nepal who created "
                "the Department of Environment; he speaks throughout as a former national "
                "policymaker/science-diplomacy actor, hence coded as national government.",
    "Problem_awareness_pop": "\"People in Nepal do not really [k]now what is eco-friendly "
                             "and what is not\"; \"there is not much awareness about "
                             "plastics\"; \"Most people think of 'plastic' only as plastic "
                             "bags, ignoring other forms\"; explicit and repeated.",
    "Problem_awareness_pol": "\"I have raised this point many times with bureaucrats and "
                             "politicians\" (implying persistent gaps in their "
                             "understanding); officials distribute plastic khada/folders at "
                             "official events without recognising the irony; the government "
                             "\"does not even know\" specifications when importing plastic "
                             "materials, and lacks testing labs to identify polymer types.",
    "Problem_concerndness": "\"Very concerned, plastic waste is increasing, with both "
                            "visible and invisible forms (microplastics) contaminating "
                            "soil, water, and air\" - explicit high concern.",
    "Problem_littering": "\"visible litter in urban and rural areas\"; the Environment "
                         "Protection Act's fines for \"discarding waste on the street\"; "
                         "government notices such as \"Don't dump here\", \"Penalty for "
                         "littering\".",
    "Problem_consumption": "\"even a customer buying only three or four tablets still "
                           "receives a fresh plastic bag\"; \"one plastic bag for one item... "
                           "another bag for a second item\"; officials handing out \"dozens, "
                           "even hundreds\" of plastic khada/folders - excessive/needless "
                           "single-use consumption repeatedly highlighted.",
    "Problem_recycling": "Only high-value plastics (e.g. PET bottles) are systematically "
                         "collected and resold; firms like Kalishishi \"collect high-value "
                         "paper, plastic, and bottles... then sell them without processing\", "
                         "i.e. genuine recycling/processing (vs. raw resale) remains limited "
                         "and mostly informal.",
    "Problem_waste_mgmt": "\"Lack of a specific, comprehensive plastics policy, only a "
                          "general Environment Protection Act exists\"; \"there is no act, "
                          "because officers cannot act without strict laws, bylaws, and "
                          "guidelines\"; management described as resting on ad hoc informal "
                          "practice rather than governance.",
    "Problem_waste_segregation": "\"Households rarely sort everything\"; segregation is "
                                 "listed among background notes and only partially practiced "
                                 "(families set aside milk pouches because they have resale "
                                 "value, not as systematic segregation).",
    "Problem_import": "\"We should know exactly what kinds of polymers are entering the "
                      "country\"; \"We lack testing laboratories of this type\"; \"Imports "
                      "keep arriving from abroad\" with no capacity to verify or regulate "
                      "what type of plastic (granules, pipes, packaging) is being imported.",
    "Impacts": "Clogging of drains/sewers; long-term persistence of litter in rivers and "
               "landscapes; microplastics in soil reducing moisture and possibly crop "
               "yields; microplastics in rivers, irrigation water and organic fertiliser; "
               "unsafe/unlabelled bottled and jar water (E. coli contamination found in a "
               "study cited by the interviewee); open burning of plastic linked to carcinogen "
               "exposure and air pollution; aquatic life and fisheries affected by river "
               "pollution.",
    NPF_VICTIMS_COL: "Consumers/public health (unlabelled low-grade plastics, unsafe "
                     "bottled/jar water); poor and rural communities (burning plastic for "
                     "fuel; historically child \"kathe\" waste-pickers); farmers/agricultural "
                     "land (soil moisture and crop-yield impacts); rivers, fisheries and "
                     "aquatic life around the Bagmati and Chitwan confluence area.",
    "NPF_villains": "Manufacturers/bottlers who do not disclose polymer type or company "
                    "name (\"many bottles do not even display the company name\"); the "
                    "absence of Extended Producer Responsibility, letting producers avoid "
                    "end-of-life costs; entrenched public habits (single bag per item, "
                    "plastic khada/folders at events, hot food in thin bags); and political "
                    "instability/weak enforcement that leaves the 40-micron ban unimplemented.",
    "NPF_hero": "The informal-to-semi-formal waste-picker/collector economy and the ~78 "
               "registered firms under the Solid Waste Management Association; grassroots "
               "movements such as the Bagmati Cleaning Movement (a people's initiative later "
               "joined by government); innovation projects PLEASE (Institute of Engineering) "
               "and Bio-Camp (upcycling multilayer plastics); and the interviewee himself, "
               "advocating EPR, research and international science diplomacy.",
    "Relevance_international_pol": "Extensive discussion of the UN Global Plastics Treaty "
                                   "negotiations (Korea meeting, regional meeting in "
                                   "Colombo), \"international agencies... amplifying regional "
                                   "voices, and that external momentum is pushing our "
                                   "government to draft legislation and commit to the treaty.\"",
    "Coordination_sectoral": "Confusion between different bodies both called \"Department of "
                             "Environment\" (one under the executive, one under Parliament); "
                             "calls for \"technical cooperation between manufacturers and "
                             "waste managers\" imply this cooperation does not yet exist.",
    "Coordination_levels": "\"No clear division of authority between federal, provincial, "
                           "and local levels\"; \"We must decide whether certain authority "
                           "should go to provincial or local governments. All of this remains "
                           "in transition.\"",
    "Unclear_responsibilities": "\"Unclear division of roles between levels of government\" "
                                "listed explicitly as a main challenge; \"an upside-down "
                                "sequence compared with most countries\" (policy is meant to "
                                "flow top-down, but in Nepal informal/local actors led first "
                                "and federal government is only now catching up).",
    "Res_nat_government": "Federal government is described as conducting the plastics "
                          "assessment, publishing notices/bans, and should own EPR and "
                          "import-testing responsibilities going forward.",
    "Res_prov_government": "\"Provincial government should coordinate among municipalities\" "
                           "- explicitly assigned a responsibility role.",
    "Res_loc_government": "\"It is slowly going to the local government\"; \"Inspection, "
                          "monitoring, and evaluation should be carried out by local "
                          "government\"; local governments increasingly set collection rules.",
    "Res_private_sector": "Private collection/recycling companies (\"birth and growth of "
                          "collectors, selling, recycling\"), the Solid Waste Management "
                          "Association, and ~78-79 registered firms are central responsible "
                          "actors in the current system.",
    "Res_civil_society": "The Bagmati Cleaning Movement began as \"a people's initiative\" "
                         "led by citizens (including the interviewee) before government "
                         "involvement.",
    "Res_science": "The PLEASE project's academic component was run by the Institute of "
                  "Engineering, cataloguing 8-9 scientific categories of plastic; a "
                  "researcher is drafting EPR modalities for Nepal.",
    "Res_edu_institutions": "Institute of Engineering (PLEASE project) and the Central "
                            "Department of Environmental Science are named as active, "
                            "responsible technical/academic contributors.",
    "Res_private_companies": "The Nepal Plastic Foundation, described as a manufacturers' "
                             "CSR consortium, and hotels that now store/sort bottles for "
                             "collection are named as engaged private-company actors.",
    "Cul_nat_government": "\"There is no act, because officers cannot act\"; political "
                          "instability (\"chief ministers, prime ministers, and even "
                          "environment ministers change frequently\") and weak enforcement "
                          "of the 40-micron ban are attributed to government failure.",
    "Cul_private_sector": "Firms such as Kalishishi \"collect high-value paper, plastic, and "
                          "bottles... then sell them without processing\", i.e. cherry-pick "
                          "profitable materials rather than managing the full waste stream.",
    "Cul_households": "\"Public perception that waste management is solely a government "
                      "responsibility, with little citizen engagement\"; overuse of "
                      "single-use items, burning of household waste including plastic, "
                      "buying bottled water without checking plastic type.",
    "Cul_private_companies": "\"Many bottles do not even display the company name\"; low "
                             "quality/unlabelled plastics in packaging; no Extended Producer "
                             "Responsibility for products like mobile phones, wires and "
                             "cables that embed large volumes of plastic.",
    "Tar_nat_government": "\"Federal: coordination; implement EPR; should support research, "
                          "regulate what kind of plastics are imported\" (Q9 answer).",
    "Tar_prov_government": "Proposed role: \"provincial government should coordinate among "
                           "municipalities\" as part of the future decentralised framework.",
    "Tar_loc_government": "\"By local gov: inspection\" (Q9 answer); \"Inspection, "
                          "monitoring, and evaluation should be carried out by local "
                          "government.\"",
    "Tar_private_sector": "\"Technical cooperation with waste collectors; and the "
                          "manufactur[ers], plus informal sector\" explicitly named as "
                          "groups that need to be brought into a coordinated future policy.",
    "Tar_science": "\"Should support research\" (Q9, re: federal government's role) and an "
                  "extended discussion of research priorities (air pollution from burning, "
                  "plastics in water, plastics in soil) positions the research/science "
                  "community as a target for future support.",
    "Tar_households": "The proposed public-awareness campaign (\"people should understand "
                      "where plastic is appropriate\"; \"first priority is public awareness "
                      "about when to use or refuse plastic\") directly targets households/"
                      "consumers.",
    "Tar_private_companies": "\"Regulate what kind of plastics are imported\" and the call "
                             "for manufacturers to specify/label polymer types both make "
                             "private companies/importers/manufacturers an explicit future "
                             "policy target.",
    "Actor_role": "As a Former Minister who created the Department of Environment and who "
                 "continues to shape policy discourse via international science diplomacy "
                 "and advocacy for EPR/a lead agency, the interviewee's role in this "
                 "interview is that of a formulation/policy actor (role 1), not a current "
                 "day-to-day implementer.",
    "Monitoring": "\"the ban has never been enforced\"; future solution explicitly proposes "
                  "that \"Inspection, monitoring, and evaluation should be carried out by "
                  "local government\", implying this is currently inadequate.",
    "Financial_resources": "\"Limited technical capacity... Lack of resources, manpower, and "
                           "technical infrastructure\" listed explicitly among the main "
                           "implementation barriers.",
    "Research": "\"We still lack rigorous scientific research\"; \"More research is needed "
               "on the impact of plastics pollution / waste burning / microplastics\" stated "
               "explicitly, with three concrete research gaps identified (air, water, soil).",
    "Infrastructure": "\"Limited technical capacity (e.g., no testing labs for imported "
                      "plastic materials)\"; \"We lack testing laboratories of this type\" - "
                      "explicit infrastructure gap.",
    "Capacity": "\"Limited technical capacity\"; \"Human Resources development -> "
               "capacity-building\" listed as a needed response to a current capacity gap.",
    "Enforcement": "\"Weak enforcement (e.g., ban on plastics <40 microns exist on paper but "
                   "is ignored)\" stated explicitly as a main challenge.",
    "Pol_epr": "\"Nepal must begin formulating EPR\"; \"a young researcher is drafting EPR "
              "modalities\" - explicitly framed as not yet in place, only in development.",
    "Pol_import": "No dedicated import-regulation regime is described; instead the "
                  "interviewee stresses Nepal currently \"do[es] not have that capacity\" to "
                  "even know what polymer types are being imported, and calls for import "
                  "regulation/testing labs as a future need.",
    "Pol_ban": "\"Nepal banned plastics thinner than 40 microns ten years ago\" - an existing "
              "(if unenforced) ban.",
    "Pol_clean_up": "The Bagmati Cleaning Movement is an ongoing/established clean-up "
                    "initiative (\"we first led the effort ourselves, removing and gathering "
                    "plastic waste... gradually a government office became involved\"), as is "
                    "the PLEASE project's bottle collection around the Chitwan river "
                    "confluence.",
    "Pol_upcycling": "Bio-Camp \"gathers multilayer pouches... to produce plastic 'plywood' "
                     "boards\" and other entrepreneurs \"make flowerpots for nurseries from "
                     "similar mixed plastics\" - active, existing upcycling activity.",
    "Pol_recycling": "An established (if largely informal/semi-formal) collection-to-"
                     "recycling value chain exists: collectors, balers, and processors "
                     "handling PET bottles and other valuable plastics; \"increasingly, "
                     "recycling and up-cycling\" among the ~78 registered private firms.",
    "Pol_waste_collection": "Extensive existing collection system described: historical "
                            "informal waste-pickers, now ~78-79 licensed firms under the "
                            "Solid Waste Management Association, competitive tenders in "
                            "Kathmandu Valley, hotels sorting bottles for pickup.",
    "Pol_effectiveness": "\"the ban has never been enforced\"; \"there is no act, because "
                         "officers cannot act without strict laws\"; ministries \"promise "
                         "that policies, programmes, bylaws, and guidelines are 'coming "
                         "soon'\" - a clear, repeated statement of policy ineffectiveness.",
    "Sol_lead_agency": "\"Who will be the main agency?\"; \"The main question, first of all, "
                       "is which agency will be the lead body. That agency must oversee the "
                       "act.\" - explicit call to establish a lead agency.",
    "Sol_responsibilities": "\"Most responsibilities should be decentralized. Inspection, "
                            "monitoring, and evaluation should be carried out by local "
                            "government. The federal government should provide overarching "
                            "guidance, while provincial government should coordinate among "
                            "municipalities\" - explicit proposed division of responsibilities.",
    "Sol_epr": "\"Implement Extended Producer Responsibility (EPR) nationwide\" listed "
              "repeatedly as a top priority solution.",
    "Sol_awareness": "\"Launch sustained public-awareness campaigns on correct plastic use\"; "
                     "\"A nationwide public-awareness campaign is urgently needed\".",
    "Sol_segregation": "\"Waste segregation skills\" listed under technical capacity "
                       "building; \"people should learn how to separate plastics correctly\".",
    "Sol_upcycling": "\"Learn from and expand successful pilot projects like PLEASE and "
                     "Bio-Camp\" (Bio-Camp being an upcycling initiative) proposed as a "
                     "model to scale up.",
    "Sol_recycling": "\"Strengthen cooperation between manufacturers, waste collectors, and "
                     "recyclers\" proposed to expand/formalise the recycling value chain.",
    "Sol_capacity": "\"Build technical capacity (labs, trained workforce, waste segregation "
                    "skills)\"; \"Human Resources development -> capacity-building\" "
                    "explicitly proposed.",
    "Sol_RD": "\"Promote research on plastic pollution impacts (air, water, soil, health)\"; "
             "three specific research priorities proposed (burning/air pollution, water "
             "contamination, soil/agriculture).",
    "Sol_infrastructure": "\"The federal government should establish [testing laboratories] "
                          "for large imports of plastic materials, granules, or pipes\" - "
                          "explicit infrastructural solution proposed.",
    "Sol_subsitutes": "\"Use of leaf plates (pat leaves, banana leaves) at feasts. Cotton/"
                      "towel bags for shopping. Encouraging reuse of plastic bags before "
                      "recycling\" - concrete substitute-based solutions proposed, rooted in "
                      "tradition.",
    "Sol_enforcement": "\"Public education and strict enforcement are essential\" - explicit "
                       "call to strengthen enforcement of existing/future bans.",
    "Sol_monitoring": "\"Inspection, monitoring, and evaluation should be carried out by "
                      "local government\" - explicit proposed solution assigning monitoring "
                      "responsibility.",
    "Traditions to build on (free-hand)": "\"Traditionally we carried cotton bags. Or we "
                                          "used a simple towel\"; \"Non-plastic plates were "
                                          "biodegradable leaves or banana[-leaf]... part of "
                                          "our culture long before plastics\"; the "
                                          "interviewee explicitly frames reduce/reuse of "
                                          "existing bags as an extension of these older, "
                                          "lower-waste habits.",
}

EXPLANATIONS = {"NPL_1": doe_expl, "NPL_2": ganesh_expl, "NPL_3": rsct_expl}
NAMES = {
    "NPL_1": "Department of Environment - Deepak Diwali, Deputy Director, Pollution "
             "Control (air & plastics), Kathmandu (Interview no. 5, 23 June 2025)",
    "NPL_2": "Former Minister, Government of Nepal - Ganesh Shah, Kathmandu "
             "(21 June 2025)",
    "NPL_3": "RSCT (savings-and-credit cooperative network, urban plastics programme)",
}

# ---------------------------------------------------------------------------
# 4. BUILD WORKBOOK
# ---------------------------------------------------------------------------

wb = openpyxl.Workbook()

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def style_header_row(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = WRAP
        cell.border = BORDER


# --- Sheet: Codebook (matches Coded_Data columns exactly) ---
ws1 = wb.active
ws1.title = "Codebook"
ws1.append(["Variable name", "Description"])
style_header_row(ws1, 1, 2)
for var in COLUMNS:
    ws1.append([var, DESCRIPTIONS.get(var, "")])
    r = ws1.max_row
    for c in range(1, 3):
        ws1.cell(row=r, column=c).alignment = WRAP
        ws1.cell(row=r, column=c).border = BORDER
ws1.column_dimensions["A"].width = 34
ws1.column_dimensions["B"].width = 100
ws1.freeze_panes = "A2"

# --- Sheet: Coded_Data (wide format, exact column order supplied) ---
ws2 = wb.create_sheet("Coded_Data")
ws2.append(COLUMNS)
style_header_row(ws2, 1, len(COLUMNS))
for interview in INTERVIEWS:
    ws2.append([interview.get(c, "NA") for c in COLUMNS])
    r = ws2.max_row
    for c in range(1, len(COLUMNS) + 1):
        ws2.cell(row=r, column=c).alignment = WRAP
        ws2.cell(row=r, column=c).border = BORDER
for i in range(1, len(COLUMNS) + 1):
    ws2.column_dimensions[get_column_letter(i)].width = 22
ws2.column_dimensions["B"].width = 10  # ID
ws2.freeze_panes = "C2"

# --- Sheet: Coding_Explanations (long format) ---
ws3 = wb.create_sheet("Coding_Explanations")
ws3.append(["ID", "Interview / Source", "Variable", "Coded value", "Evidence / Explanation"])
style_header_row(ws3, 1, 5)
for interview in INTERVIEWS:
    iid = interview["ID"]
    expl = EXPLANATIONS[iid]
    for var in COLUMNS:
        if var in expl:
            value = interview.get(var, "NA")
            ws3.append([iid, NAMES[iid], var, value, expl[var]])
            r = ws3.max_row
            for c in range(1, 6):
                ws3.cell(row=r, column=c).alignment = WRAP
                ws3.cell(row=r, column=c).border = BORDER
ws3.column_dimensions["A"].width = 10
ws3.column_dimensions["B"].width = 34
ws3.column_dimensions["C"].width = 30
ws3.column_dimensions["D"].width = 14
ws3.column_dimensions["E"].width = 100
ws3.freeze_panes = "A2"

# --- Sheet: Read_Me ---
ws4 = wb.create_sheet("Read_Me", 0)
readme_lines = [
    ("Nepal Plastic Pollution Governance - Interview Coding", True),
    ("", False),
    ("This workbook applies the supplied codebook (see 'Codebook' sheet) to three interview "
     "sources from the 'Plastic Pollution Governance in Nepal' project. Column names and "
     "order in 'Coded_Data' follow exactly the variable list supplied by the research team.", False),
    ("  1. NPL_1 - Department of Environment (DoE), Deepak Diwali, Deputy Director, "
     "Pollution Control (air & plastics), Kathmandu. Labelled \"Interview no. 5\" in the "
     "project's internal running order (23 June 2025), but coded here as NPL_1 per the "
     "research team's instruction. Coded from both the interview notes and the full "
     "verbatim transcript.", False),
    ("  2. NPL_2 - Former Minister, Government of Nepal, Ganesh Shah, Kathmandu (21 June "
     "2025). Coded from both the interview guideline notes and the full verbatim "
     "transcript, per the research team's instruction to use ID NPL_2 for this interview.", False),
    ("  3. NPL_3 - RSCT: a savings-and-credit cooperative network (est. 1991) that recently "
     "expanded into urban plastics work. Coded from the hand-written interview notes only "
     "(no transcript available). RSCT was originally ID'd NPL_2 in an earlier version of "
     "this workbook; it was renumbered to NPL_3 once the research team assigned NPL_2 to "
     "the Ganesh Shah interview, to keep IDs unique.", False),
    ("", False),
    ("Sheets in this workbook:", True),
    ("  - Codebook: the variable dictionary, listed in the exact same order as the columns "
     "in 'Coded_Data'.", False),
    ("  - Coded_Data: wide-format matrix (one row per interview, one column per codebook "
     "variable, including a free-hand 'Notes' column) with the coded values.", False),
    ("  - Coding_Explanations: long-format table giving the quote/observation used to "
     "justify every non-NA code in Coded_Data (one row per Interview x Variable).", False),
    ("", False),
    ("Coding conventions:", True),
    ("  - 'yes' / 'no' / 'NA': NA means the topic was not addressed in that interview (no "
     "evidence either way). 'no' is only used where the interviewee explicitly said the "
     "item is not an issue / not in place (e.g. Pol_epr = no for all three interviews, "
     "since each explicitly states extended producer responsibility does not yet exist in "
     "Nepal).", False),
    ("  - Free-hand fields (Impacts, NPF_victims/villains/hero, Discretion, "
     "Pol_effectiveness_example, 'Traditions to build on') contain short descriptive text "
     "derived directly from the interview content, or 'NA' if not discussed.", False),
    ("  - Actor_role: 1 = formulation/policy actor, 2 = managerial/organisational "
     "implementer, 3 = street-level implementer, 4 = target group (comma-separated if more "
     "than one applies), derived from what the interview describes the interviewee's "
     "organisation/role as actually doing - not simply from their actor type. Ganesh Shah "
     "(NPL_2) is coded as role 1 (formulation/policy actor), reflecting his role as a "
     "former Minister and policy/science-diplomacy advocate rather than a current "
     "day-to-day implementer; Discretion is therefore marked NA for that row, as the field "
     "only applies to roles 2/3.", False),
    ("", False),
    ("Note on source material: a block of longer, polished paragraph-style quotes appeared "
     "under the RSCT question list in the original material, but the content and "
     "first-person phrasing (\"we have... directives...\", \"Department of Environment is "
     "regularly monitoring...\") match the Department of Environment notes and transcript "
     "almost verbatim. These paragraphs were therefore treated as additional corroborating "
     "evidence for the DoE interview (NPL_1), not for RSCT (NPL_3), and coded accordingly.", False),
]
for text, bold in readme_lines:
    ws4.append([text])
    cell = ws4.cell(row=ws4.max_row, column=1)
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    if bold:
        cell.font = Font(bold=True, size=13 if ws4.max_row == 1 else 11)
ws4.column_dimensions["A"].width = 130

wb.save(OUT_PATH)
print(f"Workbook written to {OUT_PATH}")
print(f"Columns: {len(COLUMNS)}")
