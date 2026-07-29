"""
Generates a codebook-based qualitative coding of the Nepal plastic-pollution
governance interview set into a single Excel workbook, strictly following the
exact column list supplied by the research team (Country, ID, Actortype, all
Problem_*, Impacts, NPF victims/villains/hero, Governance/coordination
variables, Res_/Cul_/Tar_ actor grids, Actor_role, Discretion,
implementation-issue variables, Pol_* / Sol_* variables, Traditions to build
on, and Notes).

The workbook covers the FULL 32-interview master numbering supplied by the
research team (see MASTER_LIST below), one row per ID from NPL_1 to NPL_32,
in order. Only four interviews have source material (notes/transcripts)
available so far and are therefore fully coded: NPL_1 (Department of
Environment), NPL_2 (Former Minister Ganesh Shah), NPL_3 (KTM Municipal
Office / Kathmandu Metropolitan City), NPL_4 (Mayor of Dhulikhel Municipality)
and NPL_7 (Rural Self-Reliance Development Center / RSCT). The remaining 27
rows are explicit placeholders
(Country/ID/best-guess Actortype/Notes only, all substantive variables NA)
reserving the correct ID until their interview material is supplied.

Sheets produced:
  1. Read_Me            - scope, sources, coding conventions, full master
                           numbering list, colour key
  2. Codebook            - variable dictionary matching the Coded_Data columns
  3. Coded_Data          - wide-format matrix: one row per interview (NPL_1
                           through NPL_32, in order), one column per COLUMNS
                           entry (exact order/spelling as supplied). Each
                           interview's row is filled with its own colour.
  4. Coding_Explanations - long-format table (ID | Variable | Value |
                           Evidence / Explanation). The columns the research
                           team asked to always document (NPF_victims/
                           villains/hero, all Res_/Cul_/Tar_ actor-grid
                           columns, Actor_role, Discretion, Capacity,
                           Pol_effectiveness, Pol_effectiveness_example)
                           always get a row for every interview, even when
                           the value is NA (explaining why). Other columns
                           are documented wherever a specific explanation was
                           written for a fully-coded interview. Rows use the
                           same colour as their interview in Coded_Data.

Coding conventions:
  - yes / no / NA -> for fully-coded interviews, NA means the topic was not
    addressed in that interview (absence of evidence); "no" is only used
    when the interviewee explicitly indicated the item is not an issue / not
    currently in place. For placeholder interviews, NA means no source
    material has been supplied yet.
  - Free-hand fields contain short descriptive text, or "NA" if not
    discussed/not yet coded.
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

# --- NPL_7: Rural Self-Reliance Development Center (RSCT) -------------------
rsct = blank_row()
rsct.update({
    "Country": "Nepal",
    "ID": "NPL_7",
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
    "Notes": "Master list label: \"7. Rural Self-Reliance Development Center/ Narayan "
             "Nirola\". Affiliation: RSCT / Rural Self-Reliance Development Center, a "
             "savings-and-credit cooperative network (est. 1991), originally focused on "
             "poverty-alleviation microfinance for rural communities in Nepal, now also "
             "working on urban plastics. Interviewee name in the source notes is not given "
             "directly (RSCT's own remarks reference Chitwan district); per the research "
             "team's master interview list the interviewee is Narayan Nirola. Coded from "
             "interview notes only (short bullet-point answers); no full transcript "
             "available for this interview.",
})

# --- NPL_3: Kathmandu Metropolitan City (KMC) Municipal Office --------------
ktm = blank_row()
ktm.update({
    "Country": "Nepal",
    "ID": "NPL_3",
    "Actortype": "loc_government",

    "Problem_awareness_pop": "no",
    "Problem_concerndness": "NA",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",
    "Problem_import": "yes",

    "Impacts": "visual pollution (described as \"the biggest problem\"), drainage/sewer "
               "blockage, riverbank and roadside litter, open dumping (due to lack of land "
               "for proper storage/sorting), microplastics, wildlife (plastics found in "
               "animal excreta), tourism (visual pollution affecting the city's image)",
    NPF_VICTIMS_COL: "rivers and riverbanks, drainage/sewer systems, wildlife (plastic found "
                     "in excreta), the visual/tourism image of Kathmandu, and neighbouring "
                     "communities near landfill/transfer sites who bear the burden of other "
                     "municipalities' waste",
    "NPF_villains": "plastic manufacturers who keep producing bags below 40 microns despite "
                    "the ban, and importers bringing banned bags/plastic flowers in from "
                    "neighbouring countries; residents who oppose siting any waste/recovery "
                    "facility near their homes (NIMBY-style \"public opposition\"); and, to a "
                    "lesser degree, consumers who keep choosing cheap single-use plastic out "
                    "of convenience despite being aware of its harms",
    "NPF_hero": "Kathmandu Metropolitan City's own solid waste management office (formalising "
               "private-sector partnerships across seven clusters, drafting new regulations, "
               "planning micron-threshold increases); DoCoRecyclers (the NGO partner "
               "collecting/processing recyclables in Cluster 7); UNDP (technical/financial "
               "partner for mechanised recovery and storage systems); and department stores "
               "that have voluntarily stopped giving out plastic bags",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_private_sector": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_loc_government": "yes",
    "Tar_students": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "1,2",
    "Discretion": "As an autonomous local government, Kathmandu Metropolitan City drafts its "
                 "own regulations (e.g. the new solid waste management regulation enabling "
                 "formal cluster-based partnerships with private actors) and independently "
                 "negotiates MoUs/terms with private and NGO partners (e.g. the DoCoRecyclers "
                 "agreement in Cluster 7, the UNDP MoU for mechanised recovery/storage), "
                 "showing considerable discretion in how it organises implementation.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_import": "no",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Plastic bags below 40 microns have already been banned "
                                 "in Kathmandu Metropolitan City, and also nationally... But "
                                 "on the ground, the reality is very different. Plastics "
                                 "below 40 microns are still widely available, and plastic "
                                 "flowers are commonly found. The main issue lies in "
                                 "implementation... enforcement is weak\"; \"there's been no "
                                 "full assessment [of the micron ban]... it's not fully "
                                 "operational\"; \"we also try to ban the single use of "
                                 "plastic, but it's very hard for us to achieve that goal.\"",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
    "Sol_awareness": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_ban": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "When asked directly, the interviewee initially "
                                          "said no specific tradition came to mind, but then "
                                          "described \"lobti\" - leaf plates used at some "
                                          "Kathmandu community festivals instead of plastic "
                                          "plates. Effective at cutting plastic use at large "
                                          "gatherings, but costly (~20 rupees/piece vs. much "
                                          "cheaper plastic), so currently only used by "
                                          "wealthier groups; the interviewee agreed government "
                                          "subsidies for producers of such alternatives could "
                                          "help scale this practice.",
    "Notes": "Master list label: \"3. KTM Municipal Office\", affiliation \"Kathmandu "
             "Metropolitan City - Environment / Solid Waste Management Division\". "
             "Interviewee name(s) not given; the transcript suggests at least two KMC staff "
             "were present (\"our ma'am has answered...\"). Location: Kathmandu. Interview "
             "date: 23 June (year "
             "not stated on this document, but consistent with the other Nepal interviews "
             "conducted around 21-23 June 2025). Coded from both the interview guideline "
             "notes and the full verbatim transcript. Data-quality note: the bullet list "
             "under Q3a (\"2017; dry waste management...situated in Kathmandu and "
             "Pokhara...UNDP skill development training in Pokhara, 400 participants... "
             "Awareness campaigns Bhaktapur... research on policy framework in collaboration "
             "with a UK university\") describes a multi-city private/NGO recycling and "
             "research operation, not Kathmandu Metropolitan City's own government office "
             "(a single-city municipal authority would not itself be based in Pokhara or run "
             "a campaign in Bhaktapur, a separate metropolitan city). This content appears to "
             "belong to a different interview accidentally included in the same document and "
             "was therefore excluded from this row's coding; the coding instead relies on the "
             "content that is clearly attributable to the Kathmandu Metropolitan City office "
             "(Q3b onward in the notes, and the full PEGO/KTM Municipal Office transcript).",
})

# --- NPL_4: Dhulikhel Municipality - Mayor Ashok Kumar Byanju Shrestha ------
dhulikhel = blank_row()
dhulikhel.update({
    "Country": "Nepal",
    "ID": "NPL_4",
    "Actortype": "loc_government",

    "Problem_awareness_pop": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "rivers, forests, agriculture (agricultural fields), roadsides, ponds, "
               "health (plastic dangerous for the environment and human health), "
               "tourism/healthy-city image (WHO-declared healthy/LDCT town but lacking "
               "waste management)",
    NPF_VICTIMS_COL: "rivers, forests, agricultural fields, ponds and the broader "
                     "environment; human health (plastic described as dangerous for "
                     "environment and health); Dhulikhel's reputation as a WHO-declared "
                     "healthy/low-density clean town",
    "NPF_villains": "travelers on the Araniko and BP highways who throw plastic and "
                    "garbage; the national/federal government (concerned but lacking "
                    "action-oriented programmes, never monitoring locally, collapsed "
                    "Solid Waste Management Commission, controlling public land needed for "
                    "landfills); weak enforcement allowing markets to sell banned thin "
                    "plastic despite 20/40-micron rules; residents who understand waste "
                    "harms but refuse to host treatment/landfill facilities nearby",
    "NPF_hero": "Dhulikhel Municipality (LAPA project, green clubs, model ward, plastic "
               "budget); NGOs and social workers eager to collaborate; Kathmandu "
               "University; school- and community-based green clubs; women's handicraft "
               "groups repurposing plastic (small scale)",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_students": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_households": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_students": "yes",
    "Tar_private_sector": "yes",
    "Tar_civil_society": "yes",
    "Tar_households": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "1,2",
    "Discretion": "As an autonomous local government the mayor/municipality sets policy "
                 "through the municipal board, allocates its own plastic-management budget "
                 "(NPR 1.7 million plus an additional NPR 10 lakh this year), selects a "
                 "model ward, negotiates PPP/contract/MoU arrangements with private "
                 "waste-to-energy partners, and declares the final price for plastic sold "
                 "by private actors.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_RD": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_upcycling": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"The government is declaring the 20 micron, 40 micron "
                                 "plastic only used by the people. But if you want to go to "
                                 "the market area, there are many kinds of plastic\"; "
                                 "\"Before also we had the Soil Waste Management Act and the "
                                 "Soil Waste Management Commission. But the commission has "
                                 "already collapsed\"; \"20 years ago I was a deputy mayor "
                                 "of this town. I never saw the ministry people MONITOR here\"; "
                                 "\"the national government is also concerned about this "
                                 "sector. But we don't have proper action oriented "
                                 "activities.\"",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_upcycling": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Some women's groups and small-scale handicraft "
                                          "producers already repurpose plastic into goods, "
                                          "offering a local tradition of upcycling that "
                                          "could be scaled with clearer procedures and "
                                          "bylaws.",
    "Notes": "Master list label: \"4. Mayor Dhulikhel/ Ashok Kumar Byanju Shrestha\", "
             "affiliation: Dhulikhel Municipality (Mayor's Office). Interviewee: Ashok "
             "Kumar Byanju Shrestha, Mayor of Dhulikhel Municipality. Location: "
             "Dhulikhel. Interview date: 22 June 2025 (2:55pm-3:40pm), Kathmandu "
             "interview no. 4. Coded from both the interview guideline notes and the full "
             "verbatim transcript. A mayor's team member was also present and contributed "
             "additional remarks on policy fragmentation and constitutional "
             "responsibilities. Data-quality note: the source notes/transcript refer to "
             "\"NASA\" working on a national plastic-collection programme; this is "
             "treated as a national agency/programme name as stated in the source (not "
             "the US space agency) and coded accordingly. \"LDCT\" in the notes is "
             "interpreted as Low-Density Clean Town per the mayor's WHO-related "
             "remarks.",
})

INTERVIEWS = [doe, ganesh, rsct, ktm, dhulikhel]

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
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_prov_government": "NA - not mentioned. The interview discusses national (DoE/Chief "
                           "Secretary committee), local government and private/NGO actors, "
                           "but does not describe any role for provincial (state) government "
                           "in mitigating plastic pollution.",
    "Res_students": "NA - not mentioned. Students/schools are referenced only as recipients "
                    "of existing education (\"in education kids are taught about it\"), not "
                    "as an actor bearing responsibility to mitigate plastic pollution.",
    "Res_science": "NA - not mentioned. No researchers/scientific institutions are described "
                   "as currently responsible for mitigating plastic pollution in this "
                   "interview.",
    "Res_households": "NA - not mentioned as a currently responsible actor; households "
                      "appear in this interview mainly as a culprit (habits/price-"
                      "sensitivity) and as a future target group, not as an actor already "
                      "exercising mitigation responsibility.",
    "Res_private_companies": "NA - not explicitly described. Private manufacturing/retail "
                             "companies appear in this interview mainly as culprits (continued "
                             "production of non-compliant plastics) rather than as actors "
                             "currently exercising responsibility to mitigate pollution.",
    "Cul_nat_government": "NA - not explicitly blamed. The interviewee presents DoE/national "
                          "government primarily as the actor driving bans, monitoring and "
                          "planned subsidies (i.e. as part of the solution), rather than as a "
                          "culprit for the bad situation.",
    "Cul_prov_government": "NA - not mentioned. Provincial government is not discussed in "
                           "this interview at all.",
    "Cul_loc_government": "NA - not explicitly blamed. Local government is described as a "
                          "responsible/implementing actor (waste collection, market "
                          "monitoring) but is not directly blamed for the pollution problem.",
    "Cul_students": "NA - not mentioned as contributing to the problem.",
    "Cul_civil_society": "NA - not mentioned as contributing to the problem; NGOs are "
                         "described positively as running recovery facilities and awareness "
                         "campaigns.",
    "Cul_science": "NA - not mentioned as contributing to the problem.",
    "Cul_edu_institutions": "NA - not mentioned as contributing to the problem.",
    "Tar_prov_government": "NA - not mentioned as a target group for future measures.",
    "Tar_students": "NA - not explicitly named as a target group in this interview (schools "
                    "are mentioned only via \"education\" and \"edu_institutions\", not "
                    "students as individuals).",
    "Tar_civil_society": "NA - not mentioned as a target group for future measures; NGOs are "
                         "framed as current implementing partners rather than as a group "
                         "whose behaviour future policy should target.",
    "Tar_science": "NA - not mentioned as a target group for future measures.",
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
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "NPF_villains": "Coded NA - the notes do not name a specific culprit/villain actor; only "
                    "general public ignorance and unreleased government funds are referenced "
                    "as contributing factors, without a clearly personified \"bad guy\".",
    "NPF_hero": "Coded NA - no specific actor is described as the one committed to solving "
               "the problem in these short notes (RSCT describes its own broader "
               "organisational activities but does not frame itself, or anyone else, as a "
               "narrative \"hero\" of the plastics story specifically).",
    "Res_nat_government": "NA - not mentioned as currently responsible; national government "
                          "appears in these notes mainly via the (unreleased) earmarked-funds "
                          "issue, coded instead under Cul_nat_government.",
    "Res_prov_government": "NA - provincial/state government is not mentioned anywhere in "
                           "these notes.",
    "Res_students": "NA - students/schools are not mentioned in these notes at all.",
    "Res_science": "NA - researchers are not mentioned as a currently responsible actor "
                   "(only a general call for \"more research\" as a future need, captured "
                   "under Research/Sol_RD).",
    "Res_edu_institutions": "NA - schools/universities are not mentioned in these notes.",
    "Res_private_companies": "NA - hotels/shops/manufacturers are not mentioned in these "
                             "notes; the only private actor discussed is the waste-collector "
                             "incentive scheme, captured under Res_private_sector.",
    "Cul_prov_government": "NA - provincial government is not mentioned in these notes.",
    "Cul_loc_government": "NA - not explicitly blamed; the notes attribute the "
                          "unreleased-funds problem to \"they\" (interpreted as a governing/"
                          "funding authority, coded under Cul_nat_government) without "
                          "specifying local government.",
    "Cul_students": "NA - not mentioned.",
    "Cul_private_sector": "NA - waste collectors are described positively (recipients of the "
                          "300 rupee/month incentive), not as culprits.",
    "Cul_civil_society": "NA - not mentioned as contributing to the problem.",
    "Cul_science": "NA - not mentioned.",
    "Cul_edu_institutions": "NA - not mentioned.",
    "Cul_private_companies": "NA - not mentioned; no specific companies are blamed in these "
                             "short notes.",
    "Tar_nat_government": "NA - not explicitly named as a target group for future measures "
                          "in these notes.",
    "Tar_prov_government": "NA - not mentioned.",
    "Tar_students": "NA - not mentioned as a target group in these notes.",
    "Tar_civil_society": "NA - not mentioned as a target group; RSCT discusses itself as an "
                         "implementer rather than identifying other civil-society groups as "
                         "future targets.",
    "Tar_science": "NA - not mentioned as a target group.",
    "Tar_edu_institutions": "NA - not mentioned as a target group in these notes.",
    "Tar_private_companies": "NA - not mentioned as a target group; the only private actor "
                             "discussed as a target is waste collectors, captured under "
                             "Tar_private_sector.",
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
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - not mentioned as a currently responsible actor; students appear "
                    "only indirectly via general references to education/human-resource "
                    "development as a future need.",
    "Res_households": "NA - not described as currently responsible for mitigation; "
                      "households appear mainly as consumers whose habits need to change "
                      "(captured under Cul_households and Tar_households).",
    "Cul_prov_government": "NA - provincial government is not discussed as a culprit; the "
                           "interview's governance criticism focuses on unclear federal/"
                           "local division of authority and political instability at the "
                           "national level, captured under Cul_nat_government.",
    "Cul_loc_government": "NA - local government is described positively as an emerging "
                          "responsible actor (increasingly setting collection rules) rather "
                          "than being blamed for the pollution problem.",
    "Cul_students": "NA - not mentioned as contributing to the problem.",
    "Cul_civil_society": "NA - not mentioned as contributing to the problem; NGOs/"
                         "associations (e.g. the Solid Waste Management Association) are "
                         "presented as responsible/positive actors.",
    "Cul_science": "NA - not mentioned as contributing to the problem; researchers are "
                   "presented positively (e.g. drafting EPR modalities, cataloguing plastic "
                   "categories).",
    "Cul_edu_institutions": "NA - not mentioned as contributing to the problem.",
    "Tar_students": "NA - not explicitly named as a distinct target group; the interview "
                    "discusses building human-resource capacity and public awareness "
                    "broadly, without singling out students/schools as their own target "
                    "category (unlike, e.g., the KTM Municipal Office interview).",
    "Tar_civil_society": "NA - not mentioned as a target group for future measures; NGOs/"
                         "associations are discussed as current implementing partners.",
    "Tar_edu_institutions": "NA - not explicitly named as a target group for future measures "
                            "in this interview.",
}

ktm_expl = {
    "Actortype": "The interviewee(s) work for Kathmandu Metropolitan City's Solid Waste "
                "Management Office, a municipal/local government body; coded as "
                "loc_government.",
    "Problem_awareness_pop": "\"Most of the people are very well aware about the impact of "
                             "plastic pollution... they know about the microplastic, they "
                             "know about the harmful impact... Still, they choose to use the "
                             "plastic because it is very convenient.\" Unlike other Nepal "
                             "interviews, this respondent explicitly says awareness is "
                             "already high and the barrier is convenience/habit, not "
                             "ignorance - hence coded 'no' (lack of awareness was not "
                             "identified as the problem).",
    "Problem_littering": "\"Most of the plastic is multi-layer plastic, which is thrown "
                         "riverside, bank of the river, corridor, drainage\"; \"visible litter"
                         "\" is a recurring theme.",
    "Problem_consumption": "\"They don't have the habit of taking the cloth bag wherever we "
                           "go\" and continued preference for cheap single-use plastic over "
                           "costlier reusable alternatives (e.g. leaf plates) reflect ongoing "
                           "excessive single-use consumption.",
    "Problem_recycling": "\"Multi-layer plastic has no value and is often left behind\"; "
                         "\"cherry picking\" by collectors (Q3b notes) means only valuable "
                         "plastics get collected/recycled, leaving low-value plastic "
                         "unmanaged.",
    "Problem_waste_mgmt": "Extensive description of structural waste-management gaps: no "
                          "space to sort/store waste in the city, waste has to be trucked to "
                          "a landfill in another district, and \"lack of infrastructure to "
                          "manage, collect, recycle\" (Q3b notes).",
    "Problem_alternatives": "\"The plastics are so much cheaper\"; leaf plates (lobti) \"cost "
                            "20rs per piece, that is very costly\" and are only used by "
                            "wealthier groups - explicit statement that affordable "
                            "alternatives are lacking.",
    "Problem_waste_segregation": "\"Lack of the land, we cannot separate, collect [or store] "
                                 "in storage\" - segregation is explicitly hampered by space "
                                 "constraints.",
    "Problem_import": "\"Plastic companies continue to produce bags below 40 microns, and we "
                      "also import such plastic products - including plastic flowers - from "
                      "neighboring countries\" - explicit statement that imports undermine "
                      "the existing ban.",
    "Impacts": "\"The biggest problem is visual pollution\"; \"due to the plastic problem, "
               "all the drainage is blockage\"; plastic thrown \"riverside, bank of the "
               "river\"; microplastics and plastics \"in the excreta of wildlife\" (per "
               "research cited by the interviewee); visual pollution linked to tourism and "
               "lifestyle.",
    NPF_VICTIMS_COL: "Rivers/riverbanks and drainage systems (blocked and littered), "
                     "wildlife (plastic found in excreta), the city's visual/tourism "
                     "appeal, and neighbouring districts/communities that host the "
                     "Bancharedanda landfill or the Dakshinkali debris-dumping site without "
                     "being responsible for managing them.",
    "NPF_villains": "\"Plastic companies... are still producing the plastic that is less "
                    "than 40 micron\" and imported plastic bags/flowers \"from neighboring "
                    "countries\" undermine the ban; \"public opposition\" to siting any "
                    "recovery facility nearby (\"no one wants waste on their side\"); and "
                    "consumers who \"still choose to use plastic because it's convenient\" "
                    "despite being aware of the harms.",
    "NPF_hero": "Kathmandu Metropolitan City's own office (drafting new solid waste "
               "regulation, formalising 7 cluster-based private-sector partnerships, "
               "planning micron increases and subsidies); DoCoRecyclers (the NGO partner "
               "in Cluster 7); UNDP (MoU for mechanised plastic recovery/storage); and "
               "department stores that \"no longer provide plastic bags\".",
    "Coordination_sectoral": "\"Uncertain which Department should look after\" plastic waste "
                             "at the national level (Q5 notes); \"lack of coordination of the "
                             "stakeholders\" (Q3b notes); \"no clear coordination with local "
                             "communities\".",
    "Coordination_levels": "\"If land is under the jurisdiction of the national government, "
                           "it will be very difficult for the local government to reach out "
                           "to that land\"; explicit call for \"better coordination with both "
                           "the national and provincial government.\"",
    "Unclear_responsibilities": "\"Gov Nepal; uncertain which Department should look after\" "
                                "plastic waste (Q5 notes) - explicit statement of unclear "
                                "institutional ownership at the national level.",
    "Res_nat_government": "The Ministry of Forest and Environment has drafted a national "
                          "policy (under revision) and the Department of Environment is "
                          "named; central government's help is also invoked for land "
                          "identification.",
    "Res_prov_government": "\"The central or provincial governments could help us identify "
                           "suitable locations within the city or valley for plastic "
                           "recovery and management facilities\" - explicit responsibility "
                           "role assigned to provincial government.",
    "Res_loc_government": "\"The Solid Waste Management Act of 2068 clearly assigns the "
                          "responsibility for managing solid waste to the local level... "
                          "our office and the metropolitan city are responsible for "
                          "managing plastic waste.\"",
    "Res_private_sector": "DoCoRecyclers and other registered private actors now formally "
                          "collect and process recyclable waste in KMC's seven clusters "
                          "under signed MoUs.",
    "Res_civil_society": "DoCoRecyclers is explicitly described as \"a local NGO\" that "
                         "collects recyclable waste under an MoU with the municipality.",
    "Res_private_companies": "\"If you visit department stores in our city... you'll notice "
                             "they no longer provide plastic bags. Shoppers have to buy "
                             "cloth bags\" - private retail companies already exercising "
                             "responsibility by voluntarily discontinuing single-use plastic "
                             "bags.",
    "Cul_nat_government": "Land needed for waste facilities \"is under the jurisdiction of "
                          "the national government\" making it \"very difficult for the "
                          "local government to reach out to that land\"; the national policy "
                          "revision remains stalled/unclear (\"I do not know exactly\").",
    "Cul_loc_government": "The interviewee acknowledges enforcement gaps under the "
                          "municipality's own remit: \"there's been no full assessment [of "
                          "the ban]\"; \"we also try to ban the single use of plastic, but "
                          "it's very hard for us to achieve that goal.\"",
    "Cul_private_sector": "\"Cherry picking\" by collectors (Q3b notes) - only valuable "
                          "plastics are collected, leaving low-value/multi-layer plastic "
                          "behind to accumulate in rivers/dumpsites.",
    "Cul_households": "\"Public opposition\" to hosting any waste facility nearby (\"no one "
                      "wants waste on their side\"); continued convenience-driven plastic "
                      "use despite awareness of harms; \"people are not taking it very "
                      "seriously\" (Q3b notes).",
    "Cul_private_companies": "\"Plastic companies, they are still producing the plastic that "
                             "is less than 40 micron\"; plastic bags and plastic flowers "
                             "\"also being imported from neighboring countries\" despite the "
                             "ban.",
    "Tar_loc_government": "\"Implementing partners/monitoring should [be] the municipality "
                          "-> they should be made accountable\" (Q9 notes) - the "
                          "municipality itself is named as a target for accountability.",
    "Tar_students": "\"As young as you can catch them; waste management should be in the "
                    "education system\" (Q10 follow-up) - explicit call to target children/"
                    "students through the curriculum.",
    "Tar_private_sector": "\"Waste managers\" explicitly listed as a target group for future "
                          "policy (Q9 notes).",
    "Tar_households": "\"Focus reduce plastic consumers\" (Q9 notes); consumers/shoppers "
                      "targeted via cloth-bag requirements at department stores.",
    "Tar_edu_institutions": "Tied to the same call to embed waste management \"in the "
                            "education system\", implicating schools as a target for future "
                            "policy/curriculum change.",
    "Tar_private_companies": "\"Producers (EPR)\" explicitly listed as a target group (Q9 "
                             "notes); plastic industry also targeted for micron-threshold "
                             "subsidies to encourage compliance.",
    "Actor_role": "Kathmandu Metropolitan City both formulates its own local rules (\"we can "
                 "have our own legal code and formulate our own rules and regulations\", "
                 "e.g. the new solid waste management regulation - role 1) and directly "
                 "manages implementation (cluster-based MoUs, transfer stations, subsidy "
                 "planning - role 2), so both are coded.",
    "Discretion": "\"Kathmandu Metropolitan City is a local government, so we can create our "
                  "own laws and regulations. We also manage our own facilities\"; the "
                  "cluster-based MoU with DoCoRecyclers (negotiated terms, payment amount) "
                  "further shows discretion in implementation design.",
    "Monitoring": "\"There is no assessment of [the micron ban]... fully not [in] "
                  "operation[al]\" - explicit statement that current measures are not being "
                  "monitored/evaluated.",
    "Financial_resources": "Q10 notes call for \"collection chains; recycling facilities -> "
                           "financing\" and the interview highlights ongoing reliance on "
                           "donor/partner funding (UNDP, DoCoRecyclers' annual payment) to "
                           "sustain and expand operations.",
    "Research": "\"There's been no full assessment\" of the effectiveness of the micron "
               "increases; \"even if we increase it to 100 microns, we're unsure if that "
               "would be beneficial\" - explicit knowledge/evaluation gap.",
    "Infrastructure": "\"Land is the biggest challenge... lack of infrastructure to manage, "
                      "collect, recycle\"; \"we lack... technology and recycling "
                      "infrastructure\"; no space within the city for a transfer station or "
                      "material recovery facility.",
    "Capacity": "\"Technology transfer, do not reinvent the wheel\" (Q10 notes) signals a "
               "technical/knowledge capacity gap that external partners are expected to "
               "help fill.",
    "Enforcement": "\"On paper, the legal provisions are very well-written. But in practice, "
                   "enforcement is weak\"; plastic companies and importers continue supplying "
                   "banned products despite the rules.",
    "Pol_epr": "\"First policy EPR\" is listed among future/needed solutions (Q8), and \"EPR "
              "should be introduced\" (Q6 notes) - both frame EPR as not yet in place.",
    "Pol_import": "No dedicated import-control mechanism is described; instead the "
                  "interviewee states banned bags/flowers \"are also being imported from "
                  "neighboring countries\" unchecked, indicating the absence of an "
                  "effective import regulation.",
    "Pol_ban": "\"Plastic bags below 40 microns have already been banned in Kathmandu "
              "Metropolitan City, and also nationally. Plastic flowers have also been "
              "banned.\"",
    "Pol_subsitutes": "\"If you visit department stores... you'll notice they no longer "
                      "provide plastic bags. Shoppers have to buy cloth bags. Our office "
                      "also produces and occasionally distributes cloth bags.\"",
    "Pol_recycling": "DoCoRecyclers \"collect recyclable waste from that cluster and "
                     "transport it to our Deku transfer station\" under a formal, ongoing "
                     "MoU with the municipality.",
    "Pol_waste_collection": "Extensive existing collection system: city divided into 7 "
                            "clusters by ward, MoUs with registered private actors, waste "
                            "transported to Bancharedanda landfill, debris collection "
                            "arrangement with Dakshinkali Municipality.",
    "Pol_effectiveness": "\"On the ground, the reality is very different\"; \"the main issue "
                         "lies in implementation... enforcement is weak\"; \"there's been no "
                         "full assessment\" of the micron-threshold policy; \"it's very hard "
                         "for us to achieve that goal\" (regarding the single-use plastic "
                         "ban).",
    "Sol_lead_agency": "\"There is the need of a body, guiding governing body\"; \"important "
                       "missing link is implementing body\" (Q6 notes) - explicit call for "
                       "establishing/clarifying a lead implementing agency.",
    "Sol_responsibilities": "Tied to the same call for a guiding/implementing body, plus the "
                            "explicit proposal that municipalities \"should be made "
                            "accountable\" (Q9 notes) as part of a clearer division of roles.",
    "Sol_epr": "\"First policy EPR\" listed as the top future policy priority (Q8 notes); "
              "\"producers (EPR)\" named as a target group (Q9).",
    "Sol_awareness": "\"As young as you can catch them; waste management should be in the "
                     "education system; Reduce ignorance\" (Q10 follow-up) calls for future "
                     "awareness/education efforts.",
    "Sol_recycling": "\"Processing treatment/recycling\" listed as a needed future policy "
                     "focus (Q8 notes); \"policies to export recycling produces\" (Q10 notes).",
    "Sol_education": "\"Waste management should be in the education system\" (Q10 "
                     "follow-up) - explicit proposed solution.",
    "Sol_capacity": "\"Technology transfer, do not reinvent the wheel\" (Q10 notes) proposes "
                    "building capacity via external knowledge/technology transfer rather "
                    "than developing everything from scratch.",
    "Sol_ban": "Plan to progressively raise the minimum plastic thickness standard (\"we "
              "plan to increase it to 75 microns - possibly [beyond]... previously it was "
              "20 microns, then upgraded to 40, and now the goal is 75 microns\") as a "
              "strengthened ban/standard going forward.",
    "Sol_finance": "\"Subsidies should be provided to the plastic industry, especially to "
                   "increase plastic thickness\"; \"government subsidizing companies "
                   "producing alternatives\" endorsed as \"perhaps\" a good opportunity; "
                   "Q10 notes call for \"financing\" for collection chains and recycling "
                   "facilities.",
    "Sol_infrastructure": "\"If we get the land, we can [establish an] MRF... recycle\"; "
                          "call for national/provincial government help \"to identify "
                          "suitable locations... for plastic recovery and management "
                          "facilities\"; UNDP MoU for \"mechanized plastic recovery and "
                          "storage systems.\"",
    "Sol_subsitutes": "\"Government subsidizing companies producing alternatives is a good "
                      "opportunity - Yes, perhaps\", discussed in relation to scaling up "
                      "leaf-plate (lobti) style alternatives that are currently too "
                      "expensive for widespread use.",
    "Sol_enforcement": "\"Regular field inspections are needed\" - explicit proposed "
                       "enforcement solution to address the currently weak enforcement of "
                       "the micron ban.",
    "Sol_monitoring": "Tied to the same call for \"regular field inspections\" and the "
                      "acknowledged lack of assessment of policy effectiveness, implying a "
                      "need for stronger ongoing monitoring going forward.",
    "Traditions to build on (free-hand)": "When first asked directly about traditional "
                                          "practices (in a Bhutan-comparison framing), the "
                                          "interviewee said \"No.\" But later, when the "
                                          "question was revisited, described \"lobti\" - leaf "
                                          "plates used in some Kathmandu community festivals "
                                          "instead of plastic ones - as an existing, if "
                                          "underused and costly (20 rupees/piece), "
                                          "alternative practice.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - not mentioned as a currently responsible actor; students are "
                    "discussed only as a future target group via the education-system "
                    "proposal (captured under Tar_students).",
    "Res_science": "NA - no researchers/scientific institutions are described as currently "
                   "responsible for mitigating plastic pollution in this interview (the "
                   "ambiguous Q3a research-collaboration bullets were excluded from this "
                   "row's coding - see Notes).",
    "Res_households": "NA - households are not described as currently exercising "
                      "responsibility; they are discussed as a culprit (convenience-driven "
                      "plastic use) and future target group instead.",
    "Res_edu_institutions": "NA - schools/universities are not described as currently active "
                            "on plastic waste management in this interview; embedding waste "
                            "management in the curriculum is discussed only as a future "
                            "solution (captured under Sol_education/Tar_edu_institutions).",
    "Cul_prov_government": "NA - provincial government is mentioned only as a hoped-for "
                           "future collaborator on land access, not blamed for the current "
                           "situation.",
    "Cul_students": "NA - not mentioned as contributing to the problem.",
    "Cul_civil_society": "NA - not mentioned as contributing to the problem; the NGO partner "
                         "(DoCoRecyclers) is presented positively.",
    "Cul_science": "NA - not mentioned as contributing to the problem.",
    "Cul_edu_institutions": "NA - not mentioned as contributing to the problem.",
    "Tar_nat_government": "NA - national government is discussed as a needed collaborator "
                          "(land access, coordination) rather than as a target whose "
                          "behaviour a mitigation measure should change.",
    "Tar_prov_government": "NA - similarly discussed as a needed collaborator/coordination "
                           "partner rather than a target group of mitigation measures.",
    "Tar_civil_society": "NA - not mentioned as a target group for future measures.",
    "Tar_science": "NA - not mentioned as a target group for future measures.",
}

dhulikhel_expl = {
    "Actortype": "Ashok Kumar Byanju Shrestha is the Mayor of Dhulikhel Municipality, a "
                 "local government body; coded as loc_government.",
    "Problem_awareness_pop": "\"Lack of awareness\" is listed explicitly among "
                             "implementation challenges (Q6 notes); the mayor also "
                             "states that \"first of all we need a household awareness and "
                             "management system\" and that in some places \"people do not "
                             "have technically sound\" knowledge.",
    "Problem_concerndness": "\"How concerned are you? Very concerned\" (Q1 notes); the "
                            "municipal council and board have raised plastic concerns "
                            "\"so many times\"; this year the municipality allocated "
                            "dedicated budget for plastic management.",
    "Problem_littering": "Plastics are visible \"in the region nowadays, most of the "
                         "agriculture sector and some rivers and some roadside and some "
                         "forest areas\"; travelers on the Araniko and BP highways "
                         "\"throw plastics and so much garbage.\"",
    "Problem_consumption": "\"In Nepal we have a lot of internally and externally. Plastic "
                           "is used by people, some industries, and some business sectors\" "
                           "- widespread plastic use across households, industries and "
                           "businesses.",
    "Problem_recycling": "\"At the municipal level we don't have that type of manpower and "
                         "NGO properly who are working in that area\" for recycling; no "
                         "Dhopur Recycler or operating plastic-recycling company currently "
                         "in Dhulikhel; nearby pipe/water-tank factories exist but local "
                         "recycling capacity is absent.",
    "Problem_waste_mgmt": "\"Plastic and some garbage is already, we don't have a "
                          "management system\"; \"all kinds of the garbage is going to the "
                          "landfill side\"; landfill sites are not managed technically "
                          "even in Kathmandu Metropolitan City.",
    "Problem_alternatives": "The municipality is \"finding alternative ways, collection and "
                            "relative some new goods from the plastics\" and exploring "
                            "alternatives, but the mayor's team states \"properly we don't "
                            "have how to manage, how to sell, how to produce, how to "
                            "recycle for the plastic. We don't have proper laws\" - "
                            "indicating viable alternatives/procedures are still lacking.",
    "Problem_waste_segregation": "\"We have so many problems about segregating the garbage\"; "
                                 "household-level segregation is named as a first priority "
                                 "(\"segregating the household waste\").",
    "Impacts": "Agricultural fields, rivers, roadsides and forest areas littered with "
               "plastic; plastic described as \"very dangerous for the environment and "
               "human health\"; WHO has declared Dhulikhel a healthy/LDCT town but the "
               "lack of a waste-management system undermines this achievement.",
    NPF_VICTIMS_COL: "Rivers, forests, agricultural fields, ponds and the broader natural "
                     "environment are visibly polluted; human health is implicated "
                     "(\"dangerous for the environment and human health\"); and "
                     "Dhulikhel's hard-won healthy-city/LDCT status is undermined by "
                     "unmanaged plastic waste.",
    "NPF_villains": "Highway travelers on the Araniko and BP highways who discard plastic "
                    "and garbage; the federal/national government (\"concerned\" but "
                    "without \"proper action oriented activities\", never monitoring "
                    "locally, collapsed Solid Waste Management Commission, controlling "
                    "public land needed for landfills); markets that continue selling all "
                    "types of plastic despite 20/40-micron restrictions; and residents "
                    "who understand waste harms but \"nobody agrees on giving the place\" "
                    "for treatment/landfill facilities.",
    "NPF_hero": "Dhulikhel Municipality itself (LAPA project, green clubs, model ward, "
               "dedicated plastic budget, Environment and Disaster Management Unit); NGOs "
               "and social workers \"want to collaborate\"; Kathmandu University; "
               "school/community green clubs; and women's handicraft groups repurposing "
               "plastic (small scale).",
    "Coordination_sectoral": "The municipality works across sectors (energy-efficient "
                             "city, carbon neutrality, green city) and with diverse "
                             "partners (private sector, NGOs, Kathmandu University, "
                             "Dhulikhel Hospital, Bill & Melinda Gates Foundation, UCLG, "
                             "SPARC), but repeatedly stresses that \"we need to "
                             "collaborate\" because it lacks technical capacity alone.",
    "Coordination_levels": "\"Dhulikhel Municipality and the provincial and the federal "
                           "government, we can't work together nowadays\"; limited "
                           "coordination with higher levels despite constitutional local "
                           "autonomy; District Coordination Committee lacks mandate/budget "
                           "for inter-municipal waste cooperation.",
    "Unclear_responsibilities": "Mayor's team: \"the government has already managed some "
                                "policies, some laws. But who is the responsibility?\"; "
                                "Solid Waste Management Commission \"has already collapsed\" "
                                "and responsibility is \"throwing\" onto local governments "
                                "without capacity; ministry approves acts but "
                                "implementation/monitoring rests entirely on municipalities.",
    "Res_nat_government": "National government is \"concerned about this sector\"; "
                          "previously ran a plastic-collection programme (referred to as "
                          "\"NASA\" in the source); approves acts and controls public land "
                          "for landfill sites.",
    "Res_prov_government": "Provincial government \"is prioritizing this area\" with a "
                           "small project in Hetauda; historically supported private-sector/"
                           "NGO recycling projects there.",
    "Res_loc_government": "\"Constitutionally, all this kind of responsibility is on "
                          "municipal government\" - sanitation, waste management, "
                          "wastewater treatment and environmental protection; Dhulikhel "
                          "runs waste-management activities, green clubs, model ward and "
                          "plastic-budget programmes.",
    "Res_students": "Kathmandu University students have interned with the municipality "
                    "and professors have worked on joint initiatives; school-based green "
                    "clubs are being established for community plastic work.",
    "Res_private_sector": "Private waste-management/waste-to-energy company (MoU signed, "
                          "also works with Dharan Metropolitan City); PPP, contract and "
                          "MoU-based engagement models for waste management.",
    "Res_civil_society": "NGOs, social organizations and the LAPA project partner with "
                         "the municipality on plastic reduction/recycling/reuse; Bill & "
                         "Melinda Gates Foundation, UCLG and SPARC support fecal-sludge/"
                         "wastewater work.",
    "Res_edu_institutions": "Kathmandu University (within the municipality) provides "
                            "environmental expertise and has supported municipal "
                            "initiatives; students and professors have collaborated.",
    "Res_private_companies": "Tushithani Hotel operates its own wastewater treatment "
                             "plant; nearby municipalities host pipe/water-tank "
                             "manufacturers that buy recycled plastic; a private "
                             "recycling company previously operated in Dhulikhel.",
    "Cul_nat_government": "Federal government lacks \"proper action oriented activities\"; "
                          "ministry officials have never monitored Dhulikhel in 20+ years; "
                          "Solid Waste Management Commission collapsed; controls public "
                          "land so municipalities cannot manage landfill sites without "
                          "federal permission; provides insufficient technical/financial "
                          "support despite constitutional local responsibility.",
    "Cul_loc_government": "The municipality acknowledges its own gaps: \"the municipal "
                          "government does not have technical manpower\"; landfill site "
                          "not managed with adequate technical knowledge (\"technically "
                          "we do not have knowledge\").",
    "Cul_households": "Highway travelers (not only locals) discard plastic; in some areas "
                      "people \"do not have technically sound\" waste-management knowledge; "
                      "residents understand waste harms but resist hosting facilities.",
    "Tar_nat_government": "Three things needed \"to support from the federal government: "
                          "technical equipment related and some budget\" plus "
                          "infrastructure; federal government should strengthen DCC "
                          "coordination with budget and technical assistance.",
    "Tar_loc_government": "Inter-municipal collaboration proposed with 2nd, 3rd and 4th "
                          "neighbouring municipalities for joint waste/plastic management; "
                          "municipal Environment and Disaster Management Unit is the focal "
                          "point for programmes.",
    "Tar_students": "School-based green clubs are being established; Kathmandu University "
                    "students have been mobilised through internships and professor "
                    "collaborations.",
    "Tar_private_sector": "\"Private sector... we need to collaborate... they have a "
                          "management system, they are technically sound\"; waste-to-energy "
                          "company and PPP/contract/MoU models are central to future plans.",
    "Tar_civil_society": "NGOs and INGOs are key partners for this year's plastic "
                         "activities (\"the municipal government also wants to work "
                         "together with the NGO and the INGO for the plastic related\").",
    "Tar_households": "Household awareness, segregation and 3Rs are named as the first "
                      "priority (\"first of all we need a household awareness and "
                      "management system\").",
    "Tar_edu_institutions": "Kathmandu University and school green clubs are explicit "
                            "partners; one ward is being developed as a model for plastic "
                            "reduction with community education.",
    "Tar_private_companies": "Hotels and industries will be regulated by forthcoming "
                              "bylaws requiring their own wastewater treatment; private "
                              "waste-to-energy and recycling companies are sought as "
                              "implementation partners.",
    "Actor_role": "The mayor/municipal board formulates policy (green city priorities, "
                 "plastic budget, model ward, bylaws - role 1) and the municipality "
                 "directly implements programmes through its Environment and Disaster "
                 "Management Unit, green clubs, NGO partnerships and private-sector MoUs "
                 "(role 2).",
    "Discretion": "\"It's the sole right of the municipal government\" to allocate budget; "
                  "the municipal board decides policy and then an action plan is "
                  "formulated; the municipality sets \"the final price\" for plastic "
                  "sold by private actors and chooses among PPP, contract and MoU models.",
    "Monitoring": "Mayor: in 20 years as deputy mayor and now as mayor, ministry officials "
                  "have never come to monitor (\"I never saw the ministry people MONITOR "
                  "here\"); weak enforcement of micron restrictions implies absent "
                  "monitoring of markets.",
    "Financial_resources": "Waste management \"according to money... that was a lack of "
                           "budget\"; municipality has only a small plastic budget (~NPR "
                           "1.7 million plus NPR 10 lakh) and must seek federal/provincial/"
                           "international support for larger projects.",
    "Research": "LAPA project with Indian colleagues \"developing one technology to reduce "
               "plastic\"; a Detailed Project Report (DPR) for Dhulikhel's waste-management "
               "system has been prepared.",
    "Infrastructure": "\"Technically in Nepal we don't have technical machines and other "
                      "things\"; no local plastic recycler (Dhopur Recycler absent); "
                      "landfill site exists but lacks technical management knowledge; "
                      "waste-to-energy infrastructure planned via private partner.",
    "Capacity": "\"We needed more manpower and technical assistance\"; two types of "
               "capacity needed - community/household awareness and segregation, and "
               "technical personnel directly working in the field.",
    "Enforcement": "20/40-micron restrictions exist \"but if you want to go to the market "
                   "area, there are many kinds of plastic\"; mayor's team notes policies "
                   "exist on paper but implementation remains weak.",
    "Pol_epr": "Extended producer responsibility is not mentioned anywhere in this "
              "interview.",
    "Pol_awareness": "Green clubs (school- and community-based), community collaboration "
                     "with social workers, and household-awareness programmes are active "
                     "or planned policy tools.",
    "Pol_education": "Two types of capacity/education named: household awareness and "
                     "segregation, and technical education for personnel; model ward "
                     "action-oriented programme launched this year.",
    "Pol_RD": "LAPA project developing technology to reduce plastic with Indian "
              "colleagues working in Dhulikhel.",
    "Pol_ban": "Earlier government notification restricting 20-micron and 40-micron "
              "plastic use (\"I think 10 years ago\").",
    "Pol_subsitutes": "Municipality is \"finding alternative ways\" and making \"new goods "
                      "from the plastics\"; green-city and carbon-neutrality policies "
                      "frame broader substitution/reduction efforts.",
    "Pol_upcycling": "Women's groups and handicraft producers repurpose plastic into "
                     "products; municipality aims to collect plastic and \"make new goods.\"",
    "Pol_recycling": "This year's budget focuses on \"reducing the plastic, recycling the "
                     "plastic, and reuse the plastic\" in collaboration with NGOs and the "
                     "LAPA project.",
    "Pol_waste_collection": "Plastic collection through green-club/community initiatives "
                            "and NGO partnerships is part of the current year's policy and "
                            "budget.",
    "Pol_effectiveness": "Existing micron restrictions are widely ignored in markets; "
                         "Solid Waste Management Commission has collapsed; national "
                         "government lacks action-oriented programmes; ministry never "
                         "monitors local implementation despite approving acts.",
    "Pol_effectiveness_example": "See Pol_effectiveness - quoted directly in the "
                                 "Coded_Data cell.",
    "Sol_lead_agency": "Environment and Disaster Management Unit of the municipality is "
                       "the focal point; \"policy level we are deciding from our municipal "
                       "board. Then after you will find out the action plan.\"",
    "Sol_responsibilities": "Calls for clearer procedures/bylaws for managing, selling, "
                            "producing and recycling plastic; federal government should "
                            "provide technical support while respecting constitutional "
                            "local autonomy over sanitation.",
    "Sol_awareness": "Household awareness and 3Rs/segregation named as the most important "
                     "first step.",
    "Sol_segregation": "\"Segregating the household waste\" and household management "
                       "system proposed before broader waste solutions.",
    "Sol_upcycling": "Scale up women's-group/handicraft plastic repurposing with clearer "
                     "procedures and bylaws.",
    "Sol_recycling": "Continue and expand NGO/municipal collaboration on reducing, "
                     "recycling and reusing plastic; learn from Hetauda's provincially "
                     "supported model.",
    "Sol_education": "Both household/community education and technical training for waste "
                     "personnel are proposed.",
    "Sol_capacity": "More manpower and technical capacity (community awareness plus "
                    "technical personnel); federal technical equipment and budget support.",
    "Sol_RD": "Continue LAPA technology development for plastic reduction.",
    "Sol_finance": "Municipal plastic budget (~NPR 1.7 million + NPR 10 lakh); seek "
                   "federal/provincial/international funding for larger projects; "
                   "strengthen DCC with budget for inter-municipal cooperation.",
    "Sol_infrastructure": "Waste-to-energy (CNG/biogas) via private partner (DPR "
                           "prepared); improved landfill technical management; shared "
                           "regional facilities with neighbouring municipalities.",
    "Sol_subsitutes": "Green-city initiative, alternative product creation from collected "
                      "plastic, and promotion of non-plastic alternatives through green "
                      "clubs.",
    "Sol_enforcement": "Implied need to enforce existing micron restrictions and new "
                       "bylaws once prepared (markets currently sell all plastic types).",
    "Sol_monitoring": "Federal ministry should monitor and support, not only approve acts; "
                      "DCC could monitor/coordinate inter-municipal cooperation if given "
                      "mandate and resources.",
    "Traditions to build on (free-hand)": "\"Some women groups and some handicrafts "
                                          "products people are using plastic for some "
                                          "goods\" - a small but existing local practice "
                                          "of repurposing plastic that could be built on.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_science": "NA - no researchers or scientific institutions are described as "
                   "currently responsible for mitigating plastic pollution; Kathmandu "
                   "University is coded under Res_edu_institutions.",
    "Res_households": "NA - households are discussed as needing awareness/education and "
                      "as contributors to littering, not as currently responsible actors "
                      "for mitigation.",
    "Cul_prov_government": "NA - provincial government is mentioned as prioritising the "
                             "issue (Hetauda project) but not blamed for the current "
                             "plastic-pollution situation.",
    "Cul_students": "NA - students are not mentioned as contributing to the problem.",
    "Cul_private_sector": "NA - private waste actors are presented as needed partners, "
                          "not as culprits.",
    "Cul_civil_society": "NA - NGOs and social organisations are presented positively as "
                         "collaborators.",
    "Cul_science": "NA - not mentioned as contributing to the problem.",
    "Cul_edu_institutions": "NA - Kathmandu University is presented as a supportive "
                             "partner, not a culprit.",
    "Cul_private_companies": "NA - private companies (hotels, manufacturers) are "
                             "discussed as partners or regulated entities, not as "
                             "polluters.",
    "Tar_prov_government": "NA - provincial government is mentioned as a potential "
                           "funding partner (Hetauda example) but not as a primary target "
                           "group for mitigation measures.",
    "Tar_civil_society": "NGOs and INGOs are named as key partners and focus areas for "
                         "this year's plastic activities (\"the municipal government also "
                         "wants to work together with the NGO and the INGO for the plastic "
                         "related\").",
    "Tar_science": "NA - not mentioned as a target group for future measures.",
}

EXPLANATIONS = {
    "NPL_1": doe_expl,
    "NPL_2": ganesh_expl,
    "NPL_3": ktm_expl,
    "NPL_4": dhulikhel_expl,
    "NPL_7": rsct_expl,
}

# ---------------------------------------------------------------------------
# 3b. MASTER INTERVIEW LIST (full 32-interview numbering supplied by the
#     research team). NPL_1, NPL_2, NPL_3, NPL_4 and NPL_7 have source
#     material (interview notes/transcripts) provided so far and are fully
#     coded above; all other IDs are included as placeholder rows (Country/
#     ID/Actortype/Notes filled in from the master list, every substantive
#     variable left as NA) so the workbook's structure and numbering match
#     the full study design. Placeholder rows must NOT be treated as
#     completed coding - they exist purely to reserve the correct ID/row
#     and are clearly flagged as such in their 'Notes' cell.
# ---------------------------------------------------------------------------
CODED_ROWS = {
    "NPL_1": doe,
    "NPL_2": ganesh,
    "NPL_3": ktm,
    "NPL_4": dhulikhel,
    "NPL_7": rsct,
}

# (id, label as given by the research team, affiliation, inferred Actortype)
MASTER_LIST = [
    ("NPL_1", "1. Environmental Department/ Deepak Diwal", "Department of Environment", "nat_government"),
    ("NPL_2", "2. Former Minister/ Ganesh Shah", "Former Minister, Government of Nepal", "nat_government"),
    ("NPL_3", "3. KTM Municipal Office", "Kathmandu Metropolitan City - Environment / Solid Waste Management Division", "loc_government"),
    ("NPL_4", "4. Mayor Dhulikhel/ Ashok Kumar Byanju Shrestha", "Dhulikhel Municipality (Mayor's Office)", "loc_government"),
    ("NPL_5", "5. Private Sector/ Doco Recyclers", "Doco Recyclers", "private_sector"),
    ("NPL_6", "6. Restaurant and Hotel Owner Dhulikhel_Combined", "Dhulikhel Hotel / Restaurant", "private_companies"),
    ("NPL_7", "7. Rural Self-Reliance Development Center/ Narayan Nirola", "Rural Self-Reliance Development Center", "civil_society"),
    ("NPL_8", "8. Urban Development Ministry/ Kamal Adhikar", "Ministry of Urban Development, Government of Nepal", "nat_government"),
    ("NPL_9", "9. Dhulikhel Ward no.1", "Dhulikhel Municipality - Ward No. 1", "loc_government"),
    ("NPL_10", "10. Dhulikhel Ward no.7", "Dhulikhel Municipality - Ward No. 7", "loc_government"),
    ("NPL_11", "11. Private Sector Contractor/ SafaUrja Utpadan", "Safa Urja Utpadan", "private_sector"),
    ("NPL_12", "12. Ministry for Urban Planning", "Ministry of Urban Development", "nat_government"),
    ("NPL_13", "13. Creasion", "Center for Research and Sustainable Development in Nepal", "civil_society"),
    ("NPL_14", "14. Hotel Silent Park/ Rajesh Aryal", "Hotel Silent Park", "private_companies"),
    ("NPL_15", "15. Kathmandu Metropolitan City/ Suna Maya Margen", "Kathmandu Metropolitan City - Environment Division / Public Education & Inspection", "loc_government"),
    ("NPL_16", "16. Kathmandu Ward no.6", "Kathmandu Metropolitan City - Ward No. 6", "loc_government"),
    ("NPL_17", "17. Kathmandu Ward no.7/ Mr. Romy Prasad Shrestha", "Kathmandu Metropolitan City - Ward No. 7", "loc_government"),
    ("NPL_18", "18. Nepal Waste Manager", "Nepal Waste Management", "private_sector"),
    ("NPL_19", "19. TAAN Trekking Association", "TAAN Trekking Association", "civil_society"),
    ("NPL_20", "20. Chitwan Environmental Officer", "Municipal Environmental Office", "loc_government"),
    ("NPL_21", "21. Dhulikhel Household no.1", "Household", "households"),
    ("NPL_22", "22. Dhulikhel Household no.2", "Household", "households"),
    ("NPL_23", "23. Dhulikhel Shopowner no.3", "Shopowner", "private_companies"),
    ("NPL_24", "24. Dhulikhel Household no.4", "Household", "households"),
    ("NPL_25", "25. Dhulikhel Household no.5", "Household", "households"),
    ("NPL_26", "26. Dhulikhel Restaurant no.6", "Restaurant", "private_companies"),
    ("NPL_27", "27. Dhulikhel Household out of town no.7", "Household", "households"),
    ("NPL_28", "28. Chitwan Ward", "Local Government (Municipality level)", "loc_government"),
    ("NPL_29", "29. Kathmandu Former UNEP member no.8", "Former UNDP & UNEP member", "civil_society"),
    ("NPL_30", "30. Kathmandu Household Ward23 no.9", "Household woman", "households"),
    ("NPL_31", "31. Kathmandu Shopowner Babinasa no.10", "Shopowner in Babinasa", "private_companies"),
    ("NPL_32", "32. Chisapani Tour Guide", "Local tour guide", "private_companies"),
]

NAMES = {iid: f"{label} - {affil}" for iid, label, affil, _ in MASTER_LIST}
# Fill in the fuller descriptive names for the five fully-coded interviews.
NAMES["NPL_1"] = ("Department of Environment - Deepak Diwali, Deputy Director, Pollution "
                   "Control (air & plastics), Kathmandu (Interview no. 5, 23 June 2025)")
NAMES["NPL_2"] = ("Former Minister, Government of Nepal - Ganesh Shah, Kathmandu "
                   "(21 June 2025)")
NAMES["NPL_3"] = ("Kathmandu Metropolitan City (KMC), Solid Waste Management Office - "
                   "Municipality Officer(s), Kathmandu (23 June)")
NAMES["NPL_4"] = ("Dhulikhel Municipality - Mayor Ashok Kumar Byanju Shrestha, "
                   "Dhulikhel (22 June 2025, Kathmandu interview no. 4)")
NAMES["NPL_7"] = ("Rural Self-Reliance Development Center / RSCT (savings-and-credit "
                   "cooperative network, urban plastics programme)")


def placeholder_row(iid, label, affiliation, actortype):
    row = blank_row()
    row.update({
        "Country": "Nepal",
        "ID": iid,
        "Actortype": actortype,
        "Notes": f"PLACEHOLDER ROW - not yet coded. Master list entry: \"{label}\", "
                 f"affiliation: \"{affiliation}\". No interview notes or transcript for "
                 f"this respondent have been provided to the coder yet; Actortype above is "
                 f"a best-guess classification based on the affiliation label only. All "
                 f"other variables are left as NA pending the actual interview material. "
                 f"Do not interpret the NA values in this row as findings - they simply "
                 f"reflect that this interview has not been coded yet.",
    })
    return row


ALL_ROWS = {}
for iid, label, affiliation, actortype in MASTER_LIST:
    if iid in CODED_ROWS:
        ALL_ROWS[iid] = CODED_ROWS[iid]
    else:
        ALL_ROWS[iid] = placeholder_row(iid, label, affiliation, actortype)

INTERVIEWS = [ALL_ROWS[iid] for iid, _, _, _ in MASTER_LIST]

# ---------------------------------------------------------------------------
# 3c. Columns for which an explanation must ALWAYS be recorded (per the
#     research team's explicit instruction), even when the coded value is
#     NA. For fully-coded interviews without a specific reason already
#     written above, and for all placeholder interviews, a clear fallback
#     explanation is generated below.
# ---------------------------------------------------------------------------
ALWAYS_EXPLAIN_COLUMNS = [
    "NPF_victims", "NPF_villains", "NPF_hero",
    "Res_nat_government", "Res_prov_government", "Res_loc_government", "Res_students",
    "Res_private_sector", "Res_civil_society", "Res_science", "Res_households",
    "Res_edu_institutions", "Res_private_companies",
    "Cul_nat_government", "Cul_prov_government", "Cul_loc_government", "Cul_students",
    "Cul_private_sector", "Cul_civil_society", "Cul_science", "Cul_households",
    "Cul_edu_institutions", "Cul_private_companies",
    "Tar_nat_government", "Tar_prov_government", "Tar_loc_government", "Tar_students",
    "Tar_private_sector", "Tar_civil_society", "Tar_science", "Tar_households",
    "Tar_edu_institutions", "Tar_private_companies",
    "Actor_role", "Discretion", "Capacity", "Pol_effectiveness", "Pol_effectiveness_example",
]

PLACEHOLDER_EXPLANATION = ("No interview notes or transcript have been provided for this "
                            "respondent yet, so this variable could not be coded and is left "
                            "as NA. This is a data-availability gap, not a coding judgement.")


def get_explanation(iid, var, value):
    expl = EXPLANATIONS.get(iid)
    if expl and var in expl:
        return expl[var]
    if iid not in CODED_ROWS:
        return PLACEHOLDER_EXPLANATION
    return f"NA - not addressed in this interview (no mention found for {var})."


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


# ---------------------------------------------------------------------------
# 4a. One distinct colour per interview, used to fill each interview's rows
#     in both Coded_Data and Coding_Explanations so interviews are easy to
#     tell apart visually. Colours are spread evenly around the HSL colour
#     wheel; fully-coded interviews get a slightly stronger tint, while
#     not-yet-coded placeholder interviews get a light grey-tinted version
#     so they read as visually "unfinished".
# ---------------------------------------------------------------------------
import colorsys


def make_color_map(ids, coded_ids):
    n = len(ids)
    colors = {}
    for i, iid in enumerate(ids):
        hue = i / n
        if iid in coded_ids:
            # Fully-coded interviews: clearly saturated, distinct colours.
            r, g, b = colorsys.hls_to_rgb(hue, 0.78, 0.65)
        else:
            # Placeholder (not-yet-coded) interviews: very pale, low-saturation
            # tint of the same hue, so they read as "pending" at a glance.
            r, g, b = colorsys.hls_to_rgb(hue, 0.95, 0.35)
        colors[iid] = "{:02X}{:02X}{:02X}".format(int(r * 255), int(g * 255), int(b * 255))
    return colors


ALL_IDS = [iid for iid, _, _, _ in MASTER_LIST]
COLOR_MAP = make_color_map(ALL_IDS, set(CODED_ROWS.keys()))


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
    iid = interview["ID"]
    row_fill = PatternFill(start_color=COLOR_MAP[iid], end_color=COLOR_MAP[iid], fill_type="solid")
    ws2.append([interview.get(c, "NA") for c in COLUMNS])
    r = ws2.max_row
    for c in range(1, len(COLUMNS) + 1):
        cell = ws2.cell(row=r, column=c)
        cell.alignment = WRAP
        cell.border = BORDER
        cell.fill = row_fill
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
    expl = EXPLANATIONS.get(iid, {})
    row_fill = PatternFill(start_color=COLOR_MAP[iid], end_color=COLOR_MAP[iid], fill_type="solid")
    for var in COLUMNS:
        must_explain = var in ALWAYS_EXPLAIN_COLUMNS
        if must_explain or var in expl:
            value = interview.get(var, "NA")
            explanation = get_explanation(iid, var, value)
            ws3.append([iid, NAMES[iid], var, value, explanation])
            r = ws3.max_row
            for c in range(1, 6):
                cell = ws3.cell(row=r, column=c)
                cell.alignment = WRAP
                cell.border = BORDER
                cell.fill = row_fill
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
    ("This workbook applies the supplied codebook (see 'Codebook' sheet) to the Nepal "
     "interview list. 'Coded_Data' has one row for every ID in the research team's full "
     "32-interview master numbering (NPL_1-NPL_32), in that exact order. Column names and "
     "order follow exactly the variable list supplied by the research team.", False),
    ("", False),
    ("IMPORTANT - which rows are actually coded:", True),
    ("  Only 5 of the 32 rows are coded from real interview material (notes and/or "
     "transcripts) that has been provided so far: NPL_1, NPL_2, NPL_3, NPL_4 and NPL_7 (see "
     "below). The remaining 27 rows are PLACEHOLDERS: they reserve the correct ID, Country "
     "and a best-guess Actortype (inferred only from the short affiliation label supplied "
     "by the research team, e.g. \"Household\" -> households), but every substantive "
     "variable is left as NA because no interview notes/transcript for that respondent has "
     "been supplied yet. Each placeholder row's 'Notes' cell is explicitly marked "
     "\"PLACEHOLDER ROW - not yet coded\", and in 'Coding_Explanations' every one of the "
     "columns the research team asked to always document (see below) carries an explicit "
     "explanation stating that no source material is available yet - this is a "
     "data-availability gap, not a finding of \"no\" or \"not mentioned\". Please send the "
     "remaining interview notes/transcripts to have those rows properly coded.", False),
    ("", False),
    ("Fully-coded interviews:", True),
    ("  1. NPL_1 - Department of Environment (DoE), Deepak Diwali, Deputy Director, "
     "Pollution Control (air & plastics), Kathmandu. Labelled \"Interview no. 5\" in the "
     "project's internal running order (23 June 2025), but coded here as NPL_1 per the "
     "research team's master numbering. Coded from both the interview notes and the full "
     "verbatim transcript.", False),
    ("  2. NPL_2 - Former Minister, Government of Nepal, Ganesh Shah, Kathmandu (21 June "
     "2025). Coded from both the interview guideline notes and the full verbatim "
     "transcript.", False),
    ("  3. NPL_3 - KTM Municipal Office / Kathmandu Metropolitan City - Environment / Solid "
     "Waste Management Division, unnamed municipal officer(s), Kathmandu (23 June). Coded "
     "from both the interview guideline notes and the full verbatim transcript. Note: the "
     "Q3a bullet notes in the source document describing a multi-city (Kathmandu/Pokhara) "
     "dry-waste operation with UNDP training and a Bhaktapur awareness campaign appear to "
     "belong to a different, unrelated interview accidentally included in the same file, "
     "and were excluded from this row's coding (see its 'Notes' cell for detail). This "
     "interview was numbered NPL_4 in an earlier version of this workbook; it has been "
     "renumbered to NPL_3 per the research team's master list.", False),
    ("  4. NPL_4 - Mayor of Dhulikhel Municipality, Ashok Kumar Byanju Shrestha, "
     "Dhulikhel (22 June 2025, 2:55pm-3:40pm; Kathmandu interview no. 4). Coded from "
     "both the interview guideline notes and the full verbatim transcript. A member of "
     "the mayor's team also contributed additional remarks on policy fragmentation and "
     "constitutional responsibilities.", False),
    ("  7. NPL_7 - Rural Self-Reliance Development Center (RSCT), a savings-and-credit "
     "cooperative network (est. 1991) that recently expanded into urban plastics work; per "
     "the master list, the interviewee is Narayan Nirola. Coded from the hand-written "
     "interview notes only (no transcript available). This interview was numbered NPL_1, "
     "then NPL_2, then NPL_3 in earlier versions of this workbook as other interviews were "
     "added; it has now been renumbered to NPL_7 per the research team's master list.", False),
    ("", False),
    ("Placeholder (not yet coded) interviews: NPL_5, NPL_6, NPL_8-NPL_32 - see the "
     "master numbering list below and each row's 'Notes' cell in 'Coded_Data'.", False),
    ("", False),
    ("Sheets in this workbook:", True),
    ("  - Codebook: the variable dictionary, listed in the exact same order as the columns "
     "in 'Coded_Data'.", False),
    ("  - Coded_Data: wide-format matrix, one row per interview (NPL_1-NPL_32, in order), "
     "one column per codebook variable, including a free-hand 'Notes' column. Each "
     "interview's row is filled with a distinct colour (see 'Colour key' below) so "
     "interviews are easy to tell apart at a glance.", False),
    ("  - Coding_Explanations: long-format table giving the quote/observation (or, for "
     "placeholder interviews, the reason no evidence exists yet) behind each coded value, "
     "one row per Interview x Variable. As requested, the following columns ALWAYS have an "
     "explanation row for every interview, even when the value is NA: NPF_victims, "
     "NPF_villains, NPF_hero, all Res_/Cul_/Tar_ actor-grid columns, Actor_role, "
     "Discretion, Capacity, Pol_effectiveness and Pol_effectiveness_example. Other columns "
     "are documented only where a specific, non-obvious explanation was written for a "
     "fully-coded interview. Rows are colour-matched to the same interview colour used in "
     "'Coded_Data'.", False),
    ("", False),
    ("Colour key (interview -> row colour): fully-coded interviews (NPL_1, NPL_2, NPL_3, "
     "NPL_4, NPL_7) use a stronger/brighter tint of their colour; the 27 not-yet-coded "
     "placeholder interviews use a paler, greyed-down tint of their colour so they are "
     "distinguishable as \"pending\" at a glance. Colours are spread evenly across the "
     "colour wheel in master-list order (NPL_1 through NPL_32) so adjacent IDs are easy to "
     "tell apart.", False),
    ("", False),
    ("Full master numbering supplied by the research team:", True),
] + [
    (f"  {iid.replace('NPL_', '')}. {label.split('. ', 1)[1] if '. ' in label else label} "
     f"-> {iid} -> {affil}", False)
    for iid, label, affil, _ in MASTER_LIST
] + [
    ("", False),
    ("Coding conventions:", True),
    ("  - 'yes' / 'no' / 'NA': for the 5 fully-coded interviews, NA means the topic was not "
     "addressed in that interview (no evidence either way); 'no' is only used where the "
     "interviewee explicitly said the item is not an issue / not in place (e.g. Pol_epr = "
     "no for all five coded interviews where EPR was discussed or clearly absent). For the "
     "27 placeholder interviews, NA means no source material has been provided yet (see "
     "above).", False),
    ("  - Free-hand fields (Impacts, NPF_victims/villains/hero, Discretion, "
     "Pol_effectiveness_example, 'Traditions to build on') contain short descriptive text "
     "derived directly from the interview content, or 'NA' if not discussed/not yet coded.", False),
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
     "evidence for the DoE interview (NPL_1), not for RSCT (NPL_7), and coded accordingly.", False),
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
print(f"Interviews: {len(INTERVIEWS)} (coded: {len(CODED_ROWS)}, placeholder: {len(INTERVIEWS) - len(CODED_ROWS)})")
