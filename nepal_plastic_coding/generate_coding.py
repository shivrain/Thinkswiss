"""
Generates a codebook-based qualitative coding of the Nepal plastic-pollution
governance interview set into a single Excel workbook, strictly following the
exact column list supplied by the research team (Country, ID, Actortype, all
Problem_*, Impacts, NPF victims/villains/hero, Governance/coordination
variables, Res_/Cul_/Tar_ actor grids, Actor_role, Discretion,
implementation-issue variables, Pol_* / Sol_* variables, Traditions to build
on, and Notes).

The workbook contains one row per fully coded interview (currently NPL_1
through NPL_15). Unfilled placeholder rows are not included. Interviews
coded: NPL_1 (Department of Environment), NPL_2 (Former Minister Ganesh
Shah), NPL_3 (KTM Municipal Office), NPL_4 (Mayor of Dhulikhel
Municipality), NPL_5 (Doco Recyclers), NPL_6 (Dhulikhel Hotel/Restaurant
Owner), NPL_7 (Rural Self-Reliance Development Center / RSCT), NPL_8
(Ministry of Urban Development), NPL_9 (Dhulikhel Ward Office), NPL_10
(Dhulikhel Ward Chairman), NPL_11 (Safa Urja Utpadan), NPL_12 (Ministry for
Urban Planning / MoUD), NPL_13 (Creasion), NPL_14 (Hotel Silent Park /
Rajesh Aryal), NPL_15 (KTM Environment Inspector / Suna Maya Margen).

Sheets produced:
  1. Read_Me            - scope, sources, coding conventions, full master
                           numbering list, colour key
  2. Codebook            - variable dictionary matching the Coded_Data columns
  3. Coded_Data          - wide-format matrix: one row per coded interview
                           (NPL_1 through NPL_15, in order), one column per COLUMNS
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
    currently in place.
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

# --- NPL_7: Rural Self-Reliance Development Center (RSDC/RSCT) --------------
rsct = blank_row()
rsct.update({
    "Country": "Nepal",
    "ID": "NPL_7",
    "Actortype": "civil_society",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "air pollution (open burning of plastic, indoor air pollution from burning "
               "MLP wrappers), drainage/sewer blockage, landfills/dumping sites (designed "
               "for ~20 years but half-filled in ~3 years), freshwater rivers polluted "
               "(rivers used as dumping sites), microplastics (Chitwan lake; salt cited), "
               "health (burning impacts; PET bottles reused despite health warnings), "
               "aquatic life (fertility impacts from microplastics/contamination), persistent "
               "environmental degradation (plastics do not degrade easily)",
    NPF_VICTIMS_COL: "people who burn plastic at home (especially those who remain indoors "
                     "and suffer indoor air pollution), aquatic life and freshwater "
                     "ecosystems, communities relying on contaminated rivers/lakes, and the "
                     "general public who know plastic is harmful but do not grasp the full "
                     "extent of health and environmental risks",
    "NPF_villains": "plastic producers and manufacturers (highly influential — \"100%\" — "
                    "using money to influence policy); unlicensed private waste-collection "
                    "companies (operate without municipal registration/licence, earn from "
                    "household fees and recovered materials, influence local government — "
                    "e.g. Sawman under Neximac); local governments (do not implement own "
                    "policies such as segregated-waste fee reductions, freeze earmarked "
                    "environment budgets, lack political will, allow unlicensed collectors); "
                    "households/public (ignorant of extent of harm, NIMBY syndrome, mix "
                    "waste despite segregation options, reuse PET bottles against labels)",
    "NPF_hero": "RSDC/RSCT and its cooperative waste-collection model (Budhanilkantha pilot "
               "— cooperative collects segregated waste and benefits households); "
               "multi-stakeholder cooperation (government, private sector, NGOs); exposure "
               "visits that changed ward-chair attitudes; Extended Producer Responsibility "
               "(EPR) as a key proposed policy tool; leaders who \"genuinely care about "
               "people\"",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_loc_government": "yes",
    "Cul_private_sector": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",
    "Tar_civil_society": "yes",

    "Actor_role": "2,3",
    "Discretion": "RSDC independently designed and scaled its Swabalamban (self-reliance) "
                 "two-tier cooperative model since 1991, recently piloting the same approach "
                 "for urban WASH/plastics in Budhanilkantha — establishing cooperatives, "
                 "waste-segregation points, exposure visits, and household engagement without "
                 "waiting for government directives.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "no",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_ban": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"We always miss on implementation. That is our biggest "
                                 "challenge\"; Budhanilkantha policy to reduce the ~NPR 300 "
                                 "monthly waste-collection fee for households that segregate "
                                 "waste \"is not implemented\"; \"nobody is monitoring\" the "
                                 "20–40 micron plastic-bag rule (only occasional DoE checks); "
                                 "earmarked environmental budgets (e.g. NPR 1 crore) are "
                                 "\"frozen\" year after year because municipalities produce "
                                 "no action plan and do not spend them.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
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

    "Traditions to build on (free-hand)": "RSDC's Swabalamban two-tier cooperative savings-and-"
                                          "loan model (200+ cooperatives since 1991, governed "
                                          "by a general assembly) already adapted for urban "
                                          "plastic/waste management in Budhanilkantha; rural "
                                          "practice of reusing plastic bags many times (vs. "
                                          "urban single-use); long-term aspiration to revive "
                                          "indigenous practices such as bringing one's own cup "
                                          "when visiting others.",
    "Notes": "Master list label: \"7. Rural Self-Reliance Development Center/ Narayan "
             "Nirola\", affiliation: Rural Self-Reliance Development Center (RSCT/RSDC — "
             "the transcript uses \"Rural Self-Reliance Development Centre\" / RSDC). "
             "Interviewee: Narayan Nirola, project coordinator (Respondent 1); a second "
             "RSDC staff member with WASH/academic background also contributed substantially "
             "(Respondent 2). Internal project label: \"Interview no. 4\". Location: "
             "Kathmandu (interview); project work discussed in Budhanilkantha Municipality. "
             "Interview date: 23 June 2025 (12:52pm–1:30pm). Coded from both the interview "
             "guideline notes and the full verbatim transcript.",
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

# --- NPL_5: Doco Recyclers (Private Sector) -----------------------------------
doco = blank_row()
doco.update({
    "Country": "Nepal",
    "ID": "NPL_5",
    "Actortype": "private_sector",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",
    "Problem_import": "yes",

    "Impacts": "soil (plastic beneath farmland, reduced water/nutrient absorption), water "
               "(rivers polluted with plastic layers beneath soil), agriculture (decreased "
               "yield/productivity), air pollution (dioxin and furan from open burning), "
               "health (carcinogenic gases), microplastics (SAFEC study at Everest base "
               "camp cited), drainage/sewer blockage, dumping sites/landfills/leachate, "
               "aesthetics, wildlife/livestock (cows eating plastics in notes)",
    NPF_VICTIMS_COL: "farmers and agricultural productivity, human health (from toxic "
                     "burning emissions), rivers and aquatic ecosystems, soil, and society "
                     "as a whole (\"plastics is affected whole society\")",
    "NPF_villains": "plastic producers, importers and brand owners operating without "
                    "extended producer responsibility; the national government "
                    "(fragmented/uncoordinated policies, no apex body, weak enforcement "
                    "of the 40-micron ban); households and the wider public (talk about "
                    "pollution but take little action, burn waste including in government "
                    "backyards); informal scrap collectors who cherry-pick high-value "
                    "fractions and leave low-value/multi-layer plastics unmanaged",
    "NPF_hero": "Doco Recyclers and other private waste-management companies (MRFs, "
               "collection chains, consulting with municipalities); proposed national "
               "apex coordinating body; formal private-sector partners working with "
               "municipalities (e.g. KMC dry-waste inventory survey)",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_science": "yes",
    "Res_households": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_private_sector": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_students": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "2",
    "Discretion": "Doco Recyclers designs its own clientele-based collection model, "
                 "operates two MRFs with in-house sorting/processing flows, chooses "
                 "research and consulting partnerships (UNDP, GIZ, universities), and "
                 "develops municipal waste frameworks independently of government "
                 "directives.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_import": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_RD": "yes",
    "Pol_ban": "yes",
    "Pol_upcycling": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Even if we have... 40 micron has been banned... But if "
                                 "you go everywhere, you can find the plastics of below 40 "
                                 "micron\"; \"around 600 tonnes of plastic waste is generated "
                                 "in Nepal every day. But the management is going for hardly "
                                 "20 to 30 tonnes per day\"; \"fragmented policies that "
                                 "municipalities bring on themselves. If they cannot "
                                 "implement still, they have there in the policy\"; \"We were "
                                 "not engaged in this draft [Solid Waste Act]. They did not "
                                 "consult with us... this draft was incomplete and it's old "
                                 "fashioned.\"",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
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

    "Traditions to build on (free-hand)": "Using natural leaves (e.g. banana leaves) instead "
                                          "of plastic packaging, and carrying bags made from "
                                          "cotton or towels instead of plastic bags - "
                                          "traditional alternatives that could reduce plastic "
                                          "use while connecting to cultural heritage.",
    "Notes": "Master list label: \"5. Private Sector/ Doco Recyclers\", affiliation: Doco "
             "Recyclers (private waste-management/recycling company, est. 2017). "
             "Interviewee: Doco Recyclers (two respondents in the transcript - Respondent 1 "
             "and Respondent 2; individual names not given). Role: Private Sector. Location: "
             "Kathmandu Valley (MRFs at Sanothimi and Salghari/Bhaktapur; expanding to "
             "Pokhara). Interview date: 23 June 2025. Coded from both the interview "
             "guideline notes and the full verbatim transcript. Associations mentioned in "
             "the notes: Plastic Foundation Nepal and Solid Waste Management Association "
             "(likely the Solid Waste Management Association of Nepal). Data-quality note: "
             "the Q4a/Q4b blocks in the guideline notes (\"For ministry: What is new in "
             "the planned act?\" / ministry confidence questions) contain ministry-interview "
             "template content not spoken by Doco Recyclers in the transcript and were "
             "excluded from this row's coding; the transcript confirms Doco was not "
             "consulted on the draft Solid Waste Act but submitted online feedback calling "
             "it \"incomplete\" and \"old fashioned.\" The Q3a bullet content about "
             "Pokhara/UNDP training that was excluded from NPL_3 (KTM) belongs to this "
             "interview and is reflected here.",
})

# --- NPL_6: Dhulikhel Hotel/Restaurant Owner - Swasti Byanju ------------------
hotel = blank_row()
hotel.update({
    "Country": "Nepal",
    "ID": "NPL_6",
    "Actortype": "private_companies",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "environment (foul smell, visible accumulation, pollution of surroundings, "
               "aesthetics/litter, ~100-year degradation), health (limited detailed "
               "knowledge but hospital/medical waste including needles and syringes in "
               "dumping areas), agriculture (leachate affecting wards 2 and 3, downstream "
               "Panchkhal), water (drinking water no longer pure in affected areas; river "
               "dumping in Khurkot/Sindhuli road), aquatic life/fish (ocean plastic "
               "impacts cited), inter-municipal conflict (downstream Panchkhal affected)",
    NPF_VICTIMS_COL: "downstream communities and farmers (Panchkhal, wards 2 and 3 — "
                     "leachate, water and agriculture impacts), families and local "
                     "community, environment and aquatic life, and Dhulikhel Municipality "
                     "itself (major issues from the dumping site)",
    "NPF_villains": "national and municipal government (lack of strict policy/enforcement, "
                    "budget not allocated, dumping mismanaged for 15+ years, projects "
                    "announced but never followed up); the wider public (everyone is "
                    "aware but still uses plastic; awareness without behavioural change); "
                    "shops/markets still supplying goods in plastic despite ban discussions; "
                    "upstream dumping affecting downstream areas without regard for others",
    "NPF_hero": "Dankhuta's \"Waste into Money\" model (zero plastics, segregated "
               "buildings, park on former dump); ward 7 shopkeeper litter-collection rule "
               "keeping Dhulikhel streets clean; waste collectors who pick plastic; "
               "Chitwan large recycling unit; individuals making personal efforts (e.g. "
               "using alternative bags, household bag collection for reuse); Kathmandu "
               "mayor cited as example of political will improving waste management",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_households": "yes",
    "Tar_private_sector": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "4",
    "Discretion": "As a hotel owner the interviewee has limited discretion: personally uses "
                 "non-polythene shopping bags, consolidates items into fewer bags, collects "
                 "plastic bags at home for family reuse, and rolls shampoo sachets to reduce "
                 "waste, but at the hotel can only bag all waste together for municipal "
                 "collection — the only option available.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_ban": "yes",
    "Pol_clean_up": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Beat the Plastic\" on World Environment Day had mayor and "
                                 "KU student participation \"but in practice nothing has "
                                 "changed. Same work has been repeating\"; \"We keep hearing "
                                 "that a project will start soon — it's been 3 years already\"; "
                                 "\"When I go shopping... they are plastic\" for vegetables "
                                 "despite a ministry notice to ban plastic; \"waste here is not "
                                 "properly minimized, it's almost zero\" recycling in Dhulikhel; "
                                 "\"currently, it seems the budget itself is not being "
                                 "allocated.\"",

    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Thokodi — traditional newspaper paper bags made and "
                                          "sold in shops (practised in the past, now largely "
                                          "gone); paper bags distributed to households; cloth/"
                                          "cotton bags (respondent received an organic cloth bag "
                                          "as a gift and noted such bags should be made locally).",
    "Notes": "Master list label: \"6. Restaurant and Hotel Owner Dhulikhel_Combined\", "
             "affiliation: Dhulikhel Hotel / Restaurant. Interviewee: Mr. Swasti Byanju, "
             "hotel owner, Dhulikhel. Location: Dhulikhel. Interview date: 22 June 2025 "
             "(Kathmandu interview no. 3; transcript file labelled \"June 22 (Part 1 ENG)\"). "
             "Interviewers: PEGO, Ram Devi (translator/local), Deep, Swasti Byanju (local "
             "collaborator — also contributed the Dankhuta \"Waste into Money\" example in "
             "the transcript). Coded from both the interview guideline notes and the full "
             "verbatim transcript. The respondent is not personally aware of specific "
             "national plastic policies (\"No. I don't\") but describes ministry ban notices "
             "and municipal measures as relayed by the interview team.",
})

# --- NPL_8: Ministry of Urban Development (MoUD) ------------------------------
moud = blank_row()
moud.update({
    "Country": "Nepal",
    "ID": "NPL_8",
    "Actortype": "nat_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "rivers and canals (flow disturbed), drainage/sewer line blockage, "
               "landfills/dumping sites (Bancharedanda/Bonsai centre designed for 20 years "
               "but filled in 3–4 years), aesthetics/tourism, health (indirect microplastics "
               "via livestock; meat, blood and liquids carried in black plastic bags), "
               "agriculture (productivity decreasing), leachate/stench/flies at dumps, "
               "inter-municipal/transboundary social conflict",
    NPF_VICTIMS_COL: "households and the general public (public-health risks from carrying "
                     "meat, blood and liquids in low-quality black plastic; customary "
                     "practices now harmful under higher population density); communities "
                     "near landfill/dumping sites (\"weak people\" facing health hazards, "
                     "stench, flies, leachate); farmers/agriculture (productivity "
                     "decreasing due to plastic); tourism and urban aesthetics (old cities "
                     "and emerging towns facing nuisance); rivers and aquatic environments",
    "NPF_villains": "the general public/households (convenience-driven plastic use, no "
                    "source-segregation culture, indiscriminate disposal, historically "
                    "throwing waste from doors/windows while expecting municipalities to "
                    "clean up); plastic producers and the wider manufacturing/packaging "
                    "industry (\"whole business\" — polluter's-pay principle exists but "
                    "implementation is weak); municipalities that operate sanitary landfills "
                    "as open dumping sites (Bancharedanda example deterring new sites) and "
                    "dump waste across municipal boundaries; the historical policy "
                    "stand-off framed as a \"war between people and policy makers\"",
    "NPF_hero": "Ministry of Urban Development / central government (drafting the new Solid "
               "Waste Management Act with EPR, PPP, burning ban, three-tier roles); "
               "provincial/central support for landfill land acquisition; private sector "
               "under PPP with sample municipal agreements; MoUD training centre and "
               "planned municipal software system; circular-economy and source-segregation "
               "approach; NGOs/researchers/experts invited to co-create policy",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_science": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_households": "yes",
    "Cul_loc_government": "yes",
    "Cul_private_companies": "yes",

    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",
    "Tar_edu_institutions": "yes",

    "Actor_role": "1",
    "Discretion": "MoUD (cabinet-delegated authority) drafts the Solid Waste Management Act, "
                 "regulations, PPP sample agreements, plastics guidelines, municipal "
                 "monitoring software, and coordinates international donor partners — "
                 "formulation/policy discretion at national level.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_ban": "yes",
    "Pol_capacity": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Pre-2015 Solid Waste Management Act \"weaknesses in the "
                                 "federal governance context\"; no specific national plastic "
                                 "policy despite plastic being \"hazardous waste\" in "
                                 "classification; Bancharedanda sanitary landfill operated "
                                 "as a dumping site (20-year design filled in 3–4 years); "
                                 "polluter's-pay principle used but \"weak in implementation\"; "
                                 "respondent could not cite any assessment of the existing "
                                 "40-micron ban notification.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Historical pet-to-khet (stomach-to-field) "
                                          "sustainable cycle between consumption and land; "
                                          "leaf plates instead of plastic; customary organic "
                                          "waste practices that were harmless at low "
                                          "population density; emerging household practice "
                                          "of separate decomposable and non-decomposable "
                                          "dustbins; new generation already disposing waste "
                                          "properly in dustbins.",
    "Notes": "Master list label: \"8. Urban Development Ministry/ Kamal Adhikar\", "
             "affiliation: Ministry of Urban Development (MoUD), Government of Nepal. "
             "Internal project label: \"Interview 2\". Interview date: 23 June 2025 "
             "(10:55am–12:00pm). Participants: Kamal Adhikar (Senior Sociologist), "
             "Nawaraj/Nawrag (Joint Secretary, 29 years in service), Senior Division "
             "Engineer, and additional ministry staff (multiple voices in transcript: "
             "UB, UB2, UB3–UB6). Coded from both the interview guideline notes and the "
             "full verbatim transcript. NPF fields: victims, villains and hero are not "
             "named using explicit NPF narrative labels in the interview, but all three "
             "are clearly implied through attributed harms, blamed actors and proposed "
             "problem-solvers (see Coding_Explanations). Draft Solid Waste Management Act "
             "on MOUD website for consultation; separate plastics-only policy planned "
             "after the act. \"EPI\" in transcript treated as EPR (extended producer "
             "responsibility). Bonsai/Bancharedanda landfill name varies in transcript.",
})

# --- NPL_9: Dhulikhel Municipality - Ward Office (Ward No. 2 per transcript) --
ward2 = blank_row()
ward2.update({
    "Country": "Nepal",
    "ID": "NPL_9",
    "Actortype": "loc_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "agriculture (plastic in farmlands, severely affected farmland fertility in "
               "Ward No. 2), rivers and canals (blockages), health (plastic bottles; open "
               "syringes along Panchkhal stream; medical/hazardous waste mismanagement), "
               "lower settlements (dumping-site runoff due to topography), lifestyle "
               "(everything packaged in plastic)",
    NPF_VICTIMS_COL: "farmers and residents of Wards No. 2 and 3 (most affected by the Ward "
                     "No. 8 dumping site; farmland fertility loss); lower settlements "
                     "downstream of the dumping site; the general public/households "
                     "(health risks from plastic use and waste, including syringe scavenging); "
                     "rivers and agricultural environments in Ward No. 2",
    "NPF_villains": "the general public/households (lack of awareness, casual plastic use, "
                    "not following syringe-collection guidance); factories (resisted the "
                    "micron-thickness ban, no follow-through); hotels, shops and restaurants "
                    "(prioritised quick profits over cloth-bag and anti-plastic-container "
                    "campaigns); the municipality (collects annual cleaning taxes but weak "
                    "execution; no segregation in collection vehicles; unclear allocation of "
                    "funds across 12 wards); people who collect syringes and other waste "
                    "materials to sell",
    "NPF_hero": "the ward office as first contact point and coordinator (Tol Sudhar Samiti "
               "pilot projects, Environment Day rallies, awareness activities with limited "
               "ward funds); Dhulikhel Hospital/nursing college (cloth-bag campaign — ~1,000 "
               "bags to hotels/shops); former ETPC NGO (household waste collection); "
               "Tribhuvan University (market-area development collaboration); researchers/"
               "universities invited to support awareness and continue efforts",

    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_science": "yes",
    "Res_households": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_civil_society": "yes",
    "Tar_science": "yes",
    "Tar_nat_government": "yes",

    "Actor_role": "3",
    "Discretion": "The ward can coordinate grassroots activities through Tol Sudhar Samiti "
                 "(~100–150 households per committee), run awareness activities from limited "
                 "ward funds, act as first contact for NGO/researcher collaborations, and "
                 "report problems to the municipality — but has no separate waste-management "
                 "budget, no authority over municipal collection trucks, and no manpower or "
                 "materials to segregate or manage waste independently.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_tax": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_clean_up": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "\"Policies exist only on paper; they are not implemented\"; "
                                 "\"The local government collects annual cleaning taxes from us "
                                 "yearly, but there's no effective execution\"; \"The government "
                                 "tried to make a rule but couldn't enforce it\" (micron ban "
                                 "introduced ~3–4 years ago but factories resisted); \"No, "
                                 "everything collected in the vehicles goes together\" — no "
                                 "source segregation in practice.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_clean_up": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Baskets made from paper or bamboo; meat wrapped "
                                          "in straw bundles; large reusable cloth bags for "
                                          "carrying goods; earlier practice of carrying items "
                                          "in reusable bags before plastic became common.",
    "Notes": "Master list label: \"9. Dhulikhel Ward no.1\", affiliation: Dhulikhel "
             "Municipality - Ward No. 1. However, the transcript consistently identifies the "
             "respondents' focal ward as Ward No. 2 (e.g. plastic in farmlands and rivers in "
             "Ward No. 2; Wards 2 and 3 most affected by dumping site in Ward No. 8). "
             "Interviewees: Ward Head (name not given), E, Deep and F (ward/municipal staff; "
             "exact roles not fully identified). Interviewers: Ram Devi (Kathmandu University "
             "researcher) and PEGO project team. Location: Dhulikhel. Coded from both the "
             "interview guideline notes and the full verbatim transcript. NPF fields: victims, "
             "villains and hero are not named using explicit NPF narrative labels but are "
             "clearly implied (see Coding_Explanations). Hospital collaboration refers to "
             "Dhulikhel Hospital; ETPC = Environment Tourism Committee (former NGO).",
})

# --- NPL_10: Dhulikhel Municipality - Ward Chairman (Ward No. 3 per transcript) -
ward7 = blank_row()
ward7.update({
    "Country": "Nepal",
    "ID": "NPL_10",
    "Actortype": "loc_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "rivers and riverbanks (Panchkhal area; Bagmati waste through pipelines; "
               "\"no good river in Nepal\"), smell affecting residents daily (including near "
               "the ward chairman's house), ecosystem described as \"already destroyed\", "
               "dust and visual pollution (\"everywhere is dusty\"), tourism/national image "
               "(embarrassment compared with Switzerland), waste still reaching rivers despite "
               "collection efforts",
    NPF_VICTIMS_COL: "local residents and downstream communities (daily smell, Panchkhal "
                     "negative impacts); rivers and the broader environment/ecosystem; Nepal's "
                     "national image (\"people call Nepal a beautiful country, like Switzerland, "
                     "but everywhere is dusty\")",
    "NPF_villains": "households/residents who leave plastic outside despite distributed "
                    "containers, throw waste in forests, and fail to cooperate with "
                    "segregation; the central/federal government (\"big gap\" with local "
                    "government, slow progress, no master plans, draft act not reaching "
                    "wards); pervasive plastic use in food packaging and single-use items "
                    "without affordable systemic alternatives",
    "NPF_hero": "the ward chairman personally (model zero-waste household behaviour, picks "
               "up litter, organic diet); Dhulikhel Municipality (segregating plastic and "
               "composting organic waste, transfer-station/landfill plan passed, contracting "
               "waste collection); Bhutan cited as an awareness example; researchers to find "
               "plastic alternatives; 2–4 workers separating recyclables at the dumping/refining "
               "site; international donors on Bagmati (World Bank, ADB)",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_households": "yes",
    "Res_science": "yes",

    "Cul_nat_government": "yes",
    "Cul_households": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_households": "yes",
    "Tar_private_sector": "yes",
    "Tar_science": "yes",

    "Actor_role": "3",
    "Discretion": "As elected ward chairman (32 years, five terms) he supervises dumping sites, "
                 "ran door-to-door container programmes, coordinates directly with the "
                 "municipality on contractor arrangements, and requests land from the Nepal "
                 "government — but has no formal ward-level waste-management staff, no separate "
                 "ward budget (municipality contracts fund collection), and processing-company "
                 "agreements were terminated.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_clean_up": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Door-to-door household containers \"wasn't enough\" because "
                                 "people left plastic outside; processing-company agreements "
                                 "\"were terminated\"; despite Sunday plastic collection and "
                                 "refining (Rs 10–15/kg), \"even if there's a collection problem "
                                 "here, it still ends up in the river\"; \"now there's a big gap "
                                 "between the local and federal governments — they keep making "
                                 "plans and talking, but progress is very slow.\"",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_clean_up": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Baskets made from natural materials; paper bags "
                                          "used for vegetables; cloth bags encouraged; "
                                          "biodegradable vs. non-biodegradable household "
                                          "separation promoted in recent 1–2 year efforts.",
    "Notes": "Master list label: \"10. Dhulikhel Ward no.7\", affiliation: Dhulikhel "
             "Municipality - Ward No. 7. Transcript identifies Ward No. 3 (land-ownership "
             "request, transfer-station planning in Thakuri village) and Panchkhal-area "
             "impacts. Interviewee: Ward Chairman/Ward Head (name not given), 32 years as "
             "ward chairman (five terms). Also present: Ram Devi (Kathmandu University), "
             "Deep and F. Location: Dhulikhel. Coded primarily from the full verbatim "
             "transcript. DATA-QUALITY NOTE: the structured theme notes at the top of the "
             "source document (Bio-Camp pilot, Plast Foundation Nepal, Solid Waste Management "
             "Association, EPR nationwide, microplastics in fertilisers, Environmental "
             "Protection Act without plastics policy, etc.) do not appear in this ward-head "
             "transcript and were excluded as likely belonging to a different interview "
             "accidentally bundled in the same file. NPF fields implied, not explicit.",
})

# --- NPL_11: Safa Urja Utpadan (Private Sector Contractor) --------------------
safaurja = blank_row()
safaurja.update({
    "Country": "Nepal",
    "ID": "NPL_11",
    "Actortype": "private_sector",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "environment and land (improper disposal, long-term plastic impacts), "
               "agriculture (monsoon chemicals from plastic reach crops; soil pH and "
               "fertility affected; repeated tilling brings plastic to surface), health "
               "(winter home burning of plastics — smoke and carbon emissions), landfill "
               "pressure (Ratnanagar ~21 tractor loads/day; only small unusable fraction "
               "sent after company processing)",
    NPF_VICTIMS_COL: "farmers and agricultural land (soil pH/fertility loss from plastic "
                     "chemicals in monsoon); the environment and land; public health "
                     "(smoke/carbon from winter burning); future generations (\"breathe "
                     "clean air and consume safely produced food\")",
    "NPF_villains": "the general public/households (low awareness, careless disposal, night "
                    "dumping in bazaars, burning plastics at home in winter, resistance to "
                    "fines and fee increases, mentality that waste management is the "
                    "company's job); the state/government (\"we have not seen improvements "
                    "from the state\"; system not formalized); hotels/bazaar businesses "
                    "that dump waste at night and resist proper fees",
    "NPF_hero": "Safa Urja Utpadan (collection, segregation, processing, GPS/barcode "
               "tracking, awareness programmes, paying NPR 6.6M+ annually to "
               "municipalities); school students who influence parents; planned PPP model "
               "across Ratnanagar, Khaireni and Kalika municipalities; paper alternatives "
               "for feast plastics",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_households": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_households": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_students": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "2",
    "Discretion": "Safa Urja designs its own collection and segregation system (hazardous/"
                 "food/non-harmful/plastic/textile streams), sets differentiated household/"
                 "business fee schedules (NPR 25–300/month by category), deploys GPS-enabled "
                 "vehicles and a planned barcode payment-tracking system, conducts awareness "
                 "programmes on its own schedule (schools every 3 months; Sauraha tourism zone "
                 "every 5–6 months), and negotiates PPP/TOR development across three "
                 "municipalities — substantial organisational discretion within a loss-making "
                 "municipal contract.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_tax": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Fines up to NPR 11,000 exist on paper but \"can only be "
                                 "applied in certain areas; they cannot be enforced "
                                 "everywhere\"; fee-increase proposals \"aren't passing\"; "
                                 "rates unchanged for seven years (only 10%/year added); "
                                 "\"we have not seen improvements from the state\"; household "
                                 "segregation compliance inconsistent despite company sorting "
                                 "at facility.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Some traditional knowledge mentioned but \"no "
                                          "policy exists to reduce plastic\"; plan to replace "
                                          "plastic plates/cups/bowls at feasts with paper "
                                          "alternatives; jute sacks/bags for purchased-plastic "
                                          "disposal proposed.",
    "Notes": "Master list label: \"11. Private Sector Contractor/ SafaUrja Utpadan\", "
             "affiliation: Safa Urja Utpadan. Interviewee: Safa Urja representative (name "
             "not given in transcript). Location: Khaireni, Chitwan district — operates "
             "across Ratnanagar, Khaireni and Kalika municipalities (16 wards; Sauraha "
             "tourism zone). Company operating ~2 years at time of interview; 70–80 staff; "
             "German operations manager; UK/US stakeholders. Interviewers: Ram Devi "
             "(Kathmandu University) and Deep. Coded from both the interview guideline notes "
             "and the full verbatim transcript. Landfill sites: Ratnanagar (Udaypur community "
             "forest; also Ratnanagar community forest near Ladhari/Kayar Khola cited); "
             "Khaireni (riverside below bridge). Pays municipality NPR 6.6M annually "
             "(increasing 10%/year). PPP model in planning across three municipalities. "
             "NPF fields implied, not explicit.",
})

# --- NPL_12: Ministry for Urban Planning / MoUD (policy formulation) ----------
moud12 = blank_row()
moud12.update({
    "Country": "Nepal",
    "ID": "NPL_12",
    "Actortype": "nat_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "rivers and canals (flow obstructed), sewer/water-line blockages, tourism "
               "(haphazard disposal), health (low-quality black plastic for meat/yogurt; "
               "public unaware of hazards), agriculture (productivity impacted), landfill/"
               "dumping sites (Banchare Danda designed 20 years but filled in 3–4 years; "
               "leachate, bad smells, visual disturbance), livestock consuming plastic, "
               "microeconomic/household-level damage, inter-municipal protests",
    NPF_VICTIMS_COL: "the general public/households (health risks from black plastic food "
                     "packaging; microeconomic damage); communities near landfill/dumping "
                     "sites (leachate, smell, protests); farmers/agriculture (productivity "
                     "loss); tourism and urban environments (old cities and emerging towns); "
                     "rivers, canals and livestock",
    "NPF_villains": "the general public (convenience attachment to plastic, no segregation "
                    "culture, haphazard disposal); plastic producers/industries (polluter-pays "
                    "principle exists but implementation weak; EPR not yet enacted); "
                    "municipalities operating sanitary landfills as open dumping sites "
                    "(Banchare Danda deterrent); inter-municipal waste dumping and NIMBY "
                    "conflicts",
    "NPF_hero": "Ministry of Urban Development / central government (draft Solid Waste "
               "Management Act with EPR, PPP, burning ban, three-tier roles); provincial/"
               "central assistance for landfill land acquisition; private sector (domestic "
               "and international) under PPP with sample agreements; MoUD training centre "
               "and municipal data software; source-segregation and circular-economy "
               "approach; international donor coordination",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_science": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_households": "yes",
    "Cul_loc_government": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_prov_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "1",
    "Discretion": "MoUD has authority to formulate policies, acts and regulations; drafts "
                 "the Solid Waste Management Act, post-enactment regulations/guidelines "
                 "(including 5 km/10 km distance management), PPP sample agreements, "
                 "community-level plastic guidelines, municipal monitoring software, and "
                 "coordinates international donor partner approvals.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_ban": "yes",
    "Pol_capacity": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Bagmati Cleaning Campaign running ~10 years but plastic "
                                 "remains the majority problem; Kathmandu source segregation "
                                 "only partial (\"most not in the desired form\"); polluter-"
                                 "pays principle in use but \"implementation is weak\"; "
                                 "Banchare Danda landfill designed for 20 years nearly filled "
                                 "within 3–4 years; no specific national plastic policy despite "
                                 "general solid waste framework.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "At low population density, waste was mostly "
                                          "organic and nature could absorb it; traditional "
                                          "leaf plates instead of plastic; natural consumption-"
                                          "land cycle (broken by plastic intervention); "
                                          "emerging household practice of separate decomposable "
                                          "and non-decomposable bins among younger generation.",
    "Notes": "Master list label: \"12. Ministry for Urban Planning\", affiliation: Ministry "
             "of Urban Development (MoUD), Government of Nepal. Interviewee name not given; "
             "role described as involved in formulating policies, acts and regulations for "
             "solid waste management and urban planning. Interviewer: PEGO project team. "
             "Coded from both the interview guideline notes and the full verbatim transcript. "
             "Substantively overlaps NPL_8 (also MoUD, 23 June 2025 session with named "
             "officials) but this appears to be a separate/cleaned transcript of the ministry "
             "policy perspective; coded independently from this source. Local Government "
             "Operation Act 2017 cited (NPL_8 referenced Self-Governance Act 2016). NPF "
             "fields implied, not explicit. Separate plastic waste management policy planned "
             "after Solid Waste Management Act.",
})

# --- NPL_13: Creasion (recycling / waste-management CSO) ----------------------
creasion = blank_row()
creasion.update({
    "Country": "Nepal",
    "ID": "NPL_13",
    "Actortype": "civil_society",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",

    "Impacts": "end-of-life plastic fate; micro- and nano-plastics (colored PET bottles "
               "cannot be effectively reused — only transparent PET can; colored PET emits "
               "significant microplastics when processed); health hazards from recycled "
               "plastics in food-contact applications; carcinogenic emissions from poor "
               "recycling sector practices (processing at 100–200°C); exploitation of "
               "informal waste pickers (especially women) at landfills/dumpsites",
    NPF_VICTIMS_COL: "informal waste pickers, particularly women selling PET from "
                     "landfills/dumpsites for NPR 10 (vs NPR 25/kg formal rate) and facing "
                     "intimidation/violence from intermediary \"mafias\"; the public exposed "
                     "to unsafe recycled food-contact plastics and carcinogenic emissions "
                     "from substandard recyclers; the environment from end-of-life plastics "
                     "and microplastic release",
    "NPF_villains": "intermediary \"mafias\" who threaten and assault waste pickers who do "
                    "not sell to them; substandard recyclers with \"pathetic\" social and "
                    "environmental standards (100–200°C processing, carcinogenic emissions); "
                    "the scrap tax on PET at municipal level (raising raw-material and granule "
                    "costs); virgin-plastic preference among Nepali consumers and businesses; "
                    "bureaucratic barriers (Department of Industry registration, 60% export "
                    "condition after year three, repeated LG approvals on relocation); national "
                    "government for not consulting recyclers on solid waste policy (improving "
                    "only gradually)",
    "NPF_hero": "Creasion (recycling, downsizing, product conversion, safety-tested granules, "
               "High-Effluent Treatment Plant recycling ~10,000 L water with no drainage "
               "outflow); EU-funded plastic-recovery and entrepreneurship project; CAP "
               "project (World Bank/UNOPS/SACEP) with sophisticated recycling technology; "
               "municipal collaboration to identify and formalize informal collectors; Plus "
               "Nepal and informal NRR associations; CSO trainings for sector coordination",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_loc_government": "yes",
    "Cul_households": "yes",
    "Cul_private_sector": "yes",

    "Tar_private_sector": "yes",
    "Tar_loc_government": "yes",
    "Tar_households": "yes",
    "Tar_civil_society": "yes",
    "Tar_private_companies": "yes",

    "Actor_role": "2",
    "Discretion": "Creasion designs its own recycling and product-conversion processes, "
                 "develops granule quality/safety standards and government-standard testing "
                 "before market release, runs EU and CAP projects, collaborates with "
                 "municipalities to identify informal collectors, and builds entrepreneur "
                 "capacity — but faces constrained discretion on bureaucratic setup (Dept of "
                 "Industry, LG approvals) and export/tax policy conditions.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_tax": "yes",
    "Pol_recycling": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Scrap tax on PET at municipal level raised raw-material and "
                                 "granule costs, making recycled products less competitive "
                                 "than virgin plastic; social/environmental standards across "
                                 "the broader recycling sector described as \"pathetic\" with "
                                 "weak enforcement; Creasion \"not consulted\" on solid waste "
                                 "management policy development (gradually being included); "
                                 "national plastic-generation estimates (60,000 t/year; 27% "
                                 "PET per World Bank) questioned as methodology unclear.",

    "Sol_lead_agency": "yes",
    "Sol_awareness": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "NA",
    "Notes": "Master list label: \"13. Creasion\", affiliation: Center for Research and "
             "Sustainable Development in Nepal (Creasion). Interviewee name not given. "
             "Organization: recycling and waste-management CSO supporting informal waste "
             "workers and entrepreneurs. Interview date: 23 June 2025 (~5:15pm). Coded from "
             "interview guideline notes and bullet-point field notes (no full verbatim "
             "transcript supplied). ~70 t/month PET capacity cited. Two known recycling "
             "centers in Nepal mentioned. Associations: Plus Nepal (formal); Nepal Recycler "
             "and Reuse/NRR (informal/unregistered). CAP project throughput figures in "
             "source (1 t/h input vs 10 t/h granules) appear inconsistent — retained as "
             "stated. NPF fields implied, not explicit.",
})

# --- NPL_14: Hotel Silent Park - Rajesh Aryal ---------------------------------
silent_park = blank_row()
silent_park.update({
    "Country": "Nepal",
    "ID": "NPL_14",
    "Actortype": "private_companies",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",

    "Impacts": "wildlife and domestic animals (ingest salty plastic snack wrappers); air "
               "pollution (burning plastics; poor communities use plastic as fire starters); "
               "microplastics contaminating soil, water, air, rivers, irrigation water and "
               "organic fertilisers (reducing soil moisture, potentially affecting crop "
               "yields); health (low-quality packaging may leach chemicals; unsafe bottled "
               "water/labelling practices); visual pollution and clogged drains/rivers; "
               "aluminum foil food packaging (greater disposal problem than plastic because "
               "recyclers do not collect it); tourism/safari area when bins overflow",
    NPF_VICTIMS_COL: "wildlife and domestic animals; poor communities exposed to burning "
                     "and fire-starter practices; farmers/agriculture (microplastics in soil "
                     "reducing moisture and crop yields); the general public (chemical "
                     "leaching, unsafe bottled water); the tourist/safari area and local "
                     "community when collection fails and waste spreads",
    "NPF_villains": "the general public/hotels (low awareness, laziness, convenience, free "
                    "plastic bags, perception that waste management is solely government's "
                    "responsibility); national government (no comprehensive plastics policy — "
                    "only general Environmental Protection Act; weak enforcement of <40-micron "
                    "ban; political instability delaying implementation); unclear "
                    "federal/provincial/local authority division; municipal collection system "
                    "(irregular Eco-Green and municipal trucks; bins overflow); hotel sector "
                    "(thousands of Chitwan hotels with no continuous waste-management action); "
                    "aluminum foil widely used without recycling route; heavy informal-sector "
                    "reliance without formal integration",
    "NPF_hero": "nature guides and park rangers (trained not to litter, collect wrappers on "
               "safari); hotel associations and NGOs (awareness training); Eco-Green and "
               "informal collectors (NPR 10–15/kg for bottles; informal sector collects ~99% "
               "of valuable plastics from hotels); PLEASE and Bio-Camp pilot projects "
               "(segregation, recycling/upcycling); Solid Waste Management Association and "
               "Plast Foundation Nepal; safari dustbin system; proposed EPR nationwide and "
               "ward-level designated street cleaners; elephant-dung cleaning teams as "
               "operational model (twice daily)",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_private_sector": "yes",
    "Res_civil_society": "yes",
    "Res_households": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_households": "yes",
    "Cul_loc_government": "yes",
    "Cul_private_companies": "yes",

    "Tar_nat_government": "yes",
    "Tar_loc_government": "yes",
    "Tar_private_sector": "yes",
    "Tar_households": "yes",
    "Tar_private_companies": "yes",
    "Tar_civil_society": "yes",

    "Actor_role": "4",
    "Discretion": "As a hotel operator Rajesh Aryal stores plastic bottles for Eco-Green "
                 "recycling, separates other plastics for municipal trucks, empties bottles "
                 "before disposal, and participates in hotel-association discussions — but "
                 "has no authority over municipal collection schedules, national policy "
                 "formulation, or thousands of other hotels' waste practices in Chitwan.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_clean_up": "yes",
    "Pol_upcycling": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Environmental Protection Act exists but no specific plastics "
                                 "policy; <40-micron ban weakly enforced; federal government "
                                 "assessing new plastics act but guidelines unclear; Eco-Green "
                                 "bottle collection every 15–20 days but irregular; municipal "
                                 "collection not on time (bins overflow); awareness/cleanup "
                                 "programmes occasional — Environment Day \"only one day\"; "
                                 "PLEASE/Bio-Camp pilots show promise but not yet scaled; "
                                 "effectiveness limited by low enforcement, irregular "
                                 "collection and inconsistent public awareness.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_epr": "yes",
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
    "Sol_clean_up": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Natural plates from banana leaves and pat leaves "
                                          "during feasts; cotton or towel bags for shopping; "
                                          "cloth/fiber bags and clay/cement pots instead of "
                                          "plastic; reuse of plastics before recycling.",
    "Notes": "Master list label: \"14. Hotel Silent Park/ Rajesh Aryal\", affiliation: Hotel "
             "Silent Park. Interviewee: Rajesh Aryal. Location: Chitwan (Sauraha/safari "
             "tourism area). Interview date: 26 June 2025. Interviewers: Ram Devi and Deep. "
             "Coded from the full interview package: structured guideline notes (Themes A–D, "
             "including national policy challenges, PLEASE/Bio-Camp pilots, EPR, informal "
             "sector, Plast Foundation Nepal, Solid Waste Management Association) AND the "
             "verbatim Rajesh Aryal transcript. One field bullet notes greater concern about "
             "aluminum foil than plastic at the hotel level; overall interview concern coded "
             "high given visible/invisible pollution and microplastic impacts in guideline "
             "notes. NPF fields implied, not explicit.",
})

# --- NPL_15: Kathmandu Metropolitan City - Suna Maya Margen -----------------
margen = blank_row()
margen.update({
    "Country": "Nepal",
    "ID": "NPL_15",
    "Actortype": "loc_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "yes",
    "Problem_concerndness": "high",
    "Problem_littering": "yes",
    "Problem_consumption": "yes",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",
    "Problem_import": "yes",

    "Impacts": "visual pollution (described as the biggest problem visually since 1998, "
               "linked to tourists and lifestyle), drainage blockage and riverbank dumping, "
               "rising share of low-value multi-layer plastics (8% rising to 15%), "
               "microplastics impacts on health, environment and wildlife (raised as a public "
               "knowledge gap)",
    NPF_VICTIMS_COL: "rivers and riverbanks (dumping, drainage blockage), drainage/sewer "
                     "systems, the visual environment and tourism image of Kathmandu, "
                     "neighbouring municipalities/communities opposing MRF siting, wildlife "
                     "and the broader environment (microplastics), and valley residents "
                     "affected by inadequate processing capacity despite sufficient waste "
                     "volumes",
    "NPF_villains": "consumers and lifestyle/tourism-driven plastic use (visual pollution); "
                    "multi-layer plastic producers/suppliers (low-value plastics with \"no "
                    "value\" left behind); unregistered/informal private actors (KMC works "
                    "only with registered private sector and explicitly does not work with "
                    "the informal sector); residents who oppose MRF/waste facilities nearby "
                    "(\"people don't want wastes everywhere\"); and uncertain import flows "
                    "(\"import of plastics??\" raised as an open policy issue)",
    "NPF_hero": "Kathmandu Metropolitan City (public education, composting programmes at "
               "household/community/school levels, landfill site management for 23 "
               "municipalities, riverbed debris excavation/transfer); registered private "
               "partners under MoUs (Doko Recyclers collecting Cluster 7); UNDP (MoU for "
               "plastics recovery and plastic study); department stores voluntarily "
               "providing cloth bags instead of polythene; Newari festival traditions using "
               "leaf plates (Lapti); and planned measures (action plan, producer subsidies, "
               "single-use plastics ban, subsidised alternatives)",

    "Relevance_international_pol": "yes",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_nat_government": "yes",
    "Res_prov_government": "yes",
    "Res_loc_government": "yes",
    "Res_students": "yes",
    "Res_private_sector": "yes",
    "Res_households": "yes",
    "Res_edu_institutions": "yes",
    "Res_private_companies": "yes",

    "Cul_nat_government": "yes",
    "Cul_households": "yes",
    "Cul_private_sector": "yes",

    "Tar_private_companies": "yes",
    "Tar_households": "yes",
    "Tar_students": "yes",
    "Tar_edu_institutions": "yes",
    "Tar_loc_government": "yes",
    "Tar_nat_government": "yes",
    "Tar_prov_government": "yes",

    "Actor_role": "3",
    "Discretion": "As an Environment Inspector focused on public education, Ms. Margen "
                 "implements composting and plastics-reduction programmes at household, "
                 "community and school levels and works within KMC's framework of MoUs with "
                 "registered private partners (e.g. Doko Recyclers in Cluster 7) — but has "
                 "limited discretion over land procurement (federal land authority), "
                 "engagement with the informal sector (explicitly excluded), national policy "
                 "revision timelines, or siting MRF facilities opposed by neighbouring "
                 "municipalities.",

    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "yes",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Pol_epr": "no",
    "Pol_import": "no",
    "Pol_awareness": "yes",
    "Pol_education": "yes",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_clean_up": "yes",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_example": "Ministry of Forest and Environment plastics policy \"not "
                                 "implemented and currently under revision\"; KMC relies on "
                                 "the Environment Act 2077 and Natural Resource Management "
                                 "Act 2007 rather than a plastics-specific local ordinance; "
                                 "single-use plastics ban is still at the planning stage; no "
                                 "assessment yet of the 40-micron notification; sufficient "
                                 "waste is collected but KMC cannot procure land for MRF/"
                                 "processing — effectiveness limited by land, technology and "
                                 "enforcement gaps.",

    "Sol_lead_agency": "yes",
    "Sol_responsibilities": "yes",
    "Sol_awareness": "yes",
    "Sol_segregation": "yes",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_ban": "yes",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_clean_up": "yes",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions to build on (free-hand)": "Newari community festival practices that try to "
                                          "minimise plastics by using leaf plates (Lapti) "
                                          "instead of disposable plastic — effective "
                                          "culturally but expensive, with cheaper plastic "
                                          "alternatives still dominating unless subsidised.",
    "Notes": "Master list label: \"15. Kathmandu Metropolitan City/ Suna Maya Margen\", "
             "affiliation: Kathmandu Metropolitan City - Environment Division / Public "
             "Education & Inspection. Interviewee: Ms. Suna Maya Margen (30 years), "
             "Environment Inspector, Public Education. Interview date: 23 June 2025; "
             "internal interview no. 2. Coded from interview guideline/field notes only "
             "(no full verbatim transcript supplied). Data-quality notes: the follow-up on "
             "concern level (Q1) and Q2 (visible effects) were left blank in the source; "
             "ministry-template follow-ups Q4a/Q4b and Q4 effectiveness follow-up were also "
             "blank. Overlap with NPL_3 (same municipality, Doko Recyclers Cluster 7, land/"
             "MRF challenges, Lapti tradition, UNDP MoU) but distinct respondent with "
             "emphasis on public education, composting and inspection rather than SWM office "
             "policy formulation. NPF fields implied, not explicit.",
})

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
    "Actortype": "Rural Self-Reliance Development Centre (RSDC/RSCT) is an NGO/cooperative "
                 "network established in 1991, now piloting urban WASH and plastic-waste "
                 "management; coded as civil_society.",
    "Problem_awareness_pop": "\"People know that plastic affects them but they don't have a "
                             "deep understanding\"; PET bottles labelled single-use are "
                             "reused for water/milk/oil; burning plastic is assumed safe if "
                             "\"a bit far from my house\"; public described as \"ignorant\" "
                             "of waste-management practices.",
    "Problem_awareness_pol": "Ward-level officials are supportive but \"the mayor has not yet "
                            "understood or internalized it\"; local government does not "
                            "implement its own segregated-waste fee-reduction policy despite "
                            "knowing unlicensed collectors operate.",
    "Problem_concerndness": "Plastic described as a \"huge problem\" in growing cities, a "
                            "global concern (Great Pacific Garbage Patch cited), with "
                            "landfills half-filled in three years instead of twenty; RSDC "
                            "has launched its first project fully focused on WASH/plastics.",
    "Problem_littering": "Plastic \"is everywhere around us\" in main cities; rivers in "
                         "Kathmandu and other cities \"have become a dumping site\"; plastic "
                         "clogs sewer systems.",
    "Problem_consumption": "High per-capita plastic production (ADB 2013 data cited); "
                           "plastics \"have become an inseparable part of our lives\"; urban "
                           "areas trend toward single-use disposal.",
    "Problem_recycling": "Cooperative model recovers and sells segregated waste, but "
                         "earnings are minimal (\"a huge amount of plastic becomes only one "
                         "kilogramme\"); municipal incentive policy for segregators is not "
                         "implemented.",
    "Problem_waste_mgmt": "Landfills designed for ~20 years half-filled in ~3 years; "
                          "policies exist but \"we always miss on implementation\"; "
                          "stakeholders not working in a coordinated way.",
    "Problem_production": "\"Lots of plastic being produced per capita in the context of "
                          "Nepal\" (ADB 2013 report referenced).",
    "Problem_alternatives": "R&D needed for alternatives such as plant-based plastic bags "
                            "(only 20–30 NPR/kg price difference cited); cannot aim for "
                            "zero plastic production.",
    "Problem_waste_segregation": "Households mix everything for the morning garbage truck "
                                 "despite segregation points and vendors; no financial "
                                 "incentive because the NPR 300/month fee is unchanged "
                                 "whether households segregate or not.",
    "Impacts": "Indoor air pollution from burning MLP wrappers; blocked sewers; rapidly "
               "filling landfills; river pollution harming aquatic life; microplastics in "
               "Chitwan lake and even salt; persistent environmental accumulation.",
    NPF_VICTIMS_COL: "People suffering indoor air pollution from burning plastic; aquatic "
                     "life and freshwater ecosystems; communities drinking contaminated "
                     "water; public with only superficial awareness of plastic harms.",
    "NPF_villains": "Influential plastic producers/manufacturers; unlicensed private waste "
                    "collectors (Sawman/Neximac) who profit from fees and recyclables while "
                    "doing minimal real conservation work; local governments that fail to "
                    "implement policies, freeze budgets, and tolerate unlicensed operators; "
                    "ignorant/NIMBY households.",
    "NPF_hero": "RSDC and its cooperative segregation/collection model; multi-stakeholder "
               "engagement; exposure visits changing local leaders' mindsets; EPR as the "
               "key proposed policy solution; genuine leaders committed to people.",
    "Coordination_sectoral": "\"All the stakeholders related to this plastic waste management "
                             "are not working properly... not working in a coordinated way\"; "
                             "government, public, private collectors and NGOs not aligned.",
    "Coordination_levels": "Ward officials support the programme but municipal leadership "
                           "does not prioritise it; federal earmarked environmental budgets "
                           "not released or spent by local government.",
    "Unclear_responsibilities": "Local government \"doesn't even know who is collecting waste "
                                "in their jurisdiction\" and has no list of waste-management "
                                "organisations; unlicensed private companies operate despite "
                                "Solid Waste Management Act 2011 registration requirements.",
    "Res_nat_government": "National level sets policy frameworks; draft EPR policy exists; "
                          "Department of Environment occasionally monitors plastic-bag rules.",
    "Res_loc_government": "Budhanilkantha Municipality responsible for waste policy "
                          "implementation, fee rules, licensing collectors, and spending "
                          "environmental budgets.",
    "Res_private_sector": "Private companies (not government) collect household waste for "
                          "~NPR 300/month across the municipality.",
    "Res_civil_society": "RSDC runs the Budhanilkantha pilot, establishing cooperatives and "
                         "waste-segregation points; cooperatives collect segregated waste.",
    "Res_households": "Households generate waste and can segregate/recover materials for "
                      "sale; project works with 100+ households.",
    "Res_private_companies": "Plastic producers and private waste-collection companies are "
                             "named as key actors in the value chain (should be licensed "
                             "under Solid Waste Management Act 2011).",
    "Cul_loc_government": "Does not implement segregated-waste fee-reduction policy; freezes "
                          "earmarked environmental budgets; lacks political will and "
                          "ownership; allows unlicensed collectors; mayor has not "
                          "internalised the programme.",
    "Cul_private_sector": "Unlicensed waste collectors operate with municipal tacit approval; "
                          "Sawman/Neximac groups claim conservation work but \"it's mostly "
                          "just talk\"; companies compete for contracts with political "
                          "favouritism.",
    "Cul_households": "Public \"ignorant\" of proper waste management; NIMBY syndrome "
                      "(want clean surroundings but oppose nearby collection points); mix "
                      "waste despite available segregation options.",
    "Cul_private_companies": "Producers are \"100%\" influential, \"us[ing] money\" to shape "
                             "policy; producing and selling creates the problem without "
                             "end-of-life responsibility.",
    "Tar_nat_government": "Enact and enforce EPR legislation (currently only a draft).",
    "Tar_loc_government": "Implement existing policies (fee reductions for segregated waste), "
                          "release frozen budgets, license and oversee private collectors, "
                          "prioritise plastics within environmental spending.",
    "Tar_private_sector": "Private waste collectors must register, act responsibly, raise "
                          "awareness, and engage communities rather than only collecting "
                          "fees.",
    "Tar_households": "Households should segregate waste, stop burning plastic, and participate "
                      "in cooperative collection/recycling models.",
    "Tar_private_companies": "Plastic producers/importers must be held accountable through "
                             "EPR; waste-collection companies must be licensed and regulated.",
    "Tar_civil_society": "Cooperatives and community groups should be supported as local "
                         "collection/recycling partners (RSDC model).",
    "Actor_role": "RSDC both organises cooperatives, segregation points and municipal "
                 "engagement (managerial implementer — role 2) and works directly with "
                 "100+ households on the ground (street-level implementer — role 3).",
    "Discretion": "RSDC independently designed the Swabalamban cooperative model, chose to "
                  "pilot it for urban plastics/WASH, established local cooperatives and "
                  "segregation infrastructure, and organised exposure visits — all without "
                  "being mandated by government.",
    "Monitoring": "20–40 micron bag rule is \"not regular\" monitoring (only occasional DoE "
                  "checks on Facebook); implementation and policy compliance broadly "
                  "unmonitored.",
    "Financial_resources": "Earmarked environmental budgets exist but are frozen year after "
                           "year because municipalities produce no action plans; \"fund is "
                           "there\" but not released or used.",
    "Research": "\"There are not many research papers\" on plastic health/environment "
               "impacts in Nepal; \"we don't have evidence-based information\" — localized "
               "data needed to convince decision-makers.",
    "Infrastructure": "In the project area, waste-segregation points and vendors buying "
                      "segregated materials already exist — the barrier is lack of "
                      "incentives and political will, not absence of basic facilities "
                      "(though NIMBY still blocks new collection-centre siting).",
    "Capacity": "Local-government environmental staff lack motivation (fixed salaries "
               "regardless of performance); communities and officials \"don't know how "
               "to proceed\" without exposure and evidence; communication and localized "
               "know-how are missing.",
    "Enforcement": "20–40 micron plastic-bag rule not regularly enforced; Solid Waste "
                   "Management Act 2011 licensing requirement for private collectors "
                   "ignored in Budhanilkantha.",
    "Pol_epr": "\"There is a draft policy on this in Nepal, but it hasn't been enacted yet\" "
              "— EPR described as essential but not yet in force.",
    "Pol_awareness": "Project-based NGO awareness efforts; private collector associations "
                     "(Sawman) claim to run awareness/training but actual work is minimal.",
    "Pol_ban": "Regulation requiring plastic bags of at least 20–40 microns exists on paper.",
    "Pol_recycling": "RSDC cooperative collects segregated household waste and sells "
                     "recyclables to vendors.",
    "Pol_waste_collection": "~NPR 300/month household fee to private collectors; municipal "
                            "policy to reduce fee if waste is segregated (not implemented).",
    "Pol_effectiveness": "Policies exist at national and local levels but implementation is "
                         "\"very weak\"; segregated-waste incentives not applied; budgets "
                         "frozen; monitoring absent.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in the Coded_Data "
                                 "cell.",
    "Sol_lead_agency": "Need leaders who \"genuinely care about people\" and multi-stakeholder "
                       "cooperation rather than single-actor solutions.",
    "Sol_responsibilities": "Enact EPR; register/license private collectors; clarify which "
                            "organisations operate in each municipality.",
    "Sol_epr": "\"Extended Producer Responsibility (EPR) is key... EPR is essential\" — top "
              "proposed policy solution.",
    "Sol_awareness": "Communication and exposure visits to show working models; awareness "
                     "campaigns for all stakeholders including government officials.",
    "Sol_segregation": "Household source segregation with financial incentives (reduced "
                       "collection fees) and cooperative collection.",
    "Sol_recycling": "Integrate cooperatives, vendors and recyclers into municipal systems; "
                     "minimize plastic use but manage remaining waste properly.",
    "Sol_education": "Evidence-based, localized education to show households the extent of "
                     "harm and practical alternatives.",
    "Sol_capacity": "Exposure visits and localized evidence to build municipal and community "
                    "capacity; motivate bureaucrats beyond fixed salaries.",
    "Sol_RD": "R&D for plant-based plastic bags and other affordable alternatives (20–30 "
              "NPR/kg price gap cited).",
    "Sol_finance": "Use existing earmarked environmental budgets rather than letting them "
                   "freeze; practical household incentives (reduced waste fees).",
    "Sol_infrastructure": "Waste-to-energy technologies with minimal pollution; "
                           "segregation points and transfer systems (adapted to local "
                           "capacity, not imported expensive foreign models).",
    "Sol_subsitutes": "Promote plant-based and other affordable alternative products through "
                      "R&D and policy support.",
    "Sol_enforcement": "Register and license private waste collectors per Solid Waste "
                       "Management Act 2011; regular monitoring of micron-thickness rules.",
    "Sol_monitoring": "Regular government monitoring of plastic-bag rules and private "
                      "collector compliance, not occasional checks.",
    "Traditions to build on (free-hand)": "RSDC's 1991 Swabalamban cooperative savings model "
                                          "(200+ cooperatives, two-tier structure); rural "
                                          "reuse of plastic bags many times; aspiration to "
                                          "revive bringing one's own cup when visiting others.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - students are not described as currently responsible actors for "
                    "mitigating plastic pollution.",
    "Res_prov_government": "NA - provincial government is not mentioned in this interview.",
    "Res_science": "NA - researchers are discussed only as a future need for localized "
                   "evidence, not as current responsible actors.",
    "Res_edu_institutions": "NA - schools/universities are not described as currently active "
                            "on plastic waste management.",
    "Cul_nat_government": "NA - national government is discussed mainly as policy-setter and "
                          "budget allocator; the primary blame in this interview falls on "
                          "local-government non-implementation and private actors rather "
                          "than national policy design.",
    "Cul_prov_government": "NA - provincial government is not mentioned.",
    "Cul_students": "NA - not mentioned as contributing to the problem.",
    "Cul_civil_society": "NA - NGOs/cooperatives (including RSDC itself) are presented as "
                         "part of the solution, not the problem.",
    "Cul_science": "NA - not blamed for the situation.",
    "Cul_edu_institutions": "NA - not mentioned as culprits.",
    "Tar_prov_government": "NA - provincial government is not mentioned as a target group.",
    "Tar_students": "NA - youth/students are not named as a primary target group in this "
                    "interview.",
    "Tar_science": "NA - while localized research is called for, researchers are not named "
                   "as a primary target group for mitigation measures.",
    "Tar_edu_institutions": "NA - schools/universities are not named as a primary policy "
                            "target.",
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

doco_expl = {
    "Actortype": "Doco Recyclers is a private-sector waste-management and recycling company "
                 "(dry waste and e-waste), operating material recovery facilities and "
                 "collection chains; coded as private_sector.",
    "Problem_awareness_pop": "\"People don't think this serious. Everyone talks about plastic "
                             "pollution. But if I ask what you did yourself to reduce the "
                             "plastic pollution, hardly you can find any answer\"; burning "
                             "waste in government backyards despite knowing it is harmful; "
                             "notes list \"awareness\" and \"habit\" as challenges.",
    "Problem_awareness_pol": "\"The government of Nepal is still uncertain which ministries, "
                            "and which department is going to look after this plastic waste "
                            "management\"; \"Even they [Ministry of Forest and Environment] "
                            "don't know what policy we should write, because MOE Urban "
                            "development is writing one policy. Department of Environment is "
                            "writing another policy\"; no apex body to guide the sector.",
    "Problem_concerndness": "Notes state respondents express \"serious concern about the "
                            "scale and impacts\" and that end-of-life management requires "
                            "\"urgent policy and infrastructural interventions\"; transcript "
                            "describes the problem as structural and pressing throughout.",
    "Problem_littering": "Plastics found \"below my farmlands\", \"beneath the soil\" in "
                         "rivers across Nepal, and in open dumping/landfill sites; notes "
                         "mention aesthetics and unmanaged waste.",
    "Problem_consumption": "\"We are used to live with the plastic\"; \"the problem started "
                           "when we exploit using it excessively and forget the end-of-life "
                           "management\"; policy should focus on reducing single-use plastic "
                           "production and use.",
    "Problem_recycling": "Only \"hardly 20 to 30 tonnes per day\" of ~600 tonnes generated "
                         "are managed; low-value and multi-layer plastics (MLPs) lack "
                         "recycling infrastructure and market; \"cherry picking\" leaves "
                         "low-value fractions behind.",
    "Problem_waste_mgmt": "Central theme: lack of policy incentives, collection value chain, "
                          "infrastructure, investment and coordination; sanitary landfill, "
                          "leachate and land-acquisition problems noted in the guideline.",
    "Problem_production": "Future policy should focus on \"reducing the production and use "
                          "of the single use plastics\"; producers/importers/brand owners "
                          "put plastic into the market without end-of-life responsibility.",
    "Problem_alternatives": "Shift to alternatives \"cannot be overnight\"; until then only "
                            "end-of-life management is feasible; traditional leaf/cotton-bag "
                            "alternatives exist but are not yet scaled as mainstream "
                            "substitutes.",
    "Problem_waste_segregation": "Notes list \"segregation at source\" as a main challenge; "
                                 "proper collection value chain requires source segregation.",
    "Problem_import": "Some plastic fractions are \"passed on to our neighboring country "
                      "through illegal channels\"; federal government should \"regulate what "
                      "kind of plastics are imported\" (Q9 notes).",
    "Impacts": "Soil and water contamination reducing agricultural yield and nutrient/water "
               "retention; open burning releasing carcinogenic dioxin and furan; "
               "microplastics found at Everest base camp (SAFEC study cited); rivers, "
               "farmland, landfills/leachate and drainage blockage.",
    NPF_VICTIMS_COL: "Farmers (reduced crop production as rainwater/nutrients no longer "
                     "absorbed by plastic-contaminated soil), human health (toxic burning "
                     "emissions), rivers and soil ecosystems, and society broadly.",
    "NPF_villains": "Producers/importers/brand owners without EPR; national government "
                    "(uncoordinated ministries, no apex body, weak ban enforcement); "
                    "public/households (inaction, open burning); informal scrap sector "
                    "cherry-picking high-value plastics and abandoning low-value fractions.",
    "NPF_hero": "Doco Recyclers and peer private waste-management firms running MRFs and "
               "collection chains; the proposed national apex coordinating body; and "
               "private-sector partners already working with municipalities (e.g. KMC "
               "inventory survey).",
    "Coordination_sectoral": "Ministry of Forest and Environment, Department of Environment "
                             "and Ministry of Urban Development write separate policies "
                             "without synchronization; \"everyone is working in different "
                             "directions... Government is trying to move in another "
                             "direction. Private is another direction.\"",
    "Coordination_levels": "Fragmented municipal policies versus absent national plastic "
                           "policy; lack of synergy between government, private sector, "
                           "informal collectors and donor agencies.",
    "Unclear_responsibilities": "\"There is no clear synchronization of who is going to "
                                "look after what. And we don't have a body who can guide "
                                "who can drive these all factors regarding to the plastic "
                                "waste\"; solid waste folded into Municipal Solid Waste "
                                "Management Act without plastic-specific clarity.",
    "Res_nat_government": "National ministries (Environment, Urban Development) and "
                          "Department of Environment are named as policy actors, though "
                          "roles are unclear.",
    "Res_loc_government": "Municipalities are \"one of the most important authorities\" with "
                          "constitutional authority to prepare local guidelines; 753 "
                          "municipalities should be accountable implementing partners.",
    "Res_private_sector": "Doco Recyclers and other private waste-management companies handle "
                          "collection, sorting, processing and consulting; formal and "
                          "informal scrap collectors operate in the value chain.",
    "Res_civil_society": "UNDP, GIZ and NGOs (e.g. SAFEC Centre microplastics study) "
                         "partner on projects, training and research.",
    "Res_science": "Universities and international institutions (e.g. Loughborough "
                   "University UK) collaborate on EPR and end-of-life research.",
    "Res_households": "Households are clients in Doco's collection chain and generate the "
                      "waste stream that must be managed.",
    "Res_edu_institutions": "Schools and universities host awareness/green-school camps and "
                            "research partnerships with Doco.",
    "Res_private_companies": "Institutions, embassies and hotels are collection clients; "
                             "cement industry proposed as RDF destination for low-value "
                             "plastic; producers/importers should bear EPR responsibility.",
    "Cul_nat_government": "No incentive policy, no EPR, no apex body, weak enforcement of "
                          "40-micron ban, and ministries writing conflicting policies.",
    "Cul_loc_government": "Municipal plastic bans exist on paper but \"most of these "
                          "policies are not implemented\" and municipalities \"do not monitor "
                          "private waste collectors adequately.\"",
    "Cul_private_sector": "Established informal/formal scrap systems engage in \"cherry "
                          "picking\" of high-value plastics, leaving low-value/multi-layer "
                          "fractions unmanaged.",
    "Cul_households": "Public talks about pollution but takes little personal action; open "
                      "burning of waste including in government backyards.",
    "Cul_private_companies": "Producers, importers and brand owners place plastic on the "
                             "market without end-of-life responsibility (no EPR).",
    "Tar_nat_government": "Federal government should coordinate, implement EPR, support "
                          "research and regulate plastic imports (Q9 notes).",
    "Tar_loc_government": "\"753 total operations [municipalities] that should be making "
                          "accountable\"; local government inspection role noted in Q9.",
    "Tar_students": "\"It starts from young. As young you can catch them\"; waste management "
                    "should be embedded in the education system.",
    "Tar_private_sector": "Private waste-management companies should be incentivised as "
                          "collection/processing actors in the EPR system.",
    "Tar_households": "Public education needed, though reaching all 30 million people "
                      "individually is impractical compared with targeting producers.",
    "Tar_edu_institutions": "School curricula and facility visits (green-school camps) are "
                            "proposed vehicles for awareness and behaviour change.",
    "Tar_private_companies": "\"Producers either are manufacturers or they are importers or "
                             "they are the brand owners. Everyone should be responsible\"; "
                             "\"you can easily touch those thousand producer or 10,000 "
                             "producers.\"",
    "Actor_role": "Doco Recyclers is a managerial/organisational implementer (role 2): it "
                 "operates MRFs, designs collection systems, runs consulting and awareness "
                 "programmes, but does not formulate national policy.",
    "Discretion": "The company chooses its clientele model, sorting grades, research "
                  "partnerships and municipal consulting frameworks independently (e.g. "
                  "UNDP Pokhara training design, KMC inventory survey, EPR policy research "
                  "with Loughborough University).",
    "Monitoring": "40-micron ban exists but plastics below the threshold are found "
                  "\"everywhere\"; municipalities inadequately monitor private collectors; "
                  "no systematic policy monitoring described.",
    "Financial_resources": "Private entities \"lack confidence whether they will survive\" "
                           "and will not invest without external project funding; "
                           "financing needed for decentralised MRFs and collection chains.",
    "Research": "Doco conducts multiple research projects and calls for more research on "
               "burning impacts, microplastics in water and agriculture; notes list "
               "life-cycle assessment and air/water/soil impacts as knowledge gaps.",
    "Infrastructure": "No infrastructure to recycle/upcycle/dispose low-value plastics; "
                        "need decentralised MRFs, recycling plants and RDF destinations.",
    "Capacity": "\"Human Resources development - capacity-building\" listed in solutions; "
               "private-sector capital and technical capacity must be strengthened.",
    "Enforcement": "40-micron ban not enforced on the ground; burning should be penalised "
                   "but people feel \"nobody is going to punish me.\"",
    "Pol_epr": "\"We don't have policy for extended producer responsibility\"; actively "
              "pushing DoE to adopt EPR for plastic packaging and e-waste.",
    "Pol_import": "No effective import-control mechanism; some fractions exit through "
                  "\"illegal channels\" to neighbouring countries.",
    "Pol_awareness": "Doco runs awareness campaigns and green-school camps; notes describe "
                     "project-based NGO/private awareness efforts rather than continuous "
                     "government campaigns.",
    "Pol_education": "School learning programmes and facility visits where students learn "
                     "about waste processing and carbon footprint reduction.",
    "Pol_RD": "Multiple research projects (UNDP skill development, MLP/low-value plastic "
              "projects, EPR framework research with Loughborough University, NREP energy "
              "efficiency study, KMC inventory survey).",
    "Pol_ban": "Ban on single-use plastic bags below 40 microns was drafted/enacted "
              "(respondents recall ~2016/four-to-five years ago); fragmented municipal "
              "single-use plastic/cutlery bans also exist.",
    "Pol_upcycling": "Paper-to-plant pencil upcycling project in Pokhara (UNDP training); "
                     "community upcycling/recycling groups supported.",
    "Pol_recycling": "Doco's core business: sorting plastics into 10-12 grades and supplying "
                     "recyclers; World Bank-funded plastic recovery facility in Salghari.",
    "Pol_waste_collection": "Own collection chain serving Kathmandu Valley institutions, "
                            "households and embassies under a clientele model.",
    "Pol_effectiveness": "Ban enforcement weak; only ~3-5% of daily plastic generation "
                         "managed; policies fragmented, outdated and poorly monitored; draft "
                         "Solid Waste Act criticised as incomplete.",
    "Pol_effectiveness_example": "See Pol_effectiveness - quoted directly in the Coded_Data "
                                 "cell.",
    "Sol_lead_agency": "Repeated call to form an \"apex body\" as the guiding governing "
                       "body for all waste/pollution management.",
    "Sol_responsibilities": "EPR to assign producer/importer/brand-owner responsibility; "
                            "clarify ministry roles; make municipalities accountable.",
    "Sol_epr": "\"The first policy that should be coming is the extended producer "
              "responsibility\"; EPR listed as \"very important\" in notes.",
    "Sol_awareness": "Information campaigns and mandatory education from a young age; notes "
                     "call for continuous awareness beyond project-based efforts.",
    "Sol_segregation": "Segregation at source and proper collection value chain are "
                       "foundational challenges requiring policy support.",
    "Sol_upcycling": "Support community upcycling/recycling groups and low-value-plastic "
                     "product development.",
    "Sol_recycling": "Decentralised material recovery and recycling facilities in "
                     "high-generation clusters; standardise recycled products to build "
                     "markets.",
    "Sol_education": "Embed waste management in school curricula; facility-based learning "
                     "at MRFs.",
    "Sol_capacity": "Human-resource development and capacity-building for waste collectors, "
                    "manufacturers and municipalities.",
    "Sol_RD": "More research on burning impacts, microplastics in rivers/drinking/"
              "irrigation water and plastics in agricultural soil.",
    "Sol_finance": "Monetary incentives for private collectors/processors; investment in "
                   "decentralised collection and MRF infrastructure.",
    "Sol_infrastructure": "Cluster-based MRFs, recycling plants, RDF/cement-industry "
                           "co-processing for non-recyclables, and technology transfer.",
    "Sol_subsitutes": "Gradual shift from single-use plastics to alternatives; revive "
                      "leaf plates and cotton/towel bags.",
    "Sol_enforcement": "Penalise open burning; enforce EPR and collection obligations "
                       "through rules, guidelines and incentives.",
    "Sol_monitoring": "Municipal corporations as implementing/monitoring partners; central "
                      "government should hold municipalities accountable.",
    "Traditions to build on (free-hand)": "Banana leaves instead of plastic packaging and "
                                          "cotton or towel bags instead of plastic carry "
                                          "bags - traditional practices noted in the "
                                          "guideline and discussed in the traditions "
                                          "follow-up.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - students participate in green-school camps but are not described "
                    "as currently responsible actors for mitigating plastic pollution "
                    "(captured under Tar_students for future education measures).",
    "Res_prov_government": "NA - provincial government is not mentioned in this interview.",
    "Cul_prov_government": "NA - provincial government is not mentioned or blamed.",
    "Cul_students": "NA - students are not mentioned as contributing to the problem.",
    "Cul_civil_society": "NA - NGOs and donors are presented as partners, not culprits.",
    "Cul_science": "NA - researchers are not blamed for the situation.",
    "Cul_edu_institutions": "NA - schools/universities are partners in awareness, not "
                             "culprits.",
    "Tar_prov_government": "NA - provincial government is not mentioned as a target group.",
    "Tar_civil_society": "NA - NGOs are partners rather than primary policy targets in this "
                         "interview.",
    "Tar_science": "NA - while more research is called for, researchers/scientific "
                   "institutions are not named as a primary target group for mitigation "
                   "measures.",
}

hotel_expl = {
    "Actortype": "Mr. Swasti Byanju is a hotel/restaurant owner in Dhulikhel; coded as "
                 "private_companies per the master list (hotels/shops/restaurants).",
    "Problem_awareness_pop": "\"Everyone is aware of the plastic, but still uses it\"; "
                             "\"they are very aware. But as I said there's a lack of policy "
                             "and lack of awareness\"; notes state awareness exists but "
                             "\"behavioral change is lacking.\"",
    "Problem_awareness_pol": "\"There is no certain policy to make this happen\"; \"there is "
                            "no end policy\"; notes cite \"lack of policy\" and \"policy gaps\" "
                            "alongside weak implementation.",
    "Problem_concerndness": "\"I'm very concerned about this\" — both personally and for "
                            "the hotel business; notes state \"Very concerned, both as a "
                            "private person and as a business owner.\"",
    "Problem_littering": "Plastic pollution visible \"everywhere on the streets\"; waste "
                         "thrown into \"rivers, ditches, or open areas\"; dumping directly "
                         "into rivers with tractors (Khurkot, Sindhuli road).",
    "Problem_consumption": "Lifestyles changed — \"everything now comes packaged in plastic\"; "
                           "even remote villages (2–3 days' travel) show noodle/biscuit "
                           "wrappers; plastic bags are only a small portion of overall "
                           "plastic use (clothes, electronics, shampoo sachets, etc.).",
    "Problem_recycling": "\"Waste here is not properly minimized, it's almost zero\" "
                         "recycling in Dhulikhel; only reusable items picked at the dumping "
                         "site while the rest is dumped.",
    "Problem_waste_mgmt": "15+ years of mismanaged local dumping site; no proper segregation, "
                          "transfer stations or environmentally safe landfill; hospital/KU "
                          "waste mixed into general dumps; inter-municipal conflict with "
                          "downstream Panchkhal.",
    "Problem_alternatives": "Paper bags, thokodi and cloth bags are discussed as needed "
                            "substitutes but are not widely available — \"when I was a child\" "
                            "newspaper bags were sold in shops \"but nowadays it's totally "
                            "closed.\"",
    "Problem_waste_segregation": "Notes and transcript repeatedly cite lack of household and "
                                 "hotel-level segregation; municipality should run separate "
                                 "collection after source segregation.",
    "Impacts": "Foul smell on the way to the municipality office and toward Panchkhal; "
               "leachate affecting agriculture in wards 2 and 3; undrinkable water; hospital "
               "waste (needles, syringes) in dumps; aesthetic problems (birds, litter); "
               "ocean plastic impacts on fish cited; 100-year plastic degradation mentioned.",
    NPF_VICTIMS_COL: "Downstream Panchkhal (agricultural, water-scarce area), farmers in "
                     "wards 2 and 3, families and the local community, the environment, "
                     "and Dhulikhel Municipality facing major issues from the dumping site.",
    "NPF_villains": "Government at national and municipal levels (no strict/enforced policy, "
                    "budget not allocated, 15-year dumping mismanagement, stalled projects); "
                    "the public (aware but not acting); shops still using plastic; upstream "
                    "dumping without regard for downstream communities.",
    "NPF_hero": "Dankhuta \"Waste into Money\" model; ward 7 shopkeeper litter rule; plastic "
               "waste collectors; Chitwan recycling unit; personal efforts (alternative bags, "
               "household bag reuse); Kathmandu mayor as example of will + budget.",
    "Coordination_sectoral": "Ministry of Forest and Environment / Department of Environment "
                             "formulate policy while municipalities implement — but "
                             "coordination and implementation are weak.",
    "Coordination_levels": "Central budget \"has not been able to\" reach the local level; "
                           "national policy without local implementation fails; problems "
                           "differ at national, municipal and ward levels.",
    "Unclear_responsibilities": "Roles known in principle (mayor/municipality + citizens) "
                                "but \"coordination and implementation are weak\"; "
                                "respondent personally \"don't know\" specific national "
                                "policies.",
    "Res_nat_government": "Ministry of Forest and Environment / Department of Environment "
                          "formulate policy and issued ban notice; national government "
                          "responsible for enforcement and budget allocation.",
    "Res_loc_government": "Municipality (Mirvan post collection), mayor holds budget and "
                          "formulates/manages waste — \"in a major way they are responsible.\"",
    "Res_private_sector": "\"Those men who collect the plastic are also a very important part "
                          "of this thing\"; waste collectors and recyclers must be integrated.",
    "Res_households": "\"We are also responsible for this\"; citizens should use alternative "
                      "bags and collect plastic for reuse/recycling.",
    "Res_private_companies": "Hotels and businesses must manage waste and hand it to "
                             "municipal collection; shopkeepers responsible under ward 7 "
                             "litter rule.",
    "Cul_nat_government": "No strict binding national policy; ministry ban notice not "
                          "enforced (vegetables still sold in plastic); national plans "
                          "\"did not go through\" or are not enforced.",
    "Cul_loc_government": "Budget not allocated for waste management; dumping mismanaged "
                          "for 15 years; projects announced but no follow-up for years; "
                          "recycling \"almost zero.\"",
    "Cul_households": "\"No one likes plastic, but... everyone is using plastic\"; awareness "
                      "without behavioural change.",
    "Cul_private_companies": "Plastic \"comes from every shop\"; shops still supply goods in "
                             "plastic packaging.",
    "Tar_nat_government": "National-level binding policy with budget allocation (at least 5% "
                          "of waste-management budget for plastic); enforcement from "
                          "government level down.",
    "Tar_loc_government": "Mayor must prioritise dumping, allocate 2–10% of municipal "
                          "budget, organise segregation/collection/landfill management.",
    "Tar_households": "Large-scale awareness to \"every person in every house\"; household "
                      "paper-bag distribution in initial transition period.",
    "Tar_private_sector": "Waste collectors/recyclers must be integrated into systematic "
                          "management; Chitwan-style recycling requires municipal "
                          "coordination.",
    "Tar_private_companies": "Ban plastic in every shop; mandate paper/cloth bags; hotels "
                             "should segregate waste before municipal handover.",
    "Actor_role": "As a hotel/restaurant owner the interviewee is primarily a target of "
                 "future plastic/waste policies (shop/hotel bans, segregation requirements, "
                 "awareness campaigns aimed at businesses and households) rather than a "
                 "policy formulator or municipal implementer — coded as role 4.",
    "Discretion": "Limited operational discretion: uses non-polythene bags when shopping, "
                  "consolidates purchases into fewer bags, family collects bags for reuse, "
                  "but at the hotel can only bag all waste together for municipal pickup.",
    "Monitoring": "Campaigns like \"Beat the Plastic\" happen but \"nothing has changed\"; "
                  "good ward-level initiatives lack \"strong monitoring\" and long-term "
                  "results; projects announced with no follow-up for years.",
    "Financial_resources": "\"Currently, it seems the budget itself is not being allocated\"; "
                           "central budget exists but has not reached the local level; need "
                           "separate compulsory budget line (2–10% of municipal budget).",
    "Infrastructure": "No proper segregation areas, separate collection, transfer stations "
                      "or environmentally safe landfill sites despite 40 years of discussion "
                      "(Daika, Japanese team, ADB, World Bank).",
    "Capacity": "Municipality lacks organisation to send segregated waste to Chitwan "
               "recycling unit; manpower and policy transfer needed from well-managed "
               "models.",
    "Enforcement": "Strict policies needed \"in a strict way\"; national ban discussed but "
                   "shops still give vegetables in plastic; open dumping into rivers "
                   "continues.",
    "Pol_epr": "Extended producer responsibility is not mentioned anywhere in this interview.",
    "Pol_awareness": "\"Beat the Plastic\" programme on World Environment Day; clean-up "
                     "campaigns in Dhulikhel; calls for large-scale household awareness.",
    "Pol_education": "Awareness campaigns should target women/housewives who do shopping; "
                     "education on side effects of plastic to every household.",
    "Pol_ban": "Ministry of Forest and Environment / Department of Environment \"issued a "
              "notice to ban the plastic\"; national ban plans discussed but not consistently "
              "implemented.",
    "Pol_clean_up": "Ward 7 rule requiring shopkeepers to collect litter in front of their "
                    "shops; municipality collects next day; clean-up campaigns started in "
                    "Dhulikhel.",
    "Pol_waste_collection": "Municipal collection (Mirvan post); ward 7 dustbin system where "
                            "each shop has a bin and municipality picks up daily.",
    "Pol_effectiveness": "Limited effectiveness — campaigns and events but \"nothing changes "
                         "in practice\"; Dhulikhel cleaner than Kathmandu in some respects "
                         "due to ward 7 rule but recycling near zero and dumping site "
                         "unresolved for 15 years.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in the Coded_Data "
                                 "cell.",
    "Sol_awareness": "\"Campaigns to raise awareness are more important than funding\"; show "
                     "side effects to every household; target women/housewives.",
    "Sol_segregation": "Segregation at source (household and hotel level) before municipal "
                       "collection; designated segregation areas and separate collection.",
    "Sol_recycling": "Replicate Dankhuta and Chitwan models; integrate waste collectors; "
                     "mayor should focus on recycling — \"even just 2% of the annual budget "
                     "could start something meaningful.\"",
    "Sol_education": "Embed waste management in education from young age; company/community "
                     "awareness especially for women.",
    "Sol_finance": "Allocate fixed percentage of municipal budget (2%, 5%, 6% or 10%); "
                   "national level at least 5% of waste-management budget for plastic; "
                   "separate compulsory budget line.",
    "Sol_infrastructure": "Designated segregation areas, transfer stations, environmentally "
                           "safe landfill sites; decentralised facilities adapted to local "
                           "budget and capacity.",
    "Sol_subsitutes": "Ban plastic in shops; mandate paper bags and cloth bags; distribute "
                     "paper bags to every household in initial phase; revive thokodi.",
    "Sol_enforcement": "Strict policies with real enforcement; national-level binding policy "
                       "requiring budget allocation — \"unless the government enforces "
                       "policies at the national level... otherwise its quite difficult.\"",
    "Sol_monitoring": "Policies must have follow-up and monitoring, not just announcements; "
                      "continuous campaigns not one-off events.",
    "Traditions to build on (free-hand)": "Grandfathers made and sold thokodi (newspaper "
                                          "paper bags) in shops; paper bags and locally made "
                                          "cloth/cotton bags as alternatives to polythene.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - KU students participated in Beat the Plastic but students are not "
                    "described as currently responsible actors for mitigation.",
    "Res_prov_government": "NA - provincial government is not mentioned.",
    "Res_civil_society": "NA - NGOs are not described as currently responsible actors "
                         "(SAFEC study cited only as heard-about example).",
    "Res_science": "NA - researchers are not described as currently responsible actors.",
    "Res_edu_institutions": "NA - Kathmandu University mentioned only in campaign "
                            "participation, not as a responsible mitigation actor.",
    "Cul_prov_government": "NA - not mentioned.",
    "Cul_students": "NA - not mentioned as contributing to the problem.",
    "Cul_private_sector": "NA - waste collectors are presented as important partners/heroes, "
                          "not culprits.",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - KU hospital waste in dumps implicates institutions but "
                             "coded under waste-management failure rather than as a "
                             "separate culprit category.",
    "Tar_prov_government": "NA - not mentioned as a target group.",
    "Tar_students": "NA - not named as a primary target group (education is household/general).",
    "Tar_civil_society": "NA - not named as a target group.",
    "Tar_edu_institutions": "NA - schools not named as a primary policy target in this "
                            "interview.",
    "Tar_science": "NA - research needs mentioned but scientists not named as target group.",
}

moud_expl = {
    "Actortype": "Interview conducted at the Ministry of Urban Development (MoUD), "
                 "Government of Nepal — a national government ministry with cabinet-delegated "
                 "authority to formulate solid-waste policy, acts and regulations.",
    "Problem_awareness_pop": "Senior sociologist: public \"don't care about the future public "
                             "health and environmental implications\"; people feel plastic "
                             "bags are convenient; no culture of source segregation; waste "
                             "traditionally thrown from doors/windows.",
    "Problem_awareness_pol": "Local authorities confused about onboarding private sector, "
                            "land acquisition and donor support; elected municipal leaders "
                            "\"don't know the technical know-how about solid waste\" without "
                            "MoUD backing.",
    "Problem_concerndness": "\"Nepal government is very serious about this plastic waste\"; "
                            "plastic is the majority component disturbing society in urban "
                            "areas; Bagmati cleaning campaign running 10 years because "
                            "plastic is the majority problem.",
    "Problem_littering": "Plastic thrown \"everywhere... here and there, indiscriminately\"; "
                         "rivers in Kathmandu and other cities used as dumping sites; waste "
                         "at start/end of new municipalities and on riverbanks.",
    "Problem_consumption": "\"The life of people is closely attached with the life of the "
                           "plastic\"; consumeristic culture and modernization changed waste "
                           "composition from organic to plastic.",
    "Problem_recycling": "Circular economy promoted but requires source segregation; "
                         "recycling clauses in policy but no dedicated plastic policy yet.",
    "Problem_waste_mgmt": "Sanitary landfills operated as dumping sites (Bancharedanda filled "
                          "in 3–4 years vs. 20-year design); inter-municipal waste conflicts.",
    "Problem_production": "Discussion of producers across the \"whole business\" — plastic "
                          "carriers, packaging, bottled water — and weak polluter's-pay "
                          "implementation.",
    "Problem_alternatives": "Separate plastics-only policy planned; R&D for plant-based bags "
                            "(20–30 NPR/kg price gap); historical leaf plates vs. plastic.",
    "Problem_waste_segregation": "\"The main challenge is the separation of the waste... "
                                 "source segregation\"; without it plastic management is very "
                                 "difficult.",
    "Impacts": "Rivers/canals and sewer lines blocked; landfill lifespan collapse; tourism "
               "aesthetics; agricultural productivity loss; public health from black-plastic "
               "food packaging; leachate, stench and flies; social conflict.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "households/public (health risks from meat/blood/liquids in black "
                     "plastic); communities near dumps (\"weak people\", leachate, flies); "
                     "farmers (declining productivity); tourism/urban aesthetics; rivers and "
                     "aquatic life. No FLAG needed — victims are implied through attributed "
                     "harms.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "the general public (convenience, no segregation, indiscriminate "
                    "disposal); producers/manufacturing \"whole business\" (polluter's pay "
                    "weak); municipalities running landfills as dumps and cross-border "
                    "dumping; the people-vs-policymaker \"war\". No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
               "MoUD/central government (new act, EPR, PPP, guidelines); provincial/"
               "central landfill support; private sector under PPP; training centre and "
               "municipal software; circular economy; NGOs/experts. Ministry partly "
               "positions itself as the hero. No FLAG needed.",
    "Relevance_international_pol": "SDG and human-rights frameworks integrated into act "
                                   "formulation; foreign investment and international donor "
                                   "partner coordination; foreign technology for SWM.",
    "Coordination_sectoral": "Need multi-stakeholder engagement — government, private "
                             "collectors, producers, cooperatives, communities; MoUD "
                             "coordinates international donors.",
    "Coordination_levels": "New act defines central, provincial and local roles; provincial/"
                           "central help local government acquire landfill land; "
                           "inter-municipal waste conflicts.",
    "Unclear_responsibilities": "\"War between people and policy makers\" on who is "
                                "responsible; local governments confused on private-sector "
                                "onboarding, land and donors until act is promulgated.",
    "Res_nat_government": "MoUD develops policies, acts, regulations and guidelines; "
                          "coordinates international donor partners.",
    "Res_prov_government": "Provincial government role defined in new act; helps acquire "
                           "central landfill land when municipalities cannot.",
    "Res_loc_government": "Constitutionally responsible under Local Self-Governance Act "
                          "2016 for solid waste within municipal areas.",
    "Res_private_sector": "Private companies collect household waste; PPP model with sample "
                          "municipal agreements; foreign investors welcomed.",
    "Res_science": "University people and researchers consulted in act formulation; "
                   "welcomes expert/NGO ideas.",
    "Res_households": "Households generate waste; source segregation is foundational.",
    "Res_private_companies": "Producers/manufacturing industries across packaging and "
                             "plastic products; commercial activities generate byproducts.",
    "Cul_households": "Public perception prioritises convenience over environmental health; "
                      "no segregation culture; NIMBY while still dumping waste.",
    "Cul_loc_government": "Sanitary landfills operated as dumping sites; confusion and "
                          "lack of political will at municipal leadership level.",
    "Cul_private_companies": "Producers pollute but polluter's-pay \"weak in implementation\"; "
                             "industries influential through money.",
    "Tar_loc_government": "Municipalities must implement act, prepare concepts, manage "
                          "regional sanitary landfills; need orientation and technical "
                          "support from MoUD.",
    "Tar_private_sector": "Private waste collectors must register and operate under PPP "
                          "agreements; foreign/domestic investors targeted.",
    "Tar_households": "Household habits must change (20–25 years cited for habit change); "
                      "source segregation essential.",
    "Tar_private_companies": "Producers/manufacturing industries to be engaged after act "
                             "finalisation for extended producer/polluter responsibility.",
    "Tar_edu_institutions": "University people consulted; training centre to develop "
                            "orientation packages for municipalities.",
    "Actor_role": "MoUD staff speak as national policy formulators (drafting the Solid Waste "
                 "Management Act, regulations, guidelines, PPP templates — role 1). They do "
                 "not describe themselves as street-level implementers.",
    "Discretion": "MoUD chooses act content, PPP sample agreements, guideline scope, "
                  "stakeholder consultation process, and municipal monitoring software "
                  "design — significant policy-formulation discretion.",
    "Monitoring": "No monitoring technology for segregation patterns; no assessment of "
                  "40-micron ban cited; planned municipal software to track segregation, "
                  "transport and disposal.",
    "Financial_resources": "Land acquisition remains difficult; earmarked environmental "
                           "budgets may go unused/frozen at local level; circular economy "
                           "needs financial support.",
    "Research": "Welcomes expert/NGO ideas and international evidence; lacks Nepal-specific "
               "plastic impact studies; separate plastics policy in planning.",
    "Infrastructure": "Landfill land procurement difficult (Land Acquisition Act separate); "
                      "Bancharedanda deterrent effect; transfer stations exist in Kathmandu "
                      "but landfill management failing.",
    "Capacity": "Municipal elected leaders lack technical know-how; MoUD training centre "
               "and orientation packages proposed; 20–25 years for habit change cited.",
    "Enforcement": "Polluter's-pay weak; burning to be strictly prohibited in new act; "
                   "private collectors currently unlicensed in practice.",
    "Pol_epr": "EPR proposed in draft new act but \"hasn't been enacted yet\"; currently no.",
    "Pol_awareness": "Bagmati cleaning campaign (~10 years); planned provincial/district "
                     "consultations and orientations on the new act.",
    "Pol_ban": "Government banned certain 10-micron plastics; new act strictly prohibits "
              "burning any solid waste; heard plan to increase micron threshold to 100.",
    "Pol_capacity": "Department of Urban Development and Building Construction training "
                    "centre to develop orientation packages for municipalities.",
    "Pol_recycling": "Segregation and recycling provisions in solid waste policy; circular "
                     "economy as policy goal.",
    "Pol_waste_collection": "Act covers collection, transportation and disposal; PPP opens "
                            "door for private-sector collection.",
    "Pol_effectiveness": "Pre-2015 act inadequate for federal system; no dedicated plastic "
                         "policy; landfills mismanaged; polluter's-pay weak; ban impact "
                         "unassessed.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "MoUD as central coordinating ministry; council of ministers approval "
                       "path for donor partnerships.",
    "Sol_responsibilities": "Clarify three-tier government roles in new act; EPR for "
                            "producers; register private collectors.",
    "Sol_epr": "EPR \"essential\" and proposed in draft act for those who generate plastic.",
    "Sol_awareness": "Province/district consultations and orientations; not just publishing "
                     "act on website.",
    "Sol_segregation": "Source segregation as primary solution enabling circular economy.",
    "Sol_recycling": "Circular economy and recycling after segregation; separate plastics "
                     "policy later.",
    "Sol_education": "Training/orientation packages via MoUD training centre; gather "
                     "municipal feedback.",
    "Sol_capacity": "Technical backstopping of municipalities; capacity development for "
                    "equity and localization.",
    "Sol_RD": "R&D for plant-based plastic alternatives; foreign technology investment.",
    "Sol_finance": "PPP and foreign investment; provincial/central help for landfill land.",
    "Sol_infrastructure": "Regional sanitary landfill sites covering multiple municipalities; "
                          "waste-to-energy with minimal pollution mentioned.",
    "Sol_enforcement": "Strict ban on burning solid waste with punishment in new act.",
    "Sol_monitoring": "Planned municipal software tracking segregation, transport, disposal "
                      "and landfill status.",
    "Traditions to build on (free-hand)": "Pet-to-khet cycle; leaf plates; organic customary "
                                          "waste practices at low population density; "
                                          "separate household dustbins for decomposable vs. "
                                          "non-decomposable waste.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - students not mentioned as currently responsible actors.",
    "Res_civil_society": "NA - NGOs mentioned as policy contributors, not as current "
                         "responsible mitigation actors (captured under NPF_hero as "
                         "collaborators).",
    "Res_edu_institutions": "NA - universities consulted on act drafting but not described "
                            "as ongoing responsible actors.",
    "Cul_nat_government": "NA - interviewees ARE the national government; past act weaknesses "
                          "discussed but culpability attributed mainly to public, producers "
                          "and local implementation gaps.",
    "Cul_prov_government": "NA - provincial government not blamed; seen as future helper on "
                           "land acquisition.",
    "Cul_students": "NA - not mentioned.",
    "Cul_private_sector": "NA - private waste collectors targeted for regulation, not blamed "
                          "as primary culprits (contrast with producer industry).",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Tar_nat_government": "NA - interviewees are the national government setting policy, not "
                          "a target group of their own measures.",
    "Tar_prov_government": "NA - provincial government is a partner role, not a primary "
                           "target group.",
    "Tar_students": "NA - not named as a target group.",
    "Tar_civil_society": "NA - NGOs welcomed as experts but not named as primary policy "
                         "targets.",
    "Tar_science": "NA - researchers consulted but not named as primary target group.",
}

ward2_expl = {
    "Actortype": "Interview conducted at the Dhulikhel Municipality ward office with the "
                 "Ward Head and ward/municipal staff — coded as loc_government (ward-level "
                 "local government).",
    "Problem_awareness_pop": "\"People don't understand, they just collect and store waste, "
                             "unaware of the consequences\"; \"People treat plastic casually\"; "
                             "\"The biggest challenge is education. Until the community itself "
                             "becomes aware, no matter how many programs we run, it won't work.\"",
    "Problem_awareness_pol": "\"Policies exist only on paper; they are not implemented\"; "
                            "ward lacks authority and resources to act on national/municipal "
                            "rules; \"some matters fall under higher authorities, we haven't "
                            "been able to act effectively.\"",
    "Problem_concerndness": "\"It's a major problem\"; \"Plastic is creating severe problems "
                            "here\"; \"Plastic has become a serious issue regarding waste\"; "
                            "\"Plastic has severely affected farmland fertility.\"",
    "Problem_littering": "Open syringes found along the Panchkhal stream; waste dumped at "
                         "one site; dumping-site runoff affecting lower settlements.",
    "Problem_consumption": "\"Everything now comes packaged in plastic\"; \"Plastic is "
                           "convenient and cheap for just 1 rupee\"; even if households avoid "
                           "bringing plastic home, food products always arrive in plastic "
                           "packaging.",
    "Problem_recycling": "Former ETPC NGO collected household waste but could not sustain "
                         "operations; no systematic recycling at ward level.",
    "Problem_waste_mgmt": "\"The implementation aspect is weak\"; waste collected twice a "
                          "week and dumped at one site; ward tried hospital collaboration but "
                          "\"couldn't succeed\"; no ward-level systematic framework.",
    "Problem_production": "\"There's no monitoring of what kind of plastic is being produced "
                          "or used\"; factories resisted the micron ban.",
    "Problem_alternatives": "\"We don't have affordable alternatives\"; without alternatives "
                            "\"it's difficult to stop its use.\"",
    "Problem_waste_segregation": "\"No, everything collected in the vehicles goes together\"; "
                                 "\"Exactly\" — no source segregation; \"Some segregation "
                                 "happens later, but not properly.\"",
    "Impacts": "Plastic in Ward No. 2 farmlands and rivers; canal blockages; farmland "
               "fertility loss; health concerns (plastic bottles; syringes); dumping-site "
               "runoff to lower settlements.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "farmers and residents of Wards 2 and 3 (dumping site in Ward 8); lower "
                     "settlements affected by runoff; households/public facing health risks; "
                     "rivers and agricultural land in Ward No. 2. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "unaware/casual public; factories resisting the ban; profit-prioritising "
                    "hotels/restaurants/shops; municipality with weak implementation; syringe "
                    "scavengers. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: ward "
               "office (coordination, awareness, Tol Sudhar Samiti); hospital/nursing-college "
               "cloth-bag campaign; ETPC; Tribhuvan University; researchers. Ward partly "
               "positions itself and partners as problem-solvers despite limited capacity. "
               "No FLAG needed.",
    "Coordination_sectoral": "Ward coordinated with nursing college and Dhulikhel Hospital "
                             "(cloth bags), ETPC (waste collection), Tol Sudhar Samiti "
                             "(grassroots), Tribhuvan University (market development), and "
                             "welcomes researcher collaboration.",
    "Coordination_levels": "Waste management contracted at municipality level; ward reports "
                           "problems to municipality but has no separate budget or authority; "
                           "\"difficult to control from the ward level alone.\"",
    "Unclear_responsibilities": "Ward says plastic management is within its responsibility "
                                "but matters fall under higher authorities; residents pay "
                                "municipal cleaning fees but \"we don't know exactly how funds "
                                "are distributed among the 12 wards.\"",
    "Res_nat_government": "National/municipal rules on micron-thickness bans cited.",
    "Res_loc_government": "Municipality contracts waste management to a company, collects "
                          "revenue and manages disposal; ward reports problems upward.",
    "Res_private_sector": "Waste-management company contracted by municipality; private "
                          "companies proposed to collect household waste for a fee.",
    "Res_civil_society": "ETPC formerly collected waste; Tol Sudhar Samiti operates under "
                         "ward guidance.",
    "Res_science": "Researchers (Kathmandu University team) and Tribhuvan University "
                   "collaboration welcomed.",
    "Res_households": "\"Plastic originates from households\"; households should separate "
                      "waste at home.",
    "Res_edu_institutions": "Nursing college partnered on cloth-bag distribution; "
                            "Tribhuvan University coordinating on market-area development.",
    "Res_private_companies": "Hotels, shops and restaurants targeted in anti-plastic "
                             "campaigns; businesses generate plastic waste.",
    "Cul_nat_government": "\"The government tried to make a rule but couldn't enforce it\"; "
                          "policies exist on paper only.",
    "Cul_loc_government": "Municipality collects annual cleaning taxes but \"there's no "
                          "effective execution\"; unclear fund allocation to wards; collection "
                          "vehicles mix all waste.",
    "Cul_households": "Public lacks awareness; people store waste without understanding "
                      "consequences; historically throw waste expecting municipality to clean.",
    "Cul_private_companies": "Restaurants continued serving food in plastic containers; "
                             "businesses \"prioritized quick profits\"; campaigns could not "
                             "continue.",
    "Tar_loc_government": "Deep argues ward should implement rules like road building; "
                          "municipality must improve execution.",
    "Tar_private_sector": "Private companies could collect segregated waste from households "
                          "for a fee.",
    "Tar_households": "Awareness and household source segregation emphasised as essential.",
    "Tar_private_companies": "Hotels, shops and restaurants were direct targets of the "
                             "cloth-bag and anti-plastic-container campaigns.",
    "Tar_edu_institutions": "Nursing college involved in distribution campaign; TU "
                            "collaboration planned.",
    "Tar_civil_society": "Tol Sudhar Samiti and NGOs suggested for pilot projects and "
                         "awareness.",
    "Tar_science": "Researchers invited to support and continue ward efforts after the study.",
    "Tar_nat_government": "Strict policy enforcement of micron ban cited as one of three "
                          "essentials alongside awareness and resources.",
    "Actor_role": "Ward Head and staff describe grassroots coordination, Tol Sudhar Samiti "
                 "guidance, awareness rallies and reporting to the municipality — typical "
                 "street-level/local implementer role (3), not national policy formulation.",
    "Discretion": "Ward can run awareness with limited funds and coordinate pilots through "
                  "Tol Sudhar Samiti, but cannot control municipal trucks, segregation, "
                  "dumping-site location (Ward 8) or waste-contractor operations.",
    "Monitoring": "\"There's no monitoring of what kind of plastic is being produced or "
                  "used\"; ban not enforced.",
    "Financial_resources": "Ward receives no separate waste-management budget; lacks "
                           "resources and manpower despite residents paying municipal "
                           "cleaning fees.",
    "Research": "Ward Head requests researcher suggestions to improve waste management; TU "
               "collaboration discussed; researchers welcomed to continue efforts.",
    "Infrastructure": "Ward has no own dumping site (site in Ward No. 8); topography causes "
                      "runoff affecting Wards 2 and 3; hospital lacks own dumping site.",
    "Capacity": "\"The ward lacks both resources and manpower\"; no right, manpower or "
               "materials to manage waste separately.",
    "Enforcement": "Micron ban and municipal rules not enforced; factories resisted; ward "
                   "lacks authority to enforce.",
    "Pol_epr": "Extended producer responsibility not mentioned.",
    "Pol_awareness": "Environment Day rallies; cloth-bag campaign with nursing college; "
                     "syringe-collection awareness ~8–10 years ago.",
    "Pol_education": "\"The biggest challenge is education\"; awareness essential for long-"
                     "term control.",
    "Pol_tax": "Annual cleaning taxes/fees collected by the municipality from residents.",
    "Pol_ban": "Plastics below a certain micron thickness banned (~3–4 years ago; 40 microns "
              "referenced in discussion).",
    "Pol_subsitutes": "Cloth-bag distribution campaign (~1,000 bags to hotels and shops).",
    "Pol_clean_up": "Environment Day observed with rallies.",
    "Pol_waste_collection": "Household waste collected twice a week by municipal contractor "
                            "and dumped at one site.",
    "Pol_effectiveness": "Implementation weak; policies on paper; ban not enforced; no "
                         "segregation.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Ward as first contact point for collaborations; Tol Sudhar Samiti "
                       "for neighbourhood pilots.",
    "Sol_responsibilities": "Three essentials: public awareness, proper resources, strict "
                            "policy enforcement.",
    "Sol_awareness": "Awareness campaigns, Environment Day, cloth-bag distribution, "
                     "researcher-supported awareness.",
    "Sol_segregation": "Household source segregation and separate bins/cloth bags proposed.",
    "Sol_education": "Community education emphasised as prerequisite for any programme.",
    "Sol_capacity": "Technical and resource support needed for ward and municipal "
                    "implementers.",
    "Sol_finance": "Ward funds used for awareness; calls for clearer municipal budget "
                   "allocation; private collection for a fee proposed.",
    "Sol_infrastructure": "Separate bins at household level; ward lacks own disposal site.",
    "Sol_subsitutes": "Cloth bags, bamboo/paper baskets, straw bundles for meat.",
    "Sol_clean_up": "Environment Day rallies; small-scale ward activities.",
    "Sol_enforcement": "Strict policy enforcement named as essential alongside awareness "
                       "and resources.",
    "Sol_monitoring": "Implied need to monitor plastic production and usage (currently "
                      "absent).",
    "Traditions to build on (free-hand)": "Bamboo/paper baskets, straw-wrapped meat, large "
                                          "reusable cloth bags — pre-plastic carrying "
                                          "practices.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_prov_government": "NA - provincial government not mentioned.",
    "Res_students": "NA - students not named as responsible actors (nursing college is an "
                    "institution, coded under Res_edu_institutions).",
    "Cul_prov_government": "NA - not blamed.",
    "Cul_students": "NA - not blamed.",
    "Cul_private_sector": "NA - private waste contractor not blamed; private collection "
                          "proposed as solution.",
    "Cul_civil_society": "NA - NGOs not blamed (ETPC lacked resources).",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Tar_prov_government": "NA - not named as target group.",
    "Tar_students": "NA - not named as target group.",
}

ward7_expl = {
    "Actortype": "Interview with the Dhulikhel Municipality ward chairman/ward head — "
                 "coded as loc_government (ward-level elected local government).",
    "Problem_awareness_pop": "Households left plastic outside despite distributed containers "
                             "(\"they made bags and throw\"); waste thrown in forests; "
                             "awareness improving but cooperation insufficient; Deep: "
                             "\"people won't change as long as they aren't penalized.\"",
    "Problem_awareness_pol": "\"Big gap between the local and federal governments\"; draft "
                            "Waste Management Act on webpage but ward \"hasn't reached that "
                            "stage yet\"; \"there are no master plans made\"; central "
                            "government slow while municipality must act.",
    "Problem_concerndness": "\"Plastic is a very harmful material\"; pollution \"spread "
                            "everywhere\"; \"the ecosystem is already destroyed\"; rivers "
                            "embarrass Nepal internationally.",
    "Problem_littering": "Plastic spread everywhere; riverbanks covered ~20–30 years of "
                         "dumping; sweepers collect and dump at open sites; waste thrown in "
                         "forests despite highway cleaning.",
    "Problem_consumption": "\"Plastic is used everywhere, in food packaging, drinking "
                           "glasses, etc.\"; fast-food packaging plastics need alternatives.",
    "Problem_recycling": "Some refining at Rs 10–15/kg and on-site separation by 2–4 "
                         "workers, but processing-company agreements terminated; collection "
                         "still ends in rivers.",
    "Problem_waste_mgmt": "20–30 years of dumping; daily smell near residences; waste from "
                          "Banepa and Okharpauwa; contractor system but gaps remain.",
    "Problem_alternatives": "\"We also need alternatives to plastic, like the plastic used in "
                            "fast food packaging\"; ward encourages cloth bags and natural "
                            "materials but systemic alternatives lacking.",
    "Problem_waste_segregation": "Sunday plastic collection exists but household "
                                 "implementation weak; door-to-door container programme "
                                 "insufficient; Deep proposes separate collection days per "
                                 "waste type.",
    "Impacts": "Panchkhal river impacts, daily smell, destroyed ecosystem, dusty landscape, "
               "national-image harm, Bagmati pollution despite donor projects.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "local residents (smell); rivers/Panchkhal and broader environment; "
                     "Nepal's international reputation. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "uncooperative households; slow/weak central government planning; "
                    "pervasive plastic consumption. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: ward "
               "chairman as personal model; Dhulikhel Municipality (segregation, transfer "
               "station plan); Bhutan awareness example; researchers; site workers separating "
               "recyclables; World Bank/ADB on Bagmati. No FLAG needed.",
    "Relevance_international_pol": "Comparisons with Bhutan, Germany, Switzerland and the US; "
                                   "World Bank and Asian Development Bank funding on Bagmati.",
    "Coordination_sectoral": "Municipality contracts private collectors; ward coordinates "
                             "with municipality; discussed refining with a company (agreements "
                             "later terminated); 2–4 workers separate recyclables on site.",
    "Coordination_levels": "\"Big gap between the local and federal governments\"; debate over "
                           "whether central government or municipality should lead; ward "
                           "reports to municipality as main authority.",
    "Unclear_responsibilities": "RAM DEVI: central doesn't act, falls on municipality; ward "
                                "head: state must be responsible OR municipality given full "
                                "power — responsibility contested across levels.",
    "Res_nat_government": "Restricts plastic beyond certain size; draft Waste Management Act; "
                          "Bagmati projects with international donors.",
    "Res_loc_government": "Municipality provides funds through contracts; Dhulikhel segregating "
                          "plastic and composting; transfer-station plan passed.",
    "Res_private_sector": "Contractors collect/refine/manage waste; private companies visit "
                          "monthly; processing companies (agreements terminated).",
    "Res_households": "Household containers and source separation expected; ward chairman "
                      "models zero household waste.",
    "Res_science": "\"Researchers will know how to solve the problem by finding "
                   "alternatives.\"",
    "Cul_nat_government": "\"They keep making plans and talking, but progress is very slow\"; "
                          "no master plans; gap with local government.",
    "Cul_households": "People left plastic outside despite containers; throw waste in forests.",
    "Tar_nat_government": "\"The state has to be responsible — the central government must "
                          "take charge\"; also calls for municipality to be given full power.",
    "Tar_loc_government": "Municipality should make laws, set penalties and enforce them; "
                          "Dhulikhel model like Dharan.",
    "Tar_households": "Door-to-door awareness; household separation of biodegradable and "
                      "non-biodegradable waste; Germany-style differentiated fees discussed.",
    "Tar_private_sector": "Private companies should collect separately by waste type; "
                          "contractor improvements suggested.",
    "Tar_science": "Researchers to develop plastic alternatives.",
    "Actor_role": "Long-serving elected ward chairman describing grassroots supervision, "
                 "door-to-door programmes, dumping-site management and municipal coordination "
                 "— street-level/local implementer (3).",
    "Discretion": "Supervises dumping areas, distributed household containers, coordinates "
                  "with municipality on contractors, requested Nepal-government land ownership "
                  "in Ward No. 3 — but no formal ward waste staff or independent budget.",
    "Monitoring": "\"Cannot monitor 24/7\"; people throw waste in forests when unobserved.",
    "Financial_resources": "Land purchase for transfer station difficult (\"local people didn't "
                           "give it freely\"); processing agreements terminated; municipality "
                           "funds via contracts only.",
    "Research": "Researchers expected to find alternatives to fast-food packaging plastics.",
    "Infrastructure": "Transfer station and landfill planned (1–2 years); land in Thakuri "
                      "village; access roads prepared in Ward No. 3; open dumping ~20–30 years.",
    "Capacity": "No designated ward-level waste-management personnel; relies on municipal "
               "contracts and 2–4 site workers.",
    "Enforcement": "\"Enforcement is key. Lawmakers must make strict laws\"; penalties and "
                   "differentiated fees (Germany Rs 200 vs 500 example) needed — "
                   "\"just teaching won't be enough.\"",
    "Pol_epr": "Extended producer responsibility not mentioned in transcript.",
    "Pol_awareness": "Door-to-door programmes; Environment Day source collection and swap "
                     "system; Bhutan-style public announcements cited.",
    "Pol_education": "Awareness campaigns; municipal meetings on waste frequently.",
    "Pol_ban": "\"The Nepal government restricts plastic use beyond a certain size.\"",
    "Pol_subsitutes": "Cloth bags, natural-material baskets, paper bags promoted.",
    "Pol_clean_up": "Sweepers collect plastic; highway cleaning; Environment Day activities.",
    "Pol_recycling": "Plastic/paper/glass separated at site; refining at Rs 10–15/kg.",
    "Pol_waste_collection": "Sunday plastic collection; monthly private-company visits; "
                            "trucks whistle to signal collection.",
    "Pol_effectiveness": "Container programme failed; agreements terminated; waste still "
                         "reaches rivers; federal-local gap slows progress.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Central government must take charge OR municipality given full "
                       "enforcement power.",
    "Sol_responsibilities": "Clarify state vs municipal responsibility; elected ward "
                            "representative accountable to voters.",
    "Sol_awareness": "Door-to-door, Environment Day, Bhutan/Germany examples, public "
                     "announcements on separation.",
    "Sol_segregation": "Sunday plastic collection; biodegradable/non-biodegradable household "
                       "bins; separate collection days proposed.",
    "Sol_recycling": "On-site separation and refining; expand like Dharan/Dhulikhel "
                     "municipal model.",
    "Sol_education": "Awareness that separation reduces fees (Germany example).",
    "Sol_capacity": "Designated personnel and formal ward-level contacts needed (currently "
                    "none).",
    "Sol_finance": "Differentiated fees for separated vs unseparated waste; municipality "
                   "contract funding.",
    "Sol_infrastructure": "Transfer station, landfill, access roads in Ward No. 3.",
    "Sol_subsitutes": "Natural baskets, paper bags, cloth bags.",
    "Sol_clean_up": "Source collection, highway and neighbourhood cleaning.",
    "Sol_enforcement": "Strict laws, penalties, differentiated waste fees.",
    "Sol_monitoring": "24/7 monitoring impossible — need systemic enforcement not just "
                      "awareness.",
    "Traditions to build on (free-hand)": "Natural-material baskets, paper bags for "
                                          "vegetables, cloth bags, pre-plastic carrying "
                                          "practices.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_prov_government": "NA - provincial government not mentioned.",
    "Res_students": "NA - not mentioned.",
    "Res_civil_society": "NA - not mentioned as responsible actors.",
    "Res_edu_institutions": "NA - not mentioned.",
    "Res_private_companies": "NA - private collectors coded under Res_private_sector.",
    "Cul_prov_government": "NA - not blamed.",
    "Cul_loc_government": "NA - municipality portrayed partly as problem-solver (Dharan/Dhulikhel "
                           "model), not primary culprit.",
    "Cul_students": "NA - not blamed.",
    "Cul_private_sector": "NA - contractors not blamed; processing agreements terminated but "
                          "not framed as villain.",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Cul_private_companies": "NA - not blamed.",
    "Tar_prov_government": "NA - not named.",
    "Tar_students": "NA - not named.",
    "Tar_civil_society": "NA - not named.",
    "Tar_edu_institutions": "NA - not named.",
    "Tar_private_companies": "NA - not named separately from private sector.",
}

safaurja_expl = {
    "Actortype": "Safa Urja Utpadan is a private-sector waste-management contractor "
                 "(collection, segregation, processing) operating under municipal contracts "
                 "in Chitwan district.",
    "Problem_awareness_pop": "\"Awareness in the environment is lacking\"; people discard "
                             "plastics carelessly, burn at home in winter, dump at night in "
                             "bazaars, and assume waste management is the company's "
                             "responsibility.",
    "Problem_awareness_pol": "\"We have not seen improvements from the state\"; government "
                            "has not formalized the system; federal ministry policies unknown "
                            "to interviewee.",
    "Problem_concerndness": "\"Since we work in this field, we are very concerned\"; "
                            "operating in Nepal carries financial, social, ethical and "
                            "political risks.",
    "Problem_littering": "People throw plastic everywhere; bazaar waste dumped at night; "
                         "scavengers, jackals and dogs scatter waste.",
    "Problem_consumption": "Promotes minimising usage, reusable plastics, purchasing "
                           "wisely, and using less plastic in households and businesses.",
    "Problem_recycling": "Company segregates and processes at facility; only unusable "
                         "fraction to landfill; condenses 1–2 quintals to 8–10 kg.",
    "Problem_waste_mgmt": "Monsoon (Jestha–Shravan–Bhadra, ~4 months) makes collection and "
                          "landfill operations very difficult; terrain delays vehicles; "
                          "Ratnanagar alone ~21 tractor loads/day.",
    "Problem_production": "Discussion of PP vs LP plastic types and micron-thickness rules "
                          "(40-micron standard; thicker = higher cost = less usage).",
    "Problem_alternatives": "Paper plates/cups planned for feasts; jute sacks/bags for "
                            "purchased-plastic disposal; no alternative technologies cited.",
    "Problem_waste_segregation": "Household segregation tried but compliance inconsistent; "
                                 "company sorts mixed waste at facility (glass, food, "
                                 "plastic, clothes, hazardous colour-coding).",
    "Impacts": "Soil pH/fertility, crop contamination in monsoon, burning health effects, "
               "environmental pollution from improper disposal.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "farmers/agricultural land; environment; public health from burning; "
                     "future generations. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "unaware/careless public; state inaction; businesses dumping at night. "
                    "No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: Safa "
               "Urja company, school students as change agents, planned PPP model. Company "
               "positions itself as primary problem-solver. No FLAG needed.",
    "Relevance_international_pol": "German operations manager; UK/US stakeholders/owners; "
                                   "comparison with DOKO Recyclers Kathmandu.",
    "Coordination_sectoral": "Company–municipality–community communication essential; "
                             "temples, schools, cooperatives (~150+ public spots) need "
                             "coordinated management.",
    "Coordination_levels": "Calls for coordinated regulation across municipality, "
                           "provincial and central government; PPP across three "
                           "municipalities.",
    "Unclear_responsibilities": "Public assumes free collection at temples/schools; "
                                "leadership vs household responsibility debated; company "
                                "pays municipality but community resists fees.",
    "Res_nat_government": "Long-term government management needed; coordinated top-down "
                          "regulation sought.",
    "Res_loc_government": "\"Responsibility ideally lies with the leadership\"; "
                          "municipality receives NPR 6.6M+ annual payment from company.",
    "Res_private_sector": "Safa Urja collects, segregates, processes and manages landfill "
                          "operations under contract.",
    "Res_households": "Every household must understand the system for management to work.",
    "Res_edu_institutions": "Schools receive awareness programmes every 3 months.",
    "Res_private_companies": "Hotels, hardware stores, hostels, businesses pay tiered "
                             "monthly fees.",
    "Cul_nat_government": "\"We have not seen improvements from the state\"; system not "
                          "formalized.",
    "Cul_households": "Careless disposal, burning plastics, night dumping, fee resistance.",
    "Cul_private_companies": "Bazaar businesses dump at night; some assume free collection.",
    "Tar_nat_government": "Coordinated regulation from central level sought.",
    "Tar_loc_government": "Municipality must take responsibility across all areas and "
                          "submit support proposals.",
    "Tar_students": "School awareness every 3 months — children influence parents.",
    "Tar_private_sector": "PPP model for three municipalities; private sector engagement in "
                          "new waste-management system.",
    "Tar_households": "Household segregation, jute-sack disposal, fee compliance.",
    "Tar_edu_institutions": "School programmes every 3 months.",
    "Tar_private_companies": "Sauraha tourism-zone hotels and businesses targeted every "
                             "5–6 months; tiered business fees.",
    "Actor_role": "Company representative describes contract operations, fee structures, "
                 "technology deployment, PPP planning and municipal negotiations — "
                 "managerial/organisational implementer (2), not a policy formulator.",
    "Discretion": "See Coded_Data Discretion cell — GPS, barcodes, segregation streams, "
                 "awareness scheduling, PPP/TOR development.",
    "Monitoring": "GPS-enabled vehicles; barcode payment-tracking system planned within "
                  "3 months; staff attendance recorded from home.",
    "Financial_resources": "Operating at a loss every year; rates unchanged 7 years; fee "
                           "increases face uproar; NPR 300,000/month needed for public-spot "
                           "collection alone.",
    "Research": "NA - no research programmes mentioned (Deep references CREASION pellet "
                "factory visit but not attributed to Safa Urja).",
    "Infrastructure": "Need proper landfill sites in good locations; existing sites in "
                      "community forests and riverside land; terrain limits monsoon access.",
    "Capacity": "70–80 workers; monsoon terrain difficulties; 21 tractor loads/day in "
                "Ratnanagar alone.",
    "Enforcement": "Fines partial; cannot enforce everywhere; late payment up to 7 months; "
                   "no late-payment penalty currently.",
    "Pol_epr": "Extended producer responsibility not mentioned.",
    "Pol_awareness": "School programmes every 3 months; Sauraha tourism programmes every "
                     "5–6 months.",
    "Pol_education": "Children educated to influence families at home.",
    "Pol_tax": "Tiered household/business monthly fees; company pays NPR 6.6M+ annually "
               "to municipality as tax.",
    "Pol_ban": "40-micron standard discussed; above 40 microns not allowed (thicker = "
              "higher cost).",
    "Pol_subsitutes": "Jute sacks/bags; paper plates/cups for feasts planned.",
    "Pol_recycling": "Facility-level segregation into compostable/recyclable streams.",
    "Pol_waste_collection": "Contracted municipal collection across 16 wards.",
    "Pol_effectiveness": "Fines and fees weakly enforced; state improvements not seen.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "PPP model across three municipalities with consultant/TOR study.",
    "Sol_responsibilities": "Clear communication between company, municipality and "
                            "community.",
    "Sol_awareness": "Regular school and tourism-zone programmes.",
    "Sol_segregation": "Household and facility-level colour-coded segregation.",
    "Sol_recycling": "Process segregated streams; minimise landfill volume.",
    "Sol_education": "Students as family influencers.",
    "Sol_capacity": "Worker responsibility training; 70–80 staff operations.",
    "Sol_finance": "Fee increases; late-payment penalties from municipality; PPP to "
                   "address loss-making contract.",
    "Sol_infrastructure": "Proper landfill sites; GPS/barcode technology.",
    "Sol_subsitutes": "Paper feast ware; jute disposal bags; LP instead of PP plastic.",
    "Sol_enforcement": "Progressive fines (NPR 5,000–10,000); differentiated fees for "
                       "segregated vs unsegregated waste (discussed).",
    "Sol_monitoring": "GPS vehicles, barcode payment tracking, staff attendance system.",
    "Traditions to build on (free-hand)": "Some traditional knowledge acknowledged but no "
                                          "supporting policy; paper feast alternatives and "
                                          "jute bags proposed.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_prov_government": "NA - provincial government mentioned only as part of desired "
                           "coordinated regulation, not as current responsible actor.",
    "Res_students": "NA - students targeted for awareness but not described as responsible "
                    "mitigation actors (coded under Tar_students).",
    "Res_civil_society": "NA - temples/schools/cooperatives discussed but not as formal "
                         "responsible actors.",
    "Res_science": "NA - not mentioned.",
    "Cul_prov_government": "NA - not blamed.",
    "Cul_loc_government": "NA - municipalities receive payments and are partners in PPP, "
                           "not primarily blamed.",
    "Cul_students": "NA - not blamed.",
    "Cul_private_sector": "NA - Safa Urja is the interviewee; other private actors not "
                          "blamed.",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Tar_prov_government": "NA - provincial level only in desired regulation, not as "
                           "primary target group.",
    "Tar_civil_society": "NA - not named as primary target group.",
    "Tar_science": "NA - not named.",
}

moud12_expl = {
    "Actortype": "Interview at the Ministry of Urban Development (MoUD) / Ministry for Urban "
                 "Planning — national government ministry with authority to formulate solid-waste "
                 "policy, acts and regulations.",
    "Problem_awareness_pop": "People lack segregation habits; lives \"closely attached to "
                             "plastic\" for convenience; don't consider future public-health "
                             "implications; heterogeneous waste thrown together on streets.",
    "Problem_awareness_pol": "Municipalities confused about onboarding private sector and "
                            "donor partners; no specific plastic policy yet; only general "
                            "solid waste framework.",
    "Problem_concerndness": "\"Plastic is the component that disturbs society the most\"; "
                            "\"Nepal government is very serious about this plastic waste\"; "
                            "remarkable volume in waste composition.",
    "Problem_littering": "Plastic visible everywhere in urban areas; waste thrown haphazardly "
                         "without control.",
    "Problem_consumption": "Public convenience attachment to plastic for carrying goods.",
    "Problem_recycling": "Circular economy promoted but requires source segregation first.",
    "Problem_waste_mgmt": "Banchare Danda operated as dumping site (20-year design filled in "
                          "3–4 years); waste picked up and dumped without processing.",
    "Problem_production": "Producers/industries discussed under proposed EPR and weak "
                          "polluter-pays implementation.",
    "Problem_alternatives": "Separate plastics policy planned; foreign technology and "
                            "investment welcomed.",
    "Problem_waste_segregation": "\"The main challenge is the separation of waste, specifically "
                                 "source segregation\"; most Kathmandu waste not in desired form.",
    "Impacts": "Rivers/canals, sewers, tourism, health (black plastic), agriculture, landfills, "
               "livestock, household microeconomics.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "public/households; communities near dumps; farmers; tourism/urban "
                     "environments; rivers and livestock. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "convenience-driven public; weakly regulated producers; municipalities "
                    "running dumps; inter-municipal conflicts. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
               "MoUD/central government (new act, EPR, PPP); provincial/central landfill "
               "support; private sector; training centre/software; circular economy. Ministry "
               "positions itself as primary problem-solver. No FLAG needed.",
    "Relevance_international_pol": "Foreign direct investment and international donor partner "
                                   "coordination; SDG and human-rights frameworks integrated.",
    "Coordination_sectoral": "Multi-stakeholder consultation — universities, private sector, "
                             "mayors, industries; PPP between municipalities and private sector.",
    "Coordination_levels": "New act defines central, provincial and local roles; provincial/"
                           "central help acquire landfill land.",
    "Unclear_responsibilities": "Three-tier roles being clarified in new act; municipalities "
                                "need orientation beyond online posting.",
    "Res_nat_government": "MoUD develops policies, acts, regulations; coordinates donors.",
    "Res_prov_government": "Assists local governments with landfill land acquisition under "
                           "new act.",
    "Res_loc_government": "Constitutionally responsible under Local Government Operation Act "
                          "2017.",
    "Res_private_sector": "Domestic and international partners under PPP; foreign technology.",
    "Res_science": "Universities consulted in act formulation.",
    "Res_households": "Source segregation at household level essential.",
    "Res_private_companies": "Producers/industries under proposed EPR.",
    "Cul_households": "No segregation culture; convenience attachment; haphazard disposal.",
    "Cul_loc_government": "Landfills operated as dumping sites (Banchare Danda example).",
    "Cul_private_companies": "Polluter-pays \"implementation is weak.\"",
    "Tar_nat_government": "Coordination, policy framework, donor partnerships.",
    "Tar_prov_government": "Assistance with land acquisition and coordination.",
    "Tar_loc_government": "Implementation, landfill operation, PPP agreements, orientation.",
    "Tar_private_sector": "Investment and technology partnership.",
    "Tar_households": "Source segregation habits; reducing plastic attachment.",
    "Tar_private_companies": "EPR, fee payments, post-enactment industry cooperation.",
    "Actor_role": "Respondent speaks as national policy formulator drafting the Solid Waste "
                 "Management Act, regulations, guidelines and PPP templates — role 1.",
    "Discretion": "See Coded_Data Discretion cell.",
    "Monitoring": "Insufficient monitoring and technology for segregation patterns; "
                  "planned municipal software.",
    "Financial_resources": "Land acquisition difficult under separate Land Acquisition Act; "
                           "circular economy needs financial sustainability support.",
    "Research": "Universities and researchers consulted; welcomes international expertise.",
    "Infrastructure": "Transfer stations exist in Kathmandu; sanitary landfill operation "
                      "failing; PPP infrastructure provisions.",
    "Capacity": "Municipal elected leaders need technical orientation via training centre.",
    "Enforcement": "40-micron ban exists; burning prohibition with penalties in draft act; "
                   "polluter-pays weak.",
    "Pol_epr": "EPR proposed in draft act but not yet enacted.",
    "Pol_awareness": "Bagmati Cleaning Campaign (~10 years); provincial/district "
                     "consultations planned.",
    "Pol_ban": "Government banned plastic thinner than 40 microns; burning solid waste "
              "strictly prohibited in draft act.",
    "Pol_capacity": "Training centre for local-government orientation packages.",
    "Pol_recycling": "Segregation and circular economy provisions; separate plastics policy "
                     "planned.",
    "Pol_waste_collection": "Act covers collection, transportation and disposal; transfer "
                            "stations in Kathmandu.",
    "Pol_effectiveness": "Campaign, segregation, polluter-pays and landfill management all "
                         "weak in practice.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Three-tier role clarification in new act; central donor coordination.",
    "Sol_responsibilities": "EPR for producers; clear three-tier responsibilities.",
    "Sol_epr": "Producers responsible for disposal or pay fee to local body.",
    "Sol_awareness": "Provincial/district consultations; not just website posting.",
    "Sol_segregation": "Primary solution enabling circular economy and waste reduction.",
    "Sol_recycling": "Circular economy after segregation; separate plastics framework later.",
    "Sol_education": "Training centre orientation for municipalities.",
    "Sol_capacity": "Guidelines and regulations post-enactment; municipal software.",
    "Sol_RD": "Foreign technology investment welcomed.",
    "Sol_finance": "EPR fees; circular economy financial sustainability; donor coordination.",
    "Sol_infrastructure": "Sanitary landfill operation; transfer stations; PPP models.",
    "Sol_enforcement": "Burning ban with penalties; EPR enforcement after enactment.",
    "Sol_monitoring": "Municipal software tracking population, production, segregation, disposal.",
    "Traditions to build on (free-hand)": "Leaf plates; organic waste cycle at low population "
                                          "density; separate household bins emerging.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - students not mentioned as responsible actors.",
    "Res_civil_society": "NA - NGOs not described as current responsible actors.",
    "Res_edu_institutions": "NA - universities consulted but not ongoing responsible actors.",
    "Cul_nat_government": "NA - interviewee IS the national government; past weaknesses "
                          "discussed but ministry positions itself as reformer.",
    "Cul_prov_government": "NA - provincial government seen as future helper, not culprit.",
    "Cul_students": "NA - not blamed.",
    "Cul_private_sector": "NA - private sector targeted for partnership, not blamed.",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Tar_students": "NA - not named as primary target group.",
    "Tar_civil_society": "NA - not named.",
    "Tar_science": "NA - researchers consulted but not primary target group.",
    "Tar_edu_institutions": "NA - not named.",
}

creasion_expl = {
    "Actortype": "Creasion is a civil-society/recycling organization (Center for Research and "
                 "Sustainable Development in Nepal) operating plastic recycling facilities and "
                 "supporting informal waste workers and entrepreneurs.",
    "Problem_awareness_pop": "Most Nepali consumers/businesses prefer virgin over recycled "
                             "plastic; unclear market demand for recycled granules.",
    "Problem_awareness_pol": "Creasion not consulted on solid waste management policy "
                            "(gradually being included); unclear national measurement "
                            "methodology (60,000 t/year; 27% PET).",
    "Problem_concerndness": "Organization works directly on plastic end-of-life, food-packaging "
                            "safety, and microplastics — high operational concern.",
    "Problem_recycling": "Only two recycling centers identified; colored PET not effectively "
                         "recyclable; sector standards \"pathetic\"; virgin plastic displaces "
                         "recycled granules for food contact.",
    "Problem_waste_mgmt": "End-of-life plastic fate; informal pickers exploited; bureaucratic "
                          "barriers to establishing/moving recycling industry.",
    "Problem_production": "Virgin plastic preferred; 60% export condition after year three "
                          "constrains recycler business models.",
    "Problem_alternatives": "Need standardization of recycled granule quality and food-contact "
                            "safety; recycled products could be cheaper without scrap tax.",
    "Impacts": "Micro/nano plastics from colored PET; food-contact safety risks; carcinogenic "
               "emissions from poor recyclers; informal picker exploitation.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "informal waste pickers (especially women); public exposed to unsafe "
                     "recycled food plastics and emissions; environment from end-of-life "
                     "plastics. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "intermediary mafias; substandard recyclers; scrap tax; virgin-plastic "
                    "market preference; bureaucratic/export policy barriers; government "
                    "consultation gaps. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: Creasion, "
               "EU project, CAP/World Bank project, municipal formalization of informal "
               "collectors, Plus Nepal/NRR associations. Organization positions itself as "
               "primary problem-solver. No FLAG needed.",
    "Relevance_international_pol": "EU-funded project; CAP project (World Bank, UNOPS, SACEP); "
                                   "60% export condition from government.",
    "Coordination_sectoral": "Collaboration with municipalities, informal/formal associations "
                             "(Plus Nepal, NRR), CSO trainings.",
    "Coordination_levels": "Department of Industry and local government approvals required "
                           "for setup/relocation; municipal scrap tax.",
    "Unclear_responsibilities": "Who buys recycled granules unclear with only two centers; "
                                "policy consultation gap for private recyclers.",
    "Res_nat_government": "Department of Industry (registration, utilities, export condition).",
    "Res_loc_government": "LG approvals for industry setup; municipal scrap tax collection.",
    "Res_private_sector": "Recycling entrepreneurs and formal/informal recycler associations.",
    "Res_civil_society": "Creasion and CSO sector coordinating trainings and advocacy.",
    "Res_households": "Consumer preference for virgin plastic affects market.",
    "Res_private_companies": "Buyers of recycled vs virgin granules; food vs non-food "
                             "applications.",
    "Cul_nat_government": "Export conditionality; not consulting recyclers on policy; unclear "
                          "data methodology.",
    "Cul_loc_government": "Municipal scrap tax raising recycler costs.",
    "Cul_households": "Virgin plastic preference limits recycled market.",
    "Cul_private_sector": "Substandard recyclers with carcinogenic emissions; intermediary "
                          "mafias exploiting pickers.",
    "Tar_private_sector": "Recycling entrepreneurs need capacity-building and fair policy.",
    "Tar_loc_government": "Municipal collaboration to formalize informal collectors.",
    "Tar_households": "Consumer awareness on recycled vs virgin plastic.",
    "Tar_civil_society": "CSO trainings; Plus Nepal and NRR as coordination channels.",
    "Tar_private_companies": "Industry standards and food-contact safety compliance.",
    "Actor_role": "Creasion operates recycling facilities, EU/CAP projects, entrepreneur "
                 "capacity-building and municipal partnerships — managerial/organisational "
                 "implementer (2).",
    "Discretion": "Own processing standards, safety testing, water-recycling ATP, project "
                  "design — but constrained by Dept of Industry/LG bureaucracy and tax/export "
                  "rules.",
    "Monitoring": "Weak sector standards enforcement; unclear national waste quantification.",
    "Financial_resources": "Scrap tax raises costs; operating at competitive disadvantage "
                           "vs virgin plastic; export condition after year 3.",
    "Research": "Questions World Bank PET percentage methodology; developing own granule "
               "standardization.",
    "Infrastructure": "Bureaucratic hurdles for land, water, electricity on setup/relocation; "
                      "ATP and CAP technology as infrastructure solutions.",
    "Capacity": "Building entrepreneur and informal-collector capacity via EU/CAP projects.",
    "Enforcement": "Sector social/environmental standards weakly enforced.",
    "Pol_epr": "EPR not mentioned.",
    "Pol_tax": "Scrap tax on PET bottles paid at municipal level.",
    "Pol_recycling": "Creasion's recycling, downsizing and product-conversion operations.",
    "Pol_effectiveness": "Tax undermines recycled competitiveness; weak standards; "
                         "consultation gap.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Sector coordination via formal/informal associations and CSO "
                       "trainings.",
    "Sol_awareness": "Entrepreneur and informal-worker support programmes.",
    "Sol_recycling": "High-quality granule production with safety testing; CAP technology.",
    "Sol_education": "CSO trainings for formal and informal associations.",
    "Sol_capacity": "EU project entrepreneurship support; CAP entrepreneur capacity-building.",
    "Sol_finance": "Reconsider scrap tax and export conditionality.",
    "Sol_infrastructure": "Streamline Dept of Industry/LG approval processes; ATP water "
                          "recycling model.",
    "Sol_enforcement": "Standardize and enforce social/environmental standards across "
                       "recyclers.",
    "Sol_monitoring": "Clearer national plastic-waste measurement methodology.",
    "Traditions to build on (free-hand)": "NA - not discussed.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_prov_government": "NA - provincial government not mentioned.",
    "Res_students": "NA - not mentioned.",
    "Res_science": "NA - research on methodology questioned but no science actors named.",
    "Res_edu_institutions": "NA - not mentioned.",
    "Cul_prov_government": "NA - not blamed.",
    "Cul_students": "NA - not blamed.",
    "Cul_civil_society": "NA - CSOs portrayed as heroes.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Cul_private_companies": "NA - virgin preference coded under Cul_households/market.",
    "Tar_nat_government": "NA - national level discussed as policy reform need but not "
                          "primary target group in notes.",
    "Tar_prov_government": "NA - not named.",
    "Tar_students": "NA - not named.",
    "Tar_science": "NA - not named.",
    "Tar_edu_institutions": "NA - not named.",
}

silent_park_expl = {
    "Actortype": "Rajesh Aryal, Hotel Silent Park, Chitwan tourism/hotel sector — coded as "
                 "private_companies.",
    "Problem_awareness_pop": "Low awareness of eco-friendly alternatives; ignorance and "
                             "laziness; people do not reuse/recycle properly due to "
                             "convenience; plastic bags given free at shops.",
    "Problem_awareness_pol": "No comprehensive plastics policy (only Environmental Protection "
                            "Act); federal assessing new act but guidelines unclear; unclear "
                            "division of authority across three tiers.",
    "Problem_concerndness": "Guideline notes: very concerned about visible litter and "
                            "invisible microplastic contamination; transcript also flags "
                            "aluminum foil as a particular hotel-level disposal concern.",
    "Problem_littering": "Plastic everywhere urban and rural; bins overflow; wind and "
                         "animals scatter waste; drains and rivers clogged.",
    "Problem_consumption": "Plastic easy and fast; free bags at shops; widespread packaging.",
    "Problem_recycling": "Eco-Green irregular; aluminum foil not collected by recyclers; "
                         "informal sector collects 99% of valuable hotel plastics.",
    "Problem_waste_mgmt": "Irregular municipal/Eco-Green collection; heavy informal-sector "
                          "reliance without formal integration; political instability "
                          "delays implementation.",
    "Problem_production": "Upstream production and packaging implied in policy discussion; "
                          "virgin plastic preference in food-contact applications.",
    "Problem_alternatives": "Cloth bags, clay/cement pots, leaf plates promoted but "
                            "awareness low.",
    "Problem_waste_segregation": "Bottles stored for Eco-Green; wrappers via municipal trucks; "
                                 "PLEASE/Bio-Camp demonstrate segregation pilots.",
    "Impacts": "Animals, burning, microplastics (soil/water/air/crops), health/leaching, "
               "drains, aluminum foil, tourism area.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "animals; poor communities; farmers; public health; tourist/community "
                     "areas. No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "unaware public; weak national policy/enforcement; unclear governance; "
                    "irregular collection; hotels; aluminum foil use; informal sector "
                    "exploitation without integration. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: nature "
               "guides/rangers; PLEASE/Bio-Camp; Solid Waste Management Association; Plast "
               "Foundation; Eco-Green; informal collectors; proposed EPR. No FLAG needed.",
    "Relevance_international_pol": "UN plastic negotiations influence national policy "
                                   "development.",
    "Coordination_sectoral": "Solid Waste Management Association coordinates private "
                             "operators; hotel associations, NGOs, park rangers; "
                             "manufacturers/collectors/recyclers cooperation sought.",
    "Coordination_levels": "Unclear federal/provincial/local roles; municipal vs Eco-Green vs "
                           "informal collectors.",
    "Unclear_responsibilities": "Division of authority between federal, provincial and local "
                                "levels; public sees waste as only government's job.",
    "Res_nat_government": "Federal government sets broad frameworks and assessments; new "
                        "plastics act in development.",
    "Res_prov_government": "Provincial level part of three-tier responsibility framework.",
    "Res_loc_government": "Municipalities/metropolitan cities gradually taking responsibility; "
                          "municipal trucks and dustbins.",
    "Res_private_sector": "Semi-formal private companies; Eco-Green; informal pickers (~99% "
                          "hotel plastics).",
    "Res_civil_society": "Plast Foundation Nepal; NGOs training nature guides.",
    "Res_households": "Source separation and proper disposal expected.",
    "Res_private_companies": "Hotels (thousands in Chitwan); retailers/shops giving free bags.",
    "Cul_nat_government": "No specific plastics policy; weak <40-micron enforcement; political "
                          "instability.",
    "Cul_households": "Low awareness, convenience, waste seen as government's job only.",
    "Cul_loc_government": "Irregular collection; no municipal notices; streets not cleaned "
                          "regularly.",
    "Cul_private_companies": "Hotels lack continuous waste-management action.",
    "Tar_nat_government": "Coordination, EPR implementation, regulate imported plastics, "
                          "support research.",
    "Tar_loc_government": "Inspections, local enforcement, designated street cleaners, "
                          "dustbins.",
    "Tar_private_sector": "Semi-formal collectors, recyclers, hotel committees.",
    "Tar_households": "Awareness, 3R/4R, proper disposal.",
    "Tar_private_companies": "Hotels, small businesses, retailers.",
    "Tar_civil_society": "Nature guides, park rangers, NGO training programmes.",
    "Actor_role": "Hotel operator as target group (role 4) — describes ground-level hotel/"
                 "tourism practices while interview also covers national policy themes.",
    "Discretion": "Limited to own hotel waste-handling practices; no policy or municipal "
                  "authority.",
    "Monitoring": "Weak enforcement; limited testing-lab capacity nationally; no municipal "
                  "notices locally.",
    "Financial_resources": "Local governments face resource gaps; hotel committee rickshaw "
                           "collection (~NPR 50,000/month) proposed; financial support for "
                           "private collection initiatives.",
    "Research": "Research needed on microplastics, burning impacts, agriculture and water "
               "contamination.",
    "Infrastructure": "Need dustbins, shredders, transfer stations; safari buckets; ward "
                      "cleaners.",
    "Capacity": "Limited technical capacity (testing labs); no designated street cleaners; "
               "training needed for collectors and nature guides.",
    "Enforcement": "Weak ban enforcement; ward strict rules proposed.",
    "Pol_epr": "EPR proposed as future solution but not currently in place.",
    "Pol_awareness": "Irregular campaigns; Environment Day one day only; nature-guide "
                     "training.",
    "Pol_education": "Training for waste collectors and nature guides.",
    "Pol_ban": "<40-micron plastics ban weakly enforced.",
    "Pol_subsitutes": "Cloth bags, clay/cement pots, leaf plates.",
    "Pol_clean_up": "Environment Day; safari dustbin rule; elephant-dung teams.",
    "Pol_upcycling": "Bio-Camp pilots (flowerpots, plastic boards).",
    "Pol_recycling": "Eco-Green bottles; PLEASE/Bio-Camp segregation/recycling.",
    "Pol_waste_collection": "Municipal trucks; informal sector; irregular schedules.",
    "Pol_effectiveness": "Limited due to weak enforcement, irregular collection, policy gaps.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Clear three-tier responsibilities; Solid Waste Management "
                       "Association coordination.",
    "Sol_responsibilities": "Clarify federal/provincial/local roles.",
    "Sol_epr": "Implement EPR nationwide.",
    "Sol_awareness": "Sustained public-awareness campaigns; ward programmes.",
    "Sol_segregation": "Household and hotel source segregation; scale PLEASE/Bio-Camp.",
    "Sol_upcycling": "Scale Bio-Camp upcycling models.",
    "Sol_recycling": "Strengthen formal recycling; integrate informal sector.",
    "Sol_education": "3R/4R, lifecycle assessment, eco-friendly alternatives.",
    "Sol_capacity": "Labs, trained workforce, segregation skills, nature-guide training.",
    "Sol_RD": "Research on microplastics, burning, agriculture, water.",
    "Sol_finance": "Support private collection initiatives; hotel committee funding.",
    "Sol_infrastructure": "Dustbins, shredders, transfer stations.",
    "Sol_subsitutes": "Leaf plates, cotton/towel bags, clay pots.",
    "Sol_clean_up": "Designated street cleaners; daily collection model.",
    "Sol_enforcement": "Local inspections; enforce bans.",
    "Sol_monitoring": "Stronger enforcement and data systems.",
    "Traditions to build on (free-hand)": "Banana/pat leaf plates; cotton/towel bags; clay "
                                          "pots; reuse before recycle.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_students": "NA - not mentioned.",
    "Res_science": "NA - research called for but no science actors named as responsible.",
    "Res_edu_institutions": "NA - not mentioned.",
    "Cul_prov_government": "NA - provincial level not blamed specifically.",
    "Cul_students": "NA - not blamed.",
    "Cul_private_sector": "NA - informal/private collectors largely portrayed positively.",
    "Cul_civil_society": "NA - NGOs/guides portrayed as heroes.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - not blamed.",
    "Tar_prov_government": "NA - not named as primary target group.",
    "Tar_students": "NA - not named.",
    "Tar_science": "NA - research support via federal govt, not science actors as targets.",
    "Tar_edu_institutions": "NA - not named.",
}

margen_expl = {
    "Actortype": "Ms. Suna Maya Margen is an Environment Inspector in Kathmandu Metropolitan "
                 "City's Environment Division (Public Education & Inspection) — coded as "
                 "loc_government.",
    "Problem_awareness_pop": "Follow-up Q10 asks what people need to learn more about "
                             "(assessment of the 40-micron notification; microplastics "
                             "impacts) — implying public knowledge gaps despite KMC's public "
                             "education work.",
    "Problem_awareness_pol": "Ministry of Forest and Environment plastics policy is \"not "
                            "implemented and currently under revision\"; KMC has no direct "
                            "plastics ordinance and instead relies on broader environment "
                            "acts — indicating policy-framework gaps.",
    "Problem_concerndness": "Plastic described as \"the biggest problem visually since 1998\" "
                            "with rising multi-layer share (8% to 15%), riverbank dumping "
                            "and drainage blockage. The explicit follow-up on concern level "
                            "was left blank in the source notes, but the overall framing "
                            "reflects high institutional concern.",
    "Problem_littering": "\"Dumped at river-banks\"; drainage blockage; visual pollution.",
    "Problem_consumption": "Visual pollution attributed to \"tourists\" and \"life-style\" — "
                           "ongoing consumption patterns driving visible waste.",
    "Problem_recycling": "\"8% and before now 15%\" multi-layer plastics with \"no value\" — "
                         "low-value plastics not economically recycled.",
    "Problem_waste_mgmt": "\"Sufficient waste but unable to procure land\"; lack of land/"
                          "space within Kathmandu city; 23 municipalities in four districts "
                          "dispose at KMC's landfill.",
    "Problem_production": "Multi-layer plastic share rising (8% to 15%); future solutions "
                          "include \"provide subsidies to producers\" — implying upstream "
                          "production/composition is part of the problem.",
    "Problem_alternatives": "\"Lapti\" leaf plates and other alternatives \"are expensive\"; "
                            "\"alternatives are expensive\" relative to cheap plastic.",
    "Problem_waste_segregation": "\"Lack of land, we can segregate (no space within the "
                                 "Kathmandu city)\" — segregation hampered by space "
                                 "constraints.",
    "Problem_import": "\"Import of plastics??\" raised as an open policy question in future "
                      "measures — suggesting imports may undermine domestic management.",
    "Impacts": "See Coded_Data Impacts cell — visual pollution since 1998, drainage "
               "blockage, riverbank dumping, multi-layer plastic share, microplastics.",
    NPF_VICTIMS_COL: "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                     "rivers/drainage systems, tourism/visual environment, neighbouring "
                     "communities near proposed MRF sites, wildlife and public health "
                     "(microplastics). No FLAG needed.",
    "NPF_villains": "NOT named in explicit NPF narrative language, but clearly IMPLIED: "
                    "lifestyle/tourism consumers; multi-layer plastic producers; unregistered/"
                    "informal sector (excluded from KMC partnerships); NIMBY opposition to "
                    "waste facilities; possible plastic importers. No FLAG needed.",
    "NPF_hero": "NOT named in explicit NPF narrative language, but clearly IMPLIED: KMC "
               "(education, composting, landfill management, riverbed clean-up); Doko "
               "Recyclers; UNDP; stores using cloth bags; Newari Lapti traditions; planned "
               "subsidies and bans. No FLAG needed.",
    "Relevance_international_pol": "UNDP donor partnership (MoU) for plastics recovery and "
                                   "plastic study.",
    "Coordination_sectoral": "KMC coordinates with registered private sector (MoUs), UNDP "
                             "donor partner, and Mayors Forum (18 municipalities) on plastic "
                             "and solid waste management.",
    "Coordination_levels": "\"Land authority is with federal and implementation is by LG, "
                           "its difficult for LG, better coordination and collaboration with "
                           "provincial and district [needed]\"; central/provincial "
                           "governments should help provide land within the metro city or "
                           "valley.",
    "Unclear_responsibilities": "Federal government holds land authority while local "
                                "government must implement; ministry plastics policy under "
                                "revision and not implemented — unclear national-local "
                                "division of roles.",
    "Res_nat_government": "Ministry of Forest and Environment (national plastics policy "
                          "under revision); federal land authority.",
    "Res_prov_government": "\"Central and provincial government provide [land] within metro "
                           "city or in the valley\" — provincial role in land/support "
                           "explicitly invoked.",
    "Res_loc_government": "\"KTM city\" responsible; manages landfill site used by 23 "
                          "municipalities in four districts; implements environment acts and "
                          "solid waste management at city level.",
    "Res_students": "Composting programmes run at school level — students/schools are "
                    "implementation partners.",
    "Res_private_sector": "Registered private sector engaged via MoUs (e.g. Doko Recyclers "
                          "collecting Cluster 7); solid waste management provisions for "
                          "private sector engagement.",
    "Res_households": "Composting promoted at household level as part of KMC public "
                      "education work.",
    "Res_edu_institutions": "School-level composting programmes — schools as implementation "
                            "sites.",
    "Res_private_companies": "Department stores that provide cloth bags and do not give out "
                             "polythene bags — retail companies exercising responsibility.",
    "Cul_nat_government": "Land authority rests with federal government while local "
                          "government must implement; national ministry policy \"not "
                          "implemented and currently under revision.\"",
    "Cul_households": "\"Public opposition – people don't want wastes everywhere\"; "
                      "neighbouring municipalities oppose MRF facilities nearby.",
    "Cul_private_sector": "Historical note that private-sector partners \"were not "
                          "legalized\"; KMC now works only with registered private sector "
                          "and not with the informal sector.",
    "Tar_private_companies": "Q9: \"stores in Nepal but u get cloth bags and don't get "
                              "plastics (polythene)\" — retailers/stores named as focus for "
                              "future policy.",
    "Tar_households": "Household composting programmes and public education on plastics "
                      "imply households as a target group.",
    "Tar_students": "School-level composting and education programmes target students.",
    "Tar_edu_institutions": "Schools are explicit sites for composting and education "
                            "programmes.",
    "Tar_loc_government": "Better rules and regulation \"within our areas\"; MRF siting "
                            "and implementation accountability at municipal level.",
    "Tar_nat_government": "Central government should provide land and revise/implement "
                          "national plastics policy; subsidies to producers.",
    "Tar_prov_government": "Provincial government should help provide land and coordinate "
                           "with local government.",
    "Actor_role": "Environment Inspector conducting public education, composting programmes "
                 "and ground-level implementation — street-level implementer (role 3).",
    "Discretion": "See Coded_Data Discretion cell.",
    "Monitoring": "\"Assessment of the 40m notification\" needed — explicit monitoring/"
                  "evaluation gap for existing ban.",
    "Financial_resources": "Calls for subsidies to producers and to companies producing "
                           "alternatives; donor support via UNDP MoU.",
    "Research": "UNDP MoU includes \"plastic study\"; microplastics impacts need more "
               "public knowledge.",
    "Infrastructure": "\"Technology, recycling factory\" lacking; need MRF and land from "
                      "central/provincial government.",
    "Capacity": "Technology and recycling-factory capacity gaps explicitly named among "
               "main challenges.",
    "Enforcement": "\"Rules & regulation within our areas\" and need to assess 40-micron "
                   "notification — implying enforcement/monitoring gaps.",
    "Pol_epr": "Producer subsidies discussed but no extended producer responsibility "
              "framework mentioned — coded no.",
    "Pol_import": "Import of plastics raised only as an open question (\"import of "
                  "plastics??\") with no effective import-control mechanism described.",
    "Pol_awareness": "Interviewee's public-education role; KMC runs awareness-oriented "
                     "programmes on plastics reduction and composting.",
    "Pol_education": "Public education is the interviewee's field of work; composting "
                     "education at household, community and school levels.",
    "Pol_ban": "\"Single-use plastics ban — planning\" — ban under development, not yet "
              "fully operational.",
    "Pol_subsitutes": "Stores provide cloth bags instead of polythene; planned government "
                      "subsidies for companies producing alternatives; Lapti leaf plates.",
    "Pol_clean_up": "\"We did collaborated with wastes at debris near river beds "
                    "(excavate and transfer)\" — active clean-up of riverbed debris.",
    "Pol_recycling": "Doko Recyclers MoU collecting from Cluster 7; Mayors Forum provisions "
                     "for plastic management.",
    "Pol_waste_collection": "Cluster-based collection via registered private partners; "
                            "landfill site management for 23 municipalities.",
    "Pol_effectiveness": "National policy not implemented; ban only planning stage; no 40-m "
                         "assessment; land/technology gaps limit effectiveness.",
    "Pol_effectiveness_example": "See Pol_effectiveness — quoted directly in Coded_Data cell.",
    "Sol_lead_agency": "Better coordination between federal (land authority), provincial, "
                       "district and local government; action plan referenced.",
    "Sol_responsibilities": "Clarify federal land authority vs local implementation; "
                            "Mayors Forum and ministry revision to assign roles.",
    "Sol_awareness": "Continued public education on microplastics and 40-micron "
                     "notification assessment.",
    "Sol_segregation": "Segregation desired but blocked by lack of land/space — future "
                       "solution requires land and MRF infrastructure.",
    "Sol_recycling": "Recycling factory/technology and MRF needed.",
    "Sol_education": "Public education on microplastics, health/environment/wildlife "
                     "impacts and composting.",
    "Sol_capacity": "Technology and recycling-factory capacity building.",
    "Sol_ban": "Single-use plastics ban in planning.",
    "Sol_finance": "Subsidies to producers; government subsidies for alternative-product "
                   "companies; UNDP donor support.",
    "Sol_infrastructure": "MRF and land provision by central/provincial government within "
                          "metro city or valley.",
    "Sol_subsitutes": "Subsidise companies producing alternatives; promote Lapti and cloth "
                      "bags.",
    "Sol_clean_up": "Riverbed debris excavation and transfer (ongoing collaborative work).",
    "Sol_enforcement": "\"Rules & regulation within our areas\"; assess and enforce "
                      "40-micron notification.",
    "Sol_monitoring": "Assessment of 40-micron notification; plastic study with UNDP.",
    "Traditions to build on (free-hand)": "Newari festival leaf plates (Lapti) — culturally "
                                          "rooted plastic reduction but cost barrier.",
    # --- Explanations for otherwise-NA values in the "always explain" column set ---
    "Res_civil_society": "NA - Doko Recyclers coded under private sector; no NGOs/CSOs "
                         "named as responsible actors.",
    "Res_science": "NA - research needs mentioned but no scientific institutions named as "
                   "responsible actors.",
    "Cul_prov_government": "NA - provincial government portrayed as needed helper, not "
                           "primarily blamed.",
    "Cul_loc_government": "NA - interviewee represents KMC; municipality portrayed as "
                           "implementer facing external constraints rather than primary "
                           "culprit.",
    "Cul_students": "NA - not blamed.",
    "Cul_civil_society": "NA - not blamed.",
    "Cul_science": "NA - not blamed.",
    "Cul_edu_institutions": "NA - schools are programme sites, not blamed.",
    "Cul_private_companies": "NA - stores with cloth bags portrayed positively; importers "
                             "only raised as open question.",
    "Tar_private_sector": "NA - registered private partners discussed as collaborators "
                          "rather than primary policy targets.",
    "Tar_civil_society": "NA - not named as target group.",
    "Tar_science": "NA - not named.",
}

EXPLANATIONS = {
    "NPL_1": doe_expl,
    "NPL_2": ganesh_expl,
    "NPL_3": ktm_expl,
    "NPL_4": dhulikhel_expl,
    "NPL_5": doco_expl,
    "NPL_6": hotel_expl,
    "NPL_7": rsct_expl,
    "NPL_8": moud_expl,
    "NPL_9": ward2_expl,
    "NPL_10": ward7_expl,
    "NPL_11": safaurja_expl,
    "NPL_12": moud12_expl,
    "NPL_13": creasion_expl,
    "NPL_14": silent_park_expl,
    "NPL_15": margen_expl,
}

# ---------------------------------------------------------------------------
# 3b. Coded interview registry — only fully coded interviews are included in
#     the workbook (no placeholder rows).
# ---------------------------------------------------------------------------
CODED_ROWS = {
    "NPL_1": doe,
    "NPL_2": ganesh,
    "NPL_3": ktm,
    "NPL_4": dhulikhel,
    "NPL_5": doco,
    "NPL_6": hotel,
    "NPL_7": rsct,
    "NPL_8": moud,
    "NPL_9": ward2,
    "NPL_10": ward7,
    "NPL_11": safaurja,
    "NPL_12": moud12,
    "NPL_13": creasion,
    "NPL_14": silent_park,
    "NPL_15": margen,
}

# Master-list order for coded interviews only (NPL_1 through NPL_15).
CODED_ORDER = [
    "NPL_1", "NPL_2", "NPL_3", "NPL_4", "NPL_5", "NPL_6", "NPL_7", "NPL_8", "NPL_9",
    "NPL_10", "NPL_11", "NPL_12", "NPL_13", "NPL_14", "NPL_15",
]

# (id, label as given by the research team, affiliation)
CODED_META = {
    "NPL_1": ("1. Environmental Department/ Deepak Diwal", "Department of Environment"),
    "NPL_2": ("2. Former Minister/ Ganesh Shah", "Former Minister, Government of Nepal"),
    "NPL_3": ("3. KTM Municipal Office", "Kathmandu Metropolitan City - Environment / Solid Waste Management Division"),
    "NPL_4": ("4. Mayor Dhulikhel/ Ashok Kumar Byanju Shrestha", "Dhulikhel Municipality (Mayor's Office)"),
    "NPL_5": ("5. Private Sector/ Doco Recyclers", "Doco Recyclers"),
    "NPL_6": ("6. Restaurant and Hotel Owner Dhulikhel_Combined", "Dhulikhel Hotel / Restaurant"),
    "NPL_7": ("7. Rural Self-Reliance Development Center/ Narayan Nirola", "Rural Self-Reliance Development Center"),
    "NPL_8": ("8. Urban Development Ministry/ Kamal Adhikar", "Ministry of Urban Development, Government of Nepal"),
    "NPL_9": ("9. Dhulikhel Ward no.1", "Dhulikhel Municipality - Ward No. 1"),
    "NPL_10": ("10. Dhulikhel Ward no.7", "Dhulikhel Municipality - Ward No. 7"),
    "NPL_11": ("11. Private Sector Contractor/ SafaUrja Utpadan", "Safa Urja Utpadan"),
    "NPL_12": ("12. Ministry for Urban Planning", "Ministry of Urban Development"),
    "NPL_13": ("13. Creasion", "Center for Research and Sustainable Development in Nepal"),
    "NPL_14": ("14. Hotel Silent Park/ Rajesh Aryal", "Hotel Silent Park"),
    "NPL_15": ("15. Kathmandu Metropolitan City/ Suna Maya Margen",
               "Kathmandu Metropolitan City - Environment Division / Public Education & Inspection"),
}

NAMES = {iid: f"{label} - {affil}" for iid, (label, affil) in CODED_META.items()}
# Fuller descriptive names for Coding_Explanations sheet.
NAMES["NPL_1"] = ("Department of Environment - Deepak Diwali, Deputy Director, Pollution "
                   "Control (air & plastics), Kathmandu (Interview no. 5, 23 June 2025)")
NAMES["NPL_2"] = ("Former Minister, Government of Nepal - Ganesh Shah, Kathmandu "
                   "(21 June 2025)")
NAMES["NPL_3"] = ("Kathmandu Metropolitan City (KMC), Solid Waste Management Office - "
                   "Municipality Officer(s), Kathmandu (23 June)")
NAMES["NPL_4"] = ("Dhulikhel Municipality - Mayor Ashok Kumar Byanju Shrestha, "
                   "Dhulikhel (22 June 2025, Kathmandu interview no. 4)")
NAMES["NPL_5"] = ("Doco Recyclers - Private Sector waste-management/recycling company, "
                   "Kathmandu Valley (23 June 2025)")
NAMES["NPL_6"] = ("Dhulikhel Hotel/Restaurant - Mr. Swasti Byanju, hotel owner, "
                   "Dhulikhel (22 June 2025, Kathmandu interview no. 3)")
NAMES["NPL_7"] = ("Rural Self-Reliance Development Center (RSDC/RSCT) - Narayan Nirola "
                   "(project coordinator) + colleague, Kathmandu (23 June 2025)")
NAMES["NPL_8"] = ("Ministry of Urban Development - Kamal Adhikar (Senior Sociologist), "
                   "Nawaraj (Joint Secretary) + colleagues, Kathmandu (23 June 2025, "
                   "Interview 2)")
NAMES["NPL_9"] = ("Dhulikhel Municipality Ward Office - Ward Head + ward staff (E, Deep, F), "
                   "Dhulikhel (Ward No. 2 per transcript; master list: Ward no. 1)")
NAMES["NPL_10"] = ("Dhulikhel Municipality - Ward Chairman (32 years, five terms), "
                    "Dhulikhel (Ward No. 3 per transcript; master list: Ward no. 7)")
NAMES["NPL_11"] = ("Safa Urja Utpadan - Private waste-management contractor, Khaireni/Chitwan "
                    "(Ratnanagar, Khaireni & Kalika municipalities)")
NAMES["NPL_12"] = ("Ministry of Urban Development / Ministry for Urban Planning - MoUD "
                    "policy official(s), Kathmandu (PEGO interview)")
NAMES["NPL_13"] = ("Creasion - Recycling and waste-management CSO, Chitwan area "
                    "(23 June 2025, ~5:15pm)")
NAMES["NPL_14"] = ("Hotel Silent Park - Rajesh Aryal, Chitwan/Sauraha tourism area "
                    "(26 June 2025)")
NAMES["NPL_15"] = ("Kathmandu Metropolitan City - Ms. Suna Maya Margen (30 years), "
                    "Environment Inspector, Public Education & Inspection, Kathmandu "
                    "(23 June 2025, internal interview no. 2)")

INTERVIEWS = [CODED_ROWS[iid] for iid in CODED_ORDER]

# ---------------------------------------------------------------------------
# 3c. Columns for which an explanation must ALWAYS be recorded (per the
#     research team's explicit instruction), even when the coded value is
#     NA.
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

def get_explanation(iid, var, value):
    expl = EXPLANATIONS.get(iid)
    if expl and var in expl:
        return expl[var]
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
#     wheel.
# ---------------------------------------------------------------------------
import colorsys


def make_color_map(ids):
    n = len(ids)
    colors = {}
    for i, iid in enumerate(ids):
        hue = i / n
        r, g, b = colorsys.hls_to_rgb(hue, 0.78, 0.65)
        colors[iid] = "{:02X}{:02X}{:02X}".format(int(r * 255), int(g * 255), int(b * 255))
    return colors


COLOR_MAP = make_color_map(CODED_ORDER)


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
     "interview list. 'Coded_Data' contains one row per fully coded interview "
     "(NPL_1 through NPL_15, in that order). Unfilled placeholder rows are not included. "
     "Column names and order follow exactly the variable list supplied by the research team.",
     False),
    ("", False),
    ("Coded interviews:", True),
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
    ("  5. NPL_5 - Doco Recyclers, private-sector waste-management/recycling company "
     "(dry waste and e-waste; MRFs at Sanothimi and Salghari), Kathmandu Valley (23 June "
     "2025). Two respondents in the transcript (names not given). Coded from both the "
     "interview guideline notes and the full verbatim transcript. Note: Q4a/Q4b "
     "ministry-template content in the notes was excluded (see Notes cell).", False),
    ("  6. NPL_6 - Dhulikhel Hotel/Restaurant Owner, Mr. Swasti Byanju, Dhulikhel "
     "(22 June 2025; Kathmandu interview no. 3). Coded from both the interview guideline "
     "notes and the full verbatim transcript. Swasti Byanju (local collaborator) also "
     "contributed the Dankhuta waste-management example in the transcript.", False),
    ("  7. NPL_7 - Rural Self-Reliance Development Center (RSDC/RSCT), Narayan Nirola "
     "(project coordinator) and a WASH-specialist colleague, Kathmandu (23 June 2025, "
     "12:52pm–1:30pm; internal interview no. 4). NGO/cooperative network (est. 1991) "
     "piloting urban plastic/waste management in Budhanilkantha. Coded from both the "
     "interview guideline notes and the full verbatim transcript.", False),
    ("  8. NPL_8 - Ministry of Urban Development (MoUD), Kamal Adhikar (Senior Sociologist), "
     "Nawaraj (Joint Secretary) and colleagues including a Senior Division Engineer, "
     "Kathmandu (23 June 2025, 10:55am–12:00pm; internal interview no. 2). Coded from "
     "both the interview guideline notes and the full verbatim transcript. NPF victims, "
     "villains and hero are not named using explicit NPF narrative labels but are clearly "
     "implied through attributed harms, blamed actors and proposed problem-solvers (see "
     "Coding_Explanations).", False),
    ("  9. NPL_9 - Dhulikhel Municipality Ward Office, Ward Head + ward staff (E, Deep, F), "
     "Dhulikhel. Master list label: Ward no. 1; transcript focal ward: Ward No. 2. "
     "Interviewers: Ram Devi (Kathmandu University) and PEGO team. Coded from both the "
     "interview guideline notes and the full verbatim transcript.", False),
    ("  10. NPL_10 - Dhulikhel Municipality Ward Chairman (32 years, five terms), "
     "Dhulikhel. Master list label: Ward no. 7; transcript focal ward: Ward No. 3. "
     "Coded from verbatim transcript (theme notes at top of source file excluded).", False),
    ("  11. NPL_11 - Safa Urja Utpadan, private waste-management contractor, Khaireni/"
     "Chitwan. Coded from notes and transcript.", False),
    ("  12. NPL_12 - Ministry for Urban Planning / MoUD. Coded from notes and transcript.", False),
    ("  13. NPL_13 - Creasion, recycling CSO. Coded from notes and field notes.", False),
    ("  14. NPL_14 - Hotel Silent Park, Rajesh Aryal, Chitwan (26 June 2025). Coded from "
     "full interview package (guideline policy themes and verbatim transcript). NPF fields "
     "implied.", False),
    ("  15. NPL_15 - Kathmandu Metropolitan City, Ms. Suna Maya Margen (Environment Inspector, "
     "Public Education), Kathmandu (23 June 2025, internal interview no. 2). Coded from "
     "interview guideline/field notes only (no full verbatim transcript). Overlaps thematically "
     "with NPL_3 (same municipality) but distinct respondent.", False),
    ("", False),
    ("Sheets in this workbook:", True),
    ("  - Codebook: the variable dictionary, listed in the exact same order as the columns "
     "in 'Coded_Data'.", False),
    ("  - Coded_Data: wide-format matrix, one row per coded interview (NPL_1-NPL_15, in "
     "order), one column per codebook variable, including a free-hand 'Notes' column. Each "
     "interview's row is filled with a distinct colour so interviews are easy to tell apart "
     "at a glance.", False),
    ("  - Coding_Explanations: long-format table giving the quote/observation behind each "
     "coded value, one row per Interview x Variable. As requested, the following columns "
     "ALWAYS have an explanation row for every interview, even when the value is NA: "
     "NPF_victims, NPF_villains, NPF_hero, all Res_/Cul_/Tar_ actor-grid columns, "
     "Actor_role, Discretion, Capacity, Pol_effectiveness and Pol_effectiveness_example. "
     "For NPF fields, explanations explicitly note when victims/villains/hero are implied "
     "rather than named, or FLAG when none are mentioned directly or even implied. Other "
     "columns are documented only where a specific explanation was written. Rows are "
     "colour-matched to the same interview colour used in 'Coded_Data'.", False),
    ("", False),
    ("Colour key: each interview (NPL_1 through NPL_14) has a distinct colour spread "
     "evenly across the colour wheel so adjacent IDs are easy to tell apart.", False),
    ("", False),
    ("Coding conventions:", True),
    ("  - 'yes' / 'no' / 'NA': NA means the topic was not addressed in that interview "
     "(no evidence either way); 'no' is only used where the interviewee explicitly said "
     "the item is not an issue / not in place (e.g. Pol_epr = no where EPR is proposed "
     "but not yet enacted).", False),
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
    ("Note on source material: an earlier version of this workbook coded NPL_7 from short "
     "bullet-point notes only. That row has now been fully recoded from the complete "
     "interview guideline notes and verbatim transcript supplied for this respondent.", False),
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
print(f"Interviews: {len(INTERVIEWS)} coded (NPL_1 through NPL_{len(INTERVIEWS)})")
