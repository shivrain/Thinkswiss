"""
Generates a codebook-based qualitative coding of two Nepal plastic-pollution
governance interviews (RSCT and the Department of Environment / Interview No. 5)
into a single Excel workbook with three sheets:

  1. Codebook            - the variable dictionary (as supplied)
  2. Coded_Data          - wide-format matrix: one row per interview, one
                           column per codebook variable
  3. Coding_Explanations - long-format table (Interview | Variable | Value |
                           Evidence / Explanation) documenting, for every
                           variable that was coded (i.e. not left as NA),
                           the quote/observation used to justify the code.

Source material:
  - Interview guideline + hand-written notes for "RSCT" (a Nepali savings &
    credit cooperative network that recently started urban plastics work).
  - Interview No. 5 notes + full transcript with Deepak Diwal(i), Deputy
    Director, Pollution Control (air & plastics), Department of Environment
    (DoE), Government of Nepal - 23 June 2025.

Coding conventions used:
  - yes / no / NA  -> NA means the topic was not addressed in that interview
    (absence of evidence), "no" is only used when the interviewee explicitly
    indicated the item is not an issue / not in place.
  - Free-hand fields are filled with short descriptive text or "NA".
  - Actor_role: 1 = formulation/policy actor, 2 = managerial/organisational
    implementer, 3 = street-level implementer, 4 = target group (comma
    separated if more than one applies).
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT_PATH = "/workspace/nepal_plastic_coding/Nepal_Plastic_Governance_Coding.xlsx"

# ---------------------------------------------------------------------------
# 1. CODEBOOK (variable dictionary, transcribed from the supplied codebook)
# ---------------------------------------------------------------------------
CODEBOOK = [
    ("Country", "Country name", "Nepal"),
    ("ID", "InterviewID: ISO_Nr", "NPL_1"),
    ("Actortype", "nat_government, prov_government, loc_government, students, "
                  "private_sector, civil_society, science, households, "
                  "edu_institutions, private_companies", "civil_society"),
    ("--- Problems ---", "", ""),
    ("Problem_awareness_pop", "Lack of awareness among population was mentioned: yes/no", ""),
    ("Problem_awareness_pol", "Lack of awareness among policy makers was mentioned: yes/no", ""),
    ("Problem_concerndness", "How concerned is the interview partner: high, medium, low, NA", ""),
    ("Problem_littering", "Littering is a problem; was mentioned: yes/no", ""),
    ("Problem_consumption", "High consumption is a problem (midstream); was mentioned: yes/no", ""),
    ("Problem_recycling", "No or too little recycling is a problem; was mentioned: yes/no", ""),
    ("Problem_waste_mgmt", "Ill waste management is a problem; was mentioned: yes/no", ""),
    ("Problem_production", "High production rates (too much production of plastic products - "
                           "upstream); was mentioned: yes/no", ""),
    ("Problem_alternatives", "No good alternatives; was mentioned: yes/no", ""),
    ("Problem_waste_segregation", "Lack of segregation was mentioned: yes/no", ""),
    ("Problem_import", "Plastic imports are a problem; was mentioned: yes/no", ""),
    ("Problem_impacts", "Free hand - impacts mentioned on water, cities, biodiversity, health, "
                        "etc. or NA (examples: water, cities, biodiversity, health, visual "
                        "pollution, rivers, drainage blockage, forests, aquatic life, foul smell, "
                        "dumping sites, air pollution, microplastics, marine pollution, tourism, "
                        "coral damage, agriculture, wildlife, workforce, climate, soil)", ""),
    ("NPF_victims", "Free hand - who are the main victims mentioned "
                    "(e.g. human, environment, marine species, etc.)", ""),
    ("NPF_villains", "Free hand - who are the bad guys/culprits contributing most to plastic "
                     "pollution mentioned by interview partners (e.g. Chinese tourists, "
                     "households, sea nomads, etc.)", ""),
    ("NPF_hero", "Free hand - who is perceived as the actor committed to solving the problem "
                "(e.g., the children will educate their parents)", ""),
    ("--- Governance (NPF - Narrative Policy Framework) ---", "", ""),
    ("Governance", "mentioned: yes/no", ""),
    ("Relevance_international_pol", "International treaties are relevant for national level policy", ""),
    ("Coordination_sectoral", "Lack of coordination across sectors", ""),
    ("Coordination_levels", "Lack of coordination across levels", ""),
    ("Unclear_responsibilities", "Unclear responsibilities among the involved actors", ""),
    ("--- Actors I - Responsibility: Who is responsible to mitigate plastic pollution? (yes/no) ---", "", ""),
    ("Res_Nat_government", "National", ""),
    ("Res_Prov_government", "State/province (e.g. Sabah)", ""),
    ("Res_Loc_government", "Municipality and below (e.g. Semporna)", ""),
    ("Res_Students", "University and school level", ""),
    ("Res_Private_sector", "Waste collectors, recyclers (connected to waste)", ""),
    ("Res_Civil_society", "NGOs, IGOs, monks, etc.", ""),
    ("Res_Science", "Researchers", ""),
    ("Res_Households", "Individuals/wider population", ""),
    ("Res_Edu_institutions", "Schools, universities", ""),
    ("Res_Private_companies", "Hotels, shops, tour guides, cafes, restaurants, etc.", ""),
    ("--- Actors II - Culprit: Who is responsible (e.g. due to government failure/polluter) for the bad situation? (yes/no) ---", "", ""),
    ("Cul_Nat_government", "National", ""),
    ("Cul_Prov_government", "State/province (e.g. Sabah)", ""),
    ("Cul_Loc_government", "Municipality and below (e.g. Semporna)", ""),
    ("Cul_Students", "University and school level", ""),
    ("Cul_Private_sector", "Waste collectors, recyclers (connected to waste)", ""),
    ("Cul_Civil_society", "NGOs, IGOs, monks, etc.", ""),
    ("Cul_Science", "Researchers", ""),
    ("Cul_Households", "Individuals/wider population", ""),
    ("Cul_Edu_institutions", "Schools, universities", ""),
    ("Cul_Private_companies", "Hotels, shops, tour guides, cafes, etc.", ""),
    ("--- Actors III - Target group: Who should be the target group of measures to mitigate plastic pollution? (yes/no) ---", "", ""),
    ("Tar_Nat_government", "National", ""),
    ("Tar_Prov_government", "State/province (e.g. Sabah)", ""),
    ("Tar_Loc_government", "Municipality and below (e.g. Semporna)", ""),
    ("Tar_Students", "University and school level", ""),
    ("Tar_Private_sector", "Waste collectors, recyclers (connected to waste)", ""),
    ("Tar_Civil_society", "NGOs, IGOs, monks, etc.", ""),
    ("Tar_Science", "Researchers", ""),
    ("Tar_Households", "Individuals/wider population", ""),
    ("Tar_Edu_institutions", "Schools, universities", ""),
    ("Tar_Private_companies", "Hotels, shops, tour guides, cafes, etc.", ""),
    ("--- Actor role & discretion ---", "", ""),
    ("Actor_role", "Can be more than one selection, comma separated. 1 = formulation/policy actor "
                  "- not an implementer (e.g. ministry, governmental agency); "
                  "2 = managerial/organisational implementer (not street-level, e.g. district "
                  "office); 3 = street level implementer (e.g. village heads, local NGO staff, "
                  "heads of youth organisations, etc.); 4 = target group (e.g. shops, "
                  "households, etc.). Must be derived from the interview, not the actor group.", ""),
    ("Discretion", "If Actor_role = 2 or 3 (implementer) - free hand - does the interview partner "
                  "indicate a certain degree of discretion, e.g. private sector waste collectors "
                  "developing their own schedule/collection plan.", ""),
    ("--- Implementation issues (mentioned: yes/no) ---", "", ""),
    ("Implementation_issues", "Implementation issues mentioned: yes/no", ""),
    ("Monitoring", "Lack of monitoring", ""),
    ("Financial_resources", "Lack of financial resources", ""),
    ("Research", "Lack of research", ""),
    ("Infrastructure", "Lack of infrastructure", ""),
    ("Capacity", "Lack of capacity of implementers", ""),
    ("Enforcement", "Lack of enforcement", ""),
    ("--- Policies in place (mentioned that a policy is currently in place, based on their knowledge: yes/no) ---", "", ""),
    ("Policies_in_place", "Policies in place mentioned: yes/no", ""),
    ("Pol_epr", "Extended producer responsibility", ""),
    ("Pol_import", "Import regulations", ""),
    ("Pol_awarness", "Awareness campaigns", ""),
    ("Pol_education", "Education programs", ""),
    ("Pol_capacity", "Capacity building", ""),
    ("Pol_RD", "Research & development", ""),
    ("Pol_tax", "Levies or taxes on specific products (often bags)", ""),
    ("Pol_ban", "Bans of specific products", ""),
    ("Pol_subsitutes", "Alternatives like reusable bags, or banana leaf plates", ""),
    ("Pol_clean_up", "Clean-up campaigns", ""),
    ("Pol_upcycling", "Making a new product, like blacktop", ""),
    ("Pol_recycling", "Making a new raw material", ""),
    ("Pol_waste_collection", "Waste collection", ""),
    ("Pol_effectiveness", "Lack of policy effectiveness", ""),
    ("Pol_effectiveness_expample", "Freehand, provide example or copy/paste from the interview "
                                   "what they said.", ""),
    ("--- Solutions (mentioned as solution to overcome plastic pollution/strengthen plastic policy: yes/no) ---", "", ""),
    ("Solutions", "Solutions mentioned: yes/no", ""),
    ("Sol_lead_agency", "Procedural measures to establish lead agency", ""),
    ("Sol_responsibilities", "Clear responsibilities", ""),
    ("Sol_epr", "Extended producer responsibility", ""),
    ("Sol_awarness", "Awareness raising measures", ""),
    ("Sol_segregation", "Segregation of waste", ""),
    ("Sol_upcycling", "Upcycling of plastics", ""),
    ("Sol_recycling", "New raw materials", ""),
    ("Sol_education", "Education programs at schools", ""),
    ("Sol_capacity", "Capacity-building programs", ""),
    ("Sol_RD", "Research programs", ""),
    ("Sol_tax", "Taxes or levies on products", ""),
    ("Sol_ban", "Ban of products", ""),
    ("Sol_finance", "Financial measures", ""),
    ("Sol_infrastructure", "Infrastructural measures", ""),
    ("Sol_subsitutes", "Alternatives", ""),
    ("Sol_clean_up", "Clean-up campaigns", ""),
    ("Sol_enforcement", "Enforcement procedures", ""),
    ("Sol_monitoring", "Monitoring procedures", ""),
    ("Traditions_to_build_on", "Traditions to build on (free-hand) or NA", ""),
]

# Ordered list of the *actual* data columns (excludes the "--- section ---" headers)
DATA_VARS = [row[0] for row in CODEBOOK if not row[0].startswith("---")]

# ---------------------------------------------------------------------------
# 2. CODED DATA - one dict per interview: variable -> coded value
# ---------------------------------------------------------------------------

rsct = {
    "Country": "Nepal",
    "ID": "NPL_1",
    "Actortype": "civil_society",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "NA",
    "Problem_concerndness": "medium",
    "Problem_littering": "yes",
    "Problem_consumption": "NA",
    "Problem_recycling": "NA",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "yes",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "NA",
    "Problem_import": "NA",
    "Problem_impacts": "air pollution (open burning), drainage/sewer blockage, dumping "
                       "sites (landfills), health (indoor air pollution from burning)",
    "NPF_victims": "general public / urban residents (health effects of burning), "
                   "local environment in Chitwan district",
    "NPF_villains": "NA (no specific actor blamed; general public ignorance and lack of "
                    "political will/fund disbursement referenced instead)",
    "NPF_hero": "NA (not explicitly named)",

    "Governance": "yes",
    "Relevance_international_pol": "NA",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "NA",

    "Res_Nat_government": "NA",
    "Res_Prov_government": "NA",
    "Res_Loc_government": "yes",
    "Res_Students": "NA",
    "Res_Private_sector": "yes",
    "Res_Civil_society": "yes",
    "Res_Science": "NA",
    "Res_Households": "NA",
    "Res_Edu_institutions": "NA",
    "Res_Private_companies": "NA",

    "Cul_Nat_government": "yes",
    "Cul_Prov_government": "NA",
    "Cul_Loc_government": "NA",
    "Cul_Students": "NA",
    "Cul_Private_sector": "NA",
    "Cul_Civil_society": "NA",
    "Cul_Science": "NA",
    "Cul_Households": "yes",
    "Cul_Edu_institutions": "NA",
    "Cul_Private_companies": "NA",

    "Tar_Nat_government": "NA",
    "Tar_Prov_government": "NA",
    "Tar_Loc_government": "yes",
    "Tar_Students": "NA",
    "Tar_Private_sector": "yes",
    "Tar_Civil_society": "NA",
    "Tar_Science": "NA",
    "Tar_Households": "yes",
    "Tar_Edu_institutions": "NA",
    "Tar_Private_companies": "NA",

    "Actor_role": "2,3",
    "Discretion": "RSCT designs its own bottom-up cooperative loan/fund mechanisms across "
                 "its 200+ member cooperatives and recently redirected this model to urban "
                 "plastics work, indicating discretion in programme design.",

    "Implementation_issues": "yes",
    "Monitoring": "yes",
    "Financial_resources": "yes",
    "Research": "yes",
    "Infrastructure": "no",
    "Capacity": "yes",
    "Enforcement": "NA",

    "Policies_in_place": "yes",
    "Pol_epr": "NA",
    "Pol_import": "NA",
    "Pol_awarness": "NA",
    "Pol_education": "NA",
    "Pol_capacity": "NA",
    "Pol_RD": "NA",
    "Pol_tax": "NA",
    "Pol_ban": "NA",
    "Pol_subsitutes": "NA",
    "Pol_clean_up": "NA",
    "Pol_upcycling": "NA",
    "Pol_recycling": "NA",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_expample": "\"Fund is there; they don't release the funds\" / "
                                  "\"Two funds: ear-marked; free to use (if you don't use it "
                                  "it will be frozen)\" - illustrates how earmarked funds go "
                                  "unused, undermining policy effectiveness; also \"Implementation "
                                  "is always the problem: monitoring\" and \"Policies are there\" "
                                  "(but not enforced).",

    "Solutions": "yes",
    "Sol_lead_agency": "NA",
    "Sol_responsibilities": "NA",
    "Sol_epr": "NA",
    "Sol_awarness": "yes",
    "Sol_segregation": "NA",
    "Sol_upcycling": "NA",
    "Sol_recycling": "NA",
    "Sol_education": "NA",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_tax": "NA",
    "Sol_ban": "NA",
    "Sol_finance": "yes",
    "Sol_infrastructure": "NA",
    "Sol_subsitutes": "yes",
    "Sol_clean_up": "NA",
    "Sol_enforcement": "NA",
    "Sol_monitoring": "NA",

    "Traditions_to_build_on": "RSCT's existing bottom-up, grassroots cooperative "
                              "fund/loan model (built since 1991 through a two-tier "
                              "structure of 200+ member cooperatives governed by a "
                              "general assembly) could be adapted to finance and "
                              "organise community-level plastic waste management, "
                              "as RSCT itself has already begun doing in urban areas.",
}

doe = {
    "Country": "Nepal",
    "ID": "NPL_5",
    "Actortype": "nat_government",

    "Problem_awareness_pop": "yes",
    "Problem_awareness_pol": "NA",
    "Problem_concerndness": "high",
    "Problem_littering": "NA",
    "Problem_consumption": "NA",
    "Problem_recycling": "yes",
    "Problem_waste_mgmt": "yes",
    "Problem_production": "NA",
    "Problem_alternatives": "yes",
    "Problem_waste_segregation": "yes",
    "Problem_import": "NA",
    "Problem_impacts": "rivers, lakes, land, forests, water leakage/drainage, "
                       "microplastics (including residues found in Himalayan snow)",
    "NPF_victims": "rivers/lakes/water bodies, forests and land, the Himalayan "
                   "ecosystem (microplastics in snow), and by extension public health",
    "NPF_villains": "plastic-producing industries without extended producer "
                    "responsibility (\"there is no producer responsibility in "
                    "Nepal\"), informal/unregistered industries operating without "
                    "government notice, and households/consumers whose habits and "
                    "price-sensitivity keep demand for cheap conventional plastic bags high",
    "NPF_hero": "Department of Environment and the high-level committee (chaired by "
               "the Chief Secretary) driving bans/subsidies; local governments and "
               "NGOs operating material recovery facilities and awareness campaigns",

    "Governance": "yes",
    "Relevance_international_pol": "NA",
    "Coordination_sectoral": "yes",
    "Coordination_levels": "yes",
    "Unclear_responsibilities": "yes",

    "Res_Nat_government": "yes",
    "Res_Prov_government": "NA",
    "Res_Loc_government": "yes",
    "Res_Students": "NA",
    "Res_Private_sector": "yes",
    "Res_Civil_society": "yes",
    "Res_Science": "NA",
    "Res_Households": "NA",
    "Res_Edu_institutions": "yes",
    "Res_Private_companies": "NA",

    "Cul_Nat_government": "NA",
    "Cul_Prov_government": "NA",
    "Cul_Loc_government": "NA",
    "Cul_Students": "NA",
    "Cul_Private_sector": "yes",
    "Cul_Civil_society": "NA",
    "Cul_Science": "NA",
    "Cul_Households": "yes",
    "Cul_Edu_institutions": "NA",
    "Cul_Private_companies": "yes",

    "Tar_Nat_government": "yes",
    "Tar_Prov_government": "NA",
    "Tar_Loc_government": "yes",
    "Tar_Students": "NA",
    "Tar_Private_sector": "yes",
    "Tar_Civil_society": "NA",
    "Tar_Science": "NA",
    "Tar_Households": "yes",
    "Tar_Edu_institutions": "yes",
    "Tar_Private_companies": "yes",

    "Actor_role": "1,2",
    "Discretion": "DoE has discretion in which industries it assesses/monitors "
                 "(currently only formally established/registered plastic "
                 "industries and markets - informal producers fall outside its "
                 "monitoring) and is designing a new subsidy scheme for the next "
                 "fiscal year to help industries shift to compliant (>40 micron) "
                 "plastic production.",

    "Implementation_issues": "yes",
    "Monitoring": "yes",
    "Financial_resources": "NA",
    "Research": "yes",
    "Infrastructure": "NA",
    "Capacity": "yes",
    "Enforcement": "yes",

    "Policies_in_place": "yes",
    "Pol_epr": "no",
    "Pol_import": "yes",
    "Pol_awarness": "yes",
    "Pol_education": "yes",
    "Pol_capacity": "NA",
    "Pol_RD": "NA",
    "Pol_tax": "NA",
    "Pol_ban": "yes",
    "Pol_subsitutes": "yes",
    "Pol_clean_up": "NA",
    "Pol_upcycling": "NA",
    "Pol_recycling": "yes",
    "Pol_waste_collection": "yes",
    "Pol_effectiveness": "yes",
    "Pol_effectiveness_expample": "\"There is a production of biodegradable plastics, "
                                  "but price is high... due to this price gap the market "
                                  "also uses conventional plastic\"; \"lack of a strong "
                                  "surveillance mechanism\"; \"there is inadequate "
                                  "monitoring due to lack of few people. 2-3 times monitor "
                                  "in a year\"; \"no penalty in place\" - together showing "
                                  "the 40-micron bag ban and plastic-flower ban are weakly "
                                  "enforced.",

    "Solutions": "yes",
    "Sol_lead_agency": "NA",
    "Sol_responsibilities": "NA",
    "Sol_epr": "yes",
    "Sol_awarness": "yes",
    "Sol_segregation": "yes",
    "Sol_upcycling": "NA",
    "Sol_recycling": "yes",
    "Sol_education": "yes",
    "Sol_capacity": "yes",
    "Sol_RD": "yes",
    "Sol_tax": "NA",
    "Sol_ban": "NA",
    "Sol_finance": "yes",
    "Sol_infrastructure": "yes",
    "Sol_subsitutes": "yes",
    "Sol_clean_up": "NA",
    "Sol_enforcement": "yes",
    "Sol_monitoring": "yes",

    "Traditions_to_build_on": "Bhaktapur municipality was cited as a positive local "
                              "example (\"eg. Of Bhaktapur\") of waste management "
                              "practice that could be built on/replicated elsewhere; "
                              "no other traditional practices were mentioned.",
}

INTERVIEWS = [rsct, doe]

# ---------------------------------------------------------------------------
# 3. EXPLANATIONS (long format) - quote / reasoning behind each coded value
#    Only variables with a substantive (non-NA) value are documented here.
# ---------------------------------------------------------------------------

rsct_expl = {
    "Actortype": "RSCT is described as a savings-and-credit cooperative network "
                 "(\"Based on that cooperative; two tier cooperative; More than 200 "
                 "cooperative\") originally focused on poverty alleviation, now also "
                 "working on urban plastics - classified as civil society/cooperative "
                 "sector rather than government or private business.",
    "Problem_awareness_pop": "\"In general they know; [but] they don't know the extent "
                             "of it\" and \"Public is ignorant\" - population has only "
                             "partial awareness of the plastic pollution problem.",
    "Problem_concerndness": "The interviewee acknowledges the problem exists and is "
                            "generally known (\"In general they know\") but stresses "
                            "the extent is under-appreciated, and no strong emotional or "
                            "urgent language is used - rated as medium concern.",
    "Problem_littering": "\"Blocking sewer systems\" implies discarded/littered plastic "
                         "accumulating in drainage systems.",
    "Problem_waste_mgmt": "\"Land-fills are already [full]\"; \"Blocking sewer systems\"; "
                          "\"It's burned - air pollution\" collectively describe inadequate "
                          "waste management.",
    "Problem_production": "\"Lots of plastics are produced\" - explicit mention of high "
                          "upstream production volumes.",
    "Problem_alternatives": "\"R&D - alternatives are needed\" implies current alternatives "
                            "are insufficient.",
    "Problem_impacts": "Derived from: landfills (dumping sites), blocked sewers (drainage "
                       "blockage), burning of plastic causing air pollution, and indoor "
                       "health effects of burning (\"can impact health by burning; indoor\").",
    "Governance": "Extensive discussion of stakeholder coordination, fund release "
                 "processes, and policy implementation gaps constitutes governance-related "
                 "content.",
    "Coordination_sectoral": "\"Stakeholders are not collaborating; they do not coordinate\" "
                             "- explicit statement of poor cross-actor collaboration.",
    "Coordination_levels": "Same quote (\"they do not coordinate\") combined with \"so it "
                           "stops at major[levels]\" suggests breakdowns between "
                           "government levels, not only sectors.",
    "Res_Loc_government": "The described funding/facility system (earmarked funds, "
                          "facilities) is administered at a sub-national level, and \"it "
                          "stops at major\" implies responsibility sits with government "
                          "below the national tier.",
    "Res_Private_sector": "The 300 rupee/month incentive scheme is paid to \"waste "
                          "collectors\", who are private-sector actors central to "
                          "implementation.",
    "Res_Civil_society": "RSCT itself, a cooperative/NGO-type organisation, has taken on "
                         "plastics work in urban areas, i.e. it sees itself as (partly) "
                         "responsible.",
    "Cul_Nat_government": "\"Fund is there; they don't release the funds\" implicates a "
                          "governing authority (fund holder) in stalling implementation.",
    "Cul_Households": "\"Public is ignorant\" attributes part of the problem to household/ "
                      "population behaviour and lack of awareness.",
    "Tar_Loc_government": "Local-level bodies administer the waste-collector incentive "
                          "scheme and facilities described, making them a natural target "
                          "for future measures.",
    "Tar_Private_sector": "Waste collectors are directly targeted by the existing 300 "
                          "rupee/month incentive and would remain a target of future policy.",
    "Tar_Households": "The 300 rupee/month scheme is conditional on \"the household reduce "
                      "[waste]\", explicitly targeting households.",
    "Actor_role": "RSCT coordinates a network of 200+ member cooperatives (organisational/ "
                 "managerial role = 2) while also directly running grassroots urban "
                 "plastics activities (street-level role = 3).",
    "Discretion": "RSCT's cooperative model was self-designed bottom-up (\"Bottom-up and "
                 "grassroot fund to give loan\") and it independently chose to redirect "
                 "this model toward plastics work, indicating discretion.",
    "Implementation_issues": "Multiple concrete implementation problems are raised "
                             "(monitoring, funds, facilities, political will), indicating "
                             "this theme was substantively discussed.",
    "Monitoring": "\"Implementation is always the problem: monitoring\" - monitoring is "
                 "explicitly named as a recurring implementation failure.",
    "Financial_resources": "\"Fund is there; they don't release the funds\"; \"Two funds: "
                           "ear-marked; free to use (if you don't use it will be frozen)\" "
                           "- funds exist on paper but are not effectively available/used.",
    "Research": "\"Not much research on effects of plastics\" and the closing remark "
               "\"More research is needed\" both point to a research gap.",
    "Infrastructure": "\"Facility is there (at least in this [case])\" - explicitly states "
                      "infrastructure/facilities are available, i.e. not lacking.",
    "Capacity": "\"Lack of how to act\" describes implementers/communities not knowing how "
               "to act - a capacity gap.",
    "Policies_in_place": "\"Policies are there\" is stated explicitly, even though "
                         "implementation is weak.",
    "Pol_waste_collection": "\"300 Rupees per month for waste collectors, if the household "
                            "reduce[s]\" describes an active local waste-collection/ "
                            "incentive scheme.",
    "Pol_effectiveness": "\"Implementation is always the problem\" combined with the "
                         "earmarked-funds example shows existing policies are not "
                         "translating into effective action.",
    "Sol_awarness": "\"There is a need of communication\" is offered as a way forward, "
                    "i.e. more awareness/communication work.",
    "Sol_capacity": "The same \"lack of how to act\" problem is paired with an implicit "
                    "call to build the capacity/know-how to act.",
    "Sol_RD": "\"R&D - alternatives are needed\" is explicitly proposed as a way forward.",
    "Sol_finance": "The funds/disbursement problem identified above implies a needed "
                  "solution of ensuring earmarked funds are actually released and used.",
    "Sol_subsitutes": "Tied to the R&D need for \"alternatives\", i.e. viable substitute "
                      "products are called for.",
    "Traditions_to_build_on": "RSCT's own background (\"Established in 1991\"; \"Bottom-up "
                              "and grassroot fund to give loan\"; \"More than 200 "
                              "cooperative[s]\"; \"Governed by general assembly\") "
                              "describes a long-standing, locally rooted cooperative "
                              "financing tradition that has already been repurposed for "
                              "plastics work and could be scaled further.",
}

doe_expl = {
    "Actortype": "Deepak Diwal(i) is Deputy Director for Pollution Control (air and "
                "plastics) at the Department of Environment, a national government body.",
    "Problem_awareness_pop": "The interviewee states people need more awareness \"on "
                             "impacts of plastic on human health\" and that \"behavioural "
                             "change - mindset change\" is still needed, implying "
                             "currently insufficient public awareness.",
    "Problem_concerndness": "The DoE describes a wide-ranging institutional response "
                            "(policy instruments, a high-level committee chaired by the "
                            "Chief Secretary, planned subsidies, monitoring) reflecting a "
                            "high level of institutional concern about the issue.",
    "Problem_recycling": "\"Single-use plastics... are not economically viable for "
                         "recycling\" identifies a structural limitation in current "
                         "recycling capacity/viability.",
    "Problem_waste_mgmt": "\"It is lacking to manage the waste, including the plastic "
                          "waste\" and \"high level of leak[age] of plastic in our "
                          "different river systems\" directly state inadequate waste "
                          "management.",
    "Problem_alternatives": "\"The option is quite expensive\" and the ~100 rupee/kg price "
                            "gap between conventional and compostable bags show viable "
                            "alternatives are limited/unaffordable.",
    "Problem_waste_segregation": "The solutions raised later (\"systematic circulation of "
                                 "plastics waste (circular economy), segregation\"; "
                                 "\"Collection system - segregation of plastics\") imply "
                                 "that segregation is currently insufficient.",
    "Problem_impacts": "\"In our river system, in our lake system, in our land, forest, "
                       "everywhere our plastic pollution exists\"; \"high level of leak[age] "
                       "of plastic in our different river systems\"; microplastic \"residue... "
                       "in our snow\" in the Himalayan range.",
    "NPF_victims": "Rivers, lakes, land and forests are described as polluted, and "
                  "microplastics were found even \"in our snow\" in the Himalayan range - "
                  "framing the natural environment (and, by extension, the people relying "
                  "on it) as the victims.",
    "NPF_villains": "\"There is no policy [producer] responsibility in Nepal\" (extended "
                    "producer responsibility absent) points to plastic producers; "
                    "\"informal without govt notice these are the problems\" points to "
                    "unregistered industries; and \"it is our habits and needs\" plus the "
                    "price-sensitivity of consumers points to households/the public.",
    "NPF_hero": "The Department of Environment presents itself and the high-level "
               "committee (chaired by the Chief Secretary) as driving bans, monitoring and "
               "planned subsidies, while crediting local governments and NGOs for running "
               "material recovery facilities and awareness campaigns.",
    "Governance": "Detailed description of the policy/committee structure (directives, "
                 "gazette publications, a high-level committee, DoE monitoring, local "
                 "government and Department of Commerce & Supplies involvement) shows "
                 "governance was substantively discussed.",
    "Coordination_sectoral": "The interviewee is unsure of DoE's own boundary versus the "
                             "Department of Urban Development on the Solid Waste "
                             "Management Act (\"these things look after our urban "
                             "development\"; \"DoE role in plastic or waste mgmt.?? "
                             "Department of Urban Development??\"), indicating unclear "
                             "cross-sector coordination.",
    "Coordination_levels": "\"Whose responsibility is it to manage the landfills? "
                           "Municipals. Do you monitor this as well? - No.\" shows a gap "
                           "between what national (DoE) and local government each monitor.",
    "Unclear_responsibilities": "The interviewee explicitly could not clearly state "
                                "whether solid-waste-act ownership sits with DoE or the "
                                "Department of Urban Development (\"I also heard that the "
                                "solid waste management act is being revised... actually "
                                "these things look after our urban development\").",
    "Res_Nat_government": "DoE issues directives/gazette notices, chairs (via the Chief "
                          "Secretary) the high-level committee, and regularly monitors "
                          "plastic-producing industries and markets.",
    "Res_Loc_government": "\"Local government[s] lead\" waste management; \"our local "
                          "government is the governing body for waste management.\"",
    "Res_Private_sector": "\"Private companies... participate in recycling\"; the "
                          "recycling industry and private waste-collection contractors "
                          "(via tenders) are named as active responsible actors.",
    "Res_Civil_society": "\"NGOs also operate recovery facilities and run awareness "
                         "campaigns\"; \"Some NGOs are involved in conducting the material "
                         "recovery facility in Kathmandu... and getting awareness.\"",
    "Res_Edu_institutions": "\"In education kids are taught about it\" indicates schools "
                            "already play some role in addressing the issue.",
    "Cul_Private_sector": "\"There is no [extended] producer responsibility in Nepal\" and "
                          "continued production/sale of banned items due to weak "
                          "enforcement implicate plastic-producing/selling businesses.",
    "Cul_Households": "\"It is our habits and needs\" and price-sensitivity keeping demand "
                      "for cheap conventional bags high attribute part of the problem to "
                      "household/consumer behaviour.",
    "Cul_Private_companies": "\"Informal without govt notice these are the problems\" - "
                             "unregistered/informal industries selling non-compliant "
                             "plastics outside of regulatory reach.",
    "Tar_Nat_government": "Q9 answer: \"all stakeholders. Central govt and LG.\"",
    "Tar_Loc_government": "Same Q9 answer explicitly includes local government as a focus "
                          "of future policy.",
    "Tar_Private_sector": "\"Waste producers\" and \"plastic producers - how to manage "
                          "better education\" are named as groups future policy should "
                          "target.",
    "Tar_Households": "Q10 answer: support is needed \"at every household level\".",
    "Tar_Edu_institutions": "Q10 answer explicitly lists \"schools\" alongside households "
                            "as a focus for support/education measures.",
    "Tar_Private_companies": "\"Plastic producers... they need to invest as well\" frames "
                             "producing companies as a target group that must change "
                             "behaviour/investment.",
    "Actor_role": "As Deputy Director at DoE, the interviewee both drafts/oversees policy "
                 "instruments (gazette notices, directives - role 1) and personally "
                 "describes DoE's operational monitoring of industries and markets "
                 "(role 2).",
    "Discretion": "DoE \"only assess[es] formally established industries\" (informal ones "
                 "are excluded from its monitoring scope) and is independently developing "
                 "a subsidy scheme (\"next fiscal year... we will provide grant... to "
                 "shift the industry\") - both reflect discretionary decisions by DoE.",
    "Implementation_issues": "Multiple concrete barriers are discussed (weak surveillance, "
                             "understaffed monitoring, no penalties, expensive "
                             "alternatives), indicating substantive coverage of this theme.",
    "Monitoring": "\"There is inadequate monitoring due to lack of few people. 2-3 times "
                 "[we] monitor in a year\"; assessments only cover \"formally established "
                 "industries\", leaving informal ones unchecked.",
    "Research": "\"Role of DoE - issues to monitor, how to strengthen it? - types of "
               "plastics to be known\" and the call for \"a specific plastics guideline\" "
               "point to a knowledge/research gap.",
    "Capacity": "Inadequate monitoring is attributed to \"lack of [a] few people\" - a "
               "staffing/capacity constraint.",
    "Enforcement": "\"Due to the lack of a strong surveillance mechanism\" products \"are "
                  "sold even without a ban\"; \"no penalty in place\" for non-compliance.",
    "Policies_in_place": "A detailed set of current instruments is described: directives, "
                         "import/production/use restrictions, and a high-level committee.",
    "Pol_epr": "\"There is no [extended] producer responsibility in Nepal\" - explicitly "
              "stated as not currently in place.",
    "Pol_import": "\"Restriction of import and use of plastic less than 40 microns\" is an "
                 "active import-related regulation.",
    "Pol_awarness": "\"We also launched... different types of awareness campaigns\", "
                    "documentaries and social/traditional media use.",
    "Pol_education": "\"In education kids are taught about it\" indicates an existing "
                     "school-based education element.",
    "Pol_ban": "\"Ban on production, import, and use of plastic bags under 40 microns\" "
              "and \"restrictions on plastic decorative flowers\".",
    "Pol_subsitutes": "Biodegradable/starch-based bags already exist on the market as "
                      "substitutes, and a subsidy to support the shift to compliant "
                      "plastic is being planned.",
    "Pol_recycling": "\"Private companies and NGOs participate in recycling\"; material "
                     "recovery facilities feed recyclables \"to processing/recycling "
                     "units\".",
    "Pol_waste_collection": "\"Waste collection is done either by local governments "
                            "directly or via private contractors\" (tender system).",
    "Pol_effectiveness": "Weak surveillance, only 2-3 monitoring visits/year, no "
                         "penalties, and the persistent price gap for alternatives all "
                         "point to limited effectiveness of the existing ban/import "
                         "restriction.",
    "Sol_epr": "\"Extended producer responsibility is key - there is no producer "
              "responsibility in Nepal, while in other countries there is\" is offered "
              "directly as a needed solution.",
    "Sol_awarness": "Q10: \"awareness campaigns, advocacy through documentary\" listed as "
                    "helpful support going forward.",
    "Sol_segregation": "\"Systematic circulation of plastics waste (circular economy), "
                       "segregation\"; \"Collection system - segregation of plastics. "
                       "Separate collection and managed.\"",
    "Sol_recycling": "Circular-economy framing and \"manure out of biodegradable waste\" "
                     "point to expanded recycling/resource recovery as a way forward.",
    "Sol_education": "Q10: support needed \"at every household level., schools, "
                     "monitoring of public places\" and better education for plastic "
                     "producers.",
    "Sol_capacity": "\"Plastic producers - how to manage better education. Circular "
                    "economy - make sustainability. They need to invest as well\" implies "
                    "capacity-building among producers.",
    "Sol_RD": "\"Need for a specific plastics guideline for its management\" and the need "
             "to have \"types of plastics to be known\" both call for further "
             "research/knowledge development.",
    "Sol_finance": "\"Need for incentives - alternatives\"; the planned fiscal-year grant "
                  "programme to help industries shift to compliant plastic.",
    "Sol_infrastructure": "\"Make transport department for wastes along roads. Need [for] "
                          "continuity. Sustainability.\" calls for infrastructural/ "
                          "logistics investment in collection.",
    "Sol_subsitutes": "\"Need for incentives - alternatives\" - cheaper/accessible "
                      "substitute products are called for.",
    "Sol_enforcement": "\"No penalty in place\" is implicitly paired with the need to "
                       "introduce enforcement/penalty mechanisms as part of strengthening "
                       "policy.",
    "Sol_monitoring": "\"Role of DoE - issues to monitor how to strengthen it?\" and Q10's "
                      "\"monitoring of public places\" call for strengthened monitoring "
                      "going forward.",
    "Traditions_to_build_on": "\"Eg. Of Bhaktapur\" is raised as a positive existing local "
                              "example under the solutions discussion, suggesting an "
                              "existing municipal practice that could be replicated "
                              "elsewhere.",
}

EXPLANATIONS = {"NPL_1": rsct_expl, "NPL_5": doe_expl}
NAMES = {"NPL_1": "RSCT (cooperative network, urban plastics programme)",
         "NPL_5": "Department of Environment - Deepak Diwal(i), Deputy Director, "
                  "Pollution Control (air & plastics) - Interview No. 5, 23 June 2025"}

# ---------------------------------------------------------------------------
# 4. BUILD WORKBOOK
# ---------------------------------------------------------------------------

wb = openpyxl.Workbook()

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF")
SECTION_FONT = Font(bold=True, italic=True)
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


# --- Sheet 1: Codebook ---
ws1 = wb.active
ws1.title = "Codebook"
ws1.append(["Variable name", "Description", "Example"])
style_header_row(ws1, 1, 3)
for var, desc, example in CODEBOOK:
    ws1.append([var, desc, example])
    r = ws1.max_row
    if var.startswith("---"):
        for c in range(1, 4):
            ws1.cell(row=r, column=c).fill = SECTION_FILL
            ws1.cell(row=r, column=c).font = SECTION_FONT
    for c in range(1, 4):
        ws1.cell(row=r, column=c).alignment = WRAP
        ws1.cell(row=r, column=c).border = BORDER
ws1.column_dimensions["A"].width = 34
ws1.column_dimensions["B"].width = 95
ws1.column_dimensions["C"].width = 20
ws1.freeze_panes = "A2"

# --- Sheet 2: Coded_Data (wide format) ---
ws2 = wb.create_sheet("Coded_Data")
ws2.append(DATA_VARS)
style_header_row(ws2, 1, len(DATA_VARS))
for interview in INTERVIEWS:
    ws2.append([interview.get(v, "NA") for v in DATA_VARS])
    r = ws2.max_row
    for c in range(1, len(DATA_VARS) + 1):
        ws2.cell(row=r, column=c).alignment = WRAP
        ws2.cell(row=r, column=c).border = BORDER
for i in range(1, len(DATA_VARS) + 1):
    ws2.column_dimensions[get_column_letter(i)].width = 22
ws2.column_dimensions["A"].width = 12
ws2.column_dimensions["B"].width = 10
ws2.freeze_panes = "C2"

# --- Sheet 3: Coding_Explanations (long format) ---
ws3 = wb.create_sheet("Coding_Explanations")
ws3.append(["ID", "Interview / Source", "Variable", "Coded value", "Evidence / Explanation"])
style_header_row(ws3, 1, 5)
desc_lookup = {row[0]: row[1] for row in CODEBOOK}
for interview in INTERVIEWS:
    iid = interview["ID"]
    expl = EXPLANATIONS[iid]
    for var in DATA_VARS:
        if var in expl:
            value = interview.get(var, "NA")
            ws3.append([iid, NAMES[iid], var, value, expl[var]])
            r = ws3.max_row
            for c in range(1, 6):
                ws3.cell(row=r, column=c).alignment = WRAP
                ws3.cell(row=r, column=c).border = BORDER
ws3.column_dimensions["A"].width = 10
ws3.column_dimensions["B"].width = 30
ws3.column_dimensions["C"].width = 26
ws3.column_dimensions["D"].width = 14
ws3.column_dimensions["E"].width = 100
ws3.freeze_panes = "A2"

# --- Sheet 4: Read_Me ---
ws4 = wb.create_sheet("Read_Me", 0)
readme_lines = [
    ("Nepal Plastic Pollution Governance - Interview Coding", True),
    ("", False),
    ("This workbook applies the supplied codebook (see 'Codebook' sheet) to two "
     "interview sources from the 'Plastic Pollution Governance in Nepal' project:", False),
    ("  1. NPL_1 - RSCT: a savings-and-credit cooperative network (est. 1991) that "
     "recently expanded into urban plastics work. Coded from the hand-written "
     "interview notes.", False),
    ("  2. NPL_5 - Department of Environment (Deepak Diwal(i), Deputy Director, "
     "Pollution Control - air & plastics), Interview No. 5, 23 June 2025. Coded "
     "from both the interview notes and the full verbatim transcript.", False),
    ("", False),
    ("Sheets in this workbook:", True),
    ("  - Codebook: the variable dictionary as supplied, used as the coding frame.", False),
    ("  - Coded_Data: wide-format matrix (one row per interview, one column per "
     "codebook variable) with the coded values.", False),
    ("  - Coding_Explanations: long-format table giving the quote/observation used "
     "to justify every non-NA code in Coded_Data (one row per Interview x Variable).", False),
    ("", False),
    ("Coding conventions:", True),
    ("  - 'yes' / 'no' / 'NA': NA means the topic was not addressed in that "
     "interview (no evidence either way). 'no' is only used where the interviewee "
     "explicitly said the item is not an issue / not in place (e.g. Pol_epr = no "
     "for both interviews, since both explicitly state extended producer "
     "responsibility does not exist in Nepal).", False),
    ("  - Free-hand fields (Problem_impacts, NPF_victims/villains/hero, Discretion, "
     "Pol_effectiveness_expample, Traditions_to_build_on) contain short descriptive "
     "text derived directly from the interview content, or 'NA' if not discussed.", False),
    ("  - Actor_role: 1 = formulation/policy actor, 2 = managerial/organisational "
     "implementer, 3 = street-level implementer, 4 = target group (comma-separated "
     "if more than one applies), derived from what the interview describes the "
     "interviewee's organisation as actually doing - not simply from their actor "
     "type.", False),
    ("", False),
    ("Note on source material: a block of longer, polished paragraph-style quotes "
     "appeared under the RSCT question list in the original material, but the "
     "content and first-person phrasing (\"we have... directives...\", "
     "\"Department of Environment is regularly monitoring...\") match the "
     "Department of Environment notes and transcript almost verbatim. These "
     "paragraphs were therefore treated as additional corroborating evidence for "
     "the DoE interview (NPL_5), not for RSCT (NPL_1), and coded accordingly.", False),
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
