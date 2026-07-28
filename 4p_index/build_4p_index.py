#!/usr/bin/env python3
"""
Builds the 4P Index (Plastic Pollution Policy Index) v2 coding table
for Nepal's national plastic-pollution-relevant policy landscape.

Each entry in POLICIES contains the policy-level fields (A-N) plus a list of
instrument-level dicts (P-W). Columns O (policy_score) and V (instrument_score)
are computed automatically from the formulas defined in the coding prompt:
    O = (G + I + K + M) / 4
    V = S * (P + T) / 2
"""
import csv

COLUMNS = [
    "policy_name", "policy_url", "policy_year", "policy_objective",
    "policy_target", "policy_target_text", "policy_type", "policy_type_justification",
    "policy_integration", "policy_sectors_list", "policy_circularity",
    "policy_lifecycle_phases_list", "policy_budget", "policy_budget_text", "policy_score",
    "instrument_type", "instrument_lifecycle_stage", "instrument_description",
    "instrument_in_force", "instrument_implementation", "instrument_implementation_text",
    "instrument_score", "comments",
]

POLICIES = [
    dict(
        policy_name="Environment Protection Act, 2076 (2019)",
        policy_url="http://climate.mohp.gov.np/images/Environment-Protection-Act-2019-English.pdf",
        policy_year=2019,
        policy_objective=(
            "To protect and promote a clean and healthy environment in Nepal by regulating "
            "pollution (including from industry, hotels/restaurants, goods, activities and "
            "waste/hazardous-substance disposal), requiring environmental studies for proposals, "
            "and establishing the institutions, standard-setting powers, monitoring duties and "
            "penalties that underpin all subordinate plastic-waste regulation (e.g., the Plastic "
            "Bag directives)."
        ),
        policy_target=0,
        policy_target_text="",
        policy_type=1,
        policy_type_justification=(
            "Act (\"Ain\") enacted by the Federal Parliament of Nepal, repealing and replacing the "
            "Environment Protection Act, 1997 (s.47). Primary legislation, not an executive decree."
        ),
        policy_integration=0.75,
        policy_sectors_list="waste management, industry, tourism, transport, hotels & restaurants, chemicals",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="production, consumption, disposal, environmental leakage",
        policy_budget=0.5,
        policy_budget_text=(
            "Ministerial/Council function \"to manage economic resources for environmental "
            "protection and climate change and facilitate for the same\" (functions of the Council "
            "chaired by the Prime Minister, Chapter on institutional arrangements); no dedicated, "
            "ring-fenced plastics/waste fund is created by the Act itself."
        ),
        instruments=[
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "S.15 \"Control of pollution\": Government of Nepal may, by Gazette notification, "
                    "determine standards to mitigate pollution from industrial enterprises, hotels, "
                    "restaurants, goods, activities and disposal/emission of hazardous substances; "
                    "s.15(2) \"No person shall create pollution... or do... any act contrary to the "
                    "standards\"; s.35 sets fines for non-compliance; s.41 empowers Gazette orders on "
                    "precautionary matters."
                ),
                instrument_in_force=1,
                instrument_implementation=0.75,
                instrument_implementation_text=(
                    "Authority: Department of Environment/Ministry designated (+0.25). Enforcement: "
                    "s.35 fines up to NPR 5,00,000-50,00,000 and s.35(3) fines up to NPR 3,00,000 for "
                    "other contraventions (+0.25). Monitoring: s.39 mandatory monitoring/inspection duty "
                    "(+0.25). Unconditional: standards are only \"as prescribed\" by subordinate notice, "
                    "leaving discretionary scope/exemptions (0)."
                ),
                comments=(
                    "This is the parent enabling provision under which plastic-specific bans (e.g., the "
                    "Plastic Bag (Regulation and Control) Directive) are issued via s.45 - see the "
                    "separate Plastic Bag Directive policy rows for the plastics-specific ban."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Environmental leakage",
                instrument_description=(
                    "Ss.21-22: environmental inspectors \"shall carry out monitoring and inspection\" of "
                    "compliance with pollution-control standards; s.39: Ministry/Department/Provincial/"
                    "Local Level \"may carry out monitoring and inspection\" of Act implementation."
                ),
                instrument_in_force=1,
                instrument_implementation=0.75,
                instrument_implementation_text=(
                    "Authority: environmental inspectors formally designated (+0.25). Monitoring: this "
                    "is itself the monitoring/inspection mechanism, explicitly mandatory for s.39(1) "
                    "(+0.25). Enforcement: inspection findings feed into s.35 fines (+0.25). "
                    "Unconditional: not evidenced (0)."
                ),
                comments="Institutional/monitoring instrument rather than a substantive plastics rule.",
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Council (chaired by the Prime Minister) function \"to give necessary policy "
                    "guidance to the Provincial and Local levels with regard to environmental "
                    "protection\"; s.26 empowers coordinated action by Government of Nepal, Provincial "
                    "Government and Local Level on environmental/climate risk mitigation."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: Council/Ministry roles for federal-provincial-local guidance are "
                    "explicitly defined in the Act (+0.25). No explicit monitoring or enforcement "
                    "mechanism or unconditional guarantee evidenced for this coordination function "
                    "specifically."
                ),
                comments=(
                    "Coded per Rule 11 as a multi-level governance coordination instrument "
                    "(national-provincial-local policy guidance on environmental protection)."
                ),
            ),
        ],
    ),
    dict(
        policy_name="Plastic Bag (Regulation and Control) Directive, 2082 (2025)",
        policy_url="https://thehimalayantimes.com/environment/nepal-issues-new-plastic-bag-directive-bans-bags-thinner-than-40-microns",
        policy_year=2025,
        policy_objective=(
            "To prohibit the production, import, storage, sale, distribution and use of plastic bags "
            "thinner than 40 microns throughout Nepal; to set size, material and labelling standards "
            "for permitted bags; to impose Extended Producer Responsibility (EPR) obligations on "
            "producers/importers; and to assign compliance-monitoring roles across federal, "
            "provincial and local government."
        ),
        policy_target=1,
        policy_target_text=(
            "\"The directive prohibits the production, import, storage, sale, and use of plastic "
            "bags thinner than 40 microns. Only bags measuring at least 7 inches by 14 inches or "
            "larger may be produced, imported, and sold as general plastic bags, while garbage bags "
            "must be at least 14 inches by 26 inches or larger.\""
        ),
        policy_type=0.75,
        policy_type_justification=(
            "Directive (\"Nirdeshika\") issued by the Ministry of Forests and Environment under the "
            "authority of s.45 of the Environment Protection Act, 2076 - an executive/ministerial, "
            "sub-legislative instrument, not enacted by Parliament. It replaces the earlier Plastic "
            "Bag (Regulation and Control) Directive, 2068 (2011) and its 2078 (2021) revision."
        ),
        policy_integration=1,
        policy_sectors_list="production, consumption, retail, waste management, recycling, industry, trade",
        policy_circularity=1,
        policy_lifecycle_phases_list="production, consumption, recycling, disposal, environmental leakage",
        policy_budget=0.5,
        policy_budget_text=(
            "EPR clause requires producers/importers to \"collect and manage at least [a specified "
            "share of] the waste they introduce\" and submit annual reports - a producer-funded "
            "financing mechanism, but no ring-fenced government fund is created by the Directive "
            "itself."
        ),
        instruments=[
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Production",
                instrument_description=(
                    "Complete ban on production, import, storage, sale and distribution of plastic "
                    "bags thinner than 40 microns nationwide; mandatory minimum bag sizes and "
                    "virgin/recycled-granule composition standards; mandatory labelling (producer "
                    "name/address, thickness, recycling symbol)."
                ),
                instrument_in_force=1,
                instrument_implementation=0.75,
                instrument_implementation_text=(
                    "Authority: Department of Environment designated as central regulator, with "
                    "provincial/local governments empowered to monitor markets (+0.25). Enforcement: "
                    "violations actionable under Environment Protection Act s.35(3) (+0.25). "
                    "Monitoring: importers must \"maintain calibrated instruments\" and \"submit annual "
                    "reports\" on bags sold/imported (+0.25). Unconditional: not met - bags \u226540 "
                    "microns remain permitted and prior enforcement attempts (2011, 2015, 2018, 2021) "
                    "were repeatedly reported as unimplemented, indicating de-facto loopholes (0)."
                ),
                comments=(
                    "Replaces/supersedes the Action Plan for Ban on Plastic Bags, 2079 (2022) and the "
                    "2078 (2021) Gazette notice banning bags <40 microns, which pursued the same "
                    "objective via a Cabinet-approved action plan rather than a ministerial directive; "
                    "media sources note repeated past non-enforcement (e.g., Kathmandu Post, 2021)."
                ),
            ),
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="End of life",
                instrument_description=(
                    "Extended Producer Responsibility (EPR) obligation: manufacturers and importers "
                    "\"must collect and manage\" a minimum share of the plastic-bag waste they place on "
                    "the market and submit annual compliance reports to the Department."
                ),
                instrument_in_force=1,
                instrument_implementation=0.75,
                instrument_implementation_text=(
                    "Authority: Department of Environment (+0.25). Enforcement: penalties under the "
                    "Environment Protection Act apply to non-compliant producers (+0.25). Monitoring: "
                    "mandatory annual EPR reporting (+0.25). Unconditional: not evidenced (0)."
                ),
                comments="First binding, plastics-specific EPR obligation identified in Nepali policy to date.",
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Multi-level enforcement architecture: the Department of Environment regulates "
                    "compliance centrally, while \"provincial and local governments can use the "
                    "guidelines as a framework for enforcement\" and \"local bodies are empowered to "
                    "monitor markets and enforce standards within their jurisdictions.\""
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: roles for Department of Environment, provincial and local governments "
                    "explicitly assigned (+0.25). Monitoring: local-level market monitoring role "
                    "explicit (+0.25). No separate/independent enforcement or unconditional guarantee "
                    "evidenced beyond what is already counted for the ban instrument."
                ),
                comments=(
                    "Coded per Rule 11 as a distinct multi-level governance coordination instrument "
                    "(national Department of Environment - provincial - local monitoring roles), kept "
                    "separate from the ban itself to highlight the coordination dimension."
                ),
            ),
        ],
    ),
    dict(
        policy_name=(
            "Notice on Prohibition of Production, Import, Sale, Distribution and Storage of "
            "Plastic Artificial Flowers (Nepal Gazette, 2079/2022)"
        ),
        policy_url="https://research.hktdc.com/en/article/MTE0MjMzNDg4Mw",
        policy_year=2022,
        policy_objective=(
            "To completely prohibit the production, import, sale, distribution and storage of "
            "plastic artificial flowers, garlands and bouquets in Nepal, as an upstream measure to "
            "reduce non-biodegradable, non-recyclable plastic products entering the market."
        ),
        policy_target=0.5,
        policy_target_text=(
            "\"...the government imposed complete ban on production, import, sales, distribution and "
            "storage of plastic flower.\" (paraphrased in secondary reporting; exact Gazette wording "
            "not independently verified)."
        ),
        policy_type=0.75,
        policy_type_justification=(
            "Notification published in the Nepal Gazette by the Ministry of Forests and Environment "
            "under its Environment Protection Act powers - an executive Gazette order, not an Act of "
            "Parliament."
        ),
        policy_integration=0.5,
        policy_sectors_list="production, trade, retail, consumption",
        policy_circularity=0.5,
        policy_lifecycle_phases_list="production, consumption",
        policy_budget=0,
        policy_budget_text="",
        instruments=[
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Production",
                instrument_description=(
                    "Complete ban on production, import, sale, distribution and storage of plastic "
                    "artificial flowers, garlands and bouquets; forecast to halt an estimated NPR 10 "
                    "crore/year of plastic-decoration imports."
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: Ministry of Forests and Environment/Customs administer the import ban "
                    "(+0.25). Unconditional: framed as a blanket, complete ban on the product category "
                    "with no stated carve-outs (+0.25). Enforcement and monitoring: not evidenced in "
                    "the sources reviewed - press reporting (Nepal News, 2022) instead notes 15 tonnes "
                    "of plastic flowers were imported in the first month after the ban, suggesting weak "
                    "practical enforcement (0)."
                ),
                comments=(
                    "Uncertain (Rule 8): coded policy_target as 0.5 rather than 1 because the "
                    "underlying measure is framed as a categorical ban rather than a numeric "
                    "percentage/threshold target; a case could be made for scoring a 100% ban as a "
                    "quantifiable target (=1) instead."
                ),
            ),
        ],
    ),
    dict(
        policy_name="Solid Waste Management National Policy, 2079 (2022)",
        policy_url="https://mofe.gov.np/storage/files/National_Waste_Management_Policy_2079.pdf",
        policy_year=2022,
        policy_objective=(
            "To provide the national policy framework guiding solid waste management (which by "
            "definition includes plastic waste) across the federal, provincial and local tiers of "
            "government; to promote waste as a resource through the circular economy and 3R "
            "(reduce, reuse, recycle); to reduce environmental pollution and public-health risk; and "
            "to encourage private-sector participation and public-private partnership in waste "
            "infrastructure."
        ),
        policy_target=1,
        policy_target_text=(
            "\"The Waste Management National Policy 2079 aims to address the growing waste crisis... "
            "It targets zero-waste-to-landfill by 2030 through enhanced recycling and waste-to-energy "
            "initiatives.\""
        ),
        policy_type=0.5,
        policy_type_justification=(
            "A national policy (\"rashtriya niti\") approved at Council of Ministers (Cabinet) level - "
            "a strategic policy document rather than an Act of Parliament or a stand-alone "
            "regulation/ordinance - that nonetheless embeds a quantifiable target (zero waste to "
            "landfill by 2030), placing it at 0.50 rather than 0.25."
        ),
        policy_integration=1,
        policy_sectors_list="waste management, industry, municipalities, private sector, tourism, agriculture, chemicals, health",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="consumption, recycling, disposal, environmental leakage",
        policy_budget=0.5,
        policy_budget_text=(
            "Policy strategy of \"increasing public private partnership and investment in waste "
            "management\"; in practice, local bodies draw on federal/provincial conditional, "
            "complementary and special intergovernmental grants (via the National Natural Resources "
            "and Fiscal Commission and Ministry of Finance) for SWM projects - a funding source is "
            "identified but no single ring-fenced fund is created by this policy itself."
        ),
        instruments=[
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Multi-tier institutional role allocation: Federal level - development of acts, "
                    "policies, regulations, research/technology development and attracting foreign "
                    "investment; Provincial level - development of infrastructure (e.g., landfills) "
                    "and coordination of local governments; Local level - management of municipal "
                    "solid waste from source to disposal and development of municipal waste-management "
                    "by-laws."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: roles for each of the three tiers of government are explicitly and "
                    "individually defined (+0.25). No explicit monitoring, enforcement mechanism or "
                    "unconditional guarantee is specified within the policy text itself for this "
                    "allocation."
                ),
                comments=(
                    "Coded per Rule 11 as a multi-level governance coordination instrument. Coded "
                    "in force (S=1) because the role division is described as the operative framework "
                    "already guiding current practice, not a contingent/unexercised enabling power - "
                    "flagged as an area of judgement per Rule 12, since as a policy (not a binding Act) "
                    "it is not directly justiciable on its own."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Recycling",
                instrument_description=(
                    "Zero-waste-to-landfill-by-2030 target underpinned by an 3R/circular-economy "
                    "strategy: source segregation promotion, public-private partnership, baseline-data "
                    "and technology research, and capacity building of SWM stakeholders, with \"robust "
                    "monitoring and evaluation to ensure effectiveness.\""
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: Ministry of Forests and Environment/local bodies identified as "
                    "implementers (+0.25). Monitoring: policy explicitly calls for \"robust monitoring "
                    "and evaluation\" (+0.25). No specific enforcement mechanism (fines/penalties) or "
                    "unconditional guarantee tied to the 2030 target is specified."
                ),
                comments="This is a strategic planning/target-setting instrument rather than a regulatory rule.",
            ),
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="End of life",
                instrument_description=(
                    "Policy \"lays the groundwork for Extended Producer Responsibility (EPR) - "
                    "producers bearing accountability for the packaging waste they generate\" as a "
                    "forward strategic direction."
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "No sub-score evidenced: this is described as a groundwork/aspirational direction "
                    "rather than a binding obligation with an identified authority, enforcement or "
                    "monitoring mechanism within this policy."
                ),
                comments=(
                    "Enabling/aspirational EPR reference only (\"lays the groundwork for\") - not yet "
                    "operationalised within this policy, hence scored not-in-force per Rule 12. The "
                    "actually binding, operative plastics-specific EPR obligation is in the Plastic Bag "
                    "(Regulation and Control) Directive, 2082 (see separate policy rows above)."
                ),
            ),
        ],
    ),
    dict(
        policy_name="Solid Waste Management Act, 2068 (2011)",
        policy_url="https://faolex.fao.org/docs/pdf/nep137767.pdf",
        policy_year=2011,
        policy_objective=(
            "To provide Nepal's primary legal framework for systematic and effective management of "
            "household, industrial, chemical, health-care and other hazardous solid waste (which "
            "includes plastic waste) from source to final disposal, in order to protect public "
            "health and the environment, and to define the responsibilities of local bodies, the "
            "private sector and waste generators; supplemented by the Solid Waste Management Rules, "
            "2070 (2013)."
        ),
        policy_target=0,
        policy_target_text="",
        policy_type=1,
        policy_type_justification=(
            "\"Solid Waste Management Act, 2068 (2011)\" - an Act enacted by the Legislature-"
            "Parliament of Nepal; primary legislation, not an executive regulation (the accompanying "
            "Solid Waste Management Rules, 2070 (2013) are the subordinate regulation implementing "
            "it)."
        ),
        policy_integration=1,
        policy_sectors_list="waste management, municipalities, industry, health, chemicals, households, private sector",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="consumption, recycling, disposal, environmental leakage",
        policy_budget=0.5,
        policy_budget_text=(
            "Solid Waste Management Rules, 2070 (2013), Rule 4(2)(d): waste generators are directed "
            "\"to pay the prescribed service charge at the specified time\" - local bodies may levy "
            "waste-service tariffs to fund SWM operations, but no dedicated national fund is created."
        ),
        instruments=[
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Recycling",
                instrument_description=(
                    "S.6: \"The local body shall prescribe to separate the solid waste into at least "
                    "organic and inorganic including different kinds at its source\"; the liability for "
                    "transporting segregated waste to the collection centre lies with the waste "
                    "generator (individual, organisation or body)."
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: local body designated to prescribe and enforce segregation (+0.25). "
                    "Monitoring: Rule 3(3) requires local bodies to \"conduct programs for increasing "
                    "people's awareness\" in relation to source segregation (+0.25). Enforcement: no "
                    "explicit fine/penalty for individual non-segregation identified in the reviewed "
                    "text. Unconditional: hazardous/chemical waste is carved out for separate treatment "
                    "(0)."
                ),
                comments=(
                    "Secondary sources (UNDP 2020; Doko Recyclers 2026) note source segregation is "
                    "practiced in only ~25% of municipalities despite the mandatory language, "
                    "indicating an implementation gap not reflected in the formal in-force test."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Chapter 4: \"The concerned local body may issue a license to a private company "
                    "for collection, transportation and disposal of solid waste, use, reuse, recycled "
                    "use or processing of solid waste, and enhancement of public awareness in the "
                    "reduction of solid waste.\""
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: local body empowered to license private operators (+0.25). No "
                    "explicit monitoring, enforcement or unconditional guarantee specified in the "
                    "licensing provision itself."
                ),
                comments=(
                    "Although worded as an enabling power (\"may issue a license\"), this has been "
                    "widely exercised in practice (e.g., Hetauda Sub-Metropolitan, Biratnagar "
                    "Metropolitan, Lalitpur Metropolitan City PPP arrangements), so coded in force per "
                    "Rule 12's operationalisation test."
                ),
            ),
            dict(
                instrument_type=0.80,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Local bodies are mandated to construct and operate solid-waste facilities "
                    "(transfer stations, landfills/sanitary landfill sites, processing plants) and to "
                    "organise waste at collection centres and transfer stations."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: local bodies explicitly designated as responsible (+0.25). No "
                    "monitoring or enforcement mechanism specified for facility construction "
                    "obligations themselves; unconditional not met given widely reported funding and "
                    "land-availability gaps."
                ),
                comments=(
                    "Real-world implementation is weak: only 7 of 97 surveyed landfill sites had "
                    "leachate/treatment infrastructure (Central Bureau of Statistics 2020 survey, cited "
                    "in SWITCH-Asia 2025 country profile), and CREASION (2023) reports local "
                    "governments \"lack sufficient budget\" and land for sanitary landfill construction."
                ),
            ),
        ],
    ),
    dict(
        policy_name="Local Government Operation Act, 2074 (2018)",
        policy_url="https://faolex.fao.org/docs/pdf/nep187145.pdf",
        policy_year=2018,
        policy_objective=(
            "To operationalise the federal constitutional structure at the local level by "
            "specifying the functions, powers and law-making authority of municipalities and rural "
            "municipalities, including sanitation, waste collection, recycling, disposal and tariff "
            "regulation, and by encouraging municipal partnerships with the private and "
            "non-governmental sector for waste and sanitation services."
        ),
        policy_target=0,
        policy_target_text="",
        policy_type=1,
        policy_type_justification=(
            "Act enacted by the Federal Parliament of Nepal operationalising local-government "
            "authority under the 2015 Constitution; primary legislation."
        ),
        policy_integration=0.75,
        policy_sectors_list="waste management, municipalities, sanitation, private sector, non-governmental organisations",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="consumption, recycling, disposal",
        policy_budget=0.5,
        policy_budget_text=(
            "Municipalities are empowered to regulate and collect waste-management tariffs; local "
            "governments may also request additional SWM budget from provincial/federal government "
            "(no single ring-fenced fund is created by this Act)."
        ),
        instruments=[
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Assigns municipalities/rural municipalities key sanitation responsibilities: "
                    "promoting sanitation awareness, managing waste collection, recycling and "
                    "disposal, and regulating waste-service tariffs; encourages partnerships with "
                    "private and non-governmental entities."
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: local government bodies explicitly designated (+0.25). Unconditional: "
                    "the mandate is stated as a general, unqualified local function with no exemptions "
                    "identified (+0.25). No specific enforcement or monitoring mechanism is attached to "
                    "this provision itself."
                ),
                comments=(
                    "Coded per Rule 11: this is a planning/enforcement-responsibility-assignment "
                    "instrument imposed on sub-national authorities."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Constrains local law-making: local by-laws/policies on waste and sanitation "
                    "\"shouldn't encroach upon exclusive powers of Federal and Province Level,\" must "
                    "\"avoid duplication with the federal and provincial laws,\" and must \"remain "
                    "consistent with federal law\" and \"national policy or priorities.\""
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: consistency requirement is a binding legal constraint on local "
                    "governments (+0.25). No explicit monitoring or enforcement mechanism evidenced for "
                    "this specific constraint."
                ),
                comments=(
                    "Coded per Rule 11 as a multi-level governance coordination instrument governing "
                    "the vertical relationship between federal, provincial and local waste/sanitation "
                    "rule-making."
                ),
            ),
        ],
    ),
    dict(
        policy_name="Federal, Provincial and Local Level (Coordination and Interrelation) Act, 2077 (2020)",
        policy_url="https://files.creasion.org/uploads/2026/02/152b6798.pdf",
        policy_year=2020,
        policy_objective=(
            "To establish coordination and cooperation mechanisms between the federal, provincial "
            "and local levels of government, define the exclusive and concurrent lists of powers "
            "(including local sanitation, public health and waste-related regulation), and create "
            "national and provincial coordination councils."
        ),
        policy_target=0,
        policy_target_text="",
        policy_type=1,
        policy_type_justification="Act enacted by the Federal Parliament of Nepal; primary legislation.",
        policy_integration=0.5,
        policy_sectors_list="waste management, sanitation, environment, municipalities",
        policy_circularity=0.5,
        policy_lifecycle_phases_list="consumption, disposal",
        policy_budget=0,
        policy_budget_text="",
        instruments=[
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Establishes a National Coordination Council and Provincial Coordination Councils "
                    "responsible for coordination and collaboration between the three tiers of "
                    "government, including on matters within the concurrent list (which covers local "
                    "sanitation/public-health regulation, encompassing solid and plastic waste)."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: the councils are formally established with defined functions (+0.25). "
                    "No monitoring, enforcement or unconditional guarantee specifically evidenced for "
                    "the coordination function itself."
                ),
                comments=(
                    "Textbook example of the multi-level governance coordination instrument introduced "
                    "in v2 of this coding prompt (Rule 11) - a national-provincial-municipal "
                    "coordination framework."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Consumption",
                instrument_description=(
                    "Concurrent-list clause 6: \"The regulation and awareness at the local level "
                    "regarding the quality of service delivery and public health, including "
                    "consumption, public service delivery at the local level will be from the local "
                    "level or against the local law.\""
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: local-level regulatory competence over consumption-related public "
                    "health/service-delivery matters is explicitly assigned (+0.25). No further "
                    "sub-scores evidenced."
                ),
                comments="",
            ),
        ],
    ),
    dict(
        policy_name="Inter-Governmental Fiscal Transfer Act, 2074 (2017)",
        policy_url="https://files.creasion.org/uploads/2026/02/152b6798.pdf",
        policy_year=2017,
        policy_objective=(
            "To regulate the transfer of fiscal resources (revenue-sharing, and special, "
            "complementary and conditional grants) between the federal, provincial and local levels "
            "of government, including financing mechanisms that fund solid-waste-management "
            "infrastructure (e.g., landfills and transfer stations) and to regulate local taxing "
            "powers, including on waste transportation."
        ),
        policy_target=0,
        policy_target_text="",
        policy_type=1,
        policy_type_justification="Act enacted by the Federal Parliament of Nepal; primary legislation.",
        policy_integration=0.5,
        policy_sectors_list="waste management, municipalities, public finance",
        policy_circularity=0.5,
        policy_lifecycle_phases_list="recycling, disposal",
        policy_budget=1,
        policy_budget_text=(
            "\"National Natural Resources and Fiscal Commission (NNRFC) to suggest on the allocation "
            "for SWM from the corpus of 'special' or 'complementary' or 'conditional' [grants] based "
            "on the requirements received from MoFAGA; and availability of fund at MoF\"; s.19: \"The "
            "Local Level may levy tax by making a law on matters falling within its jurisdiction...\" "
            "- a dedicated intergovernmental grant corpus plus a self-generated local-tax mechanism "
            "for solid-waste financing."
        ),
        instruments=[
            dict(
                instrument_type=0.60,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Conditional/special/complementary intergovernmental grant mechanism through which "
                    "the Ministry of Finance, on NNRFC recommendation, disburses funds to provinces and "
                    "local governments for SWM infrastructure projects (e.g., transfer stations, "
                    "landfills) - e.g., construction of the Banchare Danda landfill site financed via "
                    "the Ministry of Urban Development and handed over to Kathmandu Metropolitan City."
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: NNRFC/MoF/MoFAGA roles in the disbursement process are explicitly "
                    "defined (+0.25). Monitoring: process includes \"financial sanctions, monitoring "
                    "and evaluation\" of fund transfers (+0.25). No enforcement penalty or unconditional "
                    "guarantee evidenced."
                ),
                comments="",
            ),
            dict(
                instrument_type=0.60,
                instrument_lifecycle_stage="Disposal",
                instrument_description=(
                    "S.19: local levels \"may levy tax by making a law on matters falling within its "
                    "jurisdiction\"; in practice this underpins local waste-transport levies "
                    "(colloquially \"kabadi tax\") of roughly NRS 0.5-1.5/kg charged at multiple points "
                    "along the recyclables value chain."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: local levying power explicitly granted (+0.25). Unconditional: not "
                    "met - rates and application points vary widely and are reported to be applied \"at "
                    "around four different points,\" discouraging recyclers/collectors (0). No "
                    "dedicated monitoring/enforcement mechanism evidenced for the tax itself."
                ),
                comments=(
                    "CREASION (2023) explicitly flags this uncoordinated, multi-point local taxation "
                    "as a barrier to plastic recycling economics - included per the Plastics Relevance "
                    "Filter criterion (c) as a waste-management/governance instrument commonly applying "
                    "to plastic waste streams."
                ),
            ),
        ],
    ),
    dict(
        policy_name="National Environmental Policy, 2019",
        policy_url="https://dpnet.org.np/resource-detail/1817",
        policy_year=2019,
        policy_objective=(
            "To set Nepal's overarching environmental policy direction, including reducing the "
            "environmental and health impacts of plastic waste, promoting alternatives to "
            "single-use plastics, strengthening waste segregation/recycling infrastructure, and "
            "encouraging circular-economy principles and stricter enforcement of plastic "
            "production/disposal regulation."
        ),
        policy_target=0.5,
        policy_target_text=(
            "\"[The policy] highlights the detrimental impact of plastic on ecosystems and human "
            "health, promoting a comprehensive approach to reduce environmental pollution... The "
            "importance of reducing plastic use is encouraged via the adoption of alternatives to "
            "single-use plastics.\" (paraphrased in SWITCH-Asia's 2025 country profile; no numeric "
            "percentage/threshold target identified in the sources reviewed)."
        ),
        policy_type=0.25,
        policy_type_justification=(
            "A national policy approved by the Government of Nepal (Cabinet-level) setting "
            "aspirational strategic direction; no quantifiable, plastic-specific target was "
            "identified in the sources reviewed, and it is not an Act or a regulation/decree."
        ),
        policy_integration=0.5,
        policy_sectors_list="waste management, industry, consumption, recycling",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="consumption, recycling, disposal, environmental leakage",
        policy_budget=0,
        policy_budget_text="",
        instruments=[
            dict(
                instrument_type=0.20,
                instrument_lifecycle_stage="Consumption",
                instrument_description=(
                    "Public-awareness/behaviour-change strategy encouraging adoption of alternatives "
                    "to single-use plastics and \"increased public awareness and education on the "
                    "negative effects of plastic waste... to foster behavioural changes in consumption "
                    "patterns.\""
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "No specific responsible authority, enforcement mechanism, monitoring requirement "
                    "or unconditional guarantee is evidenced in the (paraphrased) policy text reviewed; "
                    "framed as a general aspiration (\"encouraged,\" \"endorsed\") rather than a "
                    "mandatory measure."
                ),
                comments=(
                    "Coded as not in force because the language reviewed (via secondary summary) is "
                    "consistently aspirational (\"encouraged,\" \"suggests,\" \"calling for\") with no "
                    "\"shall\"/\"must\" formulation identified - flagged as uncertain (Rule 8) since the "
                    "primary Nepali-language text was not independently reviewed."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Recycling",
                instrument_description=(
                    "Strategy to enhance waste-management systems, including segregating and recycling "
                    "plastic waste, developing recycling infrastructure, and promoting circular-economy "
                    "principles; also \"suggests implementing stricter regulations and enforcement "
                    "mechanisms to control plastic production and disposal.\""
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "No specific authority, monitoring, enforcement or unconditional guarantee "
                    "evidenced; phrased as a forward-looking suggestion rather than a binding "
                    "obligation."
                ),
                comments=(
                    "Consistent with the policy_type=0.25 coding (aspirational strategy without "
                    "concrete/enforceable targets)."
                ),
            ),
        ],
    ),
    dict(
        policy_name="The Sixteenth Plan (National Five-Year Development Plan, Fiscal Year 2024/25-2028/29)",
        policy_url="https://www.switch-asia.eu/site/assets/files/4512/plastic_policies_np.pdf",
        policy_year=2024,
        policy_objective=(
            "To set Nepal's national socio-economic development priorities for FY2024/25-2028/29, "
            "including - within its environmental-sustainability chapter(s) - directions on "
            "sustainable waste management, restricting open burning/river-dumping of waste, taxing "
            "landfilled waste, regulating plastic products, and strengthening pollution monitoring "
            "and enforcement of environmental standards."
        ),
        policy_target=0.5,
        policy_target_text=(
            "\"...there is mention of imposing restrictions on burning and disposing of waste in "
            "rivers and forests, and implementing taxes on waste dumped in landfills\"; \"Implementing "
            "scientific and transparent standards for the extraction and use of natural resources, "
            "including the regulation of plastic products to reduce environmental impact, is "
            "suggested.\" (paraphrased secondary summary; primary planning-document tables/annexes "
            "with any numeric plastics target were not independently verified)."
        ),
        policy_type=0.25,
        policy_type_justification=(
            "A five-year development plan produced by the National Planning Commission and approved "
            "by the Government of Nepal - a strategic planning document, not an Act or regulation. "
            "No confirmed numeric, plastic-specific target was identified in the sources reviewed, so "
            "this is coded 0.25 rather than 0.50."
        ),
        policy_integration=0.5,
        policy_sectors_list="waste management, environment, industry, natural resources",
        policy_circularity=0.75,
        policy_lifecycle_phases_list="consumption, recycling, disposal, environmental leakage",
        policy_budget=0.5,
        policy_budget_text=(
            "\"...implementing taxes on waste dumped in landfills\" - a proposed revenue-generating "
            "mechanism, not confirmed as an operative, ring-fenced fund within the Plan text reviewed."
        ),
        instruments=[
            dict(
                instrument_type=0.60,
                instrument_lifecycle_stage="Disposal",
                instrument_description=(
                    "Proposed tax on waste dumped in landfills, framed as part of the Plan's waste "
                    "management strategy to \"minimise environmental harm\" and support reduce-reuse-"
                    "recycle with private-sector participation."
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "Framed prospectively (\"there is mention of imposing... implementing taxes\") "
                    "rather than as an already-operative levy; no authority, enforcement or monitoring "
                    "mechanism specified within the Plan text itself."
                ),
                comments=(
                    "Coded not-in-force (Rule 12): this reads as a planned/aspirational direction for "
                    "the plan period rather than a self-executing mandatory tax."
                ),
            ),
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Planned restriction on open burning of waste and disposal of waste in rivers and "
                    "forests."
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "\"There is mention of imposing restrictions\" - prospective/aspirational framing; "
                    "no identified authority, enforcement or monitoring mechanism within the Plan text "
                    "itself (distinct from municipal by-laws or the Environment Protection Act, which "
                    "are coded separately)."
                ),
                comments="",
            ),
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Production",
                instrument_description=(
                    "Proposed \"scientific and transparent standards for the extraction and use of "
                    "natural resources, including the regulation of plastic products to reduce "
                    "environmental impact.\""
                ),
                instrument_in_force=0,
                instrument_implementation=0,
                instrument_implementation_text=(
                    "Framed as a suggestion for the plan period (\"is suggested\") rather than an "
                    "operative standard; no authority, enforcement or monitoring evidenced within the "
                    "Plan text itself."
                ),
                comments=(
                    "Uncertain (Rule 8): the primary Nepali-language Plan document (with chapter/table "
                    "citations) was not independently reviewed; this row relies on a secondary English "
                    "paraphrase (SWITCH-Asia, 2025) and should be re-verified against the original text "
                    "if available."
                ),
            ),
        ],
    ),
    dict(
        policy_name=(
            "Notice of Khumbu Pasang Lhamu Rural Municipality Banning Single-Use Plastics "
            "(Everest/Solukhumbu region)"
        ),
        policy_url="https://myrepublica.nagariknetwork.com/news/everest-region-to-ban-plastic-items-below-30-microns-from-2020",
        policy_year=2019,
        policy_objective=(
            "To ban plastic items thinner than 30 microns and all plastic beverage bottles within "
            "Khumbu Pasang Lhamu Rural Municipality (the Everest/Solukhumbu trekking and "
            "mountaineering region), in order to reduce the volume of plastic waste left behind by "
            "the large seasonal influx of tourists, trekkers and mountaineers."
        ),
        policy_target=1,
        policy_target_text=(
            "\"The executive council of Khumbu Pasang Lhamu Rural Municipality... has decided to ban "
            "plastic items of less than 30 microns thickness... All types of plastic bags, bottles "
            "and items not meeting the given standard will be banned here.\""
        ),
        policy_type=0.75,
        policy_type_justification=(
            "Executive decision of the Khumbu Pasang Lhamu Rural Municipality's executive council "
            "(a local-government executive body), published as a public notice - not national "
            "legislation. IMPORTANT CAVEAT: this is a sub-national/municipal instrument, included "
            "here only because it appears to correspond to an item referenced in the source material "
            "provided; it falls outside the national-policy scope that the 4P Index is designed to "
            "capture, and its policy-level fields (A-O) should be treated with lower confidence than "
            "the national-level policies above."
        ),
        policy_integration=0.5,
        policy_sectors_list="tourism, retail, consumption, hospitality",
        policy_circularity=0.5,
        policy_lifecycle_phases_list="consumption, environmental leakage",
        policy_budget=0,
        policy_budget_text="",
        instruments=[
            dict(
                instrument_type=1.0,
                instrument_lifecycle_stage="Consumption",
                instrument_description=(
                    "Ban on plastic items thinner than 30 microns and on plastic beverage bottles "
                    "(e.g., Coke, Fanta, Sprite) within the rural municipality, effective 1 January "
                    "2020; metal-can beverages remain permitted; households given reusable bags as a "
                    "substitute."
                ),
                instrument_in_force=1,
                instrument_implementation=0.5,
                instrument_implementation_text=(
                    "Authority: rural municipality executive council designated (+0.25). "
                    "Unconditional: framed as a blanket ban on the covered product categories "
                    "(+0.25). Enforcement: explicitly absent at the time of the announcement - \"no "
                    "penalty had been agreed on for people violating the rule yet\" (BBC, 2019; The "
                    "Hindu, 2019) (0). Monitoring: none specified beyond coordination with trekking "
                    "companies/airlines (0)."
                ),
                comments=(
                    "Best-effort identification (Rule 8): this may correspond to a row in the source "
                    "spreadsheet referencing a single-use-plastics ban in Himalayan tea-house/lodge "
                    "areas, but the image legibility did not allow certain confirmation of policy name, "
                    "URL or exact article text, so this row was instead built from independently "
                    "verified news reporting."
                ),
            ),
            dict(
                instrument_type=0.40,
                instrument_lifecycle_stage="Waste management",
                instrument_description=(
                    "Coordination arrangement between the rural municipality, trekking companies, "
                    "airlines and the Nepal Mountaineering Association to enforce the ban and raise "
                    "visitor awareness."
                ),
                instrument_in_force=1,
                instrument_implementation=0.25,
                instrument_implementation_text=(
                    "Authority: municipality plus named private/sectoral partners identified (+0.25). "
                    "No monitoring, enforcement or unconditional guarantee evidenced."
                ),
                comments=(
                    "Multi-stakeholder (public-private, not strictly intergovernmental) coordination "
                    "instrument, included by analogy to Rule 11."
                ),
            ),
        ],
    ),
]


def score_policy(p):
    return round((p["policy_type"] + p["policy_integration"] + p["policy_circularity"] + p["policy_budget"]) / 4, 4)


def score_instrument(i):
    return round(i["instrument_in_force"] * (i["instrument_type"] + i["instrument_implementation"]) / 2, 4)


def main():
    rows = []
    for p in POLICIES:
        o = score_policy(p)
        for instr in p["instruments"]:
            v = score_instrument(instr)
            row = {
                "policy_name": p["policy_name"],
                "policy_url": p["policy_url"],
                "policy_year": p["policy_year"],
                "policy_objective": p["policy_objective"],
                "policy_target": p["policy_target"],
                "policy_target_text": p["policy_target_text"],
                "policy_type": p["policy_type"],
                "policy_type_justification": p["policy_type_justification"],
                "policy_integration": p["policy_integration"],
                "policy_sectors_list": p["policy_sectors_list"],
                "policy_circularity": p["policy_circularity"],
                "policy_lifecycle_phases_list": p["policy_lifecycle_phases_list"],
                "policy_budget": p["policy_budget"],
                "policy_budget_text": p["policy_budget_text"],
                "policy_score": o,
                "instrument_type": instr["instrument_type"],
                "instrument_lifecycle_stage": instr["instrument_lifecycle_stage"],
                "instrument_description": instr["instrument_description"],
                "instrument_in_force": instr["instrument_in_force"],
                "instrument_implementation": instr["instrument_implementation"],
                "instrument_implementation_text": instr["instrument_implementation_text"],
                "instrument_score": v,
                "comments": instr["comments"],
            }
            rows.append(row)

    with open("4p_index_nepal_plastic_policies.csv", "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=COLUMNS)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print(f"Wrote {len(rows)} instrument-level rows across {len(POLICIES)} policies.")


if __name__ == "__main__":
    main()
