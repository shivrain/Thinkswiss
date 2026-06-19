#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Política Nacional do Ambiente only.

Source: user-uploaded Boletim da República (National_Environmental_Policy_24ce.pdf).
Approved by Resolução n.º 5/95, de 3 de Agosto de 1995.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_Resolucao_5_95.xlsx")
WORD_OUT = Path("/workspace/Resolucao_5_95_English_Translation.docx")

COLUMNS = [
    "A_policy_name",
    "B_policy_url",
    "C_policy_year",
    "D_policy_objective",
    "E_policy_target",
    "F_policy_target_text",
    "G_policy_type",
    "H_policy_type_justification",
    "I_policy_integration",
    "J_policy_sectors_list",
    "K_policy_circularity",
    "L_policy_lifecycle_phases_list",
    "M_policy_budget",
    "N_policy_budget_text",
    "O_policy_score",
    "P_instrument_type",
    "Q_instrument_lifecycle_stage",
    "R_instrument_description",
    "S_instrument_in_force",
    "T_instrument_implementation",
    "U_instrument_implementation_text",
    "V_instrument_score",
    "W_comments",
]

POLICY = {
    "policy_name": "Política Nacional do Ambiente (Resolução n.º 5/95, de 3 de Agosto)",
    "policy_url": "not available",
    "policy_year": 1995,
    "policy_objective": (
        "Establish the national environmental policy framework for sustainable development "
        "in Mozambique, setting objectives, principles, and sectoral action strategies for "
        "environmental management — including urban solid waste collection, treatment and "
        "recycling, coastal pollution control, polluter-pays obligations, and preparation "
        "of environmental legislation applicable to plastic pollution and waste streams."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.25,
    "policy_type_justification": (
        "Council of Ministers resolution (Resolução n.º 5/95) approving the National "
        "Environmental Policy as an annex. Strategy-level instrument with aspirational "
        "commitments and action plans but no quantifiable national targets."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "production, consumption, waste management, recycling, agriculture, industry, "
        "tourism, municipalities, fisheries, water, retail, transport, urban, rural, marine"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0.5,
    "policy_budget_text": (
        "§2.3.5: 'elaborar programas específicos e disponibilizar verbas para o efeito' "
        "(environmental education). §3.3.1 immediate activities: Capacity 21 teacher-training "
        "project 'financiado pelo Programa das Nações Unidas para o Desenvolvimento'. "
        "Funding mentioned but not ring-fenced to the policy."
    ),
    "instruments": [
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "§2.2 (Principles): Polluter-pays — 'o poluidor deve repor a qualidade do "
                "ambiente danificado e/ou pagar os custos para a prevenção e eliminação da "
                "poluição por si causada.' Foundational economic principle for pollution "
                "including plastic waste and marine litter."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Principle declared as binding basis for policy implementation ('A política "
                "do ambiente será implementada de acordo com os princípios seguintes'). "
                "No specific authority or penalty mechanism in this provision (+0.25 unconditional)."
            ),
            "comments": (
                "Economic instrument at principle level; later operationalised through Lei 20/97 "
                "and sectoral decrees. Dual nature: governance principle + economic liability."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§2.2 (Principles): Public participation — 'deve ser garantida a participação "
                "pública na tomada de decisões com impactos ambientais.'"
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Mandatory policy principle ('deve ser garantida'). No enforcement mechanism "
                "specified in this section."
            ),
            "comments": "Governance/participation instrument at policy level.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§2.3.1 & §3.1.1: Strengthen the Ministry for Environmental Coordination "
                "(MICOA) to coordinate, plan, supervise and monitor environmental activities "
                "across sectors, including inspection and monitoring of socio-economic "
                "development activities."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "MICOA designated as coordinating authority (+0.25). "
                "'garantir a fiscalização, inspecção e monitorização das actividades de "
                "desenvolvimento socio-económica' (+0.25 monitoring)."
            ),
            "comments": (
                "Multi-sector governance/coordination instrument commonly applied to municipal "
                "waste and pollution control (filter criterion c)."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.1.2 (Decentralization): Establish provincial representations, Sustainable "
                "Development Centres, and provincial directorates to decentralize environmental "
                "management to local level, including technical assistance to local governments."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'estabelecer representações do Ministério para Coordenação de Acção Ambiental "
                "em todas as províncias'; 'estabelecer direcções provinciais'; 'prestar "
                "assistência técnica ambiental aos governos provinciais, órgãos locais.' "
                "Responsible authorities designated (+0.25)."
            ),
            "comments": "National–provincial–local coordination framework for waste governance.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.1.3: Propose creation of the National Council for Sustainable Development "
                "(CNDS/CONDES) subordinated to the Council of Ministers; establish sector "
                "technical secretariats and define inter-sectorial environmental responsibilities."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'propor a criação do Conselho Nacional de Desenvolvimento Sustentável (CNDS).' "
                "Enabling/planning language — proposal not yet exercised in this instrument "
                "(later operationalised via Lei 20/97 Art. 6)."
            ),
            "comments": (
                "In-force = 0 per Rule 12. Cross-reference: CONDES created by Lei 20/97 (separate policy)."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§2.3.3 & §3.1.3: Integrate environmental variables into medium- and long-term "
                "socio-economic development planning at conception and implementation stages."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'A sustentabilidade dos planos de desenvolvimento... pressupõe que aqueles... "
                "integrem também directrizes referentes às variáveis ambientais.' Policy "
                "planning obligation (+0.25 authority via planning process)."
            ),
            "comments": "Planning integration instrument applicable to industrial and waste planning.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§3.2 (Environmental legislation strategy): Propose adoption of Environmental "
                "Law (Lei do Ambiente), complementary regulations, EIA rules, and alignment "
                "of sectoral legislation; establish multisectoral working group on environmental law."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'propor a aprovação da Lei do Ambiente'; 'garantir a aprovação da Lei do "
                "Ambiente' (immediate activity). Enabling/strategy language — legislative "
                "proposal, not operative regulation (+0.25 authority: Government/MICOA)."
            ),
            "comments": (
                "Enabling instrument; Lei 20/97 adopted 1997. Basis for subsequent plastic-waste "
                "regulations (Decreto 94/2014, Decreto 16/2015)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§3.2 (Immediate activities): Finalise regulations for Environmental Impact "
                "Assessment (AIA) and marine pollution prevention; develop sectoral EIA "
                "directives in coordination with line ministries."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'finalizar a elaboração dos regulamentos de Avaliação do Impacto Ambiental, "
                "AIA, e de prevenção de poluição marinha.' Planned regulation — not yet operative "
                "in this policy text (+0.25 authority designated)."
            ),
            "comments": (
                "Regulatory instrument type assigned to planned EIA/pollution rules. "
                "Applies to plastic production facilities and coastal pollution."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "§3.2 (Immediate activities): Elaborate regulations defining standards for "
                "water quality, domestic wastewater and its recycling, toxic residues, air "
                "quality, and soil quality."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'elaborar regulamentos, incluindo a definição de padrões sobre a qualidade "
                "de água, águas residuais domésticas e sua reciclagem, resíduos tóxicos, "
                "qualidade do ar e qualidade do solo.' Enabling — standards to be elaborated."
            ),
            "comments": (
                "Explicitly mentions wastewater recycling and toxic residues — relevant to "
                "plastic waste streams. Not operative until subordinate regulations adopted."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.2 (Medium/long-term): Establish and operationalise a corps of environmental "
                "auditors, inspectors and enforcement agents responsible for verifying application "
                "of environmental laws, norms and standards."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'estabelecer e operacionalizar um corpo de auditores, inspectores e fiscais "
                "ambientais, responsável pela verificação da aplicação das leis, normas e "
                "padrões ambientais.' Future/planned institutional measure."
            ),
            "comments": "Governance/enforcement framework; planned not operative in 1995 policy.",
        },
        {
            "instrument_type": 0.80,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.2 (Medium/long-term): Establish and operationalise an Environmental "
                "Monitoring Centre including a laboratory."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'estabelecer e operacionalizar um Centro de Monitoramento Ambiental, "
                "incluindo um laboratório.' Infrastructure investment planned (+0.25 authority: MICOA)."
            ),
            "comments": "Infrastructure instrument for environmental data and pollution monitoring.",
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§3.3.1 (Formal environmental education): Promote definition of national "
                "environmental education policy; integrate environmental aspects into school "
                "curricula at all levels; establish inter-sectoral cooperation mechanisms and "
                "Environment Clubs in schools."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'promover a definição de uma política de educação ambiental'; immediate "
                "activities include curriculum revision and creation of 'Clubes do Ambiente'. "
                "Government/MICOA and Education Ministry designated (+0.25)."
            ),
            "comments": (
                "Information/voluntary instrument; supports litter reduction and waste "
                "awareness including plastics."
            ),
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§3.3.2 (Non-formal environmental education): Implement non-formal education "
                "programmes for general public, business associations, women, youth, and "
                "decision-makers; mobile audio-visual units for rural environmental outreach."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "§2.3.5: 'o governo implementará programas de educação não-formal.' "
                "Government commitment to implement programmes (+0.25 authority)."
            ),
            "comments": "Awareness instrument for improper waste disposal and pollution.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.4: Establish Environmental Information Centre with library and database; "
                "create National Environmental Information Network linking government and NGO "
                "institutions; concentrate GIS on environmental monitoring and EIA analysis."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'estabelecer no Ministério para a Coordenação de Acção Ambiental um Centro "
                "de Informação Ambiental'; 'criar uma Rede Nacional de Informação Ambiental.' "
                "Planned institutional setup (+0.25)."
            ),
            "comments": "Data/governance instrument supporting waste and pollution management.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "§3.6.3 (Coastal and marine protection): Conduct national marine pollution "
                "studies to identify pollution sources, determine pollution levels, conduct "
                "regular monitoring, and define contamination parameters; implement national "
                "contingency plan for hydrocarbon spills."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'realizar estudos da poluição marinha a nível nacional... fazer o "
                "monitoramento regular e proceder a um controlo efectivo; definir parâmetros "
                "de contaminação.' Planned monitoring regime (+0.25 monitoring mechanism proposed)."
            ),
            "comments": (
                "Applies to marine plastic litter and coastal pollution. Contingency plan "
                "reference suggests partial operationalisation possible but not evidenced in policy text alone."
            ),
        },
        {
            "instrument_type": 0.80,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "§3.6.3 (Medium/long-term): Create conditions for pre-treatment of domestic "
                "and industrial waste from Maputo and Beira before discharge to sea, or ensure "
                "discharge points allow desired dilution with seawater."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'criar condições para que os resíduos domésticos e industriais, em especial "
                "os das cidades de Maputo e Beira, sejam previamente tratados antes de "
                "descarregados ao mar.' Planned infrastructure/ treatment requirement."
            ),
            "comments": "Directly addresses urban solid waste including plastics before marine disposal.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§3.6.5 (Tourism): Promote establishment of a minimum environmental protection "
                "fee/tax as a contribution from tourism units."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'promover o estabelecimento de uma taxa mínima para a protecção do ambiente, "
                "como contribuição das unidades turísticas.' Enabling — promote establishment, "
                "not a binding levy in force (+0.25 authority: Government)."
            ),
            "comments": "Economic instrument proposed for coastal tourism areas; not operative.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.7.1: Define national urban and environmental development policy; "
                "decentralize urban environmental management competence to municipalities; "
                "integrate environmental aspects in Local Government Reform Programme."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'descentralização da competência da gestão ambiental urbana para os municípios'; "
                "immediate activity: 'integrar os aspectos ambientais no Projecto de Reforma dos "
                "Órgãos Locais.' Municipal authority framework (+0.25)."
            ),
            "comments": "Multi-level urban waste governance coordination.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "§3.7.4 (Strategy): Build municipal capacity for management of domestic and "
                "hospital solid waste, improving collection, deposition and treatment of waste "
                "(lixos); introduce community participation mechanisms with incentives."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'capacitar os concelhos municipais nas áreas de gestão dos resíduos sólidos "
                "domésticos e hospitalares, visando melhorar o sistema de recolha, deposição "
                "e tratamento de lixos.' Capacity-building strategy (+0.25 authority: municipalities)."
            ),
            "comments": (
                "Primary plastics-relevant waste management instrument in this policy. "
                "Explicit reference to lixos (waste/garbage)."
            ),
        },
        {
            "instrument_type": 0.80,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "§3.7.4 (Immediate activities): Create and manage new landfills, introducing "
                "systems for treatment and recycling of solid waste."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'criar e gerir, convenientemente, novas lixeiras, introduzindo sistemas de "
                "tratamento e reciclagem dos resíduos sólidos.' Planned infrastructure with "
                "recycling component (+0.25 authority)."
            ),
            "comments": "Infrastructure + recycling; highest applicable type = 0.80 (infrastructure investment).",
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "§3.7.4 (Immediate activities): Progressively introduce mechanisms for citizens "
                "to practice separation of domestic solid waste."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'introduzir, progressivamente, nos cidadãos, mecanismos para a prática de "
                "separação dos resíduos sólidos domésticos.' Voluntary/behavioural programme "
                "planned (+0.25 authority: municipalities/Government)."
            ),
            "comments": "Waste separation prerequisite for plastic recycling streams.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "§3.7.4 (Medium/long-term): Adopt legislative measures obliging polluters to "
                "recycle their waste ('desperdícios')."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'adoptar medidas legislativas que obriguem os poluidores a procederem à "
                "reciclagem dos seus desperdícios.' Future enabling power — legislation not "
                "yet adopted in this policy (+0.25 authority: Government)."
            ),
            "comments": (
                "Closest explicit recycling obligation in policy; anticipates producer/polluter "
                "recycling responsibility. Not in force until subordinate legislation adopted."
            ),
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "§2.3.11: Government shall create incentives for private sector actors "
                "identified with concrete environmental preservation actions."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'O Governo criará incentivos para os que neste processo se revelarem "
                "identificados, com acções concretas.' Enabling/discretionary incentive power."
            ),
            "comments": "Economic incentive enabling power; could support clean production/recycling.",
        },
    ],
}

ENGLISH_TRANSLATION = """RESOLUTION NO. 5/95 OF 3 AUGUST 1995
Approving the National Environmental Policy

Source document: User-uploaded Boletim da República, Series I, No. 49, 6 December 1995 (Supplement).

---

RESOLUTION OF THE COUNCIL OF MINISTERS

The Constitution of the Republic of Mozambique, in its Article 72, enshrines citizens' right to a balanced environment and the duty to defend it, assigning the State the task of materialising this through initiatives promoting ecological balance, conservation and preservation of nature.

The Government Programme for 1995–1999, approved by the Assembly of the Republic, recognises that human and natural resources are the basis of any country's economic and social development, and therefore need to be adequately managed so as not to degrade and compromise the future of present and future generations.

It having become necessary to establish the principles and strategy for implementation of the Government Programme for 1995–1999, exercising the competence attributed to it under subparagraph e) of paragraph 1 of Article 153 of the Constitution of the Republic, the Council of Ministers resolves:

Sole Article. The National Environmental Policy is approved, as annexed, and forms an integral part of this Resolution.

Approved by the Council of Ministers.
Published.
The Prime Minister, Pascoal Manuel Mocumbi.

---

NATIONAL ENVIRONMENTAL POLICY

I. Introduction

The national environmental policy is the instrument through which the Government clearly and unequivocally recognises the interdependence between development and the environment. It is a means for executing socio- and macro-economic policies that are environmentally acceptable in the country, aiming to promote and drive economic growth founded, as far as possible, on the universal precepts of sustainable development.

Effective use of natural resources, adoption of appropriate technologies for Mozambique's real conditions, and deep recognition of social dynamics are fundamental premises on which this policy rests, given their inseparability in the national development process.

Despite global transformations in recent years, Mozambique continues to face serious economic difficulties, with most of the population living in absolute poverty. Much of the country's natural resources are under immense pressure, as the needy population is forced to resort to unregulated exploitation of those resources as a guarantee of survival — a fact that contributes to accelerating environmental degradation.

It is in the spirit of national and international efforts tending to reduce growing environmental degradation, so that the environment remains available to future generations in the best possible conditions, that this national environmental policy is presented.

II. National Environmental Policy

Environmental policy represents the basis for sustainable development in Mozambique, aiming at the progressive eradication of poverty and improvement of Mozambicans' quality of life, as well as reduction of damage to the environment.

2.1 General objectives

The main objective of environmental policy is to ensure sustainable development of the country, considering its specific conditions, through an acceptable and realistic compromise between socio-economic progress and environmental protection. To this end, the policy aims to:

- ensure an adequate quality of life for citizens;
- ensure management of natural resources and the environment in general so that they maintain their functional and productive capacity for present and future generations;
- develop environmental awareness among the population to enable public participation in environmental management;
- ensure integration of environmental considerations in socio-economic planning;
- promote participation of local communities in planning and decision-making on use of natural resources;
- protect ecosystems and essential ecological processes;
- integrate regional and global efforts in seeking solutions to environmental problems.

2.2 Principles

Environmental policy shall be implemented according to the following principles:

- Man is an important component of the environment and the main beneficiary of its adequate management;
- use of natural resources must be optimised;
- laws, incentives and disincentives for environmental management must be applied;
- the polluter must restore damaged environmental quality and/or pay the costs of preventing and eliminating pollution caused by them;
- public participation in decision-making with environmental impacts must be guaranteed;
- local communities must benefit from distribution of revenues from rational use of natural resources;
- traditional knowledge of local communities in environmental management must be recognised and valued.

2.3 Specific aspects

2.3.1 Institutional aspects

The Ministry for Environmental Coordination has as its main objectives to coordinate, advise, control and evaluate the degree of use of the country's natural resources, promoting their preservation and rational use, as well as coordinating all environmental activities and ensuring integration of environmental variables in planning and management of socio-economic development.

For this purpose, the Ministry for Environmental Coordination will be strengthened, endowed with powers compatible with its role of planning, supervision and coordination of socio-economic activities, proceeding, as far as possible, with decentralisation and democratisation of natural resource management through instruments such as the Local Government Reform Programme, the Municipalities Law and other similar foundations.

2.3.2 Legal aspects

From the various existing sectoral laws and regulations that directly or indirectly address environmental complexity, it is found that they were elaborated according to the specificities and concerns of each State body, resulting in various environmental aspects not covered by legislation, or existing provisions needing updating in space and time, making them somewhat inadequate.

It thus becomes necessary and urgent to adopt an Environmental Law addressing environmental aspects as a whole and driving elaboration and application of specific regulations, such as environmental impact assessment, environmental quality standards, toxic products, industrial development, among others.

2.3.3 Integration of environmental aspects in socio-economic planning

Sustainability of development plans, in the medium and long term, presupposes that they integrate, beyond economic and social criteria, guidelines on environmental variables to be considered both in the conception phase and during and after implementation.

2.3.4 Development of sectoral policies

Despite existing potential in natural resources, many remain unexploited, or those under exploitation are subject to either very high exploitation rates or are under-exploited. Development policies adopted to date have proved inadequate and inconsistent from the standpoint of rational use of natural resources and establishment of an efficient productive system. In this sense, sectoral policies will play a decisive role in establishing a culture of rationality and sustainability in use of natural resources.

Population dynamics decisively influence the quantity and quality of available natural resources as well as patterns of their conservation. The defined development strategy sought to relate real availability of resources, in quantity and quality, with the number of users, their spatial distribution and the level and quality of technologies applied in resource exploitation, supported by a coherent population policy, ecologically sustainable industrial development and consistent agrarian policies.

2.3.5 Environmental education and outreach

Environmental education plays a fundamental role in creating and raising citizens' sensitivity for their growing involvement in seeking solutions to environmental problems. The national environmental education strategy therefore involves:

- providing extensive environmental education at all levels of the education system;
- elaborating specific programmes and making funds available for this purpose;
- reviewing methods and teaching practices used in public institutions responsible for environmental education initiatives.

Beyond measures directly related to formal education, the Government will implement non-formal education programmes to reach school-age children, the general public and specific groups such as business and commercial associations, women, youth, decision-makers and the social communication media.

2.3.6 Training of environmental professionals

The country needs a flexible, technically capable workforce with knowledge to face increasingly growing and complex environmental and development problems. For this, the Government will:

- train and capacitate necessary human resources;
- establish a multidisciplinary forum bringing together various academic trends, scientific institutions, research institutions, NGOs and other personalities, capable of advising and guiding competent bodies in implementation of environmental policy.

2.3.7 Environmental research and database

Research is of fundamental importance and utility for planning, implementation and timely correction of socio-economic activities. In the country, environmental research deserves special attention as it is little developed. In this area the Government prioritises:

- equipping research centres with more and better human, financial and material resources;
- informing about and disseminating, more systematically, results of research conducted;
- improving inter-institutional information flow;
- supporting environmental research through a database containing organised and systematised information;
- improving information flow with countries and international organisations, particularly SADC countries.

2.3.8 Environmental monitoring

Recent evidence shows that human activities have been affecting environmental quality. The new dynamic tending to promote economic development and general population well-being will cause more negative impacts on the environment. From this fact arises the need to establish appropriate mechanisms for surveillance and monitoring of degradation and other environmental changes, and a centre for analysis and follow-up of such changes. Knowledge resulting from permanent monitoring constitutes a valuable instrument for environmental management.

2.3.9 Role of women in environmental management

Women constitute more than half of the total population and more than half of the economically active population. Considering their direct dependence on natural resources due to responsibility for family sustenance, they have an obvious interest in maintaining such resources. However, with the current socio-economic state, women's action has become prejudicial to natural equilibrium, transforming them simultaneously into agents and victims of their own action. The environmental policy therefore dedicates special attention to women through natural resource management programmes, environmental education and other programmes tending to afford women equality of opportunity.

2.3.10 Role of the community in environmental management

Sustainability of natural resource and environmental management can only be effective through direct and active participation of communities, valuing and using their traditions and experiences. The Government will therefore create a favourable climate through knowledge of resource use patterns, traditional management forms and community lifestyles, while seeking to encourage and strengthen community capacity to know and apply principles and rules of natural resource management guiding society in general, delegating competencies and instruments facilitating closer cooperation with formal and informal structures.

2.3.11 Role of the private sector in environmental management

The private sector, whose role and contribution to national economic growth is assuming increasing importance, is called to participate actively and consciously in management of natural resources, assisting and sharing in environmental preservation efforts with concrete actions. The Government will create incentives for those identified in this process with concrete actions for these purposes.

2.3.12 International cooperation

Environmental problems know no borders between countries. Problems such as water and air pollution, droughts, climate change and many others always require a joint approach given their global character. Mozambique recognises this principle and reaffirms its commitment to seeking local solutions contributing to minimisation and/or resolution of these major international problems. Efforts will be made to strengthen exchange of information and research and to participate actively in international events requiring its participation, considering its specific responsibilities and according to its capacities.

III. Strategies and priority actions

The strategy defined for implementation of this policy rests on the following fundamental strands: identification of the priority intervention area, strategy to follow, immediate actions to be carried out, and medium- and long-term activities.

3.1 Strengthening institutional capacity for environmental management

Environmental management in general, and of natural resources in particular, is a complex, dynamic and sectorally integrated process. Clear strategies will be adopted on institutional development, decentralisation of environmental management, inter-sectorial coordination and technical-professional training.

3.1.1 Institutional development

Correct and efficient environmental management presupposes strengthening capacities of all intervening institutions. Priority in institutional capacitation lies fundamentally at the Ministry for Environmental Coordination level. The Ministry shall:

- define general policies of sustainable development, including respective legislation;
- establish and operationalise adequate mechanisms for good inter-sectorial coordination;
- promote, in coordination with related sectors, environmental education, research and outreach;
- ensure surveillance, inspection and monitoring of socio-economic development activities.

Immediate activities include structuring MICOA, defining staff framework, recruitment, training courses for decision-makers, and improving working conditions.

3.1.2 Decentralisation

Correct environmental management requires active and conscious involvement of all sectors of society, presupposing mechanisms for its effectiveness, with emphasis on decentralisation of natural resource management.

Strategy: establish MICOA representations in all provinces; establish national speciality centres for investigative and training environmental activities.

Immediate activities: establish provincial directorates; establish at least one of three Sustainable Development Centres in north, centre and south regions; provide technical environmental assistance to provincial governments, local bodies, NGOs, traditional authorities, civic associations, religious confessions, etc.

3.1.3 Inter-sectorial coordination

Coordination between sectors is an indispensable condition for efficient environmental management.

Strategy: create sector technical environmental secretariats; capacitate all sectors and decision-makers in the spirit of Agenda 21.

Immediate activities: propose creation of the National Council for Sustainable Development (CNDS) subordinated to the Council of Ministers; establish permanent secretariat; define responsibilities of various sectors in environmental management; emphasise in ministry statutes the need and obligation of sustainable use of resources.

3.1.4 Professional training

Institutions must have technically trained staff with solid environmental training. Given national shortages, the realistic alternative is environmental capacitation of current technicians through short courses and, in the medium and long term, full training.

3.2 Environmental legislation

Sustainability of socio-economic development presupposes definition and adoption of a clear and adequate environmental legal framework.

Strategy:
- propose approval of the Environmental Law;
- elaborate regulations for implementation of the Environmental Law;
- adjust sectoral laws and regulations to the Environmental Law;
- capacitate technical-professional staff in environmental legislation;
- constitute a multisectoral working group on environmental legislation.

Immediate activities:
- guarantee approval of the Environmental Law;
- finalise elaboration of Environmental Impact Assessment (EIA) regulations and marine pollution prevention regulations;
- define and elaborate directives for EIA execution by sector;
- elaborate regulations including standards on water quality, domestic wastewater and its recycling, toxic residues, air quality and soil quality;
- promote deep analysis of existing environmental legislation to identify gaps and overlaps;
- analyse level of insertion in national legislation of obligations under international agreements.

Medium- and long-term activities:
- create technical capacity in environmental legislation at MICOA and other sectors;
- establish and operationalise a corps of environmental auditors, inspectors and enforcement agents;
- establish and operationalise an Environmental Monitoring Centre including a laboratory;
- introduce environmental legislation aspects in university law curricula.

3.3 Environmental awareness and outreach

National effort and success in implementing a sustainable development policy depend on quality environmental education extensive to all sectors. Urgent promotion of information, training, awareness and sensitisation activities is needed for greater citizen involvement in identifying causes of environmental degradation and seeking solutions.

3.3.1 Formal education

Strategy: promote definition of environmental education policy; integrate environmental aspects into school programmes at all formal education levels; prioritise revision of primary school programmes and teacher training; establish functional institutional cooperation mechanisms between MICOA, Education Ministry and universities.

Immediate activities: establish cooperation mechanisms; support creation of inter-sectorial environmental education forum; define environmental education policy and implementation strategy; implement teacher training under Capacity 21 project (UNDP-funded); produce educational materials; create Environment Clubs in schools.

3.3.2 Non-formal education

Strategy: implement non-formal environmental education programmes for general public, associations, women, youth and decision-makers; use mobile audio-visual units for rural outreach.

3.4 Documentation, information and environmental research

Existence and accessibility of timely environmental information are important factors for correct management of environment and natural resources.

Strategy: establish Environmental Information Centre with library and database; create National Environmental Information Network; concentrate GIS on environmental monitoring, EIA analysis, digital cartography and database creation.

3.5 Rural zones

National development efforts necessarily pass through poverty reduction, especially in rural zones. Strategy includes rural production incentives and legal/institutional conditions for community and decentralised natural resource management.

3.6 Coastal and marine management

The Mozambican coast has great socio-economic development potential based on natural resources under strong demographic pressure. Coastal management will rest on sector coordination and an integrated coastal development plan based on coordinated investigation and data collection.

3.6.1 Fisheries

Fisheries is among the most important national economic activities. Strategy emphasises sustainable fisheries management, inter-sectorial cooperation with MICOA, legal framework improvement, and reduction of post-harvest losses.

3.6.2 Management of coastal and marine ecosystems

Gaps exist in ensuring sustainable management of coastal and marine ecosystems. Coordinated management policies and programmes are needed, especially for pilot sites.

3.6.3 Coastal and marine protection

Strategy includes coastal erosion studies, national marine pollution studies (identify sources, monitor pollution levels, define contamination parameters), protection of critical coastal habitats, and implementation of national contingency plan for hydrocarbon spill control.

Medium- and long-term activities: create conditions for pre-treatment of domestic and industrial waste from Maputo and Beira before discharge to sea; submit all coastal and marine development projects to rigorous EIA; prepare contingency plan for toxic chemical spills; promote ratification of relevant international conventions.

3.6.4 Marine parks

Strategy: define policy guiding creation and management of marine parks; develop adequate legislation; capacitate planners and managers.

3.6.5 Tourism

Strategy: approve and implement National Tourism Policy. Immediate activities include zoning of tourism areas, capacitation of local authorities, mobile technical brigades, environmental guide for tourism projects, and promote establishment of a minimum environmental protection fee as contribution from tourism units.

3.7 Urban environmental management

Urban environmental management in Mozambican cities is particularly difficult due to parallel traditional/informal and modern/formal socio-economic systems.

3.7.1 Institutional capacity and coordination mechanisms

Strategy: define national urban and environmental development policy; decentralize urban environmental management competence to municipalities; capacitate all sectors involved in urban management.

Immediate activities: integrate environmental aspects in Local Government Reform Programme covering Pemba, Nampula, Quelimane, Beira and Maputo; establish Nampula Sustainable Development Centre specialised in urban environmental management.

3.7.2 Rehabilitation/construction of urban sanitation systems and water supply

Strategy: extend water system coverage to city peripheries; research appropriate technologies; strengthen Low-Cost Sanitation National Programme in municipalities; formulate and adopt tariff policy based on real water supply cost.

3.7.3 Erosion and deforestation control programmes

Strategy: capacitate municipal councils in erosion control; recover municipal nurseries; ecological zoning; define urban energy policy promoting renewable resources.

3.7.4 Management of domestic and hospital solid waste

Correct management of domestic and hospital solid waste in national urban centres requires adequate technical capacity of municipal councils and community involvement.

Strategy:
- capacitate municipal councils in domestic and hospital solid waste management, improving collection, deposition and treatment of waste (lixos);
- introduce community participation mechanisms in solid waste management under Local Government Reform Programme and Urban Rehabilitation Project, creating incentive forms.

Immediate activities:
- create and manage new landfills, introducing treatment and recycling systems for solid waste;
- progressively introduce mechanisms for citizens to practice separation of domestic solid waste.

Medium- and long-term activities:
- adopt legislative measures obliging polluters to recycle their waste.

---

Note on source: This English translation is based on text extracted from the user-uploaded Boletim da República PDF (National_Environmental_Policy_24ce.pdf). The annex ends with a glossary (not reproduced here). Some OCR artefacts in the source were corrected against standard published versions where garbled."""


def build_rows() -> list[dict]:
    rows = []
    for inst in POLICY["instruments"]:
        rows.append({
            "A_policy_name": POLICY["policy_name"],
            "B_policy_url": POLICY["policy_url"],
            "C_policy_year": POLICY["policy_year"],
            "D_policy_objective": POLICY["policy_objective"],
            "E_policy_target": POLICY["policy_target"],
            "F_policy_target_text": POLICY["policy_target_text"],
            "G_policy_type": POLICY["policy_type"],
            "H_policy_type_justification": POLICY["policy_type_justification"],
            "I_policy_integration": POLICY["policy_integration"],
            "J_policy_sectors_list": POLICY["policy_sectors_list"],
            "K_policy_circularity": POLICY["policy_circularity"],
            "L_policy_lifecycle_phases_list": POLICY["policy_lifecycle_phases_list"],
            "M_policy_budget": POLICY["policy_budget"],
            "N_policy_budget_text": POLICY["policy_budget_text"],
            "O_policy_score": "[auto]",
            "P_instrument_type": inst["instrument_type"],
            "Q_instrument_lifecycle_stage": inst["instrument_lifecycle_stage"],
            "R_instrument_description": inst["instrument_description"],
            "S_instrument_in_force": inst["instrument_in_force"],
            "T_instrument_implementation": inst["instrument_implementation"],
            "U_instrument_implementation_text": inst["instrument_implementation_text"],
            "V_instrument_score": "[auto]",
            "W_comments": inst.get("comments", ""),
        })
    return rows


def write_excel(rows: list[dict]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index"

    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True, size=9)

    ws.append(COLUMNS)
    for col in range(1, len(COLUMNS) + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for i, row in enumerate(rows, start=2):
        ws.append([row[c] for c in COLUMNS])
        ws[f"O{i}"] = f"=AVERAGE(G{i},I{i},K{i},M{i},P{i})"
        ws[f"V{i}"] = f"=AVERAGE(P{i},T{i})"

    widths = {
        "A": 42, "B": 18, "C": 8, "D": 45, "E": 8, "F": 12, "G": 8, "H": 35,
        "I": 8, "J": 35, "K": 8, "L": 30, "M": 8, "N": 35, "O": 10,
        "P": 8, "Q": 20, "R": 50, "S": 8, "T": 8, "U": 40, "V": 10, "W": 40,
    }
    for col_letter, width in widths.items():
        ws.column_dimensions[col_letter].width = width

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if cell.column in (7, 9, 11, 13, 15, 16, 19, 20, 22):
                cell.number_format = "0.00"

    meta = wb.create_sheet("Metadata")
    meta.append(["Field", "Value"])
    meta.append(["Country", "Mozambique"])
    meta.append(["Index", "4P Index v2"])
    meta.append(["Policy", POLICY["policy_name"]])
    meta.append(["Source", "User upload: National_Environmental_Policy_24ce.pdf"])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Total instrument rows", len(rows)])
    meta.append([
        "Note",
        "Columns O and V use Excel formulas. "
        "policy_score = AVERAGE(G,I,K,M,P); instrument_score = AVERAGE(P,T).",
    ])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_translation_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading(
        "Política Nacional do Ambiente (Resolução 5/95) — English Translation",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English translation of Mozambique's National Environmental Policy, approved by "
        "Council of Ministers Resolution No. 5/95 of 3 August 1995, based on the user-uploaded "
        "Boletim da República scan. This is the sole policy covered by the accompanying "
        "4P Index coding table."
    )

    p = doc.add_paragraph()
    p.add_run("Portuguese title: ").bold = True
    p.add_run("Política Nacional do Ambiente (Resolução n.º 5/95, de 3 de Agosto)")

    p2 = doc.add_paragraph()
    p2.add_run("Source: ").bold = True
    p2.add_run("User-uploaded document (National_Environmental_Policy_24ce.pdf); policy URL not available.")

    doc.add_heading("English translation", level=2)

    for block in ENGLISH_TRANSLATION.strip().split("\n\n"):
        text = block.strip()
        if not text:
            continue
        if text.startswith("---"):
            continue
        if text.startswith("RESOLUTION") or text.startswith("NATIONAL") or text.startswith("III."):
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
        elif text.startswith(("2.", "3.")) and len(text) < 120:
            doc.add_heading(text, level=3)
        else:
            para = doc.add_paragraph(text)
            para.paragraph_format.space_after = Pt(6)

    doc.save(WORD_OUT)
    print(f"Wrote {WORD_OUT}")


def main() -> None:
    rows = build_rows()
    write_excel(rows)
    write_translation_doc()


if __name__ == "__main__":
    main()
