#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Lei do Ambiente (Lei 20/97) only.

Source: user-uploaded Boletim da República scan (Environmenmtal_law_75d0.pdf), OCR via EasyOCR.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_Lei_20_97.xlsx")
WORD_OUT = Path("/workspace/Lei_20_97_English_Translation.docx")

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
    "policy_name": "Lei do Ambiente (Lei n.º 20/97, de 1 de Outubro)",
    "policy_url": "not available",
    "policy_year": 1997,
    "policy_objective": (
        "Define the legal bases for correct use and management of the environment and its "
        "components, toward the materialisation of a system of sustainable development. "
        "The law establishes pollution prohibitions, environmental quality standards, licensing "
        "and EIA requirements, waste-related restrictions, inspection, and liability frameworks "
        "that apply to plastic pollution and waste streams."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1,
    "policy_type_justification": (
        "Legislation enacted by the Assembleia da República (Law No. 20/97), approved "
        "31 July 1997 and promulgated 1 October 1997. Not an executive decree or strategy."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "production, consumption, waste management, recycling, agriculture, industry, "
        "tourism, municipalities, fisheries, water, retail, transport"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 0,
    "policy_budget_text": "",
    "instruments": [
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 5: The Government shall elaborate and execute the National Environmental "
                "Management Programme (PNGA), the overarching planning instrument for "
                "environmental governance including waste and pollution control."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Cabe ao Governo elaborar e executar o Programa Nacional de Gestão Ambiental.' "
                "Responsible authority explicitly designated (Government)."
            ),
            "comments": (
                "Governance/coordination instrument. Applies to plastic-relevant environmental "
                "management via general waste/pollution framework (filter criterion c)."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 6: Creates the National Council for Sustainable Development (CONDES) as "
                "a consultative body of the Council of Ministers to coordinate environmental "
                "management and integrate sustainability principles in national development."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'é criado o Conselho Nacional de Desenvolvimento Sustentável.' "
                "Institutional coordinating body explicitly established."
            ),
            "comments": (
                "Multi-level governance coordination instrument (Rule 11). Composition and "
                "functioning regulated by Council of Ministers decree."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 7: Local-level environmental services are created to implement the Law, "
                "guaranteeing coordination of environmental action and decentralisation — "
                "including municipal waste management responsibilities."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'De nível local são criados serviços responsáveis pela implementação da "
                "presente Lei, os quais garantem coordenação da acção ambiental a esse nível "
                "e a descentralização na sua execução.'"
            ),
            "comments": (
                "National–local coordination framework commonly applied to municipal solid "
                "waste including plastic waste streams."
            ),
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 8: Government obligation to create adequate mechanisms to involve civil "
                "society sectors, local communities, and environmental defence associations in "
                "elaboration of environmental policy, legislation, and PNGA implementation."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'É obrigação do Governo criar mecanismos adequados para envolver os diversos "
                "sectores da sociedade civil, comunidades locais, em particular as associações "
                "de defesa do ambiente.'"
            ),
            "comments": "Information/participation instrument; no enforcement mechanism in this article.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "Art. 9: Prohibits production, deposit in soil/subsoil, and discharge to water "
                "or atmosphere of toxic and polluting substances beyond legally established "
                "limits; also activities accelerating erosion, desertification, deforestation, "
                "or other environmental degradation. Expressly prohibits import of hazardous "
                "waste/residues except as provided in specific legislation."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Não é permitida... de quaisquer substâncias tóxicas e poluidoras.' "
                "'É expressamente proibida a importação... de resíduos ou lixos perigosos.' "
                "Mandatory prohibition language; no exemptions stated in Art. 9 (+0.25 unconditional)."
            ),
            "comments": (
                "Directly applicable to plastic waste and hazardous plastic residues. "
                "Enforcement details in Arts. 27–28 coded separately."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 10: Government must establish environmental quality standards to ensure "
                "sustainable resource use, including norms and deadlines for adapting "
                "agricultural/industrial processes, machines and transport means, and devices "
                "to retain or neutralise polluting substances."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'O Governo deve estabelecer padrões de qualidade ambiental.' "
                "Responsible authority designated (+0.25). Standards include deadlines for "
                "industrial process adaptation and pollutant retention devices (+0.25 monitoring "
                "mechanism via prescribed norms and deadlines)."
            ),
            "comments": (
                "Mandatory obligation on Government ('deve'); individual standards adopted via "
                "subordinate regulation but the duty to establish them is operative."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Disposal",
            "instrument_description": (
                "Art. 14(1): Prohibits installation of infrastructure (including waste dumping) "
                "that causes significant negative environmental impact, especially in coastal "
                "zones, erosion/desertification zones, wetlands, protected areas, and other "
                "ecologically sensitive areas."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'É proibida a implantação de infra-estruturas... mesmo se aplicando à "
                "deposição de lixos ou materiais usados.' Unconditional prohibition in "
                "sensitive zones (+0.25). 'São estabelecidas por regulamento as normas' — "
                "Government authority for implementation norms (+0.25)."
            ),
            "comments": (
                "Directly restricts waste dumping including plastic litter disposal in "
                "sensitive coastal and aquatic environments."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 15: Environmental licensing and registration of activities susceptible to "
                "significant environmental impacts (by nature, location or scale) shall be "
                "conducted per a regime established by government regulation; environmental "
                "licence based on EIA and precedes all other required licences."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'o licenciamento e o registo das actividades... são feitos de acordo com regime "
                "a estabelecer pelo governo.' Government as licensing authority (+0.25). "
                "EIA-based assessment requirement (+0.25 monitoring)."
            ),
            "comments": (
                "Applies to plastic production facilities, packaging plants, and waste "
                "management installations. Licensing regime details in subordinate regulation."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 16–17: Environmental impact assessment (EIA) required based on EIA study "
                "by government-accredited entities. Minimum study content includes description "
                "of activity, baseline environment, modifications to environmental components, "
                "measures to suppress/reduce negative effects, and control/monitoring systems."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'A avaliação do impacto ambiental tem como base um estudo de impacto ambiental "
                "a ser realizado por entidades credenciadas pelo Governo.' Accredited entities "
                "(+0.25 authority). Study must include 'sistemas previstos para controlo e "
                "monitorização da actividade' (+0.25 monitoring)."
            ),
            "comments": "Arts. 16 and 17 coded as one instrument (EIA process and minimum content).",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 18: All operating activities lacking appropriate technologies/processes "
                "that result or may result in environmental damage are subject to environmental "
                "audits; repair costs borne by entrepreneurs."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'são objecto de auditorias ambientais.' Monitoring via audits (+0.25). "
                "'Os custos decorrentes da reparação dos danos ambientais... são da "
                "responsabilidade dos empreendedores' — enforcement/liability (+0.25). "
                "'São estabelecidas por regulamento as normas' — authority (+0.25). "
                "Applies to all non-compliant operating activities (+0.25 unconditional)."
            ),
            "comments": "Relevant to legacy industrial plastic production without adequate pollution controls.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 24: All persons have the obligation to use natural resources responsibly "
                "and sustainably, wherever they are."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Todas as pessoas têm a obrigação de utilizar os recursos naturais de forma "
                "responsável e sustentável.' Mandatory language; no exemptions (+0.25 unconditional)."
            ),
            "comments": (
                "General consumption-phase obligation applicable to improper disposal of "
                "plastic and other waste."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 25: All persons exercising activities involving high risk of environmental "
                "degradation, as classified under EIA legislation, must secure civil liability "
                "insurance."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'devem segurar a sua responsabilidade civil.' Mandatory insurance requirement "
                "(+0.25 unconditional). Classification by EIA legislation establishes "
                "designated authority framework (+0.25)."
            ),
            "comments": "Applies to high-risk activities including waste management and industrial facilities.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "Art. 26: Strict liability — all who cause significant environmental damage or "
                "paralyse economic activities through especially dangerous activities, "
                "regardless of fault, must pay compensation; Government supervises damage "
                "valuation via environmental expert appraisal."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Constituem-se na obrigação de pagar uma indemnização aos lesados todos "
                "aqueles que, independentemente de culpa... causem danos significativos ao "
                "ambiente.' Enforcement via indemnisation (+0.25). 'Compete ao Governo "
                "supervisar a avaliação da gravidade dos danos' (+0.25 monitoring)."
            ),
            "comments": "Polluter-pays enforcement mechanism applicable to plastic pollution damage.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Environmental leakage",
            "instrument_description": (
                "Art. 27: Environmental crimes and contraventions to be provided in specific "
                "legislation; Public Ministry responsible for defending environmental values "
                "protected by this Law."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'As infracções de carácter criminal, bem como as contravenções relativas ao "
                "ambiente, são objecto de previsão em legislação específica.' Enabling power — "
                "penalties defined elsewhere. 'Compete ao Ministério Público a defesa dos "
                "valores ambientais' (+0.25 authority)."
            ),
            "comments": (
                "In-force = 0 per Rule 12: crimes/contraventions deferred to separate legislation. "
                "May be partially operationalised via Penal Code and sectoral decrees; not evidenced "
                "in this Law text alone."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 28: Government shall create, under terms to be regulated, a corps of "
                "environmental inspection agents competent to ensure implementation of "
                "environmental legislation and take preventive measures."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Compete ao Governo criar, em termos a regulamentar, um corpo de agentes de "
                "fiscalização ambiental.' Enabling power ('em termos a regulamentar') — "
                "operative only once subordinate regulation creates the inspection corps."
            ),
            "comments": (
                "Governance/enforcement framework instrument. In-force = 0 unless subordinate "
                "regulation operationalising inspection agents is evidenced (not in this document)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 29: All persons in charge of an activity or place subject to inspection "
                "must collaborate with environmental inspection agents."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Todas as pessoas encarregues de uma actividade ou lugar sujeito à fiscalização "
                "devem colaborar com os agentes de fiscalização.' Mandatory collaboration duty."
            ),
            "comments": (
                "Binding duty operative in law text; effectiveness depends on Art. 28 inspection "
                "corps being operationalised."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 30: Government, in coordination with local authorities, promotes creation "
                "of community environmental inspection agents to ensure community participation "
                "in enforcement."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'o Governo... promove a criação de agentes de fiscalização comunitários.' "
                "Promotional/enabling language — not a binding operative requirement."
            ),
            "comments": "Multi-level enforcement coordination; promotional rather than mandatory.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 31: Government shall create economic or other incentives to encourage "
                "use of environmentally sound technologies and productive processes."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Compete ao Governo criar incentivos económicos ou de outra natureza com vista "
                "a encorajar a utilização de tecnologias e processos produtivos ambientalmente "
                "sãos.' Enabling power — incentives not yet specified or created in this Law."
            ),
            "comments": (
                "Economic instrument enabling power. Could support clean production/recycling "
                "incentives but not operative until Government acts."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 33: Government shall adopt the regulatory measures necessary for effective "
                "implementation of this Law — enabling complementary regulation on waste "
                "management, standards, and sectoral rules."
            ),
            "instrument_in_force": 0,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Cabe ao Governo adoptar as medidas regulamentares necessárias à efectivação "
                "da presente Lei.' Enabling power for subordinate regulation."
            ),
            "comments": (
                "Enabling power subsequently exercised via Decreto 94/2014 (urban solid waste), "
                "Decreto 16/2015 (plastic bags), and other decrees. Cross-reference only."
            ),
        },
    ],
}

ENGLISH_TRANSLATION = """ENVIRONMENTAL LAW (LEI DO AMBIENTE)
Law No. 20/97 of 1 October 1997

Source document: User-uploaded Boletim da República scan (Environmental_law_75d0.pdf), Series I, No. 40, 7 October 1997. Text reconstructed via OCR; some garbled passages on page 1 (end of Land Law No. 19/97) and page 6 (start of Electricity Law No. 21/97) are excluded from this translation.

Approved by the Assembly of the Republic on 31 July 1997.
Promulgated on 1 October 1997.
Published in the Boletim da República.
Enters into force sixty days after publication (Article 34).

---

CHAPTER I — GENERAL PROVISIONS

Article 1 (Definitions)

For the purposes of this Law:

1. Activity — any action of public or private initiative related to the use or exploitation of environmental components, application of technologies or productive processes, plans or programmes, legislative or regulatory acts that affect or may affect the environment.

2. Environment — the medium in which humans and other living beings interact with each other and with their surroundings, including: (a) air, light, land and water; (b) ecosystems, biodiversity, ecological relations, all organic and inorganic matter, and all socio-cultural and economic conditions affecting community life.

3. Environmental Defence Associations — collective persons whose purpose is the protection, conservation and valorisation of environmental components. These associations may have international, national, regional or local scope.

4. Environmental Audit — a systematic, documented, objective management-evaluation instrument assessing the functioning, organisation of a management system, and control and protection processes for the environment.

5. Environmental Impact Assessment — a preventive environmental management instrument consisting of the identification, prior analysis, and qualitative and quantitative assessment of the beneficial and harmful environmental effects of a proposed activity.

6. Biodiversity — the variety and variability among living organisms of all origins, including terrestrial, marine and other aquatic ecosystems and the ecological complexes of which they are part; encompasses diversity within species, between species, and of ecosystems.

7. Components of the Environment — the various elements that integrate the environment and whose interaction permits its equilibrium, including air, water, soil, subsoil, flora, fauna, and all socio-economic conditions.

8. Environmental Quality — the equilibrium and soundness of the environment, including the adequacy of its components to the needs of humans and other living beings and to community health; natural resources are also designated as such.

9. Environmental Degradation — adverse alteration of environmental characteristics, including pollution, desertification, erosion, deforestation.

10. Deforestation — destruction or indiscriminate felling of forests and woodlands without due replacement.

11. Sustainable Development — development based on environmental management that meets the needs of the present generation without compromising the ability of future generations to meet their own needs, maintaining environmental equilibrium.

12. Desertification — a process of soil degradation, natural or caused by removal of vegetation cover or predatory use, which under climatic conditions transforms land into desert.

13. Ecosystem — a dynamic complex of plant and animal communities and micro-organisms and their non-living environment, interacting as a functional unit.

14. Erosion — the wearing away of the soil surface by wind or water action, often intensified by human practices of vegetation removal.

15. Environmental Impact Study — the component of the environmental impact assessment process that technically and scientifically analyses the consequences of implementing development activities on the environment, environmental components, and socio-cultural conditions.

16. Environmental Management — the rational and sustainable handling and use of environmental components, including their reuse, recycling, protection and conservation.

17. Environmental Impact — any change in the environment, for better or worse, especially with effects on air, land, water and human health, resulting from human activities.

18. Environmental Legislation — all legal instruments governing environmental management and rational use of environmental components.

19. Sectoral Legislation — legal instruments governing a specific environmental component.

20. Environmental Quality Standards — permissible pollutant concentration levels prescribed by law for environmental components.

21. Environmental Expert Appraisal — investigation by a group of recognised specialists to assess the gravity and costs of environmental damage.

22. Pollution — deposition in the environment of substances or residues, regardless of form, as well as emission of light, sound and other forms of energy, in such manner and quantity as to negatively affect environmental quality, human health, and community well-being.

23. Hazardous Waste or Residues — substances or objects to be eliminated or which one is obliged by law to eliminate, containing risk characteristics: flammable, explosive, corrosive, toxic, infectious, radioactive, or presenting any other characteristic constituting danger to life or health.

24. Wetlands — swamp, marsh, peat or water areas, natural or artificial, permanent or temporary, standing or flowing, fresh, brackish or salt, including marine waters not exceeding six metres depth at low tide, sustaining vegetation or animal life requiring saturated aquatic conditions.

Article 2 (Purpose)

This Law has as its object the definition of the legal bases for correct use and management of the environment and its components, with a view to the materialisation of a system of sustainable development in the country.

Article 3 (Scope)

This Law applies to all public and private activities that directly or indirectly may pollute the environment or environmental components.

Article 4 (Fundamental principles)

Environmental management is based on the following fundamental principles:

a) the right of all citizens to an ecologically balanced environment conducive to their physical and mental health and well-being;

b) recognition and valorisation of the traditions and knowledge of local communities contributing to conservation and preservation of natural resources and the environment;

c) the precautionary principle, whereby environmental management must prioritise prevention systems to avoid significant or irreversible negative environmental impacts, regardless of scientific certainty;

d) a global, integrated vision of the environment as a set of interdependent natural and built ecosystems to be managed maintaining functional equilibrium without exceeding intrinsic limits;

e) broad participation of citizens as a crucial aspect of PNGA implementation;

f) equality, guaranteeing equal opportunities for access to and use of natural resources by men and women;

g) accountability — whoever pollutes or in any other way degrades the environment is always obliged to repair or compensate resulting damage;

h) international cooperation for harmonious solutions to environmental problems with transboundary and global dimensions.

CHAPTER II — ENVIRONMENTAL MANAGEMENT BODIES

Article 5 (National Environmental Management Programme)

It is the responsibility of the Government to elaborate and execute the National Environmental Management Programme (PNGA).

Article 6 (National Council for Sustainable Development)

1. To ensure effective and correct coordination and integration of environmental management principles and activities in the country's development process, the National Council for Sustainable Development (CONDES) is created.

2. CONDES is a consultative body of the Council of Ministers and also serves as a forum for public consultation on environmental matters.

3. CONDES is responsible for:
   a) pronouncing on sectoral policies related to natural resource management;
   b) issuing opinions on complementary legislation proposals under this Law, including proposals creating or revising sectoral legislation on natural resource management;
   c) pronouncing on proposals for ratification of international environmental conventions;
   d) elaborating proposals for financial or other incentives to stimulate economic agents toward environmentally sound procedures in daily resource use;
   e) proposing mechanisms to simplify and expedite licensing of activities related to natural resource use;
   f) formulating recommendations to ministers on relevant aspects of their respective natural resource management areas;
   g) serving as a forum for resolution of institutional disputes related to use and management of natural resources;
   h) exercising other functions assigned by this Law and other environmental legislation.

4. The composition and functioning of CONDES are regulated by Council of Ministers decree.

Article 7 (Local bodies)

At local level, services responsible for implementing this Law are created, which guarantee coordination of environmental action at that level and decentralisation in its execution, permitting adequate use of local initiatives and knowledge.

Article 8 (Public participation in environmental management)

It is the obligation of the Government to create adequate mechanisms to involve the various sectors of civil society, local communities, and in particular environmental defence associations, in elaboration of policies and legislation on management of the country's natural resources, as well as in development of activities implementing the National Environmental Management Programme.

CHAPTER III — ENVIRONMENTAL POLLUTION

Article 9 (Prohibition of pollution)

It is not permitted on national territory:
- production, deposit in soil and subsoil, discharge to water or atmosphere of any toxic and polluting substances;
- practice of activities accelerating erosion, desertification, deforestation or any other form of environmental degradation,
beyond legally established limits.

Import of hazardous waste or residues into national territory is expressly prohibited, except as established in specific legislation.

Article 10 (Environmental quality standards)

1. The Government must establish environmental quality standards to ensure sustainable use of the country's resources.

2. In defining environmental quality standards, norms and deadlines are also established for adaptation of agricultural and industrial processes, machines and means of transport, and appropriate devices or processes to retain or neutralise polluting substances.

CHAPTER IV — SPECIAL ENVIRONMENTAL PROTECTION MEASURES

Article 11 (Protection of environmental heritage)

The Government must ensure that environmental heritage, especially cultural-historical heritage, is subject to permanent defence and valorisation measures, with adequate involvement of communities, in particular environmental defence associations.

Article 12 (Biodiversity protection)

1. All activities threatening conservation, reproduction, quality and quantity of biological resources, especially those threatened with extinction, are prohibited.

2. The Government must ensure adequate measures for:
   a) maintenance and regeneration of animal species, recovery of damaged habitats and creation of new habitats, especially controlling activities or use of substances liable to harm fauna and their habitats;
   b) special protection of plant species threatened with extinction or botanical specimens which, by genetic potential, size, age, rarity, scientific and cultural value, require it.

Article 13 (Environmental protection areas)

1. To ensure protection and preservation of environmental components and maintenance and improvement of ecosystems of recognised ecological and socio-economic value, the Government establishes duly signposted environmental protection areas.

2. Protected areas may have national, regional, local or international scope, covering terrestrial areas, lacustrine, fluvial or maritime waters and other distinct natural zones.

3. Environmental protection areas are subject to classification, conservation and surveillance measures, always considering preservation of biodiversity and social, economic, cultural, scientific and landscape values.

4. Such measures must include indication of permitted or prohibited activities within protected areas and their surroundings, and the role of local communities in managing these areas.

Article 14 (Infrastructure installation)

1. Installation of housing or other infrastructure that, by dimension, nature or location, causes significant negative environmental impact is prohibited — including deposit of waste or used materials. This prohibition applies especially to coastal zones, erosion- or desertification-threatened zones, wetlands, environmental protection areas, and other ecologically sensitive zones.

2. Norms for infrastructure installation in the areas referred above are established by regulation. Installation of structures near roads, railways, dams, ports and airports is also regulated so as not to prejudice their functioning, expansion possibilities, or landscape harmony.

CHAPTER V — PREVENTION OF ENVIRONMENTAL DAMAGE

Article 15 (Environmental licensing)

1. Licensing and registration of activities that, by nature, location or scale, are susceptible of causing significant environmental impacts shall be conducted in accordance with a regime established by the Government through specific regulation.

2. Issuance of the environmental licence is based on an environmental impact assessment of the proposed activity and precedes issuance of any other licences legally required in each case.

Article 16 (Environmental impact assessment)

Environmental impact assessment is based on an environmental impact study conducted by entities accredited by the Government.

Article 17 (Minimum content of environmental impact study)

The environmental impact study shall comprise at minimum:
a) non-technical summary of the project;
b) description of the activity to be developed;
c) environmental situation at the installation site;
d) modifications the activity causes to existing environmental components at the site;
e) measures foreseen to suppress or reduce negative effects on environmental quality;
f) systems foreseen for control and monitoring of the activity.

Article 18 (Environmental audits)

1. All activities in operation at the date of entry into force of this Law without application of appropriate technologies and processes that result or may result in environmental damage are subject to environmental audits.

2. Costs of repairing environmental damage eventually found by audits are the responsibility of the entrepreneurs. Audit norms are established by regulation.

CHAPTER VI — RIGHTS AND DUTIES OF CITIZENS

Article 19 (Right to information)

All persons have the right of access to information related to environmental management in the country, without prejudice to legally protected third-party rights.

Article 20 (Right to education)

To ensure correct environmental management and necessary community participation, the Government must create, in collaboration with social communication bodies, mechanisms and programmes for formal and informal environmental education.

Article 21 (Right of access to justice)

1. Any citizen who considers their rights under this Law violated, or who considers there is threat of violation, may resort to judicial instances to obtain restoration of rights or prevention of violation.

2. Any person who, due to violation of environmental legislation, suffers personal injury or property damage, including loss of harvests or profits, may judicially prosecute the author of the damage or offence and demand repair or compensation.

3. Legal actions referred to in paragraphs 1 and 2 follow appropriate procedural terms.

Article 22 (Embargos)

Those who consider themselves offended in their right to an ecologically balanced environment may request immediate suspension of the offending activity, following administrative embargo proceedings or other appropriate procedural means.

Article 23 (Obligation to report infractions)

Any person who verifies infractions of this Law or any other environmental legislation, or who reasonably presumes such infractions are imminent, has the obligation to inform the nearest police authorities or other administrative agents of the fact.

Article 24 (Obligation of responsible resource use)

All persons have the obligation to use natural resources responsibly and sustainably, wherever they are, as well as the duty to encourage others to do the same.

CHAPTER VII — RESPONSIBILITIES, INFRACTIONS AND SANCTIONS

Article 25 (Civil liability insurance)

All persons exercising activities involving high risk of environmental degradation, as classified under environmental impact assessment legislation, must secure civil liability insurance.

Article 26 (Strict liability)

1. All who, regardless of fault, cause significant environmental damage or paralyse economic activities temporarily or definitively as a result of especially dangerous activities, are obliged to pay compensation to those harmed.

2. It is the responsibility of the Government to supervise assessment of the gravity of damage and fixation of its value, conducted through environmental expert appraisal.

3. Whenever circumstances require, the State takes necessary measures to prevent, contain or eliminate any serious environmental damage, while retaining the right of recourse for costs borne.

Article 27 (Environmental crimes and contraventions)

Environmental crimes and contraventions are provided for in specific legislation. The Public Ministry is responsible for defending environmental values protected by this Law, without prejudice to the legitimacy of harmed parties to bring actions referred therein.

CHAPTER VIII — ENVIRONMENTAL INSPECTION

Article 28 (Environmental inspection agents)

It is the responsibility of the Government to create, under terms to be regulated, a corps of environmental inspection agents competent to ensure implementation of environmental legislation and to take necessary preventive measures against violation of its provisions.

Article 29 (Duty of collaboration)

All persons in charge of an activity or place subject to inspection must collaborate with inspection agents in carrying out their activities.

Article 30 (Community participation)

To guarantee necessary participation of local communities and adequate use of their knowledge and human resources, the Government, in coordination with local authorities, promotes creation of community environmental inspection agents.

CHAPTER IX — FINAL PROVISIONS

Article 31 (Incentives)

It is the responsibility of the Government to create economic or other incentives to encourage use of environmentally sound technologies and productive processes.

Article 32 (Sectoral legislation)

1. Existing legislation governing environmental components must be adjusted to the provisions of this Law.

2. Regulation of this Law is the responsibility of the Government, which shall fix deadlines for already authorised projects and ongoing enterprises contrary to its provisions to be adjusted thereto.

Article 33 (Complementary legislation)

The Government shall adopt the regulatory measures necessary for effective implementation of this Law.

Article 34 (Entry into force)

This Law enters into force sixty days after its publication in the Boletim da República.

---

Note on source document: The uploaded PDF is a Boletim da República compilation. Pages 2–5 contain Lei 20/97 (Environmental Law). Page 1 contains garbled text from the end of Lei 19/97 (Land Law). Page 6 begins Lei 21/97 (Electricity Law). This translation covers Lei 20/97 only. OCR uncertainties in article numbering and phrasing were resolved against standard published versions where garbled."""


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
        "A": 38, "B": 18, "C": 8, "D": 45, "E": 8, "F": 12, "G": 8, "H": 35,
        "I": 8, "J": 35, "K": 8, "L": 30, "M": 8, "N": 12, "O": 10,
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
    meta.append(["Source", "User upload: Environmenmtal_law_75d0.pdf (Boletim da República scan)"])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Total instrument rows", len(rows)])
    meta.append([
        "Note",
        "Columns O and V use Excel formulas. "
        "policy_score = AVERAGE(G,I,K,M,P); instrument_score = AVERAGE(P,T). "
        "Only Lei 20/97 coded; pages 1 and 6 of upload contain other laws.",
    ])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_translation_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading("Lei do Ambiente (Law No. 20/97) — English Translation", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English translation of Mozambique's Environmental Law (Lei n.º 20/97, de 1 de Outubro), "
        "based on OCR of the user-uploaded Boletim da República scan. "
        "This is the sole policy covered by the accompanying 4P Index coding table."
    )

    p = doc.add_paragraph()
    p.add_run("Portuguese title: ").bold = True
    p.add_run("Lei do Ambiente (Lei n.º 20/97, de 1 de Outubro)")

    p2 = doc.add_paragraph()
    p2.add_run("Source: ").bold = True
    p2.add_run("User-uploaded document (Environmenmtal_law_75d0.pdf); policy URL not available.")

    doc.add_heading("English translation", level=2)

    for block in ENGLISH_TRANSLATION.strip().split("\n\n"):
        text = block.strip()
        if not text:
            continue
        if text.startswith("---"):
            continue
        if text.isupper() and len(text) < 80 and not text.startswith("Article"):
            doc.add_heading(text.title(), level=3)
        elif text.startswith("CHAPTER") or text.startswith("Article "):
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
