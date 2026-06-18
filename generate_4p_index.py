#!/usr/bin/env python3
"""Generate Mozambique 4P Index Excel and English-translated policy Word document."""

from __future__ import annotations

import re
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from generate_mozambique_plastic_policies_bilingual_docx import POLICIES

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index.xlsx")
WORD_OUT = Path("/workspace/Mozambique_Policies_English_Translations.docx")

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

# Each policy: policy-level dict + list of instrument dicts
CODING = [
    {
        "policy_name": "Política Nacional do Ambiente (Resolução n.º 5/95)",
        "policy_url": "not available",
        "policy_year": 1995,
        "policy_objective": "Establish national environmental policy principles for sustainable development, pollution prevention, public participation, and integration of environmental considerations into socio-economic planning.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "Council of Ministers resolution approving a national environmental policy (strategy-level instrument, not legislation).",
        "policy_integration": 1,
        "policy_sectors_list": "environment, industry, agriculture, municipalities, fisheries, tourism, water, waste management",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Establishes polluter-pays principle and framework for environmental management across sectors (Section 2.2), enabling coordinated pollution prevention including plastic waste pressures.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Assigns government responsibility for environmental policy implementation; no explicit fines in policy text itself.",
                "comments": "Full-text PDF not publicly available; coded from verified legal-framework references on sibmoz.gov.mz.",
            },
            {
                "instrument_type": 0.40,
                "instrument_lifecycle_stage": "Consumption",
                "instrument_description": "Commits to environmental education and public awareness (Section 2.3.5) to reduce pollution including litter and improper waste disposal.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Strategy calls for education programmes at all schooling levels; no enforcement mechanism in policy text.",
                "comments": "Information/voluntary instrument at policy level.",
            },
        ],
    },
    {
        "policy_name": "Lei do Ambiente (Lei n.º 20/97)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/1547563349-Lei-do-Ambiente.pdf",
        "policy_year": 1997,
        "policy_objective": "Define legal bases for correct use and management of the environment, prohibit unauthorized pollutant releases, and establish environmental licensing and sustainable development framework applicable to plastic pollution.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 1,
        "policy_type_justification": "Environmental Law enacted by the Assembly of the Republic (Parliament) on 1 October 1997.",
        "policy_integration": 1,
        "policy_sectors_list": "environment, industry, municipalities, agriculture, mining, energy, fisheries, tourism, waste management",
        "policy_circularity": 1,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Art. 7: Prohibits release of pollutants into the environment without authorization; applies to industrial activities producing plastic goods and packaging.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Establishes environmental licensing; penalties for violations; monitoring through environmental authorities.",
                "comments": "Uploaded user PDF (Environmental_law_13c6.pdf) appears to be image-based Lei 20/97; coding aligned with official sibmoz text references.",
            },
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Art. 33: Authorises executive regulation of waste management, enabling Decreto 94/2014 and sectoral plastic-waste rules.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Delegates waste regulation to Council of Ministers; exercised through Decreto 94/2014 and related decrees.",
                "comments": "Enabling power exercised via subordinate regulations.",
            },
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Creates CONDES consultative body and framework for public participation in environmental decisions affecting waste and pollution.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Institutional coordination mechanism; no direct penalties in this provision.",
                "comments": "Governance/coordination instrument.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre a Gestão e Controlo do Saco de Plástico (Decreto n.º 16/2015)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/Regulamento-Sobre-a-Gestao-e-Controlo-do-Saco-de-Plastico.pdf",
        "policy_year": 2015,
        "policy_objective": "Regulate production, import, marketing and use of plastic bags to reduce health, infrastructure, biodiversity and environmental impacts of non-biodegradable plastic bags.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree (Decreto 16/2015) approved by Council of Ministers, 5 August 2015.",
        "policy_integration": 0.50,
        "policy_sectors_list": "production, retail, industry, municipalities, waste management, packaging",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal",
        "policy_budget": 0.5,
        "policy_budget_text": "Art. 8: Fines allocated 40% State Budget, 30% Environment Fund, 30% inspecting entity.",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Art. 4(1)(a): Prohibits production, import, wholesale/retail marketing of plastic bags with thickness below 30 micrometres (exceptions for food weighing bags and municipal waste bags).",
                "instrument_in_force": 1,
                "instrument_implementation": 1.0,
                "instrument_implementation_text": "Art. 6-7: MAAP, Finance, Industry, INAE and municipalities assigned; Art. 7 fines 40-80 minimum wages; monitoring and inspection duties explicit.",
                "comments": "Core plastic-bag ban instrument; phased entry for import/production bans per Art. 3(2).",
            },
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Consumption",
                "instrument_description": "Art. 4(1)(b): Prohibits free distribution of plastic bags at commercial premises; Art. 5(4): mandatory separate price display for bags.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Fines for free distribution (25 min. wages) and failure to display price (30 min. wages); INAE inspection.",
                "comments": "",
            },
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Art. 4(1)(c) & Art. 5: Restricts bags with >40% recycled content in food retail; mandates labelling (company, thickness, recycled %, plastic symbol) per NM 596.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Labelling mandatory; fine 60 min. wages for violating recycled-content rule in food stores.",
                "comments": "Regulatory standard on bag composition and labelling.",
            },
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Art. 6: Assigns national–municipal coordination for monitoring compliance and promoting alternatives to plastic bags.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Municipalities and local state organs must ensure compliance within their jurisdiction.",
                "comments": "Multi-level governance coordination.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens (Decreto n.º 79/2017)",
        "policy_url": "not available",
        "policy_year": 2017,
        "policy_objective": "Establish extended producer responsibility for packaging including internal management systems, packaging environmental fee (TAE), and packaging standardisation to reduce pollution and fund waste management.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree approved by Council of Ministers, 28 December 2017 (Boletim da República n.º 203).",
        "policy_integration": 0.75,
        "policy_sectors_list": "packaging, industry, retail, municipalities, waste management, recycling, chemicals",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal",
        "policy_budget": 0.5,
        "policy_budget_text": "Creates Taxa Ambiental sobre a Embalagem (TAE) payable by producers/importers; Diploma Ministerial 26/2025 later defines TAE formula.",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Art. 6-8: Assigns physical/financial responsibility to producers and importers for packaging waste management and take-back obligations.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Responsibilities assigned to producers, MIC, MEF and municipalities; implementation reported weak without TAE operationalisation.",
                "comments": "Full decree text not on public .gov.mz PDF; coded from CIPMoz legal analysis and World Bank plastic flows study citing Arts. 6-8.",
            },
            {
                "instrument_type": 0.60,
                "instrument_lifecycle_stage": "Recycling",
                "instrument_description": "Art. 13-14: Establishes TAE (packaging environmental fee) with four criteria (returnability, decomposition time/impact, treatment cost, eco-design).",
                "instrument_in_force": 0,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "TAE requires supplementary ministerial diploma for formula; DM 26/2025 approved March 2025 — enabling power partially exercised.",
                "comments": "Score in_force=0 until TAE fully operational; DM 26/2025 may change status — verify separately.",
            },
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Art. 17: Producers must preferentially use recyclable materials in packaging production.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Mandatory language but limited explicit enforcement provisions in cited articles.",
                "comments": "",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre a Gestão de Resíduos Sólidos Urbanos (Decreto n.º 94/2014)",
        "policy_url": "https://gpa.co.mz/wp-content/uploads/2023/08/REGULAMENTO-DA-LEI-DA-CONCORRENCIA_2014-1.pdf",
        "policy_year": 2014,
        "policy_objective": "Establish national rules for urban solid waste management including segregation, collection, treatment, valorisation and landfill disposal, covering plastic fractions in municipal waste streams.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree (Decreto 94/2014) approved by Council of Ministers, 31 December 2014.",
        "policy_integration": 0.75,
        "policy_sectors_list": "waste management, municipalities, industry, retail, recycling, public health, environment",
        "policy_circularity": 1,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal, environmental leakage",
        "policy_budget": 0.5,
        "policy_budget_text": "Art. 5(2)(e): Municipalities may fix fees/taxes for collection, transport, treatment and disposal services.",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Art. 5-6: Assigns national–municipal–district coordination; municipalities must ensure waste not dumped in beaches, sea, water bodies (plastic litter pathway).",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Penalties for municipal managers; monitoring by environment ministry; municipal obligations explicit.",
                "comments": "Multi-level governance; directly addresses marine plastic leakage via Art. 6(a).",
            },
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Art. 8: Requires Integrated Urban Solid Waste Management Plans (PGIRSU) for all public/private waste management entities.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Planning obligation; licensing linked to plans; ministry oversight.",
                "comments": "Governance/planning instrument.",
            },
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Recycling",
                "instrument_description": "Art. 4(d): Waste hierarchy — prevention, reuse, recycling, other recovery, disposal; Art. 7: segregation at source for recycling including plastics.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Hierarchy mandatory; selective collection promoted; enforcement via municipal regulations and penalties.",
                "comments": "",
            },
            {
                "instrument_type": 0.60,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Art. 5(2)(e): Municipalities authorised to set fees for waste collection/treatment/disposal services.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Fee-setting power delegated to municipalities; not unconditional national levy.",
                "comments": "Economic instrument at municipal level.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre a Gestão de Resíduos Perigosos (Decreto n.º 83/2014)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/Aprova-o-Regulamento-sobre-Gestao-de-Residuos-perigosos-e-os-respectivos-anexos.pdf",
        "policy_year": 2014,
        "policy_objective": "Regulate safe management of hazardous wastes including contaminated plastic packaging from pesticides and chemicals.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree approved by Council of Ministers, 31 December 2014.",
        "policy_integration": 0.50,
        "policy_sectors_list": "industry, agriculture, chemicals, waste management, municipalities, packaging",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Disposal",
                "instrument_description": "Prohibits recycling of plastic packaging contaminated by agro-toxins/chemicals except below threshold in Annex IX; bans import of empty contaminated packaging.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Explicit prohibitions; licensing for recycling; fines; extended producer responsibility principle in Art. 4(g).",
                "comments": "Directly addresses plastic packaging waste streams.",
            },
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Recycling",
                "instrument_description": "Requires importers/marketers of products whose packaging becomes hazardous waste to operate take-back and collection systems.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Take-back obligation on market placers; treatment and disposal responsibility assigned.",
                "comments": "EPR-type provision for hazardous packaging.",
            },
        ],
    },
    {
        "policy_name": "Programa Nacional de Gestão Sustentável de Resíduos (ValoRe)",
        "policy_url": "not available",
        "policy_year": 2025,
        "policy_objective": "Circular-economy programme to build sustainable waste-treatment infrastructure, strengthen recycling value chains, and pilot municipal waste management.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "Government programme (not legislation); aspirational infrastructure and pilot commitments.",
        "policy_integration": 0.50,
        "policy_sectors_list": "waste management, municipalities, industry, recycling, environment",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "recycling, disposal, consumption",
        "policy_budget": 0.5,
        "policy_budget_text": "Programme mobilises ~€18.4 million for sustainable waste management infrastructure (government announcement).",
        "instruments": [
            {
                "instrument_type": 0.80,
                "instrument_lifecycle_stage": "Recycling",
                "instrument_description": "Plans construction of sustainable waste-treatment infrastructure and recycling value chains in Nampula, Nacala and Pemba pilot municipalities.",
                "instrument_in_force": 0,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Programme commitment; no standalone official PDF; implementation via projects.",
                "comments": "No direct PDF; coded from government programme description. May be in_force=1 when projects operational.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre Padrões de Qualidade Ambiental e de Emissão de Efluentes (Decreto n.º 18/2004, alterado pelo Decreto n.º 67/2010)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/Regulamento-sobre-Padroes-de-Qualidade-Ambiental-e-de-Emissao-de-Efluentes.pdf",
        "policy_year": 2010,
        "policy_objective": "Set national environmental quality and effluent/emission standards limiting pollutant releases from industrial sources including plastic production.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree (amended 2010 by Decreto 67/2010).",
        "policy_integration": 0.50,
        "policy_sectors_list": "industry, environment, energy, transport, chemicals",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Establishes mandatory emission/effluent limits for air, water and mobile sources applicable to industrial plastic manufacturing facilities.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Standards enforceable through environmental licensing; penalties for non-compliance under environmental law.",
                "comments": "Indirect plastic relevance via industrial production controls.",
            },
        ],
    },
    {
        "policy_name": "Regulamento para Prevenção da Poluição e Proteção do Ambiente Marinho e Costeiro (Decreto n.º 45/2006)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/1547623391-Regulamento-sobre-a-Prevencao-da-Poluicao-ea-Proteccao-Ambiente-Marinho-e-Costeiro-Decreto-45-barra-2006.pdf",
        "policy_year": 2006,
        "policy_objective": "Prevent, control and combat marine pollution from ships and land-based sources including plastic marine litter.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree, Council of Ministers, 30 November 2006.",
        "policy_integration": 0.50,
        "policy_sectors_list": "fisheries, maritime, ports, environment, tourism, municipalities",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 1.0,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Requires measures to prevent marine pollution from land-based sources and ships; packaging requirements for hazardous substances discharged to sea.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Compensation for pollution required; enforcement via maritime authorities.",
                "comments": "Addresses marine plastic leakage pathway.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre a Actividade de Fiscalização Ambiental (Decreto n.º 51/2024)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2025/11/Regulamento-de-fiscalizac%CC%A7a%CC%83o-ambiental.pdf",
        "policy_year": 2024,
        "policy_objective": "Establish rules for environmental inspection and enforcement of environmental legislation including plastic and waste regulations.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree approved 17 July 2024; revokes Decreto 11/2006.",
        "policy_integration": 0.75,
        "policy_sectors_list": "environment, industry, municipalities, waste management, fisheries, mining",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Defines types of environmental inspection, competencies and procedures for enforcing waste and pollution regulations nationally and subnationally.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Inspection powers, penalties reference, monitoring duties for environmental authorities.",
                "comments": "Cross-cutting enforcement framework for plastic-related decrees.",
            },
        ],
    },
    {
        "policy_name": "Regulamento da Avaliação do Impacto Ambiental (Decreto n.º 54/2015)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/1547556163-Decreto-54-2015-Regulamento-da-AIA.pdf",
        "policy_year": 2015,
        "policy_objective": "Regulate environmental impact assessment and licensing for projects that may cause significant environmental harm including plastic production facilities.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree, Council of Ministers, December 2015.",
        "policy_integration": 0.75,
        "policy_sectors_list": "industry, infrastructure, mining, energy, tourism, environment, municipalities",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Mandatory EIA and environmental licensing for Category A/A+ projects including large industrial installations; mitigation hierarchy required.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Licensing authority designated; inspections; biodiversity offsets where required; public participation.",
                "comments": "Applies to new plastic/polymer production projects.",
            },
        ],
    },
    {
        "policy_name": "Regulamento sobre o Processo de Auditoria Ambiental (Decreto n.º 45/2024)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2025/11/Regulamento-sobre-o-Processo-de-Auditoria-Ambiental.pdf",
        "policy_year": 2024,
        "policy_objective": "Establish environmental audit requirements for activities impacting the environment including waste-generating industries.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree, Council of Ministers, 26 June 2024.",
        "policy_integration": 0.50,
        "policy_sectors_list": "industry, environment, municipalities, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Production",
                "instrument_description": "Requires environmental audits and corrective action plans; RAA must be submitted within 15 days of audit.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.75,
                "instrument_implementation_text": "Mandatory reporting; corrective action plan within 30 days for non-conformities; authority oversight.",
                "comments": "Monitoring/compliance instrument.",
            },
        ],
    },
    {
        "policy_name": "Política e Estratégia do Mar – POLMAR (Resolução n.º 39/2017)",
        "policy_url": "https://www.proazul.gov.mz/wp-content/uploads/2023/10/POLITICA-E-ESTRATEGIA-DO-MAR-POLMAR-1.pdf",
        "policy_year": 2017,
        "policy_objective": "Guide sustainable integrated use of maritime and coastal spaces and blue-economy development including marine pollution control.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "Council of Ministers resolution approving sea policy and strategy (POLMAR), 14 September 2017.",
        "policy_integration": 0.75,
        "policy_sectors_list": "fisheries, maritime, tourism, ports, environment, municipalities, industry",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "POLMAR establishes integrated governance framework for maritime/coastal economic activities balancing development and marine environmental protection.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Strategic coordination instrument; implementation through sectoral plans and decrees.",
                "comments": "",
            },
        ],
    },
    {
        "policy_name": "Lei do Mar (Lei n.º 20/2019)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/Lei-do-Mar.pdf",
        "policy_year": 2019,
        "policy_objective": "Legal regime for sovereignty, jurisdiction, exploitation of marine resources and use of maritime public domain relevant to marine plastic pollution control.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 1,
        "policy_type_justification": "Law enacted by Assembly of the Republic, 8 November 2019.",
        "policy_integration": 0.50,
        "policy_sectors_list": "maritime, fisheries, mining, transport, environment, tourism",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Establishes legal framework for maritime jurisdiction enabling regulation of activities causing marine pollution including litter.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "State sovereignty provisions; exercised through Decreto 45/2006 and POEM.",
                "comments": "Framework law enabling marine plastic controls.",
            },
        ],
    },
    {
        "policy_name": "Regulamento de Gestão e Ordenamento da Zona Costeira e das Praias (Decreto n.º 97/2020)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/06/Regulamento-Zonas-costeiras.pdf",
        "policy_year": 2020,
        "policy_objective": "Define integrated sustainable management of coastal zones and beaches including controls on activities affecting coastal pollution.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree, 4 October 2020.",
        "policy_integration": 0.50,
        "policy_sectors_list": "municipalities, tourism, fisheries, environment, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Coastal zone management and licensing rules to prevent degradation of beaches and coastal ecosystems from waste and development pressures.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Municipal and national coordination; licensing requirements.",
                "comments": "Supports prevention of coastal plastic litter.",
            },
        ],
    },
    {
        "policy_name": "Plano Nacional de Ordenamento do Espaço Marítimo – POEM (Resolução n.º 63/2024)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2025/11/Resolucao-n.o-63_2024-que-aprova-o-Plano-de-Ordenamento-do-Espaco-Maritimo.pdf",
        "policy_year": 2024,
        "policy_objective": "Approve national maritime spatial plan coordinating sea uses and minimising environmental impacts including marine pollution.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.50,
        "policy_type_justification": "Council of Ministers resolution approving national spatial plan (POEM), November 2024.",
        "policy_integration": 0.75,
        "policy_sectors_list": "maritime, fisheries, energy, transport, environment, tourism, mining",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "POEM zones maritime space to minimise environmental impacts and coordinate competing uses affecting marine ecosystems.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Binding spatial plan; coordination across sectors and levels of government.",
                "comments": "Planning/coordination instrument.",
            },
        ],
    },
    {
        "policy_name": "Estratégia de Desenvolvimento da Economia Azul – EDEA (Resolução n.º 53/2024)",
        "policy_url": "https://www.proazul.gov.mz/wp-content/uploads/2024/11/EDEA-ESTRATEGIA-DE-DESENVOLVIMENTO-DA-ECONOMIA-AZUL-_FINAL_TYPO_PRINT_WEB.pdf",
        "policy_year": 2024,
        "policy_objective": "National blue-economy strategy (2024–2033) promoting sustainable marine resource use, circular economy, waste valorisation, and reducing plastic production, import and use.",
        "policy_target": 1,
        "policy_target_text": "Strategic objective (c): 'Reduzir a produção, importação e uso de plásticos, aumentando a responsabilidade alargada dos sectores da indústria e do comércio e os apoios e incentivos à reciclagem e indústria de remanufactura.' Action plan tables include quantitative targets on illegal fishing practices (50% reduction by 2027; 90% by 2040) and waste/recycling indicators.",
        "policy_type": 0.50,
        "policy_type_justification": "Council of Ministers resolution approving national strategy with quantifiable action-plan targets (2024–2033).",
        "policy_integration": 1,
        "policy_sectors_list": "fisheries, maritime, industry, tourism, environment, municipalities, recycling, waste management, energy",
        "policy_circularity": 1,
        "policy_lifecycle_phases_list": "production, consumption, recycling, disposal, environmental leakage",
        "policy_budget": 0.5,
        "policy_budget_text": "Strategy includes budget lines per strategic action (e.g., 901,355 MT; 3,465,152 MT for fisheries actions); circular-economy/waste actions budgeted in action matrix.",
        "instruments": [
            {
                "instrument_type": 0.40,
                "instrument_lifecycle_stage": "Consumption",
                "instrument_description": "Strategic objective: reduce plastic production, import and use; strengthen EPR and recycling/remanufacturing incentives (Pilar 3, circular economy).",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Strategy commitment with action matrix; implementation depends on sectoral regulations.",
                "comments": "Aspirational with action-plan indicators; plastic reduction objective qualitative at strategic level.",
            },
            {
                "instrument_type": 0.80,
                "instrument_lifecycle_stage": "Recycling",
                "instrument_description": "Action viii (Pilar 3): Establish circular-economy projects improving waste management with community involvement; invest in waste treatment and valorisation technology.",
                "instrument_in_force": 0,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Planned infrastructure investment; not yet binding regulation.",
                "comments": "Infrastructure/planning instrument in strategy.",
            },
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Coastal pillar: prevent plastic pollution through reallocation, remanufacture/recycling and reintroduction into economic circuit; discourage import/production of hard-to-degrade products.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Strategic direction; operationalised through existing decrees.",
                "comments": "References 8% plastic fraction in KaNyaka municipal waste study.",
            },
        ],
    },
    {
        "policy_name": "Estratégia de Gestão Integrada das Zonas Costeiras (2016–2025)",
        "policy_url": "not available",
        "policy_year": 2016,
        "policy_objective": "Integrated coastal zone management addressing erosion, climate vulnerability and sustainable coastal resource use linked to marine pollution pressures.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "National integrated coastal management strategy/action plan 2016–2025 (listed on sibmoz.gov.mz).",
        "policy_integration": 0.50,
        "policy_sectors_list": "municipalities, fisheries, tourism, environment, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "EGIZC action plan for integrated coastal management to reduce coastal pollution and erosion pressures affecting plastic litter pathways.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Strategic coordination; no standalone PDF available for verification.",
                "comments": "Referenced in EDEA and POEM documents; direct PDF not found.",
            },
        ],
    },
    {
        "policy_name": "Estratégia Nacional de Gestão e Conservação dos Corais (2022–2032; Resolução n.º 51/2022)",
        "policy_url": "https://faolex.fao.org/docs/pdf/moz214814.pdf",
        "policy_year": 2022,
        "policy_objective": "Protect coral reef ecological integrity by reducing anthropogenic degradation including marine litter and pollution pressures.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "Council of Ministers resolution approving coral reef strategy 2022–2032, 28 December 2022.",
        "policy_integration": 0.50,
        "policy_sectors_list": "fisheries, environment, tourism, maritime, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "environmental leakage, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.40,
                "instrument_lifecycle_stage": "Environmental leakage",
                "instrument_description": "Strategy actions on reducing anthropogenic reef degradation including marine litter, unsustainable fishing and tourism practices.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Action plan with monitoring indicators; fisheries enforcement components.",
                "comments": "Approved Resolução 51/2022; full strategy in Boletim annex.",
            },
        ],
    },
    {
        "policy_name": "Política de Saúde e Estratégia de Implementação (Resolução n.º 13/2021)",
        "policy_url": "https://sisma.misau.gov.mz/docs/Politica%20Nacional%20de%20Saude.pdf",
        "policy_year": 2021,
        "policy_objective": "National health policy including expansion of basic sanitation and management of solid waste pressures from urbanisation affecting public health.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.25,
        "policy_type_justification": "Council of Ministers resolution approving national health policy, 2021.",
        "policy_integration": 0.50,
        "policy_sectors_list": "health, municipalities, environment, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "disposal, consumption",
        "policy_budget": 0.5,
        "policy_budget_text": "Health sector budget allocations for sanitation and waste-related programmes referenced in implementation strategy.",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Health policy commitments to expand basic sanitation and address solid waste management pressures in rapidly urbanising areas.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.25,
                "instrument_implementation_text": "Sectoral policy; implementation via MISAU and municipal health programmes.",
                "comments": "Indirect plastic relevance via municipal solid waste/sanitation.",
            },
        ],
    },
    {
        "policy_name": "Plano Nacional de Desenvolvimento Territorial (Resolução n.º 7/2021)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/06/Plano-Nacional-de-Desenvolvimento-Territorial-1.pdf",
        "policy_year": 2021,
        "policy_objective": "Guide sustainable territorial development and urban growth with environmental balance at national and subnational levels affecting waste infrastructure planning.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.50,
        "policy_type_justification": "Council of Ministers resolution approving national territorial development plan and action plan, 28 December 2021.",
        "policy_integration": 0.75,
        "policy_sectors_list": "urban planning, municipalities, environment, infrastructure, industry, tourism, agriculture",
        "policy_circularity": 0.75,
        "policy_lifecycle_phases_list": "production, consumption, disposal, environmental leakage",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "PNDT coordinates territorial planning instruments at national, provincial, district and municipal levels including environmental balance and urban waste infrastructure.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Binding national plan; multi-level planning obligations.",
                "comments": "Spatial planning affects placement of waste/recycling infrastructure.",
            },
        ],
    },
    {
        "policy_name": "Lei do Ordenamento do Território (Lei n.º 19/2007)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2022/01/Lei-do-ordenamento-territorial.pdf",
        "policy_year": 2007,
        "policy_objective": "Territorial planning law promoting sustainable natural resource use and environmental balance in national development including waste facility siting.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 1,
        "policy_type_justification": "Law enacted by Assembly of the Republic, 18 July 2007.",
        "policy_integration": 0.75,
        "policy_sectors_list": "urban planning, municipalities, environment, infrastructure, industry, agriculture, tourism",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Establishes territorial planning framework requiring environmental balance in land-use decisions affecting waste management zones.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Planning instruments at all government levels; exercised via Decreto 23/2008.",
                "comments": "",
            },
        ],
    },
    {
        "policy_name": "Regulamento da Lei de Ordenamento do Território (Decreto n.º 23/2008)",
        "policy_url": "https://sibmoz.gov.mz/content/uploads/2025/11/Decreto-n.o-23_2008_Regulamento-da-Lei-de-Ordenamento-do-Territorio.pdf",
        "policy_year": 2008,
        "policy_objective": "Regulate territorial occupation and sustainable land use through national, provincial, district and municipal planning instruments coordinating environmental development.",
        "policy_target": 0,
        "policy_target_text": "",
        "policy_type": 0.75,
        "policy_type_justification": "Executive decree regulating Territorial Planning Law, 2008.",
        "policy_integration": 0.75,
        "policy_sectors_list": "municipalities, urban planning, environment, industry, infrastructure, waste management",
        "policy_circularity": 0.50,
        "policy_lifecycle_phases_list": "production, disposal",
        "policy_budget": 0,
        "policy_budget_text": "",
        "instruments": [
            {
                "instrument_type": 0.20,
                "instrument_lifecycle_stage": "Waste management",
                "instrument_description": "Defines planning instruments (POT, PDU, etc.) at national to municipal levels integrating economic, social and environmental development including waste infrastructure.",
                "instrument_in_force": 1,
                "instrument_implementation": 0.50,
                "instrument_implementation_text": "Mandatory planning procedures; municipal competencies for land-use regulation.",
                "comments": "Subnational planning coordination instrument.",
            },
        ],
    },
]

# English translations for Word document (summary translations of key provisions)
TRANSLATIONS = {
    "Regulamento sobre a Gestão e Controlo do Saco de Plástico (Decreto n.º 16/2015)": """DECREE NO. 16/2015 — REGULATION ON THE MANAGEMENT AND CONTROL OF PLASTIC BAGS

Article 1 (Definitions): Defines plastic, plastic bag, biodegradable, recycled material, micrometre, and related terms.

Article 2 (Purpose): Establishes norms for production, import, marketing and use of plastic bags to reduce negative impacts on human health and the environment.

Article 4 (Prohibitions):
1. It is prohibited to: (a) produce, import, or market plastic bags with thickness below 30 micrometres; (b) distribute plastic bags free of charge at commercial premises; (c) market bags with over 40% recycled content in food retail establishments.
2. Exceptions: bags for weighing food products; bags for municipal solid waste conditioning; export free-zone production.

Article 5 (Production and marketing): Production and import must comply with Mozambican Standard NM 596. Mandatory labelling with company name, address, volume, material, plastic symbol, thickness and recycled percentage. Separate price display for bags required.

Article 6 (Competencies): Environment Ministry monitors compliance; Finance Ministry inspects imports; Industry Ministry licenses producers; INAE inspects production and use; municipalities ensure local compliance.

Article 7 (Offences and penalties): Fines from 25 to 80 minimum wages depending on violation (production, import, sale, free distribution, labelling). Repeat offences tripled.

Article 8 (Use of fines): 40% to State Budget, 30% to Environment Fund, 30% to inspecting entity.""",

    "Regulamento sobre a Gestão de Resíduos Sólidos Urbanos (Decreto n.º 94/2014)": """DECREE NO. 94/2014 — REGULATION ON URBAN SOLID WASTE MANAGEMENT

Article 2 (Purpose): Establishes rules for urban solid waste management nationwide.

Article 4 (General principles): Includes self-sufficiency, producer responsibility, polluter-pays, and waste hierarchy (prevention, reuse, recycling, other recovery, disposal).

Article 5 (Competencies): Environment Ministry issues rules, conducts inspections, maintains national register, coordinates with municipalities. Municipalities and district governments must ensure adequate waste management, adopt municipal bylaws, define collection procedures, promote recycling and selective collection, set service fees, and penalise offenders.

Article 6 (Municipal obligations): Ensure waste is not dumped on beaches, in the sea, watercourses, or burned in open air. Maintain annual waste registers.

Article 8 (Integrated Waste Management Plans): All waste management entities must prepare PGIRSU plans based on the waste hierarchy.

Definitions include: segregation, selective collection, recycling, sanitary landfill, special waste (including contaminated plastics), and urban solid waste from domestic and commercial activities.""",

    "Estratégia de Desenvolvimento da Economia Azul – EDEA (Resolução n.º 53/2024)": """BLUE ECONOMY DEVELOPMENT STRATEGY (EDEA) 2024–2033

Vision: Promote sustainable use of marine, coastal and inland water resources through a profitable and sustainable blue economy.

Strategic Objective (Circular Economy pillar):
(c) Reduce plastic production, import and use; increase extended producer responsibility in industry and commerce; support recycling and remanufacturing industries.

Coastal management actions include preventing plastic pollution through reallocation, remanufacturing/recycling and reintroduction into the economic circuit, while discouraging import or production of hard-to-degrade products.

The strategy references municipal waste composition studies (e.g., 8% plastic fraction in KaNyaka) and plans investment in waste treatment and valorisation technology.""",

    "Lei do Ambiente (Lei n.º 20/97)": """ENVIRONMENTAL LAW (LAW NO. 20/97)

Enacted by the Assembly of the Republic, 1 October 1997.

Purpose: Establish legal bases for correct use and management of the environment and its components, toward sustainable development.

Key provisions relevant to plastic pollution:
- Article 7: Prohibits release of pollutants into the environment without proper authorisation.
- Article 33: Authorises the executive to regulate waste management (basis for Decreto 94/2014 and sectoral regulations).
- Establishes environmental impact assessment, licensing, and the polluter-pays principle.
- Creates institutional framework including the National Council for Sustainable Development (CONDES).

Note: The uploaded document (Environmental_law_13c6.pdf) is an image-based copy of this law; this English summary is based on the official legal framework.""",
}


def build_rows() -> list[dict]:
    rows = []
    for policy in CODING:
        for inst in policy["instruments"]:
            row = {
                "A_policy_name": policy["policy_name"],
                "B_policy_url": policy["policy_url"],
                "C_policy_year": policy["policy_year"],
                "D_policy_objective": policy["policy_objective"],
                "E_policy_target": policy["policy_target"],
                "F_policy_target_text": policy["policy_target_text"],
                "G_policy_type": policy["policy_type"],
                "H_policy_type_justification": policy["policy_type_justification"],
                "I_policy_integration": policy["policy_integration"],
                "J_policy_sectors_list": policy["policy_sectors_list"],
                "K_policy_circularity": policy["policy_circularity"],
                "L_policy_lifecycle_phases_list": policy["policy_lifecycle_phases_list"],
                "M_policy_budget": policy["policy_budget"],
                "N_policy_budget_text": policy["policy_budget_text"],
                "O_policy_score": "[auto]",
                "P_instrument_type": inst["instrument_type"],
                "Q_instrument_lifecycle_stage": inst["instrument_lifecycle_stage"],
                "R_instrument_description": inst["instrument_description"],
                "S_instrument_in_force": inst["instrument_in_force"],
                "T_instrument_implementation": inst["instrument_implementation"],
                "U_instrument_implementation_text": inst["instrument_implementation_text"],
                "V_instrument_score": "[auto]",
                "W_comments": inst.get("comments", ""),
            }
            rows.append(row)
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
        # policy_score = AVERAGE(G, I, K, M, P)
        ws[f"O{i}"] = f"=AVERAGE(G{i},I{i},K{i},M{i},P{i})"
        ws[f"V{i}"] = f"=AVERAGE(P{i},T{i})"

    widths = {
        "A": 38, "B": 42, "C": 8, "D": 40, "E": 8, "F": 35, "G": 8, "H": 30,
        "I": 8, "J": 30, "K": 8, "L": 28, "M": 8, "N": 30, "O": 10,
        "P": 8, "Q": 18, "R": 45, "S": 8, "T": 8, "U": 35, "V": 10, "W": 35,
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
    meta.append(["Coding date", "June 2026"])
    meta.append(["Total policies", len(CODING)])
    meta.append(["Total instrument rows", len(rows)])
    meta.append(["Note", "Columns O and V use Excel formulas. policy_score = AVERAGE(G,I,K,M,P); instrument_score = AVERAGE(P,T)."])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_translations_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading("Mozambique Plastic Pollution Policies — English Translations", level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English translations and summaries of key Mozambique national policies addressing plastic pollution. "
        "Where full official English translations do not exist, summaries are provided based on verified Portuguese legal texts. "
        "June 2026."
    )

    for policy in POLICIES:
        pt_title = policy["title_pt"]
        en_title = policy["title_en"]
        doc.add_heading(en_title, level=2)
        p = doc.add_paragraph()
        p.add_run("Portuguese title: ").bold = True
        p.add_run(pt_title)

        if policy.get("link"):
            p2 = doc.add_paragraph()
            p2.add_run("Source URL: ").bold = True
            p2.add_run(policy["link"])
        else:
            p2 = doc.add_paragraph()
            p2.add_run("Source: ").bold = True
            p2.add_run(policy.get("link_note", "not available"))

        if pt_title in TRANSLATIONS:
            doc.add_heading("English translation / summary", level=3)
            for para in TRANSLATIONS[pt_title].strip().split("\n\n"):
                doc.add_paragraph(para.strip())
        else:
            doc.add_heading("English summary", level=3)
            doc.add_paragraph(policy["description"])

        doc.add_paragraph("")

    doc.save(WORD_OUT)
    print(f"Wrote {WORD_OUT}")


def main() -> None:
    rows = build_rows()
    write_excel(rows)
    write_translations_doc()


if __name__ == "__main__":
    main()
