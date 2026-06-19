#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Decreto 79/2017 only.

Source legal text: FDUEM Colectânea compilation (user upload), section Dec_79_2017.txt.
Supplementary analysis: Lexology article (user-provided URL; Cloudflare-blocked in automation),
VdA Legal Partners flash 2018, IUCN MARPLASTICCS policy assessment 2022, Club of Mozambique /
Lusa reporting on Diploma Ministerial 26/2025, Imani Development Africa RISE project 2022–2025.

Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens —
Decreto n.º 79/2017, de 28 de Dezembro.
"""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

EXCEL_OUT = Path("/workspace/Mozambique_4P_Index_Decreto_79_2017.xlsx")
WORD_OUT = Path("/workspace/Decreto_79_2017_English_Translation.docx")

LEXOLOGY_URL = (
    "https://www.lexology.com/library/detail.aspx?g=bc88751e-ed7d-4f3f-972c-158b0f8fbbdf"
)

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
    "policy_name": (
        "Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens "
        "(Decreto n.º 79/2017, de 28 de Dezembro)"
    ),
    "policy_url": LEXOLOGY_URL,
    "policy_year": 2017,
    "policy_objective": (
        "Adopt principles, norms and guidelines assigning extended responsibility to producers "
        "and importers of packaging and packaging waste to protect public health and the environment "
        "within sustainable development, through internal management systems, the Packaging "
        "Environmental Fee (TAE), and packaging normalisation."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree (Decreto n.º 79/2017) approved by the Council of Ministers on "
        "21 November 2017 and published 28 December 2017 (Boletim da República), under "
        "Lei 20/97 Arts. 10 and 33. Sub-legislative EPR instrument, not parliamentary legislation."
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": (
        "packaging, plastics, industry, retail, import, municipalities, waste management, "
        "recycling, chemicals, finance"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        "Art. 24: TAE revenues allocated 60% State Budget / 40% FNDS; fine revenues 40% State / "
        "60% FNDS, ring-fenced for integrated waste management, selective collection, recycling, "
        "valorisation, environmental education and institutional strengthening. COMAGE operating "
        "costs borne by FNDS (Art. 23(5))."
    ),
    "instruments": [
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 4(f): Waste hierarchy principle — management must follow priority order: "
                "prevention and reduction, reuse, recycling, other recovery, then disposal; using "
                "best available economically sustainable technologies to prolong material life cycles."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'a gestão de resíduos sólidos urbanos deve respeitar a seguinte ordem de prioridades "
                "— prevenção e redução, reutilização, reciclagem, outras formas de valorização e "
                "eliminação.' Binding hierarchy in Regulation (+0.25). MTA monitors compliance "
                "(Art. 5). Limited evidence of systematic hierarchy application for plastic packaging "
                "(IUCN 2022: ~1% packaging recycled)."
            ),
            "comments": (
                "Plastics relevance: Annex I typology includes all plastic types including polystyrene "
                "foam. Lexology/VdA analysis confirms hierarchy as core EPR design element."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 4(d): Prevention and eco-design principle — all actors in the packaging life "
                "cycle must adopt ecological design and sustainable consumption practices appropriate "
                "to legal and technical norms."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'adoptando as práticas de desenho ecológico e de consumo sustentável mais adequadas.' "
                "Principle stated (+0.25 authority via Art. 5 MTA). No quantified eco-design targets or "
                "specific enforcement beyond Art. 17 normalisation."
            ),
            "comments": "Regulatory principle instrument; complements Art. 17 material standards.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 5: Assigns competencies to Environment Ministry (rules, fiscalization, "
                "monitoring, valorisation oversight), Industry/Commerce Ministry (import/production "
                "standards), and Finance Ministry (tax/fine collection, customs clearance control)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Compete ao Ministério que superintende a área do Ambiente' fiscalizar, sancionar e "
                "monitorar; MIC define normas de produção/importação; MEF garante cobrança de taxas e "
                "controlo aduaneiro. Multi-ministry coordination (+0.25). Africa RISE/Imani project "
                "2022–2025 supported TAE diploma and COMAGE operationalisation (+0.25)."
            ),
            "comments": "Multi-level governance (Rule 11); Decreto Art. 2 assigns MTA, MIC, MEF implementation.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 6: Municipal councils and local State bodies responsible for urban solid waste "
                "management in their jurisdictions; may receive financial counterparties from approved "
                "packaging waste management projects and may contract licensed public/private operators."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'As Autarquias e os Órgãos Locais do Estado são responsáveis pela gestão de resíduos "
                "sólidos urbanos.' Municipal authority designated (+0.25). Project-based counterparties "
                "require submission and approval — discretionary, not automatic funding."
            ),
            "comments": "Links EPR revenues (Art. 24 FNDS) to local waste infrastructure.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 7: Producers and importers co-responsible for packaging and packaging-waste "
                "management — must pay TAE fees and ensure return/valorisation of packaging waste "
                "directly or through recovery organisations."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Os produtores e importadores são co-responsáveis pela gestão das embalagens e "
                "resíduos de embalagens.' Mandatory co-responsibility (+0.25). TAE payment obligation "
                "(+0.25) operationalising via DM 26/2025. Take-back/valorisation weakly implemented "
                "for plastics (IUCN 2022; VdA 2018 anticipates significant business impact)."
            ),
            "comments": (
                "Core EPR obligation. Annex I explicitly covers all plastic types including polystyrene "
                "foam (item 1.2)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 8: Waste operators must ensure environmentally safe, sustainable management "
                "favouring reduction, recycling and reuse; promote environmental education; and "
                "register with the Environment Ministry."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'Assegurar uma gestão ambientalmente segura, sustentável e racional das embalagens, "
                "tendo em conta a necessidade da sua redução, reciclagem e reutilização.' Registration "
                "requirement (+0.25). Limited evidence of comprehensive operator registration for "
                "plastic packaging streams."
            ),
            "comments": "Applies to licensed operators handling plastic packaging waste.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 9–10: Three combined EPR systems — Internal Management (direct/indirect), "
                "Packaging Environmental Fee (TAE), and Packaging Normalisation — may be applied "
                "jointly for optimal packaging management."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'A responsabilidade do produtor e importador de embalagens é assumida através dos "
                "seguintes sistemas' (internal management, TAE, normalisation). Framework operative "
                "(+0.25). IUCN 2022 / Lexology-type analysis: only TAE system substantially "
                "implemented; internal management and normalisation need greater producer investment."
            ),
            "comments": "Overarching EPR architecture instrument; one row for Art. 9 system design.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 11: Direct internal management — producers/importers may individually or jointly "
                "pursue reduction, reuse, recycling, organic recovery, energy recovery or incineration; "
                "consumers pay a deposit at purchase refunded on return of used packaging."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'o consumidor de produtos que utilizem embalagens paga um determinado valor de "
                "depósito no acto da compra, que lhe é devolvido aquando da entrega da embalagem "
                "utilizada.' Deposit-return mechanism defined (+0.25). No national deposit system "
                "evidenced operational for plastic packaging (IUCN recommends Maputo pilot)."
            ),
            "comments": (
                "Deposit-return (DRS) for packaging including plastics. Art. 20(2)(a) penalises "
                "refusal to accept used packaging or refund deposits (15 minimum wages)."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 12: Indirect internal management — producers/importers may contract licensed "
                "waste operators (PRO-type entities) for collection, selective sorting, take-back and "
                "valorisation; financial counterparties required; responsibility transfers upon "
                "certified operator assumption declaration."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'A responsabilidade dos produtores ou importadores... pode ser transferida, mediante "
                "celebração de contrato, para uma entidade devidamente licenciada.' PRO framework "
                "defined (+0.25). IUCN: indirect management system not yet operational at scale; "
                "producer inventories and PRO mandates need development."
            ),
            "comments": "Mozambique PRO/contractual EPR model per VdA and IUCN analysis.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 13–14: Creates variable Packaging Environmental Fee (TAE) payable by all "
                "packaging producers and importers, based on environmental/public-health impact and "
                "treatment complexity; criteria include returnability, decomposition time/impact, "
                "treatment cost, and eco-design."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'É criada a Taxa Ambiental sobre a Embalagem (TAE).' DM 26/2025 (10 March 2025) "
                "defines formula, categories and exemptions per Art. 15 (+0.25). Levied on primary "
                "packaging including plastic (+0.25). AT collects on imports; annual production reports "
                "for domestic producers (Art. 16). Payment from 2026 after management-plan period "
                "(Club of Mozambique/Lusa). Up to 90% reduction when internal management established."
            ),
            "comments": (
                "Previously scored in_force=0 pending diploma; DM 26/2025 operationalises Art. 15. "
                "Lexology URL and Africa RISE project document TAE as primary implemented EPR pillar."
            ),
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 15–16: Ministerial diploma (DM 26/2025) defines TAE formula, procedures, packaging "
                "categories and exemptions; AT collects TAE on imported packaging at customs clearance; "
                "domestic producers pay annually based on production reports."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'Por Diploma Ministerial conjunto... são definidas a fórmula de cálculo.' DM 26/2025 "
                "joint MTA/MEF/MIC adopted (+0.25). Formula based on waste treatment cost and "
                "material/weight impact factor (+0.25). 180-day entry into force; e-TAE platform "
                "referenced in news reports (+0.25 monitoring). Non-payment fined at 15 min. wages "
                "(Art. 20(2)(c))."
            ),
            "comments": "Enabling Art. 15 now exercised; distinguishes diploma mechanics from Art. 13 creation.",
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 17: Packaging normalisation — packaging must preferentially use biodegradable "
                "materials or materials permitting reuse, recycling or recovery; producers must limit "
                "volume/weight, design for returnability, and ensure recyclability. Applies to "
                "manufacturers, material suppliers, importers and distributors."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'As embalagens devem ser produzidas com materiais preferencialmente de natureza "
                "biodegradável ou que permitam a reutilização, reciclagem ou valorização' and 'Recicláveis.' "
                "Mandatory language (+0.25). Art. 20(2)(d): market placement without packaging norms "
                "— 15 min. wages (+0.25). VdA notes significant business impact; limited enforcement "
                "data for plastic streams."
            ),
            "comments": (
                "Annex I typology item 1.2: all plastic types including polystyrene foam. "
                "Packaging Standardisation System per IUCN."
            ),
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 18: Packaging normalisation symbols for reusable, recyclable or recoverable "
                "packaging must be adopted on packaging or labels — clearly visible, legible, and "
                "durable for packaging lifetime including after opening."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'devem ser adoptados símbolos específicos para as embalagens reutilizáveis, recicláveis "
                "ou valorizáveis' on packaging/labels. Information/labelling requirement (+0.25). "
                "Symbol standards not further specified in decree — implementation depends on "
                "complementary norms."
            ),
            "comments": "Information-type labelling; scored 0.40 per mandatory visibility requirement (Rule 10).",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 19: Environment Ministry fiscalizes Regulation compliance and decides on fines "
                "and accessory sanctions; Municipal Councils and District Administrations must "
                "collaborate by providing information for enforcement."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'Compete ao Ministério que superintende a área do Ambiente fiscalizar o cumprimento "
                "do presente Regulamento.' Authority (+0.25), fine decision power linked to Arts. 20–22 "
                "(+0.25). Municipal collaboration duty (+0.25 monitoring)."
            ),
            "comments": "Primary enforcement governance for EPR including plastic packaging.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 20–22: Administrative fines — obstruction of inspection (10 min. wages); "
                "refusal of used packaging/deposit refund, placing packaged products without waste "
                "management assurance, non-payment of TAE, or non-compliant packaging (15 min. wages "
                "each); 30% cumulative increase for repeat offences; accessory sanctions include "
                "seizure, activity suspension, compulsory removal; 20-day payment deadline."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "'Constituem infracções... puníveis com sanção de multa correspondente a 10/15 Salários "
                "Mínimos.' Explicit fines (+0.25). TAE non-payment sanction (+0.25). Accessory "
                "sanctions Art. 21 (+0.25). Coercive fiscal execution via Juízo das Execuções Fiscais "
                "if unpaid (Art. 22)."
            ),
            "comments": "Economic enforcement; VdA flash confirms 10–15 minimum wage fine range.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 23: Creates COMAGE (Comissão de Monitoria e Avaliação da Gestão de Embalagens) "
                "— multistakeholder consultative body to advise on packaging management, coordinate "
                "authorities and operators, monitor TAE revenue use, and propose legal improvements; "
                "operating costs funded by FNDS."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'É criada a Comissão de Monitoria e Avaliação da Gestão de Embalagens (COMAGE).' "
                "Body created (+0.25). Africa RISE/Imani 2022–2025 supported COMAGE operationalisation "
                "under MTA (+0.25). Consultative not executive — implementation depends on functioning "
                "secretariat and stakeholder participation."
            ),
            "comments": "Lexology/Africa RISE identify COMAGE as key EPR governance institution.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 24: TAE revenue split 60% State / 40% FNDS; fine revenue 40% State / 60% FNDS; "
                "FNDS funds ring-fenced for integrated waste management, selective collection, reuse, "
                "recycling, valorisation, environmental education, and institutional strengthening "
                "for packaging waste."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "'40% para o FNDS' (taxes) and '60% para o FNDS' (fines) with enumerated eligible "
                "actions including selective collection and recycling (+0.25 ring-fenced allocation). "
                "COMAGE monitors TAE revenue application (Art. 23(2)(c)). Revenue flows depend on "
                "TAE collection commencing 2026."
            ),
            "comments": "Supports policy_budget = 1; links TAE to circular-economy waste infrastructure.",
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 4(g): Environmental education principle — provide education and training to "
                "increase citizens' capacity for sustainable development and environmental quality tasks."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "'pressupõe providenciar a educação e formação ambiental.' Principle stated (+0.25). "
                "Art. 8(1)(c) requires waste operators to promote community awareness. FNDS may fund "
                "education actions (Art. 24(3)(c)) but no mandatory programme specified."
            ),
            "comments": "Information/education instrument complementing operator duties Art. 8(1)(c).",
        },
    ],
}

ENGLISH_TRANSLATION = """DECREE NO. 79/2017 OF 28 DECEMBER 2017
Approving the Regulation on Extended Producer Responsibility of Producers and Importers of Packaging

Source document: FDUEM Colectânea de Legislação sobre o Ambiente (2020), user-uploaded compilation.
Supplementary analysis: Lexology (user-provided URL), VdA Legal Partners flash 2018, IUCN MARPLASTICCS 2022.

Approved by the Council of Ministers on 21 November 2017.
Published 28 December 2017.
The Prime Minister, Carlos Agostinho do Rosário.

---

DECREE OF THE COUNCIL OF MINISTERS

There being a need to adopt principles, norms and guidelines for the accountability of producers and importers of packaging, with a view to protecting the environment and public health within the objective of sustainable development, under the provisions of Articles 10 and 33 of Law No. 20/97 of 1 October, the Council of Ministers decrees:

Article 1. The Regulation on Extended Producer Responsibility of Producers and Importers of Packaging is approved, annexed hereto and forming an integral part thereof.

Article 2. It is the responsibility of the Ministries overseeing the areas of Environment, Industry, Commerce and Finance to ensure implementation of this Regulation.

Article 3. This Decree enters into force on the date of its publication.

---

REGULATION ON EXTENDED PRODUCER RESPONSIBILITY OF PRODUCERS AND IMPORTERS OF PACKAGING

CHAPTER I — GENERAL PROVISIONS

Article 1 (Definitions)
The meaning of terms used in this Regulation is set out in the glossary annexed hereto.

Article 2 (Purpose)
1. This Regulation establishes principles and norms relating to extended producer responsibility of producers and importers of packaging and packaging waste, to ensure protection of public health and the environment within sustainable development.
2. Extended producer responsibility for other classes of waste is subject to specific regulation.

Article 3 (Scope)
1. This Regulation applies to all public and private entities, natural and legal persons involved in production, import and management of packaging.
2. It applies to all packaging placed on the market, at domestic, industrial, agricultural or commercial levels, including offices, shops and services, regardless of material, and to packaging waste susceptible of collection and treatment.
3. This Regulation does not prejudice legislation on packaging quality, product safety and hygiene, or hazardous waste provisions.

Article 4 (Principles)
a) Polluter Pays: the polluter must bear costs of repairing environmental damage caused.
b) Extended producer/importer responsibility: physical and/or financial responsibility for packaging impacts assigned to producers and importers.
c) Systemic vision in solid waste management: environmental, social, cultural, economic, technological and public health dimensions.
d) Prevention and reduction: all life-cycle actors must adopt ecological design and sustainable consumption.
e) Protection of human health and the environment: waste operations must avoid adverse effects on water, air, soil, fauna, flora, noise, odours and landscape.
f) Waste hierarchy: urban solid waste management must follow priority order — prevention/reduction, reuse, recycling, other recovery, disposal — using best available economically sustainable technologies.
g) Environmental education: provide education and training to increase citizens' capacity for sustainable development.

CHAPTER II — COMPETENCIES AND RESPONSIBILITIES

Article 5 (Competencies)
1. Environment Ministry: elaborate and disseminate rules; fiscalize, sanction and monitor compliance; ensure environmental norms in valorisation of packaging waste.
2. Industry and Commerce Ministry: define import/production standards; ensure industry and commerce observe environmental management principles.
3. Finance Ministry: ensure collection of taxes and fines; ensure only compliant packaging enters at customs clearance.

Article 6 (Local authorities)
1. Municipal councils and local State bodies are responsible for urban solid waste management in their jurisdictions and may benefit from financial counterparties through approved packaging waste management projects.
2. They may transfer management responsibilities via contracts with licensed public or private entities.

Article 7 (Producers and importers)
Producers and importers are co-responsible for packaging and packaging-waste management, including payment of TAE fees and return/valorisation of packaging waste directly or through recovery organisations.

Article 8 (Waste operators)
Waste operators must ensure environmentally safe, sustainable and rational management favouring reduction, recycling and reuse; protect health and environment; promote community education; and register with the Environment Ministry.

CHAPTER III — PRODUCER RESPONSIBILITY REGIME

Section I — Application systems

Article 9 (Systems)
Producer/importer responsibility is assumed through: (a) Internal Management System (direct and indirect); (b) Packaging Environmental Fee (TAE); (c) Packaging Normalisation System. Systems may be combined.

Article 10 (Internal management)
Internal management may be adopted on producer initiative in direct or indirect modalities.

Article 11 (Direct internal management)
1. Producers/importers may pursue reduction, reuse, recycling, organic recovery, energy recovery or incineration individually or jointly.
2. Consumers pay a deposit at purchase refunded on return of used packaging.

Article 12 (Indirect internal management)
1. Responsibility may be transferred by contract to licensed waste operators.
2. Any duly registered waste operator, public or private, national or foreign, may provide packaging management services.
3. Producers/importers must provide financial counterparties for selective collection, sorting, take-back and valorisation.
4. Responsibility ceases upon certified operator assumption declaration.

Section II — Packaging Environmental Fee (TAE)

Article 13 (TAE)
1. The Packaging Environmental Fee (TAE) is created, payable by all packaging producers and importers.
2. TAE is variable according to environmental/public-health impact and treatment complexity.

Article 14 (TAE criteria)
Criteria: returnability; natural decomposition time and impact; treatment cost in country or abroad; eco-design.

Article 15 (Formula)
By joint Ministerial Diploma of Finance, Environment and Industry/Commerce Ministers, the calculation formula, application norms and procedures, packaging categories and exemptions are defined.

Article 16 (Collection)
1. Mozambique Tax Authority collects TAE on imported packaging.
2. For domestically produced packaging, TAE is paid annually based on company production reports.

Section III — Packaging Normalisation

Article 17 (Normalisation)
1. Packaging must preferentially use biodegradable materials or materials permitting reuse, recycling or recovery.
2. Producers must ensure packaging is: (a) limited in volume/weight; (b) designed for returnability; (c) recyclable.
3. Applies to manufacturers, material suppliers, importers and distributors at all trade stages.

Article 18 (Symbols)
Reusable, recyclable or recoverable packaging symbols must appear on packaging or labels, clearly visible, legible and durable.

CHAPTER IV — FISCALIZATION, INFRACTIONS AND SANCTIONS

Article 19 (Fiscalization)
Environment Ministry fiscalizes compliance and decides fines; municipalities and district administrations must collaborate.

Article 20 (Infractions)
1. Obstruction of inspection without just cause — fine of 10 minimum wages.
2. Fine of 15 minimum wages for: refusal to accept used packaging or refund deposits; placing imported packaged products without assured packaging management; non-payment of TAE; placing products without respecting packaging norms.
3. Repeat offences aggravated by 30% cumulatively.

Article 21 (Accessory sanctions)
May include seizure of irregular equipment/products, suspension of production activity, and compulsory removal of discarded packaging.

Article 22 (Fine payment)
20-day payment deadline; coercive fiscal execution if unpaid.

CHAPTER V — FINAL PROVISIONS

Article 23 (COMAGE)
Multistakeholder Packaging Management Monitoring and Evaluation Commission created, chaired by Environment Ministry, to advise, coordinate, monitor TAE revenues and propose improvements. Operating costs borne by FNDS.

Article 24 (Revenue allocation)
TAE: 60% State Budget, 40% FNDS. Fines: 40% State, 60% FNDS. FNDS funds selective collection, recycling, valorisation, environmental education and institutional strengthening.

Article 25 (Updates)
Environment, Industry/Commerce and Finance Ministers may update tax and fine values.

ANNEX I — PACKAGING TYPOLOGY
Materials include: glass; all plastic types including polystyrene foam; paper/cardboard (simple and multilayer); metals. Volume categories from 0–250 ml to >250,000 ml.

---

Implementation update (supplementary sources, not in decree text):
Diploma Ministerial No. 26/2025 (10 March 2025) defines TAE formula and procedures under Art. 15. Levy applies to primary packaging including plastic; payment deferred approximately one year after 180-day entry into force (reported from 2026); up to 90% TAE reduction when producer establishes internal management system (Club of Mozambique/Lusa, 2025)."""


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
        "A": 48, "B": 22, "C": 8, "D": 45, "E": 8, "F": 12, "G": 8, "H": 35,
        "I": 8, "J": 35, "K": 8, "L": 28, "M": 8, "N": 40, "O": 10,
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
    meta.append(["Primary legal text", "FDUEM Colectânea — Decreto 79/2017 section (user upload)"])
    meta.append(["Analysis source (user)", LEXOLOGY_URL])
    meta.append([
        "Supplementary sources",
        "VdA Legal Partners flash 2018; IUCN MARPLASTICCS 2022; Club of Mozambique/Lusa DM 26/2025; "
        "Imani Development Africa RISE 2022–2025",
    ])
    meta.append(["Coding date", "June 2026"])
    meta.append(["Total instrument rows", len(rows)])
    meta.append([
        "Note",
        "Columns O and V use Excel formulas. "
        "policy_score = AVERAGE(G,I,K,M,P); instrument_score = AVERAGE(P,T). "
        "TAE in_force updated to 1 following DM 26/2025 (March 2025).",
    ])

    wb.save(EXCEL_OUT)
    print(f"Wrote {EXCEL_OUT} ({len(rows)} instrument rows)")


def write_translation_doc() -> None:
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.75)
    section.right_margin = Inches(0.75)

    title = doc.add_heading(
        "Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens "
        "(Decreto 79/2017) — English Translation",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_paragraph(
        "English translation of Mozambique's Regulation on Extended Producer Responsibility for "
        "Packaging, approved by Council of Ministers Decree No. 79/2017 of 28 December 2017, "
        "based on the FDUEM legislation compilation (user upload). Coding supplemented by Lexology "
        "analysis and implementation reporting on Diploma Ministerial 26/2025."
    )

    p = doc.add_paragraph()
    p.add_run("Portuguese title: ").bold = True
    p.add_run(
        "Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens "
        "(Decreto n.º 79/2017, de 28 de Dezembro)"
    )

    p2 = doc.add_paragraph()
    p2.add_run("Lexology analysis: ").bold = True
    p2.add_run(LEXOLOGY_URL)

    p3 = doc.add_paragraph()
    p3.add_run("Legal text source: ").bold = True
    p3.add_run("FDUEM Colectânea de Legislação sobre o Ambiente (2020), user-uploaded PDF.")

    doc.add_heading("English translation", level=2)

    for block in ENGLISH_TRANSLATION.strip().split("\n\n"):
        text = block.strip()
        if not text:
            continue
        if text.startswith("---"):
            continue
        if text.startswith(("DECREE", "REGULATION", "CHAPTER", "Section", "ANNEX")):
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
        elif text.startswith("Article ") or text.startswith("Implementation"):
            doc.add_heading(text.split("\n")[0], level=3)
            rest = "\n".join(text.split("\n")[1:]).strip()
            if rest:
                doc.add_paragraph(rest)
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
