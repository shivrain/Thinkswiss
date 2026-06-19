#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Decreto 79/2017 only.

Source legal text: FDUEM Colectânea compilation (user upload), section Dec_79_2017.txt.
Primary analysis: Lexology article (user-provided text and URL).
Implementation updates: Diploma Ministerial 26/2025, Club of Mozambique/Lusa 2025,
IUCN MARPLASTICCS 2022, Imani Development Africa RISE 2022–2025.

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

LEXOLOGY_ANALYSIS = """On 29 December 2017, Decree 79/2017, of 28 December 2017, approving the Regulation on the Extended Responsibility of Producers and Importers of Packaging (the "Regulation"), came into force.

The Regulation's objective is the adoption of principles, rules and guidelines to increase the responsibility of producers and importers of packaging in order to safeguard the environment and public health, in the context of sustainable development.

To whom is it applicable?

To all public and private entities, natural or legal persons engaged in the production, import and management of packaging.

What kind of packaging falls within its scope?

Any packaging placed on the market, used or produced, namely at a domestic, industrial, agricultural or commercial level, including offices, shops and services, regardless of the materials used, as well as packaging waste which can be collected and processed by systems currently in place or to be put in place.

Who has authority and responsibility for the management of packaging and packaging waste?

Article 5 of the Regulation determines that the Ministries responsible for the Environment (e.g. the drafting and disclosure of rules and procedures in the context of the production and import of packaging and packaging waste, its supervision and the imposition of sanctions), Industry and Commerce (e.g. establishment of rules and standards applicable to import and production of packaging), and Finance (e.g. the collection of fees and fines and the supervision of the rules applicable to packaging in the context of the clearance of goods).

Article 5 of the Regulation also attributes the following responsibilities and authority on producers and importers of packaging and waste operators:

Producers and importers of packaging and packaging waste

Jointly responsible for the management of packaging and packaging waste, pursuant to the Regulation and other applicable legislation
Payment of fees for the management of packaging
Return and recovery of packaging waste, whether directly or through organisations created for waste recovery
Waste operators

Ensure the environmentally safe, sustainable and rational management of packaging, in view of the need to reduce, recycle and reuse it, including the sorting, collection, handling, transportation, storage and/or elimination of packaging waste
Foster the protection of human health and the environment against the harmful effects that may result from the disposal of packaging
Promote community education and awareness initiatives regarding the proper management of packaging
Register with the Ministry responsible for the Environment

How does the framework of producer and importer responsibility work?

Packaging producers and importers assume their responsibility through the following:

Internal Management System
Packaging Environmental Fee System
Packaging Standardisation System

Internal Management Systems (may be adopted at the producer's initiative and has two forms)

Direct Internal Management System: the producer or the importer may opt, whether individually or jointly, for any one of the following procedures: i) reduction; (ii) re-use; (iii) recycling; (iv) organic recovery; (v) energetic recovery; vi) incineration. In this system the consumer of products that uses packaging pays a specific deposit amount upon purchase which is reimbursed once the used packaging is returned.

Indirect Internal Management System: producer and importer responsibility for the treatment of packaging or packaging waste can be transferred, contractually, to an entity duly licensed to undertake that activity. In this system the producer or importer must bear the necessary financial costs of the processes of selective collection and sorting of packaging waste, as well as for the take back and recovery of packaging waste. The producer or importer's responsibility for the final destination of the packaging ends when a declaration of assumption of responsibility is issued by the certified waste operator to whom the packaging is delivered.

Packaging Environmental Fee System

Article 13 of the Regulation creates a Packaging Environmental Fee (Taxa Ambiental sobre a Embalagem, "TAE") which must be paid by all producers and importers of packaging, and it varies depending on the packaging's environmental and public health impact and the complexity of the treatment of the waste resulting from the packaging.

In the case of imported packaging, the TAE is charged by the Mozambican Tax Authority (Autoridade Tributária de Moçambique).

In the case of packaging produced in national territory, the TAE is paid annually on the basis of the company's production report.

Packaging Standardisation System

This system requires that all packaging be produced with preferably biodegradable materials or materials which can be re-used, recycled or recovered. It is for (i) producers, (ii) manufacturers of packaging or those who supply materials for the manufacture of packaging, and (iii) those who import or put packaging, materials for the manufacture of packaging or packaged goods into circulation to ensure that packaging:

Is limited to the volume and dimensions required for the protection of its contents and the sale of the good
Is designed so that it is reusable in a technically viable way which is compatible with the requirements of the good it contains
Is recyclable

Symbol and labelling of packaging

In the context of the Packaging Standardisation System, the use of specific symbols is mandatory for reusable, recyclable or packaging that can be subject to recovery, and it should be placed on the packaging itself or its label so that it is clearly visible, legible and that it shall last as long as the expected lifetime of the packaging.

Who supervises and what are the sanctions for breaches of the Regulation?

The Ministry responsible for the Environment is responsible for the supervision, with the cooperation of Municipal Councils and District Administrations.

Breaches to the Regulation are sanctionable with a fine which amount may vary between 10 and 15 times the Minimum Salary and may be increased by 30% in the event of reoccurrence. Together with the imposition of a fine, ancillary sanctions may also be imposed as stipulated in Article 21.

Observations

The Regulation creates a wide scope of responsibilities including financial charges, on producers and importers of packaging, as well as responsibilities for waste operators.

We anticipate that the Regulation will have considerable impact on companies, who will be under an obligation to assume the management of packaging waste through the three systems of producer and importer responsibility, and will be required to pay the fees and costs, as well as the fines and ancillary sanctions in the event of breach of the Regulation.

The effectiveness of the Regulation's application will depend on the approval of complementary and regulatory legislation, on the proper functioning of the Commission for the Supervision and Evaluation of the Management of Packaging ("Comissão de Monitoria e Avaliação da Gestão de Embalagens") created by this Regulation, and on the operation of the various entities involved in the extended responsibility framework."""

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
        "Adoption of principles, rules and guidelines to increase the responsibility of "
        "producers and importers of packaging in order to safeguard the environment and "
        "public health, in the context of sustainable development (Lexology; Reg. Art. 2)."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 0.75,
    "policy_type_justification": (
        "Executive decree (Decreto n.º 79/2017) approved by the Council of Ministers on "
        "21 November 2017, published 28 December 2017, entered into force 29 December 2017 "
        "(Lexology). Sub-legislative EPR instrument under Lei 20/97 Arts. 10 and 33."
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
                "Lexology: scope covers any packaging 'regardless of the materials used' plus "
                "collectable packaging waste. Annex I includes all plastic types including polystyrene foam."
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
                "Art. 5: Environment Ministry drafts/discloses rules, supervises and imposes sanctions; "
                "Industry and Commerce Ministry sets import/production standards; Finance Ministry "
                "collects fees and fines and supervises packaging rules at customs clearance (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Lexology confirms Art. 5 ministry competencies for Environment (rules, supervision, "
                "sanctions), Industry/Commerce (standards), Finance (fee/fine collection, customs). "
                "Multi-ministry coordination (+0.25). Africa RISE/Imani 2022–2025 supported TAE "
                "diploma and COMAGE operationalisation (+0.25)."
            ),
            "comments": (
                "Lexology: effectiveness depends on operation of entities in extended responsibility "
                "framework. Decreto Art. 2 assigns MTA, MIC, MEF implementation."
            ),
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
                "Art. 7: Producers and importers jointly responsible for packaging and packaging-waste "
                "management; must pay packaging management fees; must ensure return and recovery of "
                "packaging waste directly or through recovery organisations (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Lexology: producers/importers 'jointly responsible' for management, fee payment, and "
                "return/recovery (+0.25 each obligation). TAE fee system operationalising via DM 26/2025 "
                "(+0.25). Lexology anticipates 'considerable impact on companies' — take-back/recovery "
                "weakly implemented for plastics (IUCN 2022)."
            ),
            "comments": (
                "Core EPR obligation per Lexology. Scope includes all packaging materials; Annex I "
                "item 1.2 covers all plastic types including polystyrene foam."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 8: Waste operators must ensure environmentally safe, sustainable and rational "
                "packaging management (reduce, recycle, reuse; sorting, collection, handling, transport, "
                "storage and/or disposal); protect health and environment; promote community education; "
                "and register with Environment Ministry (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Lexology lists four operator duties: safe management, health/environment protection, "
                "community education, Environment Ministry registration. Registration (+0.25). Limited "
                "evidence of comprehensive operator registration for plastic packaging streams."
            ),
            "comments": "Lexology attributes separate responsibilities to waste operators alongside producers.",
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
                "Lexology: producers/importers assume responsibility through Internal Management, "
                "Packaging Environmental Fee, and Packaging Standardisation systems. Framework "
                "operative (+0.25). Lexology observations: companies must manage waste through all "
                "three systems; effectiveness depends on complementary legislation and entity operation."
            ),
            "comments": (
                "Overarching EPR architecture per Lexology. One row for Art. 9 system design."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 11: Direct Internal Management System — producer/importer may opt individually or "
                "jointly for reduction, re-use, recycling, organic recovery, energetic recovery or "
                "incineration; consumer pays deposit at purchase reimbursed on return of used packaging "
                "(Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Lexology: deposit 'reimbursed once the used packaging is returned.' Mechanism defined "
                "(+0.25). Adopted at producer initiative — no national DRS evidenced operational for "
                "plastic packaging (IUCN recommends Maputo pilot)."
            ),
            "comments": (
                "Lexology Internal Management System (direct form). Art. 20(2)(a): refusal to accept "
                "used packaging or refund deposit — 15 minimum wages."
            ),
        },
        {
            "instrument_type": 1.0,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 12: Indirect Internal Management System — producer/importer responsibility for "
                "packaging treatment transferred contractually to licensed entity; producer bears "
                "financial costs of selective collection, sorting, take-back and recovery; "
                "responsibility ends upon certified operator declaration of assumption (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Lexology: contractual transfer to licensed operator; producer bears collection/sorting/"
                "take-back/recovery costs (+0.25). Responsibility ends on operator assumption "
                "declaration (+0.25 framework). Not yet operational at scale (IUCN 2022)."
            ),
            "comments": "Lexology Indirect Internal Management System; PRO-type contractual EPR model.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Recycling",
            "instrument_description": (
                "Art. 13: Packaging Environmental Fee (TAE / Taxa Ambiental sobre a Embalagem) payable "
                "by all packaging producers and importers; fee varies by environmental/public-health "
                "impact and waste-treatment complexity. AT charges imports; domestic producers pay "
                "annually on production report (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "Lexology: TAE created Art. 13; AT collects imported packaging; annual production-report "
                "basis for domestic producers (+0.25 each). DM 26/2025 defines formula per Art. 15 "
                "(Lexology flagged need for complementary legislation) (+0.25). Levied on plastic "
                "packaging; payment from 2026; up to 90% reduction with internal management system."
            ),
            "comments": (
                "Lexology Packaging Environmental Fee System. Complementary diploma DM 26/2025 (March "
                "2025) addresses Lexology observation on regulatory legislation dependency."
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
                "Art. 17: Packaging Standardisation System — packaging produced with preferably "
                "biodegradable or re-usable/recyclable/recoverable materials; producers, manufacturers, "
                "material suppliers and importers must ensure volume/weight limits, technical "
                "returnability, and recyclability (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Lexology: three standardisation requirements — volume/dimension limits, technical "
                "reusability, recyclability (+0.25). Applies to producers, manufacturers, suppliers, "
                "importers (+0.25). Art. 20(2)(d) fine 15 min. wages for non-compliant placement."
            ),
            "comments": (
                "Lexology Packaging Standardisation System. Annex I item 1.2: all plastic types "
                "including polystyrene foam."
            ),
        },
        {
            "instrument_type": 0.40,
            "instrument_lifecycle_stage": "Consumption",
            "instrument_description": (
                "Art. 18: Mandatory symbols for reusable, recyclable or recoverable packaging on "
                "packaging or label — clearly visible, legible, lasting expected packaging lifetime "
                "(Lexology: Symbol and labelling of packaging)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.25,
            "instrument_implementation_text": (
                "Lexology: 'use of specific symbols is mandatory' on packaging or label, visible and "
                "legible for packaging lifetime (+0.25). Symbol technical standards depend on "
                "complementary norms not yet specified."
            ),
            "comments": "Lexology symbol/labelling requirement under Packaging Standardisation System.",
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 19: Environment Ministry supervises Regulation compliance; Municipal Councils "
                "and District Administrations cooperate in supervision (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Lexology: Environment Ministry 'responsible for the supervision' with municipal and "
                "district cooperation (+0.25 authority, +0.25 multi-level). Linked to Arts. 20–21 "
                "fine and ancillary sanction powers (+0.25)."
            ),
            "comments": "Lexology supervision framework for EPR including plastic packaging.",
        },
        {
            "instrument_type": 0.60,
            "instrument_lifecycle_stage": "Production",
            "instrument_description": (
                "Art. 20–21: Breaches sanctionable with fines of 10–15 times Minimum Salary; 30% "
                "increase on reoccurrence; ancillary sanctions per Art. 21 may be imposed together "
                "with fines (Lexology)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.75,
            "instrument_implementation_text": (
                "Lexology: fines 'between 10 and 15 times the Minimum Salary', +30% reoccurrence; "
                "ancillary sanctions with fines (+0.25 each). Decreto Art. 20 specifies obstruction "
                "(10), TAE non-payment and packaging breaches (15). Coercive fiscal execution "
                "if unpaid (Art. 22)."
            ),
            "comments": (
                "Lexology economic enforcement. Companies liable for 'fines and ancillary sanctions "
                "in the event of breach' per Lexology observations."
            ),
        },
        {
            "instrument_type": 0.20,
            "instrument_lifecycle_stage": "Waste management",
            "instrument_description": (
                "Art. 23: COMAGE (Comissão de Monitoria e Avaliação da Gestão de Embalagens / "
                "Commission for the Supervision and Evaluation of the Management of Packaging) — "
                "multistakeholder body; proper functioning required for Regulation effectiveness "
                "(Lexology observations)."
            ),
            "instrument_in_force": 1,
            "instrument_implementation": 0.50,
            "instrument_implementation_text": (
                "Lexology: effectiveness 'will depend on... the proper functioning of' COMAGE (+0.25). "
                "Body created Art. 23 (+0.25). Africa RISE/Imani 2022–2025 supported operationalisation "
                "under MTA (+0.25). Consultative mandate limits executive enforcement power."
            ),
            "comments": (
                "Lexology observations flag COMAGE as critical dependency for EPR effectiveness."
            ),
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
Primary analysis: Lexology article (user-provided text).
Implementation update: Diploma Ministerial 26/2025 (Club of Mozambique/Lusa, 2025).

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
    meta.append(["Primary analysis", LEXOLOGY_URL])
    meta.append(["Lexology text", "User-provided (incorporated in generator and Word doc)"])
    meta.append([
        "Implementation updates",
        "DM 26/2025 (March 2025); IUCN MARPLASTICCS 2022; Africa RISE/Imani Development 2022–2025",
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
        "Packaging, approved by Council of Ministers Decree No. 79/2017 of 28 December 2017 "
        "(entered into force 29 December 2017), based on the FDUEM legislation compilation "
        "(user upload). The accompanying 4P Index coding uses the Lexology legal analysis "
        "(user-provided) cross-checked against the decree text, with DM 26/2025 implementation notes."
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

    doc.add_heading("Lexology legal analysis (coding source)", level=2)
    for block in LEXOLOGY_ANALYSIS.strip().split("\n\n"):
        text = block.strip()
        if not text:
            continue
        if text.endswith("?") or text in (
            "To whom is it applicable?",
            "What kind of packaging falls within its scope?",
            "Who has authority and responsibility for the management of packaging and packaging waste?",
            "How does the framework of producer and importer responsibility work?",
            "Who supervises and what are the sanctions for breaches of the Regulation?",
            "Observations",
            "Producers and importers of packaging and packaging waste",
            "Waste operators",
            "Internal Management System",
            "Packaging Environmental Fee System",
            "Packaging Standardisation System",
            "Symbol and labelling of packaging",
        ):
            doc.add_heading(text.rstrip("?") if not text.endswith("?") else text, level=3)
        elif text.startswith("Direct Internal") or text.startswith("Indirect Internal"):
            doc.add_heading(text.split(":")[0], level=4)
            rest = text.split(":", 1)[1].strip() if ":" in text else ""
            if rest:
                doc.add_paragraph(rest)
        else:
            doc.add_paragraph(text)

    doc.add_heading("English translation of decree text", level=2)

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
