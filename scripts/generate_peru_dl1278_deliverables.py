#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DL 1278 (Solid Waste Management Law)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-dl-1278")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

POLICY_NAME_ES = (
    "Decreto Legislativo Nº 1278 – Ley de Gestión Integral de Residuos Sólidos"
)
POLICY_NAME_EN = (
    "Legislative Decree that approves the Comprehensive Solid Waste Management Law "
    "(Decreto Legislativo Nº 1278)"
)
POLICY_URL = "https://www.minam.gob.pe/disposiciones/decreto-legislativo-n-1278/"
COUNTRY = "Peru"

POLICY = {
    "policy_name_es": POLICY_NAME_ES,
    "policy_name_en": POLICY_NAME_EN,
    "policy_url": POLICY_URL,
    "country": COUNTRY,
    "policy_year": 2024,
    "policy_objective": (
        "National framework law for integrated solid waste management across the full "
        "life cycle — minimization, segregation, collection, valorization/recycling, and "
        "final disposal. Establishes circular economy and Extended Producer Responsibility "
        "(EPR) principles directly applicable to packaging and plastic-containing products, "
        "assigns competencies to MINAM, OEFA, sector authorities, and municipalities, and "
        "creates planning, information, financing, and sanction instruments for plastic and "
        "other waste streams."
    ),
    "policy_target": 1,
    "policy_target_text": (
        "Art. 49: 'El MINAM establece metas anuales para la valorización de residuos "
        "municipales'. Art. 13 (as amended Ley 32212): MINAM approves via supreme decree "
        "'objetivos, las metas... y los plazos' for prioritized goods EPR regimes (including "
        "packaging). DCT Tercera: municipalities without selective collection systems must "
        "approve them within one (01) year. DCF Sexta Transitoria: ten (10) years for "
        "construction/reconversion of final disposal infrastructure. Art. 39: maximum twelve "
        "(12) hours temporary storage at transfer stations."
    ),
    "policy_type": 0.75,
    "policy_type_justification": (
        "Legislative Decree Nº 1278 issued by the Executive Power (22 December 2016) under "
        "delegated legislative authority from Congress (Ley Nº 30506). Last substantively "
        "amended by parliamentary Law Nº 32212 (21 December 2024). Classified as 0.75 "
        "(executive legislation/regulation) rather than 1.0 because the base instrument was "
        "not enacted directly by Congress, though subsequent parliamentary amendments apply."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "production, consumption, retail, waste management, recycling, municipalities, "
        "industry, health, packaging, chemicals, water, agriculture, tourism"
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": (
        "production, consumption, recycling, disposal, environmental leakage"
    ),
    "policy_budget": 1,
    "policy_budget_text": (
        "Art. 70-A (incorporated Ley 32212): municipal cleaning-fee revenue 'se destina "
        "única y exclusivamente' to cleaning service and waste infrastructure. Ley 26793 Art. 3 "
        "g) (amended 32212): FONAM resources from environmental-crime reparations dedicated "
        "to degraded-area recovery and landfill financing. DCF Séptima (Ley 32212): "
        "municipalities may constitute fideicomisos for cleaning-service infrastructure."
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 2: Waste management hierarchy — first priority is prevention/minimization at "
            "source; second priority recovery and material/energy valorization (reuse, recycling, "
            "composting); final disposal is last resort in environmentally adequate infrastructure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Binding hierarchy in law ('primera finalidad', 'en segundo lugar', 'última "
            "alternativa'). MINAM/OEFA/municipal enforcement via broader framework. No single "
            "penalty article in Art. 2."
        ),
        "comments": "Foundational waste hierarchy applicable to all streams including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 3: State guarantees continuous, regular, permanent, and mandatory public "
            "cleaning service including collection, transport, and final disposal of solid waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'El Estado garantiza la prestación... obligatoria'. Municipal responsibility under "
            "Art. 24. Authorities must adopt investment measures. No explicit fines in Art. 3."
        ),
        "comments": "Public service guarantee for municipal waste including plastic packaging waste.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 4: National application to production, import, distribution of goods/services "
            "and all waste management from generation to final disposal, including cross-border "
            "waste movements and degraded areas."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Mandatory national scope. Multi-authority enforcement framework. Applies to "
            "packaging/plastics through production and consumption chains."
        ),
        "comments": "Overarching scope instrument.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 5(c) & Art. 12–14: Extended Producer Responsibility (REP/EPR) — producers, "
            "importers, distributors must use eco-efficient products/packaging minimizing waste "
            "and facilitating valorization; participate across product life cycle. Art. 13 "
            "(amended Ley 32212): MINAM approves by supreme decree prioritized goods regimes, "
            "actor obligations, objectives, targets, management systems and timelines."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "EPR principle binding in Art. 5(c). Art. 12: producers 'se involucran activamente'. "
            "Art. 13 enabling power exercised through subordinate supreme decrees for prioritized "
            "goods including packaging. OEFA/sector enforcement of EPR obligations. MINAM "
            "designated authority."
        ),
        "comments": (
            "Core EPR instrument directly applicable to plastic packaging via prioritized-goods "
            "regimes. Ley 32212 strengthened Art. 13 operative language. Specific plastic "
            "targets in subordinate EPR supreme decrees."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5(d) & Art. 6(h): Shared responsibility system for integrated waste management "
            "from generation to final disposal, involving generators, operators, and municipalities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Principle and lineamiento establishing coordination framework. No explicit penalties "
            "in these articles."
        ),
        "comments": "Governance/coordination principle underpinning EPR and municipal systems.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 8: Production of goods/services must privilege efficient material use, "
            "eco-design, process optimization, and use of discard materials as inputs — "
            "preventing waste generation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Privilegia el uso eficiente' — binding orientation on all productive sectors. "
            "Annex defines ecodiseño for products, packaging, labeling. Enforcement through "
            "sectoral norms and SEIA instruments."
        ),
        "comments": "Production-phase eco-design/material efficiency standard relevant to plastic packaging.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 11: State procurement at all government levels oriented to efficient material "
            "use, energy efficiency, and waste minimization."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Está orientada' — mandatory orientation for public purchasing. Public entities "
            "designated. No explicit penalty in Art. 11."
        ),
        "comments": "Public procurement instrument reducing plastic/waste in government consumption.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 15: MINAM national lead — coordinates application; formulates PLANRES with targets "
            "for universal cleaning service, recycler formalization, minimization and valorization; "
            "administers SIGERSOL and EO-RS registry; regulates waste infrastructure and import/"
            "export authorizations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "MINAM extensively designated with planning, regulatory, and information functions. "
            "PLANRES includes quantified targets (monitoring). OEFA enforcement linkage. Broad "
            "mandate without single penalty in Art. 15."
        ),
        "comments": "National governance hub. Bundled MINAM competencies as one coordination instrument.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 16 & 75: OEFA supervises, fiscalizes, and sanctions municipal and private "
            "infrastructure for treatment, valorization, and final disposal; typifies infractions "
            "and sanction scales; maintains National Inventory of Degraded Areas; may issue "
            "coercive fines (DCF Décima)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "OEFA designated ('supervisar, fiscalizar y sancionar'). Art. 16(d): tipificación and "
            "sanction scales. DCF Décima: 'multa coercitiva automática' for non-compliance with "
            "administrative measures. Inventory maintenance (monitoring)."
        ),
        "comments": "Core national enforcement instrument for waste infrastructure including plastic waste flows.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 23–24: Provincial municipalities plan via PIGARS; district municipalities approve "
            "PMR, ensure cleaning/collection/disposal, and progressively implement source segregation "
            "and selective collection programs (Art. 24.2d)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipalities 'deben' plan, approve, ensure services, 'promover e implementar "
            "progresivamente' segregation/selective collection. DCT Tercera: 1-year deadline for "
            "selective collection systems. Municipal fiscalization/sanctions (Arts. 78–79)."
        ),
        "comments": "Municipal planning and progressive segregation obligation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 33–34: Segregation required at source or authorized valorization infrastructure; "
            "prohibited at final disposal sites. Municipal and non-municipal generators must deliver "
            "segregated waste to authorized operators. Municipalities must define segregation criteria "
            "by legal instrument within one year if not yet done."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Obligados a entregar... debidamente segregados'. Art. 33: 'Queda prohibida la "
            "segregación' at disposal sites. Municipal fiscalization of generators (Art. 24.2g, "
            "78–79). One-year criteria deadline (quantified)."
        ),
        "comments": "Key regulatory instrument for source separation of recyclables including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 35: Collection must be selective per municipal provisions; formalized recyclers "
            "integrate into municipal selective collection systems without environmental exposure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Debe ser selectiva' — mandatory selective collection. Municipal authority sets "
            "requirements. Recycler integration mandatory. Municipal supervision/sanctions apply."
        ),
        "comments": "Regulatory selective collection standard.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 37 & 47–48: Valorization prioritized over disposal; includes reuse, recycling, "
            "composting, energy recovery. Valorization grounded in selective collection and EPR "
            "regimes for prioritized goods."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Debe priorizarse' (Art. 37). Links to EPR and selective collection systems. "
            "Enforcement via municipal/OEFA frameworks."
        ),
        "comments": "Valorization priority rule applicable to plastic material recovery.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 49: MINAM establishes annual targets for municipal waste valorization based on "
            "selective collection systems; compliance verified by MINAM."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'establece metas anuales' — quantified annual targets. MINAM verification (monitoring). "
            "MINAM designated authority. Specific target values set administratively."
        ),
        "comments": "Quantitative valorization target instrument. Annual targets include plastic recyclables.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 44: Prohibited to abandon, dump, or dispose waste in unauthorized places; "
            "illegal dumps (botaderos) must be closed by provincial municipalities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Está prohibido el abandono, vertido o disposición'. Municipal closure duty. Sanctions "
            "under Arts. 81–82 via Ley General del Ambiente Art. 136. OEFA/municipal enforcement."
        ),
        "comments": "Anti-littering/illegal dumping ban directly addressing environmental leakage.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 41–43: Non-valorizable waste must be disposed in authorized infrastructure; "
            "on-site disposal in extractive/productive facilities requires environmental instruments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deben ser aislados y/o confinados en infraestructuras debidamente autorizadas'. "
            "SEIA/municipal licensing requirements. OEFA oversight of disposal infrastructure."
        ),
        "comments": "Final disposal regulatory standard.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 42: Import and transit permitted only for valorization; export for valorization or "
            "final disposal; prior MINAM authorization required; Basel Convention compliance; no "
            "import/transit of radioactive or dangerous waste."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "MINAM authorization required ('se requerirá una autorización previa'). Competent "
            "authorities must control compliance. Applies to plastic waste shipments under Basel."
        ),
        "comments": "Transboundary waste control including plastic waste streams.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 30: Containers used for hazardous substances and expired dangerous products "
            "classified as hazardous waste unless treated to remove hazard; packaging subject to "
            "hazardous waste rules (strengthened Ley 32212)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Son considerados residuos peligrosos y deben ser manejados como tales'. Sector/MINSA "
            "enforcement. Reuse restrictions after treatment per 2024 amendment."
        ),
        "comments": "Regulatory standard for hazardous plastic/chemical containers.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Art. 55: Non-municipal waste generators must segregate, maintain appropriate storage, "
            "prioritize valorization, ensure treatment/disposal, keep internal records, report via "
            "SIGERSOL, submit annual declarations and waste management plans when required."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Se encuentran obligados' — extensive mandatory duties list (a–j). SIGERSOL reporting "
            "(monitoring). Sector authority fiscalization and sanctions. Hazardous waste manifests."
        ),
        "comments": "Industrial/commercial generator obligations covering plastic production waste.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 60–61: Waste services and commercialization through registered Empresas "
            "Operadoras de Residuos Sólidos (EO-RS) with MINAM registration, qualified technical "
            "director, and appropriate infrastructure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deben estar debidamente registradas'. MINAM registry (Art. 15q). Mandatory "
            "obligations list Art. 61. OEFA/authority supervision. Insurance for hazardous waste "
            "(Art. 62)."
        ),
        "comments": "Operator licensing/registration regulatory framework.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 64: Formalized recyclers integrate into municipal non-hazardous waste management "
            "systems; municipal supervision and fiscalization of compliance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Integration duty on municipalities/recyclers. 'Supervisado y fiscalizado por la autoridad "
            "municipal'. Links to Law 29419 framework."
        ),
        "comments": "Governance integration of waste pickers recovering plastics.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Art. 45–46 & 66: Degraded waste areas must be recovered, closed, or reconverted to "
            "authorized disposal infrastructure; polluter-pays with state fallback; OEFA maintains "
            "national inventory; requires approved environmental instruments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Deben ser recuperadas' — mandatory remediation. OEFA inventory (monitoring). "
            "Approved IGA required before operations. Polluter liability with state backup."
        ),
        "comments": "Infrastructure/remediation instrument for legacy dumps and degraded sites.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 51: Municipalities must prioritize valorization of organic waste from green areas "
            "and markets; municipal parks are priority beneficiaries of compost/biochar produced."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Deben valorizar, prioritariamente' — mandatory municipal duty. No explicit penalty in "
            "Art. 51; general sanctions apply."
        ),
        "comments": "Organic waste valorization obligation; complements plastic stream separation.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 68: SIGERSOL — official national information system for municipal and non-municipal "
            "waste planning, management, reporting, and public dissemination (part of SINIA); "
            "municipalities assign responsible officials; MINAM approves indicators and methodologies."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Mandatory reporting platform. Municipal responsible officer required. MINAM indicators "
            "(monitoring). 2024 amendment: SIGERSOL used for supervision/fiscalization."
        ),
        "comments": "National data/monitoring governance instrument.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 69: MINAM, municipalities, and sectors must promote environmental education on "
            "sustainable consumption, source segregation, arbitrio payment, and EPR participation "
            "for prioritized goods."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "'Deben promover' — binding promotional duty on MINAM/municipalities/sectors. MINAM "
            "must establish lineamientos. No explicit penalties."
        ),
        "comments": "Education/awareness instrument including EPR consumer role.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 70 & 70-A (Ley 32212): Municipalities may use utility bills to collect cleaning "
            "fees; revenue ring-fenced exclusively for cleaning service and waste infrastructure; "
            "may offer arbitrio discounts as segregation incentives."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Municipalities designated. Art. 70-A: revenue 'única y exclusivamente' for service "
            "(ring-fenced). Segregation discounts permitted (economic incentive). 2024 incorporation."
        ),
        "comments": "Economic financing instrument with ring-fenced fee revenue.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 71 & Ley 26793 Art. 3(g) (amended 32212): FONAM and PROFONANPE finance waste "
            "management projects and degraded-area recovery; FONAM receives environmental-crime "
            "reparation funds dedicated to waste infrastructure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "FONAM/PROFONANPE designated. Dedicated revenue stream for waste infrastructure from "
            "reparations (2024). Project financing rather than direct penalties in Art. 71."
        ),
        "comments": "Economic funding instrument. DCF Séptima adds municipal fideicomisos.",
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 7(m) & DCF Décima (Ley 32212): Waste Bags (Bolsas de Residuos) mechanism to "
            "facilitate commercial transactions of waste and discard materials; MINAM to implement "
            "platform within 180 days (per DS 001-2022-MINAM complementary provision)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Instrument listed in Art. 7. Implementation via MINAM platform (enabling/regulatory "
            "development). Limited enforcement detail in parent law."
        ),
        "comments": "Market facilitation instrument for recyclable materials including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 74–83 & 81–82: Authorities supervise, fiscalize, and sanction non-compliance; "
            "administrative infractions for violating law and regulation; sanctions per Ley General "
            "del Ambiente Art. 136; regional/local governments typify infractions within parameters."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": (
            "Art. 81: 'incumplimiento de las obligaciones' are infractions. Art. 82: sanctions under "
            "LGA Art. 136. Art. 80: municipalities typify fines and measures. OEFA coercive fines "
            "(DCF Décima). Multi-authority enforcement."
        ),
        "comments": "General sanctions framework for all obligations including EPR and segregation.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 14: Producers may conclude collaboration agreements with municipalities or EO-RS "
            "for source segregation, selective collection, reception/storage facilities for prioritized "
            "waste, and other EPR implementation actions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "'Pueden celebrar convenios' — voluntary agreement mechanism enabling EPR "
            "implementation. No mandatory language; complements binding EPR regimes."
        ),
        "comments": (
            "Scored 1.0 as part of EPR framework but in_force uses permissive 'pueden'; primarily "
            "enabling/voluntary unless incorporated in binding EPR DS. Implementation 0.25."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 7 instruments: PLANRES, PIGARS, PMR, annual non-domestic waste declarations, EPR "
            "recovery plans for prioritized goods, and Clean Production Agreements — national and "
            "local planning framework."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Planning instruments listed Art. 7; MINAM approves PLANRES with targets (Art. 15b). "
            "Municipal plans mandatory (Art. 23–24). Monitoring via SIGERSOL and CAM reports."
        ),
        "comments": "Bundled planning/governance instruments under Art. 7.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 36 & Annex: Storage must follow segregation criteria and NTP 900.058 color code "
            "for containers; generator responsible until municipal handover."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'Debe ser realizado siguiendo los criterios de segregación'. Mandatory NTP color "
            "standard cited. Generator responsibility until delivery. Municipal enforcement."
        ),
        "comments": "Technical storage/segregation standard for recyclable fractions including plastics.",
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "End of life",
        "instrument_description": (
            "Art. 39: Transfer stations may not store waste more than twelve (12) hours; transfer "
            "only between vehicles at authorized infrastructure."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "'No se permitirá el almacenamiento temporal... por más de doce (12) horas' — "
            "quantified unconditional prohibition. Infrastructure authorization required."
        ),
        "comments": "Operational time-limit standard at transfer infrastructure.",
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 26: Municipal Environmental Commissions (CAM) annually evaluate municipal waste "
            "management performance per MINAM guides; report to OEFA and SINIA."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "CAM 'evalúa el desempeño' annually (monitoring). Report to OEFA mandatory. MINAM "
            "issues evaluation guides."
        ),
        "comments": "Performance monitoring governance instrument.",
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": (
            "DCF Sexta Transitoria: Declaration of public necessity and national interest for public "
            "investment in final disposal infrastructure construction and reconversion over ten (10) "
            "years."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Quantified 10-year priority investment declaration. Enables infrastructure "
            "acceleration. No direct penalties; facilitates public investment."
        ),
        "comments": "Infrastructure investment promotion with 10-year quantified horizon.",
    },
]

COLUMNS = [
    ("A", "policy_name_es"),
    ("B", "policy_name_en"),
    ("C", "country"),
    ("D", "policy_url"),
    ("E", "policy_year"),
    ("F", "policy_objective"),
    ("G", "policy_target"),
    ("H", "policy_target_text"),
    ("I", "policy_type"),
    ("J", "policy_type_justification"),
    ("K", "policy_integration"),
    ("L", "policy_sectors_list"),
    ("M", "policy_circularity"),
    ("N", "policy_lifecycle_phases_list"),
    ("O", "policy_budget"),
    ("P", "policy_budget_text"),
    ("Q", "policy_score"),
    ("R", "instrument_type"),
    ("S", "instrument_lifecycle_stage"),
    ("T", "instrument_description"),
    ("U", "instrument_in_force"),
    ("V", "instrument_implementation"),
    ("W", "instrument_implementation_text"),
    ("X", "instrument_score"),
    ("Y", "comments"),
]


def policy_score_row(instrument_type: float) -> float:
    return round(
        (
            POLICY["policy_type"]
            + POLICY["policy_integration"]
            + POLICY["policy_circularity"]
            + POLICY["policy_budget"]
            + instrument_type
        )
        / 5,
        3,
    )


def instrument_score_row(instrument_type: float, implementation: float) -> float:
    return round((instrument_type + implementation) / 2, 3)


def build_excel(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "4P Index Coding"
    ws.append([f"{col} — {name}" for col, name in COLUMNS])
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    for inst in INSTRUMENTS:
        ws.append(
            [
                POLICY["policy_name_es"],
                POLICY["policy_name_en"],
                POLICY["country"],
                POLICY["policy_url"],
                POLICY["policy_year"],
                POLICY["policy_objective"],
                POLICY["policy_target"],
                POLICY["policy_target_text"],
                POLICY["policy_type"],
                POLICY["policy_type_justification"],
                POLICY["policy_integration"],
                POLICY["policy_sectors_list"],
                POLICY["policy_circularity"],
                POLICY["policy_lifecycle_phases_list"],
                POLICY["policy_budget"],
                POLICY["policy_budget_text"],
                policy_score_row(inst["instrument_type"]),
                inst["instrument_type"],
                inst["instrument_lifecycle_stage"],
                inst["instrument_description"],
                inst["instrument_in_force"],
                inst["instrument_implementation"],
                inst["instrument_implementation_text"],
                instrument_score_row(inst["instrument_type"], inst["instrument_implementation"]),
                inst["comments"],
            ]
        )

    widths = {
        "A": 34, "B": 36, "C": 10, "D": 38, "E": 8, "F": 44, "G": 8, "H": 44,
        "I": 8, "J": 40, "K": 10, "L": 40, "M": 10, "N": 36, "O": 10, "P": 42,
        "Q": 10, "R": 10, "S": 18, "T": 50, "U": 10, "V": 12, "W": 44, "X": 10, "Y": 40,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"
    wb.save(path)


def add_article(doc: Document, es: str, en: str) -> None:
    p = doc.add_paragraph()
    r = p.add_run(es)
    r.bold = True
    doc.add_paragraph(en).paragraph_format.space_after = Pt(8)


def build_word(path: Path) -> None:
    doc = Document()
    for margin in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(doc.sections[0], margin, Inches(1))

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t = title.add_run(f"{POLICY_NAME_EN}\n{POLICY_NAME_ES}")
    t.bold = True
    t.font.size = Pt(14)

    sub = doc.add_paragraph()
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub.add_run(
        "Original adoption: 22 December 2016 | Last amended: 21 December 2024 (Law Nº 32212)\n"
        "Also amended by: D.L. Nº 1501 (2020), Law Nº 31896 (2023)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Legislative Decree Nº 1278, incorporating "
        "material amendments from Law Nº 32212 (2024) where noted."
    )

    sections = [
        (
            "Article 1 — Object",
            "establece derechos, obligaciones, atribuciones y responsabilidades...",
            "This Legislative Decree establishes rights, obligations, powers, and responsibilities "
            "of society as a whole to maximize material efficiency and ensure economically, "
            "sanitarily, and environmentally adequate solid waste management.",
        ),
        (
            "Article 2 — Purpose of integrated waste management",
            "La gestión integral... tiene como primera finalidad la prevención o minimización...",
            "Integrated waste management's first purpose is prevention or minimization at source. "
            "For generated waste, recovery and material/energy valorization (reuse, recycling, "
            "composting, coprocessing) are preferred. Final disposal in appropriate infrastructure "
            "is the last alternative.",
        ),
        (
            "Article 5(c) — Extended Producer Responsibility principle",
            "Principio de responsabilidad extendida del productor...",
            "Extended Producer Responsibility principle: manufacturers, importers, distributors, "
            "and retailers shall manufacture or use products or packaging with eco-efficiency criteria "
            "minimizing waste generation and/or facilitating valorization, and are responsible for "
            "participating across life-cycle stages.",
        ),
        (
            "Title III — EPR (Articles 12–14)",
            "Los fabricantes, importadores, distribuidores y comerciantes se involucran activamente...",
            "Producers must actively engage across product life-cycle stages using eco-design and "
            "waste prevention. Art. 13 (amended 2024): MINAM approves by supreme decree prioritized "
            "consumer goods subject to special management regimes, including actor obligations, "
            "objectives, targets, management systems, and implementation timelines. Art. 14: producers "
            "may sign collaboration agreements with municipalities or waste operators for segregation, "
            "selective collection, and reception facilities.",
        ),
        (
            "Article 8 — Material efficiency and eco-design",
            "La producción de bienes y servicios... privilegia el uso eficiente de los insumos...",
            "Production of goods and services must privilege efficient use of inputs and materials "
            "through eco-design, process optimization, innovation, and use of discard materials as "
            "inputs.",
        ),
        (
            "Article 34 — Source segregation",
            "Los generadores... están obligados a entregar los residuos debidamente segregados...",
            "Municipal and non-municipal generators must deliver properly segregated waste to "
            "authorized operators. Municipal generators must classify waste to facilitate recovery. "
            "Municipalities must approve segregation criteria by legal instrument within one year if "
            "not yet established.",
        ),
        (
            "Article 35 — Selective collection",
            "La recolección de los residuos debe ser selectiva...",
            "Waste collection must be selective per municipal provisions. Formalized recyclers "
            "integrate into municipal selective collection systems.",
        ),
        (
            "Article 37 & 49 — Valorization priority and annual targets",
            "La valorización constituye la alternativa... que debe priorizarse...",
            "Valorization must be prioritized over disposal. MINAM establishes annual municipal "
            "valorization targets based on selective collection systems and verifies compliance.",
        ),
        (
            "Article 44 — Prohibition of unauthorized disposal",
            "Está prohibido el abandono, vertido o disposición de residuos en lugares no autorizados...",
            "Abandonment, dumping, or disposal of waste in unauthorized places is prohibited. "
            "Illegal dumps must be closed by provincial municipalities.",
        ),
        (
            "Article 55 — Non-municipal generator obligations",
            "Los generadores de residuos del ámbito no municipal se encuentran obligados a...",
            "Non-municipal generators must segregate waste, maintain appropriate storage, prioritize "
            "valorization, ensure treatment and disposal, keep records, report via SIGERSOL, and "
            "submit annual declarations and updated management plans.",
        ),
        (
            "Article 60 — Waste operator registration",
            "las Empresas Operadoras de Residuos Sólidos deben estar debidamente registradas ante el MINAM",
            "Waste services and commercialization must be performed by Waste Operator Companies "
            "(EO-RS) registered with MINAM, with qualified technical direction and appropriate "
            "equipment and infrastructure.",
        ),
        (
            "Articles 74–82 — Supervision and sanctions",
            "Constituyen infracciones administrativas el incumplimiento...",
            "Failure to comply with the Law and its Regulation constitutes an administrative "
            "infraction. Authorities apply sanctions under General Environment Law Art. 136. "
            "OEFA and municipalities typify infractions and impose fines within established parameters.",
        ),
        (
            "2024 Amendment — Article 70-A (cleaning fee collection)",
            "Lo recaudado por el concepto de limpieza pública se destina única y exclusivamente...",
            "Municipalities may collect cleaning fees through utility bills. Revenue from cleaning "
            "fees is dedicated exclusively to cleaning service and waste infrastructure investments "
            "(Law 32212, Art. 70-A).",
        ),
        (
            "2024 Amendment — FONAM financing (Ley 26793 Art. 3g)",
            "montos pagados por concepto de reparaciones civiles... destinados a... rellenos sanitarios",
            "FONAM receives civil reparation payments from environmental crimes related to waste "
            "mismanagement, dedicated to recovering degraded areas and financing sanitary landfills "
            "(Law 32212).",
        ),
        (
            "Annex — Key definitions (ecodesign, EPR collection center)",
            "Ecodiseño... Centro de acopio municipal...",
            "Eco-design: designing products, packaging, and labeling to minimize environmental "
            "impacts and maximize material efficiency across the life cycle. Municipal collection "
            "center: infrastructure storing non-hazardous waste recovered through source segregation, "
            "selective collection, or EPR programs.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DL 1278 derogated Law 27314 upon entry into force of its Regulation (DS "
        "014-2017-MINAM). Implementing regulation DS 014-2017-MINAM was amended by DS "
        "001-2022-MINAM. EPR targets for specific prioritized goods (including packaging and "
        "plastics) are set in subordinate supreme decrees approved under Art. 13."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DL_1278_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DL_1278_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")


if __name__ == "__main__":
    main()
