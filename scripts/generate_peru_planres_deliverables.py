#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru PLANRES 2016–2024 (RM 191-2016-MINAM)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-planres")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

POLICY_NAME_ES = (
    "Resolución Ministerial N.° 191-2016-MINAM – Plan Nacional de Gestión Integral de "
    "Residuos Sólidos (PLANRES) 2016–2024"
)
POLICY_NAME_EN = (
    "Ministerial Resolution No. 191-2016-MINAM – National Comprehensive Solid Waste "
    "Management Plan (PLANRES) 2016–2024"
)
POLICY_URL = (
    "https://sinia.minam.gob.pe/documentos/plan-nacional-gestion-integral-residuos-solidos-planres"
)
COUNTRY = "Peru"


def bi(es: str, en: str) -> str:
    """Bilingual cell text: original Spanish / English translation."""
    return f"{es} / {en}"


POLICY = {
    "policy_name_es": POLICY_NAME_ES,
    "policy_name_en": POLICY_NAME_EN,
    "policy_url": POLICY_URL,
    "country": COUNTRY,
    "policy_year": 2016,
    "policy_objective": bi(
        "Instrumento de política nacional que establece objetivos, estrategias, programas y "
        "metas para todo el ciclo de los residuos sólidos — minimización, segregación, "
        "recolección selectiva, reaprovechamiento/reciclaje y disposición final segura — "
        "incluyendo plásticos y otros materiales reciclables; articula capacidades, "
        "institucionalidad e inversión entre MINAM/CONAM, gobiernos regionales y locales, "
        "salud y sector privado.",
        "National policy instrument setting goals, strategies, programs and targets for the "
        "entire solid-waste cycle — minimization, segregation, selective collection, "
        "recovery/recycling and safe final disposal — including plastics and other recyclable "
        "materials; coordinates capacity building, institutional development and investment "
        "among MINAM/CONAM, regional and local governments, health authorities and the "
        "private sector.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "Metas cuantificadas del Plan (Anexo operativo): en 5 años 70% de funcionarios "
        "municipales capacitados y 80% de centros educativos con currícula de recolección "
        "selectiva; reducción del 10% de generación per cápita de residuos municipales; 70% "
        "de capitales regionales con planes de reaprovechamiento; 100% de manipuladores bajo "
        "salud ocupacional; 40% de municipios provinciales con recolección selectiva; en 10 "
        "años 95% de disposición adecuada, 100% de botaderos municipales clausurados/"
        "recuperados, canalización de US$ 100 millones en inversiones y 50% de municipios "
        "provinciales con tecnologías apropiadas.",
        "Quantified plan targets (operational annex): within 5 years 70% of municipal officials "
        "trained and 80% of schools with selective-collection curriculum; 10% per-capita "
        "reduction in municipal waste generation; 70% of regional capitals with reuse plans; "
        "100% of waste handlers under occupational health programs; 40% of provincial "
        "municipalities with selective collection; within 10 years 95% adequate disposal, "
        "100% municipal dumps closed/recovered, US$100 million investment program channeled, "
        "and 50% of provincial municipalities with appropriate technologies.",
    ),
    "policy_type": 0.50,
    "policy_type_justification": bi(
        "Resolución Ministerial N.° 191-2016-MINAM aprobada por el Ministerio de Ambiente "
        "(CONAM) el 30 de junio de 2016. Es un plan nacional de gestión integral con metas "
        "y programas, no legislación parlamentaria ni decreto supremo reglamentario. "
        "Actualización en curso desde 2023 (segundo entregable publicado 27 dic. 2023); "
        "nuevo PLANRES pendiente a 2026. Clasificado 0.50 (plan/instrumento de política "
        "nacional con objetivos cuantificados).",
        "Ministerial Resolution No. 191-2016-MINAM approved by the Ministry of Environment "
        "(CONAM) on 30 June 2016. It is a national integrated management plan with targets "
        "and programs, not parliamentary legislation or supreme regulatory decree. Under "
        "update since 2023 (second deliverable published 27 Dec 2023); new PLANRES pending as "
        "of 2026. Classified 0.50 (national policy plan with quantified objectives).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "ambiente, gobiernos municipales y regionales, salud, producción, consumo, comercio, "
        "industria, educación, turismo, minería, agroindustria, inversión privada, reciclaje",
        "environment, municipal and regional governments, health, production, consumption, "
        "trade, industry, education, tourism, mining, agro-industry, private investment, "
        "recycling",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "producción, consumo, reciclaje, disposición, fuga ambiental",
        "production, consumption, recycling, disposal, environmental leakage",
    ),
    "policy_budget": 1,
    "policy_budget_text": bi(
        "Sección 9: mecanismos de financiamiento — canje de deuda por medio ambiente; fondos "
        "mundiales y cooperación internacional; presupuesto del Tesoro Público; fondos "
        "especiales de desarrollo regional y municipal. Programa 3: déficit de inversión "
        "estimado en US$ 100 millones; fondo rotativo de US$ 1 millón para PYMES; líneas de "
        "crédito COFIDE; Programa Nacional de Inversiones canalizando US$ 100 millones en 10 "
        "años.",
        "Section 9: financing mechanisms — debt-for-nature swap; global funds and international "
        "cooperation; public treasury budget; special regional and municipal development "
        "funds. Program 3: estimated investment deficit of US$100 million; US$1 million "
        "revolving fund for MSMEs; COFIDE credit lines; National Investment Program channeling "
        "US$100 million over 10 years.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Sección 2 (Ámbito de aplicación): el Plan se aplica a nivel nacional a todas las "
            "actividades de gestión y manejo de residuos sólidos desde la generación hasta la "
            "disposición final, incluyendo tránsito e internamiento; cubre residuos municipales "
            "y no municipales y los niveles nacional, regional y local.",
            "Section 2 (Scope): the Plan applies nationally to all solid-waste management "
            "activities from generation to final disposal, including transit and import; covers "
            "municipal and non-municipal waste and national, regional and local government levels.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Marco planificatorio vinculado a Ley N° 27314 y su Reglamento. CONAM/MINAM "
            "coordina; gobiernos locales y regionales ejecutan. Sin sanciones directas en el "
            "Plan; cumplimiento vía seguimiento y normativa sectorial.",
            "Planning framework linked to Law No. 27314 and its Regulation. CONAM/MINAM "
            "coordinates; local and regional governments implement. No direct penalties in the "
            "Plan; compliance through monitoring and sector regulation.",
        ),
        "comments": bi(
            "Instrumento de alcance nacional para todo el ciclo, incluidos plásticos en flujos "
            "municipales e industriales.",
            "National scope instrument for the full cycle, including plastics in municipal and "
            "industrial streams.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Principio 1.2 Prevención y minimización (Sección 5): se priorizan medidas para "
            "reducir la generación de residuos sólidos y su peligrosidad en producción y "
            "consumo.",
            "Principle 1.2 Prevention and minimization (Section 5): measures to reduce solid-waste "
            "generation and hazardousness in production and consumption are prioritized.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Principio rector orientador del Plan. Metas asociadas: 10% reducción per cápita en "
            "5 años; programas de producción limpia. Sin penalidades en el Plan mismo.",
            "Guiding principle of the Plan. Associated targets: 10% per-capita reduction in 5 "
            "years; cleaner-production programs. No penalties in the Plan itself.",
        ),
        "comments": bi(
            "Jerarquía de prevención aplicable a envases y plásticos.",
            "Prevention hierarchy applicable to packaging and plastics.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Principio 1.6 Reciclaje (Sección 5): se facilitará la valorización y recuperación "
            "directa de residuos, potenciando reaprovechamiento formal y mercados de productos "
            "recuperados.",
            "Principle 1.6 Recycling (Section 5): valorization and direct recovery of waste will "
            "be facilitated, strengthening formal reuse/recycling and markets for recovered "
            "products.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Línea de política 3.6 y programa de minimización/reciclaje. Metas de "
            "reaprovechamiento en capitales regionales (70% en 5 años) y aumento de "
            "comercialización de residuos (30%/40%).",
            "Policy line 3.6 and minimization/recycling program. Reuse targets in regional "
            "capitals (70% in 5 years) and increased waste commercialization (30%/40%).",
        ),
        "comments": bi(
            "Línea base: plásticos representan parte del 20,3% de materiales altamente "
            "reciclables en residuos municipales.",
            "Baseline: plastics are part of the 20.3% of highly recyclable materials in municipal "
            "waste.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Principios 1.8–1.10 (Sección 5): autosuficiencia y contaminador pagador; "
            "responsabilidad común pero diferenciada; producción limpia y responsabilidad "
            "empresarial ('de la cuna a la tumba' para residuos peligrosos).",
            "Principles 1.8–1.10 (Section 5): self-sufficiency and polluter pays; common but "
            "differentiated responsibility; cleaner production and corporate responsibility "
            "('cradle to grave' for hazardous waste).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Principios orientadores vinculados a programa 1.2 Producción limpia. Metas: 10,000 "
            "empresas incentivadas y 5,000 con ISO 14001 en 5 años. Implementación vía "
            "programas sectoriales.",
            "Guiding principles linked to Program 1.2 Cleaner Production. Targets: 10,000 "
            "incentivized firms and 5,000 with ISO 14001 in 5 years. Implementation through "
            "sector programs.",
        ),
        "comments": bi(
            "Antecedente de responsabilidad extendida del productor en política nacional de "
            "residuos.",
            "Precursor to extended producer responsibility in national waste policy.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Lineamiento específico 3.7 (Sección 5): promoción del manejo selectivo de residuos "
            "sólidos y admisión de manejo conjunto cuando no existan riesgos significativos.",
            "Specific policy line 3.7 (Section 5): promotion of selective management of solid "
            "waste and allowance of co-management when no significant sanitary or "
            "environmental risks arise.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Vinculado a metas de recolección selectiva (40% municipios provinciales en 5 años; "
            "100 ordenanzas en 4 años). Municipalidades y MINAM responsables.",
            "Linked to selective-collection targets (40% provincial municipalities in 5 years; 100 "
            "ordinances in 4 years). Municipalities and MINAM responsible.",
        ),
        "comments": bi(
            "Instrumento clave para segregación de plásticos y reciclables.",
            "Key instrument for segregation of plastics and recyclables.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Objetivo General 1 (Sección 6): promover y alcanzar calidad y cobertura universal "
            "de los servicios de manejo de residuos sólidos con gestión integral y sostenible "
            "para prevenir contaminación y proteger la salud.",
            "General Objective 1 (Section 6): promote and achieve universal quality and coverage "
            "of solid-waste management services through integrated and sustainable systems to "
            "prevent environmental contamination and protect public health.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Objetivo estratégico del Plan con indicadores de cobertura y disposición adecuada "
            "(95% en 10 años). Seguimiento anual CONAM.",
            "Strategic Plan objective with coverage and adequate-disposal indicators (95% in 10 "
            "years). Annual CONAM monitoring.",
        ),
        "comments": bi(
            "Servicio público de limpieza incluye residuos con contenido plástico.",
            "Public cleaning service includes waste with plastic content.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Objetivo General 2 (Sección 6): promover consumo sostenible, reducir al mínimo la "
            "generación de residuos y maximizar reutilización y reciclaje ambientalmente "
            "aceptables.",
            "General Objective 2 (Section 6): promote sustainable consumption, minimize waste "
            "generation and maximize environmentally acceptable reuse and recycling.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas educativas y de consumo responsable (70% centros educativos con contenidos "
            "de consumo sostenible; 40% consumidores responsables en 10 años).",
            "Educational and responsible-consumption targets (70% of schools with sustainable-"
            "consumption content; 40% responsible consumers in 10 years).",
        ),
        "comments": bi(
            "Objetivo central para reducción de plásticos de un solo uso y envases.",
            "Core objective for reducing single-use plastics and packaging.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Objetivo General 3 (Sección 6): fortalecer la gestión integral articulando "
            "instituciones competentes, responsabilidad empresarial, participación ciudadana y "
            "libre acceso a la información.",
            "General Objective 3 (Section 6): strengthen integrated management by coordinating "
            "competent institutions, corporate responsibility, citizen participation and free "
            "access to information.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Instancia intersectorial CONAM (50 directivas consensuadas en 2 años); SIGERSOL; "
            "participación ciudadana en capitales regionales.",
            "CONAM intersectoral instance (50 consensus directives in 2 years); SIGERSOL; citizen "
            "participation in regional capitals.",
        ),
        "comments": bi(
            "Gobernanza multisectorial del sistema de residuos.",
            "Multisector governance of the waste system.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Estrategia 1 (Sección 7): concertación de objetivos de comercio y sanitario-"
            "ambientales para la gestión de residuos sólidos.",
            "Strategy 1 (Section 7): harmonization of trade and health-environment objectives "
            "for solid-waste management.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.20,
        "instrument_implementation_text": bi(
            "Estrategia orientadora sin metas cuantificadas propias. Vinculada a línea 1.3 "
            "Negociación y comercio internacional.",
            "Guiding strategy without its own quantified targets. Linked to line 1.3 International "
            "Negotiation and Trade.",
        ),
        "comments": bi(
            "Competitividad exportadora y normas ambientales sobre residuos.",
            "Export competitiveness and environmental standards on waste.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 1.1 — Objetivo específico 1: fortalecer capacidades para segregación y "
            "recolección selectiva. Indicadores: 70% funcionarios municipales capacitados en 5 "
            "años; 80% centros educativos con currícula de recolección selectiva; 30% centros "
            "con programas de reciclaje en 3 años.",
            "Program 1.1 — Specific objective 1: strengthen capacities for segregation and "
            "selective collection. Indicators: 70% of municipal officials trained in 5 years; 80% "
            "of schools with selective-collection curriculum; 30% of schools with recycling "
            "programs in 3 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas cuantificadas con plazos. MINAM/CONAM y municipios; monitoreo vía Plan "
            "Operativo y evaluación anual.",
            "Quantified targets with deadlines. MINAM/CONAM and municipalities; monitoring via "
            "Operational Plan and annual evaluation.",
        ),
        "comments": bi(
            "Capacitación para recolección selectiva de plásticos y reciclables.",
            "Training for selective collection of plastics and recyclables.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Programa 1.1 — Objetivo específico 2: cultura de consumo y producción sostenible. "
            "Meta: disminuir producción de bienes no reciclables y peligrosos en 5 años; 70% "
            "centros educativos con contenidos de consumo sostenible.",
            "Program 1.1 — Specific objective 2: sustainable consumption and production culture. "
            "Target: reduce production of non-recyclable and hazardous goods in 5 years; 70% of "
            "schools with sustainable-consumption content.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "26 programas de educación ciudadana en 3 años. Programas académicos basados en "
            "Agenda 21.",
            "26 citizen-education programs in 3 years. Academic programs based on Agenda 21.",
        ),
        "comments": bi(
            "Reducción de generación de plásticos no reciclables.",
            "Reduction of non-recyclable plastic generation.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 1.1 — Objetivo específico 3: PIGARS y planes de manejo. Indicadores: 150 "
            "municipios provinciales con PIGARS en 5 años; 50,000 empresas e instituciones con "
            "planes de manejo en 5 años.",
            "Program 1.1 — Specific objective 3: PIGARS and management plans. Indicators: 150 "
            "provincial municipalities with PIGARS in 5 years; 50,000 firms and institutions with "
            "management plans in 5 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas cuantificadas de planificación local y sectorial. Municipios y generadores "
            "responsables.",
            "Quantified local and sectoral planning targets. Municipalities and generators "
            "responsible.",
        ),
        "comments": bi(
            "PIGARS instrumento municipal de gestión integral incluyendo fracciones plásticas.",
            "PIGARS municipal integrated-management instrument including plastic fractions.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Programa 1.2 — Objetivo específico 6: minimización y reciclaje. Meta: 10% reducción "
            "per cápita de residuos municipales en 5 años; 70% capitales regionales con planes "
            "de reaprovechamiento; 40% capitales con disminución de volumen de generación.",
            "Program 1.2 — Specific objective 6: minimization and recycling. Target: 10% per-"
            "capita reduction in municipal waste in 5 years; 70% of regional capitals with reuse "
            "plans; 40% of capitals with reduced generation volume.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas cuantificadas de minimización y valorización. Sectores manufactureros y "
            "mineros con programas de minimización (30%/50% en 6 años).",
            "Quantified minimization and valorization targets. Manufacturing and mining sectors "
            "with minimization programs (30%/50% in 6 years).",
        ),
        "comments": bi(
            "Reciclaje de plásticos, metales y papeles en flujo municipal.",
            "Recycling of plastics, metals and paper in the municipal stream.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 1.2 — Objetivo específico 5: salud ocupacional de manipuladores. Meta: "
            "100% de manipuladores cubiertos bajo programa de salud ocupacional en 5 años; "
            "erradicación de participación infantil en 10 años.",
            "Program 1.2 — Specific objective 5: occupational health for waste handlers. Target: "
            "100% of handlers covered by occupational health programs in 5 years; eradication "
            "of child participation in 10 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Meta cuantificada 100% en 5 años. Municipios, EPS-RS, EC-RS y generadores "
            "implementan programas.",
            "Quantified 100% target in 5 years. Municipalities, EPS-RS, EC-RS and generators "
            "implement programs.",
        ),
        "comments": bi(
            "Protección de recicladores que recuperan plásticos y otros materiales.",
            "Protection of recyclers recovering plastics and other materials.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Programa 1.2 — Objetivo específico 7: producción limpia. Metas: 10,000 empresas "
            "incentivadas y 5,000 con ISO 14001 en 5 años; 40% consumidores responsables en "
            "10 años.",
            "Program 1.2 — Specific objective 7: cleaner production. Targets: 10,000 incentivized "
            "firms and 5,000 with ISO 14001 in 5 years; 40% responsible consumers in 10 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Programa de incentivos a producción limpia. MINAM/sector productivo coordinan.",
            "Cleaner-production incentive program. MINAM/productive sectors coordinate.",
        ),
        "comments": bi(
            "Ecoeficiencia de envases y embalajes industriales con contenido plástico.",
            "Eco-efficiency of industrial packaging with plastic content.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": bi(
            "Programa 1.2 — Objetivo específico 8 / Programa 1.3: control transfronterizo según "
            "Convenio de Basilea. Meta: 80% de fronteras controlan comercio de residuos en 5 "
            "años.",
            "Program 1.2 — Specific objective 8 / Program 1.3: transboundary control under Basel "
            "Convention. Target: 80% of borders control waste trade in 5 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "DIGESA y ADUANAS responsables. Normas de control en aduanas de ciudades "
            "fronterizas.",
            "DIGESA and CUSTOMS responsible. Control standards at border-city customs posts.",
        ),
        "comments": bi(
            "Control de movimientos de residuos plásticos y peligrosos.",
            "Control of movements of plastic and hazardous waste.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 2.1 — Objetivo específico 9: marco normativo nacional/regional/local. "
            "Meta: normar todos los aspectos técnicos y legales en 5 años; 100 municipios "
            "provinciales con ordenanzas de recolección selectiva en 4 años.",
            "Program 2.1 — Specific objective 9: national/regional/local regulatory framework. "
            "Target: regulate all technical and legal aspects in 5 years; 100 provincial "
            "municipalities with selective-collection ordinances in 4 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas de normalización y ordenanzas municipales. Gobiernos regionales y locales.",
            "Standardization and municipal ordinance targets. Regional and local governments.",
        ),
        "comments": bi(
            "Ordenanzas habilitan recolección selectiva de plásticos.",
            "Ordinances enable selective collection of plastics.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 2.3 — Objetivo específico 12: registro y supervisión de EPS-RS y EC-RS. "
            "Meta: 100% registradas y supervisadas en 5 años; 100% de instituciones contratan "
            "solo con empresas registradas en 3 años.",
            "Program 2.3 — Specific objective 12: registration and supervision of EPS-RS and "
            "EC-RS. Target: 100% registered and supervised in 5 years; 100% of institutions "
            "contract only with registered firms in 3 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "DIGESA administra registro. Contratación obligatoria con empresas registradas.",
            "DIGESA administers registry. Mandatory contracting with registered firms.",
        ),
        "comments": bi(
            "Formalización del sector de reciclaje comercial de plásticos y otros materiales.",
            "Formalization of commercial recycling of plastics and other materials.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Programa 2.3 — Objetivo específico 13: mercados de reaprovechamiento. Meta: +30% "
            "volumen comercializado en capitales regionales en 5 años y +40% en residuos no "
            "municipales en 10 años; Bolsas de Residuos en 6 regiones en 3 años.",
            "Program 2.3 — Specific objective 13: recovery markets. Target: +30% commercialized "
            "volume in regional capitals in 5 years and +40% for non-municipal waste in 10 years; "
            "Waste Exchanges in 6 regions in 3 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "SIGERSOL informa precios de mercado; 4 boletines anuales de precios de residuos.",
            "SIGERSOL reports market prices; 4 annual waste-price bulletins.",
        ),
        "comments": bi(
            "Mercados de plásticos, metales y papel recuperados.",
            "Markets for recovered plastics, metals and paper.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 2.4 — Objetivo específico 16: sistema de información SIGERSOL. Meta: 80% "
            "de actores con acceso en 2 años; información empleada para decisiones en 3 años.",
            "Program 2.4 — Specific objective 16: SIGERSOL information system. Target: 80% of "
            "actors with access in 2 years; information used for decisions in 3 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Sistema nacional de información de residuos sólidos. MINAM/CONAM coordinan; "
            "actualización anual de tecnologías.",
            "National solid-waste information system. MINAM/CONAM coordinate; annual technology "
            "updates.",
        ),
        "comments": bi(
            "Monitoreo de flujos de residuos incluyendo fracciones plásticas reciclables.",
            "Monitoring of waste flows including recyclable plastic fractions.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 2.4 — Objetivo específico 15: vigilancia sanitaria y ambiental. Meta: "
            "sistema articulado a nivel nacional en 5 años; 100% regiones con vigilancia; "
            "auditores acreditados en 100% regiones en 5 años.",
            "Program 2.4 — Specific objective 15: sanitary and environmental surveillance. Target: "
            "nationally articulated system in 5 years; 100% of regions conducting surveillance; "
            "accredited auditors in 100% of regions in 5 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "CONAM coordina vigilancia multisectorial. Sistema de acreditación de auditores en 2 "
            "años.",
            "CONAM coordinates multisector surveillance. Auditor accreditation system in 2 years.",
        ),
        "comments": bi(
            "Fiscalización de manejo inadecuado de residuos en vertederos y botaderos.",
            "Oversight of improper waste handling at dumpsites.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Programa 3.1 — Objetivo específico 17: fondos para PYMES y transferencia "
            "tecnológica. Meta: fondo rotativo de US$ 1 millón en 2 años; líneas COFIDE en 4 "
            "años.",
            "Program 3.1 — Specific objective 17: funds for MSMEs and technology transfer. "
            "Target: US$1 million revolving fund in 2 years; COFIDE credit lines in 4 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Financiamiento identificado en Sección 9. Canje de deuda por medio ambiente para "
            "micro y pequeña empresa.",
            "Financing identified in Section 9. Debt-for-nature swap for micro and small "
            "enterprises.",
        ),
        "comments": bi(
            "Inversión en infraestructura de reciclaje y disposición final.",
            "Investment in recycling and final-disposal infrastructure.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Programa 3.2 — Objetivo específico 18: Programa Nacional de Inversiones. Meta: "
            "100% proyectos con acceso a financiamiento en 5 años; canalización de US$ 100 "
            "millones en 10 años.",
            "Program 3.2 — Specific objective 18: National Investment Program. Target: 100% of "
            "projects with financing access in 5 years; channeling of US$100 million in 10 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Déficit sectorial estimado US$ 100 millones. BID, CAF y cooperación internacional "
            "como fuentes.",
            "Sector deficit estimated at US$100 million. IDB, CAF and international cooperation "
            "as sources.",
        ),
        "comments": bi(
            "Inversión en rellenos sanitarios, transferencia y plantas de valorización.",
            "Investment in sanitary landfills, transfer stations and valorization plants.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Programa 3.2 — Objetivo específico 19: infraestructura y disposición adecuada. "
            "Meta: 100% infraestructura existente opera sanitariamente en 10 años; 95% de "
            "residuos municipales y no municipales dispuestos adecuadamente en 10 años.",
            "Program 3.2 — Specific objective 19: infrastructure and adequate disposal. Target: "
            "100% of existing infrastructure operating sanitarily in 10 years; 95% of municipal "
            "and non-municipal waste adequately disposed in 10 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Metas cuantificadas de disposición final. PAMAS para municipios e instituciones (50% "
            "municipios y 100% no municipales en 8 años).",
            "Quantified final-disposal targets. PAMAS for municipalities and institutions (50% "
            "municipalities and 100% non-municipal in 8 years).",
        ),
        "comments": bi(
            "Cierre de botaderos y vertido incontrolado de plásticos y residuos mixtos.",
            "Closure of dumps and uncontrolled dumping of plastics and mixed waste.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Programa 3.2 — Objetivo específico 20: recuperación de áreas degradadas. Meta: "
            "100% botaderos municipales clausurados/recuperados y 50% botaderos no municipales "
            "en 10 años; 100% municipios con PAMAS en 3 años.",
            "Program 3.2 — Specific objective 20: recovery of degraded areas. Target: 100% of "
            "municipal dumps closed/recovered and 50% of non-municipal dumps in 10 years; 100% "
            "of municipalities with PAMAS in 3 years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "80% PAMAS con financiamiento. Recuperación ambiental, social y económica.",
            "80% of PAMAS with financing. Environmental, social and economic recovery.",
        ),
        "comments": bi(
            "Restauración de sitios con acumulación persistente de plásticos.",
            "Restoration of sites with persistent plastic accumulation.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Programa 3.2 — Objetivo específico 21: segregación y recolección selectiva "
            "municipal. Meta: porcentaje significativo de residuos segregados en 5 años; 40% "
            "municipios provinciales con sistemas de recolección selectiva.",
            "Program 3.2 — Specific objective 21: municipal segregation and selective collection. "
            "Target: significant share of waste segregated in 5 years; 40% of provincial "
            "municipalities with selective-collection systems.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Meta cuantificada 40% en 5 años. Planes municipales de recolección selectiva.",
            "Quantified 40% target in 5 years. Municipal selective-collection plans.",
        ),
        "comments": bi(
            "Instrumento operativo principal para captura de plásticos en fuente.",
            "Main operational instrument for at-source plastic capture.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Sección 9 — Mecanismos de financiamiento: (1) canje de deuda por medio ambiente; "
            "(2) fondos mundiales y cooperación; (3) presupuesto del Tesoro Público; (4) fondos "
            "especiales regionales y municipales.",
            "Section 9 — Financing mechanisms: (1) debt-for-nature swap; (2) global funds and "
            "cooperation; (3) public treasury budget; (4) special regional and municipal funds.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Mecanismos identificados para líneas de acción del Plan. Sin montos obligatorios "
            "vinculantes en la RM; programación financiera sectorial.",
            "Mechanisms identified for Plan action lines. No binding mandatory amounts in the "
            "MR; sectoral financial programming.",
        ),
        "comments": bi(
            "Financiamiento de servicios de limpieza y infraestructura de residuos.",
            "Financing of cleaning services and waste infrastructure.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Sección 10 — Seguimiento y revisión: reuniones de planificación operativa anual "
            "(último trimestre); evaluación anual del Plan; monitoreo semestral; informes anuales "
            "de autoridades sectoriales y municipalidades al CONAM (Art. 120 Reglamento Ley "
            "27314).",
            "Section 10 — Monitoring and review: annual operational planning meetings (last "
            "quarter); annual Plan evaluation; semiannual monitoring; annual reports from sector "
            "authorities and municipalities to CONAM (Art. 120 Regulation of Law 27314).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "CONAM convoca evaluación anual; CARs coordinan a nivel regional. Sistema de "
            "información para participación ciudadana.",
            "CONAM convenes annual evaluation; CARs coordinate regionally. Information system "
            "for citizen participation.",
        ),
        "comments": bi(
            "Monitoreo de cumplimiento de metas incluyendo reciclaje y disposición.",
            "Monitoring of target compliance including recycling and disposal.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": bi(
            "Lineamiento específico 3.16 (Sección 5): acciones para evitar contaminación del "
            "medio acuático, eliminando arrojo de residuos sólidos en cuerpos o cursos de agua.",
            "Specific policy line 3.16 (Section 5): actions to avoid aquatic-environment "
            "contamination by eliminating dumping of solid waste in water bodies or courses.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Línea de política del Plan. Situación base: 19,6% residuos vertidos al ambiente "
            "(ríos y playas principales receptores). Sin sanción directa en el Plan.",
            "Plan policy line. Baseline: 19.6% of waste released to the environment (rivers and "
            "beaches main receptors). No direct penalty in the Plan.",
        ),
        "comments": bi(
            "Fuga de plásticos a ríos y playas — problema de contaminación persistente.",
            "Plastic leakage to rivers and beaches — persistent pollution problem.",
        ),
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
        "Approved: 30 June 2016 (RM 191-2016-MINAM) | Plan period: 2016–2024\n"
        "Under update since 2023 (2nd deliverable published 27 December 2023); "
        "new PLANRES pending as of 2026\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of the National Comprehensive Solid Waste "
        "Management Plan (PLANRES) 2016–2024, approved by Ministerial Resolution No. "
        "191-2016-MINAM."
    )

    sections = [
        (
            "Introduction — Plan purpose",
            "El objetivo del Plan es reducir la producción nacional de residuos sólidos y controlar "
            "los riesgos sanitarios y ambientales asociados... incrementar la calidad y cobertura "
            "de los servicios... recolección selectiva; reducir, recuperar, reutilizar y reciclar "
            "los residuos...",
            "The Plan aims to reduce national solid-waste production and control associated "
            "sanitary and environmental risks, including permanent environmental education, "
            "citizen participation, improved service quality and coverage including selective "
            "collection, and reducing, recovering, reusing and recycling waste with safe final "
            "disposal of non-recovered fractions.",
        ),
        (
            "Section 2 — Scope",
            "El Plan Nacional... su ámbito de acción se extiende a nivel nacional y se aplica a "
            "todas las actividades... desde la generación hasta la disposición final...",
            "The National Plan applies nationwide to all solid-waste management activities from "
            "generation to final disposal, including transit and import, for municipal and non-"
            "municipal waste at national, regional and local levels.",
        ),
        (
            "Section 4 — Baseline (plastics in waste composition)",
            "los materiales altamente reciclables como el papel, cartón, plásticos, metales... "
            "representan el 20,3%... El reciclaje alcanza al 14,7%... corresponde principalmente "
            "a papeles, plásticos y metales...",
            "Highly recyclable materials including paper, cardboard, plastics and metals represent "
            "20.3% of municipal waste by weight. Recycling reaches 14.7% of generated waste, "
            "mainly papers, plastics and metals; metal recycling is most efficient.",
        ),
        (
            "Section 5 — Prevention and recycling principles",
            "1.2 Prevención y minimización... 1.6 Reciclaje... Se facilitará a través de la "
            "valorización de los residuos, la recuperación directa...",
            "Principles 1.2 Prevention and minimization and 1.6 Recycling: waste valorization and "
            "direct recovery will be facilitated, strengthening formal reuse/recycling and markets "
            "for recovered products.",
        ),
        (
            "Section 6 — General objectives",
            "Objetivo 1: calidad y cobertura universal... Objetivo 2: consumo sostenible y "
            "maximizar reutilización y reciclaje... Objetivo 3: fortalecer la gestión integral...",
            "Objective 1: universal quality and coverage of waste services. Objective 2: "
            "sustainable consumption and maximum reuse/recycling. Objective 3: strengthen "
            "integrated management coordinating institutions, corporate responsibility, citizen "
            "participation and information access.",
        ),
        (
            "Program 1.1 — Selective collection capacity targets",
            "En 5 años 70% de los funcionarios municipales... han recibido capacitación... "
            "80% de centros educativos... recolección selectiva...",
            "Within 5 years: 70% of municipal officials linked to solid waste trained; 80% of "
            "public schools incorporate and implement selective-collection curriculum content.",
        ),
        (
            "Program 1.2 — Minimization and occupational health",
            "reducción de la producción per cápita... alcanza el 10%... 100% de los "
            "manipuladores... programa de salud Ocupacional...",
            "Within 5 years in main cities: 10% per-capita reduction in municipal waste generation; "
            "100% of waste handlers covered by occupational health programs; eradication of child "
            "labor in waste handling within 10 years.",
        ),
        (
            "Program 2.1 — Municipal selective-collection ordinances",
            "En 4 años 100 municipios provinciales cuentan con ordenanzas para promover la "
            "recolección selectiva.",
            "Within 4 years 100 provincial municipalities have ordinances promoting selective "
            "collection.",
        ),
        (
            "Program 3.2 — Adequate disposal and dump closure",
            "En 10 años el 95% de los residuos... se dispone adecuadamente... 100% de los "
            "botaderos municipales clausurados, recuperados y restaurados...",
            "Within 10 years: 95% of municipal and non-municipal waste adequately disposed; 100% "
            "of municipal dumps closed, recovered and restored; 40% of provincial municipalities "
            "with selective-collection systems within 5 years.",
        ),
        (
            "Section 9 — Financing mechanisms",
            "Canje de deuda por medio ambiente... Fondos Mundiales... Presupuesto del Tesoro "
            "Público... Fondos especiales de desarrollo regional y municipal.",
            "Financing mechanisms: debt-for-nature swap; global sustainable-development funds and "
            "international cooperation; public treasury budget; special regional and municipal "
            "development funds. Investment program targets US$100 million over 10 years.",
        ),
        (
            "Section 10 — Monitoring and review",
            "Reuniones... evaluación anual... monitoreo semestral... informes anuales... al "
            "CONAM...",
            "Monitoring: annual operational planning meetings; annual Plan evaluation convened by "
            "CONAM; semiannual operational monitoring; annual sector and municipal waste-"
            "management reports to CONAM under Regulation Art. 120.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: PLANRES was approved under CONAM (now integrated into MINAM). It implements "
        "obligations under the former General Solid Waste Law (Law No. 27314, later replaced by "
        "DL 1278). The 2016–2024 plan remains the operative national plan while an updated "
        "PLANRES is under development (consultation deliverables published December 2023)."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


BRANCH = "cursor/peru-planres-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"


def main() -> None:
    excel = OUTPUT_DIR / "Peru_PLANRES_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_PLANRES_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(f"  Excel: {REPO_RAW}/output/peru-planres/Peru_PLANRES_4P_Index_Coding.xlsx")
    print(f"  Word:  {REPO_RAW}/output/peru-planres/Peru_PLANRES_English_Translation.docx")


if __name__ == "__main__":
    main()
