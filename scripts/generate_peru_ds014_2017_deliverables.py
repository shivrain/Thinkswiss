#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 014-2017-MINAM (DL 1278 Regulation)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-014-2017-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ds-014-2017-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    "Decreto Supremo N.° 014-2017-MINAM – Reglamento del Decreto Legislativo N.° 1278"
)
POLICY_NAME_EN = (
    "Supreme Decree No. 014-2017-MINAM – Regulation of Legislative Decree No. 1278 "
    "(Comprehensive Solid Waste Management Law)"
)
POLICY_URL = "https://www.gob.pe/institucion/minam/normas-legales/3695-014-2017-minam"


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2024,
    "policy_objective": bi(
        "Reglamento de desarrollo del Decreto Legislativo N.° 1278 que reglamenta la gestión "
        "integral de residuos sólidos: eficiencia de materiales, minimización en la fuente, "
        "instrumentos de gestión (PLANRES, planes municipales, programas de segregación y "
        "recolección selectiva, SIGERSOL), servicio de limpieza pública, valorización "
        "(incluidos plásticos y otros inorgánicos no peligrosos), infraestructura de "
        "disposición final, régimen especial de bienes priorizados (REP), registro de EO-RS "
        "y fiscalización OEFA.",
        "Implementing regulation of Legislative Decree No. 1278 on integrated solid-waste "
        "management: material efficiency, source minimization, management instruments "
        "(PLANRES, municipal plans, source-segregation and selective-collection programs, "
        "SIGERSOL), public cleaning service, valorization (including plastics and other "
        "non-hazardous inorganics), final-disposal infrastructure, special regime for "
        "prioritized goods (EPR), EO-RS registry and OEFA enforcement.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "Art. 32: frecuencia mínima de recolección de residuos no aprovechables tres (03) "
        "veces por semana. Art. 34: servicio municipal hasta 150 litros/día por fuente; "
        "derechos adicionales hasta 500 litros. Art. 38: metas nacionales de valorización "
        "en PLANAA y PLANRES. Art. 39: almacenamiento temporal en transferencia máximo "
        "doce (12) horas. Art. 55: almacenamiento de residuos peligrosos máximo doce (12) "
        "meses. Art. 10: reporte anual de planes municipales último día hábil de marzo. "
        "Art. 13: reportes SIGERSOL (febrero, trimestral, abril). Enmiendas DS 001-2022-"
        "MINAM (Arts. 34, 38, 39) y DS 002-2024-MINAM (4 abr. 2024).",
        "Art. 32: minimum collection frequency for non-recoverable waste three (03) times per "
        "week. Art. 34: municipal service up to 150 L/day per source; additional fees up to "
        "500 L. Art. 38: national valorization targets in PLANAA and PLANRES. Art. 39: "
        "temporary storage at transfer max twelve (12) hours. Art. 55: hazardous-waste "
        "storage max twelve (12) months. Art. 10: annual municipal plan report last business "
        "day of March. Art. 13: SIGERSOL reporting (February, quarterly, April). Amended by "
        "DS 001-2022-MINAM (Arts. 34, 38, 39) and DS 002-2024-MINAM (4 Apr 2024).",
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Decreto Supremo N.° 014-2017-MINAM aprobado por el Poder Ejecutivo (21 dic. 2017), "
        "refrendado por MINAM y sectores vinculados. Reglamento de desarrollo de ley delegada "
        "(DL 1278). Enmendado por DS N.° 001-2022-MINAM (9 ene. 2022) y DS N.° 002-2024-"
        "MINAM (4 abr. 2024). Clasificado 0.75 (reglamento/decreto supremo ejecutivo).",
        "Supreme Decree No. 014-2017-MINAM approved by the Executive (21 Dec 2017), "
        "countersigned by MINAM and related sectors. Implementing regulation of delegated "
        "law DL 1278. Amended by DS No. 001-2022-MINAM (9 Jan 2022) and DS No. 002-2024-"
        "MINAM (4 Apr 2024). Classified 0.75 (executive supreme-decree regulation).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "ambiente, gobiernos municipales y regionales, industria, producción, consumo, salud, "
        "construcción, comercio, transporte, reciclaje, envases y embalajes",
        "environment, municipal and regional governments, industry, production, consumption, "
        "health, construction, trade, transport, recycling, packaging",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "producción, consumo, reciclaje, disposición, fuga ambiental",
        "production, consumption, recycling, disposal, environmental leakage",
    ),
    "policy_budget": 1,
    "policy_budget_text": bi(
        "Art. 3 (aprobación): aplicación sujeta a presupuesto institucional. Art. 44: promoción "
        "de inversión pública y privada en infraestructura de residuos (SNPMGI, Ley 29230, "
        "APP DL 1224). Art. 34: cobros diferenciados por volumen de residuos. Tasas municipales "
        "de limpieza pública vinculadas al servicio.",
        "Art. 3 (approval): application subject to institutional budget. Art. 44: promotion of "
        "public and private investment in waste infrastructure (SNPMGI, Law 29230, PPP DL 1224). "
        "Art. 34: differentiated charges by waste volume. Municipal cleaning-fee revenue linked "
        "to service provision.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 1: objeto del Reglamento — reglamentar el DL 1278 para maximizar eficiencia "
            "de materiales; regular minimización en fuente, valorización material y energética, "
            "disposición final adecuada y sostenibilidad del servicio de limpieza pública.",
            "Art. 1: purpose — regulate DL 1278 to maximize material efficiency; govern source "
            "minimization, material and energy valorization, adequate final disposal and "
            "sustainability of the public cleaning service.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Marco reglamentario obligatorio de cumplimiento nacional. MINAM, municipalidades, "
            "OEFA y sectores competentes. Sin sanción específica en Art. 1.",
            "Mandatory national regulatory framework. MINAM, municipalities, OEFA and competent "
            "sectors. No specific penalty in Art. 1.",
        ),
        "instrument_score_notes": bi(
            "Instrumento marco del reglamento de gestión integral incluyendo plásticos.",
            "Overarching regulation framework for integrated management including plastics.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 3: sectores bajo SEIA establecerán disposiciones para uso eficiente de "
            "materias primas e insumos, en coordinación con MINAM, para reducir impactos "
            "ambientales durante el ciclo de vida del producto.",
            "Art. 3: SEIA sectors shall issue provisions for efficient use of raw materials and "
            "inputs, coordinating with MINAM, to reduce environmental impacts across the "
            "product life cycle.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Autoridades sectoriales designadas. Coordinación MINAM. Disposiciones sectoriales "
            "pendientes de emisión por sector.",
            "Sector authorities designated. MINAM coordination. Sectoral provisions to be issued "
            "by each sector.",
        ),
        "instrument_score_notes": bi(
            "Ecoeficiencia y diseño de productos/envases con contenido plástico.",
            "Eco-efficiency and design of products/packaging with plastic content.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 7: generadores no municipales deben incluir en el Plan de Minimización y "
            "Manejo estrategias preventivas orientadas a la minimización en la fuente.",
            "Art. 7: non-municipal generators must include preventive strategies for source "
            "minimization in the Minimization and Management Plan.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Obligación en plan de manejo dentro del IGA. Fiscalización OEFA/sectorial.",
            "Obligation in management plan within EIA instrument. OEFA/sectoral enforcement.",
        ),
        "instrument_score_notes": bi(
            "Minimización de generación de plásticos y otros residuos en fuente industrial.",
            "Minimization of plastic and other waste generation at industrial source.",
        ),
    },
    {
        "instrument_type": 0.50,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 9: PLANRES — instrumento nacional con metas del PLANAA; aprobado por "
            "Decreto Supremo a propuesta del MINAM; actualización cada cinco (05) años.",
            "Art. 9: PLANRES — national instrument with PLANAA targets; approved by Supreme "
            "Decree on MINAM proposal; updated every five (05) years.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "MINAM formula y propone; DS aprueba. Metas de valorización vinculadas Art. 38.",
            "MINAM drafts and proposes; Supreme Decree approves. Valorization targets linked "
            "via Art. 38.",
        ),
        "instrument_score_notes": bi(
            "Plan nacional con metas de reciclaje/valorización de plásticos y otros materiales.",
            "National plan with recycling/valorization targets for plastics and other materials.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 10: Plan Provincial de Gestión de Residuos Sólidos Municipales y Plan "
            "Distrital de Manejo — planificación municipal alineada al PLANRES; actualización "
            "cada cinco (05) años; reporte anual a MINAM y OEFA último día hábil de marzo.",
            "Art. 10: Provincial Integrated Municipal Waste Management Plan and District "
            "Management Plan — municipal planning aligned with PLANRES; updated every five "
            "(05) years; annual report to MINAM and OEFA last business day of March.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben contener' diagnóstico, metas y plan de acción. Plazo reporte marzo. "
            "MINAM emite guías técnicas. OEFA recibe reportes.",
            "'Must contain' diagnosis, targets and action plan. March reporting deadline. MINAM "
            "issues technical guides. OEFA receives reports.",
        ),
        "instrument_score_notes": bi(
            "PMR/PIGARS municipales para gestión de fracciones plásticas reciclables.",
            "Municipal PMR plans for management of recyclable plastic fractions.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 11: Programa de Segregación en la Fuente y Recolección Selectiva — "
            "instrumento técnico municipal con estrategias de segregación y diseño de "
            "recolección selectiva, incluyendo participación de recicladores formalizados.",
            "Art. 11: Source Segregation and Selective Collection Program — municipal technical "
            "instrument with segregation strategies and selective-collection design, including "
            "participation of formalized recycler organizations.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Municipalidades elaboran programa. Vinculado Ley 29419 y recicladores. "
            "Obligatorio establecer progresivamente (Art. 28).",
            "Municipalities prepare program. Linked to Law 29419 and recyclers. Must be "
            "established progressively (Art. 28).",
        ),
        "instrument_score_notes": bi(
            "Instrumento operativo clave para segregación de plásticos en fuente.",
            "Key operational instrument for at-source plastic segregation.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 13: SIGERSOL — MINAM administra registro y difusión de información; "
            "municipalidades reportan hasta último día hábil de febrero; EO-RS informe "
            "trimestral; generadores no municipales declaración anual primeros 15 días "
            "hábiles de abril.",
            "Art. 13: SIGERSOL — MINAM administers registration and dissemination of "
            "information; municipalities report by last business day of February; EO-RS "
            "quarterly operator report; non-municipal generators annual declaration first "
            "15 business days of April.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Están obligados a registrar información'. Plazos calendario definidos. MINAM "
            "consolida Informe Nacional del Estado del Ambiente.",
            "'Obliged to register information'. Defined calendar deadlines. MINAM consolidates "
            "National State of Environment Report.",
        ),
        "instrument_score_notes": bi(
            "Monitoreo de composición y disposición incluyendo plásticos municipales.",
            "Monitoring of composition and disposal including municipal plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 19: el generador municipal 'debe realizar' la segregación según "
            "características físicas, químicas y biológicas para facilitar valorización o "
            "disposición final; municipalidades regulan en marco del Programa de Segregación.",
            "Art. 19: municipal generators 'must carry out' segregation by physical, chemical "
            "and biological characteristics to facilitate valorization or final disposal; "
            "municipalities regulate under the Segregation Program.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Debe realizar' segregación. Permitida solo en fuente, centros de acopio y "
            "plantas autorizadas. Infracciones tipificadas Art. 135.",
            "'Must carry out' segregation. Allowed only at source, collection centers and "
            "authorized plants. Infractions typified Art. 135.",
        ),
        "instrument_score_notes": bi(
            "Obligación de segregar plásticos y reciclables en fuente domiciliaria.",
            "Obligation to segregate plastics and recyclables at household source.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 21–22: municipalidades responsables del servicio de limpieza pública "
            "(barrido, recolección, transporte, transferencia, valorización y disposición final); "
            "prestación continua, regular, permanente y obligatoria; normas técnicas MINAM "
            "de cumplimiento obligatorio.",
            "Art. 21–22: municipalities responsible for public cleaning service (sweeping, "
            "collection, transport, transfer, valorization and final disposal); continuous, "
            "regular, permanent and mandatory provision; MINAM technical standards mandatory.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Son responsables' y 'debe garantizar'. Contratos con EO-RS con penalidades (Art. "
            "23). OEFA fiscaliza.",
            "'Are responsible' and 'must guarantee'. EO-RS contracts with penalties (Art. 23). "
            "OEFA supervises.",
        ),
        "instrument_score_notes": bi(
            "Servicio público cubre residuos con contenido plástico municipal.",
            "Public service covers municipal waste with plastic content.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 28: municipalidades deben establecer progresivamente Programas de Segregación "
            "y Recolección Selectiva con rutas, horarios y frecuencias; recolección selectiva "
            "por municipalidades, EO-RS o recicladores formalizados (Ley 29419).",
            "Art. 28: municipalities must progressively establish Segregation and Selective "
            "Collection Programs with routes, schedules and frequencies; selective collection "
            "by municipalities, EO-RS or formalized recyclers (Law 29419).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben establecer progresivamente'. Integración recicladores formalizados. "
            "Programa Art. 11 obligatorio.",
            "'Must progressively establish'. Integration of formalized recyclers. Art. 11 "
            "program required.",
        ),
        "instrument_score_notes": bi(
            "Recolección selectiva de plásticos, papel, metal y vidrio.",
            "Selective collection of plastics, paper, metal and glass.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 32: frecuencia mínima de recolección y transporte de residuos no "
            "aprovechables de tres (03) veces por semana.",
            "Art. 32: minimum frequency of collection and transport of non-recoverable waste "
            "of three (03) times per week.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Meta cuantificada mínima 3x/semana. Municipalidades establecen frecuencia según "
            "Art. 31. Incumplimiento sujeto a infracciones.",
            "Quantified minimum target 3x/week. Municipalities set frequency per Art. 31. "
            "Non-compliance subject to infractions.",
        ),
        "instrument_score_notes": bi(
            "Frecuencia mínima para residuos no valorizados post-segregación.",
            "Minimum frequency for non-valorized waste after segregation.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 34: municipalidades garantizan recolección hasta 150 litros diarios por "
            "fuente; derechos adicionales por volumen 150–500 litros; sobre 500 litros el "
            "generador contrata EO-RS.",
            "Art. 34: municipalities guarantee collection up to 150 litres per day per source; "
            "additional fees for 150–500 litres; above 500 litres generator must contract EO-RS.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben garantizar' y 'debe contratar'. Umbrales cuantificados 150/500 L. Enmendado "
            "DS 001-2022-MINAM.",
            "'Must guarantee' and 'must contract'. Quantified thresholds 150/500 L. Amended by "
            "DS 001-2022-MINAM.",
        ),
        "instrument_score_notes": bi(
            "Tarifas diferenciadas por volumen de residuos incluyendo envases.",
            "Differentiated fees by waste volume including packaging.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 35: centros de acopio para residuos inorgánicos no peligrosos recuperados "
            "en programas de segregación y recolección selectiva; autorizados por municipalidad.",
            "Art. 35: collection centers for non-hazardous inorganic waste recovered under "
            "segregation and selective-collection programs; authorized by municipality.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Infraestructura para acondicionamiento de plásticos y otros inorgánicos. Art. 101 "
            "regula actividades (clasificación, compactación, embalaje).",
            "Infrastructure for conditioning plastics and other inorganics. Art. 101 governs "
            "activities (sorting, compaction, packaging).",
        ),
        "instrument_score_notes": bi(
            "Centros de acopio para plásticos recuperados en recolección selectiva.",
            "Collection centers for plastics recovered through selective collection.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 36: la valorización de residuos municipales 'debe priorizarse' frente a la "
            "disposición final; otras operaciones distintas a DL 1278 Art. 48 requieren "
            "opinión previa favorable MINAM.",
            "Art. 36: valorization of municipal waste 'must be prioritized' over final disposal; "
            "other operations beyond DL 1278 Art. 48 require prior favorable MINAM opinion.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Debe priorizarse' — jerarquía vinculante. MINAM opina sobre nuevas tecnologías. "
            "OEFA fiscaliza.",
            "'Must be prioritized' — binding hierarchy. MINAM opines on new technologies. OEFA "
            "enforces.",
        ),
        "instrument_score_notes": bi(
            "Prioridad de valorización material de plásticos y otros reciclables.",
            "Priority of material valorization of plastics and other recyclables.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 38: metas nacionales de valorización establecidas por MINAM en PLANAA y "
            "PLANRES; cumplimiento verificado con información SIGERSOL de municipalidades.",
            "Art. 38: national valorization targets set by MINAM in PLANAA and PLANRES; compliance "
            "verified using municipal SIGERSOL information.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "MINAM establece metas. Verificación vía SIGERSOL. Enmendado DS 001-2022-MINAM.",
            "MINAM sets targets. Verification via SIGERSOL. Amended by DS 001-2022-MINAM.",
        ),
        "instrument_score_notes": bi(
            "Metas nacionales de valorización incluyen fracciones plásticas.",
            "National valorization targets include plastic fractions.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 39: prohibido almacenamiento temporal en vehículo de transferencia por más "
            "de doce (12) horas; prohibido trasbordo fuera de plantas de transferencia.",
            "Art. 39: temporary storage in transfer vehicle prohibited beyond twelve (12) hours; "
            "transshipment outside transfer plants prohibited.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Bajo ninguna circunstancia' y 'está prohibido'. Plazo 12 horas cuantificado. "
            "Enmendado DS 001-2022-MINAM. Infracciones Art. 135.",
            "'Under no circumstances' and 'is prohibited'. Quantified 12-hour limit. Amended by "
            "DS 001-2022-MINAM. Infractions Art. 135.",
        ),
        "instrument_score_notes": bi(
            "Control de transferencia de residuos mixtos y reciclables.",
            "Control of transfer of mixed and recyclable waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Art. 41–42: disposición final en rellenos sanitarios; residuos peligrosos, no "
            "peligrosos y de construcción en celdas diferenciadas según programa de segregación.",
            "Art. 41–42: final disposal in sanitary landfills; hazardous, non-hazardous and "
            "construction waste in differentiated cells per segregation program.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Infraestructura con IGA aprobado. Condiciones técnicas Título IX. OEFA "
            "fiscalización.",
            "Infrastructure with approved EIA instrument. Technical conditions Title IX. OEFA "
            "enforcement.",
        ),
        "instrument_score_notes": bi(
            "Disposición final de rechazos plásticos no valorizados en rellenos sanitarios.",
            "Final disposal of non-valorized plastic rejects in sanitary landfills.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 48: obligaciones del generador no municipal — manejar residuos, registro "
            "interno, contratar EO-RS, valorización como primera opción, declaración anual y "
            "manifiesto peligrosos vía SIGERSOL.",
            "Art. 48: non-municipal generator obligations — manage waste, internal registry, "
            "contract EO-RS, valorization as first option, annual declaration and hazardous "
            "manifest via SIGERSOL.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Son obligaciones' y 'deben' en Art. 48.1. Plazos SIGERSOL abril/trimestral. "
            "Sanciones OEFA Art. 135.",
            "'Are obligations' and 'must' in Art. 48.1. SIGERSOL April/quarterly deadlines. OEFA "
            "sanctions Art. 135.",
        ),
        "instrument_score_notes": bi(
            "Obligaciones industriales para residuos de envases y plásticos no municipales.",
            "Industrial obligations for packaging and non-municipal plastic waste.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 51: generadores no municipales 'están obligados a segregar' residuos en la "
            "fuente.",
            "Art. 51: non-municipal generators 'are obliged to segregate' waste at source.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Están obligados a segregar'. Condiciones en IGA. Tipificado en tabla infracciones "
            "Art. 135.",
            "'Are obliged to segregate'. Conditions in EIA instrument. Typified in Art. 135 "
            "infraction table.",
        ),
        "instrument_score_notes": bi(
            "Segregación industrial de plásticos y otras fracciones.",
            "Industrial segregation of plastics and other fractions.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Art. 84–86: régimen especial de bienes priorizados — productores implementan "
            "sistemas de manejo individual o colectivo (REP); DS regula bienes, objetivos, "
            "metas y plazos; criterios incluyen volumen y posibilidad de valorización.",
            "Art. 84–86: special regime for prioritized goods — producers implement individual "
            "or collective management systems (EPR); Supreme Decree regulates goods, objectives, "
            "targets and timelines; criteria include volume and valorization potential.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben implementar sistemas'. DS subordinados para envases/plásticos priorizados. "
            "Persona jurídica para sistemas colectivos.",
            "'Must implement systems'. Subordinate supreme decrees for prioritized "
            "packaging/plastics. Legal entity for collective systems.",
        ),
        "instrument_score_notes": bi(
            "Marco REP/EPR para envases y bienes con contenido plástico priorizados.",
            "EPR framework for packaging and prioritized plastic-containing goods.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 87–89: Registro Autoritativo de EO-RS administrado por MINAM — inscripción "
            "previa obligatoria para operaciones de barrido, recolección, transporte, "
            "transferencia, tratamiento, valorización y disposición final.",
            "Art. 87–89: Authoritative Registry of EO-RS administered by MINAM — prior "
            "mandatory registration for sweeping, collection, transport, transfer, treatment, "
            "valorization and final disposal operations.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Deben inscribirse previamente'. Declaración jurada y fiscalización posterior. "
            "Revocación Art. 97. Municipalidades contratan solo EO-RS registradas.",
            "'Must register beforehand'. Sworn declaration and subsequent inspection. Revocation "
            "Art. 97. Municipalities contract only registered EO-RS.",
        ),
        "instrument_score_notes": bi(
            "Formalización de operadores de reciclaje comercial de plásticos.",
            "Formalization of commercial plastic-recycling operators.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 75: comercialización de residuos aprovechables por generadores, recicladores "
            "formalizados y EO-RS; generadores en programa municipal de segregación no pueden "
            "comercializar.",
            "Art. 75: commercialization of recoverable waste by generators, formalized recyclers "
            "and EO-RS; generators in municipal segregation program may not commercialize.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Restricción a generadores en programa selectivo. Medidas de seguridad Art. 76. "
            "Mercado de plásticos y metales recuperados.",
            "Restriction on generators in selective program. Safety measures Art. 76. Market for "
            "recovered plastics and metals.",
        ),
        "instrument_score_notes": bi(
            "Comercialización formal de plásticos recuperados.",
            "Formal commercialization of recovered plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": bi(
            "Art. 77–80: importación solo para valorización; autorización MINAM por embarque; "
            "sujeción Convenio de Basilea; plazo emisión autorización 20 días hábiles.",
            "Art. 77–80: import only for valorization; MINAM authorization per shipment; Basel "
            "Convention applies; authorization issued within 20 business days.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Solo está permitida la importación... para su valorización'. MINAM DGGRS autoriza. "
            "Revocación Art. 83.",
            "'Import of solid waste is permitted only for valorization'. MINAM DGGRS authorizes. "
            "Revocation Art. 83.",
        ),
        "instrument_score_notes": bi(
            "Control transfronterizo de residuos plásticos y peligrosos.",
            "Transboundary control of plastic and hazardous waste.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 101: actividades en centros de acopio — recepción, clasificación, pesaje, "
            "compactación y embalaje de residuos inorgánicos no peligrosos recuperados.",
            "Art. 101: collection-center activities — reception, sorting, weighing, compaction "
            "and packaging of recovered non-hazardous inorganic waste.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Condiciones técnicas de operación. Autorización municipal. Vinculado Art. 35.",
            "Technical operating conditions. Municipal authorization. Linked to Art. 35.",
        ),
        "instrument_score_notes": bi(
            "Operaciones de acondicionamiento de plásticos para comercialización.",
            "Conditioning operations for plastics commercialization.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Art. 15–18: proyectos de infraestructura de residuos requieren IGA aprobado; "
            "requisitos de admisibilidad para disposición final; clasificación anticipada "
            "Anexo II; modificación IGA para valorización/coprocesamiento.",
            "Art. 15–18: waste infrastructure projects require approved EIA instrument; "
            "admissibility requirements for final disposal; anticipated classification Annex II; "
            "EIA modification for valorization/coprocessing.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Debe contar con' IGA. SENACE/Gobierno Regional/Municipalidad evalúan. Trial burn "
            "para coprocesamiento Art. 17.",
            "'Must have' approved EIA instrument. SENACE/Regional Government/Province evaluate. "
            "Trial burn for coprocessing Art. 17.",
        ),
        "instrument_score_notes": bi(
            "Licenciamiento ambiental de plantas de valorización y rellenos sanitarios.",
            "Environmental licensing of valorization plants and sanitary landfills.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Art. 118–121: OEFA administra Inventario Nacional de Áreas Degradadas; "
            "recuperación y reconversión de botaderos; requisitos aprobación expedientes "
            "técnicos de recuperación.",
            "Art. 118–121: OEFA administers National Inventory of Degraded Areas; dump "
            "reconversion and recovery; technical-file approval requirements for recovery.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "OEFA elabora inventario. Municipios y EO-RS implementan planes recuperación. "
            "Categorización MINAM.",
            "OEFA prepares inventory. Municipalities and EO-RS implement recovery plans. MINAM "
            "categorization.",
        ),
        "instrument_score_notes": bi(
            "Recuperación de sitios con acumulación persistente de plásticos.",
            "Recovery of sites with persistent plastic accumulation.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "Art. 130–136: OEFA y EFA ejercen supervisión, fiscalización y sanción; tabla de "
            "infracciones Art. 135 vinculada a DL 1278 Art. 55; multas coercitivas y "
            "graduación Art. 136 según Ley 28611.",
            "Art. 130–136: OEFA and EFAs exercise supervision, enforcement and sanctions; "
            "Art. 135 infraction table linked to DL 1278 Art. 55; coercive fines and Art. 136 "
            "graduation per Law 28611.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "Autoridades designadas. Tipificación expresa de infracciones por segregación, "
            "recolección, valorización, registro EO-RS. Multas y medidas administrativas.",
            "Authorities designated. Express typification of infractions for segregation, "
            "collection, valorization, EO-RS registration. Fines and administrative measures.",
        ),
        "instrument_score_notes": bi(
            "Sanciones por manejo inadecuado de residuos plásticos y vertidos ilegales.",
            "Sanctions for improper plastic-waste management and illegal dumping.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "DS 001-2022-MINAM (enmienda): modifica Arts. 34, 38 y 39 del Reglamento; "
            "incorpora Registro Nacional de Recicladores administrado por MINAM vía SIGERSOL.",
            "DS 001-2022-MINAM (amendment): amends Regulation Arts. 34, 38 and 39; incorporates "
            "National Registry of Recyclers administered by MINAM via SIGERSOL.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "En vigor desde 9 ene. 2022. MINAM DGGRS administra registro. Municipalidades "
            "reportan formalización en 10 días hábiles.",
            "In force since 9 Jan 2022. MINAM DGGRS administers registry. Municipalities report "
            "formalization within 10 business days.",
        ),
        "instrument_score_notes": bi(
            "Registro nacional de recicladores que recuperan plásticos y otros materiales.",
            "National registry of recyclers recovering plastics and other materials.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": bi(
            "DS 002-2024-MINAM (4 abr. 2024): última enmienda sustantiva al Reglamento del "
            "DL 1278 — actualiza disposiciones operativas de gestión integral de residuos "
            "sólidos en coordinación con reformas posteriores a la Ley.",
            "DS 002-2024-MINAM (4 Apr 2024): latest substantive amendment to the DL 1278 "
            "Regulation — updates operational provisions on integrated solid-waste management "
            "coordinated with subsequent Law reforms.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Decreto supremo en vigor. MINAM autoridad de aplicación. Detalle en norma "
            "modificatoria 2024.",
            "Supreme decree in force. MINAM implementing authority. Detail in 2024 amending "
            "norm.",
        ),
        "instrument_score_notes": bi(
            "Actualización reglamentaria 2024 del marco de gestión de plásticos y residuos.",
            "2024 regulatory update of plastics and waste management framework.",
        ),
    },
]

# Columns A–X per coding prompt (no separate English name or country columns)
COLUMNS = [
    ("A", "policy_name"),
    ("B", "policy_url"),
    ("C", "policy_year"),
    ("D", "policy_objective"),
    ("E", "policy_target"),
    ("F", "policy_target_text"),
    ("G", "policy_type"),
    ("H", "policy_type_justification"),
    ("I", "policy_integration"),
    ("J", "policy_sectors_list"),
    ("K", "policy_circularity"),
    ("L", "policy_lifecycle_phases_list"),
    ("M", "policy_budget"),
    ("N", "policy_budget_text"),
    ("O", "policy_score"),
    ("P", "instrument_type"),
    ("Q", "instrument_lifecycle_stage"),
    ("R", "instrument_description"),
    ("S", "instrument_in_force"),
    ("T", "instrument_implementation"),
    ("U", "instrument_implementation_text"),
    ("V", "instrument_score"),
    ("W", "comments"),
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
                POLICY["policy_name"],
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
                inst["instrument_score_notes"],
            ]
        )

    widths = {
        "A": 40, "B": 38, "C": 8, "D": 44, "E": 8, "F": 44, "G": 8, "H": 40,
        "I": 10, "J": 40, "K": 10, "L": 36, "M": 10, "N": 42, "O": 10, "P": 10,
        "Q": 18, "R": 50, "S": 10, "T": 12, "U": 44, "V": 10, "W": 40,
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
        "Approved: 21 December 2017 | Amended: 9 January 2022 (DS 001-2022-MINAM); "
        "4 April 2024 (DS 002-2024-MINAM)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 014-2017-MINAM, "
        "Regulation of Legislative Decree No. 1278, as amended."
    )

    sections = [
        (
            "Article 1 — Purpose",
            "El presente dispositivo normativo tiene como objeto reglamentar el Decreto Legislativo "
            "N° 1278... minimización... valorización material y energética... disposición final...",
            "This regulation implements Legislative Decree No. 1278 to maximize material efficiency "
            "and govern source minimization, material and energy valorization, adequate final disposal "
            "and sustainability of the public cleaning service.",
        ),
        (
            "Article 9 — PLANRES",
            "El PLANRES es un instrumento nacional... se aprueba por Decreto Supremo... se actualiza "
            "cada cinco (05) años...",
            "PLANRES is a national instrument aligned with PLANAA targets; approved by Supreme "
            "Decree on MINAM proposal; updated every five years based on objective and target "
            "compliance analysis.",
        ),
        (
            "Article 10 — Municipal waste management plans",
            "Los Planes... deben estar alineados al PLANRES... presentan ante el MINAM y al OEFA el "
            "reporte... último día hábil del mes de marzo...",
            "Provincial and district municipal waste plans must align with PLANRES, be updated every "
            "five years, and submit annual activity reports to MINAM and OEFA by the last business "
            "day of March.",
        ),
        (
            "Article 11 — Segregation and selective collection program",
            "Programa de Segregación en la Fuente y Recolección Selectiva... estrategias para la "
            "segregación en fuente y el diseño de la recolección selectiva... participación de las "
            "organizaciones de recicladores formalizados.",
            "Source Segregation and Selective Collection Program: municipal technical instrument "
            "formulating source-segregation strategies and selective-collection design, including "
            "participation of formalized recycler organizations.",
        ),
        (
            "Article 19 — Source segregation",
            "El generador de residuos municipales debe realizar la segregación de sus residuos "
            "sólidos... para facilitar su valorización y/o disposición final.",
            "Municipal waste generators must segregate solid waste by physical, chemical and "
            "biological characteristics to facilitate valorization and/or final disposal.",
        ),
        (
            "Articles 36–38 — Valorization priority and national targets",
            "La valorización de residuos sólidos municipales debe priorizarse... Las metas nacionales "
            "de valorización... son establecidas por el MINAM, en el PLANAA y PLANRES.",
            "Valorization of municipal solid waste must be prioritized over final disposal. National "
            "valorization targets are set by MINAM in PLANAA and PLANRES and verified using SIGERSOL "
            "data.",
        ),
        (
            "Articles 84–86 — Prioritized goods / EPR regime",
            "El régimen especial... productores... deben implementar sistemas específicos de manejo... "
            "responsabilidad extendida del productor...",
            "Special regime for prioritized mass-consumption goods: producers must implement "
            "individual or collective post-consumer management systems under extended producer "
            "responsibility, regulated by supreme decree with goods, targets and timelines.",
        ),
        (
            "Articles 87–89 — EO-RS Authoritative Registry",
            "Las empresas... deben inscribirse previamente en el Registro Autoritativo de Empresas "
            "Operadoras de Residuos Sólidos administrado por el MINAM.",
            "Waste operator companies must register beforehand in the Authoritative Registry of Waste "
            "Operator Companies administered by MINAM before conducting collection, transport, "
            "valorization or disposal operations.",
        ),
        (
            "Articles 130–136 — Supervision and sanctions",
            "Las funciones de supervisión, fiscalización y sanción... a cargo del OEFA... Infracciones "
            "administrativas...",
            "OEFA and environmental enforcement agencies exercise supervision, enforcement and "
            "sanctions; Art. 135 typifies administrative infractions for non-compliance with "
            "segregation, collection, valorization and registry obligations.",
        ),
        (
            "Amendments — DS 001-2022 and DS 002-2024",
            "DS 001-2022-MINAM modifica Arts. 34, 38 y 39 e incorpora Registro Nacional de "
            "Recicladores. DS 002-2024-MINAM (4 abr. 2024) actualiza disposiciones del Reglamento.",
            "DS 001-2022-MINAM amends Arts. 34, 38 and 39 and incorporates the National Registry of "
            "Recyclers. DS 002-2024-MINAM (4 Apr 2024) updates Regulation provisions.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DS 014-2017-MINAM is the implementing regulation of Legislative Decree No. 1278. "
        "It derogated prior waste regulations upon entry into force and coordinates with Law "
        "No. 29419 on waste pickers, PLANRES, and subordinate EPR supreme decrees for prioritized "
        "goods including packaging and plastics."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_014_2017_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_014_2017_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(f"  Excel: {REPO_RAW}/output/peru-ds-014-2017-minam/Peru_DS_014_2017_MINAM_4P_Index_Coding.xlsx")
    print(
        f"  Word:  {REPO_RAW}/output/peru-ds-014-2017-minam/Peru_DS_014_2017_MINAM_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
