#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 023-2021-MINAM (PNA 2030)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-023-2021-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ds-023-2021-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = "Decreto Supremo N.° 023-2021-MINAM – Política Nacional del Ambiente al 2030"
POLICY_NAME_EN = "Supreme Decree No. 023-2021-MINAM – National Environmental Policy to 2030"
POLICY_URL = (
    "https://www.gob.pe/institucion/minam/campanas/2041-politica-nacional-del-ambiente"
)


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2021,
    "policy_objective": bi(
        "Política nacional marco que define lineamientos, objetivos prioritarios, metas al 2030 "
        "y servicios para orientar la acción ambiental del Estado, gobiernos subnacionales, "
        "sector privado y sociedad civil. Incluye el Objetivo Prioritario 4 (disposición "
        "adecuada de residuos sólidos, incl. plásticos), el Objetivo Prioritario 7 (economía "
        "circular en cadenas productivas y de consumo) y el Objetivo Prioritario 9 "
        "(comportamiento ambiental ciudadano), con diagnóstico explícito del consumo de "
        "plásticos y residuos plásticos de un solo uso.",
        "Overarching national policy defining guidelines, priority objectives, 2030 targets "
        "and services to steer environmental action by the State, subnational governments, "
        "private sector and civil society. Includes Priority Objective 4 (adequate solid-waste "
        "disposal, incl. plastics), Priority Objective 7 (circular economy in production and "
        "consumption chains) and Priority Objective 9 (citizen environmental behavior), with "
        "explicit diagnosis of plastic consumption and single-use plastic waste.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "OP4.I2: 63,24% de residuos sólidos municipales generados dispuestos en infraestructura "
        "de disposición final adecuada al 2030. OP4.I1: 2,91% de residuos municipales "
        "valorizados. OP4.I3: 2% tasa residuos no municipales en disposición final adecuada. "
        "OP7.I1: 64% energías renovables en matriz eléctrica nacional. OP9.I1: índice de "
        "comportamiento ambiental ciudadano 0,3252. OP7: tránsito hacia economía circular.",
        "OP4.I2: 63.24% of generated municipal solid waste disposed in adequate final-disposal "
        "infrastructure by 2030. OP4.I1: 2.91% of municipal waste valorized. OP4.I3: 2% rate "
        "for non-municipal waste in adequate final disposal. OP7.I1: 64% renewable energy in "
        "national electricity matrix. OP9.I1: citizen environmental behavior index 0.3252. "
        "OP7: transition toward circular economy.",
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Decreto Supremo N.° 023-2021-MINAM aprobado por el Poder Ejecutivo (22 jul. 2021, "
        "publicado 25 jul. 2021 en El Peruano), refrendado por PCM y 9 ministerios incl. "
        "MINAM. No enmendado. Política nacional de competencia exclusiva del Poder Ejecutivo "
        "bajo Ley 28611 Art. 8.1 y DS 029-2018-PCM; deroga DS 012-2009-MINAM. Clasificado "
        "0.75 (decreto supremo/política nacional ejecutiva).",
        "Supreme Decree No. 023-2021-MINAM approved by the Executive (22 Jul 2021, published "
        "25 Jul 2021 in El Peruano), countersigned by PCM and 9 ministries incl. MINAM. Not "
        "amended. National policy of exclusive Executive competence under Law 28611 Art. 8.1 "
        "and DS 029-2018-PCM; repeals DS 012-2009-MINAM. Classified 0.75 (supreme decree/"
        "executive national policy).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "todos los sectores (transversal), ambiente, gobiernos nacionales/regionales/locales, "
        "producción, consumo, residuos sólidos, educación, energía, saneamiento, agricultura, "
        "comercio, vivienda, defensa",
        "all sectors (cross-cutting), environment, national/regional/local governments, "
        "production, consumption, solid waste, education, energy, sanitation, agriculture, "
        "commerce, housing, defense",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "producción, consumo, recolección, reciclaje, valorización, disposición (marco "
        "estratégico ciclo de vida)",
        "production, consumption, collection, recycling, valorization, disposal (strategic "
        "whole-life-cycle framework)",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": bi(
        "DS Art. 6: implementación con cargo al Presupuesto Institucional de pliegos "
        "correspondientes conforme leyes anuales de presupuesto, sin demandar recursos "
        "adicionales al Tesoro Público. Servicios OP financiados por entidades responsables.",
        "DS Art. 6: implementation from Institutional Budget of corresponding entities per "
        "annual public-sector budget laws, without requiring additional Public Treasury "
        "resources. OP services financed by responsible entities.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DS Art. 1: 'Apruébase la Política Nacional del Ambiente al 2030', la misma que "
            "como Anexo forma parte integrante del decreto supremo.",
            "DS Art. 1: 'The National Environmental Policy to 2030 is approved', forming an "
            "integral annex to the supreme decree.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Apruébase' — vigencia desde publicación 25 jul. 2021. Deroga DS 012-2009-MINAM.",
            "'Is approved' — effective from publication 25 Jul 2021. Repeals DS 012-2009-MINAM.",
        ),
        "comments": bi(
            "Acto habilitante de la política ambiental marco al 2030.",
            "Enabling act for the overarching 2030 environmental policy.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DS Art. 2: la Política Nacional del Ambiente al 2030 es de 'cumplimiento "
            "obligatorio' para entidades de la Administración Pública y aplicable a personas "
            "jurídicas de derecho privado y sociedad civil en cuanto corresponda.",
            "DS Art. 2: the National Environmental Policy to 2030 is of 'mandatory compliance' "
            "for Public Administration entities and applicable to private legal persons and "
            "civil society where relevant.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Cumplimiento obligatorio' para administración pública. Roles Arts. 19–20 DS "
            "029-2018-PCM para entidades.",
            "'Mandatory compliance' for public administration. Arts. 19–20 DS 029-2018-PCM "
            "roles for entities.",
        ),
        "comments": bi(
            "Ámbito transversal de aplicación de la política.",
            "Cross-cutting policy application scope.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DS Art. 3–4: conducción a cargo de MINAM (DGPIGA); entidades responsables de "
            "objetivos prioritarios implementan y ejecutan vía planes del SINAPLAN coordinando "
            "con MINAM.",
            "DS Arts. 3–4: steering by MINAM (DGPIGA); entities responsible for priority "
            "objectives implement and execute via SINAPLAN plans in coordination with MINAM.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "MINAM designado rector. Implementación sectorial vía SINAPLAN. Art. 4 asigna "
            "responsables por OP.",
            "MINAM designated lead. Sectoral implementation via SINAPLAN. Art. 4 assigns "
            "responsible entities per PO.",
        ),
        "comments": bi(
            "Gobernanza multi-nivel de la política.",
            "Multi-level policy governance.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DS Art. 5: MINAM (DGPIGA) tiene a su cargo seguimiento y evaluación conforme "
            "pautas CEPLAN; entidades responsables brindan información oportunamente.",
            "DS Art. 5: MINAM (DGPIGA) monitors and evaluates per CEPLAN methodology; "
            "responsible entities provide information timely.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Tiene a su cargo el seguimiento y la evaluación'. Reporte de entidades "
            "obligatorio Art. 5.3.",
            "'Has charge of monitoring and evaluation'. Entity reporting mandatory Art. 5.3.",
        ),
        "comments": bi(
            "Monitoreo de cumplimiento de metas al 2030.",
            "Monitoring of 2030 target compliance.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "Objetivo Prioritario 4: 'Incrementar la disposición adecuada de los residuos "
            "sólidos' — aborda deterioro ambiental por mala disposición (ríos, mar, lagos) y "
            "lineamientos: eficiencia instrumentos LGIRS, fiscalización, mejoras gestión "
            "integral, valorización de residuos.",
            "Priority Objective 4: 'Increase adequate disposal of solid waste' — addresses "
            "environmental deterioration from poor disposal (rivers, sea, lakes) and "
            "guidelines: LGIRS instrument efficiency, enforcement, integrated-management "
            "improvements, waste valorization.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Objetivo estratégico con 4 lineamientos y servicios OP4.S1–S2. Responsable MINAM. "
            "Vinculado DL 1278 y PLANRES.",
            "Strategic objective with 4 guidelines and OP4.S1–S2 services. MINAM lead. Linked "
            "to DL 1278 and PLANRES.",
        ),
        "comments": bi(
            "Objetivo marco para gestión de residuos sólidos incl. plásticos. Nota: meta 63% "
            "es OP4 (no OP9).",
            "Framework objective for solid-waste management incl. plastics. Note: 63% target is "
            "OP4 (not OP9).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "OP4.I2 — Indicador: 'Porcentaje de residuos sólidos municipales generados, que "
            "se disponen en una infraestructura de disposición final adecuada' — logro "
            "esperado al 2030: 63,24%.",
            "OP4.I2 — Indicator: 'Percentage of generated municipal solid waste disposed in "
            "adequate final-disposal infrastructure' — expected achievement by 2030: 63.24%.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "Meta cuantitativa nacional explícita al 2030. Indicador principal OP4. Cubre "
            "flujos municipales incl. fracción plástica.",
            "Explicit national quantitative 2030 target. Main OP4 indicator. Covers municipal "
            "streams incl. plastic fraction.",
        ),
        "comments": bi(
            "Meta nacional clave citada en metadata del usuario (63% ≈ 63,24% OP4.I2).",
            "Key national target cited in user metadata (63% ≈ 63.24% OP4.I2).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "OP4.I1 — Indicador: 'Porcentaje de residuos sólidos municipales valorizados' — "
            "logro esperado al 2030: 2,91%.",
            "OP4.I1 — Indicator: 'Percentage of municipal solid waste valorized' — expected "
            "achievement by 2030: 2.91%.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Meta cuantitativa de valorización. Lineamiento 4 OP4: 'Impulsar la valorización de "
            "los residuos sólidos'.",
            "'Valorization' quantitative target. OP4 guideline 4: 'Promote solid-waste "
            "valorization'.",
        ),
        "comments": bi(
            "Meta de valorización/reciclaje de residuos municipales.",
            "Municipal waste valorization/recycling target.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "OP4 Lineamiento 1: 'Mejorar la eficiencia de los instrumentos técnico-normativos "
            "de gestión integral de los residuos sólidos' (vinculado actualización PLANRES y "
            "DL 1278).",
            "OP4 Guideline 1: 'Improve efficiency of technical-normative instruments for "
            "integrated solid-waste management' (linked to PLANRES update and DL 1278).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Lineamiento técnico-normativo sin servicios directos ('Sin servicios'). Depende "
            "de actualización normativa PLANRES/DL 1278.",
            "Technical-normative guideline without direct services ('No services'). Depends on "
            "PLANRES/DL 1278 normative update.",
        ),
        "comments": bi(
            "Marco normativo para gestión integral incl. plásticos (Ley 30884, DL 1278).",
            "Regulatory framework for integrated management incl. plastics (Law 30884, DL 1278).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Disposal",
        "instrument_description": bi(
            "OP4 Lineamiento 2: 'Fortalecer la fiscalización de la gestión y manejo de los "
            "residuos sólidos municipales y no municipales'.",
            "OP4 Guideline 2: 'Strengthen enforcement of municipal and non-municipal solid-"
            "waste management'.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Lineamiento de fiscalización ('Sin servicios'). DL 1278 Título VIII régimen "
            "supervisión/sanción.",
            "Enforcement guideline ('No services'). DL 1278 Title VIII supervision/sanction "
            "regime.",
        ),
        "comments": bi(
            "Fortalecimiento de fiscalización de flujos de residuos plásticos y sólidos.",
            "Strengthened enforcement of plastic and solid-waste streams.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "OP4.S1: 'Evaluación de los Planes de Manejo de Residuos de Bienes Priorizados' "
            "en marco REP, dirigido a empresas privadas — MINAM/DGRS.",
            "OP4.S1: 'Evaluation of Prioritized-Goods Waste Management Plans' under EPR, "
            "targeting private companies — MINAM/DGRS.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Servicio de evaluación REP para bienes priorizados (bolsas PET, RAEE, etc.). "
            "Implementación MINAM.",
            "EPR evaluation service for prioritized goods (PET bottles, WEEE, etc.). MINAM "
            "implementation.",
        ),
        "comments": bi(
            "Vincula política marco con regímenes REP de plásticos y e-waste.",
            "Links framework policy to plastic and e-waste EPR regimes.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "OP4.S2: 'Fortalecimiento de capacidades en materia de gestión integral de residuos "
            "sólidos a los gobiernos locales' — MINAM/DGRS.",
            "OP4.S2: 'Capacity building on integrated solid-waste management for local "
            "governments' — MINAM/DGRS.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Fortalecimiento de capacidades... de manera precisa'. Servicio a gobiernos locales.",
            "'Capacity building... in a targeted manner'. Service to local governments.",
        ),
        "comments": bi(
            "Capacitación municipal en gestión de residuos incl. plásticos.",
            "Municipal training on waste management incl. plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "Objetivo Prioritario 7: 'Mejorar el desempeño ambiental de las cadenas "
            "productivas y de consumo de bienes y servicios, aplicando la economía circular' — "
            "reúso de residuos sólidos, recuperación de materiales, eficiencia energética.",
            "Priority Objective 7: 'Improve environmental performance of production and "
            "consumption chains for goods and services, applying circular economy' — reuse of "
            "solid waste, material recovery, energy efficiency.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Objetivo estratégico con 3 lineamientos y servicios OP7.S1–S3. Hojas de ruta "
            "sectoriales en marcha.",
            "Strategic objective with 3 guidelines and OP7.S1–S3 services. Sectoral roadmaps "
            "underway.",
        ),
        "comments": bi(
            "Transición nacional hacia economía circular que abarca plásticos.",
            "National transition toward circular economy encompassing plastics.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "OP7 Lineamiento 1: 'Generar las condiciones en las entidades públicas y privadas "
            "para el tránsito hacia una economía circular' — extracción, transformación, "
            "distribución, uso y recuperación de materiales.",
            "OP7 Guideline 1: 'Generate conditions in public and private entities for "
            "transition toward a circular economy' — extraction, transformation, distribution, "
            "use and material recovery.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Lineamiento estratégico de economía circular. Servicios OP7.S1–S3 de apoyo "
            "(PRODUCE/MINAM).",
            "Strategic circular-economy guideline. Supporting services OP7.S1–S3 "
            "(PRODUCE/MINAM).",
        ),
        "comments": bi(
            "Marco de transición circular aplicable a materiales poliméricos.",
            "Circular transition framework applicable to polymer materials.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "OP7.S2–S3: fortalecimiento de capacidades en economía circular, buenas prácticas "
            "y producción limpia (PRODUCE agentes manufactureros/comercio; MINAM sector "
            "público/privado).",
            "OP7.S2–S3: capacity building on circular economy, good practices and cleaner "
            "production (PRODUCE for manufacturing/trade agents; MINAM for public/private "
            "sector).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Servicios de capacitación 'de manera precisa/fiable'. Dirigidos a agentes "
            "económicos manufactureros.",
            "'Targeted/reliable' training services. Directed at manufacturing economic agents.",
        ),
        "comments": bi(
            "Capacitación sectorial en economía circular y producción limpia.",
            "Sectoral circular economy and cleaner-production training.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Objetivo Prioritario 9: 'Mejorar el comportamiento ambiental de la ciudadanía' — "
            "diagnóstico cita consumo de plásticos (+50% 2008–2016), 68% residuos plásticos "
            "de un solo uso (bolsas, PET, poliestireno) y solo 58% hogares segregan residuos.",
            "Priority Objective 9: 'Improve citizen environmental behavior' — diagnosis cites "
            "plastic consumption (+50% 2008–2016), 68% single-use plastic waste (bags, PET, "
            "polystyrene) and only 58% of households segregate waste.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Objetivo con indicador OP9.I1 (índice 0,3252). Lineamientos educación formal y "
            "comunitaria. Coordinación MINEDU.",
            "Objective with indicator OP9.I1 (index 0.3252). Formal and community education "
            "guidelines. MINEDU coordination.",
        ),
        "comments": bi(
            "OP9 es comportamiento ciudadano (no la meta 63% residuos — esa es OP4). Diagnóstico "
            "explícito sobre plásticos.",
            "OP9 is citizen behavior (not the 63% waste target — that is OP4). Explicit plastic "
            "diagnosis.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OP9.I1 — Indicador: 'Índice de comportamiento ambiental de la ciudadanía' — logro "
            "esperado al 2030: 0,3252; lineamientos: integración enfoque ambiental en educación "
            "formal/comunitaria y sostenibilidad de acción ciudadana.",
            "OP9.I1 — Indicator: 'Citizen environmental behavior index' — expected achievement "
            "by 2030: 0.3252; guidelines: integrate environmental approach in formal/community "
            "education and sustainability of citizen action.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Meta cuantitativa índice comportamental. Servicios OP9.S1–S2 (MINAM gobiernos "
            "subnacionales, gestión escolar).",
            "Quantitative behavioral-index target. Services OP9.S1–S2 (MINAM subnational "
            "governments, school management).",
        ),
        "comments": bi(
            "Indicador de consumo/comportamiento sostenible incl. hábitos plásticos.",
            "Sustainable consumption/behavior indicator incl. plastic habits.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OP9 Lineamiento 1: 'Garantizar la integración del enfoque ambiental en la "
            "educación formal y comunitaria' — MINAM/DGECIA.",
            "OP9 Guideline 1: 'Ensure integration of the environmental approach in formal and "
            "community education' — MINAM/DGECIA.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Educación ambiental vinculada Política Nacional de Educación Ambiental. Servicio "
            "OP9.S1 capacidades gobiernos subnacionales.",
            "Environmental education linked to National Environmental Education Policy. "
            "Service OP9.S1 subnational government capacity.",
        ),
        "comments": bi(
            "Educación para reducir consumo insostenible de plásticos.",
            "Education to reduce unsustainable plastic consumption.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DCF Primera: MINAM 'aprueba las normas complementarias que se requieran' para "
            "implementación de la Política Nacional del Ambiente al 2030.",
            "First FCP: MINAM 'approves complementary norms as required' for implementation of "
            "the National Environmental Policy to 2030.",
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "'Aprueba las normas complementarias' — poder habilitante. Normas futuras por RM/"
            "actos administrativos.",
            "'Approves complementary norms' — enabling power. Future norms via ministerial "
            "resolutions/administrative acts.",
        ),
        "comments": bi(
            "Habilitación para normas complementarias de implementación.",
            "Enabling future complementary implementation norms.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": bi(
            "DCF Segunda: entidades involucradas 'adecúan progresivamente sus instrumentos de "
            "planeamiento' a la política conforme DS 029-2018-PCM Art. 11.4.",
            "Second FCP: involved entities 'progressively adapt their planning instruments' to "
            "the policy per DS 029-2018-PCM Art. 11.4.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Adecúan progresivamente' — obligación de alineación de planes sectoriales con "
            "metas OP4/OP7/OP9.",
            "'Progressively adapt' — obligation to align sectoral plans with OP4/OP7/OP9 "
            "targets.",
        ),
        "comments": bi(
            "Alineación de planeamiento estratégico con metas de residuos y economía circular.",
            "Strategic planning alignment with waste and circular-economy targets.",
        ),
    },
]

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
                inst["comments"],
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
        "Signed: 22 July 2021 | Published: 25 July 2021 (El Peruano)\n"
        "Not amended\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 023-2021-MINAM "
        "approving Peru's National Environmental Policy to 2030."
    )

    sections = [
        (
            "Supreme Decree Art. 1 — Approval",
            "Apruébase la Política Nacional del Ambiente al 2030, la misma que como Anexo forma "
            "parte integrante del presente Decreto Supremo.",
            "The National Environmental Policy to 2030 is approved, forming an integral annex "
            "to this Supreme Decree.",
        ),
        (
            "Supreme Decree Art. 2 — Scope",
            "La Política Nacional del Ambiente al 2030 es de cumplimiento obligatorio para las "
            "entidades de la Administración Pública...",
            "The National Environmental Policy to 2030 is of mandatory compliance for Public "
            "Administration entities...",
        ),
        (
            "Priority Objective 4 — Solid waste",
            "Incrementar la disposición adecuada de los residuos sólidos... OP4.I2: 63,24% en "
            "infraestructura de disposición final adecuada al 2030.",
            "Increase adequate disposal of solid waste... OP4.I2: 63.24% in adequate final-"
            "disposal infrastructure by 2030.",
        ),
        (
            "Priority Objective 7 — Circular economy",
            "Mejorar el desempeño ambiental de las cadenas productivas y de consumo... aplicando "
            "la economía circular.",
            "Improve environmental performance of production and consumption chains... applying "
            "the circular economy.",
        ),
        (
            "Priority Objective 9 — Citizen behavior",
            "Mejorar el comportamiento ambiental de la ciudadanía... diagnóstico: 68% residuos "
            "plásticos de un solo uso.",
            "Improve citizen environmental behavior... diagnosis: 68% single-use plastic waste.",
        ),
        (
            "Plastic consumption diagnosis",
            "El consumo nacional de plásticos se incrementó ~50% entre 2008 y 2016; el 68% de "
            "residuos plásticos son de un solo uso (bolsas, PET, poliestireno).",
            "National plastic consumption rose ~50% between 2008 and 2016; 68% of plastic waste "
            "is single-use (bags, PET, polystyrene).",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DS 023-2021-MINAM updates Peru's National Environmental Policy (repealing DS "
        "012-2009-MINAM). The 63% solid-waste disposal target is under Priority Objective 4 "
        "(OP4.I2: 63.24%), not Objective 9. Objective 9 addresses citizen environmental "
        "behavior. Links to DL 1278, PLANRES, Law 30884 and sectoral EPR regimes."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_023_2021_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_023_2021_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(
        f"  Excel: {REPO_RAW}/output/peru-ds-023-2021-minam/"
        "Peru_DS_023_2021_MINAM_4P_Index_Coding.xlsx"
    )
    print(
        f"  Word:  {REPO_RAW}/output/peru-ds-023-2021-minam/"
        "Peru_DS_023_2021_MINAM_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
