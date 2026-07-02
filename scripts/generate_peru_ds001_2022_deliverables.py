#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 001-2022-MINAM."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-001-2022-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ds-001-2022-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    "Decreto Supremo N.° 001-2022-MINAM – Modifica el Reglamento del D.L. N.° 1278 y el "
    "Reglamento de la Ley N.° 29419"
)
POLICY_NAME_EN = (
    "Supreme Decree No. 001-2022-MINAM – Amendments to the Regulation of Legislative "
    "Decree No. 1278 and the Regulation of Law No. 29419 (Waste Pickers Law)"
)
POLICY_URL = (
    "https://www.gob.pe/institucion/minam/normas-legales/2649587-001-2022-minam"
)


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2022,
    "policy_objective": bi(
        "Decreto Supremo que modifica el Reglamento del DL 1278 (gestión integral de residuos "
        "sólidos) y el Reglamento de la Ley 29419 (recicladores), fortaleciendo segregación en "
        "la fuente, programas de recolección selectiva, trazabilidad SIGERSOL, transporte de "
        "material de descarte, formalización de organizaciones de recicladores, Registro "
        "Nacional de Recicladores, plataforma Bolsas de Residuos y códigos Basel de desechos "
        "plásticos (PE, PP, PET) para reciclaje.",
        "Supreme Decree amending the DL 1278 solid-waste regulation and Law 29419 recyclers "
        "regulation, strengthening source segregation, selective-collection programs, SIGERSOL "
        "traceability, discard-material transport, formalization of recycler organizations, "
        "National Recycler Registry, Residue Bags commercial platform and Basel plastic-waste "
        "codes (PE, PP, PET) for recycling.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "DCF Primera: municipalidades aprueban/actualizan Programas de Segregación en la Fuente "
        "en 180 días. DCF Segunda: actualización cuadro de infracciones en 120 días hábiles. "
        "DCF Décima: plataforma Bolsas de Residuos en 180 días. DCF Décima Primera/Segunda: "
        "Registro Nacional de Recicladores en 90 días. Art. 34 (29419): evaluación formalización "
        "en 15 días hábiles. Art. 13-B.2: reporte formalizados al MINAM en 10 días hábiles. "
        "Art. 34 (1278): servicio municipal hasta 145 kg/día por generador.",
        "DCF Primera: municipalities approve/update Source Segregation Programs within 180 days. "
        "DCF Segunda: update sanctions chart within 120 business days. DCF Décima: Residue Bags "
        "platform within 180 days. DCF Décima Primera/Segunda: National Recycler Registry "
        "within 90 days. Art. 34 (29419): formalization evaluation within 15 business days. "
        "Art. 13-B.2: report formalized recyclers to MINAM within 10 business days. "
        "Art. 34 (1278): municipal service up to 145 kg/day per generator.",
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Decreto Supremo N.° 001-2022-MINAM aprobado por el Poder Ejecutivo (7 ene. 2022, "
        "publicado 9 ene. 2022 en El Peruano), refrendado por MINAM y ocho ministerios "
        "sectoriales. Enmienda técnica y procedimental al Reglamento DS 014-2017-MINAM (DL 1278) "
        "y al Reglamento DS 005-2010-MINAM (Ley 29419). No enmendado posteriormente como "
        "norma autónoma. Clasificado 0.75 (decreto supremo ejecutivo de modificación "
        "reglamentaria).",
        "Supreme Decree No. 001-2022-MINAM approved by the Executive (7 Jan 2022, published "
        "9 Jan 2022 in El Peruano), countersigned by MINAM and eight sector ministries. "
        "Technical and procedural amendment to DS 014-2017-MINAM regulation (DL 1278) and "
        "DS 005-2010-MINAM regulation (Law 29419). Not subsequently amended as a standalone "
        "instrument. Classified 0.75 (executive supreme-decree regulatory amendment).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "ambiente, gobiernos locales, sector reciclaje, transporte, industria, consumo, salud, "
        "vivienda, producción, comercio exterior",
        "environment, local governments, recycling sector, transport, industry, consumption, "
        "health, housing, production, foreign trade",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "consumo, recolección, reciclaje, valorización, disposición",
        "consumption, collection, recycling, valorization, disposal",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": bi(
        "Art. 6: financiamiento con cargo al presupuesto institucional de los pliegos "
        "involucrados, sin recursos adicionales del Tesoro Público. Art. 34-A: convenios "
        "interinstitucionales para recaudación de arbitrios de limpieza pública.",
        "Art. 6: financing from institutional budgets of involved entities, without additional "
        "Treasury resources. Art. 34-A: inter-institutional agreements for collection of "
        "municipal cleaning-fee revenue.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 1: objeto — modificar el Reglamento del DL 1278 (DS 014-2017-MINAM) y el "
            "Reglamento de la Ley 29419 (DS 005-2010-MINAM).",
            "Art. 1: purpose — amend the DL 1278 Regulation (DS 014-2017-MINAM) and the "
            "Law 29419 Regulation (DS 005-2010-MINAM).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Acto habilitante de las modificaciones reglamentarias. Vigencia desde publicación "
            "9 ene. 2022.",
            "Enabling act for regulatory amendments. Effective from publication 9 Jan 2022.",
        ),
        "comments": bi(
            "Marco de enmienda dual: residuos sólidos generales y recicladores.",
            "Dual amendment framework: general solid waste and recyclers.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 6 DL 1278): transporte terrestre del material de descarte debe "
            "regirse por normativa MTC, gobiernos regionales y locales en sus competencias.",
            "Art. 2 (amended Art. 6 DL 1278): land transport of discard material must follow "
            "MTC regulations and regional/local government rules within their competencies.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Obligación de cumplimiento para transportistas de material de descarte "
            "aprovechable (incl. plásticos post-consumo). Fiscalización MTC/municipal.",
            "Compliance obligation for transporters of recoverable discard material "
            "(incl. post-consumer plastics). MTC/municipal enforcement.",
        ),
        "comments": bi(
            "Relevante para cadena logística de plásticos reciclables.",
            "Relevant for recyclable plastics logistics chain.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 11 DL 1278): Programa de Segregación en la Fuente y Recolección "
            "Selectiva — instrumento técnico obligatorio de municipalidades provinciales y "
            "distritales; incluye orgánicos e inorgánicos aprovechables, rutas, horarios, "
            "identificación de recicladores y educación ambiental.",
            "Art. 2 (amended Art. 11 DL 1278): Source Segregation and Selective Collection "
            "Program — mandatory technical instrument for provincial and district "
            "municipalities; covers recoverable organics and inorganics, routes, schedules, "
            "recycler identification and environmental education.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "DCF Primera: aprobación/actualización en 180 días desde vigencia. Exceptuadas "
            "municipalidades rurales <10 000 hab. Reporte anual vía SIGERSOL.",
            "DCF Primera: approval/update within 180 days of entry into force. Rural "
            "municipalities <10,000 pop. exempt. Annual reporting via SIGERSOL.",
        ),
        "comments": bi(
            "Instrumento central para segregación de plásticos post-consumo en fuente.",
            "Central instrument for post-consumer plastic source segregation.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 13 DL 1278): SIGERSOL — sistema oficial de información para "
            "trazabilidad de residuos desde generación hasta valorización/disposición; "
            "reportes trimestrales y anuales de municipalidades sobre segregación y "
            "recolección selectiva.",
            "Art. 2 (amended Art. 13 DL 1278): SIGERSOL — official information system for "
            "waste traceability from generation to valorization/disposal; quarterly and annual "
            "municipal reports on segregation and selective collection.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "MINAM administra SIGERSOL; OEFA supervisa. Publicación en portal datos abiertos. "
            "Incumplimiento comunicado a Contraloría.",
            "MINAM administers SIGERSOL; OEFA supervises. Open-data portal publication. "
            "Non-compliance reported to Comptroller.",
        ),
        "comments": bi(
            "Habilita monitoreo de flujos de plásticos aprovechables.",
            "Enables monitoring of recoverable plastic flows.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 19 DL 1278): generadores municipales tienen obligación de "
            "segregar residuos según características físicas, químicas y biológicas en la "
            "fuente de generación para facilitar acondicionamiento, valorización y disposición.",
            "Art. 2 (amended Art. 19 DL 1278): municipal waste generators must segregate waste "
            "by physical, chemical and biological characteristics at the source to facilitate "
            "conditioning, valorization and disposal.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Obligación directa del generador. Vinculada a Programa de Segregación municipal. "
            "Fiscalización municipal/OEFA.",
            "Direct generator obligation. Linked to municipal Segregation Program. "
            "Municipal/OEFA enforcement.",
        ),
        "comments": bi(
            "Mandato de segregación en fuente aplicable a plásticos domiciliarios.",
            "Source-segregation mandate applicable to household plastics.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 28 DL 1278): recolección selectiva de residuos aprovechables "
            "puede realizarse por organizaciones de recicladores formalizados bajo Ley 29419; "
            "municipalidades deben implementar Programas de Segregación en toda su jurisdicción.",
            "Art. 2 (amended Art. 28 DL 1278): selective collection of recoverable waste may "
            "be carried out by formalized recycler organizations under Law 29419; municipalities "
            "must implement Segregation Programs throughout their jurisdiction.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Puente normativo entre DL 1278 y Ley 29419 para recolección de inorgánicos "
            "aprovechables (PET, PE, PP, etc.).",
            "Regulatory bridge between DL 1278 and Law 29419 for recoverable inorganic "
            "collection (PET, PE, PP, etc.).",
        ),
        "comments": bi(
            "Formaliza rol de recicladores en cadena de plásticos post-consumo.",
            "Formalizes recycler role in post-consumer plastics chain.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 2 (mod. Art. 47.3 DL 1278): residuos no municipales aprovechables similares "
            "a municipales pueden entregarse a organizaciones de recicladores formalizados en "
            "marco del Programa de Segregación en la Fuente y Recolección Selectiva.",
            "Art. 2 (amended Art. 47.3 DL 1278): recoverable non-municipal waste similar to "
            "municipal waste may be delivered to formalized recycler organizations under the "
            "Source Segregation and Selective Collection Program.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Extiende acceso de recicladores formalizados a flujos no municipales "
            "aprovechables hasta 145 kg/día.",
            "Extends formalized recyclers' access to recoverable non-municipal flows up to "
            "145 kg/day.",
        ),
        "comments": bi(
            "Amplía mercado de materiales plásticos recuperables para recicladores.",
            "Expands recoverable plastic materials market for recyclers.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 3 (incorp. Art. 12-A DL 1278): MINAM administra plataforma Bolsas de "
            "Residuos para facilitar transacciones comerciales de residuos aprovechables y "
            "material de descarte como materia prima.",
            "Art. 3 (incorporated Art. 12-A DL 1278): MINAM administers Residue Bags platform "
            "to facilitate commercial transactions of recoverable waste and discard material "
            "as raw material.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "DCF Décima: implementación plataforma en 180 días. Componente SINIA. Sin "
            "sanciones específicas en el artículo.",
            "DCF Décima: platform implementation within 180 days. SINIA component. No "
            "specific sanctions in the article.",
        ),
        "comments": bi(
            "Mercado secundario de residuos plásticos aprovechables.",
            "Secondary market for recoverable plastic waste.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 3 (incorp. Art. 13-B DL 1278): Registro Nacional de Recicladores administrado "
            "por DGRS/MINAM; municipalidades reportan formalizados vía SIGERSOL en 10 días "
            "hábiles; habilita inscripción en Registro Municipal para programas de segregación.",
            "Art. 3 (incorporated Art. 13-B DL 1278): National Recycler Registry administered "
            "by DGRS/MINAM; municipalities report formalized recyclers via SIGERSOL within "
            "10 business days; enables Municipal Registry inscription for segregation programs.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "DCF Décima Segunda: elaboración en 90 días con Censo Nacional de Recicladores. "
            "DCF Décima Primera: transición automática Inscripción → Formalización.",
            "DCF Décima Segunda: establishment within 90 days using National Recycler Census. "
            "DCF Décima Primera: automatic Inscripción → Formalización transition.",
        ),
        "comments": bi(
            "Formalización nacional de recicladores de plásticos y otros aprovechables.",
            "National formalization of plastics and other recoverables recyclers.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": bi(
            "Art. 3 (incorp. Anexo V Lista B código B3011): desechos plásticos PE, PP y PET "
            "destinados al reciclaje por separado, ambientalmente racional, apenas contaminados; "
            "efectivo 1 ene. 2021 (nota al pie).",
            "Art. 3 (incorporated Annex V List B code B3011): plastic waste PE, PP and PET "
            "destined for separate, environmentally sound recycling, barely contaminated; "
            "effective 1 Jan 2021 (footnote).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Alineación con Convención de Basilea para movimientos transfronterizos de "
            "plásticos reciclables. Fiscalización OEFA.",
            "Alignment with Basel Convention for transboundary movements of recyclable "
            "plastics. OEFA enforcement.",
        ),
        "comments": bi(
            "Referencia explícita a polietileno, polipropileno y PET.",
            "Explicit reference to polyethylene, polypropylene and PET.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 4 (mod. Art. 34 Ley 29419): requisitos para formalización de organizaciones "
            "de recicladores — capacitación, vacunación, equipamiento, vehículos de recolección; "
            "evaluación en 15 días hábiles con silencio negativo; constancia de formalización.",
            "Art. 4 (amended Art. 34 Law 29419): requirements for formalization of recycler "
            "organizations — training, vaccination, equipment, collection vehicles; evaluation "
            "within 15 business days with negative silence; formalization certificate.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Procedimiento administrativo municipal. Actualizaciones comunicadas al MINAM para "
            "Registro Nacional.",
            "Municipal administrative procedure. Updates communicated to MINAM for National "
            "Registry.",
        ),
        "comments": bi(
            "Vía principal de formalización de recicladores de plásticos post-consumo.",
            "Main formalization pathway for post-consumer plastics recyclers.",
        ),
    },
    {
        "instrument_type": 0.75,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 4 (mod. Art. 38 Ley 29419): Registro Municipal de Organizaciones de "
            "Recicladores autoriza recolección selectiva, transporte y acondicionamiento de "
            "residuos aprovechables; reporte al MINAM vía SIGERSOL.",
            "Art. 4 (amended Art. 38 Law 29419): Municipal Registry of Recycler Organizations "
            "authorizes selective collection, transport and conditioning of recoverable waste; "
            "reporting to MINAM via SIGERSOL.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Inscripción sujeta a necesidad municipal del programa de segregación. "
            "Evaluación 15 días hábiles.",
            "Inscription subject to municipal need for segregation program. "
            "15-business-day evaluation.",
        ),
        "comments": bi(
            "Autorización operativa para recolección de plásticos y otros aprovechables.",
            "Operational authorization for plastics and other recoverables collection.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Art. 4 (mod. Art. 39 Ley 29419): requisitos inscripción Registro Municipal — "
            "constancia de formalización, relación de miembros participantes, declaraciones "
            "juradas de equipamiento/vehículos; emisión de carné de identificación de "
            "recicladores.",
            "Art. 4 (amended Art. 39 Law 29419): Municipal Registry inscription requirements — "
            "formalization certificate, list of participating members, sworn declarations on "
            "equipment/vehicles; issuance of recycler identification cards.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Carné obligatorio durante jornada laboral. Silencio negativo en 15 días hábiles.",
            "ID card mandatory during work shift. Negative silence within 15 business days.",
        ),
        "comments": bi(
            "Identificación formal de recicladores en cadena de valor del plástico.",
            "Formal identification of recyclers in plastics value chain.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "DCF Primera: municipalidades deben aprobar y/o actualizar Programas de Segregación "
            "en la Fuente y Recolección Selectiva en 180 días calendario desde vigencia del DS.",
            "DCF Primera: municipalities must approve and/or update Source Segregation and "
            "Selective Collection Programs within 180 calendar days of the decree's entry into "
            "force.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Plazo vencido (desde ene. 2022). Programas deben cumplir Art. 11 modificado. "
            "Fiscalización OEFA/municipal.",
            "Deadline elapsed (since Jan 2022). Programs must comply with amended Art. 11. "
            "OEFA/municipal enforcement.",
        ),
        "comments": bi(
            "Plazo transitorio vinculante para implementación de segregación de plásticos.",
            "Binding transitional deadline for plastics segregation implementation.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "DCF Décima Primera: Constancias de Inscripción vigentes se convierten automáticamente "
            "en Constancias de Formalización; municipalidades remiten al MINAM en 90 días para "
            "Registro Nacional.",
            "DCF Décima Primera: valid Inscripción Certificates automatically become "
            "Formalización Certificates; municipalities submit to MINAM within 90 days for "
            "National Registry.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Transición automática sin requisitos adicionales. Plazo vencido desde 2022.",
            "Automatic transition without additional requirements. Deadline elapsed since 2022.",
        ),
        "comments": bi(
            "Regularización retroactiva de recicladores previamente inscritos.",
            "Retroactive regularization of previously registered recyclers.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Disposición Complementaria Derogatoria Única: derógase el Art. 31 del Reglamento "
            "de la Ley 29419 (DS 005-2010-MINAM).",
            "Derogatory Complementary Provision Única: Art. 31 of the Law 29419 Regulation "
            "(DS 005-2010-MINAM) is repealed.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Derogación inmediata desde vigencia 9 ene. 2022. Sustituido por nuevo marco "
            "Arts. 34, 38, 39 modificados.",
            "Immediate repeal from effective date 9 Jan 2022. Replaced by amended Arts. 34, "
            "38, 39 framework.",
        ),
        "comments": bi(
            "Elimina disposición anterior del régimen de recicladores.",
            "Removes prior provision from recycler regime.",
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
        "Signed: 7 January 2022 | Published: 9 January 2022 (El Peruano)\n"
        "Amends DS 014-2017-MINAM (DL 1278) and DS 005-2010-MINAM (Law 29419)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 001-2022-MINAM "
        "amending the solid-waste and waste-pickers regulations."
    )

    sections = [
        (
            "Art. 1 — Purpose",
            "Modificar el Reglamento del DL 1278 y el Reglamento de la Ley 29419.",
            "Amend the DL 1278 Regulation and the Law 29419 Regulation.",
        ),
        (
            "Amended Art. 11 — Segregation Program",
            "Programa de Segregación en la Fuente y Recolección Selectiva obligatorio para "
            "municipalidades; incluye inorgánicos aprovechables (plásticos), rutas, recicladores "
            "y educación ambiental.",
            "Mandatory Source Segregation and Selective Collection Program for municipalities; "
            "includes recoverable inorganics (plastics), routes, recyclers and environmental "
            "education.",
        ),
        (
            "Amended Art. 19 — Source segregation",
            "Generadores municipales tienen obligación de segregar residuos en la fuente.",
            "Municipal generators must segregate waste at source.",
        ),
        (
            "Amended Art. 28 — Selective collection",
            "Recolección selectiva de aprovechables por organizaciones de recicladores "
            "formalizados bajo Ley 29419.",
            "Selective collection of recoverables by formalized recycler organizations under "
            "Law 29419.",
        ),
        (
            "Art. 13-B — National Recycler Registry",
            "Registro Nacional de Recicladores administrado por MINAM; reporte vía SIGERSOL "
            "en 10 días hábiles desde formalización.",
            "National Recycler Registry administered by MINAM; reporting via SIGERSOL within "
            "10 business days of formalization.",
        ),
        (
            "Art. 34 (Law 29419) — Formalization",
            "Requisitos de formalización de recicladores con evaluación en 15 días hábiles.",
            "Recycler formalization requirements with 15-business-day evaluation.",
        ),
        (
            "Annex B3011 — Plastic waste",
            "Desechos plásticos PE, PP y PET destinados al reciclaje por separado.",
            "Plastic waste PE, PP and PET destined for separate recycling.",
        ),
        (
            "DCF Primera — Deadline",
            "Municipalidades aprueban Programas de Segregación en 180 días calendario.",
            "Municipalities approve Segregation Programs within 180 calendar days.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DS 001-2022-MINAM substantially updates DL 1278 and Law 29419 implementing "
        "regulations. It strengthens segregation-at-source, recycler formalization pathways, "
        "SIGERSOL traceability and transport rules relevant to post-consumer plastics recovery."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_001_2022_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_001_2022_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(
        f"  Excel: {REPO_RAW}/output/peru-ds-001-2022-minam/"
        "Peru_DS_001_2022_MINAM_4P_Index_Coding.xlsx"
    )
    print(
        f"  Word:  {REPO_RAW}/output/peru-ds-001-2022-minam/"
        "Peru_DS_001_2022_MINAM_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
