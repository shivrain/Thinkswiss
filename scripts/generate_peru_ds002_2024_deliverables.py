#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 002-2024-MINAM."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-002-2024-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ds-002-2024-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    "Decreto Supremo N.° 002-2024-MINAM – Modifica el artículo 9 del Reglamento del "
    "D.L. N.° 1278"
)
POLICY_NAME_EN = (
    "Supreme Decree No. 002-2024-MINAM – Amendment to Article 9 of the Regulation of "
    "Legislative Decree No. 1278 (PLANRES Provision)"
)
POLICY_URL = (
    "https://www.gob.pe/institucion/minam/normas-legales/5425206-002-2024-minam"
)


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2024,
    "policy_objective": bi(
        "Decreto Supremo que modifica el Art. 9 del Reglamento del DL 1278 (DS 014-2017-MINAM) "
        "para actualizar el marco del PLANRES como instrumento nacional multisectorial de "
        "obligatorio cumplimiento; alinea su aprobación y articulación al ciclo de políticas "
        "nacionales (CEPLAN); refuerza funciones MINAM de seguimiento, evaluación, medidas "
        "correctivas y reportes anuales al Congreso y CEPLAN. Implementa parcialmente la DCF "
        "de la Ley 31896 sobre funciones PLANRES del Art. 15(b) DL 1278.",
        "Supreme Decree amending Art. 9 of the DL 1278 Regulation (DS 014-2017-MINAM) to "
        "update the PLANRES framework as a mandatory national multisectoral instrument; "
        "aligns its approval and articulation to the national policy cycle (CEPLAN); "
        "strengthens MINAM follow-up, evaluation, corrective measures and annual reporting "
        "to Congress and CEPLAN. Partially implements Law 31896 DCF on PLANRES functions "
        "in DL 1278 Art. 15(b).",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "Art. 9 modificado: PLANRES se actualiza cada 10 años. Autoridades competentes "
        "remiten información al MINAM hasta último día hábil de abril de cada año. MINAM "
        "publica anualmente reportes de avance en portal institucional. MINAM remite informe "
        "anual de cumplimiento del PLANRES al CEPLAN dentro del plazo establecido por dicha "
        "entidad y a la Comisión de Ambiente del Congreso.",
        "Amended Art. 9: PLANRES updated every 10 years. Competent authorities submit "
        "information to MINAM by last business day of April each year. MINAM annually "
        "publishes progress reports on its institutional portal. MINAM submits annual PLANRES "
        "compliance report to CEPLAN within CEPLAN's established deadline and to the "
        "Congress Environment Commission.",
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Decreto Supremo N.° 002-2024-MINAM aprobado por el Poder Ejecutivo (3 abr. 2024, "
        "publicado 4 abr. 2024 en El Peruano), refrendado por MINAM. Enmienda puntual al "
        "Reglamento DS 014-2017-MINAM (Art. 9 PLANRES). Vinculado a Ley 31896. Clasificado "
        "0.75 (decreto supremo ejecutivo de modificación reglamentaria).",
        "Supreme Decree No. 002-2024-MINAM approved by the Executive (3 Apr 2024, published "
        "4 Apr 2024 in El Peruano), countersigned by MINAM. Targeted amendment to DS "
        "014-2017-MINAM Regulation (Art. 9 PLANRES). Linked to Law 31896. Classified 0.75 "
        "(executive supreme-decree regulatory amendment).",
    ),
    "policy_integration": 1,
    "policy_sectors_list": bi(
        "ambiente, planificación pública (CEPLAN), gobierno nacional, regional y local, "
        "sector privado, sociedad civil, residuos sólidos, reciclaje",
        "environment, public planning (CEPLAN), national, regional and local government, "
        "private sector, civil society, solid waste, recycling",
    ),
    "policy_circularity": 1,
    "policy_lifecycle_phases_list": bi(
        "planificación (ciclo de vida completo), consumo, recolección, reciclaje, valorización, "
        "disposición",
        "planning (whole life cycle), consumption, collection, recycling, valorization, disposal",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": bi(
        "Art. 2: implementación financiada con presupuesto institucional del MINAM y actores "
        "de los tres niveles de gobierno, sin recursos adicionales del Tesoro Público.",
        "Art. 2: implementation financed from MINAM institutional budget and actors at all "
        "three government levels, without additional Treasury resources.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 1: modifica el Art. 9 del Reglamento del DL 1278 (DS 014-2017-MINAM) sobre "
            "el Plan Nacional de Gestión Integral de Residuos Sólidos (PLANRES).",
            "Art. 1: amends Art. 9 of the DL 1278 Regulation (DS 014-2017-MINAM) on the "
            "National Comprehensive Solid Waste Management Plan (PLANRES).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Acto habilitante de la nueva redacción del Art. 9 PLANRES. Vigencia desde "
            "publicación 4 abr. 2024.",
            "Enabling act for the new Art. 9 PLANRES text. Effective from publication "
            "4 Apr 2024.",
        ),
        "comments": bi(
            "Implementa mandato Ley 31896 DCF sobre funciones PLANRES del Art. 15(b).",
            "Implements Law 31896 DCF mandate on PLANRES functions in Art. 15(b).",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (1er párrafo): PLANRES es instrumento nacional multisectorial de "
            "obligatorio cumplimiento para definir, orientar y articular acciones y metas a "
            "corto, mediano y largo plazo de gobierno nacional, regional y local, sector "
            "privado y sociedad civil, incluyendo compromisos nacionales e internacionales.",
            "Art. 9 (1st paragraph): PLANRES is a mandatory national multisectoral instrument "
            "to define, orient and articulate short-, medium- and long-term actions and "
            "targets for national, regional and local government, private sector and civil "
            "society, including national and international commitments.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Definición y propósito declarativo del instrumento de planificación nacional. "
            "Marco para metas de gestión integral de residuos/plásticos.",
            "Declaratory definition and purpose of national planning instrument. Framework "
            "for integrated waste/plastics management targets.",
        ),
        "comments": bi(
            "PLANRES abarca planificación de todo el ciclo de vida de residuos sólidos.",
            "PLANRES covers whole-life-cycle solid-waste planning.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (2º párrafo): PLANRES se aprueba por Decreto Supremo a propuesta del "
            "MINAM como ente rector, en coordinación con sectores intervinientes, en marco "
            "de normativa vigente que regula las Políticas Nacionales (alineación CEPLAN).",
            "Art. 9 (2nd paragraph): PLANRES is approved by Supreme Decree upon MINAM "
            "proposal as national steering authority, coordinating with intervening sectors, "
            "within the regulations governing National Policies (CEPLAN alignment).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Procedimiento de aprobación vinculado al ciclo de políticas nacionales. MINAM "
            "como proponente rector.",
            "Approval procedure linked to national policy cycle. MINAM as steering proponent.",
        ),
        "comments": bi(
            "Alinea PLANRES al ciclo de planificación nacional liderado por CEPLAN.",
            "Aligns PLANRES to national planning cycle led by CEPLAN.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (3er párrafo): PLANRES se actualiza cada diez (10) años, en base al "
            "análisis del cumplimiento de objetivos específicos y metas, e identificación "
            "de nuevas necesidades de solución.",
            "Art. 9 (3rd paragraph): PLANRES is updated every ten (10) years, based on "
            "analysis of compliance with specific objectives and targets, and identification "
            "of new solution needs.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Ciclo decenal de actualización. Próxima revisión según calendario PLANRES "
            "vigente (PLANRES 2016–2024 en transición).",
            "Decennial update cycle. Next revision per current PLANRES calendar "
            "(PLANRES 2016–2024 in transition).",
        ),
        "comments": bi(
            "Meta temporal explícita de revisión del plan nacional de residuos.",
            "Explicit temporal target for national waste-plan revision.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (4º párrafo): articulación intersectorial, seguimiento y evaluación del "
            "PLANRES corresponde al MINAM; autoridades competentes obligadas a remitir "
            "información requerida hasta último día hábil de abril de cada año, bajo "
            "responsabilidad.",
            "Art. 9 (4th paragraph): intersectoral articulation, follow-up and evaluation of "
            "PLANRES is MINAM's responsibility; competent authorities must submit required "
            "information by the last business day of April each year, under liability.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Obligación anual de reporte de autoridades. MINAM evalúa cumplimiento de metas "
            "PLANRES (incl. valorización/reciclaje industrial).",
            "Annual authority reporting obligation. MINAM evaluates PLANRES target compliance "
            "(incl. valorization/industrial recycling).",
        ),
        "comments": bi(
            "Mecanismo de monitoreo de metas de gestión de plásticos/residuos a nivel nacional.",
            "Monitoring mechanism for national plastics/waste management targets.",
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (4º párrafo, continuación): MINAM coordina con autoridades competentes "
            "la implementación de medidas correctivas para que cumplan metas planificadas "
            "del PLANRES a su cargo.",
            "Art. 9 (4th paragraph, continued): MINAM coordinates with competent authorities "
            "implementation of corrective measures so they meet their assigned PLANRES targets.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Poder correctivo MINAM vinculado a Ley 31896 Art. 15(b). Sin sanciones "
            "específicas en el artículo.",
            "MINAM corrective power linked to Law 31896 Art. 15(b). No specific sanctions "
            "in the article.",
        ),
        "comments": bi(
            "Herramienta de enforcement de planificación nacional de reciclaje/valorización.",
            "Enforcement tool for national recycling/valorization planning.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (5º párrafo): MINAM publica anualmente en portal institucional reportes "
            "de avance y cumplimiento de objetivos y metas del PLANRES, así como medidas "
            "correctivas, elaborados a partir de información de entidades con compromisos.",
            "Art. 9 (5th paragraph): MINAM annually publishes on its institutional portal "
            "reports on progress and compliance with PLANRES objectives and targets, as well "
            "as corrective measures, based on information from entities with commitments.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Transparencia pública anual obligatoria. Publicación en portal MINAM.",
            "Mandatory annual public transparency. Publication on MINAM portal.",
        ),
        "comments": bi(
            "Rendición pública de avances en gestión integral de residuos/plásticos.",
            "Public accountability on integrated waste/plastics management progress.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Planning",
        "instrument_description": bi(
            "Art. 9 (6º párrafo): MINAM remite al CEPLAN informe anual con reporte de "
            "cumplimiento del PLANRES dentro del plazo establecido por CEPLAN, así como a "
            "la Comisión de Pueblos Andinos, Amazónicos y Afroperuanos, Ambiente y Ecología "
            "del Congreso.",
            "Art. 9 (6th paragraph): MINAM submits to CEPLAN an annual report on PLANRES "
            "compliance within CEPLAN's established deadline, and to the Congress Commission "
            "on Andean, Amazonian and Afro-Peruvian Peoples, Environment and Ecology.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Integración PLANRES al sistema nacional de planificación (CEPLAN). Reporte "
            "parlamentario anual.",
            "PLANRES integration into national planning system (CEPLAN). Annual "
            "parliamentary reporting.",
        ),
        "comments": bi(
            "Vinculación explícita PLANRES–CEPLAN para planificación de política nacional.",
            "Explicit PLANRES–CEPLAN linkage for national policy planning.",
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
        "Signed: 3 April 2024 | Published: 4 April 2024 (El Peruano)\n"
        "Amends DS 014-2017-MINAM Art. 9 (PLANRES) | Linked to Law 31896\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 002-2024-MINAM "
        "amending Article 9 on PLANRES in the DL 1278 Regulation."
    )

    sections = [
        (
            "Art. 1 — Amendment",
            "Modifica el artículo 9 del Reglamento del Decreto Legislativo 1278.",
            "Amends Article 9 of the Regulation of Legislative Decree 1278.",
        ),
        (
            "Art. 9 — PLANRES instrument",
            "PLANRES es instrumento nacional multisectorial de obligatorio cumplimiento para "
            "definir, orientar y articular acciones y metas a corto, mediano y largo plazo.",
            "PLANRES is a mandatory national multisectoral instrument to define, orient and "
            "articulate short-, medium- and long-term actions and targets.",
        ),
        (
            "Art. 9 — Approval and CEPLAN alignment",
            "PLANRES se aprueba por Decreto Supremo a propuesta del MINAM, en marco de "
            "normativa de Políticas Nacionales.",
            "PLANRES is approved by Supreme Decree upon MINAM proposal, within National "
            "Policies regulations.",
        ),
        (
            "Art. 9 — Monitoring and reporting",
            "MINAM realiza seguimiento y evaluación; autoridades remiten información hasta "
            "abril de cada año; MINAM publica reportes anuales y remite informe al CEPLAN y "
            "al Congreso.",
            "MINAM conducts follow-up and evaluation; authorities submit information by April "
            "each year; MINAM publishes annual reports and submits report to CEPLAN and Congress.",
        ),
        (
            "Art. 9 — Update cycle",
            "PLANRES se actualiza cada diez (10) años.",
            "PLANRES is updated every ten (10) years.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DS 002-2024-MINAM updates the PLANRES provision within the DL 1278 Regulation, "
        "aligning national solid-waste planning to CEPLAN's national policy cycle and "
        "strengthening MINAM monitoring, corrective measures and annual reporting."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_002_2024_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_002_2024_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(
        f"  Excel: {REPO_RAW}/output/peru-ds-002-2024-minam/"
        "Peru_DS_002_2024_MINAM_4P_Index_Coding.xlsx"
    )
    print(
        f"  Word:  {REPO_RAW}/output/peru-ds-002-2024-minam/"
        "Peru_DS_002_2024_MINAM_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
