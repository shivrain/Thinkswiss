#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru DS 244-2019-EF (ICBP)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-ds-244-2019-ef")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-ds-244-2019-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    "Decreto Supremo N.° 244-2019-EF – Reglamento del Impuesto al Consumo de las Bolsas de "
    "Plástico"
)
POLICY_NAME_EN = (
    "Supreme Decree No. 244-2019-EF – Regulation of the Plastic Bag Consumption Tax (ICBP)"
)
POLICY_URL = "https://www.gob.pe/institucion/sunat/normas-legales/292916-244-2019-ef"


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2019,
    "policy_objective": bi(
        "Decreto supremo que aprueba el reglamento del impuesto al consumo de las bolsas de "
        "plástico (ICBP) creado por el artículo 12 de la Ley N.° 30884, implementando la "
        "gravamen sobre la adquisición de bolsas de base polimérica en establecimientos "
        "comerciales de contribuyentes del IGV, designando a los titulares de establecimientos "
        "como agentes de percepción, fijando el cronograma gradual del impuesto, las "
        "excepciones (alimentos a granel, higiene/salud, bolsas biodegradables certificadas) "
        "y los mecanismos de percepción, declaración y pago administrados por SUNAT.",
        "Supreme decree approving the regulation of the plastic-bag consumption tax (ICBP) "
        "created by Article 12 of Law No. 30884, implementing the levy on acquisition of "
        "polymer-based bags in commercial establishments of VAT taxpayers, designating "
        "establishment holders as withholding agents, setting the graduated tax schedule, "
        "exemptions (bulk food, hygiene/health, certified biodegradable bags) and "
        "collection, reporting and payment mechanics administered by SUNAT.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "Reglamento Art. 5: impuesto gradual por bolsa — S/0.10 (2019), S/0.20 (2020), "
        "S/0.30 (2021), S/0.40 (2022), S/0.50 (2023 en adelante). Ley 30884 Art. 12.10: "
        "vigencia desde 1 ago. 2019. Art. 3: grava adquisición gratuita u onerosa de bolsas "
        "de plástico para cargar o llevar bienes enajenados.",
        "Regulation Art. 5: graduated per-bag tax — PEN 0.10 (2019), PEN 0.20 (2020), "
        "PEN 0.30 (2021), PEN 0.40 (2022), PEN 0.50 (2023 onward). Law 30884 Art. 12.10: "
        "effective from 1 Aug 2019. Art. 3: taxes free or onerous acquisition of plastic bags "
        "to carry or transport sold goods.",
    ),
    "policy_type": 0.75,
    "policy_type_justification": bi(
        "Decreto Supremo N.° 244-2019-EF aprobado por el Poder Ejecutivo (2 ago. 2019, "
        "publicado 3 ago. 2019 en El Peruano), refrendado por el Ministro de Economía y "
        "Finanzas. No enmendado sustantivamente (incrementos de tasa programados directamente "
        "por Ley 30884). Reglamento tributario ejecutivo que implementa Art. 12 de Ley "
        "30884. Clasificado 0.75 (decreto supremo/regulación ejecutiva).",
        "Supreme Decree No. 244-2019-EF approved by the Executive (2 Aug 2019, published "
        "3 Aug 2019 in El Peruano), countersigned by the Minister of Economy and Finance. "
        "Not substantively amended (rate increases scheduled directly by Law 30884). "
        "Executive tax regulation implementing Law 30884 Art. 12. Classified 0.75 (supreme "
        "decree/executive regulation).",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "comercio minorista, retail, administración tributaria (SUNAT), MEF, alimentos y "
        "bebidas, servicios, comercio electrónico, manufactura de bolsas plásticas, "
        "importación de empaques",
        "retail commerce, tax administration (SUNAT), MEF, food & beverage, services, "
        "e-commerce, plastic bag manufacturing, packaging imports",
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": bi(
        "consumo, producción (bolsas biodegradables certificadas), disposición",
        "consumption, production (certified biodegradable bags), disposal",
    ),
    "policy_budget": 1.0,
    "policy_budget_text": bi(
        "Implementa el impuesto al consumo de bolsas de plástico (ICBP) creado por Ley "
        "30884 Art. 12, cuyos ingresos constituyen ingreso del tesoro público administrado "
        "por SUNAT. Art. 5 fija montos graduales S/0.10–S/0.50 por bolsa.",
        "Implements the plastic-bag consumption tax (ICBP) created by Law 30884 Art. 12, "
        "whose revenue constitutes public treasury income administered by SUNAT. Art. 5 sets "
        "graduated amounts from PEN 0.10 to PEN 0.50 per bag.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "DS Art. 1: 'Apruébase el Reglamento del impuesto al consumo de las bolsas de "
            "plástico' — seis artículos que forman parte integrante del decreto supremo.",
            "DS Art. 1: 'The Regulation of the plastic-bag consumption tax is approved' — "
            "six articles forming an integral part of the supreme decree.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Apruébase' — aprobación directa del reglamento ICBP. MEF/SUNAT. Vigente desde "
            "publicación 3 ago. 2019.",
            "'Is approved' — direct approval of ICBP regulation. MEF/SUNAT. Effective from "
            "publication 3 Aug 2019.",
        ),
        "comments": bi(
            "Acto habilitante que pone en vigor el reglamento del impuesto a bolsas plásticas.",
            "Enabling act bringing the plastic-bag tax regulation into force.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 1: objeto — reglamentar la aplicación del impuesto al consumo "
            "de las bolsas de plástico creado por Ley 30884 Art. 12.",
            "Regulation Art. 1: purpose — regulate application of the plastic-bag consumption "
            "tax created by Law 30884 Art. 12.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Objeto declarativo. Marco para instrumentos operativos Arts. 3–6. SUNAT "
            "administrador tributario.",
            "Declaratory purpose. Framework for operative Arts. 3–6. SUNAT tax administrator.",
        ),
        "comments": bi(
            "Objeto del reglamento tributario sobre bolsas de plástico.",
            "Purpose of the plastic-bag tax regulation.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 2: definiciones — 'bolsas de plástico' como bolsas de base "
            "polimérica; establecimiento físico o virtual; impuesto; Ley 30884; titular del "
            "establecimiento (contribuyente IGV).",
            "Regulation Art. 2: definitions — 'plastic bags' as polymer-based bags; physical "
            "or virtual establishment; tax; Law 30884; establishment holder (VAT taxpayer).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Definiciones vinculantes para ámbito, excepciones y agentes de percepción. "
            "Incluye comercio electrónico (establecimiento virtual).",
            "Binding definitions for scope, exemptions and withholding agents. Includes "
            "e-commerce (virtual establishment).",
        ),
        "comments": bi(
            "Define explícitamente materiales poliméricos (bolsas de base polimérica).",
            "Explicitly defines polymer materials (polymer-based bags).",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 3: 'El impuesto grava la adquisición, a título gratuito u "
            "oneroso, de bolsas de plástico' cuya finalidad sea cargar o llevar bienes "
            "enajenados en establecimientos que distribuyan dichas bolsas; incluye bolsas "
            "en comercio virtual entregadas al adquirente.",
            "Regulation Art. 3: 'The tax levies acquisition, whether free or onerous, of "
            "plastic bags' intended to carry or transport goods sold in establishments "
            "distributing such bags; includes bags in virtual commerce delivered to the "
            "purchaser.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Grava la adquisición' — obligación tributaria directa. Titulares IGV en "
            "establecimientos físicos y virtuales. Percepción Art. 6.",
            "'Levies acquisition' — direct tax obligation. VAT taxpayers in physical and "
            "virtual establishments. Withholding Art. 6.",
        ),
        "comments": bi(
            "Ámbito material del gravamen sobre bolsas plásticas en puntos de venta IGV.",
            "Material scope of plastic-bag levy at VAT-liable points of sale.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 4.1: exceptuada la adquisición de (a) bolsas para alimentos a "
            "granel u origen animal y asepsia/inocuidad; (b) bolsas por limpieza, higiene o "
            "salud; (c) bolsas biodegradables nacionales o importadas con certificado de "
            "laboratorio acreditado conforme glosario Ley 30884.",
            "Regulation Art. 4.1: acquisition exempt of (a) bags for bulk food or animal-origin "
            "food and asepsis/food safety; (b) bags required for cleaning, hygiene or health; "
            "(c) domestically produced or imported biodegradable bags with accredited-laboratory "
            "certificate per Law 30884 glossary.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Está exceptuada' — excepciones obligatorias. Certificación biodegradabilidad "
            "requerida (Art. 4.1c). Criterios sectoriales vía normas MINSA Art. 4.2.",
            "'Is exempt' — mandatory exemptions. Biodegradability certification required "
            "(Art. 4.1c). Sectoral criteria via MINSA norms Art. 4.2.",
        ),
        "comments": bi(
            "Exenciones que distinguen bolsas funcionales/higiénicas y alternativas "
            "biodegradables certificadas.",
            "Exemptions distinguishing functional/hygienic bags and certified biodegradable "
            "alternatives.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 4.2: 'Las normas sobre la materia establecen los criterios para "
            "la aplicación de las excepciones' señaladas en el Art. 4.1 (alimentos, higiene, "
            "salud).",
            "Regulation Art. 4.2: 'The norms on the matter establish the criteria for applying "
            "the exemptions' listed in Art. 4.1 (food, hygiene, health).",
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Remisión a normas sectoriales externas ('establecen los criterios'). No fija "
            "criterios directamente en el reglamento — habilitación/referencia.",
            "Referral to external sectoral norms ('establish the criteria'). Does not set "
            "criteria directly in the regulation — enabling/reference.",
        ),
        "comments": bi(
            "Criterios de exención delegados a normas MINSA/sanitarias — verificar normas "
            "aplicables para implementación completa.",
            "Exemption criteria delegated to MINSA/health norms — verify applicable norms for "
            "full implementation.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 5: monto gradual del impuesto por cada bolsa de plástico "
            "adquirida — S/0.10 (2019), S/0.20 (2020), S/0.30 (2021), S/0.40 (2022), "
            "S/0.50 (2023 en adelante).",
            "Regulation Art. 5: graduated tax amount per plastic bag acquired — PEN 0.10 "
            "(2019), PEN 0.20 (2020), PEN 0.30 (2021), PEN 0.40 (2022), PEN 0.50 "
            "(2023 onward).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 1.0,
        "instrument_implementation_text": bi(
            "Cronograma cuantitativo explícito en texto. Tasas posteriores a 2022 fijadas "
            "directamente por Ley 30884 Art. 12.5. SUNAT percepción y fiscalización.",
            "Explicit quantitative schedule in text. Post-2022 rates set directly by Law 30884 "
            "Art. 12.5. SUNAT collection and enforcement.",
        ),
        "comments": bi(
            "Instrumento central del ICBP — gravamen por bolsa con escalamiento progresivo.",
            "Central ICBP instrument — per-bag levy with progressive escalation.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 6.1: titulares de establecimientos que transfieran bolsas de "
            "plástico (gratuito u oneroso) para cargar o llevar bienes enajenados 'son "
            "agentes de percepción del impuesto'.",
            "Regulation Art. 6.1: establishment holders transferring plastic bags (free or "
            "onerous) to carry or transport sold goods 'are withholding agents of the tax'.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Son agentes de percepción' — designación obligatoria de comercios IGV. "
            "Vinculado a comprobante de pago Art. 6.2.",
            "'Are withholding agents' — mandatory designation of VAT-liable retailers. Linked "
            "to payment receipt Art. 6.2.",
        ),
        "comments": bi(
            "Mecanismo de recaudación en punto de venta para bolsas plásticas.",
            "Point-of-sale collection mechanism for plastic bags.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 6.2: 'La percepción del impuesto se efectúa en el momento en que "
            "se emita el comprobante de pago correspondiente'.",
            "Regulation Art. 6.2: 'Tax withholding is effected at the moment the corresponding "
            "payment receipt is issued'.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Se efectúa en el momento' — obligación temporal vinculada a comprobante fiscal. "
            "SUNAT supervisa agentes.",
            "'Is effected at the moment' — timing obligation linked to tax receipt. SUNAT "
            "oversees agents.",
        ),
        "comments": bi(
            "Momento de percepción alineado con facturación electrónica/comprobantes SUNAT.",
            "Withholding timing aligned with SUNAT electronic invoicing/receipts.",
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 6.3: agentes de percepción 'realizan el pago del impuesto "
            "percibido dentro del plazo aplicable a los tributos de determinación mensual' "
            "(literal b) Art. 29 Código Tributario).",
            "Regulation Art. 6.3: withholding agents 'pay the tax withheld within the deadline "
            "applicable to monthly-determination taxes' (Tax Code Art. 29(b)).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Realizan el pago... dentro del plazo' — obligación de pago mensual. Plazo "
            "Código Tributario.",
            "'Pay... within the deadline' — monthly payment obligation. Tax Code deadline.",
        ),
        "comments": bi(
            "Remisión de impuesto percibido al tesoro vía agentes.",
            "Remittance of withheld tax to treasury via agents.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Reglamento Art. 6.4: declaración y pago del agente de percepción 'se efectúan en "
            "la forma y condiciones que establezca la SUNAT mediante resolución de "
            "superintendencia'.",
            "Regulation Art. 6.4: agent declaration and payment 'are made in the form and "
            "conditions that SUNAT establishes by superintendency resolution'.",
        ),
        "instrument_in_force": 0,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Procedimientos delegados a resolución SUNAT ('establezca'). Obligación de "
            "declarar/pagar existe (Art. 6.3) pero forma depende de acto administrativo "
            "posterior — in_force 0 para habilitación procedimental.",
            "Procedures delegated to SUNAT resolution ('establishes'). Declare/pay obligation "
            "exists (Art. 6.3) but form depends on subsequent administrative act — in_force 0 "
            "for procedural enabling.",
        ),
        "comments": bi(
            "SUNAT debe emitir resolución de superintendencia con formularios/plazos. Verificar "
            "resoluciones SUNAT complementarias.",
            "SUNAT must issue superintendency resolution with forms/deadlines. Verify "
            "complementary SUNAT resolutions.",
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
        "Signed: 2 August 2019 | Published: 3 August 2019 (El Peruano)\n"
        "Not substantively amended (rate increases scheduled by Law 30884)\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Supreme Decree No. 244-2019-EF "
        "regulating the plastic-bag consumption tax (ICBP) under Law No. 30884."
    )

    sections = [
        (
            "Supreme Decree Art. 1 — Approval",
            "Apruébase el Reglamento del impuesto al consumo de las bolsas de plástico, el que "
            "consta de seis (6) artículos y forma parte integrante de este Decreto Supremo.",
            "The Regulation of the plastic-bag consumption tax is approved, consisting of six "
            "(6) articles and forming an integral part of this Supreme Decree.",
        ),
        (
            "Regulation Art. 1 — Purpose",
            "La presente Reglamento tiene por objeto reglamentar la aplicación del impuesto al "
            "consumo de las bolsas de plástico.",
            "This Regulation regulates application of the plastic-bag consumption tax.",
        ),
        (
            "Regulation Art. 2 — Definitions",
            "Bolsas de plástico: bolsas de base polimérica. Establecimiento: lugar físico o "
            "virtual de enajenación de bienes y/o prestación de servicios. Titular: "
            "contribuyente del IGV.",
            "Plastic bags: polymer-based bags. Establishment: physical or virtual place of "
            "sale of goods and/or provision of services. Holder: VAT taxpayer.",
        ),
        (
            "Regulation Art. 3 — Scope",
            "El impuesto grava la adquisición, a título gratuito u oneroso, de bolsas de "
            "plástico, cuya finalidad sea cargar o llevar bienes enajenados en los "
            "establecimientos que distribuyan dichas bolsas.",
            "The tax levies acquisition, whether free or onerous, of plastic bags intended to "
            "carry or transport goods sold in establishments distributing such bags.",
        ),
        (
            "Regulation Art. 4 — Exemptions",
            "Exceptuada la adquisición de bolsas para alimentos a granel, higiene/salud, y "
            "bolsas biodegradables con certificado de laboratorio acreditado conforme al "
            "glosario de la Ley 30884.",
            "Acquisition exempt for bulk-food bags, hygiene/health bags, and biodegradable "
            "bags with accredited-laboratory certificate per Law 30884 glossary.",
        ),
        (
            "Regulation Art. 5 — Tax schedule",
            "Monto gradual por bolsa: S/0.10 (2019), S/0.20 (2020), S/0.30 (2021), S/0.40 "
            "(2022), S/0.50 (2023 en adelante).",
            "Graduated per-bag amount: PEN 0.10 (2019), PEN 0.20 (2020), PEN 0.30 (2021), "
            "PEN 0.40 (2022), PEN 0.50 (2023 onward).",
        ),
        (
            "Regulation Art. 6 — Withholding agents",
            "Los titulares de establecimientos son agentes de percepción. Percepción al emitir "
            "comprobante de pago. Pago dentro del plazo de tributos mensuales. Declaración y "
            "pago en forma que establezca SUNAT.",
            "Establishment holders are withholding agents. Withholding at payment-receipt "
            "issuance. Payment within monthly-tax deadline. Declaration and payment as SUNAT "
            "establishes.",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: DS 244-2019-EF implements Law 30884 Art. 12 (plastic-bag consumption tax). "
        "SUNAT administers collection; tax revenue constitutes public treasury income. Rate "
        "increases from 2023 onward are set directly in Law 30884 rather than by amendment "
        "to this decree."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_DS_244_2019_EF_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_DS_244_2019_EF_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(
        f"  Excel: {REPO_RAW}/output/peru-ds-244-2019-ef/"
        "Peru_DS_244_2019_EF_4P_Index_Coding.xlsx"
    )
    print(
        f"  Word:  {REPO_RAW}/output/peru-ds-244-2019-ef/"
        "Peru_DS_244_2019_EF_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
