#!/usr/bin/env python3
"""Generate 4P Index Excel and English translation for Peru RM 122-2021-MINAM (Perú Limpio)."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

OUTPUT_DIR = Path("/workspace/output/peru-rm-122-2021-minam")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

BRANCH = "cursor/peru-rm-122-2021-4p-index-d350"
REPO_RAW = f"https://github.com/shivrain/Thinkswiss/raw/{BRANCH}"

POLICY_NAME_ES = (
    'Resolución Ministerial N.° 122-2021-MINAM – "Perú Limpio": Estrategia de educación y '
    "comunicación sobre consumo responsable, valorización y gestión integrada de los residuos "
    "sólidos"
)
POLICY_NAME_EN = (
    'Ministerial Resolution No. 122-2021-MINAM – "Peru Limpio": Education and Communication '
    "Strategy on Responsible Consumption, Valorization and Integrated Solid Waste Management"
)
POLICY_URL = (
    "https://www.gob.pe/institucion/minam/normas-legales/2024227-122-2021-minam"
)


def bi(es: str, en: str) -> str:
    return f"{es} / {en}"


POLICY = {
    "policy_name": bi(POLICY_NAME_ES, POLICY_NAME_EN),
    "policy_url": POLICY_URL,
    "policy_year": 2021,
    "policy_objective": bi(
        "Estrategia nacional de educación y comunicación aprobada por MINAM que fomenta "
        "consumo responsable, minimización, segregación y valorización de residuos sólidos "
        "hacia economía circular, dirigida a sector público (Poder Ejecutivo), gobiernos "
        "locales (programas EDUCCA), sector privado y ciudadanía; prioriza consumo "
        "responsable (incl. reducción de plásticos de un solo uso), manejo adecuado de "
        "residuos y cultura de pago de arbitrios; incluye campañas como #MenosPlásticoMásVida.",
        "MINAM-approved national education and communication strategy promoting responsible "
        "consumption, minimization, segregation and valorization of solid waste toward a "
        "circular economy, targeting public sector (Executive Branch), local governments "
        "(EDUCCA programs), private sector and citizens; prioritizes responsible consumption "
        "(incl. single-use plastic reduction), adequate waste management and municipal-fee "
        "payment culture; includes campaigns such as #MenosPlasticoMasVida.",
    ),
    "policy_target": 1,
    "policy_target_text": bi(
        "OE2: 70 municipalidades (capitales de provincia) en 2021 y 130 en 2025 con programas "
        "EDUCCA. Meta Bicentenario: transición a economía circular minimizando residuos y "
        "valorizando aprovechables. Acción 8.1b: no usar plásticos de un solo uso. Vinculada "
        "Plan Nacional de Competitividad y Productividad 2030 (Medida 9.2).",
        "OE2: 70 municipalities (provincial capitals) in 2021 and 130 in 2025 with EDUCCA "
        "programs. Bicentenary goal: transition to circular economy minimizing waste and "
        "valorizing recoverables. Action 8.1b: do not use single-use plastics. Linked to "
        "National Competitiveness and Productivity Plan 2030 (Measure 9.2).",
    ),
    "policy_type": 0.5,
    "policy_type_justification": bi(
        "Resolución Ministerial N.° 122-2021-MINAM aprobada por el Ministro del Ambiente (15 "
        "jul. 2021, publicada 17 jul. 2021 en El Peruano). No enmendada. Estrategia nacional "
        "de educación/comunicación bajo rectoría MINAM (DGECIA), vinculada DS 237-2019-EF "
        "Plan Nacional de Competitividad y Productividad Medida 9.2. Clasificado 0.5 "
        "(resolución ministerial/estrategia nacional).",
        "Ministerial Resolution No. 122-2021-MINAM approved by the Minister of Environment "
        "(15 Jul 2021, published 17 Jul 2021 in El Peruano). Not amended. National "
        "education/communication strategy under MINAM steering (DGECIA), linked to DS "
        "237-2019-EF National Competitiveness and Productivity Plan Measure 9.2. Classified "
        "0.5 (ministerial resolution/national strategy).",
    ),
    "policy_integration": 0.75,
    "policy_sectors_list": bi(
        "ambiente, educación (MINEDU), gobiernos locales, sociedad civil, sector privado, "
        "municipalidades, reciclaje, consumo, comunicación",
        "environment, education (MINEDU), local governments, civil society, private sector, "
        "municipalities, recycling, consumption, communication",
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": bi(
        "consumo, recolección, reciclaje, valorización, disposición (cambio de comportamiento)",
        "consumption, collection, recycling, valorization, disposal (behavior change)",
    ),
    "policy_budget": 0.5,
    "policy_budget_text": bi(
        "Implementación a cargo DGECIA con apoyo DGRS/DGCA vía POI institucional MINAM. "
        "Vinculada Plan Nacional de Competitividad y Productividad 2030. Sin asignación "
        "presupuestal específica en la RM.",
        "Implementation by DGECIA with DGRS/DGCA support via MINAM institutional POI. "
        "Linked to National Competitiveness and Productivity Plan 2030. No specific budget "
        "allocation in the ministerial resolution.",
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "RM Art. 1: 'Apruébase' el documento 'Perú Limpio: Estrategia de educación y "
            "comunicación sobre consumo responsable, valorización y gestión integrada de los "
            "residuos sólidos' como Anexo integrante.",
            "RM Art. 1: 'Approves' the document 'Peru Limpio: Education and Communication "
            "Strategy on Responsible Consumption, Valorization and Integrated Solid Waste "
            "Management' as integral Annex.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Apruébase' — vigencia desde publicación 17 jul. 2021. MINAM/DGECIA implementa.",
            "'Is approved' — effective from publication 17 Jul 2021. MINAM/DGECIA implements.",
        ),
        "comments": bi(
            "Acto habilitante de la estrategia nacional Perú Limpio.",
            "Enabling act for the national Peru Limpio strategy.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Estrategia §6.1 — Objetivo General: fomentar y fortalecer buenas prácticas "
            "ambientales para consumo responsable, minimización, segregación y valorización "
            "de residuos sólidos, articulando actores públicos y privados hacia economía "
            "circular.",
            "Strategy §6.1 — General Objective: foster and strengthen good environmental "
            "practices for responsible consumption, minimization, segregation and valorization "
            "of solid waste, coordinating public and private actors toward circular economy.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Objetivo declarativo estratégico. Marco para 4 objetivos específicos y 3 temas "
            "priorizados.",
            "Declaratory strategic objective. Framework for 4 specific objectives and 3 "
            "priority themes.",
        ),
        "comments": bi(
            "Marco estratégico de cambio de comportamiento en residuos/plásticos.",
            "Strategic behavior-change framework for waste/plastics.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Estrategia §8.1 — Tema priorizado Consumo responsable: acciones incluyen "
            "'Prevenir la generación de residuos sólidos: no usando plásticos de un solo uso, "
            "usando productos reutilizables' y empaques reciclables/reutilizables/compostables.",
            "Strategy §8.1 — Priority theme Responsible consumption: actions include 'Prevent "
            "solid-waste generation: not using single-use plastics, using reusable products' "
            "and recyclable/reusable/compostable packaging.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Acción explícita anti plástico de un solo uso en tema priorizado. Promoción vía "
            "campañas OE4 (incl. #MenosPlásticoMásVida).",
            "Explicit anti-single-use-plastic action in priority theme. Promotion via OE4 "
            "campaigns (incl. #MenosPlasticoMasVida).",
        ),
        "comments": bi(
            "Instrumento directo de reducción de plásticos de un solo uso.",
            "Direct single-use plastic reduction instrument.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "Estrategia §8.2 — Tema priorizado Manejo adecuado de residuos: segregación "
            "NTP 900.058.2019 (código colores), valorización/reciclaje (incl. botellas PET), "
            "formalización de recicladores, entrega segregada a operadores autorizados.",
            "Strategy §8.2 — Priority theme Adequate waste management: segregation per NTP "
            "900.058.2019 (color code), valorization/recycling (incl. PET bottles), recycler "
            "formalization, segregated delivery to authorized operators.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Mención explícita botellas PET y reciclaje. Vinculado DL 1278 y segregación "
            "obligatoria generadores municipales.",
            "Explicit mention of PET bottles and recycling. Linked to DL 1278 and mandatory "
            "municipal-generator segregation.",
        ),
        "comments": bi(
            "Comportamiento de segregación/reciclaje de flujos plásticos.",
            "Segregation/recycling behavior for plastic streams.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE1 — Promover que entidades del sector público (énfasis Poder Ejecutivo) "
            "desarrollen buenas prácticas de consumo responsable, manejo de residuos y "
            "cultura de pago de arbitrios y lideren prácticas sostenibles.",
            "OE1 — Promote that public-sector entities (emphasis Executive Branch) develop "
            "good practices on responsible consumption, waste management and municipal-fee "
            "payment culture and lead sustainable practices.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Promover' y 'lideren' — obligación orientadora para 18 ministerios y organismos "
            "adscritos. Líneas de acción con asistencia técnica MINAM.",
            "'Promote' and 'lead' — guiding obligation for 18 ministries and attached agencies. "
            "Action lines with MINAM technical assistance.",
        ),
        "comments": bi(
            "Cambio de comportamiento en administración pública.",
            "Behavior change in public administration.",
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE1 Línea b: asistencia técnica para que entidades públicas cumplan normativa de "
            "residuos sólidos, ecoeficiencia y 'reducción de plásticos de un solo uso'; "
            "articulación con estrategia de ecoeficiencia institucional.",
            "OE1 Line b: technical assistance for public entities to comply with solid-waste "
            "regulations, eco-efficiency and 'single-use plastic reduction'; linkage with "
            "institutional eco-efficiency strategy.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Reducción de plásticos de un solo uso explícita en línea de acción. Vinculado "
            "Ley 30884 y DS 013-2018-MINAM. Asistencia técnica MINAM.",
            "Single-use plastic reduction explicit in action line. Linked to Law 30884 and DS "
            "013-2018-MINAM. MINAM technical assistance.",
        ),
        "comments": bi(
            "Instrumento operativo de reducción plástica en sector público.",
            "Operative plastic-reduction instrument in public sector.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Collection",
        "instrument_description": bi(
            "OE2 — Institucionalizar educación ambiental en gobiernos locales vía programas "
            "municipales EDUCCA; meta: 70 municipalidades (capitales provincia) en 2021 y 130 "
            "en 2025.",
            "OE2 — Institutionalize environmental education in local governments via municipal "
            "EDUCCA programs; target: 70 municipalities (provincial capitals) in 2021 and 130 "
            "in 2025.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "Metas cuantificadas 70/130 municipalidades. Asistencia técnica diseño e "
            "implementación EDUCCA.",
            "Quantified targets 70/130 municipalities. Technical assistance for EDUCCA design "
            "and implementation.",
        ),
        "comments": bi(
            "Educación municipal sobre segregación y consumo responsable (plásticos).",
            "Municipal education on segregation and responsible consumption (plastics).",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE3 — Promover que el sector privado implemente buenas prácticas de consumo "
            "responsable, manejo de residuos y cultura de arbitrios entre colaboradores y "
            "usuarios, en alianza con instituciones y sociedad civil.",
            "OE3 — Promote that the private sector implements good practices on responsible "
            "consumption, waste management and fee-payment culture among staff and users, in "
            "alliance with institutions and civil society.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Promover' sector privado. Línea b: impulsar campañas con clientes y proveedores.",
            "'Promote' private sector. Line b: drive campaigns with clients and suppliers.",
        ),
        "comments": bi(
            "Engagement empresarial en reducción de residuos/plásticos.",
            "Business engagement in waste/plastic reduction.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE4 — Promover que la ciudadanía implemente buenas prácticas de consumo "
            "responsable, manejo de residuos y cultura de pago de arbitrios; público: familias, "
            "niños, adolescentes y jóvenes.",
            "OE4 — Promote that citizens implement good practices on responsible consumption, "
            "waste management and fee-payment culture; audience: families, children, "
            "adolescents and youth.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "'Promover' ciudadanía. Campañas educativas articuladas sector público/privado.",
            "'Promote' citizens. Educational campaigns coordinated with public/private sector.",
        ),
        "comments": bi(
            "Cambio de comportamiento ciudadano en consumo y residuos.",
            "Citizen behavior change on consumption and waste.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE4 Línea a: 'Realización de campañas educativas' para consumo responsable, "
            "manejo de residuos y cultura de arbitrios — incluye campaña #MenosPlásticoMásVida "
            "como componente de Perú Limpio.",
            "OE4 Line a: 'Conduct educational campaigns' for responsible consumption, waste "
            "management and fee-payment culture — includes #MenosPlasticoMasVida campaign as "
            "a Peru Limpio component.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": bi(
            "'Realización de campañas educativas' — acción operativa de comunicación. "
            "#MenosPlásticoMásVida referenciada en metadata estrategia (campaña nacional "
            "anti-plástico).",
            "'Conduct educational campaigns' — operative communication action. "
            "#MenosPlasticoMasVida referenced in strategy metadata (national anti-plastic "
            "campaign).",
        ),
        "comments": bi(
            "Campaña #MenosPlásticoMásVida no aparece textualmente en PDF pero es componente "
            "documentado de Perú Limpio.",
            "#MenosPlasticoMasVida not in PDF text but documented Peru Limpio component.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "OE4 Línea b: impulsar junto al MINEDU implementación de buenas prácticas en la "
            "comunidad educativa.",
            "OE4 Line b: drive together with MINEDU implementation of good practices in the "
            "educational community.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Coordinación MINAM-MINEDU. Vinculado PLANEA 2017-2022 y PNEA.",
            "MINAM-MINEDU coordination. Linked to PLANEA 2017-2022 and National Environmental "
            "Education Policy.",
        ),
        "comments": bi(
            "Educación escolar sobre consumo responsable y plásticos.",
            "School education on responsible consumption and plastics.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Estrategia §8.3 — Tema priorizado Cultura de pago de arbitrios municipales para "
            "servicio de limpieza; incentivos por segregación en fuente (DL 1278 Art. 70 mod. "
            "DL 1501).",
            "Strategy §8.3 — Priority theme Culture of municipal fee payment for cleaning "
            "service; incentives for source segregation (DL 1278 Art. 70 as amended DL 1501).",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "Incentivos municipales por segregación vinculados a recolección selectiva. "
            "Sostenibilidad financiera servicio limpieza.",
            "Municipal segregation incentives linked to selective collection. Financial "
            "sustainability of cleaning service.",
        ),
        "comments": bi(
            "Incentivos económicos para segregación de residuos (incl. plásticos).",
            "Economic incentives for waste segregation (incl. plastics).",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Estrategia §10 — Implementación a cargo DGECIA con asistencia DGRS/DGCA; "
            "vinculada Plan Nacional de Competitividad y Productividad 2030; líneas de acción "
            "en POI de Dirección de Educación y Ciudadanía Ambiental.",
            "Strategy §10 — Implementation by DGECIA with DGRS/DGCA assistance; linked to "
            "National Competitiveness and Productivity Plan 2030; action lines in POI of "
            "Directorate of Environmental Education and Citizenship.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": bi(
            "DGECIA/DGRS/DGCA designados. POI institucional MINAM. Medida 9.2 PNCP.",
            "DGECIA/DGRS/DGCA designated. MINAM institutional POI. PNCP Measure 9.2.",
        ),
        "comments": bi(
            "Gobernanza de implementación de la estrategia.",
            "Strategy implementation governance.",
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": bi(
            "Diagnóstico estrategia: 68% de residuos plásticos generados son de un solo uso "
            "(bolsas, PET, poliestireno); solo 58% hogares segregan residuos; consumo plástico "
            "+50% 2008–2016.",
            "Strategy diagnosis: 68% of generated plastic waste is single-use (bags, PET, "
            "polystyrene); only 58% of households segregate waste; plastic consumption +50% "
            "2008–2016.",
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": bi(
            "Diagnóstico contextual (no mandato). Fundamenta priorización de plásticos y "
            "segregación.",
            "Contextual diagnosis (not a mandate). Grounds plastic and segregation prioritization.",
        ),
        "comments": bi(
            "Evidencia de problema plástico que motiva la estrategia.",
            "Evidence of plastic problem motivating the strategy.",
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
        "Signed: 15 July 2021 | Published: 17 July 2021 (El Peruano)\n"
        "Not amended\n"
        f"Source: {POLICY_URL}"
    )

    doc.add_heading("English Translation of Key Provisions", 1)
    doc.add_paragraph(
        "English translation of principal provisions of Ministerial Resolution No. 122-2021-MINAM "
        "approving the Peru Limpio education and communication strategy."
    )

    sections = [
        (
            "RM Art. 1 — Approval",
            "Apruébase el Documento 'Perú Limpio: Estrategia de educación y comunicación...' "
            "como Anexo integrante.",
            "The document 'Peru Limpio: Education and Communication Strategy...' is approved "
            "as integral Annex.",
        ),
        (
            "General Objective",
            "Fomentar y fortalecer buenas prácticas ambientales para el consumo responsable, "
            "minimización, segregación y valorización de residuos sólidos hacia economía circular.",
            "Foster and strengthen good environmental practices for responsible consumption, "
            "minimization, segregation and valorization of solid waste toward circular economy.",
        ),
        (
            "Priority theme — Responsible consumption",
            "Prevenir la generación de residuos: no usando plásticos de un solo uso, usando "
            "productos reutilizables.",
            "Prevent waste generation: not using single-use plastics, using reusable products.",
        ),
        (
            "OE1 — Public sector",
            "Asistencia técnica para cumplir normativa de residuos, ecoeficiencia y reducción "
            "de plásticos de un solo uso.",
            "Technical assistance to comply with waste regulations, eco-efficiency and single-use "
            "plastic reduction.",
        ),
        (
            "OE2 — Local governments",
            "Institucionalizar educación ambiental vía programas EDUCCA: 70 municipalidades en "
            "2021 y 130 en 2025.",
            "Institutionalize environmental education via EDUCCA programs: 70 municipalities in "
            "2021 and 130 in 2025.",
        ),
        (
            "OE4 — Citizens and campaigns",
            "Realización de campañas educativas para consumo responsable y manejo de residuos "
            "(incl. #MenosPlásticoMásVida).",
            "Conduct educational campaigns for responsible consumption and waste management "
            "(incl. #MenosPlasticoMasVida).",
        ),
    ]

    for heading, es, en in sections:
        doc.add_heading(heading, 2)
        add_article(doc, es, en)

    doc.add_paragraph()
    note = doc.add_paragraph(
        "Note: Peru Limpio implements PNCP Measure 9.2 on integrated solid-waste education. "
        "It links to Law 30884, DL 1278, PLANRES and DS 006-2019-MINAM. The #MenosPlasticoMasVida "
        "campaign is a documented component though not spelled out in the annex PDF text."
    )
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.save(path)


def main() -> None:
    excel = OUTPUT_DIR / "Peru_RM_122_2021_MINAM_4P_Index_Coding.xlsx"
    word = OUTPUT_DIR / "Peru_RM_122_2021_MINAM_English_Translation.docx"
    build_excel(excel)
    build_word(word)
    print(f"Created: {excel}")
    print(f"Created: {word}")
    print(f"Instruments: {len(INSTRUMENTS)}")
    print()
    print("Download links:")
    print(
        f"  Excel: {REPO_RAW}/output/peru-rm-122-2021-minam/"
        "Peru_RM_122_2021_MINAM_4P_Index_Coding.xlsx"
    )
    print(
        f"  Word:  {REPO_RAW}/output/peru-rm-122-2021-minam/"
        "Peru_RM_122_2021_MINAM_English_Translation.docx"
    )


if __name__ == "__main__":
    main()
