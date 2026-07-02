#!/usr/bin/env python3
"""Generate 4P Index Excel coding table and English translation Word doc for Peru Law 29419."""

from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = Path("/workspace/output/peru-law-29419")
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

POLICY_NAME_ES = "Ley que regula la actividad de los recicladores"
POLICY_NAME_EN = "Law that Regulates the Activity of Waste Pickers (Recyclers)"
POLICY_NUMBER = "Law No. 29419"
POLICY_URL = "https://www.minam.gob.pe/wp-content/uploads/2017/04/Ley-N%C2%B0-29419.pdf"
COUNTRY = "Peru"

# Policy-level fields (A-O)
POLICY = {
    "policy_name_es": POLICY_NAME_ES,
    "policy_name_en": POLICY_NAME_EN,
    "policy_number": POLICY_NUMBER,
    "policy_url": POLICY_URL,
    "country": COUNTRY,
    "policy_year": 2009,
    "policy_objective": (
        "Establish a normative framework for regulating recycling workers' activities, "
        "promoting their protection, training, formalization, and association, and "
        "contributing to ecologically efficient solid waste management in Peru. The law "
        "integrates informal recyclers into municipal solid waste management systems, "
        "including collection, segregation, and commercialization of non-hazardous solid "
        "waste (which commonly includes plastic waste streams)."
    ),
    "policy_target": 0,
    "policy_target_text": "",
    "policy_type": 1,
    "policy_type_justification": (
        "Legislation enacted by the Congress of the Republic of Peru (Ley Nº 29419), "
        "approved on 18 September 2009 and promulgated by the President on 6 October 2009. "
        "This is parliamentary legislation, not an executive decree or regulation."
    ),
    "policy_integration": 1,
    "policy_sectors_list": (
        "waste management, recycling, municipalities, consumption, industry, education, "
        "health, retail"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": "consumption, recycling, disposal, environmental leakage",
    "policy_budget": 1,
    "policy_budget_text": (
        "Art. 11: 'El Fondo Nacional del Ambiente (FONAM), en coordinación con las "
        "instituciones privadas, crea un fondo especial orientado a facilitar el acceso "
        "al crédito a los recicladores con fines vinculados a su actividad, formalización "
        "y asociación.'"
    ),
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 2.2: The State recognizes recycler activity and promotes formalization and "
            "integration of recyclers into solid waste management systems in all cities, "
            "through the General Directorate of Environmental Health (Digesa) of the Ministry "
            "of Health and provincial municipalities."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Responsible authorities designated: Digesa (Ministry of Health) and provincial "
            "municipalities explicitly named. Monitoring: integration into municipal SWM "
            "systems implies ongoing coordination. No explicit fines/penalties or "
            "unconditional binding obligations on private actors in this article."
        ),
        "comments": (
            "Governance instrument integrating informal recyclers (including plastic waste "
            "pickers) into municipal waste systems. Does not explicitly mention plastics but "
            "applies to all non-hazardous solid waste streams."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 4: Defines institutional actors for selective collection, segregation, and "
            "commercialization of non-hazardous solid waste: Ministry of Environment "
            "(national environmental policy), Ministry of Health (sanitary waste policy), "
            "local/provincial/district governments, recycler associations, EPS-RS and EC-RS."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Art. 4 explicitly designates institutional actors across national and local "
            "levels. No enforcement, monitoring mechanisms, or unconditional obligations "
            "specified in this article."
        ),
        "comments": (
            "Institutional coordination framework. Art. 4 is definitional of roles rather "
            "than imposing new duties, but establishes the multi-level governance structure "
            "for recycler-related waste management."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5.1: Local governments regulate recycler activity as lead authorities and "
            "must incorporate recyclers into local solid waste management systems. Local "
            "governments establish promotion norms for non-hazardous solid waste recyclers "
            "in coordination with registered recycler associations."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Local governments designated as lead regulators. Art. 5.1 uses mandatory "
            "orientation to incorporate recyclers into local systems. No explicit "
            "enforcement penalties or monitoring/inspection requirements in this provision."
        ),
        "comments": (
            "Multi-level governance delegation to municipalities. Coordination with recycler "
            "associations is required."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 5.2: Solid waste management programs and projects implemented by local "
            "governments must include recycler activity ('deben incluir la actividad de "
            "los recicladores')."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Local governments are the responsible authority. Mandatory language 'deben "
            "incluir' (must include). No explicit fines, inspections, or audit mechanisms "
            "in this article."
        ),
        "comments": (
            "Regulatory obligation on municipalities. Applies to all municipal waste "
            "programs, including those handling plastic recyclables."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 5.3: Local governments maintain a registration of recycler associations "
            "operating in their jurisdiction for authorization, certification, and access to "
            "benefits established in their favor."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.75,
        "instrument_implementation_text": (
            "Local governments must maintain registration ('mantienen un registro'). "
            "Registration serves authorization/certification function (monitoring mechanism). "
            "Responsible authority: local governments. No explicit penalties stated in Art. 5.3."
        ),
        "comments": (
            "Formalization/registry instrument. Registration is a monitoring and "
            "authorization mechanism for recycler associations."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 5.4: Formalized recyclers registered with local governments have the right "
            "to exercise their activity within the framework of this Law, its regulation, "
            "solid waste legislation, and municipal norms."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Establishes legal right for formalized recyclers. Local governments administer "
            "registration (Art. 5.3). No explicit enforcement or monitoring provisions in "
            "Art. 5.4 itself."
        ),
        "comments": (
            "Regulatory rights provision protecting formalized recyclers' ability to operate "
            "in the waste/recycling sector."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 6: Regional and local governments promote formation of recycler associations "
            "and small/micro EPS-RS and EC-RS enterprises specialized in collection for "
            "recycling and commercialization; they issue provisions facilitating incorporation "
            "of independent recyclers. Regulation sets promotional regime requirements."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Regional and local governments designated. Uses 'promueven' (promote) rather "
            "than prohibit/require private action. Requirements for promotional regime deferred "
            "to regulation (DS 005-2010-MINAM, now in force)."
        ),
        "comments": (
            "Governance/promotional instrument. Regulation DS 005-2010 operationalizes "
            "requirements referenced in Art. 6."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Art. 7: Local governments implement source segregation incentive programs, "
            "which may include compensation to taxpayers through tariff reductions, delivery "
            "of lower-cost or free goods/services, or environmental certification programs "
            "for enterprises and institutions."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Local governments must implement programs ('implementan programas'). Responsible "
            "authority designated. Programs are discretionary in design ('pueden incluir'). "
            "No explicit enforcement penalties or monitoring requirements."
        ),
        "comments": (
            "Economic incentive instrument for source segregation, which commonly includes "
            "segregation of plastic packaging and containers. Incentive design is flexible."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Art. 8: Local governments promote implementation of treatment plants within "
            "sanitary landfills where organized recyclers can segregate reusable waste for "
            "commercialization."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Local governments designated ('promueven la implementación'). Promotional "
            "language rather than mandatory construction requirement. No enforcement or "
            "monitoring mechanisms specified."
        ),
        "comments": (
            "Infrastructure promotion instrument for landfill-based segregation facilities. "
            "Reusable waste includes plastics recovered at landfills."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 9.1: Ministry of Environment and local governments, coordinating with "
            "Education and Health Ministries, regional governments, universities, specialized "
            "educational institutions, and NGOs, promote training programs for recyclers."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Multiple authorities designated (MINAM, local governments, coordinating "
            "entities). 'Promueven' (promote) — no mandatory training quotas, penalties, or "
            "inspection requirements."
        ),
        "comments": "Voluntary/information instrument for recycler capacity building.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 9.2: SENATI (National Industrial Training Service) establishes a national "
            "educational and training program for recyclers to make solid waste handling "
            "ecologically efficient and technical. Program content coordinated with Ministries "
            "of Environment and Health."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "SENATI explicitly required to establish program ('establece un programa'). "
            "Content coordination with MINAM and Health Ministry (monitoring/coordination). "
            "No penalties or unconditional obligations on recyclers to participate."
        ),
        "comments": (
            "Information/training instrument with binding establishment duty on SENATI. "
            "Borderline between 0.40 and 0.20; scored 0.40 as training/awareness program."
        ),
    },
    {
        "instrument_type": 1.0,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 10.1: Authorizations or licenses granted by local governments to recyclers "
            "must comply with legal norms protecting minors, pregnant women, persons with "
            "disabilities, and elderly persons, under social responsibility and at social cost."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            "Local governments grant authorizations/licenses. Mandatory language 'deben "
            "cumplir' (must comply). No explicit fines or inspection mechanisms in this "
            "article, though general labor/child protection laws may apply."
        ),
        "comments": (
            "Regulatory social protection standard for recycler licensing. Applies to all "
            "recycler authorizations regardless of material type."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 10.2: Ministry of Health progressively implements vaccination and "
            "occupational health programs for recyclers, in coordination with local governments."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Ministry of Health designated. 'Implementa progresivamente' (progressively "
            "implements) — phased rather than immediate unconditional obligation. No "
            "enforcement penalties specified."
        ),
        "comments": (
            "Governance/health protection instrument. Progressive implementation weakens "
            "unconditional score."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Art. 11: FONAM (National Environment Fund), in coordination with private "
            "institutions, creates a special fund to facilitate credit access for recyclers "
            "for purposes linked to their activity, formalization, and association."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "FONAM designated to create special fund ('crea un fondo especial'). No explicit "
            "enforcement, monitoring, or ring-fenced budget amount specified in the law text."
        ),
        "comments": (
            "Economic instrument — dedicated credit fund. Policy-level budget score of 1 "
            "reflects creation of special fund; implementation detail may be in FONAM "
            "operational rules."
        ),
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Complementary Provision First: National Recycler Day is commemorated on 1 June "
            "each year."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Commemoration established by law. No designated implementing authority, "
            "enforcement, or monitoring specified."
        ),
        "comments": "Awareness/commemoration instrument with limited implementation detail.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Complementary Provision Second: National Recycling Award created under Ministry "
            "of Environment, awarded annually to persons/entities distinguished for recycling "
            "integrating environmental, social, and economic benefits. Local governments may "
            "grant similar awards."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "MINAM designated as responsible for national award. Local governments may "
            "grant similar prizes. No enforcement or monitoring mechanisms."
        ),
        "comments": "Voluntary recognition/awareness instrument.",
    },
    {
        "instrument_type": 0.40,
        "instrument_lifecycle_stage": "Consumption",
        "instrument_description": (
            "Complementary Provision Third: Ministry of Environment, coordinating with "
            "Ministry of Education and competent entities, promotes public education and "
            "communication programs on benefits of source segregation and recycling, "
            "highlighting recyclers' role."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "MINAM and Education Ministry designated. 'Promueve programas' — promotional "
            "language. No enforcement or monitoring requirements."
        ),
        "comments": (
            "Public awareness instrument targeting source segregation and recycling, including "
            "plastic waste streams."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Final Complementary Provision (Única): Executive Power must approve, within 120 "
            "days of the law's entry into force, the regulation of this Law by supreme decree, "
            "countersigned by Ministers of Environment and Health."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.25,
        "instrument_implementation_text": (
            "Executive Power designated. Enabling power exercised: DS Nº 005-2010-MINAM "
            "approved the regulation on 2 June 2010. Original 'debe aprobar' was mandatory "
            "on the Executive."
        ),
        "comments": (
            "Enabling power scored in_force = 1 because subordinate regulation DS "
            "005-2010-MINAM was adopted and is operative."
        ),
    },
]

COLUMNS = [
    ("A", "policy_name_es"),
    ("B", "policy_name_en"),
    ("C", "policy_number"),
    ("D", "country"),
    ("E", "policy_url"),
    ("F", "policy_year"),
    ("G", "policy_objective"),
    ("H", "policy_target"),
    ("I", "policy_target_text"),
    ("J", "policy_type"),
    ("K", "policy_type_justification"),
    ("L", "policy_integration"),
    ("M", "policy_sectors_list"),
    ("N", "policy_circularity"),
    ("O", "policy_lifecycle_phases_list"),
    ("P", "policy_budget"),
    ("Q", "policy_budget_text"),
    ("R", "policy_score"),
    ("S", "instrument_type"),
    ("T", "instrument_lifecycle_stage"),
    ("U", "instrument_description"),
    ("V", "instrument_in_force"),
    ("W", "instrument_implementation"),
    ("X", "instrument_implementation_text"),
    ("Y", "instrument_score"),
    ("Z", "comments"),
]


def policy_score_row(instrument_type: float) -> str:
  avg = (POLICY["policy_type"] + POLICY["policy_integration"] + POLICY["policy_circularity"]
         + POLICY["policy_budget"] + instrument_type) / 5
  return round(avg, 3)


def instrument_score_row(instrument_type: float, implementation: float) -> str:
  return round((instrument_type + implementation) / 2, 3)


def build_excel(path: Path) -> None:
  wb = Workbook()
  ws = wb.active
  ws.title = "4P Index Coding"

  headers = [f"{col} — {name}" for col, name in COLUMNS]
  ws.append(headers)

  header_fill = PatternFill("solid", fgColor="1F4E79")
  header_font = Font(color="FFFFFF", bold=True)
  for cell in ws[1]:
    cell.fill = header_fill
    cell.font = header_font
    cell.alignment = Alignment(wrap_text=True, vertical="top")

  for inst in INSTRUMENTS:
    row = [
      POLICY["policy_name_es"],
      POLICY["policy_name_en"],
      POLICY["policy_number"],
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
    ws.append(row)

  for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
    for cell in row:
      cell.alignment = Alignment(wrap_text=True, vertical="top")

  widths = {
    "A": 28, "B": 28, "C": 14, "D": 10, "E": 36, "F": 8, "G": 42, "H": 8, "I": 18,
    "J": 8, "K": 36, "L": 10, "M": 36, "N": 10, "O": 28, "P": 10, "Q": 36, "R": 10,
    "S": 10, "T": 18, "U": 48, "V": 10, "W": 12, "X": 42, "Y": 10, "Z": 36,
  }
  for col, width in widths.items():
    ws.column_dimensions[col].width = width

  ws.freeze_panes = "A2"
  wb.save(path)


def add_heading(doc: Document, text: str, level: int = 1) -> None:
  doc.add_heading(text, level=level)


def add_bilingual_article(doc: Document, article_es: str, article_en: str) -> None:
  p = doc.add_paragraph()
  run = p.add_run(article_es)
  run.bold = True
  run.font.size = Pt(11)
  p = doc.add_paragraph(article_en)
  p.paragraph_format.space_after = Pt(10)


def build_word(path: Path) -> None:
  doc = Document()
  section = doc.sections[0]
  section.top_margin = Inches(1)
  section.bottom_margin = Inches(1)
  section.left_margin = Inches(1)
  section.right_margin = Inches(1)

  title = doc.add_paragraph()
  title.alignment = WD_ALIGN_PARAGRAPH.CENTER
  run = title.add_run(f"{POLICY_NAME_EN}\n{POLICY_NAME_ES}")
  run.bold = True
  run.font.size = Pt(16)

  subtitle = doc.add_paragraph()
  subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
  sub_run = subtitle.add_run(f"{POLICY_NUMBER} — Republic of Peru")
  sub_run.font.size = Pt(12)

  meta = doc.add_paragraph()
  meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
  meta.add_run(
    "Adopted by Congress: 18 September 2009 | Promulgated: 6 October 2009\n"
    f"Source: {POLICY_URL}"
  )

  doc.add_paragraph()
  add_heading(doc, "English Translation of Key Provisions", 1)
  doc.add_paragraph(
    "The following is an English translation of the principal provisions of Peru's Law No. 29419. "
    "Article headings and numbering follow the official Spanish text."
  )

  add_heading(doc, "Preamble", 2)
  doc.add_paragraph(
    "THE CONGRESS OF THE REPUBLIC; Has enacted the following Law:"
  )

  add_heading(doc, "Article 1 — Purpose of the Law", 2)
  add_bilingual_article(
    doc,
    "El objeto de la presente Ley es establecer el marco normativo para la regulación de las actividades de los trabajadores del reciclaje, orientada a la protección, capacitación y promoción del desarrollo social y laboral, promoviendo su formalización, asociación y contribuyendo a la mejora en el manejo ecológicamente eficiente de los residuos sólidos en el país, en el marco de los objetivos y principios de la Ley núm. 27314, Ley General de Residuos Sólidos, y la Ley núm. 28611, Ley General del Ambiente.",
    "The purpose of this Law is to establish the normative framework for regulating the activities of recycling workers, aimed at their protection, training, and promotion of social and labor development, promoting their formalization and association, and contributing to the improvement of ecologically efficient solid waste management in the country, within the framework of the objectives and principles of Law No. 27314, General Solid Waste Law, and Law No. 28611, General Environment Law.",
  )

  add_heading(doc, "Article 2 — Scope of Application", 2)
  add_bilingual_article(
    doc,
    "2.1 Para efectos de la aplicación de la presente Ley, se considera recicladores a las personas que, de forma dependiente o independiente, se dedican a las actividades de recolección selectiva para el reciclaje, segregación y comercialización en pequeña escala de residuos sólidos no peligrosos, de acuerdo con lo dispuesto por la Ley núm. 27314, Ley General de Residuos Sólidos.",
    "2.1 For purposes of applying this Law, recyclers are persons who, whether dependently or independently employed, engage in selective collection for recycling, segregation, and small-scale commercialization of non-hazardous solid waste, in accordance with Law No. 27314, General Solid Waste Law.",
  )
  add_bilingual_article(
    doc,
    "2.2 El Estado reconoce la actividad de los recicladores, promueve su formalización e integración a los sistemas de gestión de residuos sólidos de todas las ciudades del país a través de la Dirección General de Salud Ambiental (Digesa), del Ministerio de Salud y de las municipalidades provinciales.",
    "2.2 The State recognizes the activity of recyclers, promotes their formalization and integration into the solid waste management systems of all cities in the country through the General Directorate of Environmental Health (Digesa) of the Ministry of Health and provincial municipalities.",
  )

  add_heading(doc, "Article 3 — Definitions", 2)
  definitions = [
    (
      "a) Reciclaje",
      "Recycling: Process by which waste, inputs, or final products are incorporated into transformation and production processes designed especially to eliminate or minimize their polluting effects and generate economic benefits.",
    ),
    (
      "b) Recolección selectiva para el reciclaje",
      "Selective collection for recycling: Action of collecting waste segregated at source and transferring it through an appropriate means of transport for subsequent conditioning and commercialization.",
    ),
    (
      "c) Segregación",
      "Segregation: Action of grouping specific physical components or elements of solid waste for special handling.",
    ),
    (
      "d) Residuos sólidos no peligrosos",
      "Non-hazardous solid waste: Waste not defined as hazardous under Legislative Resolution No. 26234 (Basel Convention) and included in Annex 5 of Supreme Decree No. 057-2004-PCM, Regulation of Law No. 27314.",
    ),
    (
      "e) Reciclador independiente",
      "Independent recycler: Person who formally carries out recycling activities, including selective collection and commercialization, and who has no employment relationship with solid waste service providers, solid waste trading companies, or solid waste generating companies.",
    ),
    (
      "f) Empresa Prestadora de Servicios de Residuos Sólidos (EPS-RS)",
      "Solid Waste Service Provider Company (EPS-RS): Legal entity providing solid waste services through one or more of the following activities: cleaning of roads and public spaces, collection and transport, transfer, treatment, or final disposal of solid waste.",
    ),
    (
      "g) Empresa Comercializadora de Residuos Sólidos (EC-RS)",
      "Solid Waste Trading Company (EC-RS): Legal entity engaged in commercialization of solid waste for reuse.",
    ),
  ]
  for term_es, term_en in definitions:
    p = doc.add_paragraph()
    p.add_run(f"{term_es}: ").bold = True
    p.add_run(term_en)

  add_heading(doc, "Article 4 — Institutional Actors", 2)
  add_bilingual_article(
    doc,
    "Son actores institucionales vinculados a las actividades de recolección selectiva, segregación y comercialización de residuos sólidos no peligrosos los siguientes: a) El Ministerio del Ambiente; b) El Ministerio de Salud; c) Los gobiernos locales, provinciales y distritales; d) Las asociaciones de recicladores; e) Las EPS-RS y EC-RS.",
    "The following are institutional actors linked to selective collection, segregation, and commercialization of non-hazardous solid waste: (a) the Ministry of Environment, as lead entity for national environmental policy; (b) the Ministry of Health, as lead entity for sanitary solid waste management policy; (c) local, provincial, and district governments, responsible for ecologically efficient waste management policies in their jurisdictions; (d) recycler associations; and (e) EPS-RS and EC-RS private economic units.",
  )

  add_heading(doc, "Article 5 — Local Regulation", 2)
  articles_5 = [
    (
      "5.1 La actividad de los recicladores es regulada por los gobiernos locales como entes rectores, en el marco de sus atribuciones. El régimen de regulación local se orienta a incorporar a los recicladores como parte del sistema local de gestión de residuos sólidos.",
      "5.1 Recycler activity is regulated by local governments as lead authorities within their powers. The local regulatory regime is oriented toward incorporating recyclers as part of the local solid waste management system. Local governments establish promotion norms for non-hazardous solid waste recyclers in coordination with registered recycler associations in their jurisdiction.",
    ),
    (
      "5.2 Los programas y proyectos de gestión y manejo de residuos sólidos implementados por los gobiernos locales deben incluir la actividad de los recicladores.",
      "5.2 Solid waste management programs and projects implemented by local governments must include recycler activity.",
    ),
    (
      "5.3 Los gobiernos locales mantienen un registro de inscripción de las asociaciones de recicladores, cuyos miembros operen en su jurisdicción para el otorgamiento de la autorización y certificación correspondiente.",
      "5.3 Local governments maintain a registration of recycler associations whose members operate in their jurisdiction for granting corresponding authorization and certification, which also serves access to benefits established in their favor.",
    ),
    (
      "5.4 Los recicladores formalizados a través del registro en los gobiernos locales tienen derecho a ejercer su actividad dentro del marco establecido por la presente Ley y su reglamento, la legislación de residuos sólidos y las normas municipales.",
      "5.4 Recyclers formalized through registration with local governments have the right to exercise their activity within the framework established by this Law and its regulation, solid waste legislation, and municipal norms.",
    ),
  ]
  for es, en in articles_5:
    add_bilingual_article(doc, es, en)

  add_heading(doc, "Article 6 — Formation of EPS-RS and EC-RS", 2)
  add_bilingual_article(
    doc,
    "Los gobiernos regionales y locales, en el marco de sus atribuciones legales, promueven la formación de asociaciones de recicladores y de pequeñas y microempresas EPS-RS y EC-RS, especializadas en la recolección para el reciclaje y la comercialización de residuos sólidos.",
    "Regional and local governments, within their legal powers, promote the formation of recycler associations and small and micro EPS-RS and EC-RS enterprises specialized in collection for recycling and commercialization of solid waste; they also issue provisions facilitating incorporation of independent recyclers into existing enterprises. The regulation of this Law establishes requirements for the promotional regime.",
  )

  add_heading(doc, "Article 7 — Incentives for Source Segregation", 2)
  add_bilingual_article(
    doc,
    "Los gobiernos locales implementan programas de incentivos a la segregación en la fuente, los cuales pueden incluir compensación a los contribuyentes a través de la reducción del pago de tarifas o la entrega de bienes o servicios a menos costo o de forma gratuita, o como parte de programas de certificación ambiental de empresas o instituciones en general.",
    "Local governments implement source segregation incentive programs, which may include compensation to taxpayers through reduction of fee payments, delivery of goods or services at lower cost or free of charge, or as part of environmental certification programs for enterprises or institutions in general.",
  )

  add_heading(doc, "Article 8 — Recycling at Sanitary Landfills", 2)
  add_bilingual_article(
    doc,
    "Los gobiernos locales promueven la implementación de plantas de tratamiento dentro de los rellenos sanitarios en donde los recicladores organizados puedan segregar los residuos reutilizables para su comercialización.",
    "Local governments promote the implementation of treatment plants within sanitary landfills where organized recyclers can segregate reusable waste for commercialization.",
  )

  add_heading(doc, "Article 9 — Training Programs for Recyclers", 2)
  add_bilingual_article(
    doc,
    "9.1 El Ministerio del Ambiente y los gobiernos locales, en coordinación con los Ministerios de Educación y de Salud, los gobiernos regionales, las universidades, las instituciones educativas especializadas y las organizaciones no gubernamentales, promueven el desarrollo de programas de capacitación a los recicladores.",
    "9.1 The Ministry of Environment and local governments, in coordination with the Ministries of Education and Health, regional governments, universities, specialized educational institutions, and non-governmental organizations, promote the development of training programs for recyclers.",
  )
  add_bilingual_article(
    doc,
    "9.2 El Servicio Nacional de Adiestramiento en Trabajo Industrial (Senati) establece un programa educativo y de capacitación dirigido a los recicladores en todo el país, con el objetivo de hacer ecológicamente eficiente y técnico el manejo de los residuos sólidos.",
    "9.2 The National Industrial Training Service (SENATI) establishes an educational and training program directed at recyclers throughout the country, with the objective of making solid waste handling ecologically efficient and technical. Similar programs may be developed by other educational institutions. In all cases, program content is coordinated with the Ministries of Environment and Health.",
  )

  add_heading(doc, "Article 10 — Protection of Vulnerable Groups", 2)
  add_bilingual_article(
    doc,
    "10.1 Las autorizaciones o licencias concedidas por los gobiernos locales a los recicladores deben cumplir las normas legales de protección al menor de edad, las madres gestantes, las personas con discapacidad y las personas de la tercera edad, bajo responsabilidad y a costo social.",
    "10.1 Authorizations or licenses granted by local governments to recyclers must comply with legal norms protecting minors, pregnant women, persons with disabilities, and elderly persons, under social responsibility and at social cost.",
  )
  add_bilingual_article(
    doc,
    "10.2 El Ministerio de Salud implementa progresivamente programas de vacunación y salud ocupacional para los recicladores, en coordinación con los gobiernos locales.",
    "10.2 The Ministry of Health progressively implements vaccination and occupational health programs for recyclers, in coordination with local governments.",
  )

  add_heading(doc, "Article 11 — Recycling Promotion Fund", 2)
  add_bilingual_article(
    doc,
    "El Fondo Nacional del Ambiente (FONAM), en coordinación con las instituciones privadas, crea un fondo especial orientado a facilitar el acceso al crédito a los recicladores con fines vinculados a su actividad, formalización y asociación.",
    "The National Environment Fund (FONAM), in coordination with private institutions, creates a special fund oriented toward facilitating credit access for recyclers for purposes linked to their activity, formalization, and association.",
  )

  add_heading(doc, "Complementary Provisions", 2)
  comp = [
    (
      "PRIMERA",
      "FIRST: National Recycler Day is commemorated on 1 June each year.",
    ),
    (
      "SEGUNDA",
      "SECOND: The National Recycling Award is created under the Ministry of Environment, awarded annually to natural or legal persons distinguished for recycling that integrates environmental, social, and economic advantages. Local governments grant similar awards in their jurisdictions.",
    ),
    (
      "TERCERA",
      "THIRD: The Ministry of Environment, in coordination with the Ministry of Education and other competent entities, promotes public education and communication programs aimed at showing the social, environmental, and economic benefits of source segregation and recycling, emphasizing recyclers' role in this process.",
    ),
  ]
  for es, en in comp:
    p = doc.add_paragraph()
    p.add_run(f"Complementary Provision {es}: ").bold = True
    p.add_run(en)

  add_heading(doc, "Final Complementary Provision", 2)
  add_bilingual_article(
    doc,
    "ÚNICA.- El Poder Ejecutivo, en un plazo no mayor de ciento veinte (120) días, contado a partir de la vigencia de la presente Ley, aprueba mediante decreto supremo el reglamento de la presente Ley, el cual debe ser refrendado por los Ministros del Ambiente y de Salud.",
    "SOLE: The Executive Power, within no more than one hundred twenty (120) days from the entry into force of this Law, shall approve by supreme decree the regulation of this Law, which must be countersigned by the Ministers of Environment and Health.",
  )

  add_heading(doc, "Note on Implementing Regulation", 2)
  doc.add_paragraph(
    "The regulation of Law No. 29419 was approved by Supreme Decree No. 005-2010-MINAM on 2 June 2010, "
    "within the 120-day deadline established by the Law. The regulation contains detailed provisions on "
    "registration, authorization, integration into municipal waste systems, and operational requirements "
    "for recyclers and recycler associations."
  )

  doc.add_paragraph()
  footer = doc.add_paragraph(
    "Document prepared for the Plastic Pollution Policy Index (4P Index) coding exercise. "
    "English translation of official Spanish legal text."
  )
  footer.alignment = WD_ALIGN_PARAGRAPH.CENTER

  doc.save(path)


def main() -> None:
  excel_path = OUTPUT_DIR / "Peru_Law_29419_4P_Index_Coding.xlsx"
  word_path = OUTPUT_DIR / "Peru_Law_29419_English_Translation.docx"
  build_excel(excel_path)
  build_word(word_path)
  print(f"Created: {excel_path}")
  print(f"Created: {word_path}")
  print(f"Instruments coded: {len(INSTRUMENTS)}")


if __name__ == "__main__":
  main()
