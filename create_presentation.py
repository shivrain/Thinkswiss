#!/usr/bin/env python3
"""Generate policy coding & scoring review presentation."""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

FONT = "Times New Roman"
TITLE_COLOR = RGBColor(0x1F, 0x3A, 0x5F)
HEADER_BG = RGBColor(0x1F, 0x3A, 0x5F)
HEADER_FG = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0x2E, 0x6B, 0x9E)
BODY_COLOR = RGBColor(0x22, 0x22, 0x22)
COMMENT_COLOR = RGBColor(0xC0, 0x00, 0x00)
QUOTE_COLOR = RGBColor(0x44, 0x44, 0x44)
LIGHT_BG = RGBColor(0xFA, 0xFA, 0xFA)
BORDER = RGBColor(0xBB, 0xBB, 0xBB)
DELETE_COLOR = RGBColor(0x8B, 0x00, 0x00)

BASE_FONT = 10
HEADER_FONT = 11
TITLE_FONT = 24


def set_font(run, size=BASE_FONT, bold=False, italic=False, color=BODY_COLOR):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_rich_content(tf, blocks, base_size=BASE_FONT):
    """blocks: list of dicts with keys: text, bold, italic, color, size, newline_before"""
    first = True
    for block in blocks:
        text = block.get("text", "")
        if not text:
            continue
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.space_after = Pt(block.get("space_after", 4))
        p.level = block.get("level", 0)
        r = p.add_run()
        r.text = text
        set_font(
            r,
            size=block.get("size", base_size),
            bold=block.get("bold", False),
            italic=block.get("italic", False),
            color=block.get("color", BODY_COLOR),
        )


def comment_blocks(elaboration, quote=None, proposed=None, question=None):
    blocks = []
    if elaboration:
        blocks.append({"text": elaboration, "color": COMMENT_COLOR, "space_after": 3})
    if proposed:
        blocks.append({"text": f"Proposed change: {proposed}", "color": COMMENT_COLOR, "bold": True, "space_after": 3})
    if quote:
        blocks.append({"text": f'"{quote}"', "color": QUOTE_COLOR, "italic": True, "space_after": 4})
    if question:
        blocks.append({"text": f"Question: {question}", "color": COMMENT_COLOR, "bold": True, "space_after": 6})
    return blocks


def field_blocks(title, comment_data=None):
    blocks = [{"text": title, "bold": True, "color": ACCENT, "space_after": 3}]
    if comment_data:
        if isinstance(comment_data, str):
            blocks.append({"text": comment_data, "color": COMMENT_COLOR, "space_after": 4})
        else:
            blocks.extend(comment_blocks(**comment_data))
    return blocks


def add_title_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.16))
    bar.fill.solid()
    bar.fill.fore_color.rgb = TITLE_COLOR
    bar.line.fill.background()

    box = slide.shapes.add_textbox(Inches(0.6), Inches(2.1), Inches(8.8), Inches(1.0))
    r = box.text_frame.paragraphs[0].add_run()
    r.text = "Policy Coding & Scoring Review"
    set_font(r, size=30, bold=True, color=TITLE_COLOR)

    sub = slide.shapes.add_textbox(Inches(0.6), Inches(3.1), Inches(8.8), Inches(0.6))
    sr = sub.text_frame.paragraphs[0].add_run()
    sr.text = "Country Review of Assumptions, Inclusions, Policy & Instrument Coding"
    set_font(sr, size=15, italic=True, color=ACCENT)

    date_box = slide.shapes.add_textbox(Inches(0.6), Inches(3.8), Inches(8.8), Inches(0.35))
    dr = date_box.text_frame.paragraphs[0].add_run()
    dr.text = "Review Meeting — July 2026"
    set_font(dr, size=12, color=RGBColor(0x66, 0x66, 0x66))

    note = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(8.8), Inches(0.8))
    ntf = note.text_frame
    ntf.word_wrap = True
    np = ntf.paragraphs[0]
    np.alignment = PP_ALIGN.CENTER
    nr = np.add_run()
    nr.text = (
        "Note: Norway has no coding or scoring doubts to be discussed in this review. "
        "Countries under review: Nepal, Mozambique, Peru, and São Tomé and Príncipe (STP)."
    )
    set_font(nr, size=11, italic=True, color=COMMENT_COLOR)


def add_section_header(slide, left, top, width, text):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.3))
    shape.fill.solid()
    shape.fill.fore_color.rgb = HEADER_BG
    shape.line.fill.background()
    tf = shape.text_frame
    tf.margin_left = Pt(6)
    tf.margin_top = Pt(1)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    r = tf.paragraphs[0].add_run()
    r.text = text
    set_font(r, size=HEADER_FONT, bold=True, color=HEADER_FG)


def add_content_cell(slide, left, top, width, height, blocks, base_size=BASE_FONT):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = BORDER
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(6)
    tf.margin_right = Pt(6)
    tf.margin_top = Pt(5)
    tf.margin_bottom = Pt(4)
    tf.vertical_anchor = MSO_ANCHOR.TOP
    add_rich_content(tf, blocks, base_size)


def get_active_columns(entries):
    cols = []
    keys = [
        ("assumptions", "Assumptions"),
        ("policy_inclusion", "Policy Inclusion / Exclusions"),
        ("policy_level", "Policy-Level Fields"),
        ("instrument_level", "Instrument-Level Fields"),
    ]
    for key, label in keys:
        if any(e.get(key) for e in entries):
            cols.append((key, label))
    return cols


def build_cell_blocks(entry, col_key):
    if col_key == "assumptions":
        return [{"text": entry.get("assumptions", ""), "space_after": 2}] if entry.get("assumptions") else []

    if col_key == "policy_inclusion":
        blocks = []
        if entry.get("policy_inclusion"):
            blocks.extend(field_blocks(entry["policy_inclusion"], entry.get("policy_inclusion_comment")))
        return blocks

    if col_key == "policy_level":
        blocks = []
        if entry.get("policy_level"):
            blocks.extend(field_blocks(entry["policy_level"], entry.get("policy_level_comment")))
        return blocks

    if col_key == "instrument_level":
        blocks = []
        if entry.get("instrument_level"):
            blocks.extend(field_blocks(entry["instrument_level"], entry.get("instrument_comment")))
        return blocks

    return []


def add_delete_list(slide, left, top, width, items):
    if not items:
        return top
    h = Inches(0.28 + 0.22 * len(items))
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, h)
    box.fill.solid()
    box.fill.fore_color.rgb = RGBColor(0xFD, 0xF2, 0xF2)
    box.line.color.rgb = DELETE_COLOR
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(8)
    tf.margin_right = Pt(8)
    tf.margin_top = Pt(5)
    tf.margin_bottom = Pt(4)
    add_rich_content(
        tf,
        [{"text": "Policies proposed for removal:", "bold": True, "color": DELETE_COLOR, "space_after": 4}]
        + [{"text": f"• {item}", "color": BODY_COLOR, "space_after": 2} for item in items],
        base_size=BASE_FONT,
    )
    return top + h


def add_country_slide(prs, country, entries, delete_list=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    left_margin = Inches(0.15)
    content_width = Inches(9.7)

    title_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.48))
    title_bar.fill.solid()
    title_bar.fill.fore_color.rgb = TITLE_COLOR
    title_bar.line.fill.background()
    ttf = title_bar.text_frame
    ttf.margin_left = Pt(12)
    ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tr = ttf.paragraphs[0].add_run()
    tr.text = country
    set_font(tr, size=TITLE_FONT, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    active_cols = get_active_columns(entries)
    if not active_cols:
        return

    col_top = Inches(0.52)
    col_count = len(active_cols)
    col_gap = Inches(0.04)
    col_w = (content_width - col_gap * (col_count - 1)) / col_count

    for i, (_, label) in enumerate(active_cols):
        add_section_header(slide, left_margin + (col_w + col_gap) * i, col_top, col_w, label)

    delete_h = Inches(0.35 + 0.22 * len(delete_list or [])) if delete_list else Inches(0)
    entry_top = Inches(0.86)
    footer_space = Inches(0.18)
    available_h = Inches(7.5) - entry_top - delete_h - footer_space
    n = len(entries)
    row_gap = Inches(0.04)
    row_h = (available_h - row_gap * (n - 1)) / n if n else available_h

    for ridx, entry in enumerate(entries):
        top = entry_top + (row_h + row_gap) * ridx
        for cidx, (col_key, _) in enumerate(active_cols):
            left = left_margin + (col_w + col_gap) * cidx
            blocks = build_cell_blocks(entry, col_key)
            if not blocks:
                blocks = [{"text": "—", "color": RGBColor(0xAA, 0xAA, 0xAA)}]
            add_content_cell(slide, left, top, col_w, row_h, blocks)

    if delete_list:
        add_delete_list(slide, left_margin, entry_top + (row_h + row_gap) * n - row_gap + Inches(0.04), content_width, delete_list)

    footer = slide.shapes.add_textbox(left_margin, Inches(7.28), content_width, Inches(0.15))
    fr = footer.text_frame.paragraphs[0].add_run()
    fr.text = "Policy Coding & Scoring Review  |  July 2026"
    set_font(fr, size=8, italic=True, color=RGBColor(0x99, 0x99, 0x99))


# ── Country data ──────────────────────────────────────────────────────────────

NEPAL_ENTRIES = [
    {
        "assumptions": "Mentions plastics explicitly",
        "policy_inclusion": "The Excise Duty Act, 2002",
        "policy_inclusion_comment": {
            "elaboration": "This Act was provisionally included but falls outside the inclusion criteria. It should be removed from the dataset.",
        },
        "policy_level": "Directive on the Ban of Single-Use Plastics in the Sagarmatha (Everest) Region",
        "policy_level_comment": {
            "elaboration": "Valid plastic-specific directive. Sector coding should reflect its multi-sector reach beyond waste management alone.",
            "quote": "Policy sectors: tourism, waste management, conservation, water, retail, packaging.",
        },
        "instrument_level": "Environment-friendly Local Governance Framework, 2013",
        "instrument_comment": {
            "elaboration": "Documents a comprehensive multi-level plastic bag restriction system across household, Tole, municipality, VDC, and district levels.",
            "quote": (
                "HOUSEHOLD (Basic): 'Among the inorganic (non-decomposing) wastes, plastics collected separately "
                "and sold or gathered in a fixed or specified location' (p.11, p.19). "
                "HOUSEHOLD (Advanced): 'Plastic Bag Regulation: Instead of all sorts of plastic bags; jute, cloth or paper bags used "
                "and written commitment towards it submitted to the Tole Development Organization.' (p.12, p.20). "
                "MUNICIPALITY (Basic): 'Ban on plastic use: Through the decision of municipal counsel, use of plastic restricted "
                "as per the national standard approved by the Government of Nepal.' (p.15). "
                "DISTRICT (Basic): 'Declaration of plastic bag free district… District level mechanism established to control "
                "the import of prohibited plastic bags.' (p.21)."
            ),
            "question": "How will we analyze the qualitative data from multi-level governance frameworks like this?",
        },
    },
    {
        "assumptions": "Anything remotely national",
        "policy_inclusion": "Customs Act (Law) 2007",
        "policy_level": "Action Plan for Ban on Plastic Bags, 2022",
        "policy_level_comment": {
            "elaboration": (
                "Coding relied on secondary sources because the primary Nepali source document has severely corrupted OCR. "
                "English-language reporting and SWITCH-Asia (2025) were used to validate content."
            ),
            "quote": "Sources: partial Nepali text extraction; The Himalayan Times; myRepublica; SWITCH-Asia Nepal plastic policy overview (2025).",
        },
        "instrument_level": "Local Government Operation Act",
        "instrument_comment": {
            "elaboration": "Assigns municipal and ward-level waste management functions relevant to plastic waste at the local level.",
            "quote": (
                "SCHEDULE — MUNICIPAL FUNCTIONS (Section झ): "
                "'(6) Raising sanitation awareness and management of health-related waste… "
                "(7) Collection, reuse, processing, disposal of health waste… "
                "(10) Coordination… for management of waste from sanitation and health sector.' "
                "SCHEDULE (Section ञ): '(15) Sanitation and waste management at the local level.' "
                "WARD FUNCTIONS: '(11) Collection and management of waste discharged from homes…' "
                "MUNICIPALITY CLASSIFICATION (Section 5): 'solid waste processing and management system' as a classification criterion."
            ),
        },
    },
    {
        "assumptions": "Frameworks are also included",
        "policy_inclusion": "National Urban Water Supply and Sanitation Sector Policy, 2009",
        "policy_level": "The Sixteenth Plan (FY 2024/25–2028/29)",
        "policy_level_comment": {
            "elaboration": (
                "Budget is allocated to waste management and recycling but is not ring-fenced specifically for plastics. "
                "Current budget coding should reflect general sector investment rather than plastic-specific funding."
            ),
            "quote": (
                "Water supply, sewerage, waste management and recycling — total plan investment NPR 3,991 million "
                "(Public: NPR 2,091 million / 52.4%; Private: NPR 1,852 million / 46.4%; Cooperative: NPR 48 million / 1.2%). "
                "Total plan investment across all sectors: approx. NPR 11,481 billion."
            ),
        },
        "instrument_level": "Solid Waste Management National Policy",
        "instrument_comment": {
            "elaboration": (
                "Establishes polluter-pays and EPR-like generator responsibility for industrial, chemical, and health waste. "
                "Also promotes source segregation and the 3Rs."
            ),
            "quote": (
                "Art. 10: 'The local levels will be made responsible for management of household waste, whereas the concerned "
                "organisation or institution will have to be accountable and responsible in managing hazardous, chemical, industrial "
                "and medical waste produced by them.' "
                "Strategy 9.1: 'The local level will be made responsible for household waste management and the related producers "
                "or organizations for the management of hazardous, chemical, industrial and health institutional waste.' "
                "'The policy has also encouraged segregation of waste at source, and reduction, recycling, and reuse of waste.'"
            ),
        },
    },
    {
        "assumptions": "Polluter-pay principles count as valid instruments",
        "policy_inclusion": (
            "Comprehensive Master Plan of Ghodaghodi Lake Area (2077 BS); "
            "Management Plan of Bardia National Park and its Buffer Zone; "
            "Bangladesh-Bhutan-India-Nepal (BBIN) ESMF"
        ),
        "instrument_level": "WTO Trade Policy Review — Report by the Secretariat — Nepal",
        "instrument_comment": {
            "elaboration": (
                "This is a Secretariat report, not an official national policy instrument. "
                "It is unclear whether it should substitute for the actual Excise Duty Act and Customs Act."
            ),
            "question": "Should this WTO report be included instead of the actual Excise Duty Act and Customs Act, even though it is not an official document?",
        },
    },
    {
        "policy_level": "Ban on Plastic Bottles in High-Class Hotels — Directive / Sectoral Order (2024)",
        "policy_level_comment": {
            "elaboration": "The referenced directive could not be located in official sources. Entry may be based on incomplete secondary reporting.",
            "question": "Should this entry be deleted given the actual policy document could not be found?",
        },
    },
]

NEPAL_DELETE = [
    "The Excise Duty Act, 2002 (not an included Act — provisional entry)",
    "WTO Trade Policy Review — Report by the Secretariat — Nepal (non-official document; substitute for Excise/Customs Acts?)",
    "Ban on Plastic Bottles in High-Class Hotels — Directive / Sectoral Order (2024) (source document not found)",
]

MOZAMBIQUE_ENTRIES = [
    {
        "assumptions": (
            "Creation of a governing body is an instrument (may/may not be related to plastics) — "
            "so far only coordination is explicitly mentioned in the prompt"
        ),
        "policy_level": "Regulation on Hazardous Waste Management",
        "policy_level_comment": {
            "elaboration": (
                "Art. 25 sets a general solid-waste infrastructure deadline (close open dumps within 3 years), "
                "not a quantifiable plastic-specific reduction target. Annexes I/II cover plan objectives and reporting metrics. "
                "Fines and payment periods in the annex finance implementation of the Act."
            ),
            "quote": (
                "Art. 25: close open dumps within 3 years. "
                "Annex I/II: plan objectives and reporting metrics — no numeric plastic reduction goals. "
                "Annex: fines and payment periods finance the Act."
            ),
            "proposed": "Policy target currently scored 0; proposed score 1 (general waste infrastructure target with financing mechanism).",
        },
        "instrument_level": "Regulation on Environmental Quality and Effluent Standards",
        "instrument_comment": {
            "elaboration": (
                "Art. 8 regulates atmospheric pollutant emissions from stationary industrial sources. "
                "This is an air-quality/effluent standard, not a plastic waste management instrument."
            ),
            "quote": (
                "Art. 8: 'Establishes atmospheric pollutant emission standards per industrial establishment, "
                "limiting releases from stationary industrial sources.'"
            ),
            "question": "Should this be excluded as not a relevant plastic policy instrument?",
        },
    },
    {
        "assumptions": "Original text including the translations",
        "policy_level": "National Strategy for the Management and Conservation of Coral Reefs",
        "policy_level_comment": {
            "elaboration": (
                "The action plan assigns line-item budgets across three pillars but the in-policy budget is currently scored 0.75, "
                "which is not a valid value on the scoring scale."
            ),
            "quote": (
                "Pillar budgets include: MZN 10 million (reef mapping/validation); MZN 2.1 million (fisheries fiscalization); "
                "MZN 1.1 million (extractive-industry mitigation guide); MZN 8.9 million (environmental education pilot); "
                "MZN 1.7 million (communication plan). Pillar 3 target: sustainable financing mechanism by 2025 (ProAzul, government funds)."
            ),
            "proposed": "Correct invalid in-policy budget score of 0.75 to a valid scale value.",
        },
    },
]

MOZAMBIQUE_DELETE = [
    "Regulation on Environmental Quality and Effluent Standards (air-quality standard — not a relevant plastic instrument)",
]

PERU_ENTRIES = [
    {
        "assumptions": "Original text including the translations",
        "policy_level": "Supreme Decree No. 001-2022-MINAM — Amendments to the Regulation of Legislative Decree",
        "policy_level_comment": {
            "elaboration": (
                "This decree establishes a clear institutional financing mechanism for waste management through involved entities' "
                "budgets and inter-institutional cleaning-fee agreements — qualifying as a full budget source."
            ),
            "quote": (
                "Art. 6: 'financiamiento con cargo al presupuesto institucional de los pliegos involucrados, "
                "sin recursos adicionales del Tesoro Público.' / "
                "'financing from institutional budgets of involved entities, without additional Treasury resources.' "
                "Art. 34-A: 'convenios interinstitucionales para recaudación de arbitrios de limpieza pública.' / "
                "'inter-institutional agreements for collection of municipal cleaning-fee revenue.'"
            ),
            "proposed": "In-policy budget score: 0.5 → 1.",
        },
        "instrument_level": "Law that Regulates the Activity of Waste Pickers (Recyclers)",
        "instrument_comment": {
            "elaboration": (
                "Beyond core regulatory provisions, the law includes symbolic/complementary provisions (national day, awards, "
                "regulation deadline) that may or may not qualify as policy instruments under our coding criteria."
            ),
            "quote": (
                "Complementary Provision First: 'National Recycler Day is commemorated on 1 June each year.' "
                "Complementary Provision Second: 'National Recycling Award created under Ministry of Environment, awarded annually…' "
                "Final Complementary Provision: 'Executive Power must approve, within 120 days of the law's entry into force, "
                "the regulation of this Law by supreme decree.'"
            ),
            "question": "Should these complementary legal/day-declaration provisions be included as instruments?",
        },
    },
    {
        "policy_level": "Law that Amends Legislative Decree No. 1278 — Introducing Industrialization of Recycling",
        "policy_level_comment": {
            "elaboration": (
                "The law promotes public, private, and mixed investment in valorization infrastructure but does not allocate "
                "a specific budget line. Current score of 0.75 is not on the valid scale."
            ),
            "quote": (
                "Art. 6(e): 'promotion of public, private and mixed investment in valorization infrastructure.' "
                "Art. 21(a): 'regional public, mixed or private investment programs for valorization infrastructure.'"
            ),
            "proposed": "In-policy budget score: 0.75 → 0.5.",
        },
        "instrument_level": "Regulation of Law N.° 29419 (Waste Pickers Law)",
        "instrument_comment": {
            "elaboration": (
                "Art. 37 imposes a health/safety requirement on recyclers rather than a waste-management policy instrument. "
                "Relevance to plastic policy coding is questionable."
            ),
            "quote": (
                "Art. 37: 'All recyclers must be vaccinated against Hepatitis B and Tetanus; municipalities promote mass "
                "vaccination with Health Ministry; health facilities must issue vaccination cards.'"
            ),
            "question": "Should Art. 37 (vaccination requirement) be included as an instrument?",
        },
    },
    {
        "policy_level": "Article 9 of the Regulation of Legislative Decree No. 1278 (PLANRES Provision)",
        "policy_level_comment": {
            "elaboration": (
                "PLANRES is a periodic planning/reporting framework (updated every 10 years) with annual progress reporting — "
                "not a quantifiable plastic-specific policy target."
            ),
            "quote": (
                "PLANRES updated every 10 years. Competent authorities submit information to MINAM by last business day of April. "
                "MINAM annually publishes progress reports. MINAM submits annual PLANRES compliance report to CEPLAN and Congress "
                "Environment Commission."
            ),
            "proposed": "Policy target score: 1 → 0.",
        },
    },
]

PERU_DELETE = []

STP_ENTRIES = [
    {
        "instrument_level": "National Plan for Integrated Urban Solid Waste Management (PNGIRSU) 2018–2023",
        "instrument_comment": {
            "elaboration": (
                "This is a waste projection/inventory exercise documenting future special waste streams — "
                "not a regulatory or programmatic instrument."
            ),
            "quote": (
                "Projects special waste streams: '(i) Solventes e tintas; (ii) Pneus; (iii) Óleos usados; "
                "(iv) Resíduos de Equipamentos Elétricos e Eletrónicos; e (v) Veículos em Fim de Vida' — "
                "projected to reach 742 and 758 tonnes/year in 2022 and 2023 respectively."
            ),
            "question": "Should this entry be removed as a projection/inventory rather than an instrument?",
        },
    },
    {
        "instrument_level": "PNGIRSU 2018–2023 — National Waste Characterisation Exercise",
        "instrument_comment": {
            "elaboration": (
                "Founded on approximately one year of national waste data tracking. Produces district-level generation, "
                "collection coverage, and composition data — a data collection exercise, not a policy instrument."
            ),
            "quote": (
                "National average 0.39 kg/capita/day; total 197,700 t/year (2018). Collection coverage: national average 38%. "
                "Waste composition tracks 'Plásticos de todos os tipos' (~5% by weight), biodegradables (~58%), "
                "fines/inerts (~18%), paper/cardboard, glass (~5%), textiles, metals, complex materials, hazardous waste."
            ),
            "question": "Should this data-tracking exercise be removed from the instrument list?",
        },
    },
]

STP_DELETE = [
    "National Plan for Integrated Urban Solid Waste Management (PNGIRSU) 2018–2023 (projection/inventory — not an instrument)",
    "PNGIRSU 2018–2023 — National Waste Characterisation Exercise (data collection — not an instrument)",
]


def main():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    add_title_slide(prs)

    countries = [
        ("Nepal", NEPAL_ENTRIES, NEPAL_DELETE),
        ("Mozambique", MOZAMBIQUE_ENTRIES, MOZAMBIQUE_DELETE),
        ("Peru", PERU_ENTRIES, PERU_DELETE),
        ("São Tomé and Príncipe (STP)", STP_ENTRIES, STP_DELETE),
    ]

    for name, entries, delete_list in countries:
        add_country_slide(prs, name, entries, delete_list)

    out = "/workspace/Policy_Coding_Scoring_Review.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
