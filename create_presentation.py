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
BODY_COLOR = RGBColor(0x33, 0x33, 0x33)
LIGHT_BG = RGBColor(0xF4, 0xF6, 0xF8)
BORDER = RGBColor(0xD0, 0xD7, 0xDE)


def set_font(run, size=10, bold=False, italic=False, color=BODY_COLOR):
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color


def add_agenda_slide(prs):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.55))
    bar.fill.solid()
    bar.fill.fore_color.rgb = TITLE_COLOR
    bar.line.fill.background()
    ttf = bar.text_frame
    ttf.margin_left = Pt(14)
    ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tr = ttf.paragraphs[0].add_run()
    tr.text = "Agenda"
    set_font(tr, size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    items = [
        "Review coding assumptions applied per country",
        "Discuss policy inclusion/exclusion decisions",
        "Evaluate policy-level field coding and scoring",
        "Evaluate instrument-level field coding and scoring",
        "Resolve open questions for scoring consistency",
    ]
    box = slide.shapes.add_textbox(Inches(0.9), Inches(1.1), Inches(8.2), Inches(4.5))
    tf = box.text_frame
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(14)
        p.level = 0
        num = p.add_run()
        num.text = f"{i + 1}.  "
        set_font(num, size=16, bold=True, color=ACCENT)
        body = p.add_run()
        body.text = item
        set_font(body, size=16, color=BODY_COLOR)

    framework = slide.shapes.add_textbox(Inches(0.9), Inches(5.2), Inches(8.2), Inches(1.5))
    ftf = framework.text_frame
    fp = ftf.paragraphs[0]
    fr = fp.add_run()
    fr.text = (
        "Slide structure mirrors the review spreadsheet: Assumptions → Policy Inclusion/Exclusions → "
        "Policy-Level Fields → Instrument-Level Fields → Open Questions."
    )
    set_font(fr, size=11, italic=True, color=RGBColor(0x66, 0x66, 0x66))


def add_title_slide(prs, title, subtitle):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    bg.fill.solid()
    bg.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    bg.line.fill.background()

    bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.18))
    bar.fill.solid()
    bar.fill.fore_color.rgb = TITLE_COLOR
    bar.line.fill.background()

    box = slide.shapes.add_textbox(Inches(0.75), Inches(2.0), Inches(8.5), Inches(1.2))
    tf = box.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.CENTER
    r = p.add_run()
    r.text = title
    set_font(r, size=32, bold=True, color=TITLE_COLOR)

    sub = slide.shapes.add_textbox(Inches(0.75), Inches(3.3), Inches(8.5), Inches(0.8))
    stf = sub.text_frame
    sp = stf.paragraphs[0]
    sp.alignment = PP_ALIGN.CENTER
    sr = sp.add_run()
    sr.text = subtitle
    set_font(sr, size=16, italic=True, color=ACCENT)

    date_box = slide.shapes.add_textbox(Inches(0.75), Inches(4.8), Inches(8.5), Inches(0.4))
    dtf = date_box.text_frame
    dp = dtf.paragraphs[0]
    dp.alignment = PP_ALIGN.CENTER
    dr = dp.add_run()
    dr.text = "Review Meeting — July 2026"
    set_font(dr, size=12, color=RGBColor(0x66, 0x66, 0x66))


def add_section_header(slide, left, top, width, text):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, Inches(0.28))
    shape.fill.solid()
    shape.fill.fore_color.rgb = HEADER_BG
    shape.line.fill.background()
    tf = shape.text_frame
    tf.margin_left = Pt(6)
    tf.margin_top = Pt(2)
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = text
    set_font(r, size=9, bold=True, color=HEADER_FG)


def add_text_block(slide, left, top, width, height, lines, font_size=8):
    box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    box.fill.solid()
    box.fill.fore_color.rgb = LIGHT_BG
    box.line.color.rgb = BORDER
    tf = box.text_frame
    tf.word_wrap = True
    tf.margin_left = Pt(5)
    tf.margin_right = Pt(5)
    tf.margin_top = Pt(4)
    tf.margin_bottom = Pt(4)
    tf.vertical_anchor = MSO_ANCHOR.TOP

    for i, (label, content) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(3)
        if label:
            lr = p.add_run()
            lr.text = f"{label}: "
            set_font(lr, size=font_size, bold=True, color=ACCENT)
        if content:
            cr = p.add_run()
            cr.text = content
            set_font(cr, size=font_size, color=BODY_COLOR)


def add_country_slide(prs, country, entries, questions=None):
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Country title bar
    title_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, Inches(0.55))
    title_bar.fill.solid()
    title_bar.fill.fore_color.rgb = TITLE_COLOR
    title_bar.line.fill.background()
    ttf = title_bar.text_frame
    ttf.margin_left = Pt(14)
    ttf.vertical_anchor = MSO_ANCHOR.MIDDLE
    tp = ttf.paragraphs[0]
    tr = tp.add_run()
    tr.text = country
    set_font(tr, size=22, bold=True, color=RGBColor(0xFF, 0xFF, 0xFF))

    # Column headers
    col_top = Inches(0.62)
    col_h = Inches(0.26)
    left_margin = Inches(0.2)
    content_width = Inches(9.6)
    col_w = content_width / 4

    headers = [
        "Assumptions",
        "Policy Inclusion / Exclusions",
        "Policy-Level Fields & Comments",
        "Instrument-Level Fields & Comments",
    ]
    for i, h in enumerate(headers):
        add_section_header(slide, left_margin + col_w * i, col_top, col_w - Inches(0.04), h)

    # Entry blocks
    entry_top = Inches(0.92)
    q_reserved = Inches(1.15) if questions else Inches(0.15)
    available_h = Inches(7.5) - entry_top - q_reserved - Inches(0.2)
    n = len(entries)
    gap = Inches(0.05)
    block_h = (available_h - gap * (n - 1)) / n if n else available_h

    for idx, entry in enumerate(entries):
        top = entry_top + (block_h + gap) * idx
        cols = [
            entry.get("assumptions", ""),
            _format_policy_inclusion(entry),
            _format_policy_level(entry),
            _format_instrument_level(entry),
        ]
        for i, text in enumerate(cols):
            if not text.strip():
                text = "—"
            fs = 6.5 if len(text) > 500 else 7 if len(text) > 300 else 7.5
            add_text_block(
                slide,
                left_margin + col_w * i,
                top,
                col_w - Inches(0.04),
                block_h,
                [("", text)],
                font_size=fs,
            )

    # Questions row
    if questions:
        q_top = Inches(7.5) - q_reserved
        add_section_header(slide, left_margin, q_top, content_width, "Open Questions for Discussion")
        add_text_block(
            slide,
            left_margin,
            q_top + Inches(0.28),
            content_width,
            q_reserved - Inches(0.32),
            [(f"{i+1}.", q) for i, q in enumerate(questions)],
            font_size=7.5,
        )

    # Footer
    footer = slide.shapes.add_textbox(Inches(0.2), Inches(7.22), Inches(9.6), Inches(0.2))
    ftf = footer.text_frame
    fr = ftf.paragraphs[0].add_run()
    fr.text = "Policy Coding & Scoring Review  |  July 2026  |  Confidential"
    set_font(fr, size=7, italic=True, color=RGBColor(0x99, 0x99, 0x99))


def _format_policy_inclusion(entry):
    parts = []
    if entry.get("policy_inclusion"):
        parts.append(entry["policy_inclusion"])
    if entry.get("policy_inclusion_comment"):
        parts.append(f"Comment: {entry['policy_inclusion_comment']}")
    return "\n".join(parts)


def _format_policy_level(entry):
    parts = []
    if entry.get("policy_level"):
        parts.append(entry["policy_level"])
    if entry.get("policy_level_comment"):
        parts.append(f"Comment: {entry['policy_level_comment']}")
    return "\n".join(parts)


def _format_instrument_level(entry):
    parts = []
    if entry.get("instrument_level"):
        parts.append(entry["instrument_level"])
    if entry.get("instrument_comment"):
        parts.append(entry["instrument_comment"])
    return "\n".join(parts)


NEPAL_ENTRIES = [
    {
        "assumptions": "Mentions plastics explicitly",
        "policy_inclusion": "The Excise Duty Act, 2002",
        "policy_inclusion_comment": "Not included Acts — included in sheet for now, will delete. Don't mention plastics explicitly.",
        "policy_level": "Directive on the Ban of Single-Use Plastics in the Sagarmatha (Everest) Region",
        "policy_level_comment": "Policy sectors list: tourism, waste management, conservation, water, retail, packaging",
        "instrument_level": "Environment-friendly Local Governance Framework, 2013",
        "instrument_comment": (
            "COMPREHENSIVE MULTI-LEVEL PLASTIC BAG RESTRICTION AND BAN SYSTEM:\n"
            "HOUSEHOLD (Municipal + VDC): Basic — plastics collected separately (p.11, p.19). "
            "Advanced — plastic bag regulation; jute/cloth/paper bags (p.12, p.20).\n"
            "TOLE (Advanced): No non-decomposing wastes/plastics in public places (p.13).\n"
            "MUNICIPALITY (Basic): Ban on plastic use per national standard; eco-friendly bags encouraged (p.15).\n"
            "VDC (Basic): Plastic use discouraged. VDC Advanced: eco-friendly bag manufacturing encouraged (p.21–22).\n"
            "DISTRICT (Basic): Plastic bag free district declaration; district mechanism to control import (p.21)."
        ),
    },
    {
        "assumptions": "Anything remotely national",
        "policy_inclusion": "Customs Act (Law) 2007",
        "policy_level": "Action Plan for Ban on Plastic Bags, 2022",
        "policy_level_comment": (
            "Reference to other sources — source document in Nepali with severely corrupted OCR. "
            "Coding based on: (a) partial Nepali text extraction; (b) English reports (Himalayan Times, myRepublica); "
            "(c) SWITCH-Asia Nepal plastic policy overview (2025)."
        ),
        "instrument_level": "Local Government Operation Act",
        "instrument_comment": (
            "SCHEDULE — MUNICIPAL FUNCTIONS (Health & Sanitation): waste management, health waste collection/disposal, "
            "private-sector coordination (items 6, 7, 10).\n"
            "SCHEDULE (Local Market & Environment): sanitation/waste management at local level; low-carbon development (15–16).\n"
            "WARD FUNCTIONS: collection/management of household waste (item 11).\n"
            "MUNICIPALITY CLASSIFICATION: 'solid waste processing and management system' as classification criterion."
        ),
    },
    {
        "assumptions": "Frameworks are also included",
        "policy_inclusion": "National Urban Water Supply and Sanitation Sector Policy, 2009",
        "policy_level": "The Sixteenth Plan (FY 2024/25–2028/29)",
        "policy_level_comment": (
            "Policy budget score — Water supply, sewerage, waste management and recycling: NPR 3,991 million "
            "(Public 52.4% / Private 46.4% / Cooperative 1.2%). Total plan investment ~NPR 11,481 billion. "
            "Budget exists for waste management/recycling but not ring-fenced for plastic."
        ),
        "instrument_level": "Solid Waste Management National Policy",
        "instrument_comment": (
            "Art. 10 — Generator responsibility / polluter pays: local levels responsible for household waste; "
            "producers accountable for hazardous/industrial/medical waste.\n"
            "EPR-like provision (Strategy 9.1): producers responsible for hazardous/chemical/industrial/health waste.\n"
            "Waste segregation at source + 3R encouraged. School curriculum on waste management."
        ),
    },
    {
        "assumptions": "Polluter-pay principles count as valid instruments",
        "policy_inclusion": (
            "Comprehensive Master Plan of Ghodaghodi Lake Area (2077 BS)\n"
            "Management Plan of Bardia National Park and its Buffer Zone\n"
            "Bangladesh-Bhutan-India-Nepal (BBIN) ESMF"
        ),
        "instrument_level": "WTO Trade Policy Review — Report by the Secretariat — Nepal",
        "instrument_comment": "Should I include this or the actual Excise Duty and Customs Act? Can this substitute although not an official document?",
    },
    {
        "assumptions": "",
        "policy_level": "Ban on Plastic Bottles in High-Class Hotels — Directive / Sectoral Order (2024)",
        "instrument_comment": "Couldn't find the actual policy — should I delete it?",
    },
]

NEPAL_QUESTIONS = [
    "How will we analyze the qualitative data?",
    "Should the WTO Trade Policy Review be included instead of the actual Excise Duty and Customs Act?",
    "Should the Ban on Plastic Bottles in High-Class Hotels (2024) entry be deleted (policy not found)?",
]

MOZAMBIQUE_ENTRIES = [
    {
        "assumptions": "Creation of a governing body is an instrument — so far only coordination is explicitly mentioned in the prompt",
        "policy_level": "Regulation on Hazardous Waste Management",
        "policy_level_comment": (
            "Policy target — not a quantifiable plastic-specific target. Art. 25 (close open dumps in 3 years) is a general "
            "solid waste infrastructure deadline, not a plastic target. Annex I/II refer to plan objectives and reporting "
            "metrics, not numeric plastic reduction goals. Annex mentions fines, payment periods — fines finance the act."
        ),
        "instrument_level": "Regulation on Environmental Quality and Effluent Standards",
        "instrument_comment": "Art. 8: Establishes atmospheric pollutant emission standards per industrial establishment, limiting releases from stationary industrial sources.",
    },
    {
        "assumptions": "Original text including the translations",
        "policy_level": "National Strategy for the Management and Conservation of Coral Reefs",
        "policy_level_comment": (
            "In-policy budget scored 0.75 — doesn't exist on score scale. Text: Action plan assigns line-item budgets "
            "across three pillars (e.g. MZN 10M reef mapping; MZN 2.1M fisheries fiscalization; MZN 1.1M extractive-industry "
            "mitigation; MZN 8.9M environmental education; MZN 1.7M communication). Pillar 3 target: sustainable financing by 2025."
        ),
    },
]

MOZAMBIQUE_QUESTIONS = [
    "Should the in-policy budget score of 0.75 be corrected (not on valid score scale)?",
]

PERU_ENTRIES = [
    {
        "assumptions": "Original text including the translations",
        "policy_level": "Supreme Decree No. 001-2022-MINAM — Amendments to the Regulation of Legislative Decree",
        "policy_level_comment": (
            "Budget currently rated 0.5 — proposed score: 1. "
            "Art. 6: financing from institutional budgets of involved entities, without additional Treasury resources. "
            "Art. 34-A: inter-institutional agreements for collection of municipal cleaning-fee revenue."
        ),
        "instrument_level": "Law that Regulates the Activity of Waste Pickers (Recyclers)",
        "instrument_comment": (
            "Should we include complementary provisions (mostly legal/day declarations)?\n"
            "• National Recycler Day: 1 June each year.\n"
            "• National Recycling Award under Ministry of Environment.\n"
            "• Executive Power must approve regulation within 120 days of law's entry into force."
        ),
    },
    {
        "assumptions": "",
        "policy_level": "Law that Amends Legislative Decree No. 1278 — Introducing Industrialization of Recycling",
        "policy_level_comment": (
            "Budget scored 0.75 (invalid on scale). Art. 6(e): promotion of public/private/mixed investment in "
            "valorization infrastructure. Art. 21(a): regional investment programs. No specific budget allocation. "
            "Proposed score: 0.5."
        ),
        "instrument_level": "Regulation of Law N.° 29419 (Waste Pickers Law)",
        "instrument_comment": (
            "Should we include Art. 37: All recyclers must be vaccinated against Hepatitis B and Tetanus; "
            "municipalities promote mass vaccination with Health Ministry."
        ),
    },
    {
        "assumptions": "",
        "policy_level": "Article 9 of the Regulation of Legislative Decree No. 1278 (PLANRES Provision)",
        "policy_level_comment": (
            "Currently scored 1 for policy targets — proposed score: 0. "
            "PLANRES updated every 10 years. Authorities submit info to MINAM by last business day of April. "
            "MINAM publishes annual progress reports and compliance report to CEPLAN and Congress."
        ),
    },
]

PERU_QUESTIONS = [
    "Should Supreme Decree 001-2022-MINAM budget be scored 1 instead of 0.5?",
    "Should LD 1278 amendment budget be scored 0.5 instead of 0.75?",
    "Should PLANRES provision (Art. 9) policy target score be 0 instead of 1?",
    "Should complementary provisions in the Waste Pickers Law be included as instruments?",
    "Should Art. 37 (vaccination requirement) in the Waste Pickers Regulation be included?",
]

STP_ENTRIES = [
    {
        "assumptions": "",
        "instrument_level": "National Plan for Integrated Urban Solid Waste Management (PNGIRSU) 2018–2023",
        "instrument_comment": (
            "Proposed removal — projection/inventory, not an instrument. Documents special waste streams "
            "(solvents, tyres, used oils, WEEE, end-of-life vehicles) projected at 742–758 tonnes/year in 2022–2023."
        ),
    },
    {
        "assumptions": "",
        "instrument_level": "PNGIRSU 2018–2023 (waste characterisation exercise)",
        "instrument_comment": (
            "Proposed removal — data collection exercise, not an instrument. Founded on ~1 year national waste data tracking. "
            "District-level generation data (0.39 kg/capita/day; 197,700 t/year as of 2018). Collection coverage 38% national avg. "
            "Waste composition tracks 'Plásticos de todos os tipos' (~5% by weight) alongside biodegradables (~58%), "
            "fines/inerts (~18%), paper/cardboard, glass (~5%), textiles, metals, complex materials, hazardous waste."
        ),
    },
]

STP_QUESTIONS = [
    "Should both PNGIRSU 2018–2023 entries be removed (projection/inventory vs. instrument)?",
]

NORWAY_ENTRIES = [
    {
        "assumptions": "No entries recorded in review sheet",
        "policy_inclusion": "",
        "policy_level": "",
        "instrument_level": "",
        "instrument_comment": "Awaiting data — please confirm whether Norway should be included in this review cycle.",
    },
]

NORWAY_QUESTIONS = [
    "Should Norway be included in this review? No coding/scoring entries were provided.",
]


def main():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)

    add_title_slide(
        prs,
        "Policy Coding & Scoring Review",
        "Country-by-Country Review of Assumptions, Policy Inclusions, and Instrument-Level Coding",
    )
    add_agenda_slide(prs)

    countries = [
        ("Nepal", NEPAL_ENTRIES, NEPAL_QUESTIONS),
        ("Norway", NORWAY_ENTRIES, NORWAY_QUESTIONS),
        ("Mozambique", MOZAMBIQUE_ENTRIES, MOZAMBIQUE_QUESTIONS),
        ("Peru", PERU_ENTRIES, PERU_QUESTIONS),
        ("São Tomé and Príncipe (STP)", STP_ENTRIES, STP_QUESTIONS),
    ]

    for name, entries, questions in countries:
        add_country_slide(prs, name, entries, questions)

    out = "/workspace/Policy_Coding_Scoring_Review.pptx"
    prs.save(out)
    print(f"Saved: {out}")


if __name__ == "__main__":
    main()
