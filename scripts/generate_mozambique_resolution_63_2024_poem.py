#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 63/2024 (POEM)."""

import os
from datetime import date

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

OUTPUT_DIR = "/workspace/output/mozambique"
os.makedirs(OUTPUT_DIR, exist_ok=True)

POLICY_FIELDS = {
    "policy_name": (
        "Plano Nacional de Ordenamento do Espaço Marítimo – POEM "
        "(Resolução n.º 63/2024)"
    ),
    "policy_url": (
        "https://sibmoz.gov.mz/content/uploads/2025/11/"
        "Resolucao-n.o-63_2024-que-aprova-o-Plano-de-Ordenamento-do-Espaco-Maritimo.pdf"
    ),
    "policy_year": 2024,
    "policy_objective": (
        "Establish the National Maritime Spatial Plan (POEM) to balance competing uses of "
        "Mozambique's maritime space (~572,000 km²), promote harmonious blue-economy development, "
        "protect marine and coastal ecosystems, reduce marine pollution including land-based debris, "
        "and coordinate sectoral activities through spatial zoning under Sea Law (Law 20/2019)."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'SDG 14.1 alignment: "Até 2025, prevenir e reduzir significativamente a poluição marinha '
        'de todos os tipos... incluindo a poluição de nutrientes e detritos marinhos." Action plan '
        'targets include: 50% of coastal communities involved in recycling/reuse by 2026–2027; '
        'waste management services developed for coastal and maritime activities by 2022–2026; '
        '10 international cooperation protocols on pollution and waste management by 2022–2040.'
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "Council of Ministers Resolution n.º 63/2024 under Sea Law Art. 84(3) — national maritime "
        "spatial planning instrument (POEM), not binding regulation. Sets strategic zoning, "
        "objectives and action plan; implementation through sectoral laws (REPMAR, Decree 45/2006, "
        "Decree 97/2020, etc.)."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging"
    ),
    "policy_circularity": 0.75,
    "policy_lifecycle_phases_list": (
        "consumption, disposal, environmental leakage"
    ),
    "policy_budget": 0.50,
    "policy_budget_text": (
        "Action plan references incentive policy, national operational programme and pilot recycling "
        "projects (OE II.2a, 2023–2027); no dedicated plastic-pollution budget line in resolution text. "
        "International cooperation protocols for pollution/waste management (OE V.3a, 2022–2040)."
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "SDG 14.1 / POEM framework: Adopts target to significantly prevent and reduce all types "
            "of marine pollution by 2025, especially from land-based activities including nutrient "
            "pollution and marine debris (detritos marinhos); aligns with OPRC 90 and Nairobi "
            "Convention through national/regional pollution incident preparedness systems."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'POEM enquadramento: "14.1 Até 2025, prevenir e reduzir significativamente a poluição '
            'marinha... incluindo a poluição de nutrientes e detritos marinhos"; also references '
            'poluição por hidrocarbonetos, OPRC 90 and Convenção de Nairobi.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic target — not plastic-specific but explicit marine debris (detritos marinhos). "
            "Not coded separately: SDG 14.2–14.5 ecosystem/fisheries targets bundled under planning."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "OE II.2a Action Plan: Increase economic valorization of solid waste recycling and "
            "reuse; establish incentive policy, regulatory instruments and national operational "
            "programme; community recycling action plan; targets 50% coastal communities and 80% "
            "coastal municipalities involved by 2026–2027."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'OE II.2a actions: "Estabelecer uma Política de Incentivo ao envolvimento local/ '
            'comunitário na reutilização e reciclagem"; "Operacionalizar a valorização económica e '
            'social da reciclagem e reutilização de resíduos sólidos"; targets 50% comunidades '
            'costeiras envolvidas (2026–2027).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Circular economy direction for coastal solid waste — plastics included implicitly. "
            "Not coded separately: environmental qualification videos/awareness (OE II.2b)."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "OE II.2 Action Plan: Develop waste management network and services for activities in "
            "coastal territories (sanitation, industry, ports, beaches) and at sea (navigation, "
            "platforms); improve solid and liquid waste management in river basins involving local "
            "communities (2022–2026)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action: "Desenvolver a rede e os serviços de gestão de resíduos nas actividades que '
            'correm nos territórios costeiros (saneamento do meio em geral, indústrias, portos, '
            'praias, etc.) e no mar (navegação, plataformas, etc.)"; responsible MOPHRH, MTA, MIMAIP.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary operational waste-management instrument — ports and beaches explicitly named. "
            "Links to Decree 94/2014, Decree 97/2020. Not coded: port-specific indicator monitoring."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Sector zoning provisions: Require effective waste and effluent management reducing "
            "land- and sea-based impacts on coastal ecosystems; vessels must ensure adequate "
            "onboard waste disposal per REPMAR; promote awareness against abandonment of fishing "
            "gear; port areas monitored for waste/effluent and ship fuel/oil spill risks."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Zoning: "gestão de resíduos e efluentes mais eficaz, reduzindo os impactos de origem '
            'terrestre e marinha"; fisheries: "não abandono de artes" and "deposição adequada dos '
            'resíduos criados a bordo, previstos no REPMAR"; ports: monitor waste/effluent indicators.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Spatial planning conditions for fishing gear litter and vessel waste — no explicit "
            "'plástico'. Cruise anchorage: control tourist waste/pollutant discharges near KBAs."
        ),
    },
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "OE V.3a Action Plan: Strengthen international partnerships to operationalize pollution "
            "combat and sustainable waste management; promote circular green-economy experience-sharing "
            "with neighbouring countries; target 10 cooperation protocols by 2022–2040."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'OE V.3a: "Potenciar as parcerias internacionais para operacionalizar o combate à '
            'poluição e a gestão sustentável de resíduos"; action: "Estabelecer protocolos de '
            'cooperação de comunidades/ municípios/ entidades de Moçambique com homólogos dos '
            'países vizinhos" — 10 protocolos (2022–2040).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "International cooperation instrument — complements MARPOL/OPRC/Nairobi references. "
            "Not coded separately: OE V.3b convention participation and scientific cooperation."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "OE III.2 / fisheries zoning: Priority prevention and elimination of illegal, unreported "
            "and unregulated fishing (IUU); fiscalization of permitted gear, mesh sizes and set "
            "duration; reduce destructive trawling impacts; promote non-abandonment of fishing gear "
            "and REPMAR pollution prevention compliance."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'SDG 14.4 / OE III.2: "acabar com a sobrepesca e a pesca ilegal, não reportada e não '
            'regulamentada e as práticas de pesca destrutivas"; zoning: "eliminar o uso de artes '
            'para além do número permitido" and "sensibilização para o não abandono de artes".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Ghost gear relevance indirect — gear abandonment explicitly addressed in zoning. "
            "Not coded separately: tuna/cephalopod area rules, aquaculture (OE III.3)."
        ),
    },
]

COLUMNS = [
    "policy_name", "policy_url", "policy_year", "policy_objective", "policy_target",
    "policy_target_text", "policy_type", "policy_type_justification", "policy_integration",
    "policy_sectors_list", "policy_circularity", "policy_lifecycle_phases_list",
    "policy_budget", "policy_budget_text", "policy_score",
    "instrument_type", "instrument_lifecycle_stage", "instrument_description",
    "instrument_in_force", "instrument_implementation", "instrument_implementation_text",
    "instrument_score", "comments",
]

COLUMN_LABELS = {
    "policy_name": "A — policy_name",
    "policy_url": "B — policy_url",
    "policy_year": "C — policy_year",
    "policy_objective": "D — policy_objective",
    "policy_target": "E — policy_target",
    "policy_target_text": "F — policy_target_text",
    "policy_type": "G — policy_type",
    "policy_type_justification": "H — policy_type_justification",
    "policy_integration": "I — policy_integration",
    "policy_sectors_list": "J — policy_sectors_list",
    "policy_circularity": "K — policy_circularity",
    "policy_lifecycle_phases_list": "L — policy_lifecycle_phases_list",
    "policy_budget": "M — policy_budget",
    "policy_budget_text": "N — policy_budget_text",
    "policy_score": "O — policy_score [auto]",
    "instrument_type": "P — instrument_type",
    "instrument_lifecycle_stage": "Q — instrument_lifecycle_stage",
    "instrument_description": "R — instrument_description",
    "instrument_in_force": "S — instrument_in_force",
    "instrument_implementation": "T — instrument_implementation",
    "instrument_implementation_text": "U — instrument_implementation_text",
    "instrument_score": "V — instrument_score [auto]",
    "comments": "W — comments",
}


def build_dataframe():
    return pd.DataFrame([{**POLICY_FIELDS, **i} for i in INSTRUMENTS], columns=COLUMNS)


def export_csv(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.csv")
    df.rename(columns=COLUMN_LABELS).to_csv(path, index=False, encoding="utf-8-sig")
    return path


def export_xlsx(df, basename):
    path = os.path.join(OUTPUT_DIR, f"{basename}.xlsx")
    labelled = df.rename(columns=COLUMN_LABELS)
    labelled.to_excel(path, index=False, sheet_name="4P Index Coding")
    wb = load_workbook(path)
    ws = wb.active
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(wrap_text=True, vertical="top")
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = min(45, max(12, len(str(ws.cell(1, col).value or "")) * 0.9))
    ws.freeze_panes = "A2"
    wb.save(path)
    return path


def add_heading(doc, text, level=1):
    doc.add_heading(text, level=level)


def add_para(doc, text):
    doc.add_paragraph(text)


def create_translation_doc():
    doc = Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("National Maritime Spatial Plan — POEM", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 63/2024 of 15 November")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Plano Nacional de Ordenamento do Espaço Marítimo – POEM "
        "(Resolução n.º 63/2024)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://sibmoz.gov.mz/content/uploads/2025/11/"
        "Resolucao-n.o-63_2024-que-aprova-o-Plano-de-Ordenamento-do-Espaco-Maritimo.pdf"
    )

    # (heading, body) or (heading, [(subheading, body), ...])
    sections = [
        ("Resolution No. 63/2024 — Enacting Provisions", (
            "Having regard to the need to establish a National Maritime Spatial Plan that balances "
            "different uses and promotes their harmonious development, under Article 84(3) of "
            "Law No. 20/2019 (Sea Law), the Council of Ministers resolves:\n\n"
            "Article 1. The National Maritime Spatial Plan, abbreviated POEM, is approved as an "
            "annex forming an integral part of this Resolution.\n\n"
            "Article 2. The Ministry responsible for the maritime sector shall ensure implementation "
            "of POEM.\n\n"
            "Article 3. This Resolution enters into force on the date of its publication.\n\n"
            "Approved by the Council of Ministers on 9 November 2021. Published in the Official "
            "Gazette (Boletim da República), Series I, No. 222, 15 November 2024."
        )),
        ("1. Introduction and Legal Framework", [
            ("1.1 Policy context", (
                "POEM responds to government priorities on sustainable management of natural resources "
                "and the environment, blue-economy integration, improved territorial planning, and "
                "strengthened monitoring and fiscalization. It implements Sea Law (Law 20/2019), POLMAR "
                "(Resolution 39/2017), the Regulation on Use of National Maritime Space (RJUEM, "
                "Decree 21/2017), and the Integrated Coastal Zone Management Strategy (EGIZC).\n\n"
                "Mozambique's maritime space covers approximately 572,000 km², including territorial "
                "waters, internal waters and the Exclusive Economic Zone (EEZ). Major uses include "
                "fisheries, aquaculture, maritime transport, ports, tourism, oil and gas, mining, "
                "scientific research and marine conservation."
            )),
            ("1.2 Threats identified", (
                "POEM identifies threats to marine and coastal ecosystems including: overfishing and "
                "unregulated resource exploitation; inappropriate tourism practices; pollution and "
                "marine litter (lixo marinho); coastal erosion and mangrove destruction; extreme "
                "weather and climate change; and unplanned coastal urban and industrial development.\n\n"
                "Coastal water degradation from pollution and ocean acidification adversely affects "
                "ecosystems, biodiversity and small-scale fisheries."
            )),
            ("1.3 Related instruments", (
                "POEM articulates with: POLMAR; EGIZC; Mangrove Management Strategy 2020–2024; "
                "National Environmental Management Programme (PNGA); National Biodiversity Strategy; "
                "Integrated Urban Solid Waste Management Strategy; Environmental Quality and Effluents "
                "Regulation; REPMAR (Fisheries Regulation, Decree 89/2020); Decree 45/2006 on marine "
                "pollution; and Decree 97/2020 on coastal zones and beaches."
            )),
        ]),
        ("2. POEM Purpose and Objectives", [
            ("2.1 General purpose", (
                "POEM covers the entire National Maritime Space and identifies:\n"
                "a) the spatial and temporal distribution of existing and potential uses and activities;\n"
                "b) natural and cultural values of strategic importance for environmental sustainability "
                "and intergenerational solidarity.\n\n"
                "Uses mapped include: aquaculture and fisheries; biotechnology; mineral resources; "
                "energy and renewables; scientific research; recreation, sport and tourism; underwater "
                "cultural heritage; and infrastructure and equipment."
            )),
            ("2.2 Strategic objectives (summary)", (
                "POEM pursues five core objectives:\n"
                "1. Establish maritime spatial ordering respecting environmental carrying capacity;\n"
                "2. Promote sustainable, rational and integrated economic exploitation;\n"
                "3. Ensure preservation, protection and recovery of ecosystems and biodiversity;\n"
                "4. Guarantee legal security and transparency of maritime space use;\n"
                "5. Ensure quality of information on maritime space conditions.\n\n"
                "Additional tasks include: characterizing the coastal and maritime environment; "
                "identifying high environmental value zones; preventing use conflicts; conducting "
                "Strategic Environmental and Social Assessment (SESA/AASE); and defining indicators "
                "and a monitoring and evaluation plan."
            )),
        ]),
        ("3. Guiding Principles", [
            ("3.1 Priority principles", (
                "Sustainable development and balance — integrate social, economic and environmental "
                "objectives, balancing economic development with conservation of marine and coastal "
                "resources.\n\n"
                "Multisectoral coordination and integrated management — harmonize sectors whose "
                "common denominator is the sea.\n\n"
                "Community involvement and participation — active citizen and traditional authority "
                "participation in planning, decision-making, use and conservation.\n\n"
                "Economic capacity building — improve maritime economic activities over the long term."
            )),
            ("3.2 Supplementary principles", (
                "Include: ecosystem approach; precaution; adaptive management; environmental "
                "education; co-responsibility; rational use of environmental components; social "
                "equity; polluter-pays and user-pays in coastal management options; and coordinated "
                "scientific and interdisciplinary approach."
            )),
        ]),
        ("4. International Commitments — SDG 14 and Conventions", [
            ("4.1 SDG 14 targets adopted in POEM", (
                "14.1 — By 2025, prevent and significantly reduce all types of marine pollution, "
                "in particular from land-based activities, including nutrient pollution and marine debris.\n\n"
                "14.2 — Sustainably manage and protect marine and coastal ecosystems, strengthening "
                "resilience and restoring ocean health.\n\n"
                "14.4 — By 2020, effectively regulate harvesting; end overfishing and IUU fishing "
                "and destructive practices; implement science-based management plans.\n\n"
                "14.5 — Conserve at least 10% of coastal and marine areas.\n\n"
                "14.a — Increase scientific knowledge and marine technology transfer."
            )),
            ("4.2 Conventions referenced", (
                "UNCLOS; MARPOL (International Convention for the Prevention of Pollution from Ships); "
                "OPRC 90 (oil pollution preparedness and response); UN Convention on Biological "
                "Diversity; Nairobi Convention (land-based sources of marine pollution).\n\n"
                "POEM commits to national and regional systems for pollution incident preparedness "
                "and response, international cooperation, and improved prevention and combat techniques."
            )),
        ]),
        ("5. Maritime Spatial Zoning", [
            ("5.1 Zoning framework", (
                "POEM classifies maritime space into categories reflecting conservation levels, use "
                "intensity and sectoral activities. Zoning identifies: navigation channels and traffic "
                "separation schemes; pilotage areas; dredging zones; aids to navigation; anchorage "
                "and mooring areas; ports and marinas; coastal defence works; dredge disposal zones; "
                "submarine cables and pipelines; munitions and hazardous material deposition areas; "
                "and wreck sites.\n\n"
                "Spatial planning must coordinate maritime and terrestrial coastal planning, as many "
                "maritime activities require land-based infrastructure."
            )),
            ("5.2 Conservation and sensitive ecosystems", (
                "Marine Protected Areas and sensitive ecosystems (mangroves, dunes, coral reefs, "
                "seagrass beds) receive protection classifications. Activities causing pollution, "
                "habitat alteration or waste generation require stronger controls. Effective waste "
                "and effluent management must reduce terrestrial and marine impacts from industry, "
                "mining and aquaculture."
            )),
            ("5.3 Sector-specific zoning conditions", (
                "Fisheries: IUU fishing prevention prioritized; fiscalization of permitted gear, mesh "
                "sizes and fishing duration; awareness against abandonment of fishing gear; adequate "
                "onboard waste disposal per REPMAR.\n\n"
                "Ports: Monitor indicators for port-area waste/effluents and ship fuel/oil spills.\n\n"
                "Wrecks and cultural heritage: No ship discharges or washing near identified wrecks "
                "as pollution/residues endanger heritage; no munitions deposition.\n\n"
                "Tourism/cruise anchorage: Control environmental impacts including generator use and "
                "discharge of pollutants and litter near Key Biodiversity Areas and critical habitats."
            )),
        ]),
        ("6. Strategic Axes and Action Plan — Plastic-Relevant Elements", [
            ("6.1 EE II — Marine and coastal environment", (
                "OE II.2 Environmental quality: Increase economic valorization of solid waste recycling "
                "and reuse; promote environmental qualification and circular economy practices with "
                "technological, scientific and NGO support.\n\n"
                "OE II.2a Actions (2023–2027):\n"
                "— Establish incentive policy for local/community involvement in reuse and recycling;\n"
                "— Define and publish legal and regulatory instruments (2024–2025);\n"
                "— Prepare National Action Plan for community recycling (2025–2026);\n"
                "— Support pilot recycling/reuse projects (2022–2026);\n"
                "— Target: 50% of coastal communities involved by 2026–2027; 80% of coastal municipalities.\n\n"
                "OE II.2 — Waste services: Develop waste management network and services for coastal "
                "territories (sanitation, industry, ports, beaches) and at sea (navigation, platforms) "
                "by 2022–2026 (MOPHRH, MTA, MIMAIP)."
            )),
            ("6.2 EE III — Socioeconomic development", (
                "OE III.2 Fisheries: Minimize over-exploitation and illegal/inappropriate practices; "
                "increase marine conservation areas; strengthen fiscalization; combat IUU fishing "
                "(Pesca INN); publish lists of IUU vessels; expand fisheries fiscalization coverage.\n\n"
                "OE III.3 Aquaculture: Offshore mariculture development with environmental qualification "
                "to preserve investments; adequate waste deposition from mariculture infrastructure."
            )),
            ("6.3 EE V — Sovereignty and international cooperation", (
                "OE V.3a International partnerships: Operationalize pollution combat and sustainable "
                "waste management; promote circular green-economy experience-sharing with neighbouring "
                "countries; establish 10 cooperation protocols (2022–2040); organize experience-sharing "
                "meetings and field visits.\n\n"
                "OE V.3b: Strengthen Mozambique's active participation in international commissions "
                "promoting sustainable resource use and marine protected areas."
            )),
        ]),
        ("7. Monitoring and Implementation", (
            "POEM includes indicators and verification means for each strategic objective, with "
            "responsible entities (MIMAIP, MTA, MOPHRH, MINEC, municipalities, etc.) and timelines "
            "from 2022 to 2040.\n\n"
            "The Ministry responsible for the maritime sector ensures POEM implementation. Monitoring "
            "covers pollution parameters from research and hydrocarbon installations, environmental "
            "quality mapping, and progress on waste management service development.\n\n"
            "POEM does not replace applicable laws; it provides spatial planning guidance that "
            "sectoral regulations must operationalize."
        )),
        ("8. Note on Plastics", (
            "POEM does not mention plastics or plastic bags explicitly. The only direct reference to "
            "marine litter is 'lixo marinho' in the threats analysis. Plastic-relevant elements are:\n\n"
            "— SDG 14.1 target on marine debris (detritos marinhos);\n"
            "— Solid waste recycling, reuse and circular economy action plan (OE II.2a);\n"
            "— Coastal and maritime waste management services including ports and beaches;\n"
            "— Fishing gear non-abandonment and vessel waste disposal per REPMAR;\n"
            "— Tourism/cruise litter discharge controls near sensitive areas.\n\n"
            "Binding operational rules remain in Decree 45/2006, Decree 97/2020, REPMAR, Sea Law "
            "and Decree 16/2015 (plastic bags)."
        )),
    ]

    for heading, content in sections:
        add_heading(doc, heading, level=1)
        if isinstance(content, str):
            add_para(doc, content)
        else:
            for subheading, body in content:
                add_heading(doc, subheading, level=2)
                add_para(doc, body)

    doc.add_page_break()
    add_heading(doc, "4P Index Coding Summary", level=1)
    add_para(
        doc,
        f"Policy: {POLICY_FIELDS['policy_name']}\n"
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — sector zoning chapters not overcoded)\n"
        f"See: 4P_Index_Mozambique_Resolution_63_2024_POEM.xlsx",
    )

    path = os.path.join(OUTPUT_DIR, "Mozambique_Resolution_63_2024_POEM_English_Translation.docx")
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 63/2024 POEM — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 63/2024 (POEM)</h1>
  <p>Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Resolution_63_2024_POEM_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_63_2024_poem.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_63_2024_POEM"
    df = build_dataframe()
    paths = [
        export_csv(df, basename),
        export_xlsx(df, basename),
        create_translation_doc(),
        create_index_html(basename),
    ]
    print("Generated:")
    for p in paths:
        print(f"  {p}")


if __name__ == "__main__":
    main()
