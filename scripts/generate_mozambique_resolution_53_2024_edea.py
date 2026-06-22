#!/usr/bin/env python3
"""Generate 4P Index coding and English translation for Mozambique Resolution 53/2024 (EDEA)."""

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
        "Estratégia de Desenvolvimento da Economia Azul – EDEA "
        "(Resolução n.º 53/2024)"
    ),
    "policy_url": (
        "https://www.proazul.gov.mz/wp-content/uploads/2024/11/"
        "EDEA-ESTRATEGIA-DE-DESENVOLVIMENTO-DA-ECONOMIA-AZUL-_FINAL_TYPO_PRINT_WEB.pdf"
    ),
    "policy_year": 2024,
    "policy_objective": (
        "Approve the Blue Economy Development Strategy (EDEA) 2024–2033 to harness "
        "Mozambique's marine, coastal and inland water potential (~562,000 km² maritime space, "
        "~2,700 km coastline) through six strategic pillars and four transversal axes, integrating "
        "natural capital, circular economy and coastal governance with sustainable fisheries, ports, "
        "tourism, renewable energy and maritime security — explicitly targeting plastic pollution "
        "prevention, plastic production/import reduction, extended producer responsibility and "
        "marine/coastal pollution prevention capacity."
    ),
    "policy_target": 1,
    "policy_target_text": (
        'Pilar 3 objective c: "Reduzir a produção, importação e uso de plásticos, aumentando a '
        'responsabilidade alargada dos sectores da indústria e do comércio e os apoios e incentivos '
        'à reciclagem e indústria de remanufactura" (2024–2027 action plan). Additional targets: '
        'establish incentive policy and National Action Plan for community recycling/reuse (2024–2026); '
        'SADC circular-economy/waste cooperation protocols (2024–2025); illegal/inappropriate fishing '
        'practices reduced 50% by 2027; beach/sea plastic waste context in Maputo (8% of collected waste).'
    ),
    "policy_type": 0.50,
    "policy_type_justification": (
        "Council of Ministers Resolution n.º 53/2024 — national blue-economy development strategy "
        "with integrated action plan (2024–2033), not binding sectoral regulation. Sets strategic "
        "objectives, governance (CNEA/SNEA/CTEA) and implementation targets; operational rules "
        "remain in Sea Law (20/2019), POLMAR (39/2017), POEM (63/2024), Decree 45/2006, Decree "
        "97/2020, Decree 16/2015 (plastic bags), etc."
    ),
    "policy_integration": 1.0,
    "policy_sectors_list": (
        "fisheries, waste management, industry, tourism, municipalities, water, packaging, ports"
    ),
    "policy_circularity": 1.0,
    "policy_lifecycle_phases_list": (
        "production, consumption, disposal, environmental leakage"
    ),
    "policy_budget": 0.75,
    "policy_budget_text": (
        "Action plan assigns budget lines to Pilar 3 circular-economy actions: MZN 8 billion "
        "(SADC waste cooperation, 2024–2025); MZN 15 billion (incentive policy and regulatory "
        "instruments for recycling/reuse, 2024–2026); MZN 260 billion (plastic reduction, EPR, "
        "recycling/remanufacturing pilot projects, 2024–2027). Additional Pilar 5 port "
        "sustainability investments referenced without dedicated plastic line."
    ),
    "policy_score": "[auto]",
}

INSTRUMENTS = [
    {
        "instrument_type": 0.20,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pilar 3 strategic framework: Circular economy and remanufacturing/recycling identified "
            "as the cornerstone for plastic pollution prevention; economic circularity requires "
            "public and private measures incentivizing recycling and discouraging import/production "
            "of hard-to-degrade products flowing through the environmental chain."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Pilar 3 narrative: "pedra angular para a prevenção da poluição plástica" through '
            'reassignment, remanufacturing/recycling and reintroduction into the economic circuit; '
            'measures incentivizing recycling and discouraging import/production of products of '
            'difficult degradation.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Strategic direction — primary explicit plastic-pollution-prevention framing. "
            "Not coded separately: biodiversity conservation objectives (Pilar 3 a–d) or "
            "bioexploration objectives."
        ),
    },
    {
        "instrument_type": 0.60,
        "instrument_lifecycle_stage": "Production",
        "instrument_description": (
            "Pilar 3 objective c / Action Plan: Reduce plastic production, import and use; "
            "increase extended producer responsibility (EPR) for industry and commerce sectors; "
            "provide support and incentives for recycling and remanufacturing industry; operationalize "
            "economic and social valorization of solid waste recycling and reuse (2024–2027)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Objective c: "Reduzir a produção, importação e uso de plásticos, aumentando a '
            'responsabilidade alargada dos sectores da indústria e do comércio e os apoios e '
            'incentivos à reciclagem e indústria de remanufactura"; action: "Operacionalizar a '
            'valorização económica e social da reciclagem e reutilização de resíduos sólidos" — '
            'pilot projects for industry/commerce sectors (MEF, MADER, MTA, MIC, 2024–2027).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Primary explicit plastic instrument — production/import reduction plus EPR. "
            "Indicators: operational recycling/reuse projects, industries and districts involved, "
            "economic value and employment generated. Complements Decree 16/2015 plastic bags."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Recycling",
        "instrument_description": (
            "Pilar 3 objectives a–b / Action Plan: Promote blue environmental and circular economy "
            "approaches creating employment and improving environmental quality; strengthen "
            "investment in innovation, management and waste-treatment/valorization technology; "
            "establish incentive policy for local/community involvement in reuse and recycling; "
            "prepare National Action Plan for implementation (2024–2026)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action Plan P3-b: "Estabelecer uma Política de Incentivo ao envolvimento local/ '
            'comunitário na reutilização e reciclagem"; define and publish legal/regulatory '
            'instruments (2024–2026); "preparado um Plano de Acção Nacional para a implementação '
            'dessa política"; SADC partnerships for circular-economy waste projects (2024–2025).'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Operational circular-economy/waste-valorization instrument — plastics included "
            "implicitly in solid waste. Context: 98% of collected waste dumped in uncontrolled "
            "dumps; Maputo waste composition 8% plastic. Not coded: bioexploration objectives."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Transversal axis — coastal governance objective a: Promote capacity for prevention "
            "and combat of marine and coastal pollution in inland waters; develop National Coastal "
            "Management Policy; integrate POEM zoning into coastal territorial plans; strengthen "
            "international convention implementation (MARPOL, Nairobi Convention)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Coastal governance objective a: "Promover a capacidade de prevenção e combate à '
            'poluição marinha e costeira nas águas interiores"; strategic action iv: "Elaborar e '
            'implementar a Política Nacional de Gestão Costeira"; context cites "resíduos plásticos '
            'nas praias e no mar".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Pollution-prevention capacity instrument — explicit beach/sea plastic waste in "
            "Pilar 3 context. Links to POEM, Decree 97/2020, Decree 45/2006. Not coded separately: "
            "National Blue Economy legislative code or investment directory."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Waste management",
        "instrument_description": (
            "Pilar 5 — Blue ports: Transform traditional ports into 'blue ports' with GHG reduction, "
            "sustainable liquid and solid waste management, prevention of exotic species invasion "
            "via ships, dredging impact minimization and port-city articulation; fiscal/financial "
            "incentives for environmental sustainability adaptation."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Pilar 5: "gestão sustentável de resíduos líquidos e sólidos"; objective a: "Garantir '
            'a transformação e modernização dos portos nacionais em portos azuis"; action ii: '
            '"Criar incentivos financeiros e fiscais para investimentos de adaptação e transformação '
            'dos portos para responder aos requisitos de sustentabilidade ambiental".'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Port waste-management instrument — solid/liquid waste at Maputo, Beira, Nacala ports. "
            "Carbon-neutral port target (climate axis). Not coded separately: cabotage or shipbuilding "
            "incentives."
        ),
    },
    {
        "instrument_type": 0.80,
        "instrument_lifecycle_stage": "Environmental leakage",
        "instrument_description": (
            "Pilar 1 — Fisheries action plan: Promote conservation and sustainable management of "
            "marine, coastal and inland water resources; reduce illegal/inappropriate fishing practices "
            "50% by 2027 and 90% by 2040; eliminate shore trawling by 2027; sensitize fishers on "
            "sustainable regulated gear — indirect relevance to abandoned/lost fishing gear (ghost gear)."
        ),
        "instrument_in_force": 1,
        "instrument_implementation": 0.50,
        "instrument_implementation_text": (
            'Action Plan P1-e/i: "Práticas ilegais e/ou inadequadas reduzidas em 50% até 2027"; '
            '"Pesca com redes de arrasto para terra eliminada até 2027"; "Sensibilizar os pescadores '
            'para a utilização de artes mais sustentáveis e regulamentadas" — 90% reduction by 2040.'
        ),
        "instrument_score": "[auto]",
        "comments": (
            "Bundled fisheries/IUU instrument — no explicit 'plástico' but addresses destructive "
            "gear and illegal practices linked to marine debris. Pilar 6 maritime security also "
            "references IUU fishing. Not coded separately: aquaculture disease-risk objectives."
        ),
    },
]

COLUMNS = [
    "policy_name",
    "policy_url",
    "policy_year",
    "policy_objective",
    "policy_target",
    "policy_target_text",
    "policy_type",
    "policy_type_justification",
    "policy_integration",
    "policy_sectors_list",
    "policy_circularity",
    "policy_lifecycle_phases_list",
    "policy_budget",
    "policy_budget_text",
    "policy_score",
    "instrument_type",
    "instrument_lifecycle_stage",
    "instrument_description",
    "instrument_in_force",
    "instrument_implementation",
    "instrument_implementation_text",
    "instrument_score",
    "comments",
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
        ws.column_dimensions[letter].width = min(
            45, max(12, len(str(ws.cell(1, col).value or "")) * 0.9)
        )
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

    title = doc.add_heading("Blue Economy Development Strategy — EDEA", 0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub = doc.add_paragraph("Resolution No. 53/2024 — Period 2024–2033")
    sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta = doc.add_paragraph()
    meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    meta.add_run(
        "English translation of: Estratégia de Desenvolvimento da Economia Azul – EDEA "
        "(Resolução n.º 53/2024)\n"
    ).italic = True
    meta.add_run(f"Translated for 4P Index policy analysis — {date.today().isoformat()}\n")
    meta.add_run(
        "Source: https://www.proazul.gov.mz/wp-content/uploads/2024/11/"
        "EDEA-ESTRATEGIA-DE-DESENVOLVIMENTO-DA-ECONOMIA-AZUL-_FINAL_TYPO_PRINT_WEB.pdf"
    )

    sections = [
        ("Resolution No. 53/2024 — Enacting Provisions", (
            "Having regard to the need to promote sustainable exploitation of Mozambique's marine, "
            "coastal and inland water resources and to establish an integrated national instrument "
            "for blue-economy development aligned with government priorities, international "
            "commitments and existing maritime policy (POLMAR, Sea Law), the Council of Ministers "
            "resolves:\n\n"
            "Article 1. The Blue Economy Development Strategy (EDEA) for the period 2024–2033, "
            "including its Action Plan, is approved as an annex forming an integral part of this "
            "Resolution.\n\n"
            "Article 2. The Ministry of the Sea, Inland Waters and Fisheries (MIMAIP) shall ensure "
            "coordination of EDEA implementation, supported by the National Blue Economy Council "
            "(CNEA), the National Secretariat for the Blue Economy (SNEA) and the Blue Economy "
            "Technical Committee (CTEA).\n\n"
            "Article 3. This Resolution enters into force on the date of its publication.\n\n"
            "Approved by the Council of Ministers on 17 July 2024. Launched publicly in Maputo "
            "on 12 August 2024. Published in the Official Gazette (Boletim da República), "
            "Series I, September 2024."
        )),
        ("Preface — President of the Republic", (
            '"Investing in Ocean Health is Investing in the Planet\'s Future."\n\n'
            "With the approval by the Council of Ministers of the Blue Economy Development Strategy "
            "(EDEA) 2024–2033, Mozambique strategically positions itself to maximize its immense "
            "marine, coastal and inland water capital.\n\n"
            "A vibrant Blue Economy in Mozambique rests on a vast maritime space of 562,000 km², "
            "an extensive coastline of approximately 2,700 km along the Indian Ocean, and extensive "
            "inland waters including the Zambezi, Rovuma, Lúrio, Licungo, Púngue, Búzi, Save, "
            "Limpopo and Incomati river basins, Lake Niassa, Cahora Bassa and Chicamba Real "
            "reservoirs, and numerous lagoons and wetlands.\n\n"
            "The Strategy is intrinsically linked to the Green Economy and Circular Economy, "
            "requiring integrated approaches that prioritize sustainable resource exploitation and "
            "maximize by-product use in a circular production and consumption chain without "
            "endangering ecosystems. The central aim is that abundant marine, coastal and aquatic "
            "resources become a blessing rather than a curse for communities and the national economy.\n\n"
            "Six pillars structure the Strategy: (1) Fisheries and aquaculture; (2) Renewable energy "
            "and marine extractive industry; (3) Natural capital, environment and circular economy; "
            "(4) Tourism and culture; (5) Maritime transport and port/logistics infrastructure; "
            "(6) Maritime security.\n\n"
            "EDEA consolidates policy gains including POLMAR (2017), Sea Law (2019), the Coral Reef "
            "Management and Conservation Strategy 2022–2032, and international conventions. It "
            "materializes SDG 14 and commitments under UNFCCC, CBD, the African Union Agenda 2063, "
            "the African Convention on Conservation of Nature and Natural Resources, the Nairobi "
            "Convention, and the African Blue Economy Strategy 2019.\n\n"
            "Maputo, 12 August 2024 — Filipe Jacinto Nyusi, President of the Republic of Mozambique."
        )),
        ("1. Context and Legal Framework", [
            ("1.1 Importance of the sea and justification for EDEA", (
                "The blue economy concept has become established in national strategic planning as "
                "an integrated approach balancing sustainable use of ocean resources, improved "
                "living standards, ecosystem protection, job creation, poverty eradication and "
                "climate change response — requiring close collaboration across society.\n\n"
                "Mozambique's maritime space (~562,000 km²) encompasses mineral and ecological "
                "riches supporting economic capacity and growth. The ~2,700 km coastline, where "
                "approximately 30% of the population lives (60% dependent on the sea), includes "
                "sandy beaches with dunes and lagoons, estuaries, mangroves and coral reefs. "
                "In 2015 the Government created MIMAIP to ensure sustainable governance and "
                "exploitation of marine, coastal and inland water potential across economic, social "
                "and environmental dimensions."
            )),
            ("1.2 Global and national framing", (
                "Globally, EDEA aligns with SDG 14 (life below water), the African Union Agenda "
                "2063, the African Blue Economy Strategy, UNCLOS, MARPOL, the Nairobi Convention "
                "and the Pan-African Fisheries and Aquaculture Policy Reform Strategy.\n\n"
                "Nationally, EDEA integrates with POLMAR, Sea Law (20/2019), POEM (63/2024), "
                "Integrated Coastal Zone Management Strategy (EGIZC), National Energy Strategy "
                "(ENDE 2025–2044), National Biodiversity Strategy, and sectoral regulations on "
                "fisheries (REPMAR), marine pollution (Decree 45/2006) and coastal zones "
                "(Decree 97/2020).\n\n"
                "The absence of an overarching blue-economy policy instrument has led sectors to "
                "act in isolation, undermining integrated management of marine and coastal natural "
                "capital. EDEA addresses this gap."
            )),
            ("1.3 Guiding principles", (
                "EDEA is guided by: sustainable development and balance; risk assessment; due "
                "diligence; transparency, community involvement and participation; accountability "
                "(actors assume responsibility for negative impacts and adopt prevention, "
                "minimization and elimination measures); economic capacity building and innovation; "
                "ecosystem approach; precaution; adaptive management; environmental education; "
                "co-responsibility; rational use of environmental components; social equity; "
                "polluter-pays and user-pays principles."
            )),
        ]),
        ("2. Strategic Framework", [
            ("2.1 Objective, vision and mission", (
                'Objective: "To drive aquatic economic potential through application of scientific '
                'and local knowledge, ensuring economic, environmental and social sustainability '
                'and inclusion."\n\n'
                'Vision: "Marine, coastal and inland water resources and associated activities '
                'contributing effectively to Mozambique\'s sustainable development."\n\n'
                'Mission: "Promote sustainable exploitation of marine, coastal and inland water '
                'natural capital that meets present needs and ensures conservation and value '
                'creation for future generations."'
            )),
            ("2.2 Six strategic pillars and four transversal axes", (
                "Six pillars:\n"
                "1. Fisheries and aquaculture (fisheries, aquaculture production, processing, marketing)\n"
                "2. Renewable energy and marine extractive industry (wind, solar, waves/tides, "
                "hydrocarbons, minerals)\n"
                "3. Natural capital, environment and circular economy (biodiversity, bioprospecting, "
                "circular economy, coastal management)\n"
                "4. Tourism and culture (coastal, maritime, nature and purpose-driven tourism)\n"
                "5. Maritime transport and port/logistics infrastructure (ports, cabotage, "
                "shipbuilding/repair)\n"
                "6. Maritime security (protection, fiscalization, search and rescue)\n\n"
                "Four transversal axes: Good governance (including maritime/coastal planning); "
                "Knowledge, innovation and technologies; Climate change; Communities."
            )),
        ]),
        ("3. Pilar 3 — Natural Capital, Environment and Circular Economy", [
            ("3.1 Natural capital and circular economy context", (
                "Natural capital comprises the value the environment delivers to the economy through "
                "ecosystem services. Mozambique possesses exceptional coastal, marine and inland water "
                "habitats supporting fisheries, tourism and carbon sequestration (mangroves cover "
                "nearly 3,000 km²). Marine Protected Areas cover 2.2% of national maritime space.\n\n"
                "Reintegrating end-of-life materials enables transition to a circular economy — "
                "described as the cornerstone for plastic pollution prevention and operational across "
                "multiple sources through reassignment, remanufacturing/recycling and reintroduction "
                "into the economic circuit, while creating jobs and stimulating technological innovation.\n\n"
                "Circular economy benefits include reduced environmental pressure, greater raw-material "
                "security, competitiveness, innovation and economic growth. Public and private actors "
                "must adopt measures incentivizing recycling and discouraging import or production of "
                "hard-to-degrade products."
            )),
            ("3.2 Waste and plastic pollution context", (
                "According to POEM data cited in EDEA, Maputo urban area generates on average "
                "1 kg of solid waste per person per day (0.49 kg in suburban/KaTembe areas; "
                "0.20 kg in KaNyaka). Of collected waste: 69% organic, 5% paper/cardboard, "
                "8% plastic, 9% other including rubble.\n\n"
                "More than 98% of collected waste is dumped in uncontrolled landfills, generating "
                "an estimated 1.5 million tonnes CO₂ equivalent per year (~USD 37.5 million/year "
                "at USD 25/tonne). Socioeconomic considerations include plastic waste on beaches "
                "and at sea, health costs from pollution, and impacts on tourism and fisheries value "
                "chains.\n\n"
                "EDEA must promote blue environmental, circular-economy and coastal-management "
                "initiatives that valorize environmental quality in coastal zones."
            )),
            ("3.3 Strategic objectives — circular economy and coastal management", (
                "For environmental and circular economy:\n"
                "a) Disseminate concepts and promote blue environmental and circular economy "
                "approaches creating employment and improving environmental quality, in coordination "
                "with neighbouring countries;\n"
                "b) Strengthen investment in innovation, management and waste-treatment/valorization "
                "technology;\n"
                "c) Reduce plastic production, import and use; increase extended producer "
                "responsibility (EPR) for industry and commerce; provide support and incentives for "
                "recycling and remanufacturing industry.\n\n"
                "For coastal management:\n"
                "a) Strengthen nature-based climate resilience, prevention and recovery from extreme "
                "events and maritime accidents;\n"
                "b) Promote conservation and restoration of coastal ecosystems (mangroves, coral "
                "reefs, seagrass) minimizing coastal erosion and extreme climate impacts."
            )),
            ("3.4 Strategic actions (summary)", (
                "Key actions include: mapping and economic valuation of ecosystems and services; "
                "expanding marine protected area networks; fiscal incentive packages for natural "
                "capital conservation; codes of good practice for coastal/maritime activities; "
                "safeguarding ecosystems in TUPEM and DUAT titles; community involvement in "
                "ecosystem recovery and resilient coastal infrastructure; disaster risk-transfer "
                "mechanisms including parametric climate insurance; and deepening SADC partnerships "
                "to establish circular-economy projects improving environmental management and "
                "particularly waste management with local community involvement.\n\n"
                "Expected result: Increase area, quality, resilience and protection of sensitive "
                "ecosystems through valuation of services and minimization of threats."
            )),
        ]),
        ("4. Transversal Axis — Coastal Governance and Pollution Prevention", [
            ("4.1 Planning context", (
                "Coastal governance integrates POEM (national maritime spatial plan), EGIZC, "
                "provincial/district territorial plans (PNDT, PDUT, PEU) and the forthcoming National "
                "Coastal Management Policy. POEM covers the entire national maritime space; coastal "
                "territorial plans may extend to the 12-nautical-mile territorial sea."
            )),
            ("4.2 Strategic objectives", (
                "a) Promote capacity for prevention and combat of marine and coastal pollution in "
                "inland waters;\n"
                "b) Improve business environment through legal/administrative conditions, dispute "
                "resolution, e-governance and financing lines for blue-economy activities;\n"
                "c) Improve articulation between ocean governance, blue economy and territorial "
                "planning sectors (central and local decentralization);\n"
                "d) Develop and ensure access to credible databases for blue-economy, environmental "
                "and circular investment;\n"
                "e) Promote equitable participation of women and youth in natural resource management;\n"
                "f) Ensure articulation between plans affecting maritime space, coastal zones and "
                "inland waters;\n"
                "g) Improve appropriate combination of uses supporting long-term ecosystem "
                "sustainability and community socioeconomic integration;\n"
                "h) Value/recover coastal habitats contributing to coastal resilience."
            )),
            ("4.3 Key strategic actions", (
                "i. Strengthen adherence and domestic regulation of international conventions;\n"
                "ii. Create a Blue Economy legislative code;\n"
                "iii. Prepare an investment opportunities directory;\n"
                "iv. Develop and implement the National Coastal Management Policy;\n"
                "v. Integrate coastal/maritime economic uses into PDUT and PEU per POEM guidance;\n"
                "vi. Include POEM guidance in SEA/EIA and concession contracts (DUAT, TUPEM);\n"
                "vii. Maintain updated mapping of maritime uses and marine/coastal habitats."
            )),
        ]),
        ("5. Pilar 5 — Maritime Transport and Blue Ports", [
            ("5.1 Context", (
                "The Mozambique Channel is a major route for national and international maritime "
                "trade. Three main commercial ports (Maputo, Beira, Nacala) with development corridors "
                "position Mozambique as a regional logistics hub. Maritime transport contributes "
                "relatively little to cargo movement despite high potential; port infrastructure "
                "requires modernization to meet international standards."
            )),
            ("5.2 Blue ports and waste management", (
                "EDEA must: (i) transform traditional ports into 'blue ports' meeting sustainability "
                "objectives — GHG reduction, sustainable liquid and solid waste management, prevention "
                "of exotic species invasion via ships, dredging impact minimization and port-city "
                "articulation; (ii) dynamize integrated transport (maritime, road, rail, pipeline, air) "
                "and national cabotage/shipbuilding; (iii) ensure navigation safeguards sensitive areas.\n\n"
                "Strategic objective a: Ensure transformation and modernization of national ports into "
                "blue ports contributing to sustainable national and regional port development.\n\n"
                "Strategic actions include fiscal/financial incentives for port investments meeting "
                "environmental sustainability requirements. Expected result: sustainable modernization "
                "of maritime transport and port services with reduced GHG emissions and impacts on "
                "marine biodiversity."
            )),
        ]),
        ("6. Pilar 1 — Fisheries (Marine Debris Relevance)", (
            "Fisheries and sustainable aquaculture are essential to EDEA. Mozambique has fishing "
            "potential of 937,581 tonnes (48.5% utilization; 455,544 tonnes recorded catches). "
            "POLMAR objectives orient the sector toward combating illegal, unreported and unregulated "
            "(IUU) fishing and sustainable resource management.\n\n"
            "Action Plan targets under objective e/i:\n"
            "— Illegal and/or inappropriate practices reduced 50% by 2027;\n"
            "— Shore trawling eliminated by 2027;\n"
            "— Sensitize fishers on sustainable regulated gear; illegal practices reduced 90% by 2040.\n\n"
            "Indicators include reduction of non-conventional fishing gear and adoption of more "
            "sustainable legal technologies. Pilar 6 (maritime security) also identifies IUU fishing, "
            "illegal resource trafficking and smuggling as threats requiring regional cooperation."
        )),
        ("7. Governance, Monitoring and Action Plan", [
            ("7.1 Governance model", (
                "Political coordination — National Blue Economy Council (CNEA): chaired by MIMAIP, "
                "co-chaired by MEF; meets twice yearly; decides on blue-economy projects/programmes, "
                "deliberates strategic matters, appraises regulations for Council of Ministers, "
                "validates monitoring/evaluation reports.\n\n"
                "Executive coordination — National Secretariat for the Blue Economy (SNEA): technical "
                "support unit integrating strategic actions, monitoring alignment between objectives "
                "and actor actions; composed of National Directorate of Blue Economy, National "
                "Directorate of Policies and Cooperation, and ProAzul Blue Economy Development Fund.\n\n"
                "Technical Committee (CTEA): provides scientific/technical foundation for political "
                "decisions; includes representatives from ministries and pillar lead institutions "
                "(P3 natural capital/circular economy led by MTA)."
            )),
            ("7.2 Action Plan — Pilar 3 plastic and waste targets", (
                "P3-a (circular economy dissemination): Deepen SADC partnerships for circular-economy "
                "waste projects with community involvement; targets cooperation protocols and "
                "information sharing on waste management in Zambezi, Save, Limpopo, Incomati and "
                "Maputo international river basins (2024–2025; budget MZN 8 billion).\n\n"
                "P3-b (waste technology): Establish policy/regulatory instruments for economic and "
                "social valorization of recycling/reuse; incentive policy for local/community "
                "involvement; National Action Plan prepared (2024–2026; budget MZN 15 billion; "
                "lead MTA).\n\n"
                "P3-c (plastic reduction): Operationalize economic/social valorization of solid waste "
                "recycling and reuse; define and execute Action Plan and exemplary pilot projects for "
                "industry and commerce sectors involved in recycling and remanufacturing (2024–2027; "
                "budget MZN 260 billion; lead MEF with MADER, MTA, MIC, private sector, municipalities).\n\n"
                "Indicators: operational recycling/reuse projects; industries, districts and "
                "municipalities involved; economic value generated (% GDP); employment created."
            )),
            ("7.3 International cooperation on pollution and waste", (
                "Transversal governance axis includes strengthening international partnerships to "
                "operationalize pollution combat and sustainable waste management. Mozambique "
                "develops multilateral scientific cooperation programmes to share pollution-combat "
                "and waste-management experiences with neighbouring countries and build an East "
                "African marine and coastal biodiversity knowledge network (2024–2034)."
            )),
        ]),
        ("8. Note on Plastics — Explicit and Implicit References", (
            "EDEA contains five explicit plastic-related references in the strategy text:\n\n"
            "1. Circular economy as 'cornerstone for plastic pollution prevention' (Pilar 3);\n"
            "2. Maputo waste composition: 8% plastic (POEM data cited);\n"
            "3. 'Plastic waste on beaches and at sea' listed among pollutant streams;\n"
            "4. Strategic objective c: reduce plastic production, import and use; increase EPR; "
            "incentives for recycling and remanufacturing;\n"
            "5. Action Plan operational targets and budgets for objectives a–c (2024–2027).\n\n"
            "Implicit plastic-relevant elements include: blue ports solid/liquid waste management; "
            "marine/coastal pollution prevention capacity; SADC circular-economy waste cooperation; "
            "uncontrolled dumps (98% of collected waste); fisheries gear regulation and IUU reduction "
            "(ghost gear); MARPOL and Nairobi Convention references; alignment with POEM waste/recycling "
            "targets and Decree 16/2015 (plastic bags).\n\n"
            "EDEA is a strategic planning instrument; binding operational plastic rules remain in "
            "sectoral laws and regulations."
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
        f"Instruments coded: {len(INSTRUMENTS)} (consolidated — biodiversity/tourism pillars not overcoded)\n"
        f"See: 4P_Index_Mozambique_Resolution_53_2024_EDEA.xlsx",
    )

    path = os.path.join(
        OUTPUT_DIR, "Mozambique_Resolution_53_2024_EDEA_English_Translation.docx"
    )
    doc.save(path)
    return path


def create_index_html(basename):
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Mozambique Resolution 53/2024 EDEA — 4P Index</title>
  <style>
    body {{ font-family: system-ui, sans-serif; max-width: 800px; margin: 2rem auto; padding: 0 1rem; }}
    .card {{ border: 1px solid #ddd; border-radius: 8px; padding: 1rem; margin: 1rem 0; }}
    a.download {{ display: inline-block; background: #1F4E79; color: #fff; padding: 0.6rem 1.2rem;
      border-radius: 4px; text-decoration: none; margin-top: 0.5rem; }}
  </style>
</head>
<body>
  <h1>4P Index — Resolution 53/2024 (EDEA)</h1>
  <p>Blue Economy Development Strategy 2024–2033. Generated {date.today().isoformat()}. Six consolidated instrument rows.</p>
  <div class="card">
    <a class="download" href="{basename}.xlsx" download>Download Excel</a><br><br>
    <a class="download" href="{basename}.csv" download>Download CSV</a><br><br>
    <a class="download" href="Mozambique_Resolution_53_2024_EDEA_English_Translation.docx" download>Download English translation</a>
  </div>
</body>
</html>"""
    path = os.path.join(OUTPUT_DIR, "index_resolution_53_2024_edea.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path


def main():
    basename = "4P_Index_Mozambique_Resolution_53_2024_EDEA"
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
