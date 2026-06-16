#!/usr/bin/env python3
"""Generate bilingual Mozambique plastic pollution policy list as a Word document."""

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

OUTPUT = "/workspace/Mozambique_Plastic_Pollution_Policies_Bilingual.docx"

POLICIES = [
    {
        "title_en": "National Environmental Policy (Resolution No. 5/95)",
        "title_pt": "Política Nacional do Ambiente (Resolução n.º 5/95)",
        "type": "Policy",
        "year": "1995",
        "description": "Approves Mozambique's national environmental policy, establishing principles and guidelines for sustainable use of natural resources, pollution prevention, and environmental management across sectors.",
        "sectors": "Environment; all economic sectors",
        "lifecycle": "All stages (framework)",
        "link": None,
        "link_note": "No standalone official PDF identified on publicly accessible government portals (verified via sibmoz.gov.mz legal framework, June 2026).",
    },
    {
        "title_en": "Environmental Law (Law No. 20/97)",
        "title_pt": "Lei do Ambiente (Lei n.º 20/97)",
        "type": "Law",
        "year": "1997",
        "description": "Framework environmental law prohibiting unauthorized release of pollutants, establishing bases for environmental management, licensing, and sustainable development.",
        "sectors": "Environment; industry; municipalities; agriculture",
        "lifecycle": "All stages (framework)",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/1547563349-Lei-do-Ambiente.pdf",
    },
    {
        "title_en": "Regulation on Management and Control of Plastic Bags (Decree No. 16/2015)",
        "title_pt": "Regulamento sobre a Gestão e Controlo do Saco de Plástico (Decreto n.º 16/2015)",
        "type": "Regulation",
        "year": "2015",
        "description": "Regulates production, import, marketing, and use of plastic bags; prohibits thin plastic bags, sets labelling and thickness requirements, and phases in import/production bans.",
        "sectors": "Industry; commerce; municipalities; environment",
        "lifecycle": "Production; distribution; use; end-of-life",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/Regulamento-Sobre-a-Gestao-e-Controlo-do-Saco-de-Plastico.pdf",
    },
    {
        "title_en": "Regulation on Extended Producer Responsibility for Packaging (Decree No. 79/2017)",
        "title_pt": "Regulamento sobre a Responsabilidade Alargada dos Produtores e Importadores de Embalagens (Decreto n.º 79/2017)",
        "type": "Regulation",
        "year": "2017",
        "description": "Establishes extended producer responsibility (EPR) for packaging, including internal management systems, packaging environmental fee (TAE), and packaging standardisation.",
        "sectors": "Industry; commerce; environment; municipalities",
        "lifecycle": "Production; distribution; end-of-life",
        "link": None,
        "link_note": "No standalone official PDF identified on publicly accessible government portals (verified via sibmoz.gov.mz legal framework, June 2026).",
    },
    {
        "title_en": "Regulation on Urban Solid Waste Management (Decree No. 94/2014)",
        "title_pt": "Regulamento sobre a Gestão de Resíduos Sólidos Urbanos (Decreto n.º 94/2014)",
        "type": "Regulation",
        "year": "2014",
        "description": "National regulation on urban solid waste management: segregation (including plastics), collection, transport, treatment, valorisation, landfill rules, and municipal responsibilities.",
        "sectors": "Municipalities; environment; public health; private waste operators",
        "lifecycle": "Collection; end-of-life",
        "link": "https://gpa.co.mz/wp-content/uploads/2023/08/REGULAMENTO-DA-LEI-DA-CONCORRENCIA_2014-1.pdf",
        "link_note": "Official Boletim da República PDF (Decreto 94/2014 full text).",
    },
    {
        "title_en": "Regulation on Hazardous Waste Management (Decree No. 83/2014)",
        "title_pt": "Regulamento sobre a Gestão de Resíduos Perigosos (Decreto n.º 83/2014)",
        "type": "Regulation",
        "year": "2014",
        "description": "Sets procedures for safe management, transport, storage, and disposal of hazardous wastes from industrial and other activities.",
        "sectors": "Industry; health; environment; municipalities",
        "lifecycle": "Production; end-of-life",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/Aprova-o-Regulamento-sobre-Gestao-de-Residuos-perigosos-e-os-respectivos-anexos.pdf",
    },
    {
        "title_en": "National Sustainable Waste Management Programme (ValoRe)",
        "title_pt": "Programa Nacional de Gestão Sustentável de Resíduos (ValoRe)",
        "type": "Program",
        "year": "2025",
        "description": "Government circular-economy programme to build sustainable waste-treatment infrastructure, strengthen recycling value chains, and pilot municipal waste management.",
        "sectors": "Municipalities; environment; industry; finance",
        "lifecycle": "Collection; end-of-life",
        "link": None,
        "link_note": "No standalone official programme document PDF identified on publicly accessible government portals (June 2026).",
    },
    {
        "title_en": "Regulation on Environmental Quality and Effluent Standards (Decree No. 18/2004, as amended by Decree No. 67/2010)",
        "title_pt": "Regulamento sobre Padrões de Qualidade Ambiental e de Emissão de Efluentes (Decreto n.º 18/2004, alterado pelo Decreto n.º 67/2010)",
        "type": "Regulation",
        "year": "2010",
        "description": "Sets national environmental quality and effluent/emission standards for air, water, and mobile sources; limits pollutant releases and supports enforcement.",
        "sectors": "Industry; environment; energy; transport",
        "lifecycle": "Production; end-of-life",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/Regulamento-sobre-Padroes-de-Qualidade-Ambiental-e-de-Emissao-de-Efluentes.pdf",
    },
    {
        "title_en": "Regulation on Prevention of Marine and Coastal Pollution (Decree No. 45/2006)",
        "title_pt": "Regulamento para Prevenção da Poluição e Proteção do Ambiente Marinho e Costeiro (Decreto n.º 45/2006)",
        "type": "Regulation",
        "year": "2006",
        "description": "Requires measures to prevent, control, and combat marine pollution from ships and land-based sources in Mozambique's jurisdictional waters.",
        "sectors": "Maritime; fisheries; ports; environment",
        "lifecycle": "Distribution; end-of-life",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/1547623391-Regulamento-sobre-a-Prevencao-da-Poluicao-ea-Proteccao-Ambiente-Marinho-e-Costeiro-Decreto-45-barra-2006.pdf",
    },
    {
        "title_en": "Regulation on Environmental Inspection (Decree No. 51/2024)",
        "title_pt": "Regulamento sobre a Actividade de Fiscalização Ambiental (Decreto n.º 51/2024)",
        "type": "Regulation",
        "year": "2024",
        "description": "Approves rules for environmental inspection and enforcement, replacing Decree 11/2006 and defining inspection types and competencies.",
        "sectors": "Environment; all regulated sectors",
        "lifecycle": "All stages (enforcement)",
        "link": "https://sibmoz.gov.mz/content/uploads/2025/11/Regulamento-de-fiscalizac%CC%A7a%CC%83o-ambiental.pdf",
    },
    {
        "title_en": "Environmental Impact Assessment Regulation (Decree No. 54/2015)",
        "title_pt": "Regulamento da Avaliação do Impacto Ambiental (Decreto n.º 54/2015)",
        "type": "Regulation",
        "year": "2015",
        "description": "Regulates environmental impact assessment and licensing for projects that may cause significant environmental harm, including mitigation hierarchy.",
        "sectors": "Industry; infrastructure; mining; energy; tourism",
        "lifecycle": "Production; end-of-life (project planning)",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/1547556163-Decreto-54-2015-Regulamento-da-AIA.pdf",
    },
    {
        "title_en": "Environmental Audit Process Regulation (Decree No. 45/2024)",
        "title_pt": "Regulamento sobre o Processo de Auditoria Ambiental (Decreto n.º 45/2024)",
        "type": "Regulation",
        "year": "2024",
        "description": "Establishes rules for environmental audits of public and private activities that may impact the environment.",
        "sectors": "Industry; environment; public administration",
        "lifecycle": "Production; end-of-life (compliance)",
        "link": "https://sibmoz.gov.mz/content/uploads/2025/11/Regulamento-sobre-o-Processo-de-Auditoria-Ambiental.pdf",
    },
    {
        "title_en": "Sea Policy and Strategy – POLMAR (Resolution No. 39/2017)",
        "title_pt": "Política e Estratégia do Mar – POLMAR (Resolução n.º 39/2017)",
        "type": "Policy / Strategy",
        "year": "2017",
        "description": "National sea policy and implementation strategy guiding sustainable, integrated use of maritime and coastal spaces and blue-economy development.",
        "sectors": "Maritime; fisheries; tourism; ports; environment",
        "lifecycle": "All stages (marine/coastal framework)",
        "link": "https://www.proazul.gov.mz/wp-content/uploads/2023/10/POLITICA-E-ESTRATEGIA-DO-MAR-POLMAR-1.pdf",
    },
    {
        "title_en": "Sea Law (Law No. 20/2019)",
        "title_pt": "Lei do Mar (Lei n.º 20/2019)",
        "type": "Law",
        "year": "2019",
        "description": "Legal regime for sovereignty, jurisdiction, exploitation of marine resources, and use of national maritime public domain.",
        "sectors": "Maritime; fisheries; mining; transport",
        "lifecycle": "All stages (marine framework)",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/Lei-do-Mar.pdf",
    },
    {
        "title_en": "Coastal Zone and Beaches Management Regulation (Decree No. 97/2020)",
        "title_pt": "Regulamento de Gestão e Ordenamento da Zona Costeira e das Praias (Decreto n.º 97/2020)",
        "type": "Regulation",
        "year": "2020",
        "description": "Defines principles and norms for integrated, sustainable management and development of coastal zones and beaches nationwide.",
        "sectors": "Municipalities; tourism; fisheries; environment",
        "lifecycle": "Distribution; end-of-life",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/06/Regulamento-Zonas-costeiras.pdf",
    },
    {
        "title_en": "National Maritime Spatial Plan – POEM (Resolution No. 63/2024)",
        "title_pt": "Plano Nacional de Ordenamento do Espaço Marítimo – POEM (Resolução n.º 63/2024)",
        "type": "Plan",
        "year": "2024",
        "description": "Approves national maritime spatial plan to coordinate competing sea uses, minimise environmental impacts, and support sustainable maritime economies.",
        "sectors": "Maritime; fisheries; energy; transport; environment",
        "lifecycle": "All stages (marine planning)",
        "link": "https://sibmoz.gov.mz/content/uploads/2025/11/Resolucao-n.o-63_2024-que-aprova-o-Plano-de-Ordenamento-do-Espaco-Maritimo.pdf",
    },
    {
        "title_en": "Blue Economy Development Strategy – EDEA (Resolution No. 53/2024)",
        "title_pt": "Estratégia de Desenvolvimento da Economia Azul – EDEA (Resolução n.º 53/2024)",
        "type": "Strategy",
        "year": "2024",
        "description": "National blue-economy strategy (2024–2033) promoting sustainable use of marine resources, circular economy, waste valorisation, and targets to reduce plastic production, import, and use.",
        "sectors": "Maritime; fisheries; industry; tourism; environment",
        "lifecycle": "Production; distribution; end-of-life",
        "link": "https://www.proazul.gov.mz/wp-content/uploads/2024/11/EDEA-ESTRATEGIA-DE-DESENVOLVIMENTO-DA-ECONOMIA-AZUL-_FINAL_TYPO_PRINT_WEB.pdf",
    },
    {
        "title_en": "Integrated Coastal Zone Management Strategy (2016–2025)",
        "title_pt": "Estratégia de Gestão Integrada das Zonas Costeiras (2016–2025)",
        "type": "Strategy",
        "year": "2016",
        "description": "Integrated coastal zone management action plan addressing erosion, climate vulnerability, and sustainable use of coastal resources linked to marine pollution pressures.",
        "sectors": "Municipalities; fisheries; tourism; environment",
        "lifecycle": "Distribution; end-of-life",
        "link": None,
        "link_note": "No standalone official PDF identified on publicly accessible government portals (listed on sibmoz.gov.mz, June 2026).",
    },
    {
        "title_en": "National Coral Reef Management and Conservation Strategy (2022–2032; Resolution No. 51/2022)",
        "title_pt": "Estratégia Nacional de Gestão e Conservação dos Corais (2020–2030; Resolução n.º 51/2022)",
        "type": "Strategy",
        "year": "2022",
        "description": "National strategy to protect coral reef ecological integrity and resilience by reducing anthropogenic degradation, including marine litter and pollution pressures.",
        "sectors": "Fisheries; environment; tourism; maritime",
        "lifecycle": "End-of-life (marine leakage)",
        "link": "https://faolex.fao.org/docs/pdf/moz214814.pdf",
        "link_note": "Official Boletim da República PDF (Resolução 51/2022 full text).",
    },
    {
        "title_en": "National Health Policy and Implementation Strategy (Resolution No. 13/2021)",
        "title_pt": "Política de Saúde e Estratégia de Implementação (Resolução n.º 13/2021)",
        "type": "Policy",
        "year": "2021",
        "description": "National health policy including strategies to expand basic sanitation and manage solid waste pressures linked to rapid urbanisation and public health.",
        "sectors": "Health; municipalities; environment",
        "lifecycle": "End-of-life (sanitation/waste)",
        "link": "https://sisma.misau.gov.mz/docs/Politica%20Nacional%20de%20Saude.pdf",
    },
    {
        "title_en": "National Territorial Development Plan (Resolution No. 7/2021)",
        "title_pt": "Plano Nacional de Desenvolvimento Territorial (Resolução n.º 7/2021)",
        "type": "Plan",
        "year": "2021",
        "description": "Approves national territorial development plan and action plan, guiding sustainable land use, urban growth, and environmental balance at national and subnational levels.",
        "sectors": "Urban planning; municipalities; environment; infrastructure",
        "lifecycle": "All stages (spatial planning)",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/06/Plano-Nacional-de-Desenvolvimento-Territorial-1.pdf",
    },
    {
        "title_en": "Territorial Planning Law (Law No. 19/2007)",
        "title_pt": "Lei do Ordenamento do Território (Lei n.º 19/2007)",
        "type": "Law",
        "year": "2007",
        "description": "Territorial planning law promoting rational and sustainable use of natural resources, environmental balance, and quality of life in national development.",
        "sectors": "Urban planning; municipalities; environment; infrastructure",
        "lifecycle": "All stages (planning framework)",
        "link": "https://sibmoz.gov.mz/content/uploads/2022/01/Lei-do-ordenamento-territorial.pdf",
    },
    {
        "title_en": "Territorial Planning Law Regulation (Decree No. 23/2008)",
        "title_pt": "Regulamento da Lei de Ordenamento do Território (Decreto n.º 23/2008)",
        "type": "Regulation",
        "year": "2008",
        "description": "Regulates territorial occupation and sustainable land use through national, provincial, district, and municipal planning instruments.",
        "sectors": "Municipalities; urban planning; environment",
        "lifecycle": "All stages (subnational planning)",
        "link": "https://sibmoz.gov.mz/content/uploads/2025/11/Decreto-n.o-23_2008_Regulamento-da-Lei-de-Ordenamento-do-Territorio.pdf",
    },
]

COLUMNS = [
    "Policy Title (English)",
    "Policy Title (Portuguese)",
    "Policy Type",
    "Last amendment year",
    "Short description",
    "Sectors",
    "Life-Cycle stage",
    "Link",
]


def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(
        url,
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("r:id"), r_id)
    new_run = OxmlElement("w:r")
    r_pr = OxmlElement("w:rPr")
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "0563C1")
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(color)
    r_pr.append(underline)
    new_run.append(r_pr)
    text_elem = OxmlElement("w:t")
    text_elem.text = text
    new_run.append(text_elem)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def build_document():
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.4)
    section.right_margin = Inches(0.4)

    title = doc.add_heading(
        "Mozambique: National Policies Addressing Plastic Pollution (English / Portuguese)",
        level=1,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    subtitle = doc.add_paragraph(
        "National and national-subnational instruments (laws, regulations, policies, strategies, plans, and programmes) "
        "that directly or indirectly address plastic pollution across the life cycle. International treaties excluded. "
        "Links are direct PDF/document URLs only (no news articles). Verified June 2026."
    )
    subtitle.alignment = WD_ALIGN_PARAGRAPH.LEFT

    table = doc.add_table(rows=1, cols=len(COLUMNS))
    table.style = "Table Grid"
    table.autofit = False
    widths = [
        Inches(1.15),
        Inches(1.15),
        Inches(0.65),
        Inches(0.5),
        Inches(1.45),
        Inches(0.85),
        Inches(0.8),
        Inches(1.45),
    ]
    for idx, width in enumerate(widths):
        table.columns[idx].width = width

    hdr_cells = table.rows[0].cells
    for idx, col in enumerate(COLUMNS):
        p = hdr_cells[idx].paragraphs[0]
        p.text = col
        for run in p.runs:
            run.bold = True
            run.font.size = Pt(8)

    for policy in POLICIES:
        row = table.add_row().cells
        link = policy.get("link")
        link_note = policy.get("link_note", "")
        link_display = link if link else link_note
        values = [
            policy["title_en"],
            policy["title_pt"],
            policy["type"],
            policy["year"],
            policy["description"],
            policy["sectors"],
            policy["lifecycle"],
            link_display,
        ]
        for idx, value in enumerate(values):
            cell = row[idx]
            p = cell.paragraphs[0]
            if idx == 7 and link:
                add_hyperlink(p, link, link)
            else:
                p.text = value
            for run in p.runs:
                run.font.size = Pt(7)

    note = doc.add_paragraph(
        "Note: Where a standalone full-text PDF is not published on a .gov.mz portal, the link column states that "
        "no direct official PDF was found. Where indicated, links use official Boletim da República compilations "
        "(gpa.co.mz, faolex.fao.org) hosting the full legal text. PDFs on sibmoz.gov.mz and proazul.gov.mz are "
        "official government publications."
    )
    for run in note.runs:
        run.font.size = Pt(7)

    doc.save(OUTPUT)
    linked = sum(1 for p in POLICIES if p.get("link"))
    print(f"Wrote {OUTPUT} with {len(POLICIES)} policies ({linked} with direct PDF links)")


if __name__ == "__main__":
    build_document()
