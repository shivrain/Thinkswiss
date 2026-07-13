#!/usr/bin/env python3
"""Generate Senegal plastic pollution policies Word document."""

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

OUTPUT = "/workspace/Senegal_Plastic_Pollution_Policies.docx"

# Sorted chronologically by enactment / last amendment year
POLICIES = [
    {
        "year": 2001,
        "title_fr": "Loi n° 2001-01 du 12 avril 2001 portant Code de l'Environnement",
        "title_en": "Law No. 2001-01 of 12 April 2001 establishing the Environmental Code",
        "type": "Law (Code)",
        "amendment": "2001",
        "description": "Foundational environmental code establishing pollution control, EIA requirements, waste and nuisance prevention, and State/local authority responsibilities across environmental media.",
        "sectors": "Environment; industry; agriculture; local government",
        "lifecycle": "Production; use; end-of-life; cross-cutting",
        "link": "https://www.denv.gouv.sn/telechargement/69/codes/18921/code-de-lenvironnement-2001.pdf",
    },
    {
        "year": 2001,
        "title_fr": "Décret n° 2001-282 du 12 avril 2001 portant application du Code de l'Environnement",
        "title_en": "Decree No. 2001-282 of 12 April 2001 implementing the Environmental Code",
        "type": "Decree",
        "amendment": "2001",
        "description": "Main implementing decree for the Environmental Code, notably regulating atmospheric emissions from fixed installations and vehicles, special protection zones, and pollution taxes.",
        "sectors": "Environment; industry; transport",
        "lifecycle": "Production; use; end-of-life",
        "link": "https://www.denv.gouv.sn/reglementations-et-normes-cgqa/",
    },
    {
        "year": 2001,
        "title_fr": "Norme sénégalaise NS 05-061 — Rejets d'eaux usées (juillet 2001)",
        "title_en": "Senegalese Standard NS 05-061 — Wastewater discharges (July 2001)",
        "type": "Technical standard",
        "amendment": "2001",
        "description": "Sets threshold values for industrial and municipal wastewater discharges; applied in pollution control and effluent compliance monitoring (including plastic-related industrial effluents).",
        "sectors": "Industry; sanitation; environment",
        "lifecycle": "Production; end-of-life",
        "link": "https://www.denv.gouv.sn/telechargement/74/normes-cadre-juridique/19105/norme-rejets-ns-05-061-eaux-usees-juillet-2001.pdf",
    },
    {
        "year": 2001,
        "title_fr": "Arrêtés ministériels du 28 novembre 2001 relatifs aux études d'impact environnemental (n° 9468–9472 MJEHP-DEEC)",
        "title_en": "Ministerial orders of 28 November 2001 on environmental impact assessment (Nos. 9468–9472 MJEHP-DEEC)",
        "type": "Ministerial orders (arrêtés)",
        "amendment": "2001",
        "description": "Regulatory package governing public participation, technical committee organisation, EIA consultant accreditation, terms of reference, and EIA report content for projects with environmental implications including plastic manufacturing and waste facilities.",
        "sectors": "Environment; industry; planning",
        "lifecycle": "Production; planning",
        "link": "https://www.denv.gouv.sn/decrets/",
    },
    {
        "year": 2003,
        "title_fr": "Norme sénégalaise NS 05-062 — Pollution atmosphérique (octobre 2003)",
        "title_en": "Senegalese Standard NS 05-062 — Atmospheric pollution (October 2003)",
        "type": "Technical standard",
        "amendment": "2003",
        "description": "Defines emission limits for atmospheric pollutants from industrial sources; used with NS 05-061 in industrial pollution control relevant to plastic production and waste treatment.",
        "sectors": "Industry; environment; energy",
        "lifecycle": "Production; end-of-life",
        "link": "https://www.denv.gouv.sn/telechargement/74/normes-cadre-juridique/19106/norme-rejets-ns05-062-pollution-atmospherique-octobre-2003.pdf",
    },
    {
        "year": 2007,
        "title_fr": "Arrêté interministériel n° 09311 du 5 octobre 2007 portant gestion des huiles usagées",
        "title_en": "Interministerial Order No. 09311 of 5 October 2007 on used-oil management",
        "type": "Interministerial order (arrêté)",
        "amendment": "2007",
        "description": "Regulates collection, storage, transport, treatment and disposal of used oils, reducing hazardous waste streams that often co-occur with plastic waste in industrial and municipal settings.",
        "sectors": "Industry; transport; environment",
        "lifecycle": "Use; collection; disposal",
        "link": "https://www.denv.gouv.sn/telechargement/73/arretes-cadre-juridique/19096/arrete-portant-gestion-des-huiles-usees.pdf",
    },
    {
        "year": 2008,
        "title_fr": "Décret n° 2008-1007 du 18 août 2008 portant réglementation de la gestion des déchets biomédicaux",
        "title_en": "Decree No. 2008-1007 of 18 August 2008 regulating biomedical waste management",
        "type": "Decree",
        "amendment": "2008",
        "description": "Establishes classification, segregation, packaging, storage, transport, treatment and disposal rules for biomedical waste, including plastic-containing healthcare waste.",
        "sectors": "Health; environment; local government",
        "lifecycle": "Use; collection; disposal",
        "link": "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/19089/decret-portant-reglementation-de-la-gestion-des-dechets-biomedicaux.pdf",
    },
    {
        "year": 2009,
        "title_fr": "Loi n° 2009-24 du 8 juillet 2009 portant Code de l'Assainissement",
        "title_en": "Law No. 2009-24 of 8 July 2009 establishing the Sanitation Code",
        "type": "Law (Code)",
        "amendment": "2009",
        "description": "Unified sanitation code governing wastewater, stormwater and excreta management; explicitly prohibits discharge of household waste and plastic waste into public sewers and sets planning obligations for communes and rural communities.",
        "sectors": "Sanitation; urban planning; local government",
        "lifecycle": "Use; collection; disposal",
        "link": "https://www.sante.gouv.sn/sites/default/files/3.%20Code_assainissement.pdf",
    },
    {
        "year": 2009,
        "title_fr": "Arrêté n° 07022 du 16 juillet 2009 portant organisation et fonctionnement du Plan national de lutte contre la pollution marine (POLMAR)",
        "title_en": "Order No. 07022 of 16 July 2009 on organisation and operation of the National Marine Pollution Response Plan (POLMAR)",
        "type": "Prime Ministerial order / national plan",
        "amendment": "2009",
        "description": "Organises Senegal's national marine pollution preparedness and response system, including oil and hazardous substance spills and coastal litter affecting marine ecosystems.",
        "sectors": "Fisheries; maritime transport; environment; coast",
        "lifecycle": "Use; disposal; marine environment",
        "link": "https://www.hassmar.gouv.sn/sites/default/files/reglementations/PLAN%20POLMAR.pdf",
    },
    {
        "year": 2010,
        "title_fr": "Décret n° 2010-1281 du 16 septembre 2010 réglementant l'exploitation du plomb issu des batteries usagées",
        "title_en": "Decree No. 2010-1281 of 16 September 2010 regulating recovery of lead from used batteries",
        "type": "Decree",
        "amendment": "2010",
        "description": "Regulates collection and processing of used lead-acid batteries and mercury use, addressing hazardous components often combined with plastic casings in waste electrical equipment.",
        "sectors": "Industry; waste management; environment",
        "lifecycle": "Use; collection; recycling; disposal",
        "link": "https://www.denv.gouv.sn/telechargement/72/decrets-cadre-juridique/19090/decret-num-3-2010-1281-du-16-septembre-2010-reglementant-les-conditions-dexploitation-du-plomb-issu-des-batteries-usagees.pdf",
    },
    {
        "year": 2013,
        "title_fr": "Loi n° 2013-10 du 28 décembre 2013 portant Code Général des Collectivités Territoriales (CGCT)",
        "title_en": "Law No. 2013-10 of 28 December 2013 establishing the General Code of Local Authorities (CGCT)",
        "type": "Law (Code)",
        "amendment": "2013",
        "description": "Defines competences of communes, departments, regions and other local authorities, including municipal cleanliness, household waste management and hygiene—key subnational framework for solid waste including plastics.",
        "sectors": "Local government; sanitation; environment",
        "lifecycle": "Collection; disposal; governance",
        "link": "https://www.dri.gouv.sn/sites/default/files/an-documents/LOI%20N%202013%2010%20DU%2028%20DECEMBRE%202013.pdf",
    },
    {
        "year": 2015,
        "title_fr": "Loi n° 2015-09 du 4 mai 2015 relative à l'interdiction des sachets plastiques de faible micronnage (abrogée)",
        "title_en": "Law No. 2015-09 of 4 May 2015 on prohibition of thin plastic bags (repealed)",
        "type": "Law",
        "amendment": "2015 (abrogated 2020)",
        "description": "First dedicated national plastic-bag law banning production, import, possession, distribution and use of thin plastic bags and requiring rational plastic-waste management. Repealed and replaced by Loi n° 2020-04.",
        "sectors": "Retail; manufacturing; environment; local government",
        "lifecycle": "Production; distribution; use; end-of-life",
        "link": "https://www.dri.gouv.sn/loi-n%C2%B0201509-du-4-mai-2015",
    },
    {
        "year": 2015,
        "title_fr": "Loi n° 2015-18 portant Code de la Pêche maritime",
        "title_en": "Law No. 2015-18 establishing the Maritime Fisheries Code",
        "type": "Law (Code)",
        "amendment": "2015",
        "description": "Maritime fisheries code including provisions on protection of aquatic environments and control of gear and waste affecting marine ecosystems, relevant to fishing gear and marine plastic litter.",
        "sectors": "Fisheries; maritime; environment",
        "lifecycle": "Use; disposal; marine environment",
        "link": "https://www.ditp.gouv.sn/download/file/fid/53",
    },
    {
        "year": 2016,
        "title_fr": "Décret n° 2016-1804 portant application du Code de la Pêche maritime",
        "title_en": "Decree No. 2016-1804 implementing the Maritime Fisheries Code",
        "type": "Decree",
        "amendment": "2016",
        "description": "Application decree for the 2015 Fisheries Code, detailing implementation measures for fisheries management and environmental protection at sea.",
        "sectors": "Fisheries; maritime; environment",
        "lifecycle": "Use; disposal; marine environment",
        "link": "https://www.ditp.gouv.sn/download/file/fid/76",
    },
    {
        "year": 2016,
        "title_fr": "Lettre de politique du secteur de l'environnement et du développement durable 2016–2020",
        "title_en": "Policy letter for the environment and sustainable development sector 2016–2020",
        "type": "Sector policy letter / strategy",
        "amendment": "2016",
        "description": "Government policy orientation for the environment sector covering pollution control, waste management, environmental governance and sustainable development priorities for 2016–2020.",
        "sectors": "Environment; sustainable development; cross-sectoral",
        "lifecycle": "Cross-cutting",
        "link": "https://www.denv.gouv.sn/telechargement/61/documentation/18323/lettre-de-politique-du-secteur-de-lenvironnement-et-du-developpement-durable-2016-2020.pdf",
    },
    {
        "year": 2019,
        "title_fr": "Loi n° 2019-12 du 8 juillet 2019 portant modification du Code Général des Collectivités Territoriales",
        "title_en": "Law No. 2019-12 of 8 July 2019 amending the General Code of Local Authorities",
        "type": "Amending law",
        "amendment": "2019",
        "description": "Amends the CGCT to strengthen local governance frameworks, including provisions affecting how communes and intercommunal structures organise public services such as waste management.",
        "sectors": "Local government; sanitation",
        "lifecycle": "Collection; disposal; governance",
        "link": "https://www.dri.gouv.sn/sites/default/files/LOI/LOI%202019/L-2019-12.pdf",
    },
    {
        "year": 2019,
        "title_fr": "Plan de gestion des déchets biomédicaux (réactualisation, mai 2019)",
        "title_en": "Biomedical waste management plan (updated, May 2019)",
        "type": "National plan / program",
        "amendment": "2019",
        "description": "Updated national plan for biomedical waste management covering segregation, treatment infrastructure, private-sector incinerators and hospital waste-management procedures, including plastic-containing clinical waste.",
        "sectors": "Health; environment; private sector",
        "lifecycle": "Use; collection; disposal",
        "link": "https://www.sante.gouv.sn/sites/default/files/plan_gestion_dechets_biom%C3%A9dicaux_0.pdf",
    },
    {
        "year": 2020,
        "title_fr": "Loi n° 2020-04 du 8 janvier 2020 relative à la prévention et à la réduction de l'incidence sur l'environnement des produits plastiques",
        "title_en": "Law No. 2020-04 of 8 January 2020 on prevention and reduction of environmental impacts of plastic products",
        "type": "Law",
        "amendment": "2020",
        "description": "Current flagship plastic law: bans single-use/disposable plastics and checkout bags; establishes bottle deposit-return; extended producer responsibility; recycled-content targets; plastic-waste import/export controls; and plastic tax on non-recyclable products. Replaces Loi n° 2015-09.",
        "sectors": "Manufacturing; retail; waste management; environment",
        "lifecycle": "Production; distribution; use; collection; recycling; disposal",
        "link": "https://www.dri.gouv.sn/les-actes-l%C3%A9gislatifs",
    },
    {
        "year": 2020,
        "title_fr": "Programme national « Zéro déchet »",
        "title_en": "National « Zero Waste » Programme",
        "type": "Government program",
        "amendment": "2020",
        "description": "Presidential priority programme to improve public cleanliness and promote waste reduction and proper waste management nationwide, implemented through a dedicated management unit under the Ministry of Urban Renewal, Housing and Living Environment.",
        "sectors": "Urban hygiene; local government; environment",
        "lifecycle": "Collection; recycling; disposal; awareness",
        "link": "https://urbanisme.gouv.sn/realisations/l%E2%80%99unit%C3%A9-de-gestion-du-programme-%C2%AB-z%C3%A9ro-d%C3%A9chet-%C2%BB",
    },
    {
        "year": 2021,
        "title_fr": "Projet de Promotion de la Gestion intégrée et de l'Économie des Déchets solides (PROMOGED)",
        "title_en": "Solid Waste Integrated Management and Economy Promotion Project (PROMOGED)",
        "type": "National program / project",
        "amendment": "2021",
        "description": "World Bank–supported national project to strengthen solid-waste governance, modernise collection and treatment infrastructure, and develop waste valorisation value chains across Senegalese municipalities.",
        "sectors": "Urban hygiene; waste management; local government; private sector",
        "lifecycle": "Collection; recycling; disposal; valorisation",
        "link": "https://www.urbanisme.gouv.sn/actualites/r%C3%A9union-du-comit%C3%A9-de-pilotage-du-promoged",
    },
    {
        "year": 2022,
        "title_fr": "Loi n° 2022-18 du 23 mai 2022 autorisant la création de la Société nationale de gestion intégrée des déchets (SONAGED S.A.)",
        "title_en": "Law No. 2022-18 of 23 May 2022 authorising creation of the National Integrated Waste Management Company (SONAGED S.A.)",
        "type": "Law",
        "amendment": "2022",
        "description": "Creates SONAGED S.A. as the State-owned company replacing the UCG to coordinate, collect, treat and valorise solid waste nationally, integrating PROMOGED and other public waste programmes.",
        "sectors": "Waste management; local government; private sector",
        "lifecycle": "Collection; recycling; disposal; valorisation; governance",
        "link": "https://www.urbanisme.gouv.sn/actualites/gestion-des-ordures-sonaged-sa-prend-le-relai-de-l%E2%80%99ucg",
    },
]


def set_cell_margins(cell, top=50, start=80, bottom=50, end=80):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = OxmlElement("w:tcMar")
    for side, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = OxmlElement(f"w:{side}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")
        tc_mar.append(node)
    tc_pr.append(tc_mar)


def add_hyperlink(paragraph, url, text, color="0563C1"):
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
    color_el = OxmlElement("w:color")
    color_el.set(qn("w:val"), color)
    r_pr.append(color_el)
    underline = OxmlElement("w:u")
    underline.set(qn("w:val"), "single")
    r_pr.append(underline)
    new_run.append(r_pr)
    text_el = OxmlElement("w:t")
    text_el.text = text
    new_run.append(text_el)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


def add_wrapped_text(cell, text, bold=False, size=9):
    p = cell.paragraphs[0]
    p.clear()
    run = p.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold


def build_document():
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Inches(0.45)
    section.right_margin = Inches(0.45)
    section.top_margin = Inches(0.5)
    section.bottom_margin = Inches(0.5)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("Senegal — National Policies Addressing Plastic Pollution")
    run.bold = True
    run.font.size = Pt(16)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle.add_run(
        "République du Sénégal — Politiques nationales liées à la pollution plastique (cycle de vie)\n"
        "Compiled from official .gouv.sn sources — sorted chronologically — verified July 2026"
    )
    sub_run.font.size = Pt(10)
    sub_run.italic = True

    note = doc.add_paragraph()
    note_run = note.add_run(
        "Note: International treaties are excluded per scope. Loi n° 2020-04 is in force but its full text "
        "was not published as a downloadable PDF on dri.gouv.sn at verification; the official DRI legislative "
        "portal is provided. Loi n° 2015-09 is listed as repealed (abrogée) by Loi n° 2020-04."
    )
    note_run.font.size = Pt(9)

    headers = [
        "Policy Title (FR / EN)",
        "Policy Type",
        "Last amendment year",
        "Short description",
        "Sectors",
        "Life-cycle stage",
        "Official link",
    ]

    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    widths = [Inches(1.55), Inches(0.85), Inches(0.55), Inches(1.55), Inches(0.95), Inches(0.95), Inches(1.35)]
    for i, width in enumerate(widths):
        table.columns[i].width = width

    hdr_cells = table.rows[0].cells
    for i, header in enumerate(headers):
        add_wrapped_text(hdr_cells[i], header, bold=True, size=9)
        set_cell_margins(hdr_cells[i])

    for policy in POLICIES:
        row = table.add_row().cells
        title_text = f"{policy['title_fr']}\n{policy['title_en']}"
        add_wrapped_text(row[0], title_text, size=8)
        add_wrapped_text(row[1], policy["type"], size=8)
        add_wrapped_text(row[2], policy["amendment"], size=8)
        add_wrapped_text(row[3], policy["description"], size=8)
        add_wrapped_text(row[4], policy["sectors"], size=8)
        add_wrapped_text(row[5], policy["lifecycle"], size=8)

        link_cell = row[6]
        link_cell.text = ""
        p = link_cell.paragraphs[0]
        display = policy["link"].replace("https://", "")
        if len(display) > 70:
            display = display[:67] + "..."
        add_hyperlink(p, policy["link"], display)

        for cell in row:
            set_cell_margins(cell)

    doc.add_paragraph()
    footer = doc.add_paragraph(
        "Source verification: all links were checked twice (HTTP availability and content match) against official "
        "Government of Senegal websites (*.gouv.sn) in July 2026."
    )
    footer.runs[0].font.size = Pt(8)
    footer.runs[0].italic = True

    doc.save(OUTPUT)
    print(f"Saved {OUTPUT}")


if __name__ == "__main__":
    build_document()
