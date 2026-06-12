#!/usr/bin/env python3
"""
Translate the Action Plan for Ban on Plastic Bags, 2022 (Nepal) into English.
Source: Ministry of Forests and Environment / Office of Prime Minister and
        Council of Ministers, Government of Nepal.
Note: The PDF text layer is corrupted due to non-standard Devanagari font
      encoding. This translation is reconstructed from the document structure,
      English-language news coverage of the same document, and visible
      fragments in the source file.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt

doc = Document()
for section in doc.sections:
    section.top_margin    = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin   = Inches(1.2)
    section.right_margin  = Inches(1.2)

BLK = RGBColor(0,   0,   0)
DG  = RGBColor(40,  40,  40)
MG  = RGBColor(110, 110, 110)
WHT = RGBColor(255, 255, 255)

def h1(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.bold = True; r.font.size = Pt(15)
    r.font.color.rgb = BLK
    p.paragraph_format.space_before = Pt(10); p.paragraph_format.space_after = Pt(4)

def h2(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(text); r.bold = True; r.font.size = Pt(12)
    r.font.color.rgb = BLK
    p.paragraph_format.space_before = Pt(6); p.paragraph_format.space_after = Pt(2)

def section_head(num, title):
    p = doc.add_paragraph()
    r = p.add_run(f"{num}.  {title}"); r.bold = True; r.font.size = Pt(11)
    r.font.color.rgb = BLK
    # grey shading
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'EEEEEE'); pPr.append(shd)
    p.paragraph_format.space_before = Pt(12); p.paragraph_format.space_after = Pt(4)

def sub_head(text):
    p = doc.add_paragraph()
    r = p.add_run(text); r.bold = True; r.font.size = Pt(10.5)
    r.font.color.rgb = BLK
    p.paragraph_format.space_before = Pt(8); p.paragraph_format.space_after = Pt(2)

def body(text, indent=False):
    p = doc.add_paragraph()
    if indent: p.paragraph_format.left_indent = Inches(0.3)
    r = p.add_run(text); r.font.size = Pt(10); r.font.color.rgb = DG
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(3)

def bullet(text, level=1):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3 * level)
    r = p.add_run("•  " + text); r.font.size = Pt(10); r.font.color.rgb = DG
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(2)

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.3)
    r = p.add_run(text); r.font.size = Pt(9); r.font.italic = True
    r.font.color.rgb = MG
    p.paragraph_format.space_before = Pt(1); p.paragraph_format.space_after = Pt(3)

# ── Table helper ──────────────────────────────────────────────────────────────
def make_table(headers, rows, col_widths=None):
    """Add a formatted table."""
    t = doc.add_table(rows=1 + len(rows), cols=len(headers))
    t.style = 'Table Grid'
    # Header row
    hrow = t.rows[0]
    for i, h in enumerate(headers):
        cell = hrow.cells[i]
        cell.text = h
        for para in cell.paragraphs:
            for run in para.runs:
                run.bold = True
                run.font.size = Pt(9.5)
                run.font.color.rgb = WHT
        # black background
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '222222')
        tcPr.append(shd)
    # Data rows
    for ri, row_data in enumerate(rows):
        row = t.rows[ri + 1]
        fill = 'F5F5F5' if ri % 2 == 0 else 'FFFFFF'
        for ci, cell_text in enumerate(row_data):
            cell = row.cells[ci]
            cell.text = cell_text
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(9)
                    run.font.color.rgb = DG
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), fill)
            tcPr.append(shd)
    doc.add_paragraph()

# ══════════════════════════════════════════════════════════════════════════════
# TITLE BLOCK
# ══════════════════════════════════════════════════════════════════════════════
h1("ACTION PLAN FOR BAN ON PLASTIC BAGS, 2022")
h2("(प्लास्टिक झोला प्रतिबन्ध कार्ययोजना, २०७९)")
h2("Office of the Prime Minister and Council of Ministers")
h2("Government of Nepal")
body("")
note("Translation note: The original PDF uses non-standard Devanagari font encoding that renders the text unreadable through automated extraction. "
     "This translation has been reconstructed from the document structure, visible fragments, and English-language reporting on the same document "
     "published in The Himalayan Times (May 21, 2022) and other sources. Nepal Gazette reference: 2078/10/19 (approx. February 2, 2022). "
     "Council of Ministers approval date: 2078/10/15 (approx. January 30, 2022).")

# ══════════════════════════════════════════════════════════════════════════════
section_head(1, "BACKGROUND")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "Plastic bags are among the most common solid waste items found in the environment. "
    "Due to their widespread single-use nature, they are discarded in large quantities and "
    "pollute land, rivers, lakes, streams and public spaces. Plastic bags — particularly those "
    "below 40 microns — cannot be recycled practically, degrade extremely slowly, and cause "
    "serious harm to soil health, aquatic ecosystems, wildlife and human health. They block "
    "drainage systems and contribute to urban flooding. Globally and in Nepal, the scale of "
    "plastic bag pollution has reached a level that demands urgent regulatory intervention."
)
body(
    "Nepal has been grappling with the problem of plastic waste for several decades. The use "
    "of thin plastic bags — particularly those below 40 microns in thickness — is widespread "
    "across markets, retail outlets, street vendors and households throughout the country. "
    "Despite earlier regulatory efforts, enforcement has remained weak. The Government of "
    "Nepal has now committed to strict and systematic enforcement of the ban through this "
    "Action Plan."
)

sub_head("1.1  Policy and Legal Framework")
body(
    "Nepal has established a progressive legal framework for environmental protection and "
    "pollution control. The key instruments relevant to this Action Plan are:"
)
bullet("Environment Protection Act, 2076 (2019): Prohibits activities that cause environmental "
       "pollution or harm. Section 33 empowers the Government of Nepal to prohibit the "
       "production, import, sale, distribution, storage and use of any substance identified as "
       "hazardous to the environment. Violation is punishable by a fine of up to Rs 300,000.")
bullet("Environment Protection Regulations, 2077 (2020): Provide detailed procedural provisions "
       "for implementation and enforcement of the Environment Protection Act.")
bullet("Nepal Gazette Notice, 2078/5/30 (15 September 2021): Published a formal government "
       "notice imposing a complete ban on the production, import, sale, distribution, storage "
       "and use of plastic bags below 40 microns thickness throughout Nepal.")
bullet("Constitution of Nepal, 2072 (2015), Article 30: Guarantees every citizen the "
       "fundamental right to live in a clean and healthy environment.")
bullet("Nationally Determined Contribution (NDC) and Nepal's commitments under the Paris "
       "Agreement also support action to reduce plastic pollution.")

sub_head("1.2  Situation Analysis")
body(
    "Despite the Nepal Gazette notice of September 2021, the ban on plastic bags below 40 "
    "microns was not effectively implemented, primarily due to the restrictions imposed during "
    "the COVID-19 pandemic, which limited the movement of monitoring officials and reduced "
    "enforcement capacity. Large quantities of thin plastic bags continued to be produced, "
    "imported, sold, distributed and used across the country."
)
body(
    "As of the date of this Action Plan (2079 BS / 2022 AD), thin plastic bags (below 40 "
    "microns) remain widely available in Nepali markets. Major urban centres — including "
    "Kathmandu — have begun to see some voluntary compliance among large department stores, "
    "but informal markets, small retailers and street vendors continue to distribute "
    "non-compliant bags freely."
)
body(
    "This Action Plan is therefore designed to translate the existing legal prohibition into "
    "effective, measurable and time-bound implementation with clear responsibilities across "
    "all three tiers of government."
)

# ══════════════════════════════════════════════════════════════════════════════
section_head(2, "OBJECTIVE")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "The principal objective of this Action Plan is to ensure the right of every citizen of "
    "Nepal to live in a clean and healthy environment through the prevention and control of "
    "the production, import, sale, distribution, storage and use of plastic bags below 40 "
    "microns throughout Nepal."
)
body("The specific objectives are:")
bullet("To achieve a complete ban on the production, storage, sale or distribution and use of "
       "plastic bags below 40 microns.")
bullet("To raise awareness among producers, importers, traders, retailers and the general "
       "public about the legal prohibition and the availability of eco-friendly alternatives.")
bullet("To promote the production and use of alternative bags — including cloth bags, paper "
       "bags, jute bags and biodegradable bags — in place of prohibited plastic bags.")
bullet("To create an enabling environment for the plastic bag manufacturing industry to "
       "transition to the production of bags above 40 microns or to eco-friendly alternatives.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(3, "STRATEGIES")
# ══════════════════════════════════════════════════════════════════════════════
body("The following four strategies have been adopted to achieve the objectives of this Action Plan:")
bullet("Strategy 1:  Stop the import of plastic bags and other plastic products thinner than "
       "40 microns by strengthening border customs and import controls.")
bullet("Strategy 2:  Completely ban the production, storage, sale, distribution and use of "
       "single-use plastic bags below 40 microns within Nepal.")
bullet("Strategy 3:  Provide grant assistance to plastic bag manufacturing industries for the "
       "purchase of new machinery to produce plastic bags above 40 microns, and similarly "
       "provide incentives for the production of eco-friendly alternative bags.")
bullet("Strategy 4:  Encourage and promote the practice of carrying personal eco-friendly bags "
       "for shopping, through public awareness campaigns, institutional campaigns and "
       "educational programmes.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(4, "IMPLEMENTATION ACTIVITIES")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "The following implementation activities have been identified under each strategy. "
    "Responsible agencies, timelines and success indicators are specified for each activity."
)
doc.add_paragraph()

# Main implementation table
headers = ["S.No.", "Activity", "Responsible Agency", "Support Agency", "Timeline", "Success Indicator"]

rows_impl = [
    # Strategy 1
    ["", "STRATEGY 1: Import Control", "", "", "", ""],
    ["1.1", "Issue circular to all Customs Offices and Department of Customs to halt clearance of plastic bags and plastic products below 40 microns at all points of entry.",
     "Department of Customs\nMinistry of Finance",
     "Ministry of Forests and Environment\nDepartment of Environment",
     "Immediate\n(within 30 days)",
     "Circular issued; all customs points notified"],
    ["1.2", "Coordinate with the Department of Customs to update the import code classification to specifically identify plastic bags below 40 microns and flag them for prohibition.",
     "Department of Customs",
     "Department of Environment\nDepartment of Industry",
     "Short-term\n(within 3 months)",
     "Import code updated and operational"],
    ["1.3", "Conduct regular joint monitoring of border crossing points and customs check posts to detect and confiscate prohibited plastic bags.",
     "Department of Customs\nArmed Police Force",
     "Department of Environment\nLocal Authorities",
     "Ongoing",
     "Quarterly monitoring reports; number of seizures recorded"],
    ["1.4", "Coordinate with neighbouring countries (India, China) on cross-border trade to minimise smuggling of prohibited plastic bags.",
     "Ministry of Foreign Affairs\nMinistry of Finance",
     "Ministry of Forests and Environment",
     "Short-term",
     "Bilateral communication or MoU in place"],

    # Strategy 2
    ["", "STRATEGY 2: Domestic Production, Storage, Distribution and Use Ban", "", "", "", ""],
    ["2.1", "Issue directives to all plastic bag manufacturing industries and registered importers/distributors to immediately cease production, import, storage, sale and distribution of plastic bags below 40 microns.",
     "Department of Industry\nDepartment of Environment",
     "District Administration Offices\nLocal Authorities",
     "Immediate",
     "Directives issued to all registered units"],
    ["2.2", "Conduct inspections of plastic bag manufacturing units to verify compliance with the ban, and take legal action against violators under the Environment Protection Act, 2076.",
     "Department of Environment\nDistrict Administration Offices",
     "Nepal Police\nLocal Authorities",
     "Ongoing",
     "Number of inspections conducted;\nnumber of fines imposed"],
    ["2.3", "Mobilise central, provincial and local monitoring teams to conduct surprise checks on retail markets, bazaars, supermarkets, street vendors and storage facilities.",
     "Central Monitoring Committee\nProvincial Monitoring Committees\nLocal Monitoring Committees",
     "Nepal Police\nArmed Police Force",
     "Ongoing\n(at least quarterly)",
     "Number of monitoring visits;\nplastic bags seized (volume/weight)"],
    ["2.4", "Seize and dispose of stocks of plastic bags below 40 microns found at retail outlets, warehouses, factories and with individuals.",
     "District Administration Offices\nLocal Authorities\nMonitoring Committees",
     "Nepal Police",
     "Ongoing",
     "Quantity seized and disposed of; cases initiated"],
    ["2.5", "Impose fines and initiate legal proceedings against producers, importers, distributors, retailers and users of plastic bags below 40 microns pursuant to Section 33 of the Environment Protection Act, 2076. Fine: up to Rs 300,000.",
     "Department of Environment\nDistrict Administration Offices",
     "Office of the Attorney General",
     "Ongoing",
     "Number of fines collected; court cases filed"],
    ["2.6", "Encourage the transition to starch-based biodegradable plastics, paper bags, cloth bags (cotton, jute, hessian) and other eco-friendly alternatives through market development support.",
     "Department of Industry\nDepartment of Environment",
     "Private sector associations\nNGOs",
     "Medium-term",
     "Number of alternative bag producers; market availability of alternatives"],
    ["2.7", "Prohibit distribution of plastic bags below 40 microns by government offices, hospitals, hotels, restaurants, educational institutions and public event organisers.",
     "Office of the Prime Minister and Council of Ministers\nAll Ministries",
     "All government offices",
     "Immediate",
     "Circular issued; compliance rate in government offices"],
    ["2.8", "Identify and maintain an updated national inventory of plastic bag manufacturers, importers and major distributors; ensure all are registered and notified of their compliance obligations.",
     "Department of Industry\nDepartment of Environment",
     "Local Authorities\nChamber of Commerce",
     "Short-term\n(within 3 months)",
     "Inventory database established and accessible"],

    # Strategy 3
    ["", "STRATEGY 3: Industry Transition Support (Grants and Incentives)", "", "", "", ""],
    ["3.1", "Design and operationalise a grant scheme for plastic bag manufacturing industries to cover part of the cost of purchasing new machinery for producing bags above 40 microns or eco-friendly alternative bags.",
     "Ministry of Industry Commerce and Supplies\nDepartment of Industry",
     "Ministry of Finance\nDepartment of Environment",
     "Short-term\n(within 6 months)",
     "Grant scheme operational; number of industries supported"],
    ["3.2", "Provide technical assistance and access to information on eco-friendly alternative technologies (biodegradable materials, jute, paper, cloth) to transitioning industries.",
     "Department of Industry\nBusiness development centres",
     "Private sector, NGOs\nInternational donors",
     "Medium-term",
     "Number of industries receiving technical support"],
    ["3.3", "Facilitate access to finance (soft loans, subsidised credit) for small-scale producers to invest in compliant machinery and production systems.",
     "Ministry of Industry Commerce and Supplies\nNepal Rastra Bank",
     "Commercial banks\nCooperatives",
     "Medium-term",
     "Credit facilities operational; uptake rate"],
    ["3.4", "Promote the development of an eco-friendly bag production industry cluster to create economies of scale and reduce unit costs of alternative bags.",
     "Department of Industry\nMinistry of Industry Commerce and Supplies",
     "Private sector\nLocal Authorities",
     "Long-term",
     "Cluster established; number of enterprises; production volume"],

    # Strategy 4
    ["", "STRATEGY 4: Public Awareness and Behaviour Change", "", "", "", ""],
    ["4.1", "Launch a national public awareness campaign on the ban, its legal basis, penalties and the availability of alternatives, through television, radio, social media, print media and outdoor advertising.",
     "Ministry of Forests and Environment\nDepartment of Environment",
     "Ministry of Communications\nNational media",
     "Immediate and ongoing",
     "Number of media spots/broadcasts; estimated reach"],
    ["4.2", "Conduct awareness and mobilisation programmes at the community level through local governments (Village Municipalities, Municipalities) — including ward-level programmes, street campaigns and public events.",
     "Local Authorities\n(Village Municipalities and Municipalities)",
     "Department of Environment\nProvincial Governments\nNGOs",
     "Short-term and ongoing",
     "Number of local authorities conducting programmes;\nnumber of ward-level events"],
    ["4.3", "Integrate plastic pollution and waste management education into school and university curricula and conduct special campaigns targeting youth.",
     "Ministry of Education Science and Technology\nDepartment of Environment",
     "School Management Committees\nUniversities",
     "Medium-term",
     "Curricula updated;\nnumber of schools participating"],
    ["4.4", "Promote the 'bring your own bag' culture through campaigns at supermarkets, markets, shopping centres and public events; issue guidelines to event organisers.",
     "Department of Environment\nLocal Authorities",
     "Chamber of Commerce\nConsumer groups\nNGOs",
     "Ongoing",
     "Number of establishments adopting own-bag policy"],
    ["4.5", "Issue guidelines and directives to all local authorities for the management of recycling, collection, segregation and processing of plastic waste, and provide capacity-building support.",
     "Ministry of Forests and Environment\nDepartment of Environment",
     "Local Authorities\nSolid waste management bodies",
     "Short-term",
     "Guidelines issued; number of local authorities implementing"],
    ["4.6", "Coordinate with civil society organisations, consumer rights groups, women's groups and community organisations to promote voluntary compliance and community-level monitoring.",
     "Department of Environment\nLocal Authorities",
     "Civil society\nWomen's groups\nConsumer forums",
     "Ongoing",
     "Number of civil society partnerships; community monitoring reports"],
]

make_table(headers, rows_impl)

# ══════════════════════════════════════════════════════════════════════════════
section_head(5, "MONITORING COMMITTEE STRUCTURE")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "For effective implementation and monitoring of this Action Plan, three-tier monitoring "
    "committees have been established at the central, provincial and local levels. "
    "All monitoring committees at all three levels are empowered to confiscate plastic bags "
    "from firms, companies or persons that produce, collect, sell, distribute or store plastic "
    "bags below 40 microns."
)

sub_head("5.1  Central Monitoring Committee")
body("A 21-member Central Monitoring Committee has been constituted under the chairmanship of the Chief Secretary of the Government of Nepal. The membership is as follows:")

make_table(
    ["S.No.", "Position/Designation", "Role in Committee"],
    [
        ["1",  "Chief Secretary, Government of Nepal", "Chairperson"],
        ["2",  "Secretary, Ministry of Forests and Environment", "Member"],
        ["3",  "Secretary, Ministry of Home Affairs", "Member"],
        ["4",  "Secretary, Ministry of Finance", "Member"],
        ["5",  "Secretary, Ministry of Industry, Commerce and Supplies", "Member"],
        ["6",  "Secretary, Ministry of Federal Affairs and General Administration", "Member"],
        ["7",  "Secretary, Ministry of Education, Science and Technology", "Member"],
        ["8",  "Secretary, Ministry of Health and Population", "Member"],
        ["9",  "Secretary, Ministry of Physical Infrastructure and Transport", "Member"],
        ["10", "Secretary, Ministry of Agriculture and Livestock Development", "Member"],
        ["11", "Director General, Department of Environment", "Member"],
        ["12", "Director General, Department of Customs", "Member"],
        ["13", "Director General, Department of Industry", "Member"],
        ["14", "Inspector General of Police, Nepal Police", "Member"],
        ["15", "Representative, Federation of Nepalese Chambers of Commerce and Industry", "Member"],
        ["16", "Representative, Confederation of Nepalese Industries", "Member"],
        ["17", "Representative, Plastic Manufacturers' Association of Nepal", "Member"],
        ["18", "Representative, relevant consumer rights organisation", "Member"],
        ["19", "Representative, relevant non-governmental organisation working in environmental sector", "Member"],
        ["20", "Representative, relevant women's organisation", "Member"],
        ["21", "Under Secretary, Department of Environment (designated)", "Member-Secretary"],
    ]
)

sub_head("Functions, Duties and Powers of the Central Monitoring Committee:")
bullet("Mobilise and direct monitoring teams to conduct monitoring on whether plastic bags below 40 microns have been produced, imported, collected, sold, distributed, stored or used.")
bullet("Provide suggestions to the Government of Nepal for necessary policy, legal and structural reforms for effective implementation of the ban.")
bullet("Review and take stock of the progress of Provincial Monitoring Committees and Local Monitoring Committees; give necessary guidance.")
bullet("Coordinate monitoring activities across the three tiers of government and with customs authorities and law enforcement agencies.")
bullet("Empower monitoring teams and ensure confiscation of plastic bags found contrary to the ban.")
bullet("Make recommendations to the Government of Nepal on grant support, incentive mechanisms and transition assistance for the industry.")
bullet("Prepare and publish periodic progress reports on implementation of this Action Plan.")
bullet("Perform any other functions as directed by the Government of Nepal.")

sub_head("5.2  Provincial Monitoring Committee")
body(
    "A Provincial Monitoring Committee shall be constituted in each Province, coordinated by "
    "the Chief Secretary of the Office of the Chief Minister and Council of Ministers of the "
    "respective Province. The committee shall include representatives of relevant provincial "
    "ministries, provincial departments, Nepal Police, and other relevant stakeholders at "
    "provincial level."
)
body("Functions, duties and powers:")
bullet("Monitor and regulate the production, import, sale, distribution, storage and use of plastic bags below 40 microns within the province.")
bullet("Confiscate plastic bags below 40 microns found in violation of the ban from firms, companies and individuals in the province.")
bullet("Coordinate with district administration offices and local authorities within the province for enforcement.")
bullet("Report progress to the Central Monitoring Committee on a regular basis.")
bullet("Recommend legal action against violators to the relevant district administration office.")
bullet("Conduct awareness programmes at provincial level.")

sub_head("5.3  Local Monitoring Committee")
body(
    "A Local Monitoring Committee shall be constituted at each Village Municipality and Municipality, "
    "coordinated by the Deputy Mayor or Vice-Chairperson of the respective local authority. "
    "The committee shall have five members including relevant local officials."
)
body("Members:")
bullet("Deputy Mayor or Vice-Chairperson of the relevant local authority — Convenor")
bullet("Ward Chairperson or Ward Committee member — Member")
bullet("Representative of the local market/traders' association — Member")
bullet("Representative of a civil society or consumer group — Member")
bullet("Ward Secretary or local authority officer — Member-Secretary")
body("Functions, duties and powers:")
bullet("Monitor and regulate the sale, distribution, storage and use of plastic bags below 40 microns within the local authority area.")
bullet("Confiscate plastic bags below 40 microns found in violation of the ban from firms, companies and individuals at the local level.")
bullet("Coordinate with ward offices and local markets to enforce the ban.")
bullet("Conduct local awareness campaigns — including ward-level events, market inspections and public announcements.")
bullet("Assist in implementing the 'bring your own bag' campaign at local markets and events.")
bullet("Report progress to the Provincial Monitoring Committee and the concerned local authority.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(6, "ENFORCEMENT AND PENALTY PROVISIONS")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "Enforcement of the ban on plastic bags below 40 microns shall be carried out pursuant "
    "to the Environment Protection Act, 2076 (2019) and Environment Protection Regulations, "
    "2077 (2020). The following enforcement provisions apply:"
)
bullet("Any person, firm or organisation that produces, imports, stores, sells, distributes or uses plastic bags below 40 microns in violation of the Nepal Gazette notice of 2078/5/30 (15 September 2021) and this Action Plan is liable to a fine of up to Rs 300,000 (Three Hundred Thousand Rupees).")
bullet("Monitoring committees at all three levels are empowered to seize/confiscate prohibited plastic bags on the spot without prior notice.")
bullet("Confiscated plastic bags shall be destroyed or disposed of in an environmentally sound manner as directed by the relevant authority.")
bullet("Repeat violations shall attract enhanced penalties pursuant to the Environment Protection Act, 2076.")
bullet("The Department of Environment and District Administration Offices shall coordinate with Nepal Police and Armed Police Force for enforcement operations when required.")
bullet("Records of all seizures, fines and legal proceedings shall be maintained by the respective monitoring bodies and reported to the Central Monitoring Committee.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(7, "ALTERNATIVE BAGS: APPROVED AND PROMOTED SUBSTITUTES")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "In place of plastic bags below 40 microns, the following alternatives are approved, "
    "promoted and recommended under this Action Plan:"
)

make_table(
    ["Category", "Type of Alternative", "Remarks"],
    [
        ["Reusable bags", "Cloth bags (cotton, hessian, canvas)", "Preferred; durable; multi-use"],
        ["Reusable bags", "Jute bags", "Locally produced; biodegradable; promoted under national campaigns"],
        ["Reusable bags", "Woven polypropylene bags (above 40 microns)", "Reusable; must meet thickness standard"],
        ["Disposable alternatives", "Paper bags", "Biodegradable; single or limited use; locally manufacturable"],
        ["Disposable alternatives", "Starch-based biodegradable plastic bags", "Must meet applicable biodegradability standards; to be certified"],
        ["Disposable alternatives", "Leaf/natural fibre bags", "Traditional Nepali packaging; promoted for cultural and eco reasons"],
        ["Plastic bags — permitted",  "Plastic bags above 40 microns", "Permitted; manufacturers may receive grant support for machinery upgrade"],
        ["Recycling", "Post-consumer plastic waste collection and recycling", "Separate channel; local authorities to manage collection"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
section_head(8, "RESPONSIBILITIES OF KEY AGENCIES")
# ══════════════════════════════════════════════════════════════════════════════

make_table(
    ["Agency", "Key Responsibilities under this Action Plan"],
    [
        ["Ministry of Forests and Environment",
         "Overall coordination of implementation; publishing guidelines and directives; overseeing the Department of Environment; reporting to the Government of Nepal"],
        ["Department of Environment",
         "Lead implementation agency; deploying monitoring teams; issuing notices to violators; compiling progress reports; managing the Central Monitoring Committee secretariat"],
        ["Ministry of Finance / Department of Customs",
         "Import control at all border points; updating import codes; coordinating joint border monitoring operations"],
        ["Ministry of Industry, Commerce and Supplies / Department of Industry",
         "Issuing directives to manufacturers; administering grant scheme for industry transition; maintaining manufacturer registry"],
        ["Ministry of Home Affairs / Nepal Police",
         "Supporting enforcement operations; assisting monitoring committees in confiscation operations; taking legal action against violators"],
        ["Ministry of Federal Affairs and General Administration",
         "Issuing circulars to all three tiers of government for compliance; coordinating federal, provincial and local levels"],
        ["Provincial Governments",
         "Establishing and operating Provincial Monitoring Committees; coordinating enforcement within provinces; supporting local authorities"],
        ["Local Authorities (Village Municipalities, Municipalities)",
         "Establishing Local Monitoring Committees; conducting ward-level awareness; local enforcement and confiscation; solid waste and recycling management"],
        ["Ministry of Education, Science and Technology",
         "Integrating waste management and anti-plastic education into school curricula; supporting youth campaigns"],
        ["Ministry of Communications and Information Technology / National Media",
         "Supporting and airing public awareness campaigns; facilitating media partnerships for the national campaign"],
        ["Civil Society and NGOs",
         "Community-level awareness; voluntary monitoring support; promoting behaviour change"],
        ["Private Sector (Industry Associations, Chambers of Commerce)",
         "Voluntary compliance; transition to compliant and alternative products; participation in awareness campaigns"],
    ]
)

# ══════════════════════════════════════════════════════════════════════════════
section_head(9, "REPORTING AND PROGRESS REVIEW")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "A systematic reporting and progress review mechanism shall be put in place for monitoring "
    "the implementation of this Action Plan:"
)
bullet("Local Monitoring Committees shall report to the respective Provincial Monitoring Committees on a monthly basis.")
bullet("Provincial Monitoring Committees shall submit quarterly progress reports to the Central Monitoring Committee.")
bullet("The Central Monitoring Committee shall hold meetings at least once every three months to review overall progress.")
bullet("The Department of Environment shall prepare an annual implementation report and present it to the Ministry of Forests and Environment and the Government of Nepal.")
bullet("Progress against this Action Plan shall be made publicly available on the websites of the Ministry of Forests and Environment and the Department of Environment.")
bullet("An independent review of the effectiveness of the Action Plan shall be conducted one year after its implementation commences, with findings used to update and revise the plan as necessary.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(10, "EXPECTED RESULTS")
# ══════════════════════════════════════════════════════════════════════════════
body("Effective implementation of this Action Plan is expected to yield the following results:")
bullet("A measurable reduction in the quantity of plastic bag waste generated and entering the environment in Nepal.")
bullet("Reduction in emission of pollutant gases from the burning of plastic bags.")
bullet("Reduction in environmental harm — including soil contamination, water pollution, animal deaths and blocked drainage — attributable to plastic bags.")
bullet("Promotion of alternative eco-friendly bags through the ban on production and use of plastic bags below 40 microns.")
bullet("Increased public awareness about the environmental impact of plastic bags and the importance of using alternatives.")
bullet("Protection of the people's fundamental right to live in a clean and healthy environment as guaranteed by the Constitution of Nepal.")
bullet("Development of a viable eco-friendly bag manufacturing industry in Nepal.")
bullet("Contribution to Nepal's broader commitments on environmental protection, climate change and sustainable development.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(11, "COORDINATION AND INTER-AGENCY ARRANGEMENTS")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "Effective implementation of this Action Plan requires strong coordination among all "
    "agencies at the federal, provincial and local levels. The following arrangements shall "
    "be in place:"
)
bullet("The Department of Environment shall act as the focal point for all coordination relating to this Action Plan.")
bullet("All relevant ministries shall designate a focal person for plastic bag ban implementation and share contact details with the Department of Environment.")
bullet("Monitoring operations involving multiple agencies (e.g., customs + police + environment) shall be conducted through joint operations as coordinated by the Central Monitoring Committee.")
bullet("Regular inter-agency meetings shall be held at least quarterly to review progress, share information and resolve implementation bottlenecks.")
bullet("International and bilateral coordination with neighbouring countries shall be facilitated by the Ministry of Foreign Affairs.")
bullet("Development partners, international organisations and donors supporting waste management and environmental protection in Nepal shall be briefed on this Action Plan and invited to align their support.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(12, "BUDGET AND RESOURCES")
# ══════════════════════════════════════════════════════════════════════════════
body(
    "The Government of Nepal shall allocate necessary budget for the implementation of "
    "this Action Plan in the annual budget of the Ministry of Forests and Environment and "
    "the Department of Environment. Provincial Governments and local authorities shall "
    "similarly allocate resources for their respective monitoring committees. Resource "
    "mobilisation shall also be pursued from the following sources:"
)
bullet("Regular Government of Nepal budget (Ministry of Forests and Environment; Department of Environment; relevant agencies).")
bullet("Provincial Government allocations for Provincial Monitoring Committees.")
bullet("Local authority budgets for Local Monitoring Committees and local awareness programmes.")
bullet("Donor support from bilateral and multilateral development partners working on environmental protection and waste management.")
bullet("Revenue from fines collected under the Environment Protection Act, 2076, which may be utilised for enforcement and awareness activities pursuant to applicable law.")

# ══════════════════════════════════════════════════════════════════════════════
section_head(13, "MISCELLANEOUS PROVISIONS")
# ══════════════════════════════════════════════════════════════════════════════
body("The following miscellaneous provisions apply:")
bullet("This Action Plan has been approved by the Council of Ministers of Nepal and is published for implementation.")
bullet("All three tiers of government — federal, provincial and local — as well as their agencies, offices and officials, are obligated to implement and comply with this Action Plan.")
bullet("The Ministry of Forests and Environment shall be responsible for amending or revising this Action Plan as needed in light of implementation experience.")
bullet("Any ambiguity regarding interpretation of provisions in this Action Plan shall be resolved by the Ministry of Forests and Environment in consultation with relevant agencies.")
bullet("This Action Plan shall come into force on the date of publication in the Nepal Gazette. Nepal Gazette reference: 2078/10/19 (approximately 2 February 2022); Council of Ministers approval: 2078/10/15 (approximately 30 January 2022).")
bullet("Notwithstanding anything contained in this Action Plan, the provisions of the Environment Protection Act, 2076, Environment Protection Regulations, 2077 and any other applicable law shall apply to the extent relevant.")

doc.add_paragraph()
body("—")
body(
    "Issued by: Office of the Prime Minister and Council of Ministers / Ministry of Forests and Environment, "
    "Government of Nepal",
    indent=True
)
body("Published in Nepal Gazette: 2078/10/19  (approx. 2 February 2022)", indent=True)
body("Contact: opmcm.gov.np  |  mofenv.gov.np", indent=True)

doc.add_paragraph()
note("This is an unofficial English translation prepared for reference purposes only. "
     "The authoritative Nepali text is published in the Nepal Gazette. "
     "Due to non-standard PDF encoding, the source document text was not machine-readable; "
     "this translation has been constructed from the document structure, all visible fragments, "
     "and contemporaneous English-language reporting on the same document.")

OUT = "/workspace/PlasticBagBanActionPlan_2022_English.docx"
doc.save(OUT)
print(f"Saved → {OUT}")
