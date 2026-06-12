#!/usr/bin/env python3
"""
Translate the Solid Waste Management Act, 2068 (Nepal) into English
and produce a formatted Word document.
"""

from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# ── Page margins ──────────────────────────────────────────────────────────────
for section in doc.sections:
    section.top_margin    = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin   = Inches(1.2)
    section.right_margin  = Inches(1.2)

# ── Style helpers ─────────────────────────────────────────────────────────────
def h1(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(15)
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(4)
    return p

def h2(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(12)
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after  = Pt(2)
    return p

def chapter(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.bold = True
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(0, 0, 0)
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after  = Pt(4)
    # light gray shading
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), 'EEEEEE')
    pPr.append(shd)
    return p

def section_head(num, title):
    p = doc.add_paragraph()
    run = p.add_run(f"Section {num}. {title}")
    run.bold = True
    run.font.size = Pt(10.5)
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after  = Pt(2)
    return p

def body(text, indent=False):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(3)
    if indent:
        p.paragraph_format.left_indent = Inches(0.3)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(30, 30, 30)
    return p

def note(text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.4)
    run = p.add_run(text)
    run.font.size = Pt(9.5)
    run.font.italic = True
    run.font.color.rgb = RGBColor(80, 80, 80)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(3)
    return p

def sub(text):
    """Sub-clause / indented body paragraph."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent   = Inches(0.3)
    p.paragraph_format.space_before  = Pt(1)
    p.paragraph_format.space_after   = Pt(3)
    run = p.add_run(text)
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(30, 30, 30)
    return p

def clause(letter, text):
    p = doc.add_paragraph()
    p.paragraph_format.left_indent  = Inches(0.55)
    p.paragraph_format.space_before = Pt(1)
    p.paragraph_format.space_after  = Pt(2)
    run = p.add_run(f"({letter})  {text}")
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(30, 30, 30)
    return p

def divider():
    doc.add_paragraph()


# ══════════════════════════════════════════════════════════════════════════════
# TITLE BLOCK
# ══════════════════════════════════════════════════════════════════════════════
h1("SOLID WASTE MANAGEMENT ACT, 2068 (2011)")
h2("Government of Nepal")
h2("Authenticated and Published: 2068.03.01  (approx. 17 June 2011)")

body("")
body("Amending Acts:")
sub("1.  Some Nepal Acts Amendment Act, 2072 (2015) — 2072.11.13")
sub("2.  Some Nepal Acts Amendment Act, 2075 (2018) — 2075.11.19")

body("")
body("Act No. 5 of the Constitution 2068")

h2("AN ACT TO AMEND AND CONSOLIDATE THE LAW RELATING TO SOLID WASTE MANAGEMENT")

body(
    "Preamble:  Whereas it is desirable to amend and consolidate the law relating to solid "
    "waste management as an essential service, to manage and effectively administer waste by "
    "reducing, reusing, processing or disposing of it at source, and to maintain a clean and "
    "healthy environment by minimising the adverse effects on public health and the environment "
    "from solid waste;"
)
body(
    "Now therefore, the Constituent Assembly, acting in its capacity as the Legislature-Parliament "
    "pursuant to Article 83 of the Interim Constitution of Nepal, 2063, has enacted this Act."
)

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 1
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 1 — PRELIMINARY")

section_head(1, "Short Title and Commencement")
sub("(1)  This Act shall be called the \"Solid Waste Management Act, 2068.\"")
sub("(2)  This Act shall come into force immediately.")

section_head(2, "Definitions")
body(
    "Unless the subject or context otherwise requires, in this Act:"
)
clause("a", "\"Industrial waste\" means harmful and polluting waste discharged from industrial establishments.")
clause("b", "\"Industrial establishment\" means any company, industry, firm or other entity established pursuant to prevailing law for the purpose of operating any industry, business or service.")
clause("c", "\"Container\" means any vessel, box, bucket or similar item placed at a designated location for collecting waste, and the term includes any vessel placed for producing compost fertilizer.")
clause("d", "[Repealed]")
clause("e", "\"Transportation\" means the work of transporting collected waste from its production site to a collection centre, from a collection centre to a transfer station, or from a transfer station to a waste management site or a place designated by the local authority.")
clause("f", "\"Prescribed\" or \"as prescribed\" means prescribed in rules made under this Act.")
clause("g", "\"Minimisation\" means reducing the quantity, size or impact of waste by using any technology or measure.")
clause("h", "\"Discharge\" means depositing or removing waste from its production site to a place designated by the local authority.")
clause("i", "\"Council\" means the Solid Waste Management Council formed pursuant to Section 23.")
clause("j", "\"Pollution\" means any activity that directly or indirectly affects the environment due to the combination of solid, liquid or gaseous substances emitted from waste, causing significant environmental deterioration, damage or harm, or causing damage or loss to the beneficial or useful use of the environment.")
clause("k", "\"Affected area\" means the area identified in the initial environmental examination and environmental impact assessment report for a waste management site.")
clause("l", "\"Processing\" means managing waste by changing its form or quality to produce other useful goods, fertiliser, gas, energy or other products.")
clause("m", "\"Processing site\" means a location where work is carried out to process waste to produce fertiliser, gas, energy or other products.")
clause("n", "\"Recycling\" means transforming collected waste into raw materials through technology and developing it into useful goods for reuse.")
clause("o", "\"Waste\" means domestic waste, industrial waste, chemical waste, healthcare waste or hazardous waste, and includes substances that cannot be used immediately, that are discarded or decomposed, or that are discharged causing environmental degradation — including solid, liquid, gas, sludge, smoke, dust, and materials used for electronic and information technology — as well as posters and pamphlets illegally affixed in public places, and other items designated by the Government of Nepal by gazette notification from time to time.")
clause("p", "\"Waste management site (sanitary landfill site)\" means a site designated by the local authority for the disposal or processing of waste.")
clause("q", "\"Waste collection\" means lifting waste from its production site, collecting door to door, sweeping, piling and removing waste from public places, uprooting weeds, and collecting posters and pamphlets illegally affixed in public places.")
clause("r", "\"Waste collection and transportation means\" means the means, equipment or tools used for collecting and transporting waste.")
clause("s", "\"Post-closure management\" means activities carried out to maintain environmental balance in the area after waste disposal at a waste management site has been stopped.")
clause("t", "\"Ministry\" means the Ministry of Federal Affairs and General Administration of the Government of Nepal.")
clause("u", "\"Chemical waste\" means any chemical substances from any source or process, in any form (solid, liquid, dust, sludge, gas), that have adverse effects on human health, living organisms, animals and the environment and that cannot be used as is, as well as other items designated by gazette notification as chemical waste.")
clause("v", "\"Disposal\" means the final discharge and management of waste.")
clause("w", "\"Collection centre\" means a place designated by the local authority for collecting household waste and temporarily storing it until a specified time, and includes waste collectors or waste collection means designated by the local authority for household collection.")
clause("x", "[Repealed]")
clause("y", "\"Community organisation\" means a participatory consumer group, cooperative or non-governmental organisation established pursuant to prevailing law on a non-profit basis for the benefit of the community.")
clause("z", "\"Transfer station\" means a place designated by the local authority for the intermediate holding of collected waste before it is transported to a waste management site for final disposal.")
clause("aa", "\"Local authority\" means the relevant Metropolitan City, Sub-Metropolitan City, Municipality or Village Municipality.")
clause("bb", "\"Healthcare waste\" means harmful waste produced and discharged from hospitals, clinics, pharmacies, medical shops, blood banks, pathological laboratories, veterinary institutions or health research centres.")
clause("cc", "\"Hazardous waste\" means various types of discharged goods, substances and radioactive radiation that degrade the natural environment and cause harm to human and other living beings' health.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 2
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 2 — PROVISIONS RELATING TO WASTE PRODUCTION, COLLECTION, MINIMISATION AND DISCHARGE")

section_head(3, "Responsibility for Waste Management to Lie with Local Authority")
sub("(1)  The responsibility for constructing and operating necessary infrastructure — including transfer stations, landfill sites, processing plants, compost plants, bio-gas plants and other structures required for waste collection, final disposal and processing — shall lie with the local authority.")
sub("(2)  The responsibility for managing waste deposited or kept at a collection centre, transfer station or processing site, or accumulated during cleaning operations, or for using it in any manner, shall lie with the local authority.")
sub("(3)  For the purposes of this Section, any material deposited or kept at a collection centre, transfer station or processing site, or accumulated during cleaning operations, shall be considered waste.")

section_head(4, "Duty to Manage Waste")
sub("(1)  The duty to manage waste pursuant to this Act shall lie with the local authority.")
sub("(2)  Notwithstanding sub-section (1), the duty to process and manage hazardous waste, healthcare waste, chemical waste or industrial waste within prescribed standards shall lie with the person or entity producing such waste.")
sub("(3)  If an industry or health institution requests the local authority to manage residual and other waste after processing its hazardous, healthcare, chemical or industrial waste, or requests use of the local authority's waste management site, the local authority may manage the waste or allow use of the site on payment of a prescribed service fee.")

section_head(5, "Reducing Waste Production")
sub("(1)  Any person, institution or entity conducting any business shall minimise waste generated as far as possible.")
sub("(2)  It shall be the duty of every person, institution or entity to reduce the quantity of waste by making arrangements for the disposal or reuse of waste that can be managed within their own area, discharging only the remaining waste.")
note("Explanation: \"Own area\" means the compound of a private house, the premises of an industrial area, a hospital or health institution, an industrial establishment premises, or any premises of the waste-producing person, institution or entity.")

section_head(6, "Segregation of Waste")
sub("(1)  The local authority shall direct waste to be segregated at source into at least organic and inorganic categories.")
sub("(2)  As directed by the local authority, the duty to segregate waste at source and deliver it to the collection centre shall lie with the person, institution or entity producing such waste. The local authority may provide necessary technology, materials, equipment and containers for this purpose.")

section_head(7, "Discharge of Waste")
sub("(1)  The time, place and manner of waste discharge shall be as determined by the local authority.")
sub("(2)  A person, institution or entity producing hazardous waste or chemical waste shall manage such waste as prescribed.")
sub("(3)  Hazardous waste or chemical waste shall not be discharged at collection centres or transfer stations.")

section_head(8, "Designation of Waste Collection Centres")
sub("(1)  The local authority may designate collection centres and provide necessary containers in each ward or settlement for the systematic collection of waste.")
sub("(2)  When designating collection centres, a location shall be selected as far as possible that is convenient for all residents of the ward or settlement and is environmentally suitable.")
sub("(3)  The time and manner of waste discharge and collection at collection centres shall be as determined by the local authority.")

section_head(9, "Transportation of Waste")
sub("(1)  The responsibility for transporting waste from collection centres to transfer stations or waste management sites shall lie with the local authority or an institution/entity it has arranged.")
sub("(2)  When transporting waste, the prescribed transportation means shall be used. In designating transportation means, account shall be taken of weight, capacity, method, road conditions and potential environmental impacts.")
sub("(3)  When transporting waste, the local authority shall transport separately the waste segregated at source pursuant to Section 6.")

section_head(10, "Waste Minimisation, Reuse and Recycling")
sub("(1)  The local authority shall take necessary measures to promote waste minimisation, reuse and recycling, and may formulate and implement necessary guidelines for effective implementation.")
sub("(2)  The local authority may coordinate with relevant industries to encourage reuse of industrial packaging materials to reduce waste quantity.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 3
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 3 — PROVISIONS RELATING TO TRANSFER STATIONS AND WASTE MANAGEMENT SITES")

section_head(11, "Designation of Transfer Stations")
sub("(1)  The local authority may designate any location as a transfer station to organise waste collected in the initial stages.")
sub("(2)  When designating a transfer station, it shall be located in a manner that does not adversely affect public health and the environment, and necessary measures shall be taken to prevent odour.")

section_head(12, "Waste Management Sites")
sub("(1)  The local authority may designate a waste management site (sanitary landfill site) for managing and permanently disposing of waste collected within its area, subject to prevailing environmental law.")
sub("(2)  If the local authority does not have suitable land, it may lease or purchase suitable land.")
sub("(3)  When designating a waste management site pursuant to this Section, private land may also be included for development and operation as a waste management site, subject to Section 16.")
sub("(4)  If there is a shortage of land, the local authority may request the Ministry to identify and provide a suitable site.")
sub("(5)  Upon receiving such a request, the Ministry shall acquire the land pursuant to prevailing law and provide it to the local authority.")
sub("(6)  Where a single waste management site is suitable for two or more local authorities, the Ministry may designate such site with the written consent of all relevant local authorities.")
sub("(7)  The local authority shall operate waste management sites and carry out post-closure management in accordance with prescribed environmental standards.")
sub("(8)  Post-closure management shall be carried out in accordance with recommendations of the initial environmental examination and environmental impact assessment report.")
sub("(9)  The local authority may, if necessary, declare the waste management site area as an environmentally sensitive area.")
sub("(10) In such a declared area, the local authority may prohibit unauthorised entry of livestock, animals and humans, prohibit extraction of stones, soil and sand, and issue directives and guidelines for environmental protection and proper management.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 4
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 4 — INVOLVEMENT OF PRIVATE AND COMMUNITY SECTOR IN WASTE MANAGEMENT")

section_head(13, "Provisions Relating to Permits")
sub("(1)  No one shall carry out or cause to be carried out any waste management work without obtaining permission from the local authority.")
sub("(2)  Any domestic or foreign company, institution or entity wishing to manage waste shall submit an application to the local authority disclosing:")
clause("a", "Plan for waste management,")
clause("b", "Details of human resources and technology required,")
clause("c", "Other details as prescribed.")
sub("(3)  Upon receipt of an application, the local authority may issue a permit after necessary inquiries.")
sub("(4)  If necessary technology is not available domestically, the local authority may, with Government of Nepal approval, issue a permit to a foreign company on the condition of technology transfer within the agreed period.")
sub("(5)  Other permit provisions shall be as prescribed.")

section_head(14, "Entrusting Waste Management Work to the Private Sector")
sub("(1)  The local authority may entrust waste management to licensed private sector companies or community organisations through competitive bidding (Section 15) or pursuant to private investment law.")
sub("(2)  The following work may be entrusted — to private companies, all or any; to community/NGOs, any one or more:")
clause("a", "Public awareness campaigns for waste minimisation,")
clause("b", "Waste collection,")
clause("c", "Waste transportation,")
clause("d", "Use, reuse, recycling or processing of waste,")
clause("e", "Waste disposal,")
clause("f", "Post-closure management.")

section_head(15, "Entrusting Waste Management through Competition")
sub("(1)  When entrusting waste management to the private sector or community organisations, the local authority shall invite tenders, conduct competition, select a manager and entrust the management.")
sub("(2)  Selection criteria shall include:")
clause("a", "Amount agreed to be paid to the local authority,")
clause("b", "Capacity, capital, technology and human resources for energy/fertiliser production from waste,")
clause("c", "Financial and technical capacity,")
clause("d", "Sustainability of proposed technology and environmental impact minimisation,")
clause("e", "Proposed management fee,")
clause("f", "Royalty approved to be paid to the local authority regarding use, processing or reuse of waste.")
sub("(3)  Other tender provisions shall be pursuant to prevailing law.")
sub("(4)  An entity entrusted with waste management may collect fees pursuant to Section 18, subject to the agreement with the local authority.")
sub("(5)  Other matters relating to non-governmental sector involvement shall be as prescribed.")

section_head(16, "Authority to Grant Approval for Construction and Operation of Waste Management Facilities")
sub("(1)  If the private sector requests approval to construct and operate a waste management site, processing site or other facility, the local authority may grant approval, subject to environmental and other prevailing laws.")
sub("(2)  The local authority shall monitor compliance with prescribed environmental standards by private sector waste management facilities.")
sub("(3)  If non-compliance is found, a deadline shall be set for rectification. If not rectified within that period, the local authority may revoke the approval.")

section_head(17, "Public-Private Partnership in Waste Management")
sub("(1)  The local authority may carry out waste management in partnership with the private sector and community/non-governmental organisations, subject to prevailing law.")
sub("(2)  Notwithstanding sub-section (1), in partnership with community/NGOs, only the following activities may be carried out: public awareness campaigns, waste collection, transportation, post-closure management, park development and beautification.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 5
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 5 — PROVISIONS RELATING TO WASTE MANAGEMENT SERVICE FEES")

section_head(18, "Authority to Collect Service Fees")
sub("(1)  The local authority may levy and collect service fees from persons, institutions or entities for managing their waste.")
sub("(2)  Fees shall be determined by the local authority based on quantity, weight, nature of waste and other prescribed factors.")
sub("(3)  Fees may be collected by the local authority itself or through a designated institution/entity.")
sub("(4)  An entity entrusted with waste management pursuant to Section 15 may also collect service fees, based on its agreement with the local authority. Prescribed categories of underprivileged groups shall receive prescribed exemptions.")
sub("(5)  Revenue from service fees and from private sector involvement shall be kept by the local authority in a separate budget head and spent on waste management, environmental protection and development of waste management site affected areas, subject to prescribed standards.")

section_head(19, "Authority to Suspend or Terminate Services")
sub("(1)  The local authority may suspend or terminate waste management services to service users who do not pay service fees pursuant to Section 18.")
sub("(2)  An entity entrusted with waste management pursuant to Section 15 may also suspend or terminate services of non-paying users, and shall notify the local authority of such action.")
sub("(3)  Where services are suspended or terminated, the house owner shall personally manage waste generated from their home.")
sub("(4)  Services shall be restored upon payment of outstanding fees.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 6
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 6 — PROVISIONS RELATING TO POLLUTION CONTROL AND MONITORING")

section_head(20, "Pollution Control")
sub("(1)  The local authority shall be responsible for managing collected waste in a pollution-free manner, minimising adverse environmental impacts as far as possible.")
sub("(2)  [Repealed]")
sub("(3)  The local authority shall comply with standards prescribed under this Act when discharging and managing collected waste.")

section_head(21, "Monitoring of Waste Management")
sub("(1)  The local authority shall regularly monitor or cause to be monitored waste management and discharge activities.")
sub("(2)  For monitoring purposes, the local authority may prepare and implement a necessary work plan.")
sub("(3)  A person involved in monitoring shall submit a post-monitoring report to the local authority.")
sub("(4)  The local authority shall take necessary measures regarding improvement and implementation of matters identified in the report.")
sub("(5)  [Repealed]")

section_head(22, "Provisions for Economic, Social Development and Environmental Protection of Affected Areas")
sub("(1)  The local authority shall formulate a master plan for economic, social, physical development and environmental protection of waste management site affected areas, and shall formulate and implement programs accordingly.")
sub("(2)  Programs may primarily focus on:")
clause("a", "Road construction, electricity supply, drinking water and sewerage, sanitation and environmental protection,")
clause("b", "Establishment and operation of schools and health institutions,")
clause("c", "Programs for upliftment of economically poor and socially marginalised groups in affected areas.")
sub("(3)  Plans and programs shall be formulated in consultation with the relevant community.")
sub("(4)  Programs shall be implemented with local community participation.")
sub("(5)  Operations shall comply with environmental protection standards under prevailing law.")
sub("(6)  A local-level committee may be formed as prescribed to provide advice on development and environmental protection of seriously affected areas.")
note("Explanation: \"Seriously affected area\" means an area designated by the Government of Nepal by gazette notification.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 7
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 7 — PROVISIONS RELATING TO THE SOLID WASTE MANAGEMENT COUNCIL")

section_head(23, "Formation of the Council")
sub("(1)  A Solid Waste Management Council is hereby formed for the purpose of determining policy to be adopted in relation to waste management.")
sub("(2)  The Council shall consist of the following members:")
clause("a", "Minister, Ministry of Local Development — Chairperson")
clause("b", "Secretary, Ministry of Physical Planning and Construction — Member")
clause("c", "Secretary, Ministry of Industry — Member")
clause("d", "Secretary, Ministry of Environment — Member")
clause("e", "Secretary, Ministry of Health and Population — Member")
clause("f", "Secretary, Ministry of Local Development — Member")
clause("g", "Secretary, Secretariat of National Planning Commission — Member")
clause("h", "Mayor, Kathmandu Metropolitan City — Member")
clause("i", "Five persons nominated by the Council from among municipal mayors, representing five development regions — Member")
clause("j", "Five persons nominated by the Government of Nepal from office bearers of associations/federations related to local authorities, including at least three women — Member")
clause("k", "Two persons nominated by the Council from areas designated by the Government of Nepal as most waste-affected areas, including at least one woman — Member")
clause("l", "Representative, Federation of Nepalese Chambers of Commerce and Industry — Member")
clause("m", "Two persons nominated by the Council from among waste management experts and scientists, including at least one woman — Member")
clause("n", "One representative of an organisation nominated by the Council from among community organisations working in waste management — Member")
clause("o", "Under-Secretary designated by the Ministry — Member-Secretary")
sub("(3)  The term of office of members nominated pursuant to clause (m) shall be four years, renewable.")
sub("(4)  The term of office of members nominated pursuant to clauses (i), (k) and (n) shall be one year and shall not be renewed.")
sub("(5)  [Repealed]")

section_head(24, "Functions, Duties and Powers of the Council")
body("The functions, duties and powers of the Council shall be:")
clause("a", "Formulate national policy relating to waste management and present it to the Government of Nepal for approval,")
clause("b", "Make policy arrangements for coordination among entities related to waste management,")
clause("c", "Approve standards for determining service fees to maintain uniformity when local authorities set fees,")
clause("d", "Determine the percentage of investment to be made by local authorities when investing in integrated waste management,")
clause("e", "Perform other prescribed functions.")

section_head(25, "Meetings and Decisions of the Council")
sub("(1)  The Council shall meet at least once a year at a date, time and place designated by the Chairperson.")
sub("(2)  Meetings shall be chaired by the Chairperson, or in the Chairperson's absence, by a member elected from among those present.")
sub("(3)  More than fifty percent of total members present shall constitute a quorum.")
sub("(4)  Decisions shall generally be by consensus; if not, by majority; in case of a tie, the chair shall cast the deciding vote.")
sub("(5)  The Council may invite domestic or foreign waste management experts or relevant officials as observers.")
sub("(6)  Council decisions shall be authenticated by the Member-Secretary.")
sub("(7)  Other meeting procedures shall be as determined by the Council itself.")

note("(Chapters 26–37 — Chapter 8 — are omitted from the source text.)")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 9
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 9 — OFFENCES AND PENALTIES")

section_head(38, "Offences")
body("Any person who commits any of the following acts shall be deemed to have committed an offence under this Act:")
clause("a", "Discharging waste at a time and place other than that designated by the local authority.")
clause("b", "Unauthorised use of waste kept in containers or at waste collection centres.")
clause("c", "Damaging, causing harm to, removing from its location, or causing any damage to a container or collection centre.")
clause("d", "Carrying out waste management work without obtaining a permit under this Act.")
clause("e", "Violating conditions specified in a permit issued for waste management under this Act.")
clause("f", "Dumping, placing or accumulating any type of hazardous substance at a collection centre, container or waste accumulation point.")
clause("g", "Placing, dumping or accumulating household or premises waste on roads or other public places.")
clause("h", "Contaminating the house or land of another person by seeping leachate (contaminated water from waste) or sewage.")
clause("i", "Parking or refusing to move a parked vehicle at any place during the time designated by the local authority for street cleaning and waste collection.")
clause("j", "Placing, dumping, accumulating or discharging any type of hazardous waste on roads or other public places in a manner that adversely affects public health, other than at places designated by the local authority.")
clause("k", "Inappropriately dumping, placing or discharging chemical waste, industrial waste, healthcare waste or hazardous waste.")
clause("l", "An industrial establishment or health institution inappropriately dumping, placing or discharging hazardous waste generated by it.")
clause("m", "Obstructing, hindering or impeding waste collection, transportation or waste management.")
clause("n", "Obstructing, closing, surrounding or going on strike in relation to waste collection, transportation, final disposal sites or waste management work.")
clause("o", "Producing, selling or distributing any item designated by the Government of Nepal by gazette notification as excessively waste-generating and prohibited.")
clause("p", "Importing chemical pesticides in violation of Section 44 or failing to fulfil the responsibility to destroy expired medicines.")
clause("q", "Discharging unsegregated waste contrary to Section 6.")
clause("r", "Placing, dumping or accumulating dead or slaughtered livestock and their hides, feathers, bones, fish scraps, etc., in public places, roads, lanes or junctions.")

section_head(39, "Penalties")
sub("(1)  For an offence under Section 38(a): fine of up to Rs 5,000 for the first offence; Rs 5,000–10,000 for the second; Rs 15,000 for each subsequent offence; plus cost of waste removal.")
sub("(2)  For offences under Section 38(b) and (i): fine of Rs 500–5,000.")
sub("(3)  For an offence under Section 38(c): fine of Rs 15,000–50,000, plus cost of managing the container or collection centre.")
sub("(4)  For offences under Section 38(d) and (e): fine of Rs 15,000–50,000 and prohibition on carrying out such activity until a permit is obtained.")
sub("(5)  For an offence under Section 38(f): fine of Rs 5,000–15,000; plus recovery of any damages caused.")
sub("(6)  For offences under Section 38(g), (h) and (r): fine of Rs 5,000–15,000.")
sub("(7)  For an offence under Section 38(j): fine of Rs 30,000–50,000.")
sub("(8)  For offences under Section 38(k), (l) and (p): fine of Rs 50,000–100,000; doubled for a repeat offence; and the relevant authority may be written to for revocation of the licence under prevailing law.")
sub("(9)  For offences under Section 38(m) and (n): the Chief District Officer may impose a fine of Rs 10,000–50,000, or imprisonment of 15 days to 3 months, or both.")
sub("(10) For an offence under Section 38(o): the Chief District Officer may impose a fine of Rs 5,000–10,000, or imprisonment of up to 3 months, or both.")
sub("(11) For an offence under Section 38(q): the local authority may impose a fine of Rs 500 per offence.")

section_head(40, "Authority to Freeze Services and Facilities")
sub("(1)  If a person does not pay or refuses to pay prescribed service fees, the local authority may freeze any services and facilities it provides — including electricity and telephone services — and may write to the relevant authority to freeze sale of the person's house and land. The person shall be given written notice beforehand.")
sub("(2)  Upon receiving such written communication, it shall be the duty of the relevant authority to freeze such services and facilities.")

section_head(41, "Government Prosecution")
body("For offences under Section 38(m), (n) and (o), the case shall proceed as a Government prosecution case and shall be deemed included in Schedule 1 of the National Criminal Procedure Code, 2074.")

section_head(42, "Right to Appeal")
body("A person dissatisfied with a penalty order under Section 39 may appeal to the relevant District Court within 25 days of the order.")

# ══════════════════════════════════════════════════════════════════════════════
# CHAPTER 10
# ══════════════════════════════════════════════════════════════════════════════
chapter("CHAPTER 10 — MISCELLANEOUS")

section_head(43, "Management of Healthcare Waste")
sub("(1)  An authority granting permission to establish a health institution shall first verify that adequate waste management arrangements are in place and shall grant permission only if they are.")
sub("(2)  When granting permission, the authority may also prescribe special waste management conditions or standards to be maintained by the health institution.")

section_head(44, "Provisions Relating to Chemical Pesticides")
sub("(1)  Import of chemical pesticides shall be carried out subject to prescribed standards.")
sub("(2)  The responsibility for destroying expired chemical pesticides subject to prescribed standards shall lie with the relevant person or institution.")

section_head(45, "Maintenance of Records of Community Sector Entities")
sub("(1)  Each local authority may maintain updated records of community sector entities working in waste management within its area.")
sub("(2)  Records may include the area and nature of work, human resources, financial and technical sources, and other prescribed details.")

section_head(46, "Requirement to Provide Information")
body("Any person operating a waste management programme with assistance from a foreign person, association, institution or donor shall inform the Ministry through the relevant local authority.")

section_head(47, "Delegation of Authority")
sub("(1)  Except for matters requiring policy decisions, the Council may delegate some of its powers to the Chairperson or members as necessary.")
sub("(2)–(3)  [Repealed]")

section_head(48, "Contact with the Government of Nepal")
body("The Council shall maintain contact with the Government of Nepal through the Ministry of Local Development.")

section_head(49, "Authority to Grant Awards")
sub("(1)  The Ministry may give appropriate awards to individuals or institutions that develop innovative concepts for waste management, promote such work or contribute to waste management.")
sub("(2)  A local authority may give recognition and cash awards to a person who provides evidence-based complaints against persons who inappropriately dispose of waste in violation of this Act.")

section_head(50, "Authority to Make Rules and Regulations")
sub("(1)  The Government of Nepal may formulate and implement necessary rules to give effect to this Act.")
sub("(2)  [Repealed]")

section_head(51, "Authority to Issue Standards or Guidelines")
sub("(1)  Subject to this Act and rules made thereunder, the local authority may issue separate standards for management of different types of waste.")
sub("(2)  The relevant local authority may formulate and implement necessary guidelines on waste management.")

section_head(52, "Applicability of Prevailing Law")
body("On matters relating to waste management mentioned in this Act, this Act shall apply; on other matters, prevailing law shall apply.")

section_head("52A", "Transfer of Property and Liabilities to the Government of Nepal")
body("After the dissolution of the Waste Management Centre, all its property and liabilities shall transfer to the Government of Nepal.")

section_head(53, "Repeal and Saving")
sub("(1)  The Solid Waste (Management and Resource Mobilisation) Act, 2044 is hereby repealed.")
sub("(2)  Actions and proceedings done pursuant to the repealed Act shall be deemed to have been done pursuant to this Act.")
sub("(3)  All movable and immovable property, staff and liabilities of the Solid Waste Management and Resource Mobilisation Centre formed under the repealed Act shall transfer to the centre under this Act.")

divider()
chapter("REMARKS")
body("(1)  The reference to \"Schedule 1 of the State Cases Act, 2049\" has been replaced by \"Schedule 1 of the National Criminal Procedure Code, 2074\" by subsequent amendment.")
body("(2)  The former Ministry of Local Development is now the Ministry of Federal Affairs and General Administration.")
body("(3)  The former Ministry of Physical Planning and Construction is now the Ministry of Physical Infrastructure and Transport.")
body("(4)  The former Ministry of Industry is now the Ministry of Industry, Commerce and Supplies.")
body("(5)  The former Ministry of Environment is now the Ministry of Forests and Environment.")

divider()
note("Translation prepared from the official Nepali text published at lawcommission.gov.np / faolex.fao.org (FAO LEX reference: nep137767). "
     "This is an unofficial translation for reference purposes only.")

# ── Save ──────────────────────────────────────────────────────────────────────
OUT = "/workspace/SolidWasteManagementAct_2068_English.docx"
doc.save(OUT)
print(f"Saved → {OUT}")
