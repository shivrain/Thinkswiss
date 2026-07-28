# 4P Index Coding — Environment Protection Regulation, 2077 (2020), Nepal

## Provenance / methodology

- **Source page:** https://dpnet.org.np/resource-detail/1821
- **Source document (direct PDF link found on that page):** `https://dpnet.org.np/uploads/files/Env_Regulation_rajpatra_1592360474-61778-(1)-1660721565-1665478236-compressed 2024-06-14 05-41-41.pdf`
  A local copy is archived at [`sources/environment-protection-regulation-2020-nepal-source.pdf`](sources/environment-protection-regulation-2020-nepal-source.pdf).
- **Original language / title:** Nepali — वातावरण संरक्षण नियमावली, २०७७ ("Environment Protection Regulation, 2077"), published in the *Nepal Gazette* (नेपाल राजपत्र), Part 3, Volume 70, No. 9, dated 2077/03/01 B.S. (≈ 15 June 2020 CE). It is a scanned (image-only) PDF of 167 pages with no embedded text layer.
- **Processing performed for this coding exercise:**
  1. Rendered all 167 pages to images (`pdftoppm`, 200 dpi).
  2. OCR'd every page with Tesseract using the Nepali (`nep`) language model.
  3. Combined the per-page OCR output into a single file with page markers, preserved at [`sources/nepal-environment-protection-regulation-2020-ocr-nepali.txt`](sources/nepal-environment-protection-regulation-2020-ocr-nepali.txt) for auditability (OCR of a scanned Nepali gazette is imperfect in places — where wording was ambiguous, the underlying rule number and page are cited so it can be re-checked against the source PDF).
  4. Read the OCR'd Nepali text chapter by chapter, identified all provisions relevant to plastics per the Plastics Relevance Filter, and translated the operative clauses into English for the coding table below. English quotations in Column F/N/R/U below are **translations of the actual regulation text** (not paraphrases of secondary summaries), with the Nepali rule/schedule citation given alongside each.
- No amendment to this Regulation was found in the source document (no "संशोधन" section appears); the 2020 (2077 B.S.) text is therefore the most recent version coded.

### Document structure (for reference)

| Chapter (परिच्छेद) | Nepali title | English |
|---|---|---|
| 1 | प्रारम्भिक | Preliminary (Rules 1–2) |
| 2 | वातावरणीय अध्ययन | Environmental Study — brief study / IEE / EIA screening (Rules 3–13) |
| 3 | **प्रदूषण नियन्त्रण** | **Pollution Control** (Rules 14–24) — contains the core waste/hazardous-substance provisions |
| 4 | जलवायु परिवर्तन | Climate Change (Rules 25–27) |
| 5 | कार्बन व्यापार | Carbon Trading (Rule 28) |
| 6 | राष्ट्रिय सम्पदाको संरक्षण तथा वातावरण संरक्षण क्षेत्र | National Heritage & Environment Conservation Areas, incl. the Environment Conservation Fund (Rules 29–40) |
| 7 | क्षतिपूर्ति | Compensation (Rules 41–43) |
| 8 | विविध | Miscellaneous (Rules 44–50) |
| — | अनुसूची १–२१ | Schedules 1–21 (forms, and — critically — the sectoral thresholds in Schedules 1–3 that trigger mandatory environmental screening, including explicit plastic-industry and waste-management-facility thresholds) |

Only Chapter 3 (Pollution Control), the Environment Conservation Fund (Chapter 6), and the Schedule 1–3 screening thresholds contain provisions that meet the Plastics Relevance Filter; the rest of the Regulation (EIA procedure generally, climate change, carbon trading, heritage areas, compensation) does not specifically or functionally target plastics/waste and was excluded from the instrument-level coding.

---

## Section A — Policy-level fields (identical across all rows)

| Col | Field | Value |
|---|---|---|
| A | policy_name | Environment Protection Regulation, 2077 (वातावरण संरक्षण नियमावली, २०७७) |
| B | policy_url | https://dpnet.org.np/resource-detail/1821 (full text PDF: see Provenance section above) |
| C | policy_year | 2020 |
| D | policy_objective | To operationalise the Environment Protection Act, 2076 (2019) by (i) mandating environmental screening (brief study / Initial Environmental Examination / Environmental Impact Assessment) for listed categories of development and industrial proposals, including plastic-product manufacturing and waste-management facilities; (ii) establishing pollution-control mechanisms, including a duty on local governments to segregate and manage decomposable, reusable and recyclable waste, cost-recovery for unlawful waste dumping, a licensing regime for "hazardous substances" (import, collection, storage, processing, sale, disposal and transport) and Basel-Convention-consistent controls on export of hazardous substances/waste; (iii) creating a dedicated Environment Conservation Fund to finance pollution-prevention technology and programmes; and (iv) providing compensation for environmental damage. It does not set out a plastics-specific policy objective as such, but functions as the general implementing framework within which plastic production, plastic waste and hazardous plastic-related waste streams are regulated. |
| E | policy_target | 0 |
| F | policy_target_text | *(empty — no plastic-specific quantifiable or non-quantifiable target found in the Regulation; the numeric thresholds in Schedules 1–3, e.g. ">5 metric tons/day of plastic goods", are screening/permitting triggers, not reduction or performance targets)* |
| G | policy_type | 0.75 |
| H | policy_type_justification | Issued as a "नियमावली" (Regulation) by the Government of Nepal (Council of Ministers, acting through the Ministry of Forests and Environment) under the delegated rule-making authority of Section 44 of the Environment Protection Act, 2076 (2019), and published as an executive notice in the Nepal Gazette, Part 3. It was not enacted by the federal Parliament (that status belongs to the parent 2019 Act); it is therefore a sub-legislative executive/ministerial instrument = 0.75. |
| I | policy_integration | 1 |
| J | policy_sectors_list | production, industry, mining, waste management, recycling, agriculture, forestry, tourism, transportation (roads), energy, water resources, irrigation, drinking water, housing & urban development, health, education, municipalities |
| K | policy_circularity | 0.75 |
| L | policy_lifecycle_phases_list | production, recycling, disposal, environmental leakage |
| M | policy_budget | 1 |
| N | policy_budget_text | "वातावरण संरक्षण कोष सञ्चालक समितिको गठन" — Rule 35 establishes an Environment Conservation Fund Operating Committee (chaired by the Secretary of the Ministry, with members from the Ministries of Finance; Industry, Commerce & Supplies; Physical Infrastructure & Transport; Federal Affairs & General Administration; a provincial environment secretary; the Director-General of the Department of Environment; and environment experts). Rule 37: "कोषलाई चक्रीय (रिभल्भिङ्ग) कोषको रूपमा सञ्चालन गरिनेछ" — the Fund shall be operated as a revolving fund, deposited in a "क" class commercial bank. Rule 38 restricts fund use to: (a) implementing specialised environmental-protection, pollution-prevention-and-control and heritage-conservation plans/programmes; (b) using technology for pollution study and reduction; (c) public-awareness activities; (d) environmental education, training, study and research; and (e) rewarding outstanding contributors. Rule 39 requires the Fund's income and expenditure to be accounted for under prevailing law, and Rule 40 requires the Committee to publish an annual report within 30 days of each fiscal year-end. Because a dedicated, ring-fenced Fund with its own governance committee is created and its permissible uses are restricted by rule, this is coded 1 (ring-fenced budget), even though the Fund is not plastics-exclusive. |
| O | policy_score `[auto]` | **0.875** = (G 0.75 + I 1 + K 0.75 + M 1) / 4 |

---

## Section B — Instrument-level fields (one row per instrument)

Eight distinct instruments in the Regulation meet the Plastics Relevance Filter (criteria (b) and/or (c) — none of the provisions name "plastic" explicitly except the Schedule‑2 industrial threshold, so criterion (a) applies only to Instrument 6).

| # | Instrument | P | Q | S | T | V `[auto]` |
|---|---|---|---|---|---|---|
| 1 | Ministry-arranged alternative substances/technologies for pollution control (Rule 14(1), 14(4)) | 0.40 | Production | 0 | 0.25 | **0** |
| 2 | Local-level mandatory waste segregation & management (Rule 14(2)) | 1.00 | Waste management | 1 | 0.50 | **0.75** |
| 3 | Cost recovery + 50% surcharge for unlawful waste discharge (Rule 14(3)) | 0.60 | Waste management | 1 | 0.75 | **0.675** |
| 4 | Hazardous-substance export permit & Basel Convention compliance (Rule 15) | 1.00 | End of life | 1 | 0.50 | **0.75** |
| 5 | Hazardous-substance management licence: import/collection/storage/processing/sale/disposal/transport (Rule 16) | 1.00 | Waste management | 1 | 1.00 | **1.00** |
| 6 | Mandatory environmental screening (IEE/EIA) for plastic-product manufacturing, waste-plastic/tyre-to-fuel processing & waste-management facilities (Rules 3–4; Schedules 1–3) | 0.40 | Production | 1 | 0.75 | **0.575** |
| 7 | Pollution Control Certificate & mandatory effluent/emission sampling and monitoring (Rules 21–22) | 1.00 | Waste management | 1 | 0.75 | **0.875** |
| 8 | Environment Conservation Fund (Rules 35–40) | 0.60 | Waste management | 1 | 0.50 | **0.55** |

### Detailed instrument coding

#### Instrument 1 — Alternative substances/technologies for pollution control

- **Q — instrument_lifecycle_stage:** Production
- **R — instrument_description:** Rule 14(1): "मन्त्रालयले विभागको सिफारिसमा प्रदूषण नियन्त्रण गर्न सघाउ पुऱ्याउने वैकल्पिक पदार्थ, इन्धन, औजार, यन्त्र वा उपकरणको प्रयोगका लागि आवश्यक व्यवस्था मिलाउन सक्नेछ।" — *Translation:* "The Ministry may, on the Department's recommendation, make necessary arrangements for the use of alternative substances, fuel, tools, machinery or equipment that help control pollution." Rule 14(4) additionally allows the Ministry to contract "mountain environment conservators" to control pollution and remove waste in the Himalayan/high-mountain region.
- **S — instrument_in_force:** 0. The language is enabling ("सक्नेछ" = "may"), and no subordinate order/decision operationalising this power was found in the document. Per the in-force test, an unexercised "may" power scores 0.
- **T — instrument_implementation:** 0.25 (responsible authority explicitly named — the Ministry, on Department recommendation — +0.25; no enforcement, monitoring or unconditional language is attached to this discretionary power).
- **U — instrument_implementation_text:** "मन्त्रालयले विभागको सिफारिसमा ... व्यवस्था मिलाउन सक्नेछ" (only the responsible authority is specified; the rest is discretionary).
- **V — instrument_score `[auto]`:** S × (P+T)/2 = 0 × (0.40+0.25)/2 = **0**.
- **W — comments:** Directly matches the "emphasises use of alternative substances and technologies for pollution control" element of the policy summary, but as drafted it is a discretionary enabling power, not a binding requirement — flagged as an area where the Regulation could be strengthened. Alternative substances for "pollution control" would in practice include alternatives to certain plastics/additives, but the rule text itself is substance-agnostic.

#### Instrument 2 — Local-level mandatory waste segregation & management

- **Q:** Waste management
- **R:** Rule 14(2): "स्थानीय तहले प्रदूषणको प्रकृति अनुसार कुहिने, पुनःप्रयोग र पुनर्उत्पादन (रिसाइकल) गर्न सकिने फोहर वर्गीकरण गरी उचित व्यवस्थापन गर्नु पर्नेछ।" — *Translation:* "The local level shall classify waste into decomposable, reusable, and recyclable (रिसाइकल) categories according to the nature of the pollution, and manage it appropriately." This is the clearest waste-segregation/recycling obligation in the Regulation and directly captures plastics as part of the "reusable/recyclable" waste stream.
- **S:** 1 — mandatory language ("पर्नेछ" = "shall").
- **T:** 0.50 — responsible authority explicitly designated (local level) (+0.25); no explicit enforcement/penalty for a local level that fails to segregate waste, and no explicit monitoring mechanism is attached to this specific sub-rule; the obligation itself carries no stated exemptions (+0.25 unconditional).
- **U:** "स्थानीय तहले ... फोहर वर्गीकरण गरी उचित व्यवस्थापन गर्नु पर्नेछ।" (mandatory duty on local level; no named enforcement or monitoring mechanism in the text of this sub-rule).
- **V:** 1 × (1.00+0.50)/2 = **0.75**.
- **W:** This is the single most plastics-relevant instrument in the Regulation for the "consumption/end-of-life" side of the lifecycle, since it obliges Nepal's ~753 local governments to separate recyclable waste (which in practice is dominated by plastics, paper and metal) from organic/residual waste. It is a substantive, mandatory obligation on sub-national government (not a mere coordination mechanism), so it is scored as a Regulatory approach (1.0) rather than the 0.40 "procedural/coordination" tier.

#### Instrument 3 — Cost recovery and 50% surcharge for unlawful waste discharge

- **Q:** Waste management
- **R:** Rule 14(3): "कुनै व्यक्ति, संस्था वा उद्योगले ... तोकिएको शर्त वा मापदण्ड विपरीत फोहरमैला निष्कासन नगर्न सूचना वा निर्देशन गर्दा समेत फोहरमैला निष्कासन गरेको कारणले जनस्वास्थ्य तथा वातावरणमा प्रतिकूल प्रभाव पर्न गएमा सम्बन्धित प्रदेश सरकार वा स्थानीय तहले आफ्नै खर्चमा त्यस्तो फोहरमैला हटाई सो फोहर हटाउँदा लागेको खर्चमा पचास प्रतिशत रकम थप गरी त्यस्तो फोहरमैला निष्कासन गर्ने व्यक्ति, संस्था वा उद्योगसँग प्रचलित कानुन बमोजिम असुल उपर गर्नु पर्नेछ।" — *Translation:* "If a person, institution or industry discharges waste in violation of the prescribed conditions or standards — even after being notified or directed not to — and public health or the environment is adversely affected as a result, the concerned provincial government or local level shall remove such waste at its own expense and shall recover the cost of removal, plus an additional 50 percent, from the discharger under prevailing law."
- **S:** 1 — mandatory ("पर्नेछ").
- **T:** 0.75 — responsible authority named (province/local level) (+0.25); the cost-recovery-plus-50%-surcharge mechanism is itself an enforcement/economic-penalty tool (+0.25); no explicit routine monitoring mechanism is described for detecting the violation itself (no +0.25); the duty to remove and recover is stated without carve-outs once the trigger (adverse effect from unlawful discharge) is met (+0.25 unconditional).
- **U:** as quoted above — cost recovery "पचास प्रतिशत रकम थप गरी ... असुल उपर गर्नु पर्नेछ" (recovery of costs plus 50% is mandatory once triggered).
- **V:** 1 × (0.60+0.75)/2 = **0.675**.
- **W:** This operates as a subsidiary/contingent enforcement tool (it activates only once a discharger has already ignored a notice/direction), but the duty to act, once triggered, is mandatory rather than discretionary, so it is coded in force (S=1) per Rule 12's guidance that mandatory language is in force; the contingent character is noted here rather than reflected in S.

#### Instrument 4 — Hazardous-substance export permit & Basel Convention compliance

- **Q:** End of life
- **R:** Rule 15(1): "ऐनको दफा १६ को उपदफा (३) बमोजिम नेपालबाट जोखिमपूर्ण पदार्थ निर्यातका लागि निर्यातकर्ताले मन्त्रालयबाट अनुमति लिनु पर्नेछ र सोको लागि निजले त्यस्तो पदार्थ आयात गर्ने मुलुकको सरोकारवाला निकायको सहमति समेत पेश गर्नु पर्नेछ।" — *Translation:* "An exporter exporting hazardous substances from Nepal shall obtain permission from the Ministry and shall submit the consent of the relevant authority of the importing country." Rule 15(2): "... जोखिमपूर्ण पदार्थको निर्यात गर्दा जोखिमपूर्ण फोहरको सीमापार ओसारपसार, नियन्त्रण र विसर्जन सम्बन्धी बासेल महासन्धिमा उल्लिखित व्यवस्थाको पालना गर्नु पर्नेछ।" — *Translation:* "When exporting hazardous substances, the provisions of the Basel Convention on the Control of Transboundary Movements of Hazardous Wastes and Their Disposal shall be complied with."
- **S:** 1 — mandatory ("पर्नेछ").
- **T:** 0.50 — authority explicit (Ministry) (+0.25); no explicit domestic enforcement/penalty or monitoring clause within Rule 15 itself (0); no exemptions stated for this obligation (+0.25 unconditional).
- **U:** "निर्यातकर्ताले मन्त्रालयबाट अनुमति लिनु पर्नेछ ... बासेल महासन्धिमा उल्लिखित व्यवस्थाको पालना गर्नु पर्नेछ।"
- **V:** 1 × (1.00+0.50)/2 = **0.75**.
- **W:** This provision incorporates the Basel Convention's transboundary hazardous-waste controls into domestic law — directly relevant to plastic waste exports (post-2021 Basel Convention plastic waste amendments make many plastic waste categories "hazardous"/controlled). The Regulation's own definition of "जोखिमपूर्ण पदार्थ" ("hazardous substance") is broader than "hazardous waste" alone; noted as a translation/scope caveat.

#### Instrument 5 — Hazardous-substance management licence (import, collection, storage, processing, sale, disposal, transport)

- **Q:** Waste management
- **R:** Rule 16: (1) "मन्त्रालयले विभागको सिफारिसमा जोखिमपूर्ण पदार्थको सूचीकरण गर्नेछ।" — "The Ministry shall list hazardous substances on the Department's recommendation." (2) "जोखिमपूर्ण पदार्थको आयात, सङ्कलन, भण्डारण, प्रशोधन, विक्री वितरण, विसर्जन वा ओसारपसार गर्न चाहने व्यक्ति वा संस्थाले अनुसूची-१६ बमोजिमको ढाँचामा विभाग समक्ष निवेदन दिनु पर्नेछ।" — "Any person or institution wishing to import, collect, store, process, sell/distribute, dispose of or transport hazardous substances shall apply to the Department (Schedule-16 form)." (3) The Department shall grant a permit (Schedule-17 form) with necessary conditions if satisfied on inquiry. (4)–(6) Permit is valid 5 years, renewable for a further 2 years. (7) Permit holders must file an annual report (Schedule-18 form) within 30 days of fiscal year-end. (8)–(9) The Department must order compliance where a permit holder breaches conditions or mismanages hazardous substances, and must refer continued non-compliance for legal action.
- **S:** 1 — mandatory throughout ("पर्नेछ").
- **T:** 1.00 — authority explicit (Ministry lists substances; Department licenses, monitors and enforces) (+0.25); enforcement explicit (compliance orders and referral for legal action under sub-rules (8)–(9)) (+0.25); monitoring explicit (mandatory annual reporting by every permit holder, sub-rule (7)) (+0.25); the licensing requirement itself applies to *any* person or institution engaging in these activities with no stated exemption or threshold carve-out (+0.25 unconditional).
- **U:** "अनुसूची-१६ बमोजिमको ढाँचामा विभाग समक्ष निवेदन दिनु पर्नेछ ... प्रत्येक आर्थिक वर्ष समाप्त भएको तीस दिनभित्र विभाग समक्ष पेश गर्नु पर्नेछ ... आदेश दिनु पर्नेछ ... कारबाहीका लागि सम्बन्धित निकायमा लेखी पठाउनु पर्नेछ।"
- **V:** 1 × (1.00+1.00)/2 = **1.00**.
- **W:** This is the instrument that most directly matches the "permits and responsibilities for hazardous waste handling, collection, storage, processing, and disposal" description of the Regulation. Note for accuracy: the Nepali term used throughout ("जोखिमपूर्ण पदार्थ") translates as "hazardous substance/material" rather than literally "hazardous waste" — the licence chain covers substances from import through to disposal ("विसर्जन"), which functionally includes hazardous waste management, but the Regulation does not use a term equivalent to "plastic waste" specifically. Scored as directly relevant under Plastics Relevance Filter criterion (c).

#### Instrument 6 — Mandatory environmental screening for plastic manufacturing & waste-management facilities

- **Q:** Production (also functionally covers Waste management for the waste-facility sub-items; see comments)
- **R:** Rules 3–4 require every "proposal" falling in Schedule 1 to undergo a brief environmental study, every proposal in Schedule 2 to undergo an Initial Environmental Examination (IEE), and every proposal in Schedule 3 to undergo a full Environmental Impact Assessment (EIA), including a mandatory public-scoping process (Rule 4). Schedule 2 ("प्रारम्भिक वातावरणीय परीक्षण गर्नु पर्ने प्रस्ताव"), Industry sector, explicitly lists (item 41): *"Establishing an industry producing more than 5 metric tons per day of plastic goods (including production of pipes, fittings, yarn, etc.), including [industries] based on waste-plastic raw material or flakes"* and (item 43): *"Establishing an industry that processes old tyres and plastic to extract fuel (oil)."* Schedules 2 and 3 separately list a dedicated "Waste management sector" (फोहरमैला व्यवस्थापन क्षेत्र) requiring IEE/EIA for landfilling above stated tonnages, transfer stations, resource-recovery areas, chemical/mechanical/biological waste sorting, composting plants, sewage/waste management above stated population thresholds, and (irrespective of scale) construction of any hazardous-waste plant, recovery plant, landfill/storage/treatment facility.
- **S:** 1 — mandatory ("पर्नेछ") for any proposal falling within the listed Schedule thresholds.
- **T:** 0.75 — responsible authority explicit (Department/Ministry/Province, per Rule 4(5)) (+0.25); no explicit fine/penalty is specified within Rules 3–4 themselves for proceeding without the required study (0); monitoring is explicit — Rule 45 requires proponents to self-monitor every six months and report, and the Ministry/Department/Province to monitor and inspect (+0.25); the screening duty applies to every proposal meeting the listed thresholds with no further exemption stated (+0.25 unconditional).
- **U:** Rule 45(1)–(2): "प्रस्तावकले ... प्रत्येक छ महिनामा स्वःअनुगमन गरी सोको प्रतिवेदन ... पेश गर्नु पर्नेछ ... मन्त्रालय वा विभागले कुनै आयोजनाको अनुगमन तथा निरीक्षण गर्दा ... निर्देशन दिनेछ।"
- **V:** 1 × (0.40+0.75)/2 = **0.575**.
- **W:** Scored as a Procedural measure (0.40) rather than Regulatory (1.0) because the P-scale explicitly places "planning requirements" and screening/permitting procedures in the 0.40 tier (per the worked examples in the prompt); the *effect* is nonetheless to require prior governmental approval before a qualifying plastic-manufacturing or waste facility can be built or operated. This is the only instrument in the Regulation that names "प्लाष्टिक" (plastic) explicitly (Plastics Relevance Filter criterion (a)); it is also relevant under criterion (c) for the waste-management-facility thresholds. Comment/uncertainty: reasonable analysts could instead classify this as Regulatory (1.0) on the grounds that it operates as a de facto construction/operating permit requirement — flagged here for reviewer judgement.

#### Instrument 7 — Pollution Control Certificate & mandatory sampling/monitoring

- **Q:** Waste management
- **R:** Rule 21 requires the Department or concerned authority to ascertain the pollution/waste ("फोहरमैला") status of emissions from any industry, factory, machine or vehicle, and to collect, test and analyse samples. Rule 22 sets out the Pollution Control Certificate (PCC) regime: an industry must apply with a report on its process, pollution sources and control measures, and a laboratory test report (no older than 3 months) of its effluent ("फोहर पानी"), smoke, dust, noise or light against government standards (22(1)–(2)); the Ministry/Province must form an inspection team to verify pollution control on site (22(3)); the certificate, once issued, is valid 3 years (22(6)); certificate holders must submit a fresh test report every 6 months (22(7)); if standards are exceeded the industry has 3 months to remediate and re-report (22(8)); and the certifying authority **must revoke** the certificate for continued non-compliance or repeated adverse impact (22(9)).
- **S:** 1 — mandatory throughout ("पर्नेछ").
- **T:** 0.75 — authority explicit (Ministry/Province/Department) (+0.25); enforcement explicit (mandatory revocation for non-compliance, 22(9)) (+0.25); monitoring explicit (site-inspection team, recurring 6-monthly effluent/emission testing) (+0.25); *not* scored unconditional because 22(6) allows the certificate to be issued "सशर्त वा निशर्त" ("with or without conditions"), i.e. conditions/exemptions may attach case-by-case (0).
- **U:** "उपनियम (५) बमोजिमको प्रतिवेदनको आधारमा ... तीन वर्षको लागि ... प्रदूषण नियन्त्रण प्रमाणपत्र प्रदान गर्न सक्नेछ ... प्रत्येक छ महिनामा ... परीक्षण प्रतिवेदन पेश गर्नु पर्नेछ ... प्रमाणपत्र खारेज गर्नु पर्नेछ।"
- **V:** 1 × (1.00+0.75)/2 = **0.875**.
- **W:** Borderline plastics relevance flagged: the PCC regime is a general industrial-effluent/emissions control tool, not plastics-specific, but it applies to plastic-manufacturing industries identified in Schedule 2/3 (Instrument 6) and to any industry whose wastewater/waste stream includes plastics or plastic additives — included here under Plastics Relevance Filter criterion (c) ("waste management or governance instrument that commonly applies to plastic waste streams in practice"). A different coder could reasonably exclude this row as too general; documented here for transparency.

#### Instrument 8 — Environment Conservation Fund

- **Q:** Waste management
- **R:** Rules 35–40 (see Column N above for full detail): establishes a dedicated Environment Conservation Fund with its own governing Committee (Rule 35), a Ministry-provided secretariat (Rule 36), operated as a revolving fund in a scheduled commercial bank (Rule 37), whose proceeds may be spent only on the purposes listed in Rule 38 — including "वातावरण प्रदूषण अध्ययन तथा न्यूनीकरण सम्बन्धी प्रविधिको प्रयोग गर्ने" (use of technology for pollution study and reduction) — subject to statutory accounting (Rule 39) and mandatory public annual reporting (Rule 40).
- **S:** 1 — the Committee, secretariat and Fund mechanism are established directly by the Regulation ("रहनेछ" / "गरिनेछ"), not merely enabled for future activation.
- **T:** 0.50 — authority/governance body explicit (+0.25); monitoring explicit via the accounting duty and mandatory published annual report (+0.25); no explicit enforcement/penalty provision for misuse of the Fund (0); not scored unconditional because expenditure is subject to Committee discretion/decision under Rule 38's chapeau ("...सञ्चालक समितिले निर्णय गरे बमोजिम खर्च गरिनेछ") (0).
- **U:** "कोषलाई चक्रीय (रिभल्भिङ्ग) कोषको रूपमा सञ्चालन गरिनेछ ... कोषको आय-व्ययको लेखा प्रचलित कानुन बमोजिम राखिनेछ ... वार्षिक प्रतिवेदन ... सार्वजनिक गर्नु पर्नेछ।"
- **V:** 1 × (0.60+0.50)/2 = **0.55**.
- **W:** Coded as an Economic-incentive-type instrument (0.60) because it is a dedicated financing mechanism (closest analogue in the P-scale to "a dedicated fund is created"); a case could also be made for 0.80 ("Planned government investment") if the Fund is understood to finance physical pollution-control/waste infrastructure — flagged as a judgement call. The Fund is not plastics-exclusive (it also covers heritage conservation and general pollution/environmental education) but is included under Plastics Relevance Filter criterion (c) because Rule 38(b) explicitly funds pollution-reduction technology, which in practice includes waste (incl. plastic-waste) processing technology.

---

## Summary scores

- **policy_score (O):** 0.875
- **Instrument scores (V):** 0, 0.75, 0.675, 0.75, 1.00, 0.575, 0.875, 0.55

## Key uncertainties / flags for reviewer (Column W roll-up)

1. Instrument 1 (Rule 14(1)) is a discretionary "may" power with no evidence of a subordinate order — scored not-in-force per Rule 12's language test; if such an order is later found, S should become 1 and V would need to be recomputed.
2. Instrument 6 (EIA/IEE screening) could alternatively be scored as a Regulatory instrument (P=1.0) rather than a Procedural measure (P=0.40) depending on how strictly "planning requirement" vs "permit requirement" is interpreted; both readings are defensible.
3. Instrument 7 (Pollution Control Certificate) and Instrument 8 (Conservation Fund) are general-purpose pollution/financing instruments rather than plastics-exclusive ones; included under Plastics Relevance Filter criterion (c) but flagged as the two most borderline inclusions in this coding.
4. "जोखिमपूर्ण पदार्थ" (Instruments 4–5) is translated as "hazardous substance/material," which is broader than "hazardous waste" in English; the licence chain nonetheless spans collection through disposal and therefore functions as hazardous-waste management regulation in practice.
5. OCR of a 167-page scanned Nepali gazette is imperfect; all quotations above were checked against the rendered page images during coding, and page/rule numbers are given so each quotation can be independently re-verified against [`sources/environment-protection-regulation-2020-nepal-source.pdf`](sources/environment-protection-regulation-2020-nepal-source.pdf).

See [`nepal-environment-protection-regulation-2020.csv`](nepal-environment-protection-regulation-2020.csv) for the fully flattened Columns A–W table (Policy-level fields A–O repeated identically on all 8 rows, as required by the coding prompt).
