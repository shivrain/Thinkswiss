# 4P Index Coding — Solid Waste Management Act, 2068 (2011), Nepal

## Provenance / methodology

- **Source document:** provided directly by the user (`Solid_Waste_Management_Act_a561.pdf`), archived at [`sources/solid-waste-management-act-2068-nepal-source.pdf`](sources/solid-waste-management-act-2068-nepal-source.pdf). This is byte-identical to the copy hosted by FAO's FAOLEX legal database at https://faolex.fao.org/docs/pdf/nep137767.pdf (used here as the citable public URL), which is itself sourced from Nepal's Law Commission (www.lawcommission.gov.np).
- **Original language / title:** Nepali — फोहरमैला व्यवस्थापन ऐन, २०६८ ("Solid Waste Management Act, 2068"), Act No. 5 of 2068 B.S. (2011), as amended by "केही नेपाल ऐन संशोधन गर्ने ऐन" in 2072 B.S. (2015/16) and 2075 B.S. (2018/19).
- **Processing:** the PDF is a Microsoft-Word-generated, vector-text document (not scanned), but its embedded Nepali text uses a legacy font/ToUnicode mapping that produces systematically corrupted (mojibake) Devanagari when extracted with `pdftotext` (a well-known issue with older Nepali government Word documents). Rather than rely on that corrupted extraction, all 23 pages were rendered to images (`pdftoppm`, 150 dpi) and **read and translated directly from the page images**, which render cleanly and legibly. The (corrupted) raw `pdftotext` output is nonetheless archived for reference. English-translation cross-checking was additionally done against publicly available unofficial translations of this Act (e.g. Studocu/Scribd copies of the same FAOLEX document) to confirm terminology, without relying on them as the primary source.
- Most recent amendment: 2075 B.S. (≈2018/19 CE) → policy_year coded as 2019.

## Key coding decisions

- **G (policy_type) = 1.** Enacted by the Constituent Assembly sitting as the Legislature-Parliament under Article 83 of the Interim Constitution, 2063, and twice amended by Parliament — full legislative pedigree.
- **I/J (sectors):** production, retail, waste management, recycling, agriculture, industry, municipalities, chemicals, health → 9 sectors → I = 1.
- **K/L (circularity):** all five lifecycle phases are addressed (production — the Section 38(ण) product ban; consumption — source-reduction duty on every person/institution, Section 5; recycling — Section 6/10; disposal — Sections 3/7/12; environmental leakage — the river/public-place dumping prohibitions, Sections 7(3)/38) → K = 1.
- **M/N (budget) = 1.** Section 18(5) ring-fences the Act's own service-fee revenue for waste management, environmental protection and affected-area development — a self-generated, dedicated funding stream.
- **policy_score (O) = (1 + 1 + 1 + 1) / 4 = 1.0.**

## Instruments coded (10 rows; see the CSV/XLSX for full detail, including exact Nepali quotations, translations and section citations)

1. Local-level duty to build/operate waste infrastructure and provide collection/disposal/processing services (§3) — Regulatory (1.0), in force.
2. Generator-responsibility for hazardous/health-institutional/chemical/industrial waste (§4(2)-(3)) — Regulatory (1.0), in force.
3. Mandatory source segregation at source, with dedicated criminal offence (§6, §38(थ), §39(11)) — Regulatory (1.0), in force.
4. Ban on disposing hazardous/chemical waste at collection/transfer centres (§7(3)) — Regulatory (1.0), in force.
5. Duty to promote waste minimisation, reuse and recycling, incl. local guideline-issuing power (§10) — Regulatory (1.0), in force.
6. Permit requirement + competitive tendering + royalty regime for private/community waste-management operators (§§13-17) — Regulatory (1.0), in force.
7. Ring-fenced waste-management service fee, with mandatory discount for disadvantaged groups (§§18-19) — Economic incentive (0.60), in force.
8. Pollution-control duty + mandatory routine monitoring of waste management/disposal (§§20-21) — Regulatory (1.0), in force.
9. National Solid Waste Management Council (§§23-25) — Procedural/institutional setup (0.40), in force.
10. Ban on production/sale/distribution of Gazette-notified "excessive-waste-generating" goods — the statutory basis for Nepal's plastic-bag bans (§38(ण), §39(10)) — Regulatory (1.0), in force (enabling power already exercised via subordinate instruments, e.g. the Plastic Bags (Regulation and Control) Directive, 2082).

Full A–W table: [`nepal-solid-waste-management-act-2068-2011.csv`](nepal-solid-waste-management-act-2068-2011.csv) / [`.xlsx`](nepal-solid-waste-management-act-2068-2011.xlsx).

## Uncertainties flagged

- Instrument 10's in-force scoring (S=1) relies on external knowledge that Nepal has exercised the Section 38(ण) Gazette-notice power for plastic bags (via, most recently, the Ministry of Forests & Environment's "प्लाष्टिक झोला (नियमन तथा नियन्त्रण) निर्देशिका, २०८२"). That subordinate directive was identified via web search (mofe.gov.np) but was **not** itself sourced, OCR'd or separately coded in this exercise — recommended as a priority follow-up document for the 4P Index given it is Nepal's most directly plastics-specific legal instrument.
- Several sections of the Act were repealed/renumbered by the 2072 and 2075 amendment Acts (e.g. Chapter 8 is now blank); the coding above reflects the Act **as currently in force** after both amendments, based on the footnotes printed in the source PDF itself.
