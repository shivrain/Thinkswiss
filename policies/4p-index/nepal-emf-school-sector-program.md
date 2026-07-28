# 4P Index Coding — Environmental Management Framework (EMF) for the School Sector Reform Plan / School Sector Program, Nepal

## Provenance / methodology

- **Source document:** Asian Development Bank project document, "EARF: Nepal: School Sector Program" (ADB project 35174-082-NEP), archived at [`sources/emf-school-sector-program-nepal-source.pdf`](sources/emf-school-sector-program-nepal-source.pdf). Public URL: https://www.adb.org/sites/default/files/project-documents/35174-082-nep-earfab.pdf.
- The document's own subtitle is **"Environmental Management Framework"** (its footnote 2 clarifies this is "equivalent to Environmental Assessment and Review Framework (EARF) as per Safeguard Policy Statement 2009 (SPS) of ADB"). Paragraph 6 of the document states: *"a donor harmonized EMF was prepared for SSP in May 2009. The EMF has been updated for the purpose of SSP..."* — i.e. this document is an **updated version** (prepared ~2011/2012, per its own PDF metadata) of the original May-2009 Environmental Management Framework prepared by Nepal's Department of Education for the School Sector Reform Plan (SSRP), used here because the original stand-alone 2009 DOE document could not be independently located; the two documents describe the same underlying framework and institutional/mitigation content.
- **Processing:** the PDF is a clean, text-native Microsoft Word document; text was extracted directly with `pdftotext -layout` (archived at [`sources/emf-school-sector-program-nepal-extracted-text.txt`](sources/emf-school-sector-program-nepal-extracted-text.txt)) and read in full.

## Key coding decisions

- **C (policy_year) = 2012.** Based on the source PDF's own metadata (creation/modification dates late 2011–early 2012) and its references to the School Sector Program covering FY2013–2014; the original EMF this document updates dates to May 2009 — flagged for reviewers who may prefer to code this policy as "2009" if treating it as a continuation of the original DOE document rather than this specific ADB-updated version.
- **G (policy_type) = 0.50.** A donor-harmonised safeguard "Framework," approved by the Ministry of Education, not enacted by Parliament nor Gazette-published as an executive regulation — but it does contain quantified operational elements (an itemised implementation budget, Table 8; monitoring indicators with defined frequencies, Tables 9–11), so it sits at the higher "plan/programme with quantifiable targets" tier rather than the purely aspirational tier.
- **E/F (policy_target) = 0.5.** "Ban use of plastic products in schools" is a plastics-specific, non-quantifiable commitment.
- **I/J (sectors):** education, waste management, health, water, construction, disaster risk management, energy → 7 sectors → I = 1.
- **K/L (circularity):** consumption (in-school plastic-product ban), recycling (waste minimisation/recovery/recycling awareness), disposal (solid waste management system), environmental leakage (waste-spreading/aesthetics mitigation) are addressed; production is not → K = 0.75.
- **M/N (budget) = 1.** Table 8 sets out a specific, itemised implementation budget (translation/printing, safeguard-desk setup, capacity building, per-subproject environmental-assessment and monitoring cost allocations) dedicated to this Framework's own implementation.
- **policy_score (O) = (0.50 + 1 + 0.75 + 1) / 4 = 0.8125.**

## Instruments coded (4 rows; see the CSV/XLSX for full detail)

1. Ban on use of plastic products in schools, combined with hazardous-waste awareness-raising/recycling and safe-disposal requirements (mitigation table, "Hazardous waste" row) — Regulatory (1.0), **in force**. The only provision in this document naming "plastic" explicitly.
2. Mandatory solid waste segregation, proper disposal and composting system in schools (mitigation table, "Solid waste management" row) — Regulatory (1.0), **in force**.
3. Multi-level institutional implementation/monitoring mechanism (MOE → DOE → DEO → SMC, safeguard desks, Environmental Officer, quarterly/annual compliance audits) — Procedural/coordination (0.40), **in force**.
4. Itemised EMF implementation budget (Table 8) — Planned government/donor investment (0.80), **not in force** (framed as a "tentative" estimate that "may include" the listed items — an enabling/indicative rather than binding allocation, per Rule 12's "may" test).

Full A–W table: [`nepal-emf-school-sector-program.csv`](nepal-emf-school-sector-program.csv) / [`.xlsx`](nepal-emf-school-sector-program.xlsx).

## Uncertainties flagged

- No explicit fine, penalty or other enforcement mechanism was found anywhere in this Framework for non-compliance with its environmental mitigation measures (including the plastic ban) — all three "in force" instruments therefore score 0 on the enforcement sub-criterion of Column T, capping their implementation scores at 0.75 despite explicit authority and monitoring provisions.
- This document is narrower in scope than the other four Nepali policies coded in this index (it governs only donor-financed school-infrastructure sub-projects, not the general population/economy), which is reflected in a more moderate sectoral spread (7 sectors, still enough for I=1) and a circularity score that excludes the production phase entirely.
- Whether to treat this as the "same" policy as the original May-2009 DOE-authored EMF, or as a materially updated/distinct instrument, is a judgement call — flagged above under Column C.
