# 4P Index — Nepal Policy Coding

Nepal-focused coding of national plastic pollution-relevant policies into the **Plastic Pollution Policy
Index (4P Index)** format, following the "4P Index — Policy Coding Prompt (Version 2)" strictly: one
policy-level record (Columns A–O) per policy, repeated identically across one row per distinct policy
instrument (Columns P–W).

Every policy below was sourced from an official document (either fetched from the cited government/NGO
resource page, or provided directly), OCR'd/text-extracted as needed, and coded from a direct reading and
translation of the actual document text — not from secondary summaries. See each policy's own `.md` file
for full methodology, source provenance, exact quotations (Nepali + English translation) and per-instrument
justification.

## Policies coded

| # | Policy | Year | Type | Policy score (O) | Files |
|---|---|---|---|---|---|
| 1 | Environment Protection Regulation, 2077 (2020) | 2020 | Regulation (0.75) | 0.875 | [.md](nepal-environment-protection-regulation-2020.md) · [.csv](nepal-environment-protection-regulation-2020.csv) · [.xlsx](nepal-environment-protection-regulation-2020.xlsx) |
| 2 | Waste Management National Policy, 2079 (2022) | 2022 | Policy (0.25) | 0.625 | [.md](nepal-waste-management-national-policy-2079-2022.md) · [.csv](nepal-waste-management-national-policy-2079-2022.csv) · [.xlsx](nepal-waste-management-national-policy-2079-2022.xlsx) |
| 3 | Solid Waste Management Act, 2068 (2011, as amended) | 2019 | Legislation (1.0) | 1.0 | [.md](nepal-solid-waste-management-act-2068-2011.md) · [.csv](nepal-solid-waste-management-act-2068-2011.csv) · [.xlsx](nepal-solid-waste-management-act-2068-2011.xlsx) |
| 4 | The Sixteenth Plan (F.Y. 2024/25–2028/29) | 2024 | Plan w/ targets (0.50) | 0.75 | [.md](nepal-sixteenth-plan-2024-25-2028-29.md) · [.csv](nepal-sixteenth-plan-2024-25-2028-29.csv) · [.xlsx](nepal-sixteenth-plan-2024-25-2028-29.xlsx) |

**29 instrument-level rows in total** across the four policies (8 + 6 + 10 + 5).

## Combined / master files

- [`nepal-4p-index-master.csv`](nepal-4p-index-master.csv) — all 29 rows from all 4 policies, stacked into a single flat A–W table (one continuous spreadsheet, matching the structure of a master 4P Index dataset).
- [`nepal-4p-index-master.xlsx`](nepal-4p-index-master.xlsx) — the same data as an Excel workbook, with a combined "All policies" sheet plus one sheet per individual policy.

## Sources

Original source PDFs, OCR transcripts and extracted text for every policy are archived under
[`sources/`](sources/) for full auditability — every quotation used in the coding can be independently
re-verified against these files.

## Cross-cutting notes

- **Legislative hierarchy of the four documents coded:** the *Solid Waste Management Act, 2068* is Nepal's
  foundational binding legislation for waste management (G=1); the *Environment Protection Regulation,
  2077* is the executive regulation implementing the parent Environment Protection Act (G=0.75); the *Waste
  Management National Policy, 2079* is a non-binding strategic policy (G=0.25); and *The Sixteenth Plan* is
  a five-year national development plan with quantifiable (but not plastics-specific) targets (G=0.50).
  Reading them together shows how Nepal's plastics-relevant obligations are distributed across different
  levels of legal force — from a firm, criminally-enforced statutory ban mechanism (Solid Waste Management
  Act §38(ण), the basis for Nepal's plastic bag bans) down to aspirational five-year-plan commitments (e.g.
  the explicit extended-producer-responsibility commitment in the Sixteenth Plan, which has not yet been
  operationalised in any of the other three documents).
- **Most directly plastics-specific provisions found:**
  - Solid Waste Management Act, 2068, §38(ण)/§39(10) — the Gazette-notice power to ban "excessive-waste-generating goods," exercised in practice for thin plastic bags.
  - Environment Protection Regulation, 2077, Schedule 2, Industry sector, items 41 & 43 — the only provisions across all four documents that name "प्लाष्टिक" (plastic) explicitly.
  - The Sixteenth Plan, Ch. 7 & 13 — "discourage the use of plastic materials and prohibit burning plastics"; "control plastic items"; "displacing plastic products."
- **Recommended follow-up document** (not sourced/coded in this exercise, but identified during research): the Ministry of Forests and Environment's **प्लाष्टिक झोला (नियमन तथा नियन्त्रण) निर्देशिका, २०८२** ("Plastic Bags (Regulation and Control) Directive, 2082 / 2025-26"), available at https://mofe.gov.np/content/348/plastic-bags--regulation-and-control--guidelines--2082/ — this is Nepal's most recent and most directly plastics-specific legal instrument and would substantially strengthen a future iteration of this index.
