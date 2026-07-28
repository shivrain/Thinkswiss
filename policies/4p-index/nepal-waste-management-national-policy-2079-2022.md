# 4P Index Coding — Waste Management National Policy, 2079 (2022), Nepal

## Provenance / methodology

- **Source page:** https://dpnet.org.np/resource-detail/1781
- **Source document (direct PDF link found on that page):** `https://dpnet.org.np/uploads/files/Notices-20220619152630462 2024-06-06 10-24-18.pdf`, archived at [`sources/waste-management-national-policy-2079-nepal-source.pdf`](sources/waste-management-national-policy-2079-nepal-source.pdf).
- **Original language / title:** Nepali — फोहरमैला व्यवस्थापन राष्ट्रीय नीति, २०७९ ("Waste Management National Policy, 2079"). A 10-page scanned PDF: page 1 is the Cabinet Secretariat's covering memorandum recording Council of Ministers approval on 2079/1/25 B.S.; page 2 is the title page; pages 3-10 are the policy text (Sections 1-16: Background, Past Efforts, Current Status, Problems & Challenges, Need for a New Policy, Vision, Goal, Objectives, Strategy, Working Policies (कार्यनीति) 9.1-9.8, Institutional Structure, Role of the three tiers of government, Resource Management, Legal Arrangements, Monitoring & Evaluation, and Risks).
- **Processing:** rendered all 10 pages to images (`pdftoppm`, 200 dpi) and read/translated each page directly against the image (the scan is a clean, non-rotated vector-quality print, so direct visual reading was more reliable than OCR for this short document). OCR output is nonetheless archived at [`sources/waste-management-national-policy-2079-ocr-nepali.txt`](sources/waste-management-national-policy-2079-ocr-nepali.txt) for cross-checking.
- No amendment to this policy was found.

## Key coding decisions

- **G (policy_type) = 0.25.** This is a "राष्ट्रिय नीति" (National Policy) approved by the Council of Ministers — a strategy/vision document, not a binding Regulation or Act. Having read the full text, it does **not** contain any quantifiable target (contrary to the external DPNet website summary, which claims a "zero-waste-to-landfill by 2030" target — this phrase does **not** appear anywhere in the actual 10-page policy document; see the flag in Column W of Instrument 1). Because no quantifiable target is present in the primary text itself, it is coded at the aspirational tier (0.25) rather than 0.50.
- **E/F (policy_target) = 0.** No plastic-specific target (quantifiable or otherwise) appears in the document text.
- **I/J (sectors):** waste management, industry, health, municipalities, drinking water & sanitation, education, households, private sector → 8 sectors → I = 1.
- **K/L (circularity):** production, recycling, disposal, environmental leakage are addressed; consumption is not directly addressed → K = 0.75.
- **M/N (budget) = 0.5.** Section 13 assigns implementation funding to local levels' own budgets, with co-financing/foreign-aid/PPP mobilisation — a funding source is mentioned but no dedicated fund is created by this policy.
- **policy_score (O) = (0.25 + 1 + 0.75 + 0.5) / 4 = 0.625.**

## Instruments coded (6 rows; see the CSV/XLSX for full detail)

1. Ban on waste disposal in rivers/lakes/wetlands/protected & heritage areas (Working Policy 9.2.4) — Regulatory (1.0), **not in force** (a future standard-setting commitment).
2. Commitment to enact an integrated waste-management law covering hazardous/chemical/industrial/health-institutional/household waste (Working Policy 9.1.1) — Procedural (0.40), **not in force**.
3. Promotion of source segregation, reuse, processing and composting (Working Policy 9.3.1) — Voluntary/persuasion (0.20), **not in force**.
4. Multi-level role clarification: federal policy/standards role, provincial coordination role, local-level operational responsibility (Section 12, Strategy 9.5) — Procedural/coordination (0.40), **in force** (mandatory role assignment).
5. National Waste Management Coordination Committee (Section 11) — Procedural/institutional setup (0.40), **in force**.
6. Promotion of PPP/co-financing/community investment in waste management (Strategy 9.6) — Economic incentive (0.60), **not in force**.

Full A–W table: [`nepal-waste-management-national-policy-2079-2022.csv`](nepal-waste-management-national-policy-2079-2022.csv) / [`.xlsx`](nepal-waste-management-national-policy-2079-2022.xlsx).

## Uncertainties flagged

- The DPNet website's summary of this policy claims a "zero-waste-to-landfill by 2030" target; this could not be verified in the source PDF's actual text and is therefore **not** reflected in Column E/F. If a fuller/annexed version of the policy containing such a target is located, this coding should be revisited.
- "Waste Management National Strategic Policy Framework, 2079" and "National Five-Year Plan for Solid Waste Management," which appear as separate line items in some reference spreadsheets, were not located as independently published documents distinct from this Policy during sourcing; they may be alternative titles/summaries of this same document, or unpublished internal planning documents. Only this DPNet-hosted policy document was coded here — flagged for follow-up if the other titles refer to genuinely distinct, separately published instruments.
