# 4P Index Coding — The Sixteenth Plan (F.Y. 2024/25–2028/29), Nepal

## Provenance / methodology

- **Source document:** provided directly by the user (`16th_5_Yr_Plan_Nepal_fbc3.pdf`), archived at [`sources/sixteenth-plan-2024-25-2028-29-nepal-source.pdf`](sources/sixteenth-plan-2024-25-2028-29-nepal-source.pdf). A 248-page official "Unofficial Translation" (so labelled on its own title page) published in English by the Government of Nepal, National Planning Commission, Singha Durbar, Kathmandu (May 2024).
- **Processing:** this PDF has a clean, directly extractable English text layer (no OCR needed). Full text was extracted with `pdftotext -layout` (layout mode was needed to correctly parse the plan's numerous tables, e.g. the quantitative-targets tables), archived at [`sources/sixteenth-plan-2024-25-2028-29-english-extracted-text.txt`](sources/sixteenth-plan-2024-25-2028-29-english-extracted-text.txt). The full 248-page plan covers every sector of Nepal's economy (17 substantive chapters); it was searched systematically for plastics/waste/recycling/circular-economy/EPR/landfill keywords, and every resulting passage was read in full surrounding context before coding.
- No amendment; this is the current, in-force five-year plan (approved May 2024, covering FY2024/25–2028/29).

## Key coding decisions

- **G (policy_type) = 0.50.** A National Planning Commission five-year Plan — a strategy/plan document, not enacted by Parliament and not a binding regulation, but it does incorporate quantifiable targets in dedicated "Quantitative Targets" tables for every chapter (e.g. the solid-waste-facilities target used in Instrument 5 below) — hence the higher "Plan with quantifiable targets" tier (0.50) rather than the purely aspirational tier (0.25). None of the Plan's own quantified targets is itself plastics-specific, however.
- **E/F (policy_target) = 0.5.** Plastics-specific commitments are present ("discourage the use of plastic materials and prohibit burning plastics"; "control plastic items"; "displacing plastic products"), but none is expressed as a quantified number/percentage — hence 0.5, not 1.
- **I/J (sectors):** given this is an all-of-government national development plan, the sectoral count is very high — production, consumption, retail, waste management, recycling, agriculture, industry, tourism, municipalities, packaging, chemicals, water, energy, forestry, urban development, health, transportation → I = 1.
- **K/L (circularity):** all five lifecycle phases are addressed somewhere across the Plan's 17 chapters (e.g. production — green-economy raw-material substitution; consumption — plastic-item control commitments; recycling — Ch.7 source-segregation infrastructure; disposal — landfill-facility targets; environmental leakage — river/forest dumping-and-burning prohibitions) → K = 1.
- **M/N (budget) = 0.5.** Funding mechanisms (green bonds, concessional finance, carbon/ecosystem-service payments, sectoral capital budgets) are referenced throughout, but no dedicated, ring-fenced plastics/waste fund is created by the Plan itself.
- **policy_score (O) = (0.50 + 1 + 1 + 0.5) / 4 = 0.75.**

## Instruments coded (5 rows; see the CSV/XLSX for full detail)

1. Integrated urban waste management program — source segregation infrastructure + "discourage plastic materials, prohibit burning plastics" (Ch. 7, Program 10) — Regulatory (1.0), **not in force** (Plan-level program commitment).
2. Pollution-control program — landfill-disposal tax + dumping/burning prohibition + "control plastic items" (Ch. 13, Program 4) — Regulatory (1.0), **not in force**.
3. Green-economy raw-material substitution displacing plastic products (Ch. 13, Program 1) — Economic incentive (0.60), **not in force**.
4. Extended producer responsibility commitment for producers/importers/polluters of industrial, hospital and hazardous waste (Ch. 13, Transformative Strategy 3) — Regulatory/EPR (1.0), **not in force** (strategic commitment to "implement the concept of").
5. Quantified target: 1 → 17 modern integrated solid-waste-management centres and sanitary landfill sites by FY2028/29 (Ch. 7.6, Quantitative Targets, item 8) — Planned government investment (0.80), **in force** (the Plan's own numerically specific, dated capital-investment commitment).

Full A–W table: [`nepal-sixteenth-plan-2024-25-2028-29.csv`](nepal-sixteenth-plan-2024-25-2028-29.csv) / [`.xlsx`](nepal-sixteenth-plan-2024-25-2028-29.xlsx).

## Uncertainties flagged

- Four of the five instruments are scored not-in-force (S=0) on the basis that a national development Plan's program-level directives ("prohibit...", "impose...", "control...") describe the government's intended course of action for the plan period rather than self-executing binding law, distinguishing them from the Plan's own quantified, dated resource-allocation commitments (which ARE scored in force). This treatment is a judgement call, flagged for reviewer discretion — an alternative, stricter reading of Rule 12's "shall/must" language test could score some of these Program clauses in force given their imperative phrasing, notwithstanding the non-legislative character of the source document.
- The EPR commitment (Instrument 4) is the most significant future-facing plastics-relevant signal found in Nepal's current policy landscape across all four documents coded in this index; it is recommended as a specific item to track for a future implementing regulation during FY2024/25–2028/29.
