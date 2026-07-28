# 4P Index coding — Nepal plastic-pollution policies

This folder contains a coding of Nepal's national plastic-pollution-relevant
policy landscape according to the **Plastic Pollution Policy Index (4P
Index) — Policy Coding Prompt, Version 2** (multi-level governance
coordination instruments; clarified in-force coding).

## Files

- `4p_index_nepal_plastic_policies.csv` — the coded table (columns A–W as
  defined in the prompt). One row per policy **instrument**; columns A–O are
  repeated identically for every instrument belonging to the same policy.
  Column O (`policy_score`) and Column V (`instrument_score`) are computed
  automatically using the prescribed formulas:
  - `policy_score = (policy_type + policy_integration + policy_circularity + policy_budget) / 4`
  - `instrument_score = instrument_in_force * (instrument_type + instrument_implementation) / 2`
- `build_4p_index.py` — the script that generates the CSV from a structured,
  fully-sourced dataset (run with `python3 build_4p_index.py`). Re-running it
  regenerates the CSV deterministically from the same underlying data, and is
  the place to add, correct or extend coded policies/instruments.

## Scope and important caveats

The request that produced this table supplied a screenshot of a spreadsheet
(rather than machine-readable text or links) containing what appears to be an
existing/partial coding exercise covering Nepal (plastic bag/plastic-flower
bans, the Solid Waste Management National Policy, national park management
plans, COVID-19 sector reopening guidelines, "The Sixteenth Plan," etc.),
plus a small number of Bhutan/Bangladesh multi-level ESMF framework rows.

At the resolution supplied, the screenshot's cell text (exact article
numbers, quoted target language, source URLs, and many numeric codes) could
not be read reliably enough to satisfy the coding prompt's evidentiary bar
("Each sub-score must be explicitly evidenced in the policy text — do not
infer or assume," Rule 7; "quantifiable targets" must be copy-pasted
verbatim, Col F). Rather than fabricate quotations or article citations from
an unreadable image, this table was built by independently researching and
verifying the actual, identifiable Nepali national policies from primary and
credible secondary sources (government Acts/Gazette notices, FAOLEX legal
texts, and recent policy-analysis publications by SWITCH-Asia/adelphi (2025)
and CREASION (2023)). See the `policy_url` column and in-line comments for
citations.

**Policies coded (11 policies, 26 instrument rows):**

1. Environment Protection Act, 2076 (2019) — the parent enabling legislation
   for plastics/pollution regulation.
2. Plastic Bag (Regulation and Control) Directive, 2082 (2025) — the current
   nationwide ban (<40 microns) plus EPR and multi-level enforcement.
3. Notice on Prohibition of Plastic Artificial Flowers (2079/2022 Nepal
   Gazette).
4. Solid Waste Management National Policy, 2079 (2022) — zero-waste-to-
   landfill-by-2030 target, multi-tier institutional roles.
5. Solid Waste Management Act, 2068 (2011) — source-segregation mandate,
   private-sector licensing, facility-construction duties.
6. Local Government Operation Act, 2074 (2018).
7. Federal, Provincial and Local Level (Coordination and Interrelation) Act,
   2077 (2020) — coordination councils (a direct example of the v2
   multi-level governance coordination instruments).
8. Inter-Governmental Fiscal Transfer Act, 2074 (2017) — SWM financing/grant
   mechanism and local waste-transport taxation.
9. National Environmental Policy, 2019.
10. The Sixteenth Plan (National Five-Year Development Plan, FY2024/25–
    2028/29).
11. Notice of Khumbu Pasang Lhamu Rural Municipality banning single-use
    plastics in the Everest/Solukhumbu region (2019/2020) — included as a
    plausible match for a "tea house"/Everest-region item in the source
    material, but flagged throughout as **sub-national** and therefore
    outside the national-policy scope the 4P Index is designed to capture.

**Not coded in this pass**, because their primary text could not be located
or verified with confidence from the supplied screenshot alone:
- Master Plan of Ghodaghodi Lake Area
- Management plans for Bardia, Banke and Shuklaphanta National Parks and
  buffer zones
- COVID-19 sector-reopening operational guidelines (tourism, food/retail)
- Bhutan and Bangladesh multi-sector Environmental & Social Management
  Frameworks (ESMF/SEMF)

If the underlying documents (full text, PDFs, or working links) for any of
these — or for other policies — can be supplied directly, they can be added
to `build_4p_index.py` and recoded following exactly the same v2 methodology
(plastics relevance filter, in-force language test, implementation
sub-scoring, etc.).

## Key coding judgment calls (see `comments`/Column W for the rest)

- **Solid Waste Management National Policy, 2079 (2022)**: coded
  `policy_type = 0.50` (strategy/policy with a quantifiable target — zero
  waste to landfill by 2030) rather than 0.75/1, since it is a Cabinet-level
  policy document, not an Act or a stand-alone regulation.
- **EPR under the Solid Waste Management National Policy** is only described
  as "laying the groundwork" for EPR (aspirational) and is therefore coded
  `instrument_in_force = 0`; the actually binding, operative EPR obligation
  is coded separately under the Plastic Bag (Regulation and Control)
  Directive, 2082.
- **The Sixteenth Plan** and **National Environmental Policy, 2019**
  instruments are coded `instrument_in_force = 0` throughout: the language
  reviewed (via secondary paraphrase, since the primary Nepali-language
  documents were not independently accessed) is consistently aspirational
  ("suggests," "is encouraged," "there is mention of") rather than
  "shall"/"must" — consistent with Rule 12's in-force test. This is flagged
  as an area for re-verification against primary text.
- **Khumbu Pasang Lhamu Rural Municipality ban**: an executive decision of a
  *local* government body, not a national instrument; retained only because
  it is the closest verifiable match to an item that appears to be referenced
  in the source material, and clearly labelled as sub-national throughout.
