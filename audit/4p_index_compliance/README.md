# 4P Index Coding — Compliance Audit (v2 prompt)

This audit checks **every** "4P Index" policy-coding spreadsheet that exists anywhere in this
repository's branch history against the *Plastic Pollution Policy Index (4P Index) — Policy
Coding Prompt, Version 2* (multi-level governance coordination instruments + clarified
in-force coding).

- **Scope:** 119 unique coding spreadsheets, spanning Norway (27), Mozambique (26), Senegal (21),
  Nepal/Karnali (20), Peru (18), São Tomé and Príncipe (7). These live across ~90 separate
  `cursor/...` branches/PRs (none of this data is on `main` yet).
- **Method:** an automated checker (`validate.py`) parses every sheet, maps columns to the
  canonical A–W schema by field name (robust to files that mislabel or shift columns), and
  checks every mechanical rule in the prompt (allowed value sets, O/V auto-formulas, F/N
  empty-cell rules, A–N policy-level invariance across a policy's rows, Col L/Q controlled
  vocabulary and casing, sector/lifecycle count-vs-bucket consistency). It also runs
  keyword-based heuristics for the two things Version 2 specifically clarifies — Rule 12
  ("may" vs "shall" in-force test) and Rule 11 (multi-level coordination instruments) — and
  flags those as **REVIEW** items for a human to confirm, since they require judgement the
  script cannot make reliably on its own.
- **Full results:** `findings_detail.csv` (one row per finding — branch, file, spreadsheet row,
  column, severity, description) and `PER_FILE_DETAIL.md` (the same findings, grouped and
  de-duplicated per file). Re-run with `python3 validate.py` after adjusting
  `/tmp/4p_audit/raw` (see script header) to regenerate.

**Reproducing this audit:** from a checkout with all remote branches fetched
(`git fetch origin --prune`), regenerate `files_manifest.txt` by scanning every branch for
`*4P*Index*.{xlsx,csv}` files, then run `python3 extract_files.py` (dedupes by content hash into
`dedup_manifest.tsv` + a `raw/` folder), `python3 validate.py` (writes `results.json`), and
`python3 export_report.py` (writes `findings_detail.csv` + `PER_FILE_DETAIL.md`). The scripts as
committed here reference the `/tmp/4p_audit` working directory used during this run; adjust the
`RAW`/`DEDUP_MANIFEST`/output paths at the top of each script if re-running elsewhere.

Severity key: **ERROR** = definite violation of an explicit prompt rule (wrong value on a fixed
scale, formula mismatch, non-identical policy-level fields, wrong vocabulary/casing). **WARN** =
a consistency check that usually indicates a miscount but could be a legitimate edge case.
**REVIEW** = Rule 11/12 heuristic flag — needs a human to re-read the cited text and confirm.

## Headline numbers

| Country | Unique files | Files with ≥1 finding |
|---|---|---|
| Norway | 27 | 27 |
| Mozambique | 26 | 23 |
| Senegal | 21 | 14 |
| Nepal / Karnali | 20 | 20 |
| Peru | 18 | 18 |
| São Tomé and Príncipe | 7 | 7 |

6,551 total findings: **4,439 ERROR**, **1,930 WARN**, **182 REVIEW**.

## The systemic, high-value issues (fix these first)

### 1. Column V (`instrument_score`) ignores the Col S weighting — 23 files, ~128 rows

The prompt defines `V = S × (P + T) / 2`, i.e. an instrument that is **not in force (S=0) must
score 0**. In 23 files (mostly Norway) V was instead computed as the plain `(P+T)/2`, so
enabling-power instruments that are correctly marked `S=0` still carry a non-zero score.

Example — `4P_Index_Forurensningsloven.xlsx` (branch `cursor/4p-index-forurensningsloven-65aa`),
row 6: P=1, T=0.75, **S=0**, but V=0.875 (should be 0). Same pattern in
`4P_Index_Avfallsforskriften*.xlsx`, `4P_Index_Produktforskrift*.xlsx`,
`4P_Index_Skipssikkerhet_MARPOL_V.xlsx`, `4P_Index_REACH_forskrift.xlsx`,
`Peru_DS_244_2019_EF_4P_Index_Coding.xlsx`, `Peru_Law_30884_4P_Index_Coding.xlsx`, and others —
see `findings_detail.csv` filtered on `finding` containing `S×(P+T)/2 = 0.0×`.

**Fix:** recompute Col V for every row where S=0 → V=0.

### 2. Column O (`policy_score`) is not held constant per policy — 50 files, mostly Peru + Norway

The prompt is explicit that `policy_score` is a **policy-level** field and "must therefore be
identical across all rows belonging to the same policy" — it is not meant to trend upward
row-by-row. In roughly 50 files (nearly every Peru file, and about a dozen Norway files) Col O
instead increases/varies from row to row, which looks like it was computed by blending in each
instrument's own score rather than only `(G+I+K+M)/4`.

Example — `Peru_DS_244_2019_EF_4P_Index_Coding.xlsx`: O = 0.69, 0.73, 0.85 across its rows,
none of which equal the correct constant value 0.812 = (0.75+0.75+0.75+1.0)/4.
Same pattern in `Peru_DS_005_2010_MINAM…`, `Peru_Ley_31896…`, `Peru_DS_003_2025_MINAM…`,
`Peru_DS_023_2021_MINAM…`, `Peru_PLANRES…`, `4P_Index_Forurensningsloven.xlsx`,
`4P_Index_Avfallsforskriften_Kap6.xlsx`, and more.

**Fix:** recompute Col O once per policy as `(G+I+K+M)/4` and copy the same number into every
row of that policy.

### 3. Some files add non-spec columns, shifting every subsequent letter — 24 files

11 Peru files (e.g. `Peru_DS_005_2010_MINAM_4P_Index_Coding.xlsx`,
`Peru_DL_1278_4P_Index_Coding.xlsx`) and several Nepal/Mozambique files insert extra columns —
most commonly a `policy_name_es`/`policy_name_en` split and a `country` column — **before**
`policy_url`. Because the prompt defines an exact, fixed A–W schema, this means what these files
literally label "Col C" is `country`, not `policy_year`, "Col E" is `policy_year`, not
`policy_target`, and so on all the way across the row.

**Fix:** either drop the extra columns, or keep them but append them **after** column W so A–W
still match the prompt's schema exactly.

### 4. Column Q (`instrument_lifecycle_stage`) uses free text instead of the 6 fixed values — 46 files

Rule 4 restricts Col Q to exactly: `Production, Consumption, Recycling, Waste management,
End of life, Environmental leakage` (title case). In practice, 46 files instead used one of:
`Disposal`, `End-of-life` (hyphenated), `Collection`, `Litter/Pollution`, `Planning`,
`Design`, `Use/Consumption`. Note `Disposal` is one of the *five Col L phases* but is **not**
one of the *six Col Q stages* — this is an easy trap because the two columns use overlapping but
not identical vocabularies.

**Fix:** remap — `Disposal`→`Waste management` or `End of life` (context-dependent),
`End-of-life`→`End of life`, `Collection`→`Waste management`, `Litter/Pollution`→
`Environmental leakage`, `Use/Consumption`→`Consumption`, `Planning`/`Design`→closest of the six
(often `Production`), or leave a W-column note if none genuinely fits.

### 5. Column L (`policy_lifecycle_phases_list`) uses non-English or annotated terms — 36 files

Rule 3 requires exactly the five lower-case English terms
(`production, consumption, recycling, disposal, environmental leakage`). Several
Spanish/Portuguese/French-language jurisdictions (chiefly Peru, and a few Senegal/Mozambique
files) instead wrote Spanish/Portuguese words (`consumo`, `producción`, `disposición`) or
appended parentheticals (`production (certified biodegradable bags)`), which the prompt's
fixed vocabulary does not allow.

**Fix:** translate/normalize every Col L entry to the five fixed English lower-case terms only;
move any qualifying detail into Col R/U/W instead.

### 6. A handful of scale-value typos

- **Col G = 0.2** instead of 0.25 in 4 Norway strategy/plan files (`4P_Index_Noregs_Plaststrategi.xlsx`,
  `4P_Index_Handlingsplan_Sirkulaer.xlsx`, `4P_Index_Meld_St_45_Kap7.xlsx`,
  `4P_Index_Sirkulaerokonomistrategi.xlsx`) — looks like the 0.20 instrument-type scale bled into
  the policy-type field.
- **Col P = 0.75 / 0.5 / 0.3** (not on the {0, 0.20, 0.40, 0.60, 0.80, 1.0} scale) in 5 Peru
  files (`Peru_PLANRES…`, `Peru_DS_001_2022_MINAM…` used from two branches,
  `Peru_DS_014_2017_MINAM…`, `Peru_HRNEC_2030…`).
- **Col M = 0.75** (not on the {0, 0.5, 1} scale) in 5 files, mostly Mozambique — this
  duplicates an issue already flagged in the earlier `Policy_Coding_Scoring_Review.pptx`
  (PR #112) that has not yet been corrected in the underlying spreadsheets.

## Rule 12 — "may" vs "shall" in-force test (needs a second, human pass)

The heuristic flagged **~53 files / ~109 rows** (mostly Nepal, Mozambique, Norway, and a few
Peru) where the description/implementation text and the S value look inconsistent:

- **S=1 flagged for re-check (44 rows):** the text only contains enabling-power language
  ("may…", "is authorised to…") with no mandatory language and no stated evidence that a
  subordinate regulation has already operationalised the power. Per Rule 12 these should
  usually be S=0 unless the file's own comments (Col W) already show the power has been
  exercised.
- **S=0 flagged for re-check (14 rows):** the opposite pattern — mandatory language ("shall",
  "must", "is prohibited") with no enabling-power language, so S is more likely supposed to be 1.

These are **not** auto-flagged as hard errors because the correct answer genuinely depends on
context (e.g. a provision can legitimately mix "may" and "shall" clauses). See
`findings_detail.csv` filtered on `finding` containing "Per Rule 12" for the full row-by-row
list, each with the exact quoted text driving the flag — re-read each one against the "Tip —
language test" in the prompt and correct S where warranted.

## Rule 11 — multi-level coordination instruments (needs a second, human pass)

**~39 files / ~73 rows** describe an instrument that assigns/coordinates/delegates
responsibilities across levels of government (national–regional–municipal, inter-institutional
coordination, etc.) but score Col P at something other than 0.40, with no note in Col W
explaining why a higher-scoring instrument type was chosen instead. Per Rule 11, these should
either be re-scored to 0.40, or kept at the higher score with an explicit "coordination aspect"
note added to Col W (as the worked example in the prompt itself does for combined
instruments). See `findings_detail.csv` filtered on `finding` containing "Per Rule 11".

## Other WARN-level items worth a look

- **1,331 rows** where the number of comma-separated sectors in Col J doesn't fall inside the
  bucket implied by Col I's score (e.g. I=0.75 implies 5–6 sectors, but J lists a different
  count) — often because J contains stray text (mis-pasted justification prose) rather than a
  clean sector list, or because I wasn't recomputed after J was edited.
- **599 rows** with the analogous K (policy_circularity) vs L (phase count) mismatch.
- A number of Col J sector entries are not lower-case as the coding rules require.

## Caveats / what this audit does *not* replace

- The **ERROR** items are objective violations of explicit, fixed rules in the prompt (allowed
  value sets, required formulas, required vocab/casing, required invariance) — these are safe to
  treat as "definitely fix."
- The **WARN** items are consistency heuristics that can occasionally be a false alarm (e.g. a
  sector list that legitimately mixes bilingual synonyms) — spot-check before batch-fixing.
- The **REVIEW** items (Rule 11/12) are keyword heuristics only. They are a *prioritized list of
  where to look*, not a verdict — every flagged row still needs a human to re-read the cited
  quote and apply the "Tip — language test" from the prompt.
- This audit does **not** re-derive the *substance* of each coding (e.g. whether a given
  instrument really deserves P=1.0 vs 0.8, or whether a sector/lifecycle list is exhaustive per
  Rule 9) — that requires reading each source policy text, which is outside the scope of a
  mechanical/heuristic pass across ~90 branches. The `Policy_Coding_Scoring_Review.pptx` in PR
  #112 already captures several of those substantive judgement calls for Nepal/Mozambique/Peru/STP;
  none of its proposed fixes appear to have been applied to the underlying spreadsheets yet.
