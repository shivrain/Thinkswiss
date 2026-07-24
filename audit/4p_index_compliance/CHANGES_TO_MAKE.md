# 4P Index v2 — All Changes To Make

Action list generated from strict compliance check against the 4P Index Policy Coding Prompt v2.

Each file below lists **every row/column change** required. Severity: **MUST FIX** = definite rule violation; **REVIEW** = human judgement (Rules 11/12); **CHECK** = consistency warning.

## A. Policies to remove from the dataset entirely

From the July 2026 review presentation (PR #112) — delete these policies/rows, not just edit a cell:

### Nepal

- **DELETE** `The Excise Duty Act, 2002` — Not a plastic-relevant included policy

- **DELETE** `WTO Trade Policy Review — Report by the Secretariat — Nepal` — Non-official document; not a national policy instrument

- **DELETE** `Ban on Plastic Bottles in High-Class Hotels — Directive / Sectoral Order (2024)` — Source document not found

### Mozambique

- **DELETE** `Regulation on Environmental Quality and Effluent Standards` — Air-quality standard — not a relevant plastic instrument (Art. 8)

### São Tomé and Príncipe

- **DELETE** `National Plan for Integrated Urban Solid Waste Management (PNGIRSU) 2018–2023` — Projection/inventory — not an instrument

- **DELETE** `PNGIRSU 2018–2023 — National Waste Characterisation Exercise` — Data collection exercise — not an instrument


## B. Policy-level score corrections from review (PR #112)

### Mozambique — `National Strategy for the Management and Conservation of Coral Reefs`

- **Col M**: change `0.75` → `0.5 or 1.0` — Line-item budgets exist but 0.75 is not on the M scale

### Peru — `Supreme Decree No. 001-2022-MINAM`

- **Col M**: change `0.5` → `1.0` — Institutional financing mechanism + cleaning-fee agreements = ring-fenced budget source

### Peru — `Law that Amends Legislative Decree No. 1278`

- **Col M**: change `0.75` → `0.5` — Investment promotion mentioned but no specific budget line; 0.75 invalid

### Peru — `Article 9 of the Regulation of Legislative Decree No. 1278 (PLANRES Provision)`

- **Col E**: change `1` → `0` — PLANRES is periodic planning/reporting, not a quantifiable plastic-specific target

- **Col F**: change `(text)` → `(empty)` — Clear F when E=0


## C. File-by-file spreadsheet changes


### Mozambique — `Mozambique_4P_Index.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[REVIEW]** row 2, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 11, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 14, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 17, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 24, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 26, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 26, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 27, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 28, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 28, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 29, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 33, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 33, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 34, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 36, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 36, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 37, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 1 must-fix, 10 review, 25 check_


### Mozambique — `Mozambique_4P_Index_Colectanea_FDUEM.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[REVIEW]** row 11, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 16, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 19, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 20, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 22, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 23, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 23, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 24, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 24, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 25, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 26, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 27, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 27, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 28, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 29, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 29, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 30, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 31, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 31, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 32, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 33, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 34, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 34, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 34, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 35, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 36, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 37, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 38, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 39, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 40, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 41, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 42, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 43, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 44, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 45, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 46, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 47, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 48, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 49, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 50, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 51, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 51, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 52, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 53, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 5 must-fix, 10 review, 33 check_


### Mozambique — `Mozambique_4P_Index_Decreto_16_2015.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 15 check_


### Mozambique — `Mozambique_4P_Index_Decreto_79_2017.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 5, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 9, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 16, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 19, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 0 must-fix, 5 review, 18 check_


### Mozambique — `Mozambique_4P_Index_ESMF_CRRN.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 4 must-fix, 0 review, 13 check_


### Mozambique — `Mozambique_4P_Index_Lei_20_97.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[REVIEW]** row 11, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 16, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 19, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 20, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 1 must-fix, 4 review, 0 check_


### Mozambique — `Mozambique_4P_Index_Resolucao_5_95.xlsx`

Branch: `cursor/mozambique-plastic-policies-7c97`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 9, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 24, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 1 must-fix, 2 review, 23 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_ANAC_Tourism_Concessions_Manual_2012.xlsx`

Branch: `cursor/mozambique-anac-tourism-concessions-manual-2012-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 3 must-fix, 0 review, 12 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_18_2004.xlsx`

Branch: `cursor/mozambique-decree-18-2004-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 2, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 1 review, 14 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_23_2008_LOT_Regulation.xlsx`

Branch: `cursor/mozambique-decree-23-2008-lot-regulation-04b6`

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).


_Summary: 3 must-fix, 0 review, 0 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_45_2006.xlsx`

Branch: `cursor/mozambique-decree-45-2006-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 8 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_45_2024.xlsx`

Branch: `cursor/mozambique-decree-45-2024-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 0 must-fix, 0 review, 10 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_51_2024.xlsx`

Branch: `cursor/mozambique-decree-51-2024-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 2, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 3, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 0 must-fix, 3 review, 8 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_54_2015.xlsx`

Branch: `cursor/mozambique-decree-54-2015-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 1 must-fix, 0 review, 6 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_94_2014.xlsx`

Branch: `cursor/mozambique-decree-94-2014-4p-index-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 20 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Decree_97_2020_Coastal_Zones.xlsx`

Branch: `cursor/mozambique-decree-97-2020-coastal-zones-04b6`

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 0 must-fix, 1 review, 0 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Law_19_2007_Territorial_Planning.xlsx`

Branch: `cursor/mozambique-law-19-2007-territorial-planning-04b6`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 1 must-fix, 0 review, 5 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Law_20_2019_Sea_Law.xlsx`

Branch: `cursor/mozambique-law-20-2019-sea-law-04b6`

- **[REVIEW]** row 6, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 0 must-fix, 1 review, 0 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_NDC_3_0_Provisional.xlsx`

Branch: `cursor/mozambique-ndc-3-0-4p-index-04b6`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 14, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 0 must-fix, 1 review, 13 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Resolution_13_2021_Health_Policy.xlsx`

Branch: `cursor/mozambique-resolution-13-2021-health-policy-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[REVIEW]** row 2, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 3, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 1 must-fix, 2 review, 6 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Resolution_51_2022_Coral_Reef.xlsx`

Branch: `cursor/mozambique-resolution-51-2022-coral-reef-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).


_Summary: 6 must-fix, 1 review, 12 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Resolution_53_2024_EDEA.xlsx`

Branch: `cursor/mozambique-resolution-53-2024-edea-04b6`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).


_Summary: 6 must-fix, 0 review, 6 check_


### Mozambique — `output/mozambique/4P_Index_Mozambique_Resolution_7_2021_PNDT.xlsx`

Branch: `cursor/mozambique-resolution-7-2021-pndt-04b6`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 1 must-fix, 1 review, 12 check_


### Nepal — `4p_index_16th_plan_nepal_2024.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_bbin_mpa_esmf_2022.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_bnp_management_plan_2022.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_eflg_framework_2013.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 6, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_everest_singleuse_plastic_ban_2020.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_ghodaghodi_lake_masterplan_2020.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_hcwm_guideline_2014.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 9, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 1 must-fix, 2 review, 0 check_


### Nepal — `4p_index_karnali_tourism_masterplan.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_knp_management_plan_2024.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_lgoa_2074_2025.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_nepal_7nr_cbd_2026.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_nswmp_2079_2022.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[REVIEW]** row 5, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[REVIEW]** row 6, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 4 review, 3 check_


### Nepal — `4p_index_pb_directive_2082_2026.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_plastic_bags_action_plan_2022.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_plastic_flowers_ban_2022.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 2 review, 0 check_


### Nepal — `4p_index_school_sector_emf_2009.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_ssdp_emf_2017.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_swm_act_2068_2019.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 1 must-fix, 1 review, 0 check_


### Nepal — `4p_index_tourism_health_protocol_2020.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.


_Summary: 1 must-fix, 0 review, 0 check_


### Nepal — `4p_index_wto_tpr_nepal_2025.xlsx`

Branch: `cursor/4p-index-karnali-tourism-masterplan-376f`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 6, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 1 must-fix, 2 review, 0 check_


### Norway — `4P_Index_Avfallsforskrift_Kap10a.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap10a-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 6, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[REVIEW]** row 8, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).


_Summary: 62 must-fix, 2 review, 20 check_


### Norway — `4P_Index_Avfallsforskrift_Kap7a.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap7a-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 22, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 22, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 95 must-fix, 1 review, 42 check_


### Norway — `4P_Index_Avfallsforskrift_Kap7b.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap7b-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.25)/2.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 19 must-fix, 1 review, 0 check_


### Norway — `4P_Index_Avfallsforskriften.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 13, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 19, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 63 must-fix, 2 review, 18 check_


### Norway — `4P_Index_Avfallsforskriften_Kap1.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap1-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 42 must-fix, 2 review, 0 check_


### Norway — `4P_Index_Avfallsforskriften_Kap15.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap15-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.


_Summary: 7 must-fix, 0 review, 6 check_


### Norway — `4P_Index_Avfallsforskriften_Kap4.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap4-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 38 must-fix, 1 review, 13 check_


### Norway — `4P_Index_Avfallsforskriften_Kap5.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap5-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+0.75+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).


_Summary: 39 must-fix, 0 review, 20 check_


### Norway — `4P_Index_Avfallsforskriften_Kap6.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap6-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.4+0.25)/2.

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).


_Summary: 51 must-fix, 0 review, 0 check_


### Norway — `4P_Index_Avfallsforskriften_Kap7.xlsx`

Branch: `cursor/4p-index-avfallsforskrift-kap7-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 59 must-fix, 1 review, 16 check_


### Norway — `4P_Index_Baerekraftige_Produkter_Loven.xlsx`

Branch: `cursor/4p-index-baerekraftige-produkter-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 4, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 15, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 16, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 19, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 118 must-fix, 4 review, 40 check_


### Norway — `4P_Index_Drikkevareemballasje.xlsx`

Branch: `cursor/4p-index-drikkevareemballasje-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 11, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 12, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 13 must-fix, 3 review, 0 check_


### Norway — `4P_Index_Forurensningsforskrift.xlsx`

Branch: `cursor/4p-index-forurensningsforskrift-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 62 must-fix, 1 review, 18 check_


### Norway — `4P_Index_Forurensningsforskriften_Kap20.xlsx`

Branch: `cursor/4p-index-forurensningsforskrift-kap20-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 54 must-fix, 0 review, 0 check_


### Norway — `4P_Index_Forurensningsforskriften_Kap23A.xlsx`

Branch: `cursor/4p-index-forurensningsforskrift-kap23a-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 4, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[REVIEW]** row 18, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'waste management' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).


_Summary: 101 must-fix, 1 review, 20 check_


### Norway — `4P_Index_Forurensningsloven.xlsx`

Branch: `cursor/4p-index-forurensningsloven-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 7, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 8, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.5)/2.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 18, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 19, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 20, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 24, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 29 must-fix, 7 review, 0 check_


### Norway — `4P_Index_Handlingsplan_Sirkulaer.xlsx`

Branch: `cursor/4p-index-handlingsplan-sirkulaer-65aa`

- **[MUST FIX]** row 2, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 5, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 6, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[REVIEW]** row 7, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 8, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[REVIEW]** row 8, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 9, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 10, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 11, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 12, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 14, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 16, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 20, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 21, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 22, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 22, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 23, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 23, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 23, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 24, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 24, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 24, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 25, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 25, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 25, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.


_Summary: 144 must-fix, 3 review, 48 check_


### Norway — `4P_Index_Meld_St_45_Kap7.xlsx`

Branch: `cursor/4p-index-meld-st-45-kap7-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 2, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 3, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 3, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[REVIEW]** row 3, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 4, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 5, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 6, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[REVIEW]** row 6, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 7, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[REVIEW]** row 7, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 8, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 9, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 10, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 11, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 12, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 13, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 14, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 14, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[REVIEW]** row 14, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 15, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.5)/2.

- **[MUST FIX]** row 16, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 16, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 17, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 17, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 18, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 18, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 19, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.2+1.0+0.75+0.75)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.


_Summary: 131 must-fix, 4 review, 18 check_


### Norway — `4P_Index_Noregs_Plaststrategi.xlsx`

Branch: `cursor/4p-index-noregs-plaststrategi-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 3, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 4, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 7, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 8, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 9, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 10, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 12, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 14, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 14, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 15, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 16, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 18, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 18, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 19, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 20, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 21, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 21, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 22, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 22, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 23, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 23, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 23, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 24, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 24, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 25, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 25, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 25, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 26, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 26, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 26, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 27, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 27, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 27, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 27, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 27, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[REVIEW]** row 27, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 28, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 28, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 28, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 28, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 28, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 29, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 29, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 29, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 29, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 29, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 30, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 30, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 30, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 30, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 30, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.


_Summary: 173 must-fix, 2 review, 58 check_


### Norway — `4P_Index_Produktforskrift.xlsx`

Branch: `cursor/4p-index-produktforskrift-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 13, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'use' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 17, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 67 must-fix, 3 review, 16 check_


### Norway — `4P_Index_Produktforskrift_Kap2b.xlsx`

Branch: `cursor/4p-index-produktforskrift-kap2b-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Design' → 'Production' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 15, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 78 must-fix, 2 review, 0 check_


### Norway — `4P_Index_Produktforskrift_SUP_3200.xlsx`

Branch: `cursor/4p-index-produktforskrift-sup-3200-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.


_Summary: 40 must-fix, 0 review, 0 check_


### Norway — `4P_Index_Produktkontrolloven.xlsx`

Branch: `cursor/4p-index-excel-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.5)/2.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[REVIEW]** row 12, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[REVIEW]** row 15, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (1.0+1.0+1.0+0.0)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.


_Summary: 30 must-fix, 2 review, 0 check_


### Norway — `4P_Index_REACH_forskrift.xlsx`

Branch: `cursor/4p-index-reach-forskrift-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[REVIEW]** row 9, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.


_Summary: 56 must-fix, 1 review, 0 check_


### Norway — `4P_Index_Sirkulaerokonomistrategi.xlsx`

Branch: `cursor/4p-index-sirkulaerokonomistrategi-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[REVIEW]** row 3, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 4, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 6, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 7, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 8, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 9, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 10, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 12, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[REVIEW]** row 12, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 13, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[MUST FIX]** row 14, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 14, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 15, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 16, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 16, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 17, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 17, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 18, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 18, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 18, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 18, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.75)/2.

- **[REVIEW]** row 18, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 19, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 19, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 19, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 20, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 20, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 20, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 21, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 21, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 21, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 21, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 22, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 22, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 22, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 22, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.75)/2.

- **[MUST FIX]** row 23, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 23, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 23, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.75)/2.

- **[REVIEW]** row 23, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 24, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 24, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 24, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 25, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 25, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 25, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 26, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 26, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 26, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 27, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 27, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 27, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 27, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 27, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 27, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 28, Col G: Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2).

- **[CHECK]** row 28, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 28, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 28, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.2+1.0+1.0+0.5)/4 = 0.675. and apply to ALL rows of this policy.

- **[MUST FIX]** row 28, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.


_Summary: 167 must-fix, 4 review, 54 check_


### Norway — `4P_Index_Skipssikkerhet_MARPOL_V.xlsx`

Branch: `cursor/4p-index-skipssikkerhet-marpol-v-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 9, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'End-of-life' → 'End of life' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 12, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Use/Consumption' → 'Consumption' (Rule 4).

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[MUST FIX]** row 16, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[MUST FIX]** row 16, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[REVIEW]** row 16, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 17, Col L: Replace 'use/consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'litter/pollution' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'end-of-life' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+0.75+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Litter/Pollution' → 'Environmental leakage' (Rule 4).

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 83 must-fix, 2 review, 0 check_


### Norway — `4P_Index_Stortingsvedtak_Drikkevareemballasje_2026.xlsx`

Branch: `cursor/4p-index-stortingsvedtak-2026-65aa`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.5)/2.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.25)/2.


_Summary: 18 must-fix, 0 review, 0 check_


### Peru — `output/peru-dl-1278/Peru_DL_1278_4P_Index_Coding.xlsx`

Branch: `cursor/peru-dl-1278-4p-index-d350`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 10, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 27, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 31, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 32, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 34, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 35, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 36, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 36, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).


_Summary: 16 must-fix, 4 review, 0 check_


### Peru — `output/peru-ds-001-2022-minam/Peru_DS_001_2022_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ley-32212-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 8, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).


_Summary: 142 must-fix, 1 review, 16 check_


### Peru — `output/peru-ds-002-2024-minam/Peru_DS_002_2024_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ley-32212-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'planificación (ciclo de vida completo)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición / planning (whole life cycle)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).


_Summary: 81 must-fix, 1 review, 16 check_


### Peru — `output/peru-ds-003-2025-minam/Peru_DS_003_2025_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-003-2025-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición (marco estratégico ciclo de vida completo) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposal (whole-life-cycle strategic framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).


_Summary: 133 must-fix, 1 review, 24 check_


### Peru — `output/peru-ds-005-2010-minam/Peru_DS_005_2010_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-dl-1278-4p-index-d350`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 2, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 24, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 30, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 31, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 32, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 33, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 35, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 36, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 35 must-fix, 2 review, 0 check_


### Peru — `output/peru-ds-006-2019-minam/Peru_DS_006_2019_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-006-2019-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 7, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 9, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.25)/2.

- **[REVIEW]** row 10, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 14, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 15, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 16, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 17, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 18, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[CHECK]** row 19, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 20, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 21, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 21, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 22, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 22, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 23, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 23, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 23, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 24, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 24, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 25, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 25, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 26, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 26, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 27, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 27, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 27, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 28, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 28, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 28, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 28, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.4+0.25)/2.

- **[CHECK]** row 29, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 29, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 29, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 30, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 30, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 30, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 30, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 30, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 31, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 31, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 31, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 31, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 31, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 31, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 32, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 32, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 32, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 32, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 32, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 33, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 33, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 33, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 33, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 33, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 33, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.4+0.25)/2.

- **[CHECK]** row 34, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 34, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'compostaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 34, Col L: Replace 'composting' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 34, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 34, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 34, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 339 must-fix, 5 review, 66 check_


### Peru — `output/peru-ds-009-2019-minam/Peru_DS_009_2019_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-009-2019-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 13, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 14, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 15, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 16, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 17, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 17, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 18, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 18, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 19, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 20, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 21, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 21, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 22, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 22, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 22, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 23, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 23, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 24, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 24, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 24, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[REVIEW]** row 24, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 25, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 25, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 25, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 26, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 26, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'valorización material' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'disposición / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'material valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.


_Summary: 297 must-fix, 5 review, 50 check_


### Peru — `output/peru-ds-013-2018-minam/Peru_DS_013_2018_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-013-2018-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+0.5+0.75+0.5)/4 = 0.625. and apply to ALL rows of this policy.


_Summary: 67 must-fix, 0 review, 26 check_


### Peru — `output/peru-ds-014-2017-minam/Peru_DS_014_2017_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-014-2017-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[REVIEW]** row 3, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 22, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[REVIEW]** row 23, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 24, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 25, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 26, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 26, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 27, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 27, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 27, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 27, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 28, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 28, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 29, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 29, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 29, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 30, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 30, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 30, Col O: Recalculate Col O = (0.75+1.0+1.0+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 30, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[REVIEW]** row 30, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 193 must-fix, 3 review, 29 check_


### Peru — `output/peru-ds-023-2021-minam/Peru_DS_023_2021_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-023-2021-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 16, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición (marco estratégico ciclo de vida) / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposal (strategic whole-life-cycle framework)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.75+1.0+1.0+0.5)/4 = 0.812. and apply to ALL rows of this policy.


_Summary: 197 must-fix, 2 review, 19 check_


### Peru — `output/peru-ds-244-2019-ef/Peru_DS_244_2019_EF_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ds-244-2019-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'producción (bolsas biodegradables certificadas)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'production (certified biodegradable bags)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.5)/2.

- **[REVIEW]** row 12, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 12, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 58 must-fix, 3 review, 33 check_


### Peru — `output/peru-hrnec-2030/Peru_HRNEC_2030_4P_Index_Coding.xlsx`

Branch: `cursor/peru-hrnec-2030-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 4, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 7, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 10, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 12, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 14, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 14, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 15, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 15, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 16, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 16, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 17, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 17, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 18, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 18, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.5+1.0+1.0+0.5)/4 = 0.750. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 19, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 19, Col L: Replace 'diseño' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).


_Summary: 235 must-fix, 1 review, 36 check_


### Peru — `output/peru-law-29419/Peru_Law_29419_4P_Index_Coding.xlsx`

Branch: `cursor/peru-law-29419-4p-index-d350`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 2, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 14, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 15, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 15, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.


_Summary: 17 must-fix, 5 review, 0 check_


### Peru — `output/peru-ley-30884/Peru_Law_30884_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ley-30884-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (1.0+1.0+1.0+0.5)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.2+0.25)/2.


_Summary: 122 must-fix, 0 review, 20 check_


### Peru — `output/peru-ley-31896/Peru_Ley_31896_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ley-31896-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 7, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición / recycling' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+0.75+1.0+0.75)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 43 must-fix, 2 review, 7 check_


### Peru — `output/peru-ley-32212/Peru_Ley_32212_4P_Index_Coding.xlsx`

Branch: `cursor/peru-ley-32212-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 3, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 6, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 10, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 11, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 14, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 15, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[CHECK]** row 16, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col Q: Change Col Q from 'Planning' → 'Production (or add W note if purely procedural)' (Rule 4).

- **[CHECK]** row 17, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 17, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[REVIEW]** row 17, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 18, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col M: Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid).

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (1.0+1.0+1.0+0.75)/4 = 0.938. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[REVIEW]** row 18, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 202 must-fix, 4 review, 34 check_


### Peru — `output/peru-planres/Peru_PLANRES_4P_Index_Coding.xlsx`

Branch: `cursor/peru-planres-4p-index-d350`

- **[MUST FIX]** whole file, structure: Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly.

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 2, Col P: Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only.

- **[REVIEW]** row 2, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 3, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 9, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 10, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col T: Fix Col T (instrument_implementation): instrument_implementation (T)=0.2 is not one of the allowed values {0, 0.25, 0.5, 0.75, 1}.

- **[MUST FIX]** row 11, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 13, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 16, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 16, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 16, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 16, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 17, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 17, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 17, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 17, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 18, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 18, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 18, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 18, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 19, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 19, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 19, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 19, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 20, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 20, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 20, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 20, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 21, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 21, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 21, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 21, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 21, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 22, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 22, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 22, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 22, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 22, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 23, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 23, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 23, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 23, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 23, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 24, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 24, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 24, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 24, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 24, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 25, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 25, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 25, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 25, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 25, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 26, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 26, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 26, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 26, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 26, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 27, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 27, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 27, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 27, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 28, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 28, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 28, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 28, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 29, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 29, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 29, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 29, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[REVIEW]** row 29, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[MUST FIX]** row 30, Col L: Replace 'producción' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'disposición' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 30, Col L: Replace 'fuga ambiental / production' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 30, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 30, Col O: Recalculate Col O = (0.5+1.0+1.0+1.0)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 182 must-fix, 6 review, 29 check_


### Peru — `output/peru-rm-122-2021-minam/Peru_RM_122_2021_MINAM_4P_Index_Coding.xlsx`

Branch: `cursor/peru-rm-122-2021-4p-index-d350`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 2, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 2, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 2, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 3, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 3, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[REVIEW]** row 3, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 4, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 4, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 5, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 5, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 6, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 6, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 6, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 6, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 6, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 7, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 7, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 7, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 7, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 7, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 8, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 8, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 8, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 8, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 8, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col Q: Change Col Q from 'Collection' → 'Waste management' (Rule 4).

- **[CHECK]** row 9, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 9, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 9, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 9, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 9, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 10, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 10, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 10, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 10, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 10, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[REVIEW]** row 10, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 11, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 11, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 11, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 11, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 11, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 12, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 12, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 12, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 12, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 12, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[REVIEW]** row 12, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[CHECK]** row 13, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 13, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 13, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 13, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 13, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 14, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 14, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 14, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 14, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 14, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[CHECK]** row 15, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 15, Col J: Use lower case only for all sector entries in Col J.

- **[MUST FIX]** row 15, Col L: Replace 'consumo' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'recolección' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'reciclaje' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorización' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposición (cambio de comportamiento) / consumption' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'collection' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'valorization' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[MUST FIX]** row 15, Col L: Replace 'disposal (behavior change)' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3).

- **[CHECK]** row 15, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.5+0.75+1.0+0.5)/4 = 0.688. and apply to ALL rows of this policy.


_Summary: 129 must-fix, 3 review, 42 check_


### Senegal — `4P_Index_Arrete_07022_POLMAR.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[REVIEW]** row 3, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 5, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 0 must-fix, 3 review, 0 check_


### Senegal — `4P_Index_Arretes_EIA_9468-9472_2001.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 0 must-fix, 1 review, 0 check_


### Senegal — `4P_Index_Decret_2008-1007_Dechets_Biomedicaux.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 0 must-fix, 0 review, 4 check_


### Senegal — `4P_Index_Decret_2016-1804_Code_Peche_Maritime.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 2, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 1 review, 3 check_


### Senegal — `4P_Index_Decret_2021-1115_CENPOLMAR.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[MUST FIX]** whole file, structure: Fix Col  (): No data rows found.


_Summary: 1 must-fix, 0 review, 0 check_


### Senegal — `4P_Index_Lettre_Politique_Environnement_2016-2020.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 3, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 4, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.

- **[CHECK]** row 5, Col L: Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list.


_Summary: 0 must-fix, 0 review, 4 check_


### Senegal — `4P_Index_Loi_2009-24_Code_Assainissement.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 2, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 1 review, 4 check_


### Senegal — `4P_Index_Loi_2013-10_CGCT.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 1 review, 4 check_


### Senegal — `4P_Index_Loi_2015-09_Sachets_Plastiques.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[REVIEW]** row 4, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 5, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[REVIEW]** row 6, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 0 must-fix, 3 review, 0 check_


### Senegal — `4P_Index_Loi_2015-18_Code_Peche_Maritime.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 2, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 3, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 0 must-fix, 3 review, 3 check_


### Senegal — `4P_Index_Loi_2022-18_SONAGED.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 3 check_


### Senegal — `4P_Index_NS_05-061_Eaux_Usees.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 3 check_


### Senegal — `4P_Index_Plan_Gestion_Dechets_Biomedicaux_2019.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[CHECK]** row 2, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 3, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 4, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.

- **[CHECK]** row 5, Col J: Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list.


_Summary: 0 must-fix, 0 review, 4 check_


### Senegal — `4P_Index_PROMOGED_Dechets_Solides.xlsx`

Branch: `cursor/senegal-plastic-policies-faa3`

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).


_Summary: 0 must-fix, 1 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Decreto-36-1999_Waste-Decree_4P-Index.xlsx`

Branch: `cursor/code-stp-waste-decree-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[REVIEW]** row 9, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.0)/2.

- **[REVIEW]** row 12, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 13, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[REVIEW]** row 14, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 15, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.

- **[MUST FIX]** row 15, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.25)/2.

- **[MUST FIX]** row 16, Col O: Recalculate Col O = (0.75+0.75+0.75+1.0)/4 = 0.812. and apply to ALL rows of this policy.


_Summary: 19 must-fix, 3 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Decreto-37-1999_EIA-Decree_4P-Index.xlsx`

Branch: `cursor/code-stp-eia-decree-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (0.75+1.0+0.75+1.0)/4 = 0.875. and apply to ALL rows of this policy.


_Summary: 9 must-fix, 0 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Lei-10-1999_Basic-Environmental-Law_4P-Index.xlsx`

Branch: `cursor/code-stp-basic-environmental-law-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[REVIEW]** row 3, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col Q: Change Col Q from 'Disposal' → 'Waste management or End of life (pick closest fit)' (Rule 4).

- **[MUST FIX]** row 9, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (0.6+0.25)/2.

- **[REVIEW]** row 9, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.

- **[MUST FIX]** row 10, Col V: Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = (1.0+0.0)/2.

- **[REVIEW]** row 10, Col S: Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W.


_Summary: 7 must-fix, 3 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Lei-4-2023_Customs-Duties-Law_4P-Index.xlsx`

Branch: `cursor/code-stp-customs-duties-law-3cda`

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (1.0+0.75+0.25+0.0)/4 = 0.500. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (1.0+0.75+0.25+0.0)/4 = 0.500. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (1.0+0.75+0.25+0.0)/4 = 0.500. and apply to ALL rows of this policy.


_Summary: 3 must-fix, 0 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Lei-8-2020_Plastic-Bag-Law_4P-Index.xlsx`

Branch: `cursor/code-stp-plastic-bag-law-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 9, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 12, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.

- **[MUST FIX]** row 14, Col O: Recalculate Col O = (1.0+1.0+1.0+1.0)/4 = 1.000. and apply to ALL rows of this policy.


_Summary: 8 must-fix, 0 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_Lei-9-2022_Fisheries-Aquaculture-Law_4P-Index.xlsx`

Branch: `cursor/code-stp-fisheries-law-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[REVIEW]** row 4, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[REVIEW]** row 4, Col P: Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11).

- **[REVIEW]** row 6, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 10, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 10, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 11, Col O: Recalculate Col O = (1.0+1.0+0.75+1.0)/4 = 0.938. and apply to ALL rows of this policy.

- **[REVIEW]** row 11, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.


_Summary: 3 must-fix, 5 review, 0 check_


### São Tomé and Príncipe — `4p-index/STP_PNGIRSU-2018-2023_4P-Index.xlsx`

Branch: `cursor/code-stp-pngirsu-3cda`

- **[MUST FIX]** row 2, Col O: Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row.

- **[MUST FIX]** row 2, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 3, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 4, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 5, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 6, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[MUST FIX]** row 7, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.

- **[REVIEW]** row 7, Col S: Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W.

- **[MUST FIX]** row 8, Col O: Recalculate Col O = (0.5+1.0+0.75+0.5)/4 = 0.688. and apply to ALL rows of this policy.


_Summary: 8 must-fix, 1 review, 0 check_


---

**Total files with changes: 109**
