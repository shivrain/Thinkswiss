# 4P Index Coding Compliance Audit

Automated + heuristic audit of every `4P Index` coding spreadsheet found across all repository branches, checked against the 4P Index Policy Coding Prompt (v2).

- Files audited: 119
- ERROR findings (hard rule violations): 4439
- WARN findings (count/consistency mismatches): 1930
- REVIEW findings (Rule 11/12 heuristic flags needing human judgement): 182


## `output/peru-ds-244-2019-ef/Peru_DS_244_2019_EF_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-244-2019-4p-index-d350`)

- **[ERROR] Col L** — rows 2-12 (n=11)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-12 (n=11)
  - policy_lifecycle_phases_list (L) contains 'producción (bolsas biodegradables certificadas)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-12 (n=11)
  - policy_lifecycle_phases_list (L) contains 'disposición / consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-12 (n=11)
  - policy_lifecycle_phases_list (L) contains 'production (certified biodegradable bags)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]).

- **[ERROR] Col O** — rows [2, 3, 7, 12]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col O** — rows [4]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col O** — rows [5, 6, 8, 9, 10, 11]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col V** — rows [7]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [12]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[WARN] Col J** — rows 2-12 (n=11)
  - policy_sectors_list (J) lists 16 sector(s) ('comercio minorista, retail, administración tributaria (SUNAT), MEF, alimentos y bebidas, servicios, comercio electrónico, manufactura de bolsas plásticas, importación de empaques / retail commerce, tax administration (SUNAT), MEF, food & beverage, services, e-commerce, plastic bag manufacturing, packaging imports'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows 2-12 (n=11)
  - policy_sectors_list (J) entry 'administración tributaria (SUNAT)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-12 (n=11)
  - policy_sectors_list (J) entry 'MEF' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-12 (n=11)
  - policy_sectors_list (J) entry 'tax administration (SUNAT)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-12 (n=11)
  - policy_lifecycle_phases_list (L) lists 5 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.

- **[REVIEW] Col P** — rows [7, 12]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [12]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/peru-ds-005-2010-minam/Peru_DS_005_2010_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-dl-1278-4p-index-d350`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['B — policy_name_en', 'C — country']), which shifts every subsequent column's letter out of alignment. For example: file labels its 'policy_url' column as 'D' (should be 'B' per the prompt); file labels its 'policy_year' column as 'E' (should be 'C' per the prompt); file labels its 'policy_objective' column as 'F' (should be 'D' per the prompt); file labels its 'policy_target' column as 'G' (should be 'E' per the prompt); file labels its 'policy_target_text' column as 'H' (should be 'F' per the prompt); file labels its 'policy_type' column as 'I' (should be 'G' per the prompt). Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.78, 0.82, 0.86, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36]).

- **[ERROR] Col O** — rows 2-35 (n=18)
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [3, 5, 6, 9, 11, 18, 31, 32, 36]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [7, 27, 28, 33]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [26, 29]
  - policy_score (O)=0.78 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[REVIEW] Col P** — rows [2, 24]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ley-31896/Peru_Ley_31896_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ley-31896-4p-index-d350`)

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_lifecycle_phases_list (L) contains 'disposición / recycling', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col M** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_budget (M)=0.75 is not one of the allowed values {0, 0.5, 1}.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.78, 0.82]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8]).

- **[ERROR] Col O** — rows [2]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (1.0+0.75+1.0+0.75)/4 = 0.875.

- **[ERROR] Col O** — rows [3, 4, 5, 7]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (1.0+0.75+1.0+0.75)/4 = 0.875.

- **[ERROR] Col O** — rows [6, 8]
  - policy_score (O)=0.78 does not equal (G+I+K+M)/4 = (1.0+0.75+1.0+0.75)/4 = 0.875.

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7, 8]
  - policy_sectors_list (J) lists 13 sector(s) ('ambiente, industria, gobiernos regionales, gobiernos locales, inversión pública/privada, reciclaje, valorización / environment, industry, regional governments, local governments, public/private investment, recycling, valorization'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col P** — rows [4, 7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ds-003-2025-minam/Peru_DS_003_2025_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-003-2025-4p-index-d350`)

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'disposición (marco estratégico ciclo de vida completo) / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'disposal (whole-life-cycle strategic framework)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.68, 0.72]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]).

- **[ERROR] Col O** — rows [2, 12, 13]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [3, 6, 7, 8, 10, 11]
  - policy_score (O)=0.68 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [4, 5, 9]
  - policy_score (O)=0.72 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col Q** — rows 2-13 (n=12)
  - instrument_lifecycle_stage (Q)='Planning' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-13 (n=12)
  - policy_sectors_list (J) entry 'PCM); plásticos' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-13 (n=12)
  - policy_sectors_list (J) entry 'PCM); plastics' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) lists 11 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ds-023-2021-minam/Peru_DS_023_2021_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-023-2021-4p-index-d350`)

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'disposición (marco estratégico ciclo de vida) / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) contains 'disposal (strategic whole-life-cycle framework)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]).

- **[ERROR] Col O** — rows 2-20 (n=11)
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [6, 7, 8, 10, 13, 14, 17]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [9]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows [6, 7, 9, 10, 12]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [19]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[WARN] Col L** — rows 2-20 (n=19)
  - policy_lifecycle_phases_list (L) lists 11 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [4, 16]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/mozambique/4P_Index_Mozambique_Decree_51_2024.xlsx`  (branch: `cursor/mozambique-decree-51-2024-4p-index-04b6`)

- **[WARN] Col J** — rows [2, 3, 4, 5]
  - policy_sectors_list (J) lists 8 sector(s) ('waste management, industry, municipalities, fisheries, tourism, water, agriculture, packaging'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5]
  - policy_lifecycle_phases_list (L) lists 4 phase(s), inconsistent with policy_circularity (K)=0.5 bucket.

- **[REVIEW] Col S** — rows [2, 3, 4]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/peru-planres/Peru_PLANRES_4P_Index_Coding.xlsx`  (branch: `cursor/peru-planres-4p-index-d350`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['B — policy_name_en', 'C — country']), which shifts every subsequent column's letter out of alignment. For example: file labels its 'policy_url' column as 'D' (should be 'B' per the prompt); file labels its 'policy_year' column as 'E' (should be 'C' per the prompt); file labels its 'policy_objective' column as 'F' (should be 'D' per the prompt); file labels its 'policy_target' column as 'G' (should be 'E' per the prompt); file labels its 'policy_target_text' column as 'H' (should be 'F' per the prompt); file labels its 'policy_type' column as 'I' (should be 'G' per the prompt). Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]).

- **[ERROR] Col O** — rows [2]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows 3-30 (n=28)
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+1.0)/4 = 0.875.

- **[ERROR] Col P** — rows [2]
  - instrument_type (P)=0.5 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col Q** — rows [23, 24, 25, 26]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col T** — rows [10]
  - instrument_implementation (T)=0.2 is not one of the allowed values {0, 0.25, 0.5, 0.75, 1}.

- **[WARN] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) lists 9 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [2]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.5 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [9, 16, 21, 22, 29]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4p_index_pb_directive_2082_2026.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4p_index_school_sector_emf_2009.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4P_Index_Forurensningsloven.xlsx`  (branch: `cursor/4p-index-forurensningsloven-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24]).

- **[ERROR] Col O** — rows [2, 8, 19, 24]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows [3, 16, 20, 23]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows 4-22 (n=15)
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col V** — rows [6, 14, 15]
  - instrument_score (V)=0.875 does not equal S×(P+T)/2 = 0.0×(1.0+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [12]
  - instrument_score (V)=0.75 does not equal S×(P+T)/2 = 0.0×(1.0+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [23]
  - instrument_score (V)=0.55 does not equal S×(P+T)/2 = 0.0×(0.6+0.5)/2 = 0.000.

- **[REVIEW] Col P** — rows [19]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [4, 7, 8, 18, 20, 24]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Decret_2008-1007_Dechets_Biomedicaux.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col L** — rows [2, 3, 4, 5]
  - policy_lifecycle_phases_list (L) lists 3 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.


## `4P_Index_Avfallsforskriften_Kap6.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap6-65aa`)

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]).

- **[ERROR] Col O** — rows [2, 3, 4, 8, 12]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [5, 6, 7, 9, 10, 13, 14, 15]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [11]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col Q** — rows [3, 5, 9, 12, 13, 14, 15]
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [11]
  - instrument_score (V)=0.325 does not equal S×(P+T)/2 = 0.0×(0.4+0.25)/2 = 0.000.


## `4P_Index_Loi_2015-18_Code_Peche_Maritime.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4]
  - policy_sectors_list (J) lists 5 sector(s) ('maritime fisheries, artisanal fishing, industrial fishing, marine environment, maritime trade'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [2, 3, 4]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/peru-rm-122-2021-minam/Peru_RM_122_2021_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-rm-122-2021-4p-index-d350`)

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'disposición (cambio de comportamiento) / consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'disposal (behavior change)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.59, 0.63]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]).

- **[ERROR] Col O** — rows 2-15 (n=11)
  - policy_score (O)=0.59 does not equal (G+I+K+M)/4 = (0.5+0.75+1.0+0.5)/4 = 0.688.

- **[ERROR] Col O** — rows [4, 5, 7]
  - policy_score (O)=0.63 does not equal (G+I+K+M)/4 = (0.5+0.75+1.0+0.5)/4 = 0.688.

- **[ERROR] Col Q** — rows [5, 8]
  - instrument_lifecycle_stage (Q)='Collection' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-15 (n=14)
  - policy_sectors_list (J) lists 17 sector(s) ('ambiente, educación (MINEDU), gobiernos locales, sociedad civil, sector privado, municipalidades, reciclaje, consumo, comunicación / environment, education (MINEDU), local governments, civil society, private sector, municipalities, recycling, consumption, communication'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows 2-15 (n=14)
  - policy_sectors_list (J) entry 'educación (MINEDU)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-15 (n=14)
  - policy_sectors_list (J) entry 'education (MINEDU)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) lists 9 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [3, 10, 12]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Forurensningsforskriften_Kap23A.xlsx`  (branch: `cursor/4p-index-forurensningsforskrift-kap23a-65aa`)

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'waste management', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.72, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]).

- **[ERROR] Col O** — rows [2, 3, 4, 5, 18, 19, 20, 21]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [6, 7, 8, 9, 10, 11, 13, 14, 15, 17]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [12, 16]
  - policy_score (O)=0.72 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col Q** — rows 2-21 (n=19)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [4]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[WARN] Col J** — rows 2-21 (n=20)
  - policy_sectors_list (J) lists 6 sector(s) ('sports, municipalities, construction, leisure, facility management, waste management'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [18]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_18_2004.xlsx`  (branch: `cursor/mozambique-decree-18-2004-4p-index-04b6`)

- **[WARN] Col J** — rows 2-15 (n=14)
  - policy_sectors_list (J) lists 7 sector(s) ('industry, waste management, water, chemicals, municipalities, agriculture, fisheries'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [2]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `Mozambique_4P_Index_Lei_20_97.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[ERROR] Col Q** — rows [8]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[REVIEW] Col S** — rows [11]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.

- **[REVIEW] Col S** — rows [16, 19, 20]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4P_Index_Avfallsforskriften_Kap5.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap5-65aa`)

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.78, 0.82]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]).

- **[ERROR] Col O** — rows [2, 3, 4, 9]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+0.75+1.0+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [5]
  - policy_score (O)=0.78 does not equal (G+I+K+M)/4 = (0.75+0.75+1.0+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [6, 7, 8, 10, 11]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+0.75+1.0+1.0)/4 = 0.875.

- **[ERROR] Col Q** — rows [2, 3, 5, 6, 7, 9, 10, 11]
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
  - policy_sectors_list (J) lists 8 sector(s) ('automotive, transport, waste management, retail, producers, recycling, manufacturing, environmental leakage'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.


## `4p_index_knp_management_plan_2024.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `output/peru-ds-002-2024-minam/Peru_DS_002_2024_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ley-32212-4p-index-d350`)

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'planificación (ciclo de vida completo)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'disposición / planning (whole life cycle)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.77]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9]).

- **[ERROR] Col O** — rows [2]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [3, 4, 5, 8, 9]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [6, 7]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - instrument_lifecycle_stage (Q)='Planning' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_sectors_list (J) entry 'planificación pública (CEPLAN)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_sectors_list (J) entry 'public planning (CEPLAN)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_lifecycle_phases_list (L) lists 11 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `Mozambique_4P_Index_Colectanea_FDUEM.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[ERROR] Col Q** — rows [8, 21, 22, 23, 29]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [21, 22, 23, 24]
  - policy_sectors_list (J) lists 5 sector(s) ('waste management, health, municipalities, industry, environment'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [25]
  - policy_sectors_list (J) lists 5 sector(s) ('industry, environment, energy, transport, chemicals'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [26]
  - policy_sectors_list (J) lists 6 sector(s) ('fisheries, maritime, ports, environment, tourism, municipalities'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [27, 28]
  - policy_sectors_list (J) lists 5 sector(s) ('industry, environment, waste management, municipalities, production'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [29, 30]
  - policy_sectors_list (J) lists 6 sector(s) ('industry, agriculture, chemicals, waste management, municipalities, packaging'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [31, 32, 33, 34]
  - policy_sectors_list (J) lists 7 sector(s) ('waste management, municipalities, industry, retail, recycling, public health, environment'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows 35-49 (n=15)
  - policy_sectors_list (J) lists 8 sector(s) ('production, consumption, retail, waste management, industry, municipalities, packaging, fisheries'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [50]
  - policy_sectors_list (J) lists 7 sector(s) ('industry, infrastructure, mining, energy, tourism, environment, municipalities'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [51, 52, 53]
  - policy_sectors_list (J) lists 7 sector(s) ('packaging, industry, retail, municipalities, waste management, recycling, chemicals'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col P** — rows [24, 31]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [34]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [51]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [11, 27, 34]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.

- **[REVIEW] Col S** — rows [16, 19, 20]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4p_index_swm_act_2068_2019.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ley-30884/Peru_Law_30884_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ley-30884-4p-index-d350`)

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.78, 0.82, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]).

- **[ERROR] Col O** — rows [2, 13, 19, 21]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows 3-17 (n=12)
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows [4, 18]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows [14, 20]
  - policy_score (O)=0.78 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col V** — rows [21]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[WARN] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) lists 9 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.


## `4P_Index_Produktkontrolloven.xlsx`  (branch: `cursor/4p-index-excel-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.68, 0.72, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20]).

- **[ERROR] Col O** — rows [2, 3, 4, 5, 6, 8, 9, 10, 11, 17]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.0)/4 = 0.750.

- **[ERROR] Col O** — rows [7, 12, 13, 20]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.0)/4 = 0.750.

- **[ERROR] Col O** — rows [14, 18, 19]
  - policy_score (O)=0.72 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.0)/4 = 0.750.

- **[ERROR] Col O** — rows [15, 16]
  - policy_score (O)=0.68 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.0)/4 = 0.750.

- **[ERROR] Col V** — rows [5, 6, 10]
  - instrument_score (V)=0.875 does not equal S×(P+T)/2 = 0.0×(1.0+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [7]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [9]
  - instrument_score (V)=0.75 does not equal S×(P+T)/2 = 0.0×(1.0+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [13, 20]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [14, 18, 19]
  - instrument_score (V)=0.55 does not equal S×(P+T)/2 = 0.0×(0.6+0.5)/2 = 0.000.

- **[REVIEW] Col S** — rows [12, 15]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Avfallsforskriften.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-65aa`)

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.87, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]).

- **[ERROR] Col O** — rows [2, 10, 12, 17]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows 3-19 (n=12)
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col Q** — rows [2, 3, 4, 5, 6, 9, 13, 16, 18, 19]
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [13, 19]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/peru-ds-009-2019-minam/Peru_DS_009_2019_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-009-2019-4p-index-d350`)

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'diseño', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'valorización material', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'disposición / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) contains 'material valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]).

- **[ERROR] Col O** — rows [2, 3, 5, 6, 20, 21, 22, 23, 24, 25]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows 4-19 (n=12)
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [11, 16, 26]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows 2-25 (n=19)
  - instrument_lifecycle_stage (Q)='Collection' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [24, 25]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'salud (DIGESA)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'transportes y comunicaciones (MTC)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'certificación ambiental (SENACE)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'health (DIGESA)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'transport and communications (MTC)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-26 (n=25)
  - policy_sectors_list (J) entry 'environmental certification (SENACE)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-26 (n=25)
  - policy_lifecycle_phases_list (L) lists 13 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [13, 17]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [24]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [14, 18]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_eflg_framework_2013.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [6]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Avfallsforskriften_Kap4.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap4-65aa`)

- **[ERROR] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]).

- **[ERROR] Col O** — rows [2, 3, 6, 9, 10, 11, 12, 13]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+0.75+1.0+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [4, 5, 7, 8, 14]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+0.75+1.0+1.0)/4 = 0.875.

- **[ERROR] Col Q** — rows 2-14 (n=11)
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-14 (n=13)
  - policy_sectors_list (J) lists 7 sector(s) ('automotive, waste management, producers, recycling, manufacturing, retail, public administration'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [14]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_wto_tpr_nepal_2025.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [6]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Forurensningsforskriften_Kap20.xlsx`  (branch: `cursor/4p-index-forurensningsforskrift-kap20-65aa`)

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-15 (n=14)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]).

- **[ERROR] Col O** — rows [2, 3, 4, 6, 9, 10, 13, 14, 15]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [5, 7, 8, 11, 12]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col Q** — rows 3-14 (n=11)
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).


## `output/mozambique/4P_Index_Mozambique_Law_20_2019_Sea_Law.xlsx`  (branch: `cursor/mozambique-law-20-2019-sea-law-04b6`)

- **[REVIEW] Col P** — rows [6]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ds-001-2022-minam/Peru_DS_001_2022_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ley-32212-4p-index-d350`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'disposición / consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.77, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 17]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [3, 5, 8, 10, 14, 15]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [9, 11, 16]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col P** — rows [4, 6, 7, 12, 13]
  - instrument_type (P)=0.75 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col Q** — rows 2-17 (n=13)
  - instrument_lifecycle_stage (Q)='Collection' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) lists 9 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [8]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Law_19_2007_Territorial_Planning.xlsx`  (branch: `cursor/mozambique-law-19-2007-territorial-planning-04b6`)

- **[ERROR] Col Q** — rows [2]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows [2, 3, 4, 5, 6]
  - policy_lifecycle_phases_list (L) lists 2 phase(s), inconsistent with policy_circularity (K)=0.25 bucket.


## `4p-index/STP_Decreto-37-1999_EIA-Decree_4P-Index.xlsx`  (branch: `cursor/code-stp-eia-decree-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9]).

- **[ERROR] Col O** — rows [2, 4, 5, 6, 8]
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [3, 7]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [9]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.


## `4p-index/STP_Lei-10-1999_Basic-Environmental-Law_4P-Index.xlsx`  (branch: `cursor/code-stp-basic-environmental-law-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.84, 0.92, 1.0]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10]).

- **[ERROR] Col O** — rows [7]
  - policy_score (O)=0.84 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col O** — rows [8, 9]
  - policy_score (O)=0.92 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col Q** — rows [9]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [9]
  - instrument_score (V)=0.425 does not equal S×(P+T)/2 = 0.0×(0.6+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [10]
  - instrument_score (V)=0.5 does not equal S×(P+T)/2 = 0.0×(1.0+0.0)/2 = 0.000.

- **[REVIEW] Col S** — rows [3]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.

- **[REVIEW] Col S** — rows [9, 10]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/peru-ley-32212/Peru_Ley_32212_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ley-32212-4p-index-d350`)

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col M** — rows 2-18 (n=17)
  - policy_budget (M)=0.75 is not one of the allowed values {0, 0.5, 1}.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18]).

- **[ERROR] Col O** — rows [2]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.75)/4 = 0.938.

- **[ERROR] Col O** — rows [3, 5, 6, 9]
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.75)/4 = 0.938.

- **[ERROR] Col O** — rows [4, 7, 8, 11, 12, 14, 16]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.75)/4 = 0.938.

- **[ERROR] Col O** — rows [10, 13, 15, 17, 18]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.75)/4 = 0.938.

- **[ERROR] Col P** — rows [3, 5, 6, 9]
  - instrument_type (P)=0.75 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col Q** — rows [2, 3, 6, 9, 11, 13, 14, 16, 17, 18]
  - instrument_lifecycle_stage (Q)='Planning' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-18 (n=17)
  - policy_sectors_list (J) entry 'FONAM/PROFONANPE / environment' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-18 (n=17)
  - policy_sectors_list (J) entry 'FONAM/PROFONANPE' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-18 (n=17)
  - policy_lifecycle_phases_list (L) lists 13 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [10, 11, 17, 18]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Loi_2022-18_SONAGED.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4]
  - policy_sectors_list (J) lists 5 sector(s) ('waste management, municipalities, industry, environment, recycling'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.


## `4P_Index_Lettre_Politique_Environnement_2016-2020.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col L** — rows [2, 3, 4, 5]
  - policy_lifecycle_phases_list (L) lists 4 phase(s), inconsistent with policy_circularity (K)=0.5 bucket.


## `Mozambique_4P_Index_Decreto_79_2017.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) lists 10 sector(s) ('packaging, plastics, industry, retail, import, municipalities, waste management, recycling, chemicals, finance'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [5, 9, 16, 19]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_54_2015.xlsx`  (branch: `cursor/mozambique-decree-54-2015-4p-index-04b6`)

- **[ERROR] Col Q** — rows [4]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7]
  - policy_sectors_list (J) lists 8 sector(s) ('waste management, industry, municipalities, packaging, fisheries, tourism, water, agriculture'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.


## `4p_index_bnp_management_plan_2022.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4p_index_bbin_mpa_esmf_2022.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4P_Index_Produktforskrift.xlsx`  (branch: `cursor/4p-index-produktforskrift-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'use', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.77, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 5, 7, 17]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [3, 8, 9, 15]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [4, 6, 10, 11, 12, 13, 14, 16]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows [6, 11]
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) lists 7 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [17]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [4, 13]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Decret_2021-1115_CENPOLMAR.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[ERROR] Col None** — rows []
  - No data rows found.


## `4p-index/STP_Lei-4-2023_Customs-Duties-Law_4P-Index.xlsx`  (branch: `cursor/code-stp-customs-duties-law-3cda`)

- **[ERROR] Col O** — rows [2, 3, 4]
  - policy_score (O)=0.52 does not equal (G+I+K+M)/4 = (1.0+0.75+0.25+0.0)/4 = 0.500.


## `4p_index_tourism_health_protocol_2020.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4P_Index_Avfallsforskriften_Kap1.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap1-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 6, 9, 14, 15, 16]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [4]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [5, 7, 8, 10, 11, 12, 13, 17]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col Q** — rows [2, 3, 5, 7, 8, 9, 14, 16, 17]
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[REVIEW] Col S** — rows [4, 17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Avfallsforskriften_Kap7.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap7-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.87]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 4, 7, 11, 12, 13, 14, 15, 16]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [5, 6, 8, 9, 10, 17]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col Q** — rows [2, 4, 5, 6, 7, 8, 12, 13, 15, 17]
  - instrument_lifecycle_stage (Q)='Design' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Loi_2013-10_CGCT.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4, 5]
  - policy_sectors_list (J) lists 5 sector(s) ('local government, sanitation, environment, local governance, local public finance'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [4]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Skipssikkerhet_MARPOL_V.xlsx`  (branch: `cursor/4p-index-skipssikkerhet-marpol-v-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.72, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 4, 15, 16, 17]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [5, 6, 8]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [7, 9, 10, 11, 12, 13, 14]
  - policy_score (O)=0.72 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col Q** — rows 2-17 (n=16)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [15, 16]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[REVIEW] Col S** — rows [16]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.

- **[REVIEW] Col S** — rows [17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Avfallsforskrift_Kap10a.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap10a-65aa`)

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.77]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]).

- **[ERROR] Col O** — rows [2, 3, 4, 6, 7, 8, 20, 21]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [5, 9, 10, 11, 13, 14, 15, 16, 18]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [12, 17, 19]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows 2-21 (n=18)
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [7, 8, 20]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[WARN] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) lists 2 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [8]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [6]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_everest_singleuse_plastic_ban_2020.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Loi_2015-09_Sachets_Plastiques.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[REVIEW] Col S** — rows [4, 5, 6]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4P_Index_Noregs_Plaststrategi.xlsx`  (branch: `cursor/4p-index-noregs-plaststrategi-65aa`)

- **[ERROR] Col G** — rows 2-30 (n=29)
  - policy_type (G)=0.2 is not one of the allowed values {0.25, 0.5, 0.75, 1}.

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.58, 0.62, 0.66, 0.74]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]).

- **[ERROR] Col O** — rows 2-29 (n=15)
  - policy_score (O)=0.58 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col O** — rows [8, 22, 28, 30]
  - policy_score (O)=0.62 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col O** — rows [10, 11, 16, 20, 23, 26]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col Q** — rows 2-28 (n=19)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [6, 7, 14, 19, 24, 25, 29]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [13, 15, 21]
  - instrument_score (V)=0.675 does not equal S×(P+T)/2 = 0.0×(0.6+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [17, 18]
  - instrument_score (V)=0.475 does not equal S×(P+T)/2 = 0.0×(0.2+0.75)/2 = 0.000.

- **[WARN] Col J** — rows 2-30 (n=29)
  - policy_sectors_list (J) lists 1 sector(s) ('all sectors; fisheries; aquaculture; shipping; municipalities; agriculture; transport; waste management; retail; packaging; construction; textiles; electronics; sports facilities'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [7, 27]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Avfallsforskriften_Kap15.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap15-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.77]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7]).

- **[ERROR] Col O** — rows [2, 3, 5, 6]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col O** — rows [4, 7]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7]
  - policy_sectors_list (J) lists 7 sector(s) ('municipalities, households, waste management, recycling, consumption, packaging, retail'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.


## `4P_Index_NS_05-061_Eaux_Usees.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4]
  - policy_sectors_list (J) lists 7 sector(s) ('industry, sanitation, environment, water, fisheries, agriculture, waste management'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.


## `4P_Index_Handlingsplan_Sirkulaer.xlsx`  (branch: `cursor/4p-index-handlingsplan-sirkulaer-65aa`)

- **[ERROR] Col G** — rows 2-25 (n=24)
  - policy_type (G)=0.2 is not one of the allowed values {0.25, 0.5, 0.75, 1}.

- **[ERROR] Col L** — rows 2-25 (n=24)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-25 (n=24)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-25 (n=24)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows 2-25 (n=24)
  - policy_score (O)=0.58 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col Q** — rows [7, 8, 10, 11, 12, 13, 15, 20, 24]
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [4, 5, 7, 8, 9, 11, 15, 23]
  - instrument_score (V)=0.475 does not equal S×(P+T)/2 = 0.0×(0.2+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [10, 13, 19, 20, 21, 22, 25]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[WARN] Col J** — rows 2-25 (n=24)
  - policy_sectors_list (J) lists 1 sector(s) ('all sectors; fisheries; aquaculture; packaging; plastics; waste management; municipalities; construction; textiles; electronics; food; retail; public procurement; finance'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-25 (n=24)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [7, 8]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4p_index_nswmp_2079_2022.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[WARN] Col J** — rows [5, 6, 7]
  - policy_sectors_list (J) entry 'private sector/PPP' is not lower case (Coding Rules require lower case, comma-separated).

- **[REVIEW] Col P** — rows [5, 6]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [5]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/peru-dl-1278/Peru_DL_1278_4P_Index_Coding.xlsx`  (branch: `cursor/peru-dl-1278-4p-index-d350`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['B — policy_name_en', 'C — country']), which shifts every subsequent column's letter out of alignment. For example: file labels its 'policy_url' column as 'D' (should be 'B' per the prompt); file labels its 'policy_year' column as 'E' (should be 'C' per the prompt); file labels its 'policy_objective' column as 'F' (should be 'D' per the prompt); file labels its 'policy_target' column as 'G' (should be 'E' per the prompt); file labels its 'policy_target_text' column as 'H' (should be 'F' per the prompt); file labels its 'policy_type' column as 'I' (should be 'G' per the prompt). Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87, 0.91, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36]).

- **[ERROR] Col O** — rows [6, 9, 11, 22, 25, 32, 35]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [23, 36]
  - policy_score (O)=0.91 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [26]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [27, 28, 29]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col Q** — rows [36]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[REVIEW] Col S** — rows [10, 27, 31, 34]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p-index/STP_Decreto-36-1999_Waste-Decree_4P-Index.xlsx`  (branch: `cursor/code-stp-waste-decree-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.77, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16]).

- **[ERROR] Col O** — rows [2, 4, 5, 8, 9, 10, 11, 14, 15]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col O** — rows [3, 6, 7, 13]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col O** — rows [12, 16]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+0.75+0.75+1.0)/4 = 0.812.

- **[ERROR] Col Q** — rows [9]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [12]
  - instrument_score (V)=0.3 does not equal S×(P+T)/2 = 0.0×(0.6+0.0)/2 = 0.000.

- **[ERROR] Col V** — rows [15]
  - instrument_score (V)=0.625 does not equal S×(P+T)/2 = 0.0×(1.0+0.25)/2 = 0.000.

- **[REVIEW] Col S** — rows [9, 14]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.

- **[REVIEW] Col S** — rows [12]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/mozambique/4P_Index_Mozambique_NDC_3_0_Provisional.xlsx`  (branch: `cursor/mozambique-ndc-3-0-4p-index-04b6`)

- **[WARN] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) lists 5 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.

- **[REVIEW] Col P** — rows [14]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ds-014-2017-minam/Peru_DS_014_2017_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-014-2017-4p-index-d350`)

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.85, 0.9, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30]).

- **[ERROR] Col O** — rows 2-30 (n=20)
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [3, 5]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col P** — rows 2-30 (n=20)
  - instrument_type (P)=0.75 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col P** — rows [3, 5]
  - instrument_type (P)=0.5 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col Q** — rows [18, 26, 27]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-30 (n=29)
  - policy_lifecycle_phases_list (L) lists 9 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [3]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.5 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [30]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.75 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [23]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `Mozambique_4P_Index.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[ERROR] Col Q** — rows [18]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [7, 8, 9, 10]
  - policy_sectors_list (J) lists 6 sector(s) ('production, retail, industry, municipalities, waste management, packaging'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [11, 12, 13]
  - policy_sectors_list (J) lists 7 sector(s) ('packaging, industry, retail, municipalities, waste management, recycling, chemicals'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [14, 15, 16, 17]
  - policy_sectors_list (J) lists 7 sector(s) ('waste management, municipalities, industry, retail, recycling, public health, environment'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [18, 19]
  - policy_sectors_list (J) lists 6 sector(s) ('industry, agriculture, chemicals, waste management, municipalities, packaging'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [20]
  - policy_sectors_list (J) lists 5 sector(s) ('waste management, municipalities, industry, recycling, environment'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [21]
  - policy_sectors_list (J) lists 5 sector(s) ('industry, environment, energy, transport, chemicals'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [22]
  - policy_sectors_list (J) lists 6 sector(s) ('fisheries, maritime, ports, environment, tourism, municipalities'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [24]
  - policy_sectors_list (J) lists 7 sector(s) ('industry, infrastructure, mining, energy, tourism, environment, municipalities'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [26]
  - policy_sectors_list (J) lists 7 sector(s) ('fisheries, maritime, tourism, ports, environment, municipalities, industry'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [27]
  - policy_sectors_list (J) lists 6 sector(s) ('maritime, fisheries, mining, transport, environment, tourism'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [28]
  - policy_sectors_list (J) lists 5 sector(s) ('municipalities, tourism, fisheries, environment, waste management'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [29]
  - policy_sectors_list (J) lists 7 sector(s) ('maritime, fisheries, energy, transport, environment, tourism, mining'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [33]
  - policy_sectors_list (J) lists 5 sector(s) ('municipalities, fisheries, tourism, environment, waste management'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [34]
  - policy_sectors_list (J) lists 5 sector(s) ('fisheries, environment, tourism, maritime, waste management'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [36]
  - policy_sectors_list (J) lists 7 sector(s) ('urban planning, municipalities, environment, infrastructure, industry, tourism, agriculture'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col J** — rows [37]
  - policy_sectors_list (J) lists 7 sector(s) ('urban planning, municipalities, environment, infrastructure, industry, agriculture, tourism'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[REVIEW] Col P** — rows [2, 5, 14, 26, 28, 33, 36]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [11]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [17]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_45_2006.xlsx`  (branch: `cursor/mozambique-decree-45-2006-4p-index-04b6`)

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7, 8, 9]
  - policy_sectors_list (J) lists 7 sector(s) ('fisheries, waste management, industry, tourism, municipalities, water, packaging'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.


## `4P_Index_PROMOGED_Dechets_Solides.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.8 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4p_index_plastic_flowers_ban_2022.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [5]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Decret_2016-1804_Code_Peche_Maritime.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4]
  - policy_sectors_list (J) lists 5 sector(s) ('maritime fisheries, artisanal fishing, industrial fishing, marine environment, maritime surveillance'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [2]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Arretes_EIA_9468-9472_2001.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[REVIEW] Col S** — rows [4]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_ghodaghodi_lake_masterplan_2020.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col S** — rows [5]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4p_index_plastic_bags_action_plan_2022.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col S** — rows [5]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_nepal_7nr_cbd_2026.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `output/peru-hrnec-2030/Peru_HRNEC_2030_4P_Index_Coding.xlsx`  (branch: `cursor/peru-hrnec-2030-4p-index-d350`)

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'diseño', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'recolección', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'valorización', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'regeneración de ecosistemas (marco NTP-ISO 59004:2024; Cuadro 2) / design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'collection', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'valorization', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'ecosystem regeneration (NTP-ISO 59004:2024 framework; Table 2)', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.7, 0.708, 0.717, 0.733, 0.75, 0.767]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]).

- **[ERROR] Col O** — rows [2, 3]
  - policy_score (O)=0.7 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [4, 5, 6, 7]
  - policy_score (O)=0.717 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [12, 18]
  - policy_score (O)=0.708 does not equal (G+I+K+M)/4 = (0.5+1.0+1.0+0.5)/4 = 0.750.

- **[ERROR] Col P** — rows [4, 5, 6, 7]
  - instrument_type (P)=0.3 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col P** — rows [10, 11, 14, 16]
  - instrument_type (P)=0.5 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col P** — rows [12, 18]
  - instrument_type (P)=0.25 is not one of the allowed values {0, 0.20, 0.40, 0.60, 0.80, 1.0}.

- **[ERROR] Col Q** — rows 2-19 (n=18)
  - instrument_lifecycle_stage (Q)='Planning' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'transversal (10 sectores Ejecutivo: PCM' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MINAM' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MEF' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'PRODUCE' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MINEM' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'VIVIENDA' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MINCETUR' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MIDAGRI' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MTPE' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MINEDU); plásticos/envases' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'agrario / cross-cutting (10 Executive Branch sectors: PCM' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'HOUSING' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-19 (n=18)
  - policy_sectors_list (J) entry 'MINEDU); plastics/packaging' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) lists 15 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.3 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Produktforskrift_Kap2b.xlsx`  (branch: `cursor/4p-index-produktforskrift-kap2b-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.77, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 6, 11, 15]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [4, 5, 7]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [8, 9, 10, 12, 13, 14, 17]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [16]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows 2-17 (n=11)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [6, 11]
  - instrument_score (V)=0.475 does not equal S×(P+T)/2 = 0.0×(0.2+0.75)/2 = 0.000.

- **[REVIEW] Col S** — rows [15, 17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_ssdp_emf_2017.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `output/peru-ds-013-2018-minam/Peru_DS_013_2018_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-013-2018-4p-index-d350`)

- **[ERROR] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.54, 0.58, 0.7]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]).

- **[ERROR] Col O** — rows [2, 9, 11, 12, 14]
  - policy_score (O)=0.54 does not equal (G+I+K+M)/4 = (0.75+0.5+0.75+0.5)/4 = 0.625.

- **[ERROR] Col O** — rows [3, 10, 13]
  - policy_score (O)=0.58 does not equal (G+I+K+M)/4 = (0.75+0.5+0.75+0.5)/4 = 0.625.

- **[ERROR] Col O** — rows [4, 5, 6, 7, 8]
  - policy_score (O)=0.7 does not equal (G+I+K+M)/4 = (0.75+0.5+0.75+0.5)/4 = 0.625.

- **[ERROR] Col V** — rows [9]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[WARN] Col J** — rows 2-14 (n=13)
  - policy_sectors_list (J) lists 13 sector(s) ('administración pública, ambiente, consumo, compras públicas, cultura, turismo, salud / public administration, environment, consumption, public procurement, culture, tourism, health'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) lists 7 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.


## `output/mozambique/4P_Index_Mozambique_Resolution_7_2021_PNDT.xlsx`  (branch: `cursor/mozambique-resolution-7-2021-pndt-04b6`)

- **[ERROR] Col Q** — rows [2]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7]
  - policy_sectors_list (J) lists 6 sector(s) ('municipalities, water, waste management, industry, packaging, environment'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7]
  - policy_lifecycle_phases_list (L) lists 3 phase(s), inconsistent with policy_circularity (K)=0.5 bucket.

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `Mozambique_4P_Index_ESMF_CRRN.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[ERROR] Col Q** — rows [4, 5, 6, 9]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-14 (n=13)
  - policy_lifecycle_phases_list (L) lists 5 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.


## `output/peru-law-29419/Peru_Law_29419_4P_Index_Coding.xlsx`  (branch: `cursor/peru-law-29419-4p-index-d350`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['B — policy_name_en', 'C — policy_number', 'D — country']), which shifts every subsequent column's letter out of alignment. For example: file labels its 'policy_url' column as 'E' (should be 'B' per the prompt); file labels its 'policy_year' column as 'F' (should be 'C' per the prompt); file labels its 'policy_objective' column as 'G' (should be 'D' per the prompt); file labels its 'policy_target' column as 'H' (should be 'E' per the prompt); file labels its 'policy_target_text' column as 'I' (should be 'F' per the prompt); file labels its 'policy_type' column as 'J' (should be 'G' per the prompt). Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87, 0.91, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]).

- **[ERROR] Col O** — rows [2, 3, 4, 6, 8, 14, 19]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [9, 15]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [10]
  - policy_score (O)=0.91 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [11, 12, 16, 17, 18]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[REVIEW] Col P** — rows [2, 14]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [15]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [15, 17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Stortingsvedtak_Drikkevareemballasje_2026.xlsx`  (branch: `cursor/4p-index-stortingsvedtak-2026-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.84, 0.92, 1.0]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15]).

- **[ERROR] Col O** — rows [2, 4, 5, 6, 9, 10, 12, 15]
  - policy_score (O)=0.92 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col O** — rows [3, 8, 13, 14]
  - policy_score (O)=0.84 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col V** — rows [3, 8, 13]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [5]
  - instrument_score (V)=0.55 does not equal S×(P+T)/2 = 0.0×(0.6+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [15]
  - instrument_score (V)=0.425 does not equal S×(P+T)/2 = 0.0×(0.6+0.25)/2 = 0.000.


## `4P_Index_REACH_forskrift.xlsx`  (branch: `cursor/4p-index-reach-forskrift-65aa`)

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-17 (n=16)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.72, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 4, 5, 7, 16, 17]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [6, 8, 9, 13, 14, 15]
  - policy_score (O)=0.72 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [10, 11, 12]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col Q** — rows [7, 10, 11, 12, 13, 14, 15]
  - instrument_lifecycle_stage (Q)='Use/Consumption' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[REVIEW] Col S** — rows [9]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_94_2014.xlsx`  (branch: `cursor/mozambique-decree-94-2014-4p-index-04b6`)

- **[WARN] Col J** — rows 2-21 (n=20)
  - policy_sectors_list (J) lists 7 sector(s) ('production, consumption, retail, waste management, recycling, municipalities, industry'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.


## `output/mozambique/4P_Index_Mozambique_Decree_23_2008_LOT_Regulation.xlsx`  (branch: `cursor/mozambique-decree-23-2008-lot-regulation-04b6`)

- **[ERROR] Col Q** — rows [2, 3, 4]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).


## `4P_Index_Produktforskrift_SUP_3200.xlsx`  (branch: `cursor/4p-index-produktforskrift-sup-3200-65aa`)

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-13 (n=12)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.64, 0.8]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]).

- **[ERROR] Col O** — rows [2, 3, 4, 5, 7, 8, 13]
  - policy_score (O)=0.64 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col O** — rows [6, 9, 10, 11, 12]
  - policy_score (O)=0.8 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+0.5)/4 = 0.750.

- **[ERROR] Col Q** — rows [4]
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [7, 8]
  - instrument_score (V)=0.475 does not equal S×(P+T)/2 = 0.0×(0.2+0.75)/2 = 0.000.


## `4p-index/STP_Lei-9-2022_Fisheries-Aquaculture-Law_4P-Index.xlsx`  (branch: `cursor/code-stp-fisheries-law-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.87, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11]).

- **[ERROR] Col O** — rows [10]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [11]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (1.0+1.0+0.75+1.0)/4 = 0.938.

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [4, 6, 10, 11]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Drikkevareemballasje.xlsx`  (branch: `cursor/4p-index-drikkevareemballasje-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.79, 0.83, 0.87, 0.95]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]).

- **[ERROR] Col O** — rows [2, 3, 5, 6, 7, 8, 10, 11, 12]
  - policy_score (O)=0.87 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [9]
  - policy_score (O)=0.83 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[ERROR] Col O** — rows [13, 14]
  - policy_score (O)=0.79 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+1.0)/4 = 0.938.

- **[REVIEW] Col S** — rows [11, 12, 14]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Baerekraftige_Produkter_Loven.xlsx`  (branch: `cursor/4p-index-baerekraftige-produkter-65aa`)

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]).

- **[ERROR] Col O** — rows [2, 3, 21]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows 4-20 (n=13)
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col O** — rows [11, 12, 17, 18]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+0.5)/4 = 0.875.

- **[ERROR] Col Q** — rows [2, 4, 7, 9, 10]
  - instrument_lifecycle_stage (Q)='Design' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [3]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [4, 5, 6, 7, 8, 9, 10, 20]
  - instrument_score (V)=0.875 does not equal S×(P+T)/2 = 0.0×(1.0+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [11, 17]
  - instrument_score (V)=0.675 does not equal S×(P+T)/2 = 0.0×(0.6+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [18]
  - instrument_score (V)=0.55 does not equal S×(P+T)/2 = 0.0×(0.6+0.5)/2 = 0.000.

- **[WARN] Col J** — rows 2-21 (n=20)
  - policy_sectors_list (J) lists 1 sector(s) ('all sectors; manufacturing; packaging; plastics; batteries; vehicles; textiles; electronics; retail; waste management; public procurement'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-21 (n=20)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [14, 15, 16, 19]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Resolution_51_2022_Coral_Reef.xlsx`  (branch: `cursor/mozambique-resolution-51-2022-coral-reef-04b6`)

- **[ERROR] Col M** — rows [2, 3, 4, 5, 6, 7]
  - policy_budget (M)=0.75 is not one of the allowed values {0, 0.5, 1}.

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7]
  - policy_sectors_list (J) lists 7 sector(s) ('fisheries, waste management, industry, tourism, municipalities, water, packaging'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7]
  - policy_lifecycle_phases_list (L) lists 2 phase(s), inconsistent with policy_circularity (K)=0.25 bucket.

- **[REVIEW] Col S** — rows [4]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p_index_hcwm_guideline_2014.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [9]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4p_index_lgoa_2074_2025.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4P_Index_Avfallsforskrift_Kap7b.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap7b-65aa`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.77, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17]).

- **[ERROR] Col O** — rows [2, 3, 4, 5, 7, 8, 9, 11, 13, 17]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [6, 14]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [10, 12, 15]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [16]
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col V** — rows [11]
  - instrument_score (V)=0.625 does not equal S×(P+T)/2 = 0.0×(1.0+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [12]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[REVIEW] Col S** — rows [17]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Resolution_13_2021_Health_Policy.xlsx`  (branch: `cursor/mozambique-resolution-13-2021-health-policy-04b6`)

- **[ERROR] Col Q** — rows [2]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4]
  - policy_sectors_list (J) lists 3 sector(s) ('municipalities, water, waste management'), inconsistent with policy_integration (I)=0.25 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4]
  - policy_lifecycle_phases_list (L) lists 2 phase(s), inconsistent with policy_circularity (K)=0.25 bucket.

- **[REVIEW] Col P** — rows [2]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [3]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4p-index/STP_PNGIRSU-2018-2023_4P-Index.xlsx`  (branch: `cursor/code-stp-pngirsu-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.59, 0.71]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8]).

- **[ERROR] Col O** — rows [2, 3, 4, 5]
  - policy_score (O)=0.71 does not equal (G+I+K+M)/4 = (0.5+1.0+0.75+0.5)/4 = 0.688.

- **[ERROR] Col O** — rows [6, 7, 8]
  - policy_score (O)=0.59 does not equal (G+I+K+M)/4 = (0.5+1.0+0.75+0.5)/4 = 0.688.

- **[REVIEW] Col S** — rows [7]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Avfallsforskrift_Kap7a.xlsx`  (branch: `cursor/4p-index-avfallsforskrift-kap7a-65aa`)

- **[ERROR] Col L** — rows 2-22 (n=21)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-22 (n=21)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-22 (n=21)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.77, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22]).

- **[ERROR] Col O** — rows [2, 3, 4]
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [5, 6, 7, 10, 22]
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows 8-21 (n=13)
  - policy_score (O)=0.77 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows [2, 3, 5, 6, 7, 8, 16, 17, 18, 19]
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows 2-22 (n=21)
  - policy_sectors_list (J) lists 1 sector(s) ('retail; food service; municipalities; waste management; tobacco; packaging; producer responsibility organisations'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-22 (n=21)
  - policy_lifecycle_phases_list (L) lists 3 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [22]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_45_2024.xlsx`  (branch: `cursor/mozambique-decree-45-2024-4p-index-04b6`)

- **[WARN] Col J** — rows [2, 3, 4, 5, 6]
  - policy_sectors_list (J) lists 8 sector(s) ('waste management, industry, municipalities, packaging, fisheries, tourism, water, agriculture'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6]
  - policy_lifecycle_phases_list (L) lists 3 phase(s), inconsistent with policy_circularity (K)=0.5 bucket.


## `output/mozambique/4P_Index_Mozambique_ANAC_Tourism_Concessions_Manual_2012.xlsx`  (branch: `cursor/mozambique-anac-tourism-concessions-manual-2012-04b6`)

- **[ERROR] Col Q** — rows [2, 3, 4]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col J** — rows [2, 3, 4, 5, 6, 7]
  - policy_sectors_list (J) lists 6 sector(s) ('tourism, waste management, municipalities, industry, environment, packaging'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7]
  - policy_lifecycle_phases_list (L) lists 3 phase(s), inconsistent with policy_circularity (K)=0.5 bucket.


## `4P_Index_Arrete_07022_POLMAR.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[REVIEW] Col P** — rows [3]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.2 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [4]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col P** — rows [5]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.8 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `4p_index_16th_plan_nepal_2024.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4P_Index_Meld_St_45_Kap7.xlsx`  (branch: `cursor/4p-index-meld-st-45-kap7-65aa`)

- **[ERROR] Col G** — rows 2-19 (n=18)
  - policy_type (G)=0.2 is not one of the allowed values {0.25, 0.5, 0.75, 1}.

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col M** — rows 2-19 (n=18)
  - policy_budget (M)=0.75 is not one of the allowed values {0, 0.5, 1}.

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.58, 0.66, 0.74]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]).

- **[ERROR] Col O** — rows [2, 3, 4, 11, 16, 17, 19]
  - policy_score (O)=0.58 does not equal (G+I+K+M)/4 = (0.2+1.0+0.75+0.75)/4 = 0.675.

- **[ERROR] Col O** — rows [15]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.2+1.0+0.75+0.75)/4 = 0.675.

- **[ERROR] Col Q** — rows 2-19 (n=15)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [2, 11, 16, 19]
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [3, 17]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [5, 7, 9, 12, 13, 14, 18]
  - instrument_score (V)=0.55 does not equal S×(P+T)/2 = 0.0×(0.6+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [6, 8, 10]
  - instrument_score (V)=0.675 does not equal S×(P+T)/2 = 0.0×(0.6+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [15]
  - instrument_score (V)=0.75 does not equal S×(P+T)/2 = 0.0×(1.0+0.5)/2 = 0.000.

- **[WARN] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) lists 5 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.

- **[REVIEW] Col S** — rows [3, 6, 7, 14]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `4p_index_karnali_tourism_masterplan.xlsx`  (branch: `cursor/4p-index-karnali-tourism-masterplan-376f`)

- **[ERROR] Col None** — rows []
  - This file inserts column(s) not defined in the prompt's fixed A-W schema (['Col X\n—']), which shifts every subsequent column's letter out of alignment. Per the prompt, only columns A through W (in that exact order/meaning) should be used; extra fields should either be dropped or appended after column W, not inserted earlier.


## `4p-index/STP_Lei-8-2020_Plastic-Bag-Law_4P-Index.xlsx`  (branch: `cursor/code-stp-plastic-bag-law-3cda`)

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.84, 0.88, 0.92, 1.0]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14]).

- **[ERROR] Col O** — rows [5, 6, 12]
  - policy_score (O)=0.92 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col O** — rows [9]
  - policy_score (O)=0.88 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.

- **[ERROR] Col O** — rows [10, 11, 14]
  - policy_score (O)=0.84 does not equal (G+I+K+M)/4 = (1.0+1.0+1.0+1.0)/4 = 1.000.


## `4P_Index_Sirkulaerokonomistrategi.xlsx`  (branch: `cursor/4p-index-sirkulaerokonomistrategi-65aa`)

- **[ERROR] Col G** — rows 2-28 (n=27)
  - policy_type (G)=0.2 is not one of the allowed values {0.25, 0.5, 0.75, 1}.

- **[ERROR] Col L** — rows 2-28 (n=27)
  - policy_lifecycle_phases_list (L) contains 'use/consumption', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-28 (n=27)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-28 (n=27)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.58, 0.66, 0.74]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28]).

- **[ERROR] Col O** — rows 2-28 (n=20)
  - policy_score (O)=0.58 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col O** — rows [16, 18]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.2+1.0+1.0+0.5)/4 = 0.675.

- **[ERROR] Col Q** — rows 5-27 (n=15)
  - instrument_lifecycle_stage (Q)='Litter/Pollution' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows 3-28 (n=11)
  - instrument_score (V)=0.35 does not equal S×(P+T)/2 = 0.0×(0.2+0.5)/2 = 0.000.

- **[ERROR] Col V** — rows [6, 11, 13, 23]
  - instrument_score (V)=0.475 does not equal S×(P+T)/2 = 0.0×(0.2+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [8, 12, 17, 21, 22]
  - instrument_score (V)=0.675 does not equal S×(P+T)/2 = 0.0×(0.6+0.75)/2 = 0.000.

- **[ERROR] Col V** — rows [18]
  - instrument_score (V)=0.875 does not equal S×(P+T)/2 = 0.0×(1.0+0.75)/2 = 0.000.

- **[WARN] Col J** — rows 2-28 (n=27)
  - policy_sectors_list (J) lists 1 sector(s) ('all sectors; bioeconomy; process industry; construction and buildings; retail and wholesale trade; waste management; municipalities; fisheries/aquaculture; packaging; plastics; textiles; electronics; food; finance; public procurement'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[WARN] Col L** — rows 2-28 (n=27)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col S** — rows [3, 12, 18, 23]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/mozambique/4P_Index_Mozambique_Resolution_53_2024_EDEA.xlsx`  (branch: `cursor/mozambique-resolution-53-2024-edea-04b6`)

- **[ERROR] Col M** — rows [2, 3, 4, 5, 6, 7]
  - policy_budget (M)=0.75 is not one of the allowed values {0, 0.5, 1}.

- **[WARN] Col L** — rows [2, 3, 4, 5, 6, 7]
  - policy_lifecycle_phases_list (L) lists 4 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.


## `4P_Index_Loi_2009-24_Code_Assainissement.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4, 5]
  - policy_sectors_list (J) lists 6 sector(s) ('sanitation, urban planning, local government, environment, health, industry'), inconsistent with policy_integration (I)=1.0 bucket. Recount sectors or adjust I.

- **[REVIEW] Col S** — rows [2]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `4P_Index_Plan_Gestion_Dechets_Biomedicaux_2019.xlsx`  (branch: `cursor/senegal-plastic-policies-faa3`)

- **[WARN] Col J** — rows [2, 3, 4, 5]
  - policy_sectors_list (J) lists 5 sector(s) ('health, waste management, environment, municipalities, industry'), inconsistent with policy_integration (I)=0.5 bucket. Recount sectors or adjust I.


## `4P_Index_Forurensningsforskrift.xlsx`  (branch: `cursor/4p-index-forurensningsforskrift-65aa`)

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'end-of-life', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) contains 'litter/pollution', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.74, 0.82, 0.9]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19]).

- **[ERROR] Col O** — rows [2, 5, 7, 10, 12, 14, 15, 19]
  - policy_score (O)=0.74 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [3, 4, 6, 8, 9, 11, 16, 17, 18]
  - policy_score (O)=0.82 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col O** — rows [13]
  - policy_score (O)=0.9 does not equal (G+I+K+M)/4 = (0.75+1.0+0.75+1.0)/4 = 0.875.

- **[ERROR] Col Q** — rows [6, 7, 8, 10, 11, 17, 18]
  - instrument_lifecycle_stage (Q)='End-of-life' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-19 (n=18)
  - policy_lifecycle_phases_list (L) lists 6 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.

- **[REVIEW] Col S** — rows [14]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `output/mozambique/4P_Index_Mozambique_Decree_97_2020_Coastal_Zones.xlsx`  (branch: `cursor/mozambique-decree-97-2020-coastal-zones-04b6`)

- **[REVIEW] Col P** — rows [7]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=0.6 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.


## `Mozambique_4P_Index_Resolucao_5_95.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[ERROR] Col Q** — rows [17]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[WARN] Col L** — rows 2-24 (n=23)
  - policy_lifecycle_phases_list (L) lists 5 phase(s), inconsistent with policy_circularity (K)=0.75 bucket.

- **[REVIEW] Col P** — rows [9]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [24]
  - instrument_in_force (S)=0, but the description/implementation text contains mandatory language ('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check whether S should be 1.


## `output/peru-ds-006-2019-minam/Peru_DS_006_2019_MINAM_4P_Index_Coding.xlsx`  (branch: `cursor/peru-ds-006-2019-4p-index-d350`)

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'producción', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'diseño', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'consumo', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'reciclaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'compostaje', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'disposición', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'fuga ambiental / production', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'design', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) contains 'composting', not one of the 5 allowed phases (production, consumption, recycling, disposal, environmental leakage).

- **[ERROR] Col O** — rows [2]
  - policy_score (O) varies across rows of the same policy ([0.69, 0.73, 0.85]); per the prompt O is a policy-level field and MUST be identical across every instrument row for this policy (rows [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34]).

- **[ERROR] Col O** — rows 2-31 (n=16)
  - policy_score (O)=0.69 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows 4-32 (n=11)
  - policy_score (O)=0.85 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col O** — rows [5, 19, 28, 29, 33, 34]
  - policy_score (O)=0.73 does not equal (G+I+K+M)/4 = (0.75+1.0+1.0+0.5)/4 = 0.812.

- **[ERROR] Col Q** — rows [18]
  - instrument_lifecycle_stage (Q)='Disposal' is not one of the allowed title-case values ['Consumption', 'End of life', 'Environmental leakage', 'Production', 'Recycling', 'Waste management'] (Rule 4).

- **[ERROR] Col V** — rows [10]
  - instrument_score (V)=0.625 does not equal S×(P+T)/2 = 0.0×(1.0+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [20, 24, 30, 31]
  - instrument_score (V)=0.225 does not equal S×(P+T)/2 = 0.0×(0.2+0.25)/2 = 0.000.

- **[ERROR] Col V** — rows [28, 33]
  - instrument_score (V)=0.325 does not equal S×(P+T)/2 = 0.0×(0.4+0.25)/2 = 0.000.

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'protección al consumidor (INDECOPI)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'aduanas (SUNAT)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'salud (MINSA)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'empaques PET / retail commerce' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'consumer protection (INDECOPI)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'customs (SUNAT)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'health (MINSA)' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col J** — rows 2-34 (n=33)
  - policy_sectors_list (J) entry 'PET packaging' is not lower case (Coding Rules require lower case, comma-separated).

- **[WARN] Col L** — rows 2-34 (n=33)
  - policy_lifecycle_phases_list (L) lists 13 phase(s), inconsistent with policy_circularity (K)=1.0 bucket.

- **[REVIEW] Col P** — rows [9, 10, 23]
  - instrument_description/comments mention coordination/multi-level governance responsibilities, but instrument_type (P)=1.0 is not 0.40 and Col W does not note the coordination aspect alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or add a note in W justifying the higher score.

- **[REVIEW] Col S** — rows [7, 34]
  - instrument_in_force (S)=1, but the description/implementation text only shows enabling-power language ('may' / 'is authorised to') with no mandatory language and no evidence the power has been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.


## `Mozambique_4P_Index_Decreto_16_2015.xlsx`  (branch: `cursor/mozambique-plastic-policies-7c97`)

- **[WARN] Col J** — rows 2-16 (n=15)
  - policy_sectors_list (J) lists 8 sector(s) ('production, consumption, retail, waste management, industry, municipalities, packaging, fisheries'), inconsistent with policy_integration (I)=0.75 bucket. Recount sectors or adjust I.
