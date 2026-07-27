# 4P Index coding — Senegal, Loi n° 2001-01 (Code de l'Environnement)

This folder contains a completed **Plastic Pollution Policy Index (4P Index)** coding table
(Columns A–W, per the *4P Index — Policy Coding Prompt v2*) for Senegal's foundational
environmental code, produced as a downloadable artifact.

## Files

- `senegal_loi_2001-01_4p_index.xlsx` — formatted spreadsheet (recommended for viewing/download).
- `senegal_loi_2001-01_4p_index.csv` — plain-text version of the same table.
- `build_4p_index.py` — the script that generated both files from the coded data (source of truth;
  edit this and re-run `python3 build_4p_index.py` to regenerate the outputs if the coding changes).

## What was coded

**Policy coded:** *Loi n° 2001-01 du 15 janvier 2001 portant Code de l'Environnement* (Senegal).

Primary sources consulted (full text):
- https://faolex.fao.org/docs/pdf/sen34608.pdf
- https://acbep.gouv.sn/sites/default/files/2021-12/Loi%20n%C2%B0%202001-01%20du%2012%20avril%202001%20portant%20code%20de%20l%27environnement.pdf
- https://www.sante.gouv.sn/sites/default/files/2.%2001-01-15CODENVIRONLOI.pdf
- Décret n° 2001-282 portant application du Code de l'environnement (implementing decree), for context.

**Date discrepancy note:** the task brief cited "*Loi n° 2001-01 du 12 avril 2001*". All three primary
legislative texts located show the law was deliberated by the National Assembly (29 Dec 2000) and the
Senate (4 Jan 2001) and is dated **15 January 2001**. 12 April 2001 is the date of the separate
*Décret n° 2001-282* (implementing decree for classified installations, EIA procedure, water/air/noise
pollution — issued under this law). This is flagged in Column W of the first row and in the policy
title (Column A) uses the law's own dated title.

**Supersession note:** Loi n° 2001-01 was repealed and replaced by *Loi n° 2023-15 du 2 août 2023
portant Code de l'Environnement*, adopted 7 June 2023. Per the coding rules, this table codes the 2001
law as the specific instrument requested (year = 2001, its own enactment year), since the 2023 law is a
distinct, separately-numbered replacement rather than an amendment of the 2001 text. A separate 4P Index
entry should be created for Loi n° 2023-15 if the current legal framework needs to be assessed.

**Plastics relevance:** the 2001 Code never mentions "plastique" / synthetic polymers explicitly. All
eight coded instruments qualify under filter criteria (b) and/or (c) of the coding prompt — they are
waste-management/governance instruments (classified-installations permitting, waste elimination duty,
dumping/hazardous-waste-import bans, EIA, emergency planning, State–municipality coordination) or an
enabling power (Art. L40) that directly enables future plastic-specific bans — rather than because the
Code names plastics. Senegal later exercised plastics-specific policy powers through separate,
freestanding legislation (*Loi n° 2015-09* of 4 May 2015, replaced by *Loi n° 2020-04* of 8 January
2020), which are cross-referenced in the comments column (W) and should be coded as their own policy
entries in the 4P Index rather than folded into this one.

## Summary of results

- **Policy-level score (Column O)** = (G + I + K + M) / 4 = (1 + 1 + 1 + 1) / 4 = **1.0**
  (Legislation; 13 sectors touched ⇒ 7+ bracket; all 5 life-cycle phases addressed in general terms;
  ring-fenced Environmental Protection Fund financed by the Code's own pollution levies).
- **8 instrument rows** were coded (Columns P–W), covering: classified-installations permitting,
  pollution taxes/Environmental Protection Fund, general waste-elimination duty, absolute
  dumping/hazardous-waste-import bans, the Art. L40 waste-generating-products enabling power (not yet in
  force under this law), mandatory EIA, emergency/internal-operations planning, and the
  State–municipality coordination framework.
- Instrument scores (Column V) range from **0.0** (Art. L40 enabling power, not yet operationalised) to
  **1.0** (absolute dumping/hazardous-waste-import bans).

See the spreadsheet/CSV for full article-level citations, in-force reasoning, and implementation-scoring
evidence for every row.
