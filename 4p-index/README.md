# 4P Index coding — São Tomé and Príncipe

## Lei n.º 4/2023 — Altera os Direitos Aduaneiros sobre as Lâmpadas, Plásticos, Água Mineral, Equipamentos e Matérias-Primas para a Produção de Energias Renováveis e de Sabão / Law No. 4/2023 — Amending Customs Duties on Lamps, Plastics, Mineral Water, Equipment and Raw Materials for Renewable-Energy Production and Soap

### IMPORTANT SOURCE LIMITATION — please read before using this coding

**No version of the promulgated text of Lei n.º 4/2023 — not even a fragment — could be located
online for this exercise.** This confirms the task brief's own "VERIFICATION LIMIT" note: the
official consolidated text could not be located on dre.gov.st (the site blocks automated
retrieval; repeated attempts in this exercise returned HTTP 403 errors), and other government
portals hosting similar tax legislation (e.g. impostos.financas.gov.st) likewise blocked
category-browsing access.

The coding and briefing in this folder were instead produced from:

1. The task brief provided for this exercise, citing the U.S. Department of State's 2024
   Investment Climate Statement for STP (a secondary, non-government-hosted reference).
2. **The National Assembly's own official record of the bill's parliamentary passage** — Diário
   da Assembleia Nacional, I Série, N.º 4 (20 May 2023) — which includes the full written opinion
   of the 2nd Specialised Standing Committee (17 May 2023) and the record of the article-by-article
   vote. This record describes the bill's content in **indirect/reported speech, in Portuguese**
   (i.e. paraphrasing what the articles do), rather than quoting the articles verbatim — but it
   does report **specific duty-rate figures**, which substantially strengthens this coding relative
   to a coding based on the task brief alone.
3. LegisPalop's legislative-index confirmation of the enacted law's official title ("Lei n.º
   4/2023 de 15 de Junho") and other independent secondary sources (World Bank, SE4All, Téla Nón)
   corroborating the substance of the measures.

**This coding should be treated as provisional and cross-checked against the primary text once
obtained directly from https://dre.gov.st/**, per the task brief's own instruction. Every row's
comments column documents this sourcing caveat explicitly.

### What the committee report reveals (duty-rate changes)

Per the National Assembly committee report, Article 1 ("Agravamento dos direitos aduaneiros")
gradually increases customs duties as follows:
- Mineral water: 5% → 10% (of CIF value)
- Plastic bags ("sacos de plásticos"): 20% → 25%
- Other plastic products and manufactures ("demais produtos de plásticos e suas obras"): 10% → 15%
- Traditional/incandescent lamps: 20% → 55%

Article 2 ("Isenção de direitos aduaneiros") exempts from customs duties: LED lamps; materials
and equipment for renewable-energy production; and raw materials for local soap production.
Article 3 is the entry-into-force clause. All three articles were approved unanimously without
amendment at the specialty-review stage.

### Coding files

`STP_Lei-4-2023_Customs-Duties-Law_4P-Index.csv` (and the equivalent `.xlsx` workbook) contain the
4P Index coding, with **3 instruments coded** — all Economic instruments (0.6), all `in_force = 1`:

1. Increased customs duty on plastic bags (20%→25%) — explicit plastics reference, criterion (a).
2. Increased customs duty on other plastic products/articles (10%→15%) — explicit plastics
   reference, criterion (a).
3. Increased customs duty on imported bottled mineral water (5%→10%) — **borderline plastics
   relevance, explicitly flagged**: included under criterion (c) on the basis that imported bottled
   mineral water is predominantly sold in single-use PET plastic bottles, but neither source frames
   this measure in plastics/environmental-leakage terms (it is framed as local-industry
   protection). A stricter reading of the plastics-relevance filter could exclude this row.

The lamp-duty increase/LED exemption and the renewable-energy-equipment/soap-raw-material
exemptions (Article 2) were **not** coded as separate instrument rows, since they do not
independently satisfy the plastics-relevance filter; they are described in the policy-level
objective (Column D) for completeness.

- Policy type = 1 (Legislation — approved by the National Assembly, 52 votes in favour at general
  vote, unanimous at specialty and final-global vote).
- Policy circularity = 0.25 (only the production/import life-cycle phase is addressed — this is a
  narrow upstream tariff instrument, not a comprehensive plastics policy).
- Policy budget = 0 (no dedicated fund or ring-fenced revenue use was identified; this is a general
  tariff-revenue measure).
- All score columns (E, G, I, K, M, P, S, T, O, V) contain **numbers only**; O and V are
  **calculated**, consistent with the score-format fix applied across the other STP PRs
  (#104–#108).
- Given the absence of primary statutory text, implementation sub-scores (Column T) were coded
  conservatively (T = 0.5 for all rows: +0.25 for Customs-administered authority, +0.25 for
  unconditional application; no points awarded for enforcement or monitoring, since no such
  provisions were identified in the located material).

## English-language briefing (not a translation)

`Lei_4-2023_Customs_Duties_Law_EN_Briefing.docx` is **explicitly not a translation** of the law —
since no verbatim statutory text was located, no translation could honestly be produced. Instead,
it is an English-language briefing reconstructing the law's content, legislative history, and
constitutional/policy framing from the National Assembly's official record and the other secondary
sources listed above, with a detailed methodology note explaining these limitations and
recommending verification against the primary text at https://dre.gov.st/.
