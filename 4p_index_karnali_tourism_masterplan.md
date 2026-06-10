# Plastic Pollution Policy Index (4P Index) — Coding Table

**Document:** Karnali Province Tourism Master Plan 2076/77 – 2085/86 BS (2020/21–2029/30)
**Country:** Nepal (Karnali Province)
**Source URL:** https://faolex.fao.org/docs/pdf/nep220153.pdf
**Coded:** June 2026

---

## Instruments identified (3 rows total)

Three instruments passed the plastics relevance filter:
1. Ban on plastic use in Protected Areas (explicitly mentions plastic — Criterion a)
2. Solid waste management facilities and guidelines for tourism areas (waste management instrument that commonly applies to plastic waste streams — Criterion c)
3. Site-specific plastic/bottle bans at Syarpu Tal and Kubhinde Daha eco-tourism lake sites (explicitly mentions plastic — Criterion a)

---

## Section A — Policy-Level Fields (identical across all three rows)

| Column | Field | Value |
|--------|-------|-------|
| A | **policy_name** | Karnali Province Tourism Master Plan 2076/77 – 2085/86 BS (2020/21–2029/30) |
| B | **policy_url** | https://faolex.fao.org/docs/pdf/nep220153.pdf |
| C | **policy_year** | 2020 |
| D | **policy_objective** | To develop sustainable tourism in Karnali Province (Nepal) as an engine for economic and social transformation while preserving natural and cultural heritages; includes aspirational measures to ban plastic use in protected areas, establish solid waste management facilities and guidelines in tourism areas, and prevent plastic pollution at specific lake and wetland eco-tourism sites. |
| E | **policy_target** | 0 |
| F | **policy_target_text** | *(empty — no quantifiable plastic-specific targets)* |
| G | **policy_type** | **0.25** |
| H | **policy_type_justification** | Sub-national provincial tourism master plan published in January 2020 by the Ministry of Industry, Tourism, Forest and Environment (MoITFE) of Karnali Province, with technical assistance from WWF Nepal. Formulated through a participatory planning process and issued as a guiding policy document — not enacted by any legislature. The plan contains aspirational commitments on plastic bans and waste management, but without quantifiable plastic-specific targets. |
| I | **policy_integration** | **1** |
| J | **policy_sectors_list** | tourism, waste management, conservation, water, energy, transport, agriculture, municipalities, infrastructure |
| K | **policy_circularity** | **0.75** |
| L | **policy_lifecycle_phases_list** | consumption, disposal, environmental leakage |
| M | **policy_budget** | **0.5** |
| N | **policy_budget_text** | "The plan allocates a total budget of NRs. 10,300,000,000. Program 7 (Cultural and Natural Heritage Conservation), which encompasses the plastic ban and solid waste management measures, is assigned NRs. 1,030,000,000 (10%). Program 2 (Tourism Infrastructure Development and Upgrading), which includes waste management infrastructure, is assigned NRs. 3,605,000,000 (35%). Funding is to be drawn from government budgets at three tiers plus development partners/donors." (Executive Summary, p. iv–vii) |
| O | **policy_score** | [auto — average of G, I, K, M, and P] |

**Notes on policy-level fields:**
- **Col E = 0** because the plan's quantifiable targets are all tourism performance indicators (visitor arrivals, employment, GDP contribution) and contain no plastic-specific reduction targets.
- **Col I = 1** (7+ sectors): tourism, waste management, conservation, water, energy, transport, agriculture, municipalities, infrastructure.
- **Col K = 0.75** (3 life-cycle phases): consumption (plastic use by tourists in PAs and lake sites), disposal (waste management facilities and guidelines), environmental leakage (plastic in lakes, rivers, and protected natural areas).
- **Col M = 0.5**: The plan has a multi-programme budget with funding from three tiers of government plus donors, but no ring-fenced plastic-specific fund.

---

## Section B — Instrument-Level Fields

### Instrument 1 — Ban on plastic use in Protected Areas

| Column | Field | Value |
|--------|-------|-------|
| P | **instrument_type** | **1.0** — Regulatory (explicit ban on plastic use) |
| Q | **instrument_lifecycle_stage** | Consumption |
| R | **instrument_description** | Section 5.7.5 (p. 99), Programme 7 — Cultural and Natural Heritage Conservation: The plan commits to banning the use of plastic in Protected Areas as one of a suite of measures to minimise and mitigate potential negative impacts of tourism on natural heritage in Karnali Province. Text: *"Ban on use of plastic in Protected Area."* The relevant designated Protected Areas in Karnali Province are Rara National Park and Shey-Phoksundo National Park. |
| S | **instrument_in_force** | **0** — The ban appears in a planning strategy document, not an enacted regulation or binding ordinance. No subordinate regulation operationalising the ban was found. Language test: the imperative "Ban on…" sits within a bulleted action list in a provincial tourism master plan and does not carry direct legal force. |
| T | **instrument_implementation** | **0.50** (+0.25 responsible authority; +0.25 unconditional) |
| U | **instrument_implementation_text** | *Section 5.7.5 (p. 99):* "– Ban on use of plastic in Protected Area. Ensure safe and adequate waste and sewage management facilities in touristic areas." *Action Plan table (p. 127):* "7.5 Set the appropriate measures to minimise and mitigate the potential negative impacts of tourism / Implement solid waste management guidelines / **Leading: MoITFE, M/RM** / Partner: DFO, BZUC, PAs." Responsible authority designated at programme level (+0.25). No enforcement penalties or monitoring mechanism explicitly stated. No exemptions stated (+0.25). |
| V | **instrument_score** | [auto — average of P = 1.0 and T = 0.50] → **0.75** |
| W | **comments** | Coded as Regulatory (ban) based on explicit prohibition language, but instrument_in_force = 0 because the plan is not binding legislation and no subordinate regulation has been found that operationalises this ban. If a subsequent provincial or local government ordinance later enacts this ban, instrument_in_force should be updated to 1. The responsible authority (+0.25) is designated at the broader Programme 7.5 level (MoITFE, M/RM, DFO, BZUC, PAs), not for this specific instrument alone — this is a borderline case. No exemptions are stated (+0.25). Cross-reference: Nepal's nationwide Plastic Bag Control Directive 2082 (replacing Directive 2068) imposes standards for plastic bag production, import, and use nationally; this plan's PA-level ban is narrower in scope but may be more stringent. |

---

### Instrument 2 — Solid waste management facilities and guidelines for tourism areas

| Column | Field | Value |
|--------|-------|-------|
| P | **instrument_type** | **0.80** — Infrastructure (commitment to waste and sewage management facilities in touristic areas; waste collection and management system in high-altitude areas) |
| Q | **instrument_lifecycle_stage** | Waste management |
| R | **instrument_description** | Sections 5.7.5 and Chapter 7 Action Plan item 7.5: The plan commits to (a) ensuring safe and adequate waste and sewage management facilities in touristic areas; (b) controlling and managing solid waste and sludge in high-altitude areas affected by unmanaged Yarsagumba (Cordyceps) collection expeditions, including providing a waste collection and management system; and (c) implementing solid waste management guidelines. MoITFE and Municipalities/Rural Municipalities (M/RM) are the designated leading implementing institutions; DFO, BZUC, and Protected Area offices are partners. |
| S | **instrument_in_force** | **0** — Planning commitments in a strategy document; no specific waste facility has been constructed and no solid waste management guideline has been formally adopted as a binding instrument under this plan. |
| T | **instrument_implementation** | **0.50** (+0.25 responsible authority; +0.25 unconditional) |
| U | **instrument_implementation_text** | *Section 5.7.5 (p. 99):* "– Ensure safe and adequate waste and sewage management facilities in touristic areas." "– Control and manage solid waste, sludge and increasing pressure on grassland and vegetation cover in the high altitude areas resulted due to unplanned and unmanaged Yarsagumba collection. Use alternative energy/fuel, safe huts/shelter, waste collection and management system and first aid/health facilities." *Action Plan table (p. 127):* "7.5 Set the appropriate measures to minimise and mitigate the potential negative impacts of tourism / Implement solid waste management guidelines / **Leading: MoITFE, M/RM** / Partner: DFO, BZUC, PAs." Responsible authority explicitly designated (+0.25). No enforcement penalties stated. No monitoring mechanism explicitly stated. No exemptions stated (+0.25). |
| V | **instrument_score** | [auto — average of P = 0.80 and T = 0.50] → **0.65** |
| W | **comments** | This instrument combines infrastructure (waste and sewage facilities) and governance (solid waste guidelines) elements; the highest-scoring type — Infrastructure (0.80) — is applied per coding rules, with the governance/coordination aspect noted here. The instrument covers plastic waste streams implicitly as part of general solid waste management in tourism areas (no explicit mention of plastic in this specific instrument; the plastic references are in instruments 1 and 3). **Borderline case:** A case could be made for P = 0.20 (governance & coordination / planning requirement) if the waste facility commitment is viewed as a general aspiration rather than a specific infrastructure investment. Section 2.8.1 (p. 27) explicitly names plastic as a problem: "Plastic wrapper, mineral plastic bottles, glass bottles and tin that are not biodegradable pose threats to environment." This contextual passage confirms the relevance to plastic waste, even though the instrument itself uses general solid waste language. |

---

### Instrument 3 — Plastic/bottle bans at Syarpu Tal and Kubhinde Daha eco-tourism lake sites

| Column | Field | Value |
|--------|-------|-------|
| P | **instrument_type** | **1.0** — Regulatory (explicit ban on plastic and bottles) |
| Q | **instrument_lifecycle_stage** | Consumption |
| R | **instrument_description** | Appendix 1 — Strategic Tourism Projects (p. 138): The development plans for two lake-based eco-tourism hubs include explicit bans on plastic and bottles as lake conservation measures. (1) **Syarpu Tal Eco-tourism Hub** (West Rukum, Banfikot Rural Municipality): plan includes "ban on use of plastic and bottles" to conserve the lake ecosystem. (2) **Kubhinde Daha Eco-tourism Hub** (Salyan): same "ban on use of plastic and bottles." Both bans are listed alongside siltation control, drainage management, invasive species removal, and habitat restoration. |
| S | **instrument_in_force** | **0** — Project-level commitments in a planning document. The bans are embedded in descriptions of future eco-tourism development projects and are not currently operative regulations. |
| T | **instrument_implementation** | **0.25** (+0.25 unconditional only) |
| U | **instrument_implementation_text** | *Appendix 1, Syarpu Tal and Kubhinde Daha sections (p. 138):* "To conserve the lake, develop siltation controlling mechanism at the mouth (head), control draining house-hold sewage into lake water (develop outer drainage), re-store endemic fish by removing invasive imported fish species; plant local but ornamental trees and flowers around the lake road; **ban on use of plastic and bottles**." No responsible authority, enforcement mechanism, or monitoring mechanism is explicitly designated for these site-specific bans. No exemptions are stated (+0.25). |
| V | **instrument_score** | [auto — average of P = 1.0 and T = 0.25] → **0.625** |
| W | **comments** | The plastic/bottle bans at Syarpu Tal and Kubhinde Daha are project-level planning commitments embedded in descriptions of future eco-tourism development. These lakes are **not** currently designated national protected areas (Rara NP and Shey-Phoksundo NP are the PAs in Karnali Province), making this instrument distinct from Instrument 1. If these sites are subsequently declared provincial conservation areas, this ban could be subsumed under Instrument 1. T = 0.25 because only the 'unconditional' criterion is met (no exemptions stated); no designated authority, enforcement, or monitoring are explicitly evidenced. The identical phrasing at both lake sites suggests a template/boilerplate commitment rather than a site-specific policy decision. |

---

## Summary scoring

| | Instrument 1 | Instrument 2 | Instrument 3 |
|--|--|--|--|
| **instrument_type (P)** | 1.00 | 0.80 | 1.00 |
| **instrument_in_force (S)** | 0 | 0 | 0 |
| **instrument_implementation (T)** | 0.50 | 0.50 | 0.25 |
| **instrument_score (V) [auto]** | 0.75 | 0.65 | 0.625 |
| **policy_score (O) [auto, using avg P]** | *uses each row's P in formula* | | |

> **policy_score formula:** average of G (0.25), I (1.0), K (0.75), M (0.5), and P (varies by row)
> - Row 1: (0.25 + 1.0 + 0.75 + 0.5 + 1.0) / 5 = **0.70**
> - Row 2: (0.25 + 1.0 + 0.75 + 0.5 + 0.80) / 5 = **0.66**
> - Row 3: (0.25 + 1.0 + 0.75 + 0.5 + 1.0) / 5 = **0.70**

---

## Key evidence passages (verbatim from source document)

**Plastic named as a threat (p. 27, Section 2.8.1):**
> "Plastic wrapper, mineral plastic bottles, glass bottles and tin that are not biodegradable pose threats to environment."

**SWOT — Threats (p. 32):**
> "Increasing market nitche of industrial products geopardizing the pristine nature (i.e.environmental pollution such as water pollution, land/soil by the use of plastic, tin, glass and sewage to river/wetland etc.)"

**Plastic ban in PAs + waste management (p. 99, Section 5.7.5):**
> "– Ban on use of plastic in Protected Area. Ensure safe and adequate waste and sewage management facilities in touristic areas."
> "– Control and manage solid waste, sludge and increasing pressure on grassland and vegetation cover in the high altitude areas resulted due to unplanned and unmanaged Yarsagumba collection. Use alternative energy/fuel, safe huts/shelter, waste collection and management system and first aid/health facilities."

**Action Plan responsible authority (p. 127, Chapter 7):**
> "7.5 Set the appropriate measures to minimise and mitigate the potential negative impacts of tourism / Implement solid waste management guidelines / Leading: MoITFE, M/RM / Partner: DFO, BZUC, PAs"

**Site-specific plastic bans (p. 138, Appendix 1 — Syarpu Tal and Kubhinde Daha):**
> "To conserve the lake, develop siltation controlling mechanism at the mouth (head), control draining house-hold sewage into lake water (develop outer drainage), re-store endemic fish by removing invasive imported fish species; plant local but ornamental trees and flowers around the lake road; **ban on use of plastic and bottles**."
