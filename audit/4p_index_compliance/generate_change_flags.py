#!/usr/bin/env python3
"""Turn findings_detail.csv into an actionable per-file change list."""
import csv, re
from collections import defaultdict, OrderedDict

IN_CSV = "findings_detail.csv"
OUT_MD = "CHANGES_TO_MAKE.md"
OUT_CSV = "changes_flagged.csv"

COL_NAMES = {
    "A": "policy_name", "B": "policy_url", "C": "policy_year", "D": "policy_objective",
    "E": "policy_target", "F": "policy_target_text", "G": "policy_type",
    "H": "policy_type_justification", "I": "policy_integration", "J": "policy_sectors_list",
    "K": "policy_circularity", "L": "policy_lifecycle_phases_list", "M": "policy_budget",
    "N": "policy_budget_text", "O": "policy_score", "P": "instrument_type",
    "Q": "instrument_lifecycle_stage", "R": "instrument_description",
    "S": "instrument_in_force", "T": "instrument_implementation",
    "U": "instrument_implementation_text", "V": "instrument_score", "W": "comments",
}

Q_REMAP = {
    "Disposal": "Waste management or End of life (pick closest fit)",
    "End-of-life": "End of life",
    "Collection": "Waste management",
    "Litter/Pollution": "Environmental leakage",
    "Planning": "Production (or add W note if purely procedural)",
    "Design": "Production",
    "Use/Consumption": "Consumption",
}

def country_of(branch):
    b = branch.lower()
    if "peru" in b: return "Peru"
    if "mozambique" in b: return "Mozambique"
    if "senegal" in b: return "Senegal"
    if "code-stp" in b or "stp" in b: return "São Tomé and Príncipe"
    if "karnali" in b or "nepal" in b: return "Nepal"
    if "4p-index-" in b and "65aa" in b: return "Norway"
    return "Other"

def action_for(finding, col, severity):
  f = finding
  if "inserts column(s) not defined" in f:
      return "Restructure sheet: remove extra columns (e.g. policy_name_en, country) OR move them after column W so A–W match the prompt exactly."
  if "varies across rows of the same policy" in f and col == "O":
      return "Set Col O to ONE constant value for the whole policy: O = (G+I+K+M)/4. Copy identical value into every row."
  if "does not equal (G+I+K+M)/4" in f and col == "O":
      m = re.search(r"= \(([^)]+)\)/4 = ([0-9.]+)", f)
      if m:
          return f"Recalculate Col O = ({m.group(1)})/4 = {m.group(2)} and apply to ALL rows of this policy."
      return "Recalculate Col O = (G+I+K+M)/4 and apply to ALL rows."
  if "does not equal S×(P+T)/2 = 0.0×" in f and col == "V":
      m = re.search(r"S×\(P\+T\)/2 = 0\.0×\(([0-9.]+)\+([0-9.]+)\)/2 = 0\.000", f)
      if m:
          return f"Set Col V = 0 (because S=0). Current V wrongly uses (P+T)/2 = ({m.group(1)}+{m.group(2)})/2."
      return "Set Col V = 0 because S=0 (prompt: V = S×(P+T)/2)."
  if "does not equal S×(P+T)/2" in f and col == "V":
      m = re.search(r"S×\(P\+T\)/2 = ([0-9.]+)×\(([0-9.]+)\+([0-9.]+)\)/2 = ([0-9.]+)", f)
      if m:
          return f"Set Col V = {m.group(4)} (= S×(P+T)/2 = {m.group(1)}×({m.group(2)}+{m.group(3)})/2)."
      return "Recalculate Col V = S×(P+T)/2."
  if "Policy-level field" in f and "differs across rows" in f:
      cn = COL_NAMES.get(col, col)
      return f"Copy the same Col {col} ({cn}) value into every instrument row for this policy (Rule 1)."
  if "should be EMPTY when E=0" in f:
      return "Clear Col F — leave empty when E=0 (Rule 5; do not write N/A)."
  if "is empty but E=" in f and col == "F":
      return "Paste exact plastic-specific target text into Col F (required when E > 0)."
  if "should be EMPTY when M=0" in f:
      return "Clear Col N — leave empty when M=0."
  if "is empty but M=" in f and col == "N":
      return "Insert relevant budget/funding text into Col N (required when M > 0)."
  if col == "G" and "not one of the allowed values" in f:
      return "Change Col G to one of: 0.25, 0.50, 0.75, 1.0 (likely 0.25 if currently 0.2)."
  if col == "M" and "not one of the allowed values" in f:
      return "Change Col M to one of: 0, 0.5, 1.0 only (0.75 is invalid)."
  if col == "P" and "not one of the allowed values" in f:
      return "Change Col P to one of: 0, 0.20, 0.40, 0.60, 0.80, 1.0 only."
  if col == "Q" and "not one of the allowed title-case values" in f:
      m = re.search(r"\(Q\)='([^']*)'", f)
      bad = m.group(1) if m else "?"
      fix = Q_REMAP.get(bad, "one of: Production, Consumption, Recycling, Waste management, End of life, Environmental leakage")
      return f"Change Col Q from '{bad}' → '{fix}' (Rule 4)."
  if col == "L" and "not one of the 5 allowed phases" in f:
      m = re.search(r"contains '([^']*)'", f)
      bad = m.group(1) if m else "?"
      return f"Replace '{bad}' in Col L with one of the 5 allowed lower-case terms: production, consumption, recycling, disposal, environmental leakage (Rule 3)."
  if col == "L" and "entry" in f and "must be lower case" in f:
      return "Use lower case only in Col L (Rule 3)."
  if col == "J" and "not lower case" in f:
      return "Use lower case only for all sector entries in Col J."
  if col == "J" and "inconsistent with policy_integration" in f:
      return "Recount sectors in Col J and adjust Col I to match bucket (0→0 sectors, 0.25→1–2, 0.5→3–4, 0.75→5–6, 1→7+), OR fix J list."
  if col == "L" and "inconsistent with policy_circularity" in f:
      return "Recount lifecycle phases in Col L and adjust Col K to match bucket (0.25→1, 0.5→2, 0.75→3–4, 1→5), OR fix L list."
  if "Per Rule 12" in f and "S)=1" in f:
      return "Re-check Col S: enabling-power language only → set S=0 unless subordinate instrument already operationalised (Rule 12). Document in Col W."
  if "Per Rule 12" in f and "S)=0" in f:
      return "Re-check Col S: mandatory language present → likely set S=1 (Rule 12). Document in Col W."
  if "Per Rule 11" in f:
      return "Either set Col P=0.40 (procedural/coordination) OR keep higher P and add coordination note in Col W (Rule 11)."
  if col == "C" and "4-digit year" in f:
      return "Col C must be a 4-digit year (amendment year if amended)."
  if col == "E" and "not one of the allowed values" in f:
      return "Col E must be exactly 0, 0.5, or 1."
  if severity == "ERROR":
      return f"Fix Col {col} ({COL_NAMES.get(col, col)}): {f[:200]}"
  if severity == "WARN":
      return f"Review Col {col}: {f[:200]}"
  return f[:250]

# Presentation removals (PR #112) — policies to delete entirely
REMOVALS = {
    "Nepal": [
        ("The Excise Duty Act, 2002", "Not a plastic-relevant included policy"),
        ("WTO Trade Policy Review — Report by the Secretariat — Nepal", "Non-official document; not a national policy instrument"),
        ("Ban on Plastic Bottles in High-Class Hotels — Directive / Sectoral Order (2024)", "Source document not found"),
    ],
    "Mozambique": [
        ("Regulation on Environmental Quality and Effluent Standards", "Air-quality standard — not a relevant plastic instrument (Art. 8)"),
    ],
    "São Tomé and Príncipe": [
        ("National Plan for Integrated Urban Solid Waste Management (PNGIRSU) 2018–2023", "Projection/inventory — not an instrument"),
        ("PNGIRSU 2018–2023 — National Waste Characterisation Exercise", "Data collection exercise — not an instrument"),
    ],
}

# Presentation score corrections
SCORE_FIXES = {
    ("Mozambique", "National Strategy for the Management and Conservation of Coral Reefs"): [
        ("M", "0.75", "0.5 or 1.0", "Line-item budgets exist but 0.75 is not on the M scale"),
    ],
    ("Peru", "Supreme Decree No. 001-2022-MINAM"): [
        ("M", "0.5", "1.0", "Institutional financing mechanism + cleaning-fee agreements = ring-fenced budget source"),
    ],
    ("Peru", "Law that Amends Legislative Decree No. 1278"): [
        ("M", "0.75", "0.5", "Investment promotion mentioned but no specific budget line; 0.75 invalid"),
    ],
    ("Peru", "Article 9 of the Regulation of Legislative Decree No. 1278 (PLANRES Provision)"): [
        ("E", "1", "0", "PLANRES is periodic planning/reporting, not a quantifiable plastic-specific target"),
        ("F", "(text)", "(empty)", "Clear F when E=0"),
    ],
}

rows = list(csv.DictReader(open(IN_CSV, newline="", encoding="utf-8")))

by_file = OrderedDict()
for r in rows:
    key = (r["branch"], r["file_path"])
    by_file.setdefault(key, []).append(r)

# Sort: country, then file
sorted_files = sorted(by_file.keys(), key=lambda k: (country_of(k[0]), k[1].lower()))

md = []
md.append("# 4P Index v2 — All Changes To Make\n")
md.append("Action list generated from strict compliance check against the 4P Index Policy Coding Prompt v2.\n")
md.append("Each file below lists **every row/column change** required. Severity: **MUST FIX** = definite rule violation; **REVIEW** = human judgement (Rules 11/12); **CHECK** = consistency warning.\n")

csv_out = []

# Global removals section
md.append("## A. Policies to remove from the dataset entirely\n")
md.append("From the July 2026 review presentation (PR #112) — delete these policies/rows, not just edit a cell:\n")
for country, items in REMOVALS.items():
    md.append(f"### {country}\n")
    for name, reason in items:
        md.append(f"- **DELETE** `{name}` — {reason}\n")
        csv_out.append({
            "country": country, "branch": "(any)", "file_path": "(search dataset)",
            "spreadsheet_row": "", "column": "—", "severity": "DELETE_POLICY",
            "change_required": f"Remove policy entirely: {name}", "detail": reason,
        })

md.append("\n## B. Policy-level score corrections from review (PR #112)\n")
for (country, policy), fixes in SCORE_FIXES.items():
    md.append(f"### {country} — `{policy}`\n")
    for col, old, new, reason in fixes:
        md.append(f"- **Col {col}**: change `{old}` → `{new}` — {reason}\n")
        csv_out.append({
            "country": country, "branch": "(see file)", "file_path": policy,
            "spreadsheet_row": "all rows", "column": col, "severity": "MUST FIX",
            "change_required": f"Change {col} from {old} to {new}", "detail": reason,
        })

md.append("\n## C. File-by-file spreadsheet changes\n")

file_count = 0
for (branch, path) in sorted_files:
    issues = by_file[(branch, path)]
    country = country_of(branch)
    file_count += 1

    # Group issues by row+col+action to dedupe
    seen = OrderedDict()
    for iss in issues:
        act = action_for(iss["finding"], iss["column"], iss["severity"])
        key = (iss["spreadsheet_row"], iss["column"], act)
        if key not in seen:
            seen[key] = iss["severity"]
    if not seen:
        continue

    md.append(f"\n### {country} — `{path}`\n")
    md.append(f"Branch: `{branch}`\n")

    must = review = check = 0
    for (row, col, act), sev in seen.items():
        if sev == "ERROR":
            label = "MUST FIX"
            must += 1
        elif sev == "REVIEW":
            label = "REVIEW"
            review += 1
        else:
            label = "CHECK"
            check += 1
        rowtxt = f"row {row}" if row else "whole file"
        coltxt = f"Col {col}" if col else "structure"
        md.append(f"- **[{label}]** {rowtxt}, {coltxt}: {act}\n")
        csv_out.append({
            "country": country, "branch": branch, "file_path": path,
            "spreadsheet_row": row, "column": col, "severity": label,
            "change_required": act, "detail": "",
        })
    md.append(f"\n_Summary: {must} must-fix, {review} review, {check} check_\n")

md.append(f"\n---\n\n**Total files with changes: {file_count}**\n")

with open(OUT_MD, "w", encoding="utf-8") as f:
    f.write("\n".join(md))

with open(OUT_CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, fieldnames=["country","branch","file_path","spreadsheet_row","column","severity","change_required","detail"])
    w.writeheader()
    w.writerows(csv_out)

print(f"Wrote {OUT_MD} ({len(md)} lines) and {OUT_CSV} ({len(csv_out)} rows)")
