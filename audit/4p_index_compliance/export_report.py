#!/usr/bin/env python3
import json, csv, re
from collections import OrderedDict

d = json.load(open("/tmp/4p_audit/results.json"))

def pick_primary_ref(fname, refs):
    # Prefer the (branch, path) pair whose branch name most closely matches the filename.
    base = re.sub(r"[^a-z0-9]+", "", fname.lower())
    best = refs[0]
    best_score = -1
    for branch, path in refs:
        pathbase = re.sub(r"[^a-z0-9]+", "", path.rsplit("/", 1)[-1].lower())
        score = len(set(pathbase) & set(base))
        # Prefer .xlsx path text similarity to branch name too
        branchbase = re.sub(r"[^a-z0-9]+", "", branch.lower())
        score += len(set(branchbase) & set(pathbase))
        if score > best_score:
            best_score = score
            best = (branch, path)
    return best

rows_out = []
for fname, v in d.items():
    refs = v["refs"]
    if not refs:
        continue
    primary_branch, primary_path = pick_primary_ref(fname, refs)
    other_refs = [f"{b}:{p}" for b, p in refs if (b, p) != (primary_branch, primary_path)]
    for iss in v["issues"]:
        rows_out.append({
            "branch": primary_branch,
            "file_path": primary_path,
            "also_appears_in": " | ".join(other_refs),
            "spreadsheet_row": iss["row"],
            "column": iss["col"],
            "severity": iss["severity"],
            "finding": iss["issue"],
        })

with open("/tmp/4p_audit/findings_detail.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=["branch","file_path","also_appears_in","spreadsheet_row","column","severity","finding"])
    w.writeheader()
    for r in rows_out:
        w.writerow(r)

print(f"Wrote {len(rows_out)} findings to findings_detail.csv")

# Build a per-file summary markdown, grouped by primary branch/file
by_file = OrderedDict()
for fname, v in d.items():
    refs = v["refs"]
    if not refs:
        continue
    primary_branch, primary_path = pick_primary_ref(fname, refs)
    by_file[(primary_branch, primary_path)] = v["issues"]

def norm_msg(msg):
    m = re.sub(r"row \d+", "row N", msg)
    m = re.sub(r"='[^']*'", "='...'", m)
    return m

md = []
md.append("# 4P Index Coding Compliance Audit\n")
md.append("Automated + heuristic audit of every `4P Index` coding spreadsheet found across all repository branches, "
           "checked against the 4P Index Policy Coding Prompt (v2).\n")

n_err = sum(1 for issues in by_file.values() for i in issues if i["severity"]=="ERROR")
n_warn = sum(1 for issues in by_file.values() for i in issues if i["severity"]=="WARN")
n_review = sum(1 for issues in by_file.values() for i in issues if i["severity"]=="REVIEW")
md.append(f"- Files audited: {len(by_file)}\n- ERROR findings (hard rule violations): {n_err}\n"
          f"- WARN findings (count/consistency mismatches): {n_warn}\n"
          f"- REVIEW findings (Rule 11/12 heuristic flags needing human judgement): {n_review}\n")

for (branch, path), issues in by_file.items():
    if not issues:
        continue
    md.append(f"\n## `{path}`  (branch: `{branch}`)\n")
    sig_seen = OrderedDict()
    for iss in issues:
        sig = (iss["col"], iss["severity"], norm_msg(iss["issue"]))
        sig_seen.setdefault(sig, {"rows": [], "example": iss["issue"]})
        if iss["row"] is not None:
            sig_seen[sig]["rows"].append(iss["row"])
    for (col, sev, sig), info in sorted(sig_seen.items(), key=lambda kv: (kv[0][1] != "ERROR", kv[0][0] or "")):
        rows = sorted(set(info["rows"]))
        rowtxt = f"rows {rows}" if len(rows) <= 10 else f"rows {rows[0]}-{rows[-1]} (n={len(rows)})"
        md.append(f"- **[{sev}] Col {col}** — {rowtxt}\n  - {info['example']}\n")

with open("/tmp/4p_audit/AUDIT_REPORT.md", "w") as f:
    f.write("\n".join(md))

print("Wrote AUDIT_REPORT.md")
