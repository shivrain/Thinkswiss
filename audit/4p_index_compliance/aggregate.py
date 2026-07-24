#!/usr/bin/env python3
import json, re
from collections import defaultdict, OrderedDict

d = json.load(open("/tmp/4p_audit/results.json"))

def norm_msg(msg):
    # strip specific values/quotes to get a generic signature
    m = re.sub(r"row \d+", "row N", msg)
    m = re.sub(r"='[^']*'", "='...'", m)
    m = re.sub(r"\(O\)=[0-9.]+", "(O)=X", m)
    m = re.sub(r"\(V\)=[0-9.]+", "(V)=X", m)
    m = re.sub(r"= \([^)]*\)/4 = [0-9.]+", "= (...)/4 = Y", m)
    m = re.sub(r"contains '[^']*'", "contains '...'", m)
    m = re.sub(r"entry '[^']*'", "entry '...'", m)
    m = re.sub(r"'[^']{0,120}'", "'...'", m)
    m = re.sub(r"\d+ sector\(s\)", "N sector(s)", m)
    m = re.sub(r"lists \d+ phase", "lists N phase", m)
    return m

summary = OrderedDict()
for fname, v in d.items():
    refs = v["refs"]
    issues = v["issues"]
    if not issues:
        continue
    sigmap = defaultdict(lambda: {"count":0, "rows": [], "severity": None, "col": None, "example": None})
    for iss in issues:
        sig = norm_msg(iss["issue"])
        key = (iss["col"], iss["severity"], sig)
        e = sigmap[key]
        e["count"] += 1
        if iss["row"] is not None:
            e["rows"].append(iss["row"])
        e["severity"] = iss["severity"]
        e["col"] = iss["col"]
        if e["example"] is None:
            e["example"] = iss["issue"]
    summary[fname] = {"refs": refs, "signatures": sigmap}

out_lines = []
n_files_with_error = 0
n_files_with_review = 0
for fname, info in summary.items():
    refs = info["refs"]
    sigmap = info["signatures"]
    has_error = any(v["severity"]=="ERROR" for v in sigmap.values())
    has_review = any(v["severity"] in ("REVIEW","WARN") for v in sigmap.values())
    if has_error: n_files_with_error += 1
    if has_review: n_files_with_review += 1
    ref_str = "; ".join(f"{b} :: {p}" for b,p in refs[:3])
    if len(refs) > 3:
        ref_str += f" (+{len(refs)-3} more refs)"
    out_lines.append(f"\n### FILE: {fname}")
    out_lines.append(f"refs: {ref_str}")
    for (col, sev, sig), e in sorted(sigmap.items(), key=lambda kv: (kv[0][1]!="ERROR", kv[0][0] or "")):
        rows = e["rows"]
        rowstr = ""
        if rows:
            rows_sorted = sorted(set(rows))
            if len(rows_sorted) > 8:
                rowstr = f" [rows {rows_sorted[0]}-{rows_sorted[-1]}, n={len(rows_sorted)}]"
            else:
                rowstr = f" [rows {rows_sorted}]"
        out_lines.append(f"  - ({sev}) col {col}: {sig}{rowstr} (x{e['count']})")
        out_lines.append(f"      e.g. {e['example'][:260]}")

with open("/tmp/4p_audit/summary.txt", "w") as f:
    f.write("\n".join(out_lines))

print(f"Files with >=1 ERROR: {n_files_with_error} / {len(summary)}")
print(f"Files with >=1 REVIEW/WARN: {n_files_with_review} / {len(summary)}")
print(f"Total unique files with issues: {len(summary)}")
