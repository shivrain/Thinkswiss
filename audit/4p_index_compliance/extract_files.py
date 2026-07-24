#!/usr/bin/env python3
import subprocess, hashlib, os, sys

MANIFEST = "/tmp/4p_audit/files_manifest.txt"
OUTDIR = "/tmp/4p_audit/raw"
os.makedirs(OUTDIR, exist_ok=True)

seen_hash = {}
records = []

with open(MANIFEST) as f:
    lines = [l.rstrip("\n") for l in f if l.strip()]

for line in lines:
    branch, path = line.split("\t", 1)
    try:
        data = subprocess.run(
            ["git", "show", f"origin/{branch}:{path}"],
            cwd="/workspace", capture_output=True, check=True
        ).stdout
    except subprocess.CalledProcessError as e:
        print(f"FAILED {branch}:{path}: {e.stderr.decode(errors='replace')[:200]}", file=sys.stderr)
        continue
    h = hashlib.sha256(data).hexdigest()
    if h not in seen_hash:
        ext = path.rsplit(".", 1)[-1]
        fname = f"{h[:12]}.{ext}"
        with open(os.path.join(OUTDIR, fname), "wb") as out:
            out.write(data)
        seen_hash[h] = fname
    records.append((branch, path, seen_hash[h]))

with open("/tmp/4p_audit/dedup_manifest.tsv", "w") as f:
    for branch, path, fname in records:
        f.write(f"{branch}\t{path}\t{fname}\n")

print(f"Total refs: {len(records)}, unique files: {len(seen_hash)}")
