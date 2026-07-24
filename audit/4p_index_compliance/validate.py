#!/usr/bin/env python3
"""Strict structural/mechanical + heuristic semantic validator for 4P Index coding files."""
import os, re, csv, json
import openpyxl

RAW = "/tmp/4p_audit/raw"
DEDUP_MANIFEST = "/tmp/4p_audit/dedup_manifest.tsv"

COLS = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W"]
COLNAMES = {
    "A":"policy_name","B":"policy_url","C":"policy_year","D":"policy_objective",
    "E":"policy_target","F":"policy_target_text","G":"policy_type","H":"policy_type_justification",
    "I":"policy_integration","J":"policy_sectors_list","K":"policy_circularity","L":"policy_lifecycle_phases_list",
    "M":"policy_budget","N":"policy_budget_text","O":"policy_score","P":"instrument_type",
    "Q":"instrument_lifecycle_stage","R":"instrument_description","S":"instrument_in_force",
    "T":"instrument_implementation","U":"instrument_implementation_text","V":"instrument_score","W":"comments",
}

ALLOWED_LIFECYCLE = {"production","consumption","recycling","disposal","environmental leakage"}
ALLOWED_Q = {"Production","Consumption","Recycling","Waste management","End of life","Environmental leakage"}

def load_manifest():
    file_to_refs = {}
    with open(DEDUP_MANIFEST) as f:
        for line in f:
            branch, path, fname = line.rstrip("\n").split("\t")
            file_to_refs.setdefault(fname, []).append((branch, path))
    return file_to_refs

def read_xlsx(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    # Prefer a sheet whose name suggests the main coding table; else pick the sheet
    # with the most rows / that contains the header pattern.
    best = None
    best_score = -1
    for ws in wb.worksheets:
        rows = list(ws.iter_rows(values_only=True))
        score = 0
        title_l = ws.title.lower()
        if "4p" in title_l or "coding" in title_l or "index" in title_l:
            score += 100
        idx, _ = map_header(rows)
        if idx is not None:
            score += 50
        score += min(len(rows), 50)
        if score > best_score:
            best_score = score
            best = rows
    return best if best is not None else []

def read_csv(path):
    with open(path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        return [tuple(r) for r in reader]

def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s == "" or s.lower() in ("n/a", "na", "none", "-"):
        return None
    try:
        return float(s)
    except ValueError:
        # try to extract leading number like "1 (Regulatory)"
        m = re.match(r"^\s*([0-9]*\.?[0-9]+)", s)
        if m:
            return float(m.group(1))
        return None

def is_empty(v):
    if v is None:
        return True
    s = str(v).strip()
    return s == "" or s.lower() in ("n/a", "na", "none")

def approx(a, b, tol=0.02):
    if a is None or b is None:
        return False
    return abs(a - b) <= tol

HEADER_PATTERNS = [
    re.compile(r"^\s*([A-Z])\s*[—\-–]\s*(.+)$", re.S),
    re.compile(r"^\s*Col\.?\s*([A-Z])\s*[\n:—\-–]\s*(.+)$", re.S | re.I),
    re.compile(r"^\s*([A-Z])_(\w+)$", re.S),
]

# Canonical field-name key -> our fixed schema letter. Matched against the
# de-lettered header text so files that insert extra (non-spec) columns such
# as "country" or a split policy_name_es/policy_name_en, which shift every
# subsequent letter, are still mapped to the RIGHT semantic field rather than
# trusting a stale/incorrect letter the file itself printed in its header.
FIELD_KEY_TO_LETTER = [
    ("policy_target_text", "F"), ("policy_type_justification", "H"),
    ("policy_sectors_list", "J"), ("policy_lifecycle_phases_list", "L"),
    ("policy_budget_text", "N"), ("policy_score", "O"),
    ("policy_objective", "D"), ("policy_integration", "I"),
    ("policy_circularity", "K"), ("policy_budget", "M"),
    ("policy_target", "E"), ("policy_type", "G"), ("policy_year", "C"),
    ("policy_url", "B"), ("policy_name_en", "A"), ("policy_name", "A"),
    ("instrument_lifecycle_stage", "Q"), ("instrument_description", "R"),
    ("instrument_implementation_text", "U"), ("instrument_implementation", "T"),
    ("instrument_in_force", "S"), ("instrument_type", "P"),
    ("instrument_score", "V"), ("comments", "W"),
]

def _norm_key(s):
    s = s.lower()
    s = re.sub(r"[\s\-]+", "_", s)
    s = re.sub(r"[^a-z_]", "", s)
    return s

def map_header(rows):
    """Find header row and build a col-letter -> cell-index map keyed off recognised
    4P Index field names (robust to files that add extra non-spec columns or use a
    letter prefix that no longer matches the canonical schema)."""
    for hi, row in enumerate(rows[:8]):
        by_letter_literal = {}
        by_key = {}
        for idx, cell in enumerate(row):
            if cell is None:
                continue
            s = str(cell)
            colname_text = s
            for pat in HEADER_PATTERNS:
                m = pat.match(s)
                if m:
                    by_letter_literal[m.group(1)] = idx
                    colname_text = m.group(2)
                    break
            key = _norm_key(colname_text)
            for field_key, letter in FIELD_KEY_TO_LETTER:
                if field_key in key and letter not in by_key:
                    by_key[letter] = idx
                    break
        # Prefer the semantic (field-name) mapping since it is robust to extra
        # inserted columns; fall back to the literal letter mapping otherwise.
        mapping = by_key if len(by_key) >= 15 else by_letter_literal
        if len(mapping) >= 15:
            return hi, mapping
    return None, None

def analyze_file(fname, rows, refs):
    issues = []
    header_idx, colmap = map_header(rows)
    if colmap is None:
        issues.append({"row": None, "col": None, "severity": "ERROR",
                        "issue": "Could not locate/parse header row with 'X — colname' pattern."})
        return issues

    # Detect non-spec structural drift: does this file's own printed letter for a
    # given field differ from where the prompt's canonical schema says it must be?
    header_row = rows[header_idx]
    literal_letters = {}
    for idx, cell in enumerate(header_row):
        if cell is None:
            continue
        for pat in HEADER_PATTERNS:
            m = pat.match(str(cell))
            if m:
                literal_letters[idx] = m.group(1)
                break
    extra_cols = [idx for idx in range(len(header_row))
                  if idx not in colmap.values() and header_row[idx] not in (None, "")]
    drift = any(colmap.get(L) is not None and literal_letters.get(colmap[L]) not in (L, None) for L in colmap)
    if drift or extra_cols:
        detail = []
        for L, idx in sorted(colmap.items()):
            lit = literal_letters.get(idx)
            if lit and lit != L:
                detail.append(f"file labels its '{COLNAMES[L]}' column as '{lit}' (should be '{L}' per the prompt)")
        extra_names = [str(header_row[idx]) for idx in extra_cols]
        msg = ("This file inserts column(s) not defined in the prompt's fixed A-W schema "
               f"({extra_names}), which shifts every subsequent column's letter out of alignment. ")
        if detail:
            msg += "For example: " + "; ".join(detail[:6]) + ". "
        msg += ("Per the prompt, only columns A through W (in that exact order/meaning) should be used; "
                "extra fields should either be dropped or appended after column W, not inserted earlier.")
        issues.append({"row": None, "col": None, "severity": "ERROR", "issue": msg})

    data_rows = rows[header_idx+1:]
    # drop fully-empty trailing rows
    data_rows = [r for r in data_rows if any(not is_empty(c) for c in r)]

    def get(row, letter):
        idx = colmap.get(letter)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    if not data_rows:
        issues.append({"row": None, "col": None, "severity": "ERROR", "issue": "No data rows found."})
        return issues

    # group by policy_name (col A) to check policy-level invariance
    groups = {}
    for i, row in enumerate(data_rows):
        pname = get(row, "A")
        groups.setdefault(pname, []).append((i, row))

    policy_level_letters = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N"]

    for pname, items in groups.items():
        ref_row = items[0][1]
        ref_vals = {L: get(ref_row, L) for L in policy_level_letters}
        for i, row in items:
            rn = i + header_idx + 2  # 1-indexed spreadsheet row incl header
            for L in policy_level_letters:
                v = get(row, L)
                rv = ref_vals[L]
                # normalize whitespace for string compare
                nv = str(v).strip() if v is not None else None
                nrv = str(rv).strip() if rv is not None else None
                if nv != nrv:
                    issues.append({"row": rn, "col": L, "severity": "ERROR",
                                    "issue": f"Policy-level field {COLNAMES[L]} differs across rows of the same policy "
                                             f"(row {rn}='{v}' vs row {items[0][0]+header_idx+2}='{rv}'). "
                                             f"Columns A-O must be identical across all instrument rows for a policy."})
        # Col O must also be identical across rows (policy-level, though excluded above since
        # it is often left blank/[auto] on some rows and filled on others).
        o_vals = []
        for i, row in items:
            rawo = get(row, "O")
            if not is_empty(rawo) and not re.match(r"^\[?auto\]?$", str(rawo).strip(), re.I):
                o_vals.append((i + header_idx + 2, num(rawo)))
        if len(o_vals) > 1:
            distinct = set(v for _, v in o_vals if v is not None)
            if len(distinct) > 1:
                rowlist = [r for r, _ in o_vals]
                issues.append({"row": rowlist[0], "col": "O", "severity": "ERROR",
                                "issue": f"policy_score (O) varies across rows of the same policy ({sorted(distinct)}); "
                                         f"per the prompt O is a policy-level field and MUST be identical across every "
                                         f"instrument row for this policy (rows {rowlist})."})

    for i, row in enumerate(data_rows):
        rn = i + header_idx + 2
        A = get(row, "A"); C = get(row, "C")
        E = num(get(row, "E")); F = get(row, "F")
        G = num(get(row, "G")); I_ = num(get(row, "I")); J = get(row, "J")
        K = num(get(row, "K")); L = get(row, "L")
        M = num(get(row, "M")); N = get(row, "N"); O = num(get(row, "O"))
        P = num(get(row, "P")); Q = get(row, "Q"); R = get(row, "R")
        S = num(get(row, "S")); T = num(get(row, "T")); U = get(row, "U")
        V = num(get(row, "V")); W = get(row, "W")

        def bad(col, msg, sev="ERROR"):
            issues.append({"row": rn, "col": col, "severity": sev, "issue": msg})

        # C: 4-digit year
        if C is not None and not re.match(r"^\d{4}$", str(C).strip()):
            bad("C", f"policy_year should be a 4-digit year; found '{C}'.")

        # E
        if E is None:
            bad("E", "policy_target (E) is empty; must be 0, 0.5, or 1.")
        elif E not in (0, 0.5, 1):
            bad("E", f"policy_target (E)={E} is not one of the allowed values {{0, 0.5, 1}}.")

        # F vs E
        if E == 0 and not is_empty(F):
            bad("F", f"policy_target_text (F) should be EMPTY when E=0 (Rule 5), but contains text: '{str(F)[:80]}...'")
        if E is not None and E > 0 and is_empty(F):
            bad("F", f"policy_target_text (F) is empty but E={E}>0; exact target text must be copied in when Col E > 0.")

        # G
        if G not in (0.25, 0.5, 0.75, 1.0, None):
            bad("G", f"policy_type (G)={G} is not one of the allowed values {{0.25, 0.5, 0.75, 1}}.")
        elif G is None:
            bad("G", "policy_type (G) is empty.")

        # I
        if I_ not in (0, 0.25, 0.5, 0.75, 1.0, None):
            bad("I", f"policy_integration (I)={I_} is not one of the allowed values {{0, 0.25, 0.5, 0.75, 1}}.")
        elif I_ is None:
            bad("I", "policy_integration (I) is empty.")
        else:
            n_sectors = 0 if is_empty(J) else len([s for s in str(J).split(",") if s.strip()])
            bucket_ok = {
                0: n_sectors == 0, 0.25: 1 <= n_sectors <= 2, 0.5: 3 <= n_sectors <= 4,
                0.75: 5 <= n_sectors <= 6, 1.0: n_sectors >= 7,
            }.get(I_, True)
            if not bucket_ok:
                bad("J", f"policy_sectors_list (J) lists {n_sectors} sector(s) ('{J}'), inconsistent with policy_integration "
                          f"(I)={I_} bucket. Recount sectors or adjust I.", sev="WARN")
            if not is_empty(J):
                # lowercase check
                items_ = [s.strip() for s in str(J).split(",") if s.strip()]
                for it in items_:
                    if it != it.lower():
                        bad("J", f"policy_sectors_list (J) entry '{it}' is not lower case (Coding Rules require lower case, comma-separated).", sev="WARN")

        # K / L
        if K not in (0.25, 0.5, 0.75, 1.0, None):
            bad("K", f"policy_circularity (K)={K} is not one of the allowed scale values {{0.25, 0.5, 0.75, 1}}.")
        elif K is None:
            bad("K", "policy_circularity (K) is empty.")
        if not is_empty(L):
            phases = [s.strip() for s in str(L).split(",") if s.strip()]
            for p in phases:
                if p.lower() not in ALLOWED_LIFECYCLE:
                    bad("L", f"policy_lifecycle_phases_list (L) contains '{p}', not one of the 5 allowed phases "
                              f"(production, consumption, recycling, disposal, environmental leakage).")
                elif p != p.lower():
                    bad("L", f"policy_lifecycle_phases_list (L) entry '{p}' must be lower case (Rule 3).")
            if K is not None:
                n = len(phases)
                bucket_ok = {0.25: n == 1, 0.5: n == 2, 0.75: 3 <= n <= 4, 1.0: n == 5}.get(K, True)
                if not bucket_ok:
                    bad("L", f"policy_lifecycle_phases_list (L) lists {n} phase(s), inconsistent with policy_circularity "
                              f"(K)={K} bucket.", sev="WARN")
        elif K is not None:
            bad("L", f"policy_lifecycle_phases_list (L) is empty but policy_circularity (K)={K}>0; list must be exhaustive (Rule 9).")

        # M / N
        if M not in (0, 0.5, 1.0, None):
            bad("M", f"policy_budget (M)={M} is not one of the allowed values {{0, 0.5, 1}}.")
        elif M is None:
            bad("M", "policy_budget (M) is empty.")
        if M == 0 and not is_empty(N):
            bad("N", "policy_budget_text (N) should be EMPTY when M=0, but contains text.")
        if M is not None and M > 0 and is_empty(N):
            bad("N", f"policy_budget_text (N) is empty but M={M}>0; relevant budget/funding text must be inserted.")

        # O formula check — blank or an explicit "[auto]"-style placeholder is compliant
        raw_O = get(row, "O")
        if not is_empty(raw_O):
            is_placeholder = bool(re.match(r"^\[?auto\]?$", str(raw_O).strip(), re.I))
            if O is None and not is_placeholder:
                bad("O", f"policy_score (O) contains non-numeric, non-placeholder content '{str(raw_O)[:60]}' "
                          f"instead of a computed score or an '[auto]' placeholder.")
            elif O is not None and None not in (G, I_, K, M):
                expected = (G + I_ + K + M) / 4
                if not approx(O, expected):
                    bad("O", f"policy_score (O)={O} does not equal (G+I+K+M)/4 = ({G}+{I_}+{K}+{M})/4 = {expected:.3f}.")

        # P
        if P not in (0, 0.2, 0.4, 0.6, 0.8, 1.0, None):
            bad("P", f"instrument_type (P)={P} is not one of the allowed values {{0, 0.20, 0.40, 0.60, 0.80, 1.0}}.")
        elif P is None:
            bad("P", "instrument_type (P) is empty.")

        # Q
        if not is_empty(Q) and str(Q).strip() not in ALLOWED_Q:
            bad("Q", f"instrument_lifecycle_stage (Q)='{Q}' is not one of the allowed title-case values "
                      f"{sorted(ALLOWED_Q)} (Rule 4).")
        elif is_empty(Q):
            bad("Q", "instrument_lifecycle_stage (Q) is empty.")

        # S
        if S not in (0, 1, None):
            bad("S", f"instrument_in_force (S)={S} is not 0 or 1.")
        elif S is None:
            bad("S", "instrument_in_force (S) is empty.")

        # T
        if T not in (0, 0.25, 0.5, 0.75, 1.0, None):
            bad("T", f"instrument_implementation (T)={T} is not one of the allowed values {{0, 0.25, 0.5, 0.75, 1}}.")
        elif T is None:
            bad("T", "instrument_implementation (T) is empty.")

        # V formula check — blank or an explicit "[auto]"-style placeholder is compliant
        raw_V = get(row, "V")
        if not is_empty(raw_V):
            is_placeholder = bool(re.match(r"^\[?auto\]?$", str(raw_V).strip(), re.I))
            if V is None and not is_placeholder:
                bad("V", f"instrument_score (V) contains non-numeric, non-placeholder content '{str(raw_V)[:60]}' "
                          f"instead of a computed score or an '[auto]' placeholder.")
            elif V is not None and None not in (S, P, T):
                expected_v = S * (P + T) / 2
                if not approx(V, expected_v):
                    bad("V", f"instrument_score (V)={V} does not equal S×(P+T)/2 = {S}×({P}+{T})/2 = {expected_v:.3f}.")

        # --- Heuristic semantic checks (Rule 12: may/shall in-force test) ---
        text_blob = " ".join(str(x) for x in (R, U, W) if x)
        enabling_pat = re.compile(r"\b(may|is authoris?ed to|is empowered to|authoris?ed to|compet(e|ência)... ?aprovar|poderá|podrá|facultad(a|o) para)\b", re.I)
        mandatory_pat = re.compile(r"\b(shall|must|is required to|is prohibited|may only|deve[rm]?|obrigad[oa]|debe(rá)?|prohibi(do|ción))\b", re.I)
        has_enabling = bool(enabling_pat.search(text_blob))
        has_mandatory = bool(mandatory_pat.search(text_blob))
        already_operationalised = bool(re.search(r"(already|has been|were) (exercised|operationalis|operationaliz|implemented)|subordinate (regulation|ordinance|decision)", text_blob, re.I))

        if S == 1 and has_enabling and not has_mandatory and not already_operationalised:
            bad("S", "instrument_in_force (S)=1, but the description/implementation text only shows enabling-power "
                      "language ('may' / 'is authorised to') with no mandatory language and no evidence the power has "
                      "been operationalised by a subordinate instrument. Per Rule 12, re-check whether S should be 0.",
                sev="REVIEW")
        if S == 0 and has_mandatory and not has_enabling:
            bad("S", "instrument_in_force (S)=0, but the description/implementation text contains mandatory language "
                      "('shall' / 'must' / 'is prohibited') with no enabling-power language. Per Rule 12, re-check "
                      "whether S should be 1.", sev="REVIEW")

        # --- Heuristic semantic checks (Rule 11: coordination instruments) ---
        coord_pat = re.compile(
            r"(coordinat|multi-?level governance|levels? of government|national[\s\-–—]*(and|to|–|-)*\s*(regional|cantonal|municipal)|"
            r"delegat(e|ion|es)|assign(s|ed)? .*(responsibilit|function)|inter-?institutional coordination)", re.I)
        if coord_pat.search(text_blob) and P is not None:
            if P not in (0.4,):
                # only flag if not already the higher-type + comment-noted case per Rule 11
                noted = bool(W) and re.search(r"coordinat", str(W), re.I)
                if not noted:
                    bad("P", f"instrument_description/comments mention coordination/multi-level governance responsibilities, "
                              f"but instrument_type (P)={P} is not 0.40 and Col W does not note the coordination aspect "
                              f"alongside a higher-scoring instrument type. Per Rule 11, verify P should be 0.40, or "
                              f"add a note in W justifying the higher score.", sev="REVIEW")

    return issues


def main():
    file_to_refs = load_manifest()
    all_results = {}
    xlsx_files = sorted(f for f in os.listdir(RAW) if f.endswith(".xlsx"))
    for fname in xlsx_files:
        path = os.path.join(RAW, fname)
        try:
            rows = read_xlsx(path)
        except Exception as e:
            all_results[fname] = {"refs": file_to_refs.get(fname, []), "issues": [
                {"row": None, "col": None, "severity": "ERROR", "issue": f"Failed to read xlsx: {e}"}
            ]}
            continue
        issues = analyze_file(fname, rows, file_to_refs.get(fname, []))
        all_results[fname] = {"refs": file_to_refs.get(fname, []), "issues": issues}

    with open("/tmp/4p_audit/results.json", "w") as f:
        json.dump(all_results, f, indent=2)

    total_files = len(all_results)
    total_issues = sum(len(v["issues"]) for v in all_results.values())
    by_sev = {}
    for v in all_results.values():
        for iss in v["issues"]:
            by_sev[iss["severity"]] = by_sev.get(iss["severity"], 0) + 1
    print(f"Analyzed {total_files} unique xlsx files, {total_issues} findings.")
    print("By severity:", by_sev)

if __name__ == "__main__":
    main()
