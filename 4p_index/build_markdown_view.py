#!/usr/bin/env python3
"""Generates a human-readable Markdown view of the 4P Index CSV, grouped by policy."""
import csv

CSV_PATH = "4p_index_nepal_plastic_policies.csv"
MD_PATH = "4p_index_nepal_plastic_policies.md"


def main():
    with open(CSV_PATH, newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))

    lines = [
        "# 4P Index -- Nepal Plastic Pollution Policies (coded table, human-readable view)",
        "",
        "See 4p_index_nepal_plastic_policies.csv for the machine-readable A-W column table. "
        "This file groups the same data by policy for easier reading.",
        "",
    ]

    detail_lines = []
    current_policy = None
    instrument_idx = 0

    for row in rows:
        if row["policy_name"] != current_policy:
            if current_policy is not None:
                lines.append("")
            current_policy = row["policy_name"]
            instrument_idx = 0
            lines.append(f"## {current_policy}")
            lines.append("")
            lines.append(f"- **Year (C):** {row['policy_year']}")
            lines.append(f"- **URL (B):** {row['policy_url']}")
            lines.append(f"- **Objective (D):** {row['policy_objective']}")
            lines.append(f"- **Target (E/F):** {row['policy_target']} -- {row['policy_target_text'] or '(none)'}")
            lines.append(f"- **Type (G/H):** {row['policy_type']} -- {row['policy_type_justification']}")
            lines.append(f"- **Integration (I/J):** {row['policy_integration']} -- sectors: {row['policy_sectors_list']}")
            lines.append(f"- **Circularity (K/L):** {row['policy_circularity']} -- phases: {row['policy_lifecycle_phases_list']}")
            lines.append(f"- **Budget (M/N):** {row['policy_budget']} -- {row['policy_budget_text'] or '(none)'}")
            lines.append(f"- **Policy score (O, auto):** {row['policy_score']}")
            lines.append("")
            lines.append("| # | Instrument type (P) | Lifecycle stage (Q) | In force (S) | Implementation (T) | Score (V, auto) |")
            lines.append("|---|---|---|---|---|---|")

            detail_lines.append(f"### {current_policy} -- instrument details (R, U, W)")
            detail_lines.append("")

        instrument_idx += 1
        lines.append(
            f"| {instrument_idx} | {row['instrument_type']} | {row['instrument_lifecycle_stage']} | "
            f"{row['instrument_in_force']} | {row['instrument_implementation']} | {row['instrument_score']} |"
        )

        detail_lines.append(f"**Instrument {instrument_idx} -- description (R):** {row['instrument_description']}")
        detail_lines.append("")
        detail_lines.append(f"*Implementation evidence (U):* {row['instrument_implementation_text']}")
        detail_lines.append("")
        if row["comments"]:
            detail_lines.append(f"*Comments (W):* {row['comments']}")
            detail_lines.append("")

    lines.append("")

    with open(MD_PATH, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))
        f.write("\n---\n\n")
        f.write("\n".join(detail_lines))
        f.write("\n")

    print(f"Wrote {MD_PATH}")


if __name__ == "__main__":
    main()
