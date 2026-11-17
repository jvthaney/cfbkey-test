#!/usr/bin/env python3
"""
export_feeding_traits.py

Exports the workbook's feeding_groups sheet to a CSV for the food web tools.

This is the ONE file that crosses between the two projects. The key owns the
zooplankton feeding traits, because Jim and Amanda edit them in the workbook; the
food web tools need them but should not have to reach into another folder to find
them. So the key exports, and the counter repository holds the exported copy.

    conda activate video
    python export_feeding_traits.py \\
        --workbook "$ZK/data/cfbkey_workbook.xlsx" \\
        --out "$HOME/Documents/GitHub/cyanocounter/data/zooplankton_feeding.csv"

Then commit the CSV in the counter repository and rebuild the dashboard there.

WHY A CSV RATHER THAN THE WORKBOOK
  The workbook carries draft notes, a to-do list and the whole identification key,
  none of which the dashboard needs or should publish. A CSV of one sheet is small,
  readable, and shows a plain difference in version control when a trait changes.
"""

import argparse
import csv
import datetime
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

# The columns the food web build actually reads. Anything else in the sheet is the
# key's own working material and is left behind.
WANTED = ["species_id", "display_name", "phylum", "class", "order", "suborder", "family",
          "genus", "feeding_mode", "trophic_group", "feeding_mode_draft",
          "trophic_group_draft", "body_length_min_mm", "body_length_max_mm",
          "body_length_mm_draft", "food_size_um_draft", "prey_size_min_um",
          "prey_size_max_um", "life_stage", "draft_basis", "confidence", "source"]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--all-columns", action="store_true",
                    help="export every column, not just the ones the food web build reads")
    a = ap.parse_args()

    wbp = Path(a.workbook).expanduser()
    if not wbp.exists():
        sys.exit(f"ERROR: workbook not found: {wbp}")
    wb = load_workbook(wbp, read_only=True, data_only=True)
    if "feeding_groups" not in wb.sheetnames:
        sys.exit("ERROR: no 'feeding_groups' sheet. Run draft_feeding_groups.py first.")
    ws = wb["feeding_groups"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        sys.exit("ERROR: the feeding_groups sheet is empty.")
    heads = [str(c).strip() if c is not None else "" for c in rows[0]]
    lower = {h.lower(): h for h in heads if h}

    if a.all_columns:
        cols = [h for h in heads if h]
    else:
        cols = [lower[w] for w in WANTED if w in lower]
        missing = [w for w in WANTED if w not in lower]
        if missing:
            print(f"  note: not in the sheet, so not exported: {', '.join(missing)}")

    outp = Path(a.out).expanduser()
    outp.parent.mkdir(parents=True, exist_ok=True)
    n = with_len = 0
    with outp.open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.writer(fh)
        w.writerow(cols)
        for r in rows[1:]:
            d = dict(zip(heads, r))
            if not d.get(lower.get("species_id", "species_id")):
                continue
            w.writerow(["" if d.get(c) is None else d.get(c) for c in cols])
            n += 1
            if d.get(lower.get("body_length_max_mm", "")) or d.get(lower.get("body_length_mm_draft", "")):
                with_len += 1

    print(f"Wrote {outp}")
    print(f"  {n} taxa, {len(cols)} columns, {with_len} with a body length")
    print(f"  exported from {wbp.name} on {datetime.date.today():%d %B %Y}")
    print("\nNext, in the counter repository:")
    print("   git add data/zooplankton_feeding.csv")
    print('   git commit -m "Update the zooplankton feeding traits from the key"')
    print("   bash build/rebuild_dashboard.sh")


if __name__ == "__main__":
    main()
