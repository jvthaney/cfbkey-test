#!/usr/bin/env python3
"""
add_nla_to_traits.py

Folds the NLA-derived cell biovolumes into phyto_traits.xlsx, after the comparison
script has shown that they are volumes per CELL.

    conda activate video
    python add_nla_to_traits.py \\
        --traits "$ZK/data/phyto_traits.xlsx" \\
        --nla "$ZK/data/nla_biovolumes.csv"

WHAT IT DOES
  * For a genus the trait table does NOT have, it adds one row: a cell-level entry
    with the NLA median biovolume, marked as coming from the NLA.
  * For a genus the table DOES have, it changes nothing. Rimet stays authoritative.
    The NLA figure is written into two new columns beside it, so both numbers are
    visible and can be compared later.
  * Nothing is overwritten, and every added row carries its source and the number of
    NLA records behind it.

THE ONE APPROXIMATION, AND IT MATTERS
  The NLA gives a volume but no dimensions, while the feeding rules need the largest
  dimension of the particle. This script derives a SPHERE-EQUIVALENT diameter:

      diameter = (6 * volume / pi) ** (1/3)

  For a round cell that is about right. For anything elongated it is a serious
  UNDERESTIMATE: a needle-shaped Asterionella cell of 600 cubic micrometres is some
  60 um long but only 10 um as a sphere. Since the grazing rules compare the largest
  dimension against a window, an underestimate makes a taxon look MORE edible than it
  is. Every such row says so in its unit_note, and the column
  'max_dimension_is_sphere_equivalent' marks them, so they can be corrected as real
  measurements arrive.
"""

import argparse
import csv
import math
import sys
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

REF_NLA = ("U.S. EPA National Lakes Assessment 2017 phytoplankton counts; cell biovolume "
           "derived as BIOVOLUME / DENSITY. Validated against Rimet & Druart 2018: "
           "colonial genera agree with Rimet's per-cell figure, with NLA about 1.9-fold "
           "lower overall; that offset is NOT corrected for.")


def fnum(v):
    if v in (None, ""):
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def sphere_diameter(volume_um3):
    """The diameter of a sphere of this volume. A LOWER BOUND on the real largest
    dimension for anything that is not round."""
    if not volume_um3 or volume_um3 <= 0:
        return None
    return (6 * volume_um3 / math.pi) ** (1 / 3)


def size_fraction(max_um):
    if max_um is None:
        return ""
    if max_um < 2:
        return "<2 um (pico)"
    if max_um < 20:
        return "2-20 um (nano)"
    if max_um < 50:
        return "20-50 um"
    if max_um < 100:
        return "50-100 um"
    return ">100 um"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--traits", required=True, help="phyto_traits.xlsx")
    ap.add_argument("--nla", required=True, help="nla_biovolumes.csv from compare_nla_rimet.py")
    ap.add_argument("--min-records", type=int, default=5,
                    help="skip an NLA genus with fewer records than this (default 5)")
    a = ap.parse_args()

    tp = Path(a.traits).expanduser()
    if not tp.exists():
        sys.exit(f"ERROR: trait table not found: {tp}")
    nla_rows = []
    with Path(a.nla).expanduser().open(encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            nla_rows.append(r)
    if not nla_rows:
        sys.exit("ERROR: the NLA file has no rows. Run compare_nla_rimet.py with --out first.")
    print(f"NLA genera read: {len(nla_rows)}")

    wb = load_workbook(tp)
    ws = wb["phyto_traits"]
    heads = [c.value for c in ws[1]]

    # two new columns, so the NLA figure sits beside Rimet's rather than replacing it
    for col in ("nla_cell_biovolume_um3", "nla_records", "max_dimension_is_sphere_equivalent"):
        if col not in heads:
            i = ws.max_column + 1
            c = ws.cell(row=1, column=i, value=col)
            c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            ws.column_dimensions[get_column_letter(i)].width = 22
            heads = [x.value for x in ws[1]]
    ci = {h: i + 1 for i, h in enumerate(heads)}

    have_genus = {}
    for row in ws.iter_rows(min_row=2):
        g = str(row[ci["genus"] - 1].value or "").strip().title()
        if g:
            have_genus.setdefault(g, []).append(row)
    print(f"trait table genera already present: {len(have_genus)}")

    annotated = added = skipped = not_cell = 0
    for r in nla_rows:
        genus = str(r.get("genus", "")).strip().title()
        bv = fnum(r.get("nla_median_biovolume_um3"))
        n = int(fnum(r.get("nla_records")) or 0)
        if not genus or not bv or n < a.min_records:
            skipped += 1
            continue

        if genus in have_genus:
            # Rimet stays authoritative: record the NLA figure beside it, change nothing
            # else. Only CELL rows get it: a per-cell number written next to a colony row
            # would invite exactly the confusion this project keeps trying to avoid.
            touched = 0
            for row in have_genus[genus]:
                unit = str(row[ci["counting_unit"] - 1].value or "").strip().lower()
                if unit != "cell":
                    continue
                row[ci["nla_cell_biovolume_um3"] - 1].value = round(bv, 3)
                row[ci["nla_records"] - 1].value = n
                touched += 1
            if touched:
                annotated += 1
            else:
                not_cell += 1
            continue

        d = sphere_diameter(bv)
        ws.append([None] * len(heads))
        new = ws[ws.max_row]

        def put(col, val):
            if col in ci:
                new[ci[col] - 1].value = val
        put("row_type", "nla_genus")
        put("role", "phytoplankton")
        put("role_note", "added from the NLA: not in Rimet & Druart")
        put("taxon", f"{genus} (NLA cell median)")
        put("genus", genus)
        put("counting_unit", "cell")
        put("unit_note", f"median of {n} NLA records. The largest dimension is a "
                         f"SPHERE-EQUIVALENT diameter derived from the volume, so for any "
                         f"elongated cell it is an underestimate and makes the taxon look "
                         f"more edible than it is.")
        put("max_dimension_um", round(d, 2) if d else None)
        put("biovolume_um3", round(bv, 3))
        put("n_species_in_genus", n)
        put("size_fraction", size_fraction(d))
        put("edible_lt_50um", "yes" if d and d < 50 else ("no" if d else ""))
        put("nla_cell_biovolume_um3", round(bv, 3))
        put("nla_records", n)
        put("max_dimension_is_sphere_equivalent", "yes")
        put("measurement_note", "derived: BIOVOLUME / DENSITY")
        put("reference", REF_NLA)
        added += 1

    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.font is None or c.font.name != "Arial":
                c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top",
                                    wrap_text=heads[c.column - 1] in ("unit_note", "reference"))
    wb.save(tp)

    print(f"\ngenera added as NLA cell rows: {added}")
    print(f"genera already in the table, NLA figure recorded beside Rimet: {annotated}")
    print(f"skipped (too few records or no biovolume): {skipped}")
    if not_cell:
        print(f"in the table but with no cell row to annotate: {not_cell}")
    print(f"\nWrote {tp}")
    print("\nTwo things to keep in mind:")
    print("  * Every added row is a CELL. A colonial taxon counted as colonies still needs")
    print("    a colony row, which the NLA cannot give. Use the extras file for those.")
    print("  * The largest dimension on added rows is a sphere-equivalent diameter. Filter")
    print("    on max_dimension_is_sphere_equivalent to find the rows that need a real")
    print("    measurement, starting with anything needle-shaped or filamentous.")
    print("\nRerun build_dashboard.py to pick these up.")


if __name__ == "__main__":
    main()
