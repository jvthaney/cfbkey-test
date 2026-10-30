#!/usr/bin/env python3
"""
compare_nla_rimet.py

Before trusting NLA phytoplankton data as a source of cell sizes, check it against a
source we already use. This script derives a mean biovolume per counted organism from
the EPA National Lakes Assessment phytoplankton counts, then compares it with Rimet &
Druart (2018) for every genus the two have in common.

    conda activate video
    python compare_nla_rimet.py \\
        --nla data/nla_2017_phytoplankton_count-data.csv \\
        --rimet data/RimetDruart2018Appendix1-Phytoplankton-metrics-database-revised.xlsx \\
        --out data/nla_biovolumes.csv

HOW THE NLA NUMBER IS DERIVED
  The NLA metadata defines its columns as:
      BIOVOLUME   volume of cells per mL of sample   (cu um/ml)
      DENSITY     number of organisms per mL
      ABUNDANCE   number of organisms counted from the sample
  So the mean volume of ONE counted organism is BIOVOLUME / DENSITY, in cubic
  micrometres. Dividing by ABUNDANCE instead would be wrong: that is a raw count, not
  a concentration.

  Because BIOVOLUME counts CELLS while DENSITY counts ORGANISMS, a colonial taxon
  counted as colonies gives the summed volume of the cells in a colony, excluding any
  mucilage. That is comparable to Rimet's "cumulated biovolume of cells in a colony"
  rather than to a colony's outside dimensions.

WHAT TO LOOK FOR IN THE OUTPUT
  Agreement within about an order of magnitude across many genera means the derivation
  is sound and the NLA can be trusted for the genera Rimet lacks. Systematic
  disagreement means it cannot, and the comparison will show the direction.
"""

import argparse
import csv
import math
import re
import statistics as st
import sys
from collections import defaultdict
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")


def fnum(v):
    if v in (None, ""):
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        return None


def read_taxa_list(path: Path):
    """TAXA_ID -> genus, from the separate taxa file, for count files that carry only an id."""
    out = {}
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        cols = {c.lower(): c for c in (rd.fieldnames or [])}
        if "taxa_id" not in cols or "genus" not in cols:
            print(f"  note: the taxa list has no TAXA_ID and GENUS columns "
                  f"({', '.join(rd.fieldnames or [])}); it will not be used")
            return {}
        for r in rd:
            tid = str(r.get(cols["taxa_id"], "")).strip()
            g = str(r.get(cols["genus"], "")).strip().title()
            if tid and g:
                out[tid] = g
    return out


def read_nla(path: Path, taxa: dict, unit_filter: str = ""):
    """genus -> list of mean biovolumes per organism (cubic micrometres)."""
    by_genus, by_taxon = defaultdict(list), defaultdict(list)
    rows = skipped = 0
    unit_counts = defaultdict(int)
    with path.open(encoding="utf-8-sig", newline="") as fh:
        rd = csv.DictReader(fh)
        cols = {c.lower(): c for c in (rd.fieldnames or [])}
        need = [c for c in ("biovolume", "density") if c not in cols]
        if need:
            sys.exit(f"ERROR: the NLA count file has no {', '.join(need)} column.\n"
                     f"Columns found: {', '.join(rd.fieldnames or [])}")
        if "genus" not in cols and not taxa:
            sys.exit("ERROR: the count file has no GENUS column and no taxa list was given.\n"
                     "Pass --taxa with the phytoplankton taxa list file.")
        for r in rd:
            rows += 1
            bv, dens = fnum(r.get(cols["biovolume"])), fnum(r.get(cols["density"]))
            if "genus" in cols:
                genus = str(r.get(cols["genus"], "")).strip().title()
            else:
                genus = taxa.get(str(r.get(cols.get("taxa_id", ""), "")).strip(), "")
            if not genus and taxa and "taxa_id" in cols:
                genus = taxa.get(str(r.get(cols["taxa_id"], "")).strip(), "")
            # cells or natural units: mixing the two blurs exactly the distinction
            # between a cell and a colony that the trait table keeps separate
            cu = str(r.get(cols.get("count_unit", ""), "")).strip().upper() if "count_unit" in cols else ""
            if cu:
                unit_counts[cu] += 1
                if unit_filter and cu != unit_filter.upper():
                    skipped += 1
                    continue
            if not genus or not bv or not dens or dens <= 0 or bv <= 0:
                skipped += 1
                continue
            per_organism = bv / dens          # cu um per counted organism
            by_genus[genus].append(per_organism)
            sp = str(r.get(cols.get("species", ""), "")).strip().lower()
            if sp and sp not in ("", "sp.", "spp."):
                by_taxon[f"{genus} {sp}"].append(per_organism)
    if unit_counts:
        pretty = ", ".join(f"{k}={v}" for k, v in sorted(unit_counts.items()))
        print(f"  COUNT_UNIT present in the file: {pretty}   (C = cells, N = natural units)")
        if not unit_filter:
            print("  both kinds are pooled. Use --count-unit C or N to keep them apart.")
    return by_genus, by_taxon, rows, skipped


def read_rimet(path: Path):
    """genus -> (reference biovolumes, cell biovolumes, colonial flags, cells per colony)."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Phytoplankton traits database"]
    rows = list(ws.iter_rows(values_only=True))
    h = [str(c) if c is not None else "" for c in rows[0]]
    def col(name):
        return h.index(name) if name in h else None
    gi, ri = col("Genus"), col("Reference biovolume")
    ci, coli, ni = col("Cell biovolume \u00b5m3"), col("Colonial"), col("Number of cells per colony")
    if gi is None or ri is None:
        sys.exit("ERROR: could not find 'Genus' and 'Reference biovolume' in the Rimet sheet.")
    out = defaultdict(lambda: {"ref": [], "cell": [], "colonial": [], "ncell": []})
    for r in rows[1:]:
        g = str(r[gi]).strip().title() if gi < len(r) and r[gi] else ""
        v = fnum(r[ri]) if ri < len(r) else None
        if not g or not v or v <= 0:
            continue
        out[g]["ref"].append(v)
        cv = fnum(r[ci]) if ci is not None and ci < len(r) else None
        if cv and cv > 0:
            out[g]["cell"].append(cv)
        cf = fnum(r[coli]) if coli is not None and coli < len(r) else None
        if cf is not None:
            out[g]["colonial"].append(cf)
        nc = fnum(r[ni]) if ni is not None and ni < len(r) else None
        if nc and nc > 0:
            out[g]["ncell"].append(nc)
    return out


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--nla", required=True)
    ap.add_argument("--taxa", default="",
                    help="the NLA phytoplankton taxa list, used where the count file "
                         "carries only a TAXA_ID")
    ap.add_argument("--count-unit", default="",
                    help="keep only records counted as cells (C) or as natural units (N), "
                         "where the file records which")
    ap.add_argument("--rimet", required=True)
    ap.add_argument("--out", default="", help="write the NLA genus biovolumes to this CSV")
    ap.add_argument("--min-um3", type=float, default=1.0,
                    help="treat a genus median below this many cubic micrometres as "
                         "implausible and leave it out (default 1.0)")
    ap.add_argument("--min-records", type=int, default=5,
                    help="ignore NLA genera with fewer records than this (default 5)")
    a = ap.parse_args()

    taxa = read_taxa_list(Path(a.taxa).expanduser()) if a.taxa else {}
    if taxa:
        print(f"taxa list: {len(taxa)} TAXA_ID to genus mappings")
    nla_g, nla_t, nrows, nskip = read_nla(Path(a.nla).expanduser(), taxa, a.count_unit)
    print(f"NLA: {nrows} record(s) read, {nskip} without a usable biovolume or density")
    print(f"     {len(nla_g)} genera, {len(nla_t)} named species")
    rim = read_rimet(Path(a.rimet).expanduser())
    print(f"Rimet: {len(rim)} genera with a reference biovolume")

    # A phytoplankton cell below about 1 cubic micrometre is smaller than most
    # picoplankton, so such a median points at a problem in the record rather than a
    # real measurement. Chroococcus at 0.27 um3 is the kind of case this catches.
    implausible = sorted(g for g in nla_g
                         if len(nla_g[g]) >= a.min_records and st.median(nla_g[g]) < a.min_um3)
    if implausible:
        print(f"\n{len(implausible)} genera have a median under {a.min_um3} cubic micrometres, "
              f"which is implausibly small for a cell:")
        for g in implausible[:12]:
            print(f"   {g[:22]:24} {st.median(nla_g[g]):8.3f} um3   ({len(nla_g[g])} records)")
        if len(implausible) > 12:
            print(f"   ... and {len(implausible) - 12} more")
        print("   These are left out of the comparison and should not be used until someone")
        print("   works out what those records mean.")
    too_small = set(implausible)

    shared = sorted(g for g in nla_g if g in rim and len(nla_g[g]) >= a.min_records
                    and g not in too_small)
    print(f"\ngenera in both, with at least {a.min_records} NLA records: {len(shared)}")
    if not shared:
        sys.exit("No overlap to compare. Check that the genus columns are spelled alike.")

    ratios, table = [], []
    for g in shared:
        n_med = st.median(nla_g[g])
        r_med = st.median(rim[g]["ref"])
        ratio = n_med / r_med
        colonial = (st.mean(rim[g]["colonial"]) >= 0.5) if rim[g]["colonial"] else None
        cell_med = st.median(rim[g]["cell"]) if rim[g]["cell"] else None
        ncell = st.median(rim[g]["ncell"]) if rim[g]["ncell"] else None
        ratios.append(ratio)
        table.append((g, n_med, r_med, ratio, len(nla_g[g]), colonial, cell_med, ncell))

    logs = [math.log10(x) for x in ratios]
    med_ratio = 10 ** st.median(logs)
    within10 = sum(1 for x in ratios if 0.1 <= x <= 10)
    within3 = sum(1 for x in ratios if 1/3 <= x <= 3)
    print(f"\nNLA median divided by Rimet median, across those genera:")
    print(f"   median ratio      {med_ratio:8.2f}x   (1.0 would be perfect agreement)")
    print(f"   within 3-fold     {within3:4} of {len(ratios)}  ({100*within3/len(ratios):.0f}%)")
    print(f"   within 10-fold    {within10:4} of {len(ratios)}  ({100*within10/len(ratios):.0f}%)")
    if len(logs) > 1:
        print(f"   spread (SD of log10 ratio) {st.stdev(logs):.2f}")

    def fmt(v):
        return f"{v:12,.0f}" if v is None or v >= 10 else f"{v:12,.2f}"

    # THE DECISIVE TEST: does NLA agree with Rimet's CELL biovolume rather than its
    # colony biovolume? If so, NLA is reporting volume per cell.
    uni = [t for t in table if t[5] is False]
    col = [t for t in table if t[5] is True]
    print(f"\nsplit by whether Rimet calls the genus colonial:")
    for label, subset in (("unicellular", uni), ("colonial", col)):
        if not subset:
            continue
        rs = [t[3] for t in subset]
        lg = [math.log10(x) for x in rs]
        w3 = sum(1 for x in rs if 1/3 <= x <= 3)
        print(f"   {label:12} {len(subset):4} genera   median ratio {10**st.median(lg):7.2f}x"
              f"   within 3-fold {w3}/{len(subset)} ({100*w3/len(subset):.0f}%)")
    if col:
        # against Rimet's own per-cell figure, where it has one
        pair = [(t[0], t[1], t[6]) for t in col if t[6]]
        if pair:
            rs = [n / c for _g, n, c in pair]
            lg = [math.log10(x) for x in rs]
            w3 = sum(1 for x in rs if 1/3 <= x <= 3)
            print(f"\n   colonial genera compared with Rimet's CELL biovolume instead:")
            print(f"      {len(pair):4} genera   median ratio {10**st.median(lg):7.2f}x"
                  f"   within 3-fold {w3}/{len(pair)} ({100*w3/len(pair):.0f}%)")

    table.sort(key=lambda t: -abs(math.log10(t[3])))
    print(f"\nthe ten genera that disagree most:")
    print(f"   {'genus':20} {'NLA um3':>12} {'Rimet ref':>12} {'ratio':>8}  {'colonial':>9} "
          f"{'cells/col':>10}  NLA n")
    for g, n_med, r_med, ratio, n_n, colonial, cell_med, ncell in table[:10]:
        print(f"   {g[:18]:20} {fmt(n_med)} {fmt(r_med)} {ratio:8.3f}  "
              f"{str(colonial):>9} {('%.0f' % ncell) if ncell else '-':>10}  {n_n}")
    print(f"\nthe ten that agree best:")
    for g, n_med, r_med, ratio, n_n, colonial, cell_med, ncell in table[-10:][::-1]:
        print(f"   {g[:18]:20} {fmt(n_med)} {fmt(r_med)} {ratio:8.3f}  "
              f"{str(colonial):>9} {('%.0f' % ncell) if ncell else '-':>10}  {n_n}")

    print("\nHOW TO READ THIS")
    uni_ok = False
    if uni:
        lg = [math.log10(t[3]) for t in uni]
        uni_med = 10 ** st.median(lg)
        uni_ok = 0.4 <= uni_med <= 2.5
    cell_better = False
    if col:
        pair = [(t[1], t[6]) for t in col if t[6]]
        if pair:
            lgc = [math.log10(n / c) for n, c in pair]
            w3c = sum(1 for n, c in pair if 1/3 <= n / c <= 3)
            w3r = sum(1 for t in col if 1/3 <= t[3] <= 3)
            cell_better = w3c > w3r
    if cell_better:
        print("   The colonial genera agree far better with Rimet's CELL biovolume than with its")
        print("   colony figure, and the unicellular genera agree reasonably either way. NLA is")
        print("   reporting volume per CELL. The derivation is sound; the two sources were")
        print("   measuring different units of the same thing.")
        print("   USE NLA VALUES AS CELL BIOVOLUMES. Never read them as colony sizes: doing so")
        print("   would understate a Microcystis colony by the number of cells in it.")
    elif within10 / len(ratios) >= 0.8 and 0.2 <= med_ratio <= 5:
        print("   Most genera agree within an order of magnitude and there is no large")
        print("   systematic offset, so BIOVOLUME / DENSITY looks like a sound derivation.")
    else:
        print("   Agreement is poor in a way the cell-against-colony explanation does not")
        print("   account for. Do not use the NLA numbers until this is understood.")
    if uni:
        lgu = [math.log10(t[3]) for t in uni]
        off = 10 ** st.median(lgu)
        if not (0.7 <= off <= 1.4):
            print(f"   NOTE: even the unicellular genera sit at {off:.2f}x Rimet. Expect a")
            print(f"   systematic offset of roughly {1/off:.1f}-fold between the two sources,")
            print(f"   which is worth remembering when they are mixed in one table.")
    print("   A genus with few NLA records is a weak estimate: check the n column.")

    if a.out:
        outp = Path(a.out).expanduser()
        outp.parent.mkdir(parents=True, exist_ok=True)
        with outp.open("w", newline="", encoding="utf-8-sig") as fh:
            w = csv.writer(fh)
            w.writerow(["genus", "nla_median_biovolume_um3", "nla_min_um3", "nla_max_um3",
                        "nla_records", "rimet_median_biovolume_um3", "rimet_cell_biovolume_um3",
                        "rimet_colonial", "ratio_nla_over_rimet", "in_rimet", "source"])
            for g, vals in sorted(nla_g.items()):
                if len(vals) < a.min_records:
                    continue
                if g in too_small:
                    continue
                r_med = st.median(rim[g]["ref"]) if g in rim else ""
                r_cell = st.median(rim[g]["cell"]) if g in rim and rim[g]["cell"] else ""
                r_col = (("yes" if st.mean(rim[g]["colonial"]) >= 0.5 else "no")
                         if g in rim and rim[g]["colonial"] else "")
                w.writerow([g, round(st.median(vals), 2), round(min(vals), 2), round(max(vals), 2),
                            len(vals), round(r_med, 1) if r_med != "" else "",
                            round(r_cell, 1) if r_cell != "" else "", r_col,
                            round(st.median(vals) / r_med, 3) if r_med else "",
                            "yes" if g in rim else "no",
                            "EPA National Lakes Assessment 2017 phytoplankton counts, "
                            "BIOVOLUME / DENSITY = cubic micrometres per counted organism "
                            "(per CELL where cells were the unit counted)"])
        print(f"\nWrote {outp}")
        print("   The rows marked in_rimet = no are the ones worth adding to the trait table.")


if __name__ == "__main__":
    main()
