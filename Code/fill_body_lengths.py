#!/usr/bin/env python3
"""
fill_body_lengths.py

Fills body_length_min_mm and body_length_max_mm in the workbook's feeding_groups
sheet from published measurements. Only EMPTY cells are filled: anything already
entered by a person is left exactly as it is, and the source of every filled value
is recorded beside it.

    conda activate video
    python fill_body_lengths.py --workbook "$ZK/data/cfbkey_workbook.xlsx" \\
        --stemberger "$ZK/data/stemberger_1979_measurements.csv"

    # once the Hebert supplement has been unzipped:
    python fill_body_lengths.py --workbook ... --stemberger ... \\
        --hebert "$ZK/data/hebert_2016_traits.csv"

SOURCES
  --stemberger  the transcribed tables from Stemberger 1979 (rotifers: Keratella,
                Notholca, Polyarthra). Values are in micrometres and are converted.
  --hebert      Hebert, Beisner & Maranger 2016, Ecology 97:1081, the freshwater
                crustacean trait compilation (copepods and cladocerans). Give it
                the CSV from the Supporting Information zip; the script looks for
                columns naming a taxon and a body length and reports what it found.

WHICH MEASUREMENT
  Rotifer lorica length (LL) EXCLUDES the posterior spine, which can be a third of
  the animal again. Pass --add-spines to use lorica length plus posterior spine as
  the maximum, which is closer to what a predator has to handle. Either way the
  measurement used is written into the source column, so the choice is visible.
"""

import argparse
import csv
import datetime
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

REF_STEM = ("Stemberger 1979, EPA-600/4-79-021, A Guide to Rotifers of the Laurentian "
            "Great Lakes")
REF_HEB = ("Hebert, Beisner & Maranger 2016, Ecology 97:1081, doi:10.1890/15-1275.1")


def make_backup(wbp: Path) -> Path:
    folder = wbp.parent / "workbook_backups"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = f"{datetime.datetime.now():%Y%m%d-%H%M%S}"
    dest = folder / f"{wbp.stem}_backup_{stamp}{wbp.suffix}"
    n = 2
    while dest.exists():
        dest = folder / f"{wbp.stem}_backup_{stamp}-{n}{wbp.suffix}"
        n += 1
    shutil.copy2(wbp, dest)
    return dest


def norm(name) -> str:
    """A comparable form of a taxon name: lower case, no authority, no sub-rank words."""
    s = str(name or "").strip().lower()
    s = re.sub(r"\(.*?\)", " ", s)
    s = re.sub(r"\b(f|var|forma|subsp|ssp|sp|spp|cf|aff)\b\.?", " ", s)
    s = re.sub(r"[^a-z ]", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def fnum(v):
    if v in (None, ""):
        return None
    try:
        return float(str(v).replace(",", "."))
    except ValueError:
        m = re.search(r"-?\d+(?:\.\d+)?", str(v))
        return float(m.group(0)) if m else None


def read_stemberger(path: Path, add_spines: bool):
    """taxon -> (min_mm, max_mm, note). Lorica length for loricate rotifers, body
    length for Polyarthra; the posterior spine added only if asked for."""
    by_taxon = defaultdict(dict)
    genus_vals = defaultdict(list)
    with path.open(encoding="utf-8-sig", newline="") as fh:
        for r in csv.DictReader(fh):
            key = (str(r.get("taxon", "")).strip(), str(r.get("form", "")).strip())
            by_taxon[key][str(r.get("measurement_code", "")).strip()] = r
    out = {}
    for (taxon, form), m in by_taxon.items():
        row = m.get("LL") or m.get("BL")
        if not row:
            continue
        lo, hi = fnum(row.get("range_min_um")), fnum(row.get("range_max_um"))
        if lo is None or hi is None:
            continue
        which = "lorica length" if "LL" in m else "body length"
        note = which
        if add_spines and "PS" in m:
            ps_hi = fnum(m["PS"].get("range_max_um"))
            if ps_hi:
                hi += ps_hi
                note = f"{which} plus the longest posterior spine ({ps_hi:g} um)"
        out[norm(f"{taxon} {form}")] = (lo / 1000, hi / 1000, note)
        out.setdefault(norm(taxon), (lo / 1000, hi / 1000, note + ", from the "
                                     + (form or "nominate") + " form"))
        genus_vals[norm(taxon.split()[0])].append((lo / 1000, hi / 1000, which))
    # a genus-level span, for a species the guide did not measure
    for g, vals in genus_vals.items():
        if g not in out and len(vals) >= 1:
            out[g] = (min(v[0] for v in vals), max(v[1] for v in vals),
                      f"{vals[0][2]} spanning {len(vals)} measured taxa in the genus")
    return out


def read_hebert(path: Path, include_marine: bool = False):
    """Hebert, Beisner & Maranger 2016. The file is semicolon-separated with decimal
    COMMAS, one row per individual measured, and Body.length is a mean in mm. Returns
    (lengths, trophic, shape) where lengths maps a name to (min_mm, max_mm, note)."""
    raw = path.read_text(encoding="utf-8-sig", errors="replace")
    sep = ";" if raw.count(";") > raw.count(",") else ","
    rows = list(csv.DictReader(raw.splitlines(), delimiter=sep))
    if not rows:
        return {}, {}, "the file has no rows"

    def dec(v):
        """'1,14' -> 1.14; 'NA' and blanks -> None."""
        t = str(v or "").strip()
        if not t or t.upper() == "NA":
            return None
        try:
            return float(t.replace(",", "."))
        except ValueError:
            return None

    vals, troph, kept, skipped_marine = defaultdict(list), defaultdict(list), 0, 0
    for r in rows:
        habitat = str(r.get("Habitat", "")).strip().lower()
        if habitat.startswith("marine") and not include_marine:
            skipped_marine += 1
            continue
        genus = str(r.get("Genus", "")).strip()
        species = re.sub(r"-?\(M\)", "", str(r.get("Species", "")).strip()).strip()
        if not genus:
            continue
        name = f"{genus} {species}".strip() if species and species.lower() not in ("sp.", "spp.", "na") else genus
        for col in ("Body.length", "Min.bl", "Max.bl"):
            v = dec(r.get(col))
            if v and v > 0:
                vals[norm(name)].append(v)
                vals[norm(genus)].append(v)
        tg = str(r.get("Trophic.group", "")).strip()
        if tg and tg.upper() != "NA":
            troph[norm(name)].append(tg)
        kept += 1

    lengths = {}
    for k, v in vals.items():
        if k and v:
            lengths[k] = (min(v), max(v),
                          f"body length in mm, {len(v)} measurement(s)")
    trophic = {k: max(set(v), key=v.count) for k, v in troph.items() if v}
    shape = (f"semicolon-separated, decimal commas, {kept} freshwater row(s) used"
             + (f", {skipped_marine} marine row(s) skipped" if skipped_marine else ""))
    return lengths, trophic, shape


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--stemberger", default="")
    ap.add_argument("--hebert", default="")
    ap.add_argument("--add-spines", action="store_true",
                    help="use lorica length plus posterior spine as the maximum for rotifers")
    ap.add_argument("--include-marine", action="store_true",
                    help="also use the marine rows of the Hebert file (freshwater only by default)")
    ap.add_argument("--overwrite", action="store_true",
                    help="also replace lengths that are already filled in (a backup is made)")
    a = ap.parse_args()
    wbp = Path(a.workbook).expanduser()
    if not wbp.exists():
        sys.exit(f"ERROR: workbook not found: {wbp}")

    stem = read_stemberger(Path(a.stemberger).expanduser(), a.add_spines) if a.stemberger else {}
    if a.stemberger:
        print(f"Stemberger 1979: {len(stem)} name(s) with a length"
              + (" (spines included)" if a.add_spines else " (lorica length, spines NOT included)"))
    heb, heb_troph, shape = ({}, {}, "")
    if a.hebert:
        heb, heb_troph, shape = read_hebert(Path(a.hebert).expanduser(), a.include_marine)
        print(f"Hebert 2016: {len(heb)} name(s) with a length  [{shape}]")
        if not heb:
            print("  nothing usable was read: tell me the column names and I will fix the reader")

    backup = make_backup(wbp)
    print(f"Backup saved: {backup.parent.name}/{backup.name}")
    wb = load_workbook(wbp)
    if "feeding_groups" not in wb.sheetnames:
        sys.exit("ERROR: no 'feeding_groups' sheet. Run draft_feeding_groups.py first.")
    ws = wb["feeding_groups"]
    heads = [c.value for c in ws[1]]

    for col in ("body_length_min_mm", "body_length_max_mm", "source"):
        if col not in heads:
            i = ws.max_column + 1
            c = ws.cell(row=1, column=i, value=col)
            c.font = Font(name="Arial", bold=True, size=10, color="000000")
            c.fill = PatternFill("solid", fgColor="FFD966")
            ws.column_dimensions[get_column_letter(i)].width = 40
            heads = [x.value for x in ws[1]]
    ci = {h: i for i, h in enumerate(heads)}

    filled = kept = unmatched = 0
    by_source = defaultdict(int)
    missing = []
    for row in ws.iter_rows(min_row=2):
        name = row[ci.get("display_name", 1)].value or row[ci["species_id"]].value
        key = norm(name)
        genus = str(row[ci["genus"]].value or "").strip() if "genus" in ci else ""
        hit = stem.get(key) or heb.get(key)
        src = REF_STEM if stem.get(key) else (REF_HEB if heb.get(key) else "")
        if not hit and genus:                       # fall back to the genus
            gk = norm(genus)
            hit = stem.get(gk) or heb.get(gk)
            if hit:
                src = (REF_STEM if stem.get(gk) else REF_HEB) + " (genus-level value)"
        has_value = row[ci["body_length_max_mm"]].value not in (None, "")
        if has_value and not a.overwrite:
            kept += 1
            continue
        if not hit:
            unmatched += 1
            if name:
                missing.append(str(name))
            continue
        lo, hi, note = hit
        row[ci["body_length_min_mm"]].value = round(lo, 4)
        row[ci["body_length_max_mm"]].value = round(hi, 4)
        row[ci["source"]].value = f"{src}; {note}"
        filled += 1
        by_source[src.split(",")[0]] += 1

    disagree = []
    if heb_troph:
        for row in ws.iter_rows(min_row=2):
            nm = row[ci.get("display_name", 1)].value
            theirs = heb_troph.get(norm(nm))
            if not theirs:
                continue
            ours = str(row[ci["trophic_group_draft"]].value or "").strip().lower() \
                if "trophic_group_draft" in ci else ""
            t = theirs.lower()
            same = (("carn" in t and ours == "carnivore") or ("herb" in t and ours == "herbivore")
                    or ("omni" in t and ours == "omnivore"))
            if ours and not same:
                disagree.append(f"{nm}: we say {ours}, Hebert says {theirs}")

    wb.save(wbp)
    print(f"\nbody lengths filled: {filled}")
    for s_, n in by_source.items():
        print(f"    {n:4}  {s_}")
    print(f"already had a value, left alone: {kept}")
    print(f"still without a length: {unmatched}")
    if missing:
        print("   " + ", ".join(sorted(missing)[:14]) + (" ..." if len(missing) > 14 else ""))
    if disagree:
        print(f"\nfeeding groups where Hebert disagrees with our draft: {len(disagree)}")
        for d_ in disagree[:12]:
            print(f"    {d_}")
        if len(disagree) > 12:
            print(f"    ... and {len(disagree) - 12} more")
        print("    Hebert's trophic groups are published; ours are mostly taxonomy defaults,")
        print("    so these are worth reviewing in the feeding_groups sheet.")
    print("\nRerun build_food_web.py to see the difference.")


if __name__ == "__main__":
    main()
