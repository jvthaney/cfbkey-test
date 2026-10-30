#!/usr/bin/env python3
"""
build_food_web.py

Turns the two trait tables into a list of who eats whom, by applying published
feeding rules rather than guesswork. Every link records the rule that produced it,
so a reviewer can disagree with a link by disagreeing with a rule.

    conda activate video
    python build_food_web.py \\
        --zooplankton "$ZK/data/cfbkey_workbook.xlsx" \\
        --phytoplankton "$ZK/data/phyto_traits.xlsx" \\
        --out "$ZK/data/food_web.xlsx"

Optional:
    --counts counts.csv     only taxa in this file (a CyanoCounter export, or any
                            CSV with a 'taxon' column) are included
    --rules rules.csv       override the rule parameters; written for you if absent

THE RULES

1. GRAZING by filter and suspension feeders on phytoplankton.
   Upper limit on edible particle size, from Burns (1968):
       largest particle (um) = 22 x body length (mm) + 4.87
   Lower limit from the filter mesh: the gap between the setulae. Brendelberger
   (1991) measured 0.2 um in Diaphanosoma to 4.7 um in Sida. Default 1.0 um.
   A link is made when the food's largest dimension falls between the two.

2. MECHANICAL INTERFERENCE. Long filaments and large colonies foul the filtering
   combs rather than being eaten, so they are excluded above a length threshold
   even when the cells themselves are small. This is why a Microcystis colony is
   not food although a Microcystis cell is.

3. RAPTORIAL FEEDING on other animals. A link is made when the prey's body length
   falls within a window set as a fraction of the predator's length. Raptorial
   feeders also take larger algal cells than filterers of the same size, since
   they handle particles individually.

4. SCRAPERS feed on periphyton and detritus, which are not plankton counts, so
   they get a link to a 'detritus/periphyton' node instead.

5. PARASITES are linked to their hosts, which are usually fish and so outside
   anything either app counts. Those links are reported, not invented.

WHAT IT CANNOT DO WITHOUT MORE DATA
   Rules 1 and 3 need the animal's BODY LENGTH. Where the zooplankton table has
   none, no link can be made and the taxon is listed in the report. Filling those
   in, from Hebert et al. 2016 or from the key's own measurements, is the single
   change that most improves the web.
"""

import argparse
import csv
import re
import sys
from collections import defaultdict
from pathlib import Path

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

REF_BURNS = ("Burns, C.W. 1968. The relationship between body size of filter-feeding Cladocera "
             "and the maximum size of particle ingested. Limnol. Oceanogr. 13:675-678 "
             "(largest particle um = 22 x body length mm + 4.87)")
REF_MESH = ("Brendelberger, H. The filtration apparatus of Cladocera: filter mesh-sizes and "
            "their implications on food selectivity. Setule spacing measured in 11 species, "
            "0.2 um (Diaphanosoma brachyurum) to 4.7 um (Sida crystallina), grouped as fine, "
            "medium and coarse mesh. The per-class numbers used here are interpolated from that "
            "range and are not published values")

DEFAULT_RULES = """# Rule parameters for the food web. Change a value and rebuild.
# Every number here is a judgement that someone should be able to argue with,
# which is why they are in a file rather than buried in the code.
parameter,value,note
burns_slope,22,"Burns 1968: largest particle um = slope x body length mm + intercept"
burns_intercept,4.87,"Burns 1968"
mesh_fine_um,0.5,"fine mesh (Diaphanosoma, Ceriodaphnia, Chydorus, D. cucullata, D. magna). INTERPOLATED from Brendelberger's 0.2-4.7 um range, not a published class value"
mesh_medium_um,1.0,"medium mesh (most Daphnia, Bosmina). INTERPOLATED, see above"
mesh_coarse_um,3.0,"coarse mesh (Holopedium, Sida). INTERPOLATED, see above"
mesh_default_um,1.0,"used where the mesh class is unknown, and for rotifers and copepods, which do not sieve like cladocerans: a placeholder"
interference_length_um,100,"colonies at or above this size foul the filter rather than being eaten. OUR PLACEHOLDER: the effect is well established for filamentous cyanobacteria but we have not verified a published threshold. Filaments are excluded at any size, because the source standardises them to 100 um"
raptorial_particle_multiplier,2.0,"raptorial feeders handle particles this many times larger than the Burns limit"
prey_ratio_min,0.05,"smallest prey a predator takes, as a fraction of its own body length"
prey_ratio_max,0.50,"largest prey a predator takes, as a fraction of its own body length"
rotifer_max_particle_um,20,"upper limit for microphagous rotifers, which do not follow Burns"
"""

# Filter mesh classes, following the species Brendelberger (1991) actually measured:
#   fine    Diaphanosoma brachyurum, Ceriodaphnia quadrangula, Chydorus sphaericus,
#           Daphnia cucullata, Daphnia magna
#   medium  Daphnia galeata, D. hyalina, D. pulicaria, Bosmina coregoni
#   coarse  Holopedium gibberum, Sida crystallina
# Setule spacing ranged from 0.2 um (Diaphanosoma) to 4.7 um (Sida).
#
# TWO THINGS TO KNOW:
#  * Daphnia spans fine and medium by species, so the genus default below is medium
#    and the per-species table overrides it for the two he found to be fine.
#  * The numeric values for each class (in the rules file) are interpolated from his
#    0.2-4.7 um range. They are NOT published class values. Change them if you have
#    better figures.
# Mesh size governs retention EFFICIENCY rather than a hard cutoff, so treating it as
# a lower bound is a simplification. The concept also belongs to cladoceran filtering
# combs: rotifers and copepods do not sieve this way, and for them the lower bound is
# little more than a placeholder.
MESH_CLASS = {
    "fine": ("Diaphanosoma", "Ceriodaphnia", "Chydorus"),
    "coarse": ("Holopedium", "Sida", "Latona", "Latonopsis"),
}
MESH_SPECIES = {                      # overrides the genus default
    "Daphnia cucullata": "fine",
    "Daphnia magna": "fine",
}


def rows_of(path, sheet=None):
    wb = load_workbook(Path(path).expanduser(), read_only=True, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb[wb.sheetnames[0]]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    h = [str(c).strip() if c is not None else "" for c in rows[0]]
    return [dict(zip(h, r)) for r in rows[1:] if any(x is not None for x in r)]


def get(d, *names, default=""):
    low = {str(k).strip().lower(): v for k, v in d.items() if k}
    for n in names:
        v = low.get(n.lower())
        if v not in (None, ""):
            return v
    return default


def fnum(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"(\d+(?:\.\d+)?)", str(v).replace(",", "."))
    return float(m.group(1)) if m else None


def frange(v):
    """'0.8-2.2' -> (0.8, 2.2); '1.5' -> (1.5, 1.5)."""
    if v in (None, ""):
        return (None, None)
    s = str(v)
    m = re.findall(r"(\d+(?:\.\d+)?)", s.replace(",", "."))
    if not m:
        return (None, None)
    if len(m) == 1:
        return (float(m[0]), float(m[0]))
    return (float(m[0]), float(m[1]))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--zooplankton", required=True, help="the key's workbook (feeding_groups sheet)")
    ap.add_argument("--phytoplankton", required=True, help="phyto_traits.xlsx")
    ap.add_argument("--out", required=True)
    ap.add_argument("--counts", default="", help="restrict to the taxa in this CSV")
    ap.add_argument("--rules", default="", help="rule parameters (written for you if absent)")
    a = ap.parse_args()

    # ---- rules ----
    R = {}
    notes = {}
    if a.rules:
        rp = Path(a.rules).expanduser()
        if not rp.exists():
            rp.parent.mkdir(parents=True, exist_ok=True)
            rp.write_text(DEFAULT_RULES, encoding="utf-8")
            print(f"(Wrote the rule parameters to {rp}: edit and rerun to change the web.)")
        text = rp.read_text(encoding="utf-8")
    else:
        text = DEFAULT_RULES
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or line.lower().startswith("parameter,"):
            continue
        bits = next(csv.reader([line]))
        if len(bits) >= 2 and fnum(bits[1]) is not None:
            R[bits[0].strip()] = fnum(bits[1])
            notes[bits[0].strip()] = bits[2] if len(bits) > 2 else ""

    wanted = None
    if a.counts:
        wanted = set()
        with Path(a.counts).expanduser().open(encoding="utf-8-sig", newline="") as fh:
            for row in csv.DictReader(fh):
                t = get(row, "taxon", "name", "species", "genus")
                if t:
                    wanted.add(str(t).strip().lower())
        print(f"restricting to {len(wanted)} taxa from {a.counts}")

    # ---- the animals ----
    zoo = rows_of(a.zooplankton, "feeding_groups")
    if not zoo:
        sys.exit("No 'feeding_groups' sheet found. Run draft_feeding_groups.py first.")
    animals, no_size = [], []
    for r in zoo:
        name = str(get(r, "display_name", "species_id")).strip()
        if not name or (wanted and name.lower() not in wanted
                        and str(get(r, "genus")).lower() not in wanted):
            continue
        mode = str(get(r, "feeding_mode", "feeding_mode_draft")).strip().lower()
        group = str(get(r, "trophic_group", "trophic_group_draft")).strip().lower()
        lo, hi = frange(get(r, "body_length_min_mm"))
        hi2 = fnum(get(r, "body_length_max_mm"))
        if hi2:
            hi = hi2
        if lo is None:
            lo, hi = frange(get(r, "body_length_mm_draft"))
        genus = str(get(r, "genus")).strip()
        mesh = "medium"
        for cls, genera in MESH_CLASS.items():
            if genus in genera:
                mesh = cls
        mesh = MESH_SPECIES.get(name, mesh)       # a species Brendelberger measured himself
        animals.append({"name": name, "genus": genus, "mode": mode, "group": group,
                        "len_min": lo, "len_max": hi, "mesh": mesh,
                        "phylum": str(get(r, "phylum", "Phylum")).strip(),
                        "basis": str(get(r, "draft_basis")).strip(),
                        "confidence": str(get(r, "confidence")).strip()})
        if hi is None:
            no_size.append(name)

    # ---- the algae ----
    phyto = []
    for r in rows_of(a.phytoplankton, "phyto_traits"):
        if str(get(r, "role")).strip().lower() not in ("phytoplankton", "protozoan consumer", ""):
            continue
        taxon = str(get(r, "taxon")).strip()
        if wanted and taxon.lower() not in wanted and str(get(r, "genus")).lower() not in wanted:
            continue
        mx = fnum(get(r, "max_dimension_um"))
        if mx is None:
            continue
        phyto.append({"taxon": taxon, "genus": str(get(r, "genus")).strip(),
                      "unit": str(get(r, "counting_unit")).strip(),
                      "max_um": mx, "biovolume": fnum(get(r, "biovolume_um3")),
                      "morph": str(get(r, "morphology_phycokey")).strip().lower(),
                      "role": str(get(r, "role")).strip().lower(),
                      "row_type": str(get(r, "row_type")).strip()})

    print(f"animals: {len(animals)}  ({len(no_size)} without a body length)")
    print(f"food items (counting units with a size): {len(phyto)}")

    # ---- the links ----
    links = []
    interference = 0
    for an in animals:
        L = an["len_max"] or an["len_min"]
        mode, group = an["mode"], an["group"]
        is_rotifer = an["phylum"].lower().startswith("rotifer")

        if mode in ("filter", "suspension") or group in ("herbivore", "omnivore"):
            if is_rotifer:
                upper = R["rotifer_max_particle_um"]
                rule = f"microphagous rotifer: particles up to {upper:g} um"
                ref = "rotifer feeding limit (rule parameter)"
            elif L:
                upper = R["burns_slope"] * L + R["burns_intercept"]
                rule = f"Burns 1968: 22 x {L:g} mm + 4.87 = {upper:.0f} um"
                ref = REF_BURNS
            else:
                continue
            lower = R.get(f"mesh_{an['mesh']}_um", R["mesh_default_um"])
            if mode == "raptorial":
                upper *= R["raptorial_particle_multiplier"]
            for ph in phyto:
                if ph["role"] == "protozoan consumer":
                    continue
                if not (lower <= ph["max_um"] <= upper):
                    continue
                # A filament fouls the filtering comb whatever its recorded length:
                # Rimet standardises filaments to 100 um, but real ones are longer.
                # A colony blocks only once it is bigger than the threshold.
                is_filament = (ph["unit"] == "filament_100um" or "filament" in ph["morph"])
                big_colony = (ph["unit"] == "colony"
                              and ph["max_um"] >= R["interference_length_um"])
                blocked = is_filament or big_colony
                if blocked:
                    interference += 1
                    continue
                links.append({"consumer": an["name"], "resource": ph["taxon"],
                              "resource_unit": ph["unit"], "link_type": "grazing",
                              "rule": rule, "mesh_lower_um": lower, "upper_um": round(upper, 1),
                              "resource_size_um": ph["max_um"],
                              "confidence": an["confidence"] or "", "reference": ref,
                              "note": f"{an['mesh']} mesh; {ref == REF_BURNS and 'size rule' or 'fixed limit'}"})

        # Predation needs a RAPTORIAL mode: an omnivorous filter feeder takes protists,
        # not other crustaceans, so "omnivore" alone must not trigger this rule.
        if mode in ("raptorial", "parasitic") or (
                group in ("carnivore", "parasite")
                and mode not in ("filter", "suspension", "scraper")):
            if not L:
                continue
            lo_p, hi_p = R["prey_ratio_min"] * L, R["prey_ratio_max"] * L
            for pr in animals:
                if pr["name"] == an["name"]:
                    continue
                pl = pr["len_max"] or pr["len_min"]
                if pl and lo_p <= pl <= hi_p:
                    links.append({"consumer": an["name"], "resource": pr["name"],
                                  "resource_unit": "animal", "link_type":
                                  "parasitism" if mode == "parasitic" else "predation",
                                  "rule": f"prey {R['prey_ratio_min']:g}-{R['prey_ratio_max']:g} "
                                          f"x predator length ({lo_p:.2f}-{hi_p:.2f} mm)",
                                  "mesh_lower_um": "", "upper_um": "", "resource_size_um": pl * 1000,
                                  "confidence": an["confidence"] or "",
                                  "reference": "predator:prey body size window (rule parameter)",
                                  "note": ""})

        if mode == "scraper" or group == "detritivore":
            links.append({"consumer": an["name"], "resource": "detritus / periphyton",
                          "resource_unit": "non-plankton", "link_type": "detritivory",
                          "rule": "scrapers and browsers feed on surfaces, not the water column",
                          "mesh_lower_um": "", "upper_um": "", "resource_size_um": "",
                          "confidence": an["confidence"] or "",
                          "reference": "feeding mode", "note": "not counted by either app"})

    # ---- write ----
    outp = Path(a.out).expanduser()
    outp.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "links"
    HEADS = ["consumer", "resource", "resource_unit", "link_type", "resource_size_um",
             "mesh_lower_um", "upper_um", "rule", "confidence", "reference", "note"]
    ws.append(HEADS)
    for r in links:
        ws.append([r.get(k) for k in HEADS])

    ws2 = wb.create_sheet("nodes")
    ws2.append(["node", "kind", "body_length_mm", "feeding_mode", "trophic_group", "mesh_class", "note"])
    for an in animals:
        ws2.append([an["name"], "animal", an["len_max"] or an["len_min"], an["mode"],
                    an["group"], an["mesh"], "" if (an["len_max"] or an["len_min"])
                    else "NO BODY LENGTH: cannot be linked"])
    seen_r = set()
    for ph in phyto:
        if ph["taxon"] in seen_r:
            continue
        seen_r.add(ph["taxon"])
        ws2.append([ph["taxon"], "phytoplankton", "", "", "", "", ph["unit"]])

    ws3 = wb.create_sheet("rules_used")
    ws3.append(["parameter", "value", "note"])
    for k, v in R.items():
        ws3.append([k, v, notes.get(k, "")])
    ws3.append([])
    ws3.append(["reference", REF_BURNS])
    ws3.append(["reference", REF_MESH])

    for sheet, widths in ((ws, {"consumer": 30, "resource": 34, "rule": 44, "reference": 70, "note": 28}),
                          (ws2, {"node": 34, "note": 34}), (ws3, {"parameter": 30, "note": 60})):
        heads = [c.value for c in sheet[1]]
        for i, k in enumerate(heads, 1):
            c = sheet.cell(row=1, column=i)
            c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            sheet.column_dimensions[get_column_letter(i)].width = widths.get(str(k), 16)
        for row in sheet.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=False)
        sheet.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(HEADS))}{ws.max_row}"
    wb.save(outp)

    kinds = defaultdict(int)
    for l in links:
        kinds[l["link_type"]] += 1
    print(f"\nWrote {outp}")
    print(f"  {len(links)} links")
    for k, n in sorted(kinds.items(), key=lambda kv: -kv[1]):
        print(f"    {n:6}  {k}")
    print(f"  {interference} grazing links withheld: the food is too long or too "
          f"large a colony and fouls the filter instead")
    consumers = {l["consumer"] for l in links}
    print(f"\nanimals with at least one link: {len(consumers)} of {len(animals)}")
    if no_size:
        print(f"\n{len(no_size)} animals have NO BODY LENGTH, so no rule can apply to them:")
        print("   " + ", ".join(sorted(no_size)[:12]) + (" ..." if len(no_size) > 12 else ""))
        print("   Filling in body_length_min_mm and body_length_max_mm in the feeding_groups")
        print("   sheet is the single change that most improves the web.")


if __name__ == "__main__":
    main()
