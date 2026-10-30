#!/usr/bin/env python3
"""
build_phyto_traits.py

Builds one phytoplankton trait table for the food-web tool, joining:

  * Rimet & Druart (2018) for cell, colony and filament dimensions, biovolumes,
    geometric shapes, functional groups and the nano/micro split;
  * the PhycoKey mirror for genus-level morphology, taken from the folder path
    each genus sits in (colony, filament, branched, heterocysts and so on);
  * the geometric shape equations, read out of the Rimet workbook's own
    "read me 2" sheet, so the formulas come from the source and not from us;
  * size fractions, including the 50 um "edible" threshold.

    conda activate video
    python build_phyto_traits.py \\
        --rimet "RimetDruart2018Appendix-1-Phytoplankton_metrics_database-revised.xlsx" \\
        --phycokey "cfb.unh.edu/phycokey" \\
        --out "$ZK/data/phyto_traits.xlsx"

WHAT YOU GET
  One row per COUNTING UNIT, not per species, because a count of "Microcystis: 40"
  means nothing until you know whether those were cells or colonies: the two differ
  eighty-fold for the same species. Each row says which unit it describes.

  Rows come in three kinds, marked in 'row_type':
    species_unit  a species measured in Rimet, one row per unit (cell / colony /
                  filament_100um), carrying that unit's own dimensions.
    genus_summary one row per genus per unit: the median, standard deviation and
                  sample size across the species Rimet holds for that genus. Use
                  this when a count was only taken to genus, which is common.
    genus_only    a genus present in the PhycoKey but absent from Rimet: morphology
                  and shape are given, sizes are not.

  A species not in Rimet has no row of its own. Fall back to its genus_summary row:
  that is what the median, SD and n are for, and 'n_species_in_genus' tells you how
  much to trust it.

SIZE FRACTIONS (on the maximum dimension of the counting unit)
  <2 um picoplankton, 2-20 nano, 20-50, 50-100, >100 micro. 'edible_lt_50um' flags
  the fraction most often treated as available to zooplankton grazing. The threshold
  has moved around in the literature, so the raw dimension is kept beside the flag.
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
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

REF_RIMET = "Rimet & Druart 2018, Ann. Limnol. 54:18, doi:10.1051/limn/2018009 (Zenodo 10.5281/zenodo.1164834)"
# A Rimet entry that is not a genus: a higher rank, or two look-alikes recorded
# together. These can never match a PhycoKey genus folder.
NOT_A_GENUS = re.compile(r"\+|\s|aceae$|phyceae$|ales$|phyta$|idae$|inae$", re.I)

STARTER_ALIASES = """# One line per name difference between Rimet & Druart and the PhycoKey.
# rimet_name,phycokey_name,note
# Every line needs checking by someone who knows the algae: these are starting
# points only. Add your own as you find them.
Aulacoseira,Melosira,"Aulacoseira was separated from Melosira; the PhycoKey uses the older name. CHECK"
Dolichospermum,Anabaena,"planktic Anabaena moved to Dolichospermum in 2009. CHECK"
Woronichinia,Gomphosphaeria,"segregated from Gomphosphaeria. CHECK"
Snowella,Gomphosphaeria,"segregated from Gomphosphaeria. CHECK"
Cyanobium,Synechococcus,"small picocyanobacteria, often lumped with Synechococcus. CHECK"
Planktothrix,Oscillatoria,"segregated from Oscillatoria. CHECK"
Limnothrix,Oscillatoria,"segregated from Oscillatoria. CHECK"
"""

REF_PHYCO = "Baker, A.L. et al. 2012. PhycoKey, UNH Center for Freshwater Biology, cfb.unh.edu/phycokey"

# Words that appear in PhycoKey folder names, and what they tell us. Order matters
# only for readability; every match is recorded.
# Words that appear in PhycoKey folder names and what they tell us.
# Each rule is (pattern, label, labels it rules out), so "unbranched" does not
# also record "branched", and "no_visible_sheath" does not record "sheath".
PATH_TRAITS = [
    (r"unbranched", "unbranched", ("branched", "pseudobranched")),
    (r"pseudobranch", "pseudobranched", ("branched",)),
    (r"branch", "branched", ()),
    (r"untapered", "untapered", ("tapered",)),
    (r"tapered", "tapered", ()),
    (r"no_visible_sheath|no_sheath", "no visible sheath", ("sheath present",)),
    (r"sheath", "sheath present", ()),
    (r"colonial|colonies|colony", "colonial", ()),
    (r"filament", "filamentous", ()),
    (r"heterocyst", "heterocysts present", ()),
    (r"centric", "centric", ()),
    (r"pennate", "pennate", ()),
    (r"coeno", "coenobial", ()),
    (r"flagellate", "flagellate", ()),
    (r"single|unicell|1cell|one_cell", "unicellular", ()),
    (r"mucilag", "mucilaginous", ()),
    (r"spine", "spines", ()),
    (r"star|astero", "stellate", ()),
    (r"spiral|helic", "spiral", ()),
    (r"nonps|non-ps|non_ps|colorless|colourless|heterotroph", "non-photosynthetic", ("photosynthetic",)),
    (r"(^|[/_ -])ps[_ -]|photosynthetic", "photosynthetic", ()),
    (r"naked|athecate", "naked (no theca)", ()),
    (r"thecate|armou?red", "thecate (armoured)", ()),
    (r"desmid", "desmid", ()),
    (r"diatom", "diatom", ()),
    (r"cyst|resting", "cyst or resting stage", ()),
]
SKIP_DIR = re.compile(r"^(choices|templates|images?|files|photos?|pics?|_notes|page|key|"
                      r"image_page|thumbnails?)$", re.I)


def read_shapes(wb):
    """The shape equations, from the workbook's own read me sheet."""
    out = {}
    for name in wb.sheetnames:
        if "read me" not in name.lower():
            continue
        rows = list(wb[name].iter_rows(values_only=True))
        for r in rows:
            cells = [str(c).strip() for c in r if c is not None]
            if len(cells) >= 3 and cells[0] and re.fullmatch(r"[a-z']{2,7}", cells[0]):
                out[cells[0].lower()] = {"shape_name": cells[1],
                                         "volume_equation": cells[2],
                                         "surface_equation": cells[3] if len(cells) > 3 else ""}
    # the workbook writes some acronyms one way in the table and another in the data
    for a, b in (("tub", "tube"), ("con", "cone"), ("prisme", "prism")):
        if a in out and b not in out:
            out[b] = out[a]
    return out


def phycokey_morphology(root: Path):
    """genus -> {'morphology': ..., 'class': ..., 'path': ...} from the folder tree."""
    out = {}
    if not root or not root.is_dir():
        return out
    import fnmatch
    keyfiles = [f for f in root.rglob("*") if f.is_file()
                and fnmatch.fnmatch(f.name.lower(), "*_key.htm*")]
    for f in keyfiles:
        parts = list(f.relative_to(root).parts[:-1])
        if not parts:
            continue
        genus_dir = parts[-1]
        # a genus folder is ALL CAPS or matches the file name before _key
        stem = re.sub(r"_key$", "", f.stem, flags=re.I)
        if not re.fullmatch(r"[A-Za-z][A-Za-z\-]+", genus_dir) or SKIP_DIR.match(genus_dir):
            continue
        if genus_dir.upper() != genus_dir and genus_dir.lower() != stem.lower():
            continue                       # a grouping folder, not a genus
        genus = genus_dir.capitalize() if genus_dir.upper() == genus_dir else genus_dir
        genus = re.sub(r"[^A-Za-z].*$", "", genus).capitalize()
        if len(genus) < 3:
            continue
        # only the folders ABOVE the genus describe it: the genus name itself must not
        # be read as a trait (Microcystis contains "cyst", Nitzschia contains "schia")
        joined = "/".join(parts[:-1]).lower()
        traits, seen, blocked = [], set(), set()
        for pat, label, rules_out in PATH_TRAITS:
            if label in blocked or label in seen:
                continue
            if re.search(pat, joined):
                seen.add(label)
                traits.append(label)
                blocked.update(rules_out)
        cls = parts[1] if len(parts) > 1 and parts[0].lower() == "choices" else parts[0]
        prev = out.get(genus)
        if prev and len(prev["morphology"]) >= len("; ".join(traits)):
            continue
        # the folder names between the class and the genus, tidied, so a human can read
        # the placement even where our own vocabulary has nothing to say about it
        between = [x.replace("_", " ").replace("-", " ").strip() for x in parts[2:-1]]
        out[genus] = {"morphology": "; ".join(traits),
                      "phycokey_group": "; ".join(x for x in between if x),
                      "phycokey_class": cls.replace("_", " "),
                      "phycokey_path": "/".join(parts)}
    return out


# What a PhycoKey class means for a freshwater food web. Anything not listed is
# treated as phytoplankton. Starting points for the team to correct.
CLASS_ROLE = {
    "anomalous items": ("not an organism", "detritus, bubbles, crystals and other objects"),
    "amoebae flagellates ciliates": ("protozoan consumer", "a consumer as well as prey, not a producer"),
    "rhodophyceae": ("macroalga", "red algae: nearly all marine, mostly macroalgae"),
    "fucophyceae": ("macroalga", "brown algae: nearly all marine, mostly macroalgae"),
    "charophyceae": ("uncertain", "desmids are planktonic; Chara and relatives are macroalgae"),
    "cyanidiales": ("uncertain", "extremophile red algae, not typical lake plankton"),
    "videos": ("not an organism", ""),
    "mix": ("not an organism", "mixed sample illustrations"),
    "commonnh": ("phytoplankton", "from the New Hampshire waterbody pages"),
    "unidentified": ("uncertain", ""),
}


def role_for(phyco_class, phylum):
    key = str(phyco_class or "").strip().lower()
    if key in CLASS_ROLE:
        return CLASS_ROLE[key]
    if key or str(phylum or "").strip():
        return ("phytoplankton", "")
    return ("", "")


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
    ap.add_argument("--rimet", required=True, help="the Rimet & Druart 2018 spreadsheet")
    ap.add_argument("--phycokey", default="", help="the PhycoKey mirror folder (optional)")
    ap.add_argument("--out", required=True, help="the trait table to write (.xlsx)")
    ap.add_argument("--csv", action="store_true", help="also write a .csv beside it")
    ap.add_argument("--aliases", default="",
                    help="a CSV of name differences (rimet_name,phycokey_name,note). "
                         "If the file does not exist it is created with a few starting points.")
    a = ap.parse_args()

    wb_in = load_workbook(Path(a.rimet).expanduser(), read_only=True, data_only=True)
    shapes = read_shapes(wb_in)
    ws = wb_in["Phytoplankton traits database"]
    rows = list(ws.iter_rows(values_only=True))
    h = [str(x) if x is not None else "" for x in rows[0]]

    def idx(name):
        return h.index(name) if name in h else None
    C = {k: idx(v) for k, v in {
        "name": "Genus + species name", "phylum": "Phylum", "cls": "Class", "order": "Order",
        "family": "Family", "genus": "Genus", "species": "Species",
        "cell_l": "Cell length µm", "cell_w": "Cell width µm", "cell_t": "Cell thickness µm",
        "cell_shape": "Geometrical shape of the cell", "cell_bv": "Cell biovolume µm3",
        "ncell": "Number of cells per colony",
        "col_l": "Colony length (takes into account mucilage) µm",
        "col_w": "Colony width (takes into account mucilage) µm",
        "col_t": "Colony thickness (takes into account mucilage) µm",
        "col_shape": "Geometrical shape of the colony",
        "col_bv": "Colony biovolume µm3 (without mucilage)",
        "cum_bv": "Cumulated biovolume of cells in a colony µm3",
        "ref_bv": "Reference biovolume", "maxlen": "Maximal length of the algal object",
        "notes": "Notes on biovolumes", "carbon": "Organic carbon ratio (Wetzel et Likens 2000)",
        "reynolds": "Functional groups (Reynolds 2002)", "padisak": "Functional groups (Padisak 2009)",
        "kruk": "Morpho-classification (Kruk 2010)", "nano": "Nano/microphytoplankton",
        "colonial": "Colonial", "filament": "Filament", "hetero": "Heterotrophic",
        "mixo": "Mixotrophic", "auto": "Autotrophic"}.items()}
    missing = [k for k, v in C.items() if v is None]
    if missing:
        print(f"  note: columns not found in the Rimet sheet: {', '.join(missing)}")

    def val(r, key):
        i = C.get(key)
        return r[i] if i is not None and i < len(r) else None

    def num(r, key):
        try:
            return float(val(r, key))
        except (TypeError, ValueError):
            return None

    phyco = phycokey_morphology(Path(a.phycokey).expanduser() if a.phycokey else None)
    print(f"PhycoKey genera found in the folder tree: {len(phyco)}")

    # names that differ between the two sources
    aliases = {}
    if a.aliases:
        ap_ = Path(a.aliases).expanduser()
        if not ap_.exists():
            ap_.parent.mkdir(parents=True, exist_ok=True)
            ap_.write_text(STARTER_ALIASES, encoding="utf-8")
            print(f"(Wrote a starter alias list to {ap_}: check and extend it.)")
        for line in ap_.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            bits = next(csv.reader([line]))
            if len(bits) >= 2 and bits[0].strip() and bits[1].strip():
                aliases[bits[0].strip().capitalize()] = bits[1].strip().capitalize()
        used = 0
        for rimet_name, phyco_name in aliases.items():
            if rimet_name not in phyco and phyco_name in phyco:
                phyco[rimet_name] = {**phyco[phyco_name],
                                     "morphology": phyco[phyco_name]["morphology"],
                                     "alias_of": phyco_name}
                used += 1
        print(f"name differences applied from the alias list: {used} of {len(aliases)}")

    HEADS = ["row_type", "role", "role_note", "taxon", "genus", "species", "counting_unit", "unit_note",
             "phylum", "class", "order", "family",
             "length_um", "width_um", "thickness_um", "max_dimension_um",
             "biovolume_um3", "biovolume_sd", "n_species_in_genus",
             "shape_code", "shape_name", "volume_equation",
             "size_fraction", "edible_lt_50um", "nano_micro",
             "cells_per_colony", "morphology_phycokey", "phycokey_group", "phycokey_class",
             "trophic_mode", "reynolds_group", "padisak_group", "kruk_group",
             "organic_carbon_ratio", "measurement_note", "reference"]

    out_rows = []
    by_genus_unit = defaultdict(list)          # (genus, unit) -> [biovolume, ...]
    genus_meta = {}

    for r in rows[1:]:
        if not r or not val(r, "name"):
            continue
        genus = str(val(r, "genus") or "").strip()
        species = str(val(r, "species") or "").strip()
        trophic = ("heterotrophic" if num(r, "hetero") == 1 else
                   "mixotrophic" if num(r, "mixo") == 1 else
                   "autotrophic" if num(r, "auto") == 1 else "")
        base = {"genus": genus, "species": species, "taxon": str(val(r, "name")).strip(),
                "phylum": val(r, "phylum"), "class": val(r, "cls"), "order": val(r, "order"),
                "family": val(r, "family"), "nano_micro": val(r, "nano"),
                "trophic_mode": trophic, "reynolds_group": val(r, "reynolds"),
                "padisak_group": val(r, "padisak"), "kruk_group": val(r, "kruk"),
                "organic_carbon_ratio": num(r, "carbon"),
                "measurement_note": val(r, "notes"), "reference": REF_RIMET}
        if genus:
            genus_meta.setdefault(genus, {k: base[k] for k in
                                          ("phylum", "class", "order", "family")})
        units = []
        # the cell, always
        if num(r, "cell_bv"):
            units.append(("cell", "", num(r, "cell_l"), num(r, "cell_w"), num(r, "cell_t"),
                          val(r, "cell_shape"), num(r, "cell_bv"), None))
        # the colony, where the taxon is colonial and not a filament
        if num(r, "col_bv") and num(r, "filament") != 1:
            units.append(("colony", "includes mucilage in the dimensions",
                          num(r, "col_l"), num(r, "col_w"), num(r, "col_t"),
                          val(r, "col_shape"), num(r, "col_bv"), num(r, "ncell")))
        # a filament, standardised to 100 um of length by the source
        if num(r, "col_bv") and num(r, "filament") == 1:
            units.append(("filament_100um", "biovolume is for 100 um of filament",
                          num(r, "col_l"), num(r, "col_w"), num(r, "col_t"),
                          val(r, "col_shape"), num(r, "col_bv"), num(r, "ncell")))
        for unit, note, l, w, t, shape, bv, ncell in units:
            dims = [x for x in (l, w, t) if x]
            mx = max(dims) if dims else None
            if unit == "cell" and num(r, "maxlen") and num(r, "colonial") != 1:
                mx = num(r, "maxlen")
            sh = shapes.get(str(shape or "").strip().lower(), {})
            _role, _rnote = role_for(phyco.get(genus, {}).get("phycokey_class"), base["phylum"])
            out_rows.append({**base, "row_type": "species_unit", "role": _role,
                             "role_note": _rnote, "counting_unit": unit,
                             "unit_note": note, "length_um": l, "width_um": w,
                             "thickness_um": t, "max_dimension_um": mx,
                             "biovolume_um3": bv, "shape_code": shape,
                             "shape_name": sh.get("shape_name", ""),
                             "volume_equation": sh.get("volume_equation", ""),
                             "size_fraction": size_fraction(mx),
                             "edible_lt_50um": ("yes" if mx and mx < 50 else
                                                "no" if mx else ""),
                             "cells_per_colony": ncell,
                             "morphology_phycokey": phyco.get(genus, {}).get("morphology", ""),
                             "phycokey_group": phyco.get(genus, {}).get("phycokey_group", ""),
                             "phycokey_class": phyco.get(genus, {}).get("phycokey_class", "")})
            if genus and bv:
                by_genus_unit[(genus, unit)].append((bv, mx, shape))

    # ---- one summary row per genus per unit, for counts taken only to genus ----
    for (genus, unit), vals in sorted(by_genus_unit.items()):
        bvs = [v[0] for v in vals]
        mxs = [v[1] for v in vals if v[1]]
        med = st.median(bvs)
        sd = st.stdev(bvs) if len(bvs) > 1 else None
        med_mx = st.median(mxs) if mxs else None
        common_shape = max(set(v[2] for v in vals if v[2]), default="",
                           key=lambda s: sum(1 for v in vals if v[2] == s)) if any(v[2] for v in vals) else ""
        sh = shapes.get(str(common_shape or "").strip().lower(), {})
        meta = genus_meta.get(genus, {})
        _role, _rnote = role_for(phyco.get(genus, {}).get("phycokey_class"), meta.get("phylum"))
        out_rows.append({"row_type": "genus_summary", "role": _role, "role_note": _rnote,
                         "taxon": f"{genus} (genus median)",
                         "genus": genus, "species": "", "counting_unit": unit,
                         "unit_note": f"median of {len(bvs)} species in Rimet & Druart",
                         "phylum": meta.get("phylum"), "class": meta.get("class"),
                         "order": meta.get("order"), "family": meta.get("family"),
                         "length_um": None, "width_um": None, "thickness_um": None,
                         "max_dimension_um": med_mx, "biovolume_um3": med,
                         "biovolume_sd": sd, "n_species_in_genus": len(bvs),
                         "shape_code": common_shape, "shape_name": sh.get("shape_name", ""),
                         "volume_equation": sh.get("volume_equation", ""),
                         "size_fraction": size_fraction(med_mx),
                         "edible_lt_50um": ("yes" if med_mx and med_mx < 50 else
                                            "no" if med_mx else ""),
                         "nano_micro": "", "cells_per_colony": None,
                         "morphology_phycokey": phyco.get(genus, {}).get("morphology", ""),
                         "phycokey_group": phyco.get(genus, {}).get("phycokey_group", ""),
                         "phycokey_class": phyco.get(genus, {}).get("phycokey_class", ""),
                         "trophic_mode": "", "reynolds_group": "", "padisak_group": "",
                         "kruk_group": "", "organic_carbon_ratio": None,
                         "measurement_note": "", "reference": REF_RIMET})

    # ---- genera the PhycoKey has and Rimet does not ----
    rimet_genera = {g for (g, _u) in by_genus_unit}
    for genus, info in sorted(phyco.items()):
        if genus in rimet_genera:
            continue
        _role, _rnote = role_for(info["phycokey_class"], "")
        out_rows.append({"row_type": "genus_only", "role": _role, "role_note": _rnote,
                         "taxon": f"{genus} (no size data)",
                         "genus": genus, "species": "", "counting_unit": "",
                         "unit_note": "in the PhycoKey but not in Rimet & Druart: no size available",
                         "phylum": "", "class": info["phycokey_class"], "order": "", "family": "",
                         "length_um": None, "width_um": None, "thickness_um": None,
                         "max_dimension_um": None, "biovolume_um3": None, "biovolume_sd": None,
                         "n_species_in_genus": 0, "shape_code": "", "shape_name": "",
                         "volume_equation": "", "size_fraction": "", "edible_lt_50um": "",
                         "nano_micro": "", "cells_per_colony": None,
                         "morphology_phycokey": info["morphology"],
                         "phycokey_group": info.get("phycokey_group", ""),
                         "phycokey_class": info["phycokey_class"], "trophic_mode": "",
                         "reynolds_group": "", "padisak_group": "", "kruk_group": "",
                         "organic_carbon_ratio": None, "measurement_note": "",
                         "reference": REF_PHYCO})

    # ---- write ----
    outp = Path(a.out).expanduser()
    outp.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    sh_ = wb.active
    sh_.title = "phyto_traits"
    sh_.append(HEADS)
    for r in out_rows:
        sh_.append([r.get(k) for k in HEADS])
    widths = {"taxon": 36, "role": 20, "role_note": 46, "genus": 18, "species": 18, "counting_unit": 15, "unit_note": 42,
              "volume_equation": 46, "morphology_phycokey": 34, "phycokey_group": 30, "measurement_note": 28,
              "reference": 60, "size_fraction": 15, "shape_name": 18, "row_type": 14}
    for i, k in enumerate(HEADS, 1):
        c = sh_.cell(row=1, column=i)
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        sh_.column_dimensions[get_column_letter(i)].width = widths.get(k, 14)
    for row in sh_.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top",
                                    wrap_text=HEADS[c.column - 1] in ("unit_note", "reference", "role_note"))
    sh_.freeze_panes = "C2"
    sh_.auto_filter.ref = f"A1:{get_column_letter(len(HEADS))}{sh_.max_row}"

    # a second sheet listing the shape equations, so the table explains itself
    sh2 = wb.create_sheet("shape_equations")
    sh2.append(["shape_code", "shape_name", "volume_equation", "surface_equation", "source"])
    for code, d in sorted(shapes.items()):
        sh2.append([code, d["shape_name"], d["volume_equation"], d["surface_equation"], REF_RIMET])
    for i, w in enumerate((12, 20, 42, 60, 60), 1):
        sh2.cell(row=1, column=i).font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        sh2.cell(row=1, column=i).fill = PatternFill("solid", fgColor="1F3864")
        sh2.column_dimensions[get_column_letter(i)].width = w
    wb.save(outp)

    if a.csv:
        with outp.with_suffix(".csv").open("w", newline="", encoding="utf-8-sig") as fh:
            w_ = csv.DictWriter(fh, fieldnames=HEADS)
            w_.writeheader()
            w_.writerows(out_rows)

    kinds = defaultdict(int)
    for r in out_rows:
        kinds[r["row_type"]] += 1
    print(f"\nWrote {outp}  ({len(out_rows)} rows)")
    for k in ("species_unit", "genus_summary", "genus_only"):
        print(f"  {kinds[k]:5}  {k}")
    units = defaultdict(int)
    for r in out_rows:
        if r["row_type"] == "species_unit":
            units[r["counting_unit"]] += 1
    print("\nspecies rows by counting unit:")
    for k, n in sorted(units.items(), key=lambda kv: -kv[1]):
        print(f"  {n:5}  {k}")
    frac = defaultdict(int)
    for r in out_rows:
        if r["size_fraction"]:
            frac[r["size_fraction"]] += 1
    print("\nrows by size fraction:")
    for k in ("<2 um (pico)", "2-20 um (nano)", "20-50 um", "50-100 um", ">100 um"):
        if frac[k]:
            print(f"  {frac[k]:5}  {k}")
    with_morph = sum(1 for r in out_rows if r["morphology_phycokey"])
    with_any = sum(1 for r in out_rows if r["morphology_phycokey"] or r["phycokey_group"]
                   or r["phycokey_class"])
    print(f"\nrows carrying PhycoKey morphology terms: {with_morph} of {len(out_rows)}")
    print(f"rows carrying any PhycoKey placement (class or folder group): {with_any}")
    rim_gen = {r["genus"] for r in out_rows if r["row_type"] == "species_unit" and r["genus"]}
    placed = {r["genus"] for r in out_rows
              if r["genus"] and (r["morphology_phycokey"] or r["phycokey_group"])}
    missing = sorted(g for g in rim_gen - placed if not NOT_A_GENUS.search(str(g)))
    higher = sorted(g for g in rim_gen - placed if NOT_A_GENUS.search(str(g)))
    print(f"\nRimet genera: {len(rim_gen)}   with a PhycoKey placement: {len(rim_gen & placed)}")
    print(f"  not in the PhycoKey under that name: {len(missing)}")
    print(f"  not a genus (higher rank, or two taxa recorded together): {len(higher)}")
    if missing:
        print("  worth checking for a different name in the PhycoKey:")
        print("   ", ", ".join(missing[:28]) + (" ..." if len(missing) > 28 else ""))
    roles = defaultdict(int)
    for r in out_rows:
        roles[r.get("role") or "(none)"] += 1
    print("\nrows by role:")
    for k, n in sorted(roles.items(), key=lambda kv: -kv[1]):
        print(f"  {n:5}  {k}")
    print("\nFilter role = phytoplankton for the rows a freshwater count can contain.")


if __name__ == "__main__":
    main()
