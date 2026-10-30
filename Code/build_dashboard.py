#!/usr/bin/env python3
"""
build_dashboard.py

Writes a self-contained food web dashboard: one HTML file that carries the trait
lookup and the feeding rules inside it, so a user can load a count, edit it, add or
remove taxa, change the rule parameters, and watch the web redraw. No server, no
Python, nothing to install. Open the file, or publish it beside the counters.

    conda activate video
    python build_dashboard.py \\
        --zooplankton "$ZK/data/cfbkey_workbook.xlsx" \\
        --phytoplankton "$ZK/data/phyto_traits.xlsx" \\
        --rules "$ZK/data/foodweb_rules.csv" \\
        --out "$ZK/docs/foodweb.html"

WHAT THE PAGE DOES
  * Loads a CSV exported by the Food Web Counter (taxon, group, counting_unit, count).
  * Lets the user edit counts, change a counting unit, and add or remove taxa, with
    name completion from the trait tables.
  * Applies the same rules as build_food_web.py, in the browser, and redraws.
  * Reports degree centrality, betweenness centrality, connectance and link count.
  * Node colour is the group, node size is abundance, so a taxon that is both
    abundant and well connected stands out.

WHAT TRAVELS WITH IT
  Only the fields the rules need: for an animal its body length, feeding mode and
  mesh class; for an alga its largest dimension and whether it is a colony or a
  filament. No taxonomy, no references, no images. The file is a few hundred
  kilobytes.

HONEST LIMITS, SHOWN ON THE PAGE ITSELF
  A taxon with no body length cannot be linked. A link means "physically possible
  on these rules", not "observed". The rule parameters that have no published
  source are marked as such in the panel where they are edited.
"""

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

# Names people actually stop at when counting, mapped to the rank we can aggregate.
# "Cyclopoid" is an order, but it is where a counter often stops, so it needs to work.
RANK_ALIASES = {
    "cyclopoid": ("Cyclopoida", "order"), "cyclopoids": ("Cyclopoida", "order"),
    "cyclopoida": ("Cyclopoida", "order"),
    "calanoid": ("Calanoida", "order"), "calanoids": ("Calanoida", "order"),
    "calanoida": ("Calanoida", "order"),
    "harpacticoid": ("Harpacticoida", "order"), "harpacticoida": ("Harpacticoida", "order"),
    "cladoceran": ("Cladocera", "suborder"), "cladocerans": ("Cladocera", "suborder"),
    "cladocera": ("Cladocera", "suborder"),
    "rotifer": ("Rotifera", "phylum"), "rotifers": ("Rotifera", "phylum"),
    "rotifera": ("Rotifera", "phylum"),
    "ostracod": ("Ostracoda", "class"), "ostracoda": ("Ostracoda", "class"),
}

MESH_CLASS = {"fine": ("Diaphanosoma", "Ceriodaphnia", "Chydorus"),
              "coarse": ("Holopedium", "Sida", "Latona", "Latonopsis")}
MESH_SPECIES = {"Daphnia cucullata": "fine", "Daphnia magna": "fine"}

DEFAULT_RULES = {
    "burns_slope": 22, "burns_intercept": 4.87,
    "mesh_fine_um": 0.5, "mesh_medium_um": 1.0, "mesh_coarse_um": 3.0, "mesh_default_um": 1.0,
    "interference_length_um": 100, "raptorial_particle_multiplier": 2.0,
    "prey_ratio_min": 0.05, "prey_ratio_max": 0.50, "rotifer_max_particle_um": 20,
}
RULE_NOTES = {
    "burns_slope": "Burns 1968, published",
    "burns_intercept": "Burns 1968, published",
    "mesh_fine_um": "interpolated from Brendelberger's 0.2-4.7 um range, not published",
    "mesh_medium_um": "interpolated, not published",
    "mesh_coarse_um": "interpolated, not published",
    "mesh_default_um": "placeholder for rotifers and copepods, which do not sieve",
    "interference_length_um": "OUR PLACEHOLDER: no verified published threshold",
    "raptorial_particle_multiplier": "NO SOURCE: a plausible guess",
    "prey_ratio_min": "predator:prey size window, rule of thumb",
    "prey_ratio_max": "predator:prey size window, rule of thumb",
    "rotifer_max_particle_um": "reasonable for microphagous rotifers, wrong for Asplanchna",
}


STARTER_EXTRAS = """# Taxa that no published trait table covers, but that people count anyway.
# EVERY LINE NEEDS CHECKING by someone who knows the organisms. Sizes are in mm for
# zooplankton (body length) and um for phytoplankton (largest dimension).
#
# name,group,counting_unit,min,max,feeding_mode,trophic_group,source
Cyclopoida,zoop,individual,0.40,1.80,raptorial,omnivore,"a stand-in for a count that stopped at \"Cyclopoid\". Used only where the key\'s own table has no cyclopoids to aggregate. CHECK"
Calanoida,zoop,individual,0.50,2.00,suspension,omnivore,"a stand-in for a count that stopped at \"Calanoid\". CHECK"
Nauplii,zoop,individual,0.08,0.35,filter,herbivore,"copepod nauplii across stages: a wide span because N1 to N6 differ several-fold. CHECK, and prefer a measured value"
Copepodite,zoop,individual,0.15,0.80,suspension,omnivore,"copepodite stages, between nauplius and adult. CHECK"
Protozoa,phyto,cell,20,200,,,"a stand-in for ciliates and flagellates so they appear as food: see the protozoan rows of the trait table for real sizes. CHECK"
Gloeotrichia,phyto,colony,500,2000,,,"Gloeotrichia echinulata colonies are visible to the eye and far larger than the interference threshold, so they are not grazed. CHECK against local measurements"
Eucapsis,phyto,colony,20,60,,,"small cubical colonial cyanobacterium. CHECK"
Micrasterias,phyto,cell,150,300,,,"a large flat desmid. CHECK"
Sphaerozosma,phyto,filament,100,400,,,"filamentous desmid. CHECK"
Peridinium,phyto,cell,30,70,,,"spelt Perindium in some counts. CHECK"
"""


def rows_of(path, sheet):
    wb = load_workbook(Path(path).expanduser(), read_only=True, data_only=True)
    if sheet not in wb.sheetnames:
        return []
    ws = wb[sheet]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    h = [str(c).strip() if c is not None else "" for c in rows[0]]
    return [dict(zip(h, r)) for r in rows[1:] if any(x is not None for x in r)]


def get(d, *names):
    low = {str(k).strip().lower(): v for k, v in d.items() if k}
    for n in names:
        v = low.get(n.lower())
        if v not in (None, ""):
            return v
    return ""


def fnum(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    m = re.search(r"\d+(?:\.\d+)?", str(v).replace(",", "."))
    return float(m.group(0)) if m else None


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--zooplankton", required=True)
    ap.add_argument("--phytoplankton", required=True)
    ap.add_argument("--rules", default="")
    ap.add_argument("--extras", default="",
                    help="a CSV of taxa no published table covers (nauplii, Protozoa, "
                         "Gloeotrichia). Written for you with starting points if absent.")
    ap.add_argument("--aliases", default="",
                    help="the name-difference CSV (rimet_name,phycokey_name,note), so a "
                         "current name such as Dolichospermum finds Anabaena in the table")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()

    rules = dict(DEFAULT_RULES)
    if a.rules and Path(a.rules).expanduser().exists():
        for line in Path(a.rules).expanduser().read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or line.lower().startswith("parameter,"):
                continue
            bits = next(csv.reader([line]))
            if len(bits) >= 2 and fnum(bits[1]) is not None and bits[0].strip() in rules:
                rules[bits[0].strip()] = fnum(bits[1])

    # ---- animals ----
    animals = []
    for r in rows_of(a.zooplankton, "feeding_groups"):
        name = str(get(r, "display_name", "species_id")).strip()
        if not name:
            continue
        lo = fnum(get(r, "body_length_min_mm"))
        hi = fnum(get(r, "body_length_max_mm")) or lo
        if lo is None:
            # fall back to the length mined from the key's own text, as
            # build_food_web.py does: "0.8-2.2" or a single "0.5"
            draft = str(get(r, "body_length_mm_draft"))
            nums = re.findall(r"\d+(?:\.\d+)?", draft.replace(",", "."))
            if nums:
                lo = float(nums[0])
                hi = float(nums[1]) if len(nums) > 1 else lo
        genus = str(get(r, "genus")).strip()
        mesh = "medium"
        for cls, gs in MESH_CLASS.items():
            if genus in gs:
                mesh = cls
        mesh = MESH_SPECIES.get(name, mesh)
        animals.append({
            "n": name, "g": "zoop",
            "lo": lo, "hi": hi, "mesh": mesh,
            "m": str(get(r, "feeding_mode", "feeding_mode_draft")).strip().lower(),
            "t": str(get(r, "trophic_group", "trophic_group_draft")).strip().lower(),
            "r": 1 if str(get(r, "phylum", "Phylum")).lower().startswith("rotifer") else 0,
        })

    # ---- group-level stand-ins, aggregated from the species above ----
    # A count that stops at "Cyclopoid" is common, so give that name the span of the
    # cyclopoids in this table rather than reporting it as unknown.
    rank_rows = defaultdict(list)
    for r in rows_of(a.zooplankton, "feeding_groups"):
        for col in ("order", "suborder", "class", "phylum"):
            v = str(get(r, col)).strip()
            if v:
                rank_rows[v].append(r)
    for rank_name, members in rank_rows.items():
        lens = []
        for r in members:
            lo = fnum(get(r, "body_length_min_mm"))
            hi = fnum(get(r, "body_length_max_mm")) or lo
            if lo is None:
                nums = re.findall(r"\d+(?:\.\d+)?",
                                  str(get(r, "body_length_mm_draft")).replace(",", "."))
                if nums:
                    lo = float(nums[0]); hi = float(nums[1]) if len(nums) > 1 else lo
            if lo:
                lens += [lo, hi]
        if not lens:
            continue
        modes = Counter(str(get(r, "feeding_mode", "feeding_mode_draft")).strip().lower()
                        for r in members if get(r, "feeding_mode", "feeding_mode_draft"))
        groups = Counter(str(get(r, "trophic_group", "trophic_group_draft")).strip().lower()
                         for r in members if get(r, "trophic_group", "trophic_group_draft"))
        meshes = Counter()
        for r in members:
            g_ = str(get(r, "genus")).strip()
            m_ = "medium"
            for cls, gs in MESH_CLASS.items():
                if g_ in gs:
                    m_ = cls
            meshes[MESH_SPECIES.get(str(get(r, "display_name")).strip(), m_)] += 1
        animals.append({
            "n": rank_name, "g": "zoop",
            "lo": min(lens), "hi": max(lens),
            "mesh": meshes.most_common(1)[0][0] if meshes else "medium",
            "m": modes.most_common(1)[0][0] if modes else "",
            "t": groups.most_common(1)[0][0] if groups else "",
            "r": 1 if rank_name.lower().startswith("rotifer") else 0,
            "rank": 1, "ns": len([x for x in members if x]),
        })

    # ---- algae, one entry per counting unit ----
    algae = []
    for r in rows_of(a.phytoplankton, "phyto_traits"):
        role = str(get(r, "role")).strip().lower()
        # Protozoa ARE eaten, by rotifers and by the larger cladocerans, so they belong
        # here as food. They are also consumers, which the page says when one is selected.
        if role not in ("phytoplankton", "uncertain", "protozoan consumer", ""):
            continue      # macroalgae and the anomalous items are not food
        size = fnum(get(r, "max_dimension_um"))
        if size is None:
            continue
        unit = str(get(r, "counting_unit")).strip()
        morph = str(get(r, "morphology_phycokey")).strip().lower()
        algae.append({
            "n": str(get(r, "taxon")).strip(), "g": "phyto", "u": unit or "cell",
            "s": size, "bv": fnum(get(r, "biovolume_um3")),
            "f": 1 if (unit == "filament_100um" or "filament" in morph) else 0,
            "c": 1 if unit == "colony" else 0,
            "gen": str(get(r, "genus")).strip(),
            "gs": 1 if str(get(r, "row_type")).strip() == "genus_summary" else 0,
            "pz": 1 if role == "protozoan consumer" else 0,
            "ns": fnum(get(r, "n_species_in_genus")),
        })

    extras = []
    if a.extras:
        xp = Path(a.extras).expanduser()
        if not xp.exists():
            xp.parent.mkdir(parents=True, exist_ok=True)
            xp.write_text(STARTER_EXTRAS, encoding="utf-8")
            print(f"(Wrote a starter extras file to {xp}: check every line.)")
        for line in xp.read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or line.lower().startswith("name,"):
                continue
            b = next(csv.reader([line]))
            if len(b) < 6 or not b[0].strip():
                continue
            grp = b[1].strip().lower()
            if grp.startswith("zoo"):
                animals.append({"n": b[0].strip(), "g": "zoop",
                                "lo": fnum(b[3]), "hi": fnum(b[4]) or fnum(b[3]),
                                "mesh": "medium", "m": b[5].strip().lower(),
                                "t": (b[6].strip().lower() if len(b) > 6 else ""),
                                "r": 0, "extra": 1,
                                "src": (b[7] if len(b) > 7 else "")})
            else:
                algae.append({"n": b[0].strip(), "g": "phyto", "u": b[2].strip() or "cell",
                              "s": fnum(b[4]) or fnum(b[3]), "bv": None,
                              "f": 1 if "filament" in b[2].lower() else 0,
                              "c": 1 if "colony" in b[2].lower() else 0,
                              "gen": b[0].strip().split()[0], "gs": 0, "ns": None,
                              "extra": 1, "src": (b[7] if len(b) > 7 else "")})
            extras.append(b[0].strip())
    if extras:
        print(f"  extra taxa added from the extras file: {len(extras)}")

    aliases = {}
    if a.aliases and Path(a.aliases).expanduser().exists():
        for line in Path(a.aliases).expanduser().read_text(encoding="utf-8").splitlines():
            line = line.strip()
            if not line or line.startswith("#") or line.lower().startswith("rimet_name"):
                continue
            bits = next(csv.reader([line]))
            if len(bits) >= 2 and bits[0].strip() and bits[1].strip():
                aliases[bits[0].strip().lower()] = bits[1].strip()
                aliases.setdefault(bits[1].strip().lower(), bits[0].strip())
    if aliases:
        print(f"  name aliases loaded: {len(aliases)}")

    payload = {"animals": animals, "algae": algae, "rules": rules, "notes": RULE_NOTES,
               "aliases": aliases, "rank_aliases": {k: v[0] for k, v in RANK_ALIASES.items()}}
    html = PAGE.replace("/*__DATA__*/", json.dumps(payload, separators=(",", ":")))
    outp = Path(a.out).expanduser()
    outp.parent.mkdir(parents=True, exist_ok=True)
    outp.write_text(html, encoding="utf-8")
    size_kb = outp.stat().st_size / 1024
    with_len = sum(1 for x in animals if x["hi"])
    print(f"Wrote {outp}  ({size_kb:.0f} KB)")
    print(f"  animals in the lookup: {len(animals)}  ({with_len} with a body length)")
    print(f"  algal counting units:  {len(algae)}")
    print("\nOpen it in a browser, or publish it beside the counters.")


PAGE = r"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Food Web Dashboard</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css?family=Cabin|News Cycle">
<script src="https://cdn.jsdelivr.net/npm/d3@7.8.5/dist/d3.min.js"></script>
<style>
 :root { --phyto:#2e8b57; --zoop:#F8BD7F; --other:#727171; --ink:#222; --line:#ccc; }
 body { font-family:'Cabin',sans-serif; color:var(--ink); margin:0; background:#fafaf8; }
 header { background:#333; color:#fff; padding:14px 20px; }
 header h1 { font-family:'News Cycle',sans-serif; font-size:26px; margin:0; color:lightgreen; }
 header p { margin:4px 0 0; font-size:14px; color:#ddd; }
 header nav { display:flex; flex-wrap:wrap; gap:6px; margin-top:10px; }
 header nav a { color:#fff; font-size:15px; text-decoration:none; padding:4px 11px;
                background:rgba(255,255,255,0.10); border:1px solid rgba(255,255,255,0.45);
                border-radius:6px; white-space:nowrap; line-height:1.3; }
 header nav a:hover { background:rgba(255,255,255,0.24); }
 header nav a:focus-visible { outline:3px solid #ffbf47; outline-offset:2px; }
 header nav a.fw { background:#009988; border-color:#00bfa8; }
 header nav a.fw:hover { background:#00b3a0; }
 header nav a[aria-current="page"] { background:#fff; color:#00695f; border-color:#fff; font-weight:600; }
 .wrap { display:grid; grid-template-columns:minmax(360px,460px) minmax(0,1fr); gap:20px;
         padding:18px 20px; align-items:start; }
 @media (max-width:1000px){ .wrap { grid-template-columns:1fr; } }
 .panel { background:#fff; border:1px solid var(--line); border-radius:8px; padding:14px; }
 .panel h2 { font-size:16px; text-transform:uppercase; letter-spacing:.06em; margin:0 0 10px;
             background:#333; color:#fff; padding:6px 10px; border-radius:4px; }
 table { border-collapse:collapse; width:100%; font-size:14px; }
 th,td { border:1px solid #ddd; padding:4px 6px; text-align:left; }
 th { background:#f0f0ee; }
 input,select,button { font-family:'Cabin',sans-serif; font-size:14px; }
 input[type=number] { width:76px; }
 .taxon-in { width:100%; min-width:120px; }
 button { cursor:pointer; border:1px solid #444; background:#fff; border-radius:5px; padding:4px 10px; }
 button.primary { background:#1f6f9c; color:#fff; border-color:#1f6f9c; }
 button:focus-visible,input:focus-visible,select:focus-visible { outline:3px solid #ffbf47; outline-offset:1px; }
 .rm { color:#a00; border-color:#a00; padding:2px 7px; }
 #graph { width:100%; height:620px; background:#fff; border:1px solid var(--line); border-radius:8px; }
 .legend span { display:inline-block; margin-right:14px; font-size:14px; }
 .sw { display:inline-block; width:12px; height:12px; border-radius:50%; margin-right:5px;
        vertical-align:middle; box-shadow:0 0 0 1px rgba(0,0,0,0.25); }
 .warn { background:#fff8e1; border-left:4px solid #EE7733; padding:8px 12px; font-size:14px; margin:8px 0; }
 .notelist { max-height:170px; overflow-y:auto; margin:4px 0; }
 .notelist div { padding:1px 0; }
 .muted { color:#666; font-size:13px; }
 .from-counter { background:#eaf4ec; border-left:4px solid #2e8b57; padding:8px 12px;
                 font-size:14px; margin:8px 0; }
 details summary { cursor:pointer; font-weight:bold; margin:6px 0; }
 .stat { font-size:14px; }
 .stat b { font-size:20px; }
 node text { pointer-events:none; }
</style>
</head>
<body>
<header>
  <h1>Food Web Dashboard</h1>
  <p>Load a count from the Food Web Counter, or build one by hand. Edit anything and the web redraws.</p>
  <nav aria-label="Counters and tools">
    <a href="index.html">Cyano Counter</a>
    <a href="Plankton.html">Plankton Counter</a>
    <a href="Taxon.html">Taxon Counter</a>
    <a href="FoodWeb.html" class="fw">Food Web Counter</a>
    <a href="dashboard.html" class="fw" aria-current="page">Food Web Dashboard</a>
  </nav>
</header>

<div class="wrap">
 <div>
  <div class="panel">
    <h2>1. The count</h2>
    <p class="muted">Load a CSV exported by the Food Web Counter, or add taxa by hand. Start typing a
       name for suggestions from the trait tables.</p>
    <input type="file" id="file" accept=".csv">
    <button onclick="addRow()">Add taxon</button>
    <button onclick="loadExample()">Load an example</button>
    <button class="rm" onclick="clearAll()">Clear</button>
    <div id="handoff-note"></div>
    <div id="count-warn"></div>
    <table id="count-table">
      <thead><tr><th>Taxon</th><th>Group</th><th>Unit</th><th>Count</th><th></th></tr></thead>
      <tbody id="count-body"></tbody>
    </table>
    <datalist id="taxon-list"></datalist>
  </div>

  <div class="panel" style="margin-top:16px;">
    <h2>2. The rules</h2>
    <p class="muted">Change a number and the web redraws. The note says how much to trust it.</p>
    <table><tbody id="rule-body"></tbody></table>
    <button onclick="resetRules()" style="margin-top:8px;">Reset to defaults</button>
  </div>

  <div class="panel" style="margin-top:16px;">
    <h2>3. Centrality</h2>
    <p class="muted">Degree is how many links a taxon has. Betweenness is how often it sits on the
       shortest path between two others: a high value means energy routes through it.</p>
    <table><thead><tr><th>Taxon</th><th>Degree</th><th>Betweenness</th></tr></thead>
      <tbody id="cent-body"></tbody></table>
  </div>
 </div>

 <div>
  <div class="panel">
    <h2>The web</h2>
    <div class="legend">
      <span><i class="sw" style="background:var(--phyto)"></i>Phytoplankton</span>
      <span><i class="sw" style="background:var(--zoop)"></i>Zooplankton</span>
      <span><i class="sw" style="background:var(--other)"></i>Other</span>
      <span class="muted">node size = abundance &nbsp;|&nbsp; arrows point from the food to the
      feeder &nbsp;|&nbsp; dashed line = grazing, solid = predation &nbsp;|&nbsp;
      a dashed red ring means no links: select it to see why</span>
    </div>
    <div id="summary" style="margin:8px 0;"></div>
    <svg id="graph" role="img" aria-label="Food web diagram"></svg>
    <div id="detail" class="muted" style="margin-top:8px;">Select a node to see its links.</div>
    <details>
      <summary>What a link means, and what it does not</summary>
      <p class="muted">A link means the rules make feeding <b>physically possible</b>: the food falls
      inside the consumer's size window and is not a filament or a large colony. It does not mean the
      interaction has been observed. Size rules over-predict, so treat the web as a set of candidate
      links to check, not a finding. A taxon with no body length in the trait table cannot be linked
      at all and is drawn without edges.</p>
    </details>
  </div>
 </div>
</div>

<script>
const DATA = /*__DATA__*/;
let rules = Object.assign({}, DATA.rules);
let rows  = [];          // { taxon, group, unit, count }
let nextRow = 1;

const byName = {};
DATA.animals.forEach(a => { byName[a.n.toLowerCase()] = a; });
const algaeByName = {}, algaeByGenus = {};
DATA.algae.forEach(p => {
  const k = p.n.toLowerCase();
  (algaeByName[k] = algaeByName[k] || []).push(p);
  if (p.gen) { const g = p.gen.toLowerCase();
               (algaeByGenus[g] = algaeByGenus[g] || []).push(p); }
});
const animalsByGenus = {};
DATA.animals.forEach(a => {
  const g = (a.n || '').trim().split(/\s+/)[0].toLowerCase();
  if (g) (animalsByGenus[g] = animalsByGenus[g] || []).push(a);
});

// A counter name often carries the unit in it ("Synura Colony") and is often just
// the genus ("Dinobryon"). Both are handled rather than reported as unknown.
const UNIT_WORDS = { colony:'colony', colonies:'colony', filament:'filament_100um',
                     filaments:'filament_100um', cell:'cell', cells:'cell',
                     'single cell':'cell', trichome:'filament_100um' };

function splitUnitFromName(raw) {
  let name = (raw || '').trim(), unit = null;
  const m = name.match(/[\s,(]+(single cell|colonies|colony|filaments?|cells?|trichome)s?\)?$/i);
  if (m) {
    unit = UNIT_WORDS[m[1].toLowerCase()] || null;
    name = name.slice(0, m.index).trim();
  }
  return { name, unit };
}

// name completion
(function fillList(){
  const dl = document.getElementById('taxon-list');
  const names = [...new Set([...DATA.animals.map(a=>a.n), ...DATA.algae.map(p=>p.n)])].sort();
  dl.innerHTML = names.slice(0, 4000).map(n=>`<option value="${n.replace(/"/g,'&quot;')}">`).join('');
})();

// ─── the count table ──────────────────────────────────────────────────────────
function addRow(taxon, group, unit, count) {
  const id = nextRow++;
  rows.push({ id, taxon: taxon||'', group: group||'phyto', unit: unit||'cell', count: count||0 });
  const tr = document.createElement('tr');
  tr.id = 'r'+id;
  tr.innerHTML = `
    <td><input class="taxon-in" list="taxon-list" value="${(taxon||'').replace(/"/g,'&quot;')}"
         oninput="setField(${id},'taxon',this.value)" aria-label="Taxon name"></td>
    <td><select onchange="setField(${id},'group',this.value)" aria-label="Group">
          <option value="phyto">Phyto</option><option value="zoop">Zoop</option>
          <option value="other">Other</option></select></td>
    <td><select onchange="setField(${id},'unit',this.value)" aria-label="Counting unit" id="u${id}"></select></td>
    <td><input type="number" min="0" value="${count||0}"
         oninput="setField(${id},'count',this.value)" aria-label="Count"></td>
    <td><button class="rm" onclick="removeRow(${id})" aria-label="Remove">&#10005;</button></td>`;
  document.getElementById('count-body').appendChild(tr);
  tr.querySelector('select').value = group || 'phyto';
  renderUnits(id);
  redraw();
}

function renderUnits(id) {
  const r = rows.find(r=>r.id===id); if (!r) return;
  const sel = document.getElementById('u'+id);
  const opts = r.group === 'phyto' ? ['cell','colony','filament_100um'] : ['individual'];
  sel.innerHTML = opts.map(o=>`<option value="${o}">${o==='filament_100um'?'filament':o}</option>`).join('');
  if (!opts.includes(r.unit)) r.unit = opts[0];
  sel.value = r.unit;
}

function setField(id, field, value) {
  const r = rows.find(r=>r.id===id); if (!r) return;
  r[field] = field === 'count' ? Math.max(0, parseFloat(value)||0) : value;
  if (field === 'group') renderUnits(id);
  redraw();
}

function removeRow(id) {
  rows = rows.filter(r=>r.id!==id);
  document.getElementById('r'+id)?.remove();
  redraw();
}

function clearAll() {
  rows = []; document.getElementById('count-body').innerHTML = ''; redraw();
}

function loadExample() {
  clearAll();
  [['Microcystis aeruginosa','phyto','colony',40],
   ['Microcystis aeruginosa (single cell)','phyto','cell',120],
   ['Cyclotella','phyto','cell',200],
   ['Asterionella formosa','phyto','cell',150],
   ['Daphnia pulex','zoop','individual',18],
   ['Bosmina longirostris','zoop','individual',44],
   ['Keratella cochlearis','zoop','individual',60],
   ['Mesocyclops edax','zoop','individual',7],
   ['Leptodora kindtii','zoop','individual',2]].forEach(r=>addRow(...r));
}

// ─── the rules panel ──────────────────────────────────────────────────────────
function renderRules() {
  document.getElementById('rule-body').innerHTML = Object.keys(rules).map(k=>`
    <tr><td><label for="rule-${k}">${k.replace(/_/g,' ')}</label></td>
        <td><input type="number" step="any" id="rule-${k}" value="${rules[k]}"
             oninput="setRule('${k}',this.value)"></td>
        <td class="muted">${DATA.notes[k]||''}</td></tr>`).join('');
}
function setRule(k, v) { const n = parseFloat(v); if (!isNaN(n)) { rules[k] = n; redraw(); } }
function resetRules() { rules = Object.assign({}, DATA.rules); renderRules(); redraw(); }

// ─── the rules themselves, matching build_food_web.py ─────────────────────────
function lookupAnimal(name) { return byName[(name||'').trim().toLowerCase()] || null; }

function pickUnit(list, unit) {
  return list.find(p => p.u === unit) || list.find(p => p.u === 'cell') || list[0];
}

// Returns { trait, how } so the page can say HOW a name was matched.
function lookupAlga(rawName, unit) {
  const exact = algaeByName[(rawName||'').trim().toLowerCase()];
  if (exact) return { trait: pickUnit(exact, unit), how: 'name' };

  const { name, unit: unitFromName } = splitUnitFromName(rawName);
  const u = unitFromName || unit;
  const stripped = algaeByName[name.toLowerCase()];
  if (stripped) return { trait: pickUnit(stripped, u), how: 'name without the unit word' };

  // the genus: prefer the genus median row, which carries the spread across species
  // a different name for the same alga (Dolichospermum / Anabaena)
  const alt = aliasOf(name.split(/\s+/)[0]);
  if (alt) {
    const byAlt = algaeByGenus[alt.toLowerCase()];
    if (byAlt) {
      const med = byAlt.filter(p => p.gs);
      const hit = pickUnit(med.length ? med : byAlt, u);
      if (hit) return { trait: hit, how: `the name ${alt}, which the trait table uses` };
    }
  }

  const g = algaeByGenus[name.toLowerCase().split(/\s+/)[0]];
  if (g) {
    const medians = g.filter(p => p.gs);
    const pool = medians.length ? medians : g;
    const hit = pickUnit(pool, u);
    if (hit) return { trait: hit,
        how: `genus median${hit.ns ? ' of ' + hit.ns + ' species' : ''}` };
  }
  return { trait: null, how: null };
}

const ALIASES = DATA.aliases || {};
const RANK_ALIASES = DATA.rank_aliases || {};
function aliasOf(name) {
  const a = ALIASES[(name||'').trim().toLowerCase()];
  return a && a.toLowerCase() !== (name||'').trim().toLowerCase() ? a : null;
}

// ─── "did you mean" ───────────────────────────────────────────────────────────
// Counter sheets carry typos (Wornichinia), group names (Protozoa), life stages
// (nauplii) and stray keystrokes. Reporting "not in the table" is true but useless,
// so look for a near match and offer it.
const ALL_NAMES = (() => {
  const set = new Set();
  DATA.algae.forEach(p => { if (p.gen) set.add(p.gen); });
  DATA.algae.forEach(p => set.add(p.n.replace(/\s*\(genus median\)$/, '')));
  DATA.animals.forEach(a => { set.add(a.n); const g = a.n.split(/\s+/)[0]; if (g) set.add(g); });
  return [...set].filter(n => n && n.length > 2);
})();

// Words that are a group or a life stage rather than a taxon in the tables.
const NOT_A_TAXON = {
  'protozoa': 'a group, not a taxon: count the genus, or set the group to Other',
  'protozoan': 'a group, not a taxon',
  'ciliate': 'a group of protozoa: set the group to Other',
  'ciliates': 'a group of protozoa: set the group to Other',
  'diatom': 'a group: count the genus',
  'diatoms': 'a group: count the genus',
  'green algae': 'a group: count the genus',
  'detritus': 'not an organism',
  'unknown': 'not a taxon',
};

function editDistance(a, b) {
  const m = a.length, n = b.length;
  if (!m || !n) return Math.max(m, n);
  let prev = Array.from({length: n + 1}, (_, j) => j), cur = new Array(n + 1);
  for (let i = 1; i <= m; i++) {
    cur[0] = i;
    for (let j = 1; j <= n; j++) {
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1,
                        prev[j - 1] + (a[i-1] === b[j-1] ? 0 : 1));
    }
    [prev, cur] = [cur, prev];
  }
  return prev[n];
}

function didYouMean(raw) {
  const name = (raw || '').trim();
  const key = name.toLowerCase();
  if (NOT_A_TAXON[key]) return { kind: 'explain', text: NOT_A_TAXON[key] };
  if (name.length < 4 || !/[a-z]{3}/i.test(name)) return null;
  let best = null, bestD = Infinity;
  for (const cand of ALL_NAMES) {
    const c = cand.toLowerCase();
    if (Math.abs(c.length - key.length) > 4) continue;
    const d = editDistance(key, c);
    if (d < bestD) { bestD = d; best = cand; if (d === 0) break; }
  }
  if (!best) return null;
  const sim = 1 - bestD / Math.max(key.length, best.length);
  return sim >= 0.72 ? { kind: 'suggest', name: best, sim } : null;
}

function lookupAnimalLoose(rawName) {
  const a = lookupAnimal(rawName);
  if (a) return { trait: a, how: a.rank ? `the ${a.n} group, spanning ${a.ns} taxa here`
                                        : (a.extra ? 'the extras file' : 'name') };
  // "Cyclopoid", "Calanoid", "Nauplii": the rank a counter stopped at
  const key = (rawName || '').trim().toLowerCase();
  const rank = RANK_ALIASES[key];
  if (rank) {
    const hit = lookupAnimal(rank);
    if (hit) return { trait: hit,
        how: `the ${rank} group, spanning ${hit.ns || '?'} taxa in the trait table` };
  }
  // a name the alias list maps to something we do have (misspellings included)
  const alt = aliasOf(rawName);
  if (alt) {
    const hit = lookupAnimal(alt) || (RANK_ALIASES[alt.toLowerCase()]
                                      ? lookupAnimal(RANK_ALIASES[alt.toLowerCase()]) : null);
    if (hit) return { trait: hit, how: `the name ${alt}` };
  }
  const genus = (rawName||'').trim().split(/\s+/)[0].toLowerCase();
  const list = animalsByGenus[genus];
  if (list && list.length) {
    const withLen = list.filter(x => x.hi);
    const pool = withLen.length ? withLen : list;
    const lo = Math.min(...pool.map(x => x.lo || x.hi)),
          hi = Math.max(...pool.map(x => x.hi || x.lo));
    return { trait: Object.assign({}, pool[0], { lo, hi, n: rawName }),
             how: `genus, spanning ${pool.length} species` };
  }
  return { trait: null, how: null };
}

function buildLinks() {
  const nodes = [], notes = [], matched = [];
  rows.forEach(r => {
    if (!r.taxon.trim()) return;
    if (r.group === 'zoop') {
      const { trait: a, how } = lookupAnimalLoose(r.taxon);
      if (!a) notes.push({ row: r.id, taxon: r.taxon,
                           msg: 'not in the zooplankton trait table',
                           hint: didYouMean(r.taxon) });
      else if (how !== 'name') matched.push(`${r.taxon}: matched by ${how}`);
      if (a && !a.hi) notes.push({ row: r.id, taxon: r.taxon,
                                  msg: 'no body length, so it cannot be linked', hint: null });
      nodes.push({ id:r.taxon, group:'zoop', count:r.count, unit:r.unit, trait:a });
    } else if (r.group === 'phyto') {
      const { trait: p, how } = lookupAlga(r.taxon, r.unit);
      if (!p) notes.push({ row: r.id, taxon: r.taxon,
                           msg: 'not in the phytoplankton trait table',
                           hint: didYouMean(r.taxon) });
      else if (how !== 'name') matched.push(`${r.taxon}: matched by ${how}`);
      const shownUnit = p ? p.u : r.unit;
      nodes.push({ id:`${r.taxon} (${shownUnit==='filament_100um'?'filament':shownUnit})`,
                   group:'phyto', count:r.count, unit:shownUnit, trait:p });
    } else {
      nodes.push({ id:r.taxon, group:'other', count:r.count, unit:r.unit, trait:null });
      notes.push({ row:r.id, taxon:r.taxon,
                   msg:'"other" taxa have no traits yet, so they are drawn unlinked', hint:null });
    }
  });

  const links = [];
  nodes.forEach(cons => {
    const a = cons.trait;
    if (cons.group !== 'zoop' || !a) return;
    const L = a.hi || a.lo;
    const herb = (a.m === 'filter' || a.m === 'suspension' ||
                  a.t === 'herbivore' || a.t === 'omnivore');
    if (herb) {
      let upper = null, why = '';
      if (a.r) { upper = rules.rotifer_max_particle_um; why = `microphagous rotifer, up to ${upper} um`; }
      else if (L) { upper = rules.burns_slope*L + rules.burns_intercept;
                    why = `Burns 1968: 22 x ${L} mm + 4.87 = ${upper.toFixed(0)} um`; }
      if (upper !== null) {
        const lower = rules['mesh_'+(a.mesh||'default')+'_um'] ?? rules.mesh_default_um;
        const up = (a.m === 'raptorial') ? upper*rules.raptorial_particle_multiplier : upper;
        nodes.forEach(res => {
          if (res.group !== 'phyto' || !res.trait) return;
          const s = res.trait.s;
          if (s < lower || s > up) return;
          if (res.trait.f) return;                                   // a filament fouls the filter
          if (res.trait.c && s >= rules.interference_length_um) return;  // too big a colony
          links.push({ source:cons.id, target:res.id, type:'grazing',
                       why:`${why}; mesh ${lower} um` });
        });
      }
    }
    // Predation needs a RAPTORIAL mode. An omnivorous filter feeder such as Daphnia
    // takes protists and bacteria, not other crustaceans, so the trophic group alone
    // must not hand it the predator-prey rule.
    const carn = (a.m === 'raptorial' || a.m === 'parasitic' ||
                  ((a.t === 'carnivore' || a.t === 'parasite') && a.m !== 'filter'
                   && a.m !== 'suspension' && a.m !== 'scraper'));
    if (carn && L) {
      const lo = rules.prey_ratio_min*L, hi = rules.prey_ratio_max*L;
      nodes.forEach(res => {
        if (res.group !== 'zoop' || res.id === cons.id || !res.trait) return;
        const pl = res.trait.hi || res.trait.lo;
        if (pl && pl >= lo && pl <= hi) {
          links.push({ source:cons.id, target:res.id, type:'predation',
                       why:`prey ${lo.toFixed(2)}-${hi.toFixed(2)} mm of predator length` });
        }
      });
    }
  });
  return { nodes, links, notes, matched };
}

// ─── centrality ───────────────────────────────────────────────────────────────
function centrality(nodes, links) {
  const idx = {}; nodes.forEach((n,i)=>idx[n.id]=i);
  const adj = nodes.map(()=>[]);
  links.forEach(l => {
    const s = idx[l.source], t = idx[l.target];
    if (s === undefined || t === undefined) return;
    if (!adj[s].includes(t)) adj[s].push(t);
    if (!adj[t].includes(s)) adj[t].push(s);
  });
  const degree = adj.map(a=>a.length);
  // Brandes, on the undirected graph
  const bc = nodes.map(()=>0);
  for (let s=0; s<nodes.length; s++) {
    const S=[], P=nodes.map(()=>[]), sigma=nodes.map(()=>0), d=nodes.map(()=>-1);
    sigma[s]=1; d[s]=0; const Q=[s];
    while (Q.length) {
      const v=Q.shift(); S.push(v);
      adj[v].forEach(w=>{
        if (d[w]<0) { d[w]=d[v]+1; Q.push(w); }
        if (d[w]===d[v]+1) { sigma[w]+=sigma[v]; P[w].push(v); }
      });
    }
    const delta=nodes.map(()=>0);
    while (S.length) {
      const w=S.pop();
      P[w].forEach(v=>{ delta[v]+= (sigma[v]/sigma[w])*(1+delta[w]); });
      if (w!==s) bc[w]+=delta[w];
    }
  }
  const n = nodes.length;
  const norm = n>2 ? 2/((n-1)*(n-2)) : 1;     // undirected, so halve
  return nodes.map((nd,i)=>({ id:nd.id, group:nd.group,
      degree:degree[i], betweenness:+(bc[i]*norm/2).toFixed(3) }));
}

// ─── drawing ──────────────────────────────────────────────────────────────────
let sim = null;
function redraw() {
  const { nodes, links, notes, matched } = buildLinks();
  const cent = centrality(nodes, links);
  const centById = {}; cent.forEach(c=>centById[c.id]=c);

  const possible = nodes.length*(nodes.length-1);
  document.getElementById('summary').innerHTML =
    `<span class="stat"><b>${nodes.length}</b> taxa &nbsp; <b>${links.length}</b> links &nbsp; ` +
    `connectance <b>${possible? (links.length/possible).toFixed(3):'0'}</b></span>`;

  const esc = t => String(t).replace(/[&<>"]/g, ch =>
      ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[ch]));
  const noteLine = n => {
    let extra = '';
    if (n.hint && n.hint.kind === 'suggest') {
      extra = ` &mdash; did you mean <b>${esc(n.hint.name)}</b>? ` +
              `<button style="font-size:12px;padding:1px 7px;" ` +
              `onclick="applySuggestion(${n.row}, '${esc(n.hint.name).replace(/'/g, "\\'")}')">Use</button>`;
    } else if (n.hint && n.hint.kind === 'explain') {
      extra = ` &mdash; <span class="muted">${esc(n.hint.text)}</span>`;
    }
    return `<div><b>${esc(n.taxon)}</b>: ${esc(n.msg)}${extra}</div>`;
  };
  const listBox = (items, cls, heading, footer, render) => items.length ? `<div class="${cls}">` +
      `<b>${heading} (${items.length})</b>` +
      `<button onclick="copyNotes('${cls}')" style="float:right;font-size:12px;padding:2px 8px;">Copy list</button>` +
      `<div class="notelist" id="list-${cls}">` + items.map(render).join('') + `</div>` +
      (footer ? `<span class="muted">${footer}</span>` : '') + `</div>` : '';
  document.getElementById('count-warn').innerHTML =
    listBox(notes, 'warn', 'Not found or unusable', '', noteLine) +
    listBox(matched, 'from-counter', 'Matched less exactly',
            'A genus value is the median across the species in the trait table, so it is a '
            + 'rough stand-in for a species measurement.',
            m => `<div>${esc(m)}</div>`);

  document.getElementById('cent-body').innerHTML = cent
    .sort((a,b)=>b.degree-a.degree || b.betweenness-a.betweenness)
    .map(c=>`<tr><td>${c.id.replace(/</g,'&lt;')}</td><td>${c.degree}</td><td>${c.betweenness}</td></tr>`)
    .join('') || '<tr><td colspan="3" class="muted">No taxa yet.</td></tr>';

  const svg = d3.select('#graph');
  svg.selectAll('*').remove();
  if (!nodes.length) return;
  const box = document.getElementById('graph').getBoundingClientRect();
  const W = box.width, H = box.height;
  svg.attr('viewBox', `0 0 ${W} ${H}`);

  nodes.forEach(n => { n.degree = centById[n.id] ? centById[n.id].degree : 0; });
  const maxCount = d3.max(nodes, n=>n.count) || 1;
  const rOf = n => 6 + 18*Math.sqrt((n.count||0)/maxCount);
  const colour = g => g==='phyto' ? '#2e8b57' : g==='zoop' ? '#F8BD7F' : '#727171';

  // arrowheads: the line runs from the PREY to the PREDATOR, the way energy moves
  const defs = svg.append('defs');
  [['arrow-graz', '#999'], ['arrow-pred', '#555']].forEach(([id, fill]) => {
    defs.append('marker')
        .attr('id', id).attr('viewBox', '0 -5 10 10')
        .attr('refX', 9).attr('refY', 0)
        .attr('markerWidth', 6).attr('markerHeight', 6)
        .attr('orient', 'auto-start-reverse')
      .append('path')
        .attr('d', 'M0,-4L9,0L0,4')
        .attr('fill', fill);
  });

  const link = svg.append('g').attr('stroke','#999').attr('stroke-opacity',0.55)
    .selectAll('line').data(links).join('line')
      .attr('stroke', d=>d.type==='predation'?'#555':'#999')
      .attr('stroke-width', d=>d.type==='predation'?2:1)
      .attr('stroke-dasharray', d=>d.type==='predation'?'':'3,2')
      .attr('marker-end', d=>`url(#${d.type==='predation'?'arrow-pred':'arrow-graz'})`);

  const node = svg.append('g').selectAll('g').data(nodes).join('g')
      .attr('tabindex', 0).style('cursor','pointer')
      .on('click', (e,d)=>showDetail(d, links, centById, nodes))
      .on('keydown', (e,d)=>{ if(e.key==='Enter') showDetail(d, links, centById, nodes); })
      .call(d3.drag()
        .on('start', (e,d)=>{ if(!e.active) sim.alphaTarget(0.25).restart(); d.fx=d.x; d.fy=d.y; })
        .on('drag',  (e,d)=>{ d.fx=e.x; d.fy=e.y; })
        .on('end',   (e,d)=>{ if(!e.active) sim.alphaTarget(0); d.fx=null; d.fy=null; }));
  node.append('circle').attr('r', rOf).attr('fill', d=>colour(d.group))
      .attr('stroke', d=> d.degree ? (d.group==='zoop' ? '#b07a3c' : '#fff') : '#c0392b')
      .attr('stroke-width', d=> d.degree ? 1.5 : 2)
      .attr('stroke-dasharray', d=> d.degree ? '' : '3,2')
      .append('title').text(d=>`${d.id}\ncount ${d.count}\ndegree ${centById[d.id]?.degree ?? 0}`);
  node.append('text').text(d=>d.id).attr('font-size',11).attr('dx', d=>rOf(d)+3).attr('dy',4);

  sim = d3.forceSimulation(nodes)
    .force('link', d3.forceLink(links).id(d=>d.id).distance(90).strength(0.35))
    .force('charge', d3.forceManyBody().strength(d => d.degree ? -260 : -90))
    .force('x', d3.forceX(W/2).strength(0.06))
    .force('y', d3.forceY(H/2).strength(0.06))
    .force('center', d3.forceCenter(W/2, H/2))
    .force('collide', d3.forceCollide().radius(d=>rOf(d)+14))
    .on('tick', () => {
      // keep every node inside the panel, labels included, so nothing is clipped
      nodes.forEach(d => {
        const r = rOf(d) + 4;
        const labelRoom = Math.min(8 * d.id.length + 20, W / 3);
        d.x = Math.max(r, Math.min(W - r - labelRoom, d.x));
        d.y = Math.max(r, Math.min(H - r, d.y));
      });
      // d.source is the consumer and d.target the food, so the line is drawn the
      // other way round and stops at the edge of the consumer's circle, leaving
      // the arrowhead visible rather than buried under the node.
      link.each(function (d) {
        const dx = d.source.x - d.target.x, dy = d.source.y - d.target.y;
        const len = Math.hypot(dx, dy) || 1;
        const stop = rOf(d.source) + 7;
        this.setAttribute('x1', d.target.x);
        this.setAttribute('y1', d.target.y);
        this.setAttribute('x2', d.source.x - (dx / len) * stop);
        this.setAttribute('y2', d.source.y - (dy / len) * stop);
      });
      node.attr('transform', d=>`translate(${d.x},${d.y})`);
    });
}

function applySuggestion(rowId, name) {
  const r = rows.find(r => r.id === rowId);
  if (!r) return;
  r.taxon = name;
  const input = document.querySelector(`#r${rowId} input.taxon-in`);
  if (input) input.value = name;
  // the corrected name may belong to the other group
  const grpSel = document.querySelector(`#r${rowId} select`);
  if (lookupAnimalLoose(name).trait && !lookupAlga(name, 'cell').trait) {
    r.group = 'zoop'; if (grpSel) grpSel.value = 'zoop'; renderUnits(rowId);
  }
  redraw();
}

function copyNotes(cls) {
  const box = document.getElementById('list-' + cls);
  if (!box) return;
  const text = [...box.querySelectorAll('div')].map(d => d.textContent).join('\n');
  const done = () => { const b = box.parentElement.querySelector('button');
                       const was = b.textContent; b.textContent = 'Copied';
                       setTimeout(() => { b.textContent = was; }, 1500); };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(text).then(done, () => fallbackCopy(text, done));
  } else { fallbackCopy(text, done); }
}

function fallbackCopy(text, done) {
  const ta = document.createElement('textarea');
  ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
  document.body.appendChild(ta); ta.select();
  try { document.execCommand('copy'); done(); } catch (e) { window.prompt('Copy this list:', text); }
  document.body.removeChild(ta);
}

// Why does this node have no links? Silence is the least informative thing a
// diagram can do, so say which rule stopped it.
function whyNoLinks(d, nodes) {
  const t = d.trait;
  if (!t) return 'It is not in the trait table, so no rule can be applied to it. '
              + 'Check the spelling, or add it to the local additions file.';
  if (d.group === 'zoop') {
    const L = t.hi || t.lo;
    if (!L) return 'It has <b>no body length</b> in the trait table, and every rule needs one. '
                 + 'Fill in body_length_min_mm and body_length_max_mm in the feeding_groups sheet, '
                 + 'then rebuild.';
    const herb = (t.m === 'filter' || t.m === 'suspension' || t.t === 'herbivore' || t.t === 'omnivore');
    if (!herb && t.m !== 'raptorial' && t.t !== 'carnivore')
      return `Its feeding mode is "${t.m || 'unknown'}", which no rule covers yet.`;
    const upper = t.r ? rules.rotifer_max_particle_um
                      : rules.burns_slope * L + rules.burns_intercept;
    const lower = rules['mesh_' + (t.mesh || 'default') + '_um'] ?? rules.mesh_default_um;
    const algae = nodes.filter(n => n.group === 'phyto');
    if (!algae.length) return 'There is no phytoplankton in this count for it to eat.';
    const sized = algae.filter(n => n.trait);
    const tooBig = sized.filter(n => n.trait.s > upper);
    const blocked = sized.filter(n => n.trait.f ||
                                 (n.trait.c && n.trait.s >= rules.interference_length_um));
    let out = `Its window is <b>${lower} to ${upper.toFixed(0)} um</b>` +
              (t.r ? ' (the fixed rotifer limit)' : ` (Burns: 22 x ${L} mm + 4.87)`) + '. ';
    if (blocked.length) out += `${blocked.length} of the algae here are filaments or large `
                             + `colonies, so they foul the filter rather than being eaten. `;
    if (tooBig.length) out += `${tooBig.length} are larger than ${upper.toFixed(0)} um. `;
    if (!sized.length) out += 'None of the algae here have a size in the trait table. ';
    return out;
  }
  if (d.group === 'phyto') {
    if (t.pz) return 'It is a protozoan. It is available as food here, but protozoa are also '
                   + 'consumers, which this model does not yet represent: no links out of it '
                   + 'are drawn. Its size is ' + t.s + ' um.';
    if (t.f) return 'It is a <b>filament</b>, so every grazing link is withheld: filaments foul '
                  + 'the filtering combs rather than being eaten. Counting it as cells instead '
                  + 'would make it food.';
    if (t.c && t.s >= rules.interference_length_um)
      return `It is a <b>colony of ${t.s} um</b>, at or above the interference threshold of `
           + `${rules.interference_length_um} um, so it is too large to be filtered. The same `
           + `species counted as single cells would be food.`;
    return `It is ${t.s} um across. No grazer in this count has a window that includes that size.`;
  }
  return 'Taxa in the "other" group have no traits yet.';
}

function showDetail(d, links, centById, nodes) {
  const eats = links.filter(l=>(l.source.id||l.source)===d.id);
  const eaten = links.filter(l=>(l.target.id||l.target)===d.id);
  const name = s => s.id || s;
  const none = !eats.length && !eaten.length;
  document.getElementById('detail').innerHTML =
    `<b>${d.id}</b> &mdash; count ${d.count}, degree ${centById[d.id]?.degree ?? 0}, ` +
    `betweenness ${centById[d.id]?.betweenness ?? 0}<br>` +
    (eats.length ? `<b>Eats:</b> ${eats.map(l=>name(l.target)).join(', ')}<br>` : '') +
    (eaten.length ? `<b>Eaten by:</b> ${eaten.map(l=>name(l.source)).join(', ')}<br>` : '') +
    (d.trait && d.trait.pz ? `<span class="muted">A protozoan: eaten here, but also a consumer, `
      + `which the model does not yet represent.</span><br>` : '') +
    (eats[0] ? `<span class="muted">Rule: ${eats[0].why}</span>` : '') +
    (none ? `<div class="warn" style="margin-top:6px;"><b>No links.</b> ` +
            `${whyNoLinks(d, nodes)}</div>` : '');
}

// ─── loading a counter CSV ────────────────────────────────────────────────────
document.getElementById('file').addEventListener('change', ev => {
  const f = ev.target.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = () => { parseCsv(rd.result); };
  rd.readAsText(f);
});

function parseCsv(text) {
  const lines = text.split(/\r?\n/);
  let head = -1;
  for (let i=0;i<lines.length;i++) {
    if (/(^|,)"?taxon( name)?"?(,|$)/i.test(lines[i])) { head = i; break; }
  }
  if (head < 0) { alert('No "taxon" column found. Export from the Food Web Counter, or add taxa by hand.'); return; }
  const cols = splitCsvLine(lines[head]).map(c=>c.toLowerCase().trim());
  const ci = n => {
    let i = cols.indexOf(n);
    if (i < 0 && n === 'taxon') i = cols.indexOf('taxon name');   // older exports
    if (i < 0 && n === 'count') i = cols.findIndex(c => c.startsWith('count'));
    return i;
  };
  clearAll();
  let added = 0;
  for (let i=head+1;i<lines.length;i++) {
    if (!lines[i].trim()) continue;
    const c = splitCsvLine(lines[i]);
    const taxon = (c[ci('taxon')]||'').trim();
    if (!taxon) continue;
    const grpRaw = (c[ci('group')]||'').toLowerCase();
    // With no group column, work it out from the trait tables rather than assuming
    // phytoplankton: a bare genus like "Keratella" is a rotifer.
    const guess = () => {
      const { name } = splitUnitFromName(taxon);
      if (lookupAnimalLoose(name).trait) return 'zoop';
      if (lookupAlga(taxon, 'cell').trait) return 'phyto';
      return 'phyto';
    };
    const group = grpRaw.startsWith('zoo') ? 'zoop' : grpRaw.startsWith('phyt') ? 'phyto'
                : grpRaw ? 'other' : guess();
    let unit = (c[ci('counting_unit')]||'').trim().toLowerCase() || (group==='phyto'?'cell':'individual');
    if (unit === 'filament') unit = 'filament_100um';
    addRow(taxon, group, unit, parseFloat(c[ci('count')]) || 0);
    added++;
  }
  if (!added) alert('No taxon rows were found in that file.');
}

function splitCsvLine(line) {
  const out=[]; let cur='', q=false;
  for (let i=0;i<line.length;i++) {
    const ch=line[i];
    if (q) { if (ch==='"' && line[i+1]==='"') { cur+='"'; i++; } else if (ch==='"') q=false; else cur+=ch; }
    else if (ch==='"') q=true;
    else if (ch===',') { out.push(cur); cur=''; }
    else cur+=ch;
  }
  out.push(cur); return out;
}

// ─── a count handed over by the Food Web Counter on this same site ──────────────────
const HANDOFF_KEY = 'cfb.foodweb.count';

function loadHandoff() {
  let raw = null;
  try { raw = localStorage.getItem(HANDOFF_KEY); } catch (e) { return false; }
  if (!raw) return false;
  let p;
  try { p = JSON.parse(raw); } catch (e) { return false; }
  if (!p || !Array.isArray(p.rows) || !p.rows.length) return false;
  clearAll();
  p.rows.forEach(r => addRow(r.taxon, r.group, r.unit, r.count));
  const who = [p.meta && p.meta.name, p.meta && p.meta.location, p.meta && p.meta.date]
              .filter(Boolean).join(', ');
  document.getElementById('handoff-note').innerHTML =
    `<div class="from-counter">Loaded <b>${p.rows.length}</b> taxa from the Food Web Counter` +
    (who ? ` (${who.replace(/</g,'&lt;')})` : '') +
    `. Edit anything below; the counter is not changed.` +
    ` <button onclick="forgetHandoff()" style="margin-left:6px;">Start fresh</button></div>`;
  return true;
}

function forgetHandoff() {
  try { localStorage.removeItem(HANDOFF_KEY); } catch (e) {}
  document.getElementById('handoff-note').innerHTML = '';
  clearAll();
}

renderRules();
if (!loadHandoff()) redraw();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
