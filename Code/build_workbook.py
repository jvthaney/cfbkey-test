#!/usr/bin/env python3
"""
build_workbook.py

Builds the DRAFT editable workbook for the new CFB Key from the CSV files made by
inventory_site.py plus the video manifest made by upload_to_ia.py.

  python build_workbook.py --inventory "$ZK/inventory" \\
         --manifest "$ZK/video_manifest.csv" --out "$ZK/cfbkey_workbook.xlsx"

It only READS the CSVs. Re-running it makes a fresh file, so once people start
editing the workbook, copy it under a new name first (it will not be overwritten
unless --out points at the same name).

SHEETS
  README          how to use the workbook + a live progress tracker
  species         one row per species (taxonomy, ids, counts)
  species_sections  the text of each species page, one row per bullet/paragraph
  image_maps      pictures with clickable regions, one row per region
  images          every picture on a species page, with a column for NEW alt text
  figures         the popup "detail figure" links, with a column for NEW alt text
  videos          every clip: type, old descriptions, archive.org URL, NEW description
  species_videos  which species page shows which clip
  other_video_links  clips shown on key steps, methods and other pages
  key_steps       the identification key, one row per choice
  anatomy         the anatomy diagrams: one row per label (write alt_text_new for each view)
  groups          the Groups outline (group names, where each leads)
  key_steps_info  one row per key step: prompt, taxonomy so far, optional new heading
  key_help_links  help/method links on key steps
  key_images_of_text  non-choice images on key steps (often words saved as pictures)
  taxonomy        every rank value with counts; near-duplicates flagged
  extra_pages     supplementary pages (figures, data, movies) attached to a species
  pages           the remaining pages (about, history, anatomy ...)
  fixes           known broken links and the proposed correction
  to_do           problems the inventory could not resolve by itself

YELLOW columns are for people to fill in. Everything else is generated.
"""

import argparse
import csv
import difflib
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from pathlib import Path, PurePosixPath

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

FONT = "Arial"
HEAD_FILL = PatternFill("solid", fgColor="1F3864")
INPUT_HEAD_FILL = PatternFill("solid", fgColor="FFD966")
INPUT_CELL_FILL = PatternFill("solid", fgColor="FFF2CC")
RANKS = ["Phylum", "Subphylum", "Class", "Subclass", "Order", "Suborder", "Family"]

# Known broken links -> proposed correction. EDIT after Amanda confirms.
KNOWN_FIXES = [
    {"kind": "video link", "old_value": "media/chonochilus1/chonochilus1.htm",
     "new_value": "media/conochilus2/conochilus2",
     "note": "Folder name is misspelled on the old site. The clip description matches "
             "conochilus2 ('Movement of individual Conochilus within a colony'). CONFIRM WITH AMANDA."},
]
VIDEO_TYPES = {
    "organisms videos": "Live organism",
    "virtual focus": "Virtual Focus",
    "sampling methods": "Sampling method",
    "dissection demonstration": "Dissection demonstration",
}


def read(inv: Path, name: str):
    f = inv / name
    if not f.exists():
        print(f"  (missing {name}: that sheet will be empty)")
        return []
    with f.open(encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def slug(s: str) -> str:
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    s = s.replace("(", " ").replace(")", " ")
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def truthy(v) -> bool:
    return str(v).strip().lower() in ("true", "1", "yes")


def step_key(path: str) -> str:
    """html/choices/cladocera/FChydoridae/25/25.html -> cladocera/FChydoridae/25/25"""
    p = PurePosixPath(path)
    base = str(p.with_suffix(""))
    if base.startswith("html/choices/"):
        return base[len("html/choices/"):]
    if base.startswith("html/"):
        return "entry/" + base[len("html/"):]
    return base


class Book:
    def __init__(self):
        self.wb = Workbook()
        self.wb.remove(self.wb.active)
        self.cols = {}   # sheet -> {header: letter}

    def sheet(self, name, headers, rows, widths=None, inputs=(), wrap=(), dropdowns=None):
        ws = self.wb.create_sheet(name)
        ws.append(headers)
        for r in rows:
            ws.append([r.get(h, "") if isinstance(r, dict) else r[i] for i, h in enumerate(headers)])
        self.cols[name] = {h: get_column_letter(i + 1) for i, h in enumerate(headers)}
        for i, h in enumerate(headers, 1):
            c = ws.cell(row=1, column=i)
            is_in = h in inputs
            c.font = Font(name=FONT, bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = INPUT_HEAD_FILL if is_in else HEAD_FILL
            c.alignment = Alignment(vertical="center", wrap_text=True)
            w = (widths or {}).get(h) or min(max(len(h) + 4, 12), 40)
            ws.column_dimensions[get_column_letter(i)].width = w
        n = len(rows) + 1
        for row in ws.iter_rows(min_row=2, max_row=n):
            for c in row:
                h = headers[c.column - 1]
                c.font = Font(name=FONT, size=10)
                c.alignment = Alignment(vertical="top", wrap_text=h in wrap)
                if h in inputs:
                    c.fill = INPUT_CELL_FILL
        ws.freeze_panes = "A2"
        if rows:
            ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{n}"
        for h, opts in (dropdowns or {}).items():
            dv = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
            ws.add_data_validation(dv)
            col = self.cols[name][h]
            dv.add(f"{col}2:{col}{max(n, 2000)}")
        return ws


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--inventory", required=True)
    ap.add_argument("--manifest", default="")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    inv = Path(a.inventory).expanduser()
    out = Path(a.out).expanduser()
    if not inv.is_dir():
        sys.exit(f"ERROR: inventory folder not found: {inv}")

    org = read(inv, "organism_pages.csv")
    sections = read(inv, "organism_sections.csv")
    images = read(inv, "organism_images.csv")
    figs = read(inv, "figure_links.csv")
    videos_inv = read(inv, "videos.csv")
    sp_videos = read(inv, "species_videos.csv")
    other_vl = read(inv, "page_video_links.csv")
    steps = read(inv, "key_steps.csv")
    step_links = read(inv, "key_step_links.csv")
    step_imgs = read(inv, "key_step_extra_images.csv")
    other_pages = read(inv, "other_pages.csv")
    status = {r["page"]: r["status"] for r in read(inv, "page_status.csv")}
    problems = read(inv, "problems.csv")
    manifest = {}
    if a.manifest and Path(a.manifest).expanduser().exists():
        with Path(a.manifest).expanduser().open(encoding="utf-8-sig", newline="") as fh:
            manifest = {r["video_id"]: r for r in csv.DictReader(fh)}
    else:
        print("  (no video manifest given: archive_url will be empty)")

    # ---------------- species ids ----------------
    species_rows, id_of_page, used = [], {}, Counter()
    live_species = [r for r in org if r["page_type"] == "species"]
    for r in sorted(live_species, key=lambda r: (r["index_name"] or r["name"]).lower()):
        disp = r["index_name"] or r["name"] or r["title"]
        sid = slug(disp) or slug(PurePosixPath(r["page"]).stem)
        used[sid] += 1
        if used[sid] > 1:
            sid = f"{sid}-{used[sid]}"
        id_of_page[r["page"]] = sid
        species_rows.append({
            "species_id": sid, "display_name": disp.strip(),
            "genus": re.sub(r"[()]", "", disp).split()[0] if disp.split() else "",
            **{k: r.get(k, "") for k in RANKS}, "tax_other": r.get("tax_other", ""),
            "old_path": r["page"], "path_folders": r.get("path_dirs", ""),
            "n_images": r.get("n_images", ""), "n_images_no_alt": r.get("n_images_no_alt", ""),
            "n_sections": r.get("n_sections", ""), "status": "", "notes": ""})
    extra_rows = []
    for r in sorted((r for r in org if r["page_type"] == "supplementary"), key=lambda r: r["page"]):
        parent = id_of_page.get(r["parent_page"], "")
        pid = f"{parent or 'unattached'}--{slug(PurePosixPath(r['page']).stem)}"
        id_of_page[r["page"]] = parent or pid     # unattached pages keep their own id
        extra_rows.append({"extra_page_id": pid, "parent_species_id": parent,
                           "title": r["title"] or r["name"], "old_path": r["page"],
                           "new_slug": "", "status": "", "notes": ""})
    n_extra = Counter(r["parent_species_id"] for r in extra_rows)
    for r in species_rows:
        r["n_extra_pages"] = n_extra.get(r["species_id"], 0)

    bk = Book()
    readme = bk.wb.create_sheet("README")      # filled in at the end

    sp_headers = ["species_id", "display_name", "genus"] + RANKS + ["tax_other", "old_path",
                  "path_folders", "n_images", "n_images_no_alt", "n_sections", "n_extra_pages",
                  "status", "notes"]
    bk.sheet("species", sp_headers, species_rows, widths={"species_id": 30, "display_name": 30,
             "old_path": 50, "path_folders": 40, "notes": 40},
             inputs=("status", "notes"), dropdowns={"status": ["draft", "reviewed", "ready"]})

    # ---------------- sections ----------------
    sec_rows = []
    for r in sections:
        sid = id_of_page.get(r["page"])
        if sid is None or sid == "":
            continue
        sec_rows.append({"species_id": sid, "from_page": "" if r["page"] in
                         {x["old_path"] for x in species_rows} else r["page"],
                         "section": r["section"], "subsection": r["subsection"],
                         "item_no": r["item_no"], "text": r["text"]})
    bk.sheet("species_sections", ["species_id", "from_page", "section", "subsection", "item_no", "text"],
             sec_rows, widths={"species_id": 28, "from_page": 12, "section": 28, "subsection": 26,
             "item_no": 8, "text": 110}, inputs=("text",), wrap=("text",))

    # ---------------- images & figures ----------------
    counters, img_rows = Counter(), []
    for r in images:
        owner = id_of_page.get(r["page"])
        if not owner:
            continue
        counters[owner] += 1
        img_rows.append({"image_id": f"{owner}-{counters[owner]:02d}", "species_id": owner,
                         "file": r["src"], "exists": r["exists"], "alt_on_old_site": r["alt"],
                         "alt_text_new": "", "width": r["width"], "height": r["height"],
                         "links_to_larger": r["links_to"], "hover_image": r.get("hover_image", ""),
                         "on_extra_page": "" if r["page"] in {x["old_path"] for x in species_rows} else r["page"]})
    bk.sheet("images", ["image_id", "species_id", "file", "exists", "alt_on_old_site", "alt_text_new",
                        "width", "height", "links_to_larger", "hover_image", "on_extra_page"], img_rows,
             widths={"image_id": 30, "species_id": 28, "file": 60, "alt_on_old_site": 28,
                     "alt_text_new": 50, "links_to_larger": 50, "hover_image": 50, "on_extra_page": 30},
             inputs=("alt_text_new",), wrap=("alt_text_new",))
    sp_paths = {x["old_path"] for x in species_rows}
    fig_rows = [{"species_id": id_of_page.get(r["page"], ""), "caption_on_old_site": r["label"],
                 "image": r["image"], "exists": r["exists"], "alt_text_new": "",
                 "on_extra_page": "" if r["page"] in sp_paths else r["page"]}
                for r in figs if id_of_page.get(r["page"])]
    bk.sheet("figures", ["species_id", "caption_on_old_site", "image", "exists", "alt_text_new", "on_extra_page"],
             fig_rows, widths={"species_id": 28, "caption_on_old_site": 40, "image": 70,
                               "alt_text_new": 50, "on_extra_page": 30}, inputs=("alt_text_new",),
             wrap=("alt_text_new",))
    imap_rows = [{"page": r["page"], "map_name": r["map_name"], "seq": r["seq"],
                  "label": r["label"], "label_new": "", "shape": r["shape"], "coords": r["coords"],
                  "goes_to": id_of_page.get(r["goes_to"]) or r["goes_to"],
                  "goes_to_type": ("species" if id_of_page.get(r["goes_to"]) else "other")}
                 for r in read(inv, "image_maps.csv")]
    bk.sheet("image_maps", ["page", "map_name", "seq", "label", "label_new", "shape", "coords",
                            "goes_to", "goes_to_type"], imap_rows,
             widths={"page": 50, "label": 32, "label_new": 30, "coords": 26, "goes_to": 40},
             inputs=("label_new",))

    # ---------------- videos ----------------
    dup_of = {v["video_id"]: v["duplicate_of"] for v in videos_inv if v.get("duplicate_of")}
    canon = lambda vid: dup_of.get(vid, vid)
    vid_rows = []
    for v in sorted(videos_inv, key=lambda v: v["video_id"]):
        vid = v["video_id"]
        m = manifest.get(canon(vid), {})
        sect = (v.get("catalog_section") or "").strip().lower()
        vtype = next((t for k, t in VIDEO_TYPES.items() if sect.startswith(k.split(":")[0])), "")
        if not vtype and vid.startswith("media/methods/"):
            vtype = "Sampling method"
        if not vtype and "/VF/" in vid:
            vtype = "Virtual Focus"
        old = v.get("caption_from_wrapper") or v.get("catalog_title") or v.get("caption_from_link_text") or ""
        flags = []
        if truthy(v.get("orphan_not_referenced")):
            flags.append("not linked from any page")
        if v.get("duplicate_of"):
            flags.append(f"duplicate of {v['duplicate_of']}")
        if not old:
            flags.append("no description anywhere")
        vid_rows.append({
            "video_id": vid, "video_type": vtype, "title_on_old_site": v.get("catalog_title", ""),
            "description_on_old_site": old, "is_new_flag": v.get("catalog_is_new", ""),
            "description_new": "", "archive_url": m.get("url", ""),
            "duration_s": m.get("duration_s", ""), "has_audio": m.get("has_audio", ""),
            "caption_file": "", "duplicate_of": v.get("duplicate_of", ""),
            "local_source": v["mov_path"], "flags": "; ".join(flags), "status": "", "notes": ""})
    bk.sheet("videos", ["video_id", "video_type", "title_on_old_site", "description_on_old_site",
                        "is_new_flag", "description_new", "archive_url", "duration_s", "has_audio",
                        "caption_file", "duplicate_of", "local_source", "flags", "status", "notes"],
             vid_rows, widths={"video_id": 46, "video_type": 22, "title_on_old_site": 40,
             "description_on_old_site": 60, "description_new": 60, "archive_url": 70,
             "local_source": 50, "flags": 36, "notes": 30},
             inputs=("video_type", "description_new", "caption_file", "status", "notes"),
             wrap=("description_on_old_site", "description_new", "title_on_old_site", "flags"),
             dropdowns={"status": ["draft", "reviewed", "ready"],
                        "video_type": sorted(set(VIDEO_TYPES.values()))})
    known_ids = {r["video_id"] for r in vid_rows}

    alias = {f["old_value"]: f["new_value"] for f in KNOWN_FIXES if f["kind"] == "video link"}
    sv_rows = []
    for r in sp_videos:
        sid = id_of_page.get(r["page"])
        if not sid:
            continue
        vid, note = r["video_id"], ""
        if not vid and r["link"] in alias and alias[r["link"]] in known_ids:
            vid, note = alias[r["link"]], "fixed via 'fixes' sheet"
        elif not vid:
            note = "link is broken: no clip found"
        sv_rows.append({"species_id": sid, "video_id": canon(vid) if vid else "", "label_on_old_site": r["label"],
                        "old_link": r["link"], "note": note})
    bk.sheet("species_videos", ["species_id", "video_id", "label_on_old_site", "old_link", "note"],
             sv_rows, widths={"species_id": 30, "video_id": 46, "label_on_old_site": 40, "old_link": 50, "note": 30},
             inputs=("note",))
    ov_rows = []
    for r in other_vl:
        if status.get(r["page"]) == "dead" or not r["video_id"]:
            continue
        kind = "key step" if r["page"].startswith("html/choices/") or r["page"] in ("html/begin.html", "html/rotifers.htm") else "page"
        ov_rows.append({"shown_on_type": kind, "shown_on": step_key(r["page"]) if kind == "key step" else r["page"],
                        "video_id": canon(r["video_id"]), "label_on_old_site": r.get("label", "")})
    bk.sheet("other_video_links", ["shown_on_type", "shown_on", "video_id", "label_on_old_site"],
             ov_rows, widths={"shown_on": 50, "video_id": 46, "label_on_old_site": 60})

    # ---------------- key ----------------
    live_steps = [r for r in steps if truthy(r.get("live", "True"))]
    sid_by_path = id_of_page
    key_rows = []
    for r in live_steps:
        t, tt = r["goes_to"], r["goes_to_type"]
        if tt == "key_step":
            goes = step_key(t)
        elif tt in ("organism_page", "organism_page_unlisted"):
            goes = sid_by_path.get(t) or t
        else:
            goes = t
        key_rows.append({
            "step_key": step_key(r["step_path"]), "group": r["key_group"], "choice_no": r["choice_no"],
            "prompt": r["step_prompt"], "choice_text": r["choice_text"],
            "label_needed": "yes" if not r["choice_text"] else "", "choice_label_new": "",
            "image": r["image"], "extra_images": r["extra_images"], "hover_image": r["image_hover"],
            "alt_text_new": "", "goes_to_type": tt, "goes_to": goes, "target_exists": r["target_exists"]})
    bk.sheet("key_steps", ["step_key", "group", "choice_no", "prompt", "choice_text", "label_needed",
                           "choice_label_new", "image", "extra_images", "hover_image", "alt_text_new",
                           "goes_to_type", "goes_to", "target_exists"], key_rows,
             widths={"step_key": 32, "choice_text": 70, "choice_label_new": 40, "image": 54,
                     "extra_images": 40, "hover_image": 40, "alt_text_new": 44, "goes_to": 44, "prompt": 30},
             inputs=("choice_label_new", "alt_text_new"), wrap=("choice_text", "choice_label_new", "alt_text_new"))
    info_rows, seen_k = [], set()
    for r in live_steps:
        k = step_key(r["step_path"])
        if k not in seen_k:
            seen_k.add(k)
            info_rows.append({"step_key": k, "old_path": r["step_path"], "step_prompt": r["step_prompt"],
                              "taxonomy_at_step": r["taxonomy_at_step"], "step_heading_new": "", "notes": ""})
    bk.sheet("key_steps_info", ["step_key", "old_path", "step_prompt", "taxonomy_at_step", "step_heading_new", "notes"],
             info_rows, widths={"step_key": 34, "old_path": 52, "step_prompt": 50, "taxonomy_at_step": 56,
                                "step_heading_new": 36, "notes": 30},
             inputs=("step_heading_new", "notes"), wrap=("step_prompt", "taxonomy_at_step"))
    bk.sheet("key_help_links", ["step_key", "text", "target", "exists"],
             [{"step_key": step_key(r["step_path"]), "text": r["text"], "target": r["target"], "exists": r["exists"]}
              for r in step_links if status.get(r["step_path"], "live") != "dead"],
             widths={"step_key": 32, "text": 60, "target": 60})
    live_step_paths = {r["step_path"] for r in live_steps}
    bk.sheet("key_images_of_text", ["step_key", "image", "alt_on_old_site", "alt_text_new"],
             [{"step_key": step_key(r["step_path"]), "image": r["image"], "alt_on_old_site": r["alt"], "alt_text_new": ""}
              for r in step_imgs if r["step_path"] in live_step_paths],
             widths={"step_key": 32, "image": 60, "alt_text_new": 50}, inputs=("alt_text_new",))

    # ---------------- anatomy diagrams ----------------
    ana_rows = [{"group": a["group"], "seq": a["seq"], "label": a["label"], "label_new": "",
                 "image": a["image"], "is_base": a["is_base"], "alt_text_new": "",
                 "old_path": a["page"], "intro": a["intro"]}
                for a in read(inv, "anatomy.csv")]
    bk.sheet("anatomy", ["group", "seq", "label", "label_new", "image", "is_base", "alt_text_new",
                         "old_path", "intro"], ana_rows,
             widths={"group": 12, "label": 34, "label_new": 30, "image": 56, "alt_text_new": 55,
                     "old_path": 40, "intro": 46},
             inputs=("label_new", "alt_text_new"), wrap=("alt_text_new", "intro"))

    # ---------------- groups outline ----------------
    grp_rows = []
    for g in read(inv, "groups.csv"):
        t, tt = g["goes_to"], g["goes_to_type"]
        goes = step_key(t) if tt == "key_step" else (id_of_page.get(t) or t)
        grp_rows.append({"group_page": g["group_page"], "seq": g["seq"], "level": g["level"], "rank": g["rank"],
                         "name": g["name"], "name_new": "", "goes_to_type": tt, "goes_to": goes, "notes": ""})
    bk.sheet("groups", ["group_page", "seq", "level", "rank", "name", "name_new", "goes_to_type", "goes_to", "notes"],
             grp_rows, widths={"group_page": 14, "name": 40, "name_new": 34, "goes_to": 60, "notes": 30},
             inputs=("name_new", "notes"))

    # ---------------- taxonomy ----------------
    tax_rows = []
    key_only = defaultdict(Counter)                      # values found on key steps (so typos there can be fixed too)
    seen_tax = set()
    for r in live_steps:
        k = step_key(r["step_path"])
        if k in seen_tax:
            continue
        seen_tax.add(k)
        for part in str(r["taxonomy_at_step"]).split(";"):
            if "=" in part:
                lab, val = (x.strip() for x in part.split("=", 1))
                if lab in RANKS and val:
                    key_only[lab][val] += 1
    for rank in RANKS:
        counts = Counter(r[rank] for r in species_rows if r[rank])
        for val in key_only[rank]:
            counts.setdefault(val, 0)                    # n_species 0 = used on key pages only
        vals = sorted(counts, key=lambda v: -counts[v])
        for v in vals:
            sim = [o for o in vals if o != v and difflib.SequenceMatcher(None, v.lower(), o.lower()).ratio() >= 0.85]
            best = max(sim, key=lambda o: counts[o]) if sim else ""
            tax_rows.append({"rank": rank, "value": v, "n_species": counts[v],
                             "similar_to": ", ".join(sim),
                             "suggested": best if best and counts[best] > counts[v] else "",
                             "preferred_spelling": ""})
    bk.sheet("taxonomy", ["rank", "value", "n_species", "similar_to", "suggested", "preferred_spelling"],
             tax_rows, widths={"value": 28, "similar_to": 30, "suggested": 24, "preferred_spelling": 28},
             inputs=("preferred_spelling",))

    # ---------------- pages ----------------
    bk.sheet("extra_pages", ["extra_page_id", "parent_species_id", "title", "old_path", "new_slug", "status", "notes"],
             extra_rows, widths={"extra_page_id": 50, "parent_species_id": 28, "title": 36, "old_path": 70},
             inputs=("new_slug", "status", "notes"), dropdowns={"status": ["draft", "reviewed", "ready"]})
    pg_rows = [{"old_path": r["page"], "title": r["title"], "status_on_old_site": status.get(r["page"], ""),
                "new_slug": "", "decision": "", "notes": ""}
               for r in other_pages if status.get(r["page"], "live") != "dead"]
    bk.sheet("pages", ["old_path", "title", "status_on_old_site", "new_slug", "decision", "notes"], pg_rows,
             widths={"old_path": 50, "title": 50, "notes": 40}, inputs=("new_slug", "decision", "notes"),
             dropdowns={"decision": ["keep", "rewrite", "merge", "drop"]})
    fixes = [{**f, "confirmed_by": "", "date": ""} for f in KNOWN_FIXES]
    bk.sheet("fixes", ["kind", "old_value", "new_value", "note", "confirmed_by", "date"], fixes,
             widths={"old_value": 44, "new_value": 44, "note": 80}, wrap=("note",),
             inputs=("confirmed_by", "date"))
    todo = [{"where": step_key(p["page"]) if p["page"].startswith("html/choices/") else p["page"],
             "kind": p["kind"], "detail": p["detail"]}
            for p in problems if p.get("status") != "dead"]
    bk.sheet("to_do", ["where", "kind", "detail"], todo, widths={"where": 60, "kind": 34, "detail": 70})

    # ---------------- README + progress tracker ----------------
    c = bk.cols
    def rng(sheet, header):
        col = c[sheet][header]
        return f"{sheet}!{col}2:{col}20000"
    ws = readme
    lines = [
        ("CFB Zooplankton Key: draft content workbook", "title"),
        ("Generated from the old site by inventory_site.py and upload_to_ia.py. It is a DRAFT: "
         "check it before relying on it.", ""),
        ("", ""),
        ("How to use", "h"),
        ("1. YELLOW columns are for people to fill in; everything else was generated from the old site.", ""),
        ("2. Edit text in the 'text' column of species_sections. *asterisks* = italics, **double** = bold; "
         "[words](fig:path) links words to a figure; <sup> and <sub> are allowed.", ""),
        ("3. Write NEW alt text in images, figures and key_steps; describe what the picture shows, "
         "not the file name. Run draft_alt_text.py once to pre-fill suggestions in the grey "
         "'alt_text_draft' columns, then write the final wording in 'alt_text_new'.", ""),
        ("4. For every clip write 'description_new' (what a person who cannot see the video needs to know).", ""),
        ("5. Use the taxonomy sheet to settle spellings (e.g. Brachiopoda / Branchiopoda). "
         "Note the ranks reflect the 2013 edition and may need updating by a specialist.", ""),
        ("6. Do not rename species_id or video_id values: other sheets refer to them. "
         "(For an extra page with no parent species, 'species_id' holds its extra_page_id.)", ""),
        ("", ""),
        ("Progress (updates automatically)", "h"),
    ]
    r = 1
    for text, kind in lines:
        ws.cell(row=r, column=1, value=text)
        ws.cell(row=r, column=1).font = Font(name=FONT, size=14 if kind == "title" else 11 if kind == "h" else 10,
                                             bold=kind in ("title", "h"))
        ws.cell(row=r, column=1).alignment = Alignment(wrap_text=True, vertical="top")
        r += 1
    prog = [
        ("Species in the key", f"=COUNTA({rng('species','species_id')})", None),
        ("Species marked ready", f'=COUNTIF({rng("species","status")},"ready")', f"=B{r}"),
        ("Images with new alt text", f"=COUNTA({rng('images','alt_text_new')})", f"=COUNTA({rng('images','image_id')})"),
        ("Figures with new alt text", f"=COUNTA({rng('figures','alt_text_new')})", f"=COUNTA({rng('figures','image')})"),
        ("Clips with a new description", f"=COUNTA({rng('videos','description_new')})", f"=COUNTA({rng('videos','video_id')})"),
        ("Key choices still needing a label", f'=COUNTIF({rng("key_steps","label_needed")},"yes")-COUNTA({rng("key_steps","choice_label_new")})', f'=COUNTIF({rng("key_steps","label_needed")},"yes")'),
    ]
    for lab, done, total in prog:
        ws.cell(row=r, column=1, value=lab)
        ws.cell(row=r, column=2, value=done)
        if total:
            ws.cell(row=r, column=3, value=total)
        for col in (1, 2, 3):
            ws.cell(row=r, column=col).font = Font(name=FONT, size=10)
        r += 1
    ws.cell(row=r, column=1, value="(column B = done, column C = out of)").font = Font(name=FONT, size=9, italic=True)
    ws.column_dimensions["A"].width = 100
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 12
    # README first
    bk.wb.move_sheet("README", offset=-(len(bk.wb.sheetnames) - 1))
    bk.wb.save(out)

    print(f"Wrote {out}")
    print(f"  species: {len(species_rows)}   extra pages: {len(extra_rows)}   section rows: {len(sec_rows)}")
    print(f"  images: {len(img_rows)}   figures: {len(fig_rows)}   videos: {len(vid_rows)}")
    print(f"  key choices: {len(key_rows)} ({sum(1 for k in key_rows if k['label_needed'])} need a label)")
    print(f"  taxonomy values flagged as near-duplicates: {sum(1 for t in tax_rows if t['similar_to'])}")


if __name__ == "__main__":
    main()
