#!/usr/bin/env python3
"""
update_workbook.py

Brings an EXISTING workbook up to date with newer inventory output, without
overwriting anyone's edits. (build_workbook.py makes a brand-new workbook; this
script only adds and refreshes.)

  python inventory_site.py --root "$ZK/cfbkey_local" --out "$ZK/inventory"     # first, refresh the reports
  python update_workbook.py --workbook "$ZK/cfbkey_workbook.xlsx" --inventory "$ZK/inventory"

WHAT IT DOES
  1. Adds the sheet 'key_steps_info' (one row per key step: prompt, where you are in
     the classification, a place to override the page heading) if it is not there.
  2. In 'key_steps', refreshes 'choice_text' so the colored letters that match the
     arrows in the pictures, such as (p), (m), (d), keep their colors. A cell is only
     changed if it still holds the original, unedited text. Cells someone has edited
     are left alone and counted.
  3. Adds the 'image_maps' sheet, the new 'hover_image' and 'on_extra_page' columns,
     the 'anatomy' sheet (the diagram labels) and the 'groups' sheet (the Groups outline) if it is not there.
  4. Adds to the 'taxonomy' sheet any rank value that occurs only on key pages (so a
     typo there can be corrected with 'preferred_spelling' like any other).
  5. Makes a dated backup copy of the workbook first.
It is safe to run more than once.
"""

import argparse
import csv
import datetime
import difflib
import re
import shutil
import sys
from collections import Counter
from pathlib import Path, PurePosixPath

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

MARK = re.compile(r"\{(?:red|blue|green|orange)\|([^}]*)\}")


def same_text(new: str, old: str) -> bool:
    """Equal once colour marks are removed and the line separator ' | ' is read as ' / '."""
    return MARK.sub(r"\1", new).replace(" | ", " / ") == old.replace(" | ", " / ")


def step_key(path: str) -> str:
    p = PurePosixPath(path)
    base = str(p.with_suffix(""))
    if base.startswith("html/choices/"):
        return base[len("html/choices/"):]
    if base.startswith("html/"):
        return "entry/" + base[len("html/"):]
    return base


def read_csv(inv: Path, name: str):
    f = inv / name
    if not f.exists():
        sys.exit(f"ERROR: {f} not found. Run inventory_site.py first.")
    with f.open(encoding="utf-8-sig", newline="") as fh:
        return list(csv.DictReader(fh))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--inventory", required=True)
    ap.add_argument("--overwrite-key-text", action="store_true",
                    help="also replace key text cells that differ from the old site (a backup is made first)")
    a = ap.parse_args()
    wbp = Path(a.workbook).expanduser()
    inv = Path(a.inventory).expanduser()
    if not wbp.exists():
        sys.exit(f"ERROR: workbook not found: {wbp}")
    backup = wbp.with_name(f"{wbp.stem}_backup_{datetime.datetime.now():%Y%m%d-%H%M%S}{wbp.suffix}")
    shutil.copy2(wbp, backup)
    print(f"Backup saved: {backup.name}")

    steps = [r for r in read_csv(inv, "key_steps.csv") if str(r.get("live", "True")).lower() == "true"]
    wb = load_workbook(wbp)

    # ---- 1. key_steps_info ----
    if "key_steps_info" in wb.sheetnames:
        print("key_steps_info: already present, left as it is")
    else:
        ws = wb.create_sheet("key_steps_info")
        heads = ["step_key", "old_path", "step_prompt", "taxonomy_at_step", "step_heading_new", "notes"]
        ws.append(heads)
        seen = set()
        for r in steps:
            k = step_key(r["step_path"])
            if k in seen:
                continue
            seen.add(k)
            ws.append([k, r["step_path"], r["step_prompt"], r["taxonomy_at_step"], None, None])
        widths = {"step_key": 34, "old_path": 52, "step_prompt": 50, "taxonomy_at_step": 56,
                  "step_heading_new": 36, "notes": 30}
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=1, column=i)
            is_in = h in ("step_heading_new", "notes")
            c.font = Font(name="Arial", bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = PatternFill("solid", fgColor="FFD966" if is_in else "1F3864")
            ws.column_dimensions[get_column_letter(i)].width = widths[h]
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=True)
                if heads[c.column - 1] in ("step_heading_new", "notes"):
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{ws.max_row}"
        print(f"key_steps_info: added with {ws.max_row - 1} steps")

    # ---- 2. refresh coloured key text where nobody has edited it ----
    if "key_steps" not in wb.sheetnames:
        print("key_steps: sheet not found, skipped")
    else:
        ws = wb["key_steps"]
        heads = [c.value for c in ws[1]]
        ci = {h: i for i, h in enumerate(heads)}
        new_text = {(step_key(r["step_path"]), str(r["choice_no"])): r["choice_text"] for r in steps}
        refreshed = edited = same = filled = overwritten = 0
        diffs = []
        for row in ws.iter_rows(min_row=2):
            key = (str(row[ci["step_key"]].value), str(row[ci["choice_no"]].value))
            new = new_text.get(key)
            cell = row[ci["choice_text"]]
            old = "" if cell.value is None else str(cell.value)
            if new is None or new == old:
                same += 1
            elif same_text(new, old) and new != old:
                cell.value = new
                refreshed += 1
            elif not old.strip() and new.strip():
                cell.value = new                      # an empty cell: nobody has typed anything, so fill it
                filled += 1
            elif old:
                if a.overwrite_key_text and new.strip():       # never replace existing text with nothing
                    cell.value = new
                    overwritten += 1
                else:
                    edited += 1
                    diffs.append((key, old, new))
        print(f"key_steps text: {refreshed} cells refreshed with colored letters, {filled} empty cells filled, "
              f"{edited} left alone because the text differs from the old site, {same} unchanged"
              + (f", {overwritten} overwritten" if overwritten else ""))
        if diffs:
            print("  These cells differ from what the old site now yields (edited by a person, or the parser improved):")
            for (sk, no), old_t, new_t in diffs[:12]:
                print(f"   {sk} choice {no}\n      workbook : {old_t[:95]}\n      old site : {new_t[:95]}")
            if len(diffs) > 12:
                print(f"   ... and {len(diffs) - 12} more")
            print("  To replace them with the old-site text, rerun with --overwrite-key-text (a backup was made).")
            print("  (A cell is never overwritten with empty text, so cells the old site yields nothing for are always kept.)")

    # ---- 2b. the Groups outline ----
    if "groups" in wb.sheetnames:
        print("groups: already present, left as it is")
    elif not (inv / "groups.csv").exists():
        print("groups: groups.csv not found in the inventory (rerun inventory_site.py)")
    else:
        sp_by_old = {}
        if "species" in wb.sheetnames:
            hs = [c.value for c in wb["species"][1]]
            for row in wb["species"].iter_rows(min_row=2, values_only=True):
                d = dict(zip(hs, row))
                if d.get("old_path"):
                    sp_by_old[d["old_path"]] = d["species_id"]
        ws = wb.create_sheet("groups")
        heads = ["group_page", "seq", "level", "rank", "name", "name_new", "goes_to_type", "goes_to", "notes"]
        ws.append(heads)
        for g in read_csv(inv, "groups.csv"):
            t, tt = g["goes_to"], g["goes_to_type"]
            goes = step_key(t) if tt == "key_step" else (sp_by_old.get(t) or t)
            ws.append([g["group_page"], int(g["seq"]), int(g["level"]), g["rank"], g["name"], None, tt, goes, None])
        widths = {"group_page": 14, "seq": 6, "level": 7, "rank": 12, "name": 40, "name_new": 34,
                  "goes_to_type": 18, "goes_to": 60, "notes": 30}
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=1, column=i)
            is_in = h in ("name_new", "notes")
            c.font = Font(name="Arial", bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = PatternFill("solid", fgColor="FFD966" if is_in else "1F3864")
            ws.column_dimensions[get_column_letter(i)].width = widths[h]
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                if heads[c.column - 1] in ("name_new", "notes"):
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{ws.max_row}"
        print(f"groups: added with {ws.max_row - 1} rows")

    # ---- 2c. the anatomy diagrams ----
    if "anatomy" in wb.sheetnames:
        print("anatomy: already present, left as it is")
    elif not (inv / "anatomy.csv").exists():
        print("anatomy: anatomy.csv not found in the inventory (rerun inventory_site.py)")
    else:
        ws = wb.create_sheet("anatomy")
        heads = ["group", "seq", "label", "label_new", "image", "is_base", "alt_text_new", "old_path", "intro"]
        ws.append(heads)
        for r in read_csv(inv, "anatomy.csv"):
            ws.append([r["group"], int(r["seq"]), r["label"], None, r["image"], r["is_base"], None,
                       r["page"], r["intro"]])
        widths = {"group": 12, "seq": 6, "label": 34, "label_new": 30, "image": 56, "is_base": 9,
                  "alt_text_new": 55, "old_path": 40, "intro": 46}
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=1, column=i)
            is_in = h in ("label_new", "alt_text_new")
            c.font = Font(name="Arial", bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = PatternFill("solid", fgColor="FFD966" if is_in else "1F3864")
            ws.column_dimensions[get_column_letter(i)].width = widths[h]
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=heads[c.column - 1] in ("alt_text_new", "intro"))
                if heads[c.column - 1] in ("label_new", "alt_text_new"):
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{ws.max_row}"
        print(f"anatomy: added with {ws.max_row - 1} label rows")

    # ---- 2d. new columns and the image_maps sheet ----
    def add_cols(sheet, cols, source_csv, key_cols):
        """Add missing columns to an existing sheet and fill them from the inventory."""
        if sheet not in wb.sheetnames:
            return
        ws_ = wb[sheet]
        heads = [c.value for c in ws_[1]]
        added = [c for c in cols if c not in heads]
        if not added:
            print(f"{sheet}: columns already present")
            return
        for name in added:
            col = ws_.max_column + 1
            c = ws_.cell(row=1, column=col, value=name)
            c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F3864")
            ws_.column_dimensions[get_column_letter(col)].width = 44
            heads = [x.value for x in ws_[1]]
        src = {}
        for r in read_csv(inv, source_csv):
            src[tuple(str(r[k]) for k in key_cols[1])] = r
        ci = {h: i for i, h in enumerate(heads)}
        n = 0
        for row in ws_.iter_rows(min_row=2):
            k = tuple(str(row[ci[h]].value) for h in key_cols[0])
            r = src.get(k)
            if not r:
                continue
            for name in added:
                row[ci[name]].value = r.get(name) or None
                row[ci[name]].font = Font(name="Arial", size=10)
            n += 1
        print(f"{sheet}: added {', '.join(added)} and filled {n} row(s)")

    add_cols("images", ["hover_image"], "organism_images.csv", (["file"], ["src"]))
    if "figures" in wb.sheetnames and "on_extra_page" not in [c.value for c in wb["figures"][1]]:
        ws_ = wb["figures"]
        col = ws_.max_column + 1
        c = ws_.cell(row=1, column=col, value="on_extra_page")
        c.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        ws_.column_dimensions[get_column_letter(col)].width = 36
        heads = [x.value for x in ws_[1]]
        ci = {h: i for i, h in enumerate(heads)}
        src = {str(r["image"]): r["page"] for r in read_csv(inv, "figure_links.csv")}
        sp_paths = set()
        if "species" in wb.sheetnames:
            hs = [x.value for x in wb["species"][1]]
            for row in wb["species"].iter_rows(min_row=2, values_only=True):
                d = dict(zip(hs, row))
                if d.get("old_path"):
                    sp_paths.add(str(d["old_path"]))
        n = 0
        for row in ws_.iter_rows(min_row=2):
            pg = src.get(str(row[ci["image"]].value))
            if pg and pg not in sp_paths:
                row[ci["on_extra_page"]].value = pg
                row[ci["on_extra_page"]].font = Font(name="Arial", size=10)
                n += 1
        print(f"figures: added on_extra_page and filled {n} row(s)")

    if "image_maps" in wb.sheetnames:
        print("image_maps: already present, left as it is")
    elif (inv / "image_maps.csv").exists():
        rows_im = read_csv(inv, "image_maps.csv")
        sp_by_old = {}
        if "species" in wb.sheetnames:
            hs = [c.value for c in wb["species"][1]]
            for row in wb["species"].iter_rows(min_row=2, values_only=True):
                d = dict(zip(hs, row))
                if d.get("old_path"):
                    sp_by_old[str(d["old_path"])] = d["species_id"]
        ws_ = wb.create_sheet("image_maps")
        heads = ["page", "map_name", "seq", "label", "label_new", "shape", "coords", "goes_to", "goes_to_type"]
        ws_.append(heads)
        for r in rows_im:
            sid = sp_by_old.get(r["goes_to"])
            ws_.append([r["page"], r["map_name"], int(r["seq"]), r["label"], None, r["shape"],
                        r["coords"], sid or r["goes_to"], "species" if sid else "other"])
        for i, h in enumerate(heads, 1):
            c = ws_.cell(row=1, column=i)
            is_in = h == "label_new"
            c.font = Font(name="Arial", bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = PatternFill("solid", fgColor="FFD966" if is_in else "1F3864")
            ws_.column_dimensions[get_column_letter(i)].width = {"page": 50, "label": 32, "label_new": 30,
                                                                 "coords": 26, "goes_to": 40}.get(h, 12)
        for row in ws_.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                if heads[c.column - 1] == "label_new":
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
        ws_.freeze_panes = "A2"
        print(f"image_maps: added with {ws_.max_row - 1} region(s)")

    # ---- 3. taxonomy values that appear only on key steps ----
    RANKS = ["Phylum", "Subphylum", "Class", "Subclass", "Order", "Suborder", "Family"]
    if "taxonomy" in wb.sheetnames:
        ws = wb["taxonomy"]
        heads = [c.value for c in ws[1]]
        ci = {h: i for i, h in enumerate(heads)}
        have = {}
        for row in ws.iter_rows(min_row=2):
            if row[ci["rank"]].value:
                have[(str(row[ci["rank"]].value), str(row[ci["value"]].value))] = int(row[ci["n_species"]].value or 0)
        seen_steps, found = set(), Counter()
        for r in steps:
            k = step_key(r["step_path"])
            if k in seen_steps:
                continue
            seen_steps.add(k)
            for part in str(r["taxonomy_at_step"]).split(";"):
                if "=" in part:
                    lab, val = (x.strip() for x in part.split("=", 1))
                    if lab in RANKS and val:
                        found[(lab, val)] += 1
        added = 0
        for (lab, val), n in sorted(found.items()):
            if (lab, val) in have:
                continue
            rank_vals = {v: c for (rk, v), c in have.items() if rk == lab}
            sim = [o for o in rank_vals if difflib.SequenceMatcher(None, val.lower(), o.lower()).ratio() >= 0.85]
            best = max(sim, key=lambda o: rank_vals[o]) if sim else ""
            ws.append([lab, val, 0, ", ".join(sim), best if best and rank_vals[best] > 0 else "", None])
            added += 1
            have[(lab, val)] = 0
        print(f"taxonomy: {added} value(s) used only on key pages added for review")

    wb.save(wbp)
    print("Saved.")


if __name__ == "__main__":
    main()
