#!/usr/bin/env python3
"""
audit_old_site.py

Finds every INTERACTIVE BEHAVIOR and every distinct PAGE LAYOUT in the old site, so we
can be sure nothing that visitors rely on gets lost in the rebuild. It reads the old
site only; it changes nothing.

  python audit_old_site.py --root "$ZK/cfbkey_local" --out "$ZK/inventory"

It prints a summary and writes two files into --out:
  behavior_audit.csv   one row per page: which behaviors it uses
  layout_families.csv  pages grouped into families that share the same look and
                       behaviors, with ONE example page per family

Why it matters: if the audit says 230 pages share a family, one screenshot of one
example page shows me that whole family. A handful of screenshots can then cover
the entire site.
"""

import argparse
import csv
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("Needs beautifulsoup4:  conda install -c conda-forge beautifulsoup4 lxml")

NAV_IMG = re.compile(r"(images/buttons/|indexmain|groupbrackets|shim|images/group/|images/extras/)", re.I)
IMG_EXT = (".jpg", ".jpeg", ".png", ".gif", ".bmp")

BEHAVIORS = [
    ("hover_swap_picture", "pointing at a picture swaps it for another (not the menu buttons)"),
    ("click_swap_picture", "clicking text or a picture swaps a picture (labeled diagrams)"),
    ("hover_other", "pointing at something else triggers an effect (text, layers)"),
    ("popup_picture", "a link opens a picture in a small popup window"),
    ("popup_page", "a link opens another page in a small popup window"),
    ("popup_video", "a link opens a video in a small popup window"),
    ("image_map", "a picture with clickable regions (image map)"),
    ("form", "a form or input box"),
    ("frames", "frames or iframes"),
    ("plugin_embed", "an embedded plugin object (Flash, QuickTime ...)"),
    ("custom_script", "custom JavaScript beyond the standard menu-rollover and analytics code"),
    ("tooltip_titles", "tooltips (title attributes on links or pictures)"),
    ("new_window_links", "links that open in a new window or tab (target=_blank)"),
    ("positioned_layers", "absolutely positioned or hidden/shown layers"),
    ("old_tags", "obsolete effects (marquee, blink, applet)"),
]


def read(path: Path) -> str:
    return path.read_bytes().decode("cp1252", errors="replace")


def strip_mm_functions(txt: str) -> str:
    """Remove the standard Dreamweaver MM_* functions (any nesting depth) by matching braces."""
    pat = re.compile(r"function\s+MM_\w+\s*\([^)]*\)\s*\{")
    out, i = [], 0
    while True:
        m = pat.search(txt, i)
        if not m:
            out.append(txt[i:])
            return "".join(out)
        out.append(txt[i:m.start()])
        depth, j = 1, m.end()
        while j < len(txt) and depth:
            depth += (txt[j] == "{") - (txt[j] == "}")
            j += 1
        i = j


def analyse(path: Path):
    raw = read(path)
    soup = BeautifulSoup(raw, "lxml")
    f = Counter()

    for tag in soup.find_all(True):
        om = tag.get("onmouseover") or ""
        oc = tag.get("onclick") or ""
        if "MM_swapImage(" in om:
            if tag.name == "img" and NAV_IMG.search(tag.get("src", "")):
                pass
            elif tag.name == "a" and tag.find("img") and NAV_IMG.search(tag.find("img").get("src", "")):
                pass
            else:
                f["hover_swap_picture"] += 1
        elif om:
            f["hover_other"] += 1
        if "MM_swapImage(" in oc:
            f["click_swap_picture"] += 1
        for t in re.findall(r"MM_openBrWindow\(\s*'([^']+)'", oc):
            low = t.lower().split("?")[0]
            if low.endswith(IMG_EXT):
                f["popup_picture"] += 1
            elif low.endswith(".mov") or "/media/" in "/" + low.replace("\\", "/") or low.startswith("media/"):
                f["popup_video"] += 1
            else:
                f["popup_page"] += 1
        if tag.name == "a" and (tag.get("target") or "").lower() == "_blank":
            f["new_window_links"] += 1
        if tag.name in ("a", "img", "area") and tag.get("title"):
            f["tooltip_titles"] += 1
        if tag.name in ("map", "area") or tag.get("usemap"):
            f["image_map"] += 1
        if tag.name in ("form", "input", "select", "textarea", "button"):
            f["form"] += 1
        if tag.name in ("frameset", "frame", "iframe"):
            f["frames"] += 1
        if tag.name in ("embed", "object", "applet"):
            src = (tag.get("src") or tag.get("data") or tag.get("classid") or "").lower()
            f["plugin_embed"] += 1
        if tag.name in ("marquee", "blink", "applet"):
            f["old_tags"] += 1
        style = (tag.get("style") or "").lower()
        if "position:absolute" in style.replace(" ", "") or "visibility:" in style.replace(" ", "") or tag.name == "layer":
            f["positioned_layers"] += 1
    # popups on <img onclick> / <li onclick> are already covered above (all tags)

    # custom script = what remains after removing the standard menu-rollover functions and analytics
    for sc in soup.find_all("script"):
        txt = sc.get_text()
        if not txt.strip() or "google-analytics" in txt or "_gat" in txt or "gaJsHost" in txt:
            continue
        txt = strip_mm_functions(txt)
        txt = re.sub(r"//.*|<!--|//-->|-->", "", txt)
        if len(re.sub(r"\s+", "", txt)) > 40:
            f["custom_script"] += 1
        if "document.write" in txt:
            f["custom_script"] += 1
    return f, soup


def category(rel: str) -> str:
    low = rel.lower()
    for key, name in (("html/organisms/", "species/organism page"), ("html/choices/", "key step"),
                      ("media/", "video wrapper"), ("html/anatomy/", "anatomy diagram"),
                      ("html/idpages/", "id page")):
        if low.startswith(key):
            return name
    return "other page"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", required=True)
    ap.add_argument("--out", default="inventory")
    a = ap.parse_args()
    root = Path(a.root).expanduser().resolve()
    out = Path(a.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    live = None
    ps = out / "page_status.csv"
    if ps.exists():
        with ps.open(encoding="utf-8-sig", newline="") as fh:
            live = {r["page"] for r in csv.DictReader(fh) if r["status"] == "live"}

    pages = sorted(p for p in root.rglob("*")
                   if p.suffix.lower() in (".htm", ".html") and not p.name.startswith("._"))
    rows, fam = [], defaultdict(list)
    for p in pages:
        rel = p.relative_to(root).as_posix()
        if live is not None and rel.startswith("html/") and rel not in live:
            continue                                   # skip dead pages when we know which are dead
        feats, soup = analyse(p)
        used = tuple(k for k, _ in BEHAVIORS if feats.get(k))
        classes = {c for tag in soup.find_all(class_=True) for c in tag.get("class", [])}
        looks = tuple(sorted(classes & {"headerblack", "headergray", "textsizeheadline", "textsize12left"}))
        key = (category(rel), used, looks)
        fam[key].append(rel)
        rows.append({"page": rel, "category": category(rel), **{k: feats.get(k, 0) for k, _ in BEHAVIORS}})

    with (out / "behavior_audit.csv").open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["page", "category"] + [k for k, _ in BEHAVIORS])
        w.writeheader()
        w.writerows(rows)

    fam_rows = []
    for i, (key, members) in enumerate(sorted(fam.items(), key=lambda kv: -len(kv[1])), 1):
        cat, used, looks = key
        example = sorted(members, key=lambda m: (len(m), m))[0]
        fam_rows.append({"family": i, "pages": len(members), "category": cat,
                         "behaviors": "; ".join(used) or "(none)", "page_styles": "; ".join(looks),
                         "example_page": example})
    with (out / "layout_families.csv").open("w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["family", "pages", "category", "behaviors", "page_styles", "example_page"])
        w.writeheader()
        w.writerows(fam_rows)

    print(f"Audited {len(rows)} pages" + (" (live pages only)" if live is not None else ""))
    print("\nBEHAVIORS FOUND (pages that use each):")
    for k, desc in BEHAVIORS:
        n = sum(1 for r in rows if r[k])
        ex = next((r["page"] for r in rows if r[k]), "")
        if n:
            print(f"  {n:4} pages  {desc}\n              e.g. {ex}")
    print(f"\nLAYOUT FAMILIES: {len(fam_rows)}  (one example page each, largest first)")
    for r in fam_rows[:25]:
        print(f"  #{r['family']:<3} {r['pages']:4} pages  {r['category']:22}  {r['behaviors'][:60]}\n"
              f"        example: {r['example_page']}")
    if len(fam_rows) > 25:
        print(f"  ... and {len(fam_rows) - 25} more in layout_families.csv")
    print(f"\nFiles written to {out}")


if __name__ == "__main__":
    main()
