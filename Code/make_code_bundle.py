#!/usr/bin/env python3
"""
make_code_bundle.py

Packs the old site into ONE text file that is small enough to hand to Claude but
still representative of all 600+ pages, so unseen layouts stop causing surprises.

  python make_code_bundle.py --root "$ZK/cfbkey_local" --inventory "$ZK/inventory" \\
         --out "$ZK/code_bundle.txt"

WHAT IT CONTAINS
  1. The shared stylesheet.
  2. A STRUCTURE DIGEST of every live page: a one-line fingerprint of its layout
     (which table shapes, behaviors and styles it uses). Pages that share a
     fingerprint are listed together, so Claude can see that, say, 80 pages are
     the same shape.
  3. The FULL source of one example page per fingerprint group, with the
     boilerplate (the Dreamweaver rollover functions, analytics, the navigation
     bar) removed, so what is left is the part that differs between pages.
  4. Every page whose layout is unique or rare (fewer than --rare copies), in full.

Typical result: a few hundred kilobytes instead of tens of megabytes, covering
every distinct layout in the site.

OPTIONS
  --max-chars N    stop adding examples past this size (default 600000)
  --rare N         include in full any layout family with fewer than N pages (default 3)
  --only-dirs A,B  restrict to these top folders (e.g. html/choices,html/anatomy)
  --fine           group by exact table shapes instead of the coarse grouping (many more groups)
"""

import argparse
import csv
import hashlib
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    sys.exit("Needs beautifulsoup4:  conda install -c conda-forge beautifulsoup4 lxml")

NAV_IMG = re.compile(r"(images/buttons/|indexmain|groupbrackets|shim|images/group/)", re.I)
BOILER_FUNCS = re.compile(
    r"function\s+MM_(swapImgRestore|preloadImages|findObj|swapImage|openBrWindow)\s*\([^)]*\)\s*\{", re.I)


def read(p: Path) -> str:
    return p.read_bytes().decode("cp1252", errors="replace")


def strip_funcs(txt: str) -> str:
    out, i = [], 0
    while True:
        m = BOILER_FUNCS.search(txt, i)
        if not m:
            out.append(txt[i:])
            return "".join(out)
        out.append(txt[i:m.start()])
        depth, j = 1, m.end()
        while j < len(txt) and depth:
            depth += (txt[j] == "{") - (txt[j] == "}")
            j += 1
        out.append("/* [standard rollover function removed] */")
        i = j


def trim(raw: str) -> str:
    """Remove what every page repeats: rollover functions, analytics, preload lists, the nav bar."""
    t = strip_funcs(raw)
    t = re.sub(r"(?is)<script[^>]*>\s*var gaJsHost.*?</script>", "", t)
    t = re.sub(r"(?is)<script[^>]*>\s*try\s*\{\s*var pageTracker.*?</script>", "", t)
    t = re.sub(r"(?is)onLoad\s*=\s*\"MM_preloadImages\([^\"]*\)\"", 'onLoad="[preload list removed]"', t)
    # navigation rows: any table row that contains a menu button image
    soup = BeautifulSoup(t, "lxml")
    for tr in soup.find_all("tr"):
        if any(NAV_IMG.search(im.get("src", "")) for im in tr.find_all("img")) and not tr.find("table"):
            tr.replace_with(BeautifulSoup("<!-- [navigation row removed] -->", "lxml").comment
                            if False else BeautifulSoup("<tr><td>[navigation row removed]</td></tr>", "lxml").tr)
    t = str(soup)
    t = re.sub(r"(?is)<!--\s*-->", "", t)
    t = re.sub(r"\n\s*\n+", "\n", t)
    return t.strip()


# Classes that change what a page MEANS, not merely how big it is.
MEANING_CLASSES = {"headerblack", "headergray", "textsizeheadline", "textsize12left", "text12italics",
                   "red", "blue", "green", "mainstyle"}


def fingerprint(raw: str, soup: BeautifulSoup, coarse: bool = True) -> str:
    """A signature of a page's KIND of layout. Coarse on purpose: two pages that differ only
    by a few extra rows or a one-off style class should land in the same group."""
    bits = []
    body = soup.body or soup
    cells = [td for td in body.find_all("td")]
    widest = max((len(tr.find_all("td", recursive=False)) for tr in body.find_all("tr")), default=0)
    if coarse:
        bits.append(f"cols:{'1' if widest <= 1 else '2' if widest == 2 else '3+'}")
        bits.append(f"nest:{'deep' if len(body.find_all('table')) > 3 else 'flat'}")
    else:
        shapes = []
        for tbl in body.find_all("table")[:6]:
            rows = tbl.find_all("tr", recursive=False) or tbl.find_all("tr")
            cols = max((len(tr.find_all("td", recursive=False)) for tr in rows), default=0)
            shapes.append(f"{min(len(rows), 30)}x{cols}")
        bits.append("tables:" + ",".join(shapes))
    classes = {c for tag in body.find_all(class_=True) for c in tag.get("class", [])}
    bits.append("cls:" + ",".join(sorted(classes & MEANING_CLASSES)) if coarse
                else "cls:" + ",".join(sorted(classes)[:12]))
    beh = []
    for name, test in (
        ("hoverswap", lambda t: "MM_swapImage(" in (t.get("onmouseover") or "")),
        ("clickswap", lambda t: "MM_swapImage(" in (t.get("onclick") or "")),
        ("popup_pic", lambda t: bool(re.search(r"MM_openBrWindow\(\s*'[^']+\.(jpg|jpeg|png|gif)",
                                               (t.get("onclick") or "") + (t.get("href") or ""), re.I))),
        ("popup_other", lambda t: "MM_openBrWindow(" in ((t.get("onclick") or "") + (t.get("href") or ""))),
        ("embed", lambda t: t.name in ("embed", "object", "applet")),
        ("imagemap", lambda t: t.name in ("map", "area") or bool(t.get("usemap"))),
        ("form", lambda t: t.name in ("form", "input", "select", "textarea")),
        ("frames", lambda t: t.name in ("frameset", "frame", "iframe")),
        ("rowspan", lambda t: bool(t.get("rowspan"))),
        ("list", lambda t: t.name in ("ul", "ol")),
        ("movlink", lambda t: ".mov" in ((t.get("href") or "") + (t.get("onclick") or "")).lower()),
        ("taxonomy", lambda t: t.name == "font" and (t.get("color") or "").lower() in ("#ff0000", "red")),
    ):
        if any(test(t) for t in body.find_all(True)):
            beh.append(name)
    bits.append("beh:" + ",".join(beh))
    imgs = [im for im in body.find_all("img") if not NAV_IMG.search(im.get("src", ""))]
    n = len(imgs)
    bits.append("imgs:" + ("0" if n == 0 else "1" if n == 1 else "2-3" if n <= 3 else "4-9" if n <= 9 else "10+")
                if coarse else f"imgs:{min(n, 20)}")
    text_len = len(re.sub(r"\s+", " ", body.get_text(" ", strip=True)))
    bits.append("text:" + ("none" if text_len < 20 else "small" if text_len < 400 else
                           "medium" if text_len < 2000 else "large"))
    if not any(b.split(":", 1)[1] for b in bits[:3]):
        bits.append("EMPTY-OR-UNPARSEABLE")
    return " | ".join(bits)


def category(rel: str) -> str:
    low = rel.lower()
    for key, name in (("html/organisms/", "organism"), ("html/choices/", "keystep"), ("media/", "wrapper"),
                      ("html/anatomy/", "anatomy"), ("html/idpages/", "idpage")):
        if low.startswith(key):
            return name
    return "other"


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", required=True)
    ap.add_argument("--inventory", default="")
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-chars", type=int, default=600000)
    ap.add_argument("--rare", type=int, default=3)
    ap.add_argument("--only-dirs", default="")
    ap.add_argument("--fine", action="store_true",
                    help="group pages by exact table shapes (many more, smaller groups)")
    a = ap.parse_args()
    root = Path(a.root).expanduser().resolve()
    outp = Path(a.out).expanduser()

    live = None
    if a.inventory:
        ps = Path(a.inventory).expanduser() / "page_status.csv"
        if ps.exists():
            with ps.open(encoding="utf-8-sig", newline="") as fh:
                live = {r["page"] for r in csv.DictReader(fh) if r["status"] == "live"}

    only = [d.strip().lower() for d in a.only_dirs.split(",") if d.strip()]
    pages = []
    for p in sorted(root.rglob("*")):
        if p.suffix.lower() not in (".htm", ".html") or p.name.startswith("._"):
            continue
        rel = p.relative_to(root).as_posix()
        if live is not None and rel.startswith("html/") and rel not in live:
            continue
        if only and not any(rel.lower().startswith(d) for d in only):
            continue
        pages.append((rel, p))

    groups = defaultdict(list)
    fps = {}
    for rel, p in pages:
        raw = read(p)
        soup = BeautifulSoup(raw, "lxml")
        fp = f"[{category(rel)}] " + fingerprint(raw, soup, coarse=not a.fine)
        fps[rel] = fp
        groups[fp].append(rel)

    order = sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0]))
    chunks = [
        "CFB OLD SITE: CODE BUNDLE",
        f"{len(pages)} live pages, {len(groups)} distinct layout fingerprints.",
        "Boilerplate (rollover functions, analytics, preload lists, navigation rows) has been removed from the",
        "example sources below; everything else is verbatim.",
        "",
        "=" * 78, "PART 1: THE SHARED STYLESHEET", "=" * 78,
    ]
    css = root / "css" / "mainstyle.css"
    chunks.append(read(css) if css.exists() else "(css/mainstyle.css not found)")

    chunks += ["", "=" * 78, "PART 2: LAYOUT FINGERPRINTS (largest group first)", "=" * 78]
    for i, (fp, members) in enumerate(order, 1):
        ex = sorted(members, key=lambda m: (len(m), m))[0]
        chunks.append(f"\nGROUP {i}: {len(members)} page(s)\n  {fp}\n  example: {ex}")
        if len(members) > 1:
            others = ", ".join(members[1:9]) + (f", ... (+{len(members) - 9})" if len(members) > 9 else "")
            chunks.append(f"  others: {others}")

    chunks += ["", "=" * 78, "PART 3: ONE EXAMPLE PAGE PER LAYOUT (boilerplate removed)", "=" * 78]
    size = sum(len(c) for c in chunks)
    included, skipped = 0, []
    for i, (fp, members) in enumerate(order, 1):
        for rel in sorted(members, key=lambda m: (len(m), m))[: (1 if len(members) >= a.rare else len(members))]:
            body = trim(read(root / rel))
            head = f"\n\n----- GROUP {i} ({len(members)} pages) : {rel} -----\n"
            if size + len(body) + len(head) > a.max_chars:
                skipped.append(rel)
                continue
            chunks.append(head + body)
            size += len(body) + len(head)
            included += 1

    if skipped:
        chunks.append("\n\n(Not included to stay within the size limit: "
                      + ", ".join(skipped[:40]) + (f", ... (+{len(skipped) - 40})" if len(skipped) > 40 else "") + ")")
    text = "\n".join(chunks) + "\n"
    outp.write_text(text, encoding="utf-8")

    print(f"Wrote {outp}  ({len(text) / 1000:.0f} KB)")
    print(f"  live pages: {len(pages)}   distinct layouts: {len(groups)}   full examples included: {included}")
    if skipped:
        print(f"  left out for size: {len(skipped)} (raise --max-chars to include them)")
    print("\nBiggest layout groups:")
    for i, (fp, members) in enumerate(order[:10], 1):
        print(f"  {len(members):4} pages  group {i}: {fp[:92]}")
    cat = Counter(category(rel) for rel, _ in pages)
    print("\nPages by kind:", dict(cat))


if __name__ == "__main__":
    main()
