#!/usr/bin/env python3
"""
inventory_site.py

Reads the OLD hand-coded CFB Key site (your local copy, cfbkey_local) and
writes a set of CSV files that become the first draft of the new spreadsheet.
It changes NOTHING in the site folder. Re-run it any time.

USAGE (inside the conda 'video' environment):

  conda install -c conda-forge beautifulsoup4 lxml      # one time
  python inventory_site.py --root "$ZK/cfbkey_local" --out "$ZK/inventory"

OUTPUT FILES (in --out; open them in Excel or Google Sheets):

  organism_pages.csv   one row per page under html/Organisms: name, taxonomy,
                       path folders, image/figure/video counts, flags
  organism_sections.csv  long format: page, section header, item number, text
                       (Distinguishing Characteristics, Habitats, References...)
  organism_images.csv  every content image: file, alt text, size, linked large
                       version, whether the file exists
  figure_links.csv     the popup "detail figure" links (text that opens a picture)
  species_index.csv    what html/species.html lists (the authoritative species list)
  key_steps.csv        the identification key: one row per choice on each step page
                       (includes the entry pages begin.html and rotifers.htm)
  key_step_extra_images.csv  non-choice images on step pages (often images of text)
  key_step_links.csv   help/method links on step pages (e.g. dissection instructions)
  video_wrappers.csv   media wrapper pages: which .mov each embeds + its caption
  image_maps.csv       pictures with clickable regions, one row per region
  anatomy.csv          the anatomy diagrams: one row per label, with the picture it shows
  groups.csv           the Groups pages as an outline (group name, level, where it leads)
  duplicate_videos.csv identical clips stored under several names, with a suggested keeper
  videos.csv           every .mov: video_id, caption, who references it, orphans
  species_videos.csv   which organism page links to which video (many-to-many)
  movies_catalog.csv   html/movies.html, the real video index: section, title,
                       description, '*New*' flag, target
  page_video_links.csv video links found on NON-organism pages (anatomy, methods...)
  popup_pages.csv      popups that open some other page (glossary, anatomy...)
  other_pages.csv      remaining pages (about, history, anatomy...) for the
                       static-pages backlog
  unreferenced_files.csv  files on disk that NO page links to (candidates to leave
                       out of the new repo), with a check for a web-size twin
  page_status.csv      every page marked live (reachable from home) or dead
  problems.csv         things to look at by hand (status column: live/dead) (missing files, case mismatches,
                       pages the parser could not read, ...)

NOTES
  * Case matters on GitHub Pages and archive.org, but not on a Mac or (often)
    an old Windows server. Any link whose capitalization does not match the
    real file name is flagged 'case_mismatch' - those links would BREAK on a
    new host.
  * video_id uses the same naming as upload_to_ia.py (path without extension,
    spaces replaced with _), so it matches the video manifest.
"""

import argparse
import csv
import hashlib
import posixpath
import re
import sys
from collections import defaultdict
from pathlib import Path
from urllib.parse import unquote

try:
    from bs4 import BeautifulSoup, NavigableString, Comment
except ImportError:
    sys.exit("Needs beautifulsoup4:  conda install -c conda-forge beautifulsoup4 lxml")

NAV_IMG = re.compile(r"(images/buttons/|indexmain|groupbrackets|images/group/|shim)", re.I)
NAV_PAGE = re.compile(r"(^|/)(index|instructions|begin|groups|species|anatomy)\.html?$", re.I)
IMG_EXT = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tif", ".tiff")
RED = re.compile(r"^#?ff0000$", re.I)
HEADER_BG = {"#000000": 1, "#cccccc": 2}


def clean_remote(rel: str) -> str:
    """Same rule as upload_to_ia.py so video_ids match the video manifest."""
    return re.sub(r"[^A-Za-z0-9._/\-]", "_", rel.replace(" ", "_"))


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").replace("\xa0", " ")).strip()


def text_of(el) -> str:
    return norm(el.get_text("")) if el is not None else ""


def own_text(li) -> str:
    """Text of a list item, excluding any nested lists."""
    parts = []
    for c in li.children:
        if getattr(c, "name", None) in ("ul", "ol"):
            continue
        parts.append(c.get_text("") if hasattr(c, "get_text") else str(c))
    return norm("".join(parts))


def _wrap(mark, inner):
    core = inner.strip()
    if not core or not re.search(r"\w", core):   # nothing to emphasise (e.g. a bold period)
        return inner
    lead = inner[:len(inner) - len(inner.lstrip())]
    trail = inner[len(inner.rstrip()):]
    return f"{lead}{mark}{core}{mark}{trail}"


def _link(a, inner, site, page_rel):
    """[text](kind:path) for popups / internal links, [text](url) for external."""
    if not inner.strip():
        return inner
    oc = re.findall(r"MM_openBrWindow\(\s*'([^']+)'", a.get("onclick") or "")
    cand = oc[0] if oc else (a.get("href") or "")
    if not cand or cand.startswith("#"):
        return inner
    r = site.resolve(page_rel, cand)
    if r is None:
        return f"[{inner.strip()}]({cand})"
    low = r[0].lower()
    kind = "fig" if low.endswith(IMG_EXT) else ("video" if (low.endswith(".mov") or low.startswith("media/")) else "page")
    return f"[{inner.strip()}]({kind}:{r[0]})"


def colour_name(value):
    """A CSS colour -> red / blue / green / orange, or '' for black, grey and anything else."""
    v = (value or "").strip().lower()
    named = {"red": "red", "blue": "blue", "green": "green", "orange": "orange"}
    if v in named:
        return named[v]
    m = re.fullmatch(r"#?([0-9a-f]{3}|[0-9a-f]{6})", v)
    if not m:
        return ""
    h = m.group(1)
    if len(h) == 3:
        h = "".join(c * 2 for c in h)
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    if r > 190 and g < 110 and b < 110:
        return "red"
    if b > 150 and r < 110 and g < 110:
        return "blue"
    if g > 90 and r < 110 and b < 110:
        return "green"
    if r > 190 and 110 <= g <= 190 and b < 90:
        return "orange"
    return ""


def css_class_colours(soup):
    """{class name: colour name} read from the page's own <style> block (and the shared .red/.blue/.green)."""
    out = {"red": "red", "blue": "blue", "green": "green"}
    for st in soup.find_all("style"):
        for name, body in re.findall(r"\.([\w-]+)\s*\{([^}]*)\}", st.get_text()):
            m = re.search(r"(?<![-\w])color\s*:\s*([#\w]+)", body)
            if m and colour_name(m.group(1)):
                out[name] = colour_name(m.group(1))
    return out


def _root(el):
    while getattr(el, "parent", None) is not None:
        el = el.parent
    return el


def hover_swap_target(img, anchor=None):
    """The picture a hover swaps in. The handler may sit on the <img>, on its <a>,
    or on an enclosing cell, and may name the image it swaps."""
    name = (img.get("name") or img.get("id") or "").strip()
    cands = []
    for t in (img, anchor, img.find_parent("td"), img.find_parent("div")):
        if t is None:
            continue
        for ev in ("onmouseover", "onmouseout", "onclick"):
            v = t.get(ev) or ""
            if ev == "onmouseover" or not cands:
                cands += [(m.group(1), m.group(2)) for m in
                          re.finditer(r"MM_swapImage\('([^']*)','',\s*'([^']+)'", v)]
        if cands:
            break
    if not cands:
        return ""
    for target, src in cands:                       # prefer the one that names this picture
        if name and target == name:
            return src
    return cands[0][1]


def tag_colour(c):
    """The colour name a <font>/<span> asks for, or ''."""
    if c.name == "font" and c.get("color"):
        return colour_name(c["color"])
    if c.name in ("span", "font", "b", "strong", "em", "i"):
        m = re.search(r"(?<![-\w])color\s*:\s*([#\w]+)", c.get("style") or "")
        if m:
            return colour_name(m.group(1))
        table = getattr(_root(c), "css_colours", {})
        for cls in (c.get("class") or []):
            if cls in table:
                return table[cls]
    return ""


def md_inline(el, site, page_rel, skip_lists, links, colours=False):
    parts = []
    for c in el.children:
        if isinstance(c, Comment):
            continue
        if isinstance(c, NavigableString):
            parts.append(str(c))
            continue
        name = c.name
        if name in ("script", "style") or (skip_lists and name in ("ul", "ol")):
            continue
        inner = md_inline(c, site, page_rel, skip_lists, links, colours)
        col = tag_colour(c) if colours else ""
        if col and inner.strip() and name not in ("em", "i", "strong", "b"):
            lead = inner[:len(inner) - len(inner.lstrip())]
            trail = inner[len(inner.rstrip()):]
            parts.append(f"{lead}{{{col}|{inner.strip()}}}{trail}")
        elif name in ("em", "i"):
            parts.append(_wrap("*", inner))
        elif name in ("strong", "b"):
            parts.append(_wrap("**", inner))
        elif name in ("sup", "sub"):
            parts.append(f"<{name}>{inner.strip()}</{name}>" if inner.strip() else "")
        elif name == "a" and links:
            parts.append(_link(c, inner, site, page_rel))
        else:
            parts.append(inner)
    return "".join(parts)


def md_of(el, site, page_rel, skip_lists=False, links=True, colours=False):
    """Text with light Markdown: *italic*, **bold**, <sup>/<sub>, [text](kind:path), {red|coloured text}."""
    if el is None:
        return ""
    return norm(md_inline(el, site, page_rel, skip_lists, links, colours))


def block_texts(cell, site, page_rel, colours=False):
    """Separate paragraphs or list items stay separate; otherwise one text."""
    ps = [p for p in cell.find_all("p") if not p.find("p")]
    if len(ps) < 2:
        ps = [li for li in cell.find_all("li") if not li.find("li")]
    if len(ps) >= 2:
        out = [md_of(x, site, page_rel, skip_lists=True, links=False, colours=colours) for x in ps]
    else:
        out = [md_of(cell, site, page_rel, links=False, colours=colours)]
    return [t for t in out if t]


def load_soup(path: Path):
    txt = path.read_bytes().decode("cp1252", errors="replace")
    try:
        soup = BeautifulSoup(txt, "lxml")
    except Exception:
        soup = BeautifulSoup(txt, "html.parser")
    for br in soup.find_all("br"):
        br.replace_with(" ")
    soup.css_colours = css_class_colours(soup)
    for s in soup.find_all(["script", "style"]):
        s.decompose()
    return soup


class Site:
    def __init__(self, root: Path):
        self.root = root
        self.lower = {}
        for p in root.rglob("*"):
            if p.is_file() and not p.name.startswith("._") and p.name != ".DS_Store":
                rel = p.relative_to(root).as_posix()
                self.lower[rel.lower()] = rel

    def resolve(self, page_rel: str, target: str):
        """-> (rel_path, exists, case_mismatch) or None if not a local file link."""
        t = unquote((target or "").strip().replace("\\", "/"))
        t = t.split("#")[0].split("?")[0]
        if not t or re.match(r"^[a-z][a-z0-9+.-]*:", t, re.I) or t.startswith("//"):
            return None
        joined = posixpath.normpath(posixpath.join(posixpath.dirname(page_rel), t))
        if joined.startswith(".."):
            return (joined, False, False)
        actual = self.lower.get(joined.lower())
        if actual is None:
            return (joined, False, False)
        return (actual, True, actual != joined)


def taxonomy_pairs(soup):
    pairs = []
    for tr in soup.find_all("tr"):
        if tr.find("table"):
            continue
        tds = tr.find_all("td", recursive=False)
        if len(tds) == 2 and tds[0].find("font", attrs={"color": RED}):
            label, value = text_of(tds[0]), text_of(tds[1])
            if label and value:
                pairs.append((label, value))
    return pairs


def header_level(tr):
    """1 (black bar) / 2 (grey bar) / 0 (not a header) and header text."""
    tds = tr.find_all("td", recursive=False)
    if not tds:
        return 0, ""
    first = tds[0]
    if first.find(class_="textsizeheadline"):
        return 0, ""  # the species-name bar, not a section header
    cls = first.get("class") or []
    # Priority: the cell's own CSS class, then the cell's bgcolor, then the row's
    # bgcolor (rows sometimes carry a stale bgcolor that the cell overrides).
    if "headerblack" in cls:
        level = 1
    elif "headergray" in cls or "headergrey" in cls:
        level = 2
    elif first.get("bgcolor"):
        level = HEADER_BG.get(first["bgcolor"].lower(), 0)
    else:
        level = HEADER_BG.get((tr.get("bgcolor") or "").lower(), 0)
    txt = text_of(first)
    if level and txt:
        return level, txt
    return 0, ""


def onclick_targets(tag):
    oc = tag.get("onclick") or ""
    return re.findall(r"MM_openBrWindow\(\s*'([^']+)'", oc)


class Collector:
    def __init__(self, site):
        self.site = site
        self.problems = []
        self.organism_pages, self.sections, self.images, self.figs = [], [], [], []
        self.species_videos_raw = []   # (page, label, target, resolved...)
        self.key_steps, self.wrappers = [], []
        self.movies_catalog_raw = []
        self.other_pages = []
        self.popups = []
        self.step_extra_images = []
        self.step_links = []

    def problem(self, page, kind, detail=""):
        self.problems.append({"page": page, "kind": kind, "detail": detail})

    # ---------------- organism pages ----------------
    def organism(self, rel, soup, in_index, index_name):
        title = text_of(soup.title)
        name_el = soup.select_one(".textsizeheadline")
        name = text_of(name_el) or title
        tax = taxonomy_pairs(soup)
        known = ["Phylum", "Subphylum", "Class", "Subclass", "Order", "Suborder", "Family"]
        taxd = {k: "" for k in known}
        other = []
        for label, value in tax:
            if label in taxd and not taxd[label]:
                taxd[label] = value
            else:
                other.append(f"{label}={value}")

        parts = rel.split("/")
        path_dirs = "/".join(parts[2:-1])  # between html/Organisms and the file

        # images
        n_img = n_noalt = n_nameonly = 0
        for i, img in enumerate(soup.find_all("img"), 1):
            src = img.get("src") or ""
            if NAV_IMG.search(src):
                continue
            r = self.site.resolve(rel, src)
            a = img.find_parent("a")
            link = ""
            if a is not None and a.get("href"):
                lr = self.site.resolve(rel, a["href"])
                link = lr[0] if lr else a["href"]
            alt = norm(img.get("alt"))
            n_img += 1
            if not alt:
                n_noalt += 1
            elif alt.lower() == name.lower():
                n_nameonly += 1
            sw = hover_swap_target(img, a)
            hov = ""
            if sw:
                hr = self.site.resolve(rel, sw)
                hov = hr[0] if (hr and hr[1]) else ""
            usemap = (img.get("usemap") or "").lstrip("#")
            self.images.append({
                "page": rel, "order": n_img, "src": r[0] if r else src,
                "exists": r[1] if r else "", "case_mismatch": r[2] if r else "",
                "alt": alt, "width": img.get("width", ""), "height": img.get("height", ""),
                "links_to": link, "hover_image": hov, "usemap": usemap})
            if r and not r[1]:
                self.problem(rel, "missing_image", src)
            if r and r[2]:
                self.problem(rel, "case_mismatch", src)

        # sections (long format)
        cur1 = cur2 = ""
        counter = defaultdict(int)
        n_sections = 0
        for tr in soup.find_all("tr"):
            if tr.find("table"):
                continue
            lvl, htxt = header_level(tr)
            if lvl == 1:
                cur1, cur2 = htxt, ""
                n_sections += 1
                continue
            if lvl == 2:
                cur2 = htxt
                n_sections += 1
                continue
            tds = tr.find_all("td", recursive=False)
            if len(tds) != 1 or tr.find(["img"]) and not text_of(tr):
                continue
            cell = tds[0]
            items = [md_of(li, self.site, rel, skip_lists=True) for li in cell.find_all("li")]
            if not items:
                items = [md_of(p, self.site, rel) for p in cell.find_all("p")]
            if not items:
                items = [md_of(cell, self.site, rel)]
            for it in items:
                if not it:
                    continue
                key = (cur1, cur2)
                counter[key] += 1
                self.sections.append({"page": rel, "section": cur1, "subsection": cur2,
                                      "item_no": counter[key], "text": it})

        if not counter:                              # no table rows at all: keep the page's own text
            body = soup.body or soup
            if not body.find("table"):
                txt_ = md_of(body, self.site, rel)
                if txt_:
                    counter[("", "")] = 1
                    self.sections.append({"page": rel, "section": "", "subsection": "",
                                          "item_no": 1, "text": txt_})

        # popup figure links + video links
        n_fig = n_vid = 0
        for tag in soup.find_all(True):
            targets = onclick_targets(tag)
            href = tag.get("href", "") if tag.name == "a" else ""
            if href.lower().endswith(".mov"):
                targets = targets + [href]
            for t in targets:
                low = t.lower().split("?")[0]
                r = self.site.resolve(rel, t)
                if r is None:
                    continue
                label = text_of(tag)[:200]
                if low.endswith(IMG_EXT):
                    n_fig += 1
                    self.figs.append({"page": rel, "label": label, "image": r[0],
                                      "exists": r[1], "case_mismatch": r[2]})
                    if not r[1]:
                        self.problem(rel, "missing_figure", t)
                    elif r[2]:
                        self.problem(rel, "case_mismatch", t)
                elif low.endswith(".htm") or low.endswith(".html"):
                    if not r[0].lower().startswith("media/"):
                        # a popup to some other page (glossary, anatomy...), not a video
                        self.popups.append({"page": rel, "label": label, "target": r[0],
                                            "exists": r[1], "case_mismatch": r[2]})
                        continue
                    n_vid += 1
                    self.species_videos_raw.append({"page": rel, "label": label, "target": t,
                                                    "resolved": r[0], "exists": r[1],
                                                    "case_mismatch": r[2]})
                    if not r[1]:
                        self.problem(rel, "missing_video_link", t)
                    elif r[2]:
                        self.problem(rel, "case_mismatch", t)
                elif low.endswith(".mov"):
                    n_vid += 1
                    self.species_videos_raw.append({"page": rel, "label": label, "target": t,
                                                    "resolved": r[0], "exists": r[1],
                                                    "case_mismatch": r[2]})
                    if not r[1]:
                        self.problem(rel, "missing_video_link", t)
                    elif r[2]:
                        self.problem(rel, "case_mismatch", t)

        if not tax:
            self.problem(rel, "no_taxonomy_table_found", "check by hand (genus/index page?)")
        if n_sections == 0:
            self.problem(rel, "no_section_headers_found", "")
        flags = []
        if re.search(r"(_old|\bold\b|backup|copy)", rel, re.I):
            flags.append("looks_like_backup")
        if not in_index:
            flags.append("not_in_species_index")
        self.organism_pages.append({
            "page": rel, "title": title, "name": name, "index_name": index_name,
            **taxd, "tax_other": "; ".join(other), "path_dirs": path_dirs,
            "n_images": n_img, "n_images_no_alt": n_noalt, "n_images_alt_is_name_only": n_nameonly,
            "n_figure_links": n_fig, "n_video_links": n_vid, "n_sections": n_sections,
            "flags": "; ".join(flags)})

    # ---------------- key steps ----------------
    def key_step(self, rel, soup, organism_set, group=None):
        """Handles side-by-side choices, stacked choices (one per row), grids, and
        choices shown with two images linking to the same page."""
        parts = rel.split("/")
        if group is None:
            group = parts[2] if len(parts) > 3 else ""
        step_id = Path(rel).stem
        prompt = ""
        for el in soup.find_all(["strong", "b", "div", "p", "td"]):
            if el.find("table"):
                continue
            t = text_of(el)
            if re.search(r"\bclick\b", t, re.I) and len(t) < 200:
                prompt = t
                break
        tax = "; ".join(f"{a}={b}" for a, b in taxonomy_pairs(soup))

        choices = []
        for a in soup.find_all("a", href=True):
            img = a.find("img")
            if img is None or NAV_IMG.search(img.get("src", "")):
                continue
            if NAV_PAGE.search(a["href"]):
                continue
            choices.append((a, img))
        if not choices:
            self.problem(rel, "no_choices_found", "key step with no image links - check by hand")
            return

        # group anchors: same table row AND same destination = one choice with several images
        groups = []
        for a, img in choices:
            tr, td = a.find_parent("tr"), a.find_parent("td")
            res = self.site.resolve(rel, a["href"])
            tkey = res[0] if res else a["href"]
            row_tds = tr.find_all("td", recursive=False) if tr is not None else []
            col = next((i for i, x in enumerate(row_tds) if x is td), 0)
            for g in groups:
                if g["tr"] is tr and g["tkey"] == tkey:
                    g["anchors"].append((a, img))
                    break
            else:
                groups.append({"tr": tr, "tkey": tkey, "col": col, "res": res, "anchors": [(a, img)]})
        choice_trs = {id(g["tr"]) for g in groups}
        choice_imgs = {id(img) for _, img in choices}

        groups_per_row = defaultdict(int)
        for g in groups:
            groups_per_row[id(g["tr"])] += 1
        choice_anchor_ids = {id(a) for a, _ in choices}

        # text links that are NOT choices (help / example links): listed separately, so keep them out of the choice text
        help_texts = set()
        for ha in soup.find_all("a"):
            if ha.find("img") or id(ha) in choice_anchor_ids:
                continue
            ht = re.sub(r"[\s*]+", " ", md_of(ha, self.site, rel, links=False)).strip().lower()
            if ht:
                help_texts.add(ht)

        def usable(cell):
            """A cell that holds text for a choice. A cell that holds only a video icon and its short caption
            ("Intermediate size pecten") is not choice text; long text that merely shares a cell with an icon is."""
            if cell is None:
                return []
            out_ = []
            for t in block_texts(cell, self.site, rel, colours=True):
                plain_t = re.sub(r"[\s*]+", " ", re.sub(r"\{\w+\|([^}]*)\}", r"\1", t)).strip().lower()
                if re.search(r"\bclick\b", t, re.I) or plain_t in help_texts:
                    continue
                out_.append(t)
            if cell.find("img", attrs={"onclick": True}):
                total = sum(len(re.sub(r"\{\w+\|([^}]*)\}", r"\1", t)) for t in out_)
                if total < 45:
                    return []
            return out_

        for n, g in enumerate(groups, 1):
            texts = []
            # same-row text: only when this row holds a single choice (otherwise the
            # neighbouring cells are the other choices)
            if g["tr"] is not None and groups_per_row[id(g["tr"])] == 1:
                for td in g["tr"].find_all("td", recursive=False):
                    if any(id(a) in choice_anchor_ids for a in td.find_all("a")):
                        continue
                    texts.extend(usable(td))
            if g["tr"] is not None:
                for sib in g["tr"].find_next_siblings("tr"):
                    if id(sib) in choice_trs or header_level(sib)[0]:
                        break
                    tds = sib.find_all("td", recursive=False)
                    if not tds:
                        continue
                    if len(tds) == 1:
                        cell = tds[0]
                    elif g["col"] < len(tds):
                        cell = tds[g["col"]]
                    else:
                        cell = None
                    got = usable(cell)
                    if not got and len(tds) > 1 and groups_per_row[id(g["tr"])] == 1:
                        # stacked layout where a cell above spans several rows: the text sits one column over
                        for other in tds:
                            if other is not cell:
                                got = usable(other)
                                if got:
                                    break
                    texts.extend(got)
            a0, img0 = g["anchors"][0]
            ir = [self.site.resolve(rel, im.get("src", "")) for _, im in g["anchors"]]
            images = [x[0] if x else im.get("src", "") for x, (_, im) in zip(ir, g["anchors"])]
            hover = ""
            sw = hover_swap_target(img0, a0)
            if sw:
                hr = self.site.resolve(rel, sw)
                hover = hr[0] if hr else sw
            res = g["res"]
            target = res[0] if res else a0["href"]
            if res is None:
                ttype = "external"
            elif target.startswith("html/choices/") or target.lower() == "html/rotifers.htm":
                ttype = "key_step"
            elif target.startswith("html/idpages/"):
                ttype = "id_page"
            elif target in organism_set:
                ttype = "organism_page"
            elif target.startswith("html/Organisms/"):
                ttype = "organism_page_unlisted"
            else:
                ttype = "other"
            if res and not res[1]:
                self.problem(rel, "missing_target", a0["href"])
            elif res and res[2]:
                self.problem(rel, "case_mismatch", a0["href"])
            self.key_steps.append({
                "step_path": rel, "key_group": group, "step_id": step_id, "choice_no": n,
                "choice_text": " | ".join(texts),
                "image": images[0], "extra_images": "; ".join(images[1:]),
                "image_exists": ir[0][1] if ir[0] else "", "image_hover": hover,
                "alt_in_old_site": norm(img0.get("alt")),
                "goes_to": target, "goes_to_type": ttype, "target_exists": res[1] if res else "",
                "step_prompt": prompt, "taxonomy_at_step": tax})
            if not texts:
                self.problem(rel, "choice_without_text", f"choice {n} (image only: a label must be written)")

        # text links on the step that are not image choices (help / method links)
        for a in soup.find_all("a"):
            if a.find("img"):
                continue
            txt = md_of(a, self.site, rel, links=False)
            oc = [t for tag in [a] + a.find_all(True) for t in onclick_targets(tag)]
            cand = oc[0] if oc else (a.get("href") or "")
            if not txt or not cand or cand.startswith("#") or NAV_PAGE.search(cand):
                continue
            r = self.site.resolve(rel, cand)
            self.step_links.append({"step_path": rel, "text": txt,
                                    "target": r[0] if r else cand, "exists": r[1] if r else ""})

        # images on the step that are not choices (often images OF TEXT: needs real text)
        for img in soup.find_all("img"):
            src = img.get("src", "")
            if NAV_IMG.search(src) or id(img) in choice_imgs:
                continue
            if img.get("onclick") or re.search(r"(^|/)VF[^/]*\.gif$", src, re.I) or \
                    ((to_i := re.sub(r"\D", "", str(img.get("width", "")))) and int(to_i) <= 40):
                continue                                   # small clickable icons (video buttons)
            r = self.site.resolve(rel, src)
            self.step_extra_images.append({
                "step_path": rel, "image": r[0] if r else src, "alt": norm(img.get("alt")),
                "width": img.get("width", ""), "height": img.get("height", "")})

    # ---------------- video wrapper pages ----------------
    def wrapper(self, rel, soup):
        mov = None
        for tag in soup.find_all(["embed", "object", "param", "video", "source", "a"]):
            for attr in ("src", "value", "href", "data"):
                v = tag.get(attr) or ""
                if v.lower().split("?")[0].endswith(".mov"):
                    mov = (tag, v)
                    break
            if mov:
                break
        body = soup.body or soup
        caption = md_of(body, self.site, rel, links=False)
        if not mov:
            self.problem(rel, "wrapper_without_mov", caption[:100])
            return
        tag, v = mov
        r = self.site.resolve(rel, v)
        self.wrappers.append({
            "wrapper_path": rel, "mov_path": r[0] if r else v, "mov_exists": r[1] if r else "",
            "case_mismatch": r[2] if r else "", "embed_width": tag.get("width", ""),
            "embed_height": tag.get("height", ""), "caption": caption})
        if r and not r[1]:
            self.problem(rel, "missing_mov", v)
        elif r and r[2]:
            self.problem(rel, "case_mismatch", v)

    # ---------------- movies.html catalog ----------------
    def movies_page(self, rel, soup):
        """The real video index: sections (Organisms Videos, Virtual Focus, Sampling
        Methods, Dissection Demonstration), each item = title + description + link."""
        section = ""
        for tr in soup.find_all("tr"):
            if tr.find("table"):
                continue
            tds = tr.find_all("td", recursive=False)
            if not tds:
                continue
            first_cls = tds[0].get("class") or []
            if not tr.find("li") and text_of(tds[0]) and (header_level(tr)[0] or "textsize14" in first_cls):
                section = text_of(tds[0])
                continue
            for li in tr.find_all("li"):
                targets = []
                for tag in [li] + li.find_all(True):
                    targets += onclick_targets(tag)
                    if tag.name == "a" and tag.get("href", "").lower().endswith((".html", ".htm", ".mov")):
                        targets.append(tag["href"])
                cands = [(t, self.site.resolve(rel, t)) for t in targets]
                cands = [(t, r) for t, r in cands if r is not None]
                if not cands:
                    continue
                t, r = next(((t, r) for t, r in cands if r[1]), cands[0])
                is_new = bool(re.search(r"\*new\*", li.get_text(""), re.I))
                for tn in li.find_all(string=re.compile(r"\*New\*", re.I)):
                    tn.replace_with(re.sub(r"\*New\*", "", str(tn), flags=re.I))
                label = md_of(li, self.site, rel, skip_lists=True, links=False)
                self.movies_catalog_raw.append({
                    "category": section, "index_page": rel, "label": label, "is_new": is_new,
                    "target": r[0], "target_exists": r[1], "case_mismatch": r[2]})
                if not r[1]:
                    self.problem(rel, "missing_video_link", t)
                elif r[2]:
                    self.problem(rel, "case_mismatch", t)


RANK_ORDER = ["Phylum", "Class", "Subclass", "Superorder", "Order", "Suborder", "Family", "Subfamily", "Genus"]
RANK_RE = re.compile(r"^(Phylum|Class|Subclass|Superorder|Order|Suborder|Family|Subfamily|Genus)\s+(.+)$")
GROUPS_RE = re.compile(r"^html/groups(?:_(\w+))?\.html?$", re.I)
RANKED_PAGES = {"cladocera", "copepoda", "rotifera", "arthropoda"}   # full outlines: nest by taxonomic rank


def parse_image_maps(site, root, pages):
    """Pictures with clickable regions: one row per region, so they can become plain links."""
    rows = []
    for rel in pages:
        soup = load_soup(root / rel)
        for mp in soup.find_all("map"):
            name = (mp.get("name") or mp.get("id") or "").strip()
            for i, ar in enumerate(mp.find_all("area"), 1):
                tgt = ar.get("href") or ""
                oc = onclick_targets(ar)
                if oc:
                    tgt = oc[0]
                r = site.resolve(rel, tgt)
                rows.append({"page": rel, "map_name": name, "seq": i,
                             "label": norm(ar.get("alt") or ar.get("title") or ""),
                             "shape": ar.get("shape", ""), "coords": ar.get("coords", ""),
                             "goes_to": r[0] if r else tgt, "target_exists": r[1] if r else ""})
    return rows


def parse_anatomy(site, root):
    """The anatomy pages: a diagram plus labels that swap which parts are marked on it."""
    rows = []
    for rel in sorted(site.lower.values()):
        if not re.match(r"^html/anatomy/[^/]+/[^/]+\.html?$", rel, re.I):
            continue
        soup = load_soup(root / rel)
        main = next((i for i in soup.find_all("img") if i.get("name") and not NAV_IMG.search(i.get("src", ""))), None)
        base = site.resolve(rel, main.get("src", "")) if main is not None else None
        group = rel.split("/")[2]
        intro = ""
        for td in soup.find_all(["td", "div"]):
            t = text_of(td)
            if re.search(r"\bclick\b", t, re.I) and len(t) < 200:
                intro = t
                break
        seen, seq = set(), 0
        for tag in soup.find_all(True):
            m = re.search(r"MM_swapImage\('([^']*)','',\s*'([^']+)'", tag.get("onclick") or "")
            if not m:
                continue
            label = md_of(tag, site, rel, links=False)
            if not label or label in seen:
                continue
            r = site.resolve(rel, m.group(2))
            if r is None or not r[1]:
                continue
            seen.add(label)
            seq += 1
            rows.append({"page": rel, "group": group, "seq": seq, "label": label,
                         "image": r[0], "is_base": r[0] == (base[0] if base else None),
                         "base_image": base[0] if base else "", "intro": intro})
    return rows


def parse_groups(site, root):
    """The old Groups pages: an outline of group names (each linking into the key) and species."""
    rows = []
    for rel in sorted(site.lower.values()):
        m = GROUPS_RE.match(rel)
        if not m:
            continue
        page = (m.group(1) or "overview").lower()
        soup = load_soup(root / rel)
        stack, seq = [], 0
        for a in soup.find_all("a", href=True):
            im = a.find("img")
            if (im is not None and NAV_IMG.search(im.get("src", ""))) or NAV_PAGE.search(a["href"]) \
                    or re.search(r"(^|/)groups\w*\.html?$", a["href"], re.I):
                continue
            txt = text_of(a)
            if not txt:
                continue
            mm = RANK_RE.match(txt)
            depth = len(a.find_parents(["ul", "ol"]))
            if page in RANKED_PAGES:
                if mm:
                    ri = RANK_ORDER.index(mm.group(1))
                    while stack and stack[-1] >= ri:
                        stack.pop()
                    level = len(stack)
                    stack.append(ri)
                else:
                    level = len(stack)
            else:
                level = depth
            r = site.resolve(rel, a["href"])
            target, exists = (r[0], r[1]) if r else (a["href"], "")
            low = target.lower()
            if low.startswith("html/choices/") or low == "html/rotifers.htm":
                ttype = "key_step"
            elif low.startswith("html/organisms/"):
                ttype = "organism_page"
            elif low.startswith("html/idpages/"):
                ttype = "id_page"
            else:
                ttype = "other"
            seq += 1
            rows.append({"group_page": page, "seq": seq, "level": level,
                         "rank": mm.group(1) if mm else ("Species" if page in RANKED_PAGES else "Group"),
                         "name": txt, "goes_to": target, "goes_to_type": ttype, "target_exists": exists})
    return rows


def write_csv(path: Path, rows, fields=None):
    if not rows and not fields:
        path.write_text("", encoding="utf-8")
        return
    fields = fields or list(rows[0].keys())
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader()
        w.writerows(rows)


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", required=True, help="the cfbkey_local folder (contains html/, media/, ...)")
    ap.add_argument("--out", default="inventory")
    args = ap.parse_args()

    root = Path(args.root).expanduser().resolve()
    if not (root / "html").is_dir():
        sys.exit(f"ERROR: {root} has no html/ folder. Point --root at cfbkey_local.")
    out = Path(args.out).expanduser().resolve()
    out.mkdir(parents=True, exist_ok=True)

    site = Site(root)
    C = Collector(site)

    def pages(prefix):
        return sorted(r for r in site.lower.values()
                      if r.startswith(prefix) and r.lower().endswith((".htm", ".html")))

    # species index (authoritative list)
    index_rows, index_map = [], {}
    idx_rel = "html/species.html"
    if idx_rel in site.lower.values():
        soup = load_soup(root / idx_rel)
        for a in soup.find_all("a", href=True):
            r = site.resolve(idx_rel, a["href"])
            if r and r[0].startswith("html/Organisms/"):
                nm = text_of(a)
                index_map.setdefault(r[0], nm)
                index_rows.append({"page": r[0], "display_name": nm, "exists": r[1],
                                   "case_mismatch": r[2]})
                if not r[1]:
                    C.problem(idx_rel, "missing_target", a["href"])
    else:
        C.problem(idx_rel, "species_index_not_found", "")

    org_pages = pages("html/Organisms/")
    organism_set = set(org_pages)
    for rel in org_pages:
        try:
            C.organism(rel, load_soup(root / rel), rel in index_map, index_map.get(rel, ""))
        except Exception as e:
            C.problem(rel, "parse_error", repr(e))
    for p in index_map:
        if p not in organism_set:
            C.problem(idx_rel, "index_lists_page_not_found", p)

    for rel in pages("html/choices/"):
        try:
            C.key_step(rel, load_soup(root / rel), organism_set)
        except Exception as e:
            C.problem(rel, "parse_error", repr(e))
    # the key's entry pages: Cladocera/Copepoda/Rotifera/Other, then the rotifer groups
    for name in ("html/begin.html", "html/rotifers.htm"):
        rel = site.lower.get(name.lower())
        if rel:
            try:
                C.key_step(rel, load_soup(root / rel), organism_set, group="entry")
            except Exception as e:
                C.problem(rel, "parse_error", repr(e))

    for rel in pages("media/"):
        try:
            C.wrapper(rel, load_soup(root / rel))
        except Exception as e:
            C.problem(rel, "parse_error", repr(e))

    mv = site.lower.get("html/movies.html")
    if mv:
        try:
            C.movies_page(mv, load_soup(root / mv))
        except Exception as e:
            C.problem(mv, "parse_error", repr(e))
    else:
        C.problem("html/movies.html", "movies_index_not_found", "")

    handled_prefix = ("html/organisms/", "html/choices/", "media/")
    for rel in pages("html/"):
        low = rel.lower()
        if (low.startswith(handled_prefix) or low in ("html/species.html", "html/movies.html",
                                                       "html/begin.html", "html/rotifers.htm")):
            continue
        try:
            C.other_pages.append({"page": rel, "title": text_of(load_soup(root / rel).title),
                                  "size_bytes": (root / rel).stat().st_size})
        except Exception as e:
            C.problem(rel, "parse_error", repr(e))
    for rel in pages(""):
        if rel.lower().startswith("start") or "/" not in rel:
            C.other_pages.append({"page": rel, "title": "", "size_bytes": (root / rel).stat().st_size})

    # ---- video join: wrapper -> mov, and who references what ----
    wrapper_to_mov = {w["wrapper_path"]: w["mov_path"] for w in C.wrappers}
    movs = sorted(r for r in site.lower.values() if r.lower().endswith(".mov"))
    vid_of = {m: clean_remote(m[:-4]) for m in movs}

    def to_mov(target):
        if target.lower().endswith(".mov"):
            return target
        return wrapper_to_mov.get(target, "")

    ref_by = defaultdict(set)
    for w in C.wrappers:
        ref_by[w["mov_path"]].add(("wrapper", w["wrapper_path"]))

    species_videos = []
    for v in C.species_videos_raw:
        mov = to_mov(v["resolved"])
        if mov:
            ref_by[mov].add(("page", v["page"]))
        species_videos.append({"page": v["page"], "video_id": vid_of.get(mov, ""),
                               "via": "direct_mov" if v["resolved"].lower().endswith(".mov") else "wrapper",
                               "label": v["label"], "link": v["resolved"], "link_exists": v["exists"],
                               "case_mismatch": v["case_mismatch"]})
        if v["exists"] and not mov:
            C.problem(v["page"], "video_link_does_not_lead_to_mov", v["resolved"])

    movies_catalog = []
    for m in C.movies_catalog_raw:
        mov = to_mov(m["target"])
        if mov:
            ref_by[mov].add(("movies_page", m["index_page"]))
        movies_catalog.append({**m, "video_id": vid_of.get(mov, "")})

    # Scan EVERY page on the site (anatomy, rotifer, methods pages...) for links to
    # video wrappers or .mov files, so a clip used only there is not called an orphan.
    page_video_links = []
    link_labels = defaultdict(list)
    for rel in pages(""):
        if rel.startswith("media/"):
            continue
        try:
            soup = load_soup(root / rel)
        except Exception:
            continue
        seen = set()
        for tag in soup.find_all(True):
            cands = onclick_targets(tag)
            if tag.name == "a" and tag.get("href"):
                cands = cands + [tag["href"]]
            for t in cands:
                r = site.resolve(rel, t)
                if r is None:
                    continue
                mov = to_mov(r[0])
                if mov and (rel, mov) not in seen:
                    seen.add((rel, mov))
                    ref_by[mov].add(("page", rel))
                    li = tag if tag.name == "li" else tag.find_parent("li")
                    label = md_of(li if li is not None else tag, site, rel, skip_lists=True, links=False)[:200]
                    if not label and tag.name == "img":
                        label = md_of(tag.find_parent("td"), site, rel, links=False)[:200]
                    if label:
                        link_labels[mov].append(label)
                    if not rel.startswith("html/Organisms/"):
                        page_video_links.append({"page": rel, "video_id": vid_of.get(mov, ""),
                                                 "link": r[0], "label": label})

    cap_by_mov = defaultdict(list)
    for w in C.wrappers:
        cap_by_mov[w["mov_path"]].append(w["caption"])
    cat_by_mov = {}
    for m_ in movies_catalog:
        mv_ = to_mov(m_["target"])
        if mv_ and mv_ not in cat_by_mov:
            cat_by_mov[mv_] = m_
    video_rows = []
    for m in movs:
        refs = ref_by.get(m, set())
        ref_pages = {p for k, p in refs if k in ("page", "movies_page")}
        video_rows.append({
            "video_id": vid_of[m], "mov_path": m,
            "expected_mp4": clean_remote(m[:-4]) + ".mp4",
            "caption_from_wrapper": " || ".join(dict.fromkeys(cap_by_mov.get(m, []))),
            "catalog_section": cat_by_mov.get(m, {}).get("category", ""),
            "catalog_title": cat_by_mov.get(m, {}).get("label", ""),
            "catalog_is_new": cat_by_mov.get(m, {}).get("is_new", ""),
            "caption_from_link_text": " || ".join(dict.fromkeys(link_labels.get(m, [])))[:500],
            "n_wrappers": sum(1 for k, _ in refs if k == "wrapper"),
            "n_organism_pages": sum(1 for p in ref_pages if p.startswith("html/Organisms/")),
            "n_other_pages": sum(1 for p in ref_pages if not p.startswith("html/Organisms/")),
            "in_movies_catalog": any(k == "movies_page" for k, _ in refs),
            "orphan_not_referenced": not ref_pages})
        if not ref_pages:
            C.problem(m, "orphan_video", "no page on the site links to this clip" +
                      (" (a wrapper page exists but nothing links to it)" if m in cap_by_mov else ""))
        if m in cap_by_mov and not cap_by_mov[m][0]:
            C.problem(m, "wrapper_has_no_caption", "")
        if m not in cap_by_mov:
            if m in cat_by_mov or link_labels.get(m):
                C.problem(m, "caption_from_link_text_only", "check the wording; no wrapper page")
            else:
                C.problem(m, "video_without_any_caption", "needs a hand-written description")

    # ---- identical clips stored under more than one name ----
    def _md5(path):
        h = hashlib.md5()
        with open(path, "rb") as f:
            for blk in iter(lambda: f.read(1 << 20), b""):
                h.update(blk)
        return h.hexdigest()

    by_hash = defaultdict(list)
    md5_of_mov = {}
    for m in movs:
        md5_of_mov[m] = _md5(root / m)
        by_hash[md5_of_mov[m]].append(m)

    def keeper_rank(m):
        refs = ref_by.get(m, set())
        npages = len({q for k, q in refs if k in ("page", "movies_page")})
        return (-npages, 0 if m in cap_by_mov else 1, 1 if " " in posixpath.basename(m) else 0, len(m), m)

    dup_rows, dup_of, n_dup_groups = [], {}, 0
    for h, members in by_hash.items():
        if len(members) < 2:
            continue
        n_dup_groups += 1
        members.sort(key=keeper_rank)
        for i, m in enumerate(members):
            refs = ref_by.get(m, set())
            if i:
                dup_of[m] = vid_of[members[0]]
            dup_rows.append({"group": n_dup_groups, "mov_path": m,
                             "suggested_keeper": i == 0,
                             "pages_linking_to_it": len({q for k, q in refs if k in ("page", "movies_page")}),
                             "has_wrapper_page": m in cap_by_mov,
                             "size_mb": round((root / m).stat().st_size / 1e6, 2)})
    for r in video_rows:
        r["md5"] = md5_of_mov.get(r["mov_path"], "")
        r["duplicate_of"] = dup_of.get(r["mov_path"], "")

    # ---- which files does ANY page actually use? (to decide what to leave out) ----
    EXT = r"(?:jpe?g|png|gif|bmp|tiff?|psd|mov|mp4|swf|pdf|css|js|docx?|xlsx?|pptx?|txt|htm|html|wav|mp3|ico)"
    q_re = re.compile(r"""["']([^"'<>\n]+?\.""" + EXT + r""")(?:[#?][^"'<>\n]*)?["']""", re.I)
    u_re = re.compile(r"""=\s*([^\s"'<>=]+?\.""" + EXT + r""")(?=[\s>])""", re.I)
    css_re = re.compile(r"""url\(\s*["']?([^)"'\s][^)"']*?)["']?\s*\)""", re.I)
    referenced = set()
    refs_by_page = defaultdict(set)
    for rel in site.lower.values():
        low = rel.lower()
        if not low.endswith((".htm", ".html", ".css")):
            continue
        raw = (root / rel).read_bytes().decode("cp1252", errors="replace")
        toks = q_re.findall(raw) + u_re.findall(raw) + css_re.findall(raw)
        for t in toks:
            r = site.resolve(rel, t)
            if r and r[1] and r[0] != rel:
                referenced.add(r[0])
                refs_by_page[rel].add(r[0])

    # Follow links outward from the home page(s). Pages reached = LIVE; others = DEAD
    # (old copies, drafts, pages only linked from other dead pages).
    live_pages, stack = set(), []
    for name in ("Start.html", "html/index.html", "html/begin.html", "html/species.html"):
        actual = site.lower.get(name.lower())
        if actual:
            stack.append(actual)
    while stack:
        pg = stack.pop()
        if pg in live_pages:
            continue
        live_pages.add(pg)
        for t in refs_by_page.get(pg, ()):
            if t.lower().endswith((".htm", ".html")) and t not in live_pages:
                stack.append(t)
    publish = set(live_pages)
    for pg in live_pages:
        publish |= refs_by_page.get(pg, set())
    for t in list(publish):                      # assets used by stylesheets
        if t.lower().endswith(".css"):
            publish |= refs_by_page.get(t, set())

    def status_of(pg):
        return ("live" if pg in live_pages else "dead") if pg.lower().endswith((".htm", ".html")) else ""

    def stem_key(rel):
        d = posixpath.dirname(rel)
        if d.lower().endswith("/png"):
            d = d[:-4]
        stem = posixpath.splitext(posixpath.basename(rel))[0].lower()
        return (d.lower(), re.sub(r"_?\d+x\d+(_\w+)?$", "", stem))

    ref_img_keys = {stem_key(r) for r in referenced
                    if r.lower().endswith((".jpg", ".jpeg", ".gif", ".png"))}
    unref_rows = []
    sizes = {"ref": 0, "unref": 0}
    by_ext = defaultdict(lambda: [0, 0])
    for rel in sorted(site.lower.values()):
        size = (root / rel).stat().st_size
        ext = posixpath.splitext(rel)[1].lower() or "(none)"
        if rel in publish:
            sizes["ref"] += size
            continue
        sizes["unref"] += size
        by_ext[ext][0] += 1
        by_ext[ext][1] += size
        counterpart = ""
        if ext in (".png", ".bmp", ".tif", ".tiff", ".psd"):
            counterpart = "yes" if stem_key(rel) in ref_img_keys else "NO"
        unref_rows.append({"path": rel, "ext": ext, "size_bytes": size,
                           "only_linked_from_dead_pages": rel in referenced,
                           "size_mb": round(size / 1e6, 2),
                           "has_web_counterpart": counterpart,
                           "top_folder": "/".join(rel.split("/")[:2])})
    unref_rows.sort(key=lambda r: -r["size_bytes"])

    for r in C.organism_pages:
        pg = r["page"]
        r["live"] = pg in live_pages
        r["page_type"] = ("species" if r["index_name"] else ("supplementary" if pg in live_pages else "dead"))
    species_by_dir = {}
    for r in C.organism_pages:
        if r["page_type"] == "species":
            species_by_dir.setdefault(posixpath.dirname(r["page"]), r["page"])
    for r in C.organism_pages:
        r["parent_page"] = "" if r["page_type"] == "species" else species_by_dir.get(posixpath.dirname(r["page"]), "")
    supp = {r["page"] for r in C.organism_pages if r["page_type"] == "supplementary"}
    C.problems = [q for q in C.problems
                  if not (q["page"] in supp and q["kind"] in ("no_taxonomy_table_found", "no_section_headers_found"))]
    for r in C.key_steps:
        r["live"] = r["step_path"] in live_pages
    for pr in C.problems:
        pr["status"] = status_of(pr["page"])

    # ---- write everything ----
    write_csv(out / "organism_pages.csv", C.organism_pages)
    write_csv(out / "organism_sections.csv", C.sections,
              ["page", "section", "subsection", "item_no", "text"])
    write_csv(out / "organism_images.csv", C.images,
              ["page", "order", "src", "exists", "case_mismatch", "alt", "width", "height",
               "links_to", "hover_image", "usemap"])
    write_csv(out / "figure_links.csv", C.figs, ["page", "label", "image", "exists", "case_mismatch"])
    write_csv(out / "species_index.csv", index_rows, ["page", "display_name", "exists", "case_mismatch"])
    write_csv(out / "key_steps.csv", C.key_steps)
    write_csv(out / "key_step_links.csv", C.step_links, ["step_path", "text", "target", "exists"])
    write_csv(out / "key_step_extra_images.csv", C.step_extra_images,
              ["step_path", "image", "alt", "width", "height"])
    write_csv(out / "video_wrappers.csv", C.wrappers)
    write_csv(out / "videos.csv", video_rows)
    write_csv(out / "duplicate_videos.csv", dup_rows,
              ["group", "mov_path", "suggested_keeper", "pages_linking_to_it", "has_wrapper_page", "size_mb"])
    write_csv(out / "species_videos.csv", species_videos,
              ["page", "video_id", "via", "label", "link", "link_exists", "case_mismatch"])
    write_csv(out / "movies_catalog.csv", movies_catalog)
    write_csv(out / "page_video_links.csv", page_video_links, ["page", "video_id", "link", "label"])
    write_csv(out / "popup_pages.csv", C.popups, ["page", "label", "target", "exists", "case_mismatch"])
    write_csv(out / "other_pages.csv", C.other_pages, ["page", "title", "size_bytes"])
    all_html = [r for r in sorted(site.lower.values()) if r.lower().endswith((".htm", ".html"))]
    imap_rows = parse_image_maps(site, root, all_html)
    write_csv(out / "image_maps.csv", imap_rows,
              ["page", "map_name", "seq", "label", "shape", "coords", "goes_to", "target_exists"])
    anatomy_rows = parse_anatomy(site, root)
    write_csv(out / "anatomy.csv", anatomy_rows,
              ["page", "group", "seq", "label", "image", "is_base", "base_image", "intro"])
    groups_rows = parse_groups(site, root)
    write_csv(out / "groups.csv", groups_rows,
              ["group_page", "seq", "level", "rank", "name", "goes_to", "goes_to_type", "target_exists"])
    write_csv(out / "problems.csv", C.problems, ["page", "status", "kind", "detail"])
    write_csv(out / "page_status.csv",
              [{"page": r, "status": status_of(r)} for r in sorted(site.lower.values())
               if r.lower().endswith((".htm", ".html"))], ["page", "status"])
    write_csv(out / "unreferenced_files.csv", unref_rows,
              ["path", "ext", "size_bytes", "only_linked_from_dead_pages", "size_mb",
               "has_web_counterpart", "top_folder"])

    kinds = defaultdict(int)
    dead_n = 0
    for p in C.problems:
        if p["status"] == "dead":
            dead_n += 1
        else:
            kinds[p["kind"]] += 1
    print(f"Scanned {len(site.lower)} files under {root}")
    print(f"  organism pages: {len(C.organism_pages)}   (listed in species.html: {len(index_map)})")
    print(f"  key-step choices: {len(C.key_steps)} from {len({k['step_path'] for k in C.key_steps})} step pages")
    print(f"  video wrappers: {len(C.wrappers)}   .mov files: {len(movs)}   "
          f"orphans: {sum(r['orphan_not_referenced'] for r in video_rows)}")
    print(f"  image maps: {len(imap_rows)} clickable region(s) on "
          f"{len({r['page'] for r in imap_rows})} page(s)")
    print(f"  anatomy: {len(anatomy_rows)} labels on {len({a['page'] for a in anatomy_rows})} diagram page(s)")
    print(f"  groups outline: {len(groups_rows)} rows from {len({g['group_page'] for g in groups_rows})} Groups pages")
    print(f"  duplicate clips: {n_dup_groups} groups ({len(dup_of)} extra copies)")
    print(f"  images: {len(C.images)}   popup figure links: {len(C.figs)}")
    html_total = sum(1 for r in site.lower.values() if r.lower().endswith((".htm", ".html")))
    vid_mb = sum((root / f).stat().st_size for f in publish if f.lower().endswith((".mov", ".mp4"))) / 1e6
    print(f"  LIVE pages (reachable from home): {len(live_pages)} of {html_total}   "
          f"dead: {html_total - len(live_pages)}")
    print(f"  publish set: {len(publish)} files, {sizes['ref']/1e6:.0f} MB "
          f"({vid_mb:.0f} MB of that is video, which goes to Internet Archive)")
    print(f"  NOT in publish set: {len(unref_rows)} files, {sizes['unref']/1e6:.0f} MB")
    for e, (n, b) in sorted(by_ext.items(), key=lambda x: -x[1][1])[:6]:
        print(f"      unused {e}: {n} files, {b/1e6:.0f} MB")
    print(f"  problems on LIVE pages / non-page files: {sum(kinds.values())}   "
          f"(plus {dead_n} on dead pages, ignored)")
    for k, n in sorted(kinds.items(), key=lambda x: -x[1]):
        print(f"      {k}: {n}")
    print(f"\nCSV files written to: {out}")


if __name__ == "__main__":
    main()
