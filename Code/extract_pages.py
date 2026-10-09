#!/usr/bin/env python3
"""
extract_pages.py

Converts the OLD static pages (Instructions, Purpose, History, References, Contact,
Copyright, Special Features, Anatomy, Barcoding ...) into DRAFT Markdown files that
build_site.py turns into the new pages. Re-run-safe: it never overwrites a page that
already exists in the content folder unless you add --force.

  python extract_pages.py --root "$ZK/cfbkey_local" --content "$ZK/content"

For every page it writes   <content>/<slug>.md   and copies the images the page uses
to   <content>/assets/<slug>/ . Read each draft against the old page; the conversion
keeps the words, links, lists and images but not the old table layout.

Links in the drafts use four small prefixes that build_site.py understands:
    old:html/begin.html      -> wherever that old page lives in the new site
    page:contact             -> a content page by its slug
    video:media/x/y          -> the clip with that video_id (the Video page)
    assets/slug/file.jpg     -> an image copied next to the page
A link whose target does not exist yet is shown as plain text and listed in the
build report, so drafts can reference pages that are still to come.
"""

import argparse
import posixpath
import re
import shutil
import sys
from pathlib import Path
from urllib.parse import unquote

try:
    from bs4 import BeautifulSoup, Comment, NavigableString
except ImportError:
    sys.exit("Needs beautifulsoup4 and lxml:  conda install -c conda-forge beautifulsoup4 lxml")

# old file -> how it appears in the new site
PAGES = {
    "instructions.html":        dict(slug="instructions", nav="Instructions", order=10),
    "anatomy.html":             dict(slug="anatomy", nav="Anatomy", order=50, title="Anatomy"),
    "information.htm":          dict(slug="purpose", footer="Purpose", order=10, indent="yes"),
    "history.htm":              dict(slug="history", footer="History of Key", order=20),
    "references.htm":           dict(slug="references", footer="References", order=30),
    "barcoding.html":           dict(slug="barcoding", footer="Barcoding", order=50),
    "barcodingmethods.html":    dict(slug="barcoding-methods"),
    "barcodingnucleotides.html": dict(slug="barcoding-nucleotides"),
    "special_features.html":    dict(slug="special-features"),
    "contact.html":             dict(slug="contact", align="center"),
    "copyright.html":           dict(slug="copyright", footer="Copyright", order=60),
}
NAV_IMG = re.compile(r"(images/buttons/|indexmain|groupbrackets|shim)", re.I)
IMG_EXT = (".jpg", ".jpeg", ".png", ".gif", ".bmp")
BR = "\x00BR\x00"


def norm(s):
    s = (s or "").replace("\xa0", " ")
    return re.sub(r"\s+", " ", s).strip()


def clean_remote(rel):
    """Same rule as upload_to_ia.py / inventory_site.py, so video ids match."""
    return re.sub(r"[^A-Za-z0-9._/\-]", "_", rel.replace(" ", "_"))


class Converter:
    def __init__(self, root: Path, page_rel: str, slug: str, content: Path):
        self.root, self.page_rel, self.slug, self.content = root, page_rel, slug, content
        self.lower = {}
        for p in root.rglob("*"):
            if p.is_file() and not p.name.startswith("._"):
                rel = p.relative_to(root).as_posix()
                self.lower[rel.lower()] = rel
        self.notes = []
        self.assets = []

    # ---- paths ----
    def resolve(self, target):
        t = unquote((target or "").strip().replace("\\", "/")).split("#")[0].split("?")[0]
        if not t or re.match(r"^[a-z][a-z0-9+.-]*:", t, re.I) or t.startswith("//"):
            return None
        joined = posixpath.normpath(posixpath.join(posixpath.dirname(self.page_rel), t))
        return self.lower.get(joined.lower()), joined

    def wrapper_to_video_id(self, wrapper_rel):
        f = self.root / wrapper_rel
        if not f.exists():
            return None
        soup = BeautifulSoup(f.read_bytes().decode("cp1252", errors="replace"), "lxml")
        for tag in soup.find_all(["embed", "object", "param", "video", "source", "a"]):
            for attr in ("src", "value", "href", "data"):
                v = tag.get(attr) or ""
                if v.lower().split("?")[0].endswith(".mov"):
                    r = self.resolve_from(wrapper_rel, v)
                    if r:
                        return clean_remote(r[:-4])
        return None

    def resolve_from(self, page_rel, target):
        t = unquote(target.strip().replace("\\", "/")).split("#")[0]
        joined = posixpath.normpath(posixpath.join(posixpath.dirname(page_rel), t))
        return self.lower.get(joined.lower())

    def copy_asset(self, rel):
        name = posixpath.basename(rel)
        dest = self.content / "assets" / self.slug / name
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(self.root / rel, dest)
        self.assets.append(name)
        return f"assets/{self.slug}/{name}"

    # ---- inline ----
    def link_target(self, a):
        oc = re.findall(r"MM_openBrWindow\(\s*'([^']+)'", a.get("onclick") or "")
        cand = oc[0] if oc else (a.get("href") or "")
        if not cand or cand.startswith("#"):
            return None
        if re.match(r"^(https?:|mailto:)", cand, re.I):
            return cand
        r = self.resolve(cand)
        if r is None:
            return cand
        actual, joined = r
        if actual is None:
            self.notes.append(f"link to a file that does not exist: {cand}")
            return f"old:{joined}"
        low = actual.lower()
        if low.endswith(IMG_EXT):
            return self.copy_asset(actual)
        if low.endswith(".mov"):
            return f"video:{clean_remote(actual[:-4])}"
        if low.startswith("media/") and low.endswith((".htm", ".html")):
            vid = self.wrapper_to_video_id(actual)
            if vid:
                return f"video:{vid}"
            self.notes.append(f"video wrapper with no clip found: {actual}")
        return f"old:{actual}"

    def inline(self, el):
        return "".join(self.node(c) for c in el.children)

    def node(self, c):
        """Convert ONE node (text or tag) to Markdown, keeping the tag's own formatting."""
        if isinstance(c, Comment):
            return ""
        if isinstance(c, NavigableString):
            return str(c)
        n = c.name
        if n in ("script", "style"):
            return ""
        if n == "br":
            return BR
        if n == "img":
            return self.image(c)
        if n in ("em", "i"):
            return self.wrap("*", self.inline(c))
        if n in ("strong", "b"):
            return self.wrap("**", self.inline(c))
        if n in ("sup", "sub"):
            inner = self.inline(c).strip()
            return f"<{n}>{inner}</{n}>" if inner else ""
        if n == "a":
            inner = self.inline(c)
            tgt = self.link_target(c)
            return f"[{inner.strip()}]({tgt})" if tgt and inner.strip() else inner
        return self.inline(c)

    @staticmethod
    def wrap(mark, inner):
        core = inner.strip()
        if not core or not re.search(r"\w", core):
            return inner
        return f"{inner[:len(inner) - len(inner.lstrip())]}{mark}{core}{mark}{inner[len(inner.rstrip()):]}"

    def image(self, img):
        src = img.get("src", "")
        if NAV_IMG.search(src):
            return ""
        r = self.resolve(src)
        if r is None or r[0] is None:
            self.notes.append(f"image file not found: {src}")
            return ""
        alt = norm(img.get("alt"))
        if not alt:
            self.notes.append(f"image needs alt text: {src}")
        return f"![{alt}]({self.copy_asset(r[0])})"

    # ---- blocks ----
    def para(self, buf):
        t = norm("".join(buf).replace(BR, " " + BR + " "))
        t = re.sub(r"(\s*" + re.escape(BR) + r"\s*)+", "  \n", t).strip()
        return t

    def lst(self, el, depth=0):
        bullet = "- " if el.name == "ul" else "1. "
        lines = []
        for li in el.find_all("li", recursive=False):
            parts, nested = [], []
            for c in li.children:
                if getattr(c, "name", None) in ("ul", "ol"):
                    nested.append(self.lst(c, depth + 1))
                elif isinstance(c, NavigableString):
                    parts.append(str(c))
                elif c.name == "br":
                    parts.append(" ")
                else:
                    parts.append(self.node(c) + (" " if c.name in ("p", "div") else ""))
            text = norm("".join(parts))
            if text:
                lines.append("    " * depth + bullet + text)
            lines.extend(x for x in nested if x)
        return "\n".join(lines)

    def blocks(self, el):
        out, buf = [], []

        def flush():
            t = self.para(buf)
            if t:
                out.append(t)
            buf.clear()

        for c in el.children:
            if isinstance(c, Comment):
                continue
            if isinstance(c, NavigableString):
                buf.append(str(c))
                continue
            n = c.name
            if n in ("script", "style"):
                continue
            if n in ("p", "div", "blockquote", "center", "form"):
                flush()
                out.extend(self.blocks(c))
            elif n in ("ul", "ol"):
                flush()
                t = self.lst(c)
                if t:
                    out.append(t)
            elif n == "table":
                flush()
                out.extend(self.table(c))
            elif n == "hr":
                flush()
                out.append("---")
            elif n == "br":
                buf.append(BR)
            else:
                buf.append(self.node(c))
        flush()
        return out

    def table(self, tbl):
        out = []
        for tr in tbl.find_all("tr"):
            if tr.find("table"):
                continue
            out.extend(self.row(tr))
        return out

    def row(self, tr):
        tds = tr.find_all("td", recursive=False)
        if not tds:
            return []
        if any(NAV_IMG.search(i.get("src", "")) for i in tr.find_all("img")):
            return []                                    # navigation bar rows
        head = tr.find(class_="textsizeheadline")
        if head is not None:
            txt_h = norm(head.get_text(" ", strip=True))
            if txt_h and txt_h == getattr(self, "title_text", None) and not getattr(self, "title_used", False):
                self.title_used = True
                return []                                # the page title, handled separately
            return [f"## {txt_h}"] if txt_h else []      # a later headline is a section heading
        rowbg = (tr.get("bgcolor") or "").lower()
        out = []
        for td in tds:
            bg = (td.get("bgcolor") or rowbg).lower()
            txt = norm(td.get_text(" ", strip=True))
            cls = td.get("class") or []
            if txt and (bg in ("#000000", "#cccccc") or "headerblack" in cls or "headergray" in cls):
                level = "##" if (bg == "#000000" or "headerblack" in cls) else "###"
                out.append(f"{level} {txt}")
            else:
                out.extend(self.blocks(td))
        return out

    def convert(self, soup):
        title_el = soup.select_one(".textsizeheadline")
        title = norm(title_el.get_text(" ", strip=True)) if title_el else norm(soup.title.get_text() if soup.title else "")
        self.title_text, self.title_used = title, False
        main = soup.select_one("td.mainstyle") or soup.body
        blocks = []
        for tr in main.find_all("tr"):
            if tr.find("table"):
                continue
            blocks.extend(self.row(tr))
        for td in main.find_all("td"):                # cells the old HTML left outside any row
            if td.parent is not None and td.parent.name != "tr" and not td.find("table"):
                blocks.extend(self.blocks(td))
        text = "\n\n".join(b for b in blocks if b.strip())
        text = re.sub(r"\n{3,}", "\n\n", text).strip() + "\n"
        return title, text


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", required=True, help="the cfbkey_local folder")
    ap.add_argument("--content", required=True, help="the content folder (created if needed)")
    ap.add_argument("--pages", nargs="*", help="old page file names (default: the known static pages)")
    ap.add_argument("--force", action="store_true", help="overwrite pages that already exist")
    a = ap.parse_args()
    root = Path(a.root).expanduser().resolve()
    content = Path(a.content).expanduser().resolve()
    content.mkdir(parents=True, exist_ok=True)

    names = a.pages or list(PAGES)
    for name in names:
        cfg = PAGES.get(name, dict(slug=re.sub(r"[^a-z0-9]+", "-", Path(name).stem.lower()).strip("-")))
        rel = f"html/{name}"
        f = root / rel
        if not f.exists():
            print(f"  skip {name}: not found")
            continue
        slug = cfg["slug"]
        dest = content / f"{slug}.md"
        if dest.exists() and not a.force:
            print(f"  skip {name}: {dest.name} already exists (use --force to overwrite)")
            continue
        soup = BeautifulSoup(f.read_bytes().decode("cp1252", errors="replace"), "lxml")
        for x in soup(["script", "style"]):
            x.decompose()
        conv = Converter(root, rel, slug, content)
        title, body = conv.convert(soup)
        title = cfg.get("title") or title
        import re as _re
        if _re.search(r"image-?based key", title, _re.I):   # the generic site title tells us nothing
            m2 = _re.search(r"^##+\s*(.+)$", body, _re.M)
            better = m2.group(1).strip() if m2 else slug.replace("-", " ").title()
            print(f"      ! '{title}' is the generic site title; using '{better}' instead")
            if m2:
                body = body[:m2.start()] + body[m2.end():].lstrip("\n")
            title = better
        fm = ["---", f"title: {title}", f"old: {rel}"]
        if cfg.get("nav"):
            fm += [f"nav: {cfg['nav']}"]
        if cfg.get("footer"):
            fm += ["footer: yes", f"label: {cfg['footer']}"]
        if cfg.get("order"):
            fm += [f"order: {cfg['order']}"]
        if cfg.get("align"):
            fm += [f"align: {cfg['align']}"]
        if cfg.get("indent"):
            fm += [f"indent: {cfg['indent']}"]
        fm += ["---", ""]
        dest.write_text("\n".join(fm) + body, encoding="utf-8")
        words = len(body.split())
        print(f"  wrote {dest.name}: '{title}', {words} words, {len(conv.assets)} image(s)")
        for n in dict.fromkeys(conv.notes):
            print(f"      ! {n}")


if __name__ == "__main__":
    main()
