#!/usr/bin/env python3
"""
build_site.py   (first slice: species pages + species index + home page)

Reads the editable workbook and writes a plain static website: HTML, CSS, images.
Nothing in the output needs a server, a database, or JavaScript, so it can be
opened straight from disk, or hosted on GitHub Pages or anywhere else.

USAGE (inside the conda 'video' environment):

  conda install -c conda-forge jinja2          # once, if needed
  python build_site.py --workbook "$ZK/cfbkey_workbook.xlsx" \\
         --old-site "$ZK/cfbkey_local" --out "$ZK/site" --limit 5
  open "$ZK/site/index.html"

OPTIONS
  --only id1,id2   build just these species_ids
  --limit N        build only the first N species (quick design checks)
  --content DIR    folder of Markdown pages (default: <out>/../content). The first
                   run writes a starter index.md there (the home page). Each page can
                   start with a small header:   ---  title: ...  nav: Label  order: 20
                   footer: yes  label: Text  ---    ('nav' puts it in the top menu,
                   'footer: yes' in the bottom links; use 'label' instead of 'nav' for a
                   footer-only link; put {{aboutlinks}} in a page to show the footer links as
                   buttons there (the footer row is then left off that page); 'indent: no' switches off the first-line indent that text
                   pages get by default; 'width: wide' widens the text column for pages with long
                   single-line entries; 'style: references' indents each reference under its
                   italic sub-heading; 'indent: yes' indents the first line of every
                   paragraph; 'align: center' centres the whole page). Link prefixes
                   in the text:  page:slug,
                   old:html/begin.html (an old address), video:<video_id>. A link
                   shows only once its target exists. Images go in content/assets/.
  --noindex        add a "do not index" tag to every page and a robots.txt, so a test
                   copy does not turn up in search results alongside the live site
  --init           write the page templates and stylesheet to <out>/../site_src
                   so they can be edited; files there override the built-in ones

WHAT IT CHECKS (printed at the end, saved to <out>/build_report.txt)
  * pictures with no new alt text yet (they get a plain fallback so the page works)
  * videos with no description, or not uploaded
  * links in the text it could not resolve
  * every generated page: one <h1>, alt on every image, no duplicate ids,
    every local link and image file actually exists
"""

import argparse
import html
import posixpath
import re
import shutil
import sys
from collections import OrderedDict, defaultdict
from html.parser import HTMLParser
from pathlib import Path, PurePosixPath

try:
    from jinja2 import ChoiceLoader, DictLoader, Environment, FileSystemLoader, select_autoescape
    from markupsafe import Markup
except ImportError:
    sys.exit("Needs jinja2:  conda install -c conda-forge jinja2")
try:
    from openpyxl import load_workbook
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")
try:
    import markdown as md_lib
except ImportError:
    sys.exit("Needs markdown:  conda install -c conda-forge markdown")

# The current version and its release year. Change these two lines when a new version
# is released; they appear in the citation line in the footer of every page.
VERSION = "6.0"
RELEASE_YEAR = "2026"
SITE = {
    "title": "An Image-Based Key to the Zooplankton of North America",
    "version": VERSION,
    "citation": f"To reference the key: Haney, J.F. et al. \"An Image-Based Key to the Zooplankton of "
                f"North America\" version {VERSION} released {RELEASE_YEAR}. University of New Hampshire "
                f"Center for Freshwater Biology <cfb.unh.edu>",
    "copyright": "Copyright \u00a9 2003\u20132013 Center for Freshwater Biology, Department of Biological "
                 "Sciences, University of New Hampshire, Durham, NH 03824 USA",
}
RANKS = ["Phylum", "Subphylum", "Class", "Subclass", "Order", "Suborder", "Family"]
SKIP_SECTIONS = {"quicktime movies", "additional pictures"}

# ----------------------------------------------------------------------------
# Built-in templates and stylesheet (override by putting files in site_src/)
# ----------------------------------------------------------------------------
TEMPLATES = {
"base.html": """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{% if title != site.title %}{{ title }} | {% endif %}{{ site.title }}</title>
<meta name="description" content="{{ description }}">
{% if noindex %}<meta name="robots" content="noindex, nofollow">
{% endif %}
<link rel="stylesheet" href="{{ root }}css/site.css">
</head>
<body>
<a class="skip-link" href="#main">Skip to main content</a>
<div class="frame">
<header class="site-header">
  <div class="header-bar">
    {% if logo %}<img class="logo" src="{{ logo }}" alt="" width="153" height="35">{% endif %}
    <nav aria-label="Main">
      <ul>
        {% for n in nav %}<li><a href="{{ n.href }}"{% if n.current %} aria-current="page"{% endif %}>{{ n.label }}</a></li>
        {% endfor %}
      </ul>
    </nav>
  </div>
</header>
<main id="main" class="wrap">
{% block content %}{% endblock %}
</main>
<footer class="site-footer">
  <div class="wrap">
    {% if footer_nav and not hide_footer_nav %}
    <nav aria-label="More about the key">
      <ul class="footer-nav">
        {% for n in footer_nav %}<li><a href="{{ n.href }}"{% if n.current %} aria-current="page"{% endif %}>{{ n.label }}</a></li>
        {% endfor %}
      </ul>
    </nav>
    {% endif %}
    <p>{{ site.citation }} <span id="accessed"></span></p>
    <p>{{ site.copyright }}</p>
  </div>
</footer>
</div>
<script>
(function () {
  var el = document.getElementById("accessed");
  if (!el) return;
  var m = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"], d = new Date();
  el.textContent = "Accessed " + d.getDate() + " " + m[d.getMonth()] + " " + d.getFullYear() + ".";
})();
</script>
</body>
</html>
""",
"page.html": """{% extends "base.html" %}
{% block content %}
<h1>{{ heading }}</h1>
<div class="prose{% if centered %} centered{% endif %}{% if indent %} indent-paragraphs{% endif %}{% if page_style %} style-{{ page_style }}{% endif %}{% if wide %} prose-wide{% endif %}">
{{ body }}
</div>
{% if clip_groups %}
<div class="prose clips">
<p class="jump"><strong>Jump to:</strong></p>
<nav aria-label="Video groups"><ul>
{% for g in clip_groups %}<li><a href="#grp-{{ loop.index }}">{{ g.title }} ({{ g.clips|length }})</a></li>
{% endfor %}</ul></nav>
{% for g in clip_groups %}
<section aria-labelledby="grp-{{ loop.index }}">
  <h2 id="grp-{{ loop.index }}">{{ g.title }}</h2>
  {% for c in g.clips %}
  <details class="clip" id="{{ c.anchor }}">
    <summary><h3 class="clip-title">{{ c.title }}</h3></summary>
    <div class="clip-body">
      <video controls preload="none" aria-describedby="{{ c.anchor }}-d">
        <source src="{{ c.url }}" type="video/mp4">
        Your browser cannot play this video. <a href="{{ c.url }}">Download it</a>.
      </video>
      {% if c.description %}<p id="{{ c.anchor }}-d">{{ c.description }}</p>{% endif %}
    </div>
  </details>
  {% endfor %}
</section>
{% endfor %}
</div>
{% endif %}
{% endblock %}
""",
"figure_dialog.html": """<dialog id="figure-dialog" aria-labelledby="fd-caption">
  <form method="dialog"><button class="close" aria-label="Close picture">Close</button></form>
  <img id="fd-img" src="" alt="">
  <p id="fd-caption"></p>
</dialog>
<script>
(function () {
  var dlg = document.getElementById("figure-dialog");
  if (!dlg || !dlg.showModal) return;          // older browsers: the link just opens the picture
  var img = document.getElementById("fd-img"), cap = document.getElementById("fd-caption");
  document.querySelectorAll("a[data-figure-alt]").forEach(function (a) {
    a.addEventListener("click", function (e) {
      e.preventDefault();
      img.src = a.href; img.alt = a.dataset.figureAlt; cap.textContent = a.dataset.figureAlt;
      dlg.showModal();
    });
  });
  dlg.addEventListener("click", function (e) { if (e.target === dlg) dlg.close(); });
})();
</script>
""",
"key_step.html": """{% extends "base.html" %}
{% block content %}
<article class="key-step">
  {% if heading_hidden %}<h1 class="hidden-heading">{{ heading }}</h1>
    {% if not preview_mode %}<p class="key-prompt">{{ prompt }}</p>{% endif %}
  {% else %}<h1 class="key-prompt">{{ prompt }}</h1>{% endif %}
  {% for g in guides %}
  <figure class="key-guide">
    <a href="{{ g.src }}" data-figure-alt="{{ g.alt }}"><img src="{{ g.src }}" alt="{{ g.alt }}"></a>
    {% if g.caption %}<figcaption>{{ g.caption }}</figcaption>{% endif %}
  </figure>
  {% endfor %}
  {% if group_grid %}
  <ul class="group-grid">
    {% for c in choices %}
    <li>
      {% if c.url %}<a href="{{ c.url }}">{% else %}<span>{% endif %}
        <img src="{{ c.images[0].src }}" alt="" data-decorative>
        <span class="group-grid-label">{{ c.name if c.name else "Group " + loop.index|string }}</span>
      {% if c.url %}</a>{% else %}</span>{% endif %}
    </li>
    {% endfor %}
  </ul>
  {% elif preview_mode %}
  <div class="entry-layout">
    <ul class="entry-groups">
      {% for c in choices %}
      <li>
        {% if c.url %}<a class="entry-link" href="{{ c.url }}" data-preview="{{ c.preview }}" data-preview-alt="{{ c.preview_alt }}">{% else %}<span class="entry-link">{% endif %}
          <img src="{{ c.images[0].src }}" alt="" data-decorative width="150" height="150">
          <span class="entry-label">{{ c.name }}</span>
        {% if c.url %}</a>{% else %}</span>{% endif %}
        {% if c.preview %}<button type="button" class="preview-btn" aria-expanded="false">Show examples<span class="sr-only"> of {{ c.name }}</span></button>{% endif %}
      </li>
      {% endfor %}
    </ul>
    <div class="entry-preview" id="entry-preview" aria-live="polite">
      <p class="entry-default" id="entry-default">{{ prompt }}</p>
      <img id="entry-img" src="" alt="" hidden>
    </div>
  </div>
  {% else %}
  <ol class="choices{% if choices|length <= 4 %} cols-{{ choices|length }}{% else %} compact{% endif %}">
    {% for c in choices %}
    <li class="choice">
      <p class="choice-num">Choice {{ loop.index }} of {{ choices|length }}</p>
      {% if c.url %}<a class="choice-link" href="{{ c.url }}">{% else %}<span class="choice-link no-target">{% endif %}
        {% for im in c.images %}<img src="{{ im.src }}" alt="{{ im.alt }}"{% if not im.alt %} data-decorative{% endif %}{% if im.plain %} data-labeled="{{ im.src }}" data-plain="{{ im.plain }}"{% endif %}>{% endfor %}
        <span class="choice-text">{% for line in c.lines %}<span class="choice-line">{{ line }}</span>{% endfor %}</span>
        {% if c.go %}<span class="choice-go">{{ c.go }}</span>{% endif %}
      {% if c.url %}</a>{% else %}</span>{% endif %}
    </li>
    {% endfor %}
  </ol>
  {% endif %}
  {% if has_plain %}
  <p class="key-tools"><button type="button" id="toggle-arrows" aria-pressed="false">Hide arrows and labels on the pictures</button></p>
  {% endif %}
  {% if videos %}
  <section class="key-extras" aria-labelledby="kv-h">
    <h2 id="kv-h">Videos for this step</h2>
    {% for v in videos %}
    <details class="key-video">
      <summary>{{ v.label }}</summary>
      <video controls preload="none"><source src="{{ v.url }}" type="video/mp4">Your browser cannot play this video. <a href="{{ v.url }}">Download it</a>.</video>
      {% if v.description %}<p>{{ v.description }}</p>{% endif %}
    </details>
    {% endfor %}
  </section>
  {% endif %}
  {% if helps %}
  <ul class="key-help">{% for h in helps %}<li>{{ h }}</li>{% endfor %}</ul>
  {% endif %}
  <nav class="key-nav" aria-label="Key navigation">
    {% if back %}<a href="{{ back }}">&larr; Previous step</a>{% endif %}
    {% if show_start %}<a href="{{ start }}">Start over</a>{% endif %}
  </nav>
  {% if taxonomy %}
  <section class="key-extras" aria-labelledby="kt-h">
    <h2 id="kt-h">Taxonomy</h2>
    <dl class="taxonomy">{% for label, value in taxonomy %}<div><dt>{{ label }}</dt><dd>{{ value }}</dd></div>{% endfor %}</dl>
  </section>
  {% endif %}
</article>
{% include "figure_dialog.html" %}
{% if has_plain %}
<script>
(function () {
  var btn = document.getElementById("toggle-arrows");
  if (!btn) return;
  var imgs = document.querySelectorAll("img[data-plain]"), on = false;
  function set(im, plain) { im.src = plain ? im.dataset.plain : im.dataset.labeled; }
  btn.addEventListener("click", function () {
    on = !on;
    btn.setAttribute("aria-pressed", on ? "true" : "false");
    btn.textContent = on ? "Show arrows and labels on the pictures" : "Hide arrows and labels on the pictures";
    imgs.forEach(function (im) { set(im, on); });
  });
  imgs.forEach(function (im) {                     /* pointing at a picture also shows it without arrows, as before */
    var box = im.closest(".choice-link") || im;
    box.addEventListener("mouseenter", function () { if (!on) set(im, true); });
    box.addEventListener("mouseleave", function () { if (!on) set(im, false); });
  });
})();
</script>
{% endif %}
{% if preview_mode %}
<script>
(function () {
  var panel = document.getElementById("entry-preview");
  if (!panel) return;
  var img = document.getElementById("entry-img"), def = document.getElementById("entry-default");
  var links = [].slice.call(document.querySelectorAll(".entry-link[data-preview]"));
  links.forEach(function (a) { var p = new Image(); p.src = a.dataset.preview; });    /* preload */
  function show(a) { img.src = a.dataset.preview; img.alt = a.dataset.previewAlt; img.hidden = false; def.hidden = true; }
  function clear() { img.hidden = true; img.removeAttribute("src"); img.alt = ""; def.hidden = false; }
  links.forEach(function (a) {
    a.addEventListener("mouseenter", function () { show(a); });
    a.addEventListener("focus", function () { show(a); });
    a.addEventListener("mouseleave", clear);
    a.addEventListener("blur", clear);
  });
  var btns = [].slice.call(document.querySelectorAll(".preview-btn"));
  btns.forEach(function (b) {
    b.addEventListener("click", function () {
      var open = b.getAttribute("aria-expanded") === "true";
      btns.forEach(function (o) { o.setAttribute("aria-expanded", "false"); });
      if (open) { clear(); return; }
      b.setAttribute("aria-expanded", "true");
      show(b.parentNode.querySelector(".entry-link"));
      panel.scrollIntoView({ block: "nearest" });
    });
  });
})();
</script>
{% endif %}
{% endblock %}
""",
"extra_page.html": """{% extends "base.html" %}
{% block content %}
<article class="extra-page">
  <h1>{{ heading }}</h1>
  {% if parent %}<p class="extra-parent">Part of <a href="{{ parent.url }}"><em>{{ parent.name }}</em></a></p>{% endif %}
  {% if has_plain %}
  <p class="key-tools"><button type="button" id="toggle-arrows" aria-pressed="false">Hide arrows and labels on the pictures</button></p>
  {% endif %}
  {% for sec in sections %}
  <section class="text-section">
    {% if sec.title %}<h2>{{ sec.title }}</h2>{% endif %}
    {% for blk in sec.blocks %}
      {% if blk.title %}<h3>{{ blk.title }}</h3>{% endif %}
      <ul>{% for it in blk.entries %}<li>{{ it }}</li>{% endfor %}</ul>
    {% endfor %}
  </section>
  {% endfor %}
  {% if images %}
  <div class="extra-gallery">
    {% for im in images %}
    <figure>
      {% if im.large %}<a href="{{ im.large }}" data-figure-alt="{{ im.alt }}">{% endif %}
      <img src="{{ im.src }}" alt="{{ im.alt }}"{% if im.plain %} data-labeled="{{ im.src }}" data-plain="{{ im.plain }}"{% endif %}>
      {% if im.large %}</a>{% endif %}
      {% if im.regions %}
      <figcaption>Select a part: {% for r in im.regions %}{% if r.url %}<a href="{{ r.url }}">{{ r.label }}</a>{% else %}{{ r.label }}{% endif %}{% if not loop.last %} &middot; {% endif %}{% endfor %}</figcaption>
      {% endif %}
    </figure>
    {% endfor %}
  </div>
  {% endif %}
  {% if videos %}
  <section class="text-section" aria-labelledby="xv-h">
    <h2 id="xv-h">Videos</h2>
    {% for v in videos %}
    <figure class="video">
      <video controls preload="none" aria-describedby="xvd-{{ loop.index }}"><source src="{{ v.url }}" type="video/mp4">
        Your browser cannot play this video. <a href="{{ v.url }}">Download it</a>.</video>
      <figcaption id="xvd-{{ loop.index }}">{{ v.description }}</figcaption>
    </figure>
    {% endfor %}
  </section>
  {% endif %}
</article>
{% include "figure_dialog.html" %}
{% if has_plain %}
<script>
(function () {
  var btn = document.getElementById("toggle-arrows");
  var imgs = document.querySelectorAll("img[data-plain]"), on = false;
  function set(im, plain) { im.src = plain ? im.dataset.plain : im.dataset.labeled; }
  if (btn) btn.addEventListener("click", function () {
    on = !on;
    btn.setAttribute("aria-pressed", on ? "true" : "false");
    btn.textContent = on ? "Show arrows and labels on the pictures" : "Hide arrows and labels on the pictures";
    imgs.forEach(function (im) { set(im, on); });
  });
  imgs.forEach(function (im) {
    im.addEventListener("mouseenter", function () { if (!on) set(im, true); });
    im.addEventListener("mouseleave", function () { if (!on) set(im, false); });
  });
})();
</script>
{% endif %}
{% endblock %}
""",
"anatomy_index.html": """{% extends "base.html" %}
{% block content %}
<h1 class="hidden-heading">Anatomy</h1>
<p class="anatomy-intro">{{ intro }}</p>
<div class="ana-grid">
{% for sec in sections %}
  <section aria-labelledby="ana-{{ loop.index }}">
    <h2 id="ana-{{ loop.index }}">{{ sec.title }}</h2>
    <ul class="ana-cards">
      {% for c in sec.cards %}
      <li><a href="{{ c.url }}"><img src="{{ c.src }}" alt="" data-decorative><span>{{ c.label }}</span></a></li>
      {% endfor %}
    </ul>
  </section>
{% endfor %}
</div>
{% endblock %}
""",
"anatomy.html": """{% extends "base.html" %}
{% block content %}
<h1>{{ heading }}</h1>
<p class="anatomy-intro">{{ intro }}</p>
{% if tabs|length > 1 %}
<nav class="group-tabs" aria-label="Anatomy diagrams"><ul>
  {% for t in tabs %}<li><a href="{{ t.href }}"{% if t.current %} aria-current="page"{% endif %}>{{ t.label }}</a></li>
  {% endfor %}
</ul></nav>
{% endif %}
<div class="anatomy">
  <div class="anatomy-controls">
    <h2 id="parts-h" class="sr-only">Parts</h2>
    <ul class="part-list" aria-labelledby="parts-h">
      {% for v in views %}
      <li><button type="button" class="part-btn{% if v.is_base %} is-current{% endif %}" data-img="{{ v.src }}"
          data-alt="{{ v.alt }}" aria-pressed="{{ 'true' if v.is_base else 'false' }}">{{ v.label }}</button></li>
      {% endfor %}
    </ul>
  </div>
  <figure class="anatomy-figure">
    <img id="anatomy-img" src="{{ base.src }}" alt="{{ base.alt }}">
    <figcaption id="anatomy-caption" aria-live="polite">{{ base.label }}</figcaption>
  </figure>
</div>
<noscript><p>Each view is shown separately below.</p>
  {% for v in views %}<figure class="anatomy-figure"><img src="{{ v.src }}" alt="{{ v.alt }}"><figcaption>{{ v.label }}</figcaption></figure>{% endfor %}
</noscript>
<script>
(function () {
  var img = document.getElementById("anatomy-img"), cap = document.getElementById("anatomy-caption");
  var btns = [].slice.call(document.querySelectorAll(".part-btn"));
  btns.forEach(function (b) { var p = new Image(); p.src = b.dataset.img; });
  btns.forEach(function (b) {
    b.addEventListener("click", function () {
      btns.forEach(function (o) { o.setAttribute("aria-pressed", "false"); o.classList.remove("is-current"); });
      b.setAttribute("aria-pressed", "true"); b.classList.add("is-current");
      img.src = b.dataset.img; img.alt = b.dataset.alt; cap.textContent = b.textContent;
    });
  });
})();
</script>
{% endblock %}
""",
"groups.html": """{% extends "base.html" %}
{% block content %}
<h1 class="hidden-heading">Groups</h1>
<p class="groups-intro">Almost every group name below is linked to the point in the key at which an organism has been identified as a member of that group. Select a group name to begin the key from that point; select a species to see its page.</p>
<nav class="group-tabs" aria-label="Groups"><ul>
  {% for t in tabs %}<li><a href="{{ t.href }}"{% if t.current %} aria-current="page"{% endif %}>{{ t.label }}</a></li>
  {% endfor %}
</ul></nav>
{% if note %}<p class="groups-note"><strong>Note:</strong> {{ note }}</p>{% endif %}
<div class="outline">{{ outline }}</div>
{% endblock %}
""",
"species_index.html": """{% extends "base.html" %}
{% block content %}
<h1 class="hidden-heading">Species</h1>
<p>{{ n_species }} species, grouped by family.</p>
{% for fam, items in groups %}
<section aria-labelledby="fam-{{ loop.index }}">
  <h2 id="fam-{{ loop.index }}">{{ fam }}</h2>
  <ul class="species-list">
  {% for s in items %}
    <li>{% if s.built %}<a href="{{ s.url }}"><em>{{ s.display_name }}</em></a>{% else %}<em>{{ s.display_name }}</em>{% endif %}</li>
  {% endfor %}
  </ul>
</section>
{% endfor %}
{% endblock %}
""",
"species.html": """{% extends "base.html" %}
{% block content %}
<article class="species">
  <h1 class="species-name"><em>{{ sp.display_name }}</em></h1>

  <div class="hero">
    <div class="hero-images">
      {% if main %}
      <a class="main-img" href="{{ main.large }}" data-figure-alt="{{ main.alt }}"><img src="{{ main.src }}" alt="{{ main.alt }}"{% if main.w %} width="{{ main.w }}" height="{{ main.h }}"{% endif %}></a>
      {% endif %}
    </div>
    <div class="hero-side">
      {% if thumbs %}
      <div class="thumbs">
        {% for t in thumbs %}
        <a href="{{ t.large }}" data-figure-alt="{{ t.alt }}"><img src="{{ t.src }}" alt="{{ t.alt }}"{% if t.w %} width="{{ t.w }}" height="{{ t.h }}"{% endif %}></a>
        {% endfor %}
      </div>
      {% endif %}
      <dl class="taxonomy">
        {% for label, value in taxonomy %}
        <div><dt>{{ label }}</dt><dd>{{ value }}</dd></div>
        {% endfor %}
      </dl>
    </div>
  </div>

  {% for sec in sections %}
  <section class="text-section">
    {% if sec.title %}<h2>{{ sec.title }}</h2>{% endif %}
    {% for blk in sec.blocks %}
      {% if blk.title %}<h3>{{ blk.title }}</h3>{% endif %}
      {% if sec.is_refs %}
        {% for it in blk.entries %}<p class="ref">{{ it }}</p>{% endfor %}
      {% else %}
        <ul>{% for it in blk.entries %}<li>{{ it }}</li>{% endfor %}</ul>
      {% endif %}
    {% endfor %}
  </section>
  {% endfor %}

  {% if videos %}
  <section class="text-section" aria-labelledby="videos-h">
    <h2 id="videos-h">Videos</h2>
    {% for v in videos %}
    <figure class="video" id="v-{{ loop.index }}">
      <video controls preload="none" aria-describedby="vd-{{ loop.index }}">
        <source src="{{ v.url }}" type="video/mp4">
        Your browser cannot play this video. <a href="{{ v.url }}">Download it</a>.
      </video>
      <figcaption id="vd-{{ loop.index }}">{% if v.title %}<strong>{{ v.title }}</strong> {% endif %}{{ v.description }}</figcaption>
    </figure>
    {% endfor %}
  </section>
  {% endif %}

  {% if more %}
  <section class="text-section" aria-labelledby="more-h">
    <h2 id="more-h">Additional pictures</h2>
    <div class="gallery">
      {% for t in more %}
      <a href="{{ t.large }}" data-figure-alt="{{ t.alt }}"><img src="{{ t.src }}" alt="{{ t.alt }}"{% if t.w %} width="{{ t.w }}" height="{{ t.h }}"{% endif %}></a>
      {% endfor %}
    </div>
  </section>
  {% endif %}
</article>
{% include "figure_dialog.html" %}
{% endblock %}
""",
}

CSS = """\
:root { --ink:#15181c; --bg:#fff; --bar:#15181c; --bar-text:#fff; --sub:#e8e8e6; --label:#9b0000;
  --link:#0b4f9c; --line:#b9b9b4;
  /* ONE web-safe family for the whole site. Arial first (as on the old site), then
     Helvetica for Macs without it and Liberation Sans for Linux: all three share the
     same metrics, so the layout is identical on every machine. To change the look of
     the site, change these two lines only. */
  --display:Arial, Helvetica, "Liberation Sans", sans-serif;
  --body:Arial, Helvetica, "Liberation Sans", sans-serif; }
* { box-sizing: border-box; }
html { font-size: 100%; }
body { margin:0; background:#e4e4e4; color:var(--ink); font:1rem/1.6 var(--body); }
button, input, select, textarea { font-family:inherit; }
.frame { max-width:64rem; margin:1rem auto; background:#fff; border:3px solid #666; padding:.6rem; }
.wrap { padding: 0 1rem; }
main.wrap { background: var(--bg); padding-top: 1rem; padding-bottom: 2rem; }
.skip-link { position:absolute; left:-999px; top:0; background:#fff; color:#000; padding:.5rem 1rem; z-index:10; }
.skip-link:focus { left:0; }
a { color: var(--link); text-decoration: underline; }
a:visited { color:#5b2a86; }
a:hover { color:#000; }
[hidden] { display:none !important; }
.hidden-heading { position:absolute; width:1px; height:1px; margin:-1px; padding:0; overflow:hidden; clip:rect(0 0 0 0); white-space:nowrap; border:0; }
.sr-only, .choice-num, .choice-go { position:absolute; width:1px; height:1px; margin:-1px; padding:0; overflow:hidden; clip:rect(0 0 0 0); white-space:nowrap; border:0; }
a:focus-visible, video:focus-visible, button:focus-visible { outline: 3px solid #ffbf47; outline-offset: 2px; }
.site-header { background:#000; color:#fff; }
.header-bar { display:flex; flex-wrap:wrap; align-items:center; gap:.2rem 1.75rem; padding:.25rem .6rem; }
.site-header .logo { flex:none; margin:0; display:block; }
.site-header nav ul { list-style:none; display:flex; flex-wrap:wrap; gap:.1rem 1.7rem; margin:0; padding:0; }
.site-header nav a, .site-header nav a:visited { display:inline-block; padding:.55rem 0; color:#fff; text-decoration:none;
  font:700 1.25rem/1 var(--display); letter-spacing:.005em; }
.site-header nav a:hover { text-decoration:underline; }
.site-header nav a[aria-current="page"] { text-decoration:underline; text-decoration-thickness:3px; text-underline-offset:.4rem; }
/* Type scale: 2rem page title, 1.85rem species name, 1.25rem menu, 1.12rem prompts,
   1.02rem section bars (uppercase), 0.95rem sub-bars, 1rem body. */
h1 { font:700 2rem/1.2 var(--display); letter-spacing:-.01em; margin:.6rem 0 1.1rem; text-align:center; }
.species-name { background:var(--bar); color:var(--bar-text); padding:.6rem .75rem; margin:.25rem 0 1.4rem;
  font:700 1.85rem/1.2 var(--display); letter-spacing:-.01em; }
h2 { font:700 1.02rem/1.3 var(--display); letter-spacing:.06em; text-transform:uppercase;
  background:var(--bar); color:var(--bar-text); padding:.42rem .7rem; margin:1.9rem 0 .7rem; }
h3 { font:700 .95rem/1.3 var(--display); letter-spacing:.03em; background:var(--sub);
  padding:.3rem .7rem; margin:1.15rem 0 .45rem; }
.lede { font-size:1.15rem; }
.prose p { max-width:42rem; margin:0 0 .85rem; }
.indent-paragraphs p { text-indent:2.5em; }
.prose-wide { max-width:58rem; }
.release-line { text-indent:0 !important; margin:0 0 .55rem !important; font-size:.9rem; }
.prose-wide p, .prose-wide li { max-width:none; }
.style-references { max-width:46rem; }
.style-references p { margin:0 0 .8rem 2.6rem; max-width:none; }
.style-references p.ref-head { margin:1.4rem 0 .6rem; font-style:italic; }
.style-references h2 + p.ref-head { margin-top:.6rem; }
.indent-paragraphs li p, .indent-paragraphs blockquote p { text-indent:0; }
.indent-paragraphs .ref-head, .style-references p { text-indent:0; }
.centered .prose p, .centered p { margin-left:auto; margin-right:auto; }
.hero { display:grid; grid-template-columns: minmax(0, 360px) 1fr; gap:1rem 1.5rem; align-items:start; }
.hero-images { display:grid; gap:.5rem; }
.hero img, .gallery img { max-width:100%; height:auto; border:1px solid var(--line); background:#fff; display:block; }
.hero-side { display:grid; gap:1rem; align-content:start; }
.thumbs { display:grid; grid-template-columns: repeat(2, minmax(0, 175px)); gap:.75rem; justify-content:center; }
.taxonomy { margin:0; }
.taxonomy div { display:flex; gap:1rem; padding:.2rem 0; border-bottom:1px solid #ddd; }
.taxonomy dt { font-weight:bold; color:var(--label); min-width:7rem; }
.taxonomy dd { margin:0; font-weight:bold; }
.text-section ul { margin:.3rem 0; padding-left:1.4rem; }
.text-section li { margin:.2rem 0; }
.ref { margin:.3rem 0 .3rem 2rem; text-indent:-2rem; }
.gallery { display:flex; flex-wrap:wrap; gap:.75rem; }
figure.video { margin:1rem 0; }
figure.video video { width:100%; max-width:640px; height:auto; background:#000; display:block; }
figcaption { max-width:640px; margin-top:.35rem; }
.species-list { columns: 2 18rem; column-gap:2rem; padding-left:1.2rem; }
.logo { vertical-align:middle; margin-right:.75rem; max-width:100%; height:auto; }
.prose { max-width: 48rem; margin-left:auto; margin-right:auto; }
.centered { text-align:center; }
.centered h2, .centered h3 { text-align:center; }
.centered img { display:block; margin-left:auto; margin-right:auto; }
.centered ul:not(.group-strip):not(.link-buttons), .centered ol { display:inline-block; text-align:left; }
.banner-link { display:block; max-width:54rem; margin:1.5rem auto; text-decoration:none; }
.banner-strip { position:relative; display:grid; grid-template-columns:repeat(4, 1fr); gap:.35rem;
  border:3px solid var(--ink); background:#fff; padding:.35rem; }
.banner-strip img { display:block; width:100%; height:auto; }
.banner-overlay { position:absolute; left:12.5%; right:12.5%; top:50%; transform:translateY(-50%);
  background:rgba(0,0,0,.52); color:#fff; text-align:center; padding:.6rem .5rem;
  font:700 1.9rem/1.1 var(--display); letter-spacing:.02em; text-decoration:none; }
.banner-link:hover .banner-overlay, .banner-link:focus-visible .banner-overlay { background:rgba(0,0,0,.68); }
.banner-link:focus-visible { outline:3px solid #ffbf47; outline-offset:3px; }
@media (max-width:34rem) { .banner-strip { grid-template-columns:repeat(2, 1fr); } .banner-overlay { font-size:1.3rem; } }
.prose h2 { margin-top:1.8rem; }
.prose img { max-width:100%; height:auto; }
.prose table { border-collapse:collapse; margin:1rem 0; }
.prose th, .prose td { border:1px solid var(--line); padding:.35rem .6rem; text-align:left; vertical-align:top; }
.link-buttons { list-style:none; display:flex; flex-wrap:wrap; justify-content:center;
  gap:.6rem .7rem; margin:1.1rem 0 .4rem; padding:0; }
.link-buttons a.link-button, .link-buttons a.link-button:visited { display:block; padding:.6rem 1.15rem;
  background:var(--bar); color:#fff; text-decoration:none; border:2px solid var(--bar); border-radius:4px;
  font:700 .95rem/1.2 var(--display); letter-spacing:.02em; }
.link-buttons a.link-button:hover { background:#fff; color:var(--bar); }
.link-buttons a.link-button:focus-visible { outline:3px solid #ffbf47; outline-offset:2px; }
@media (max-width:44rem) { .link-buttons a.link-button { padding:.55rem .8rem; font-size:.9rem; } }
.footer-nav { list-style:none; display:flex; flex-wrap:wrap; justify-content:center; gap:.5rem 1.5rem; margin:0 0 .75rem; padding:0; }
.site-footer { text-align:center; background:#ddd; font-size:.9rem; padding:1rem 0; margin-top:0; }
.site-footer p { margin:.3rem 0; }
.prose, .text-section, .ref { overflow-wrap:break-word; }
img, video { max-width:100%; }
.prose table { display:block; overflow-x:auto; }
.mk-red { color:#b00000; font-weight:bold; }
.mk-blue { color:#0033cc; font-weight:bold; }
.mk-green { color:#066a00; font-weight:bold; }
.mk-orange { color:#a85400; font-weight:bold; }
.key-prompt { font:700 1.12rem/1.45 var(--display); letter-spacing:0; text-align:center;
  margin:.3rem auto .9rem; max-width:46rem; }
.key-guide { margin:0 auto 1rem; text-align:center; }
.key-guide img { display:inline-block; max-width:min(100%, 22rem); height:auto; border:1px solid var(--line); }
.key-tools { text-align:center; }
.key-tools button, .key-nav a { font:inherit; }
.key-tools button { padding:.5rem 1rem; cursor:pointer; border:2px solid var(--ink); background:#fff; border-radius:4px; }
.key-tools button[aria-pressed="true"] { background:#111; color:#fff; }
.choices { list-style:none; margin:.75rem 0; padding:0; display:grid; gap:.6rem .7rem;
  grid-template-columns:repeat(auto-fit, minmax(min(100%, 24rem), 1fr)); }
.choices.cols-1 { grid-template-columns:minmax(0, 46rem); justify-content:center; }
.choices.cols-2 { grid-template-columns:repeat(2, minmax(0, 1fr)); }
.choices.cols-3 { grid-template-columns:repeat(3, minmax(0, 1fr)); }
.choices.cols-4 { grid-template-columns:repeat(4, minmax(0, 1fr)); gap:.6rem .5rem; }
.choices.cols-3 .choice-text, .choices.cols-4 .choice-text { font-size:.92rem; line-height:1.45; }
.choices.compact { grid-template-columns:repeat(auto-fit, minmax(min(100%, 15rem), 1fr)); gap:.6rem .8rem; }
.choices.compact .choice-text { font-size:.95rem; }
.choice { margin:0; display:flex; flex-direction:column; }
.choice-link, .choice-link:visited { display:block; flex:1; padding:0 .15rem; text-align:center; color:var(--ink); text-decoration:none; background:none; }
a.choice-link:hover { background:none; }
a.choice-link:hover img { border-color:var(--link); }
a.choice-link:focus-visible { outline:3px solid #ffbf47; outline-offset:4px; }
.choice-link img { display:block; max-width:100%; height:auto; margin:0 auto; border:2px solid #00008b; border-bottom-width:1px; }
.choice-link img:last-of-type { border-bottom-width:2px; margin-bottom:.5rem; }
.choice-link img + img { border-top-width:1px; }
a.choice-link:visited img { border-color:#5b2a86; }
a.choice-link:hover img { border-color:var(--link); }
.choice-text { font-weight:bold; }
.choice-line { display:block; margin:.25rem 0; }
.entry-layout { display:grid; grid-template-columns:11rem minmax(0, 1fr); gap:1rem 1.75rem;
  align-items:start; margin:1rem auto 1.5rem; max-width:48rem; }
.entry-groups { list-style:none; margin:0; padding:0; display:grid; gap:.55rem; }
.entry-groups li { text-align:center; }
.entry-link, .entry-link:visited { display:block; color:var(--link); text-decoration:none; }
.entry-link:visited { color:#5b2a86; }
.entry-link img { display:block; width:160px; height:160px; max-width:100%; margin:0 auto .15rem; border:2px solid currentColor; }
.entry-label { display:block; color:var(--ink); font:700 .95rem/1.3 var(--display); }
a.entry-link:hover .entry-label { text-decoration:underline; }
a.entry-link:focus-visible { outline:3px solid #ffbf47; outline-offset:4px; }
.group-grid { list-style:none; margin:1rem 0; padding:0; display:grid;
  grid-template-columns:repeat(auto-fill, minmax(10.5rem, 1fr)); gap:.9rem 1rem; }
.group-grid li { text-align:center; }
.group-grid a, .group-grid span, .group-grid a:visited { display:block; text-decoration:none; color:var(--ink); }
.group-grid img { display:block; width:100%; height:auto; border:2px solid #00008b; margin:0 auto .3rem; }
.group-grid a:visited img { border-color:#5b2a86; }
.group-grid a:hover img { border-color:var(--link); }
.group-grid a:focus-visible { outline:3px solid #ffbf47; outline-offset:3px; }
.group-grid-label { font:700 .95rem/1.3 var(--display); }
.group-grid a:hover .group-grid-label { text-decoration:underline; }
.entry-preview { min-height:38rem; display:flex; align-items:center; justify-content:center; text-align:center; }
.entry-default { max-width:22rem; margin:0 auto; font:700 1.05rem/1.5 var(--display); }
.entry-preview img { display:block; max-width:100%; height:auto; border:4px solid #000; }
.preview-btn { display:none; font:inherit; margin:.3rem auto 0; padding:.35rem .75rem; cursor:pointer; border:2px solid var(--ink); background:#fff; border-radius:4px; }
.preview-btn[aria-expanded="true"] { background:#111; color:#fff; }
@media (hover:none), (max-width:44rem) { .preview-btn { display:inline-block; } }
@media (max-width:44rem) { .entry-layout { grid-template-columns:1fr; } .entry-groups { grid-template-columns:repeat(2, 1fr); }
  .entry-preview { min-height:0; } }
.extra-parent { text-align:center; margin:-.5rem 0 1rem; }
.extra-gallery { display:flex; flex-wrap:wrap; justify-content:center; gap:1.25rem; margin:1rem 0; }
.extra-gallery figure { margin:0; max-width:100%; text-align:center; }
.extra-gallery img { display:block; max-width:100%; height:auto; border:1px solid var(--line); margin:0 auto; }
.extra-gallery figcaption { margin-top:.4rem; max-width:34rem; }
.anatomy-intro { max-width:46rem; margin:0 auto 1rem; text-align:center; font-weight:700; }
.ana-grid { display:flex; flex-wrap:wrap; justify-content:center; align-items:flex-start;
  gap:1rem 1.75rem; margin:1rem 0 1.5rem; }
.ana-grid section { flex:0 1 auto; max-width:100%; }
.ana-grid h2 { font:700 1.2rem/1.3 var(--display); letter-spacing:.03em; margin:0 0 .6rem; }
.ana-cards { list-style:none; display:flex; flex-wrap:wrap; justify-content:center; gap:.5rem; margin:0; padding:0; }
.ana-cards li { text-align:center; flex:0 0 auto; max-width:100%; }
.ana-cards a, .ana-cards a:visited { display:block; text-decoration:none; color:var(--ink); }
.ana-cards img { display:block; height:15rem; width:auto; max-width:100%; background:#fff;
  margin:0 auto .35rem; border:2px solid #00008b; }
.ana-cards a:visited img { border-color:#5b2a86; }
.ana-cards a:hover img { border-color:var(--link); }
.ana-cards a:focus-visible { outline:3px solid #ffbf47; outline-offset:3px; }
.ana-cards span { font:700 .85rem/1.25 var(--display); display:block; max-width:11rem; margin:0 auto; }
.ana-cards a:hover span { text-decoration:underline; }
.anatomy { display:grid; grid-template-columns:13rem minmax(0, 1fr); gap:.5rem 1.5rem; align-items:start; margin:.5rem 0 1.5rem; }
.part-list { list-style:none; margin:0; padding:0; display:grid; gap:.1rem; }
.part-btn { font:700 .95rem/1.35 var(--body); width:100%; padding:.45rem .3rem; cursor:pointer;
  background:none; color:var(--link); border:0; border-bottom:1px solid #e3e3e0; text-align:center; }
.part-btn:hover { text-decoration:underline; color:#000; }
.part-btn.is-current, .part-btn[aria-pressed="true"] { color:#000; font-weight:700; background:var(--sub); }
.part-btn:focus-visible { outline:3px solid #ffbf47; outline-offset:-2px; }
.anatomy-figure { margin:0; text-align:center; }
.anatomy-figure img { display:block; max-width:100%; height:auto; margin:0 auto; }
.anatomy-figure figcaption { font:700 1rem/1.3 var(--display); margin-top:.6rem; }
@media (max-width:44rem) { .ana-cards img { height:11rem; }
  .anatomy { grid-template-columns:1fr; }
  .part-list { grid-template-columns:repeat(2, 1fr); gap:.2rem .5rem; }
  .part-btn { border:1px solid #ddd; border-radius:4px; } }
.groups-intro { max-width:46rem; margin:0 auto 1rem; text-align:center; }
.groups-note { max-width:46rem; margin:1rem auto; }
.group-tabs { background:#333; margin:1rem 0; }
.group-tabs ul { display:flex; flex-wrap:wrap; justify-content:center; list-style:none; margin:0; padding:0; }
.group-tabs a, .group-tabs a:visited { display:block; padding:.6rem 1.1rem; color:#fff; font-weight:bold; text-decoration:none; }
.group-tabs a:hover { color:#ffd54a; text-decoration:underline; }
.group-tabs a[aria-current="page"] { color:#ffd54a; text-decoration:underline; text-decoration-thickness:3px; text-underline-offset:.35rem; }
.outline { max-width:46rem; margin:0 auto; }
.outline ul { list-style:none; margin:0; padding-left:1.6rem; }
.outline > ul { padding-left:0; }
.outline li { margin:.2rem 0; }
.outline li.rank-order > a, .outline li.rank-family > a, .outline li.rank-suborder > a, .outline li.rank-class > a, .outline li.rank-phylum > a { font-weight:bold; }
.outline li.rank-order, .outline li.rank-suborder, .outline li.rank-class, .outline li.rank-phylum { margin-top:.8rem; }
.key-extras h2 { margin-top:1.75rem; }
.key-step .taxonomy div { border-bottom:0; padding:.1rem 0; }
.clip { margin:.3rem 0; }
.clip > summary { list-style:none; cursor:pointer; }
.clip > summary::-webkit-details-marker { display:none; }
.clip > summary .clip-title { background:var(--sub); margin:0; padding:.45rem .7rem .45rem 2rem;
  position:relative; font:700 .95rem/1.35 var(--display); letter-spacing:.01em; }
.clip > summary .clip-title::before { content:''; position:absolute; left:.8rem; top:.95em;
  width:0; height:0; border-top:.33rem solid transparent; border-bottom:.33rem solid transparent;
  border-left:.5rem solid #555; }
.clip[open] > summary .clip-title::before { top:1.05em; border-left:.33rem solid transparent;
  border-right:.33rem solid transparent; border-top:.5rem solid #555; border-bottom:0; }
.clip > summary:hover .clip-title { background:#dcdcd8; }
.clip > summary:focus-visible .clip-title { outline:3px solid #ffbf47; outline-offset:-3px; }
.clip-body { padding:.6rem 0 1rem 2rem; }
.clip-body video { display:block; width:100%; max-width:640px; background:#000; }
.clip-body p { max-width:640px; margin:.4rem 0 0; }
.clips .jump { margin:1rem 0 .2rem; }
.key-video { margin:.5rem 0; }
.key-video summary { cursor:pointer; padding:.35rem 0; font-weight:bold; }
.key-video video { display:block; width:100%; max-width:640px; background:#000; }
.key-help { list-style:none; text-align:center; padding:0; margin:1rem 0; }
.key-nav { display:flex; justify-content:center; gap:2rem; margin:1.5rem 0 .5rem; flex-wrap:wrap; }
.key-nav a { padding:.5rem .25rem; }
dialog { max-width:min(92vw, 56rem); max-height:92vh; overflow:auto; border:3px solid #333; padding:1rem; }
dialog::backdrop { background:rgba(0,0,0,.65); }
dialog img { display:block; max-width:100%; height:auto; margin:0 auto .75rem; }
dialog .close { float:right; font:inherit; padding:.4rem .9rem; margin:0 0 .5rem .5rem; cursor:pointer; }
#fd-caption { clear:both; margin:0; }
@media (max-width: 44rem) {
  .choices.cols-3, .choices.cols-4 { grid-template-columns:repeat(2, minmax(0, 1fr)); }
  .frame { margin:0; border-width:2px; }
  .hero { grid-template-columns: 1fr; }
  .header-bar { justify-content:center; }
  .site-header nav ul { flex-wrap:wrap; justify-content:center; gap:.25rem 1rem; }
  .site-header nav a, .footer-nav a { display:inline-block; padding:.55rem .2rem; }
  h1 { font-size:1.45rem; }
  .species-name { font-size:1.35rem; }
  .group-strip { gap:1rem; }
  .group-strip li { width:calc(50% - .5rem); }
  .group-strip img { width:100%; max-width:150px; height:auto; }
  .species-list { columns:1; }
}
@media (prefers-reduced-motion: reduce) { * { scroll-behavior:auto !important; } }
"""


STARTER_INDEX = """---
title: Home
heading: An Image-Based Key to the Zooplankton of North America
align: center
---
**[Version 6.0](page:history)**

{{banner}}

## Contributors

Maria A. Aliberti, Elisha Allan, Stephanie Allard, Darren J. Bauer, Wendy Beagen, Shane R. Bradt, Brady Carlson, Sonya C. Carlson, Uyen Mai Doan, Jon Dufresne, W. Travis Godkin, Sara Greene, James F. Haney, Amy Kaplan, Erin Maroni, Shawn Melillo, Amanda L. Murby, Juliette L. Smith (Nowak), Brian Ortman, Judith E. Quist, Shayle Reed, Tiffany Rowin, Monika Schmuck, Richard S. Stemberger, Breanna Travers

[Contact](page:contact)

## About this key

Our key will be continually enlarged, updated and improved. We welcome others to contribute comments and images as well as additional taxa. Those interested should contact James Haney at [jfhaney@unh.edu](mailto:jfhaney@unh.edu).

{{aboutlinks}}
"""


STARTER_VIDEO = """---
title: Video Clips
heading: Video clips
old: html/movies.html
footer: yes
label: Video
order: 40
---
Short clips of living zooplankton, microscope views and sampling methods. The clips for each species also appear near the bottom of that species' page.

## Types of video in this key

1. **Live videos of organisms.** Videos were taken of organisms freshly collected in the field.
2. **Virtual Focus.** Videos capturing the microscopic anatomy of an individual organism. These interactive videos allow you to discern key features by adjusting the depth of focus on a particular characteristic of the organism.
3. **Sampling methods.** Videos covering a range of methods for collecting plankton. These videos demonstrate sampling techniques and include underwater views of how standard sampling devices operate in situ.
4. **Dissection demonstration.** To assist with dissection of calanoid and cyclopoid copepods, we have included video clips of procedures as seen through a dissecting microscope.
"""
VIDEO_GROUPS = [("Live organism", "Live videos of organisms"), ("Virtual Focus", "Virtual Focus"),
                ("Sampling method", "Sampling methods"), ("Dissection demonstration", "Dissection demonstration")]


BANNER_ITEMS = [
    ("cladocera", "Cladocera", "html/choices/01/frame1cladocera.jpg", "old:html/choices/cladocera/10/10.html"),
    ("copepoda", "Copepoda", "html/choices/01/frame1copepoda.jpg", "old:html/choices/copepoda/5/5.html"),
    ("rotifera", "Rotifera", "html/choices/01/frame1rotifera.jpg", "old:html/rotifers.htm"),
    ("other", "Other arthropods", "html/choices/01/frame1other.jpg", "old:html/choices/otherarthropods/51/51.html"),
]
CENTER_BY_DEFAULT = {"index", "contact"}     # pages the old site centered completely


def indent_wanted(meta, slug_, style):
    """Text pages indent the first line of each paragraph. "indent: no" turns it off;
    centred pages and reference lists never indent."""
    v = str(meta.get("indent", "")).strip().lower()
    if v[:1] in ("n", "f", "0"):
        return False
    if v[:1] in ("y", "t", "1"):
        return True
    if style == "references":
        return False
    align = str(meta.get("align", "")).strip().lower()
    if align == "center" or (not align and slug_ in CENTER_BY_DEFAULT):
        return False
    return True


def parse_front_matter(text):
    """Tiny front-matter reader: lines of 'key: value' between two '---' lines."""
    meta = {}
    if text.startswith("---"):
        end = text.find("\n---", 3)
        if end != -1:
            for line in text[3:end].strip("\n").splitlines():
                if ":" in line:
                    k, v = line.split(":", 1)
                    meta[k.strip()] = v.strip()
            return meta, text[end + 4:].lstrip("\n")
    return meta, text


def load_content(folder: Path):
    pages = {}
    for f in sorted(folder.glob("*.md")):
        meta, body = parse_front_matter(f.read_text(encoding="utf-8"))
        slug_ = meta.get("slug") or f.stem
        pages[slug_] = {"meta": meta, "body": body, "file": f.name}
    return pages


# ----------------------------------------------------------------------------
def slug(s):
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def sheet_rows(wb, name):
    if name not in wb.sheetnames:
        return []
    it = wb[name].iter_rows(values_only=True)
    try:
        hdr = [str(h) if h is not None else "" for h in next(it)]
    except StopIteration:
        return []
    out = []
    for row in it:
        if row is None or all(v in (None, "") for v in row):
            continue
        out.append({h: ("" if v is None else v) for h, v in zip(hdr, row)})
    return out


def truthy(v) -> bool:
    return str(v).strip().lower() in ("true", "1", "yes")


def to_int(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def render_text(text, resolver):
    """Light Markdown (*italic*, **bold**, [text](target), <sup>/<sub>) -> safe HTML."""
    s = html.escape(str(text), quote=False)
    s = re.sub(r"&lt;(/?)(sup|sub)&gt;", r"<\1\2>", s)

    def link(m):
        label, target = m.group(1), html.unescape(m.group(2))
        res = resolver(target)
        if not res:
            return label
        href, attrs = res if isinstance(res, tuple) else (res, "")
        return f'<a href="{html.escape(href, quote=True)}"{attrs}>{label}</a>'

    s = re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", link, s)
    s = re.sub(r"\{(red|blue|green|orange)\|([^}]*)\}", r'<span class="mk mk-\1">\2</span>', s)
    s = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", s)
    s = re.sub(r"(?<![\w*])\*(?!\s)(.+?)(?<!\s)\*(?!\*)", r"<em>\1</em>", s)
    return Markup(s)


class PageChecker(HTMLParser):
    def __init__(self):
        super().__init__()
        self.h1 = 0
        self.imgs_no_alt = 0
        self.ids, self.dup_ids, self.refs = set(), set(), []
        self.videos = 0

    def handle_starttag(self, tag, attrs):
        a = dict(attrs)
        if tag == "h1":
            self.h1 += 1
        if tag == "img":
            decorative = ("logo" in (a.get("class") or "").split()) or ("data-decorative" in a) or a.get("id") in ("fd-img", "entry-img")   # alt="" is right for these
            if "alt" not in a or (not (a.get("alt") or "").strip() and not decorative):
                self.imgs_no_alt += 1
            if a.get("src"):
                self.refs.append(a["src"])
        if tag == "video":
            self.videos += 1
        if tag == "source" and a.get("src"):
            self.refs.append(a["src"])
        if tag == "a" and a.get("href"):
            self.refs.append(a["href"])
        if tag == "link" and a.get("href"):
            self.refs.append(a["href"])
        if a.get("id"):
            if a["id"] in self.ids:
                self.dup_ids.add(a["id"])
            self.ids.add(a["id"])


def check_site(out: Path):
    problems = []
    n_pages = 0
    for f in sorted(out.rglob("*.html")):
        n_pages += 1
        p = PageChecker()
        p.feed(f.read_text(encoding="utf-8"))
        rel = f.relative_to(out).as_posix()
        if p.h1 != 1:
            problems.append(f"{rel}: has {p.h1} <h1> (should be exactly 1)")
        if p.imgs_no_alt:
            problems.append(f"{rel}: {p.imgs_no_alt} image(s) with empty alt")
        for d in sorted(p.dup_ids):
            problems.append(f"{rel}: duplicate id '{d}'")
        for ref in p.refs:
            if re.match(r"^(https?:|mailto:|#)", ref):
                continue
            target = (f.parent / ref.split("#")[0]).resolve()
            if not target.exists():
                problems.append(f"{rel}: broken local link/file -> {ref}")
    return n_pages, problems


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--old-site", required=True, help="the cfbkey_local folder (source of image files)")
    ap.add_argument("--out", required=True)
    ap.add_argument("--only", default="")
    ap.add_argument("--limit", type=int, default=0)
    ap.add_argument("--content", default="")
    ap.add_argument("--noindex", action="store_true",
                    help="ask search engines to ignore this copy (use for test and preview sites)")
    ap.add_argument("--init", action="store_true")
    a = ap.parse_args()

    out = Path(a.out).expanduser().resolve()
    old = Path(a.old_site).expanduser().resolve()
    src = out.parent / "site_src"

    if a.init:
        for name, text in TEMPLATES.items():
            f = src / "templates" / name
            f.parent.mkdir(parents=True, exist_ok=True)
            f.write_text(text, encoding="utf-8")
        (src / "static" / "css").mkdir(parents=True, exist_ok=True)
        (src / "static" / "css" / "site.css").write_text(CSS, encoding="utf-8")
        print(f"Templates and stylesheet written to {src}")
        return

    env = Environment(
        loader=ChoiceLoader([FileSystemLoader(str(src / "templates")), DictLoader(TEMPLATES)]),
        autoescape=select_autoescape(["html"]), trim_blocks=True, lstrip_blocks=True)

    wb = load_workbook(a.workbook, data_only=True, read_only=True)
    species = sheet_rows(wb, "species")
    sections = sheet_rows(wb, "species_sections")
    images = sheet_rows(wb, "images")
    figures = sheet_rows(wb, "figures")
    videos = {r["video_id"]: r for r in sheet_rows(wb, "videos")}
    sp_videos = sheet_rows(wb, "species_videos")
    n_extra = len(sheet_rows(wb, "extra_pages"))

    if a.only:
        wanted = {x.strip() for x in a.only.split(",")}
        species = [s for s in species if s["species_id"] in wanted]
    if a.limit:
        species = species[:a.limit]
    all_species = sheet_rows(wb, "species")
    tax_pref = {(str(r["rank"]), str(r["value"])): str(r["preferred_spelling"]).strip()
                for r in sheet_rows(wb, "taxonomy") if str(r.get("preferred_spelling") or "").strip()}
    n_tax_fixed = 0
    for lst in (all_species, species):
        for sp_ in lst:
            for rk in RANKS:
                fixed = tax_pref.get((rk, str(sp_.get(rk) or "")))
                if fixed:
                    sp_[rk] = fixed
                    n_tax_fixed += 1
    page_of_old = {s["old_path"]: s["species_id"] for s in all_species}
    built_ids = {s["species_id"] for s in species}

    report = defaultdict(list)
    extra_path_ref = {}

    # ---- Markdown content pages (home, instructions, history, ...) ----
    content_dir = Path(a.content).expanduser().resolve() if a.content else out.parent / "content"
    content_dir.mkdir(parents=True, exist_ok=True)
    if not (content_dir / "index.md").exists():
        (content_dir / "index.md").write_text(STARTER_INDEX, encoding="utf-8")
        print(f"(Wrote a starter home page to {content_dir / 'index.md'}: edit it freely.)")
    if not (content_dir / "video.md").exists():
        (content_dir / "video.md").write_text(STARTER_VIDEO, encoding="utf-8")
    content = load_content(content_dir)
    registry = {"species": "species/index.html"}
    for slug_ in content:
        registry[slug_] = "index.html" if slug_ == "index" else f"{slug_}/index.html"

    def nav_items(kind):
        items = []
        if kind == "main":
            items.append({"label": "Home", "slug": "index", "order": 0})
            items.append({"label": "Species", "slug": "species", "order": 40})
            if "key" in registry:
                items.append({"label": "Use the Key", "slug": "key", "order": 20})
            if "groups" in registry:
                items.append({"label": "Groups", "slug": "groups", "order": 30})
            for sl, pg in content.items():
                if sl != "index" and pg["meta"].get("nav"):
                    items.append({"label": pg["meta"]["nav"], "slug": sl,
                                  "order": int(pg["meta"].get("order") or 50)})
        else:
            for sl, pg in content.items():
                if pg["meta"].get("footer", "").lower() in ("yes", "true") and sl != "index":
                    items.append({"label": pg["meta"].get("label") or pg["meta"].get("nav") or pg["meta"].get("title", sl),
                                  "slug": sl, "order": int(pg["meta"].get("order") or 50)})
        return sorted(items, key=lambda i: i["order"])

    def common(page_dir, current):
        depth = 0 if not page_dir else page_dir.count("/") + 1
        mk = lambda items: [{"label": i["label"], "current": i["slug"] == current,
                             "href": posixpath.relpath(registry[i["slug"]], page_dir or ".")}
                            for i in items if i["slug"] in registry]
        has_logo = (out / "img" / "site" / "logo.jpg").exists()
        return {"site": SITE, "root": "../" * depth, "noindex": a.noindex, "nav": mk(nav_items("main")),
                "footer_nav": mk(nav_items("footer")), "current": current,
                "logo": posixpath.relpath("img/site/logo.jpg", page_dir or ".") if has_logo else ""}

    # ---- the identification key ----
    key_rows = sheet_rows(wb, "key_steps")
    key_info = {str(r["step_key"]): r for r in sheet_rows(wb, "key_steps_info")}
    key_help = sheet_rows(wb, "key_help_links")
    key_imgs = sheet_rows(wb, "key_images_of_text")
    other_vl = sheet_rows(wb, "other_video_links")
    steps = OrderedDict()
    for r in key_rows:
        steps.setdefault(str(r["step_key"]), []).append(r)
    step_lookup = {k.lower(): k for k in steps}
    parents = defaultdict(list)
    for k, rows_ in steps.items():
        for r in rows_:
            if r.get("goes_to_type") == "key_step" and r.get("goes_to") in steps and k not in parents[r["goes_to"]]:
                parents[r["goes_to"]].append(k)

    def step_path(sk):
        if sk == "entry/begin":
            return "key/index.html"
        if sk.startswith("entry/"):
            return f"key/{sk[6:].lower()}/index.html"
        parts = sk.split("/")
        if len(parts) >= 2 and parts[-1] == parts[-2]:
            parts = parts[:-1]
        return "key/" + "/".join(p.lower() for p in parts) + "/index.html"

    def old_to_step(ref):
        r = re.sub(r"\.html?$", "", ref.lower())
        if r == "html/begin":
            return "entry/begin" if "entry/begin" in steps else None
        if r == "html/rotifers":
            return "entry/rotifers" if "entry/rotifers" in steps else None
        if r.startswith("html/choices/"):
            return step_lookup.get(r[len("html/choices/"):])
        return None

    if steps:
        registry["key"] = step_path("entry/begin") if "entry/begin" in steps else step_path(next(iter(steps)))

    # ---- anatomy diagrams ----
    ana_rows = sheet_rows(wb, "anatomy")
    by_ana = defaultdict(list)
    for r in ana_rows:
        by_ana[str(r["group"])].append(r)
    ANA_TITLES = {"bosmina": "Bosmina", "daphnia": "Daphnia", "copepod": "Copepod",
                  "rotiferA": "Rotifer (A)", "rotiferB": "Rotifer (B)"}
    ANA_SECTIONS = [("Cladocera", ["bosmina", "daphnia"]), ("Copepoda", ["copepod"]),
                    ("Rotifera", ["rotiferA", "rotiferB"])]
    ana_groups = sorted(by_ana, key=lambda g: list(ANA_TITLES).index(g) if g in ANA_TITLES else 99)

    def ana_path(g):
        return f"anatomy/{slug(g)}/index.html"

    if by_ana:
        registry["anatomy-diagrams"] = ana_path(ana_groups[0])

    # ---- the Groups outline ----
    groups_rows = sheet_rows(wb, "groups")
    by_group = defaultdict(list)
    for r in groups_rows:
        by_group[str(r["group_page"])].append(r)
    GROUP_TITLES = {"overview": "All groups", "cladocera": "Cladocera", "copepoda": "Copepoda",
                    "rotifera": "Rotifera", "arthropoda": "Other arthropods", "genera": "Genera"}
    group_pages = [g for g in GROUP_TITLES if g in by_group] + [g for g in by_group if g not in GROUP_TITLES]

    def groups_path(g):
        return "groups/index.html" if g == "overview" else f"groups/{g}/index.html"

    if by_group:
        registry["groups"] = groups_path("overview" if "overview" in by_group else group_pages[0])
    # names from the Groups outline, used as draft labels for picture-only key choices
    group_label_for_step = {}
    for g in [x for x in group_pages if x != "overview"] + (["overview"] if "overview" in by_group else []):
        for r in sorted(by_group[g], key=lambda r_: int(r_["seq"])):
            if r.get("goes_to_type") == "key_step" and r.get("goes_to") not in group_label_for_step:
                group_label_for_step[str(r["goes_to"])] = str(r.get("name_new") or r["name"])

    # old addresses -> new pages (content pages declare theirs with  old: html/x.html)
    old_to_slug = {"html/species.html": "species", "html/movies.html": "video",
                   "html/index.html": "index", "start.html": "index"}
    for sl, pg in content.items():
        for o in re.split(r"[,;]\s*", pg["meta"].get("old", "")):
            if o:
                old_to_slug[o.lower()] = sl

    # clips for the Video page (one per distinct clip; duplicates point to their keeper)
    clip_groups, clip_anchor = [], {}
    by_type = defaultdict(list)
    for vid, v in videos.items():
        if v.get("duplicate_of") or not v.get("archive_url"):
            continue
        if not (v.get("title_on_old_site") or v.get("description_new")):
            # not on the old video index: list it only once someone has written a description
            report["clips left off the Video page (not on the old index; write a description_new to list one)"].append(vid)
            continue
        by_type[v.get("video_type") or "Other"].append((vid, v))
    for key, title in VIDEO_GROUPS + [("Other", "Other clips")]:
        clips = []
        for vid, v in sorted(by_type.get(key, []), key=lambda kv: str(kv[1].get("title_on_old_site") or kv[0]).lower()):
            anchor = slug(vid)
            clip_anchor[vid] = anchor
            desc = v.get("description_new") or v.get("description_on_old_site") or ""
            if not v.get("description_new"):
                report["clips on the Video page without a new description"].append(vid)
            if not desc and not v.get("title_on_old_site"):
                report["clips with no title or description at all"].append(vid)
            plain = lambda t: re.sub(r"[\s*]+", " ", str(t)).strip().lower()
            if desc and plain(desc) == plain(v.get("title_on_old_site") or ""):
                desc = ""                      # the caption would only repeat the title
            clips.append({"anchor": anchor, "url": v["archive_url"],
                          "title": render_text(v.get("title_on_old_site") or posixpath.basename(vid).replace("_", " "),
                                               lambda t: None),
                          "description": render_text(desc, lambda t: None)})
        if clips:
            clip_groups.append({"title": title, "clips": clips})

    def fix_page_links(text, page_dir, where):
        def target(kind, ref):
            if kind == "page":
                return registry.get(ref)
            if kind == "old":
                am = re.match(r"^html/anatomy/([^/]+)/[^/]+\.html?$", ref, re.I)
                if am:
                    for g_ in by_ana:
                        if g_.lower() == am.group(1).lower():
                            return ana_path(g_)
                    return None
                gm = re.match(r"^html/groups(?:_(\w+))?\.html?$", ref, re.I)
                if gm and by_group:
                    g_ = (gm.group(1) or "overview").lower()
                    return groups_path(g_) if g_ in by_group else None
                sk = old_to_step(ref)
                if sk:
                    return step_path(sk)
                sl = old_to_slug.get(ref.lower())
                if sl in registry:
                    return registry[sl]
                if ref in page_of_old and page_of_old[ref] in built_ids:
                    return f"species/{page_of_old[ref]}/index.html"
                if ref in extra_path_ref:
                    return extra_path_ref[ref]
                return None
            v = videos.get(ref)
            cid = (v.get("duplicate_of") or ref) if v else None
            if cid in clip_anchor and "video" in registry:
                return registry["video"] + "#" + clip_anchor[cid]
            return None

        def sub2(m):
            pre, kind, ref, post, inner = m.group(1), m.group(2), html.unescape(m.group(3)), m.group(4), m.group(5)
            url = target(kind, ref)
            if url is None:
                report["links to pages or clips that do not exist yet (shown as plain text)"].append(
                    f"{where}: {kind}:{ref}")
                return inner
            path, _, frag = url.partition("#")
            rel_url = posixpath.relpath(path, page_dir or ".") + (("#" + frag) if frag else "")
            return f'<a {pre}href="{html.escape(rel_url, quote=True)}"{post}>{inner}</a>'

        def sub(m):
            kind, ref, inner = m.group(1), html.unescape(m.group(2)), m.group(3)
            url = target(kind, ref)
            if url is None:
                report["links to pages or clips that do not exist yet (shown as plain text)"].append(
                    f"{where}: {kind}:{ref}")
                return inner
            path, _, frag = url.partition("#")
            rel_url = posixpath.relpath(path, page_dir or ".") + (("#" + frag) if frag else "")
            return f'<a href="{html.escape(rel_url, quote=True)}">{inner}</a>'
        return re.sub(r'<a ([^>]*?)href="(page|old|video):([^"]*)"([^>]*)>(.*?)</a>',
                      lambda m: sub2(m), text, flags=re.S)

    for d in ("species", "img", "key", "groups", "anatomy", "pages"):
        shutil.rmtree(out / d, ignore_errors=True)
    (out / "css").mkdir(parents=True, exist_ok=True)
    css_src = src / "static" / "css" / "site.css"
    (out / "css" / "site.css").write_text(css_src.read_text(encoding="utf-8") if css_src.exists() else CSS,
                                          encoding="utf-8")

    shutil.rmtree(out / "assets", ignore_errors=True)
    if (content_dir / "assets").exists():
        shutil.copytree(content_dir / "assets", out / "assets")
    banner_have = []
    for key, label, rel_img, target in BANNER_ITEMS:
        if (old / rel_img).exists():
            (out / "img" / "site").mkdir(parents=True, exist_ok=True)
            shutil.copy2(old / rel_img, out / "img" / "site" / f"group-{key}.jpg")
            banner_have.append((key, label, target))
    logo_src = old / "images" / "indexmain1.jpg"
    if logo_src.exists():
        (out / "img" / "site").mkdir(parents=True, exist_ok=True)
        shutil.copy2(logo_src, out / "img" / "site" / "logo.jpg")

    imgs_by_sp, figs_by_sp, secs_by_sp, vids_by_sp = (defaultdict(list) for _ in range(4))
    for r in images:
        imgs_by_sp[r["species_id"]].append(r)
    for r in figures:
        figs_by_sp[r["species_id"]].append(r)
    for r in sections:
        secs_by_sp[r["species_id"]].append(r)
    for r in sp_videos:
        vids_by_sp[r["species_id"]].append(r)
    video_of_wrapper = {r["old_link"]: r["video_id"] for r in sp_videos if r.get("video_id")}

    copied = {}

    def copy_image(old_rel, owner):
        """Copy an image into the site once; return its path relative to the site root."""
        if not old_rel:
            return ""
        key = (owner, old_rel)
        if key in copied:
            return copied[key]
        name = posixpath.basename(old_rel)
        dest_rel = f"img/{owner}/{name}"
        n = 1
        while dest_rel in copied.values():
            n += 1
            dest_rel = f"img/{owner}/{n}-{name}"
        srcf = old / old_rel
        if srcf.exists():
            (out / dest_rel).parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(srcf, out / dest_rel)
        else:
            report["missing image file"].append(f"{owner}: {old_rel}")
        copied[key] = dest_rel
        return dest_rel

    def build_species(sp):
        sid = sp["species_id"]
        page_dir = f"species/{sid}"
        root = "../../"
        rel = lambda site_path: posixpath.relpath(site_path, page_dir) if site_path else ""

        # images: main, two thumbnails, then the rest
        imgs = [r for r in imgs_by_sp.get(sid, []) if not r.get("on_extra_page")]
        cards = []
        for i, r in enumerate(imgs, 1):
            alt = str(r.get("alt_text_new") or "").strip()
            if not alt:
                alt = str(r.get("alt_text_draft") or "").strip()
                if alt:
                    report["pictures using DRAFT alt text (needs review)"].append(r["image_id"])
            if not alt:
                alt = sp["display_name"] if i == 1 else f"{sp['display_name']}, additional view {i}"
                report["pictures with no alt text at all"].append(r["image_id"])
            src_rel = copy_image(r["file"], sid)
            large_rel = copy_image(r.get("links_to_larger"), sid) or src_rel
            cards.append({"src": rel(src_rel), "large": rel(large_rel), "alt": alt,
                          "w": to_int(r.get("width")), "h": to_int(r.get("height"))})
        main = cards[0] if cards else None
        thumbs = [c for c in cards[1:3] if (c["w"] or 0) <= 250]
        more = [c for c in cards[1:] if c not in thumbs]

        # videos shown on this page
        vlist, seen = [], set()
        for r in vids_by_sp.get(sid, []):
            vid = r.get("video_id")
            if not vid or vid in seen:
                continue
            seen.add(vid)
            v = videos.get(vid)
            if not v or not v.get("archive_url"):
                report["videos not uploaded / unknown"].append(f"{sid}: {vid}")
                continue
            desc = v.get("description_new") or v.get("description_on_old_site") or ""
            if not v.get("description_new"):
                report["videos with no new description (old text used)" if desc else
                       "videos with no description at all"].append(vid)
            vlist.append({"url": v["archive_url"],
                          "title": render_text(v.get("title_on_old_site") or "", lambda t: None),
                          "description": render_text(desc, lambda t: None), "id": vid})
        anchor_of_video = {v["id"]: f"#v-{i}" for i, v in enumerate(vlist, 1)}

        figure_url, figure_alt = {}, {}
        for r in figs_by_sp.get(sid, []):
            figure_url[r["image"]] = rel(copy_image(r["image"], sid))
            fa = str(r.get("alt_text_new") or "").strip()
            if not fa:
                fa = str(r.get("alt_text_draft") or "").strip()
                report["detail figures using DRAFT alt text (needs review)" if fa else
                       "detail figures with no alt text at all"].append(f"{sid}: {r['image']}")
            figure_alt[r["image"]] = fa or f"{sp['display_name']}: detail figure"
        for r, c in zip(imgs, cards):
            if r.get("links_to_larger"):
                figure_url.setdefault(r["links_to_larger"], rel(copy_image(r["links_to_larger"], sid)))
                figure_alt.setdefault(r["links_to_larger"], c["alt"])

        def resolver(target):
            if re.match(r"^(https?:|mailto:)", target):
                return target
            kind, _, path = target.partition(":")
            if kind == "fig":
                alt_attr = f' data-figure-alt="{html.escape(figure_alt.get(path, sp["display_name"] + ": detail figure"), quote=True)}"'
                if path in figure_url:
                    return figure_url[path], alt_attr
                if (old / path).exists():
                    return rel(copy_image(path, sid)), alt_attr
            elif kind == "page" and path in extra_path_ref:
                return rel(extra_path_ref[path])
            elif kind == "page" and path in page_of_old:
                other = page_of_old[path]
                if other in built_ids:
                    return posixpath.relpath(f"species/{other}/index.html", page_dir)
            elif kind == "video":
                vid = video_of_wrapper.get(path) or (path[:-4] if path.lower().endswith(".mov") else "")
                if vid in anchor_of_video:
                    return anchor_of_video[vid]
            report["links in text that could not be resolved"].append(f"{sid}: {target}")
            return None

        # text sections (only the species page's own text; extra pages come later)
        secs = OrderedDict()
        for r in secs_by_sp.get(sid, []):
            if r.get("from_page"):
                continue
            title = str(r.get("section") or "")
            if title.lower() in SKIP_SECTIONS:
                continue
            sec = secs.setdefault(title, OrderedDict())
            sec.setdefault(str(r.get("subsection") or ""), []).append(render_text(r["text"], resolver))
        sec_list = [{"title": t, "is_refs": t.lower() == "references",
                     "blocks": [{"title": bt, "entries": entries} for bt, entries in blocks.items()]}
                    for t, blocks in secs.items()]

        taxonomy = [(k, sp[k]) for k in RANKS if sp.get(k)]
        html_out = env.get_template("species.html").render(
            **common(page_dir, "species"), title=sp["display_name"],
            description=f"{sp['display_name']}: images, diagnostic characters, ecology and videos.",
            sp=sp, main=main, thumbs=thumbs, more=more, taxonomy=taxonomy,
            sections=sec_list, videos=vlist)
        f = out / page_dir / "index.html"
        f.parent.mkdir(parents=True, exist_ok=True)
        f.write_text(html_out, encoding="utf-8")

    for sp in species:
        build_species(sp)

    # species index (all species; only the built ones are links)
    groups = defaultdict(list)
    for s in all_species:
        groups[s.get("Family") or "Family not recorded"].append({
            "display_name": s["display_name"], "built": s["species_id"] in built_ids,
            "url": f"{s['species_id']}/index.html"})
    glist = sorted(groups.items(), key=lambda kv: (kv[0] == "Family not recorded", kv[0].lower()))
    for _, items in glist:
        items.sort(key=lambda x: x["display_name"].lower())
    (out / "species").mkdir(parents=True, exist_ok=True)
    (out / "species" / "index.html").write_text(env.get_template("species_index.html").render(
        **common("species", "species"), title="Species",
        description="All species in the key, grouped by family.",
        groups=glist, n_species=len(all_species)), encoding="utf-8")

    for sl, pg in content.items():
        page_dir = "" if sl == "index" else sl
        body_html = md_lib.markdown(pg["body"], extensions=["extra", "sane_lists", "attr_list", "toc"])
        body_html = fix_page_links(body_html, page_dir, f"{pg['file']}")
        style = str(pg["meta"].get("style", "")).lower()
        if not style:
            # a page with several standalone italic sub-headings among many paragraphs is a
            # reference list, whether or not anyone set "style: references"
            n_italic = len(re.findall(r"<p>\s*<em>[^<]*</em>\s*</p>", body_html))
            n_paras = body_html.count("<p>")
            if n_italic >= 2 and n_paras >= n_italic + 2:
                style = "references"
                report["pages formatted as reference lists (add 'style: none' to turn this off)"].append(pg["file"])
        # a paragraph that starts "1st:", "2nd:" ... is a release line: keep it on one line
        body_html = re.sub(r'<p>(\s*(?:<a href="#release-[^"]*">[^<]*</a>|\d+(?:st|nd|rd|th))\s*:)',
                           r'<p class="release-line">\1', body_html)
        if style == "references":
            # a paragraph that is entirely italic is a sub-heading; the rest are references
            body_html = re.sub(r"<p>(\s*<em>[^<]*</em>\s*)</p>", r'<p class="ref-head">\1</p>', body_html)
        prefix = "../" * (0 if not page_dir else page_dir.count("/") + 1)
        has_buttons = "{{aboutlinks}}" in body_html
        if has_buttons:
            btns = "".join(
                f'<li><a class="link-button" href="{html.escape(n["href"], quote=True)}">{html.escape(n["label"])}</a></li>'
                for n in common(page_dir, sl)["footer_nav"])
            body_html = re.sub(r"<p>\s*\{\{aboutlinks\}\}\s*</p>|\{\{aboutlinks\}\}",
                               f'<ul class="link-buttons">{btns}</ul>' if btns else "", body_html)
            if not btns:
                report["{{aboutlinks}} used but no footer links exist yet"].append(pg["file"])
        if "{{banner}}" in body_html:
            pics = "".join(f'<img src="{prefix}img/site/group-{k}.jpg" alt="" data-decorative '
                           f'width="200" height="200">' for k, l, t in banner_have)
            strip = (f'<a class="banner-link" href="page:key"><span class="banner-strip">{pics}'
                     f'<span class="banner-overlay">Use the key</span></span></a>') if pics else ""
            if not pics:
                report["home page banner: group pictures not found in html/choices/01/"].append("frame1*.jpg")
            body_html = re.sub(r"<p>\s*{{banner}}\s*</p>|{{banner}}", lambda m: fix_page_links(strip, page_dir, "home banner"), body_html)
        elif sl == "index" and banner_have:
            report["home page banner not shown (add the line {{banner}} to content/index.md)"].append("index.md")
        body_html = re.sub(r'(src|href)="assets/', lambda m: f'{m.group(1)}="{prefix}assets/', body_html)
        title = pg["meta"].get("title") or sl
        if re.search(r"image-?based key", title, re.I) and sl != "index":
            m_h = re.search(r"<h2[^>]*>(.*?)</h2>", body_html, re.S)
            if m_h:
                title = re.sub(r"<[^>]+>", "", m_h.group(1)).strip()
                body_html = body_html[:m_h.start()] + body_html[m_h.end():]
                report["pages whose title came from their own first heading (the old <title> was the "
                       "generic site name)"].append(f"{pg['file']} -> {title}")
        heading = pg["meta"].get("heading") or title
        dest = out / registry[sl]
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(env.get_template("page.html").render(
            **common(page_dir, sl), title=SITE["title"] if sl == "index" else title,
            heading=heading, body=Markup(body_html),
            # text pages indent the first line of each paragraph by default; "indent: no" turns
            # it off, and centred pages and reference lists are never indented
            indent=indent_wanted(pg["meta"], sl, style),
            page_style=("" if style == "none" else style),
            wide=(str(pg["meta"].get("width", "")).lower() in ("wide", "full")),
            centered=(pg["meta"].get("align", "").lower() == "center")
                     or (not pg["meta"].get("align") and sl in CENTER_BY_DEFAULT),
            clip_groups=clip_groups if sl == "video" else None,
            hide_footer_nav=has_buttons,
            description=pg["meta"].get("description") or SITE["title"]), encoding="utf-8")


    # ================= the identification key: one page per step =================
    GROUP_LABEL = {"cladocera": "Cladocera", "copepoda": "Copepoda", "rotifera": "Rotifera",
                   "otherarthropods": "Other arthropods"}
    species_by_id = {s_["species_id"]: s_ for s_ in all_species}

    def tax_pairs(txt):
        out_ = []
        for part in str(txt or "").split(";"):
            if "=" in part:
                a_, b_ = part.split("=", 1)
                out_.append((a_.strip(), b_.strip()))
        return out_

    def popup_image(target):
        """An old help link that opens a page which is only a picture -> that picture's path."""
        f_ = old / target if target else None
        if not (f_ and f_.exists() and target.lower().endswith((".htm", ".html"))):
            return None
        t_ = f_.read_bytes().decode("cp1252", errors="replace")
        imgs_ = re.findall(r'<img[^>]+src="([^"]+)"', t_, flags=re.I)
        text_ = re.sub(r"(?is)<(script|style).*?</\1>", " ", t_)
        text_ = html.unescape(re.sub(r"<[^>]+>", " ", text_))
        if len(imgs_) == 1 and len(re.sub(r"\s+", " ", text_).strip()) < 60:
            return posixpath.normpath(posixpath.join(posixpath.dirname(target), imgs_[0].replace("\\", "/")))
        return None

    n_key_pages = 0
    for sk, rows_ in steps.items():
        info = key_info.get(sk, {})
        out_rel = step_path(sk)
        page_dir = posixpath.dirname(out_rel)
        rel_ = lambda p_, _d=page_dir: posixpath.relpath(p_, _d) if p_ else ""
        parts_ = sk.split("/")
        entry = sk.startswith("entry/")
        taxp = [(a_, tax_pref.get((a_, b_), b_)) for a_, b_ in tax_pairs(info.get("taxonomy_at_step"))]
        if str(info.get("step_heading_new") or "").strip():
            heading = str(info["step_heading_new"]).strip()
        elif sk == "entry/begin":
            heading = "Use the key"
        elif sk == "entry/rotifers":
            heading = "Rotifera: choose a group"
        else:
            heading = f"{GROUP_LABEL.get(parts_[0], parts_[0].title())} key, step {parts_[-1]}"
        prompt = str(info.get("step_prompt") or rows_[0].get("prompt") or
                     "Click on the image which your specimen most closely resembles.")

        choices, any_plain, preview_mode = [], False, False
        # a page of group pictures: no choice carried its own text on the old site
        picture_only = all(not str(r_.get("choice_text") or "").strip() for r_ in rows_)
        if sk == "entry/begin" and not str(info.get("step_prompt") or "").strip():
            prompt = ("Point at, or tab to, a group on the left to see example organisms from it. "
                      "On a touch screen, use the Show examples buttons. "
                      "Then select the group that most closely resembles your specimen.")
        for r in sorted(rows_, key=lambda r_: int(r_.get("choice_no") or 0)):
            imgs_out = []
            for k_, old_img in enumerate([str(r.get("image") or "")] +
                                         [x.strip() for x in str(r.get("extra_images") or "").split(";")]):
                if not old_img.strip():
                    continue
                plain_ = ""
                if k_ == 0 and not entry and str(r.get("hover_image") or "").strip():
                    plain_ = rel_(copy_image(str(r["hover_image"]).strip(), "key"))
                    any_plain = True
                imgs_out.append({"src": rel_(copy_image(old_img.strip(), "key")), "plain": plain_, "alt": ""})
            if imgs_out:
                imgs_out[0]["alt"] = str(r.get("alt_text_new") or "").strip()
            label = str(r.get("choice_label_new") or "").strip()
            text = label or str(r.get("choice_text") or "").strip()
            if not text and not (imgs_out and imgs_out[0]["alt"]):
                tg_, tgt_type = str(r.get("goes_to") or ""), str(r.get("goes_to_type") or "")
                if tgt_type == "key_step" and tg_ in group_label_for_step:
                    text = group_label_for_step[tg_]
                    report["picture-only key choices labeled from the Groups outline (check, then fill choice_label_new)"].append(f"{sk} choice {r.get('choice_no')}: {text}")
                elif tgt_type in ("organism_page", "organism_page_unlisted") and tg_ in species_by_id:
                    text = "*" + str(species_by_id[tg_]["display_name"]) + "*"
                    report["picture-only key choices labeled with the species name (check, then fill choice_label_new)"].append(f"{sk} choice {r.get('choice_no')}: {text}")
                elif tgt_type == "key_step" and tg_ in steps:
                    text = f"Continue to step {tg_.split('/')[-1]}"
                    report["picture-only key choices with only a placeholder label (fill choice_label_new)"].append(f"{sk} choice {r.get('choice_no')}")
            lines = [render_text(x, lambda t_: None) for x in text.split(" | ") if x.strip()] if text else []
            if not text and not imgs_out[0]["alt"] if imgs_out else not text:
                report["key choices with no visible text and no alt text (need a label)"].append(f"{sk} choice {r.get('choice_no')}")

            t_, tt_ = str(r.get("goes_to") or ""), str(r.get("goes_to_type") or "")
            url, go = None, ""
            if tt_ == "key_step" and t_ in steps:
                url = rel_(step_path(t_))
                go = "Go to the rotifer groups" if t_ == "entry/rotifers" else f"Go to step {t_.split('/')[-1]}"
            elif tt_ in ("organism_page", "organism_page_unlisted"):
                sid_ = t_ if t_ in species_by_id else page_of_old.get(t_)
                if sid_ in built_ids:
                    url = rel_(f"species/{sid_}/index.html")
                    go = Markup("Species: <em>" + html.escape(species_by_id[sid_]["display_name"]) + "</em>")
                elif sid_ in species_by_id:
                    report["key choices pointing at species pages not built in this run"].append(f"{sk} -> {sid_}")
                else:
                    report["key choices whose destination page was not found"].append(f"{sk} choice {r.get('choice_no')}: {t_}")
            else:
                sl_ = old_to_slug.get(t_.lower())
                if sl_ in registry:
                    url = rel_(registry[sl_])
                    go = "Read more"
                else:
                    report["key choices whose destination page does not exist yet"].append(f"{sk} choice {r.get('choice_no')}: {t_}")
            name = re.sub(r"\{\w+\|([^}]*)\}", r"\1", text.replace(" | ", " ")).replace("*", "").strip()
            preview = preview_alt = ""
            if sk == "entry/begin" and str(r.get("hover_image") or "").strip() and imgs_out:
                preview = rel_(copy_image(str(r["hover_image"]).strip(), "key"))
                preview_alt = str(r.get("alt_text_new") or "").strip()
                if not preview_alt:
                    preview_alt = f"Example photographs of organisms in the group {name}"
                    report["group preview pictures with only a generic description (write alt_text_new)"].append(name)
                preview_mode = True
            choices.append({"images": imgs_out, "lines": lines, "url": url, "go": go, "name": name,
                            "preview": preview, "preview_alt": preview_alt})

        vl, seen_v = [], set()
        for r in other_vl:
            if r.get("shown_on_type") == "key step" and r.get("shown_on") == sk and r.get("video_id") not in seen_v:
                seen_v.add(r["video_id"])
                v = videos.get(r["video_id"])
                v = videos.get(v.get("duplicate_of")) if v and v.get("duplicate_of") else v
                if v and v.get("archive_url"):
                    lab = str(r.get("label_on_old_site") or v.get("title_on_old_site") or "").strip()
                    lab = re.sub(r"[*]", "", lab) or posixpath.basename(r["video_id"]).replace("_", " ")
                    desc = v.get("description_new") or v.get("description_on_old_site") or ""
                    vl.append({"label": "Video: " + lab, "url": v["archive_url"],
                               "description": render_text(desc, lambda t_: None) if desc and desc != lab else ""})

        helps = []
        for r in key_help:
            if str(r.get("step_key")) != sk:
                continue
            txt_, tgt_ = str(r.get("text") or ""), str(r.get("target") or "")
            pic = tgt_ if tgt_.lower().endswith((".jpg", ".jpeg", ".png", ".gif")) else popup_image(tgt_)
            if pic and (old / pic).exists():
                src_ = rel_(copy_image(pic, "key"))
                plain_txt = re.sub(r"\*", "", txt_)
                helps.append(Markup(f'<a href="{html.escape(src_, quote=True)}" data-figure-alt="{html.escape(plain_txt, quote=True)}">'
                                    f'{render_text(txt_, lambda t_: None)}</a>'))
            else:
                linked = fix_page_links(f'<a href="old:{html.escape(tgt_, quote=True)}">{render_text(txt_, lambda t_: None)}</a>', page_dir, f"key {sk}")
                if "<a " in linked:
                    helps.append(Markup(linked))
                else:
                    report["key help links whose page is not built yet (left out)"].append(f"{sk}: {tgt_}")

        guides = []
        for r in key_imgs:
            img_ = str(r.get("image") or "")
            if str(r.get("step_key")) != sk or not img_ or sk.startswith("entry/") \
                    or re.search(r"(blank|shim|spacer|listgroup|(^|/)VF[^/]*\.gif$)", img_, re.I):
                continue
            alt_ = str(r.get("alt_text_new") or "").strip()
            if not alt_ and re.search(r"guide", img_, re.I):
                alt_ = "Diagram of the specimen with the area to examine for this step highlighted"
                report["guide pictures using the standard description (write alt_text_new where a step needs its own)"].append(f"{sk}: {posixpath.basename(img_)}")
            elif not alt_:
                alt_ = "Guide picture for this step"
                report["pictures on key steps with no alt text (images of text)"].append(f"{sk}: {posixpath.basename(img_)}")
            guides.append({"src": rel_(copy_image(img_, "key")), "alt": alt_, "caption": ""})

        back = rel_(step_path(parents[sk][0])) if len(parents.get(sk, [])) == 1 else ""
        html_k = env.get_template("key_step.html").render(
            **common(page_dir, "key"), title=heading,
            description=f"{heading}. Choose the picture that best matches your specimen.",
            heading=heading, prompt=prompt, guides=guides, has_plain=any_plain,
            heading_hidden=(sk.startswith("entry/")),
            choices=choices, videos=vl, helps=helps, back=back, start=rel_(registry["key"]),
            preview_mode=preview_mode, show_start=(sk != "entry/begin"),
            group_grid=(sk.startswith("entry/") and sk != "entry/begin"
                        and picture_only and len(choices) >= 6),
            taxonomy=[] if entry else taxp)
        dest_k = out / out_rel
        dest_k.parent.mkdir(parents=True, exist_ok=True)
        dest_k.write_text(html_k, encoding="utf-8")
        n_key_pages += 1


    # ================= supplementary (extra) species pages =================
    extra_rows = sheet_rows(wb, "extra_pages")
    imap_rows = sheet_rows(wb, "image_maps")
    imaps = defaultdict(list)
    for r in imap_rows:
        imaps[(str(r["page"]), str(r["map_name"]))].append(r)
    figs_by_page = defaultdict(list)
    for r in figures:
        if r.get("on_extra_page"):
            figs_by_page[str(r["on_extra_page"])].append(r)
    imgs_by_page = defaultdict(list)
    for r in images:
        if r.get("on_extra_page"):
            imgs_by_page[str(r["on_extra_page"])].append(r)
    secs_by_page = defaultdict(list)
    for r in sections:
        if r.get("from_page"):
            secs_by_page[str(r["from_page"])].append(r)
    vids_by_page = defaultdict(list)
    for r in other_vl:
        vids_by_page[str(r.get("shown_on") or "")].append(r)

    extra_path = {}
    for r in extra_rows:
        old_p = str(r["old_path"])
        sl = str(r.get("new_slug") or "").strip() or slug(PurePosixPath(old_p).stem)
        parent = str(r.get("parent_species_id") or "")
        extra_path[old_p] = (f"species/{parent}/{sl}/index.html" if parent and parent in built_ids
                             else f"pages/{sl}/index.html")

    extra_path_ref.update(extra_path)
    n_extra_built = 0
    for r in extra_rows:
        old_p = str(r["old_path"])
        out_rel = extra_path[old_p]
        page_dir = posixpath.dirname(out_rel)
        rel_x = lambda p_: posixpath.relpath(p_, page_dir) if p_ else ""
        owner = slug(PurePosixPath(old_p).stem)
        parent_id = str(r.get("parent_species_id") or "")
        parent = None
        if parent_id in built_ids:
            parent = {"name": species_by_id[parent_id]["display_name"],
                      "url": rel_x(f"species/{parent_id}/index.html")}

        imgs_x, any_plain = [], False
        for im in imgs_by_page.get(old_p, []):
            alt = str(im.get("alt_text_new") or "").strip() or str(im.get("alt_text_draft") or "").strip()
            if not alt:
                alt = str(r.get("title") or owner)
                report["pictures on supplementary pages with no description (write alt_text_new)"].append(f"{old_p}: {posixpath.basename(str(im['file']))}")
            src = rel_x(copy_image(str(im["file"]), owner))
            large = rel_x(copy_image(str(im.get("links_to_larger") or ""), owner))
            plain = ""
            hv = str(im.get("hover_image") or "").strip()
            if hv:
                w, h = to_int(im.get("width")), None
                hw = re.search(r"_(\d+)x\d+\.", hv)
                if hw and w and int(hw.group(1)) > w * 1.5:
                    large = large or rel_x(copy_image(hv, owner))   # hovering just magnified it
                else:
                    plain = rel_x(copy_image(hv, owner))
                    any_plain = True
            regions = []
            for reg in imaps.get((old_p, str(im.get("usemap") or "")), []):
                tgt = str(reg["goes_to"])
                url = (rel_x(f"species/{tgt}/index.html") if tgt in built_ids
                       else rel_x(extra_path[tgt]) if tgt in extra_path else None)
                lab = str(reg.get("label_new") or reg["label"] or f"Part {reg['seq']}")
                regions.append({"label": lab, "url": url})
                if not url:
                    report["clickable regions whose destination was not found"].append(f"{old_p}: {lab}")
            imgs_x.append({"src": src, "large": large, "alt": alt, "plain": plain, "regions": regions})

        secs_x = OrderedDict()
        for sr in secs_by_page.get(old_p, []):
            title = str(sr.get("section") or "")
            if title.lower() in SKIP_SECTIONS:
                continue
            fa = {f["image"]: f for f in figs_by_page.get(old_p, [])}

            def res_x(target, _own=owner):
                kind, _, path = target.partition(":")
                if kind == "fig" and (old / path).exists():
                    f_ = fa.get(path, {})
                    alt_ = str(f_.get("alt_text_new") or f_.get("alt_text_draft") or "").strip() or "Figure"
                    return rel_x(copy_image(path, _own)), f' data-figure-alt="{html.escape(alt_, quote=True)}"'
                if kind == "page" and path in extra_path:
                    return rel_x(extra_path[path])
                if kind == "page" and path in page_of_old and page_of_old[path] in built_ids:
                    return rel_x(f"species/{page_of_old[path]}/index.html")
                return None
            secs_x.setdefault(title, OrderedDict()).setdefault(
                str(sr.get("subsection") or ""), []).append(render_text(sr["text"], res_x))
        sec_list_x = [{"title": t, "blocks": [{"title": bt, "entries": e} for bt, e in b.items()]}
                      for t, b in secs_x.items()]

        vids_x, seen_x = [], set()
        for vr in vids_by_page.get(old_p, []):
            vid = str(vr.get("video_id") or "")
            v = videos.get(vid)
            v = videos.get(v.get("duplicate_of")) if v and v.get("duplicate_of") else v
            if not v or not v.get("archive_url") or vid in seen_x:
                continue
            seen_x.add(vid)
            desc = v.get("description_new") or v.get("description_on_old_site") or v.get("title_on_old_site") or ""
            vids_x.append({"url": v["archive_url"], "description": render_text(desc, lambda t_: None)})

        if not (imgs_x or sec_list_x or vids_x):
            report["supplementary pages with nothing to show (skipped)"].append(old_p)
            continue
        heading = str(r.get("title") or "").strip()
        if not heading or heading.lower() in ("untitled document", "untitled"):
            first_sec = next((x["title"] for x in sec_list_x if x["title"]), "")
            heading = first_sec or owner.replace("-", " ").replace("_", " ").title()
            report["supplementary pages whose title came from the page itself (set 'title' in extra_pages if wrong)"].append(f"{old_p} -> {heading}")
        html_x = env.get_template("extra_page.html").render(
            **common(page_dir, "species"), title=heading, description=heading,
            heading=heading, parent=parent, sections=sec_list_x, images=imgs_x,
            videos=vids_x, has_plain=any_plain)
        dest_x = out / out_rel
        dest_x.parent.mkdir(parents=True, exist_ok=True)
        dest_x.write_text(html_x, encoding="utf-8")
        n_extra_built += 1

    # ================= anatomy diagrams =================
    n_ana_pages, ana_thumb = 0, {}
    for g in ana_groups:
        out_rel = ana_path(g)
        page_dir = posixpath.dirname(out_rel)
        rel_a = lambda p_: posixpath.relpath(p_, page_dir)
        views, base = [], None
        for r in sorted(by_ana[g], key=lambda r_: int(r_["seq"])):
            label = str(r.get("label_new") or r["label"])
            alt = str(r.get("alt_text_new") or "").strip()
            if not alt:
                alt = f"{ANA_TITLES.get(g, g.title())} diagram: {label.lower()}"
                report["anatomy views using a standard description (write alt_text_new)"].append(f"{g}: {label}")
            v = {"label": label, "alt": alt, "src": rel_a(copy_image(str(r["image"]), "anatomy")),
                 "is_base": truthy(r.get("is_base"))}
            views.append(v)
            if v["is_base"] and base is None:
                base = v
        if not views:
            continue
        base = base or views[-1]
        # 1) your own thumbnail in content/assets/anatomy/ (any picture whose name contains
        #    the group name), 2) the old site's *choice* picture, 3) the first diagram view
        mine = []
        for folder in (content_dir / "assets" / "anatomy", content_dir / "assets"):
            if folder.is_dir():
                mine = sorted(f for f in folder.glob("*")
                              if f.is_file() and f.suffix.lower() in (".jpg", ".jpeg", ".png", ".gif")
                              and g.lower() in f.stem.lower())
                if mine:
                    break
        gdir = old / "html" / "anatomy" / g
        thumbs = sorted(gdir.glob("*hoice*.jpg")) if gdir.is_dir() else []
        if mine:
            dest_t = out / "img" / "anatomy" / f"thumb-{g}{mine[0].suffix.lower()}"
            dest_t.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(mine[0], dest_t)
            ana_thumb[g] = f"img/anatomy/thumb-{g}{mine[0].suffix.lower()}"
            report["anatomy thumbnails taken from content/assets/anatomy/"].append(f"{g}: {mine[0].name}")
        else:
            ana_thumb[g] = (thumbs[0].relative_to(old).as_posix() if thumbs
                            else str(sorted(by_ana[g], key=lambda r_: int(r_["seq"]))[0]["image"]))
        intro = str(by_ana[g][0].get("intro") or "")
        intro = re.sub(r"(?i)click on the options to the left", "Select a part", intro) or \
            "Select a part to show where it is on the diagram."
        tabs = [{"label": ANA_TITLES.get(x, x.title()), "current": x == g, "href": rel_a(ana_path(x))}
                for x in ana_groups]
        html_a = env.get_template("anatomy.html").render(
            **common(page_dir, "anatomy"), title=f"{ANA_TITLES.get(g, g.title())} anatomy",
            description=f"Labeled anatomy diagram of {ANA_TITLES.get(g, g.title())}.",
            heading=f"{ANA_TITLES.get(g, g.title())} anatomy", intro=intro, views=views, base=base, tabs=tabs)
        dest_a = out / out_rel
        dest_a.parent.mkdir(parents=True, exist_ok=True)
        dest_a.write_text(html_a, encoding="utf-8")
        n_ana_pages += 1

    # ================= Groups pages =================
    n_group_pages = 0

    def group_url(r, from_dir):
        """-> (relative url or None, reason it is missing: '' / 'not_built' / 'missing')"""
        t_, tt_ = str(r.get("goes_to") or ""), str(r.get("goes_to_type") or "")
        url_, why = None, "missing"
        if tt_ == "key_step" and t_ in steps:
            url_ = step_path(t_)
        elif tt_ in ("organism_page", "organism_page_unlisted"):
            sid_ = t_ if t_ in species_by_id else page_of_old.get(t_)
            if sid_ in built_ids:
                url_ = f"species/{sid_}/index.html"
            elif sid_ in species_by_id:
                why = "not_built"
        elif tt_ == "id_page":
            sl_ = old_to_slug.get(t_.lower())
            url_ = registry.get(sl_) if sl_ else None
        return (posixpath.relpath(url_, from_dir), "") if url_ else (None, why)

    for g in group_pages:
        out_rel = groups_path(g)
        page_dir = posixpath.dirname(out_rel)
        roots_, stack_ = [], [(-1, None)]
        top = []
        stack_ = [(-1, top)]
        for r in sorted(by_group[g], key=lambda r_: int(r_["seq"])):
            node = {"r": r, "kids": []}
            lvl = int(r["level"])
            while stack_[-1][0] >= lvl:
                stack_.pop()
            stack_[-1][1].append(node)
            stack_.append((lvl, node["kids"]))

        def render(nodes):
            if not nodes:
                return ""
            items = []
            for n in nodes:
                r = n["r"]
                name = str(r.get("name_new") or r["name"])
                label = f"<em>{html.escape(name)}</em>" if r["rank"] == "Species" else html.escape(name)
                u, why = group_url(r, page_dir)
                inner = f'<a href="{html.escape(u, quote=True)}">{label}</a>' if u else label
                if not u:
                    if why == "not_built":
                        report["Groups entries pointing at species pages not built in this run (a full build links them)"].append(f"{g}: {name}")
                    else:
                        report["Groups entries whose destination does not exist (check goes_to, or convert the page)"].append(f"{g}: {name}")
                items.append(f'<li class="rank-{str(r["rank"]).lower()}">{inner}{render(n["kids"])}</li>')
            return "<ul>" + "".join(items) + "</ul>"

        tabs = [{"label": GROUP_TITLES.get(t, t.title()), "current": t == g,
                 "href": posixpath.relpath(groups_path(t), page_dir)} for t in group_pages]
        note = ("The copepods do not conform to the format of this key as well as the other groups. Because species of the "
                "same genus can differ biologically, selecting a genus name might not take you to a point where every "
                "following step is in that genus.") if g == "copepoda" else ""
        html_g = env.get_template("groups.html").render(
            **common(page_dir, "groups"), title=f"Groups: {GROUP_TITLES.get(g, g.title())}",
            description="The groups in the key, each linked to the point in the key where that group is reached.",
            tabs=tabs, note=note, outline=Markup(render(top)))
        dest_g = out / out_rel
        dest_g.parent.mkdir(parents=True, exist_ok=True)
        dest_g.write_text(html_g, encoding="utf-8")
        n_group_pages += 1

    if by_ana:
        ana_dir = "anatomy"
        rel_i = lambda p_: posixpath.relpath(p_, ana_dir)
        secs_i, placed = [], set()
        for title, members in ANA_SECTIONS:
            cards = []
            for g in members:
                if g in by_ana:
                    placed.add(g)
                    cards.append({"label": f"{ANA_TITLES.get(g, g.title())} anatomy", "url": rel_i(ana_path(g)),
                                  "src": rel_i(ana_thumb[g] if ana_thumb[g].startswith("img/")
                                               else copy_image(ana_thumb[g], "anatomy"))})
            if cards:
                secs_i.append({"title": title, "cards": cards})
        for g in by_ana:
            if g not in placed:
                secs_i.append({"title": ANA_TITLES.get(g, g.title()),
                               "cards": [{"label": f"{ANA_TITLES.get(g, g.title())} anatomy",
                                          "url": rel_i(ana_path(g)),
                                          "src": rel_i(ana_thumb[g] if ana_thumb[g].startswith("img/")
                                                       else copy_image(ana_thumb[g], "anatomy"))}]})
        if "anatomy" in content:
            report["the Anatomy index page is now generated from the workbook (content/anatomy.md is no longer used)"].append("anatomy.md")
        (out / ana_dir).mkdir(parents=True, exist_ok=True)
        (out / ana_dir / "index.html").write_text(env.get_template("anatomy_index.html").render(
            **common(ana_dir, "anatomy"), title="Anatomy",
            description="Labeled anatomy diagrams for the main groups in the key.",
            intro="Select one of the pictures below to explore that group's anatomy.",
            sections=secs_i), encoding="utf-8")

    if a.noindex:
        (out / "robots.txt").write_text("User-agent: *\nDisallow: /\n", encoding="utf-8")
    (out / ".nojekyll").write_text("", encoding="utf-8")      # GitHub Pages: serve the files as they are

    n_pages, problems = check_site(out)
    if tax_pref:
        report["taxonomy spellings replaced from the taxonomy sheet"].append(f"{len(tax_pref)} rule(s) in use")
    lines = [f"Built {len(species)} species page(s) of {len(all_species)}, {n_key_pages} key step page(s), {n_group_pages} Groups page(s), {n_ana_pages} anatomy page(s), {n_extra_built} supplementary page(s) and {len(content)} content page(s); "
             f"{n_pages} HTML pages in {out}",
             f"Supplementary pages in the workbook: {n_extra}", ""]
    for k, v in report.items():
        lines.append(f"{k}: {len(v)}")
        for x in v[:5]:
            lines.append(f"    {x}")
        if len(v) > 5:
            lines.append(f"    ... and {len(v) - 5} more")
    lines += ["", f"Page checks: {len(problems)} problem(s)"] + [f"    {p}" for p in problems[:30]]
    if len(problems) > 30:
        lines.append(f"    ... and {len(problems) - 30} more")
    text = "\n".join(lines)
    (out / "build_report.txt").write_text(text, encoding="utf-8")
    full = [f"{k}: {len(v)}" + "".join(f"\n    {x}" for x in v) for k, v in report.items()]
    (out / "build_report_full.txt").write_text("\n\n".join(full) + "\n", encoding="utf-8")
    print(text)


if __name__ == "__main__":
    main()
