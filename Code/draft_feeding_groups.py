#!/usr/bin/env python3
"""
draft_feeding_groups.py

Builds a DRAFT table of feeding traits for every taxon in the key, using the key's own
text first and taxonomy-based defaults only where the text says nothing. Writes a new
sheet, 'feeding_groups', into the workbook. Nothing else in the workbook is touched, and
a dated backup is made first.

    conda activate video
    python draft_feeding_groups.py --workbook "$ZK/data/cfbkey_workbook.xlsx"
    python draft_feeding_groups.py --workbook ... --force     # redo the draft columns

WHAT IT DOES
  1. Reads the species, species_sections, key_steps and videos sheets.
  2. Looks for feeding evidence: words like filter, raptorial, predatory, detritus, and
     for rotifers the trophi type (malleate, incudate, virgate, forcipate, ramate,
     uncinate), which maps closely onto feeding mode.
  3. Pulls body length out of the text where it is given (mm or micrometres).
  4. Falls back to a genus, then family, then order default where the text is silent.
  5. Writes one row per taxon with the DRAFT values, the evidence they came from, and
     empty yellow columns for the reviewed values.

WHAT IT DOES NOT DO
  Decide anything. Every draft needs checking by someone who knows the animals: the
  genus and family defaults in particular are coarse, and congeneric species often
  differ. Treat 'draft_basis' and 'confidence' as the guide to what needs most attention.

Columns written (grey = computed, yellow = yours to fill in):
  species_id, display_name, Order, Family, Genus
  feeding_mode_draft, trophic_group_draft, body_length_mm_draft, draft_basis, confidence,
  evidence
  feeding_mode, trophic_group, body_length_min_mm, body_length_max_mm,
  prey_size_min_um, prey_size_max_um, life_stage, source, notes
"""

import argparse
import datetime
import re
import shutil
import sys
from collections import Counter
from pathlib import Path

try:
    from openpyxl import load_workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.exit("Needs openpyxl:  conda install -c conda-forge openpyxl")

# ----------------------------------------------------------------------------
# What counts as evidence. Each entry: (pattern, feeding_mode, trophic_group).
# Earlier entries win. Edit freely: this is a starting point, not a rule book.
# ----------------------------------------------------------------------------
EVIDENCE = [
    (r"\bparasit", "parasitic", "parasite"),
    (r"\b(raptorial|seizes|grasps|pierc\w*|sucks? out|suctorial)\b", "raptorial", "carnivore"),
    (r"\b(predator|predacious|predatory|preys? (?:up)?on|carnivor)", "raptorial", "carnivore"),
    (r"\b(scrap\w+|brows\w+|periphyton|epiphyt\w+|surfaces? of (?:plants|sediment))", "scraper", "detritivore"),
    (r"\b(detritus|detritiv\w+|sediment\w* feed)", "scraper", "detritivore"),
    (r"\b(filter[- ]?feed\w*|filtering|filters? (?:particles|algae|bacteria))", "filter", "herbivore"),
    (r"\b(suspension[- ]?feed\w*|current[- ]?feed\w*)", "suspension", "omnivore"),
    (r"\b(graz\w+ on (?:algae|phytoplankton)|herbivor)", "filter", "herbivore"),
    (r"\b(omnivor)", "suspension", "omnivore"),
    (r"\b(bacteriv\w+|feeds? on bacteria)", "filter", "herbivore"),
]

# Rotifer trophi type -> feeding mode. The rotifer key already sorts on these.
TROPHI = {
    "incudate": ("raptorial", "carnivore"),
    "forcipate": ("raptorial", "carnivore"),
    "virgate": ("raptorial", "omnivore"),      # piercing and sucking cell contents
    "malleate": ("filter", "herbivore"),
    "malleoramate": ("filter", "herbivore"),
    "ramate": ("filter", "herbivore"),
    "uncinate": ("raptorial", "carnivore"),
    "cardate": ("raptorial", "omnivore"),
}

# ----------------------------------------------------------------------------
# Defaults by taxon, used ONLY where the text says nothing. Genus beats family
# beats order. These are deliberately coarse and all need review.
# ----------------------------------------------------------------------------
GENUS_DEFAULT = {
    # cladoceran predators
    "Leptodora": ("raptorial", "carnivore"), "Bythotrephes": ("raptorial", "carnivore"),
    "Cercopagis": ("raptorial", "carnivore"), "Polyphemus": ("raptorial", "carnivore"),
    # predatory and omnivorous copepods
    "Mesocyclops": ("raptorial", "carnivore"), "Acanthocyclops": ("raptorial", "carnivore"),
    "Epischura": ("raptorial", "carnivore"), "Limnocalanus": ("raptorial", "carnivore"),
    "Senecella": ("raptorial", "carnivore"), "Cyclops": ("raptorial", "omnivore"),
    "Diacyclops": ("raptorial", "omnivore"), "Thermocyclops": ("raptorial", "omnivore"),
    "Microcyclops": ("raptorial", "omnivore"), "Macrocyclops": ("raptorial", "omnivore"),
    "Eucyclops": ("scraper", "omnivore"), "Tropocyclops": ("suspension", "omnivore"),
    "Ergasilus": ("parasitic", "parasite"), "Argulus": ("parasitic", "parasite"),
    # rotifer predators and piercers
    "Asplanchna": ("raptorial", "carnivore"), "Ploesoma": ("raptorial", "carnivore"),
    "Collotheca": ("raptorial", "carnivore"), "Synchaeta": ("raptorial", "omnivore"),
    "Trichocerca": ("raptorial", "omnivore"), "Ascomorpha": ("raptorial", "omnivore"),
    "Gastropus": ("raptorial", "omnivore"), "Polyarthra": ("raptorial", "omnivore"),
    # other arthropods
    "Chaoborus": ("raptorial", "carnivore"), "Mysis": ("raptorial", "omnivore"),
    "Eubranchipus": ("filter", "herbivore"),
}
FAMILY_DEFAULT = {
    "Daphniidae": ("filter", "herbivore"), "Sididae": ("filter", "herbivore"),
    "Bosminidae": ("filter", "herbivore"), "Holopediidae": ("filter", "herbivore"),
    "Moinidae": ("filter", "herbivore"), "Chydoridae": ("scraper", "detritivore"),
    "Macrothricidae": ("scraper", "detritivore"), "Ilyocryptidae": ("scraper", "detritivore"),
    "Leptodoridae": ("raptorial", "carnivore"), "Cercopagididae": ("raptorial", "carnivore"),
    "Polyphemidae": ("raptorial", "carnivore"), "Diaptomidae": ("suspension", "omnivore"),
    "Temoridae": ("suspension", "omnivore"), "Centropagidae": ("suspension", "omnivore"),
    "Cyclopidae": ("raptorial", "omnivore"), "Brachionidae": ("filter", "herbivore"),
    "Synchaetidae": ("raptorial", "omnivore"), "Asplanchnidae": ("raptorial", "carnivore"),
    "Conochilidae": ("filter", "herbivore"), "Filiniidae": ("filter", "herbivore"),
    "Hexarthridae": ("filter", "herbivore"), "Trichocercidae": ("raptorial", "omnivore"),
    "Lecanidae": ("scraper", "detritivore"), "Euchlanidae": ("filter", "herbivore"),
    "Gastropodidae": ("raptorial", "omnivore"), "Testudinellidae": ("filter", "herbivore"),
    "Chaoboridae": ("raptorial", "carnivore"),
}
ORDER_DEFAULT = {
    "Calanoida": ("suspension", "omnivore"), "Cyclopoida": ("raptorial", "omnivore"),
    "Harpacticoida": ("scraper", "detritivore"), "Ploima": ("filter", "herbivore"),
    "Flosculariaceae": ("filter", "herbivore"), "Collothecaceae": ("raptorial", "carnivore"),
    "Ostracoda": ("scraper", "detritivore"), "Hydracarina": ("raptorial", "carnivore"),
    "Diptera": ("raptorial", "carnivore"), "Anostraca": ("filter", "herbivore"),
    "Cladocera": ("filter", "herbivore"),
}

# Sentences about being EATEN, or about defending against being eaten, which must not
# be read as evidence of being a predator. The hard case is phrasing like "lacks
# structures for defense against invertebrate predators", which names predators while
# describing prey.
PREYED_UPON = re.compile(
    # passive: "is preyed on", "are eaten by"
    r"\b(is|are|was|were|being|been)\s+(?:\w+\s+){0,3}?"
    r"(preyed|eaten|consumed|ingested|taken|selected|captured|grazed)\b"
    r"|\bprey (?:of|for|item|items|species)\b|\bfood (?:of|for|source for)\b"
    r"|\beaten by\b|\bpreyed (?:up)?on by\b|\bconsumed by\b"
    # defence and vulnerability: the taxon is the one being hunted
    r"|\b(defen[cs]\w*|defend\w*|protect\w*|deterren\w*|avoid\w*|escap\w*|evad\w*"
    r"|vulnerab\w*|susceptib\w*|refug\w*|camouflag\w*|crypsis|helmet|neck ?tooth"
    r"|tail ?spine|lacks?|resist\w*)\b"
    # "against / from / by ... predators", and predation as a pressure on it
    r"|\b(against|from|by|to|of)\s+(?:\w+\s+){0,3}?predat"
    r"|\bpredation (?:pressure|risk|rate|rates|mortality|intensity|by)\b"
    r"|\bpredator[- ]induced\b|\bpredator avoidance\b"
    # "anti-predatory", and predator phrases that describe the HUNTER OF this taxon
    r"|\banti[- ]?predat\w*"
    r"|\b(tactile|visual|invertebrate|vertebrate|planktivor\w*|gape[- ]limited|"
    r"size[- ]selective|fish|larval|ambush)\s+predat\w*"
    r"|\bsinks? past\b|\bpredator[- ]free\b"
    r"|\bvulnerable to\b|\bpredators? of (?:it|them|this)\b", re.I)

FEEDING_WORDS = re.compile(
    r"(feed|diet|food|graz|filter|raptorial|predat|prey|carnivor|herbivor|omnivor|detrit|"
    r"bacteri|algae|phytoplankton|scrap|brows|pierc|parasit|ingest|trophi|malleate|incudate|"
    r"virgate|forcipate|ramate|uncinate|cardate)", re.I)

# Body length. Tried in order: a range in mm, a single value in mm, then a range in
# micrometres ONLY when the sentence is about the animal rather than its food.
NOT_BODY = re.compile(r"particle|food|prey|algae|cell|mesh|egg|spine|seta", re.I)
SIZE_RE = [
    (re.compile(r"(\d+(?:\.\d+)?)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?)\s*mm\b", re.I), 1.0, False),
    (re.compile(r"(?:approximately|about|ca\.?|roughly|up to)?\s*(\d+(?:\.\d+)?)\s*mm\b", re.I), 1.0, True),
    (re.compile(r"(\d+(?:\.\d+)?)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?)\s*(?:µm|um|microns?)\b", re.I), 0.001, False),
]


def make_backup(wbp: Path) -> Path:
    folder = wbp.parent / "workbook_backups"
    folder.mkdir(parents=True, exist_ok=True)
    stamp = f"{datetime.datetime.now():%Y%m%d-%H%M%S}"
    dest = folder / f"{wbp.stem}_backup_{stamp}{wbp.suffix}"
    n = 2
    while dest.exists():
        dest = folder / f"{wbp.stem}_backup_{stamp}-{n}{wbp.suffix}"
        n += 1
    shutil.copy2(wbp, dest)
    return dest


def col(row: dict, *names):
    """A column's value, matching the name however it is capitalised."""
    lower = {str(k).strip().lower(): v for k, v in row.items() if k}
    for n in names:
        v = lower.get(n.lower())
        if v not in (None, ""):
            return str(v).strip()
    return ""


def rows_of(wb, name):
    if name not in wb.sheetnames:
        return []
    ws = wb[name]
    heads = [c.value for c in ws[1]]
    return [dict(zip(heads, r)) for r in ws.iter_rows(min_row=2, values_only=True)]


def plain(text) -> str:
    """Strip the workbook's light markup so the lexicon sees ordinary words."""
    s = str(text or "")
    s = re.sub(r"\[([^\]]+)\]\([^)]*\)", r"\1", s)
    s = re.sub(r"\{(?:red|blue|green|orange)\|([^}]*)\}", r"\1", s)
    return re.sub(r"[*_`]", "", s).replace(" | ", " ")


def eaten_in(sentence: str, display_name: str) -> bool:
    """True where THIS taxon is the object: "Chaoborus and fish prey on D. catawba".
    Active voice with another subject, which the passive guard cannot see."""
    parts = [p for p in re.split(r"\s+", str(display_name or "").strip()) if p]
    if not parts:
        return False
    toks = [re.escape(parts[0])]                                  # the genus
    if len(parts) >= 2:
        toks.append(re.escape(parts[-1]))                         # the epithet
        toks.append(re.escape(parts[0][0]) + r"\.?\s*" + re.escape(parts[-1]))  # "D. catawba"
    pat = (r"\b(prey|preys|preyed|preying|feed|feeds|feeding|eat|eats|eaten|consume|consumes|"
           r"ingest|ingests|graze|grazes|select|selects)\b[^.]{0,12}?\b(?:up)?on\s+"
           r"(?:\w+[\s,]+){0,5}?(" + "|".join(toks) + r")\b")
    return bool(re.search(pat, sentence, re.I))


def sentences(text):
    return [s.strip() for s in re.split(r"(?<=[.;:])\s+", text) if s.strip()]


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--workbook", required=True)
    ap.add_argument("--force", action="store_true", help="rewrite the draft columns of an existing sheet")
    a = ap.parse_args()
    wbp = Path(a.workbook).expanduser()
    if not wbp.exists():
        sys.exit(f"ERROR: workbook not found: {wbp}")
    backup = make_backup(wbp)
    print(f"Backup saved: {backup.parent.name}/{backup.name}")

    wb = load_workbook(wbp)
    species = rows_of(wb, "species")
    if not species:
        sys.exit("ERROR: no 'species' sheet found.")

    # --- gather every scrap of text about each taxon ---
    SKIP = ("reference", "literature", "quicktime", "movie", "additional picture", "link",
            "barcode", "accession", "sequence")
    text_by_id, skipped_rows = {}, 0
    for r in rows_of(wb, "species_sections"):
        sid = col(r, "species_id")
        section = col(r, "section")
        if not sid:
            continue
        if any(w in section.lower() for w in SKIP):
            skipped_rows += 1          # citations and video titles are not feeding evidence
            continue
        text_by_id.setdefault(sid, []).append((section, plain(r.get("text"))))
    for r in rows_of(wb, "species_videos"):
        sid = str(r.get("species_id") or "")
        if sid:
            text_by_id.setdefault(sid, []).append(("video", plain(r.get("label_on_old_site"))))

    # the key's own wording, which often describes the mouthparts and trophi
    key_text = []
    for r in rows_of(wb, "key_steps"):
        key_text.append((str(r.get("goes_to") or ""), plain(r.get("choice_text"))))

    heads = ["species_id", "display_name", "Order", "Family", "Genus",
             "feeding_mode_draft", "trophic_group_draft", "body_length_mm_draft",
             "food_size_um_draft", "draft_basis", "confidence", "evidence",
             "feeding_mode", "trophic_group", "body_length_min_mm", "body_length_max_mm",
             "prey_size_min_um", "prey_size_max_um", "life_stage", "source", "notes"]
    DRAFTS = ("feeding_mode_draft", "trophic_group_draft", "body_length_mm_draft",
              "food_size_um_draft", "draft_basis", "confidence", "evidence")
    INPUTS = ("feeding_mode", "trophic_group", "body_length_min_mm", "body_length_max_mm",
              "prey_size_min_um", "prey_size_max_um", "life_stage", "source", "notes")

    out_rows, basis_count, mode_count = [], Counter(), Counter()
    prey_guarded = 0
    for sp in species:
        sid = col(sp, "species_id")
        if not sid:
            continue
        name = col(sp, "display_name") or sid
        genus = col(sp, "genus")
        family = col(sp, "family")
        order = col(sp, "order", "suborder", "subclass")

        # the species' own pages first; the key's wording about it second
        own = " ".join(t for _, t in text_by_id.get(sid, []))
        from_key = " ".join(t for g, t in key_text if g == sid)
        all_text = (own + " " + from_key).strip()

        mode = group = ""
        basis, conf, evidence = "", "", ""

        # 1. the trophi type, for rotifers: the species' own text wins over the key's
        for source_text, label in ((own, "trophi"), (from_key, "trophi, from the key")):
            for word, (m_, g_) in TROPHI.items():
                if re.search(rf"\b{word}\b", source_text, re.I):
                    mode, group, basis, conf = m_, g_, f"{label}: {word}", "medium"
                    evidence = next((s for s in sentences(source_text)
                                     if re.search(word, s, re.I)), "")[:300]
                    break
            if mode:
                break

        # 2. a plain statement about feeding in the key's own text
        if not mode:
            for pat, m_, g_ in EVIDENCE:
                hit = next((s for s in sentences(all_text)
                            if re.search(pat, s, re.I)
                            and not PREYED_UPON.search(s)
                            and not eaten_in(s, name)), None)
                if hit:
                    mode, group, basis, conf = m_, g_, "the key's own text", "high"
                    evidence = hit[:300]
                    break

        # 3. taxonomy defaults
        if not mode:
            for label, table, keyv in (("genus", GENUS_DEFAULT, genus),
                                       ("family", FAMILY_DEFAULT, family),
                                       ("order", ORDER_DEFAULT, order)):
                if keyv in table:
                    mode, group = table[keyv]
                    basis, conf = f"{label} default ({keyv})", "low" if label != "genus" else "medium"
                    break
        if not mode:
            basis, conf = "nothing found", "none"

        # any feeding-ish sentence, so a reviewer has something to read
        if not evidence:
            evidence = next((s for s in sentences(all_text) if FEEDING_WORDS.search(s)), "")[:300]

        prey_size = ""
        m_p = re.search(r"particles?\s+of\s+(\d+(?:\.\d+)?)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?)\s*"
                        r"(?:µm|um|microns?)", all_text, re.I)
        if not m_p:
            m_p = re.search(r"(\d+(?:\.\d+)?)\s*(?:-|–|to)\s*(\d+(?:\.\d+)?)\s*(?:µm|um|microns?)"
                            r"[^.]{0,40}(?:particles?|food|prey|algae|cells)", all_text, re.I)
        if m_p:
            prey_size = f"{float(m_p.group(1)):g}-{float(m_p.group(2)):g}"

        size = ""
        for rx, factor, single in SIZE_RE:
            for m in rx.finditer(all_text):
                before = all_text[max(0, m.start() - 30):m.start()]
                around = all_text[max(0, m.start() - 40):m.end() + 25]
                says_body = re.search(r"\b(body|length|long|size|measur\w*)\b", before, re.I)
                if NOT_BODY.search(around) and not says_body:
                    continue                      # this number describes the food, not the animal
                if single:
                    size = f"{float(m.group(1)) * factor:g}"
                else:
                    size = (f"{float(m.group(1)) * factor:g}-"
                            f"{float(m.group(2)) * factor:g}")
                break
            if size:
                break

        if any(PREYED_UPON.search(x) for x in sentences(all_text)):
            prey_guarded += 1
        basis_count[basis.split(" (")[0]] += 1
        mode_count[mode or "(none)"] += 1
        out_rows.append({"species_id": sid, "display_name": name, "Order": order,
                         "Family": family, "Genus": genus,
                         "feeding_mode_draft": mode, "trophic_group_draft": group,
                         "body_length_mm_draft": size, "food_size_um_draft": prey_size,
                         "draft_basis": basis,
                         "confidence": conf, "evidence": evidence})

    # --- write the sheet ---
    if "feeding_groups" in wb.sheetnames and not a.force:
        print("feeding_groups: already present. Add --force to redo the draft columns "
              "(your own columns are kept).")
    else:
        if "feeding_groups" in wb.sheetnames:          # keep what a person has filled in
            ws = wb["feeding_groups"]
            old_heads = [c.value for c in ws[1]]
            keep = {}
            for row in ws.iter_rows(min_row=2, values_only=True):
                d = dict(zip(old_heads, row))
                keep[str(d.get("species_id"))] = {k: d.get(k) for k in INPUTS}
            del wb["feeding_groups"]
        else:
            keep = {}
        ws = wb.create_sheet("feeding_groups")
        ws.append(heads)
        for r in out_rows:
            saved = keep.get(r["species_id"], {})
            ws.append([r.get(h) if h in r else saved.get(h) for h in heads])
        widths = {"species_id": 26, "display_name": 28, "Order": 16, "Family": 18, "Genus": 18,
                  "feeding_mode_draft": 16, "trophic_group_draft": 16, "body_length_mm_draft": 16,
                  "food_size_um_draft": 16,
                  "draft_basis": 24, "confidence": 11, "evidence": 70, "feeding_mode": 15,
                  "trophic_group": 15, "body_length_min_mm": 15, "body_length_max_mm": 15,
                  "prey_size_min_um": 15, "prey_size_max_um": 15, "life_stage": 14,
                  "source": 30, "notes": 30}
        for i, h in enumerate(heads, 1):
            c = ws.cell(row=1, column=i)
            is_in = h in INPUTS
            c.font = Font(name="Arial", bold=True, size=10, color="000000" if is_in else "FFFFFF")
            c.fill = PatternFill("solid", fgColor="FFD966" if is_in else "1F3864")
            ws.column_dimensions[get_column_letter(i)].width = widths.get(h, 14)
        for row in ws.iter_rows(min_row=2):
            for c in row:
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(vertical="top", wrap_text=heads[c.column - 1] == "evidence")
                if heads[c.column - 1] in INPUTS:
                    c.fill = PatternFill("solid", fgColor="FFF2CC")
                elif heads[c.column - 1] in DRAFTS:
                    c.fill = PatternFill("solid", fgColor="EDEDED")
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{ws.max_row}"
        wb.save(wbp)
        print(f"feeding_groups: {ws.max_row - 1} taxa written")

    if skipped_rows:
        print(f"(ignored {skipped_rows} rows in reference, video and sequence sections)")
    if prey_guarded:
        print(f"({prey_guarded} taxa have sentences about being eaten or defending themselves; "
              f"those are not read as evidence of being a predator)")
    print("\nWhere each draft came from:")
    for k, n in basis_count.most_common():
        print(f"  {n:4}  {k}")
    print("\nDraft feeding modes:")
    for k, n in mode_count.most_common():
        print(f"  {n:4}  {k}")
    weak = basis_count.get("nothing found", 0) + basis_count.get("family default", 0) \
        + basis_count.get("order default", 0)
    print(f"\n{basis_count.get(chr(116) + 'he key' + chr(39) + 's own text', 0)} taxa have real evidence in the key; "
          f"{weak} rest on a coarse default or nothing at all.\nFilter the sheet on "
          f"confidence = low or none to see what needs attention first.")


if __name__ == "__main__":
    main()
