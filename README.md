# CFB Zooplankton Key — rebuild (Version 6.0, in progress)

A rebuild of [An Image-Based Key to the Zooplankton of North America](https://cfb.unh.edu/cfbkey/html/begin.html),
the taxonomic key published by the University of New Hampshire Center for Freshwater Biology.

**Preview of this copy: https://jvthaney.github.io/cfbkey-test/**

This is a **test copy for review**, not the live key. The live key is still the UNH
site linked above and is unchanged. Pages here carry a "do not index" tag so this
copy does not appear in search results beside the real one.

## What is different from the old site

- Every page is generated from one spreadsheet, so a fact is edited once, in one place,
  instead of in hand-written HTML on hundreds of pages.
- The QuickTime videos are converted to MP4 and play in the browser with no plugin.
  They are hosted at the Internet Archive ([cfbkey-videos](https://archive.org/details/cfbkey-videos)).
- The pages work on phones and tablets, which the old fixed-width pages could not.
- Accessibility work toward WCAG 2.2 AA: real headings, keyboard access to everything that
  used to need a mouse, text alternatives for pictures, and no Flash or plugins.
  **Still in progress:** most pictures carry draft text alternatives awaiting review.

## How to read this repository

| Folder | What it is |
| --- | --- |
| `docs/` | **The website.** Generated: never edit by hand, it is rewritten on every build. This is the folder GitHub Pages publishes. |
| `data/` | The workbook (`cfbkey_workbook.xlsx`) that holds the species text, taxonomy, picture lists, video descriptions and the whole identification key, plus `video_manifest.csv`, the web address of each video. |
| `content/` | The ordinary text pages (Home, Instructions, Purpose, History, References, Contact, Copyright, Barcoding, Video introduction) as plain text files, and the pictures used on them. |
| `Code/` | The scripts that read the workbook and write `docs/`. |
| `site_src/` | Optional: editable copies of the page templates and stylesheet. Present only if someone has run `build_site.py --init`. |
| `guide/` | The editing guide (`cfbkey-editing-guide.docx`). |

Not in this repository, because of size: the copy of the old site (about 700 MB, the
source of the pictures) and the converted video files (about 170 MB, now at the Internet
Archive). Both stay on the maintainer's computer and in the project's OneDrive folder.

## Making a change

Content changes are spreadsheet or text-file edits; no programming:

1. Edit `data/cfbkey_workbook.xlsx` (species text, taxonomy, alt text, video descriptions,
   the key) or a file in `content/` (ordinary page text).
2. Rebuild: `bash Code/rebuild.sh`
3. Read the short report it prints, then check the preview it opens.
4. Commit and push. GitHub Pages republishes within about a minute.

The editing guide in `guide/` explains each job in detail, including which sheet to
use for what and how to read the build report.

## Still to do

- Review the draft text alternatives for about 1,100 pictures and 315 detail figures.
- Write descriptions for the video clips that do not have one.
- Label the picture-only choices in the identification key.
- Settle the licence and credit wording with UNH.
- Decide on visitor analytics (the old Google Analytics tag stopped working in 2023).
- Decide where the key will live permanently, and keep the old web addresses working.

## Credit

The key is the work of many people at the UNH Center for Freshwater Biology; see the
Contributors list on the home page and the History of the Key page. Images remain the
property of the organisations and individuals credited under Purpose and General
Information. Please read the Copyright page before reusing anything here.
