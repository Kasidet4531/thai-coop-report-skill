---
name: coop-report-docx
description: Build a Thai cooperative-education final report (รายงานสหกิจศึกษาฉบับสมบูรณ์ / เล่มสหกิจ / เล่มรายงานฝึกงาน) as a correctly formatted Word .docx — covers, Thai and English abstracts, acknowledgement, automatic table of contents, chapters 1–5, references, appendices, author biography — with TH SarabunPSK, 3 cm left margin, Roman-numbered front matter and Arabic-numbered body. Use this skill whenever the user wants to make, format, reformat, fix, or start a สหกิจศึกษา / co-op / internship report or its template in Word, wants their draft text or existing .docx turned into the department format, or asks for "เล่มสหกิจ", "ทำเล่ม", "จัดรูปแบบรายงาน", "เทมเพลตรายงานสหกิจ", even if they don't say "docx" or name the format.
---

# Co-op report (เล่มสหกิจศึกษา) builder

Students lose hours fighting Word over margins, Thai fonts, and page numbering that
switches from Roman to Arabic and disappears on chapter openers. This skill does the
formatting deterministically with a script, so your effort goes into the content.

The department format lives in `references/format-rules.md`. The script already
implements every rule in it — you only need to read it if the user asks why something
looks the way it does or you have to edit a document by hand.

## Workflow

### 1. Work out which job this is

- **Empty template** — the user wants a blank book to fill in. Ask for any cover details
  they're willing to give (names, title, department, company, year); everything else
  becomes a red `[placeholder]`. Even with zero details, build it.
- **Format existing content** — the user gives a draft (.docx, .txt, .md, pasted text,
  notes). Your job is to move *their* words into the structure, not rewrite them.
- **Help write content** — the user wants help drafting sections from facts about their
  internship. Write from what they told you; where you lack facts (results, numbers,
  names, dates), leave a `[placeholder]` instead of inventing them — a fabricated result
  in a submitted report is far worse for the student than a visible gap.

These mix freely (e.g. format their chapter 1–2 drafts and leave 3–5 as template).

### 2. Read the user's material (if any)

For a .docx draft, extract text with headings so you can map it:

```bash
python -c "import docx,sys; d=docx.Document(sys.argv[1]); [print(f'[{p.style.name}] {p.text}') for p in d.paragraphs if p.text.strip()]" draft.docx
```

Tables in a draft: `for t in d.tables: ...` — carry them over as `table` blocks.
Images: extract from `word/media/` in the .docx zip into a working folder and reference
them as `figure` blocks.

Map the draft onto the book: cover info → `meta`; abstract/keywords; chapters →
`chapters[].sections[]`. Keep the student's wording, chapter titles and section count —
the handout explicitly allows adapting titles and number of sections to the project.
Only strip numbering they typed by hand ("1.1", "บทที่ 2"), because the script numbers
headings itself and you'd otherwise get "1.1 1.1 …".

### 3. Write content.json

Follow `references/content-schema.md` (read it before writing the file — it lists every
key and the block types for paragraphs, lists, tables and figures). Put the JSON in the
user's working folder next to the output so they can edit and rebuild later. Write it
as UTF-8.

### 4. Build

```bash
python <skill-dir>/scripts/build_report.py content.json "<output>.docx"
# or, for a blank template with no content file:
python <skill-dir>/scripts/build_report.py --template "<output>.docx"
```

Requires `python-docx` and `pythainlp` (`pip install python-docx pythainlp` if either
import fails). Body text uses Thai Distributed alignment, the usual look for Thai
reports. Thai has no spaces between words, and many Word installs can't find Thai word
boundaries, so they wrap only at real spaces and stretch the letters across the line.
The script prevents this by marking every Thai word boundary with an invisible
zero-width space (via pythainlp). If pythainlp is missing, the script warns and falls
back to left alignment. Install pythainlp rather than accepting the fallback.

### 5. Verify

```bash
python <skill-dir>/scripts/check_format.py "<output>.docx"
```

It checks margins, font, sizes, Roman/Arabic numbering, no number on chapter first
pages, and required parts, then lists the placeholders still to fill. Fix any FAIL
before handing over.

When you extract text from the built .docx yourself, remember the body contains
zero-width spaces (U+200B): strip them before comparing text, and print with UTF-8
output on Windows, because the console code page can't encode them.

On Windows with Microsoft Word installed, also render it — this refreshes the table of
contents and gives you a PDF to look at:

```bash
powershell -NoProfile -ExecutionPolicy Bypass -File <skill-dir>/scripts/finalize_with_word.ps1 -Docx "<output>.docx" -Pdf "<output>.pdf"
```

Look at the cover, an abstract page and a chapter opener in the PDF. If Word reports
the file opened read-only, the TOC in the .docx itself isn't refreshed — that's fine,
the file asks Word to update fields when the student opens it.

### 6. Hand over

Tell the user, briefly and in their language (usually Thai):
- where the .docx (and content.json) are;
- how many placeholders remain and what kind (the checker's list);
- that when Word asks "update fields?" on open they should click **Yes** — that fills
  the สารบัญ; later they can right-click the TOC → Update Field → Update entire table;
- the logo spot on the cover is a placeholder unless they gave a logo file;
- if TH SarabunPSK isn't installed on their machine, Word substitutes another font on
  screen — install it before printing (free from SIPA / f0nt.com);
- to change content later they can edit content.json and rerun step 4, or edit the
  .docx directly in Word (the styles Heading 1/2/3 keep the TOC working).

## Editing the generated document directly

If the user asks for small tweaks after the build, prefer editing content.json and
rebuilding — the formatting stays guaranteed. Edit the .docx by hand (python-docx)
only for things the schema can't express, and re-run the checker afterwards.
