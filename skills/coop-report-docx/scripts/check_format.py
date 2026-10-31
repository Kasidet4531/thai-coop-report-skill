#!/usr/bin/env python3
"""Check a co-op report .docx against the สหกิจศึกษา formatting rules.

Usage: python check_format.py report.docx [--allow-chapter-page-numbers]

Prints one PASS/FAIL line per rule, then lists any "[...]" placeholders left
for the student to fill in. Exit code 1 if any rule fails.
"""
import re
import sys

from docx import Document
from docx.oxml.ns import qn

FONT = "TH SarabunPSK"
CM = 360000  # EMU per cm
# same rule as build_report.py: "[1]", "[2, 5]" are citations, not gaps to fill
PLACEHOLDER_RE = r"\[(?![\d\s,\-–]+\])[^\]]+\]"


def near(a, b, tol=0.05):
    return a is not None and abs(a / CM - b) <= tol


def main(path, allow_chapter_numbers=False):
    try:  # Windows consoles default to a Thai/ANSI code page
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    doc = Document(path)
    results = []

    def check(name, ok, detail=""):
        results.append((ok, name, detail))

    # 1. Page setup on every section
    bad = []
    for i, s in enumerate(doc.sections):
        if not (near(s.page_width, 21.0, 0.1) and near(s.page_height, 29.7, 0.1)):
            bad.append(f"section {i + 1}: not A4")
        if not near(s.left_margin, 3.0):
            bad.append(f"section {i + 1}: left {s.left_margin / CM:.2f} cm")
        for side in ("right_margin", "top_margin", "bottom_margin"):
            if not near(getattr(s, side), 2.54):
                bad.append(f"section {i + 1}: {side} {getattr(s, side) / CM:.2f} cm")
    check("A4, left margin 3 cm, top/bottom/right 2.54 cm", not bad, "; ".join(bad))

    # 2. Fonts: every style and run that names a font uses TH SarabunPSK (incl. complex script)
    other = set()
    for el in doc.element.iter(qn("w:rFonts")):
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            v = el.get(qn(attr))
            if v and v != FONT:
                other.add(v)
    for el in doc.styles.element.iter(qn("w:rFonts")):
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            v = el.get(qn(attr))
            if v and v != FONT and el.getparent().getparent().tag == qn("w:style"):
                style_name = el.getparent().getparent().find(qn("w:name"))
                sname = style_name.get(qn("w:val")) if style_name is not None else "?"
                if sname in ("Normal", "heading 1", "heading 2", "heading 3"):
                    other.add(f"{v} (style {sname})")
    check(f"Single font throughout: {FONT}", not other, ", ".join(sorted(other)))

    # 3. Sizes: Normal 16, Heading 1 18 bold centred
    def size_of(style):
        rPr = style.element.find(qn("w:rPr"))
        szcs = rPr.find(qn("w:szCs")) if rPr is not None else None
        return (style.font.size.pt if style.font.size else None,
                int(szcs.get(qn("w:val"))) / 2 if szcs is not None else None)
    n = size_of(doc.styles["Normal"])
    check("Body text 16 pt (Latin + Thai)", n == (16, 16), str(n))
    h1 = doc.styles["Heading 1"]
    h1s = size_of(h1)
    check("Chapter titles 18 pt bold, centred",
          h1s == (18, 18) and h1.font.bold and h1.paragraph_format.alignment == 1, str(h1s))

    # 4. Page numbering
    fmts = []
    for s in doc.sections:
        pg = s._sectPr.find(qn("w:pgNumType"))
        fmts.append(pg.get(qn("w:fmt")) if pg is not None else None)
    check("Front matter numbered with Roman numerals",
          "upperRoman" in fmts or "lowerRoman" in fmts, str(fmts))
    check("Body numbered with Arabic numerals", "decimal" in fmts, str(fmts))

    first_dec = fmts.index("decimal") if "decimal" in fmts else None
    pg = doc.sections[first_dec]._sectPr.find(qn("w:pgNumType")) if first_dec is not None else None
    check("Arabic numbering restarts at 1 for Chapter 1",
          pg is not None and pg.get(qn("w:start")) == "1")

    chapter_sections = []
    for i, s in enumerate(doc.sections):
        first_text = ""
        for p in s.iter_inner_content() if hasattr(s, "iter_inner_content") else []:
            if hasattr(p, "text") and p.text.strip():
                first_text = p.text.strip()
                break
        if first_text.startswith("บทที่"):
            chapter_sections.append((i, s.different_first_page_header_footer))
    if allow_chapter_numbers:  # built with options.chapter_first_page_number
        check("Each chapter starts on a new section", bool(chapter_sections), str(chapter_sections))
    else:
        check("No page number on each chapter's first page",
          bool(chapter_sections) and all(t for _, t in chapter_sections),
          str(chapter_sections))

    footers_ok = True
    for s in doc.sections[1:]:
        xml = s.footer._element.xml
        if "PAGE" not in xml or 'w:val="center"' not in xml:
            footers_ok = False
    check("Page numbers bottom-centre", footers_ok)

    # 5. Required parts present
    texts = [p.text for p in doc.paragraphs]
    joined = "\n".join(texts)
    required = ["รายงานสหกิจศึกษาฉบับสมบูรณ์", "บทคัดย่อ", "ABSTRACT", "กิตติกรรมประกาศ", "สารบัญ",
                "บทที่ 1", "เอกสารอ้างอิง", "ภาคผนวก", "ประวัติผู้เขียน"]
    missing = [r for r in required if r not in joined]
    check("All required parts present", not missing, ", ".join(missing))

    toc = "TOC \\o" in doc.element.xml or "TOC \\\\o" in doc.element.xml
    check("Automatic table of contents field", toc)

    # Keywords line 14 pt
    kw_ok = False
    for p in doc.paragraphs:
        if p.text.startswith("คำสำคัญ") and p.runs and p.runs[0].font.size and p.runs[0].font.size.pt == 14:
            kw_ok = True
    check("Thai keywords line 14 pt", kw_ok)

    failed = 0
    for ok, name, detail in results:
        print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  -> {detail}" if detail and not ok else ""))
        failed += not ok

    ph = []
    for p in doc.paragraphs:
        ph += re.findall(PLACEHOLDER_RE, p.text.replace(chr(0x200B), ""))
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                ph += re.findall(PLACEHOLDER_RE, c.text.replace(chr(0x200B), ""))
    print(f"\n{len(results) - failed}/{len(results)} rules passed")
    if ph:
        print(f"{len(ph)} placeholder(s) still to fill in:")
        for x in dict.fromkeys(ph):
            print(f"  {x}")
    return 1 if failed else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) != 1:
        print(__doc__)
        sys.exit(2)
    sys.exit(main(args[0], allow_chapter_numbers="--allow-chapter-page-numbers" in sys.argv))
