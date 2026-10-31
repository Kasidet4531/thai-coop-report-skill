#!/usr/bin/env python3
"""Build a Thai cooperative-education (สหกิจศึกษา) final report as .docx.

Usage:
    python build_report.py content.json output.docx
    python build_report.py --template output.docx          # empty template with placeholders

The content file is JSON (UTF-8). Every field is optional: anything missing is
filled with a visible "[...]" placeholder so the student knows what to complete.
See references/content-schema.md for the full schema.
"""
import argparse
import json
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

FONT = "TH SarabunPSK"
BODY_PT = 16
CHAPTER_PT = 18
KEYWORD_PT = 14
FIRST_LINE_INDENT = Cm(1.5)  # "เริ่มพิมพ์ตัวอักษรตัวที่ 9" ≈ 8 character widths
PLACEHOLDER_COLOR = RGBColor(0xC0, 0x00, 0x00)

THAI_APPENDIX_LETTERS = "กขคงจฉชซฌญฎฏฐฑฒณดตถทธนบปผฝพฟภมยรลวศษสหฬอฮ"

DEFAULT_CHAPTERS = [
    {"title": "บทนำ", "sections": [
        {"heading": "ความเป็นมาและความสำคัญ"},
        {"heading": "วัตถุประสงค์ของการวิจัย"},
        {"heading": "ขอบเขตของการวิจัย"},
        {"heading": "วิธีการดำเนินการวิจัย"},
        {"heading": "ประโยชน์ที่คาดว่าจะได้รับ"},
    ]},
    {"title": "แนวคิด ทฤษฎีและงานวิจัยที่เกี่ยวข้อง", "sections": [
        {"heading": "ทฤษฎีที่เกี่ยวข้อง"},
        {"heading": "งานวิจัยที่เกี่ยวข้อง"},
    ]},
    {"title": "วิธีดำเนินการวิจัย", "sections": [{"heading": None}, {"heading": None}]},
    {"title": "ผลการวิจัย", "sections": [{"heading": None}, {"heading": None}]},
    {"title": "สรุปผลการวิจัยและข้อเสนอแนะ", "sections": [{"heading": None}, {"heading": None}]},
]


# --------------------------------------------------------------------------- low-level XML helpers

def set_run_font(run, size=BODY_PT, bold=False, placeholder=False):
    """Thai text is 'complex script' in Word, so cs font/size/bold must be set too,
    otherwise Thai renders in the theme font at the wrong size."""
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    rPr = run._element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)
    _set_cs_props(rPr, size, bold)
    if placeholder:
        run.font.color.rgb = PLACEHOLDER_COLOR


def _set_cs_props(rPr, size, bold):
    for tag in ("w:bCs", "w:szCs"):
        el = rPr.find(qn(tag))
        if el is not None:
            rPr.remove(el)
    if bold:
        rPr.append(OxmlElement("w:bCs"))
    szCs = OxmlElement("w:szCs")
    szCs.set(qn("w:val"), str(int(size * 2)))
    rPr.append(szCs)


def style_font(style, size, bold=False):
    style.font.name = FONT
    style.font.size = Pt(size)
    style.font.bold = bold
    style.font.italic = False
    style.font.color.rgb = RGBColor(0, 0, 0)
    rPr = style.element.get_or_add_rPr()
    rFonts = rPr.find(qn("w:rFonts"))
    if rFonts is None:
        rFonts = OxmlElement("w:rFonts")
        rPr.insert(0, rFonts)
    for attr in list(rFonts.attrib):
        del rFonts.attrib[attr]
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        rFonts.set(qn(attr), FONT)
    _set_cs_props(rPr, size, bold)
    lang = rPr.find(qn("w:lang"))
    if lang is None:
        lang = OxmlElement("w:lang")
        rPr.append(lang)
    lang.set(qn("w:bidi"), "th-TH")


ZWSP = chr(0x200B)  # zero-width space
_THAI = re.compile("[" + chr(0x0E00) + "-" + chr(0x0E7F) + "]")  # Thai Unicode block
try:
    from pythainlp.tokenize import word_tokenize as _thai_tokenize
except ImportError:  # optional: without it Thai Distributed can't be used safely
    _thai_tokenize = None


def thai_word_breaks(text):
    """Insert invisible zero-width spaces between Thai words.

    Thai is written without spaces, and Word on many machines doesn't know where
    Thai words end, so it can only wrap lines at real spaces. Under Thai Distributed
    the short lines that result get their letters spread across the page. Marking
    every word boundary gives Word normal wrap points, so lines fill up and only
    need a little stretching."""
    if not _thai_tokenize or not _THAI.search(text):
        return text
    out = []
    for tok in _thai_tokenize(text, engine="newmm", keep_whitespace=True):
        if out and not tok.isspace() and not out[-1].isspace() and                 (_THAI.search(tok) or _THAI.search(out[-1])):
            out.append(ZWSP)
        out.append(tok)
    return "".join(out)


def justify(paragraph, mode):
    """Body-text alignment: "thaiDistribute" (default), "both" or "left".
    Justified Thai relies on the zero-width word breaks from thai_word_breaks();
    without them Word wraps only at spaces and stretched lines look broken."""
    if mode == "left":
        return
    pPr = paragraph._p.get_or_add_pPr()
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), mode)


def add_field(paragraph, instr, placeholder_text="", size=BODY_PT):
    def fld(t):
        r = paragraph.add_run()
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), t)
        r._r.append(el)
        return r
    fld("begin")
    r = paragraph.add_run()
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    r._r.append(it)
    fld("separate")
    r = paragraph.add_run(placeholder_text)
    set_run_font(r, size)
    fld("end")


def set_page_number_format(section, fmt, start=None):
    sectPr = section._sectPr
    pg = sectPr.find(qn("w:pgNumType"))
    if pg is None:
        pg = OxmlElement("w:pgNumType")
        sectPr.append(pg)
    pg.set(qn("w:fmt"), fmt)
    if start is not None:
        pg.set(qn("w:start"), str(start))
    elif qn("w:start") in pg.attrib:
        del pg.attrib[qn("w:start")]


def clear_footer(footer):
    footer.is_linked_to_previous = False
    for p in footer.paragraphs:
        for r in list(p.runs):
            r._r.getparent().remove(r._r)


def page_number_footer(footer):
    clear_footer(footer)
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    add_field(p, "PAGE", "1")


def setup_section(section):
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(3.0)
    section.right_margin = Cm(2.54)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.header_distance = Cm(1.25)
    section.footer_distance = Cm(1.25)


# --------------------------------------------------------------------------- builder

class ReportBuilder:
    def __init__(self, data):
        self.data = data
        self.meta = data.get("meta", {})
        self.doc = Document()
        self.table_count = {}
        self.figure_count = {}
        self.justify_mode = data.get("options", {}).get("justify", "thaiDistribute")
        if self.justify_mode == "thaiDistribute" and _thai_tokenize is None:
            print("warning: pythainlp is not installed, so Thai words can't be marked for "
                  "line wrapping; body text falls back to left alignment "
                  "(pip install pythainlp to get Thai Distributed)", file=sys.stderr)
            self.justify_mode = "left"
        # word-break marks are needed only when lines get stretched to the margin
        self.wrap_thai = self.justify_mode != "left"
        self._setup_styles()
        setup_section(self.doc.sections[0])
        settings = self.doc.settings.element
        upd = OxmlElement("w:updateFields")
        upd.set(qn("w:val"), "true")
        settings.append(upd)

    # ---- styles
    def _setup_styles(self):
        st = self.doc.styles
        normal = st["Normal"]
        style_font(normal, BODY_PT)
        pf = normal.paragraph_format
        pf.space_before = Pt(0)
        pf.space_after = Pt(0)
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE

        h1 = st["Heading 1"]
        style_font(h1, CHAPTER_PT, bold=True)
        h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        h1.paragraph_format.space_before = Pt(0)
        h1.paragraph_format.space_after = Pt(BODY_PT)
        h1.paragraph_format.keep_with_next = True

        h2 = st["Heading 2"]
        style_font(h2, BODY_PT, bold=True)
        h2.paragraph_format.space_before = Pt(BODY_PT * 0.5)
        h2.paragraph_format.space_after = Pt(0)
        h2.paragraph_format.keep_with_next = True

        h3 = st["Heading 3"]
        style_font(h3, BODY_PT, bold=True)
        h3.paragraph_format.left_indent = Cm(0.75)
        h3.paragraph_format.space_before = Pt(0)
        h3.paragraph_format.space_after = Pt(0)
        h3.paragraph_format.keep_with_next = True

        fh = st.add_style("Front Heading", WD_STYLE_TYPE.PARAGRAPH)
        fh.base_style = normal
        style_font(fh, CHAPTER_PT, bold=True)
        fh.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fh.paragraph_format.space_after = Pt(BODY_PT)

        for name in ("Caption",):
            style_font(st[name], BODY_PT)
        for name in ("Table Caption", "Figure Caption"):
            s = st.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
            s.base_style = normal
            style_font(s, BODY_PT)
        st["Table Caption"].paragraph_format.keep_with_next = True
        st["Figure Caption"].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for i in range(1, 4):
            name = f"TOC {i}" if f"TOC {i}" in [s.name for s in st] else None
            if name is None:
                s = st.add_style(f"TOC {i}", WD_STYLE_TYPE.PARAGRAPH)
                s.base_style = normal
            style_font(st[f"TOC {i}"], BODY_PT)
            st[f"TOC {i}"].paragraph_format.space_after = Pt(0)
            st[f"TOC {i}"].paragraph_format.space_before = Pt(0)
        for name in ("List Bullet", "List Number"):
            style_font(st[name], BODY_PT)

    # ---- paragraph helpers
    def para(self, text="", size=BODY_PT, bold=False, align=None, indent_first=False,
             style=None, placeholder=False):
        p = self.doc.add_paragraph(style=style)
        if text:
            self._add_runs(p, text, size, bold, placeholder, wrap=(align == "justify"))
        if align == "center":
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        elif align == "justify":
            justify(p, self.justify_mode)
        if indent_first:
            p.paragraph_format.first_line_indent = FIRST_LINE_INDENT
        return p

    def _add_runs(self, p, text, size, bold, placeholder=False, wrap=False):
        """Supports **bold** spans inside plain strings; [..] spans are placeholders."""
        for part in re.split(r"(\*\*[^*]+\*\*|\[[^\]]*\])", text):
            if not part:
                continue
            if wrap and self.wrap_thai and not part.startswith("["):
                # placeholders stay unmarked: the student types over them anyway
                part = thai_word_breaks(part)
            if part.startswith("**") and part.endswith("**"):
                r = p.add_run(part[2:-2])
                set_run_font(r, size, True, placeholder)
            elif part.startswith("[") and part.endswith("]"):
                r = p.add_run(part)
                set_run_font(r, size, bold, True)
            else:
                r = p.add_run(part)
                set_run_font(r, size, bold, placeholder)

    def blank(self, size=BODY_PT, n=1):
        for _ in range(n):
            p = self.doc.add_paragraph()
            r = p.add_run("")
            set_run_font(r, size)
            p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
            # empty paragraph height follows the paragraph mark size
            rPr = p._p.get_or_add_pPr()
            mark = OxmlElement("w:rPr")
            sz = OxmlElement("w:sz"); sz.set(qn("w:val"), str(size * 2))
            szcs = OxmlElement("w:szCs"); szcs.set(qn("w:val"), str(size * 2))
            mark.append(sz); mark.append(szcs)
            rPr.append(mark)

    def page_break(self):
        p = self.doc.add_paragraph()
        p.add_run().add_break(WD_BREAK.PAGE)

    def new_section(self):
        sec = self.doc.add_section(WD_SECTION.NEW_PAGE)
        setup_section(sec)
        return sec

    def m(self, key, placeholder):
        v = self.meta.get(key)
        return v if v else f"[{placeholder}]"

    # ---- front parts
    def cover(self, label):
        meta = self.meta
        # Logo placeholder — the student pastes the institute logo here.
        logo = meta.get("logo_path")
        if logo and os.path.exists(logo):
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(logo, height=Cm(3.5))
        else:
            self.blank(20, 2)
            self.para("[วางโลโก้สถาบันที่นี่ สูงประมาณ 3.5 ซม.]", size=16, align="center")
            self.blank(20, 2)
        self.blank(20)
        self.para("รายงานสหกิจศึกษาฉบับสมบูรณ์", size=22, bold=True, align="center")
        self.blank(20)
        self.para(self.m("title_th", "ชื่อโครงงานภาษาไทย"), size=20, bold=True, align="center")
        self.blank(20)
        self.para(self.m("title_en", "Project title in English"), size=20, bold=True, align="center")
        self.blank(20, 3)
        students = meta.get("students") or [meta.get("student") or "[ชื่อ-สกุล นักศึกษา]"]
        for s in students:
            self.para(s, size=18, bold=True, align="center")
        self.blank(20, 4 - min(len(students), 3))
        self.para(f"ภาควิชา{self.m('department', '...')}", size=18, bold=True, align="center")
        self.para(f"คณะ{self.meta.get('faculty', 'วิศวกรรมศาสตร์')}", size=18, bold=True, align="center")
        self.para(self.meta.get("institute", "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง"),
                  size=18, bold=True, align="center")
        self.blank(18)
        self.para(f"ปีการศึกษา {self.m('academic_year', '25..')}", size=18, bold=True, align="center")

    def abstract_th(self):
        meta = self.meta
        rows = [
            ("ชื่อโครงงานสหกิจศึกษา", self.m("title_th", "ชื่อโครงงานภาษาไทย")),
            ("ชื่อ-สกุล นักศึกษา", ", ".join(meta.get("students") or []) or self.m("student", "ชื่อ-สกุล")),
        ]
        for label, val in rows:
            self._label_line(label, val)
        p = self.doc.add_paragraph()
        self._label_runs(p, "คณะ", meta.get("faculty", "วิศวกรรมศาสตร์"))
        p.add_run("  ")
        self._label_runs(p, "ภาควิชา", self.m("department", "ภาควิชา"))
        self._label_line("ชื่อ-สกุล อาจารย์นิเทศ", self.m("advisor", "ชื่ออาจารย์นิเทศ"))
        self._label_line("ชื่อ-สกุล ผู้นิเทศงาน", self.m("mentor", "ชื่อผู้นิเทศงาน"))
        self._label_line("สถานประกอบการ", self.m("company", "ชื่อสถานประกอบการ"))
        self.blank()
        self.doc.add_paragraph("บทคัดย่อ", style="Heading 1").paragraph_format.space_after = Pt(0)
        self.blank()
        self._abstract_body(self.data.get("abstract_th"), "เนื้อหาบทคัดย่อภาษาไทย ไม่เกิน 1 หน้า")
        self.blank()
        self._keywords("คำสำคัญ", self.data.get("keywords_th"), "คำสำคัญ 3-5 คำ")

    def abstract_en(self):
        meta = self.meta
        self._label_line("Cooperative Title", self.m("title_en", "Project title in English"), sep=": ")
        self._label_line("Student intern name",
                         ", ".join(meta.get("students_en") or []) or self.m("student_en", "Student name"), sep=": ")
        p = self.doc.add_paragraph()
        self._label_runs(p, "Faculty", meta.get("faculty_en", "Engineering"), sep=": ")
        p.add_run("  ")
        self._label_runs(p, "Department", self.m("department_en", "Department"), sep=": ")
        self._label_line("Advisor name", self.m("advisor_en", "Advisor name"), sep=": ")
        self._label_line("Mentor name", self.m("mentor_en", "Mentor name"), sep=": ")
        self._label_line("Company", self.m("company_en", "Company name"), sep=": ")
        self.blank()
        self.doc.add_paragraph("ABSTRACT", style="Heading 1").paragraph_format.space_after = Pt(0)
        self.blank()
        self._abstract_body(self.data.get("abstract_en"), "English abstract, no longer than one page")
        self.blank()
        self._keywords("Keywords", self.data.get("keywords_en"), "3-5 keywords")

    def _label_runs(self, p, label, value, sep="  "):
        r = p.add_run(f"{label}{sep}")
        set_run_font(r, BODY_PT, bold=True)
        self._add_runs(p, value, BODY_PT, False)

    def _label_line(self, label, value, sep="  "):
        p = self.doc.add_paragraph()
        self._label_runs(p, label, value, sep)

    def _abstract_body(self, text, ph):
        paras = text if isinstance(text, list) else ([text] if text else [f"[{ph}]"])
        for t in paras:
            self.para(t, align="justify", indent_first=True)

    def _keywords(self, label, kws, ph):
        p = self.doc.add_paragraph()
        r = p.add_run(f"{label} : ")
        set_run_font(r, KEYWORD_PT, bold=True)
        val = ", ".join(kws) if isinstance(kws, list) else (kws or f"[{ph}]")
        self._add_runs(p, val, KEYWORD_PT, False)

    def acknowledgement(self):
        self.doc.add_paragraph("กิตติกรรมประกาศ", style="Heading 1")
        ack = self.data.get("acknowledgement")
        paras = ack if isinstance(ack, list) else ([ack] if ack else
                                                   ["[ข้อความขอบคุณสถานประกอบการ ผู้นิเทศงาน อาจารย์นิเทศ และผู้ที่ให้ความช่วยเหลือ]"])
        for t in paras:
            self.para(t, align="justify", indent_first=True)
        self.blank()
        signer = ", ".join(self.meta.get("students") or []) or self.meta.get("student") or "[ชื่อ-สกุล นักศึกษา]"
        p = self.para(signer, align=None)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    def toc_page(self, title, instr, empty_hint, in_toc=True):
        self.doc.add_paragraph(title, style="Heading 1" if in_toc else "Front Heading")
        p = self.doc.add_paragraph()
        r = p.add_run("หน้า")
        set_run_font(r, BODY_PT, bold=True)
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p = self.doc.add_paragraph()
        add_field(p, instr, empty_hint)

    # ---- body
    def chapter(self, idx, ch):
        p = self.doc.add_paragraph(style="Heading 1")
        r = p.add_run(f"บทที่ {idx}")
        r.add_break()
        p.add_run(ch.get("title") or "[ชื่อบท]")
        for b in ch.get("intro", []) or []:
            self.block(b, idx)
        sections = ch.get("sections") or []
        for si, sec in enumerate(sections, 1):
            num = f"{idx}.{si}"
            heading = sec.get("heading") or "[หัวข้อ]"
            self.doc.add_paragraph(f"{num}  {heading}", style="Heading 2")
            self._blocks_or_placeholder(sec, idx)
            for ssi, sub in enumerate(sec.get("subsections") or [], 1):
                self.doc.add_paragraph(f"{num}.{ssi}  {sub.get('heading') or '[หัวข้อย่อย]'}", style="Heading 3")
                self._blocks_or_placeholder(sub, idx)

    def _blocks_or_placeholder(self, sec, chap):
        blocks = sec.get("content")
        if not blocks:
            if not sec.get("subsections"):
                self.para("[เนื้อหาหัวข้อนี้]", align="justify", indent_first=True)
            return
        for b in blocks:
            self.block(b, chap)

    def block(self, b, chap):
        if isinstance(b, str):
            self.para(b, align="justify", indent_first=True)
            return
        t = b.get("type")
        if t in ("bullets", "numbered"):
            style = "List Bullet" if t == "bullets" else "List Number"
            num_id = self._restart_numbering() if t == "numbered" else None
            for item in b.get("items", []):
                p = self.doc.add_paragraph(style=style)
                if num_id is not None:
                    numPr = p._p.get_or_add_pPr().get_or_add_numPr()
                    numPr.get_or_add_ilvl().val = 0
                    numPr.get_or_add_numId().val = num_id
                self._add_runs(p, item, BODY_PT, False, wrap=True)
        elif t == "table":
            self.table(b, chap)
        elif t == "figure":
            self.figure(b, chap)
        elif t == "paragraph":
            self.para(b.get("text", ""), align="justify", indent_first=b.get("indent", True))
        else:
            self.para(json.dumps(b, ensure_ascii=False))

    def _restart_numbering(self):
        """Every "List Number" paragraph shares the style's numbering instance, so a
        second list would carry on counting from the first. Give each list its own
        instance of the same list definition, restarting at 1."""
        numbering = self.doc.part.numbering_part.element
        style_numId = self.doc.styles["List Number"].element.pPr.numPr.numId.val
        abstract_id = numbering.num_having_numId(style_numId).abstractNumId.val
        num = numbering.add_num(abstract_id)
        num.add_lvlOverride(ilvl=0).add_startOverride(1)
        return num.numId

    def table(self, b, chap):
        n = self.table_count[chap] = self.table_count.get(chap, 0) + 1
        cap = self.doc.add_paragraph(style="Table Caption")
        r = cap.add_run(f"ตารางที่ {chap}.{n}  ")
        set_run_font(r, BODY_PT, bold=True)
        self._add_runs(cap, b.get("caption", "[ชื่อตาราง]"), BODY_PT, False)
        rows = b.get("rows") or [["[หัวตาราง]", "[หัวตาราง]"], ["", ""]]
        ncols = max(len(r) for r in rows)
        tbl = self.doc.add_table(rows=len(rows), cols=ncols)
        tbl.style = "Table Grid"
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for i, row in enumerate(rows):
            for j in range(ncols):
                cell = tbl.cell(i, j)
                cell.paragraphs[0].text = ""
                self._add_runs(cell.paragraphs[0], str(row[j]) if j < len(row) else "", BODY_PT, i == 0, wrap=True)
                if i == 0:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        self.blank()

    def figure(self, b, chap):
        n = self.figure_count[chap] = self.figure_count.get(chap, 0) + 1
        path = b.get("path")
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if path and os.path.exists(path):
            p.add_run().add_picture(path, width=Cm(b.get("width_cm", 12)))
        else:
            self._add_runs(p, f"[ใส่รูปภาพ: {path or 'ไฟล์รูป'}]", BODY_PT, False)
        p.paragraph_format.keep_with_next = True
        cap = self.doc.add_paragraph(style="Figure Caption")
        r = cap.add_run(f"รูปที่ {chap}.{n}  ")
        set_run_font(r, BODY_PT, bold=True)
        self._add_runs(cap, b.get("caption", "[คำอธิบายรูป]"), BODY_PT, False)
        self.blank()

    # ---- back matter
    def references(self):
        self.doc.add_paragraph("เอกสารอ้างอิง", style="Heading 1")
        refs = self.data.get("references") or ["[รายการเอกสารอ้างอิงตามรูปแบบที่ภาควิชากำหนด]"]
        for ref in refs:
            p = self.para(ref)
            p.paragraph_format.left_indent = Cm(1.27)
            p.paragraph_format.first_line_indent = Cm(-1.27)

    def appendices(self):
        apps = self.data.get("appendices") or [{"title": "[ชื่อภาคผนวก]"}]
        for i, app in enumerate(apps):
            if i:
                self.page_break()
            letter = THAI_APPENDIX_LETTERS[i] if i < len(THAI_APPENDIX_LETTERS) else str(i + 1)
            p = self.doc.add_paragraph(style="Heading 1")
            r = p.add_run(f"ภาคผนวก {letter}")
            r.add_break()
            p.add_run(app.get("title") or "[ชื่อภาคผนวก]")
            for b in app.get("content") or ["[เนื้อหาภาคผนวก]"]:
                self.block(b, letter)

    def biography(self):
        self.doc.add_paragraph("ประวัติผู้เขียน", style="Heading 1")
        bio = self.data.get("biography")
        if isinstance(bio, dict):
            for k, v in bio.items():
                self._label_line(k, str(v))
        else:
            paras = bio if isinstance(bio, list) else ([bio] if bio else [
                "[ชื่อ-สกุล]", "[วัน เดือน ปีเกิด]", "[ประวัติการศึกษา]", "[ที่อยู่ / ช่องทางติดต่อ]"])
            for t in paras:
                self.para(t)

    # ---- assembly
    def build(self):
        opts = self.data.get("options", {})
        doc = self.doc

        # Section 1: covers — no page numbers
        sec = doc.sections[0]
        clear_footer(sec.footer)
        if opts.get("outer_cover", True):
            self.cover("ปกนอก")
            if opts.get("flyleaf", True):
                self.page_break()
                self.para("[กระดาษรองปก — หน้าว่าง ใช้กระดาษ A4 สีขาว]", size=12, align="center")
            self.page_break()
        self.cover("ปกใน")

        # Section 2: front matter — Roman numerals I, II, III … bottom centre
        sec = self.new_section()
        set_page_number_format(sec, "upperRoman", start=1)
        page_number_footer(sec.footer)
        self.abstract_th()
        self.page_break()
        self.abstract_en()
        self.page_break()
        self.acknowledgement()
        self.page_break()
        self.toc_page("สารบัญ", r'TOC \o "1-2" \h \z \u',
                      "[สารบัญ — เปิดไฟล์ใน Word แล้วคลิกขวาที่นี่ > อัปเดตฟิลด์ (Update Field)]", in_toc=False)
        chapters = self.data.get("chapters") or DEFAULT_CHAPTERS
        has = lambda kind: kind in json.dumps(self.data.get("chapters", []) + self.data.get("appendices", []), ensure_ascii=False)
        if opts.get("list_of_tables", has('"table"')):
            self.page_break()
            self.toc_page("สารบัญตาราง", r'TOC \h \z \t "Table Caption,1"',
                          "[สารบัญตาราง — อัปเดตฟิลด์ใน Word]")
        if opts.get("list_of_figures", has('"figure"')):
            self.page_break()
            self.toc_page("สารบัญรูป", r'TOC \h \z \t "Figure Caption,1"',
                          "[สารบัญรูป — อัปเดตฟิลด์ใน Word]")

        # Body: one section per chapter so the chapter's first page has no number
        for i, ch in enumerate(chapters, 1):
            sec = self.new_section()
            set_page_number_format(sec, "decimal", start=1 if i == 1 else None)
            sec.different_first_page_header_footer = True
            page_number_footer(sec.footer)
            clear_footer(sec.first_page_footer)
            self.chapter(i, ch)

        # Back matter — continues Arabic numbering
        for part in (self.references, self.appendices, self.biography):
            sec = self.new_section()
            set_page_number_format(sec, "decimal")
            sec.different_first_page_header_footer = False
            page_number_footer(sec.footer)
            part()
        return doc


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("content", nargs="?", help="content JSON file (omit with --template)")
    ap.add_argument("output", help="output .docx path")
    ap.add_argument("--template", action="store_true", help="build an empty template with placeholders")
    args = ap.parse_args()

    if args.template or not args.content:
        data = {}
    else:
        with open(args.content, encoding="utf-8") as f:
            data = json.load(f)
    doc = ReportBuilder(data).build()
    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    doc.save(args.output)
    print(f"saved {args.output}")


if __name__ == "__main__":
    sys.exit(main())
