#!/usr/bin/env python3
"""Smoke test for the coop-report-docx scripts. Runs in CI and locally:

    python tests/smoke_test.py

Builds a blank template and a sample report that uses every block type, runs
check_format.py on both, and checks for the bugs found while developing the
skill: numbered lists continuing each other, citations coloured as placeholders,
missing Thai word-break marks, and TOC styles Word would ignore.
"""
import json
import os
import re
import subprocess
import sys
import tempfile

from docx import Document
from docx.oxml.ns import qn

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SKILL = os.path.join(ROOT, "skills", "coop-report-docx")
SCRIPTS = os.path.join(SKILL, "scripts")
ZWSP = chr(0x200B)

SAMPLE = {
    "meta": {
        "title_th": "การพัฒนาระบบทดสอบอัตโนมัติ",
        "title_en": "Development of an Automated Test System",
        "students": ["นายทดสอบ ระบบดี"],
        "students_en": ["Mr. Test Systemdee"],
        "student_id": "65000000",
        "department": "วิศวกรรมคอมพิวเตอร์",
        "department_en": "Computer Engineering",
        "program": "หลักสูตรวิศวกรรมศาสตรบัณฑิต",
        "academic_year": "2568",
        "advisor": "ผศ.ดร.ตัวอย่าง ใจดี",
        "mentor": "คุณพี่เลี้ยง ตัวอย่าง",
        "company": "บริษัท ตัวอย่าง จำกัด",
    },
    "abstract_th": ["บทคัดย่อภาษาไทยสำหรับการทดสอบระบบสร้างเล่มรายงานสหกิจศึกษาโดยอัตโนมัติ"],
    "keywords_th": ["ทดสอบ", "ระบบ"],
    "abstract_en": ["English abstract used by the smoke test."],
    "keywords_en": ["test", "system"],
    "acknowledgement": ["ขอขอบคุณทุกท่านที่ช่วยเหลือในการทดสอบ"],
    "chapters": [
        {
            "title": "บทนำ",
            "intro": ["บทนี้อธิบายที่มาของโครงงานตามแนวคิดในเอกสาร [1] และ [2, 3]"],
            "sections": [
                {"heading": "วัตถุประสงค์", "content": [
                    {"type": "numbered", "items": ["ข้อแรก", "ข้อสอง", "ข้อสาม"]}]},
                {"heading": "ขั้นตอน", "content": [
                    {"type": "numbered", "items": ["ขั้นแรก", "ขั้นสอง"]},
                    {"type": "bullets", "items": ["จุดหนึ่ง", "จุดสอง"]},
                    {"type": "table", "caption": "ตารางตัวอย่าง",
                     "rows": [["หัว 1", "หัว 2"], ["ก", "ข"]]},
                    {"type": "figure", "path": "missing.png", "caption": "รูปตัวอย่าง"},
                    "ผลลัพธ์ [ยังไม่มีข้อมูล รอเก็บผล]",
                ]},
            ],
        }
    ],
    "references": ["Reference one.", "เอกสารอ้างอิงภาษาไทยลำดับที่สอง", "Reference three."],
    "appendices": [{"title": "ภาคผนวกทดสอบ", "content": ["เนื้อหาภาคผนวก"]}],
}

failures = []


def check(name, ok, detail=""):
    print(f"{'PASS' if ok else 'FAIL'}  {name}" + (f"  -> {detail}" if detail and not ok else ""))
    if not ok:
        failures.append(name)


def run(*args):
    return subprocess.run([sys.executable, *args], capture_output=True, text=True, encoding="utf-8")


def main():
    tmp = tempfile.mkdtemp()
    content = os.path.join(tmp, "content.json")
    with open(content, "w", encoding="utf-8") as f:
        json.dump(SAMPLE, f, ensure_ascii=False)
    template, report = os.path.join(tmp, "template.docx"), os.path.join(tmp, "report.docx")

    for label, args, out in (("template", ["--template", template], template),
                             ("sample report", [content, report], report)):
        r = run(os.path.join(SCRIPTS, "build_report.py"), *args)
        check(f"build {label}", r.returncode == 0 and os.path.exists(out), r.stderr[-500:])
        r = run(os.path.join(SCRIPTS, "check_format.py"), out)
        check(f"check_format passes on {label}", r.returncode == 0, r.stdout[-800:] + r.stderr[-300:])

    doc = Document(report)

    # numbered lists each get their own numbering instance (restart at 1)
    ids = []
    for p in doc.paragraphs:
        if p.style.name == "List Number":
            nid = p._p.pPr.numPr.numId.val
            if not ids or ids[-1] != nid:
                ids.append(nid)
    check("each numbered list restarts at 1", len(ids) == 2 and len(set(ids)) == 2, str(ids))

    # citations stay black; real placeholders are red
    citation_red, placeholder_red = False, False
    for p in doc.paragraphs:
        for r in p.runs:
            text = r.text.replace(ZWSP, "")
            red = r.font.color is not None and r.font.color.rgb is not None
            if re.search(r"\[\d[\d ,]*\]", text) and red and not p.text.startswith("["):
                citation_red = True
            if "ยังไม่มีข้อมูล" in text and red:
                placeholder_red = True
    check("in-text citations [1] are not coloured as placeholders", not citation_red)
    check("[placeholders] are coloured red", placeholder_red)

    # Thai body text carries word-break marks (needs pythainlp)
    body = "\n".join(p.text for p in doc.paragraphs)
    check("Thai word-break marks inserted", ZWSP in body, "is pythainlp installed?")

    # reference numbers are plain [n] at the start of each entry
    refs = [p.text for p in doc.paragraphs if re.match(r"\[\d+\]\t", p.text)]
    check("references numbered [1]..[3]", len(refs) == 3, str(refs))

    # TOC styles use Word's built-in names, otherwise Word ignores them
    toc = [s for s in doc.styles.element.findall(qn("w:style")) if s.get(qn("w:styleId")) == "TOC1"]
    name = toc[0].find(qn("w:name")).get(qn("w:val")) if toc else None
    check("TOC 1 style uses the built-in name 'toc 1'", name == "toc 1", str(name))

    # biography carries the student ID
    check("student ID in biography", "65000000" in body)

    print(f"\n{'all checks passed' if not failures else f'{len(failures)} check(s) failed'}")
    return 1 if failures else 0


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    sys.exit(main())
