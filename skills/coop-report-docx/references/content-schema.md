# Content file schema (content.json)

`scripts/build_report.py` reads one UTF-8 JSON file. Every key is optional; anything
missing becomes a red `[...]` placeholder in the document, so a partial file still
builds a complete, correctly formatted book.

Inside any text string:
- `**bold**` makes a bold span.
- `[anything]` is rendered red as a placeholder — use it for gaps the student must fill.

```json
{
  "meta": {
    "title_th": "การพัฒนาระบบติดตามสินค้าคงคลังด้วย RFID",
    "title_en": "Development of an RFID Inventory Tracking System",
    "students": ["นายสมชาย ใจดี"],            // one or more names on the cover
    "students_en": ["Mr. Somchai Jaidee"],
    "student_id": "65010000",                  // shown in ประวัติผู้เขียน
    "program": "หลักสูตรวิศวกรรมศาสตรบัณฑิต",     // optional; appended to the department line on the cover
    "department": "วิศวกรรมคอมพิวเตอร์",       // printed as "ภาควิชา" + value
    "department_en": "Computer Engineering",
    "faculty": "วิศวกรรมศาสตร์",                // default วิศวกรรมศาสตร์
    "faculty_en": "Engineering",               // default Engineering
    "institute": "สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง",  // default KMITL
    "academic_year": "2568",
    "advisor": "ผศ.ดร.สมหญิง รักเรียน",          // อาจารย์นิเทศ
    "advisor_en": "Asst. Prof. Dr. Somying Rakrian",
    "mentor": "นายวิชัย เก่งงาน",                // ผู้นิเทศงาน (company supervisor)
    "mentor_en": "Mr. Wichai Kengngan",
    "company": "บริษัท ตัวอย่าง จำกัด",
    "company_en": "Example Co., Ltd.",
    "logo_path": "C:/path/to/logo.png"         // optional; otherwise a placeholder box
  },

  "abstract_th": ["ย่อหน้าที่ 1", "ย่อหน้าที่ 2"],   // string or list of paragraphs
  "keywords_th": ["RFID", "สินค้าคงคลัง"],
  "abstract_en": ["Paragraph 1"],
  "keywords_en": ["RFID", "inventory"],
  "acknowledgement": ["ย่อหน้าขอบคุณ"],

  "chapters": [                                 // omit to get the standard 5-chapter skeleton
    {
      "title": "บทนำ",
      "intro": ["optional paragraphs before 1.1"],
      "sections": [
        {
          "heading": "ความเป็นมาและความสำคัญ",   // numbered automatically → "1.1"
          "content": [ BLOCK, BLOCK, ... ],
          "subsections": [                       // → "1.1.1"
            {"heading": "...", "content": [ BLOCK ]}
          ]
        }
      ]
    }
  ],

  "references": ["รายการอ้างอิง 1", "Reference 2"],   // numbered [1], [2] … — cite in the text as [1]
                                                     // (don't type the [n] yourself; the script adds it)
  "appendices": [                                    // ภาคผนวก ก, ข, ค …
    {"title": "คู่มือการใช้งาน", "content": [ BLOCK ]}
  ],
  "biography": {"ชื่อ-สกุล นักศึกษา": "นายสมชาย ใจดี", "รหัสนักศึกษา": "65010000",
                "ภูมิลำเนา": "...", "การติดต่อ": "...",
                "ประวัติการศึกษา": ["มัธยม …", "ปริญญาตรี …"]},
                // dict → bold label + value column (a list value = several lines);
                // omit it to get these five rows as placeholders (student_id filled from meta)

  "options": {
    "outer_cover": true,       // ปกนอก + กระดาษรองปก before ปกใน
    "flyleaf": true,           // the blank กระดาษรองปก page
    "list_of_tables": null,    // default: only if the report has a table
    "list_of_figures": null,   // default: only if the report has a figure
    "reference_style": "numbered", // "numbered" ([1] …, default) or "apa" (hanging indent, no numbers)
    "chapter_first_page_number": false, // handout: no number on a chapter's first page.
                                    // true = show it (some reviewed books do); then run
                                    // check_format.py with --allow-chapter-page-numbers
    "justify": "thaiDistribute" // body alignment: "thaiDistribute" (default), "both" or "left".
                                // Justified modes depend on the zero-width Thai word breaks
                                // the script inserts (needs pythainlp); use "left" only if
                                // the student asks for ragged-right text.
  }
}
```

(Comments above are for reading only — real JSON must not contain them.)

## BLOCK types

| Block | Result |
|---|---|
| `"plain string"` | Body paragraph, 16 pt, Thai-distributed, first-line indent |
| `{"type": "paragraph", "text": "...", "indent": false}` | Paragraph without first-line indent |
| `{"type": "bullets", "items": ["a", "b"]}` | Bulleted list |
| `{"type": "numbered", "items": ["a", "b"]}` | Numbered list |
| `{"type": "table", "caption": "...", "rows": [["หัว1","หัว2"],["a","b"]]}` | Caption "ตารางที่ 3.1 …" above, grid table, first row bold |
| `{"type": "figure", "path": "img.png", "caption": "...", "width_cm": 12}` | Centered image, caption "รูปที่ 3.1 …" below; missing file → placeholder |

Table and figure numbers restart per chapter (3.1, 3.2 …) and feed the
สารบัญตาราง / สารบัญรูป pages automatically.
