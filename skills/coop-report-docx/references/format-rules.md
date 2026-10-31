# รูปแบบรายงานสหกิจศึกษาฉบับสมบูรณ์ (source of truth)

Transcribed from the department's handout "การเขียนรายงานสหกิจศึกษาฉบับสมบูรณ์".
`build_report.py` already implements all of this; read this file when the student
asks *why* something is formatted a certain way, or when editing a document by hand.

## ส่วนประกอบ (order of the book)

**1. ส่วนนำ (front matter)**
1. สันปก — spine text: ชื่อโครงงานสหกิจศึกษา + ปีการศึกษา (printed by the binder, not in the .docx)
2. ปกนอก — outer cover (same layout as the inner cover)
3. กระดาษรองปก — blank white A4 flyleaf
4. ปกในภาษาไทย — inner cover
5. บทคัดย่อภาษาไทย
6. บทคัดย่อภาษาอังกฤษ (ABSTRACT)
7. กิตติกรรมประกาศ
8. สารบัญ
9. สารบัญตาราง (ถ้ามี)
10. สารบัญภาพ / สารบัญรูป (ถ้ามี)

**2. ส่วนเนื้อหา (body)** — chapter titles and the number of sections may be adapted to the project
- บทที่ 1 บทนำ: 1.1 ความเป็นมาและความสำคัญ, 1.2 วัตถุประสงค์ของการวิจัย, 1.3 ขอบเขตของการวิจัย, 1.4 วิธีการดำเนินการวิจัย, 1.5 ประโยชน์ที่คาดว่าจะได้รับ
- บทที่ 2 แนวคิด ทฤษฎีและงานวิจัยที่เกี่ยวข้อง: 2.1 ทฤษฎีที่เกี่ยวข้อง, 2.2 งานวิจัยที่เกี่ยวข้อง
- บทที่ 3 วิธีดำเนินการวิจัย
- บทที่ 4 ผลการวิจัย
- บทที่ 5 สรุปผลการวิจัยและข้อเสนอแนะ

**3. ส่วนอ้างอิง (back matter)**: เอกสารอ้างอิง, ภาคผนวก, ประวัติผู้เขียน, กระดาษรองปกหลัง, ปกหลัง (the last two are physical binding, not in the .docx)

## การพิมพ์ (typesetting)

- White A4 paper, 70–80 gsm; black ink, colour allowed for pictures
- Margins: left 3 cm (1.5"), top/bottom/right 2.54 cm (1")
- One font for the whole book: TH SarabunPSK (alternatives allowed by the handout: Angsana New/UPC, Browallia New/UPC, Cordia New/UPC)
- Page numbers, bottom centre:
  - Front matter from บทคัดย่อ onward: Roman numerals I II III IV V …
  - Body: Arabic 1 2 3 …, **no number on the first page of each chapter**
- Chapter titles: 18 pt bold, centred. Body text: 16 pt regular.

## ปก (cover, inner cover same as outer)

Top to bottom, centred: logo → blank line (20 pt) → "รายงานสหกิจศึกษาฉบับสมบูรณ์" (bold 22) →
blank line (20) → Thai title (bold 20) → English title (bold 20) → ~3 blank lines (20) →
student name (bold 18) → space → ภาควิชา…, คณะวิศวกรรมศาสตร์, สถาบันเทคโนโลยีพระจอมเกล้าเจ้าคุณทหารลาดกระบัง (bold 18) → ปีการศึกษา 25.. (bold 18).

## บทคัดย่อ (Thai) / ABSTRACT (English)

Header lines — label 16 bold, value 16 regular, label and value separated by 2 spaces:
ชื่อโครงงานสหกิจศึกษา / ชื่อ-สกุล นักศึกษา / คณะ … ภาควิชา … / ชื่อ-สกุล อาจารย์นิเทศ / ชื่อ-สกุล ผู้นิเทศงาน / สถานประกอบการ
(English: Cooperative Title / Student intern name / Faculty: Engineering  Department: / Advisor name / Mentor name / Company)

Then: blank line → "บทคัดย่อ" / "ABSTRACT" 18 bold centred → blank line → body 16 regular,
first line starts at the 9th character, max one page → blank line → "คำสำคัญ :" / "Keywords :" in 14 pt.
