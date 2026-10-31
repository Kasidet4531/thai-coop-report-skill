# เล่มสหกิจศึกษา — Agent Skill สำหรับ Claude / Codex

Skill นี้ให้ AI สร้าง **รายงานสหกิจศึกษาฉบับสมบูรณ์** เป็นไฟล์ Word (.docx) ตามรูปแบบที่ภาควิชากำหนด
คุณไม่ต้องตั้งค่าขอบกระดาษ ฟอนต์ หรือเลขหน้าเอง

- ปกนอก กระดาษรองปก ปกใน
- บทคัดย่อภาษาไทยและภาษาอังกฤษ (ABSTRACT) กิตติกรรมประกาศ
- สารบัญ สารบัญตาราง สารบัญรูป แบบอัตโนมัติ
- บทที่ 1–5 เลขหัวข้อ 1.1, 1.2 … เรียงให้เอง
- เอกสารอ้างอิง ภาคผนวก ก ข ค … ประวัติผู้เขียน

**รูปแบบที่ตั้งไว้:**
- กระดาษ A4 ขอบซ้าย 3 ซม. ขอบบน ล่าง และขวา 2.54 ซม.
- ใช้ฟอนต์ TH SarabunPSK ทั้งเล่ม เนื้อหาขนาด 16 ชื่อบทขนาด 18 ตัวหนา
- ส่วนนำใช้เลขหน้าโรมัน (I II III) ส่วนเนื้อหาใช้เลขอารบิก เลขหน้าอยู่กลางล่าง หน้าแรกของแต่ละบทไม่แสดงเลขหน้า
- เนื้อหาจัดแบบ Thai Distributed โดยตัดคำไทยไว้ก่อน ตัวอักษรจึงไม่ถูกยืด

## ใช้ทำอะไรได้บ้าง

พิมพ์บอก AI ตามปกติ เช่น

> ช่วยทำเทมเพลตเล่มสหกิจให้หน่อย ชื่อ นายสมชาย ใจดี ภาควิชาวิศวกรรมไฟฟ้า ฝึกที่บริษัท ABC ปีการศึกษา 2568

> ไฟล์ draft.docx เป็นร่างเล่มสหกิจ ช่วยจัดให้ถูก format หน่อย เนื้อหาห้ามเปลี่ยน

> ช่วยเขียนเล่มสหกิจจากโน้ตนี้ … (ข้อมูลที่ยังไม่มี AI จะเว้นเป็น [ช่องสีแดง] ให้ ไม่แต่งตัวเลขขึ้นเอง)

## ติดตั้ง

### 1. ติดตั้งโปรแกรมที่ต้องใช้ (ทำครั้งเดียว)

- **Python 3.9 ขึ้นไป** แล้วรันคำสั่ง
  ```bash
  pip install python-docx pythainlp
  ```
- **ฟอนต์ TH SarabunPSK** (ฟรี ค้นหาคำว่า "TH SarabunPSK ดาวน์โหลด" ได้จาก SIPA หรือ f0nt.com) ถ้าไม่มี Word จะแสดงเป็นฟอนต์อื่นแทน
- (ไม่บังคับ) **Windows + Microsoft Word** ถ้าต้องการให้สร้าง PDF ตัวอย่างอัตโนมัติ

### 2. ติดตั้ง skill

**Claude Code**: พิมพ์ 2 คำสั่งนี้ในช่องแชท

```
/plugin marketplace add Kasidet4531/thai-coop-report-skill
/plugin install coop-report-docx@thai-coop-report
```

**Claude Desktop / Claude.ai**:
1. กดปุ่ม **Code → Download ZIP** ที่หน้า repo นี้ แล้วแตกไฟล์
2. บีบอัดโฟลเดอร์ `skills/coop-report-docx` ให้เป็นไฟล์ zip
3. ไปที่ **Settings → Capabilities → Skills → Upload skill** แล้วเลือกไฟล์ zip นั้น

หมายเหตุ: Claude.ai บนเว็บอาจยังไม่มี `pythainlp` ติดตั้งไว้ ถ้าไม่มี skill จะจัดเนื้อหาชิดซ้ายแทน Thai Distributed

**Codex** หรือ agent อื่นที่รองรับ Agent Skills (`SKILL.md`):

```bash
git clone https://github.com/Kasidet4531/thai-coop-report-skill.git
```

แล้วคัดลอกโฟลเดอร์ `skills/coop-report-docx` ไปไว้ในโฟลเดอร์ skills ของโปรแกรมนั้น สำหรับ Codex คือ `~/.codex/skills/` ตำแหน่งนี้อาจต่างกันตามเวอร์ชัน ควรเช็กในเอกสารของ Codex อีกครั้ง

### ใช้สคริปต์โดยตรง ไม่ผ่าน AI

```bash
# เทมเพลตเปล่า
python skills/coop-report-docx/scripts/build_report.py --template เล่มสหกิจ.docx

# สร้างจากไฟล์เนื้อหา (ดูรูปแบบไฟล์ได้ที่ skills/coop-report-docx/references/content-schema.md)
python skills/coop-report-docx/scripts/build_report.py content.json เล่มสหกิจ.docx

# ตรวจ format
python skills/coop-report-docx/scripts/check_format.py เล่มสหกิจ.docx
```

## หลังได้ไฟล์แล้ว

- ตอนเปิดไฟล์ ถ้า Word ถามว่า "อัปเดตฟิลด์ไหม" ให้กด **Yes** สารบัญจะขึ้นมาเอง
  ถ้าแก้หัวข้อทีหลัง ให้คลิกขวาที่สารบัญ เลือก **Update Field** แล้วเลือก **Update entire table**
- ข้อความ **[สีแดงในวงเล็บ]** คือส่วนที่ต้องเติมเอง
- โลโก้บนปกเว้นที่ไว้ให้วางรูปเอง
- ถ้าจะเพิ่มหัวข้อใหม่ ให้ใช้สไตล์ Heading 1/2/3 หัวข้อนั้นจึงจะขึ้นในสารบัญ

## ปรับให้เข้ากับมหาวิทยาลัยอื่น

ค่าเริ่มต้นของ skill นี้ทำตามคู่มือของคณะวิศวกรรมศาสตร์ สจล. (KMITL)
- **ชื่อสถาบันและคณะ:** แค่บอก AI ว่าเรียนที่ไหน หรือแก้ `institute` / `faculty` ใน content.json
- **ขอบกระดาษ ขนาดตัวอักษร และค่าอื่น ๆ:** แก้ค่าคงที่ด้านบนของ `scripts/build_report.py` และฟังก์ชัน `setup_section`
  ส่วนกฎทั้งหมดดูได้ที่ `references/format-rules.md`

## License

MIT
