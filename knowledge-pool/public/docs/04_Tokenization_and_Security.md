# ServiceTransfer: Tokenization & Security

ข้อมูลบัตรเครดิตและบัตรสมาชิกมีความอ่อนไหวสูง (Sensitive Data) เพื่อให้สอดคล้องกับมาตรฐานความปลอดภัย (PCI-DSS) ระบบ ServiceTransfer เดิมจึงทำหน้าที่เสมือน Data Protection Gateway ที่จะจับเลขบัตรจริงจาก Local DB มาแปลงเป็นค่า **Token** ก่อนที่จะส่งขึ้น Central DB

---

## 1. Tokenization Flow (กระบวนการแปลงรหัส)

ระบบทำงานร่วมกับ **SafeNet Tokenizer** ซึ่งเป็นบริการภายนอกผ่าน SOAP Web Service (เรียกใช้โดยไลบรารี `MSSOAP.SoapClient30`) รูปแบบ Token ที่ได้คือแบบ **FIRST_SIX_TOKEN** (คงเลข 6 หลักแรกของบัตรไว้ ส่วนที่เหลือจะถูกสุ่มใหม่)

```diagram
tokenization
```

> [!TIP]
> **การค้นหา Token ซ้ำ:** เพื่อไม่ให้ Web Service ทำงานหนักเกินไป ระบบจะหาตาราง `TPSTTokenLst` จาก Central Server ก่อนเสมอ หากพบเลขบัตรเดิมที่เคยแปลงแล้ว (WHERE FTValue = CardNo) จะนำ Token เก่ามาใช้ซ้ำทันที

---

## 2. เงื่อนไขการ Tokenize (Token Conditions Matrix)

การ Tokenize ไม่ได้ทำกับทุกเรคอร์ด แต่มีเงื่อนไขตาม "ชนิดการชำระเงิน" (Tender Type) หรือ "ชนิดส่วนลด" (Discount Type) ดังนี้:

| ตารางที่ถูกประมวลผล (Table) | ฟิลด์เป้าหมาย (Field) | เงื่อนไขที่ต้องแปลง Token (Conditions) |
| --- | --- | --- |
| **TPSTSalHD** (หัวบิล) | `FTCstCardCode` | ทำเสมอ (ALWAYS) |
| **TPSTSalRC** (รับชำระ) | `FTSrcCardNo` | เฉพาะ FTTdmCode (การชำระ) เป็น:<br>`T002` (Credit Card), `T003` (Hire Purchase), `T017` (TC Card) |
| **TPSTSalCD** (ส่วนลด) | `FTScdCardID` | เฉพาะ FNDctNo (ส่วนลด) เป็น:<br>`11`, `12`, `15`, `28` |
| **TPSTSalePoint** (คะแนน) | `FTSpoMemID` | ทำเสมอ (ALWAYS) |

---

## 3. ปัญหาด้านความปลอดภัยของระบบเดิม (Security Vulnerabilities)

แม้ระบบจะมีการทำ Tokenization แต่สถาปัตยกรรมโดยรวมยังพบปัญหาช่องโหว่ความปลอดภัยที่รุนแรง และ**จำเป็นต้องได้รับการแก้ไขทั้งหมดในโครงการ Revamp:**

### 3.1 Plain-Text Credentials (รหัสผ่านฐานข้อมูลและ Token เก็บแบบไม่ได้เข้ารหัส)
- **ปัญหา:** รหัสผ่านของ SafeNet Web Service ถูกเก็บไว้ในตาราง `TSysConfig` แบบข้อความธรรมดา (Plain Text)
- **ปัญหา:** ไฟล์ `AdaIni.Ada` ที่เก็บ Connection String สู่ฐานข้อมูล ถูกล็อกด้วยรหัสผ่านระดับ MS Access เท่านั้น ซึ่งถอดรหัสได้ง่าย
- **เป้าหมายแก้ไข:** ย้ายความลับ (Secrets) ทั้งหมดไปเก็บไว้ใน **Secure Vault** (เช่น Azure Key Vault, HashiCorp Vault) หรือ Environment Variables ที่มีกลไกป้องกัน

### 3.2 SQL Injection Surface (การต่อสตริงสร้าง SQL)
- **ปัญหา:** ฟังก์ชันสร้างคำสั่ง `INSERT`/`UPDATE` ของระบบเดิมอาศัยการนับตัวอักษรและต่อ String ทำให้เสี่ยงต่อการถูกเจาะฐานข้อมูล (SQL Injection) หากระบบ Escape (`'` เป็น `''`) พลาดเพียงจุดเดียว
- **เป้าหมายแก้ไข:** ระบบใหม่ต้องใช้ **Parameterized Queries** (เช่น `cmd.Parameters.Add()`) หรือ **ORM** (เช่น Entity Framework, GORM) ในทุกจุดที่เข้าถึงฐานข้อมูล

### 3.3 No Transport Encryption (ไม่มีการบังคับ SSL/TLS)
- **ปัญหา:** การเปิด Connection แบบตรงข้ามเครือข่ายไปยัง Central DB โดยไม่มีการเข้ารหัส ทำให้ผู้ไม่หวังดีในเครือข่าย (Man-in-the-Middle) ดักจับข้อมูลได้
- **เป้าหมายแก้ไข:** การสื่อสารทั้งหมดจากสาขา (Branch) มาที่ศูนย์กลาง (HQ) จะต้องส่งผ่าน **HTTPS (API Gateway)** หรือเชื่อมต่อ MQ ด้วย **TLS** เท่านั้น ไม่มีการยิงตรงเข้าฐานข้อมูลอีกต่อไป
