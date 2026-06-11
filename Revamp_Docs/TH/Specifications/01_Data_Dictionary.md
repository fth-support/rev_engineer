# พจนานุกรมข้อมูล (Data Dictionary)

เอกสารนี้แสดงโครงสร้างของตารางในฐานข้อมูลที่ถูกวิเคราะห์มาจากการทำงานของ ADODB ในระบบ `ServiceTransfer`

| ชื่อตาราง (Table Name) | ชื่อคอลัมน์ (Column Name) | ประเภทข้อมูล (Data Type) | ข้อจำกัด (Constraints) | คำอธิบาย (Description) |
| :--- | :--- | :--- | :--- | :--- |
| **TPSTSalHD** | FTTmnNum | VARCHAR | PK | รหัสเครื่องจุดขาย (Terminal Number) |
| | FTShdTransNo | VARCHAR | PK | เลขที่ธุรกรรม (Transaction Number) |
| | FDShdTransDate | DATE | PK | วันที่ทำรายการ |
| | FTShdTransType | VARCHAR | Nullable | ประเภทธุรกรรม |
| | FTShdStaDoc | VARCHAR | Nullable | สถานะเอกสาร (1 = ใช้งาน) |
| | FTShdStaPrcGL | VARCHAR | Nullable | สถานะการประมวลผลบัญชีแยกประเภท |
| **TPSTSalDT** | FTTmnNum | VARCHAR | PK, FK | รหัสเครื่องจุดขาย |
| | FTShdTransNo | VARCHAR | PK, FK | เลขที่ธุรกรรม |
| | FDShdTransDate | DATE | PK, FK | วันที่ทำรายการ |
| | FNSdtSeqNo | NUMERIC | PK | ลำดับรายการ |
| | FTStaSentOnOff | VARCHAR | Nullable | สถานะการซิงค์ข้อมูลเข้า Server (Online/Offline) |
| **TPSTSalRC** | FTTmnNum | VARCHAR | PK, FK | รหัสเครื่องจุดขาย |
| | FTShdTransNo | VARCHAR | PK, FK | เลขที่ธุรกรรม |
| | FDShdTransDate | DATE | PK, FK | วันที่ทำรายการ |
| | FNSrcSeqNo | NUMERIC | PK | ลำดับการรับชำระ |
| | FTStaSentOnOff | VARCHAR | Nullable | สถานะการซิงค์ข้อมูลเข้า Server |
| **TPSTSalCD** | FTTmnNum | VARCHAR | PK, FK | รหัสเครื่องจุดขาย |
| | FTShdTransNo | VARCHAR | PK, FK | เลขที่ธุรกรรม |
| | FDShdTransDate | DATE | PK, FK | วันที่ทำรายการ |
| | FNSdtSeqNo | NUMERIC | PK, FK | ลำดับรายการอ้างอิง |
| | FNScdSeqNo | NUMERIC | PK | ลำดับบัตร |
| | FTStaSentOnOff | VARCHAR | Nullable | สถานะการซิงค์ข้อมูลเข้า Server |
| **TPSTSalePoint** | FTTmnNum | VARCHAR | PK, FK | รหัสเครื่องจุดขาย |
| | FTShdTransNo | VARCHAR | PK, FK | เลขที่ธุรกรรม |
| | FDShdTransDate | DATE | PK, FK | วันที่ทำรายการ |
| | FTSpoMemID | VARCHAR | PK | รหัสสมาชิก |
| | FCSpoPoint | NUMERIC | Nullable | คะแนนที่ได้รับ |
| | FTStaSentOnOff | VARCHAR | Nullable | สถานะการซิงค์ข้อมูลเข้า Server |
| | FTRemark | VARCHAR | Nullable | แฟล็กตรวจสอบ (0 = ยังไม่ประมวลผล, 1 = ประมวลผลแล้ว) |
| | FDDateUpd | DATE | Nullable | วันที่อัปเดต |
| | FTWhoUpd | VARCHAR | Nullable | รหัสผู้ใช้งาน/ระบบที่อัปเดต |
| **TPSTTokenLst** | FTToken | VARCHAR | PK | รหัส Token |
| | FTValue | VARCHAR | PK | ข้อมูลต้นฉบับ (เช่น เลขบัตร) |
| | FTTblName | VARCHAR | PK | ชื่อตารางแหล่งที่มา |
| | FTFldName | VARCHAR | PK | ชื่อฟิลด์แหล่งที่มา |
| | FDDateIns | DATE | Nullable | วันที่สร้าง Token |
| | FTWhoIns | VARCHAR | Nullable | ผู้ที่สร้าง Token |
| | FTMcdCode | VARCHAR | Nullable | รหัสบัตรสมาชิก (เชื่อมโยง) |
| | FDMcdExpDate | DATE | Nullable | วันหมดอายุบัตร |
| | FTTmnNum | VARCHAR | FK | รหัสเครื่องจุดขายอ้างอิง |
| | FTShdTransNo | VARCHAR | FK | เลขที่ธุรกรรมอ้างอิง |
| | FDShdTransDate | DATE | FK | วันที่ทำรายการอ้างอิง |
| | FNStatus | NUMERIC | Nullable | สถานะ Token |
| | FTSrvName | VARCHAR | Nullable | ชื่อ Server |
| **TCNMMallCard** | FTMcdCode | VARCHAR | PK | รหัสบัตรสมาชิก / Token |
| | FDMcdExpDate | DATE | PK | วันหมดอายุ |
| | FTMcdStaAct | VARCHAR | Nullable | สถานะการใช้งาน (Active) |
| | FTMcdName | VARCHAR | Nullable | ชื่อสมาชิก |
| | FDMcdBirthday | DATE | Nullable | วันเกิดสมาชิก |
| | FCMcdAmountPoint | NUMERIC | Nullable | คะแนนยอดซื้อ |
| | FCMcdTotalPoint | NUMERIC | Nullable | คะแนนรวมก่อนหน้า |
| | FCEarned | NUMERIC | Nullable | คะแนนที่ได้รับเพิ่ม |
| | FCBalance | NUMERIC | Nullable | คะแนนคงเหลือ |
| | FDDateUpd | DATE | Nullable | วันที่อัปเดตข้อมูล |
| | FTTimeUpd | VARCHAR | Nullable | เวลาที่อัปเดตข้อมูล |
| | FTWhoUpd | VARCHAR | Nullable | ผู้ที่ทำการอัปเดต |
| **TPSTBPHis** | FTTmnNum | VARCHAR | PK, FK | รหัสเครื่องจุดขาย |
| | FTShdTransNo | VARCHAR | PK, FK | เลขที่ธุรกรรม |
| | FDShdTransDate | DATE | PK, FK | วันที่ทำรายการ |
| | FTShdMemberID | VARCHAR | PK | รหัสสมาชิก |
| | FTShdTransTime | VARCHAR | Nullable | เวลาที่ทำรายการ |
| | FCCstCardPoint | NUMERIC | Nullable | คะแนนบัตรที่ใช้ |
| | FCTotalBP | NUMERIC | Nullable | คะแนนสะสมรวม |
| | FCBalanceBP | NUMERIC | Nullable | คะแนนสะสมคงเหลือ |
| | FDDateUpd | DATE | Nullable | วันที่อัปเดตประวัติ |
| | FTTimeUpd | VARCHAR | Nullable | เวลาที่อัปเดตประวัติ |
| **TSysSync** | FTSscTable | VARCHAR | PK | ตารางต้นทางที่จะซิงค์ |
| | FTSscDocType | VARCHAR | PK | ประเภทเอกสาร |
| | FTSscTableDest | VARCHAR | PK | ตารางปลายทาง |
| | FTSscStaActive | VARCHAR | Nullable | สถานะเปิด/ปิดการซิงค์ตารางนี้ |
| **TCNMTerminalMtn** | FTTmnCode | VARCHAR | PK | รหัสจุดขายย่อย |
| | FTTmnNum | VARCHAR | Nullable | รหัสเครื่องจุดขาย |
| | FTTmnName | VARCHAR | Nullable | ชื่อเครื่องจุดขาย |
