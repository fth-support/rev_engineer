Attribute VB_Name = "mMS"
Option Explicit
Public Const cCS_Top800x600 = 990 '930
Public Const cCS_Left800x600 = 0
Public Const cCS_Top680x480 = 850
Public Const cCS_Left680x480 = 0
'
''general  msg
'Public Const tCS_MSG0 = "ไม่สามารถเปิดแฟ้มข้อมูลได้;Can not open the Table"
'Public Const tCS_MSG1 = "เกิดเหตุการณ์ ดังนี้:;This occur happened:"
'Public Const tCS_MSG2 = "หมายเลข:;No:"
'Public Const tCS_MSG3 = "ไม่มีแฟ้ม;Not found" 'tms_CN008026
'Public Const tCS_MSG4 = "มีอยู่แล้วคุณต้องการเขียนทับแฟ้มดังกล่าวหรือไม่?;has been exist, Do you want to replace?"
'Public Const tCS_MSG5 = "** ข้อมูลเก่าของคุณจะถูกลบทิ้งหมด !!!;** Your old data will be lost !!!"
'
''record msg
'Public Const tCS_MSG6 = "คุณต้องการลบข้อมูลนี้หรือไม่?;Do you want to delete this data?"
'Public Const tCS_MSG7 = "ข้อมูลมีการเปลี่ยนแปลง คุณต้องการออกจากหน้าจอหรือไม่?;Data have been change. Do you really want to exit?"
'Public Const tCS_MSG8 = "เพิ่มข้อมูลใหม่...;Adding..."
'Public Const tCS_MSG9 = "แก้ไขข้อมูล...;Editing..."
'Public Const tCS_MSG10 = "ค้นหาข้อมูล...;Finding..."
'
''warning msg
'Public Const tCS_MSG11 = "ข้อมูลนี้มีอยู่แล้ว กรุณาเปลี่ยนใหม่;This data has been exist. Please retry again."
'Public Const tCS_MSG12 = "ข้อมูลไม่สมบูรณ์ กรุณาใส่ใหม่;Data is not complete. Please fill in properly."
'Public Const tCS_MSG13 = "ต้องการออกจากระบบ จริงๆหรือไม่?;Do you really want to close?"
'Public Const tCS_MSG14 = "ข้อมูลมีการเปลี่ยนแปลง คุณต้องการบันทึก ก่อนหรือไม่?;Data have been changed. Do you want to save changed?"
'Public Const tCS_MSG15 = "ลบข้อมูล...;Deleting..."
'Public Const tCS_MSG16 = "ไม่พบข้อมูล หรือยังไม่มีการกำหนดข้อมูล;Not found data or not defined yet." 'tms_CN005108
'Public Const tCS_MSG17 = "รหัสผ่าน หรือรหัสผู้ใช้ผิดพลาด;Login or Password incorrect"
'Public Const tCS_MSG18 = "รหัสผ่านไม่ถูกต้อง;Password incorrect"
'Public Const tCS_MSG19 = "คุณต้องการออกจากระบบชั่วคราว จริงหรือไม่?;Do you really want to temporary exit?"
'Public Const tCS_MSG20 = "ยินดีต้อนรับกลับสู่การทำงาน;Welcome Back process"
'Public Const tCS_MSG21 = "ข้อมูลเกิดความผิดพลาดกรุณาติดต่อผู้ควบคุมระบบ;Invalid data Please contact Administrator."
'Public Const tCS_MSG22 = "ฟังก์ชันนี้ยังไม่ได้กำหนด หรือไม่ได้รับอนุญาตให้ใช้งาน;This function is undefined or not allowed to use." ''tms_CN005107
'Public Const tCS_MSG23 = "คุณต้องการที่จะปิดการขาย (sign out)  และออกจากระบบจริงหรือไม่;Do you really want to close your selling and signout?"
'Public Const tCS_MSG24 = "คุณต้องการที่จะยกเลิก รายการทั้งหมดนี้หรือไม่ ?;Do you want to IPV ?" 'tms_CN005110
'Public Const tCS_MSG25 = "ข้อมูลสินค้าเป็นประเภทใช้รหัสราคา" & vbCrLf & "แต่ยังไม่มีการกำหนดราคา คุณต้องทำการกำหนดราคาก่อน" & vbCrLf & "ราคาที่ใช้ปัจจุบันจะเป็นราคาขายปกติ"
'Public Const tCS_MSG26 = "This product has been set as price code type." & vbCrLf & "But you did not select the price code." & vbCrLf & "So, the current price will be set to normal price."
'Public Const tCS_MSG27 = "ข้อมูลไม่ถูกต้อง;Data invalid" 'tms_CN005111
'Public Const tCS_MSG28 = "เงินทอนมีค่าเกินกว่า 1000 บาท กรุณาใส่ใหม่;The Change is over than 1000 unit."  'check if receive >1000
'Public Const tCS_MSG29 = "ยังไม่มีการกำหนดหัวท้ายใบเสร็จ;Head and Foot of slip message have not been set yet."
'Public Const tCS_MSG30 = "รายการโปรโมชันของแถมยังไม่ครบ ตามรายการข้างล่างนี้;Some items of promotion 'Buy X Get Y Free' are not complete as follow:"
'Public Const tCS_MSG31 = "คุณต้องการจะกลับไปยังหน้าจอการขายใหม่ หรือไม่ ?;Do you want to go back to get the Free items?"
'Public Const tCS_MSG32 = "คุณต้องการพักรายการปัจจุบัน หรือไม่ ?;Do you want to Suspend Transaction now?" 'tms_CN005116
'Public Const tCS_MSG33 = "คุณต้องการยกเลิก (Canceled) สินค้ารายการสุดท้ายหรือไม่?;Do you want to canceled the this item?"
'Public Const tCS_MSG34 = "คุณต้องการจัดสินค้ารายการนี้ เป็นสินค้าฟรีหรือไม่?;Do you want this product to be free item?"
'Public Const tCS_MSG35 = "สินค้าที่จะคืน เป็นการขายวันนี้หรือไม่?;Are the products has been sold today?"
'Public Const tCS_MSG36 = "คุณต้องการคืนรายการทั้งหมด หรือไม่?;Do you want to refund all items in bill?"
'Public Const tCS_MSG37 = "มีการชำระเงินไปแล้วบางส่วน ไม่สามารถทำการลด/ชาร์จได้อีก;You can not charge or discount when payment was received"
'Public Const tCS_MSG38 = "เครื่องพิมพ์นี้ไม่สามารถพิมพ์ใบเสร็จเต็มรูปได้;This printer can not support full slip printing."
'Public Const tCS_MSG39 = "คุณต้องการยกเลิก (Canceled) สินค้ารายการที่เลือกนี้หรือไม่?;Do you want to canceled this selected item?"
'Public Const tCS_MSG40 = "ไม่พบข้อมูล หรือบิลนี้มีการคืนสินค้าไปแล้ว;Not found data or this receipt has been refunded."
'Public Const tCS_MSG41 = "มีรายการขายไปแล้ว ไม่สามารถพิมพ์ข้อความช่วยเหลือได้;Some items were saled so can not print helping."
''Public Const tCS_MSG42 = "จำนวนเงินรับ มากกว่าเงินที่ต้องชำระ, ชำระโดยคูปอง,เช็ค,บัตรเครดิต จะไม่มีเงินทอน" & vbCrLf & "ดังนั้นคุณต้องการที่จะ รับชำระด้วยยอดทั้งหมดที่เหลือหรือไม่?"
'Public Const tCS_MSG42 = "จำนวนเงินรับชำระมากกว่า มูลค่าสินค้า/บริการ ชำระด้วย คูปอง, บัตรเครดิต, บัตรกำนัล และเช็คจะไม่มีเงินทอน. " & vbCrLf & " คุณต้องการที่จะรับชำระต่อหรือไม่ ?;"
'Public Const tCS_MSG43 = ";Tender amount more than total bill,Paid by Coupon,Credit Card,Gift Voucher and cheque will not change." & vbCrLf & " Do you want to continue?"
''Public Const tCS_MSG43 = "Paying by Coupon, Cheque and Credit Card, receiving money must less than or equal total amount." & vbCrLf & "So, do you want to accept all balance?"
'Public Const tCS_MSG44 = "ไม่มีข้อมูล เพื่อที่จะทำรายการได้;No any items in grid to do the function." 'tms_CN005104
'Public Const tCS_MSG45 = "ตรวจสอบลิ้นชัก ว่าทำการต่อกับคอมพอร์ตถูกต้องหรือไม่;Please check your drawer, does it connect with com port properly."
'Public Const tCS_MSG46 = "มีรายการบางส่วนคงค้างอยู่ จึงไม่สามารถทำงานฟังก์ชันนี้ได้;This function unavailable if there are some items in grid." 'tms_CN005109
'Public Const tCS_MSG47 = "จำนวนเงินสำหรับรายการคืน จะต้องน้อยกว่าหรือพอดีกันกับยอดรวมทั้งหมด;Money for refunding must be less than or equal total amount."
'Public Const tCS_MSG48 = "สินค้ารายการนี้ได้ถูกยกเลิกรายการไปแล้ว ไม่สามารถเลือกเพื่อคืนรายการได้;This item has been voided, so can not be selected."
'Public Const tCS_MSG49 = "คุณต้องมีระดับการทำงานสูงกว่านี้;You must have higher level."
'Public Const tCS_MSG50 = "คุณต้องการที่จะแสดงราคาต้นทุนด้วยหรือไม่?;Do you really want to show the product cost?"
'Public Const tCS_MSG51 = "ข้อมูลนี้เป็นรายการคืนสินค้า ไม่สามารถที่จะพักรายการได้;This bill is a refunding bill, Can not hold bill." 'tms_CN005115
'Public Const tCS_MSG52 = "จำนวนเงินส่วนลด เกินกว่าจำนวนเงินรวม ที่สามารถลดได้;Discount amount is more than Amount. can discount"        '*Tee 46-10-06
'Public Const tCS_MSG53 = "สินค้านี้มีรายการโปรโมชั่น ไม่สามารถทำการ ลด,ชาร์จ หรือ ฟรี ได้;This is a promotion item, price could not be discount ,charge or free"    '*Tee 49-09-19
'Public Const tCS_MSG54 = "สินค้านี้มีการยกเลิก ไม่สามารถทำการลดได้อีก;This item has been voided, can not discount the price."
'Public Const tCS_MSG55 = "คุณต้องการคืนสินค้า โดยอ้างถึงเลขที่ใบเสร็จอย่างย่อหรือไม่?;Do you want to return product that refer to bill number or not?" 'tms_CN005118
'Public Const tCS_MSG56 = "ต้องใส่ข้อมูลตัวเลขก่อน แล้วจึงกดฟังก์ชันทำงาน หรือข้อมูลที่ใส่ไม่ใช่ตัวเลข;You must input number before press function, or data you input is not numeric." 'tms_CN005112
'Public Const tCS_MSG57 = "บิลปัจจุบันเป็นการคืนสินค้าอยู่แล้ว จึงไม่สามารถทำการคืนซ้ำได้อีก;Current document is refurn bill. So can not do it again." 'tms_CN005117
'Public Const tCS_MSG58 = "!!! บิลปัจจุบันเป็นการคืนสินค้า !!!;!!! Current bill is refunding !!!"
'Public Const tCS_MSG59 = "เครื่องพิมพ์ไม่สมบูรณ์ กรุณาติดต่อผู้ดูแลระบบ;Printer is not working. Contract Administrator."
'Public Const tCS_MSG60 = "จำนวนคูณ จะต้องไม่เกิน 6 หลัก (สูงสุด=999999);Number of multiple must less than 6 digits(Max=999999)." 'tms_CN005113
Public Const tCS_MSG61 = "โปรแกรมหน้าร้านถูกเปิดใช้อยู่แล้ว ไม่สามารถเปิดซ้ำได้อีก;This program is already opened, so can not run it again."
'Public Const tCS_MSG62 = "ท่านได้เข้าสู่ระบบการขายไปแล้ว แต่ยังไม่ได้ปิดรอบ ณ เครื่องจุดขาย;You are already signed in at the pos no:"      '*GP 46-11-03 ยกเลิก lognow
'Public Const tCS_MSG63 = "แคชเชียร์ยังไม่ปิดรอบการขาย(Sign out) ของวันที่ผ่านมา วันที่ขายจะยังเป็นของวันที่เดิมอยู่ คุณควรจะทำการปิดรอบการขายก่อนทำงานอย่างอื่นต่อ;Some cashiers did not signed out on the last sale day. Therefore before working others, you must sign out cashier as follow:"
'Public Const tCS_MSG64 = "เลขที่บิลไม่ครบ 19 หลัก;Bill number is than 19 letters."
'Public Const tCS_MSG65 = "ไม่พบหมายเลขเครื่อง หรือหมายเลขเครื่องได้ถูกใช้ไปแล้ว;Not found this machine number or it has been used."
'Public Const tCS_MSG66 = "สินค้าเป็นรายการประเภท Serial คุณต้องกำหนด Serial ก่อน;Type of product is serial you should define serial before"
'Public Const tCS_MSG67 = "เลือกรายงานไม่ถูกต้องกับรูปแบบ;Selected report unmatching format."
'Public Const tCS_MSG68 = "เลขที่บิลเกินรูปแบบที่กำหนดไว้ ไม่สามารถทำการขายต่อได้ โปรแกรมจะทำการปิดรอบการขายให้;Bill  number "
'Public Const tCS_MSG69 = "คุณต้องการแสดงหน้าจอให้เป็น แบบสัมผัส หรือไม่?;Do you want to change screen into touching style?"
'Public Const tCS_MSG70 = "หมายเลขเครื่องขายนี้ จะถูกกำหนดให้ใช้กับเครื่องนี้ตลอดไป" & vbCr & "คุณต้องการที่จะทำงานต่อไปหรือไม่?;This Pos machine will be fixed with this computer forever." & vbCr & "Do you want to continue..?"
'Public Const tCS_MSG71 = "ระดับราคาขายต่างกัน ไม่สามารถแก้ไขได้(เปลี่ยนโดยการยกเลิกเอกสาร อย่างเดียว);Difference price level cannot change (just change by cancel document)"
'Public Const tCS_MSG72 = "<- รับของแถมฟรี ณ. จุดรับของแถม ->;<- Get the gifts free at the counter ->"
'Public Const tCS_MSG73 = "ตรวจสอบขนาดพื้นที่ของฮาร์ดดิสก์ เพราะเหลือพื้นที่ใช้งานน้อย;check free space of harddisk because to minimum space" 'tms_CN005122
'Public Const tCS_MSG74 = "บิลนี้มีการคืนสินค้าไปแล้ว " & vbCrLf & "คุณต้องการคืนสินค้าอีกหรือไม่?;This receipt has been refunded." & vbCrLf & "Do you want to refund product or not?"
'Public Const tCS_MSG75 = "สินค้าไม่เคลื่อนไหว ไม่สามารถทำการขายได้;This product is inactive, not for sale."
'Public Const tCS_MSG76 = "กรุณากลับไป ปิดรอบการขาย(Sign Out) ณ เครื่องขาย ข้างต้นก่อน ถึงจะทำการขายต่อได้;Please go back to sign out the machine above before continue." '*GP 46-11-03
'Public Const tCS_MSG77 = "สินค้ารายการนี้ ราคาขาย เท่ากับ ศูนย์ คุณยืนยันที่จะขายสินค้าด้วยราคาศูนย์บาทหรือไม่ ?;This Product   sale price is zero Do you want to sale it with price zero?"     '*Tee 46-11-19
'Public Const tCS_MSG78 = "ไม่สามารถขายสินค้ารายการนี้ได้เนื่องจากยังไม่กำหนดราคาต้นทุน;Cannot sale this product because cost is zero"     '*Tee 46-11-21
'Public Const tCS_MSG79 = "ไม่สามารถขายสินค้าจำนวนน้อยกว่าหรือเท่ากับศูนย์ กรุณาใส่จำนวนใหม่;Cannot sale product quantity less than zero or equal zero  Please input quantity again"     'tms_CN005114
'Public Const tCS_MSG80 = "จำนวนคูณ จะต้องไม่เป็นจำนวนทศนิยม;Number of multiple is not a decimal point number."       '*GP 46-12-26
'Public Const tCS_MSG81 = "ไม่สามารถขายได้เพราะ วันที่ขายที่เลือกน้อยกว่าวันที่ ที่เคยขาย;cannot sale because  sale date less than data in database"     '*tee 47-02-26
'Public Const tCS_MSG82 = "คุณต้องบันทึกเหตุผลการยกเลิก;You must save void remark"       '*Tee 49-03-03
'Public Const tCS_MSG83 = "!!! บิลปัจจุบันเป็นเอกสารรายรับ !!!;!!! Current bill is Income document !!!"
'Public Const tCS_MSG84 = "!!! บิลปัจจุบันเป็นเอกสารการคืนรายรับ !!!;!!! Current bill is refunding Income document !!!"
'Public Const tCS_MSG85 = "!!! บิลปัจจุบันเป็นเอกสารรายจ่าย !!!;!!! Current bill is Expenses document !!!"
'Public Const tCS_MSG86 = "!!! บิลปัจจุบันเป็นเอกสารการคืนรายจ่าย !!!;!!! Current bill is refunding Expenses document !!!"
'Public Const tCS_MSG87 = "ขอบคุณ สำหรับการลงทะเบียน  " & vbCr & "ยินดีต้อนรับสู่การใช้งาน ADA POS ; Thank you for your Registration " & vbCr & " Welcome to enjoy ADA POS"
'Public Const tCS_MSG88 = "หมายเลข การลงทะเบียนไม่ถูกต้อง หรือยังไม่ได้ ลงทะเบียน " & vbCr & "กรุณาติดต่อบริษัท เอด้าซอฟต์ เพื่อรับหมายเลขลงทะเบียน;Invalid Register number or  Yor do not register program " & vbCr & " Please contact ADA SOFT Company  for new register number "
'Public Const tCS_MSG89 = "Tel. 660-25301681-4"
'Public Const tCS_MSG90 = "จำนวนผู้การเข้าใช้งานเกินกำหนด ที่ลงทะเบียนไว้; Overflow login more than your register "
'Public Const tCS_MSG91 = "ไม่มีพบฐานข้อมูล คุณต้องการสร้างใหม่หรือไม่;Not found database Do you want to create new database"
'
''For AdaEJ
'Public Const tCS_MSG92 = "สร้างฐานข้อมูล EJ ไม่สำเร็จ ต้องออกจากโปรแกรมก่อน แล้วลองใหม่อีกครั้ง;Create data EJ not successfull exit from program and try again"
Public Const tCS_MSG93 = "ยังไม่ได้กำหนดที่ตั้งแฟ้มข้อมูล สำหรับไฟล์ AdaEJ กรุณาตรวจสอบที่ตั้งแฟ้มข้อมูล จากโปรแกรมAdaTools ดังนี้ " & vbCr & " (สามารถดูได้จาก เครื่องมือ/ค่าเริ่มต้น.../Tab เอด้า EJ);Not define  destination of  folder AdaEJ, Please recheck again" & vbCr & " (Program AdaTools Menu Tools/Initial Setting.../Tab AdaEJ)"
Public Const tCS_MSG94 = "ไม่พบที่ตั้งแฟ้มข้อมูล สำหรับไฟล์ AdaEJ กรุณาตรวจสอบที่ตั้งแฟ้มข้อมูล จากโปรแกรมหลังร้าน ดังนี้ " & vbCr & " (สามารถดูได้จาก เมนู/เครื่องมือ/ตัวเลือก/ไฟล์ AdaEJ);Not found folder AdaEJ, Please recheck again"
'Public Const tCS_MSG95 = "สร้างไฟล์ EJ ไม่สำเร็จ ต้องออกจากโปรแกรมก่อน แล้วกรุณาติดต่อผู้ดูแลระบบ;Create file EJ not successfull exit from program and please contact Administrator." 'tms_CN005121
'Public Const tCS_MSG96 = "สินค้ายกเว้น ไม่อนุญาตให้ลด/ชาร์จ สินค้านี้;Excpt product, Not allow to discount or charge this item"        '*Tee 48-05-31
'Public Const tCS_MSG97 = "ไม่สามารถชำระคูปองแบบส่วนลด(หักภาษี)มากกว่ายอดเอกสาร;Cannot pay discount coupon(Cal. Vat) more than total of document"
'Public Const tCS_MSG98 = "ไม่สามารถชำระคูปองแบบส่วนลด(หักภาษี)มากกว่ายอดสินค้าที่ลดได้;Cannot pay discount coupon(Cal. Vat) more than total of product that allow discount"
'Public Const tCS_MSG99 = "!!! บิลปัจจุบันเป็นการขายสินค้า !!!;!!! Current bill is sale document !!!"
'Public Const tCS_MSG100 = "คุณต้องการจะกลับไปยังรูปแบบเอกสารการขายใหม่ หรือไม่ ?;Do you want to go back to selling document?"
'
'Public Const tCS_MSG001 = "คุณไม่ได้รับสิทธิ์อนุญาตให้ใช้โปรแกรมนี้;You do not have any permission to use this program."
'Public Const tCS_MSG002 = "R:กรุณาติดต่อ บริษัทเอด้า ซอฟท์;R:Please contact ADA Soft Co., Ltd."
'Public Const tCS_MSG003 = "W:กรุณาติดต่อ บริษัทเอด้า ซอฟท์;W:Please contact ADA Soft Co., Ltd."
'Public Const tCS_MSG004 = "โปรแกรมนี้เป็นชุดทดลองใช้ คุณสามารถบันทึกรายการได้เพียง 50 รายการ;This is demo version. You have limit only 50 records."
'Public Const tCS_MSG005 = "ข้อมูลนี้มีการอ้างอิงจากแฟ้มข้อมูลอื่นจึงไม่สามารถลบได้;This data has been referred from another file. Therefore can not be deleted."
'Public Const tCS_MSG008 = "ยังไม่ได้กำหนดที่ตั้งแฟ้มข้อมูล กรุณาตรวจสอบอีกครั้ง;Not define  destination of  file Please recheck again"
'
'Public Const tMS_CN000 = "ไม่สามารถเปิดแฟ้มข้อมูลได้;Can not open the file."
'Public Const tMS_CN005 = ";ประเภทการติดต่อ ไม่ตรงกับการติดต่อเดิม, คุณต้องการที่จะทำงานต่อหรือไม่?" & vbCr & vbCr & "Type of Connection is not the same last selected. Do you really want to continue?"
'Public Const tMS_CN006 = "คุณต้องการลบข้อมูลนี้หรือไม่?;Do you want to delete this data?"
'Public Const tMS_CN011 = "ข้อมูลนี้มีอยู่แล้ว กรุณาเปลี่ยนใหม่;This data has been exist. Please try again."
'Public Const tMS_CN012 = "ข้อมูลไม่สมบูรณ์ กรุณาใส่ใหม่;Data is not complete. Please fill in properly."
'Public Const tMS_CN013 = "ต้องการออกจากระบบ จริงๆหรือไม่?;Do you really want to exit?"
'Public Const tMS_CN014 = "ข้อมูลมีการเปลี่ยนแปลง คุณต้องการบันทึก ก่อนหรือไม่?;Data have been changed. Do you want to save changed?"
'Public Const tMS_CN016 = "ไม่พบข้อมูล หรือยังไม่มีการกำหนดข้อมูลนี้;Not found data or this data has not defined yet." 'tms_CN006003
'Public Const tMS_CN017 = "รหัสผ่าน หรือรหัสผู้ใช้ผิดพลาด;User's code or password incorrect."
'Public Const tMS_CN018 = "มีการชำระเงินบางส่วนแล้วไม่สามารถใช้คูปองได้;This invoice has paid some part cannot used coupon"
'Public Const tMS_CN019 = "ไม่มีข้อมูลในฐานข้อมูล;No data exist in database."
'Public Const tMS_CN020 = "ไม่สามารถลดเอกสารภายหลังการชาร์จท้ายบิลได้;Cannot discount this invoice after charge "        '*Tee 49-01-18
'Public Const tMS_CN024 = "ข้อมูลนี้มีการอ้างอิงจากแฟ้มข้อมูลอื่นจึงไม่สามารถลบได้;This data has been refered from another file. Therefore can not be deleted."
'Public Const tMS_CN040 = "การประมวลผลเสร็จสมบูรณ์;Processing completed. !!!"
'Public Const tMS_CN043 = "คุณไม่ได้รับสิทธิ์อนุญาตให้กระทำการนี้;You do not have any permission to do this process."
'Public Const tMS_CN048 = "รหัสรายการปัจจุบันถูกใช้งานไปแล้ว หรือ รหัสรายการไม่ถูกต้อง" & vbCr & "คุณต้องการให้โปรแกรมสร้างรหัสใหม่ตามรูปแบบที่ใช้งานปัจจุบันหรือไม่?" & " ;This code has been used. Do you want  to  generate code again?"
'Public Const tMS_CN056 = "ยังไม่มีการกำหนด การตั้งรหัสอัติโนมัติ กรุณาติดต่อผู้ควบคุมระบบ;The auto number has not defined yet, Please contact Administrator."
'Public Const tMS_CN070 = "ข้อมูลเกินขนาดรูปแบบที่กำลังใช้อยู่ กรุณาเปลี่ยนรูปแบบรหัสใหม่;Overflow data in this format, Please change data format of automation code." 'tms_CN005120
'Public Const tMS_CN076 = "ไม่สามารถติดต่อกับฐานข้อมูลได้ กรุณาตรวจสอบฐานข้อมูลเซอร์วิส หรือติดต่อผู้ควบคุมระบบ;Can not connect the database, please check database service or contract Administrators." 'tms_CN005027
'Public Const tMS_CN080 = "การประมวลผลไม่สมบูรณ์ หรือ ไม่มีข้อมูลในช่วงที่ต้องการ;Processing is not successful or No data that you were selected."
'Public Const tMS_CN106 = "การบันทึกข้อมูลสมบูรณ์;Save complete"
'Public Const tMS_CN129 = "หน้าจอนี้ถูกเปิดใช้อยู่แล้ว ไม่สามารถเปิดซ้ำได้อีก;This document  is already opened, so you can not open it again."
'Public Const tMS_CN130 = "ในลิ้นชักของคุณ มีเงินสดไม่พอ;Not enough cash in your cash drawer"
'Public Const tMS_CN131 = "ไม่สามารถปิดรอบการขายได้เนื่องจากมีรายการพักบิลค้างอยู่;You cannot signout because you not yet return bill some record"
'Public Const tMS_CN132 = "กรุณาเลือกเลขที่เอกสาร;Please select document number."
'Public Const tMS_CN133 = "เปลี่ยนฐานข้อมูลเป็นแบบ Off-Line, คุณต้องการเข้าสู่โหมดการขายแบบ Off-Line หรือไม่;Change database to Off-Line mode, Do you want to chage to Off-Line mode"
'Public Const tMS_CN134 = "การคืนสินค้าแบบอ้างถึงบิล ไม่สามารถเพิ่ม/ลด รายการสินค้าได้;Refund refer bill cannot increase or decrease  items " 'tms_CN005123
'Public Const tMS_CN135 = "ไม่สามารถทำงานต่อได้เนื่องจากวันที่ของระบบน้อยกว่าวันที่การขาย;This process not work  because system data less than sale date"
'Public Const tMS_CN136 = "ยังไม่มีการกำหนด การตั้งรหัสสาขา กรุณาติดต่อผู้ควบคุมระบบ;The branch code has not defined yet, Please contact Administrator."
'Public Const tMS_CN137 = "ไม่สามารถกำหนดฐานข้อมูล On-line/Off-line ซ้ำกันได้ กรุณาติดต่อผู้ควบคุมระบบ;Dupplicate On-line and  Off-line  database , Please contact Administrator."
'Public Const tMS_CN138 = "ไม่พบประเภทของ EDI ว่าเป็นการรับหรือการคืน กรุณาตรวจสอบแล้วลองใหม่อีกครั้ง หรือติดต่อสำนักงานใหญ่;Not found type of EDI please try again or contact Head Office."
'Public Const tMS_CN139 = "ไม่พบหมายเลข EDI คุณต้องการที่จะให้โปรแกรมสร้างเอกสารให้อัตโนมัติหรือไม่?;Not found EDI code. Do you want to generate automatic document?"
'Public Const tMS_CN140 = "ไม่พบหมายเลข EDI กรุณาติดต่อสำนักงานใหญ่;Not found EDI code, please contact Head Office."
'Public Const tMS_CN141 = "เนื่องจากไม่พบหมายเลข EDI โปรแกรมได้สร้างเอกสารอัตโนมัตให้แล้ว;Because not found EDI code, Therefore program has already automatic created document."
'Public Const tMS_CN142 = "พบหมายเลข EDI และได้ประทับวันที่และเวลาเรียบร้อยแล้ว;Found EDI code and completely stamped."
'Public Const tMS_CN143 = "เอกสารนี้ ได้ทำการประทับวันที่และเวลาไปแล้ว;This document has already been stamped."
'Public Const tMS_CN144 = "หมายเลข EDI นี้ ได้ทำการบันทึกไปแล้ว;This EDI code has already been saved."
'Public Const tMS_CN145 = "สินค้าทั่วไปไม่สามารถใส่จำนวนขายแบบมีทศนิยมได้;General product can't define decimal quantity"
'
'Public Const tMS_CN154 = "ไม่สามารถตั้งค่าชนิดของปฏิทิน เป็นแบบ คริสต์ศักราช และ รูปแบบของวันที่แบบย่อ เป็นแบบ dd/MM/yyyy ได้;Can not setting the calendar type are Christian Era and short date format = dd/MM/yyyy."
'Public Const tMS_CN501 = "คุณต้องการที่จะลบรูปภาพนี้หรือไม่;Do you want to remove picture ?"
'
'Public Const tMS_CN155 = "ไม่สามารถใช้งานส่วนนี้ได้เนื่องจากระบบยังไม่มีการกำหนดสกุลเงินท้องถิ่น;Unable to process this section. Because the local currency was undefined."
'
'
'Public Const tCS_MSG101 = "หมายเลขคูปองไม่ถูกต้อง;Invalid Coupon No." '*TW 50-02-22
'Public Const tCS_MSG102 = "คูปองหมดอายุ หรือยังไม่ถึงวันกำหนดเริ่มใช้งาน;Coupon has been expired or coupon not ready to use now." '*TW 50-02-22
'
'Public Const tCS_MSG103 = "หมายเลขเครื่องจุดขายไม่ตรงกับเอกสารที่จะพิมพ์;Pos code incorrect." '*TW 50-03-06
'
''*TW 50-04-27
Public Const tMS_CN502 = "The POS machine is not work properly." & vbCrLf & "Please contact system administrator," & vbCrLf & "Or restart program and try again.;The POS machine is not work properly." & vbCrLf & "Please contact system administrator," & vbCrLf & "Or restart program and try again." 'tms_CN005127
''*TW 50-06-06
'Public Const tMS_CN503 = "สินค้ารายการนี้ ราคาขาย เท่ากับ ศูนย์;This Product   sale price is zero"
'
''*Pla 50-08-08
'Public Const tMS_CN504 = "คุณต้องการที่จะพักบิลขณะทำรายการหรือไม่ ?;Do you want to Suspend Transaction now ?"
'Public Const tMS_CN505 = "คุณต้องการที่จะเรียกบิลที่พักไว้มาทำการขายต่อหรือไม่ ?;Do you want to Recall Transaction now ?" 'tms_CN005124
'Public Const tMS_CN506 = "คุณต้องการที่จะขายรายการสินค้านี้หรือไม่ ?;Are you will want to sell goods this list ?" 'tms_CN005125
'
'Public Const tMS_CN507 = "การค้นคืนผิดพลาด;Recalled failed" ''tms_CN005126
''Public Const tMS_CN508 = "การค้นคืนผิดพลาด;IPV and Void failed"
''Public Const tMS_CN509 = "จำนวนเป็นศูนย์;Amount zero or then zero"
''Public Const tMS_CN510 = "ไม่สามารถทำรายการยกเลิกการขายได้;Already IPV'd"
''Public Const tMS_CN511 = "ยอดรวมผิดพลาด;Transaction Total Mismatch"
'
'Public Const tMS_CN512 = "จำนวนที่ระบุไม่ถูกต้อง;Bad Amount" 'tms_CN005064

''*TW 51-03-17
'Public Const tMS_CN513 = "การสร้าง Trigger มีความผิดพลาด กรุณาติดต่อผู้ควบคุมระบบ;There are error occur on creating trigger, Please contact your administrator." 'tms_CN005119
'
'
'
'
'
'*KT 52-05-16  *TW 51-07-15 Service
Public Const tMS_CN533 = "Unable to run AutoClear Service. Please contact administrator. ! ;ไม่สามารถเปิด Service AutoClear ได้ กรุณาติดต่อผู้ควบคุมระบบ ! "
Public Const tMS_CN534 = "Scroll item failed.Please contact administrator. ! ;การเลื่อนรายการสินค้าเกิดผิดพลาด กรุณาติดต่อผู้ควบคุมระบบ ! "
Public Const tMS_CN535 = "Please check file AdaIni.ada or contract your Administrator.;กรุณาตรวจสอบไฟล์ AdaIni.ada หรือติดต่อ ผู้ควบคุมระบบ"

'Public Const tMS_CN550 = "Update header  unsuccessful. Please transaction again;Update HD ไม่สำเร็จขอให้ทำรายการใหม่อีกครั้ง"

'*KT 54-11-10 V.4.0.3
Public Const tCS_MSG59 = "เครื่องพิมพ์ไม่สมบูรณ์ กรุณาติดต่อผู้ดูแลระบบ;Printer is not working. Contract Administrator."

'*KT 54-11-15 V.4.0.4
Public Const tCS_MSG60 = "การรับชำระผิดพลาด;Payment errors."
Public Const tMS_CN518 = "ระบบไม่พบวัตถุต่อไปนี้;Object required"
'*BG 56-10-28
Public Const tMS_CN519 = "NET Terminal not reponding !!" & vbCrLf & "Please check NETs Terminal"
Public Const tMS_CN520 = "Invalid Credit type ! "
'*BG 57-04-18 PRJ10001TK-PH2-External Spec-MaxQTY9999-00.02.00.docx
Public Const tMS_CN521 = "The max Qty is 9999."
Public Const tMS_CN522 = "SLIP print is not allowed for this transaction"
Public Const tMS_CN523 = "Hot Key allow for Item Entry Screen only."
