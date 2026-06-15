Attribute VB_Name = "mCNVB"
Option Explicit
'MAIN
Public ocnAll As New ADODB.Connection       '*Tee 51-02-19
Public ocnCfg As ADODB.Connection
Public ocnMember As ADODB.Connection
Public tVB_CNSqlLastCmd As String       '*Tee 51-02-19
Public oDB As New cCNDB       '*Tee 51-02-19
Public oTransPrn As New cTransection
Public oPrn As New cPrinter
'Public bVB_CNCurOnline As Boolean       '*Tee 51-02-19
'*BG 56-06-06 TK-ISS3000-078
Public nC_Line As Long
Public bVB_LastLine As Boolean
Public nVB_Actual As Long
Public bVB_OverShort As Boolean
'*****

'Public oRe As New cRecovery
Public oTerminal  As New cTerminal
Public oSKU As New cSKU
'Public oMSR As New MSRLib.MSR

'Public tVB_Option As String
'Public aVB_Pdt() As String
'Public aVB_Suspend() As String
'Public nVB_MaxPdt As Long
'Public oPing  As Object

'Public cSubDisc As Double

'Public ocnAll As ADODB.Connection
'Public ocnOffLine As ADODB.Connection '*KT 51-04-29 ไปใช้ oDB.ocnOffLine
Public ocnEJ As ADODB.Connection
Public orsAll As ADODB.Recordset
Public orsPmt As ADODB.Recordset
Public nVB_CNBase As eEN_DBMS                 'Access or SQL Sever

'Public bVB_CNAppDemo As Boolean
'Public bVB_CNAppDemoFull As Boolean
Public bVB_CNPrgFsh As Boolean
'Public nVB_CNResolution As Long
'Public oVB_CNBrwRet As Collection

'general
Public bVB_CNShwThai As Boolean

'for color
Public nVB_CNClrBOtbCal As Long
Public nVB_CNClrBFrmMain As Long
Public nVB_CNClrBOlaTag4 As Long
Public nVB_CNClrBOlaTag5 As Long
Public nVB_CNClrBLock As Long
Public nVB_CNClrBUnLock As Long
'STRING
'special
Public tVB_CNAppPath As String
Public tVB_CNPwdAdaUsrP As String             'Ada user password login
Public tVB_CNUserCode As String               'user code from TSysUser
Public tVB_CNUserName As String              'name of user login
'general
Public tVB_CNBchCode As String                  'company branching
'Public cVB_CNCmpVatRate As Double
'Public bVB_CNCmpRetVatEx As Boolean
Public wMain As Form
Public wSub As Form
Public wDisc As Form
'Public nVB_CNDbMode As eEN_DbMode
'Public bVB_CNDbOnLine As Boolean
'Public bVB_CNDbOffLine As Boolean

'Public nVB_CNDocMode As eEN_DocMode

'Public tVB_CNPthEJ As String
'Public tVB_CNOnPthEJ As String
'Public tVB_CNOffPthEJ As String
'Public bVB_CNEjTypePIC As Boolean
'Public nVB_CNLenCodeScale As Integer     '*Tee 48-10-13
'Public bVB_CNAlwDisPmt As Boolean       '*Tee 48-05-30
'Public bVB_CNAlwRepack As Boolean       '*Tee 48-05-30
'Public bVB_CNAlwUpdQtyNow As Boolean       '*Tee 49-02-08
'Public bVB_CNUsedBankNoet As Boolean    '*Tee 48-08-17
'Public tVB_CNCalPmtOne As String                    '*WD 48-07-08
'Public bVB_CNSumRepack As Boolean       '*WD 48-07-11
'Public bVB_CNAlwCalVatAfCpn As Boolean       '*Tee 48-09-29 คำนวน ภาษีหลังใช้ คูปอง
Public bVB_CNEdcPos As Boolean      '*Tee 48-10-05 CreditCard
'Public bVB_CNMsrPos As Boolean      'Magnetic strip reader
Public bVB_CNNETSPos As Boolean
'Public tVB_CNEdcPos As String      'CreditCard
'Public bVB_CNAlwAviDC As Boolean      '*Tee 49-04-05 Available DC
'Public bVB_CNPmtCalTheBest As Boolean      '*Tee 49-09-19 การให้ส่วนลดโปรโมชั่น
'Public bVB_CNAlwChkSaleStartStop As Boolean      '*Tee 49-09-20 การเช็คอณุญาตขายตามวันที่ให้ส่วนลดโปรโมชั่น
'Public bVB_CNDocPmt As Boolean      '*Tee 49-09-25 ตรวจสอบว่าบิลมี Pmt หรือ ไม่
Public bVB_CNDrwPos As Boolean      'Drawer
Public bVB_CNDisplayPos As Boolean      'Customer Display
Public bVB_CNEJPos As Boolean      'EJ
Public bVB_CNPrnOPOS  As Boolean      'EJ

'========================================
'Public cUT As New AdaUT.cUT
'Public cUT_EDC As Object    '*Gon 50-06-25
Public cUT_Print As Object   '*Gon 50-06-21
Public cUT As Object        'for AdaUT dll only
'Public oPicDll As New AdaScreen800.Pic30FshHpm
'Public oPicDll As New AdaScreen1024.Pic30FshHpm
Public oPicDll As Object        'SM 48-09-16

Public tVB_CNComName As String
Public tVB_CNAppName As String
Public tVB_CNAppPrg As String               'program name
Public tVB_CNAppTitle As String
'Public tVB_CNDptCode As String             'department code, by user
'Public tVB_CNDptName As String            'department name

'*GP 46-10-09
Public tVB_CNDbCurFile As String           'db current file (Access: path+ "\" + file, SQL:file)
Public tVB_CNOnDbCurFile As String           'db current file (Access: path+ "\" + file, SQL:file)
Public tVB_CNOffDbCurFile As String           'db current file (Access: path+ "\" + file, SQL:file)
Public tVB_CNDbSrc As String                  'db file source(SQL)
Public tVB_CNOnDbSrc As String                  'db file source(SQL)
Public tVB_CNOffDbSrc As String                  'db file source(SQL)
Public tVB_CNDbPwd As String
Public tVB_CNDbLoginType As String
Public tVB_CNDbSQLAuthUsrC As String
Public tVB_CNOnDbSQLAuthUsrC As String
Public tVB_CNOffDbSQLAuthUsrC As String
Public tVB_CNDbSQLAuthUsrP As String
Public tVB_CNOnDbSQLAuthUsrP As String
Public tVB_CNOffDbSQLAuthUsrP As String
Public tVB_CNDbRptPath As String          'keep path of report (SQL)
Public tVB_CNDbAdaIni As String     '*GP 46-09-24

'*KT 51-04-08
Public tVB_CNOnMemDbSrc As String
Public tVB_CNOnMemDbSQLAuthUsrC As String
Public tVB_CNOnMemDbSQLAuthUsrP As String
Public tVB_CNOnMemDbCurFile As String

Public tVB_CNOffMemDbSrc As String
Public tVB_CNOffMemDbSQLAuthUsrC As String
Public tVB_CNOffMemDbSQLAuthUsrP As String
Public tVB_CNOffMemDbCurFile As String

'*Duang 51-03-17
Public tVB_CNDbSQLAuthSecondC As String    'User Name  DBSecond
Public tVB_CNDbSQLAuthSecondP As String    'User Password DBSecond
Public tVB_CNDbSecondSrc As String    'Server Name DBSecond
Public tVB_CNDbSecondCurFile As String    'Database Name DBSecond
Public tVB_CNDbLoginTypeMain As String    'login type Windows Authentic,SQL Authentic
Public tVB_CNDbLoginTypeSecond As String
Public bVB_CNDbSQLAuthSrvDepartment As Boolean    'user pswd SQL authentic+
'Public tVB_CNTermNo As String           'Terminal Number
'Public nVB_CNLastBillS As Double        '*GP 46-09-23

'*KT 51-04-08

'*Tee 49-03-15
Public tVB_CNDbOffLineCurFile As String           'db current file (Access: path+ "\" + file, SQL:file)
Public tVB_CNDbOffLineSrc As String                  'db file source(SQL)
Public tVB_CNDbOffLinePwd As String
Public tVB_CNDbOffLineLoginType As String
Public tVB_CNDbOffLineSQLAuthUsrC As String
Public tVB_CNDbOffLineSQLAuthUsrP As String
Public tVB_CNDbOffLineRptPath As String          'keep path of report (SQL)

'Public tVB_CNDbLastConType As String
'Public tVB_CNDbLastUser As String

'Public tVB_CNWahCode As String                'รหัสคลังขาย    46-06-19
Public tVB_CNRetVATInOrEx As String

'Public tVB_CNBrwTable As String
Public tVB_CNBrwField As String
Public tVB_CNBrwHead As String
'Public tVB_CNBrwSome As String
'Public tVB_CNBrwRet As String
'Public tVB_CNBrwWhe As String
'Public tVB_CNBrwCaption As String
'Public tVB_CNBrwFltLast As String
'Public tVB_CNBrwOrder As String
'Public tVB_CNBrwRetSpc As String
'Public bVB_CNBrwFltImmed As Boolean
'Public tVB_CNBrwFilterCond As String
'Public nVB_CNPmtPdtBarType As Integer       '*Tee 49-10-26
Public nVB_CNDecAmtForShw As Long                           'จำนวนทศนิยมสำหรับแสดง
Public nVB_CNDecAmtForSav As Long                           'จำนวนทศนิยมสำหรับบันทึก        '*Tee 48-10-15
'Public nVB_CNThemeStyle As Integer      '*Tee 49-02-21
Public nVB_CNBrwCol As Long
Public nVB_CNCutMsg As Integer               'cut message Tha=1, Eng=2
Public nVB_CNCutTag As Integer               'cut tag
Public nVB_CNCutConst As Integer            'cut constant
'Public nVB_CNBrwCstLev As Integer
'Public nVB_CNPdtPriLevWhs As Integer
'Public nVB_CNPdtPriLevRet As Integer
'Public nVB_CNCstPriLev As Integer
'Public nVB_CNPriLev As Integer
Public nVB_CNLastLogin As Long      '*Tee 49-03-22

'Public nVB_CNLastDocS As Long        '*GP 46-09-25
'Public nVB_CNLastDocR As Long
'Public nVB_CNLastDocN As Long        '*Tee 46-12-22
'Public nVB_CNLastDocU As Long        '*Tee 46-12-22
'Public nVB_CNLastDocI As Long        '*Tee 49-03-04
'Public nVB_CNLastDocJ As Long        '*Tee 49-03-04
'Public nVB_CNLastDocX As Long        '*Tee 49-03-04
'Public nVB_CNLastDocY As Long        '*Tee 49-03-04
'Public nVB_CNLastDocV As Long        '*Tee 49-03-04
'Public nVB_CNStartDocV As Long        '*Tee 49-10-05

Public nVB_CNTimeLock As Long        '*Tee 49-03-18

'Public tVB_CNLastSaleDate As String        '*Tee 47-01-05
'Public bVB_CNPosVatSend As Boolean      '*Tee 46-12-22
'Public bVB_CNBillVatSend As Boolean      '*Tee 46-12-22
'Public cVB_CNCstPntBuy As Double '*Tee 46-10-25
'Public cVB_CNCstPntGet As Double '*Tee 46-10-25

'browser position
Public nVB_CNBrwPosGrdRow As Long
Public nVB_CNBrwPosGrdCol     As Long
Public nVB_CNBrwPosCurX      As Long
Public nVB_CNBrwPosCurY      As Long
Public nVB_CNBrwPosBtnW  As Long
Public nVB_CNBrwPosBtnH  As Long
Public nVB_CNBrwPosParentCellW  As Long
Public nVB_CNBrwPosParentCellH  As Long
Public nVB_CNBrwTop As Long     '*GP 47-09-09

'สินค้าชุด
'Public tVB_CNPdtSetSta As String        'สถานะ สินค้า 1:ทั่วไป, 2:เป็นลูก, 3:เป็นพ่อ
'Public tVB_CNPdtSetParent As String     'รหัสสินค้าพ่อ (สินค้าชุด)
'Public nVB_CNPdtSetLev As String        'ระดับสินค้า (สินค้าชุด)
'Public cVB_CNPdtSetQty As String        'จำนวน สินค้าเป็นชุด
'Public bVB_CNChkCostZero As Boolean     '*GP 47-08-13
Public tVB_CNBrwPdtField As String                  'browser product        '*Tee 47-12-03
Public tVB_CNBrwPdtHeadTha As String
Public tVB_CNBrwPdtHeadEng As String
Public tVB_CNRptMasSpec As String
Public tVB_CNRptMasRange As String
Public bVB_CNPwdAllow As Boolean
'Public tVB_CNLastBillNo As String
Public tVB_CNBrwColHide As String
'numeric
Public nVB_CNSysDefLCID As Long
Public nVB_CNRetValue As Long
'text
Public tVB_CNRetValue As String
'object
Public oThisCombo As ComboBox
'*WD 48-06-27
Public nVB_CNBaseEJ As Integer
Public tVB_CNDbSrcEJ As String                  'db file source(SQL)
Public tVB_CNOnDbSrcEJ As String                  'db file source(SQL)
Public tVB_CNOffDbSrcEJ As String                  'db file source(SQL)

Public tVB_CNDbSQLAuthUsrCEJ As String
Public tVB_CNDbSQLAuthUsrPEJ As String
Public tVB_CNDbCurFileEJ As String           'db current file (Access: path+ "\" + file, SQL:file)

'*TW 50-02-08
'Public nVB_CNQtySalBillYear As Long 'จำนวนบิลในตาราง Sale ในรอบปี
'Public nVB_CNQtySalBillMonth As Long 'จำนวนบิลในตาราง Sale ในรอบเดือน
'Public nVB_CNQtySalBillDay As Long 'จำนวนบิลในตาราง Sale ในรอบวัน

'*TW 50-02-12
'Public nVB_CNPmtCpnQty As Integer 'จำนวนคูปองที่จะได้รับท้ายบิล
'Public cVB_CNPmtCpnAmt As Double 'มูลค่าคูปองต่อ 1 ใบ
'Public dVB_CNPmtCpnStart As Date 'วันที่เริ่มใช้คูปองได้
'Public dVB_CNPmtCpnStop As Date 'วันที่สิ้นสุดการใช้คูปอง
'Public tVB_CNPmtCpnCode As String 'รหัสคูปองโปรโมชั่น
'Public nVB_CNCpnSeq As Integer 'ลำดับของคูปอง ใช้ในการพิมพ์คูปองลง EJ
'Public nVB_CNCpnMaxQty As Integer 'จำนวนสูงสุดที่อนุญาตให้แจกคูปอง
'Public bVB_CNPntCpn As Boolean 'สถานะการพิมพ์คูปอง

'*TW 50-02-17
'Public dVB_CNLastOnlineDate As Date 'วันที่ล่าสุดที่ Olline
'Public nVB_CNStartBillDay As Long 'บิลแรกของวันที่ Online
'Public nVB_CNStartBillMonth As Long 'บิลแรกของเดือนล่าสุด Online
'Public nVB_CNStartBillYear As Long 'บิลแรกของปีล่าสุด Online

'*TW 50-02-21
'Public tVB_CNPlusSaleWording As String 'เก็บค่าที่จะแสดงส่งเสริมการขายแบบบิล
'Public cVB_CNPlusSalePer As Double 'เป็นเซ็นต์ในการคำนวณเรื่องส่งเสริมการขาย

'*TW 50-03-07
'Public nVB_CNShiftSumCopy As Integer 'จำนวนพิมพ์ใบสำเนาปิดกะ
'Public bVB_CNIsPrnCopy As Boolean 'เป็นการพิมพ์สำเนาหรือไม่
'Public nVB_CNShiftSumCopyNow As Integer 'ลำดับใบสำเนาปิดกะ

'*TW 50-03-15
'Promotion 92
'Public nVB_CNPmtTypeOrder As Integer 'เก็บประเภทการลดโปรโมชั่นซ้อนในรูปแบบตัวเลขฐาน 10 ซึ่งค่ามาจากตัวเลขฐาน 2 ที่เอาข้อมูล True & False ของตาราง Cross ประเภทโปรโมชั่นมาคิด เช่น
'                                                                                   101 ฐาน 2 เท่ากับ 5 ฐาน 10 หมายความว่า ถ้าเกิดโปรโมชั่นประเภท 1 แล้ว จะต้องไม่เกิดโปรโมชั่นประเภท 2 แต่เกิดโปรโมชั่น 3 ได้
'                                                                                   110 ฐาน 2 เท่ากับ 6 ฐาน 10 หมายความว่า ถ้าเกิดโปรโมชั่นประเภท 1 แล้ว จะต้องไม่เกิดโปรโมชั่นประเภท 3 แต่เกิดโปรโมชั่น 2 ได้
'                                                                                   111 ฐาน 2 เท่ากับ 7 ฐาน 10 หมายความว่า ถ้าเกิดโปรโมชั่นประเภท 1 แล้ว สามารถเกิดโปรโมชั่นประเภท 2 และเกิดโปรโมชั่น 3 ได้
'Public tVB_CNPmt2DocNo As String
'Public cVB_CNPmt2CpnQty As Double
'Public cVB_CNPmt2CpnAmt As Double
'Public nVB_CNPmt2CpnMaxQty As Integer
'Public tVB_CNPmt2Type As String

'Public aVB_CNPmt2PdtName() As String
'Public aVB_CNPmt2CpnGetQty() As Double
'Public aVB_CNPmt2CpnAmt() As Double
'Public bVB_CNPmt2CpnActive As Boolean

'promotion 93
'Public aVB_CNPmt93PdtName() As String 'ชื่อสินค้าที่ได้รับโปรโมชั่น
'Public aVB_CNPmt93DisType() As Boolean 'ลดบาท หรือลดเปอร์เซ็นต์
Public aVB_CNPmt93Dis() As Double 'จำนวนลด
Public aVB_CNPmt93DisAmt() As Double 'มูลค่ารวมที่ลด
Public bVB_CNPmt93CpnActive As Boolean

'*TW 50-03-22 จำนวนการนำเข้าและเบิกออกเงิน
'Public nVB_CNNoOfMnyInOut As Integer


'*Gon 50-06-26
'Public tVB_TypeOfEDC As String      '1: Hyper com 2: PosNet

''****** register ******
'Public tVB_CNRegModule As String                          'for register
'Public tVB_CNRegCompLen As String
'Public tVB_CNRegCompSum As String
'Public tVB_CNRegDateExp As String
'Public tVB_CNRegSerial As String
'Public tVB_CNRegLicBack As String
'Public tVB_CNRegLicBackChk As String
'Public tVB_CNRegLicFront As String
'Public tVB_CNRegLicFrontChk As String
'Public tVB_CNRegLanAlone As String   'alone=1 lan=2
'Public tVB_CNRegValue As String

'*Pla 50-09-03  Form Mall Card
'Public tVB_Caption As String
'Public nVB_ColSelect As String
'Public tVB_Filter As String
'Public tVB_MallCardID As String
'Public tVB_IDCard As String
'Public tVB_Tel As String
'Public tVB_TranType As String           ' เก็บค่าประเภท Voucher


'*Em 50-10-22
'Public tVB_CNAppVer As String
'Public tVB_CNEvnCode  As String
'Public tVB_CNEvnDesc  As String
'Public tVB_CNEvnNew As String
'Public tVB_CNEvnOld As String
Public tVB_CNPthFileINI As String
'Public tVB_Form As String

Public nVB_CNCutLang As Long

'Public bVB_CNPrcAll As Boolean
'Public bVB_CNPrcDone As Boolean
'Public bVB_CNPrcUsedAvg As Boolean

'Public nVB_CNVersion As eEN_Version

'T.Saisawaat 05-Nov-2007
'Public bVB_CNTrainingMode As Boolean

'*New

Public nVB_RTD As Byte

Public bVB_PromptSale As Boolean

Public oDisc As New cPSDisc
Public bVB_CNOpenTmn As Boolean
Public bVB_CNOpenTmnOff As Boolean 'เก็บสถานการ Open ของ Store
Public bVB_TrainingMode As Boolean
Public bVB_ReEntryMode As Boolean
Public tVB_LEmpID  As String
Public tVB_LEmpPW  As String
Public tVB_LRunning  As String
Public tVB_LSaleDate  As String '*Nirut 08-03-26 add
Public tVB_CNTblHD As String
Public tVB_CNTblDT As String
Public tVB_CNTblDTTmp As String
Public tVB_CNTblCDTmp As String
Public tVB_CNTblRC As String
Public tVB_CNTblCD As String
Public bMainFunc As Boolean
'Public tVB_CNTblDTOntop As String
'Public tVB_CNTblCDOntop As String


Public tVB_CNMnuCurrent As String             'keep last click menu for check user right
Public bVb_SetResorution As Boolean

'Permision User
Public tVB_CNUserLevel As String          'user level
Public tVB_CNUserPwd As String             'user password login
Public tVB_CNUserAlwC As String            'allowed user code
Public tVB_CNUserAlwN As String            'allowed user name
Public dVB_CNUserAlwDate As Date            'allowed Date
Public tVB_CNUserAlwTime As String            'allowed Time
Public bVB_CNUserOveride  As Boolean            'allowed Time

Public tVB_CNIP As String
Public tVB_CNSrvIP As String
'Text Promt
Public bVB_SetFocus As Boolean

Public Type FUNMNU
    FNSpmSeq  As String
    FTSpmID As String
    FTSpmNameUser As String
End Type
Public aVB_FunMnu() As FUNMNU

'Transaction
Public Type TRANSMODE
    ID  As String
    Name As String
    NameABB As String
    PrintType As String
End Type
Public aVB_TransMode() As TRANSMODE

Public Type MAINMNU
    ID  As String
    Seq As String
    Name As String
    DefLangPri As String
    DefLangSec As String
    UsrLangPri As String
    UsrLangSec As String
    FuncName As String
    NameUsr As String
    SmnID As String
    Show As Boolean
    Use  As Boolean
    Autho As Boolean
    SeqNo As String
End Type
Public aVB_MainMnu() As MAINMNU
Public aVB_MainMnuAll() As MAINMNU
Public aVB_TransMnu() As MAINMNU
Public aVB_TransMnuAll() As MAINMNU
Public aVB_DiscMnu() As MAINMNU

Public Type TENDERMENU
    ID  As String
    Name As String
    DefLangPri As String
    DefLangSec As String
    UsrLangPri As String
    UsrLangSec As String
    DescSale As String
    DescReturn As String
    SmnID As String
    Show As Boolean
    Use  As Boolean
    Autho As Boolean
    StaSale As Boolean
    StaChange As Boolean
    StaReturn As Boolean
    StaMallCard As Boolean
    Menu As String '*KT 53-10-15 Commsheet หมายเลข PH3-ML-SF-131
    StaPoint As Boolean '*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
    StaEmpDisc As Boolean
    StaBPSale As Boolean
    StaBPReturn As Boolean
    StaExpDate As Boolean
    StaReEntry  As Boolean
    
End Type
Public aVB_TenderMnu() As TENDERMENU

Public Type PRNNAMEABB
    ID  As String
    Name As String
    NameUsr As String
    DefLangPri As String
    DefLangSec As String
    UsrLangPri As String
    UsrLangSec As String
    SmnID As String
End Type
Public aVB_NameAbb() As PRNNAMEABB

Public Type MALLCRD
    ID As String
    IDLastPoint  As String
    Type As String
    Point As Double
    LPoint As Double
    Amt As Double
    CardID As String
    tEvnCode As String
    tFTMcdCardType As String '*Bmp 61-02-15 Phase Token
    tCardIDToken As String '*Bmp 61-02-15 หมายเลขบัตรที่ Token แล้ว
End Type

Public Type TAX
    TAXCODE As String
    TAXRATE As Double
    TAXAMT As Double
End Type

Public Type MALLPRORATE
    MALLRATE As Double
    MALLAMT As Double
End Type

Public Type CardType
    IDCard  As String
    Amt As Double
End Type

Public aVB_MallCard() As MALLCRD
Public aVB_MallCardReg() As MALLCRD

Public nVB_CurentMainMnu  As Long
Public nVB_CurentTrnMnu  As Long

Public tVB_Transaction As String
Public tVB_TransName  As String
Public tVB_TransNameABB  As String
Public tVB_TransactionSub  As String
Public tVB_TransactionMode  As String '*KT 52-03-10 เก็บ Flag ในการทำ Open แบบ

Public tVB_OldTransaction As String
Public tVB_OldTerminalNum  As String
Public tVB_OldRunning As String
Public tVB_OldTypePrinter  As String

'Transaction No
Public tVB_Running As String
'AutoGen
Public tVB_AUTTblName As String
Public tVB_AUTFedCode As String            'Field
Public tVB_AUTFedType As String                 'ประเภทของเอกสาร เฉพาะตารางที่มีการใช้งานมากกว่า 1 รูปแบบ  เช่น รหัสเอกสารสินค้า ,รหัสเอกสารบริการ
Public tVB_AUTFedTypeName As String            'ชื่อ Field ประเภทของเอกสาร เฉพาะตารางที่มีการใช้งานมากกว่า 1 รูปแบบ  เช่น รหัสเอกสารสินค้า ,รหัสเอกสารบริการ
Public nVB_AUTFedSize As Long            'ขนาดของ Field รหัสเอกสาร
Public tVB_AUTStaCallNew As String         'สถานะการเรียกปุ่ม add new ถายหลังการบันทึก
Public tVB_AUTStaResetBill As String         'สถานะการกำหนดการเปลี่ยนเลขที่ เอกสาร 0 ต่อเนื่อง  1 เปลี่ยนรอบปี
Public tVB_AUTChar As String                   'อักขระนำ
Public tVB_AUTYear As String
Public tVB_AUTMonth As String
Public tVB_AUTDay As String
Public tVB_AUTNum As String
Public tVB_AUTCodeFormat As String
Public tVB_AUTChkBch As String
'Terminal
Public tVB_CNTerminalIndex  As String 'ลำดับเครื่อง
Public tVB_CNTerminalFloor  As String 'ลำดับเครื่อง
Public tVB_CNTerminalNum  As String 'Terminal Number
Public tVB_CNTerminalDrv  As String 'ชื่อ Driver
Public tVB_CNTerminalReg As String 'หมายเลขทะเบียนเครื่อง Terminal
Public tVB_CNPhysicalTerminalNum  As String 'Physical Terminal Number

'Online/Offline
Public tVB_CNOnCaption As String
Public tVB_CNOffCaption As String

'Text Promt
Public bW_LockKeyAscii As Boolean
Public bW_Ok As Boolean
Public nW_TimerLock As Long


Public tVB_CNStoreCode As String                'รหัส Store   46-06-19
Public tVB_CNStoreNameAbb As String                'ชื่อ Store ที่แสดงบนใบ Ticket
Public tVB_CNStoreTel As String ' เบอร์โทร Store
Public tVB_CNStoreTaxID As String
Public tVB_CNStoreStaClose As String

Public tVB_CNEventPrice As String
Public tVB_CNDpcCode  As String
Public tVB_CNSkuDepCode As String
Public tVB_CNSkuBrdCode As String
Public tVB_CNSkuStaAlw As String 'เก็บสถานการ Allow ของสินค้าต่าง เป็น Degit โดย 0000000000000000000
'Degit 1 : Disc Line Item,Degit 2 : Disc Trans,Degit 3 : Disc Line Mkd,Degit 4 : Disc Mkd Trans,Degit 5 : Coop Disc,
'Degit 6 : Coop Cpn,Degit 7 : Coop Brd,Degit 8 : Mall Cpn,Degit 9 : Mall/Emp ,Degit 10 : Auto Trans,Degit 11 : On Top,Degit 12 : Free Items,13 :,14: Sku Point Allow Discount,15:Disc Mkd Trans1,16: Disc Mkd Trans2
'.17: Disc Mkd Trans3,18:Disc Mkd Trans4,19:Disc Mkd Trans5
'Alw
'Rouding
Public bVB_CNAlwRnd  As Boolean 'ตัวแปรให้มีการทำ Rouding
Public nVB_CNTypeRnd  As Byte 'ตัวแปรประเภท Rouding
'Sex,Nat,Saleman
'Public bVB_CNAlwInputCstVIP As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการจะใส่ข้อมูล Customer VIP หรือไม่
Public bVB_CNAlwInputSex As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการจะใส่ข้อมูล Sex Customer หรือไม่
Public bVB_CNAlwInputNat As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการจะใส่ข้อมูล National Customer หรือไม่
Public bVB_CNAlwInputAge  As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการจะใส่ข้อมูลช่วง Age Customer หรือไม่
Public bVB_CNAlwInputSpn  As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการจะใส่ข้อมูล Saleman หรือไม่
'Prn Point
Public bVB_CNAlwPrnLPoint  As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการพิมพ์จำนวนเต้มล่าสุด
Public bVB_CNAlwPrnToDayPoint  As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการพิมพ์จำนวนเต้มรวมก่อนหน้า
'Prn Barcode
Public bVB_CNAlwPrnBarTicket  As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการพิมพ์Barcode ของ Ticket
'Double Login Policy
Public bVB_CNAlwDLogInPolicy As Boolean 'ตัวแปรเก็บว่ามีการเก็บว่าต้องการทำงาน Double Login Policy
Public tVB_CNLevelDLogInPolicy  As String 'ตัวแปรเก็บว่ามีการเก็บว่า Level ของ Employee ที่ Double Login Policy
'Sale Name Journal
Public bVB_CNAlwDepJournal As Boolean 'ตัวแปรเก็บว่ามีการเก็บรหัสแผนกของฐานข้อมูล เช่น D8 : Depart บางกะปิ ,S8 : Super บางกะปิ ,TS : take home สยามพารากอน
Public tVB_CNDepJournalID   As String 'ตัวแปรเก็บว่ามีการเก็บ รหัสแผนกของฐานข้อมูล เช่น D8 : Depart บางกะปิ ,S8 : Super บางกะปิ ,TS : take home สยามพารากอน
'Config Print SKU Name
Public bVB_CNAlwPrnSKUName  As Boolean 'ตัวแปรเก็บว่ามีการเก็บรหัสแผนกของฐานข้อมูล เช่น D8 : Depart บางกะปิ ,S8 : Super บางกะปิ ,TS : take home สยามพารากอน
Public tVB_CNPrnSKUName   As String 'ตัวแปรเก็บว่ามีการเก็บ รหัสแผนกของฐานข้อมูล เช่น D8 : Depart บางกะปิ ,S8 : Super บางกะปิ ,TS : take home สยามพารากอน
'Config Check ExpireDate from OperationDate
Public bVB_CNAlwChkExpDate As Boolean 'ตัวแปรเก็บว่ามีการตรวจสอบวันที่หมดอายุจากวันที่ operation date หรือ syatem date มีผลต่อ mixningthsale
'Config Check ExpireDate Point from OperationDate
Public bVB_CNAlwChkExpDatePoint As Boolean 'ตัวแปรเก็บว่ามีการตรวจสอบวันที่หมดอายุจากวันที่ operation date หรือ syatem date มีผลต่อ mixningthsale

Public bVB_CNAlwBeOrAftDisc As Boolean
Public tVB_CNDiscPrcAftDisc   As String

Public bVB_CNCalAutoTran As Boolean '*KT 51-05-29 ML-SF-275

'*Screen 1024*768
Public nVB_OldPositionScreen As Long 'เก็บตำแหน่งของ Resolution ก่อนการเข้าโปรแกรม
Public nVB_ScreenDefault  As Boolean 'เก็บว่า Resolution ก่อนเข้าโปรแกรมเป้นแบบที่ต้องการหรือไม่
Public nVB_Resorution As Integer

'OPOS
Public tVB_CashDrawer As String 'ชื่อ Driver ของลิ้นชักที่ใช้กับ OPOS
Public tVB_CashPrinter  As String 'ชื่อ Driver ของ Printer ที่ใช้กับ OPOS
Public tVB_CashSlipPrinter  As String 'ชื่อ Driver ของ Printer ที่ใช้กับ OPOS
Public tVB_CashDisplay  As String 'ชื่อ Driver ของ Display ที่ใช้กับ OPOS
Public tVB_CashMsr  As String 'ชื่อ Driver ของ Msr  ที่ใช้กับ OPOS
Public tVB_Scanner  As String 'ชื่อ Driver ของ Msr  ที่ใช้กับ OPOS

Public aVB_TSlipMag() As String 'array เก็บ ข้อความท้าย Ticket
Public aVB_HSlipMag() As String 'array เก็บ ข้อความหัว Ticket
Public aVB_DSlipMag() As String 'array เก็บ ข้อความหัว Ticket
Public nVB_CashCodePage As Integer 'Code page ของ OPOS
Public tVB_PathImageSlp As String
Public nVB_MaxChar  As Integer 'Code page ของ OPOS
Public nVB_MaxCharSlip  As Integer 'Code page ของ OPOS

'Payment
Public cVB_Total As Double 'ยอดเงินทั้งหมดก่อนหักส่วนลดที่เกิดจากหน้าจอการระบุ Disccout หลังจากการ Sub จากหน้าจอ wMain
Public cVB_TotalDisc As Double 'ยอดรวมส่วนลด
Public cVB_TotalDiscB As Double 'ยอดรวมส่วนลดก่อนรายตัว Deal,Promo,Promo Time,Coo Disc,Line Disc,Line Mark
Public cVB_TotalDiscAft As Double 'ยอดรวมส่วนลดที่เกิดจากหน้าจอการระบุ Disccout หลังจากการ Sub จากหน้าจอ wMain
Public cVB_TotalDiscMallCrd As Double 'ยอดรวมส่วนลดที่เกิดจากการใช้ Mall Card
Public cVB_TotalDiscEmpCrd As Double ''ยอดรวมส่วนลดที่เกิดจากการใช้ Emp Card
Public cVB_TotalDiscOnTop As Double 'ยอดรวมส่วนลดที่เกิดจากการใช้ On Top
Public cVB_NetSale  As Double 'ยอดรวมสุทธิหลังหักส่วนลด
Public cVB_RoundRcv As Double 'ยอดปัดเศษ
Public cVB_RoundChg As Double 'ยอดปัดเศษของการทอน
'Public cVB_RoundRcvDps As Double 'ยอดปัดเศษ
'Public cVB_RoundDiffAmt As Double 'เศษจากยอดรวมก่อนปัดเศษ
Public cVB_NetSaleRound  As Double 'ยอดรวมสุทธิหลังหักส่วนลด+เศษ(Round Diff)
Public cVB_Balance As Double 'ยอดคงเหลือที่มีการรวม Roundiff แล้ว
Public cVB_NetDutyFreeDep  As Double 'ยอดรวมสุทธิมัดจำในการขาย Duty Free
Public tVB_TenderName As String 'ชื่อ Tender
Public tVB_TenderNameCreditRE As String 'ชื่อ Tender Bump 57-03-05 สำหรับเอาไว้ตรวจสอบ
Public tVB_TenderCode As String 'รหัส Tender
Public tVB_TenderType As String 'ประเภท Tender
Public cVB_Change As Double 'ยอดเงินทอน
Public nVB_CurTenderSeqNo As Double 'ลำดับของ Tender ปัจจุบันที่เลือกรับชำระ
Public cVB_SpcGiftVoucherAmtAffect   As Double 'ยอดที่มีการบวกขึ้นมากรณีใช้ Spacial gift Voucher กับ Mall Card
Public cVB_SpcCpnAmtAffect   As Double 'ยอดที่มีการบวกขึ้นมากรณีใช้ Spacial Coupon กับ Mall Card
Public cVB_CrdCardAmtAffect   As Double 'ยอดที่มีการบวกขึ้นมากรณีใช้ Credit Card กับ Mall Card
Public cVB_RegTotAmtAfDisc As Double 'ยอดรวมสินค้าหลังหัดส่วนลดทุกชนิดเฉพาะสินค้าป้ายน้ำเงิน
Public cVB_ProTotAmtAfDisc As Double 'ยอดรวมสินค้าหลังหัดส่วนลดทุกชนิดเฉพาะสินค้าป้ายเหลือง

'Line Display
Public bVB_OPOSLineDisplayEnable As Boolean

'CashDrawer
Public bVB_OPOSCashDrawerEnable As Boolean

'Scanner
Public bVB_OPOSScannerEnable As Boolean

'MSR
Public bVB_OPOSMSREnable As Boolean

'Deposit
Public cVB_BalAftDisc  As Double 'ยอดเงินที่เหลือหลังหักส่วนลดท้ายบิล

Public Enum eEN_SLPMSG
    PrintSlpHead = 0
    PrintSlpTrail = 1
    PrintSlpDisplay = 2
    PrintSlpDetail = 3
End Enum

'weight
Public cVB_Weight As Double
Public cVB_Price  As Double
Public tVB_Weight As String
Public tVB_Price  As String

Public nVB_TypePay As Byte '0: persent 1:amount 2: new price
Public cVB_Pay As Double
Public tVB_Total As String
Public tVB_FileGrid As String
Public tVB_Plu As String
Public tVB_MallCardVIP As String

'Return
Public bVB_TrnReturn  As Boolean
Public tVB_TerminalNoRet  As String
Public tVB_TransNoRet  As String
Public tVB_TransTypeRet  As String
Public dVB_TransDateRet  As Date

Public bVB_Return As Boolean
Public bVB_Reference As Boolean
'*KT 66-08-17 TTK-ISS3000-839
Public bVB_ReturnAll As Boolean
Public bVB_ReturnAllBill As Boolean 'คืนทั้งบิลในครั้งเดียว
'Public bVB_ReturnFrmBck  As Boolean
'Public bVB_ReturnHaveDisc As Boolean

'IPV
Public bVB_IPV As Boolean
Public tVB_TerminalIPV  As String
Public tVB_TransNoIPV  As String
Public tVB_TransTypeIPV  As String
Public dVB_TransDateIPV   As Date
Public cVB_AmountIPV  As Double
Public cVB_AmountIPVDisc  As Double         '*KT 53-02-26  PH3-ML-SF-089 IPV Dat File don't have TableDT detail เป้นตัวแปรเก็บส่วนลดว่าเอกสารต้นทางที่ IPV มีส่วนลด เพื่อใช้ในขั้นตอนการ Copy Data CD ถ้า cVB_AmountIPVDisc > 0 ต้อง Copy ให้สำเร้จ
'*KT 66-08-17 TTK-ISS3000-839
Public tVB_TerminalSal  As String
Public tVB_TransNoSal As String
Public tVB_TransTypeSal As String
Public tVB_TransFlagSalSta As String

'Price Overide
Public bVB_PriceOveride As Boolean

'Qty Change
Public bVB_FKeyQtyChange  As Boolean 'ใช้สำหรับเป็นตัวแปรบอกว่ามีการ Key ตัวเลขครั้งแรกที่ที่ Column จำนวน เพื่อไม่ใช้ทำการเช้คขนาดการระบุจำนวนสินค้าได้ 6 digit

'Start Bank
Public cVB_CoinsStartBank As Double
Public cVB_CurrencyStartBank As Double
Public cVB_TotalStartBank As Double

'Tax
Public cVB_SKUTaxRate  As Double
Public tVB_SKUTaxCode As String
Public tVB_SKUTaxStaSign As String
Public tVB_SKUTaxSign As String

'Cst
Public tVB_CstID As String
Public tVB_SalemanID As String
Public bVB_SalemanIDHD As Boolean '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
Public tVB_SalemanIDOld As String '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย ตัวแปรที่เก็บหมายเลขเก่าก่อนการ Change One/Change All
Public bVB_SalemanIDChg As Boolean '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย
Public bVB_Saleman As Boolean '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย ตัวแปรที่แยกว่าหน้าจอ wCNSaleman ทำงานจาก Function Ctrl+F2 หรือ Ctrl+F3
Public tVB_SalemanIDChg  As String '*KT 54-05-26 V.4.0.0 Req.15  การเก็บ Sale Man ID ในแต่ละ Item ของรายการสินค้าที่ขาย ตัวแปรที่เก็บหมายเลขจากการ Change One/Change All

'Nagative Scan
Public bVB_NagativeScan As Boolean

'Recall
Public bVB_Recall  As Boolean
Public tVB_TerminalRecall  As String
Public tVB_TransNoRecall  As String
Public tVB_TransTypeRecall  As String

'copy sale
Public tVB_TerminalCpySale As String
Public tVB_TransNoCpySale As String
Public tVB_TransTypeCpySale As String
Public bVB_CpySale As Boolean

'Deposit

Public bVB_TrnDepositRedeem As Boolean

Public bVB_TrnDeposit As Boolean
Public tVB_DepositNo As String
Public cVB_DepositAmt  As Double
Public dVB_DepositDueDate As Date
Public bVB_DepositAll As Boolean 'สถานะการมัดจำทั้งหมด PH1.8-ML-SF-042
Public dVB_CurrentDate As Date

'Redeem
Public bVB_TrnRedeem As Boolean
Public tVB_TerminalNoRedeem As String
Public tVB_TransNoRedeem As String
Public dVB_TransDateRedeem  As Date
Public tVB_TransTypeRedeem  As String
Public tVB_DepossitNoRedeem As String
Public cVB_DepossitAmtRedeem As String
Public cVB_DepossitAmtRedeemRnd  As String
Public cVB_AccBalRedeem  As Double
Public cVB_SunTotalAllRedeem  As Double
Public cVB_SunTotalAllRedeemBDisc  As Double

'Deposit Cancel
Public bVB_TrnDepositCancel As Boolean
Public tVB_TerminalNoDptCancel As String
Public tVB_TransNoDptCancel As String
Public dVB_TransDateDptCancel  As Date
Public tVB_TransTypeDptCancel  As String
Public tVB_DepossitNoDptCancel As String
Public cVB_DepossitAmtDptCancel As Double
Public cVB_DepossitAmtDptCancelRnd  As Double
Public cVB_AccBalDptCancel  As Double
Public cVB_SumTotalAllDptCancel  As Double
Public cVB_SumTotalAllDptCancelBDis  As Double

'Add Deposit
Public bVB_TrnAddDeposit  As Boolean
Public tVB_TerminalNoAddDpt  As String
Public tVB_TransNoAddDpt  As String
Public dVB_TransDateAddDpt   As Date
Public tVB_TransTypeAddDpt   As String
Public tVB_DepossitNoAddDpt As String
Public cVB_DepossitAddDptRnd  As Double
Public cVB_DepossitAmtAddDpt As Double
Public cVB_DepossitAmtAddDptOrg As String 'เก็บค่าดังเดิมของ cVB_DepossitAmtAddDpt ก่อนโดนลบด้วย ontop
Public dVB_DepositDueDateAddDpt  As Date
Public cVB_DepossitAmtBAddDpt As Double
Public cVB_DepossitVatBAddDpt As Double
Public cVB_AccBalAddDpt  As Double
Public cVB_SumTotalAllAddDpt As Double
Public cVB_SumTotalAllAddDptBDisc  As Double
Public dVB_TransDateDueDeposit  As Date

'Voucher
Public bVB_TrnVoucher  As Boolean
Public bVB_AddVoucher  As Boolean
Public tVB_VoucherNo  As String
Public cVB_VoucherAmt  As Double

'Cancel Voucher
Public bVB_CancelVoucher  As Boolean
Public tVB_CancelVoucherNo  As String
Public cVB_CancelVoucherAmt As Double
Public tVB_CancelTmnNumRef As String
Public tVB_CancelTransNoRef As String
Public dVB_CancelTransDateRef As Date

'Sale Duty Free
Public cVB_DutyFreePrice As Double 'ราคาต่อหน่วยของสินค้าDuty Free
Public cVB_DutyFreeDeposit  As Double 'จำนวนเงินมัดจำต่อหน่วยของสินค้า Duty Free

'Other Receive
Public bVB_TrnOthReceive  As Boolean
Public cVB_OthReceiveAmt  As Double

'Other Disbursement
Public bVB_TrnOthDisb As Boolean
Public cVB_OthDisbAmt  As Double

'Create Sale VIP
Public bVB_TrnVIP  As Boolean
Public bVB_TrnCreateVIPSale As Boolean

'Trainning Mode
Public bVB_TrnTrainningMode As Boolean

Public bVB_CNTrnPay  As Boolean

'Reprint
Public bVB_TrnReprint  As Boolean
Public tVB_CNClosingReprint  As String

'Credit Card Tender
Public bVB_IsCrdCard As Boolean
Public tVB_CrdCardNo  As String '*หมายเลขอ้างอิง (เครดิต:เลขบัตร, เช็ค:เลขที่เช็ค …) (Number of  This Tender (Such as if Tender is Credit Card, it mean Credit Card Number and if Tender is Gift Voucher, It mean Gift Voucher Number. )
Public tVB_CrdExpD As String
Public tVB_CrdExpM As String
Public tVB_CrdExpY As String
Public dVB_CrdDateExp As Date
Public tVB_CrdType As String
Public tVB_CrdTypeName  As String
Public tVB_CrdKey As String
Public cVB_CrdAmt   As Double
Public tVB_CrdApprCode As String 'Authorization Number . (Approved Code from Bank.)
Public cVB_CrdRate  As Double
'*KT 54-06-01 V.4.0.0  Req.19 OnTop Specify POSNO
Public bVB_CrdPOSNoOnTopFlag As Boolean
Public tVB_CrdPOSNoOnTopAlwType As String
Public tVB_CrdFloorAlwOnTop As String
Public tVB_CrdStaTierOntopType As String '0 or ว่าง : Tier Amt 1: Tier %
Public tVB_CrdStaAlwTierAmtLast  As String 'Y : Allow Calculate from Last Tier N: Not Allow Calculate from Last Tie   *** Tier Amount only
Public tVB_CrdEvnID As String '*KT 54-11-17 V.4.0.5 DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc Event ID ของบัตรเครดิตการ์ด
'*Ao 2013/02/22
Public tVB_CrdStaBPSale As String
Public tVB_CrdStaBPReturn As String
Public tVB_CrdStaReEntry As String
Public tVB_CrdStaMallCard As String
'Gift voucher Tender
Public tVB_GiftVoucherNo   As String

'Special Gift Voucher
Public tVB_SpcGiftVoucherNo   As String

'Voucher
Public tVB_VoucherNoRet   As String
Public cVB_VoucherAmtRet   As Double
Public tVB_VoucherTmnNoRet  As String
Public tVB_VoucherTrnNoRet   As String
Public dVB_VoucherTrnDateRet   As Date

'Vendor Coupon
Public tVB_VendorCpnNo  As String
Public cVB_VendorCpnAmt   As Double
Public cVB_TotalDiscVendorCpnAmt   As Double
Public bVB_PayVendoeCancel  As Boolean
Public nVB_VendorQty  As Integer   '*Ao 53-10-16

'Special Coupon
Public tVB_SpecialCpnNo As String

'Credit Store
Public tVB_CrdStoreNo As String

'Store Coupon
Public tVB_StoreCpnNo As String

'Credit Note
Public tVB_CreditNoteNo As String

'Foreign Currency
Public tVB_ForeignCurName  As String
Public tVB_ForeignCurSign As String
Public tVB_FTRteCode As String 'รหัสสกุลเงิน
Public cVB_FCSrcRteAmt As Double 'จำนวนเงิน
Public cVB_FCSrcRteFac As Double 'อัตราแลกเปลี่ยน
Public tVB_FTRteCodeLocal As String 'รหัสรหัสสกุลเงินท้องถิ่น

'Credit Voucher
Public tVB_CNCVNo As String
Public dVB_CNCVDate As Date
Public tVB_CNCVNoReIssue As String
Public dVB_CNCVDateReIssue As Date
Public cVB_CNCVAmtReIssue As Double
Public cVB_CNCVAmt As String

Public bVB_DiscAccept  As Boolean

'Promotion/Promotion Time sale/Deal
Public cVB_TotolPmtAmt As Double

'Discount Type
Public bVB_CNTrnDisc As Boolean
Public bVB_CNTrnDiscEx As Boolean '*KT 57-05-10 TK-ISS3000-188 แก้ไขกรณีกด Esc จากหน้า Payment แล้วแสดงหน้าใส่ GST
Public tVB_CNTrnDiscType As String

'Disccount Mall Card
Public tVB_DiscMallCardID As String
Public dVB_DiscMallCardExp As Date
Public tVB_DiscMallCardKey As String
Public tVB_DiscMallCardType As String
Public cVB_DiscMallRate  As Double
Public cVB_DiscMallRanking  As Double
Public cVB_DiscMallGrp As String
Public nVB_DiscMallCount As Long
Public tVB_MallCardGroup As String 'รหัสกลุ่มของ Mall Card
Public nVB_RegMallCount As Long
Public tVB_DiscMallCardIDByTel As String
Public cVB_BalMallDisc  As Double
Public tVB_MallCardIDRef As String 'เก็บค่าที่มาของ Mall Card ID ที่ได้ (Manual="*", EDC="") '*Tao 60-04-20 TK-ISS3000-632

'Disccount Employee Card
Public tVB_DiscEmpCardID As String
Public dVB_DiscEmpCardExp As Date
Public tVB_DiscEmpCardKey As String
Public tVB_DiscEmpCardType As String
Public cVB_DiscEmpRate  As Double
Public cVB_DiscEmpRanking  As Double
Public tVB_DiscEmpChkLimit  As String

'Mall Coupon
Public cVB_MallCpnAmt  As Double

'Transaction Discount/Markdown
Public cVB_DiscTransAmt  As Double

'OnTop

Public cVB_CNOntopMaxAlw As Double 'มูลค่าที่นำไปคำนวนส่วนลดได้
 'ยอดรวมที่สามารถนำไปคิดส่วนลด On Top ได้ เช่น
 'a ราคา 1000 = OnTop
 'b ราคา 50 =  ไม่เป็น OnTop
 'cVB_CNOntopMaxCal = 1000
Public cVB_CNOntopMaxCal As Double
Public cVB_CNOntopMaxDisc As Double
Public bVB_IsOnTop As Boolean
Public nVB_OnTopCount As Integer 'Count การรูดบัตรเครดิต
Public aVB_CreditCrd()  As String 'เก็บ Array บัตรเครดิต

Public cVB_AdjustMalCrdOfOntop As Double 'มูลค่าที่ไปส่วนลดของ Ontop ของการรูดบัตรเครดิต
Public cVB_AdjustDiscMalCrdOfCrd As Double 'มูลค่าที่ไปตัดการลดของ Mall Card กรณีที่มีการรูดบัตรเครดิต

Public Type SKULNDISC
    nSeqNo As Long
    tSKU As String
    tPLU As String
    cQty As Double
    cSalePrice As Double
End Type

'Line Item Discount/Markdown
Public xVB_SKULnDisc As SKULNDISC
Public nVB_TypeDiscLineItem  As Byte '0: persent 1:amount 2: new price
Public cVB_DiscLineItem  As Double
Public cVB_SaleLineItemAmt   As Double
Public bVB_LineDiscItems As Boolean

Public bVB_SKUisPromotion As Boolean
Public bVB_SKUisChkPromo  As Boolean
Public bVB_SKUisChkPromoTime  As Boolean

Public tVB_SKUSymbol As String '*KT 55-08-17 Takashimaya

Public nVB_TypeDisc  As Byte '0: persent 1:amount 2: new price
Public cVB_Disc As Double

'setting pos
'Public tVB_OldFontName As String
'Public tVB_PrintType As String
Public tVB_Display As String
'Public tVB_DspPort As String
Public tVB_Drawer As String
'Public tVB_DrwPort As String
'Public tVB_Printer As String
Public tVB_PrnDevice As String
'Public tVB_Prn950Device As String
'Public tVB_Prn950SlipPrn As String
    
Public tVB_CNPthBnkTr As String
Public tVB_CNPthDrive  As String
Public tVB_CNPthReScrip  As String

Public tVB_FileName() As WIN32_FIND_DATA '*zeah จัดเก็บรายชื่อไฟล์ในการค้นหา

Public nVB_TimeDisplayMsg As Long 'Interval of display message on wMainRet
'*AO51-03-21
Public dVB_CNDateAll As Date
Public tVB_SelConfigDate As String
Public bVB_ConfigDate As Boolean
Public bVB_ConfigBanknote As Boolean
Public aVB_CNStang() As Double 'ค่าสตางค์
'Public oRep As New cReplicate
Public oCnRep As ADODB.Connection
Public bVB_CNRepChk As Boolean
Public cVB_CNDiscAmt As Double              'Dis Return
Public aVB_CNMallCard() As CardType
Public aVB_CNOnTop() As CardType
Public aVB_CNEmpCard() As CardType
Public bW_MallANDTopALL As Boolean   'ใช้สำหรับเช็คว่ายกเลิก MallCard หรือ OnTop ทั้งหมด
Public nVB_ResolutionX As Long
Public nVB_ResolutionY As Long
Public bVB_Res640X480 As Boolean
'____________________________________________________________________________
'Store
Public bVB_StoreStatusOpen As Boolean

Public bVB_FormPayAct As Boolean 'ตัว Flag บอกว่ามีการทำงานในช่วง Form wCNSubPay

Public aVB_CNStangRndAlw() As Double 'ค่าสตางค์ที่อนุญาติให้ทำการ Round Diff

Public nVB_HisSelTrn As Long 'เก็บค่าการเลือกแถวหน้าจอ wCNBegin(Transaction Menu)

Public bVB_CNStaMain As Boolean '*TW 51-04-23 for Discount Tax Except Sale

Public oVB_TextClear As TextBox

'KeyCode & Shift  ใน wShowKB
Public nVB_FKBKeyCode As Integer
Public nVB_FKBShift As Integer
Public bVB_FKBCancel As Boolean
Public bVB_FactQtyChange As Boolean

Public oVB_PicLogo   As IPictureDisp 'ใน Receipt
Public oVB_PicLogoCpn   As IPictureDisp 'Logo Promotion Coupon

Public Type PicLogoCpn
    tPicName As String
    oPic As IPictureDisp
    tID As String
End Type
Public aVB_PicLogoCpn() As PicLogoCpn

Public nVB_lServiceOnOffWnd As Long 'เก็บค่า Handle ของ ServiceOnOff.exe
Public nVB_lServiceTransWnd As Long 'เก็บค่า Handle ของ ServiceTranfer.exe
Public nVB_lServiceAutoClaerWnd As Long 'เก็บค่า Handle ของ ServiceAutoClear.exe
Public nVB_lServiceAutoReplicateWnd As Long 'เก็บค่า Handle ของ ServiceAutoReplicate.exe

Public tVB_CNMainLogo  As String 'Path ของ Logo POSFront
Public nVB_CNCnnTimeOutSrv As Long 'ConecttionTimeOut ของ Server
Public nVB_CNCnnTimeOutLocal  As Long  'ConecttionTimeOut ของ Local
Public nVB_CNCnnTimeOutMem As Long  'ConecttionTimeOut ของ Member
Public nVB_CNCmdTimeOutSrv As Long 'CommandTimeOut ของ Server
Public nVB_CNCmdTimeOutLocal  As Long  'CommandTimeOut ของ Local
Public nVB_CNCmdTimeOutMem As Long  'CommandTimeOut ของ Member

Public bVB_BrwItem As Boolean

'Unlock
Public bVB_CNUnLock As Boolean

'System Date Local
Public bVB_CNSysDateLocal  As Boolean 'Ture : มีการ Change System Date บน Local   False :  ยังไม่ทำการ Change

Public cVB_CNSumOntopMaxCal As Double '*TW 51-08-15
Public tVB_CNUpdateEmpLimit As String '*TW 51-08-18 เก็บ Script ที่ใช้ Update Employee Limit Credit

Public nVB_CNScollItem As Long 'หมายเลขแถวที่ถูก Scoll ที่ Form Item Emptry
Public tVB_CNWheBill As String

Public tVB_TerminalNoAddDptLast  As String
Public tVB_TransNoAddDptLast As String
Public dVB_TransDateAddDptLast  As Date
Public tVB_TransTypeAddDptLast As String

'*TW 51-09-03
Public cVB_CNAddSubTotal As Double
'Public arrVB_CNKeepDT()
'Public arrVB_CNKeepCD()
Public bVB_CNReSumAmount  As Boolean

'*KT
Public nVB_TotalTmnNumCopy As Long
Public tVB_TrnNumCopyMin  As String
Public tVB_TrnNumCopyMax  As String

Public bVB_CNSplash As Boolean '*KT 51-09-17 ป้องการ Key กรณีที่ Splash ขึ้น

Public bVB_CNTenderAct  As Boolean '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero

Public bVB_CNStartBill  As Boolean '*KT 51-09-17 แก้ไขเรื่องการยิงสินค้าขึ้น Zero

Public bVB_TextPromptAct As Boolean

Public aVB_PathFileBck(4) As String

Public bVB_ImageLogo As Boolean
Public bVB_ImageLogoCpn As Boolean
Public bVB_SetImageLogoCpn As Boolean

Public bVB_CNCreateTblTmp As Boolean 'Flag ที่บอกว่ามีการสร้างตารางขายที่เป็น Temp แล้ว*KT 51-10-08 V 3.0.131/V 3.0.200
Public bVB_CNCreateTblTmpFrmCloseRpt  As Boolean 'Flag ที่บอกว่ามีการสร้างตารางขายที่เป็น Temp แล้ว จาก Closing Report เพื่อไปเป็นเงื่อนตอน Open Terminal*KT 51-10-10 V 3.0.132/V 3.0.200

Public bVB_CNZero As Boolean

Public bVB_CNPayCrdNoSuccess As Boolean 'Flag ที่แสดงว่าการจ่าย Credit Card/Credit Store แล้วยังเหลือเศษต่ำกว่า .12

Public bVB_CNRecovryAct As Boolean 'Flag ที่แสดงว่ามีการทำ Recovery ถ้า True ยังอยู่ระหว่างทำงาน False ทำงานเสร็จแล้ว ดัก  Error กรณีที่ดึงสาย แล้วกด Total เบิ้ล 2 ครั้ง

Public bVB_CNAddDpsCLast As Boolean 'Flag ที่แสดงว่ามีการทำรายการ Add Deposit Cancel ตัว Add deposit ที่เกิดขึ้นสุดท้าย  True : ใช่ Flase : ไม่ใช่

Public bVB_CNTrnOnlineSuccess As Boolean 'Flag ที่แสดงว่ามีการรายการที่ทำได้เฉพาะ Online ได้ผ่านตลอดทาง True : ใช่ Flase : ไม่ใช่
Public bVB_CNNoUpd  As Boolean
Public bVB_CNTrnComplete As Boolean '*KT 53-03-02 PH3-ML-SF-051  No EJ, Dat File 1 kb, No data on Database but it's Sale Transaction '*KT 53-05-29

Public bVB_CNChkTrnIPV As Boolean   'Mikky 53-08-11 PH3-ML-SF-126   เพิ่มการตัวแปรการเช็ค IPV เนื่องจากใช้ Function W_SetToComplete ร่วมกับ Trn อื่นๆ

Public nVB_CNCountFolderBck As Long 'เป็นตัวนับการซ้ำในการสร้าง Folder ของการเก็บ Dat '*KT 52-01-20

'Member Price
Public bVB_CNMemberPrice As Boolean 'Flag แสดงว่ารายการที่ทำอยุ่ปัจจุบันเป็นการระบุ member เพื่อต้องการ Fuature Member Price
Public Type MEMBERP 'Data Type กรเก็บข้อมูล member price
    ID  As String 'หมายเลข member
    FName As String 'ชื่อ member
    LName As String 'สกุล member
    Tel As String 'หมายเลขโทรศัพท์ที่ทำการค้นหา
    Expire As Date
    Status As String
    '---------------------------------------------
    MallCardType As String
    MallCardExpire As Date
    MallCardMemberPrice As String
    '--------------------------------------------
    'Point
    MallCardMemberPoint As String
    CardIDToken As String '*Bmp 61-02-15 Token
End Type
Public xVB_CNMemberP As MEMBERP 'ตัวแปรสำหรับเก็บข้อมูลกรรีการ Register เข้าใช้ Member Price
Public tVB_StaMemPriceAlw As String 'Flag แสดงสถานะการอนุญาตให้มีการใช้ Member price หรือไม่ของ SKU นั้น
Public tVB_StaMemPricePromoAlw As Boolean 'Flag แสดงสถานะว่ามีการ update ค่า member price ที่ตาราง DT แล้ว
Public tVB_CNEventMemPrice As String 'เก็บค่า Event ID นด Member Price
Public nVB_CNMemPriceMaxQty As Long
Public tVB_CNMemPriceType As String '*KT 52-09-07 Member Price + Consignment ประเภท  %,Amount off,New Price ของ Member Price
Public cVB_CNMemPriceAmt As Double '*KT 52-09-07 Member Price + Consignment  จำนวน %,Amount off,New Price ของ Member Price
Public cVB_CNMemPriceAmtDis  As Double '*KT 52-09-07 Member Price + Consignment  จำนวน %,Amount off,New Price ของ Member Price

Public Type POINTDPT
    tDptCode As String
    tBrdCode As String
    cRank As Double
    cPoint As Double
    cAmt  As Double
    cMin As Double
    cMax As Double
    tCode As String
    tTier As String
    tUseTierCoo As String
    tAdd As String
    tAlw As String
    tEvnID As String '*KT 54-11-18 V.4.0.5
End Type
'
Public Type POINTSKU
    tSKUCode As String
    cRank As Double
    cPoint As Double
    cAmt  As Double
    cMin As Double
    cMax As Double
    tCode As String
    tTier As String
    tUseTierCoo As String
    tAdd As String
    tAlw As String
     tEvnID As String '*KT 54-11-18 V.4.0.5
End Type

Public Type POINTSEQ
    nSeqNo As Long
    cRank As Double
    cPoint As Double
    cAmt  As Double
    cMin As Double
    cMax As Double
    tType As String '0: Cooperate Point 2: Section Point
    dST As Date
    dED As Date
    tStaAct As String
    tEvnID As String '*KT 54-11-18 V.4.0.5
End Type

Public Type CARDPreDay
    tTdmCode As String
    tCrdType As String
    tStaDeptCode As String
    tStaDataSKU As String
    tStaDataDept  As String
    tStaDataSKUAlw As String
    tStaDataDeptAlw  As String
    tStaDataSKUNotAlw As String
    tStaDataDeptNotAlw  As String
End Type

'member point
Public tVB_StaMemPointAlw As String 'Flag แสดงสถานะการอนุญาตให้มีการใช้ Member point หรือไม่ของ Mall card
Public tVB_MemPointID As String 'เก็บค่า Member ID ที่ได้ member point
Public tVB_MemPointIDTmp As String 'เก็บค่า Member ID ที่ได้ member point ในการ Redeem,IPV
Public nVB_CurMemPointCount  As Long 'เก็บค่าลำดับ Member ID ที่ได้ Point
Public nVB_MemPointMnu  As Long 'เก็บ Member ID ที่ได้ Point จากการรูดส่วนลดหรือ Register 1 : discount 2 : register
Public aW_PointDpt() As POINTDPT
Public aW_PointMall() As POINTDPT
Public aW_PointSKU() As POINTSKU
Public aW_PointDpt2() As POINTDPT
Public aW_PointSKU2() As POINTSKU
Public cVB_MemLPointRet As Double
Public tVB_MemNameRet As String
Public tVB_MemEvnCode As String  '*KT 54-11-18 V.4.0.5  DCR_5411003 PRJ09010TheMall-R4-SE-External Spec.Upload Bonus Points-00.04.00-FSBT.doc
Public tVB_IsMemJIS As String

'find member
Public tVB_FindMember As String 'หมายเลขสมาชิกที่มีการค้นหาสมาชิกด้วยเบอร์โทรแล้วมากกว่า 1 หมายเลขแล้วมีการเลือก

Public bVB_ReConnectMem As Boolean 'Flag แสดงให้ทำการ reconnect database member กรณีที่รูด mall card
Public bVB_ReConnectEmp As Boolean 'Flag แสดงให้ทำการ reconnect database member กรณีที่รูด employee card
Public tVB_InStoreBar As String             'InStoreBar ของ SKU สินค้าประเภทน้ำหนัก

'Public cVB_VatMemberPrice As Double
'Public cVB_VatableMemberPrice As Double

'PH 2.0.0
Public bVB_CNFlagRepSta As Boolean 'Flag แสดงสถานะการทำงานของ ServiceAutoReplicate.exe

'previwe day
Public bVB_CNModePreviewDay As Boolean 'Flag แสดงสถานการใช้ Function Preview Day
Public tVB_CNTblSku  As String 'TCNMSku
Public tVB_CNTblSkuBarcode  As String 'TCNMSkuBarcode
Public tVB_CNTblPromo  As String 'TCNMPromo
Public tVB_CNTblPromoTime  As String 'TCNMPromoTime
Public tVB_CNTblDealAct  As String 'TCNMDealAct
Public tVB_CNTblMBPrice As String 'TCNMMemberPrice
Public tVB_CNTblSkuPoint As String 'TCNMSkuPoint
Public nVB_RegPreType As Long 'Flag บอกประเภทการ Register กรใช้ Preview Day 1: Credit Card 2 : Mall Card 3 : Non Card
Public tVB_CNPreDayIDReg As String  'หมายเลขบัตรที่ผ่านการ Register ทั้ง Credit Card,Mall Card,Non Card
Public bVB_CNTrStaRegPreviewDay As Boolean 'Flag แสดงสถานการผ่านการ Register Preview day
Public aW_CardPreDay(1) As CARDPreDay
Public orsVB_CNSKUCrd As ADODB.Recordset
Public orsVB_CNDeptCrd As ADODB.Recordset
Public orsVB_CNSKUMall As ADODB.Recordset
Public orsVB_CNDeptMall  As ADODB.Recordset
Public orsVB_CNSKUNonCard As ADODB.Recordset
Public orsVB_CNDeptNonCard As ADODB.Recordset
Public tVB_CNSKUFlagPrv As String 'Flag แสดงสถานะว่าสินค้านั้นได้รับ Preview Day
Public tVB_CNSKUFlagNotAlwPrvUse As String 'Flag แสดงสถานะว่าสินค้านั้นเข้าเงื่อนไข Allow /Not Allow Preview Day ของบัตร Credit Card,Member Card,Non Card
Public tVB_CNSKUFlagAlwPrvUsePromo  As String 'Flag แสดงสถานะว่าสินค้านั้นเข้าเงื่อนไข Allow /Not Allow Preview Day ของบัตร Credit Card,Member Card,Non Card สำหรับ Promotion Time,Promotion
Public tVB_CNSKUFlagAlwPrvUseMBPri As String 'Flag แสดงสถานะว่าสินค้านั้นเข้าเงื่อนไข Allow /Not Allow Preview Day ของบัตร Credit Card,Member Card,Non Card สำหรับ  Member Price

'member point
Public tVB_CNSkuPointAlw  As String 'Y = Allow Member Point  or N = Not Allow Member Point
Public tVB_CNSkuEventPoint As String 'Event ID that triggers activation
Public cVB_CNSkuPointAmt  As Double 'Member Point by SKU
Public dVB_CNSkuDatePoint  As Date 'End Date of Member Point  Allow
Public tVB_CNSkuPointCooAlw  As String 'Y: Allow Member Point(Cooperate[General]) N: Not Allow Member Point(Cooperate[General])
Public cVB_TotalSecsionTierAmt As Double 'กรณีที่ Secsion มีการ set ว่าให้ส่วนของ Tier ใช้ Tier ร่วมกับ Tier ของ Coopearte นั้น ให้ทำการเก็บค่าไว้ในตัวแปร Global แล้วไป+ เพิ่มที่ส่วนการคำนวนแต้มของ Tier Cooperate
Public bVB_CNChgDptDate As Boolean
Public bVB_CNPastOffline As Boolean
Public cVB_CNTotalOntopByTicket As Double 'ยอดรวมส่วนลด Ontop ที่ได้ต่อ Ticket
Public cVB_CNTotalSpecialByTicket As Double 'ยอดรวมที่ Mall Card ลดลงจากการทำ Speial Gift Voucher/Special Cpn/Credit Card ที่ Not Allow Mall Card ที่ได้ต่อ Ticket
Public tVB_CNSkuPointAlwDiscountMall  As String 'สถานว่าได้ SKU Point แลั้วได้รับส่วนลด Mall Card อีกหรือไม่ Y = Allow Discount or N = Not Allow Discount *KT 54-06-20
Public tVB_CNSKUStaFixPointFlag As String 'สถานะ SKU Point ว่าเป็น Time หรือ Point (Y: Point N: Time)
Public cVB_CNSkuPointTime  As Double 'SKU Point by time

Public tVB_SKUConsignFlag As String ' เก็บค่า Flagของ SKU ว่าเป็นสินค้า Consignment หรือไม่ N: ไม่ใช่ Y: เป็น *KT 52-08-06
'*KT 52-08-19 เกิด Offline ขณะทำรายการ tEN_TrnIPV, tEN_TrnIPVDutyFreeSale, tEN_TrnIPVRedeem, tEN_TrnIPVReturn, tEN_TrnIPVTaxExceptSale, tEN_TrnIPVVIPService, tEN_TrnIPVOtherReceive, tEN_TrnIPVOtherDisbursement
'tEN_TrnDeposit, tEN_TrnRedeem, tEN_TrnAddDeposit, tEN_TrnDepositCancel, tEN_TrnAddDepositCancel
'tEN_TrnVoucher, tEN_TrnCancelVoucher
Public bVB_TicketNotSuccess As Boolean

Public nVB_CNWaitTranfer As Long 'ตัวแปรเก็บเวลาในการรอการทำงานของ ServiceTranfer เป็นวินาที

Public nVB_CNCountBegin As Long 'ตัวแปรเก็บจำนวนครั้งในการเปิด Begin Trans ของ Connection Local
Public nVB_CNCountBeginSrv  As Long 'ตัวแปรเก็บจำนวนครั้งในการเปิด Begin Trans ของ Connection Server

Public oVB_LogFile As TextStream 'ตัวแปร  TextStream สำหรับ Write Log ExitPrgYYYYMMDD.log
Public oVB_LogMall As TextStream 'ตัวแปร  TextStream สำหรับ Write Log MallCardLogYYYYMMDD.log

'Public bVB_CNbHistoryHaveDael As Boolean  '*KT 53-01-08 กรณรก่อนหน้ามีกด Total แล้วเคยได้ Deal ต้องการให้มีการแสดงข้อมูลใน Table ใหม่

Public cVB_CNRetSaleAmt As Double 'ตัวแปรเก็บราคารวมของสินค้าก่อนลด *KT 53-01-26 Communication Sheet PH3-ML-SB-073 Deal & Free  in Return Transaction.xls
Public cVB_CNRemainDiscAmt As Double 'ตัวแปรเก็บยอดที่ไปหัก Remain Discount *KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
Public tVB_CNRemainDiscID As String 'ตัวแปรเก็บเลขที่พนักงานที่ไปหัก Remain Discount *KT 53-01-28 Communication Sheet PH3-ML-SB-074 Remain Discount for Employee.xls
Public bVB_CNReturnOk As Boolean

'*KT 53-04-09 PH3-ML-SF-112 Allow Scanner when alert message
Public tVB_CNMsgFrm As String
Public tVB_CNMsgStyle As String
Public tVB_CNMsgAppTitle As String

'*Ao 53-10-13   New vendor
Public bVB_CheckVendor As Boolean
Public cVB_TotalVendor As Double
Public nVB_RowVendor As Integer
Public nVB_SeqRC As String
Public tVB_VendorCnpType As String
Public bVB_VedFrom As Boolean
Public tVB_CardNo As String
Public cVB_Amt As Double
Public nVB_CnpQty As Integer
Public nVB_SeqNo As Integer
Public cVB_SrcAmt As Double
Public tVB_SkuIndex As String

'*Ao 2010/12/16 Price Override

Public cVB_AFGroupPM As Double
Public cVB_OverrideAmt As Double
Public tVB_CNSkuIndex As String

Public tVB_CNTranMDownNo As String '*KT 54-07-01 V.4.0.0 Req.18 Add 5 Types of TransMarkdow

'*KT 54-07-25 V.4.0.1 Req.6 Phase4.0(Member Point) Promotion Coupon
Public Type PmtCpn
    ID As String
    Amt As Double
    Val As Double
    Qty As Long
End Type
'Public tVB_CNPmtCpn As PmtCpn

'*KT 54-08-03 V.4.0.1 Req.2 Phase4.0(Member Point) Member Point Based on Tender
Public cVB_CNTenderAmtAlwPoint As Double 'ยอดทั้งหมดของ Tender ที่อนุญาตให้ได้รับ Point

Public Type POINTSKUTime
    tSKU As String
    cRank As Double
    cPoint As Double
    cAmt  As Double
    cMin As Double
    cMax As Double
    nTime As Long
    nQty As Long
    tType As String '0: Cooperate Point 1: Section Point
End Type

Public Type SKUPointByTime
    nSeqNo As Long
    tSKU As String
    cPoint As Double
    tType As String '0: Cooperate Point 1: Section Point
End Type

'Public aVB_CNSKUPointByTime()  As SKUPointByTime 'ยอดทั้งหมดของ Tender ที่อนุญาตให้ได้รับ Point

Public tVB_CNErrDescript As String
Public nVB_CNErrNo As Long

Public bVB_CNShwChange As Boolean

Public nVB_CNTrackNoMSR As Long

'NETS/EDC
Public tVB_TypeOfEDC As String      '1: EDC    2: NETS
Public tVB_CNEdcPort As String      'Port No. EDC
Public tVB_CNNETSPort As String  'Port No. NETS
Public tVB_StepAckOfEDC As String

'Issue,ReIssue CV
Public bVB_CNReIssueCancel As Boolean
Public tVB_CNCVNumRefCancel As String
Public tVB_CNCVNumRefRCCancel As String
Public cVB_CNCVNumRefCancelAmt  As Double
Public orsVB_CNCVNumRefCancel As ADODB.Recordset
Public cVB_CNDiscEmpAmt As Double
Public tVB_CNCVType As String

'Checker Mode
Public bVB_CheckerMode  As Boolean
Public bVB_ChkMode As Boolean

Public tVB_FTTmnNum  As String
Public tVB_FTShdTransNo  As String

Public bVB_CNStartBank As Boolean
Public bVB_CNForCurrency As Boolean

Public tVB_CNTypePrinter As String '0:Slip Printer 1:Receipt Printer
Public tVB_CNReEntryIDMain As String

Public bVB_CNStartBankDT As Boolean
'*Eaw 55-11-27
Public tVB_NETSApprCode As String  'NETS
Public tVB_CashCrdApprCode As String  'Cash Card
Public tVB_FlashApprCode As String   'NETS
'*Ao 2012/11/27
Public cVB_CNRegPrice As Double
Public cVB_PayCredit As Double
Public tVB_NETType As String
'Public aVB_EDC() As String 'array เก็บ ข้อความหัว Ticket
Public tVB_EDCDetail As String
Public tVB_EDCComport As String
'*Eaw 55-11-29
Public tVB_CrdCardType As String  'เก็บ credit card Type
Public tVB_CouponNo As String
Public bVB_Discount91 As Boolean

Public bVB_CNTestSlipPrn As Boolean
Public oVB_CNReIssue As eEN_ReIssue
'*KT 56-01-15
Public bVB_CNOverShortSta As Boolean
Public bVB_CNInputGSTReEntry As Boolean

Public tVB_Free0020 As String

'*Eaw 56-02-04
Public tVB_2LineCode As String
Public cVB_2LinePrice As Double
'*Eaw 56-02-06 For Recovery Process
Public tVB_CNTblHDR As String
Public tVB_CNTblDTR As String
Public tVB_CNTblCDR As String
Public tVB_CNTblRCR As String
Public tVB_CNTblCDRTmp As String
Public tVB_CNTblDTRTmp As String
Public bVB_CNPaid As Boolean
Public bVB_CNDiscount As Boolean
Public bVB_IsRecovery As Boolean
Public nVB_CNPrintUPC As Integer '*Bump 56-03-01 TK-ISS3000-055 สำหรับตรวจสอบค่าจาก StoreParameter ว่า การพิมพ์ Receipt/EJ นั้นจะให้พิมพ์ Barcode หรือรหัสสินค้า
Public nVB_CNPerOffDiskSpace As Integer '*Bump 56-03-01 TK-ISS3000-033 สำหรับตรวจสอบค่าจาก StoreParameter ว่าเช็คพื้นที่ HDD เหลือกี่ % จึงแสดงข้อความเตือน
Public bVB_IsFirstMsgDiskSpace As Boolean '*Bump 56-03-08 ป้องกันการแสดงข้อความซ้ำตอนเปิดครั้งแรก

Public bVB_CNCheckCV As Boolean
Public tVB_Tender As String
Public tVB_MemExp As String
Public tVB_ChkMemExp As String
Public bVB_CNCheckMem As Boolean
Public tVB_PdtNoDis As String
Public tVB_CNAccTypeReEntryMode As String

'*Eaw 56-04-29
Public nVB_CNNETSType As Integer  'NETS Port  เก็บประเภทของ NETS =>1.NET 2.NET BCA 3.NET Gift Card/Prepaid 4. NET Prepaid
Public cVB_RoundChgRC As Double    'CommSheet TK-ISS3000-105-FSBT Answer.xlsx
Public cVB_CouponAmtRet As Double '*Eaw 56-05-08
Public tVB_IsFreeRange As String  '*Eaw 56-05-05 CommSheet TK-ISS3000-159-IT2 00.00.03.xls
Public bVB_CNShortCutSpn As Boolean  '*Eaw 56-06-12 หากกด Hot Key ของ Salesman ให้ Pop up หน้า salesman เสมอ
Public nVB_NETPort As Integer
Public bVB_CNNETPos As Boolean  '*Eaw 56-06-17
Public tVB_CNVirTerminalNum As String  '*Eaw 56-06-28
Public tVB_CNSmnID As String '*Eaw 56-07-03 CommSheet TK-ISS3000-263.xls
Public cVB_CVAmtEx As Double '*Eaw 56-07-04 CommSheet TK-ISS3000-257.xls
Public tVB_CNReEntryDate As String
Public cVB_GSTVatAmt As Double '*Eaw 56-07-31CommSheet TK-ISS3000-283 TR FAPL-063.xls
Public tVB_CNCVTel As String '*Eaw 56-07-31 CommSheet TK-ISS3000-202-ST 00.00.03.xls
Public tVB_CNCVTelReIssue As String
Public cVB_MinDeno As Double  '*Eaw 56-08-08 CommSheet TK-ISS3000-271-TR.xls
Public cVB_MaxChange As Double
Public cVB_NeedAmt As Double
Public bVB_ChkEmpPoint As Boolean '*Eaw 56-08-21
Public bVB_ChkMinEmpDisc As Boolean
Public bVB_ChkMarkX As Boolean
Public tVB_OptionMarkX As String
Public cVB_MinValue As Double
Public tVB_NetDesc As String
Public bVB_CNCrdEzLink As Boolean

Public tVB_EJDate As String
Public tVB_LastPosNo As String '*BG 56-10-31

Public tVB_CreditGroup As String    '*Ao 2013/11/05
Public tVB_OverSh As String    '*Ao 2014/07/31

Public tVB_EDCWait As String  '*BG 56-11-11
Public tVB_EDCTimeOut As String  '*BG 56-11-11
Public nVB_SQLVersion As Integer '*56-02-07 Bump หา Version ของ MS SQL ใช้ที่ wCNMoveFile
Public bVB_CNRunning As Boolean '*BG 57-04-19
Public tVB_CNRunning As String '*BG 57-04-19

Public Type SUMDisCount         '*Ao 2014/02/29
    FNIN  As Integer
    FNDctNo As Integer
    FTDctName As String
End Type

Public Type MemberGroup        '*Ao 2014/02/29
    FTMemberID As String
    FTMcdCode As String
    FTMcdCardType As String
    FTMbgGroup As String
    FTMcdStaChkExp As String
    FTDateExp As String
End Type

Public aMBGroup() As MemberGroup

Public bVB_ChkEDC As Boolean
Public bVB_ChkViod As Boolean
Public tVB_CNMnuTemp  As String
'*Bump 57-11-21 CommSheet TK-ISS3000-567(USB).xlsx Wait confirm
Public Enum eDriveType
    Unknown = 0
    Removable = 1
    Fixed = 2
    Network = 3
    CDRom = 4
    RamDisk = 5
End Enum
Public tVB_SdtTaxInv As String  '*Bump 58-01-09 ใช้เก็บ Flag 'Y' ถ้า สินค้ามี Line Discount สำหรับการทำ Return Refer Sale
Public bVB_Mark_X_CRexp    As String  '*Bump 58-02-12 CommSheet TK-ISS3000-578.xls          version 4.2.165 12/02/2015
'* bVB_Mark_X_CRexp = true คือ วันที่บัตรเครดิตหมดอายุ จะถูก mark เป็น XX/XX ตอนพิมพ์ และ EJ
Public bVB_C209    As Boolean

Public bVB_CNCloseTmnSta As Boolean '*KT 58-05-27 TK-ISS3000-587
Public bVB_CNClosingRptSta As Boolean '*KT 58-05-27 TK-ISS3000-587
Public bVB_CNClosingRptAgain As Boolean '*KT 58-05-27 TK-ISS3000-587
Public bVB_FirstLoad As Boolean '*Tao 60-01-20

'*Tao 60-08-25 Wallet Payments
Public tVB_WalletHostURL As String
Public tVB_WalletBrand As String
Public tVB_WalletStoreID As String
Public tVB_WalletDevideID As String
Public tVB_WalletAllowCancelTime As String
Public nVB_WalletTimeout As Integer
Public Const tVB_DllKey = "wrYBnP5gqLcJN1SMiRZwutFGRTzAtjPd5nlWu1gmc2E="
Public Const tVB_DllPass = "TakaShimaya"

Public Type WalletLog
    ReqType As String
    BizDate As String
    Status As String
    OrderID As String
    ReqDT As String
    ReqAmt As Double
    OrigTransID  As String
    RemainingAmt As Double
    Token As String
    Brand As String
    StoreID As String
    DeviceID As String
    CashierID As String
    ResCode  As String
    ErrCode As String
    ErrDesc As String
    TransDT As String
    CancelDT As String
    WalletName As String
    WalletTransID As String
    PartnerTransID As String
    WPSTransID As String
    WalletCurr As String
    WalletCurrAmt As Double
    ExchRate As Double
    ExchRateUnit As String
    WPSRefundID As String
    RefundDT As String
    RefundAmt As Double
    Result  As String
    ReceiptText As String
    RefDoc As String
    StaDoc As String
End Type
Public aVB_WalletLog As WalletLog

'Wallet Type
Public Type WalletType
    TenderCode  As String
    TenderType As String
    TenderSale As String
    TenderReturn As String
    StaSale As Boolean
    StaReturn As Boolean
    StaPoint As Boolean
    StaMallCard As Boolean
    StaReEntry As Boolean
    StaExpDate As Boolean
    StaBPReturn As Boolean
    StaEmpDisc As Boolean
End Type
Public aVB_WalletType() As WalletType
Public bVB_AlwWalletPay As Boolean '*Tao 60-09-04 Allow Wallet Payment
Public bVB_PrintRegSale As Boolean '*Tao 60-09-04 Allow printing of "REGULAR","SALES"(2lines)

'*Bmp 61-02-15 Token
Public oTKToken As Object 'CreateObject("MSSOAP.SoapClient30")
Public bVB_AbleToken As Boolean
Public bVB_AlwToken As Boolean
Public tVB_TokenURL As String 'http://192.11.15.248:8080/axis2/services/SafeNetTokenizer?wsdl
Public tVB_TokenNAEUsr As String 'nae_user
Public tVB_TokenNAEPwd As String 'Pass1234
Public tVB_TokenDB As String 'CARDNO
Public tVB_TokenDBUsr As String 'tkuser
Public tVB_TokenDBPwd As String '!@#Token
Public tVB_TokenFormat As String '4
Public tVB_TokenLuhCheck As String '0
'*Bmp เก็บว่า จำเป็นต้องเช็ค token web service ตอนจบบิล หรือไม่ โดย จะเป็น True ถ้า บิลนั้นมี POINT เพื่อพิมพ์ "OFFLINE*"
Public bVB_ChkWSOFFLINE As Boolean
Public Const nVB_ChkDigitToken As Integer = 7 '*Tao 61-10-05 จำนวนหลักที่จำเป็นต้อง Token

'*Amorn 65-07-11 ประกาศตัวแปรเพื่อการทำ Return 7%
Public bVB_ChkRetGST7 As Boolean
Public cVB_OldGstRate As Double
Public tVB_StoreParaAlwOldGstRate As String '*KT 66-08-02
Public tVB_DepositNoGST7 As String
Public cVB_CNGSTRate As Double '*Bump 66-01-04

Public cVB_CreditRedeemAmt As Double '*Tao 62-07-19 Credit Redeem
Public Const tVB_CreditRedeemName As String = "Redeemed Instant$" '*Tao 62-07-19 ชื่อ Credit Redeem
Public bVB_RePrint_DeToken As Boolean '*Thirapong 2020-10-14 Re-print Previous Transaction must Detoken, refer.CommSheet TTK-ISS3000-769
Public bVB_GVAlwStaffDisc As Boolean '*Thirapong 2021-03-17 Comsheet781 TK Request - Allow Gift Voucher payment for staff discount
Public tVB_VirtualPosNo '*Thirapong 64-06-09 CommSheet TTK-ISS3000-789 TK ISSUE - Re-entry Closing Report printed wrong Virtual POSID_7823.xlsx
Public bVB_Disallow_AddDeposit_OldVatRate '*Bump 65-10-13 STK-POS ComSheet CR803 - 006.xlsx

Public orsVB_DTReturn   As ADODB.Recordset '*KT 63-03-11 ML-POSC-1139

'*KT 66-12-06 Mobile Voucher
Public tVB_CNMVoucherAuthURL As String
Public tVB_CNMVoucherVerifyURL As String
Public tVB_CNMVoucherUpdateURL As String
Public tVB_CNMVoucherClient_id As String
Public tVB_CNMVoucherClient_secretL As String
Public tVB_CNMVoucherGrant_type As String
Public tVB_CNMVoucherScope As String
Public tVB_CNMVoucherContentTypeAuth As String
Public nVB_CNMVoucherTimeout  As Long

Public tVB_CNMVoucherToken  As String
Public nVB_CNMVoucherExpireIn As Long
Public dVB_CNMVoucherExpireInTime  As Date
Public Type oApiHKey
    tHkeyName  As String
    tHkeyVal  As String
End Type
Public aVB_MVoucherHKey()  As oApiHKey
