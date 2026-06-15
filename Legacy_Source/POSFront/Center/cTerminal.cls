VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTerminal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents ocnServer As ADodb.Connection
Attribute ocnServer.VB_VarHelpID = -1
Private bW_Online As Boolean
Private bW_StoreOpen As Boolean
Private dW_StoreSaleDate As Date


Private tW_Server As String
Private tW_User As String
Private tW_Password As String
Private tW_DB As String

Private nW_SaleMode As eEN_SaleMode
Private dW_TmnSaleDate As Date


'SaleMode
Public Enum eEN_SaleMode
        nEN_OnlineNormal = 0
        nEN_OfflineNormal = 1
        nEN_OfflineOpenWith = 2
End Enum

'SaleMode
Private Enum eEN_ServerMode
        nEN_SaleOffline = 0
        nEN_SaleOnline = 1
End Enum

Dim orsTmp As ADodb.Recordset


Public Function AvailableSaleMode(Optional ByVal pbOpenWith As Boolean) As eEN_SaleMode
Dim tTmp As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -

On Error GoTo ErrHandler

            If bW_Online Then
                    If bW_StoreOpen Then
                            If pbOpenWith Then
                                    AvailableSaleMode = nEN_OfflineOpenWith
                            Else
                                    AvailableSaleMode = nEN_OnlineNormal
                            End If
                    Else
                            AvailableSaleMode = nEN_OfflineNormal
                    End If
            Else
                    AvailableSaleMode = nEN_OfflineNormal
            End If
    


Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
End Function

Public Function TryToOnline() As Boolean
Dim tTmp As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler

            If Not bW_Online Then
                    If Not ReConnect() Then Exit Function
            End If
            
            If Not StoreOpen() Then Exit Function
            
            Select Case nW_SaleMode
            Case nEN_OfflineNormal
            Case nEN_OfflineOpenWith
                    dW_StoreSaleDate = SalesDate
                    If Format(dW_StoreSaleDate, "dd/mm/yyyy") < Format(dW_TmnSaleDate, "dd/mm/yyyy") Then Exit Function
            End Select
            
TryToOnline = True
Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Property Let SaleMode(pnMode As eEN_SaleMode)
        nW_SaleMode = pnMode
End Property

Public Property Get SaleMode() As eEN_SaleMode
        SaleMode = nW_SaleMode
End Property

Public Property Let SaleDate(pnDate As Date)
        dW_TmnSaleDate = pnDate
End Property

Public Property Get SaleDate() As Date
        SaleDate = dW_TmnSaleDate
End Property

Public Function Connect(ByVal ptServer As String, ByVal ptUser As String, ByVal ptPassword As String, ByVal ptDB As String) As Boolean
Dim tSql As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
        
        tW_Server = ptServer
        tW_User = ptUser
        tW_Password = ptPassword
        tW_DB = ptDB
        
        bW_Online = False
        tSql = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" & ptDB & ";Data Source=" & ptServer
        Set ocnServer = New ADodb.Connection
        ocnServer.CursorLocation = adUseServer
        ocnServer.ConnectionTimeout = 5
        Call ocnServer.Open(tSql, ptUser, ptPassword)
        
Connect = True
bW_Online = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Public Function ReConnect() As Boolean
Dim tSql As String
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
        bW_Online = False

                tSql = "Provider=SQLOLEDB.1;Persist Security Info=False;Initial Catalog=" & tW_DB & ";Data Source=" & tW_Server
                Set ocnServer = New ADodb.Connection
                ocnServer.CursorLocation = adUseServer
                ocnServer.ConnectionTimeout = 5
                Call ocnServer.Open(tSql, tW_User, tW_Password)
        
ReConnect = True
bW_Online = True

Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

End Function

Private Function DB_LANbPingByIP(ByVal ptIP As String) As Boolean
'-----------------------------------------------------------
'   Call :
'   Cmt :
'   Developer : SOMJAI
'   Date Create : 02-20-2008 :  08:14:03
'-----------------------------------------------------------
On Error GoTo ErrHandle
Dim tComputer$
Dim oWMIService As Object
Dim oColPing As Object
Dim oStatus As Object

tComputer = "."
Set oWMIService = GetObject("winmgmts:\\" & tComputer & "\root\cimv2")
Set oColPing = oWMIService.ExecQuery("Select * From Win32_PingStatus where Address = '" & ptIP & "'")

For Each oStatus In oColPing
    If IsNull(oStatus.StatusCode) Or oStatus.StatusCode <> 0 Then
        DB_LANbPingByIP = False
    Else
        DB_LANbPingByIP = True
    End If
Next

    Exit Function
ErrHandle:
    DB_LANbPingByIP = False
End Function


Public Property Get StoreOpen() As Boolean
Dim tSql As String
Dim orsTmp As ADodb.Recordset
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
       
       If Not bW_Online Then Exit Function
       tSql = "SELECT FTTmsOpnStore  FROM TsysTerminalSta "
       Set orsTmp = ocnServer.Execute(tSql)
       If orsTmp.EOF Then Exit Function
       If SP_FEDtChkString(orsTmp, "FTTmsOpnStore") = "Y" Then
                bW_StoreOpen = True
       Else
                bW_StoreOpen = False
       End If
        
        Set orsTmp = Nothing
        
StoreOpen = bW_StoreOpen
Exit Function '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:

        Set orsTmp = Nothing
        
End Property

Public Property Get SalesDate() As Date
Dim tSql As String
Dim orsTmp As ADodb.Recordset
'- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
On Error GoTo ErrHandler
       
       If Not bW_Online Then Exit Property
       
       tSql = "SELECT FDCdtDate,FTCdtTime  FROM TSysChgDateTime "
       Set orsTmp = ocnServer.Execute(tSql)
       If orsTmp.EOF Then Exit Property
       dW_StoreSaleDate = SP_FEDtChkString(orsTmp, "FDCdtDate")
        
        Set orsTmp = Nothing
        
SalesDate = dW_StoreSaleDate
Exit Property '- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ErrHandler:
        Set orsTmp = Nothing
        
End Property

Public Property Get Online() As Boolean
       Online = bW_Online
End Property

Public Property Let Online(pbOnline As Boolean)
       bW_Online = pbOnline
       
End Property

Private Function SP_DTEtStrSTD(pdDate As Date, Optional ByVal pbChrRet As Boolean = False, Optional ByVal ptTime$, Optional ByVal pbRetWithTime As Boolean = False) As String
'----------------------------------------------------------
'   Call:   pdDate as selected date
'   Ret:   string of date format "YYYY/MM/DD"
'----------------------------------------------------------
    Dim tRet$
    
    SP_DTEtStrSTD = Format(Year(pdDate), "0000") & "/" & Format(Month(pdDate), "00") & "/" & Format(Day(pdDate), "00")
    tRet = SP_DTEtStrSTD
    If pbChrRet Then

                  If pbRetWithTime Then
                       SP_DTEtStrSTD = "'" & tRet & " " & ptTime & "'"
                  Else
                       SP_DTEtStrSTD = "'" & tRet & "'"
                  End If

    Else
        If pbRetWithTime Then
            SP_DTEtStrSTD = tRet & " " & ptTime
        End If
    End If
    
End Function

Private Function SP_FEDtChkString(ByVal poRs As ADodb.Recordset, ByVal ptField$, Optional pvRetOld) As String
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    variant, only text and numeric
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            If Len(Trim(poRs.Fields(ptField))) = 0 Then
                bEmpty = True
            Else
                SP_FEDtChkString = poRs.Fields(ptField)
            End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDtChkString = ""
        Else
            SP_FEDtChkString = pvRetOld
        End If
    End If
    
End Function

Private Function SP_FEDdChkDate(ByVal poRs As ADodb.Recordset, ByVal ptField$, Optional pvRetOld) As Date
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            'If Len(Trim(poRs.Fields(ptField))) = 0 Then
            '    bEmpty = True
            'Else
                SP_FEDdChkDate = poRs.Fields(ptField)
            'End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDdChkDate = Date
        Else
            SP_FEDdChkDate = pvRetOld
        End If
    End If
    
End Function

Private Function SP_FEDvChkNumeric(ByVal poRs As ADodb.Recordset, ByVal ptField$, Optional pvRetOld) As Variant
'----------------------------------------------------------
'   Call:   poRs as ADODB.Recordset
'               ptField is field to check
'               pvRetOld(variant) if empty, or null want to return old
'   Ret:    numeric value as variant(integer,long,single,double ...)
'----------------------------------------------------------
    Dim bEmpty As Boolean
    
    If poRs.BOF And poRs.EOF Then
        bEmpty = True
    Else
        If IsNull(poRs.Fields(ptField)) Then
            bEmpty = True
        Else
            If Len(Trim(poRs.Fields(ptField))) = 0 Then
                bEmpty = True
            Else
                SP_FEDvChkNumeric = poRs.Fields(ptField)
            End If
        End If
    End If
    If bEmpty Then
        If IsMissing(pvRetOld) Then
            SP_FEDvChkNumeric = 0
        Else
            SP_FEDvChkNumeric = pvRetOld
        End If
    End If
End Function



Private Sub ocnServer_Disconnect(adStatus As ADodb.EventStatusEnum, ByVal pConnection As ADodb.Connection)
'Debug.Print "dis : "

End Sub

Private Sub ocnServer_InfoMessage(ByVal pError As ADodb.Error, adStatus As ADodb.EventStatusEnum, ByVal pConnection As ADodb.Connection)
'Debug.Print "info : " & pError.Description

End Sub
