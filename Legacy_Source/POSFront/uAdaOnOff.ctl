VERSION 5.00
Begin VB.UserControl uAdaOnOff 
   ClientHeight    =   375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2190
   Picture         =   "uAdaOnOff.ctx":0000
   ScaleHeight     =   375
   ScaleWidth      =   2190
   Begin VB.Image oimOnLine 
      Height          =   285
      Left            =   0
      Picture         =   "uAdaOnOff.ctx":24D2
      Stretch         =   -1  'True
      Top             =   30
      Width           =   300
   End
   Begin VB.Label olbOnline 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "On-Line"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   222
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   345
      Left            =   495
      TabIndex        =   1
      Tag             =   "On-Line;Off-Line"
      Top             =   0
      Width           =   1125
   End
   Begin VB.Image oimOffLine 
      Height          =   285
      Left            =   0
      Picture         =   "uAdaOnOff.ctx":8548
      Stretch         =   -1  'True
      Top             =   30
      Width           =   300
   End
   Begin VB.Label olaEvent 
      Caption         =   "olaEvent"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   1110
      Width           =   1815
   End
End
Attribute VB_Name = "uAdaOnOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_ShowPic = 1
Const m_def_OnLineCaption = "On-Line"
Const m_def_OffLineCaption = "Off-Line"
Const m_def_OnLine = True
'Property Variables:
Dim m_Picture As Picture
Dim m_OnLine As Boolean
Dim m_TimeOut As Integer
Dim m_OnLineCaption As String
Dim m_OffLineCaption As String
Dim m_ShowPic As Boolean
Event OnlineOffline() 'MappingInfo=olaEvent,olaEvent,-1,Change

Private Sub EventPing()

End Sub

Private Sub olaEvent_Change()
    RaiseEvent OnlineOffline
End Sub

Private Sub otmPing_Timer()
    Call EventPing
End Sub

Private Sub UserControl_Initialize()
     Call UserControl_Resize
End Sub

Private Sub UserControl_Resize()
    oimOffLine.Height = UserControl.Height
    oimOnLine.Height = UserControl.Height
    oimOffLine.Width = oimOffLine.Height
    oimOnLine.Width = oimOnLine.Height
    oimOffLine.Visible = m_ShowPic
    oimOnLine.Visible = m_ShowPic
    oimOffLine.Stretch = True
    oimOnLine.Stretch = True
    olbOnline.Height = UserControl.Height
    If m_ShowPic Then
        olbOnline.Width = UserControl.Width - (oimOnLine.Width + 30)
        olbOnline.Left = oimOnLine.Width + 30
    Else
        olbOnline.Width = UserControl.Width
        olbOnline.Left = oimOnLine.Left
    End If
End Sub

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_OnLine = m_def_OnLine
    m_OnLineCaption = m_def_OnLineCaption
    m_OffLineCaption = m_def_OffLineCaption
    m_ShowPic = m_def_ShowPic
    Set m_Picture = LoadPicture("")
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    olbOnline.ForeColor = PropBag.ReadProperty("ForeColor", &H404080)
    olbOnline.Font = PropBag.ReadProperty("Font", Ambient.Font)
    oimOffLine.Picture = PropBag.ReadProperty("PicOffLine", Nothing)
    oimOnLine.Picture = PropBag.ReadProperty("PicOnLine", Nothing)
    m_OnLine = PropBag.ReadProperty("OnLine", m_def_OnLine)
    m_OnLineCaption = PropBag.ReadProperty("OnLineCaption", m_def_OnLineCaption)
    m_OffLineCaption = PropBag.ReadProperty("OffLineCaption", m_def_OffLineCaption)
    m_ShowPic = PropBag.ReadProperty("ShowPic", m_def_ShowPic)
    olbOnline.Alignment = PropBag.ReadProperty("Alignment", 0)
    Set m_Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ForeColor", olbOnline.ForeColor, &H404080)
    Call PropBag.WriteProperty("Font", olbOnline.Font, Ambient.Font)
    Call PropBag.WriteProperty("PicOnLine", oimOnLine.Picture, Nothing)
    Call PropBag.WriteProperty("PicOffLine", oimOffLine.Picture, Nothing)
    Call PropBag.WriteProperty("OnLineCaption", m_OnLineCaption, m_def_OnLineCaption)
    Call PropBag.WriteProperty("OffLineCaption", m_OffLineCaption, m_def_OffLineCaption)
    Call PropBag.WriteProperty("ShowPic", m_ShowPic, m_def_ShowPic)
    Call PropBag.WriteProperty("Alignment", olbOnline.Alignment, 0)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=olbOnline,olbOnline,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = olbOnline.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    olbOnline.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=olbOnline,olbOnline,-1,Font
Public Property Get Font() As Font
    Set Font = olbOnline.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set olbOnline.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=oimOnLine,oimOnLine,-1,Picture
Public Property Get PicOnLine() As Picture
    Set PicOnLine = oimOnLine.Picture
End Property

Public Property Set PicOnLine(ByVal New_PicOnLine As Picture)
    Set oimOnLine.Picture = New_PicOnLine
    oimOnLine.Stretch = True
    PropertyChanged "PicOnLine"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=oimOffLine,oimOffLine,-1,Picture
Public Property Get PicOffLine() As Picture
    Set PicOffLine = oimOffLine.Picture
End Property

Public Property Set PicOffLine(ByVal New_PicOffLine As Picture)
    Set oimOffLine.Picture = New_PicOffLine
    oimOffLine.Stretch = True
    PropertyChanged "PicOffLine"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,
Public Property Get Online() As Boolean
    Online = m_OnLine
End Property

Public Property Let Online(ByVal New_Online As Boolean)
    If New_Online Then
        olbOnline.Caption = m_OnLineCaption
        oimOnLine.ZOrder 0
    Else
        olbOnline.Caption = m_OffLineCaption
        oimOffLine.ZOrder 0
    End If
    m_OnLine = New_Online
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get OnLineCaption() As String
    OnLineCaption = m_OnLineCaption
End Property

Public Property Let OnLineCaption(ByVal New_OnLineCaption As String)
    m_OnLineCaption = New_OnLineCaption
    olbOnline.Caption = New_OnLineCaption
    PropertyChanged "OnLineCaption"
End Property

Public Property Get OffLineCaption() As String
    OffLineCaption = m_OffLineCaption
End Property

Public Property Let OffLineCaption(ByVal New_OffLineCaption As String)
    m_OffLineCaption = New_OffLineCaption
    olbOnline.Caption = New_OffLineCaption
    PropertyChanged "OffLineCaption"
End Property


'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,1
Public Property Get ShowPic() As Boolean
    ShowPic = m_ShowPic
End Property

Public Property Let ShowPic(ByVal New_ShowPic As Boolean)
    m_ShowPic = New_ShowPic
    PropertyChanged "ShowPic"
    oimOffLine.Visible = m_ShowPic
    oimOnLine.Visible = m_ShowPic
    oimOffLine.Stretch = True
    oimOnLine.Stretch = True
    If m_ShowPic Then
        olbOnline.Left = oimOnLine.Width + 30
    Else
        olbOnline.Left = oimOnLine.Left
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=olbOnline,olbOnline,-1,Alignment
Public Property Get Alignment() As Integer
    Alignment = olbOnline.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    olbOnline.Alignment = New_Alignment
    PropertyChanged "Alignment"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get Picture() As Picture
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set m_Picture = New_Picture
    PropertyChanged "Picture"
End Property





