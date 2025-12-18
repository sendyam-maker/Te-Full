VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030602 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標公報資料維護"
   ClientHeight    =   6375
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   9150
   Begin VB.TextBox textTMBM08 
      Height          =   620
      Left            =   1680
      MaxLength       =   134
      MultiLine       =   -1  'True
      ScrollBars      =   2  '垂直捲軸
      TabIndex        =   9
      Top             =   2580
      Width           =   7272
   End
   Begin VB.TextBox textTA02 
      Height          =   300
      Left            =   1680
      MaxLength       =   4
      TabIndex        =   7
      Top             =   2280
      Width           =   735
   End
   Begin VB.TextBox textNA01 
      Height          =   300
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1980
      Width           =   735
   End
   Begin VB.TextBox textTMBM04 
      Height          =   264
      Left            =   1680
      MaxLength       =   9
      TabIndex        =   4
      Top             =   1710
      Width           =   1332
   End
   Begin VB.TextBox textTMBM02_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1170
      Width           =   4332
   End
   Begin VB.TextBox textTMBM02 
      Height          =   264
      Left            =   1680
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1170
      Width           =   492
   End
   Begin VB.TextBox textTMBM07 
      Height          =   264
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   0
      Top             =   630
      Width           =   732
   End
   Begin VB.TextBox textTMBM01 
      Height          =   264
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   1
      Top             =   900
      Width           =   1332
   End
   Begin VB.TextBox textTMBM03 
      Height          =   264
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   3
      Top             =   1440
      Width           =   1332
   End
   Begin MSComctlLib.ImageList ImgList 
      Left            =   8520
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm030602.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImgList"
      DisabledImageList=   "ImgList"
      HotImageList    =   "ImgList"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textTMBM09 
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   3180
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM10 
      Height          =   300
      Left            =   1680
      TabIndex        =   11
      Top             =   3480
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM11 
      Height          =   300
      Left            =   1680
      TabIndex        =   12
      Top             =   3780
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM12 
      Height          =   300
      Left            =   1680
      TabIndex        =   13
      Top             =   4080
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM13 
      Height          =   300
      Left            =   1680
      TabIndex        =   14
      Top             =   4380
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM14 
      Height          =   300
      Left            =   1680
      TabIndex        =   15
      Top             =   4680
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM15 
      Height          =   300
      Left            =   1680
      TabIndex        =   16
      Top             =   4980
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM16 
      Height          =   300
      Left            =   1680
      TabIndex        =   17
      Top             =   5280
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM17 
      Height          =   300
      Left            =   1680
      TabIndex        =   18
      Top             =   5580
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM18 
      Height          =   300
      Left            =   1680
      TabIndex        =   19
      Top             =   5880
      Width           =   7275
      VariousPropertyBits=   679493659
      MaxLength       =   150
      Size            =   "12832;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM06 
      Height          =   300
      Left            =   2520
      TabIndex        =   8
      Top             =   2280
      Width           =   6432
      VariousPropertyBits=   679493659
      MaxLength       =   12
      Size            =   "11345;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTMBM05 
      Height          =   300
      Left            =   2520
      TabIndex        =   6
      Top             =   1980
      Width           =   6432
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "11345;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱1 :"
      Height          =   240
      Left            =   420
      TabIndex        =   40
      Top             =   3210
      Width           =   1080
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱2 :"
      Height          =   240
      Left            =   420
      TabIndex        =   39
      Top             =   3510
      Width           =   1080
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱3 :"
      Height          =   240
      Left            =   420
      TabIndex        =   38
      Top             =   3810
      Width           =   1080
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱4 :"
      Height          =   240
      Left            =   420
      TabIndex        =   37
      Top             =   4110
      Width           =   1080
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱5 :"
      Height          =   240
      Left            =   420
      TabIndex        =   36
      Top             =   4410
      Width           =   1080
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱6 :"
      Height          =   210
      Left            =   420
      TabIndex        =   35
      Top             =   4710
      Width           =   1080
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱7 :"
      Height          =   240
      Left            =   420
      TabIndex        =   34
      Top             =   5010
      Width           =   1080
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱8 :"
      Height          =   240
      Left            =   420
      TabIndex        =   33
      Top             =   5310
      Width           =   1080
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱9 :"
      Height          =   240
      Left            =   420
      TabIndex        =   32
      Top             =   5610
      Width           =   1080
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "申請人名稱10 :"
      Height          =   240
      Left            =   420
      TabIndex        =   31
      Top             =   5910
      Width           =   1170
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "PS. 申請人名稱至43卷01期(105/01/01)開始匯入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   420
      TabIndex        =   30
      Top             =   6180
      Width           =   3840
   End
   Begin VB.Label Label8 
      Caption         =   "商品類別 :"
      Height          =   240
      Left            =   420
      TabIndex        =   28
      Top             =   2580
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "代理人編號 :"
      Height          =   240
      Left            =   420
      TabIndex        =   27
      Top             =   2310
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "地區編號 :"
      Height          =   240
      Left            =   420
      TabIndex        =   26
      Top             =   2010
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "申請案號 :"
      Height          =   240
      Left            =   420
      TabIndex        =   25
      Top             =   1740
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "商標種類 :"
      Height          =   240
      Left            =   420
      TabIndex        =   24
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "公報卷期:"
      Height          =   240
      Left            =   420
      TabIndex        =   23
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "審定號 :"
      Height          =   240
      Left            =   420
      TabIndex        =   22
      Top             =   930
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "正商標審定號 :"
      Height          =   240
      Left            =   420
      TabIndex        =   21
      Top             =   1470
      Width           =   1215
   End
End
Attribute VB_Name = "frm030602"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/10 Form2.0已修改 textTMBM05/textTMBM06/textTMBM09~18
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

Const MAX_FIELD = 18

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList(MAX_FIELD) As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer

' 第一筆資料的本所案號
Dim m_FirstTM(2) As String
' 最後一筆資料的本所案號
Dim m_LastTM(2) As String
' 目前正在顯示的本所案號
Dim m_CurrTM(2) As String

Dim m_LastTMBM01 As String
Dim m_LastTMBM02 As String
Dim m_LastTMBM07 As String


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT TMBM01,TMBM02 FROM TMBULLETIN " & _
            "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN) AND " & _
                  "TMBM02 = (SELECT MIN(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_FirstTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_FirstTM(1) = rsTmp.Fields("TMBM02")
   End If
   rsTmp.Close

   strSql = "SELECT TMBM01,TMBM02 FROM TMBULLETIN " & _
            "WHERE TMBM01 = (SELECT MAX(TMBM01) FROM TMBULLETIN) AND " & _
                  "TMBM02 = (SELECT MAX(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = (SELECT MAX(TMBM01) FROM TMBULLETIN)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_LastTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_LastTM(1) = rsTmp.Fields("TMBM02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' Load Form
Private Sub Form_Load()
   m_EditMode = 0
   MoveFormToCenter Me

   m_LastTMBM01 = Empty
   m_LastTMBM02 = Empty
   m_LastTMBM07 = Empty

   textTMBM02_2.BackColor = &H8000000F
   'textTMBM05.BackColor = &H8000000F
   'textTMBM06.BackColor = &H8000000F

   InitialField
   RefreshRange
   ShowFirstRecord
   SetCtrlReadOnly True
   UpdateToolbarState
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "TMBM" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, ByVal strData As String)
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
      If strName = m_FieldList(nIndex).fiName Then
         m_FieldList(nIndex).fiNewData = strData
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
   SetFieldNewData "TMBM01", textTMBM01
   SetFieldNewData "TMBM02", textTMBM02
   SetFieldNewData "TMBM03", textTMBM03
   SetFieldNewData "TMBM04", textTMBM04
   SetFieldNewData "TMBM05", textTMBM05
   SetFieldNewData "TMBM06", textTMBM06
   SetFieldNewData "TMBM07", textTMBM07
   SetFieldNewData "TMBM08", textTMBM08
   'Add By Sindy 2017/4/27
   SetFieldNewData "TMBM09", textTMBM09
   SetFieldNewData "TMBM10", textTMBM10
   SetFieldNewData "TMBM11", textTMBM11
   SetFieldNewData "TMBM12", textTMBM12
   SetFieldNewData "TMBM13", textTMBM13
   SetFieldNewData "TMBM14", textTMBM14
   SetFieldNewData "TMBM15", textTMBM15
   SetFieldNewData "TMBM16", textTMBM16
   SetFieldNewData "TMBM17", textTMBM17
   SetFieldNewData "TMBM18", textTMBM18
   '2017/4/27 END
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            'add by nickc 2007/03/03
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   
End Sub

' 新增資料前
Private Sub OnPrepareAddRecord()
   If IsEmptyText(m_LastTMBM01) = False Then
      textTMBM01 = Mid(m_LastTMBM01, 1, 2) & Format(CStr(Val(Mid(m_LastTMBM01, 3, 6)) + 1), "000000")
   End If
   If IsEmptyText(m_LastTMBM02) = False Then
      textTMBM02 = m_LastTMBM02
      textTMBM02_LostFocus
   End If
   If IsEmptyText(m_LastTMBM07) = False Then
      textTMBM07 = m_LastTMBM07
   End If
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textNA01 = Empty
   textTA02 = Empty
   textTMBM01 = Empty
   textTMBM02 = Empty
   textTMBM02_2 = Empty
   textTMBM03 = Empty
   textTMBM04 = Empty
   textTMBM05 = Empty
   textTMBM06 = Empty
   'edit by nickc 2005/04/15 若是新增時卷期不變
   If m_EditMode <> 1 Then
      textTMBM07 = Empty
   End If
   textTMBM08 = Empty
   'Add By Sindy 2017/4/27
   textTMBM09 = Empty
   textTMBM10 = Empty
   textTMBM11 = Empty
   textTMBM12 = Empty
   textTMBM13 = Empty
   textTMBM14 = Empty
   textTMBM15 = Empty
   textTMBM16 = Empty
   textTMBM17 = Empty
   textTMBM18 = Empty
   '2017/4/27 END
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textTMBM01.Locked = bEnable
   textTMBM02.Locked = bEnable
   textTMBM03.Locked = bEnable
   textTMBM04.Locked = bEnable
   textTMBM05.Locked = bEnable
   textTMBM06.Locked = bEnable
   textTMBM07.Locked = bEnable
   textTMBM08.Locked = bEnable
   'Add By Sindy 2017/4/27
   textTMBM09.Locked = bEnable
   textTMBM10.Locked = bEnable
   textTMBM11.Locked = bEnable
   textTMBM12.Locked = bEnable
   textTMBM13.Locked = bEnable
   textTMBM14.Locked = bEnable
   textTMBM15.Locked = bEnable
   textTMBM16.Locked = bEnable
   textTMBM17.Locked = bEnable
   textTMBM18.Locked = bEnable
   '2017/4/27 END
   textNA01.Locked = bEnable
   textTA02.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textTMBM01.Locked = bEnable
   textTMBM02.Locked = bEnable
End Sub

' 取得國家的代碼
Private Function GetNationNo(ByVal strData As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   GetNationNo = Empty
   'Modify By Sindy 2013/8/19 + AND length(na01)=3
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA03 = '" & strData & "' AND length(na01)=3 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NA01")) = False Then
         GetNationNo = rsTmp.Fields("NA01")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 取得公報代理人的代碼
Private Function GetTAgentNo(ByVal strData As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   GetTAgentNo = Empty
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' AND " & _
                  "TA03 = '" & Trim(strData) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA02")) = False Then
         GetTAgentNo = rsTmp.Fields("TA02")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 取得公報代理人的代碼
Private Function GetTAgentName(ByVal strData As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   GetTAgentName = Empty
   strSql = "SELECT * FROM TAGENT " & _
            "WHERE TA01 = 'T' AND " & _
                  "TA02 = '" & strData & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TA03")) = False Then
         GetTAgentName = rsTmp.Fields("TA03")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Function GetMaxTMBM01() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   GetMaxTMBM01 = "0000000001"
   strSql = "SELECT TMBM01 FROM TMBULLETIN " & _
            "WHERE TMBM01 = (SELECT MAX(TMBM01) FROM TMBULLETIN) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetMaxTMBM01 = CStr(Val(rsTmp.Fields("TMBM01")) + 1)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ClearField
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TMBM02 = '" & m_CurrTM(1) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("TMBM01")) Then: textTMBM01 = rsTmp.Fields("TMBM01")
      If Not IsNull(rsTmp.Fields("TMBM02")) Then: textTMBM02 = rsTmp.Fields("TMBM02")
      If Not IsNull(rsTmp.Fields("TMBM03")) Then: textTMBM03 = rsTmp.Fields("TMBM03")
      If Not IsNull(rsTmp.Fields("TMBM04")) Then: textTMBM04 = rsTmp.Fields("TMBM04")
      If Not IsNull(rsTmp.Fields("TMBM05")) Then: textTMBM05 = rsTmp.Fields("TMBM05")
      If Not IsNull(rsTmp.Fields("TMBM06")) Then: textTMBM06 = rsTmp.Fields("TMBM06")
      If Not IsNull(rsTmp.Fields("TMBM07")) Then: textTMBM07 = rsTmp.Fields("TMBM07")
      If Not IsNull(rsTmp.Fields("TMBM08")) Then: textTMBM08 = rsTmp.Fields("TMBM08")
      'Add By Sindy 2017/4/27
      If Not IsNull(rsTmp.Fields("TMBM09")) Then: textTMBM09 = rsTmp.Fields("TMBM09")
      If Not IsNull(rsTmp.Fields("TMBM10")) Then: textTMBM10 = rsTmp.Fields("TMBM10")
      If Not IsNull(rsTmp.Fields("TMBM11")) Then: textTMBM11 = rsTmp.Fields("TMBM11")
      If Not IsNull(rsTmp.Fields("TMBM12")) Then: textTMBM12 = rsTmp.Fields("TMBM12")
      If Not IsNull(rsTmp.Fields("TMBM13")) Then: textTMBM13 = rsTmp.Fields("TMBM13")
      If Not IsNull(rsTmp.Fields("TMBM14")) Then: textTMBM14 = rsTmp.Fields("TMBM14")
      If Not IsNull(rsTmp.Fields("TMBM15")) Then: textTMBM15 = rsTmp.Fields("TMBM15")
      If Not IsNull(rsTmp.Fields("TMBM16")) Then: textTMBM16 = rsTmp.Fields("TMBM16")
      If Not IsNull(rsTmp.Fields("TMBM17")) Then: textTMBM17 = rsTmp.Fields("TMBM17")
      If Not IsNull(rsTmp.Fields("TMBM18")) Then: textTMBM18 = rsTmp.Fields("TMBM18")
      '2017/4/27 END
      
      UpdateFieldOldData rsTmp
      
      ' 更新控制項帶出的相關內容
      'textTMBM02_Validate False
      textTMBM02_LostFocus
      'textNA01 = textTMBM05
      'textNA01_Validate False
      If IsEmptyText(textTMBM05) = False Then
         textNA01 = GetNationNo(textTMBM05)
      End If
      'textTA02 = textTMBM06
      'textTA02_Validate False
      If IsEmptyText(textTMBM06) = False Then
         textTA02 = GetTAgentNo(textTMBM06)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strTMBM01 As String, ByVal strTMBM02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strTMBM01, strTMBM02) = True Then
      m_CurrTM(0) = strTMBM01
      m_CurrTM(1) = strTMBM02
   Else
      strSql = "SELECT * FROM TMBULLETIN " & _
               "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                     "TMBM02 = (SELECT MIN(TMBM02) FROM TMBULLETIN " & _
                               "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                                     "TMBM02 > '" & m_CurrTM(1) & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
         If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT * FROM TMBULLETIN " & _
               "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN " & _
                               "WHERE TMBM01 > '" & m_CurrTM(0) & "') AND " & _
                     "TMBM02 = (SELECT MIN(TMBM02) FROM TMBULLETIN " & _
                               "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN " & _
                                               "WHERE TMBM01 > '" & m_CurrTM(0) & "')) "
                               
                     
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
         If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
      Else
         rsTmp.Close
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
      UpdateCtrlData
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrTM(0) = m_FirstTM(0)
   m_CurrTM(1) = m_FirstTM(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTM(0) = m_FirstTM(0) And m_CurrTM(1) = m_FirstTM(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TMBM02 = (SELECT MAX(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                                  "TMBM02 < '" & m_CurrTM(1) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = (SELECT MAX(TMBM01) FROM TMBULLETIN " & _
                            "WHERE TMBM01 < '" & m_CurrTM(0) & "') AND " & _
                  "TMBM02 = (SELECT MAX(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = (SELECT MAX(TMBM01) FROM TMBULLETIN " & _
                                            "WHERE TMBM01 < '" & m_CurrTM(0) & "')) "
                            
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示下一筆資料
Private Sub ShowNextRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrTM(0) = m_LastTM(0) And m_CurrTM(1) = m_LastTM(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                  "TMBM02 = (SELECT MIN(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = '" & m_CurrTM(0) & "' AND " & _
                                  "TMBM02 > '" & m_CurrTM(1) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN " & _
                            "WHERE TMBM01 > '" & m_CurrTM(0) & "') AND " & _
                  "TMBM02 = (SELECT MIN(TMBM02) FROM TMBULLETIN " & _
                            "WHERE TMBM01 = (SELECT MIN(TMBM01) FROM TMBULLETIN " & _
                                            "WHERE TMBM01 > '" & m_CurrTM(0) & "')) "
                            
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("TMBM01")) = False Then: m_CurrTM(0) = rsTmp.Fields("TMBM01")
      If IsNull(rsTmp.Fields("TMBM02")) = False Then: m_CurrTM(1) = rsTmp.Fields("TMBM02")
   End If
   rsTmp.Close
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrTM(0) = m_LastTM(0)
   m_CurrTM(1) = m_LastTM(1)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         tlbar.Buttons(1).Enabled = True
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(3).Enabled = True
         tlbar.Buttons(4).Enabled = True
         tlbar.Buttons(6).Enabled = True
         tlbar.Buttons(7).Enabled = True
         tlbar.Buttons(8).Enabled = True
         tlbar.Buttons(9).Enabled = True
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_EditMode = 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyCode = 0
            OnAction vbKeyF9
            KeyCode = 0
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
            KeyCode = 0
         Else
            OnAction vbKeyF10
            KeyCode = 0
         End If
   End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         'textTMBM01 = GetMaxTMBM01()
         OnPrepareAddRecord
         If IsRecordExist(textTMBM01, textTMBM02) = True Then
            strTit = "檢核資料"
            strMsg = "該筆記錄已經存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM01.SetFocus
         Else
            SetInputEntry
         End If
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
         textNA01.SetFocus 'Add By Sindy 2017/8/15
      ' 刪除
      Case vbKeyF5:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            OnWork
            UpdateToolbarState
         End If
      ' 查詢
      Case vbKeyF4:
         m_EditMode = 4
         SetCtrlReadOnly True
         SetKeyReadOnly False
         ClearField
         textTMBM02 = "1" 'Add By Sindy 2014/8/18 預設為1
         UpdateToolbarState
         SetInputEntry
      ' 第一筆
      Case vbKeyHome:
         ShowFirstRecord
      ' 前一筆
      Case vbKeyPageUp:
         ShowPrevRecord
      ' 後一筆
      Case vbKeyPageDown:
         ShowNextRecord
      ' 最後一筆
      Case vbKeyEnd:
         ShowLastRecord
      ' 確定
      Case vbKeyF9:
'         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         UpdateFieldNewData
         OnWork
         UpdateToolbarState
      ' 取消
      Case vbKeyF10:
         Select Case m_EditMode
            Case 1, 2:
               strTit = "詢問"
               strMsg = "你並未存檔, 確定離開嗎?"
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
               If nResponse = vbYes Then
                  m_EditMode = 0
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030602 = Nothing
End Sub

Private Sub textNA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTMBM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2014/7/7
Private Sub textTMBM01_Validate(Cancel As Boolean)
   If textTMBM01 <> "" Then
      If textTMBM02 = "" Then
         textTMBM02 = "1"
         Call textTMBM02_LostFocus
      End If
   End If
End Sub

Private Sub textTMBM04_LostFocus()
'add by nickc 2005/04/15 以申請案號抓tm  且 tm10='000'，tm28='1' 將代理人自動上'01'
If Trim(textTMBM04) <> "" Then
      Dim tmpRs As New ADODB.Recordset
      Set tmpRs = New ADODB.Recordset
      tmpRs.CursorLocation = adUseClient
      tmpRs.Open "select count(*) from trademark where tm10='000' and tm28='1' and tm12='" & textTMBM04.Text & "' ", cnnConnection, adOpenStatic, adLockReadOnly
      If tmpRs.Fields(0).Value <> 0 Then
         textTA02 = "01"
         Call textTA02_Validate(False)
      End If
      Set tmpRs = Nothing
End If
End Sub

Private Sub textTMBM04_Validate(Cancel As Boolean)
   'add by nickc 2007/01/04
   If Trim(textTMBM04) <> "" Then
       If m_EditMode = 1 Then
           CheckOC3
           strSql = "select * from tmbulletin where tmbm04='" & textTMBM04 & "' "
           AdoRecordSet3.CursorLocation = adUseClient
           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If AdoRecordSet3.RecordCount <> 0 Then
               MsgBox "申請案號重複！" & vbCrLf & "審定號：" & CheckStr(AdoRecordSet3.Fields("tmbm01")), vbOKOnly, "錯誤！"
               textTMBM04.SetFocus
               CheckOC3
               Cancel = True
           End If
           CheckOC3
       ElseIf m_EditMode = 2 Then
           CheckOC3
           strSql = "select * from tmbulletin where tmbm04='" & textTMBM04.Text & "' and tmbm01||tmbm02<>'" & textTMBM01.Text & "'||'" & textTMBM02.Text & "' "
           AdoRecordSet3.CursorLocation = adUseClient
           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
           If AdoRecordSet3.RecordCount <> 0 Then
               MsgBox "申請案號重複！" & vbCrLf & "審定號：" & CheckStr(AdoRecordSet3.Fields("tmbm01")), vbOKOnly, "錯誤！"
               textTMBM04.SetFocus
               CheckOC3
               Cancel = True
           End If
           CheckOC3
       End If
      
      '2010/9/14 ADD BY SONIA 檢查基本檔若已有審定號且與textTMBM01不同則提醒操作者
      CheckOC3
      strSql = "select * from trademark where tm10='000' and tm28='1' and tm12='" & textTMBM04.Text & "' "
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount <> 0 Then
         If "" & AdoRecordSet3.Fields("TM15") <> "" Then
            If "" & AdoRecordSet3.Fields("TM15") <> textTMBM01 Then
               MsgBox "基本檔審定號：" & CheckStr(AdoRecordSet3.Fields("TM15")) & ", 與公報不符, 請再確認 !", vbOKOnly, "錯誤！"
               textTMBM04.SetFocus
               CheckOC3
               Cancel = True
            End If
         End If
      End If
      '2010/9/14 END
   End If
End Sub

' 地區名稱
Private Sub textTMBM05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTMBM05) = False Then
      If StrLength(textTMBM05) > 20 Then
         Cancel = True
         textNA01 = Empty
         strTit = "檢核資料"
         strMsg = "地區名稱內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM05_GotFocus
         GoTo EXITSUB
      End If
      
      strTemp = GetNationNo(textTMBM05)
      If IsEmptyText(strTemp) = True Then
         strTemp = GetNationNo(textTMBM05 & "縣")
         If IsEmptyText(strTemp) = False Then
            textTMBM05 = textTMBM05 & "縣"
         End If
      End If
      If IsEmptyText(strTemp) = True Then
         Cancel = True
         textNA01 = Empty
         strTit = "檢核資料"
         strMsg = "地區名稱不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM05_GotFocus
         GoTo EXITSUB
      Else
         textNA01 = strTemp
      End If
   End If
EXITSUB:
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTMBM05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 代理人名稱
Private Sub textTMBM06_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strFreeAgentCode As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTMBM06) = False Then
      If StrLength(textTMBM06) > 12 Then
         Cancel = True
         textTA02 = Empty
         strTit = "檢核資料"
         strMsg = "代理人名稱太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM06_GotFocus
         GoTo EXITSUB
      End If
      strTemp = GetTAgentNo(textTMBM06)
      'If IsEmptyText(strTemp) = True Then
      '   Cancel = True
      '   textTA02 = Empty
      '   strTit = "檢核資料"
      '   strMsg = "代理人名稱不存在"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textTMBM06_GotFocus
      '   GoTo ExitSub
      'Else
         textTA02 = strTemp
      'End If
   End If
   
   'Add By Sindy 2010/02/02
   ' 代理人是鍵入名稱時
   If IsEmptyText(textTA02) = True And IsEmptyText(textTMBM06) = False Then
      strSql = "SELECT * FROM TAGENT " & _
                    "WHERE TA01 = 'T' AND " & _
                        "TA03 = '" & Trim(textTMBM06) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("TA02")) = False Then
            textTA02 = rsTmp.Fields("TA02")
         End If
      Else
         On Error GoTo ErrorHandler
         cnnConnection.BeginTrans
         strTit = "代理人"
         strFreeAgentCode = GetFreeAgentCode
         strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                        "　　　　　代理人名稱 <" & textTMBM06 & "> "
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
         If nResponse = vbYes Then
            strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('T','" & strFreeAgentCode & "','" & textTMBM06 & "','" & textTMBM06 & "')"
            cnnConnection.Execute strSql
            textTA02 = strFreeAgentCode
            ' 儲存公告日
            If IsEmptyText(textTMBM07) = False Then
               strTemp = GetTA05
               strSql = "UPDATE TAgent SET TA05 = " & DBDATE(strTemp) & " " & _
                            "WHERE TA01 = 'T' AND " & _
                                 "TA02 = '" & strFreeAgentCode & "' "
               cnnConnection.Execute strSql
            End If
            cnnConnection.CommitTrans
         '若不新增
         Else
            'Cancel = True
            cnnConnection.RollbackTrans
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing
   '2010/02/02 End
   
EXITSUB:
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textTMBM06.IMEMode = 2
   If Cancel = False Then CloseIme
   'Add By Sindy 2010/02/02
   If Cancel = True Then textTMBM06_GotFocus
   Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "(" & Err.Number & ")" & Err.Description
    '2010/02/02 End
End Sub

'Add By Sindy 2010/02/02
Private Function GetTA05() As String
Dim strTemp As String
   
   GetTA05 = CStr(Val(Left(Trim(textTMBM07), 2)) + 62)
   Select Case CStr(Right(Trim(textTMBM07), 2))
   Case "01"
      strTemp = "0101"
   Case "02"
      strTemp = "0116"
   Case "03"
      strTemp = "0201"
   Case "04"
      strTemp = "0216"
   Case "05"
      strTemp = "0301"
   Case "06"
      strTemp = "0316"
   Case "07"
      strTemp = "0401"
   Case "08"
      strTemp = "0416"
   Case "09"
      strTemp = "0501"
   Case "10"
      strTemp = "0516"
   Case "11"
      strTemp = "0601"
   Case "12"
      strTemp = "0616"
   Case "13"
      strTemp = "0701"
   Case "14"
      strTemp = "0716"
   Case "15"
      strTemp = "0801"
   Case "16"
      strTemp = "0816"
   Case "17"
      strTemp = "0901"
   Case "18"
      strTemp = "0916"
   Case "19"
      strTemp = "1001"
   Case "20"
      strTemp = "1016"
   Case "21"
      strTemp = "1101"
   Case "22"
      strTemp = "1116"
   Case "23"
      strTemp = "1201"
   Case "24"
      strTemp = "1216"
   End Select
   GetTA05 = GetTA05 & strTemp
End Function

'Add By Sindy 2010/02/02
Private Function GetFreeAgentCode() As String
   Dim strLastAgent As String
   Dim nNumber As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strLastAgent = "01"
   strSql = "SELECT * FROM TAgent WHERE TA01 = 'T'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("TA02")) = False Then
            If Val(rsTmp.Fields("TA02")) > Val(strLastAgent) Then
               strLastAgent = rsTmp.Fields("TA02")
            End If
         End If
         rsTmp.MoveNext
      Loop
   End If
   
   nNumber = Val(strLastAgent) + 1
   Select Case Len(strLastAgent)
      Case 1:
         If Len(nNumber) > 1 Then
            GetFreeAgentCode = Format(nNumber, "00")
         Else
            GetFreeAgentCode = Format(nNumber, "0")
         End If
      Case 2:
         If Len(nNumber) > 2 Then
            GetFreeAgentCode = Format(nNumber, "000")
         Else
            GetFreeAgentCode = Format(nNumber, "00")
         End If
      Case 3:
         GetFreeAgentCode = Format(nNumber, "000")
      Case Else
         GetFreeAgentCode = nNumber
   End Select
   
   Set rsTmp = Nothing
End Function

' 公報卷期
Private Sub textTMBM07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rs1 As New ADODB.Recordset
   
   Cancel = False
   If IsEmptyText(textTMBM07) = False Then
      If IsNumeric(textTMBM07) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公報卷期只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_GotFocus
         Exit Sub
      End If
      
      'Add By Cheng 2002/01/22
      '自動帶出公報卷期的最大審定號數加一
      If m_EditMode = 1 Then
         If rs1.State <> adStateClosed Then rs1.Close
         rs1.CursorLocation = adUseClient
         rs1.Open "Select Max(TMBM01) From TMBulletin Where TMBM07 = '" & Me.textTMBM07.Text & "'", _
                  cnnConnection, adOpenStatic, adLockReadOnly
         If rs1.EOF Then
            Me.textTMBM01.Text = Format(1, "00000000")
         Else
            'Modify By Cheng 2002/11/05
'            Me.textTMBM01.Text = Format(rs1.Fields(0).Value + 1, "0000000")
            Me.textTMBM01.Text = Format(Val(rs1.Fields(0).Value) + 1, "00000000")
         End If
         If rs1.State <> adStateClosed Then rs1.Close
         Set rs1 = Nothing
      End If
   End If
   
End Sub

Private Sub textTMBM08_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeyReturn Then
      KeyAscii = 0
   End If
End Sub

' 商品類別
Private Sub textTMBM08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim strTemp As String
   Dim nPos As Integer
   Dim nChar As String
   Cancel = False
   
   ' 無資料時不做任何檢查
   If IsEmptyText(textTMBM08) = True Then
      GoTo EXITSUB
   End If
   
   ' 刪除非文字的字元
   strTemp = Empty
   For nPos = 1 To Len(textTMBM08)
      If Mid(textTMBM08, nPos, 1) >= " " Then
         strTemp = strTemp & Mid(textTMBM08, nPos, 1)
      End If
   Next nPos
   textTMBM08 = strTemp
 
   nCount = GetSubStringCount(textTMBM08)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTMBM08, nIndex)
        'Modify By Cheng 2002/12/30
        '商品類別只能輸二碼
'      If Len(strTemp) < 3 Or Len(strTemp) > 6 Then
      If Len(strTemp) <> 2 Then
         Cancel = True
         strTit = "檢核資料"
            'Modify By Cheng 2002/12/30
'         strMsg = "商品類別<" & strTemp & ">不正確"
         strMsg = "商品類別<" & strTemp & ">不正確，只能輸兩碼"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM08_GotFocus
         GoTo EXITSUB
      End If
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textTMBM08, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textTMBM08, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "商品類別<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTMBM08_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Select Case Button.Index
      ' 新增
      Case 1: OnAction vbKeyF2
      ' 修改
      Case 2: OnAction vbKeyF3
      ' 刪除
      Case 3: OnAction vbKeyF5
      ' 查詢
      Case 4: OnAction vbKeyF4
      ' 第一筆
      Case 6: OnAction vbKeyHome
      ' 前一筆
      Case 7: OnAction vbKeyPageUp
      ' 後一筆
      Case 8: OnAction vbKeyPageDown
      ' 最後一筆
      Case 9: OnAction vbKeyEnd
      ' 確定
      Case 11: OnAction vbKeyF9
      ' 取消
      Case 12: OnAction vbKeyF10
      ' 離開
      Case 14: OnAction vbKeyEscape
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strTMBM01 As String, ByVal strTMBM02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM TMBULLETIN " & _
            "WHERE TMBM01 = '" & strTMBM01 & "' AND " & _
                  "TMBM02 = '" & strTMBM02 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsRecordExist = True
   Else
      IsRecordExist = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strTMBM01 As String
   Dim strTMBM02 As String
   
   strTMBM01 = textTMBM01
   strTMBM02 = textTMBM02
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strTMBM01, strTMBM02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO TMBULLETIN ("
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         strTmp = m_FieldList(nIndex).fiName
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            strTmp = "'" & m_FieldList(nIndex).fiNewData & "'"
         Else
            strTmp = m_FieldList(nIndex).fiNewData
         End If
      End If
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ")"
   
   cnnConnection.Execute strSql
   
   ' 更新新增的資料暫存區
   m_LastTMBM01 = textTMBM01
   m_LastTMBM02 = textTMBM02
   m_LastTMBM07 = textTMBM07
   
   If ((strTMBM01 & strTMBM02) < (m_FirstTM(0) & m_FirstTM(1))) Or ((strTMBM01 & strTMBM02) > (m_LastTM(0) & m_LastTM(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strTMBM01, strTMBM02
EXITSUB:
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strTMBM01 As String
   Dim strTMBM02 As String
   
   strTMBM01 = m_CurrTM(0)
   strTMBM02 = m_CurrTM(1)
   
   strSql = "UPDATE TMBULLETIN SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To MAX_FIELD - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = '" & m_FieldList(nIndex).fiNewData & "'"
            End If
         Else
            If m_FieldList(nIndex).fiNewData = Empty Then
               strTmp = m_FieldList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_FieldList(nIndex).fiName & " = " & m_FieldList(nIndex).fiNewData
            End If
         End If
      End If
      If strTmp <> Empty Then
         bDifference = True
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   
   strSql = strSql & " " & _
            "WHERE TMBM01 = '" & strTMBM01 & "' AND " & _
                  "TMBM02 = '" & strTMBM02 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      
      'Add By Sindy 2017/4/26 更新國內商標公報開拓定稿商標權人檔
      strExc(0) = "select * from tmbulletindata" & _
                  " where tbd01='" & textTMBM07 & "'" & _
                  " and tbd02='" & textTMBM01 & "'" & _
                  " and tbd03='" & textTMBM02 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If textTMBM09 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM09) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=1 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM10 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM10) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=2 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM11 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM11) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=3 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM12 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM12) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=4 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM13 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM13) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=5 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM14 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM14) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=6 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM15 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM15) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=7 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM16 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM16) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=8 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM17 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM17) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=9 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
         If textTMBM18 <> "" Then
            strSql = "UPDATE tmbulletinOwner SET TBOR03='" & ChgSQL(textTMBM18) & "'" & _
                     " WHERE TBOR01 = '" & textTMBM01 & "' AND TBOR02=10 AND TBOR06 = '" & textTMBM02 & "'"
            cnnConnection.Execute strSql
         End If
      End If
      '2017/4/26 END
      
      ShowCurrRecord strTMBM01, strTMBM02
   End If

End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strTMBM01 As String
   Dim strTMBM02 As String
   
   strTMBM01 = m_CurrTM(0)
   strTMBM02 = m_CurrTM(1)
   
   strSql = "DELETE FROM TMBULLETIN " & _
            "WHERE TMBM01 = '" & strTMBM01 & "' AND " & _
                  "TMBM02 = '" & strTMBM02 & "' "
                  
   cnnConnection.Execute strSql
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strTMBM01 = m_LastTM(0) And strTMBM02 = m_LastTM(1)) Or (strTMBM01 = m_FirstTM(0) And strTMBM02 = m_FirstTM(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strTMBM01, strTMBM02
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSql As String
   Dim nIndex As Integer
   Dim nPos As Integer
   Dim bFind As Boolean
   Dim strTMBM01 As String
   Dim strTMBM02 As String
   
   strTMBM01 = textTMBM01
   strTMBM02 = textTMBM02
   
   If IsRecordExist(strTMBM01, strTMBM02) = True Then
      m_CurrTM(0) = strTMBM01
      m_CurrTM(1) = strTMBM02
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If
   
   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Select Case m_EditMode
      Case 1:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            UpdateFieldNewData
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   'Modify By Sindy 2014/7/16 桂英說修改完不要自動帶下一筆資料
   m_EditMode = 0
   SetCtrlReadOnly True
'   'Modify By Cheng 2002/01/22
'   '修改存檔時, 自動帶下一筆資料
'   If m_EditMode <> 1 And m_EditMode <> 2 Then
'      m_EditMode = 0
'      SetCtrlReadOnly True
'   ElseIf m_EditMode = 2 Then
'      ShowNextRecord
'      OnAction vbKeyF3
'   Else
'      OnAction vbKeyF2
'      'ClearField
'      'OnPrepareAddRecord
'      'If IsRecordExist(textTMBM01, textTMBM02) = True Then
'      '   strTit = "檢核資料"
'      '   strMsg = "該筆資料已經存在"
'      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      '   textTMBM01.SetFocus
'      'Else
'      '   SetInputEntry
'      'End If
'   End If
   '2014/7/16 END
EXITSUB:
End Sub

' 商標種類
Private Sub textTMBM02_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   textTMBM02_2 = Empty
   If IsEmptyText(textTMBM02) = False Then
      textTMBM02_2 = GetTradeMarkName(textTMBM02, 0)
      If IsEmptyText(textTMBM02_2) = True Then
         Select Case m_EditMode
            Case 1, 4:
               strTit = "檢核資料"
               strMsg = "商標種類不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               'textTMBM02_GotFocus
               textTMBM02.SetFocus
               GoTo EXITSUB
         End Select
      End If
      If m_EditMode = 1 Then
         If IsEmptyText(textTMBM01) = False Then
            If IsRecordExist(textTMBM01, textTMBM02) = True Then
               strTit = "檢核資料"
               strMsg = "該筆資料已經存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTMBM01.SetFocus
               GoTo EXITSUB
            End If
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub textNA01_Change()
   If IsEmptyText(textNA01) = False Then
      textTMBM05.TabStop = False
   Else
      textTMBM05.TabStop = True
   End If
End Sub

' 地區別
Private Sub textNA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   
   If IsEmptyText(textNA01) = False Then
      textTMBM05.TabStop = False
      If textNA01 <= "010" Then
         textTMBM05 = Empty
         strTit = "檢核資料"
         strMsg = "地區別不正確"
         Cancel = True
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNA01.SetFocus
      End If
      strTemp = GetNationName(textNA01, 0)
      If IsEmptyText(strTemp) = False Then
         textTMBM05 = strTemp
      Else
         Select Case m_EditMode
            Case 1, 2:
               textTMBM05 = Empty
               strTit = "檢核資料"
               strMsg = "地區名稱不存在"
               Cancel = True
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNA01.SetFocus
         End Select
      End If
   Else
      textTMBM05.TabStop = True
   End If
End Sub

Private Sub textTA02_Change()
   If IsEmptyText(textTA02) = False Then
      textTMBM06.TabStop = False
   Else
      textTMBM06.TabStop = True
   End If
End Sub

' 代理人編號
Private Sub textTA02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   
   If IsEmptyText(textTA02) = False Then
      'EnableTextBox textTMBM06, False
      textTMBM06.TabStop = False
      strTemp = Empty
      If IsNumeric(textTA02) = True Then
         strTemp = GetTAgentName(textTA02)
         If IsEmptyText(strTemp) = False Then
            textTMBM06 = strTemp
         Else
            Select Case m_EditMode
               Case 1, 2:
                  textTMBM06 = Empty
                  strTit = "檢核資料"
                  strMsg = "代理人編號不存在"
                  Cancel = True
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTA02.SetFocus
            End Select
         End If
      Else
         strTemp = GetTAgentNo(textTA02)
         If IsEmptyText(strTemp) = False Then
            textTMBM06 = textTA02
            textTA02 = strTemp
         Else
            Select Case m_EditMode
               Case 1, 2:
                  textTMBM06 = Empty
                  strTit = "檢核資料"
                  strMsg = "代理人編號不存在"
                  Cancel = True
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  'textTA02_GotFocus
                  textTA02.SetFocus
            End Select
         End If
      End If
   Else
      textTMBM06.TabStop = True
   End If
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1:
         If IsEmptyText(textTMBM01) Then
            textTMBM01.SetFocus
         ElseIf IsEmptyText(textTMBM02) Then
            textTMBM02.SetFocus
         ElseIf IsEmptyText(textTMBM03) Then
            textTMBM03.SetFocus
         Else
            textTMBM04.SetFocus
         End If
      Case 2: textTMBM07.SetFocus
      Case 4: textTMBM01.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTemp As String
Dim nPos As Integer
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strFreeAgentCode As String
   
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2:
         ' 刪除非文字的字元
         strTemp = Empty
         For nPos = 1 To Len(textTMBM08)
            If Mid(textTMBM08, nPos, 1) >= " " Then
               strTemp = strTemp & Mid(textTMBM08, nPos, 1)
            End If
         Next nPos
         textTMBM08 = strTemp
         
         ' 審定號不可空白
         If IsEmptyText(textTMBM01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入審定號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM01.SetFocus
            GoTo EXITSUB
         End If
         ' 商標種類不可空白
         If IsEmptyText(textTMBM02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入商標種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM02.SetFocus
            GoTo EXITSUB
         End If
         ' 公報卷期
         If IsEmptyText(textTMBM07) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入公報卷期"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM07.SetFocus
            GoTo EXITSUB
         End If
         ' 商標種類為聯合商標, 防護商標, 聯合服務標章, 防護服務標章時正商標號數不可空白
         If IsEmptyText(textTMBM03) = True Then
            Select Case textTMBM02
               Case "2", "3", "5", "6":
                  strTit = "檢核資料"
                  strMsg = "請輸入正商標號數"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textTMBM03.SetFocus
                  GoTo EXITSUB
               Case Else:
            End Select
         End If
         ' 申請案號
         If IsEmptyText(textTMBM04) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入申請案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTMBM04.SetFocus
            GoTo EXITSUB
         End If
         ' 地區編號
         If IsEmptyText(textTMBM05) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入地區"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNA01.SetFocus
            GoTo EXITSUB
         End If
         ' 商品類別 (非團體標章及證明標章時不可空白)
        'Modify By Cheng 2003/01/07
'         If IsEmptyText(textTMBM08) = True Or Len(textTMBM08) < 3 Then
         If IsEmptyText(textTMBM08) = True Or Len(textTMBM08) < 2 Then
            Select Case textTMBM02
               Case "7", "8":
               Case Else
                    'Modify By Cheng 2003/01/07
                    If Me.textTMBM08.Text = "" Then
                        strTit = "檢核資料"
                        strMsg = "請輸入商品類別"
                        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                        textTMBM08.SetFocus
                        GoTo EXITSUB
                    Else
                        strTit = "檢核資料"
                        strMsg = "商品類別輸入錯誤"
                        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                        textTMBM08.SetFocus
                        GoTo EXITSUB
                    End If
            End Select
         End If
         
         'Add By Sindy 2010/02/02
         ' 代理人
         If IsEmptyText(textTA02) = False Then
            If IsEmptyText(GetTAgentName(Trim(textTA02))) = True Then
               strMsg = "無此代理人資料"
               strTit = "錯誤"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTA02.SetFocus
               GoTo EXITSUB
            End If
         End If
         If IsEmptyText(textTA02) = True And IsEmptyText(textTMBM06) = False Then
            Set rsTmp = New ADODB.Recordset
            strSql = "SELECT * FROM TAGENT " & _
                          "WHERE TA01 = 'T' AND " & _
                               "TA03 = '" & Trim(textTMBM06) & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("TA02")) = False Then
                  textTA02 = rsTmp.Fields("TA02")
               End If
            Else
On Error GoTo ErrorHandler
               cnnConnection.BeginTrans
               strTit = "代理人"
               strFreeAgentCode = GetFreeAgentCode
               strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                              "　　　　　代理人名稱 <" & textTMBM06 & "> "
               nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
               If nResponse = vbYes Then
                  strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('T','" & strFreeAgentCode & "','" & textTMBM06 & "','" & textTMBM06 & "')"
                  cnnConnection.Execute strSql
                  textTA02 = strFreeAgentCode
                  ' 儲存公告日
                  If IsEmptyText(textTMBM07) = False Then
                     strTemp = GetTA05
                     strSql = "UPDATE TAgent SET TA05 = " & DBDATE(strTemp) & " " & _
                              "WHERE TA01 = 'T' AND " & _
                                    "TA02 = '" & strFreeAgentCode & "' "
                     cnnConnection.Execute strSql
                  End If
                  cnnConnection.CommitTrans
               '若不新增
               Else
                  cnnConnection.RollbackTrans
                  rsTmp.Close
                  Set rsTmp = Nothing
                  GoTo EXITSUB
               End If
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         End If
         '2010/02/02 End
      Case Else:
   End Select
   
   CheckDataValid = True
   
EXITSUB:

   Exit Function
ErrorHandler:
   cnnConnection.RollbackTrans
   MsgBox "(" & Err.Number & ")" & Err.Description
   
End Function

Private Sub textTMBM01_GotFocus()
   InverseTextBox textTMBM01
End Sub

Private Sub textTMBM02_GotFocus()
   InverseTextBox textTMBM02
End Sub

Private Sub textTMBM03_GotFocus()
   InverseTextBox textTMBM03
End Sub

Private Sub textTMBM04_GotFocus()
   InverseTextBox textTMBM04
End Sub

Private Sub textTMBM05_GotFocus()
   InverseTextBox textTMBM05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTMBM05.IMEMode = 1
   OpenIme
End Sub

Private Sub textTMBM06_GotFocus()
   InverseTextBox textTMBM06
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textTMBM06.IMEMode = 1
   OpenIme
End Sub

Private Sub textTMBM07_GotFocus()
   InverseTextBox textTMBM07
End Sub

Private Sub textTMBM08_GotFocus()
   InverseTextBox textTMBM08
End Sub

'Add By Sindy 2017/4/27
Private Sub textTMBM09_GotFocus()
   OpenIme
   InverseTextBox textTMBM09
End Sub
'Modify by Amy 2022/01/10 原:Integer,+textTMBM09
Private Sub textTMBM09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM09)
End Sub
Private Sub textTMBM09_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM09.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM09, textTMBM09.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM10_GotFocus()
   OpenIme
   InverseTextBox textTMBM10
End Sub

'Modify by Amy 原:Integer +textTMBM10
Private Sub textTMBM10_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM10)
End Sub
Private Sub textTMBM10_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM10.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM10, textTMBM10.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM11_GotFocus()
   OpenIme
   InverseTextBox textTMBM11
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM11
Private Sub textTMBM11_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM11)
End Sub
Private Sub textTMBM11_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM11.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM11, textTMBM11.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM12_GotFocus()
   OpenIme
   InverseTextBox textTMBM12
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM12
Private Sub textTMBM12_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM12)
End Sub
Private Sub textTMBM12_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM12.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM12, textTMBM12.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM13_GotFocus()
   OpenIme
   InverseTextBox textTMBM13
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM13
Private Sub textTMBM13_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM13)
End Sub

Private Sub textTMBM13_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM13.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM13, textTMBM13.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM14_GotFocus()
   OpenIme
   InverseTextBox textTMBM14
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM14
Private Sub textTMBM14_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM14)
End Sub
Private Sub textTMBM14_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM14.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM14, textTMBM14.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM15_GotFocus()
   OpenIme
   InverseTextBox textTMBM15
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM15
Private Sub textTMBM15_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM15)
End Sub
Private Sub textTMBM15_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM15.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM15, textTMBM15.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM16_GotFocus()
   OpenIme
   InverseTextBox textTMBM16
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM16
Private Sub textTMBM16_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM16)
End Sub
Private Sub textTMBM16_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM16.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM16, textTMBM16.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM17_GotFocus()
   OpenIme
   InverseTextBox textTMBM17
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM17
Private Sub textTMBM17_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM17)
End Sub
Private Sub textTMBM17_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM17.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM17, textTMBM17.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
Private Sub textTMBM18_GotFocus()
   OpenIme
   InverseTextBox textTMBM18
End Sub

'Modify by Amy 2022/01/10 原:Integer,+textTMBM18
Private Sub textTMBM18_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii, textTMBM18)
End Sub
Private Sub textTMBM18_Validate(Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   If textTMBM18.Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(textTMBM18, textTMBM18.MaxLength - 1) Then
      Cancel = True
   End If
End Sub
'2017/4/27 END

Private Sub textNA01_GotFocus()
   InverseTextBox textNA01
End Sub

Private Sub textTA02_GotFocus()
   InverseTextBox textTA02
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

'Add by Amy 2022/01/10檢查畫面的 TextBox是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If

'Add By Sindy 2017/8/15
If IsEmptyText(textNA01) = False Then
   Cancel = False
   textNA01_Validate Cancel
   If Cancel = True Then
      textNA01.SetFocus
      Exit Function
   End If
End If
If IsEmptyText(textTA02) = False Then
   Cancel = False
   textTA02_Validate Cancel
   If Cancel = True Then
      textTA02.SetFocus
      Exit Function
   End If
End If
'2017/8/15 END

TxtValidate = False
If Me.textTMBM05.Enabled = True Then
   Cancel = False
   textTMBM05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTMBM06.Enabled = True Then
   Cancel = False
   textTMBM06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Modify By Cheng 2002/11/05
'If Me.textTMBM07.Enabled = True Then
'   Cancel = False
'   textTMBM07_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If

If Me.textTMBM08.Enabled = True Then
   Cancel = False
   textTMBM08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'add by nickc 2006/06/22 加入檢查申請案號不可以重複
'edit by nickc 2007/01/04
'If m_EditMode = 1 Then
'    CheckOC3
'    strSQL = "select * from tmbulletin where tmbm04='" & textTMBM04 & "' "
'    AdoRecordSet3.CursorLocation = adUseClient
'    AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If AdoRecordSet3.RecordCount <> 0 Then
'        MsgBox "申請案號重複！" & vbCrLf & "審定號：" & CheckStr(AdoRecordSet3.Fields("tmbm01")), vbOKOnly, "錯誤！"
'        textTMBM04.SetFocus
'        CheckOC3
'        Exit Function
'    End If
'    CheckOC3
'ElseIf m_EditMode = 2 Then
'    CheckOC3
'    strSQL = "select * from tmbulletin where tmbm04='" & textTMBM04.Text & "' and tmbm01||tmbm02<>'" & textTMBM01.Text & "'||'" & textTMBM02.Text & "' "
'    AdoRecordSet3.CursorLocation = adUseClient
'    AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'    If AdoRecordSet3.RecordCount <> 0 Then
'        MsgBox "申請案號重複！" & vbCrLf & "審定號：" & CheckStr(AdoRecordSet3.Fields("tmbm01")), vbOKOnly, "錯誤！"
'        textTMBM04.SetFocus
'        CheckOC3
'        Exit Function
'    End If
'    CheckOC3
'End If
If Me.textTMBM04.Enabled = True Then
   Cancel = False
   textTMBM04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
