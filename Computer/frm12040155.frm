VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040155 
   BorderStyle     =   1  '單線固定
   Caption         =   "非本所實質客戶資料維護"
   ClientHeight    =   5745
   ClientLeft      =   105
   ClientTop       =   930
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9047.999
   Begin VB.TextBox textNC01_1 
      Height          =   264
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   19
      Text            =   "CR"
      Top             =   840
      Width           =   345
   End
   Begin VB.TextBox textNC01 
      Height          =   264
      Left            =   1920
      MaxLength       =   4
      TabIndex        =   0
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   60
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
            Picture         =   "frm12040155.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040155.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   660
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   1164
      ButtonWidth     =   1138
      ButtonHeight    =   1111
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
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
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7620
         Top             =   120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8040
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin MSForms.TextBox textNC05 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Top             =   1740
      Width           =   3360
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC02 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1140
      Width           =   7125
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "12568;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC03 
      Height          =   285
      Left            =   1560
      TabIndex        =   2
      Top             =   1440
      Width           =   3360
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC16 
      Height          =   705
      Left            =   1560
      TabIndex        =   7
      Top             =   2325
      Width           =   7140
      VariousPropertyBits=   -1466941413
      MaxLength       =   500
      ScrollBars      =   2
      Size            =   "12594;1244"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC10 
      Height          =   285
      Left            =   5295
      TabIndex        =   10
      Top             =   3360
      Width           =   3390
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5980;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC11 
      Height          =   285
      Left            =   1560
      TabIndex        =   11
      Top             =   3660
      Width           =   3360
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC12 
      Height          =   285
      Left            =   5295
      TabIndex        =   12
      Top             =   3660
      Width           =   3390
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5980;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC13 
      Height          =   285
      Left            =   1560
      TabIndex        =   13
      Top             =   3960
      Width           =   3360
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC15 
      Height          =   285
      Left            =   1560
      TabIndex        =   15
      Top             =   4260
      Width           =   7125
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "12568;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC09 
      Height          =   285
      Left            =   1560
      TabIndex        =   9
      Top             =   3360
      Width           =   3360
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5927;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC08 
      Height          =   285
      Left            =   1590
      TabIndex        =   8
      Top             =   3060
      Width           =   7095
      VariousPropertyBits=   671105051
      MaxLength       =   70
      Size            =   "12515;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC14 
      Height          =   285
      Left            =   5295
      TabIndex        =   14
      Top             =   3960
      Width           =   3390
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5980;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC07 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   7125
      VariousPropertyBits=   671105051
      MaxLength       =   80
      Size            =   "12568;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC04 
      Height          =   285
      Left            =   5295
      TabIndex        =   3
      Top             =   1440
      Width           =   3390
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5980;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textNC06 
      Height          =   285
      Left            =   5295
      TabIndex        =   5
      Top             =   1740
      Width           =   3390
      VariousPropertyBits=   671105051
      MaxLength       =   30
      Size            =   "5980;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   2730
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   690
      Width           =   6075
      VariousPropertyBits=   679493663
      Size            =   "8555;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   15
      Left            =   5160
      TabIndex        =   36
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   14
      Left            =   1410
      TabIndex        =   35
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   13
      Left            =   5160
      TabIndex        =   34
      Top             =   1500
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   10
      Left            =   1410
      TabIndex        =   33
      Top             =   1500
      Width           =   90
   End
   Begin VB.Label Label30 
      Caption         =   "名稱(日)："
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   32
      Top             =   2070
      Width           =   915
   End
   Begin VB.Label Label29 
      Caption         =   "名稱(英)："
      Height          =   255
      Left            =   450
      TabIndex        =   31
      Top             =   1500
      Width           =   915
   End
   Begin VB.Label Label27 
      Caption         =   "名稱(中)："
      Height          =   255
      Left            =   450
      TabIndex        =   30
      Top             =   1170
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "備註："
      Height          =   255
      Left            =   450
      TabIndex        =   29
      Top             =   2400
      Width           =   915
   End
   Begin VB.Label Label18 
      Caption         =   "地址(中)："
      Height          =   255
      Left            =   450
      TabIndex        =   28
      Top             =   3120
      Width           =   915
   End
   Begin VB.Label Label16 
      Caption         =   "地址(英)："
      Height          =   255
      Left            =   450
      TabIndex        =   27
      Top             =   3420
      Width           =   915
   End
   Begin VB.Label Label13 
      Caption         =   "地址(日)："
      Height          =   255
      Left            =   450
      TabIndex        =   26
      Top             =   4260
      Width           =   915
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "1"
      Height          =   180
      Index           =   32
      Left            =   1410
      TabIndex        =   25
      Top             =   3390
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "3"
      Height          =   180
      Index           =   2
      Left            =   1410
      TabIndex        =   24
      Top             =   3690
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "5"
      Height          =   180
      Index           =   3
      Left            =   1410
      TabIndex        =   23
      Top             =   3990
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "2"
      Height          =   180
      Index           =   4
      Left            =   5160
      TabIndex        =   22
      Top             =   3390
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "4"
      Height          =   180
      Index           =   5
      Left            =   5160
      TabIndex        =   21
      Top             =   3690
      Width           =   90
   End
   Begin VB.Label Label41 
      AutoSize        =   -1  'True
      Caption         =   "6"
      Height          =   180
      Index           =   6
      Left            =   5160
      TabIndex        =   20
      Top             =   3990
      Width           =   90
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   255
      Index           =   0
      Left            =   450
      TabIndex        =   18
      Top             =   870
      Width           =   915
   End
End
Attribute VB_Name = "frm12040155"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/29 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2012/4/10
Option Explicit

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer

' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String

'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim TF_NC As Integer


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT count(*) FROM NOTCustomer "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 And rsTmp.Fields(0) > 0 Then
      rsTmp.Close
      strSql = "SELECT MIN(NC01) FROM NOTCustomer "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_FirstKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
      strSql = "SELECT MAX(NC01) FROM NOTCustomer "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_LastKEY(0) = rsTmp.Fields(0)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Initialize()
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from NOTCustomer where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_NC = AdoRecordSet3.Fields.Count
   CheckOC3
   ReDim m_FieldList(TF_NC) As FIELDITEM
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm12040155", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040155", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040155", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040155", strFind, False)
   
   MoveFormToCenter Me
   
   textCUID.BackColor = &H8000000F
   
   m_EditMode = 0
   
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To TF_NC
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "NC" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 21:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To TF_NC - 1
      If strName = m_FieldList(nIndex).fiName Then
         If strData = "#==#" Then
            m_FieldList(nIndex).fiNewData = m_FieldList(nIndex).fiOldData
         Else
            m_FieldList(nIndex).fiNewData = strData
         End If
         Exit For
      End If
   Next nIndex
End Sub

' 更新欄位的內容
Private Sub UpdateFieldNewData()
Dim strTmp  As String
   
   '若新增資料
   If m_EditMode = 1 Then
        '若未輸入編號
      If IsEmptyText(textNC01) = True Then
         If ClsPDGetAutoNumber("CR", strTmp, True, False) Then
            textNC01 = Format(strTmp, "0000")
         Else
            ShowMsg "讀取自動編號檔錯誤，請洽系統管理者 !"
            Exit Sub
         End If
      End If
   End If
   
   '編號
   If IsEmptyText(textNC01) = False Then
      SetFieldNewData "NC01", textNC01_1 & Format(textNC01, "0000")
   Else
      SetFieldNewData "NC01", textNC01_1 & textNC01
   End If
   SetFieldNewData "NC02", textNC02
   SetFieldNewData "NC03", textNC03
   SetFieldNewData "NC04", textNC04
   SetFieldNewData "NC05", textNC05
   SetFieldNewData "NC06", textNC06
   SetFieldNewData "NC07", textNC07
   SetFieldNewData "NC08", textNC08
   SetFieldNewData "NC09", textNC09
   SetFieldNewData "NC10", textNC10
   SetFieldNewData "NC11", textNC11
   SetFieldNewData "NC12", textNC12
   SetFieldNewData "NC13", textNC13
   SetFieldNewData "NC14", textNC14
   SetFieldNewData "NC15", textNC15
   SetFieldNewData "NC16", textNC16
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To TF_NC - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
            m_FieldList(nIndex).fiNewData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
            m_FieldList(nIndex).fiNewData = Empty
         End If
      End If
   Next nIndex
EXITSUB:
End Sub

' 讀取資料庫所有的資料
Private Sub QueryDB()
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   textNC01 = Empty
   textNC02 = Empty
   textNC03 = Empty
   textNC04 = Empty
   textNC05 = Empty
   textNC06 = Empty
   textNC07 = Empty
   textNC08 = Empty
   textNC09 = Empty
   textNC10 = Empty
   textNC11 = Empty
   textNC12 = Empty
   textNC13 = Empty
   textNC14 = Empty
   textNC15 = Empty
   textNC16 = Empty
   
   For nIndex = 0 To TF_NC - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   textCUID = ""
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textNC01.Locked = bEnable
   textNC02.Locked = bEnable
   textNC03.Locked = bEnable
   textNC04.Locked = bEnable
   textNC05.Locked = bEnable
   textNC06.Locked = bEnable
   textNC07.Locked = bEnable
   textNC08.Locked = bEnable
   textNC09.Locked = bEnable
   textNC10.Locked = bEnable
   textNC11.Locked = bEnable
   textNC12.Locked = bEnable
   textNC13.Locked = bEnable
   textNC14.Locked = bEnable
   textNC15.Locked = bEnable
   textNC16.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textNC01.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM NOTCustomer " & _
            "WHERE NC01 = '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NC01")) = False Then: textNC01 = Right(rsTmp.Fields("NC01"), 4)
      If IsNull(rsTmp.Fields("NC02")) = False Then: textNC02 = rsTmp.Fields("NC02")
      If IsNull(rsTmp.Fields("NC03")) = False Then: textNC03 = rsTmp.Fields("NC03")
      If IsNull(rsTmp.Fields("NC04")) = False Then: textNC04 = rsTmp.Fields("NC04")
      If IsNull(rsTmp.Fields("NC05")) = False Then: textNC05 = rsTmp.Fields("NC05")
      If IsNull(rsTmp.Fields("NC06")) = False Then: textNC06 = rsTmp.Fields("NC06")
      If IsNull(rsTmp.Fields("NC07")) = False Then: textNC07 = rsTmp.Fields("NC07")
      If IsNull(rsTmp.Fields("NC08")) = False Then: textNC08 = rsTmp.Fields("NC08")
      If IsNull(rsTmp.Fields("NC09")) = False Then: textNC09 = rsTmp.Fields("NC09")
      If IsNull(rsTmp.Fields("NC10")) = False Then: textNC10 = rsTmp.Fields("NC10")
      If IsNull(rsTmp.Fields("NC11")) = False Then: textNC11 = rsTmp.Fields("NC11")
      If IsNull(rsTmp.Fields("NC12")) = False Then: textNC12 = rsTmp.Fields("NC12")
      If IsNull(rsTmp.Fields("NC13")) = False Then: textNC13 = rsTmp.Fields("NC13")
      If IsNull(rsTmp.Fields("NC14")) = False Then: textNC14 = rsTmp.Fields("NC14")
      If IsNull(rsTmp.Fields("NC15")) = False Then: textNC15 = rsTmp.Fields("NC15")
      If IsNull(rsTmp.Fields("NC16")) = False Then: textNC16 = rsTmp.Fields("NC16")
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close
   
   textNC02.Tag = textNC02.Text
   textNC03.Tag = textNC03.Text
   textNC04.Tag = textNC04.Text
   textNC05.Tag = textNC05.Text
   textNC06.Tag = textNC06.Text
   textNC07.Tag = textNC07.Text
   textNC08.Tag = textNC08.Text
   textNC09.Tag = textNC09.Text
   textNC10.Tag = textNC10.Text
   textNC11.Tag = textNC11.Text
   textNC12.Tag = textNC12.Text
   textNC13.Tag = textNC13.Text
   textNC14.Tag = textNC14.Text
   textNC15.Tag = textNC15.Text
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("NC17")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC17")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("NC17"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NC18")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC18")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NC18"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NC19")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC19")) = False Then
         strTemp = rsSrcTmp.Fields("NC19")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NC20")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC20")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("NC20"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NC21")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC21")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NC21"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NC22")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NC22")) = False Then
         strTemp = rsSrcTmp.Fields("NC22")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT NC01 FROM NOTCustomer " & _
               "WHERE NC01 = '" & m_CurrKEY(0) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT NC01 FROM NOTCustomer " & _
               "WHERE NC01 = (SELECT MIN(NC01) FROM NOTCustomer " & _
                              "WHERE NC01 > '" & m_CurrKEY(0) & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
      Else
         ShowLastRecord
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrKEY(0) = m_FirstKEY(0)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
'   strSql = "SELECT NC01 FROM NOTCustomer " & _
'            "WHERE NC01 = '" & m_CurrKEY(0) & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT NC01 FROM NOTCustomer " & _
            "WHERE NC01 = (SELECT MAX(NC01) FROM NOTCustomer " & _
                           "WHERE NC01 < '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
'   strSql = "SELECT NC01 FROM NOTCustomer " & _
'            "WHERE NC01 = '" & m_CurrKEY(0) & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT NC01 FROM NOTCustomer " & _
            "WHERE NC01 = (SELECT MIN(NC01) FROM NOTCustomer " & _
                           "WHERE NC01 > '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NC01")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(6).Enabled = True
            tlbar.Buttons(7).Enabled = True
            tlbar.Buttons(8).Enabled = True
            tlbar.Buttons(9).Enabled = True
         Else
            tlbar.Buttons(6).Enabled = False
            tlbar.Buttons(7).Enabled = False
            tlbar.Buttons(8).Enabled = False
            tlbar.Buttons(9).Enabled = False
         End If
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
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 修改
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 查詢
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 刪除
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      ' 第一筆, 上一筆, 下一筆, 最後一筆
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF9, vbKeyF10:
         If m_EditMode <> 0 Then
            OnAction KeyCode
            KeyCode = 0
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
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
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
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
         PUB_FilterFormText Me '修正畫面所有含跳行符號的文字框
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         If CheckDataValid() = True Then
            If m_EditMode = 1 Or m_EditMode = 2 Then
               '重新檢查欄位有效性
               If TxtValidate = False Then Exit Sub
            End If
            UpdateFieldNewData
            OnWork
            UpdateToolbarState
        End If
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
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
      'SSTab1.Tab = 0
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm12040155 = Nothing
End Sub

Private Sub textNC01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textNC02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textNC07_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textNC08_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'日文地址要轉全形
Private Sub textNC15_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
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
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM NOTCustomer " & _
            "WHERE NC01 = '" & strKEY01 & "' "
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 檢查讀取的資料筆數
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
   Dim strNC01 As String
   Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   strNC01 = textNC01_1 & Format(textNC01, "0000")
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strNC01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo ErrHand
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO NOTCustomer ("
   For nIndex = 0 To TF_NC - 1
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
   For nIndex = 0 To TF_NC - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
            '字串中有單引號的處理
            strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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
   
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
      
   cnnConnection.CommitTrans
   If (strNC01 < m_FirstKEY(0)) Or (strNC01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strNC01
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
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
   Dim strNC01 As String
   Dim iErr As Integer, sErrMsg As String
   Dim arrFile1
   Dim ii As Integer
   Dim bolRemove As Boolean
   
On Error GoTo ErrHand
   
   strNC01 = m_CurrKEY(0)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE NOTCustomer SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_NC - 1
      strTmp = Empty
      If nIndex < 45 Or nIndex > 50 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     '字串中有單引號的處理
                     strTmp = m_FieldList(nIndex).fiName & " = '" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
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
      End If
   Next nIndex
   strSql = strSql & " " & _
                  "WHERE NC01 = '" & strNC01 & "'; end; "

   If bDifference = True Then
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
            
      cnnConnection.CommitTrans
      ShowCurrRecord strNC01
   End If
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strNC01 As String
   Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   strNC01 = m_CurrKEY(0)

   strSql = "DELETE FROM NOTCustomer " & _
            "WHERE NC01 = '" & strNC01 & "' "
   
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆
   If (strNC01 = m_LastKEY(0)) Or (strNC01 = m_FirstKEY(0)) Then
      RefreshRange
   End If
   ShowCurrRecord strNC01
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strNC01 As String
   
   QueryRecord = False
   strNC01 = textNC01_1 & Format(textNC01, "0000")
   
   textCUID = ""
   If IsRecordExist(strNC01) = True Then
      m_CurrKEY(0) = strNC01
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
      Case 1: '新增
'         '重新檢查欄位有效性
'         If TxtValidate = False Then Exit Sub
         AddRecord
         RefreshRange
      Case 2: '修改
'         '重新檢查欄位有效性
'         If TxtValidate = False Then Exit Sub
         ModRecord
      Case 3: '刪除
         DelRecord
         RefreshRange
      Case 4: '查詢
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textNC01.SetFocus
      Case 2: textNC02.SetFocus
      Case 4: textNC01.SetFocus
   End Select
End Sub

'編號
Private Sub textNC01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim LongMaxNum As Long
   
   Cancel = False
   '若有輸入編號
   If IsEmptyText(textNC01) = False Then
      '補滿4碼
      textNC01 = Right("0000" & textNC01, 4)
      '在新增時輸入的編號
      Select Case m_EditMode
         Case 1 '新增
            '不可大於目前檔案的最大號數
            LongMaxNum = Val(GetMaxNum)
            If Val(textNC01) > LongMaxNum Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "編號不可大於" & LongMaxNum
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNC01_GotFocus
               Exit Sub
            End If
            '不可已存在
            If IsRecordExist(textNC01_1 & textNC01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆編號已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNC01_GotFocus
               Exit Sub
            End If
      End Select
   End If
EXITSUB:
End Sub

'名稱(中)
Private Sub textNC02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNC02) = False Then
      If StrLength(textNC02) > textNC02.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "名稱(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNC02_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'名稱(日)
Private Sub textNC07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNC07) = False Then
      If StrLength(textNC07) > textNC07.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "名稱(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNC07_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'地址(中)
Private Sub textNC08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNC08) = False Then
      If StrLength(textNC08) > textNC08.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "地址(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNC08_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'地址(日)
Private Sub textNC15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNC15) = False Then
      If StrLength(textNC15) > textNC15.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "地址(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNC15_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim strTmp As String
   Dim nResponse
   CheckDataValid = False

   Select Case m_EditMode
      Case 4:
         ' 編號不可空白
         If IsEmptyText(textNC01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNC01.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
      
   Select Case m_EditMode
      Case 1, 2:
         '中文名稱, 英文名稱, 日文名稱不可全為空白
         If IsEmptyText(textNC02) = True And IsEmptyText(textNC03) = True And IsEmptyText(textNC07) = True Then
            strTit = "檢核資料"
            strMsg = "中文名稱, 英文名稱, 日文名稱不可全為空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNC02.SetFocus
            GoTo EXITSUB
         End If
'         '備註
'         If IsEmptyText(textNC16) = True Then
'            strTit = "檢核資料"
'            strMsg = "備註不可空白！"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textNC16.SetFocus
'            GoTo EXITSUB
'         End If
      Case Else:
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textNC01_GotFocus()
   InverseTextBox textNC01
End Sub

Private Sub textNC02_GotFocus()
   InverseTextBox textNC02
   OpenIme
End Sub

Private Sub textNC03_GotFocus()
   InverseTextBox textNC03
End Sub

Private Sub textNC04_GotFocus()
   InverseTextBox textNC04
End Sub

Private Sub textNC05_GotFocus()
   InverseTextBox textNC05
End Sub

Private Sub textNC06_GotFocus()
   InverseTextBox textNC06
End Sub

Private Sub textNC07_GotFocus()
   InverseTextBox textNC07
   OpenIme
End Sub

Private Sub textNC08_GotFocus()
   InverseTextBox textNC08
   OpenIme
End Sub

Private Sub textNC09_GotFocus()
   InverseTextBox textNC09
End Sub

Private Sub textNC10_GotFocus()
   InverseTextBox textNC10
End Sub

Private Sub textNC11_GotFocus()
   InverseTextBox textNC11
End Sub

Private Sub textNC12_GotFocus()
   InverseTextBox textNC12
End Sub

Private Sub textNC13_GotFocus()
   InverseTextBox textNC13
End Sub

Private Sub textNC14_GotFocus()
   InverseTextBox textNC14
End Sub

Private Sub textNC15_GotFocus()
   InverseTextBox textNC15
   OpenIme
End Sub

Private Sub textNC16_GotFocus()
   InverseTextBox textNC16
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
TxtValidate = False

If Me.textNC02.Enabled = True Then
   Cancel = False
   textNC02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNC07.Enabled = True Then
   Cancel = False
   textNC07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNC08.Enabled = True Then
   Cancel = False
   textNC08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNC15.Enabled = True Then
   Cancel = False
   textNC15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Sindy 2021/11/29 檢查畫面上的物件是否含有Unicode文字
If PUB_ChkUniText(Me, True, True) = False Then
   Cancel = True
   Exit Function
End If

TxtValidate = True
End Function

Private Function GetMaxNum() As String
   GetMaxNum = "0"
   strSql = "select count(*) from NOTCustomer "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         strSql = "select max(NC01) from NOTCustomer "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            GetMaxNum = Right(RsTemp.Fields(0), 4)
         End If
      End If
   End If
End Function
