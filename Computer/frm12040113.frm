VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040113 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利年費資料檔"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7605
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   7605
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6780
      Top             =   750
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
            Picture         =   "frm12040113.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040113.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   1085
      ButtonWidth     =   1138
      ButtonHeight    =   1032
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
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
   Begin MSForms.TextBox textYF15 
      Height          =   760
      Left            =   1440
      TabIndex        =   7
      Top             =   3240
      Width           =   5895
      VariousPropertyBits=   -1466941413
      MaxLength       =   100
      ScrollBars      =   2
      Size            =   "10398;1341"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF04_2 
      Height          =   300
      Left            =   2160
      TabIndex        =   19
      Top             =   1800
      Width           =   4452
      VariousPropertyBits=   671105055
      Size            =   "7853;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF03_2 
      Height          =   300
      Left            =   2760
      TabIndex        =   18
      Top             =   1440
      Width           =   3852
      VariousPropertyBits=   671105055
      Size            =   "6794;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF02_2 
      Height          =   300
      Left            =   1920
      TabIndex        =   17
      Top             =   1080
      Width           =   4692
      VariousPropertyBits=   671105055
      Size            =   "8276;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF01_2 
      Height          =   300
      Left            =   2160
      TabIndex        =   16
      Top             =   720
      Width           =   4452
      VariousPropertyBits=   671105055
      Size            =   "7853;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF07 
      Height          =   300
      Left            =   1440
      TabIndex        =   6
      Top             =   2880
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF06 
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   2520
      Width           =   855
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1508;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF05 
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   612
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1080;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF04 
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   1800
      Width           =   612
      VariousPropertyBits=   671105051
      MaxLength       =   4
      Size            =   "1080;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF03 
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   1440
      Width           =   1212
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "2138;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF02 
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   1080
      Width           =   372
      VariousPropertyBits=   671105051
      MaxLength       =   1
      Size            =   "656;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textYF01 
      Height          =   300
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   612
      VariousPropertyBits=   671105051
      MaxLength       =   3
      Size            =   "1080;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年費年度說明 :"
      Height          =   180
      Index           =   7
      Left            =   240
      TabIndex        =   20
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "規費 :"
      Height          =   180
      Index           =   6
      Left            =   240
      TabIndex        =   14
      Top             =   2880
      Width           =   576
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "服務費 :"
      Height          =   180
      Index           =   5
      Left            =   240
      TabIndex        =   13
      Top             =   2520
      Width           =   732
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年度 :"
      Height          =   180
      Index           =   4
      Left            =   240
      TabIndex        =   12
      Top             =   2160
      Width           =   576
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   3
      Left            =   240
      TabIndex        =   11
      Top             =   1800
      Width           =   816
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Index           =   2
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   852
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利種類 :"
      Height          =   180
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1080
      Width           =   936
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請國家 :"
      Height          =   180
      Index           =   0
      Left            =   240
      TabIndex        =   8
      Top             =   720
      Width           =   936
   End
End
Attribute VB_Name = "frm12040113"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/17 改成Form2.0 ;所有TextBox
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit
'Modify By Sindy 2009/06/26
'Const MAX_FIELD = 7
Const MAX_FIELD = 15

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
Dim m_SubMode As Integer

' 辦識其為外商還是內商的程式
' 0 表內商
' 1 表外商
Dim m_SysKind As Integer

' 第一筆資料的本所案號
Dim m_FirstYF(5) As String
' 最後一筆資料的本所案號
Dim m_LastYF(5) As String
' 目前正在顯示的本所案號
Dim m_CurrYF(5) As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT YF01,YF02,YF03,YF04,YF05 FROM PATENTYEARFEE " & _
            "WHERE YF01 || YF02 || YF03 || YF04 || YF05 = (SELECT MIN(YF01 || YF02 || YF03 || YF04 || YF05) FROM PATENTYEARFEE ) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YF01")) = False Then: m_FirstYF(0) = rsTmp.Fields("YF01")
      If IsNull(rsTmp.Fields("YF02")) = False Then: m_FirstYF(1) = rsTmp.Fields("YF02")
      If IsNull(rsTmp.Fields("YF03")) = False Then: m_FirstYF(2) = rsTmp.Fields("YF03")
      If IsNull(rsTmp.Fields("YF04")) = False Then: m_FirstYF(3) = rsTmp.Fields("YF04")
      If IsNull(rsTmp.Fields("YF05")) = False Then: m_FirstYF(4) = rsTmp.Fields("YF05")
   End If
   rsTmp.Close

   strSql = "SELECT YF01,YF02,YF03,YF04,YF05 FROM PATENTYEARFEE " & _
            "WHERE YF01 || YF02 || YF03 || YF04 || YF05 = (SELECT MAX(YF01 || YF02 || YF03 || YF04 || YF05) FROM PATENTYEARFEE ) "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YF01")) = False Then: m_LastYF(0) = rsTmp.Fields("YF01")
      If IsNull(rsTmp.Fields("YF02")) = False Then: m_LastYF(1) = rsTmp.Fields("YF02")
      If IsNull(rsTmp.Fields("YF03")) = False Then: m_LastYF(2) = rsTmp.Fields("YF03")
      If IsNull(rsTmp.Fields("YF04")) = False Then: m_LastYF(3) = rsTmp.Fields("YF04")
      If IsNull(rsTmp.Fields("YF05")) = False Then: m_LastYF(4) = rsTmp.Fields("YF05")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'Added by Lydia 2021/11/17
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

'Memo by Lydia 2021/11/17 從Form_KeyDown搬來
Select Case KeyCode
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
      Case vbKeyF2:
         If m_bInsert Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF3:
         If m_bUpdate Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF4:
         If m_bQuery Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
      Case vbKeyF5:
         If m_bDelete Then
            If m_EditMode = 0 Then
               OnAction KeyCode
               KeyCode = 0
            End If
         End If
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
'edit by nickc 2006/11/13
'      Case vbKeyReturn:
'         If m_EditMode <> 0 Then
'            OnAction vbKeyF9
'         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

' Load Form
Private Sub Form_Load()
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm12040113", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040113", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040113", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040113", strFind, False)
   
   textYF01_2.BackColor = &H8000000F
   textYF02_2.BackColor = &H8000000F
   textYF03_2.BackColor = &H8000000F
   textYF04_2.BackColor = &H8000000F
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
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
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "YF" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 5, 6, 7:
            m_FieldList(nIndex - 1).fiType = 1
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To MAX_FIELD - 1
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
   SetFieldNewData "YF01", textYF01
   SetFieldNewData "YF02", textYF02
   ' 代理人補足9碼
   If IsEmptyText(textYF03) = False Then
      SetFieldNewData "YF03", textYF03 & String(9 - Len(textYF03), "0")
   Else
      SetFieldNewData "YF03", textYF03
   End If
   SetFieldNewData "YF04", textYF04
   SetFieldNewData "YF05", textYF05
   SetFieldNewData "YF06", textYF06
   SetFieldNewData "YF07", textYF07
   'Add By Sindy 2009/06/26
   SetFieldNewData "YF15", textYF15
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
      If m_FieldList(nIndex).fiName <> Empty Then
         If IsNull(rsTmp.Fields(m_FieldList(nIndex).fiName)) = False Then
            m_FieldList(nIndex).fiOldData = rsTmp.Fields(m_FieldList(nIndex).fiName)
         Else
            m_FieldList(nIndex).fiOldData = Empty
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
   textYF01 = Empty
   textYF01_2 = Empty
   textYF02 = Empty
   textYF02_2 = Empty
   textYF03 = Empty
   textYF03_2 = Empty
   textYF04 = Empty
   textYF04_2 = Empty
   textYF05 = Empty
   textYF06 = Empty
   textYF07 = Empty
   'Add By Sindy 2009/06/26
   textYF15 = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textYF01.Locked = bEnable
   textYF02.Locked = bEnable
   textYF03.Locked = bEnable
   textYF04.Locked = bEnable
   textYF05.Locked = bEnable
   textYF06.Locked = bEnable
   textYF07.Locked = bEnable
   'Add By Sindy 2009/06/26
   textYF15.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textYF01.Locked = bEnable
   textYF02.Locked = bEnable
   textYF03.Locked = bEnable
   textYF04.Locked = bEnable
   textYF05.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   If m_CurrYF(0) = Empty Or m_CurrYF(1) = Empty Or m_CurrYF(2) = Empty Or m_CurrYF(3) = Empty Or m_CurrYF(4) = Empty Then
      GoTo EXITSUB
   End If
   
   strSql = "SELECT * FROM PATENTYEARFEE " & _
            "WHERE YF01 = '" & m_CurrYF(0) & "' AND " & _
                  "YF02 = '" & m_CurrYF(1) & "' AND " & _
                  "YF03 = '" & m_CurrYF(2) & "' AND " & _
                  "YF04 = '" & m_CurrYF(3) & "' AND " & _
                  "YF05 = '" & m_CurrYF(4) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("YF01")) = False Then
         textYF01 = rsTmp.Fields("YF01")
      End If
      If IsNull(rsTmp.Fields("YF02")) = False Then
         textYF02 = rsTmp.Fields("YF02")
      End If
      If IsNull(rsTmp.Fields("YF03")) = False Then
         textYF03 = rsTmp.Fields("YF03")
      End If
      If IsNull(rsTmp.Fields("YF04")) = False Then
         textYF04 = rsTmp.Fields("YF04")
      End If
      If IsNull(rsTmp.Fields("YF05")) = False Then
         textYF05 = rsTmp.Fields("YF05")
      End If
      If IsNull(rsTmp.Fields("YF06")) = False Then
         textYF06 = rsTmp.Fields("YF06")
      End If
      If IsNull(rsTmp.Fields("YF07")) = False Then
         textYF07 = rsTmp.Fields("YF07")
      End If
      'Add By Sindy 2009/06/26
      If IsNull(rsTmp.Fields("YF15")) = False Then
         textYF15 = rsTmp.Fields("YF15")
      End If
      
      ' 更新暫存區內的資料
      UpdateFieldOldData rsTmp
      ' 帶出相關資料
      textYF01_Validate False
      textYF02_Validate False
      textYF03_Validate False
      textYF04_Validate False
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strYF01 As String, ByVal strYF02 As String, ByVal strYF03 As String, ByVal strYF04 As String, ByVal strYF05 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If PUB_PYFIsExists(strYF01, strYF02, strYF03, strYF04, strYF05) = True Then
      m_CurrYF(0) = strYF01
      m_CurrYF(1) = strYF02
      m_CurrYF(2) = strYF03
      m_CurrYF(3) = strYF04
      m_CurrYF(4) = strYF05
   Else
      strSql = "SELECT YF01,YF02,YF03,YF04,YF05 FROM PATENTYEARFEE " & _
               "WHERE YF01 || YF02 || YF03 || YF04 || YF05 = (SELECT MIN(YF01 || YF02 || YF03 || YF04 || YF05) FROM PATENTYEARFEE " & _
                                                         "WHERE (YF01 || YF02 || YF03 || YF04 || YF05) > '" & m_CurrYF(0) & m_CurrYF(1) & m_CurrYF(2) & m_CurrYF(3) & m_CurrYF(4) & "') "
                  
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("YF01")) = False Then: m_CurrYF(0) = rsTmp.Fields("YF01")
         If IsNull(rsTmp.Fields("YF02")) = False Then: m_CurrYF(1) = rsTmp.Fields("YF02")
         If IsNull(rsTmp.Fields("YF03")) = False Then: m_CurrYF(2) = rsTmp.Fields("YF03")
         If IsNull(rsTmp.Fields("YF04")) = False Then: m_CurrYF(3) = rsTmp.Fields("YF04")
         If IsNull(rsTmp.Fields("YF05")) = False Then: m_CurrYF(4) = rsTmp.Fields("YF05")
      End If
      rsTmp.Close
   End If
   UpdateCtrlData
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示第一筆資料
Private Sub ShowFirstRecord()
   m_CurrYF(0) = m_FirstYF(0)
   m_CurrYF(1) = m_FirstYF(1)
   m_CurrYF(2) = m_FirstYF(2)
   m_CurrYF(3) = m_FirstYF(3)
   m_CurrYF(4) = m_FirstYF(4)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrYF(0) = m_FirstYF(0) And m_CurrYF(1) = m_FirstYF(1) And m_CurrYF(2) = m_FirstYF(2) And m_CurrYF(3) = m_FirstYF(3) And m_CurrYF(4) = m_FirstYF(4) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   'Modified by Morgan 2023/2/9 年度改前面補0後排序否則10會跳1而不是9
   strSql = "SELECT YF01,YF02,YF03,YF04,YF05 FROM PATENTYEARFEE " & _
            "WHERE YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0') = (SELECT MAX(YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0')) FROM PATENTYEARFEE " & _
                                                      "WHERE (YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0')) < '" & m_CurrYF(0) & m_CurrYF(1) & m_CurrYF(2) & m_CurrYF(3) & Right("0" & m_CurrYF(4), 2) & "') "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YF01")) = False Then: m_CurrYF(0) = rsTmp.Fields("YF01")
      If IsNull(rsTmp.Fields("YF02")) = False Then: m_CurrYF(1) = rsTmp.Fields("YF02")
      If IsNull(rsTmp.Fields("YF03")) = False Then: m_CurrYF(2) = rsTmp.Fields("YF03")
      If IsNull(rsTmp.Fields("YF04")) = False Then: m_CurrYF(3) = rsTmp.Fields("YF04")
      If IsNull(rsTmp.Fields("YF05")) = False Then: m_CurrYF(4) = rsTmp.Fields("YF05")
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
   
   If m_CurrYF(0) = m_LastYF(0) And m_CurrYF(1) = m_LastYF(1) And m_CurrYF(2) = m_LastYF(2) And m_CurrYF(3) = m_LastYF(3) And m_CurrYF(4) = m_LastYF(4) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   'Modified by Morgan 2023/2/9 年度改前面補0後排序否則1會跳10而不是2
   strSql = "SELECT YF01,YF02,YF03,YF04,YF05 FROM PATENTYEARFEE " & _
            "WHERE YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0') = (SELECT MIN(YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0')) FROM PATENTYEARFEE " & _
                                                      "WHERE (YF01 || YF02 || YF03 || YF04 || lpad(YF05,2,'0')) > '" & m_CurrYF(0) & m_CurrYF(1) & m_CurrYF(2) & m_CurrYF(3) & Right("0" & m_CurrYF(4), 2) & "') "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("YF01")) = False Then: m_CurrYF(0) = rsTmp.Fields("YF01")
      If IsNull(rsTmp.Fields("YF02")) = False Then: m_CurrYF(1) = rsTmp.Fields("YF02")
      If IsNull(rsTmp.Fields("YF03")) = False Then: m_CurrYF(2) = rsTmp.Fields("YF03")
      If IsNull(rsTmp.Fields("YF04")) = False Then: m_CurrYF(3) = rsTmp.Fields("YF04")
      If IsNull(rsTmp.Fields("YF05")) = False Then: m_CurrYF(4) = rsTmp.Fields("YF05")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrYF(0) = m_LastYF(0)
   m_CurrYF(1) = m_LastYF(1)
   m_CurrYF(2) = m_LastYF(2)
   m_CurrYF(3) = m_LastYF(3)
   m_CurrYF(4) = m_LastYF(4)
   
   UpdateCtrlData
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         ' 90.07.13 modify by louis (依照權限設定其工具列的按紐狀態)
         'tlbar.Buttons(1).Enabled = True
         'tlbar.Buttons(2).Enabled = True
         'tlbar.Buttons(3).Enabled = True
         'tlbar.Buttons(4).Enabled = True
         'tlbar.Buttons(6).Enabled = True
         'tlbar.Buttons(7).Enabled = True
         'tlbar.Buttons(8).Enabled = True
         'tlbar.Buttons(9).Enabled = True
         'tlbar.Buttons(11).Enabled = False
         'tlbar.Buttons(12).Enabled = False
         'tlbar.Buttons(14).Enabled = True
         
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
   'Memo by Lydia 2021/11/17 原程式搬到Form_KeyUp
   
End Sub
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    'Remove by Lydia 2021/11/17 改成Form2.0; 取消Enter鍵=確定(Toolbar),因為MsgBox也會回傳Enter鍵
    'Select Case KeyAscii
    '  Case vbKeyReturn:
    '     If m_EditMode <> 0 Then
    '        KeyAscii = 0
    '        OnAction vbKeyF9
    '     End If
    'End Select
    'end 2021/11/17
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   m_SubMode = 0
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
         PUB_FilterFormText Me 'Add by Morgan 2008/6/20 修正畫面所有含跳行符號的文字框
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
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
   'Add By Cheng 2002/07/18
   Set frm12040113 = Nothing
End Sub

'Modified by Lydia 2021/11/17 改成Form 2.0
'Private Sub textYF03_KeyPress(KeyAscii As Integer)
Private Sub textYF03_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
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
'Modified by Lydia 2015/01/05 改成共用模組PUB_PYFIsExists
'' 檢查記錄是否已經存在
'Private Function IsRecordExist(ByVal strYF01 As String, ByVal strYF02 As String, ByVal strYF03 As String, ByVal strYF04 As String, ByVal strYF05 As String) As Boolean
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   IsRecordExist = False
'   strSql = "SELECT * FROM PATENTYEARFEE " & _
'            "WHERE YF01 = '" & strYF01 & "' AND " & _
'                  "YF02 = '" & strYF02 & "' AND " & _
'                  "YF03 = '" & strYF03 & "' AND " & _
'                  "YF04 = '" & strYF04 & "' AND " & _
'                  "YF05 = '" & strYF05 & "' "
'
'   ' 讀取資料庫
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenDynamic
'   ' 檢查讀取的資料筆數
'   If rsTmp.RecordCount > 0 Then
'      IsRecordExist = True
'   Else
'      IsRecordExist = False
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

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
   Dim strYF01 As String
   Dim strYF02 As String
   Dim strYF03 As String
   Dim strYF04 As String
   Dim strYF05 As String
   
   strYF01 = textYF01
   strYF02 = textYF02
   strYF03 = textYF03
   strYF04 = textYF04
   strYF05 = textYF05
   
   If Len(strYF03) < 9 Then: strYF03 = strYF03 & String(9 - Len(strYF03), "0")
   
   ' 檢查記錄是否已存在
   If PUB_PYFIsExists(strYF01, strYF02, strYF03, strYF04, strYF05) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO PATENTYEARFEE ("
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
   
   If ((strYF01 & strYF02 & strYF03 & strYF04 & strYF05) < (m_FirstYF(0) & m_FirstYF(1) & m_FirstYF(2) & m_FirstYF(3) & m_FirstYF(4))) Or ((strYF01 & strYF02 & strYF03 & strYF04 & strYF05) > (m_LastYF(0) & m_LastYF(1) & m_LastYF(2) & m_LastYF(3) & m_LastYF(4))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strYF01, strYF02, strYF03, strYF04, strYF05
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
   Dim strYF01 As String
   Dim strYF02 As String
   Dim strYF03 As String
   Dim strYF04 As String
   Dim strYF05 As String
   
   strYF01 = m_CurrYF(0)
   strYF02 = m_CurrYF(1)
   strYF03 = m_CurrYF(2)
   strYF04 = m_CurrYF(3)
   strYF05 = m_CurrYF(4)
   
   strSql = "UPDATE PATENTYEARFEE SET "
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
                  "WHERE YF01 = '" & strYF01 & "' AND " & _
                        "YF02 = '" & strYF02 & "' AND " & _
                        "YF03 = '" & strYF03 & "' AND " & _
                        "YF04 = '" & strYF04 & "' AND " & _
                        "YF05 = '" & strYF05 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      ShowCurrRecord strYF01, strYF02, strYF03, strYF04, strYF05
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strYF01 As String
   Dim strYF02 As String
   Dim strYF03 As String
   Dim strYF04 As String
   Dim strYF05 As String
   
   strYF01 = m_CurrYF(0)
   strYF02 = m_CurrYF(1)
   strYF03 = m_CurrYF(2)
   strYF04 = m_CurrYF(3)
   strYF05 = m_CurrYF(4)

   strSql = "DELETE FROM PATENTYEARFEE " & _
            "WHERE YF01 = '" & strYF01 & "' AND " & _
                  "YF02 = '" & strYF02 & "' AND " & _
                  "YF03 = '" & strYF03 & "' AND " & _
                  "YF04 = '" & strYF04 & "' AND " & _
                  "YF05 = '" & strYF05 & "' "

   cnnConnection.Execute strSql

   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If ((strYF01 & strYF02 & strYF03 & strYF04 & strYF05) = (m_FirstYF(0) & m_FirstYF(1) & m_FirstYF(2) & m_FirstYF(3) & m_FirstYF(4))) Or ((strYF01 & strYF02 & strYF03 & strYF04 & strYF05) = (m_LastYF(0) & m_LastYF(1) & m_LastYF(2) & m_LastYF(3) & m_LastYF(4))) Then
      RefreshRange
   End If
   ShowCurrRecord strYF01, strYF02, strYF03, strYF04, strYF05
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strYF01 As String
   Dim strYF02 As String
   Dim strYF03 As String
   Dim strYF04 As String
   Dim strYF05 As String
   
   QueryRecord = False
   
   strYF01 = textYF01
   strYF02 = textYF02
   If IsEmptyText(textYF03) = False Then
      strYF03 = textYF03 & String(9 - Len(textYF03), "0")
   Else
      strYF03 = textYF03
   End If
   strYF04 = textYF04
   strYF05 = textYF05

   If PUB_PYFIsExists(strYF01, strYF02, strYF03, strYF04, strYF05) = True Then
      m_CurrYF(0) = strYF01
      m_CurrYF(1) = strYF02
      m_CurrYF(2) = strYF03
      m_CurrYF(3) = strYF04
      m_CurrYF(4) = strYF05
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
            
            ModRecord
         Else
            GoTo EXITSUB
         End If
      Case 3:
         DelRecord
         RefreshRange
      Case 4:
         If CheckDataValid() = True Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            GoTo EXITSUB
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textYF01.SetFocus
      Case 2: textYF06.SetFocus
              textYF06_GotFocus
      Case 4: textYF01.SetFocus
   End Select
End Sub

' 申請國家
Private Sub textYF01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textYF01_2 = Empty
   If IsEmptyText(textYF01) = False Then
      textYF01_2 = GetNationName(textYF01, 0)
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textYF01_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "申請國家代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textYF01_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 專利種類
Private Sub textYF02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textYF02_2 = Empty
   If IsEmptyText(textYF02) = False Then
      If textYF01 < "010" Then
         textYF02_2 = GetPatentName(textYF02, 0)
      Else
         textYF02_2 = GetPatentName(textYF02, 1)
      End If
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textYF02_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "專利種類不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textYF02_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 代理人
Private Sub textYF03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textYF03_2 = Empty
   If IsEmptyText(textYF03) = False Then
      textYF03_2 = GetFAgentName(textYF03)
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textYF03_2) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "代理人代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textYF03_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 案件性質
Private Sub textYF04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textYF04_2 = Empty
   If IsEmptyText(textYF04) = False Then
      ' 先以P類的取得案件性質的名稱
      '2009/7/9 modify by sonia 以申請國家判斷
      'If textYF01 > "010" Then
      '   textYF04_2 = GetCaseTypeName("P", textYF04, 1)
      'Else
      '   textYF04_2 = GetCaseTypeName("P", textYF04, 0)
      'End If
      ' 若抓不到案件性質則再以CFP類再去取一次
      'If IsEmptyText(textYF04_2) = True Then
      '   If textYF01 > "010" Then
      '      textYF04_2 = GetCaseTypeName("CFP", textYF04, 1)
      '   Else
      '      textYF04_2 = GetCaseTypeName("CFP", textYF04, 0)
      '   End If
      'End If
      Select Case textYF01
         Case "000", "013", "020", "044", "056"
            textYF04_2 = GetCaseTypeName("P", textYF04, 0)
         Case "000", "013", "020", "044", "056"
            textYF04_2 = GetCaseTypeName("P", textYF04, 1)
         Case Else
            textYF04_2 = GetCaseTypeName("CFP", textYF04, 1)
      End Select
      '2009/7/9 end
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsEmptyText(textYF04_2) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "案件性質代號不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textYF04_GotFocus
            End If
         Case Else:
      End Select
   End If
End Sub

' 年度
Private Sub textYF05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textYF05) = False Then
      Select Case m_EditMode
         Case 1, 2, 4:
            If IsNumeric(textYF05) = False Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "年度請輸入數值資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textYF05_GotFocus
            End If
         Case Else:
      End Select
      Select Case m_EditMode
         Case 1:
            If IsEmptyText(textYF01) = False And IsEmptyText(textYF02) = False And IsEmptyText(textYF03) = False And IsEmptyText(textYF01) = False Then
               If PUB_PYFIsExists(textYF01, textYF02, textYF03, textYF04, textYF05) = True Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "該筆記錄已經存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textYF05_GotFocus
               End If
            End If
      End Select
   End If
End Sub

' 服務費
Private Sub textYF06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textYF06) = False Then
      If IsNumeric(textYF06) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "服務費請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textYF06_GotFocus
      End If
   End If
End Sub

' 規費
Private Sub textYF07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textYF07) = False Then
      If IsNumeric(textYF07) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "規費請輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textYF07_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2009/06/26
'年費年度說明
Private Sub textYF15_Validate(Cancel As Boolean)
   If CheckLengthIsOK(textYF15, textYF15.MaxLength) = False Then
      Call textYF15_GotFocus
      Cancel = True
      Exit Sub
   End If
   CloseIme
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 申請國家不可空白
         If IsEmptyText(textYF01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入申請國家"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF01.SetFocus
            GoTo EXITSUB
         End If
         ' 專利種類不可為空白
         If IsEmptyText(textYF02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入專利種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF02.SetFocus
            GoTo EXITSUB
         End If
         ' 代理人不可為空白
         If IsEmptyText(textYF03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF03.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   Select Case m_EditMode
      Case 1, 2:
         ' 申請國家不可空白
         If IsEmptyText(textYF01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入申請國家"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF01.SetFocus
            GoTo EXITSUB
         End If
         ' 專利種類不可為空白
         If IsEmptyText(textYF02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入專利種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF02.SetFocus
            GoTo EXITSUB
         End If
         ' 代理人不可為空白
         If IsEmptyText(textYF03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入代理人"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF03.SetFocus
            GoTo EXITSUB
         End If
         ' 年費種類不可為空白
         If IsEmptyText(textYF04) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入年費種類"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF04.SetFocus
            GoTo EXITSUB
         End If
         ' 年度不可為空白
         If IsEmptyText(textYF05) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入年度"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF05.SetFocus
            GoTo EXITSUB
         End If
         ' 服務費及規費不可同時空白
         If IsEmptyText(textYF06) = True And IsEmptyText(textYF07) = True Then
            strTit = "檢核資料"
            strMsg = "服務費及規費不可同時空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textYF06.SetFocus
            GoTo EXITSUB
         End If
         Case Else:
   End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textYF01_GotFocus()
   InverseTextBox textYF01
End Sub

Private Sub textYF02_GotFocus()
   InverseTextBox textYF02
End Sub

Private Sub textYF03_GotFocus()
   InverseTextBox textYF03
End Sub

Private Sub textYF04_GotFocus()
   InverseTextBox textYF04
End Sub

Private Sub textYF05_GotFocus()
   InverseTextBox textYF05
End Sub

Private Sub textYF06_GotFocus()
   InverseTextBox textYF06
End Sub

Private Sub textYF07_GotFocus()
   InverseTextBox textYF07
End Sub

'Add By Sindy 2009/06/26
Private Sub textYF15_GotFocus()
   InverseTextBox textYF15
   OpenIme
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textYF01.Enabled = True Then
   Cancel = False
   textYF01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF02.Enabled = True Then
   Cancel = False
   textYF02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF03.Enabled = True Then
   Cancel = False
   textYF03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF04.Enabled = True Then
   Cancel = False
   textYF04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF05.Enabled = True Then
   Cancel = False
   textYF05_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF06.Enabled = True Then
   Cancel = False
   textYF06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textYF07.Enabled = True Then
   Cancel = False
   textYF07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2009/06/26
If Me.textYF15.Enabled = True Then
   Cancel = False
   textYF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function
