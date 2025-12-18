VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050707 
   BorderStyle     =   1  '單線固定
   Caption         =   "延期記錄資料維護"
   ClientHeight    =   5760
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9144
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8580
      Top             =   600
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050707.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
      _ExtentY        =   1016
      ButtonWidth     =   1101
      ButtonHeight    =   974
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4872
      Left            =   120
      TabIndex        =   12
      Top             =   720
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   8594
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050707.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "textDL06"
      Tab(0).Control(1)=   "textDL05"
      Tab(0).Control(2)=   "textKEY02_2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textKEY04"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textKEY03"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textKEY01"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textDL04"
      Tab(0).Control(7)=   "textDL03"
      Tab(0).Control(8)=   "textDL02"
      Tab(0).Control(9)=   "textDL01"
      Tab(0).Control(10)=   "textKEY02"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(4)"
      Tab(0).Control(12)=   "Label1(3)"
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(14)=   "Label1(1)"
      Tab(0).Control(15)=   "Label1(2)"
      Tab(0).Control(16)=   "Label2"
      Tab(0).Control(17)=   "Label1(0)"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm050707.frx":2110
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "grdList"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textCP02"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdQuery"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textCP02_2"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textCP04"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textCP03"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textCP01"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      Begin VB.TextBox textDL06 
         Height          =   270
         Left            =   -73380
         MaxLength       =   10
         TabIndex        =   5
         Top             =   2760
         Width           =   1215
      End
      Begin VB.TextBox textDL05 
         Height          =   270
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   4
         Top             =   2370
         Width           =   375
      End
      Begin VB.TextBox textKEY02_2 
         Height          =   264
         Left            =   -72360
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   912
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox textKEY04 
         Height          =   264
         Left            =   -71880
         MaxLength       =   2
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   912
         Width           =   492
      End
      Begin VB.TextBox textKEY03 
         Height          =   264
         Left            =   -72120
         MaxLength       =   1
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   912
         Width           =   252
      End
      Begin VB.TextBox textKEY01 
         Height          =   264
         Left            =   -73560
         MaxLength       =   3
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   912
         Width           =   492
      End
      Begin VB.TextBox textCP01 
         Height          =   264
         Left            =   1080
         MaxLength       =   3
         TabIndex        =   6
         Top             =   432
         Width           =   612
      End
      Begin VB.TextBox textCP03 
         Height          =   264
         Left            =   2640
         MaxLength       =   1
         TabIndex        =   9
         Top             =   432
         Width           =   252
      End
      Begin VB.TextBox textCP04 
         Height          =   264
         Left            =   2880
         MaxLength       =   2
         TabIndex        =   10
         Top             =   432
         Width           =   492
      End
      Begin VB.TextBox textCP02_2 
         Height          =   264
         Left            =   2400
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   432
         Visible         =   0   'False
         Width           =   252
      End
      Begin VB.TextBox textDL04 
         Height          =   270
         Left            =   -73560
         Locked          =   -1  'True
         MaxLength       =   7
         TabIndex        =   3
         Top             =   1992
         Width           =   975
      End
      Begin VB.TextBox textDL03 
         Height          =   270
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   2
         Top             =   1632
         Width           =   975
      End
      Begin VB.TextBox textDL02 
         Height          =   270
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   1
         Top             =   1272
         Width           =   975
      End
      Begin VB.TextBox textDL01 
         Height          =   270
         Left            =   -73560
         MaxLength       =   9
         TabIndex        =   0
         Top             =   552
         Width           =   1212
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Left            =   7860
         TabIndex        =   11
         Top             =   360
         Width           =   912
      End
      Begin VB.TextBox textCP02 
         Height          =   264
         Left            =   1680
         MaxLength       =   6
         TabIndex        =   7
         Top             =   432
         Width           =   972
      End
      Begin VB.TextBox textKEY02 
         Height          =   264
         Left            =   -73080
         MaxLength       =   6
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   912
         Width           =   972
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3852
         Left            =   72
         TabIndex        =   27
         Top             =   816
         Width           =   8652
         _ExtentX        =   15261
         _ExtentY        =   6795
         _Version        =   393216
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         Caption         =   "下一程序序號 : "
         Height          =   255
         Index           =   4
         Left            =   -74760
         TabIndex        =   26
         Top             =   2790
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "資料來源 :                     (1.案件進度檔 2.下一程序檔)"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   25
         Top             =   2370
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "本所案號 : "
         Height          =   252
         Left            =   -74760
         TabIndex        =   19
         Top             =   912
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "原法定期限 : "
         Height          =   252
         Index           =   1
         Left            =   -74760
         TabIndex        =   17
         Top             =   1992
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "原本所期限 :"
         Height          =   252
         Index           =   2
         Left            =   -74760
         TabIndex        =   16
         Top             =   1632
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "延期日 :"
         Height          =   252
         Left            =   -74760
         TabIndex        =   15
         Top             =   1272
         Width           =   732
      End
      Begin VB.Label Label1 
         Caption         =   "總收文號 :"
         Height          =   252
         Index           =   0
         Left            =   -74760
         TabIndex        =   14
         Top             =   552
         Width           =   972
      End
      Begin VB.Label Label4 
         Caption         =   "本所案號 :"
         Height          =   252
         Left            =   120
         TabIndex        =   13
         Top             =   432
         Width           =   852
      End
   End
End
Attribute VB_Name = "frm050707"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Lydia 2021/11/22 Form2.0已檢查 (無需修改的物件)
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'Modify By Cheng 2002/06/26
'Modify By Cheng 2002/06/20
'Const MAX_FIELD = 4
'Const MAX_FIELD = 5
Const MAX_FIELD = 6

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
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
'
Dim m_CurrSel As Integer

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT ) AND " & _
                  "DL02 = (SELECT MIN(DL02) FROM DATELIMIT " & _
                           "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("DL02")
   End If
   rsTmp.Close

   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = (SELECT MAX(DL01) FROM DATELIMIT ) AND " & _
                  "DL02 = (SELECT MAX(DL02) FROM DATELIMIT " & _
                           "WHERE DL01 = (SELECT MAX(DL01) FROM DATELIMIT)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_LastKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_LastKEY(1) = rsTmp.Fields("DL02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If IsEmptyText(textCP01) = True Or IsEmptyText(textCP02) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入本所案號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   If textCP01 = "TF" Then
      If IsEmptyText(textCP02_2) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
   End If
   
   ' 查詢資料
   If QueryDLFromCP() = False Then
      strTit = "查詢資料"
      strMsg = "無資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
EXITSUB:
End Sub

' Load Form
Private Sub Form_Load()
   SSTab1.Tab = 1

   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050707", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050707", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050707", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050707", strFind, False)
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   textKEY01.BackColor = &H8000000F
   textKEY02.BackColor = &H8000000F
   textKEY02_2.BackColor = &H8000000F
   textKEY03.BackColor = &H8000000F
   textKEY04.BackColor = &H8000000F
   
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
      m_FieldList(nIndex - 1).fiName = "DL" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 2, 3, 4:
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
   SetFieldNewData "DL01", textDL01
   If IsEmptyText(textDL02) = False Then
      SetFieldNewData "DL02", DBDATE(textDL02)
   Else
      SetFieldNewData "DL02", textDL02
   End If
   If IsEmptyText(textDL03) = False Then
      SetFieldNewData "DL03", DBDATE(textDL03)
   Else
      SetFieldNewData "DL03", textDL03
   End If
   If IsEmptyText(textDL04) = False Then
      SetFieldNewData "DL04", DBDATE(textDL04)
   Else
      SetFieldNewData "DL04", textDL04
   End If
   'Add By Cheng 2002/06/20
   SetFieldNewData "DL05", "" & textDL05
   'Add By Cheng 2002/06/26
   SetFieldNewData "DL06", textDL06
   
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
   'RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   textDL01 = Empty
   textDL02 = Empty
   textDL03 = Empty
   textDL04 = Empty
   textDL05 = Empty
   textDL06 = Empty
   
   textKEY01 = Empty
   textKEY02 = Empty
   textKEY02_2 = Empty
   textKEY03 = Empty
   textKEY04 = Empty
   
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textDL01.Locked = bEnable
   textDL02.Locked = bEnable
   textDL03.Locked = bEnable
   textDL04.Locked = bEnable
   'Add By Cheng 2002/06/20
   textDL05.Locked = bEnable
   'Add By Cheng 2002/06/26
   Me.textDL06.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textDL01.Locked = bEnable
   textDL02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM DATELIMIT " & _
            "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                  "DL02 = '" & m_CurrKEY(1) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("DL01")) = False Then
         textDL01 = rsTmp.Fields("DL01")
      End If
      If IsNull(rsTmp.Fields("DL02")) = False Then
         If rsTmp.Fields("DL02") <> "0" Then
            textDL02 = TAIWANDATE(rsTmp.Fields("DL02"))
         End If
      End If
      If IsNull(rsTmp.Fields("DL03")) = False Then
         If rsTmp.Fields("DL03") <> "0" Then
            textDL03 = TAIWANDATE(rsTmp.Fields("DL03"))
         End If
      End If
      If IsNull(rsTmp.Fields("DL04")) = False Then
         If rsTmp.Fields("DL04") <> "0" Then
            textDL04 = TAIWANDATE(rsTmp.Fields("DL04"))
         End If
      End If
      'Add By Cheng 2002/06/20
      If IsNull(rsTmp.Fields("DL05")) = False Then
         textDL05 = "" & rsTmp.Fields("DL05")
      End If
      'Add By Cheng 2002/06/26
      If IsNull(rsTmp.Fields("DL06")) = False Then
         textDL06 = rsTmp.Fields("DL06")
      End If
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
   End If
   rsTmp.Close
   
   strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & textDL01 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If rsTmp.Fields("CP01") = "TF" Then
         textKEY02_2.Visible = True
         If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = Mid(rsTmp.Fields("CP02"), 1, 5)
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02_2 = Mid(rsTmp.Fields("CP02"), 6, 1)
         If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
         If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
      Else
         textKEY02_2.Visible = False
         If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
         If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = rsTmp.Fields("CP02")
         If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
         If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
      End If
   Else
      textKEY02_2.Visible = False
      textKEY01 = Empty
      textKEY02 = Empty
      textKEY02_2 = Empty
      textKEY03 = Empty
      textKEY04 = Empty
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
               "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                     "DL02 = (SELECT MIN(DL02) FROM DATELIMIT " & _
                             "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                                   "DL02 > " & m_CurrKEY(1) & " )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
         If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
               "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT " & _
                              "WHERE DL01 > '" & m_CurrKEY(0) & "') AND " & _
                     "DL02 = (SELECT MIN(DL02) FROM DATELIMIT " & _
                              "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT " & _
                                             "WHERE DL01 > '" & m_CurrKEY(0) & "')) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
         If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
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
   m_CurrKEY(1) = m_FirstKEY(1)
   
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                  "DL02 = (SELECT MAX(DL02) FROM DATELIMIT " & _
                          "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                                "DL02 < " & m_CurrKEY(1) & " )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = (SELECT MAX(DL01) FROM DATELIMIT " & _
                           "WHERE DL01 < '" & m_CurrKEY(0) & "') AND " & _
                  "DL02 = (SELECT MAX(DL02) FROM DATELIMIT " & _
                           "WHERE DL01 = (SELECT MAX(DL01) FROM DATELIMIT " & _
                                          "WHERE DL01 < '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                  "DL02 = (SELECT MIN(DL02) FROM DATELIMIT " & _
                          "WHERE DL01 = '" & m_CurrKEY(0) & "' AND " & _
                                "DL02 > " & m_CurrKEY(1) & " )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT DL01,DL02 FROM DATELIMIT " & _
            "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT " & _
                           "WHERE DL01 > '" & m_CurrKEY(0) & "') AND " & _
                  "DL02 = (SELECT MIN(DL02) FROM DATELIMIT " & _
                           "WHERE DL01 = (SELECT MIN(DL01) FROM DATELIMIT " & _
                                          "WHERE DL01 > '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("DL01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("DL01")
      If IsNull(rsTmp.Fields("DL02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("DL02")
   End If
   rsTmp.Close
   
   UpdateCtrlData
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 顯示最後一筆資料
Private Sub ShowLastRecord()
   m_CurrKEY(0) = m_LastKEY(0)
   m_CurrKEY(1) = m_LastKEY(1)
   
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

Private Sub Text1_Change()

End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm050707 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
   If SSTab1.Tab = 1 Then
      CmdQuery.Default = True
      textCP01.SetFocus
   Else
      CmdQuery.Default = False
   End If
End Sub

Private Sub textCP01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 本所案號的系統別
Private Sub textCP01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse

   Cancel = False
   If IsEmptyText(textCP01) = False Then
      ' 使用者沒有權限
      If IsUserHasRightOfSystem(strUserNum, textCP01) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "您沒有使用該系統類別的權限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP01_GotFocus
         GoTo EXITSUB
      End If
      
      Select Case textCP01
         Case "TF":
            textCP02_2.Visible = True
            textCP02_2.Locked = False
            textCP02_2.TabStop = True
            textCP02.MaxLength = 5
         Case Else:
            textCP02_2.Visible = False
            textCP02_2.Locked = True
            textCP02_2.TabStop = False
            textCP02.MaxLength = 6
      End Select
   Else
      textCP02_2.Visible = False
      textCP02_2.Locked = True
      textCP02_2.TabStop = False
      textCP02.MaxLength = 6
   End If
EXITSUB:
End Sub

Private Sub textDL01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 總收文號
Private Sub textDL01_Validate(Cancel As Boolean)
   Dim rsTmp As ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textDL01) = False Then
      Select Case m_EditMode
         Case 1:
            If IsEmptyText(textDL02) = False Then
               If IsRecordExist(textDL01, DBDATE(textDL02)) = True Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "該筆延期記錄資料已存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textDL01_GotFocus
                  GoTo EXITSUB
               End If
            End If
            strSql = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS " & _
                     "WHERE CP09 = '" & textDL01 & "' "
            Set rsTmp = New ADODB.Recordset
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆收文記錄不存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textDL01_GotFocus
            Else
               If rsTmp.Fields("CP01") = "TF" Then
                  textKEY02_2.Visible = True
                  If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = Mid(rsTmp.Fields("CP02"), 1, 5)
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02_2 = Mid(rsTmp.Fields("CP02"), 6, 1)
                  If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
                  If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
               Else
                  textKEY02_2.Visible = False
                  If IsNull(rsTmp.Fields("CP01")) = False Then: textKEY01 = rsTmp.Fields("CP01")
                  If IsNull(rsTmp.Fields("CP02")) = False Then: textKEY02 = rsTmp.Fields("CP02")
                  If IsNull(rsTmp.Fields("CP03")) = False Then: textKEY03 = Mid(rsTmp.Fields("CP03"), 1, 5)
                  If IsNull(rsTmp.Fields("CP04")) = False Then: textKEY04 = Mid(rsTmp.Fields("CP04"), 1, 5)
               End If
            End If
            rsTmp.Close
            Set rsTmp = Nothing
      End Select
   End If
EXITSUB:
End Sub

' 延期日
Private Sub textDL02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDL02) = False Then
      If CheckIsTaiwanDate(textDL02, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延期日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDL02_GotFocus
         GoTo EXITSUB
      End If
      Select Case m_EditMode
         Case 1:
            If IsEmptyText(textDL01) = False Then
               If IsRecordExist(textDL01, DBDATE(textDL02)) = True Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "該筆延期記錄資料已存在"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  textDL02_GotFocus
                  GoTo EXITSUB
               End If
            End If
         Case Else:
      End Select
   End If
EXITSUB:
End Sub

' 原本所期限
Private Sub textDL03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDL03) = False Then
      If CheckIsTaiwanDate(textDL03, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原本所期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDL03_GotFocus
      End If
   End If
End Sub

' 原法定期限
Private Sub textDL04_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textDL04) = False Then
      If CheckIsTaiwanDate(textDL04, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原法定期限日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDL04_GotFocus
      End If
   End If
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      ' 90.07.13 modify by louis
      ' 新增
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd:
      '   If m_EditMode = 0 Then
      '      OnAction KeyCode
      '      KeyCode = 0
      '   End If
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
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
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

Private Sub textDL05_GotFocus()
   InverseTextBox textDL05
End Sub

Private Sub textDL05_KeyPress(KeyAscii As Integer)
'Add By Cheng 2002/06/20
If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textDL06_GotFocus()
  TextInverse Me.textDL06
End Sub

Private Sub textDL06_Validate(Cancel As Boolean)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

'若為新增或修改模式
If m_EditMode = 1 Or m_EditMode = 2 Then
   If Me.textDL05.Text = "2" Then
      If Me.textDL06.Text = "" Then
         MsgBox "下一程序序號不可為空白!!!", vbExclamation + vbOKOnly
         Cancel = True
         Me.textDL06.SetFocus
         TextInverse Me.textDL06
      Else
         StrSQLa = "Select * From NextProgress Where NP01='" & Me.textDL01.Text & "' And NP22=" & Val(Me.textDL06.Text)
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <= 0 Then
            MsgBox "下一程序序號輸入錯誤, 請重新輸入!!!", vbExclamation + vbOKOnly
            Cancel = True
            Me.textDL06.SetFocus
            TextInverse Me.textDL06
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   End If
End If
End Sub

Private Sub textKEY01_GotFocus()
  TextInverse textKEY01
End Sub

Private Sub textKEY02_2_GotFocus()
  TextInverse textKEY02_2
End Sub

Private Sub textKEY02_GotFocus()
  TextInverse textKEY02
End Sub

Private Sub textKEY03_GotFocus()
  TextInverse textKEY03
End Sub

Private Sub textKEY04_GotFocus()
  TextInverse textKEY04
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
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM DATELIMIT " & _
            "WHERE DL01 = '" & strKEY01 & "' AND " & _
                  "DL02 = " & strKEY02 & " "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
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
   Dim strDL01 As String
   Dim strDL02 As String
   
   strDL01 = textDL01
   strDL02 = DBDATE(textDL02)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(textDL01, textDL02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo EXITSUB
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO DATELIMIT ("
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
   
   If ((strDL01 & strDL02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strDL01 & strDL02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strDL01, strDL02
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
   Dim strDL01 As String
   Dim strDL02 As String
   
   strDL01 = m_CurrKEY(0)
   strDL02 = m_CurrKEY(1)
   
   strSql = "UPDATE DATELIMIT SET "
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
                  "WHERE DL01 = '" & strDL01 & "' AND " & _
                        "DL02 = '" & strDL02 & "' "
   
   If bDifference = True Then
      cnnConnection.Execute strSql
      ShowCurrRecord strDL01, strDL02
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strDL01 As String
   Dim strDL02 As String
   
   strDL01 = m_CurrKEY(0)
   strDL02 = m_CurrKEY(1)

   strSql = "DELETE FROM DATELIMIT " & _
            "WHERE DL01 = '" & strDL01 & "' AND " & _
                  "DL02 = '" & strDL02 & "' "

   cnnConnection.Execute strSql

   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strDL01 = m_LastKEY(0) And strDL02 = m_LastKEY(1)) Or (strDL01 = m_FirstKEY(0) And strDL02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strDL01, strDL02
   
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strDL01 As String
   Dim strDL02 As String
   
   strDL01 = textDL01
   strDL02 = DBDATE(textDL02)
   
   QueryRecord = False

   If IsRecordExist(strDL01, strDL02) = True Then
      m_CurrKEY(0) = strDL01
      m_CurrKEY(1) = strDL02
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
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
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
      Case 1: textDL01.SetFocus
      Case 2: textDL03.SetFocus
      Case 4: textDL01.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   
   Select Case m_EditMode
      Case 1, 2, 4:
         ' 總收文號不可空白
         If IsEmptyText(textDL01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入總收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL01.SetFocus
            GoTo EXITSUB
         End If
         ' 延期日不可為空白
         If IsEmptyText(textDL02) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入延期日"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL02.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   
   Select Case m_EditMode
      Case 1, 2:
         ' 原本所期限不可為空白
         If IsEmptyText(textDL03) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入原本所期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL03.SetFocus
            GoTo EXITSUB
         End If
         ' 原法定期限不可為空白
         If IsEmptyText(textDL04) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入原法定期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL04.SetFocus
            GoTo EXITSUB
         End If
         If Val(DBDATE(textDL03)) > Val(DBDATE(textDL04)) Then
            strTit = "檢核資料"
            strMsg = "原本所期限不可超過原法定期限"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL03.SetFocus
            GoTo EXITSUB
         End If
         ' 資料來源不可為空白
         If IsEmptyText(textDL05) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入資料來源"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDL05.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textDL01_GotFocus()
   InverseTextBox textDL01
End Sub

Private Sub textDL02_GotFocus()
   InverseTextBox textDL02
End Sub

Private Sub textDL03_GotFocus()
   InverseTextBox textDL03
End Sub

Private Sub textDL04_GotFocus()
   InverseTextBox textDL04
End Sub

Private Sub textCP01_GotFocus()
   InverseTextBox textCP01
End Sub

Private Sub textCP02_GotFocus()
   InverseTextBox textCP02
End Sub

Private Sub textCP02_2_GotFocus()
   InverseTextBox textCP02_2
End Sub

Private Sub textCP03_GotFocus()
   InverseTextBox textCP03
End Sub

Private Sub textCP04_GotFocus()
   InverseTextBox textCP04
End Sub

' 初始化列表
Public Sub InitialGridList()
   grdList.Clear
   grdList.Rows = 1
'   grdList.Cols = 5
   grdList.Cols = 6

   grdList.ColWidth(0) = 300
   grdList.row = 0

   grdList.col = 0
   grdList.ColAlignment(0) = flexAlignCenterCenter
   grdList.col = 1
   grdList.Text = "總收文號"
   grdList.ColWidth(1) = 1000
   grdList.ColAlignment(1) = flexAlignCenterCenter
   grdList.col = 2
   grdList.Text = "延期日"
   grdList.ColWidth(2) = 1000
   grdList.ColAlignment(2) = flexAlignCenterCenter
   grdList.col = 3
   grdList.Text = "原本所期限"
   grdList.ColWidth(3) = 1000
   grdList.ColAlignment(3) = flexAlignCenterCenter
   grdList.col = 4
   grdList.Text = "原法定期限"
   grdList.ColWidth(4) = 1000
   grdList.ColAlignment(4) = flexAlignCenterCenter
   'Add By Cheng 2002/06/20
   grdList.col = 5
   grdList.Text = "資料來源"
   grdList.ColWidth(5) = 1200
   grdList.ColAlignment(5) = flexAlignCenterCenter
End Sub

Private Sub grdList_Click()
   grdList_ShowSelection
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
   grdList_ShowSelection
   
   If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
      nRow = grdList.row
      ShowCurrRecord grdList.TextMatrix(nRow, 1), DBDATE(grdList.TextMatrix(nRow, 2))
   End If
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

Private Function QueryDLFromCP() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim nRow As Integer
   
   QueryDLFromCP = False
   
   ' 組成本所案號
   strCP01 = textCP01
   strCP02 = textCP02
   If strCP01 = "TF" Then: strCP02 = strCP02 & textCP02_2
   strCP03 = textCP03
   If IsEmptyText(strCP03) = True Then: strCP03 = "0"
   strCP04 = textCP04
   If IsEmptyText(strCP04) = True Then: strCP04 = "00"
   
   InitialGridList
   
   'Modify By Cheng 2002/06/20
'   strSQL = "SELECT DL01, NVL(DL02 - 19110000, NULL) AS DL02, NVL(DL03 - 19110000, NULL) AS DL03, NVL(DL04 - 19110000, NULL) AS DL04, CP09 FROM DATELIMIT, CASEPROGRESS " & _
'            "WHERE DL01 = CP09 AND " & _
'                  "CP01 = '" & strcp01 & "' AND " & _
'                  "CP02 = '" & strcp02 & "' AND " & _
'                  "CP03 = '" & strcp03 & "' AND " & _
'                  "CP04 = '" & strcp04 & "' " & _
'            "ORDER BY DL01, DL02 "
   strSql = "SELECT DL01, NVL(DL02 - 19110000, NULL) AS DL02, NVL(DL03 - 19110000, NULL) AS DL03, NVL(DL04 - 19110000, NULL) AS DL04, CP09,DECODE(DL05,'1','案件進度檔','2','下一程序檔','') AS DL05 FROM DATELIMIT, CASEPROGRESS " & _
            "WHERE DL01 = CP09 AND " & _
                  "CP01 = '" & strCP01 & "' AND " & _
                  "CP02 = '" & strCP02 & "' AND " & _
                  "CP03 = '" & strCP03 & "' AND " & _
                  "CP04 = '" & strCP04 & "' " & _
            "ORDER BY DL01, DL02 "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      QueryDLFromCP = True
      UpdateGridList rsTmp
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
   Dim nRow As Integer
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      grdList.Rows = grdList.Rows + 1
      nRow = grdList.Rows - 1
      If IsNull(rsTmp.Fields("DL01")) = False Then
         grdList.TextMatrix(nRow, 1) = rsTmp.Fields("DL01")
      End If
      If IsNull(rsTmp.Fields("DL02")) = False Then
         grdList.TextMatrix(nRow, 2) = rsTmp.Fields("DL02")
      End If
      If IsNull(rsTmp.Fields("DL03")) = False Then
         grdList.TextMatrix(nRow, 3) = rsTmp.Fields("DL03")
      End If
      If IsNull(rsTmp.Fields("DL04")) = False Then
         grdList.TextMatrix(nRow, 4) = rsTmp.Fields("DL04")
      End If
      'Add By Cheng 2002/06/20
      If IsNull(rsTmp.Fields("DL05")) = False Then
         grdList.TextMatrix(nRow, 5) = rsTmp.Fields("DL05")
      End If
      rsTmp.MoveNext
   Loop
   
   grdList.FixedRows = 1 'Added by Lydia 2023/10/16
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCP01.Enabled = True Then
   Cancel = False
   textCP01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDL01.Enabled = True Then
   Cancel = False
   textDL01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDL02.Enabled = True Then
   Cancel = False
   textDL02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDL03.Enabled = True Then
   Cancel = False
   textDL03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDL04.Enabled = True Then
   Cancel = False
   textDL04_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textDL06.Enabled = True Then
   Cancel = False
   textDL06_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

