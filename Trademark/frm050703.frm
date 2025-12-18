VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050703 
   BorderStyle     =   1  '單線固定
   Caption         =   "信函Initial資料維護"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7545
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8580
      Top             =   600
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
            Picture         =   "frm050703.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050703.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   660
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   7545
      _ExtentX        =   13309
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4872
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   6972
      _ExtentX        =   12303
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm050703.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "LabelID01"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "LabelID02"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textID02"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textID01"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textID03"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm050703.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(1)"
      Tab(1).Control(1)=   "txt1(0)"
      Tab(1).Control(2)=   "cmdQuery"
      Tab(1).Control(3)=   "GRD1"
      Tab(1).Control(4)=   "LabelName(1)"
      Tab(1).Control(5)=   "LabelName(0)"
      Tab(1).Control(6)=   "Label4"
      Tab(1).Control(7)=   "Label1(1)"
      Tab(1).ControlCount=   8
      Begin VB.TextBox txt1 
         Height          =   276
         Index           =   1
         Left            =   -71490
         MaxLength       =   6
         TabIndex        =   4
         Top             =   540
         Width           =   795
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74190
         MaxLength       =   6
         TabIndex        =   3
         Top             =   540
         Width           =   795
      End
      Begin VB.TextBox textID03 
         Height          =   276
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1290
         Width           =   1305
      End
      Begin VB.TextBox textID01 
         Height          =   276
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   0
         Top             =   585
         Width           =   795
      End
      Begin VB.TextBox textID02 
         Height          =   270
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   1
         Top             =   945
         Width           =   795
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Left            =   -69144
         TabIndex        =   5
         Top             =   384
         Width           =   912
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm050703.frx":212C
         Height          =   3885
         Left            =   -74820
         TabIndex        =   16
         Top             =   870
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   6853
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSForms.Label LabelID02 
         Height          =   270
         Left            =   2340
         TabIndex        =   18
         Top             =   960
         Width           =   1695
         Caption         =   "LabelID02"
         Size            =   "2990;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LabelID01 
         Height          =   270
         Left            =   2340
         TabIndex        =   17
         Top             =   600
         Width           =   1665
         Caption         =   "LabelID01"
         Size            =   "2937;476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "請輸入完整Initial（Ex: FC/phc）"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2790
         TabIndex        =   15
         Top             =   1320
         Width           =   2460
      End
      Begin VB.Label LabelName 
         AutoSize        =   -1  'True
         Caption         =   "LabelName"
         Height          =   180
         Index           =   1
         Left            =   -70650
         TabIndex        =   14
         Top             =   570
         Width           =   795
      End
      Begin VB.Label LabelName 
         AutoSize        =   -1  'True
         Caption         =   "LabelName"
         Height          =   180
         Index           =   0
         Left            =   -73350
         TabIndex        =   13
         Top             =   570
         Width           =   795
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "組員 : "
         Height          =   180
         Left            =   -74700
         TabIndex        =   12
         Top             =   570
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "判發主管 :"
         Height          =   180
         Index           =   1
         Left            =   -72330
         TabIndex        =   11
         Top             =   570
         Width           =   810
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Initial : "
         Height          =   180
         Left            =   450
         TabIndex        =   10
         Top             =   1320
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "組員 : "
         Height          =   180
         Left            =   450
         TabIndex        =   9
         Top             =   615
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "判發主管 :"
         Height          =   180
         Index           =   0
         Left            =   450
         TabIndex        =   7
         Top             =   975
         Width           =   810
      End
   End
End
Attribute VB_Name = "frm050703"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/13 改成Form2.0 ; LabelID01、LabelID02
'Create By Sindy 2020/9/14
Option Explicit

Const MAX_FIELD = 3

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

Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean


Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData ) AND " & _
                  "ID02 = (SELECT MIN(ID02) FROM InitialData " & _
                           "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData)) "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("ID02")
   End If
   rsTmp.Close

   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = (SELECT MAX(ID01) FROM InitialData ) AND " & _
                  "ID02 = (SELECT MAX(ID02) FROM InitialData " & _
                           "WHERE ID01 = (SELECT MAX(ID01) FROM InitialData)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_LastKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_LastKEY(1) = rsTmp.Fields("ID02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   
'   If IsEmptyText(Txt1(0)) = True And IsEmptyText(Txt1(1)) = True Then
'      strTit = "檢核資料"
'      strMsg = "請至少輸入一項查詢條件"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      Txt1(0).SetFocus
'      GoTo EXITSUB
'   End If
   
   ' 查詢資料
   If QueryList() = False Then
      If Me.SSTab1.Tab = 1 Then
         strTit = "查詢資料"
         strMsg = "無資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         If txt1(0) <> "" Then
            TextInverse txt1(0)
         ElseIf txt1(1) <> "" Then
            TextInverse txt1(1)
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm050703", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050703", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050703", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050703", strFind, False)
   
   m_EditMode = 0
   m_SubMode = 0
   MoveFormToCenter Me
   
   ClearField
   LabelName(0).Caption = ""
   LabelName(1).Caption = ""
   
   SetGrd
   InitialField
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   'tlbar.Buttons(2).Visible = False '修改圖示不顯示
   
   Me.SSTab1.Tab = 0
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To MAX_FIELD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "ID" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0
      Select Case nIndex
         Case 4:
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
   SetFieldNewData "ID01", textID01
   SetFieldNewData "ID02", textID02
   SetFieldNewData "ID03", textID03
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To MAX_FIELD - 1
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

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   textID01 = Empty
   textID02 = Empty
   textID03 = Empty
   LabelID01.Caption = ""
   LabelID02.Caption = ""
   For nIndex = 0 To MAX_FIELD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textID01.Locked = bEnable
   textID02.Locked = bEnable
   textID03.Locked = bEnable
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textID01.Locked = bEnable
   textID02.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM InitialData " & _
            "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ID02 = '" & m_CurrKEY(1) & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      textID01 = rsTmp.Fields("ID01"): textID01_Validate False
      textID02 = rsTmp.Fields("ID02"): textID02_Validate False
      textID03 = "" & rsTmp.Fields("ID03")
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
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
      strSql = "SELECT ID01,ID02 FROM InitialData " & _
               "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                     "ID02 = (SELECT MIN(ID02) FROM InitialData " & _
                             "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                                   "ID02 > '" & m_CurrKEY(1) & "' )"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
         If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT ID01,ID02 FROM InitialData " & _
               "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData " & _
                              "WHERE ID01 > '" & m_CurrKEY(0) & "') AND " & _
                     "ID02 = (SELECT MIN(ID02) FROM InitialData " & _
                              "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData " & _
                                             "WHERE ID01 > '" & m_CurrKEY(0) & "')) "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
         If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
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
   
   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ID02 = (SELECT MAX(ID02) FROM InitialData " & _
                          "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                                "ID02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = (SELECT MAX(ID01) FROM InitialData " & _
                           "WHERE ID01 < '" & m_CurrKEY(0) & "') AND " & _
                  "ID02 = (SELECT MAX(ID02) FROM InitialData " & _
                           "WHERE ID01 = (SELECT MAX(ID01) FROM InitialData " & _
                                          "WHERE ID01 < '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
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
   
   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                  "ID02 = (SELECT MIN(ID02) FROM InitialData " & _
                          "WHERE ID01 = '" & m_CurrKEY(0) & "' AND " & _
                                "ID02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT ID01,ID02 FROM InitialData " & _
            "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData " & _
                           "WHERE ID01 > '" & m_CurrKEY(0) & "') AND " & _
                  "ID02 = (SELECT MIN(ID02) FROM InitialData " & _
                           "WHERE ID01 = (SELECT MIN(ID01) FROM InitialData " & _
                                          "WHERE ID01 > '" & m_CurrKEY(0) & "')) "

   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ID01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("ID01")
      If IsNull(rsTmp.Fields("ID02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("ID02")
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

Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
   If SSTab1.Tab = 1 Then
      cmdQuery.Default = True
      txt1(0).SetFocus
   Else
      cmdQuery.Default = False
   End If
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
         SSTab1.Tab = 0
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         If textID01 = "" Or textID02 = "" Then Exit Sub
         ' 檢查記錄是否已存在
         If IsRecordExist(textID01, textID02) = True Then
            strTit = "詢問"
            strMsg = "是否要刪除此筆資料?"
            nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
            If nResponse = vbYes Then
               m_EditMode = 3
               OnWork
               UpdateToolbarState
            End If
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
   strSql = "SELECT * FROM InitialData " & _
            "WHERE ID01 = '" & strKEY01 & "' AND " & _
                  "ID02 = '" & strKEY02 & "'"
                  
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
   Dim strID01 As String
   Dim strID02 As String
   
   strID01 = textID01
   strID02 = textID02
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO InitialData ("
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
   
   If ((strID01 & strID02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strID01 & strID02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   ShowCurrRecord strID01, strID02
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
   Dim strID01 As String
   Dim strID02 As String
   
   strID01 = m_CurrKEY(0)
   strID02 = m_CurrKEY(1)
   
'   If IsRecordExist(textID01.Text, textID02.Text) = True Then
'      strTit = "修改資料"
'      strMsg = "該筆記錄已存在"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      UpdateCtrlData
'      Exit Sub
'   End If
   
   strSql = "UPDATE InitialData SET "
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
                  "WHERE ID01 = '" & strID01 & "' AND " & _
                        "ID02 = '" & strID02 & "' "
   
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      ShowCurrRecord strID01, strID02
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strID01 As String
   Dim strID02 As String
   
   strID01 = m_CurrKEY(0)
   strID02 = m_CurrKEY(1)

   strSql = "DELETE FROM InitialData " & _
            "WHERE ID01 = '" & strID01 & "' AND " & _
                  "ID02 = '" & strID02 & "' "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   Call ClearField 'Add By Sindy 2020/9/15
   
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆的本所案號
   If (strID01 = m_LastKEY(0) And strID02 = m_LastKEY(1)) Or (strID01 = m_FirstKEY(0) And strID02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strID01, strID02
   
   If Me.SSTab1.Tab = 1 Then Call cmdQuery_Click
EXITSUB:
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strID01 As String
   Dim strID02 As String
   
   strID01 = textID01
   strID02 = textID02
   
   QueryRecord = False

   If IsRecordExist(strID01, strID02) = True Then
      m_CurrKEY(0) = strID01
      m_CurrKEY(1) = strID02
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
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            
            AddRecord
            RefreshRange
         Else
            GoTo EXITSUB
         End If
      Case 2:
         If CheckDataValid() = True Then
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
      Case 1: textID01.SetFocus
      Case 2: textID03.SetFocus
      Case 4: textID01.SetFocus
   End Select
End Sub

Private Function CheckDataValid() As Boolean
Dim Cancel As Boolean
Dim s As Integer
   
   CheckDataValid = False
   
   If m_EditMode = 1 Then
      ' 檢查記錄是否已存在
      If IsRecordExist(textID01, textID02) = True Then
         s = MsgBox("該筆記錄已存在！", , "新增資料！")
         Exit Function
      End If
   End If
   
   If textID01 = textID02 Then
      'Modify By Sindy 2021/10/20
      'MsgBox "組員不可與判發主管同一人！", vbInformation, "資料錯誤"
      If MsgBox("確定組員與判發主管同一人嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         Exit Function
      End If
      '2021/10/20 END
   End If
   
   If Me.textID01.Enabled = True Then
      Cancel = False
      textID01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textID02.Enabled = True Then
      Cancel = False
      textID02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textID01_GotFocus()
   InverseTextBox textID01
End Sub
Private Sub textID01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textID01_Validate(Cancel As Boolean)
Dim s As Integer
   Cancel = False
   LabelID01.Caption = ""
   If IsEmptyText(textID01) = False Then
      LabelID01.Caption = GetPrjSalesNM(textID01)
      If LabelID01.Caption = "" Then
         Cancel = True
         s = MsgBox("組員輸入錯誤！", , "錯誤！")
         textID01.SetFocus
         textID01_GotFocus
         Exit Sub
      End If
   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If IsEmptyText(textID01) = True Then
         Cancel = True
         s = MsgBox("組員不可空白！", , "錯誤！")
         textID01.SetFocus
         textID01_GotFocus
         Exit Sub
      Else
'         If m_EditMode = 1 Then
'            ' 檢查記錄是否已存在
'            If IsRecordExist(textID01, textID02) = True Then
'               s = MsgBox("該筆記錄已存在！", , "新增資料！")
'               textID01.SetFocus
'               textID01_GotFocus
'               Exit Sub
'               'UpdateCtrlData
'               'GoTo EXITSUB
'            End If
'         End If
'
'         If textID01 = textID02 Then
'            MsgBox "組員不可與判發主管同一人！", vbInformation, "資料錯誤"
'            textID01.SetFocus
'            textID01_GotFocus
'            Exit Sub
'         End If
         If textID03 = "" Then
            textID03 = QueryInitial
         End If
      End If
   End If
End Sub

Private Sub textID02_GotFocus()
   InverseTextBox textID02
End Sub
Private Sub textID02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textID02_Validate(Cancel As Boolean)
Dim s As Integer
   
   Cancel = False
   LabelID02.Caption = ""
   If IsEmptyText(textID02) = False Then
      LabelID02.Caption = GetPrjSalesNM(textID02)
      If LabelID02.Caption = "" Then
         Cancel = True
         s = MsgBox("判發主管輸入錯誤！", , "錯誤！")
         textID02.SetFocus
         textID02_GotFocus
         Exit Sub
      End If
   End If
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If IsEmptyText(textID02) = True Then
         Cancel = True
         s = MsgBox("判發主管不可空白！", , "錯誤！")
         textID02.SetFocus
         textID02_GotFocus
         Exit Sub
      Else
'         If m_EditMode = 1 Then
'            ' 檢查記錄是否已存在
'            If IsRecordExist(textID01, textID02) = True Then
'               s = MsgBox("該筆記錄已存在！", , "新增資料！")
'               textID02.SetFocus
'               textID02_GotFocus
'               Exit Sub
'               'UpdateCtrlData
'               'GoTo EXITSUB
'            End If
'         End If
'
'         If textID01 = textID02 Then
'            MsgBox "組員不可與判發主管同一人！", vbInformation, "資料錯誤"
'            textID02.SetFocus
'            textID02_GotFocus
'            Exit Sub
'         End If
         If textID03 = "" Then
            textID03 = QueryInitial
         End If
      End If
   End If
End Sub

Private Sub textID03_GotFocus()
   InverseTextBox textID03
End Sub

'組出Initial
Private Function QueryInitial() As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If textID01 = "" Or textID02 = "" Then Exit Function
   QueryInitial = ""
   
   strSql = "select st17 from staff where st01='" & textID02 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Trim("" & rsTmp.Fields("st17")) = "" Then
         Exit Function
      End If
      QueryInitial = UCase(rsTmp.Fields("st17"))
      
      rsTmp.Close
      strSql = "select st17 from staff where st01='" & textID01 & "'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If Trim("" & rsTmp.Fields("st17")) = "" Then
            Exit Function
         End If
         QueryInitial = QueryInitial & "/" & LCase(rsTmp.Fields("st17"))
      End If
   End If
   rsTmp.Close
End Function

Private Function QueryList() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   QueryList = False
   
   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and ID01='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and ID02='" & txt1(1) & "' "
   End If
   
   GRD1.Clear
   strSql = "SELECT ID01,s1.st02,ID02,s2.st02,ID03 FROM InitialData,staff s1,staff s2" & _
            " WHERE 1=1 " & strSql & " and ID01=s1.st01(+) and ID02=s2.st01(+)" & _
            " Order by ID01,ID02"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      QueryList = True
   End If
   SetGrd
   GRD1.row = 1
   GRD1.col = 0
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
Dim tmpMouseRow
Dim i, j
GRD1.Visible = False
tmpMouseRow = GRD1.row
GRD1.Visible = True
If tmpMouseRow <> 0 Then
    GRD1.row = tmpMouseRow
    GRD1.col = 0
    If GRD1.CellBackColor <> &HFFC0C0 Then
                  GRD1.Visible = False
         For j = 1 To GRD1.Rows - 1
             GRD1.row = j
             For i = 0 To GRD1.Cols - 1
                  GRD1.col = i
                  GRD1.CellBackColor = QBColor(15)
             Next i
        Next j
        GRD1.row = tmpMouseRow
         For i = 0 To GRD1.Cols - 1
             GRD1.col = i
             GRD1.CellBackColor = &HFFC0C0
         Next i
         textID01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textID02.Text = GRD1.TextMatrix(tmpMouseRow, 2)
         QueryRecord
         GRD1.Visible = True
    End If
End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("組員編號", "姓名", "判發主管", "姓名", "Initial")
   arrGridHeadWidth = Array(1000, 1200, 1000, 1200, 1500)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Function TxtValidate() As Boolean
Dim s As Integer
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   TxtValidate = False
   
   If IsEmptyText(textID03) = True Then
      s = MsgBox("Initial不可空白！", , "資料錯誤！")
      textID03.SetFocus
      textID03_GotFocus
      Exit Function
   End If
   If InStr(textID03.Text, "/") = 0 Then
      s = MsgBox("請輸入完整的Initial！", , "資料錯誤！")
      textID03.SetFocus
      textID03_GotFocus
      Exit Function
   End If
   
   'Add By Sindy 2021/10/20 檢查是否有已存在的 Initial
   strSql = "select st01,st02,st17 from staff where st17='" & textID03 & "' and st04='1'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      s = MsgBox("此Initial已存在,是 " & rsTmp.Fields("st02") & " 在使用！", , "資料錯誤！")
      textID03.SetFocus
      textID03_GotFocus
      Exit Function
   End If
   rsTmp.Close
   '2021/10/20 END
   
   Set rsTmp = Nothing
   TxtValidate = True
End Function

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
'      Case 2, 3
'         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         ' 檢查員工編號規則
         LabelName(Index) = ""
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            Else
               LabelName(Index) = GetPrjSalesNM(txt1(Index))
            End If
         End If
'         If Index = 0 Then
'            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
'               txt1(Index + 1) = txt1(Index)
'            End If
'         ElseIf Index = 1 Then
'            If RunNick(txt1(Index - 1), txt1(Index)) Then
'               Call txt1_GotFocus(Index)
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         
'      Case 2, 3
'         ' 2008/12/16 MODIFY BY SINDY
'         'If CheckIsTaiwanDate(txt1(Index), False) = False Then
'         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
'         ' 2008/12/16 END
'            Call txt1_GotFocus(Index)
'            Cancel = True
'            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
'            Exit Sub
'         End If
'
'         ' 2008/12/17 ADD BY SINDY
'         If Index = 2 Then
'            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
'               txt1(Index + 1) = txt1(Index)
'            End If
'         ' 2008/12/17 END
'         ElseIf Index = 3 Then
'            If RunNick2(txt1(Index - 1), txt1(Index)) Then
'               Call txt1_GotFocus(Index)
'               Cancel = True
'               Exit Sub
'            End If
'         End If
         
      Case Else
   End Select
End Sub
