VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm140422 
   BorderStyle     =   1  '單線固定
   Caption         =   "設定颱風假作業"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8190
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   30
      TabIndex        =   11
      Top             =   660
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm140422.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(4)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textWD03"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textWD01"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textWD02"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textWD04"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textWD05"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textWD06"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textWD07"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm140422.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "cmdok"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "txt1(1)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txt1(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label16"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Line5"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.TextBox textWD07 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   6
         Top             =   2430
         Width           =   345
      End
      Begin VB.TextBox textWD06 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   5
         Top             =   2095
         Width           =   345
      End
      Begin VB.TextBox textWD05 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   4
         Top             =   1763
         Width           =   345
      End
      Begin VB.TextBox textWD04 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1431
         Width           =   345
      End
      Begin VB.TextBox textWD02 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   1
         Top             =   767
         Width           =   345
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm140422.frx":0038
         Height          =   3615
         Left            =   -75000
         TabIndex        =   12
         Top             =   720
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6368
         _Version        =   393216
         FixedCols       =   0
         ScrollTrack     =   -1  'True
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
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -71760
         TabIndex        =   9
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73080
         MaxLength       =   7
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74130
         MaxLength       =   7
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textWD01 
         Height          =   285
         Left            =   2370
         TabIndex        =   0
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox textWD03 
         Height          =   270
         Left            =   2370
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1099
         Width           =   345
      End
      Begin VB.Label Label2 
         Caption         =   $"frm140422.frx":004D
         ForeColor       =   &H000000FF&
         Height          =   1420
         Left            =   360
         TabIndex        =   21
         Top             =   2880
         Width           =   7270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否手動下載電子公文：          (Y:是)"
         Height          =   180
         Index           =   4
         Left            =   360
         TabIndex        =   20
         Top             =   2520
         Width           =   2830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否補班：          (Y:是)"
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   19
         Top             =   2190
         Width           =   1870
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "高所颱風假：          (Y:是)"
         Height          =   180
         Index           =   2
         Left            =   1260
         TabIndex        =   18
         Top             =   1830
         Width           =   2050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "南所颱風假：          (Y:是)"
         Height          =   180
         Index           =   1
         Left            =   1260
         TabIndex        =   17
         Top             =   1500
         Width           =   2050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "北所颱風假：          (Y:是)"
         Height          =   180
         Index           =   0
         Left            =   1260
         TabIndex        =   16
         Top             =   840
         Width           =   2050
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期起："
         Height          =   180
         Left            =   -74850
         TabIndex        =   15
         Top             =   390
         Width           =   720
      End
      Begin VB.Line Line5 
         X1              =   -73410
         X2              =   -72810
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   1800
         TabIndex        =   14
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "中所颱風假：          (Y:是)"
         Height          =   180
         Index           =   17
         Left            =   1260
         TabIndex        =   13
         Top             =   1170
         Width           =   2050
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm140422.frx":018C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":04A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":07C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":09A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":0CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":0FD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":12F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":1610
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":192C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":1C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm140422.frx":1F64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
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
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frm140422"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2025/5/9
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_WD As Integer
Dim MyKind As String


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) <> "" Then
       If RunNick(txt1(0), txt1(1)) Then
           txt1(0).SetFocus
           Exit Sub
       End If
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
       txt1(0).SetFocus
   End If
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from WorkDay where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_WD = rsA.Fields.Count
   SetGrd
End Sub

' 按下按鍵
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
'         If m_bInsert Then
'            If m_EditMode = 0 Then
'               OnAction KeyCode
'               KeyCode = 0
'            End If
'         End If
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
'         If m_bDelete Then
'            If m_EditMode = 0 Then
'               OnAction KeyCode
'               KeyCode = 0
'            End If
'         End If
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
'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

Private Sub Form_Load()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ReDim m_FieldList(tf_WD) As FIELDITEM

'   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
'   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   textWD01.BackColor = &H8000000F
   
   MoveFormToCenter Me

   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
   
   Tbar1_ButtonClick TBar1.Buttons(4) '設定為按下查詢鍵
   textWD01.Text = strSrvDate(2)
   OnAction vbKeyF9 '按確定
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140422 = Nothing
End Sub

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
            textWD01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            QueryRecord
            GRD1.Visible = True
       End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdok.SetFocus
      cmdok.Default = True
   Else
      cmdok.Default = False
   End If
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub ShowMsg(ByVal St As String)
   MsgBox St, vbInformation
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

   TxtValidate = False

   If textWD01.Text = "" Then
       MsgBox "日期不可以空白！", vbExclamation
       textWD01.SetFocus
       Exit Function
   End If

   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer

   For nIndex = 0 To tf_WD - 1
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

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
Dim nIndex As Integer
Dim strTmp As String

   For nIndex = 0 To tf_WD - 1
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

' 修改記錄
Private Function ModRecord() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strWD01 As String
   Dim strWD02 As String

   ModRecord = False

   strWD01 = m_CurrKEY(0)

   strSql = "begin user_data.user_enabled:=1; UPDATE WorkDay SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_WD - 1
      strTmp = Empty
      'If nIndex < 3 Or nIndex > 8 Then
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
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
        'End If
   Next nIndex

   strSql = strSql & " " & _
                  "WHERE WD01 = " & strWD01 & "; end; "
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord DBDATE(strWD01)

   ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)

End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   QueryRecord = False
   If IsRecordExist(textWD01) = True Then
      m_CurrKEY(0) = DBDATE(textWD01)
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Function OnWork() As Boolean
Dim strMsg As String
Dim strTit As String
Dim nResponse

   OnWork = False
   Select Case m_EditMode
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 4: '查詢
         If textWD01 <> "" Then
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
   OnWork = True
EXITSUB:
End Function

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
'      Case 1: If Me.Visible = True Then textWD01.SetFocus
      Case 2: If Me.Visible = True Then textWD02.SetFocus
'      Case 4: If Me.Visible = True Then textWD01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   IsRecordExist = False
   strSql = "SELECT * FROM WorkDay " & _
            "WHERE WD01 = " & DBDATE(strKEY01)

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

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT WD01 FROM WorkDay " & _
               "WHERE WD01 = " & m_CurrKEY(0)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("WD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("WD01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close

      strSql = "SELECT min(WD01) FROM WorkDay "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("WD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("WD01")
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

   strSql = "SELECT MAX(WD01) WD01 FROM WorkDay " & _
            "WHERE WD01 < " & m_CurrKEY(0)
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("WD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("WD01")
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

   strSql = "SELECT MIN(WD01) WD01 FROM WorkDay " & _
            "WHERE WD01 > " & m_CurrKEY(0)
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("WD01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("WD01")
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

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse

   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
'         m_EditMode = 1
'         ClearField
'         Me.SSTab1.TabEnabled(1) = False
'         SSTab1.Tab = 0
'         SetCtrlReadOnly False
'         UpdateToolbarState
'         SetInputEntry
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
'         strTit = "詢問"
'         strMsg = "是否要刪除此筆資料?"
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'         If nResponse = vbYes Then
'            m_EditMode = 3
'            If OnWork = True Then
'                UpdateToolbarState
'            Else
'                Exit Sub
'            End If
'         End If
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
         If OnWork = True Then
            Me.SSTab1.TabEnabled(1) = True
            UpdateToolbarState
         Else
            Exit Sub
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
                  Me.SSTab1.TabEnabled(1) = True
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               End If
            Case Else
               m_EditMode = 0
               Me.SSTab1.TabEnabled(1) = True
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset

   strSql = "SELECT MIN(WD01) WD01 FROM WorkDay "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("WD01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("WD01")
   End If
   rsTmp.Close

   strSql = "SELECT MAX(WD01) WD01 FROM WorkDay "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("WD01")) = False Then: m_LastKEY(0) = rsTmp.Fields("WD01")
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer

   strSql = "SELECT * FROM WorkDay " & _
            "WHERE WD01=" & DBDATE(m_CurrKEY(0))
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("WD01")) = False Then: textWD01 = TAIWANDATE(rsTmp.Fields("WD01"))
      If IsNull(rsTmp.Fields("WD02")) = False Then: textWD02 = rsTmp.Fields("WD02")
      If IsNull(rsTmp.Fields("WD03")) = False Then: textWD03 = rsTmp.Fields("WD03")
      If IsNull(rsTmp.Fields("WD04")) = False Then: textWD04 = rsTmp.Fields("WD04")
      If IsNull(rsTmp.Fields("WD05")) = False Then: textWD05 = rsTmp.Fields("WD05")
      If IsNull(rsTmp.Fields("WD06")) = False Then: textWD06 = rsTmp.Fields("WD06")
      If IsNull(rsTmp.Fields("WD07")) = False Then: textWD07 = rsTmp.Fields("WD07")
      
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close

EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and WD01>=" & DBDATE(txt1(0))
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and WD01<=" & DBDATE(txt1(1))
   End If
   '抓取資料
   strSql = "SELECT sqldateT(WD01),WD02,WD03,WD04,WD05,WD06,WD07" & _
            " FROM WorkDay where 1=1 " & strSql & _
            " order by WD01 "
   If rsTmp.State = 1 Then rsTmp.Close
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set GRD1.Recordset = rsTmp
   SetGrd
End Sub

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            TBar1.Buttons(1).Enabled = True
         Else
            TBar1.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            TBar1.Buttons(2).Enabled = True
         Else
            TBar1.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            TBar1.Buttons(3).Enabled = True
         Else
            TBar1.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
            TBar1.Buttons(6).Enabled = True
            TBar1.Buttons(7).Enabled = True
            TBar1.Buttons(8).Enabled = True
            TBar1.Buttons(9).Enabled = True
         Else
            TBar1.Buttons(6).Enabled = False
            TBar1.Buttons(7).Enabled = False
            TBar1.Buttons(8).Enabled = False
            TBar1.Buttons(9).Enabled = False
         End If
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         TBar1.Buttons(1).Enabled = False
         TBar1.Buttons(2).Enabled = False
         TBar1.Buttons(3).Enabled = False
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
   End Select
End Sub

Private Function CheckDataValid() As Boolean
Dim nResponse As Boolean
Dim strTmp  As String

   CheckDataValid = False
   
'   nResponse = False
'   textWD01_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
'   nResponse = False
'   textWD02_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
'   nResponse = False
'   textWD03_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textWD01.Locked = bEnable
   If bEnable Then textWD01.BackColor = &H8000000F Else textWD01.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer

   textWD01.Locked = bEnable
   If bEnable Then textWD01.BackColor = &H8000000F Else textWD01.BackColor = &H80000005
   textWD02.Locked = bEnable
   textWD03.Locked = bEnable
   textWD04.Locked = bEnable
   textWD05.Locked = bEnable
   textWD06.Locked = bEnable
   textWD07.Locked = bEnable
   If bEnable Then textWD02.BackColor = &H8000000F Else textWD02.BackColor = &H80000005
   If bEnable Then textWD03.BackColor = &H8000000F Else textWD03.BackColor = &H80000005
   If bEnable Then textWD04.BackColor = &H8000000F Else textWD04.BackColor = &H80000005
   If bEnable Then textWD05.BackColor = &H8000000F Else textWD05.BackColor = &H80000005
   If bEnable Then textWD06.BackColor = &H8000000F Else textWD06.BackColor = &H80000005
   If bEnable Then textWD07.BackColor = &H8000000F Else textWD07.BackColor = &H80000005
End Sub

Private Sub ClearField()
Dim nIndex As Integer

   textWD01 = Empty
   textWD02 = Empty
   textWD03 = Empty
   textWD04 = Empty
   textWD05 = Empty
   textWD06 = Empty
   textWD07 = Empty
   SetGrd
   For nIndex = 0 To tf_WD - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "WD01", DBDATE(textWD01)
   End If
   SetFieldNewData "WD02", textWD02
   SetFieldNewData "WD03", textWD03
   SetFieldNewData "WD04", textWD04
   SetFieldNewData "WD05", textWD05
   SetFieldNewData "WD06", textWD06
   SetFieldNewData "WD07", textWD07
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String

   ' 初始化欄位陣列
   For nIndex = 1 To tf_WD
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "WD" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 1:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textWD02_GotFocus()
   If textWD02.Enabled = True And textWD02.Locked = False Then
      InverseTextBox textWD02
      CloseIme
   End If
End Sub

Private Sub textWD02_KeyPress(KeyAscii As Integer)
   If textWD02.Enabled = True And textWD02.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textWD03_GotFocus()
   If textWD03.Enabled = True And textWD03.Locked = False Then
      InverseTextBox textWD03
      CloseIme
   End If
End Sub

Private Sub textWD03_KeyPress(KeyAscii As Integer)
   If textWD03.Enabled = True And textWD03.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textWD04_GotFocus()
   If textWD04.Enabled = True And textWD04.Locked = False Then
      InverseTextBox textWD04
      CloseIme
   End If
End Sub

Private Sub textWD04_KeyPress(KeyAscii As Integer)
   If textWD04.Enabled = True And textWD04.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textWD05_GotFocus()
   If textWD05.Enabled = True And textWD05.Locked = False Then
      InverseTextBox textWD05
      CloseIme
   End If
End Sub

Private Sub textWD05_KeyPress(KeyAscii As Integer)
   If textWD05.Enabled = True And textWD05.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textWD06_GotFocus()
   If textWD06.Enabled = True And textWD06.Locked = False Then
      InverseTextBox textWD06
      CloseIme
   End If
End Sub

Private Sub textWD06_KeyPress(KeyAscii As Integer)
   If textWD06.Enabled = True And textWD06.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub textWD07_GotFocus()
   If textWD07.Enabled = True And textWD07.Locked = False Then
      InverseTextBox textWD07
      CloseIme
   End If
End Sub

Private Sub textWD07_KeyPress(KeyAscii As Integer)
   If textWD07.Enabled = True And textWD07.Locked = False Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   '                        0       1             2             3             4             5         6
   arrGridHeadText = Array("日期", "北所颱風假", "中所颱風假", "南所颱風假", "高所颱風假", "有補班", "手動下載電子公文")
   arrGridHeadWidth = Array(900, 1000, 1000, 1000, 1000, 900, 1500)
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

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
         If CheckIsTaiwanDate(txt1(Index), False) = False Then
             Cancel = True
             MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
             Exit Sub
         End If
         If txt1(0) <> "" And txt1(1) = "" Then
            txt1(1) = txt1(0)
         End If
         If Index = 1 Then
             If RunNick2(txt1(Index - 1), txt1(Index)) Then
                 Cancel = True
                 Exit Sub
             End If
         End If
      Case Else
   End Select
End Sub
