VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160008 
   BorderStyle     =   1  '單線固定
   Caption         =   "獎懲資料"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8190
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8190
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
            Picture         =   "frm160008.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160008.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   615
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   8190
      _ExtentX        =   14446
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
      BorderStyle     =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4380
      Left            =   30
      TabIndex        =   11
      Top             =   660
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160008.frx":20F4
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(17)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "textSR04"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textSR01_2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textSR02"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textSR03"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textSR01"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textSR11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160008.frx":2110
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label15"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Line4"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Line5"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txt1(3)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "txt1(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txt1(1)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "txt1(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.TextBox textSR11 
         Height          =   270
         Left            =   -73980
         MaxLength       =   1
         TabIndex        =   3
         Top             =   1410
         Width           =   300
      End
      Begin VB.TextBox textSR01 
         Height          =   270
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   0
         Top             =   510
         Width           =   735
      End
      Begin VB.ComboBox textSR03 
         Height          =   300
         ItemData        =   "frm160008.frx":212C
         Left            =   -73980
         List            =   "frm160008.frx":2151
         TabIndex        =   2
         Top             =   1110
         Width           =   1695
      End
      Begin VB.TextBox textSR02 
         Height          =   270
         Left            =   -73980
         MaxLength       =   7
         TabIndex        =   1
         Top             =   810
         Width           =   945
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   5
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   2070
         MaxLength       =   6
         TabIndex        =   6
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   3960
         MaxLength       =   7
         TabIndex        =   7
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   4950
         MaxLength       =   7
         TabIndex        =   8
         Top             =   390
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   6330
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160008.frx":21BE
         Height          =   3615
         Left            =   30
         TabIndex        =   12
         Top             =   750
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6376
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
      Begin MSForms.Label textSR01_2 
         Height          =   225
         Left            =   -73170
         TabIndex        =   20
         Top             =   540
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   -74850
         TabIndex        =   21
         Top             =   4020
         Width           =   7785
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13732;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSR04 
         Height          =   945
         Left            =   -73980
         TabIndex        =   4
         Top             =   1710
         Width           =   6525
         VariousPropertyBits=   -1466939365
         MaxLength       =   200
         ScrollBars      =   3
         Size            =   "11509;1667"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "次數："
         Height          =   180
         Index           =   1
         Left            =   -74520
         TabIndex        =   19
         Top             =   1455
         Width           =   540
      End
      Begin VB.Line Line5 
         X1              =   4680
         X2              =   5280
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line4 
         X1              =   1710
         X2              =   2400
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   90
         TabIndex        =   18
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   3340
         TabIndex        =   17
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Left            =   -74520
         TabIndex        =   16
         Top             =   1755
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Index           =   17
         Left            =   -74520
         TabIndex        =   15
         Top             =   855
         Width           =   540
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "類別："
         Height          =   180
         Left            =   -74520
         TabIndex        =   14
         Top             =   1155
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   13
         Top             =   555
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm160008"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by Sindy 2009/01/16 copy from frm160005
'2009/11/30 MODIFY BY SONIA 加SR11次數
Option Explicit

Dim RcMain As New ADODB.Recordset, RsAdo As New ADODB.Recordset
' 變數宣告區
Dim m_EditMode As Integer
Dim m_SubMode As Integer
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
Dim m_FirstKEY(3) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(3) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(3) As String
Dim rsA As New ADODB.Recordset
Dim tf_SR As Integer


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) & txt1(2) & txt1(3) <> "" Then
       If RunNick(txt1(0), txt1(1)) Then
           txt1(0).SetFocus
           Exit Sub
       End If
       If RunNick2(txt1(2), txt1(3)) Then
           txt1(2).SetFocus
           Exit Sub
       End If
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
   End If
End Sub

Private Sub Form_Initialize()
   Set rsA = New ADODB.Recordset
   If rsA.State = 1 Then rsA.Close
   rsA.CursorLocation = adUseClient
   rsA.Open "select * from staff_reward where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SR = rsA.Fields.Count
   SetGrd
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
'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
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
   ReDim m_FieldList(tf_SR) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSR01.BackColor = &H8000000F
   textSR02.BackColor = &H8000000F
   
   MoveFormToCenter Me
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160008 = Nothing
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
            textSR01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSR02.Text = Trim(ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2)))
            textSR03.Text = GRD1.TextMatrix(tmpMouseRow, 3) 'Add By Sindy 2019/3/4
            QueryRecord
            GRD1.Visible = True
       End If
   End If
End Sub

'Add By Sindy 2019/8/27
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("SR05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SR05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SR06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SR06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SR07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR07")) = False Then
         strTemp = rsSrcTmp.Fields("SR07")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SR08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SR08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SR09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SR09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SR10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SR10")) = False Then
         strTemp = rsSrcTmp.Fields("SR10")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.textSR01.Enabled = True Then
      Cancel = False
      textSR01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSR01.Text = "" Then
       MsgBox "員工編號不可以空白！", vbExclamation
       textSR01.SetFocus
       Exit Function
   End If
   If Me.textSR02.Enabled = True Then
      Cancel = False
      textSR02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSR02.Text = "" Then
       MsgBox "日期不可以空白！", vbExclamation
       textSR02.SetFocus
       Exit Function
   End If
   
   'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSR01, True, textSR02) = True Then
      textSR01.SetFocus
      Exit Function
   End If
   
   If Me.textSR03.Enabled = True Then
      Cancel = False
      textSR03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSR03.Text = "" Then
       MsgBox "類別不可以空白！", vbExclamation
       textSR03.SetFocus
       Exit Function
   End If
   If Me.textSR04.Enabled = True Then
      Cancel = False
      textSR04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/9/1 END
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_SR - 1
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
   
   For nIndex = 0 To tf_SR - 1
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

' 新增記錄
Private Function AddRecord() As Boolean
Dim strSql As String
Dim strTmp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim nIndex As Integer
Dim bDifference As Boolean
Dim bFirst As Boolean
Dim strSR01 As String
Dim strSR02 As String
Dim strSR03 As String
   
   AddRecord = False
   
   strSR01 = textSR01
   strSR02 = DBDATE(textSR02)
   strSR03 = Left(Trim(textSR03), 2)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSR01, strSR02, strSR03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_reward ("
   For nIndex = 0 To tf_SR - 1
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
   For nIndex = 0 To tf_SR - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
         If m_FieldList(nIndex).fiType = 0 Then
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
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If ((strSR01 & strSR02 & strSR03) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2))) Or ((strSR01 & strSR02 & strSR03) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strSR01, DBDATE(strSR02), strSR03
   
   AddRecord = True
   Exit Function
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
   
End Function

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
Dim strSR01 As String
Dim strSR02 As String
Dim strSR03 As String
   
   ModRecord = False
   
   strSR01 = m_CurrKEY(0)
   strSR02 = m_CurrKEY(1)
   strSR03 = m_CurrKEY(2)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_reward SET "
   
   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SR - 1
      strTmp = Empty
      'If nIndex < 4 Or nIndex > 9 Then
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
                  "WHERE SR01 = '" & strSR01 & "' and SR02='" & strSR02 & "' and SR03='" & strSR03 & "' ; end; "
   
On Error GoTo ErrHand
      cnnConnection.BeginTrans
      If bDifference = True Then
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
      cnnConnection.CommitTrans
      
      ShowCurrRecord strSR01, DBDATE(strSR02), strSR03
      
    ModRecord = True
   Exit Function
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
   
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strSR01 As String
Dim strSR02 As String
Dim strSR03 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSR01 = m_CurrKEY(0)
   strSR02 = m_CurrKEY(1)
   strSR03 = m_CurrKEY(2)
   
   strSql = "DELETE FROM staff_reward " & _
            "WHERE SR01 = '" & strSR01 & "' and SR02='" & strSR02 & "' and SR03='" & strSR03 & "'"
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (strSR01 = m_LastKEY(0) And strSR02 = m_LastKEY(1) And strSR03 = m_LastKEY(2)) Or (strSR01 = m_FirstKEY(0) And strSR02 = m_FirstKEY(1) And strSR03 = m_FirstKEY(2)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strSR01, DBDATE(strSR02), strSR03
   
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
   
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSR01 As String
Dim strSR02 As String
Dim strSR03 As String
   
   QueryRecord = False
   strSR01 = textSR01
   strSR02 = DBDATE(Trim(textSR02))
   strSR03 = Left(Trim(textSR03), 2)
   
   If IsRecordExist(strSR01, strSR02, strSR03) = True Then
      m_CurrKEY(0) = strSR01
      m_CurrKEY(1) = strSR02
      m_CurrKEY(2) = strSR03
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
      Case 1: '新增
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If AddRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
         Else
            GoTo EXITSUB
         End If
      Case 2: '修改
         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSR01 <> "" And textSR02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            If textSR01 = "" Or textSR02 = "" Then
               MsgBox "須輸入員工代號及日期才可進行查詢動作！", vbInformation
            End If
            
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
      '新增
      Case 1: If Me.Visible = True Then textSR01.SetFocus
      '修改
      Case 2: If Me.Visible = True Then textSR11.SetFocus
      '查詢
      Case 4: If Me.Visible = True Then textSR01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   
   '比較員編和日期
   strSql = "SELECT * FROM staff_reward " & _
            "WHERE SR01 = '" & strKEY01 & "' and SR02='" & strKEY02 & "' and SR03='" & strKEY03 & "'"
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02, strKEY03) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
      m_CurrKEY(2) = strKEY03
   Else
      strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
               "WHERE SR01 = '" & m_CurrKEY(0) & "' and SR02='" & m_CurrKEY(1) & "' and SR03='" & m_CurrKEY(2) & "' "
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
         If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
         If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
               "WHERE SR02 = (SELECT MIN(SR02) FROM staff_reward " & _
                             "where SR01=(select min(SR01) from staff_reward) ) " & _
                 "and SR01=(select min(SR01) from staff_reward) " & _
               "order by SR01,SR02,SR03 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
         If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
         If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
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
   m_CurrKEY(2) = m_FirstKEY(2)
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = '" & m_CurrKEY(0) & "' " & _
              "AND SR02 = (SELECT MAX(SR02) FROM staff_reward " & _
                          "WHERE SR01 = '" & m_CurrKEY(0) & "' " & _
                            "AND SR02 < '" & m_CurrKEY(1) & "') " & _
            "order by SR01,SR02,SR03 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = (SELECT MAX(SR01) FROM staff_reward " & _
                          "WHERE SR01 < '" & m_CurrKEY(0) & "') " & _
                            "AND SR02 = (SELECT MAX(SR02) FROM staff_reward " & _
                                        "WHERE SR01 = (SELECT MAX(SR01) FROM staff_reward " & _
                                                      "WHERE SR01 < '" & m_CurrKEY(0) & "')) " & _
            "order by SR01,SR02,SR03 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
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
   
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = '" & m_CurrKEY(0) & "' " & _
              "AND SR02 = (SELECT MIN(SR02) FROM staff_reward " & _
                          "WHERE SR01 = '" & m_CurrKEY(0) & "' " & _
                            "AND SR02 > '" & m_CurrKEY(1) & "') " & _
            "order by SR01,SR02,SR03 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = (SELECT MIN(SR01) FROM staff_reward " & _
                          "WHERE SR01 > '" & m_CurrKEY(0) & "') " & _
                            "AND SR02 = (SELECT MIN(SR02) FROM staff_reward " & _
                                        "WHERE SR01 = (SELECT MIN(SR01) FROM staff_reward " & _
                                                      "WHERE SR01 > '" & m_CurrKEY(0) & "')) " & _
            "order by SR01,SR02,SR03 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("SR03")
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
   m_CurrKEY(2) = m_LastKEY(2)
   UpdateCtrlData
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
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
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
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         If nResponse = vbYes Then
            m_EditMode = 3
            If OnWork = True Then
                UpdateToolbarState
            Else
                Exit Sub
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
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = (SELECT MIN(SR01) FROM staff_reward) " & _
              "AND SR02 = (SELECT MIN(SR02) FROM staff_reward " & _
                          "WHERE SR01 = (SELECT MIN(SR01) FROM staff_reward)) " & _
            "order by SR01,SR02,SR03 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("SR03")
   End If
   rsTmp.Close
   
   strSql = "SELECT SR01,SR02,SR03 FROM staff_reward " & _
            "WHERE SR01 = (SELECT MAX(SR01) FROM staff_reward) " & _
              "AND SR02 = (SELECT MAX(SR02) FROM staff_reward " & _
                           "WHERE SR01 = (SELECT MAX(SR01) FROM staff_reward)) " & _
            "order by SR01,SR02,SR03 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SR01")) = False Then: m_LastKEY(0) = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: m_LastKEY(1) = rsTmp.Fields("SR02")
      If IsNull(rsTmp.Fields("SR03")) = False Then: m_LastKEY(2) = rsTmp.Fields("SR03")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM staff_reward " & _
            "WHERE SR01='" & m_CurrKEY(0) & "' and SR02 = '" & m_CurrKEY(1) & "' and SR03 = '" & m_CurrKEY(2) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("SR01")) = False Then: textSR01 = rsTmp.Fields("SR01")
      If IsNull(rsTmp.Fields("SR02")) = False Then: textSR02 = TAIWANDATE(rsTmp.Fields("SR02"))
      If IsNull(rsTmp.Fields("SR03")) = False Then: textSR03 = rsTmp.Fields("SR03")
      If IsNull(rsTmp.Fields("SR04")) = False Then: textSR04 = rsTmp.Fields("SR04")
      If IsNull(rsTmp.Fields("SR11")) = False Then: textSR11 = rsTmp.Fields("SR11")   '2009/11/30 ADD BY SONIA
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
       textSR03_Validate False
       textSR01_2 = GetStaffName(textSR01, True)
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
   
   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and SR01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and SR01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and SR02>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and SR02<='" & DBDATE(txt1(3)) & "' "
   End If
   '抓取資料
   '2009/11/30 MODIFY BY SONIA 加SR11次數
   strSql = "SELECT SR01,s1.st02,sqldateT(SR02),ac02||' '||ac03,SR11,SR04 FROM staff_reward,staff s1,allcode where SR01=s1.st01(+) and '08'=ac01(+) and SR03=ac02(+) " & strSql & _
                  "order by SR02,SR01 "
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
   
   nResponse = False
   textSR01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSR02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSR03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSR04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
   
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSR01.Locked = bEnable
   textSR02.Locked = bEnable
   textSR03.Locked = bEnable
   If bEnable Then textSR01.BackColor = &H8000000F Else textSR01.BackColor = &H80000005
   If bEnable Then textSR02.BackColor = &H8000000F Else textSR02.BackColor = &H80000005
   If bEnable Then textSR03.BackColor = &H8000000F Else textSR03.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSR01.Locked = bEnable
   textSR02.Locked = bEnable
   textSR03.Locked = bEnable
   If bEnable Then textSR01.BackColor = &H8000000F Else textSR01.BackColor = &H80000005
   If bEnable Then textSR02.BackColor = &H8000000F Else textSR02.BackColor = &H80000005
   If bEnable Then textSR03.BackColor = &H8000000F Else textSR03.BackColor = &H80000005
   textSR04.Locked = bEnable
   textSR11.Locked = bEnable   '2009/11/30 ADD BY SONIA
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSR01 = Empty
   textSR01_2 = Empty
   textSR02 = Empty
   textSR03 = Empty
   textSR04 = Empty
   textSR11 = 1   '2009/11/30 ADD BY SONIA
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SR - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SR01", textSR01
      SetFieldNewData "SR02", DBDATE(textSR02)
   End If
   If textSR03.Text <> "" Then
        MyArr = Split(textSR03, " ")
        SetFieldNewData "SR03", MyArr(0)
   Else
        SetFieldNewData "SR03", Empty
   End If
   SetFieldNewData "SR04", textSR04
   '2009/11/30 ADD BY SONIA
   If textSR11 <> "" Then
      SetFieldNewData "SR11", textSR11
   Else
      SetFieldNewData "SR11", 1
   End If
   '2009/11/30 END
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SR
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SR" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 11:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
Dim MyRs As New ADODB.Recordset
   
   textSR03.Clear
   
   Set MyRs = New ADODB.Recordset
   If MyRs.State = 1 Then MyRs.Close
   strSql = "select ac02||' '||ac03 from allcode where ac01='08' order by ac02"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
       While Not MyRs.EOF
           textSR03.AddItem "" & MyRs.Fields(0).Value
           MyRs.MoveNext
       Wend
   End If
   SetGrd
End Sub

Private Sub textSR01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSR01
End If
End Sub

Private Sub textSR01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSR01_Validate(Cancel As Boolean)
   If textSR01.Text = "" Then textSR01_2 = ""
   
   If m_EditMode <> 0 And textSR01 <> "" Then
       textSR01_2 = GetStaffName(textSR01, True)
       ' 檢查員工編號規則
       If ChkStaffID(textSR01) Then
          Call textSR01_GotFocus
          Cancel = True
          Exit Sub
       End If
       If textSR01_2 = "" Then
           MsgBox "員工編號錯誤！查無此員工！", vbInformation
           Call textSR01_GotFocus
           Cancel = True
           Exit Sub
       End If
   End If
   
   If m_EditMode = 1 And textSR01 <> "" Then
       If textSR02 <> "" And textSR03 <> "" Then
         If IsRecordExist(textSR01, DBDATE(textSR02), Left(Trim(textSR03), 2)) = True And textSR01.Enabled = True And textSR01.Locked = False Then
             MsgBox "該員工當天已有資料，請修改！", vbInformation
             Call textSR01_GotFocus
             Cancel = True
             Exit Sub
         End If
       End If
   End If
End Sub

Private Sub textSR02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSR02
    CloseIme
End If
End Sub

Private Sub textSR02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSR02_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSR02 <> "" Then
       If CheckIsTaiwanDate(textSR02, False) = False Then
           Call textSR02_GotFocus
           Cancel = True
           MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
           Exit Sub
       End If
   End If
   If m_EditMode = 1 And textSR02 <> "" Then
       If textSR01 <> "" And textSR02 <> "" And textSR03 <> "" Then
         If IsRecordExist(textSR01, DBDATE(textSR02), Left(Trim(textSR03), 2)) = True And textSR01.Enabled = True And textSR01.Locked = False Then
             MsgBox "該員工當天已有資料，請修改！", vbInformation
             Call textSR02_GotFocus
             Cancel = True
             Exit Sub
         End If
       End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   '2009/11/30 MODIFY BY SONIA 加SR11次數
   arrGridHeadText = Array("員工編號", "姓名", "日期", "類別", "次數", "備註")
   arrGridHeadWidth = Array(800, 1200, 1200, 800, 500, 2000)
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

Private Sub textSR03_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSR03
End If
End Sub

Private Sub textSR03_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSR03_Validate(Cancel As Boolean)
Dim MyRs As New ADODB.Recordset
Dim MyArr As Variant
   
   If textSR03.Text <> "" Then
       MyArr = Split(textSR03, " ")
       Set MyRs = New ADODB.Recordset
       If MyRs.State = 1 Then MyRs.Close
       strSql = "select ac02||' '||ac03 from allcode where ac01='08' and ac02='" & MyArr(0) & "' order by ac02"
       MyRs.CursorLocation = adUseClient
       MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If MyRs.RecordCount <> 0 Then
            textSR03.Text = "" & MyRs.Fields(0).Value
       Else
           If m_EditMode <> 0 Then
               Call textSR03_GotFocus
               MsgBox "獎懲代號輸入錯誤!!!", vbExclamation + vbOKOnly
               Cancel = True
               Exit Sub
           End If
       End If
   End If
End Sub

Private Sub textSR04_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSR04
End If
End Sub

Private Sub textSR04_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSR04_Validate(Cancel As Boolean)
   '若不是修改狀態，不需檢查
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   If textSR04.Text = "" Then Exit Sub
   If Not CheckLengthIsOK(textSR04, textSR04.MaxLength) Then
      Cancel = True
   End If
End Sub

Private Sub textSR11_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And (Not IsNumeric(Chr(KeyAscii)) Or KeyAscii < 49) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = UpperCase(KeyAscii)
      Case 2, 3
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case 2, 3
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case Else
   End Select
End Sub
