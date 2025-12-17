VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160022 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作所在地資料"
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
      Left            =   60
      TabIndex        =   10
      Top             =   690
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160022.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Line1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "textSP02_2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label23"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "textSP02"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "textSP01(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "textSP01(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textSP03"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160022.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line5"
      Tab(1).Control(1)=   "Line4"
      Tab(1).Control(2)=   "Label15"
      Tab(1).Control(3)=   "Label16"
      Tab(1).Control(4)=   "txt1(0)"
      Tab(1).Control(5)=   "txt1(1)"
      Tab(1).Control(6)=   "txt1(2)"
      Tab(1).Control(7)=   "txt1(3)"
      Tab(1).Control(8)=   "cmdok"
      Tab(1).Control(9)=   "GRD1"
      Tab(1).ControlCount=   10
      Begin VB.ComboBox textSP03 
         Height          =   260
         ItemData        =   "frm160022.frx":0038
         Left            =   1080
         List            =   "frm160022.frx":003A
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   1050
         Width           =   2205
      End
      Begin VB.TextBox textSP01 
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   1
         Top             =   420
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160022.frx":003C
         Height          =   3615
         Left            =   -75000
         TabIndex        =   11
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
         Height          =   255
         Left            =   -68610
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69990
         MaxLength       =   7
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70980
         MaxLength       =   7
         TabIndex        =   6
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72870
         MaxLength       =   6
         TabIndex        =   5
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   4
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textSP01 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   420
         Width           =   975
      End
      Begin VB.TextBox textSP02 
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   2
         Top             =   750
         Width           =   735
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   450
         TabIndex        =   18
         Top             =   3990
         Width           =   7395
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13044;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label textSP02_2 
         Height          =   225
         Left            =   1860
         TabIndex        =   17
         Top             =   780
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Line Line1 
         X1              =   1830
         X2              =   2520
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期起："
         Height          =   180
         Left            =   -71700
         TabIndex        =   16
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   15
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73230
         X2              =   -72540
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line5 
         X1              =   -70260
         X2              =   -69660
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   390
         TabIndex        =   14
         Top             =   480
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   795
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所在地："
         Height          =   180
         Index           =   17
         Left            =   210
         TabIndex        =   12
         Top             =   1110
         Width           =   720
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
            Picture         =   "frm160022.frx":0051
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":036D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":0689
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":0865
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":0B81
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":0E9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":11B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":14D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":17F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":1B0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160022.frx":1E29
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   9
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
Attribute VB_Name = "frm160022"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Create by Sindy 2020/4/13
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
Dim m_FirstKEY(2) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(2) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(2) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_SP As Integer


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
   rsA.Open "select * from STAFF_WORKPLACE where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SP = rsA.Fields.Count
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

ReDim m_FieldList(tf_SP) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSP01(0).BackColor = &H8000000F
   textSP01(1).BackColor = &H8000000F
   textSP02.BackColor = &H8000000F
   
   MoveFormToCenter Me
   SetTextSP03
   
   InitialField
   InitialData
   RefreshRange
   ShowFirstRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   
   Me.SSTab1.Tab = 0
End Sub

Private Sub SetTextSP03()
   textSP03.Clear
   textSP03.AddItem "01 居家辦公"
   textSP03.AddItem "02 大都會"
   textSP03.AddItem "03 108號4F"
   textSP03.AddItem "04 108號5F"
   textSP03.AddItem "05 北所"
   textSP03.AddItem "06 中所"
   textSP03.AddItem "07 南所"
   textSP03.AddItem "08 高所"
   textSP03.AddItem "09 108號8F(805室)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160022 = Nothing
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
            '2008/12/12 ADD BY SONIA
            textSP01(0).Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSP01(1).Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSP02.Text = GRD1.TextMatrix(tmpMouseRow, 2)
            textSP03.ListIndex = -1
            QueryRecord
            '2008/12/12 END
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("SP04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP04")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SP04"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP05")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP05"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP06")) = False Then
         strTemp = rsSrcTmp.Fields("SP06")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP07")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SP07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP08"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP09")) = False Then
         strTemp = rsSrcTmp.Fields("SP09")
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
Dim strTit As String, strMsg As String
Dim nResponse
   
   TxtValidate = False
   
   If textSP02.Text = "" Then
       MsgBox "員工編號不可以空白！", vbExclamation
       textSP02.SetFocus
       Exit Function
   End If
   If Me.textSP02.Enabled = True Then
      Cancel = False
      textSP02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textSP01(0).Text = "" Then
       MsgBox "起始日期不可以空白！", vbExclamation
       textSP01(0).SetFocus
       Exit Function
   End If
   If textSP01(1).Text = "" Then
       MsgBox "迄止日期不可以空白！", vbExclamation
       textSP01(1).SetFocus
       Exit Function
   End If
   If Me.textSP01(0).Enabled = True Then
      Cancel = False
      textSP01_Validate 0, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSP01(1).Enabled = True Then
      Cancel = False
      textSP01_Validate 1, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   '增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSP02, True, textSP01(0)) = True Then
      textSP02.SetFocus
      Exit Function
   End If
   
   If m_EditMode = 1 Then '新增
      ' 檢查記錄是否已存在
      If IsRecordExist(DBDATE(textSP01(0)), textSP02) = True Then
         strTit = "新增資料"
         strMsg = "該筆記錄已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP01(0).SetFocus
         Exit Function
      End If
      If IsRecordExist(DBDATE(textSP01(1)), textSP02) = True Then
         strTit = "新增資料"
         strMsg = "該筆記錄已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textSP01(1).SetFocus
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_SP - 1
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
   
   For nIndex = 0 To tf_SP - 1
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

' 儲存記錄
Private Function SaveRecord() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Long
   
On Error GoTo ErrHand
   
   SaveRecord = False
   cnnConnection.BeginTrans
   strDate = textSP01(0)
   For i = textSP01(0) To textSP01(1)
      i = strDate
      If ChkWorkDay(DBDATE(strDate)) Then
         strSql = "SELECT * FROM STAFF_WORKPLACE " & _
                  "WHERE SP01=" & DBDATE(strDate) & " and SP02='" & textSP02 & "'  "
         ' 讀取資料庫
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         ' 檢查讀取的資料筆數
         If rsTmp.RecordCount > 0 Then
            '檢查是否需要更新資料
            If rsTmp.Fields("SP03") <> Left(textSP03.Text, 2) Then
               '更新
               strSql = "UPDATE STAFF_WORKPLACE SET SP03='" & Left(textSP03.Text, 2) & "'" & _
                        " WHERE SP01=" & DBDATE(strDate) & " and SP02='" & textSP02 & "'"
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
            End If
         Else
            '新增
            strSql = "INSERT INTO STAFF_WORKPLACE (SP01,SP02,SP03) VALUES (" & _
                     DBDATE(strDate) & "," & CNULL(textSP02) & "," & CNULL(Left(textSP03.Text, 2)) & ")"
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
         rsTmp.Close
      End If
      strDate = ChangeWStringToTString(DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(CStr(i))))))
   Next i
   
   cnnConnection.CommitTrans
   
   ShowCurrRecord DBDATE(textSP01(0)), textSP02
      
   SaveRecord = True
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
'Dim strSP01 As String
'Dim strSP02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
'   strSP01 = m_CurrKEY(0)
'   strSP02 = m_CurrKEY(1)

   strSql = "DELETE FROM STAFF_WORKPLACE" & _
            " WHERE SP01>=" & DBDATE(textSP01(0)) & " and SP01<=" & DBDATE(textSP01(1)) & " and SP02='" & textSP02 & "'"
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

'   If (strSP01 = m_LastKEY(0) And strSP02 = m_LastKEY(1)) Or (strSP01 = m_FirstKEY(0) And strSP02 = m_FirstKEY(1)) Then
      RefreshRange
'   End If
   ShowCurrRecord m_LastKEY(0), m_LastKEY(1)
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSP01 As String
Dim strSP02 As String
   
   QueryRecord = False
   strSP01 = DBDATE(textSP01(0))
   strSP02 = textSP02
   If IsRecordExist(strSP01, strSP02) = True Then
      m_CurrKEY(0) = strSP01
      m_CurrKEY(1) = strSP02
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
'         If CheckDataValid() = True Then
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If SaveRecord = True Then
                RefreshRange
            Else
                Exit Function
            End If
'         Else
'            GoTo EXITSUB
'         End If
      Case 2: '修改
'         If CheckDataValid() = True Then
            'Add By Cheng 2002/05/22
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Function
            UpdateFieldNewData
            If SaveRecord = False Then Exit Function
'         Else
'            GoTo EXITSUB
'         End If
      Case 3: '刪除
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSP01(0) <> "" And textSP02 <> "" Then
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
      Case 1: If Me.Visible = True Then textSP01(0).SetFocus
      Case 2: If Me.Visible = True Then textSP03.SetFocus
      Case 4: If Me.Visible = True Then textSP01(0).SetFocus
   End Select
End Sub
' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = '" & strKEY01 & "'  and SP02='" & strKEY02 & "'  "
                  
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
Private Sub ShowCurrRecord(ByVal strKEY01 As String, ByVal strKEY02 As String)
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01, strKEY02) = True Then
      m_CurrKEY(0) = strKEY01
      m_CurrKEY(1) = strKEY02
   Else
      strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
               "WHERE SP01 = '" & m_CurrKEY(0) & "' and SP02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
               "WHERE SP02 = (SELECT MIN(SP02) FROM STAFF_WORKPLACE where SP01=(select MIN(SP01) from STAFF_WORKPLACE) ) and SP01=(select MIN(SP01) from STAFF_WORKPLACE) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
         If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
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
   
   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SP02 = (SELECT MAX(SP02) FROM STAFF_WORKPLACE " & _
                          "WHERE SP01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SP02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = (SELECT MAX(SP01) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 < '" & m_CurrKEY(0) & "') AND " & _
                  "SP02 = (SELECT MAX(SP02) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 = (SELECT MAX(SP01) FROM STAFF_WORKPLACE " & _
                                          "WHERE SP01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
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
   
   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SP02 = (SELECT MIN(SP02) FROM STAFF_WORKPLACE " & _
                          "WHERE SP01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SP02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = (SELECT MIN(SP01) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 > '" & m_CurrKEY(0) & "') AND " & _
                  "SP02 = (SELECT MIN(SP02) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 = (SELECT MIN(SP01) FROM STAFF_WORKPLACE " & _
                                          "WHERE SP01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SP02")
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
         textSP01(1).Locked = False 'Add By Sindy 2020/4/27
         textSP01(1).BackColor = &H80000005 'Add By Sindy 2020/4/27
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
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = (SELECT MIN(SP01) FROM STAFF_WORKPLACE) AND " & _
                  "SP02 = (SELECT MIN(SP02) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 = (SELECT MIN(SP01) FROM STAFF_WORKPLACE)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("SP02")
   End If
   rsTmp.Close

   strSql = "SELECT SP01,SP02 FROM STAFF_WORKPLACE " & _
            "WHERE SP01 = (SELECT MAX(SP01) FROM STAFF_WORKPLACE) AND " & _
                  "SP02 = (SELECT MAX(SP02) FROM STAFF_WORKPLACE " & _
                           "WHERE SP01 = (SELECT MAX(SP01) FROM STAFF_WORKPLACE)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SP01")) = False Then: m_LastKEY(0) = rsTmp.Fields("SP01")
      If IsNull(rsTmp.Fields("SP02")) = False Then: m_LastKEY(1) = rsTmp.Fields("SP02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM STAFF_WORKPLACE " & _
            "WHERE SP01='" & m_CurrKEY(0) & "' and SP02 = '" & m_CurrKEY(1) & "'   "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("SP01")) = False Then: textSP01(0) = TAIWANDATE(rsTmp.Fields("SP01"))
      If IsNull(rsTmp.Fields("SP01")) = False Then: textSP01(1) = TAIWANDATE(rsTmp.Fields("SP01"))
      If IsNull(rsTmp.Fields("SP02")) = False Then: textSP02 = rsTmp.Fields("SP02")
      If IsNull(rsTmp.Fields("SP03")) = False Then
         textSP03.ListIndex = Val(rsTmp.Fields("SP03")) - 1
      End If
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

      textSP02_2 = GetStaffName(textSP02, True)
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and SP02>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and SP02<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and SP01>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and SP01<='" & DBDATE(txt1(3)) & "' "
   End If
   '抓取資料
   'Modify By Sindy 2023/12/22 部門調整改抓ST93
   strSql = "SELECT sqldateT(SP01),nvl(A0922,'(舊)'||A0902),SP02,st02," & SP03WorkPlace & _
            " From STAFF_WORKPLACE, staff, acc090, acc090NEW" & _
            " where SP02=st01(+) and A0921(+)=st93 and A0901(+)=st03" & strSql & _
            " order by SP01,st93,SP02"
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

'Private Function CheckDataValid() As Boolean
'Dim nResponse As Boolean
'Dim strTmp  As String
'
'   CheckDataValid = False
'
'   nResponse = False
'   textSP01_Validate 0, nResponse
'   If nResponse = True Then GoTo EXITSUB
'   nResponse = False
'   textSP01_Validate 1, nResponse
'   If nResponse = True Then GoTo EXITSUB
'   nResponse = False
'   textSP02_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
'
'   CheckDataValid = True
'EXITSUB:
'End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSP01(0).Locked = bEnable
   'textSP01(1).Locked = bEnable
   textSP02.Locked = bEnable
   If bEnable Then textSP01(0).BackColor = &H8000000F Else textSP01(0).BackColor = &H80000005
   'If bEnable Then textSP01(1).BackColor = &H8000000F Else textSP01(1).BackColor = &H80000005
   If bEnable Then textSP02.BackColor = &H8000000F Else textSP02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSP01(0).Locked = bEnable
   textSP01(1).Locked = bEnable
   textSP02.Locked = bEnable
   If bEnable Then textSP01(0).BackColor = &H8000000F Else textSP01(0).BackColor = &H80000005
   If bEnable Then textSP01(1).BackColor = &H8000000F Else textSP01(1).BackColor = &H80000005
   If bEnable Then textSP02.BackColor = &H8000000F Else textSP02.BackColor = &H80000005
   textSP03.Locked = bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSP01(0) = Empty
   textSP01(1) = Empty
   textSP02_2 = Empty
   textSP02 = Empty
   textSP03.ListIndex = -1

   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SP - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SP01", DBDATE(textSP01(0))
      SetFieldNewData "SP02", textSP02
   End If
   SetFieldNewData "SP03", textSP03
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SP
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SP" & strTmp
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

Private Sub textSP01_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
       InverseTextBox textSP01(Index)
   End If
End Sub

Private Sub textSP01_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSP01_Validate(Index As Integer, Cancel As Boolean)
   If m_EditMode = 1 And textSP01(Index) <> "" Then
       If IsRecordExist(DBDATE(textSP01(Index)), textSP02) = True And textSP01(Index).Enabled = True And textSP01(Index).Locked = False Then
           MsgBox "該員工當天已有資料，請修改！", vbInformation
           Cancel = True
           textSP01(Index) = ""
           Exit Sub
       End If
       If CheckIsTaiwanDate(textSP01(Index), False) = False Then
           Cancel = True
           MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
           Exit Sub
       End If
       If Index = 0 Then
         If textSP01(0) <> "" And textSP01(1) = "" Then
            textSP01(1) = textSP01(0)
         End If
       End If
       '2008/12/23 cancel by sonia 劉經理說此處不必控管工作天
       'If ChkWorkDay(DBDATE(textSP01)) = False Then
       '    Cancel = True
       '    MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
       '    Exit Sub
       'End If
       '2008/12/23 end
   End If
End Sub

Private Sub textSP02_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSP02
       CloseIme
   End If
End Sub

Private Sub textSP02_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSP02_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSP02 <> "" Then
        textSP02_2 = GetStaffName(textSP02, True)
       If IsRecordExist(DBDATE(textSP01(0)), textSP02) = True And textSP02.Enabled = True And textSP02.Locked = False Then
           MsgBox "該員工當天已有資料，請修改！", vbInformation
           Cancel = True
           textSP02 = ""
           Exit Sub
       End If
       If textSP02_2 = "" Then
           MsgBox "員工編號錯誤！查無此員工！", vbInformation
           Cancel = True
           Exit Sub
       End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("日期", "部門", "員工代號", "姓名", "地點")
   arrGridHeadWidth = Array(1200, 1200, 1200, 1200, 1200)
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
           KeyAscii = UpperCase(KeyAscii)
   Case 2, 3
           KeyAscii = Pub_NumAscii(KeyAscii)
   Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      '2011/2/22 modify by sonia
'      Case 1
'              If RunNick(txt1(Index - 1), txt1(Index)) Then
'                  Cancel = True
'                  Exit Sub
'              End If
      Case 0, 1
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
         '2011/2/22 end
      Case 2, 3
              If CheckIsTaiwanDate(txt1(Index), False) = False Then
                  Cancel = True
                  MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
                  Exit Sub
              End If
              If Index = 3 Then
                  If RunNick2(txt1(Index - 1), txt1(Index)) Then
                      Cancel = True
                      Exit Sub
                  End If
              End If
      Case Else
   End Select
End Sub
