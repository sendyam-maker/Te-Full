VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160021 
   BorderStyle     =   1  '單線固定
   Caption         =   "旅遊補助金維護"
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
      TabIndex        =   15
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
      TabPicture(0)   =   "frm160021.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textSTM02(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textSTM03(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textSTM02(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textSTM03(1)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textSTM02(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textSTM03(2)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textSTM02(3)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textSTM03(3)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textSTM02(4)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textSTM03(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textSTM01"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160021.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1"
      Tab(1).Control(1)=   "cmdok"
      Tab(1).Control(2)=   "txt1(1)"
      Tab(1).Control(3)=   "txt1(0)"
      Tab(1).Control(4)=   "Label15"
      Tab(1).Control(5)=   "Line4"
      Tab(1).ControlCount=   6
      Begin VB.TextBox textSTM01 
         Height          =   285
         Left            =   2100
         MaxLength       =   3
         TabIndex        =   0
         Top             =   810
         Width           =   585
      End
      Begin VB.TextBox textSTM03 
         Height          =   270
         Index           =   4
         Left            =   3570
         TabIndex        =   10
         Top             =   2550
         Width           =   1005
      End
      Begin VB.TextBox textSTM02 
         Height          =   270
         Index           =   4
         Left            =   2550
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2550
         Width           =   375
      End
      Begin VB.TextBox textSTM03 
         Height          =   270
         Index           =   3
         Left            =   3570
         TabIndex        =   8
         Top             =   2280
         Width           =   1005
      End
      Begin VB.TextBox textSTM02 
         Height          =   270
         Index           =   3
         Left            =   2550
         MaxLength       =   2
         TabIndex        =   7
         Top             =   2280
         Width           =   375
      End
      Begin VB.TextBox textSTM03 
         Height          =   270
         Index           =   2
         Left            =   3570
         TabIndex        =   6
         Top             =   2010
         Width           =   1005
      End
      Begin VB.TextBox textSTM02 
         Height          =   270
         Index           =   2
         Left            =   2550
         MaxLength       =   2
         TabIndex        =   5
         Top             =   2010
         Width           =   375
      End
      Begin VB.TextBox textSTM03 
         Height          =   270
         Index           =   1
         Left            =   3570
         TabIndex        =   4
         Top             =   1740
         Width           =   1005
      End
      Begin VB.TextBox textSTM02 
         Height          =   270
         Index           =   1
         Left            =   2550
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1740
         Width           =   375
      End
      Begin VB.TextBox textSTM03 
         Height          =   270
         Index           =   0
         Left            =   3570
         TabIndex        =   2
         Top             =   1470
         Width           =   1005
      End
      Begin VB.TextBox textSTM02 
         Height          =   270
         Index           =   0
         Left            =   2550
         MaxLength       =   2
         TabIndex        =   1
         Top             =   1470
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160021.frx":0038
         Height          =   3615
         Left            =   -74970
         TabIndex        =   16
         Top             =   690
         Width           =   8010
         _ExtentX        =   14129
         _ExtentY        =   6368
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "補助年度|年資以上|補助金額"
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
         _Band(0).Cols   =   3
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   285
         Left            =   -68310
         TabIndex        =   13
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -73260
         MaxLength       =   6
         TabIndex        =   12
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   11
         Top             =   360
         Width           =   525
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   540
         TabIndex        =   26
         Top             =   4050
         Width           =   7395
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13044;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         Caption         =   "備註：年資由小至大輸入"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   630
         TabIndex        =   25
         Top             =   3180
         Width           =   3855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   2100
         TabIndex        =   24
         Top             =   2595
         Width           =   2700
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   2100
         TabIndex        =   23
         Top             =   2325
         Width           =   2700
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   2100
         TabIndex        =   22
         Top             =   2055
         Width           =   2700
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   2100
         TabIndex        =   21
         Top             =   1785
         Width           =   2700
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "年資            年以上                        元"
         Height          =   180
         Left            =   2100
         TabIndex        =   20
         Top             =   1515
         Width           =   2700
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "補助金額："
         Height          =   180
         Left            =   1200
         TabIndex        =   19
         Top             =   1500
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補助年度：              年 (108)"
         Height          =   180
         Index           =   1
         Left            =   1200
         TabIndex        =   18
         Top             =   840
         Width           =   2145
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "年度："
         Height          =   180
         Left            =   -74450
         TabIndex        =   17
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line4 
         X1              =   -73650
         X2              =   -72960
         Y1              =   480
         Y2              =   480
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
            Picture         =   "frm160021.frx":004D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":0369
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":0685
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":0861
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":0B7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":0E99
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":11B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":14D1
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":17ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":1B09
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160021.frx":1E25
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   14
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
Attribute VB_Name = "frm160021"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Create by Sindy 2019/7/31
Option Explicit

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
Dim oTxt2 As Object, oTxt3 As Object
Dim ii As Integer


Private Sub cmdok_Click()
   If txt1(0) & txt1(1) <> "" Then
       If RunNick(txt1(0), txt1(1)) Then
           txt1(0).SetFocus
           Exit Sub
       End If
       GetData
   Else
       MsgBox "查詢條件不可以空白！", vbExclamation, "操作錯誤！"
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
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   MoveFormToCenter Me
   
   InitialData
   RefreshRange
   ShowLastRecord
   UpdateToolbarState
   SetCtrlReadOnly True
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160021 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim nCol As Long, nRow As Long
   
   getGrdColRow GRD1, X, Y, nCol, nRow
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
         textSTM01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
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

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String

   If IsNull(rsSrcTmp.Fields("STM04")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STM04")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("STM04"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STM05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STM05")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("STM05"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("STM06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("STM06")) = False Then
         strTemp = rsSrcTmp.Fields("STM06")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim jj As Integer

   TxtValidate = False
   
   If textSTM02(0).Text = "" Then
       MsgBox "年資以上不可以空白！", vbExclamation
       textSTM02(0).SetFocus
       Exit Function
   End If
   If textSTM03(0).Text = "" Then
       MsgBox "補助金額不可以空白！", vbExclamation
       textSTM03(0).SetFocus
       Exit Function
   End If
   
   For ii = 4 To 0 Step -1
      For jj = ii - 1 To 0 Step -1
         If Len(textSTM02(ii)) > 0 And Len(textSTM02(jj)) = 0 Then
            MsgBox "資料請按照順序輸入！", vbExclamation
            Exit Function
         End If
         If Len(textSTM03(ii)) > 0 And Len(textSTM03(jj)) = 0 Then
            MsgBox "資料請按照順序輸入！", vbExclamation
            Exit Function
         End If
      Next jj
   Next ii
   
   For ii = 0 To 4
      If Len(textSTM02(ii)) > 0 And Val(textSTM03(ii)) = 0 Then
         MsgBox "請輸入補助金額！", vbExclamation
         textSTM03(ii).SetFocus
         Exit Function
      End If
      If Len(textSTM02(ii)) = 0 And Val(textSTM03(ii)) > 0 Then
         MsgBox "年資以上不可以空白！", vbExclamation
         textSTM02(ii).SetFocus
         Exit Function
      End If
   Next ii
   
   For ii = 4 To 1 Step -1
      If Len(textSTM02(ii)) > 0 And Val(textSTM02(ii)) <= Val(textSTM02(ii - 1)) Then
         MsgBox "等級 " & ii + 1 & " 的年資應該比等級 " & ii & " 的年資大！", vbExclamation
         textSTM02(ii).SetFocus
         Exit Function
      End If
      If Len(textSTM03(ii)) > 0 And Val(textSTM03(ii)) <= Val(textSTM03(ii - 1)) Then
         MsgBox "等級 " & ii + 1 & " 的補助金應該比等級 " & ii & " 的補助金大！", vbExclamation
         textSTM03(ii).SetFocus
         Exit Function
      End If
   Next ii
   
   TxtValidate = True
End Function

' 新增記錄
Private Function AddRecord() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strSTM01 As String

   AddRecord = False

   strSTM01 = Val(textSTM01) + 1911
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSTM01) = True Then
      strTit = "新增資料"
      strMsg = "該年度資料已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   For ii = 0 To 4
      If Len(Trim(textSTM02(ii))) > 0 Then
         strSql = "INSERT INTO staff_TravelMoney(STM01,STM02,STM03) VALUES(" & strSTM01 & "," & Val(textSTM02(ii)) & "," & Val(textSTM03(ii)) & ")"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   Next ii
   
   If strSTM01 < m_FirstKEY(0) Or strSTM01 > m_LastKEY(0) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strSTM01
   AddRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 新增失敗！" & vbCrLf & Err.Description
End Function

' 修改記錄
Private Function ModRecord() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strSTM01 As String

   ModRecord = False

   strSTM01 = Val(textSTM01) + 1911
   
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   
   For ii = 0 To 4
      If Len(Trim(textSTM02(ii))) > 0 Then
         If CDbl(textSTM02(ii).Text) <> CDbl(textSTM02(ii).Tag) Or _
            CDbl(textSTM03(ii).Text) <> CDbl(textSTM03(ii).Tag) Then
            strSql = "UPDATE staff_TravelMoney SET STM02=" & Val(textSTM02(ii).Text) & ",STM03=" & Val(textSTM03(ii).Text) & _
                     " WHERE STM01='" & strSTM01 & "' and STM02=" & Val(textSTM02(ii).Tag)
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
      End If
   Next ii
   
   cnnConnection.CommitTrans
   
   ShowCurrRecord strSTM01
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox " 修改失敗！" & vbCrLf & Err.Description
End Function

' 刪除記錄
Private Function DelRecord(strSTM01 As String) As Boolean
Dim strSql As String

   DelRecord = False

On Error GoTo ErrHand

   cnnConnection.BeginTrans
   
   strSTM01 = Val(textSTM01) + 1911

   strSql = "DELETE FROM staff_TravelMoney " & _
            "WHERE STM01 = " & strSTM01
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If strSTM01 = m_LastKEY(0) Or strSTM01 = m_FirstKEY(0) Then
      RefreshRange
   End If
   ShowCurrRecord strSTM01
   DelRecord = True
   cnnConnection.CommitTrans

   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSTM01 As String
   
   QueryRecord = False
   strSTM01 = Val(textSTM01) + 1911
   If IsRecordExist(strSTM01) = True Then
      m_CurrKEY(0) = strSTM01
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
            If ModRecord = False Then Exit Function
         Else
            GoTo EXITSUB
         End If
      Case 3: '刪除
         If DelRecord(m_CurrKEY(0)) = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSTM01 <> "" Then
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
      Case 1: If Me.Visible = True Then textSTM01.SetFocus
      Case 2: If Me.Visible = True Then textSTM02(0).SetFocus
      Case 4: If Me.Visible = True Then textSTM01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String

   IsRecordExist = False
   
   strSql = "SELECT * FROM staff_TravelMoney " & _
            "WHERE STM01 = " & strKEY01
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
      strSql = "SELECT STM01 FROM staff_TravelMoney " & _
               "WHERE STM01 = " & m_CurrKEY(0)
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("STM01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("STM01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close

      strSql = "SELECT max(STM01) FROM staff_TravelMoney"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
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

   strSql = "SELECT max(STM01) FROM staff_TravelMoney " & _
            "WHERE STM01 < " & m_CurrKEY(0)
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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

   strSql = "SELECT min(STM01) FROM staff_TravelMoney " & _
            "WHERE STM01 > " & m_CurrKEY(0)
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_CurrKEY(0) = rsTmp.Fields(0)
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
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
         m_EditMode = 1
         ClearField
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         UpdateToolbarState
         SetInputEntry
      ' 修改
      Case vbKeyF3:
         '檢查此年度是否已有輸入補助金資料
         strExc(0) = "SELECT STF03 FROM staff_TravelFee" & _
                     " WHERE STF03='" & Val(textSTM01) + 1911 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox textSTM01 & "年已有輸入補助金資料，不可再調整！", vbExclamation
            Exit Sub
         End If
      
         m_EditMode = 2
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         SetCtrlReadOnly False
         SetKeyReadOnly True
         UpdateToolbarState
         SetInputEntry
      ' 刪除
      Case vbKeyF5:
         '檢查此年度是否已有輸入補助金資料
         strExc(0) = "SELECT STF03 FROM staff_TravelFee" & _
                     " WHERE STF03='" & Val(textSTM01) + 1911 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox textSTM01 & "年已有輸入補助金資料，不可刪除！", vbExclamation
            Exit Sub
         End If
         
         strTit = "詢問"
         strMsg = "是否要刪除此年度資料?"
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

   strSql = "SELECT min(STM01) FROM staff_TravelMoney "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_FirstKEY(0) = rsTmp.Fields(0)
   End If
   rsTmp.Close

   strSql = "SELECT max(STM01) FROM staff_TravelMoney "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then: m_LastKEY(0) = rsTmp.Fields(0)
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If m_CurrKEY(0) <> "" Then
      strSql = "SELECT * FROM staff_TravelMoney " & _
               "WHERE STM01 = " & m_CurrKEY(0) & " order by STM02 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         ClearField
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("STM01")) = False Then: textSTM01 = rsTmp.Fields("STM01") - 1911
         ' 更新CUID
         UpdateCUID rsTmp
         ii = 0
         Do While Not rsTmp.EOF
            If IsNull(rsTmp.Fields("STM02")) = False Then: textSTM02(ii).Text = rsTmp.Fields("STM02"): textSTM02(ii).Tag = rsTmp.Fields("STM02")
            If IsNull(rsTmp.Fields("STM03")) = False Then: textSTM03(ii).Text = rsTmp.Fields("STM03"): textSTM03(ii).Tag = rsTmp.Fields("STM03")
            ii = ii + 1
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
   
   strSql = ""
   If txt1(0) <> "" Then
      If strSql <> "" Then strSql = strSql & " and "
      strSql = strSql & " STM01>='" & Val(txt1(0)) + 1911 & "' "
   End If
   If txt1(1) <> "" Then
      If strSql <> "" Then strSql = strSql & " and "
      strSql = strSql & " STM01<='" & Val(txt1(1)) + 1911 & "' "
   End If
   '抓取資料
   strSql = "SELECT STM01-1911,STM02,to_char(STM03,'999,999') FROM staff_TravelMoney where " & strSql & _
           " order by STM01,STM02 "
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
   textSTM01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   For ii = 0 To 4
      textSTM02_Validate ii, nResponse
      If nResponse = True Then GoTo EXITSUB
   Next ii
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSTM01.Locked = bEnable
   If bEnable Then textSTM01.BackColor = &H8000000F Else textSTM01.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSTM01.Locked = bEnable
   If bEnable Then textSTM01.BackColor = &H8000000F Else textSTM01.BackColor = &H80000005
   For Each oTxt2 In textSTM02
      oTxt2.Locked = bEnable
      If bEnable Then oTxt2.BackColor = &H8000000F Else oTxt2.BackColor = &H80000005
   Next
   For Each oTxt3 In textSTM03
      oTxt3.Locked = bEnable
      If bEnable Then oTxt3.BackColor = &H8000000F Else oTxt3.BackColor = &H80000005
   Next
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSTM01 = Empty
   For Each oTxt2 In textSTM02
      oTxt2.Text = Empty
      oTxt2.Tag = Empty
   Next
   For Each oTxt3 In textSTM03
      oTxt3.Text = Empty
      oTxt3.Tag = Empty
   Next
   Label23 = Empty
   SetGrd
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSTM01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSTM01
      CloseIme
   End If
End Sub

Private Sub textSTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTM01_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset

   If m_EditMode = 1 And textSTM01 <> "" Then
      If CheckIsTaiwanDate(textSTM01 & "0101", False) = False Then
         Cancel = True
         MsgBox "請輸入民國度！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      ' 檢查記錄是否已存在
      If IsRecordExist(Val(textSTM01) + 1911) = True Then
         Cancel = True
         MsgBox "該年度資料已存在！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
   End If
End Sub

Private Sub textSTM02_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
      InverseTextBox textSTM02(Index)
      CloseIme
   End If
End Sub

Private Sub textSTM02_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSTM02_Validate(Index As Integer, Cancel As Boolean)
Dim ii As Integer
Dim intText As Integer
   
   If Len(Trim(textSTM02(Index).Text)) = 0 Then
      Exit Sub
   Else
      intText = Val(textSTM02(Index).Text)
   End If
   ii = 0
   For Each oTxt2 In textSTM02
      If ii <> Index And Len(Trim(oTxt2.Text)) > 0 Then
         If Val(oTxt2.Text) = intText Then
            MsgBox "年資不可重覆輸入", vbExclamation
            textSTM02(Index).SetFocus
            Cancel = True
            Exit For
         End If
      End If
      ii = ii + 1
   Next
End Sub

Private Sub textSTM03_GotFocus(Index As Integer)
   If m_EditMode <> 0 Then
      InverseTextBox textSTM03(Index)
      CloseIme
   End If
End Sub

Private Sub textSTM03_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer

   arrGridHeadText = Array("補助年度", "年資以上", "補助金額")
   arrGridHeadWidth = Array(1500, 1500, 1500)
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
      Case Else
   End Select
End Sub
