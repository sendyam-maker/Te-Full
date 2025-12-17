VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160019 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工健檢報告資料"
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
      TabIndex        =   9
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
      TabPicture(0)   =   "frm160019.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "textSH01_2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "textSH04"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textSH03"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textSH02"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textSH01"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textOld"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textST68"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textST68_2"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Text4"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160019.frx":001C
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
      Begin VB.TextBox Text4 
         Appearance      =   0  '平面
         BackColor       =   &H8000000F&
         BorderStyle     =   0  '沒有框線
         ForeColor       =   &H000000C0&
         Height          =   630
         Left            =   2370
         Locked          =   -1  'True
         MaxLength       =   6
         MultiLine       =   -1  'True
         TabIndex        =   23
         Text            =   "frm160019.frx":0038
         Top             =   3270
         Width           =   2895
      End
      Begin VB.TextBox textST68_2 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1590
         Width           =   615
      End
      Begin VB.TextBox textST68 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   1860
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1590
         Width           =   615
      End
      Begin VB.TextBox textOld 
         Appearance      =   0  '平面
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Left            =   3510
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   870
         Width           =   615
      End
      Begin VB.TextBox textSH01 
         Height          =   270
         Left            =   1500
         MaxLength       =   6
         TabIndex        =   0
         Top             =   570
         Width           =   735
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160019.frx":0087
         Height          =   3615
         Left            =   -74970
         TabIndex        =   10
         Top             =   690
         Width           =   8010
         _ExtentX        =   14129
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
         Height          =   285
         Left            =   -68310
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -69630
         MaxLength       =   7
         TabIndex        =   6
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -70620
         MaxLength       =   7
         TabIndex        =   5
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72870
         MaxLength       =   6
         TabIndex        =   4
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73920
         MaxLength       =   6
         TabIndex        =   3
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textSH02 
         Height          =   270
         Left            =   1500
         MaxLength       =   7
         TabIndex        =   1
         Top             =   870
         Width           =   975
      End
      Begin VB.TextBox textSH03 
         Height          =   270
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   2
         Top             =   1200
         Width           =   345
      End
      Begin MSForms.TextBox textSH04 
         Height          =   1320
         Left            =   1500
         TabIndex        =   28
         Top             =   1890
         Width           =   5895
         VariousPropertyBits=   -1466939365
         MaxLength       =   200
         ScrollBars      =   3
         Size            =   "10398;2328"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label23 
         Height          =   225
         Left            =   510
         TabIndex        =   27
         Top             =   4020
         Width           =   7335
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "12938;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label textSH01_2 
         Height          =   225
         Left            =   2310
         TabIndex        =   26
         Top             =   600
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "請自行判斷是否補助費用"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   2550
         TabIndex        =   25
         Top             =   1230
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "備　　註："
         Height          =   180
         Left            =   600
         TabIndex        =   24
         Top             =   1920
         Width           =   900
      End
      Begin VB.Label Label6 
         Caption         =   "繳交健檢報告的挸則："
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   510
         TabIndex        =   22
         Top             =   3270
         Width           =   1875
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "下次應繳年齡："
         Height          =   180
         Left            =   2820
         TabIndex        =   21
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "下次應繳年度："
         Height          =   180
         Left            =   600
         TabIndex        =   19
         Top             =   1590
         Width           =   1260
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "年齡："
         Height          =   180
         Left            =   2940
         TabIndex        =   17
         Top             =   930
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號："
         Height          =   180
         Index           =   0
         Left            =   600
         TabIndex        =   15
         Top             =   615
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "健檢日期："
         Height          =   180
         Left            =   -71595
         TabIndex        =   14
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74850
         TabIndex        =   13
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73230
         X2              =   -72540
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Line Line5 
         X1              =   -69900
         X2              =   -69300
         Y1              =   480
         Y2              =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "健檢日期："
         Height          =   180
         Left            =   600
         TabIndex        =   12
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "補助費用：          (Y.是)"
         Height          =   180
         Index           =   17
         Left            =   600
         TabIndex        =   11
         Top             =   1230
         Width           =   1815
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
            Picture         =   "frm160019.frx":009C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":03B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":06D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":08B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":0BCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":0EE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":1204
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":1520
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":183C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":1B58
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160019.frx":1E74
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   8
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
Attribute VB_Name = "frm160019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Create by Sindy 2015/8/11
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
Dim tf_SH As Integer


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
   rsA.Open "select * from staff_health where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SH = rsA.Fields.Count
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

ReDim m_FieldList(tf_SH) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSH01.BackColor = &H8000000F
   textSH02.BackColor = &H8000000F
   
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
   Set frm160019 = Nothing
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
            textSH01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSH02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
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
   
   If IsNull(rsSrcTmp.Fields("sh05")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh05")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("sh05"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sh06")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh06")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sh06"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sh07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh07")) = False Then
         strTemp = rsSrcTmp.Fields("sh07")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sh08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh08")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("sh08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sh09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sh09"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sh10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sh10")) = False Then
         strTemp = rsSrcTmp.Fields("sh10")
         strUTime = Format(strTemp, "##:##:##")
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
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textSH01.Enabled = True Then
      Cancel = False
      textSH01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSH01.Text = "" Then
       MsgBox "員工代號不可以空白！", vbExclamation
       textSH01.SetFocus
       Exit Function
   End If
   If Me.textSH02.Enabled = True Then
      Cancel = False
      textSH02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSH02.Text = "" Then
       MsgBox "健檢日期不可以空白！", vbExclamation
       textSH02.SetFocus
       Exit Function
   End If
   
   '增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSH01, True, textSH02) = True Then
      textSH01.SetFocus
      Exit Function
   End If
   
   If Me.textSH03.Enabled = True Then
      Cancel = False
      textSH03_Validate Cancel
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
   
   For nIndex = 0 To tf_SH - 1
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
   
   For nIndex = 0 To tf_SH - 1
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
Dim strSH01 As String
Dim strSH02 As String
   
   AddRecord = False
   
   strSH01 = textSH01
   strSH02 = DBDATE(textSH02)
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSH01, strSH02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_health("
   For nIndex = 0 To tf_SH - 1
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
   For nIndex = 0 To tf_SH - 1
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
   
   '存下次應繳年度
   strSql = "update staff set st68=" & Val(textST68) + 1911 & " where st01='" & textSH01 & "'"
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If ((strSH01 & strSH02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSH01 & strSH02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strSH01, DBDATE(strSH02)
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
   Dim strSH01 As String
   Dim strSH02 As String
       
   ModRecord = False
   
   strSH01 = m_CurrKEY(0)
   strSH02 = m_CurrKEY(1)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_health SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SH - 1
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
                  "WHERE SH01 = '" & strSH01 & "' and SH02='" & strSH02 & "' ; end; "
On Error GoTo ErrHand
   cnnConnection.BeginTrans
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   cnnConnection.CommitTrans

   ShowCurrRecord strSH01, DBDATE(strSH02)
      
   ModRecord = True
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strSH01 As String
Dim strSH02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSH01 = m_CurrKEY(0)
   strSH02 = m_CurrKEY(1)
   
   strSql = "DELETE FROM staff_health " & _
            "WHERE SH01 = '" & strSH01 & "'  and SH02='" & strSH02 & "' "
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Add By Sindy 2018/3/2
   '恢復下次應繳健檢報告年度
   strExc(0) = "select max(SH02) from staff_health where SH01='" & textSH01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If Not IsNull(RsTemp.Fields(0)) Then
         textSH02 = RsTemp.Fields(0) - 19110000
         Call CountOld
         strSql = "update staff set st68=" & Val(textST68) + 1911 & " where st01='" & textSH01 & "'"
      Else
         strSql = "update staff set st68=null where st01='" & textSH01 & "'"
      End If
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   '2018/3/2 END
   
   If (strSH01 = m_LastKEY(0) And strSH02 = m_LastKEY(1)) Or (strSH01 = m_FirstKEY(0) And strSH02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strSH01, DBDATE(strSH02)
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSH01 As String
Dim strSH02 As String
   
   QueryRecord = False
   strSH01 = textSH01
   strSH02 = DBDATE(textSH02)
   If IsRecordExist(strSH01, strSH02) = True Then
      m_CurrKEY(0) = strSH01
      m_CurrKEY(1) = strSH02
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
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSH01 <> "" And textSH02 <> "" Then
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
      Case 1: If Me.Visible = True Then textSH01.SetFocus
      Case 2: If Me.Visible = True Then textSH03.SetFocus
      Case 4: If Me.Visible = True Then textSH01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM staff_health " & _
            "WHERE SH01 = '" & strKEY01 & "'  and SH02='" & strKEY02 & "'  "
                  
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
      strSql = "SELECT SH01,SH02 FROM staff_health " & _
               "WHERE SH01 = '" & m_CurrKEY(0) & "' and SH02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
         If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT SH01,SH02 FROM staff_health " & _
               "WHERE SH02 = (SELECT MIN(SH02) FROM staff_health where SH01=(select min(SH01) from staff_health) ) and SH01=(select min(SH01) from staff_health) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
         If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
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
   
   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sh02 = (SELECT MAX(SH02) FROM staff_health " & _
                          "WHERE SH01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sh02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = (SELECT MAX(SH01) FROM staff_health " & _
                           "WHERE SH01 < '" & m_CurrKEY(0) & "') AND " & _
                  "sh02 = (SELECT MAX(SH02) FROM staff_health " & _
                           "WHERE SH01 = (SELECT MAX(SH01) FROM staff_health " & _
                                          "WHERE SH01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
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
   
   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sh02 = (SELECT MIN(SH02) FROM staff_health " & _
                          "WHERE SH01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sh02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = (SELECT MIN(SH01) FROM staff_health " & _
                           "WHERE SH01 > '" & m_CurrKEY(0) & "') AND " & _
                  "sh02 = (SELECT MIN(SH02) FROM staff_health " & _
                           "WHERE SH01 = (SELECT MIN(SH01) FROM staff_health " & _
                                          "WHERE SH01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sh02")
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
   
   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = (SELECT MIN(SH01) FROM staff_health) AND " & _
                  "sh02 = (SELECT MIN(SH02) FROM staff_health " & _
                           "WHERE SH01 = (SELECT MIN(SH01) FROM staff_health)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("sh02")
   End If
   rsTmp.Close

   strSql = "SELECT SH01,SH02 FROM staff_health " & _
            "WHERE SH01 = (SELECT MAX(SH01) FROM staff_health) AND " & _
                  "sh02 = (SELECT MAX(SH02) FROM staff_health " & _
                           "WHERE SH01 = (SELECT MAX(SH01) FROM staff_health)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sh01")) = False Then: m_LastKEY(0) = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: m_LastKEY(1) = rsTmp.Fields("sh02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT staff_health.*,ST68 FROM staff_health,staff " & _
            "WHERE SH01='" & m_CurrKEY(0) & "' and SH02 = '" & m_CurrKEY(1) & "' " & _
              "and SH01=ST01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("sh01")) = False Then: textSH01 = rsTmp.Fields("sh01")
      If IsNull(rsTmp.Fields("sh02")) = False Then: textSH02 = TAIWANDATE(rsTmp.Fields("sh02"))
      If IsNull(rsTmp.Fields("sh03")) = False Then: textSH03 = rsTmp.Fields("sh03")
      If IsNull(rsTmp.Fields("sh04")) = False Then: textSH04 = rsTmp.Fields("sh04")
      If IsNull(rsTmp.Fields("ST68")) = False Then: textST68 = Val(rsTmp.Fields("ST68")) - 1911
      
      strExc(0) = "select st23 from staff where st01='" & textSH01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Val("" & RsTemp.Fields("st23")) > 0 Then
            '年齡=健檢日期的年-出生年
            textOld = Left(DBDATE(textSH02), 4) - Left(RsTemp.Fields("st23"), 4)
            '下次應繳年齡=下次應繳年度-出生年
            If Val(textST68) > 0 Then
               textST68_2 = Val(textST68) - (Val(Left(RsTemp.Fields("st23"), 4)) - 1911)
            End If
         End If
      End If
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

      textSH01_2 = GetStaffName(textSH01, True)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and SH01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and SH01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and SH02>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and SH02<='" & DBDATE(txt1(3)) & "' "
   End If
   '抓取資料
   strSql = "SELECT SH01,st02,sqldateT(SH02),SH03,SH04 FROM staff_health,staff where SH01=st01(+) " & strSql & _
           " order by SH01,SH02 "
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
   textSH01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSH02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSH03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSH01.Locked = bEnable
   textSH02.Locked = bEnable
   If bEnable Then textSH01.BackColor = &H8000000F Else textSH01.BackColor = &H80000005
   If bEnable Then textSH02.BackColor = &H8000000F Else textSH02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSH01.Locked = bEnable
   textSH02.Locked = bEnable
   If bEnable Then textSH01.BackColor = &H8000000F Else textSH01.BackColor = &H80000005
   If bEnable Then textSH02.BackColor = &H8000000F Else textSH02.BackColor = &H80000005
   textSH03.Locked = bEnable
   textSH04.Locked = bEnable
End Sub

Private Sub ClearField()
Dim nIndex As Integer
   
   textSH01 = Empty
   textSH01_2 = Empty
   textSH02 = Empty
   textSH03 = Empty
   textSH04 = Empty
   textOld = Empty
   textST68 = Empty
   textST68_2 = Empty
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SH - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SH01", textSH01
      SetFieldNewData "SH02", DBDATE(textSH02)
   End If
   SetFieldNewData "SH03", textSH03
   SetFieldNewData "SH04", textSH04
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SH
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SH" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 6, 7, 9, 10:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSH01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSH01
      CloseIme
   End If
End Sub

Private Sub textSH01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSH01_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2017/2/6
   
   If m_EditMode = 1 And textSH01 <> "" Then
      textSH01_2 = GetStaffName(textSH01, True)
      If IsRecordExist(textSH01, DBDATE(textSH02)) = True And textSH01.Enabled = True And textSH01.Locked = False Then
         MsgBox "該員工當天已有健檢資料，請修改！", vbInformation
         Cancel = True
         Exit Sub
      End If
      If textSH01_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Cancel = True
         Exit Sub
      Else
         If ChkStaffST04(textSH01, False) = True Then
            MsgBox "此員工已離職！", vbInformation
         End If
      End If
      If textSH02 <> "" Then
         Call CountOld
      'Add by Sindy 2017/2/6 輸入員工編號時,顯示最近的健檢日期
      Else
         '抓取資料
         strSql = "SELECT SH02 FROM staff_health where SH01=" & CNULL(textSH01) & _
                  " order by SH02 desc"
         If rsTmp.State = 1 Then rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            textSH02.Text = Val(rsTmp.Fields(0)) - 19110000
         End If
         rsTmp.Close
         Set rsTmp = Nothing
      '2017/2/6 END
      End If
   End If
End Sub

Private Sub textSH02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSH02
   End If
End Sub

Private Sub textSH02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'計算年齡，下次繳交年度，下次繳交年齡
Private Sub CountOld()
Dim strNextOld As String
Dim strAddYear As String
   
   textOld = ""
   textST68 = ""
   textST68_2 = ""
   If textSH01 <> "" And textSH02 <> "" Then
      strExc(0) = "select st23 from staff where st01='" & textSH01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Val("" & RsTemp.Fields("st23")) > 0 Then
            '年齡=健檢日期的年-出生年
            textOld = Left(DBDATE(textSH02), 4) - Left(RsTemp.Fields("st23"), 4)
            '所屬區間:
            '1.  ～３９歲，每５年１次
            If textOld <= 39 Then
               strNextOld = 39
               strAddYear = 5
            '2. ４０～６４歲，每３年１次
            ElseIf textOld >= 40 And textOld <= 64 Then
               strNextOld = 40
               strAddYear = 3
            '3. ６５～歲，每年１次
            Else
               strNextOld = 65
               strAddYear = 1
            End If
            textST68_2 = Val(textOld) + Val(strAddYear)
            '檢查是否有超過下一階段的年齡
            If textST68_2 > strNextOld Then
               If textST68_2 >= 65 Then
                  strNextOld = 65
                  strAddYear = 1
               Else
                  strNextOld = 40
                  strAddYear = 3
               End If
               textST68_2 = Val(textOld) + Val(strAddYear)
               '計算出來的應繳年齡是否沒有超過下一階段的年齡,若沒有,則以下一階段的年齡為下次應繳年齡
               If textST68_2 < strNextOld Then
                  textST68_2 = strNextOld
               End If
            End If
            '下次應繳年度
            textST68 = textST68_2 + (Left(RsTemp.Fields("st23"), 4) - 1911)
         End If
      End If
   End If
End Sub

Private Sub textSH02_Validate(Cancel As Boolean)
   If m_EditMode = 1 And textSH02 <> "" Then
      If IsRecordExist(textSH01, DBDATE(textSH02)) = True And textSH02.Enabled = True And textSH02.Locked = False Then
         MsgBox "該員工當天已有健檢資料，請修改！", vbInformation
         Cancel = True
         Exit Sub
      End If
      If CheckIsTaiwanDate(textSH02, False) = False Then
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      If textSH01 <> "" And m_EditMode = 1 Then '新增
         Call CountOld
      End If
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("員工編號", "姓名", "健檢日期", "補助費用", "備註")
   arrGridHeadWidth = Array(900, 1000, 900, 900, 4000)
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

Private Sub textSH03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSH03
      CloseIme
   End If
End Sub

Private Sub textSH03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSH03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSH03 <> "" Then
      Select Case textSH03
      Case "Y", ""
      Case Else
         MsgBox "原因請輸入 Y 或 空白！", vbExclamation, "輸入錯誤！"
         Cancel = True
         Exit Sub
      End Select
   End If
   CloseIme
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
