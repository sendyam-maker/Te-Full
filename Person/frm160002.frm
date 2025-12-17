VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160002 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤資料"
   ClientHeight    =   5070
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8170
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8170
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
            Picture         =   "frm160002.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160002.frx":1DD8
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
      Width           =   8170
      _ExtentX        =   14411
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   4230
      Left            =   15
      TabIndex        =   12
      Top             =   780
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7461
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160002.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label23"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "textSA01_2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "textSA02"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textSA03"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textSA05"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textSA06"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textSA04"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textSA01"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160002.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label10"
      Tab(1).Control(1)=   "Line1"
      Tab(1).Control(2)=   "Label11"
      Tab(1).Control(3)=   "Line2"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "txt1(0)"
      Tab(1).Control(6)=   "txt1(1)"
      Tab(1).Control(7)=   "txt1(2)"
      Tab(1).Control(8)=   "txt1(3)"
      Tab(1).Control(9)=   "cmdok"
      Tab(1).ControlCount=   10
      Begin VB.TextBox textSA01 
         Height          =   270
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   0
         Top             =   480
         Width           =   945
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -68670
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70050
         MaxLength       =   7
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71040
         MaxLength       =   7
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72930
         MaxLength       =   6
         TabIndex        =   7
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   6
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox textSA04 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1380
         Width           =   585
      End
      Begin VB.TextBox textSA06 
         Height          =   285
         Left            =   2070
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1710
         Width           =   585
      End
      Begin VB.TextBox textSA05 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1710
         Width           =   585
      End
      Begin VB.TextBox textSA03 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1050
         Width           =   585
      End
      Begin VB.TextBox textSA02 
         Height          =   270
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   1
         Top             =   750
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160002.frx":212C
         Height          =   3285
         Left            =   -74970
         TabIndex        =   13
         Top             =   780
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5786
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
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "注意: 不能輸超過480分鐘 (8小時)"
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   3090
         TabIndex        =   27
         Top             =   1770
         Width           =   2570
      End
      Begin MSForms.Label textSA01_2 
         Height          =   225
         Left            =   2040
         TabIndex        =   26
         Top             =   525
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
         Left            =   150
         TabIndex        =   25
         Top             =   3900
         Width           =   7785
         VariousPropertyBits=   27
         Caption         =   "CREATE :                                                    UPDATE : "
         Size            =   "13732;344"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   180
         Index           =   0
         Left            =   250
         TabIndex        =   24
         Top             =   525
         Width           =   720
      End
      Begin VB.Line Line2 
         X1              =   -70410
         X2              =   -69660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71640
         TabIndex        =   23
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line1 
         X1              =   -73260
         X2              =   -72660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   22
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "曠職"
         Height          =   180
         Left            =   630
         TabIndex        =   21
         Top             =   1755
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "遲到"
         Height          =   180
         Left            =   630
         TabIndex        =   20
         Top             =   1440
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "次"
         Height          =   180
         Left            =   1700
         TabIndex        =   19
         Top             =   1440
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "分"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   2760
         TabIndex        =   18
         Top             =   1760
         Width           =   190
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "日"
         Height          =   180
         Left            =   1700
         TabIndex        =   17
         Top             =   1755
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "次"
         Height          =   180
         Left            =   1700
         TabIndex        =   16
         Top             =   1110
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "忘打卡"
         Height          =   180
         Left            =   450
         TabIndex        =   15
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "日期"
         Height          =   180
         Index           =   17
         Left            =   630
         TabIndex        =   14
         Top             =   810
         Width           =   360
      End
   End
End
Attribute VB_Name = "frm160002"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/15 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2006/11/01 copy from frm140401
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
Dim tf_SA As Integer
Dim MyKind As String
Dim m_intSA03 As Integer, m_intSA04 As Integer 'Add By Sindy 2014/3/17


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
   rsA.Open "select * from staff_assist where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SA = rsA.Fields.Count
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
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   ReDim m_FieldList(tf_SA) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSA01.BackColor = &H8000000F
   textSA02.BackColor = &H8000000F
   
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
   Set frm160002 = Nothing
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
   '         textSA01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
   '         textSA01_2 = GetStaffName(textSA01, True)
   '         textSA02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
   '         textSA03.Text = GRD1.TextMatrix(tmpMouseRow, 3)
   '         textSA04.Text = GRD1.TextMatrix(tmpMouseRow, 4)
   '         textSA05.Text = GRD1.TextMatrix(tmpMouseRow, 5)
   '         textSA06.Text = GRD1.TextMatrix(tmpMouseRow, 6)
            '2008/12/12 ADD BY SONIA
            textSA01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
            textSA02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
            QueryRecord
            '2008/12/12 END
            GRD1.Visible = True
       End If
   End If
End Sub

'Add By Sindy 2019/8/27
Private Sub SSTab1_Click(PreviousTab As Integer)
   If PreviousTab = 0 Then
      cmdOK.SetFocus
      cmdOK.Default = True
   Else
      cmdOK.Default = False
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

   If IsNull(rsSrcTmp.Fields("sa07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("sa07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sa08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa09")) = False Then
         strTemp = rsSrcTmp.Fields("sa09")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("sa10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sa11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sa12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sa12")) = False Then
         strTemp = rsSrcTmp.Fields("sa12")
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
   If Me.textSA01.Enabled = True Then
      Cancel = False
      textSA01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSA01.Text = "" Then
       MsgBox "員工編號不可以空白！", vbExclamation
       textSA01.SetFocus
       Exit Function
   End If
   If Me.textSA02.Enabled = True Then
      Cancel = False
      textSA02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSA02.Text = "" Then
       MsgBox "日期不可以空白！", vbExclamation
       textSA02.SetFocus
       Exit Function
   End If
   
   'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
   If ChkStaffST04(textSA01, True, textSA02) = True Then
      textSA01.SetFocus
      Exit Function
   End If
   
   If Me.textSA03.Enabled = True Then
      Cancel = False
      textsa03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSA04.Enabled = True Then
      Cancel = False
      textSA04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSA05.Enabled = True Then
      Cancel = False
      textSA05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSA06.Enabled = True Then
      Cancel = False
      textSA06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   TxtValidate = True
End Function

'add by nickc 2006/10/24
' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
Dim nIndex As Integer
   
   For nIndex = 0 To tf_SA - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   
   For nIndex = 0 To tf_SA - 1
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
Dim strSA01 As String
Dim strSA02 As String
   
   AddRecord = False
   
   strSA01 = textSA01
   strSA02 = DBDATE(textSA02)

   ' 檢查記錄是否已存在
   If IsRecordExist(strSA01, strSA02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_assist ("
   For nIndex = 0 To tf_SA - 1
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
   For nIndex = 0 To tf_SA - 1
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
   
   If ((strSA01 & strSA02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSA01 & strSA02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   cnnConnection.CommitTrans
   
   ShowCurrRecord strSA01, DBDATE(strSA02)
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
Dim strSA01 As String
Dim strSA02 As String
       
   ModRecord = False
   
   strSA01 = m_CurrKEY(0)
   strSA02 = m_CurrKEY(1)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_assist SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SA - 1
      strTmp = Empty
      'If nIndex < 7 Or nIndex > 12 Then
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
                  "WHERE sa01 = '" & strSA01 & "' and sa02='" & strSA02 & "' ; end; "
On Error GoTo ErrHand
      cnnConnection.BeginTrans
         If bDifference = True Then
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
           
            'Add By Sindy 2014/3/17 恢復打卡異常至未確認狀態
            If Val(DBDATE(textSA02)) >= 20140301 Then
               If (m_intSA03 > 0 And m_intSA03 > Val(textSA03)) Or _
                  (m_intSA04 > 0 And m_intSA04 > Val(textSA04)) Then
                  If RestoreABS014 = False Then GoTo ErrHand
               End If
            End If
            '2014/3/17 END
         End If
         cnnConnection.CommitTrans

      ShowCurrRecord strSA01, DBDATE(strSA02)
      
   ModRecord = True
   Exit Function
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
Dim strSql As String
Dim strSA01 As String
Dim strSA02 As String
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSA01 = m_CurrKEY(0)
   strSA02 = m_CurrKEY(1)

   strSql = "DELETE FROM staff_assist " & _
            "WHERE sa01 = '" & strSA01 & "'  and sa02='" & strSA02 & "' "
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Add By Sindy 2014/3/17 恢復打卡異常至未確認狀態
   If Val(DBDATE(textSA02)) >= 20140301 Then
      If RestoreABS014 = False Then GoTo ErrHand
   End If
   '2014/3/17 END
   
   If (strSA01 = m_LastKEY(0) And strSA02 = m_LastKEY(1)) Or (strSA01 = m_FirstKEY(0) And strSA02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strSA01, DBDATE(strSA02)
   DelRecord = True
   cnnConnection.CommitTrans
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
Dim strSA01 As String
Dim strSA02 As String
   
   QueryRecord = False
   strSA01 = textSA01
   strSA02 = DBDATE(textSA02)
   If IsRecordExist(strSA01, strSA02) = True Then
      m_CurrKEY(0) = strSA01
      m_CurrKEY(1) = strSA02
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
            'Add By Cheng 2002/05/22
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
         If textSA01 <> "" And textSA02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            ' 2008/12/17 ADD BY SINDY
            If textSA01 = "" Or textSA02 = "" Then
               MsgBox "須輸入員工代號及日期才可進行查詢動作！", vbInformation
            End If
            ' 2008/12/17 END
            
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
      Case 1: If Me.Visible = True Then textSA01.SetFocus
      Case 2: If Me.Visible = True Then textSA03.SetFocus
      Case 4: If Me.Visible = True Then textSA01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM staff_assist " & _
            "WHERE sa01 = '" & strKEY01 & "'  and sa02='" & strKEY02 & "'  "
                  
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
      strSql = "SELECT sa01,sa02 FROM staff_assist " & _
               "WHERE sa01 = '" & m_CurrKEY(0) & "' and sa02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
         If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT sa01,sa02 FROM staff_assist " & _
               "WHERE sa02 = (SELECT MIN(sa02) FROM staff_assist where sa01=(select min(sa01) from staff_assist) ) and sa01=(select min(sa01) from staff_assist) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
         If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
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
   
   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sa02 = (SELECT MAX(sa02) FROM staff_assist " & _
                          "WHERE sa01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sa02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = (SELECT MAX(sa01) FROM staff_assist " & _
                           "WHERE sa01 < '" & m_CurrKEY(0) & "') AND " & _
                  "sa02 = (SELECT MAX(sa02) FROM staff_assist " & _
                           "WHERE sa01 = (SELECT MAX(sa01) FROM staff_assist " & _
                                          "WHERE sa01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
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
   
   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = '" & m_CurrKEY(0) & "' AND " & _
                  "sa02 = (SELECT MIN(sa02) FROM staff_assist " & _
                          "WHERE sa01 = '" & m_CurrKEY(0) & "' AND " & _
                                "sa02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = (SELECT MIN(sa01) FROM staff_assist " & _
                           "WHERE sa01 > '" & m_CurrKEY(0) & "') AND " & _
                  "sa02 = (SELECT MIN(sa02) FROM staff_assist " & _
                           "WHERE sa01 = (SELECT MIN(sa01) FROM staff_assist " & _
                                          "WHERE sa01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("sa02")
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
'      tabCustomer.Tab = 0
   End If
End Sub

Private Sub RefreshRange()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = (SELECT MIN(sa01) FROM staff_assist) AND " & _
                  "sa02 = (SELECT MIN(sa02) FROM staff_assist " & _
                           "WHERE sa01 = (SELECT MIN(sa01) FROM staff_assist)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("sa02")
   End If
   rsTmp.Close

   strSql = "SELECT sa01,sa02 FROM staff_assist " & _
            "WHERE sa01 = (SELECT MAX(sa01) FROM staff_assist) AND " & _
                  "sa02 = (SELECT MAX(sa02) FROM staff_assist " & _
                           "WHERE sa01 = (SELECT MAX(sa01) FROM staff_assist)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("sa01")) = False Then: m_LastKEY(0) = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: m_LastKEY(1) = rsTmp.Fields("sa02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim i As Integer, j As Integer
   
   strSql = "SELECT * FROM staff_assist " & _
            "WHERE sa01='" & m_CurrKEY(0) & "' and sa02 = '" & m_CurrKEY(1) & "'   "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("sa01")) = False Then: textSA01 = rsTmp.Fields("sa01")
      If IsNull(rsTmp.Fields("sa02")) = False Then: textSA02 = TAIWANDATE(rsTmp.Fields("sa02"))
      If IsNull(rsTmp.Fields("sa03")) = False Then: textSA03 = rsTmp.Fields("sa03")
      If IsNull(rsTmp.Fields("sa04")) = False Then: textSA04 = rsTmp.Fields("sa04")
      If IsNull(rsTmp.Fields("sa05")) = False Then: textSA05 = rsTmp.Fields("sa05")
      If IsNull(rsTmp.Fields("sa06")) = False Then: textSA06 = rsTmp.Fields("sa06")
      'Add By Sindy 2014/3/17
      m_intSA03 = Val(textSA03)
      m_intSA04 = Val(textSA04)
      '2014/3/17 END
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

      textSA01_2 = GetStaffName(textSA01, True)
      Call Pub_GetSpecWorkHour(textSA01, textSA02) '特殊人員的工作時數 Add By Sindy 2025/9/4
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub
Sub GetData()
Dim rsTmp As New ADODB.Recordset

   strSql = ""
   If txt1(0) <> "" Then
       strSql = strSql & " and sa01>='" & txt1(0) & "' "
   End If
   If txt1(1) <> "" Then
       strSql = strSql & " and sa01<='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       strSql = strSql & " and sa02>='" & DBDATE(txt1(2)) & "' "
   End If
   If txt1(3) <> "" Then
       strSql = strSql & " and sa02<='" & DBDATE(txt1(3)) & "' "
   End If
   '抓取資料
   ' 2008/12/18 Modify BY SINDY
'   strSQL = "SELECT sa01,st02,sqldateT(sa02),sa03,sa04,sa05,sa06,sa07  FROM staff_assist,staff where sa01=st01(+)  " & strSQL & _
'           " order by sa01,sa02 "
   'Modify By Sindy 2025/9/4 sa06 => round(sa06,0)
   strSql = "SELECT sa01,st02,sqldateT(sa02),sa03,sa04,sa05,round(sa06,0) FROM staff_assist,staff where sa01=st01(+)  " & strSql & _
           " order by sa02,sa01 "
   ' 2008/12/18 END
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
Dim strTit As String
Dim strMsg As String
   
   CheckDataValid = False
   
   'Add By Sindy 2014/3/14
   If textSA03 = "0" Then textSA03 = ""
   If textSA04 = "0" Then textSA04 = ""
   If textSA05 = "0" Then textSA05 = ""
   If textSA06 = "0" Then textSA06 = ""
   If Val(textSA03) = 0 And Val(textSA04) = 0 And _
      Val(textSA05) = 0 And Val(textSA06) = 0 Then
      strTit = "檢核資料"
      strMsg = "請至少輸入一項出缺勤資料!!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSA03.SetFocus
      GoTo EXITSUB
   End If
   '2014/3/14 END
   
   nResponse = False
   textSA01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textsa03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA05_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSA06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   
   CheckDataValid = True
EXITSUB:
End Function
' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSA01.Locked = bEnable
   textSA02.Locked = bEnable
   If bEnable Then textSA01.BackColor = &H8000000F Else textSA01.BackColor = &H80000005
   If bEnable Then textSA02.BackColor = &H8000000F Else textSA02.BackColor = &H80000005
End Sub
' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   
   textSA01.Locked = bEnable
   textSA02.Locked = bEnable
   If bEnable Then textSA01.BackColor = &H8000000F Else textSA01.BackColor = &H80000005
   If bEnable Then textSA02.BackColor = &H8000000F Else textSA02.BackColor = &H80000005
   textSA03.Locked = bEnable
   textSA04.Locked = bEnable
   textSA05.Locked = bEnable
   textSA06.Locked = bEnable
End Sub
Private Sub ClearField()
Dim nIndex As Integer
   
   textSA01 = Empty
   textSA01_2 = Empty
   textSA02 = Empty
   textSA03 = Empty
   textSA04 = Empty
   textSA05 = Empty
   textSA06 = Empty
   Label23 = Empty
   
'   ' 2008/12/17 ADD BY SINDY
'   txt1(0) = Empty
'   txt1(1) = Empty
'   txt1(2) = Empty
'   txt1(3) = Empty
'   Grd1.Clear
'   ' 2008/12/17 END
   
   SetGrd
   For nIndex = 0 To tf_SA - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SA01", textSA01
      SetFieldNewData "SA02", DBDATE(textSA02)
   End If
   SetFieldNewData "SA03", textSA03
   SetFieldNewData "SA04", textSA04
   SetFieldNewData "SA05", textSA05
   SetFieldNewData "SA06", textSA06
End Sub

' 初始化欄位陣列
Private Sub InitialField()
Dim nIndex As Integer
Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SA
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SA" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 3, 4, 5, 6, 7:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
   SetGrd
End Sub

Private Sub textSA01_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSA01
   End If
End Sub

Private Sub textSA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSA01_Validate(Cancel As Boolean)
   If textSA01.Text = "" Then textSA01_2 = "" ' 2008/12/18 ADD BY SINDY
   
   If m_EditMode <> 0 And textSA01 <> "" Then
      ' 2008/12/17 ADD BY SINDY
      ' 檢查員工編號規則
      If ChkStaffID(textSA01) Then
         Call textSA01_GotFocus
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 END
      textSA01_2 = GetStaffName(textSA01, True)
      If textSA01_2 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call textSA01_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
   End If
   
   If m_EditMode = 1 And textSA01 <> "" Then
      If IsRecordExist(textSA01, DBDATE(textSA02)) = True And textSA01.Enabled = True And textSA01.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSA01_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
   End If
   If textSA01 <> "" Then
      Call Pub_GetSpecWorkHour(textSA01, textSA02) '特殊人員的工作時數 Add By Sindy 2025/9/4
   End If
End Sub

Private Sub textSA02_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSA02
      CloseIme
   End If
End Sub

Private Sub textSA02_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA02_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSA02 <> "" Then
      If CheckIsTaiwanDate(textSA02, False) = False Then
         Call textSA02_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      ElseIf ChkWork(ChangeTStringToWString(textSA02)) = False Then
         Call textSA02_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
   End If
   
   If m_EditMode = 1 And textSA02 <> "" Then
      If IsRecordExist(textSA01, DBDATE(textSA02)) = True And textSA02.Enabled = True And textSA02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSA02_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
   End If
   If textSA02 <> "" Then
      Call Pub_GetSpecWorkHour(textSA01, textSA02) '特殊人員的工作時數 Add By Sindy 2025/9/4
   End If
End Sub

Private Sub textsa03_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSA03
      CloseIme
   End If
End Sub

Private Sub textSA03_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textsa03_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSA03 <> "" Then
       If CheckLengthIsOK(textSA03, textSA03.MaxLength) = False Then
           Call textsa03_GotFocus ' 2008/12/17 ADD BY SINDY
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 ADD BY SINDY
       If textSA03.Text > 31 Then
           Call textsa03_GotFocus
           MsgBox "忘打卡(次)不可超過31次!", vbExclamation + vbOKOnly
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   ' 2008/12/18 Modify BY SINDY
'   arrGridHeadText = Array("員工編號", "姓名", "日期", "忘打卡", "遲到", "曠職日", "曠職時", "曠職分")
'   arrGridHeadWidth = Array(800, 1200, 1200, 800, 800, 800, 800, 800)
   'Modify By Sindy 2025/8/28 曠職時 => 改為 曠職分
   'arrGridHeadText = Array("員工編號", "姓名", "日期", "忘打卡", "遲到", "曠職日", "曠職時")
   arrGridHeadText = Array("員工編號", "姓名", "日期", "忘打卡", "遲到", "曠職日", "曠職分")
   arrGridHeadWidth = Array(800, 1200, 1200, 800, 800, 800, 800)
   ' 2008/12/18 END
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

Private Sub textSA04_GotFocus()
   If m_EditMode <> 0 Then
       InverseTextBox textSA04
       CloseIme
   End If
End Sub

Private Sub textSA04_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA04_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSA04 <> "" Then
      If CheckLengthIsOK(textSA04, textSA04.MaxLength) = False Then
         Call textSA04_GotFocus ' 2008/12/17 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
      ' 2008/12/17 ADD BY SINDY
       If textSA04.Text > 31 Then
           Call textSA04_GotFocus
           MsgBox "遲到(次)不可超過31次!", vbExclamation + vbOKOnly
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub textSA05_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSA05
      CloseIme
   End If
End Sub

Private Sub textSA05_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSA05_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSA05 <> "" Then
       If CheckLengthIsOK(textSA05, textSA05.MaxLength) = False Then
           Call textSA05_GotFocus ' 2008/12/17 ADD BY SINDY
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 ADD BY SINDY
       If textSA05.Text > 31 Or (textSA05.Text = 31 And textSA06.Text <> "") Then
           Call textSA05_GotFocus
           MsgBox "曠職(日)不可超過31天!", vbExclamation + vbOKOnly
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 END
   End If
   CloseIme
End Sub

Private Sub textSA06_GotFocus()
   If m_EditMode <> 0 Then
      InverseTextBox textSA06
      CloseIme
   End If
End Sub

Private Sub textSA06_KeyPress(KeyAscii As Integer)
   ' 2008/12/18 Modify BY SINDY
   'KeyAscii = Pub_NumAscii(KeyAscii)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
   ' 2008/12/18 END
End Sub

Private Sub textSA06_Validate(Cancel As Boolean)
   If m_EditMode <> 0 And textSA06 <> "" Then
       If CheckLengthIsOK(textSA06, textSA06.MaxLength) = False Then
           Call textSA06_GotFocus ' 2008/12/17 ADD BY SINDY
           Cancel = True
           Exit Sub
       End If
       ' 2008/12/17 ADD BY SINDY
       'Modify By Sindy 2025/8/28
'       If textSA06.Text >= 8 Then
'           Call textSA06_GotFocus
'           MsgBox "曠職(時)不可超過8小時!!!", vbExclamation + vbOKOnly
'           Cancel = True
'           Exit Sub
'       End If
       strExc(10) = PUB_intWkHour * 60
       If Val(textSA06.Text) >= Val(strExc(10)) Then
           Call textSA06_GotFocus
           MsgBox "曠職(分)不可超過 " & strExc(10) & " 分鐘!!!", vbExclamation + vbOKOnly
           Cancel = True
           Exit Sub
       End If
       '2025/8/28 END
       ' 2008/12/17 END
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
   Select Case Index
      Case 0, 1
         ' 2008/12/17 ADD BY SINDY
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
         ' 2008/12/17 END
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         
      Case 2, 3
         ' 2008/12/16 MODIFY BY SINDY
         'If CheckIsTaiwanDate(txt1(Index), False) = False Then
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
         ' 2008/12/16 END
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         
         ' 2008/12/17 ADD BY SINDY
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ' 2008/12/17 END
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

'Add By Sindy 2014/3/17 恢復打卡異常至未確認狀態
Private Function RestoreABS014() As Boolean
Dim rsTmp As New ADODB.Recordset
   
On Error GoTo ErrHand
   
   RestoreABS014 = True
   strSql = "SELECT * FROM ABS014 " & _
            "WHERE b1401 = '" & textSA01 & "' and b1402=" & DBDATE(textSA02) & " and b1405 in('2','3') and b1411 is not null"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'Modify By Sindy 2015/4/2 個人確認原因也要清空(b1405=null)
      strSql = "UPDATE ABS014 SET " & _
               "b1405=null,b1411=null,b1412=null,b1413=null " & _
               "WHERE b1401 = '" & textSA01 & "' and b1402=" & DBDATE(textSA02) & " and b1405 in('2','3') and b1411 is not null"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   Exit Function
   
ErrHand:
   RestoreABS014 = False
   Set rsTmp = Nothing
   MsgBox "恢復打卡異常至未確認狀態失敗！" & vbCrLf & Err.Description
End Function
