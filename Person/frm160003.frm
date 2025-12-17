VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160003 
   BorderStyle     =   1  '單線固定
   Caption         =   "加班資料"
   ClientHeight    =   5060
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8200
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5060
   ScaleWidth      =   8200
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
            Picture         =   "frm160003.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160003.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4350
      Left            =   30
      TabIndex        =   15
      Top             =   660
      Width           =   8115
      _ExtentX        =   14323
      _ExtentY        =   7673
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160003.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(17)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Line1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Line2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label12"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "textSO01_2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label23"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtNote"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textSO04_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textSO04_1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textSO03_1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textSO02"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textSO03_2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textSO15"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textSO01"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textSO0506"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Frame1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdABS"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160003.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Line3"
      Tab(1).Control(1)=   "Label11"
      Tab(1).Control(2)=   "Line4"
      Tab(1).Control(3)=   "Label10"
      Tab(1).Control(4)=   "GRD1"
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(6)=   "txt1(3)"
      Tab(1).Control(7)=   "txt1(2)"
      Tab(1).Control(8)=   "txt1(1)"
      Tab(1).Control(9)=   "txt1(0)"
      Tab(1).ControlCount=   10
      Begin VB.CommandButton cmdABS 
         BackColor       =   &H00C0FFFF&
         Caption         =   "簽核資料"
         Height          =   315
         Left            =   5610
         Style           =   1  '圖片外觀
         TabIndex        =   35
         Top             =   1710
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   1575
         Left            =   90
         TabIndex        =   29
         Top             =   2310
         Visible         =   0   'False
         Width           =   2175
         Begin VB.TextBox textSO13 
            Height          =   285
            Left            =   870
            MaxLength       =   8
            TabIndex        =   8
            Top             =   60
            Width           =   1095
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "表單編號"
            Height          =   180
            Left            =   120
            TabIndex        =   30
            Top             =   90
            Width           =   720
         End
      End
      Begin VB.TextBox textSO0506 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1950
         Width           =   885
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -73980
         MaxLength       =   6
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72930
         MaxLength       =   6
         TabIndex        =   10
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71040
         MaxLength       =   7
         TabIndex        =   11
         Top             =   360
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70050
         MaxLength       =   7
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -68670
         TabIndex        =   13
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox textSO01 
         Height          =   270
         Left            =   960
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox textSO15 
         Height          =   315
         Left            =   960
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1620
         Width           =   885
      End
      Begin VB.TextBox textSO03_2 
         Height          =   285
         Left            =   1860
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1230
         Width           =   585
      End
      Begin VB.TextBox textSO02 
         Height          =   270
         Left            =   960
         MaxLength       =   7
         TabIndex        =   1
         Top             =   660
         Width           =   945
      End
      Begin VB.TextBox textSO03_1 
         Height          =   285
         Left            =   1020
         MaxLength       =   2
         TabIndex        =   2
         Top             =   1230
         Width           =   585
      End
      Begin VB.TextBox textSO04_1 
         Height          =   285
         Left            =   3270
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1230
         Width           =   585
      End
      Begin VB.TextBox textSO04_2 
         Height          =   285
         Left            =   4140
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1230
         Width           =   585
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160003.frx":212C
         Height          =   3615
         Left            =   -74970
         TabIndex        =   16
         Top             =   690
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
      Begin MSForms.TextBox txtNote 
         Height          =   1095
         Left            =   3750
         TabIndex        =   34
         Top             =   2370
         Width           =   3345
         VariousPropertyBits=   -1466939365
         MaxLength       =   100
         ScrollBars      =   3
         Size            =   "5900;1931"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   150
         TabIndex        =   33
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
      Begin MSForms.Label textSO01_2 
         Height          =   225
         Left            =   1770
         TabIndex        =   32
         Top             =   405
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "修改/刪除原因："
         Height          =   180
         Left            =   2400
         TabIndex        =   31
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "假日-共                        時"
         Height          =   180
         Left            =   210
         TabIndex        =   28
         Top             =   1980
         Width           =   1860
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74910
         TabIndex        =   27
         Top             =   390
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -73260
         X2              =   -72660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71640
         TabIndex        =   26
         Top             =   390
         Width           =   540
      End
      Begin VB.Line Line3 
         X1              =   -70410
         X2              =   -69660
         Y1              =   450
         Y2              =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工代號"
         Height          =   180
         Index           =   0
         Left            =   210
         TabIndex        =   25
         Top             =   405
         Width           =   720
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   3270
         X2              =   4950
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   1020
         X2              =   2700
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "實際時數"
         Height          =   180
         Left            =   210
         TabIndex        =   24
         Top             =   1650
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "～"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2850
         TabIndex        =   23
         Top             =   1260
         Width           =   255
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "分"
         Height          =   180
         Left            =   2490
         TabIndex        =   22
         Top             =   1290
         Width           =   180
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "時間                起                                              迄"
         Height          =   180
         Left            =   600
         TabIndex        =   21
         Top             =   990
         Width           =   3510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "加班日期"
         Height          =   180
         Index           =   17
         Left            =   210
         TabIndex        =   20
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "時"
         Height          =   180
         Left            =   1650
         TabIndex        =   19
         Top             =   1290
         Width           =   180
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "時"
         Height          =   180
         Left            =   3900
         TabIndex        =   18
         Top             =   1290
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "分"
         Height          =   180
         Left            =   4770
         TabIndex        =   17
         Top             =   1290
         Width           =   180
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8200
      _ExtentX        =   14464
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
Attribute VB_Name = "frm160003"
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
Dim m_FirstKEY(3) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(3) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(3) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tf_so As Integer
Dim MyKind As String
'Add By Sindy 2011/9/22
Dim m_B1019 As String, m_B1004 As String, m_B1005 As String
Dim m_B1007 As String, m_B1012 As String, m_B1013 As String
Dim m_B1017 As String, m_B1030 As String
Dim m_KeyCode As String 'Add By Sindy 2011/10/7
'Add By Sindy 2012/5/23
Dim m_SO02 As String, m_SO03_1 As String, m_SO03_2 As String
Dim m_SO04_1 As String, m_SO04_2 As String
'2012/5/23 End


'Add By Sindy 2022/10/28
Private Sub cmdABS_Click()
   Me.Hide
   Call frm180301_03.SetParent(Me)
   frm180301_03.txtB1001 = textSO13.Text
   frm180301_03.QueryData
   frm180301_03.Show
End Sub

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
rsA.Open "select * from staff_overtime where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_so = rsA.Fields.Count
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

   ReDim m_FieldList(tf_so) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSO01.BackColor = &H8000000F
   textSO02.BackColor = &H8000000F
   ' 2008/12/22 Add BY SINDY
   textSO03_1.BackColor = &H8000000F
   textSO03_2.BackColor = &H8000000F
   ' 2008/12/22 END
   
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
'         textSO01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
'         textSO01_2 = GetStaffName(textSO01, True)
'         textSO02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
'         textso03.Text = GRD1.TextMatrix(tmpMouseRow, 3)
'         textSA04.Text = GRD1.TextMatrix(tmpMouseRow, 4)
'         textSA05.Text = GRD1.TextMatrix(tmpMouseRow, 5)
'         textSA06.Text = GRD1.TextMatrix(tmpMouseRow, 6)
'         textSA07.Text = GRD1.TextMatrix(tmpMouseRow, 7)
         '2008/12/12 ADD BY SONIA
         textSO01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textSO02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
         ' 2008/12/22 Add BY SINDY
         textSO03_1.Text = Left(Trim(GRD1.TextMatrix(tmpMouseRow, 3)), 2)
         textSO03_2.Text = Right(Trim(GRD1.TextMatrix(tmpMouseRow, 3)), 2)
         ' 2008/12/22 END
         QueryRecord
         '2008/12/12 END
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
   If IsNull(rsSrcTmp.Fields("so07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("so07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("so08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so09")) = False Then
         strTemp = rsSrcTmp.Fields("so09")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("so10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("so11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so12")) = False Then
         strTemp = rsSrcTmp.Fields("so12")
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
If Me.textSO01.Enabled = True Then
   Cancel = False
   textSO01_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSO01.Text = "" Then
    MsgBox "員工編號不可以空白！", vbExclamation
    textSO01.SetFocus
    Exit Function
End If

'Add By Sindy 2011/9/22
If Me.Frame1.Visible = True And Me.textSO13.Enabled = True Then
   If m_EditMode = 1 And textSO13 = "" Then
      MsgBox "表單編號不可空白！", vbExclamation
      textSO13.SetFocus
      Exit Function
   End If
   
   Cancel = False
   textSO13_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textSO02.Enabled = True Then
   Cancel = False
   textSO02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textSO02.Text = "" Then
    MsgBox "加班日期不可以空白！", vbExclamation
    textSO02.SetFocus
    Exit Function
End If

'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
If ChkStaffST04(textSO01, True, textSO02) = True Then
   textSO01.SetFocus
   Exit Function
End If

If Me.textSO03_1.Enabled = True Then
   Cancel = False
   textso03_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/22 Add BY SINDY
If textSO03_1.Text = "" Then
    MsgBox "起始時間(時)不可以空白！", vbExclamation
    textSO03_1.SetFocus
    Exit Function
End If
' 2008/12/22 END
If Me.textSO03_2.Enabled = True Then
   Cancel = False
   textso03_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/22 Add BY SINDY
If textSO03_2.Text = "" Then
    MsgBox "起始時間(分)不可以空白！", vbExclamation
    textSO03_2.SetFocus
    Exit Function
End If
' 2008/12/22 END
If Me.textSO04_1.Enabled = True Then
   Cancel = False
   textSo04_1_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSO04_2.Enabled = True Then
   Cancel = False
   textSo04_2_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textSO15.Enabled = True Then
   Cancel = False
   textSo15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
' 2008/12/18 Add BY SINDY
'If Me.textSO06.Enabled = True Then
'   Cancel = False
'   textSo06_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'If (textSO05.Text <= "" And textSO06.Text <= "") Or _
'      (textSO05.Text = "0" And textSO06.Text = "0") Or _
'      (Trim(textSO05.Text) = "" And textSO06.Text = "0") Or _
'      (textSO05.Text = "0" And Trim(textSO06.Text) = "") Then
'    MsgBox "無加班時數！", vbExclamation
'    'textSO05.SetFocus
'    Exit Function
'End If
If Val(textSO15.Text) <= 0 Then
   MsgBox "無加班時數！", vbExclamation
   Exit Function
End If
' 2008/12/18 END

'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me) = False Then
   Exit Function
End If
'2021/9/1 END

TxtValidate = True
End Function

'add by nickc 2006/10/24
' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To tf_so - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   
   For nIndex = 0 To tf_so - 1
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
   Dim strSO01 As String
   Dim strSO02 As String
   Dim strSO03 As String
   Dim strContent As String, strSubject As String
   Dim rsTmp As New ADODB.Recordset
   
   AddRecord = False
   
   strSO01 = textSO01
   strSO02 = DBDATE(textSO02)
   strSO03 = Trim(textSO03_1.Text & textSO03_2.Text) ' 2008/12/22 Add BY SINDY
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strSO01, strSO02, strSO03) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_overtime ("
   For nIndex = 0 To tf_so - 1
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
   For nIndex = 0 To tf_so - 1
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
   
'   'Add By Sindy 2018/3/2 當月的薪資已計算過,發E-MAIL通知財務處
'   strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSO02), 6) & " group by sm02"
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      strContent = GetEMailContent(textSO13, strSubject) '先取得E-Mail主旨,本文內容
'      PUB_SendMail strUserNum, "71005", "", textSO01 & GetPrjSalesNM(textSO01) & "補輸加班資料，請重新做每月(" & Left(DBDATE(textSO02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'   End If
'   '2018/3/2 END
   
   'Add By Sindy 2011/9/22
   If Frame1.Visible = True And textSO13 <> "" Then
      Call ProABSData
   Else
      'Add By Sindy 2019/5/23 假單完成,後續資料檢查及SendMail
      Call PUB_AutoM21Receive_SendMail(IIf(textSO13 <> "", textSO13, ""), 表單類別_加班, textSO01, DBDATE(textSO02), Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00")), _
         DBDATE(textSO02), "", Left(DBDATE(textSO02), 6), , , , m_EditMode)
   End If
   
   cnnConnection.CommitTrans
   ' 2008/12/22 Modify BY SINDY
   'If ((strSO01 & strSO02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSO01 & strSO02) > (m_LastKEY(0) & m_LastKEY(1))) Then
   If ((strSO01 & strSO02 & strSO03) < (m_FirstKEY(0) & m_FirstKEY(1) & m_FirstKEY(2))) Or ((strSO01 & strSO02 & strSO03) > (m_LastKEY(0) & m_LastKEY(1) & m_LastKEY(2))) Then
   ' 2008/12/22 END
      RefreshRange
   End If
   
   ' 2008/12/22 Modify BY SINDY
   'ShowCurrRecord strSO01, DBDATE(strSO02)
   ShowCurrRecord strSO01, DBDATE(strSO02), strSO03
   ' 2008/12/22 END
   
   Set rsTmp = Nothing
   AddRecord = True
   Exit Function
   
ErrHand:
   Set rsTmp = Nothing
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
   Dim strSO01 As String
   Dim strSO02 As String
   Dim strSO03 As String
   Dim strContent As String, strSubject As String
   Dim rsTmp As New ADODB.Recordset
   
   ModRecord = False
   
   strSO01 = m_CurrKEY(0)
   strSO02 = m_CurrKEY(1)
   strSO03 = m_CurrKEY(2) ' 2008/12/22 Add BY SINDY
   
   'Modify By Sindy 2023/11/1 mark,前面有檢查,此處應該不用
'   'Add By SINDY 2011/12/5
'   If strSO01 <> textSO01 Or _
'      strSO02 <> DBDATE(textSO02) Or _
'      Val(strSO03) <> Val(Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00"))) Then
'      ' 檢查記錄是否已存在
'      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00"))) = True Then
'         strTit = "新增資料"
'         strMsg = "該筆記錄已存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         'UpdateCtrlData
'         textSO02.SetFocus
'         Exit Function
'      End If
'   End If
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_overtime SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_so - 1
      strTmp = Empty
      'If nIndex < 6 Or nIndex > 10 Then
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
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = strSQL & " " & _
'                  "WHERE so01 = '" & strSO01 & "' and so02='" & strSO02 & "' ; end; "
   strSql = strSql & " " & _
                  "WHERE so01 = '" & strSO01 & "' and so02='" & strSO02 & "' and so03='" & strSO03 & "' ; end; "
   ' 2008/12/22 END
   
On Error GoTo ErrHand
      cnnConnection.BeginTrans
        If bDifference = True Then
           Pub_SeekTbLog strSql
           cnnConnection.Execute strSql
           
'            'Add By Sindy 2018/3/2 當月的薪資已計算過,發E-MAIL通知財務處
'            strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSO02), 6) & " group by sm02"
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strContent = GetEMailContent(textSO13, strSubject) '先取得E-Mail主旨,本文內容
'               PUB_SendMail strUserNum, "71005", "", textSO01 & GetPrjSalesNM(textSO01) & "修改加班資料，請重新做每月(" & Left(DBDATE(textSO02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'            End If
'            '2018/3/2 END
            
           'Add By Sindy 2011/9/22
           'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
           'If Frame1.Visible = True And textSO13 <> "" Then
               Call ProABSData
           'End If
        End If
        cnnConnection.CommitTrans
      
       ' 2008/12/22 Modify BY SINDY
      'ShowCurrRecord strSO01, DBDATE(strSO02)
      ShowCurrRecord strSO01, DBDATE(strSO02), strSO03
       ' 2008/12/22 END
   
   Set rsTmp = Nothing
   ModRecord = True
   Exit Function
   
ErrHand:
   Set rsTmp = Nothing
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strSO01 As String
   Dim strSO02 As String
   Dim strSO03 As String
   Dim strContent As String, strSubject As String
   Dim rsTmp As New ADODB.Recordset
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   strSO01 = m_CurrKEY(0)
   strSO02 = m_CurrKEY(1)
   strSO03 = m_CurrKEY(2) ' 2008/12/22 Add BY SINDY
   
   cnnConnection.BeginTrans
   
'   'Add By Sindy 2018/3/12 當月的薪資已計算過,發E-MAIL通知財務處
'   strSql = "select sm02,count(*) from SalaryMonth where sm02=" & Left(DBDATE(textSO02), 6) & " group by sm02"
'   intI = 1
'   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      strContent = GetEMailContent(textSO13, strSubject) '先取得E-Mail主旨,本文內容
'      PUB_SendMail strUserNum, "71005", "", textSO01 & GetPrjSalesNM(textSO01) & "刪除加班資料，請重新做每月(" & Left(DBDATE(textSO02), 6) & ")薪資計算！", strContent, , , , , , , , , , True
'   End If
'   '2018/3/12 END
   
   'Add By Sindy 2011/9/22
   'Modify By Sindy 2019/5/24 電子紙本均要考慮發信問題
'   If Frame1.Visible = True And textSO13.Text <> "" And m_B1019 <> "" Then
      PUB_FilterFormText Me 'Add by Sindy 2011/10/14 修正畫面所有含跳行符號的文字框
      'MsgBox "電子表單人事處已簽收,不可在此作業刪除！", vbExclamation
      'Exit Function
      Call DelMark
'   End If
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "DELETE FROM staff_overtime " & _
'            "WHERE so01 = '" & strSO01 & "'  and so02='" & strSO02 & "' "
   strSql = "DELETE FROM staff_overtime " & _
            "WHERE so01 = '" & strSO01 & "' and so02='" & strSO02 & "' and so03='" & strSO03 & "' "
   ' 2008/12/22 END
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   ' 2008/12/22 Modify BY SINDY
   'If (strSO01 = m_LastKEY(0) And strSO02 = m_LastKEY(1)) Or (strSO01 = m_FirstKEY(0) And strSO02 = m_FirstKEY(1)) Then
   If (strSO01 = m_LastKEY(0) And strSO02 = m_LastKEY(1) And strSO03 = m_LastKEY(2)) Or (strSO01 = m_FirstKEY(0) And strSO02 = m_FirstKEY(1) And strSO03 = m_FirstKEY(2)) Then
   ' 2008/12/22 END
      RefreshRange
   End If
   
   ' 2008/12/22 Modify BY SINDY
   'ShowCurrRecord strSO01, DBDATE(strSO02)
   ShowCurrRecord strSO01, DBDATE(strSO02), strSO03
   ' 2008/12/22 END
   
   DelRecord = True
   
   Set rsTmp = Nothing
   Exit Function
   
ErrHand:
    cnnConnection.RollbackTrans
    Set rsTmp = Nothing
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSO01 As String
   Dim strSO02 As String
   Dim strSO03 As String
   
   QueryRecord = False
   strSO01 = textSO01
   strSO02 = DBDATE(textSO02)
   strSO03 = Trim(textSO03_1.Text & textSO03_2.Text) ' 2008/12/22 Add BY SINDY
   
   If IsRecordExist(strSO01, strSO02, strSO03) = True Then
      m_CurrKEY(0) = strSO01
      m_CurrKEY(1) = strSO02
      m_CurrKEY(2) = strSO03 ' 2008/12/22 Add BY SINDY
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
         'Add By Sindy 2012/6/20
         If Trim(txtNote) = "" And Frame1.Visible = True Then
            MsgBox "刪除原因不可空白！", vbExclamation
            txtNote.SetFocus
            Exit Function
         End If
         '2012/6/20 End
         If DelRecord = True Then
            RefreshRange
            ClearField
            ShowCurrRecord m_CurrKEY(0), m_CurrKEY(1), m_CurrKEY(2)
         Else
            Exit Function
         End If
      Case 4: '查詢
         If textSO01 <> "" And textSO02 <> "" And _
            textSO03_1 <> "" And textSO03_2 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            ' 2008/12/17 ADD BY SINDY
            If textSO01 = "" Or textSO02 = "" Or _
               textSO03_1 = "" Or textSO03_2 = "" Then
               MsgBox "須輸入員工代號及日期和起始時間才可進行查詢動作！", vbInformation
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
      Case 1: If Me.Visible = True Then textSO01.SetFocus
      Case 2: If Me.Visible = True Then textSO03_1.SetFocus
      Case 4: If Me.Visible = True Then textSO01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String, ByVal strKEY03 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT * FROM staff_overtime " & _
'            "WHERE so01 = '" & strKEY01 & "'  and so02='" & strKEY02 & "'  "
   strSql = "SELECT so01 FROM staff_overtime " & _
            "WHERE so01='" & strKEY01 & "' and so02='" & strKEY02 & "' and so03='" & strKEY03 & "'"
   ' 2008/12/22 END
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   If m_EditMode = 2 Then
      strSql = strSql & " and so13<>'" & textSO13 & "'"
   End If
   '2023/11/1 END
   'Add By Sindy 2022/7/19
   'Modify By Sindy 2023/11/29 +已核准
   strSql = strSql & " union SELECT B1003 FROM abs010" & _
            " WHERE B1003 = '" & strKEY01 & "'" & _
            " and (" & strKEY02 & Right("0000" & strKEY03, 4) & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
            " or " & strKEY02 & Format("0" & textSO04_1, "00") & Format("0" & textSO04_2, "00") & " between B1004||substr('0'||B1005,-4) and B1006||substr('0'||B1007,-4)" & _
            ") and B1018 not in('" & 註銷 & "','" & 已核准 & "')"
   '2022/7/19 END
   'Add By Sindy 2023/11/1 修改時,排除檢查此表單
   'Modify By Sindy 2025/1/6 新增時,排除檢查此表單
   'Modify By Sindy 2025/1/21 + And Frame1.Visible = True
   If (m_EditMode = 2 Or m_EditMode = 1) And Frame1.Visible = True Then
      strSql = strSql & " and B1001<>'" & textSO13 & "'"
   End If
   '2023/11/1 END
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
      m_CurrKEY(2) = strKEY03 ' 2008/12/22 Add BY SINDY
   Else
      ' 2008/12/22 Modify BY SINDY
'      strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'               "WHERE so01 = '" & m_CurrKEY(0) & "' and so02='" & m_CurrKEY(1) & "' "
      strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
               "WHERE so01 = '" & m_CurrKEY(0) & "' and so02='" & m_CurrKEY(1) & "' and so03='" & m_CurrKEY(2) & "' "
      ' 2008/12/22 END
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
         If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
         ' 2008/12/22 Add BY SINDY
         If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
         ' 2008/12/22 END
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      ' 2008/12/22 Modify BY SINDY
'      strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'               "WHERE so02 = (SELECT MIN(so02) FROM staff_overtime where so01=(select min(so01) from staff_overtime) ) and so01=(select min(so01) from staff_overtime) "
      strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
               "WHERE so02 = (SELECT MIN(so02) FROM staff_overtime where so01=(select min(so01) from staff_overtime) ) and so01=(select min(so01) from staff_overtime) Order BY so03 ASC "
      ' 2008/12/22 END
      
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
         If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
         ' 2008/12/22 Add BY SINDY
         If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
         ' 2008/12/22 END
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
   m_CurrKEY(2) = m_FirstKEY(2) ' 2008/12/22 Add BY SINDY
   UpdateCtrlData
End Sub

' 顯示上一筆資料
Private Sub ShowPrevRecord()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 2008/12/22 Modify BY SINDY
   If m_CurrKEY(0) = m_FirstKEY(0) And m_CurrKEY(1) = m_FirstKEY(1) And m_CurrKEY(2) = m_FirstKEY(2) Then
   ' 2008/12/22 END
      ShowMsg MsgText(9008)
      GoTo EXITSUB
   End If
   
   ' 2008/12/22 Add BY SINDY
   strSql = "SELECT So01,So02,So03 FROM staff_overtime " & _
            "WHERE So01 = '" & m_CurrKEY(0) & "' AND " & _
                          "So02 = '" & m_CurrKEY(1) & "' AND " & _
                  "So03 = (SELECT MAX(So03) FROM staff_overtime " & _
                          "WHERE So01 = '" & m_CurrKEY(0) & "' AND " & _
                                        "So02 = '" & m_CurrKEY(1) & "' AND " & _
                                        "So03 < '" & m_CurrKEY(2) & "' ) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
      If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   ' 2008/12/22 END
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
'                          "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "so02 < '" & m_CurrKEY(1) & "' )"
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
                          "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
                                "so02 < '" & m_CurrKEY(1) & "' ) Order By so03 DESC "
   ' 2008/12/22 END
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime " & _
'                           "WHERE so01 < '" & m_CurrKEY(0) & "') AND " & _
'                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
'                           "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime " & _
'                                          "WHERE so01 < '" & m_CurrKEY(0) & "')) "
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime " & _
                           "WHERE so01 < '" & m_CurrKEY(0) & "') AND " & _
                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
                           "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime " & _
                                          "WHERE so01 < '" & m_CurrKEY(0) & "')) Order BY so03 DESC "
   ' 2008/12/22 END
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
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
   
   ' 2008/12/22 Modify BY SINDY
   If m_CurrKEY(0) = m_LastKEY(0) And m_CurrKEY(1) = m_LastKEY(1) And m_CurrKEY(2) = m_LastKEY(2) Then
   ' 2008/12/22 END
      ShowMsg MsgText(9009)
      GoTo EXITSUB
   End If
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
'                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
'                          "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
'                                "so02 > '" & m_CurrKEY(1) & "' )"
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
                          "WHERE so01 = '" & m_CurrKEY(0) & "' AND " & _
                                "so02 > '" & m_CurrKEY(1) & "' ) Order by so03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime " & _
'                           "WHERE so01 > '" & m_CurrKEY(0) & "') AND " & _
'                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
'                           "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime " & _
'                                          "WHERE so01 > '" & m_CurrKEY(0) & "')) "
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime " & _
                           "WHERE so01 > '" & m_CurrKEY(0) & "') AND " & _
                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
                           "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime " & _
                                          "WHERE so01 > '" & m_CurrKEY(0) & "')) Order BY so03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_CurrKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
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
   m_CurrKEY(2) = m_LastKEY(2) ' 2008/12/22 Add BY SINDY
   UpdateCtrlData
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   m_SubMode = 0
   m_KeyCode = KeyCode 'Add By Sindy 2011/10/7
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
         'Add By Sindy 2013/2/1
         Me.SSTab1.TabEnabled(1) = False
         SSTab1.Tab = 0
         '2013/2/1 End
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
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime) AND " & _
'                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
'                           "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime)) "
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime) AND " & _
                  "so02 = (SELECT MIN(so02) FROM staff_overtime " & _
                           "WHERE so01 = (SELECT MIN(so01) FROM staff_overtime)) Order BY so03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_FirstKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
   End If
   rsTmp.Close
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT so01,so02 FROM staff_overtime " & _
'            "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime) AND " & _
'                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
'                           "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime)) "
   strSql = "SELECT so01,so02,so03 FROM staff_overtime " & _
            "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime) AND " & _
                  "so02 = (SELECT MAX(so02) FROM staff_overtime " & _
                           "WHERE so01 = (SELECT MAX(so01) FROM staff_overtime)) Order BY so03 ASC "
   ' 2008/12/22 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("so01")) = False Then: m_LastKEY(0) = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: m_LastKEY(1) = rsTmp.Fields("so02")
      ' 2008/12/22 Add BY SINDY
      If IsNull(rsTmp.Fields("so03")) = False Then: m_LastKEY(2) = rsTmp.Fields("so03")
      ' 2008/12/22 END
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   
   ' 2008/12/22 Modify BY SINDY
'   strSQL = "SELECT * FROM staff_overtime " & _
'            "WHERE so01='" & m_CurrKEY(0) & "' and so02 = '" & m_CurrKEY(1) & "'   "
   strSql = "SELECT * FROM staff_overtime " & _
            "WHERE so01='" & m_CurrKEY(0) & "' and so02 = '" & m_CurrKEY(1) & "' and so03 = '" & m_CurrKEY(2) & "' "
   ' 2008/12/22 END
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("so01")) = False Then: textSO01 = rsTmp.Fields("so01")
      If IsNull(rsTmp.Fields("so02")) = False Then: textSO02 = TAIWANDATE(rsTmp.Fields("so02"))
      If IsNull(rsTmp.Fields("so03")) = False Then: textSO03_1 = Mid(Format(rsTmp.Fields("so03"), "0000"), 1, 2): textSO03_2 = Mid(Format(rsTmp.Fields("so03"), "0000"), 3, 2)
      If IsNull(rsTmp.Fields("so04")) = False Then: textSO04_1 = Mid(Format(rsTmp.Fields("so04"), "0000"), 1, 2): textSO04_2 = Mid(Format(rsTmp.Fields("so04"), "0000"), 3, 2)
'      If IsNull(rsTmp.Fields("so05")) = False Then: textSO05 = rsTmp.Fields("so05")
'      ' 2008/12/18 Add BY SINDY
'      If IsNull(rsTmp.Fields("so06")) = False Then: textSO06 = rsTmp.Fields("so06")
'      ' 2008/12/18 END
      'Add By Sindy 2016/12/26
      If Not IsNull(rsTmp.Fields("So05")) Then
         Label3.Caption = "平日-共                        時"
         textSO0506.Text = rsTmp.Fields("So05")
         m_B1012 = rsTmp.Fields("So05") 'Add By Sindy 2019/5/24
      ElseIf Not IsNull(rsTmp.Fields("So06")) Then
         Label3.Caption = "假日-共                        時"
         textSO0506.Text = rsTmp.Fields("So06")
         m_B1013 = rsTmp.Fields("So06") 'Add By Sindy 2019/5/24
      End If
      If Not IsNull(rsTmp.Fields("So15")) Then
         textSO15 = rsTmp.Fields("So15")
      Else
         textSO15 = textSO0506
      End If
      m_B1030 = textSO15 'Add By Sindy 2019/5/24
      '2016/12/26 END
      'Add By Sindy 2011/9/22
      If IsNull(rsTmp.Fields("so13")) = False Then
         textSO13 = rsTmp.Fields("so13")
         Call GetABS010(True)
         Frame1.Visible = True
      Else
         'Add By Sindy 2019/5/24 記錄原始資料
         m_B1004 = DBDATE(textSO02)
         m_B1005 = textSO03_1 & textSO03_2
         m_B1007 = textSO04_1 & textSO04_2
         '2019/5/24 END
         Frame1.Visible = False
      End If
      
      'Add By Sindy 2012/5/23
      m_SO02 = TAIWANDATE(rsTmp.Fields("so02"))
      m_SO03_1 = Mid(Format(rsTmp.Fields("so03"), "0000"), 1, 2)
      m_SO03_2 = Mid(Format(rsTmp.Fields("so03"), "0000"), 3, 2)
      m_SO04_1 = Mid(Format(rsTmp.Fields("so04"), "0000"), 1, 2)
      m_SO04_2 = Mid(Format(rsTmp.Fields("so04"), "0000"), 3, 2)
      '2012/5/23 End
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp

       textSO01_2 = GetStaffName(textSO01, True)
   End If

   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
strSql = ""
If txt1(0) <> "" Then
    strSql = strSql & " and so01>='" & txt1(0) & "' "
End If
If txt1(1) <> "" Then
    strSql = strSql & " and so01<='" & txt1(1) & "' "
End If
If txt1(2) <> "" Then
    strSql = strSql & " and so02>='" & DBDATE(txt1(2)) & "' "
End If
If txt1(3) <> "" Then
    strSql = strSql & " and so02<='" & DBDATE(txt1(3)) & "' "
End If
' 2008/12/18 Modify BY SINDY
'抓取資料
'strSQL = "SELECT so01,st02,sqldateT(so02),substr(ltrim(to_char('0000'||to_char(so03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(so03),'0000')),3,2),substr(ltrim(to_char('0000'||to_char(so04),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(so04),'0000')),3,2),so05  FROM staff_overtime,staff where so01=st01(+)  " & strSQL & _
'        " order by so01,so02 "
'Modify By Sindy 2016/12/27
strSql = "SELECT so01,st02,sqldateT(so02),substr(ltrim(to_char('0000'||to_char(so03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(so03),'0000')),3,2),substr(ltrim(to_char('0000'||to_char(so04),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(so04),'0000')),3,2),so05,so06,so15 FROM staff_overtime,staff where so01=st01(+)  " & strSql & _
        " order by so02,so03,so01 "
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
   
   nResponse = False
   textSO01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSO02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textso03_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textso03_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSo04_1_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSo04_2_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSo15_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
'   ' 2008/12/18 Add BY SINDY
'   nResponse = False
'   textSo06_Validate nResponse
'   If nResponse = True Then GoTo EXITSUB
'   ' 2008/12/18 END
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSO01.Locked = bEnable
   If bEnable Then textSO01.BackColor = &H8000000F Else textSO01.BackColor = &H80000005
   If m_EditMode <> "2" Then 'Modify By Sindy 2011/12/5
      textSO02.Locked = bEnable
      If bEnable Then textSO02.BackColor = &H8000000F Else textSO02.BackColor = &H80000005
      ' 2008/12/22 Add BY SINDY
      textSO03_1.Locked = bEnable
      textSO03_2.Locked = bEnable
      If bEnable Then textSO03_1.BackColor = &H8000000F Else textSO03_1.BackColor = &H80000005
      If bEnable Then textSO03_2.BackColor = &H8000000F Else textSO03_2.BackColor = &H80000005
      ' 2008/12/22 END
   End If
   'Add By Sindy 2011/9/22
   textSO13.Locked = bEnable
   If bEnable Then textSO13.BackColor = &H8000000F Else textSO13.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textSO01.Locked = bEnable
   If bEnable Then textSO01.BackColor = &H8000000F Else textSO01.BackColor = &H80000005
   textSO02.Locked = bEnable
   If bEnable Then textSO02.BackColor = &H8000000F Else textSO02.BackColor = &H80000005
   textSO03_1.Locked = bEnable
   textSO03_2.Locked = bEnable
   ' 2008/12/22 Add BY SINDY
   If bEnable Then textSO03_1.BackColor = &H8000000F Else textSO03_1.BackColor = &H80000005
   If bEnable Then textSO03_2.BackColor = &H8000000F Else textSO03_2.BackColor = &H80000005
   ' 2008/12/22 END
   textSO04_1.Locked = bEnable
   textSO04_2.Locked = bEnable
   textSO15.Locked = bEnable
'   ' 2008/12/18 Add BY SINDY
'   textSO06.Locked = bEnable
'   ' 2008/12/18 END
   'Add By Sindy 2011/9/22
   textSO13.Locked = bEnable
   If bEnable Then textSO13.BackColor = &H8000000F Else textSO13.BackColor = &H80000005
'   txtNote.Locked = bEnable
   
   'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'   '有表單編號的資料時,時數欄位鎖住
'   If textSO13 <> "" Then
'      textSO05.Enabled = False
'      textSO06.Enabled = False
'   Else
'      textSO05.Enabled = True
'      textSO06.Enabled = True
'   End If
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   textSO01 = Empty
   textSO01_2 = Empty
   textSO02 = Empty
   textSO03_1 = Empty
   textSO03_2 = Empty
   textSO04_1 = Empty
   textSO04_2 = Empty
   textSO15 = Empty
   ' 2008/12/18 Add BY SINDY
   textSO0506 = Empty
   ' 2008/12/18 END
   
   'Add By Sindy 2011/9/22
   textSO13 = Empty
   Frame1.Visible = False
   m_B1019 = Empty: m_B1004 = Empty: m_B1005 = Empty: m_B1007 = Empty
   m_B1012 = Empty: m_B1013 = Empty: m_B1017 = Empty: m_B1030 = Empty
   txtNote = Empty
   
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_so - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   
   cmdABS.Visible = False 'Add By Sindy 2022/10/28
End Sub

Private Sub UpdateFieldNewData()
    Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SO01", textSO01
      'Modify By Sindy 2011/12/5 起始日期及起始時間開放修改
'      SetFieldNewData "SO02", DBDATE(textSO02)
'      SetFieldNewData "SO03", textSO03_1 & Format("00" & textSO03_2, "00")
   End If
   SetFieldNewData "SO02", DBDATE(textSO02)
   SetFieldNewData "SO03", textSO03_1 & Format("00" & textSO03_2, "00")
   SetFieldNewData "SO04", textSO04_1 & Format("00" & textSO04_2, "00")
   'Modify By Sindy 2016/12/27
   If ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Then '假日
      SetFieldNewData "SO05", ""
      SetFieldNewData "SO06", textSO0506
   Else '平日
      SetFieldNewData "SO05", textSO0506
      SetFieldNewData "SO06", ""
   End If
   '2016/12/27 END
   'SetFieldNewData "SO05", textSO05
   ' 2008/12/18 Add BY SINDY
   'SetFieldNewData "SO06", textSO06
   ' 2008/12/18 END
   'Add By Sindy 2011/9/22
   SetFieldNewData "SO13", textSO13
   SetFieldNewData "SO15", textSO15 'Add By Sindy 2016/12/27
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To tf_so
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SO" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2, 3, 4, 5:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
SetGrd
End Sub

Private Sub textSO01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO01
End If
End Sub

Private Sub textSO01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

'Add By Sindy 2011/9/22
Private Sub textSO01_LostFocus()
   '若輸入的員工代號為可寄信者,必須輸入表單編號
   If Frame1.Visible = True Then If textSO13.Enabled = True Then textSO13.SetFocus
   '新增狀態將游標停在員工代號的欄位
   If m_EditMode = 1 And textSO01 = "" Then textSO01.SetFocus
End Sub

Private Sub textSO01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

If textSO01.Text = "" Then
   textSO01_2 = "" ' 2008/12/18 ADD BY SINDY
   'Add By Sindy 2011/9/22 預設值
   Frame1.Visible = False
End If

If m_EditMode <> 0 And textSO01 <> "" Then
    textSO01_2 = GetStaffName(textSO01, True)
    ' 2008/12/18 ADD BY SINDY
    ' 檢查員工編號規則
    If ChkStaffID(textSO01) Then
       Call textSO01_GotFocus
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
    If textSO01_2 = "" Then
        MsgBox "員工編號錯誤！查無此員工！", vbInformation
        Call textSO01_GotFocus ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    If m_KeyCode = vbKeyF2 Then '按新增時
      'Add By Sindy 2011/9/22 檢查此員工是否為"不寄信"
      If ChkStaffST14(textSO01, False) = False Then
        strTit = "詢問"
        strMsg = "是否要補電子表單？" & vbCrLf & vbCrLf & _
                 "（注意：要補簽核流程，請先輸入表單編號）"
        nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
        If nResponse = vbYes Then
           '電子表單請假
           Frame1.Visible = True
        Else
           '紙本請假
           Frame1.Visible = False
        End If
      Else
        '不寄信,紙本請假
        Frame1.Visible = False
      End If
    End If
End If

If m_EditMode = 1 And textSO01 <> "" And Val(textSO03_1) > 0 And Val(textSO03_2) > 0 Then
    If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO01.Enabled = True And textSO01.Locked = False Then
        MsgBox "該員工當天已有資料，請修改！", vbInformation
        Call textSO01_GotFocus ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textSO02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO02
    CloseIme
End If
End Sub

Private Sub textSO02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSO02_Validate(Cancel As Boolean)
    If m_EditMode <> 0 And textSO02 <> "" Then
      If CheckIsTaiwanDate(textSO02, False) = False Then
         Call textSO02_GotFocus ' 2008/12/18 ADD BY SINDY
         Cancel = True
         MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
         Exit Sub
      End If
      'Modify By Sindy 2016/12/27
      If ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Then '假日
         Label3.Caption = "假日-共                        時"
      Else '平日
         Label3.Caption = "平日-共                        時"
      End If
      '2016/12/27 END
      
      'Add By Sindy 2025/11/3 下班逾30分鐘原因確認非處理公務時,不可填寫加班單
      strSql = "select * from abs015 " & _
                "where B1501='" & textSO01 & "' and B1502=" & DBDATE(textSO02) & _
                  "and B1504<>'2'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         MsgBox "此日期已確認下班逾30分鐘原因為 非處理公務，因此不可填寫加班單！請洽人事處。", vbInformation, "輸入日期錯誤"
         Call textSO02_GotFocus
         Cancel = True
         Exit Sub
      End If
      '2025/11/3 END
      
      'Add By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
      If m_EditMode = 1 Or m_EditMode = 2 Then
         If textSO15 = "" Then 'Add By Sindy 2017/6/19 實際時數空白時才要計算
            Call CountHour
         End If
         'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'         If Frame1.Visible = True And textSO13 <> "" Then
'            '有表單編號的資料時,時數欄位鎖住
'            textSO05.Enabled = False
'            textSO06.Enabled = False
'         End If
      End If
    End If
    
    If m_EditMode = 1 And textSO02 <> "" And Val(textSO03_1) > 0 And Val(textSO03_2) > 0 Then
      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO02.Enabled = True And textSO02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSO02_GotFocus ' 2008/12/18 ADD BY SINDY
         Cancel = True
         Exit Sub
      End If
    End If
End Sub

Private Sub textso03_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO03_1
    CloseIme
End If
End Sub

Private Sub textso03_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textso03_1_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSO03_1 <> "" Then
    If CheckLengthIsOK(textSO03_1, textSO03_1.MaxLength) = False Then
        Call textso03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/18 ADD BY SINDY
'    If ChkWorkDay(ChangeTStringToWString(textSO02)) = True Then
'      If textSO03_1.Text < 17 Then
'         Call textso03_1_GotFocus
'         MsgBox "工作天起始(時)不可小於17時!!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'    Else
'      If textSO03_1.Text < 7 Then
'         Call textso03_1_GotFocus
'         MsgBox "假日天起始(時)不可小於7時!!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'    End If
    If textSO03_1.Text > 24 Then
       Call textso03_1_GotFocus
       MsgBox "不可超過24時!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If m_EditMode = 1 And textSO02 <> "" And Val(textSO03_1) > 0 And Val(textSO03_2) > 0 Then
      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO02.Enabled = True And textSO02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textso03_1_GotFocus
         Cancel = True
         Exit Sub
      End If
    End If
    ' 2008/12/18 END
    
    If textSO03_1 <> "" And textSO04_1 <> "" Then
      If RunNick(textSO03_1, textSO04_1) Then
          Call textso03_1_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
   'Add By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textSO15 = "" Then 'Add By Sindy 2017/6/19 實際時數空白時才要計算
         Call CountHour
      End If
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      If Frame1.Visible = True And textSO13 <> "" Then
'         '有表單編號的資料時,時數欄位鎖住
'         textSO05.Enabled = False
'         textSO06.Enabled = False
'      End If
   End If
End If
CloseIme
End Sub

Private Sub textso03_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO03_2
    CloseIme
End If
End Sub

Private Sub textso03_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textso03_2_Validate(Cancel As Boolean)
If textSO03_2 = "" Then textSO03_2 = "00"

If m_EditMode <> 0 And textSO03_2 <> "" Then
    If CheckLengthIsOK(textSO03_2, textSO03_2.MaxLength) = False Then
        Call textso03_2_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/18 ADD BY SINDY
    If textSO03_2.Text > 59 Then
       Call textso03_2_GotFocus
       MsgBox "不可超過59分!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    If m_EditMode = 1 And textSO02 <> "" And Val(textSO03_1) > 0 And Val(textSO03_2) > 0 Then
      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO02.Enabled = True And textSO02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textso03_2_GotFocus
         Cancel = True
         Exit Sub
      End If
    End If
    ' 2008/12/18 END
   'Add By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textSO15 = "" Then 'Add By Sindy 2017/6/19 實際時數空白時才要計算
         Call CountHour
      End If
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      If Frame1.Visible = True And textSO13 <> "" Then
'         '有表單編號的資料時,時數欄位鎖住
'         textSO05.Enabled = False
'         textSO06.Enabled = False
'      End If
   End If
End If

CloseIme
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("員工編號", "姓名", "日期", "加班起", "加班迄", "平日時數", "假日時數")
   arrGridHeadWidth = Array(800, 1200, 1200, 800, 800, 800, 800)
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

Private Sub textSo04_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO04_1
    textSO04_1.SetFocus ' 2008/12/18 ADD BY SINDY
    CloseIme
End If
End Sub

Private Sub textSo04_1_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSo04_1_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSO04_1 <> "" Then
    If CheckLengthIsOK(textSO04_1, textSO04_1.MaxLength) = False Then
        Call textSo04_1_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/18 ADD BY SINDY
    If textSO04_1.Text > 24 Then
       Call textSo04_1_GotFocus
       MsgBox "不可超過24時!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
    
    If textSO03_1 <> "" And textSO04_1 <> "" Then
      If RunNick(textSO03_1, textSO04_1) Then
          Call textSo04_1_GotFocus   ' 2008/12/18 ADD BY SINDY
          Cancel = True
          Exit Sub
      End If
    End If
    
    'Add By Sindy 2022/7/19
    If m_EditMode = 1 And textSO02 <> "" And Val(textSO04_1) > 0 And Val(textSO04_2) > 0 Then
      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO02.Enabled = True And textSO02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSo04_1_GotFocus
         Cancel = True
         Exit Sub
      End If
    End If
    '2022/7/19 END
    
    'Add By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textSO15 = "" Then 'Add By Sindy 2017/6/19 實際時數空白時才要計算
         Call CountHour
      End If
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      If Frame1.Visible = True And textSO13 <> "" Then
'         '有表單編號的資料時,時數欄位鎖住
'         textSO05.Enabled = False
'         textSO06.Enabled = False
'      End If
   End If
End If
CloseIme
End Sub

Private Sub textSo04_2_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO04_2
    textSO04_2.SetFocus ' 2008/12/18 ADD BY SINDY
    CloseIme
End If
End Sub

Private Sub textSo04_2_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSo04_2_Validate(Cancel As Boolean)
If textSO04_2 = "" Then textSO04_2 = "00"

If m_EditMode <> 0 And textSO04_2 <> "" Then
    If CheckLengthIsOK(textSO04_2, textSO04_2.MaxLength) = False Then
        Call textSo04_2_GotFocus   ' 2008/12/18 ADD BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/18 ADD BY SINDY
    If textSO04_2.Text > 59 Then
       Call textSo04_2_GotFocus
       MsgBox "不可超過59分!", vbExclamation + vbOKOnly
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
    
    'Add By Sindy 2022/7/19
    If m_EditMode = 1 And textSO02 <> "" And Val(textSO04_1) > 0 And Val(textSO04_2) > 0 Then
      If IsRecordExist(textSO01, DBDATE(textSO02), Trim(textSO03_1.Text & textSO03_2.Text)) = True And textSO02.Enabled = True And textSO02.Locked = False Then
         MsgBox "該員工當天已有資料，請修改！", vbInformation
         Call textSo04_2_GotFocus
         Cancel = True
         Exit Sub
      End If
    End If
    '2022/7/19 END
    
   'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If textSO15 = "" Then 'Add By Sindy 2017/6/19 實際時數空白時才要計算
         Call CountHour
      End If
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      If Frame1.Visible = True And textSO13 <> "" Then
'         '有表單編號的資料時,時數欄位鎖住
'         textSO05.Enabled = False
'         textSO06.Enabled = False
'      End If
   End If
End If

CloseIme
End Sub

'Add By Sindy 2016/12/27
Private Sub textSo15_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO15
    CloseIme
End If
End Sub
Private Sub textSo15_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub
Private Sub textSo15_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSO15 <> "" Then
   If CheckLengthIsOK(textSO15, textSO15.MaxLength) = False Then
      Call textSo15_GotFocus
      Cancel = True
      Exit Sub
   Else
      'Add By Sindy 2017/12/4
      If DBDATE(textSO02) >= 20161223 Then '(2016/12/23開始實施)
         'Modify By Sindy 2017/1/3 換算加班時數
         textSO0506 = PUB_Overtime_TransDay(textSO02, textSO01, textSO15)
      '原加班多少時數,就是多少時數
      Else
         textSO0506 = textSO15
      End If
      '2017/12/4
   End If
End If
CloseIme
End Sub

'Private Sub textSo06_GotFocus()
'If m_EditMode <> 0 Then
'    InverseTextBox textSO06
'    CloseIme
'End If
'End Sub
'
'Private Sub textSo06_KeyPress(KeyAscii As Integer)
'KeyAscii = Pub_NumAscii(KeyAscii, True)
'End Sub
'
'Private Sub textSo06_Validate(Cancel As Boolean)
'If m_EditMode <> 0 And textSO06 <> "" Then
'    If CheckLengthIsOK(textSO06, textSO06.MaxLength) = False Then
'        Call textSo06_GotFocus   ' 2008/12/18 ADD BY SINDY
'        Cancel = True
'        Exit Sub
'    End If
'End If
'CloseIme
'End Sub

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

'計算時數
Private Function CountHour()
'Dim calTime
Dim dHour As Double
Dim strSo0506 As String
   
   If textSO02.Text <> "" And textSO03_1.Text <> "" And _
         textSO03_2.Text <> "" And textSO04_1.Text <> "" And _
         textSO04_2.Text <> "" Then
      'Modify By Sindy 2012/5/23
      'If textSO02.Text <> "" Then
      If textSO02.Text <> "" And _
         (textSO15 = "" Or ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Or _
            (Val(m_SO02) <> Val(textSO02.Text) Or _
            Val(m_SO03_1) <> Val(textSO03_1.Text) Or _
            Val(m_SO03_2) <> Val(textSO03_2.Text) Or _
            Val(m_SO04_1) <> Val(textSO04_1.Text) Or _
            Val(m_SO04_2) <> Val(textSO04_2.Text))) Then
      '2012/5/23 End
         'TimeSerial : 此函數不可使用, 因它會依各PC的時間設定格式有其不同的結果
         'calTime = Trim(TimeSerial(Val(textSO04_1) - Val(textSO03_1), Val(textSO04_2) - Val(textSO03_2), 0))
         'dHour = Val(Mid(calTime, 1, 2)) + ((Val(Mid(calTime, 4, 2)) \ 30) * 0.5)
         
         '以半小時為單位
         'dHour = ((((Val(textSO04_1) * 60) + Val(textSO04_2)) - ((Val(textSO03_1) * 60) + Val(textSO03_2))) \ 30) * 0.5
         '以分鐘為單位, 取至小數第一位, 四捨五入
'         dHour = Round((((Val(textSO04_1) * 60) + Val(textSO04_2)) - ((Val(textSO03_1) * 60) + Val(textSO03_2))) / 60, 1)
'         If dHour < 0 Then dHour = 0
         dHour = PUB_CountHour_Overtime(textSO02, textSO01, textSO04_1, textSO04_2, textSO03_1, textSO03_2, strSo0506)
         textSO15.Text = dHour
         textSO0506.Text = strSo0506
''         'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
''         If textSO15.Text = "" Or textSO15.Text <= "0" Or (Frame1.Visible = True And textSO13 <> "") Then
''            textSO15.Text = dHour
''         End If
'         If ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Then '假日加班
'            '非週六的加班計算
'            If Weekday(Format(ChangeTStringToWString(textSO02), "####-##-##")) <> 7 Then
'               If Val(textSO15) <= 8 Then
'                  textSO0506.Text = 8
'               Else
'                  textSO0506.Text = textSO15
'               End If
'            '週六的加班計算
'            Else
'               If Val(textSO15) <= 4 Then
'                  textSO0506.Text = 4
'               ElseIf Val(textSO15) <= 8 Then
'                  textSO0506.Text = 8
'               ElseIf Val(textSO15) <= 12 Then
'                  textSO0506.Text = 12
'               Else
'                  textSO0506.Text = textSO15
'               End If
'            End If
'         Else '平日加班(工作日加班沒有換算的問題)
'            textSO0506.Text = textSO15
'         End If
'         '假日時數
'         'Modify By Sindy 2012/8/15 增加檢查颱風假
'         'If ChkWorkDay(ChangeTStringToWString(textSO02)) = False Then
'         If ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Then
'         '2012/8/15 End
'             'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
'             'If textSO06.Text = "" Or textSO06.Text <= "0" Then
'             If textSO06.Text = "" Or textSO06.Text <= "0" Or (Frame1.Visible = True And textSO13 <> "") Then
'               textSO05.Text = ""
'               textSO05.Enabled = False
'               textSO06.Text = dHour
'               textSO06.Enabled = True
'             End If
'         Else '平日時數
'             'Modify By Sindy 2011/9/22 電子表單資料,日及時欄位值一律由系統計算
'             If textSO05.Text = "" Or textSO05.Text <= "0" Then 'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'             'If textSO05.Text = "" Or textSO05.Text <= "0" Or (Frame1.Visible = True And textSO13 <> "") Then
'               textSO05.Text = dHour
'               textSO05.Enabled = True
'               textSO06.Text = ""
'               textSO06.Enabled = False
'             End If
'         End If
      'Add By Sindy 2016/12/27
      ElseIf ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = True Then '工作日加班沒有換算的問題
         textSO0506.Text = textSO15
      '2016/12/27 END
      End If
   End If
End Function

Private Sub textSO13_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSO13
    CloseIme
End If
End Sub

Private Sub textSO13_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSO13_LostFocus()
   '新增狀態時可以輸入表單編號做查詢
   If m_EditMode = 1 And textSO13 <> "" Then
      If GetABS010 = True Then
         textSO13.Enabled = False
      End If
   End If
End Sub

Private Sub textSO13_Validate(Cancel As Boolean)
   If Frame1.Visible = False Then Exit Sub
   
   If m_EditMode = 1 And textSO13 <> "" Then
      If CheckLengthIsOK(textSO13, textSO13.MaxLength) = False Then
         Call textSO13_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkAbsSysB1001Exist(textSO13, "02", textSO01) = False Then
         Call textSO13_GotFocus
         Cancel = True
         Exit Sub
      End If
      If ChkPerSysB1001Exist(textSO13, textSO01, False) = True Then
         MsgBox "表單編號重覆！", vbExclamation
         Call textSO13_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub

'Add By Sindy 2011/9/22
Private Function GetABS010(Optional bolOnlyQrySETime As Boolean = False) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, i As Integer
   
   Screen.MousePointer = vbHourglass
   GetABS010 = False
   cmdABS.Visible = False 'Add By Sindy 2022/10/28
   
   '出缺勤電子簽核主檔
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,substr(ltrim(to_char('0000'||to_char(B1028),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1028),'0000')),3,2) B1028,substr(ltrim(to_char('0000'||to_char(B1029),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1029),'0000')),3,2) B1029,B1030 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) " & _
            "and B1001='" & textSO13 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      GetABS010 = True
      'Modify By Sindy 2011/10/13 人事不簽收,事後發現有問題都在此作業修改
'      '有表單編號的資料時,時數欄位鎖住
'      textSO05.Enabled = False
'      textSO06.Enabled = False
      
      '記錄原始資料 : 註.m_變數值必須在ClearField函數裡清值
      If Not IsNull(rsTmp.Fields("B1019")) Then m_B1019 = rsTmp.Fields("B1019")
      If Not IsNull(rsTmp.Fields("B1004")) Then m_B1004 = rsTmp.Fields("B1004")
      If Not IsNull(rsTmp.Fields("B1005")) Then m_B1005 = IIf(Format(rsTmp.Fields("B1005"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1005"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1007")) Then m_B1007 = IIf(Format(rsTmp.Fields("B1007"), "hhmm") = "0000", "", Format(rsTmp.Fields("B1007"), "hhmm"))
      If Not IsNull(rsTmp.Fields("B1012")) Then m_B1012 = rsTmp.Fields("B1012")
      If Not IsNull(rsTmp.Fields("B1013")) Then m_B1013 = rsTmp.Fields("B1013")
      If Not IsNull(rsTmp.Fields("B1030")) Then m_B1030 = rsTmp.Fields("B1030") 'Add By Sindy 2016/12/27
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      
      '顯示其他資料至畫面上
      'Add By Sindy 2022/10/28 + if 已簽核不要顯示於畫面上,已人事資料為主
      If Not IsNull(rsTmp.Fields("B1019")) Then
         '為防止簽核後又修改,抓人事資料
         strSql = "select * from abs012 where b1201='" & textSO13 & "' and substr(b1207,1,4)='修改資料'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            cmdABS.Visible = True
            '記錄畫面上的資料 : 註.m_變數值必須在ClearField函數裡清值
            m_B1004 = DBDATE(textSO02)
            m_B1005 = Format(textSO03_1 & textSO03_2, "0000")
            m_B1007 = Format(textSO04_1 & textSO04_2, "0000")
            If Not IsNull(rsTmp.Fields("B1012")) Then m_B1012 = textSO0506
            If Not IsNull(rsTmp.Fields("B1013")) Then m_B1013 = textSO0506
            m_B1030 = textSO15
         End If
      Else
      '2022/10/28 END
         If Not IsNull(rsTmp.Fields("B1004")) Then textSO02 = ChangeWStringToTString(rsTmp.Fields("B1004"))
         If Not IsNull(rsTmp.Fields("B1005")) Then textSO03_1 = Left(rsTmp.Fields("B1005"), 2): textSO03_2 = Right(rsTmp.Fields("B1005"), 2)
         If Not IsNull(rsTmp.Fields("B1007")) Then textSO04_1 = Left(rsTmp.Fields("B1007"), 2): textSO04_2 = Right(rsTmp.Fields("B1007"), 2)
         If Not IsNull(rsTmp.Fields("B1012")) Then textSO0506 = rsTmp.Fields("B1012")
         If Not IsNull(rsTmp.Fields("B1013")) Then textSO0506 = rsTmp.Fields("B1013")
         If Not IsNull(rsTmp.Fields("B1030")) Then textSO15 = rsTmp.Fields("B1030")
      End If
      
      '後面資料不顯示至畫面上
      If bolOnlyQrySETime = True Then
         GoTo EXITSUB
      End If
      
   Else
'      Screen.MousePointer = vbDefault
'      ShowNoData
'      rsTmp.Close
'      Set rsTmp = Nothing
'      Exit Sub
   End If
   
EXITSUB:
   rsTmp.Close
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

'Add By Sindy 2011/9/22
Private Sub ProABSData()
Dim strUpdDate As String, strUpdTime As String
Dim strB1004 As String, strB1005 As String
Dim strB1007 As String, strB1012 As String, strB1013 As String
Dim strOldData As String, strNowData As String, strNote As String
'Dim strTo As String 'Add By Sindy 2012/8/23
Dim strSubject As String, strContent As String 'Add By Sindy 2019/5/24
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '檢查有無異動資料:
   '畫面上的欄位值
   strB1004 = DBDATE(textSO02)
   strB1005 = textSO03_1 & Format("00" & textSO03_2, "00")
   strB1007 = textSO04_1 & Format("00" & textSO04_2, "00")
   'Modify By Sindy 2016/12/27
   If ChkWorkDay(ChangeTStringToWString(textSO02), textSO01, True) = False Then '假日
      strB1012 = ""
      strB1013 = textSO0506
   Else
      strB1012 = textSO0506
      strB1013 = ""
   End If
   '2016/12/27 END
   '串原始資料
   strOldData = strOldData & "," & ChangeWStringToTDateString(m_B1004) & "," & Format(m_B1005, "##:##")
   strOldData = strOldData & "," & Format(m_B1007, "##:##")
   strOldData = strOldData & ",實際時數" & m_B1030 & ";" & IIf(m_B1012 <> "", "平日" & m_B1012 & "時", "假日" & m_B1013 & "時")
   '串目前畫面上資料
   strNowData = strNowData & "," & ChangeWStringToTDateString(strB1004) & "," & Format(strB1005, "##:##")
   strNowData = strNowData & "," & Format(strB1007, "##:##")
   strNowData = strNowData & ",實際時數" & textSO15 & ";" & IIf(strB1012 <> "", "平日" & strB1012 & "時", "假日" & strB1013 & "時")
   If Left(strOldData, 1) = "," Then strOldData = Right(strOldData, Len(strOldData) - 1)
   If Left(strNowData, 1) = "," Then strNowData = Right(strNowData, Len(strNowData) - 1)
   
   '流程備註檔
   If txtNote.Text <> "" And textSO13 <> "" Then
      strSql = GetInsertABS012Sql(Trim(textSO13), 人事處, strUpdDate, strUpdTime, "", txtNote)
      cnnConnection.Execute strSql
   End If
   
   If strOldData <> strNowData And textSO13 <> "" Then '電子簽核的,非紙本
      '人事處尚未簽收時,在人事系統已先建立此表單編號資料,須一併更新出缺勤電子簽核主檔資料
      If m_B1019 = "" Then
         strSql = "update ABS010 set " & _
                  "B1004= " & CNULL(DBDATE(strB1004)) & _
                  ",B1005= " & CNULL(strB1005) & _
                  ",B1007= " & CNULL(strB1007) & _
                  ",B1012= " & CNULL(strB1012) & _
                  ",B1013= " & CNULL(strB1013) & _
                  ",B1030= " & CNULL(textSO15) & _
                  " where B1001=" & CNULL(textSO13)
         cnnConnection.Execute strSql
      End If
      '檢查有異動資料時,須記錄異動資訊到表單流程備註
      strNote = "修改資料" & strOldData & "->" & strNowData
      strSql = GetInsertABS012Sql(Trim(textSO13), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   End If
   
   If m_B1019 = "" And m_EditMode = 1 And textSO13 <> "" Then
      '寄E-Mail通知當事人
      PUB_SendMail strUserNum, textSO01, "", "表單人事處已先行作業，請儘速簽核。", _
      "表單內容為，" & strNowData & vbCrLf & _
      "(表單編號：" & textSO13 & ")", , , , , , , , , , True
   ElseIf m_B1019 <> "" And m_EditMode = 1 And textSO13 <> "" Then
      strSql = "update ABS010 set " & _
               "B1018='" & 已核准 & "'" & _
               " where B1001=" & CNULL(textSO13)
      cnnConnection.Execute strSql
      
      '記錄資訊到表單流程備註
      strNote = "補入資料"
      strSql = GetInsertABS012Sql(Trim(textSO13), "M21", strUpdDate, strUpdTime, "", strNote)
      cnnConnection.Execute strSql
   Else
      If strOldData <> strNowData Then
'         '寄E-Mail通知當事人有異動內容
'         'Modify By Sindy 2012/8/23 發E-Mail通知當事人之外，已簽核的職代及審核主管亦也要通知
'         strTo = GetBossB1107_All(textSO13)
'         '專利處P10-P14,必須另外E-Mail通知71011王副總
'         If (GetStaffDepartment(textSO01) >= "P10" And GetStaffDepartment(textSO01) <= "P14") And _
'            InStr(strTo, "71011") = 0 Then
'            strTo = strTo + ";71011"
'         End If
         
         If textSO13 <> "" Then '電子簽核的,非紙本
            strSubject = "[通知]人事處有修改資料(表單編號：" & textSO13 & ")"
         Else
            strSubject = "[通知]人事處有修改資料"
         End If
         
'         PUB_SendMail strUserNum, textSO01, "", "[通知]人事處有修改資料(表單編號：" & textSO13 & ")", _
'         "異動前資料：" & strOldData & vbCrLf & _
'         "異動後資料：" & strNowData & vbCrLf & _
'         "(表單編號：" & textSO13 & ")" & vbCrLf & vbCrLf & _
'         "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
         strContent = "異動前資料：" & strOldData & vbCrLf & _
                      "異動後資料：" & strNowData & vbCrLf & _
                      IIf(textSO13 <> "", "(表單編號：" & textSO13 & ")" & vbCrLf & vbCrLf, "") & _
                      "人事處修改原因：" & txtNote
         '2012/8/23 End
         'Add By Sindy 2019/5/23 假單完成,後續資料檢查及SendMail
         Call PUB_AutoM21Receive_SendMail(IIf(textSO13 <> "", textSO13, ""), 表單類別_加班, textSO01, DBDATE(textSO02), Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00")), _
            DBDATE(textSO02), "", Left(DBDATE(textSO02), 6), , strSubject, strContent, m_EditMode)
      End If
   End If
   
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Sub

Private Sub txtNote_GotFocus()
   InverseTextBox txtNote
   OpenIme
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
If txtNote <> "" Then
   If CheckLengthIsOK(txtNote, txtNote.MaxLength) = False Then
      Call txtNote_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
End Sub

Private Sub DelMark()
'Dim strTo As String
Dim nResponse
Dim m_B1018 As String, strUpdDate As String, strUpdTime As String
Dim strSubject As String, strContent As String
   
On Error GoTo ErrHand
   
'   If txtNote.Text = "" Then
'      MsgBox "原因不可以空白！", vbExclamation
'      txtNote.SetFocus
'      Exit Sub
'   End If
'
'   nResponse = MsgBox("註銷會將人事系統的相關資料一併刪除，確定要註銷嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問")
'   If nResponse = vbNo Then Exit Sub
   
   If textSO13 <> "" Then m_B1018 = 註銷 '(06)
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   'Modify By Sindy 2019/5/24
   'strContent = GetEMailContent(textSA09, strSubject)
   strContent = GetEMailContent(IIf(textSO13 <> "", textSO13, ""), strSubject, m_B1018, , , "02", textSO01, DBDATE(textSO02), Val(Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00"))), m_EditMode)
   strContent = strContent & vbCrLf & vbCrLf & _
                "人事處修改原因：" & txtNote
   
'   cnnConnection.BeginTrans
   
   If textSO13 <> "" Then
      '流程備註檔
      If txtNote.Text <> "" Then
         strSql = GetInsertABS012Sql(Trim(textSO13), 人事處, strUpdDate, strUpdTime, m_B1018, txtNote)
         cnnConnection.Execute strSql
      End If
      '主檔
      strSql = "update ABS010 set " & _
               "B1018='" & m_B1018 & "'" & _
               " where B1001='" & textSO13 & "' "
      cnnConnection.Execute strSql
   End If
'   '刪除人事系統該筆表單資料, 並寫Log
'   If Left(CboB1002, 2) = 表單類別_請假 Then
'      strSql = "delete from Staff_Absence where SA09='" & Trim(txtB1001) & "'"
'   ElseIf Left(CboB1002, 2) = 表單類別_加班 Then
'      strSql = "delete from Staff_Overtime where So13='" & Trim(txtB1001) & "'"
'   ElseIf Left(CboB1002, 2) = 表單類別_出差 Then
'      strSql = "delete from Staff_Busi_Trip where SB10='" & Trim(txtB1001) & "'"
'   End If
'   Pub_SeekTbLog strSql '記錄刪除Log
'   cnnConnection.Execute strSql
'
'   cnnConnection.CommitTrans
   
'   '發E-Mail通知當事人及已簽核的職代及審核主管
'   strTo = GetBossB1107_All(textSO13)
'   'Add By Sindy 2012/8/23 專利處P10-P14,必須另外E-Mail通知71011王副總
'   If (GetStaffDepartment(textSO01) >= "P10" And GetStaffDepartment(textSO01) <= "P14") And _
'      InStr(strTo, "71011") = 0 Then
'      strTo = strTo + ";71011"
'   End If
'   '2012/8/23 End
'   strContent = GetEMailContent(textSO13, strSubject)
'   'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
'   PUB_SendMail strUserNum, textSO01, "", strSubject, strContent & vbCrLf & vbCrLf & _
'         "人事處修改原因：" & txtNote, , , , , , strTo, , , , True
   
   'Add By Sindy 2019/5/23 假單完成,後續資料檢查及SendMail
   Call PUB_AutoM21Receive_SendMail(IIf(textSO13 <> "", textSO13, ""), 表單類別_加班, textSO01, DBDATE(textSO02), Trim(Format("00" & textSO03_1, "00") & Format("00" & textSO03_2, "00")), _
      DBDATE(textSO02), "", Left(DBDATE(textSO02), 6), , strSubject, strContent, m_EditMode)
      
'   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   MsgBox "註銷失敗！" & vbCrLf & Err.Description
End Sub
