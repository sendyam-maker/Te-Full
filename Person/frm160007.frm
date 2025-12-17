VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160007 
   BorderStyle     =   1  '單線固定
   Caption         =   "人事異動資料"
   ClientHeight    =   5076
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8184
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5076
   ScaleWidth      =   8184
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7500
      Top             =   0
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
            Picture         =   "frm160007.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160007.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   8184
      _ExtentX        =   14436
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
      TabIndex        =   20
      Top             =   690
      Width           =   8115
      _ExtentX        =   14309
      _ExtentY        =   7726
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "單筆資料"
      TabPicture(0)   =   "frm160007.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(17)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(5)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(6)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(7)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(8)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(9)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(10)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(11)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(12)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "textSC01_2"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label23"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textSC02"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textSC01"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textSC07_1"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textSC07"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textSC03"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textSC04_1"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textSC05_1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textSC06_1"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textSC04"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textSC05"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textSC06"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textSC14_1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textSC14"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "多筆瀏覽"
      TabPicture(1)   =   "frm160007.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textSC03_Q"
      Tab(1).Control(1)=   "txt1(0)"
      Tab(1).Control(2)=   "txt1(1)"
      Tab(1).Control(3)=   "txt1(2)"
      Tab(1).Control(4)=   "txt1(3)"
      Tab(1).Control(5)=   "cmdok"
      Tab(1).Control(6)=   "GRD1"
      Tab(1).Control(7)=   "Label1(13)"
      Tab(1).Control(8)=   "Line5"
      Tab(1).Control(9)=   "Line4"
      Tab(1).Control(10)=   "Label15"
      Tab(1).Control(11)=   "Label16"
      Tab(1).ControlCount=   12
      Begin VB.ComboBox textSC03_Q 
         Height          =   260
         ItemData        =   "frm160007.frx":212C
         Left            =   -74010
         List            =   "frm160007.frx":212E
         TabIndex        =   17
         Top             =   660
         Width           =   2835
      End
      Begin VB.ComboBox textSC14 
         Height          =   260
         ItemData        =   "frm160007.frx":2130
         Left            =   1020
         List            =   "frm160007.frx":2143
         TabIndex        =   12
         Top             =   2880
         Width           =   2835
      End
      Begin VB.ComboBox textSC14_1 
         Height          =   260
         ItemData        =   "frm160007.frx":2174
         Left            =   1020
         List            =   "frm160007.frx":2187
         TabIndex        =   5
         Top             =   1290
         Width           =   2835
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   0
         Left            =   -74010
         MaxLength       =   6
         TabIndex        =   13
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   1
         Left            =   -72960
         MaxLength       =   6
         TabIndex        =   14
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   2
         Left            =   -71070
         MaxLength       =   7
         TabIndex        =   15
         Top             =   390
         Width           =   915
      End
      Begin VB.TextBox txt1 
         Height          =   270
         Index           =   3
         Left            =   -70080
         MaxLength       =   7
         TabIndex        =   16
         Top             =   390
         Width           =   915
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "查詢"
         Height          =   345
         Left            =   -68700
         TabIndex        =   18
         Top             =   390
         Width           =   915
      End
      Begin VB.ComboBox textSC06 
         Height          =   260
         Left            =   5100
         TabIndex        =   10
         Top             =   2250
         Width           =   2835
      End
      Begin VB.ComboBox textSC05 
         Height          =   260
         Left            =   1020
         TabIndex        =   9
         Top             =   2250
         Width           =   2835
      End
      Begin VB.ComboBox textSC04 
         Height          =   260
         Left            =   1020
         TabIndex        =   8
         Top             =   1920
         Width           =   2835
      End
      Begin VB.ComboBox textSC06_1 
         Height          =   260
         Left            =   5100
         TabIndex        =   3
         Top             =   660
         Width           =   2835
      End
      Begin VB.ComboBox textSC05_1 
         Height          =   260
         Left            =   1020
         TabIndex        =   2
         Top             =   660
         Width           =   2835
      End
      Begin VB.ComboBox textSC04_1 
         Height          =   260
         Left            =   5100
         TabIndex        =   1
         Top             =   360
         Width           =   2835
      End
      Begin VB.ComboBox textSC03 
         Height          =   260
         Left            =   5100
         TabIndex        =   7
         Top             =   1620
         Width           =   2835
      End
      Begin VB.TextBox textSC07 
         Height          =   270
         Left            =   1020
         MaxLength       =   80
         TabIndex        =   11
         Top             =   2580
         Width           =   6885
      End
      Begin VB.TextBox textSC07_1 
         Height          =   270
         Left            =   1020
         MaxLength       =   80
         TabIndex        =   4
         Top             =   990
         Width           =   6885
      End
      Begin VB.TextBox textSC01 
         Height          =   270
         Left            =   1020
         MaxLength       =   6
         TabIndex        =   0
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox textSC02 
         Height          =   270
         Left            =   1020
         MaxLength       =   7
         TabIndex        =   6
         Top             =   1620
         Width           =   945
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm160007.frx":21B8
         Height          =   3345
         Left            =   -74970
         TabIndex        =   21
         Top             =   990
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   5906
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "異動原因："
         Height          =   180
         Index           =   13
         Left            =   -74940
         TabIndex        =   40
         Top             =   690
         Width           =   900
      End
      Begin MSForms.Label Label23 
         Height          =   195
         Left            =   150
         TabIndex        =   39
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
      Begin MSForms.Label textSC01_2 
         Height          =   225
         Left            =   1800
         TabIndex        =   38
         Top             =   390
         Width           =   1395
         BackColor       =   12632256
         VariousPropertyBits=   27
         Size            =   "2461;397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新所別："
         Height          =   180
         Index           =   12
         Left            =   270
         TabIndex        =   37
         Top             =   2940
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "所別："
         Height          =   180
         Index           =   11
         Left            =   450
         TabIndex        =   36
         Top             =   1350
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "註：存檔後會更新員工基本資料"
         Height          =   180
         Index           =   10
         Left            =   90
         TabIndex        =   35
         Top             =   3300
         Width           =   2520
      End
      Begin VB.Line Line5 
         X1              =   -70350
         X2              =   -69750
         Y1              =   510
         Y2              =   510
      End
      Begin VB.Line Line4 
         X1              =   -73320
         X2              =   -72630
         Y1              =   540
         Y2              =   540
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Left            =   -74940
         TabIndex        =   34
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日期："
         Height          =   180
         Left            =   -71690
         TabIndex        =   33
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新職稱說明："
         Height          =   180
         Index           =   9
         Left            =   30
         TabIndex        =   32
         Top             =   2610
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新  職  位："
         Height          =   180
         Index           =   8
         Left            =   4110
         TabIndex        =   31
         Top             =   2310
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新  職  稱："
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   30
         Top             =   2310
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "新  部  門："
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   29
         Top             =   1980
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "異動原因："
         Height          =   180
         Index           =   5
         Left            =   4110
         TabIndex        =   28
         Top             =   1650
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "職稱說明："
         Height          =   180
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   1050
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "職　　位："
         Height          =   180
         Index           =   3
         Left            =   4110
         TabIndex        =   26
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "職　　稱："
         Height          =   180
         Index           =   1
         Left            =   90
         TabIndex        =   25
         Top             =   750
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "員工編號："
         Height          =   180
         Index           =   0
         Left            =   90
         TabIndex        =   24
         Top             =   405
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部　　門："
         Height          =   180
         Index           =   17
         Left            =   4110
         TabIndex        =   23
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "異動日期："
         Height          =   180
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   1650
         Width           =   900
      End
   End
End
Attribute VB_Name = "frm160007"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/20 日期欄已修改
'Create by nickc 2006/12/04 copy from frm140401
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
Dim tf_SC As Integer
Dim MyKind As String
Dim m_ST29 As String 'Add By Sindy 2011/10/26
Dim m_bolKillFingerPrinter As Boolean 'Added by Morgan 2013/8/2


Private Sub cmdok_Click()
'Modify By Sindy 2023/2/24 + & textSC03_Q
If txt1(0) & txt1(1) & txt1(2) & txt1(3) & textSC03_Q <> "" Then
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
rsA.Open "select * from staff_change where rownum <2 ", cnnConnection, adOpenStatic, adLockReadOnly
tf_SC = rsA.Fields.Count
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

   ReDim m_FieldList(tf_SC) As FIELDITEM
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   textSC02.BackColor = &H8000000F
   textSC01.BackColor = &H8000000F
   
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
   
   Set frm160007 = Nothing
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
         textSC01.Text = GRD1.TextMatrix(tmpMouseRow, 0)
         textSC02.Text = ChangeTDateStringToTString(GRD1.TextMatrix(tmpMouseRow, 2))
         QueryRecord
         '2008/12/12 END
         GRD1.Visible = True
    End If
End If
End Sub

'Add By Sindy 2019/8/27
Private Sub SSTab1_Click(PreviousTab As Integer)
   Dim oCtrl As Control
   If PreviousTab = 0 Then
      cmdok.SetFocus
      cmdok.Default = True
   Else
      cmdok.Default = False
      'Added by Morgan 2024/5/23 修正ComboBox欄位會自動被反白導致看不見資料問題'
      For Each oCtrl In Me.Controls
         If TypeName(oCtrl) = "ComboBox" Then
            oCtrl.SelLength = 0
         End If
      Next
      'end 2024/5/23
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
   
   If IsNull(rsSrcTmp.Fields("sc08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc08")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("sc08"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sc09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc09")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sc09"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sc10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc10")) = False Then
         strTemp = rsSrcTmp.Fields("sc10")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sc11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc11")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("sc11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sc12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("sc12"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("sc13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("sc13")) = False Then
         strTemp = rsSrcTmp.Fields("sc13")
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
   If Me.textSC01.Enabled = True Then
      Cancel = False
      textSC01_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSC01.Text = "" Then
       MsgBox "員工編號不可以空白！", vbExclamation
       textSC01.SetFocus
       Exit Function
   End If
   If Me.textSC02.Enabled = True Then
      Cancel = False
      textSC02_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSC02.Text = "" Then
       MsgBox "異動日不可以空白！", vbExclamation
       textSC02.SetFocus
       Exit Function
   End If
   If Me.textSC03.Enabled = True Then
      Cancel = False
      textSC03_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textSC03.Text = "" Then
       MsgBox "異動原因不可以空白！", vbExclamation
       textSC03.SetFocus
       Exit Function
   End If
   
   'Add By Sindy 2011/10/17 增加判斷員工代號+日期是否人員已離職
   If Left(textSC03.Text, 2) <> "02" Then '不等於復職時才檢查
      'add by sonia 2014/11/17 留職停薪者可直接輸離職(90012洪丹怡)
      If ChkStaffST04(textSC01, False, textSC02) = True Then
         If Left(textSC03.Text, 2) = "03" Then
            strExc(0) = "select sd02 from salarydata where sd01='" & textSC01 & "' and sd02='S' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               GoTo Nextstep
            End If
         End If
      End If
      'end 2014/11/17
      If ChkStaffST04(textSC01, True, textSC02) = True Then
         textSC01.SetFocus
         Exit Function
      End If
   End If
   
Nextstep:   'add by sonia 2014/11/17
   ' 2008/12/22 Add BY SINDY
   '異動原因是07.調職時，新部門不可與原部門相同！
   If Left(Trim(textSC03.Text), 2) = "07" Then
      If textSC04.Text = textSC04_1.Text And textSC14.Text = textSC14_1.Text Then
         '2011/5/11 MODIFY BY SONIA 改為提醒仍可輸入
         'MsgBox "異動原因是07.調職時，新部門、所別不可與原部門及所別相同！", vbExclamation
         'textSC04.SetFocus
         'Exit Function
         If MsgBox("請注意！異動原因是07.調職時，新部門、所別與原部門及所別都相同, 是否仍要存檔？", vbExclamation + vbYesNo, "提醒！") = vbNo Then
            textSC04.SetFocus
            Exit Function
         End If
         '2011/5/11 END
      End If
   End If
   'Modify By Sindy 2014/5/2 Mark 因顏經理由中一區經理變成台中區經理
   ''異動原因是05.晉升或11.歸籍時，新職稱或新職位至少一項與原資料不同！
   'If Left(Trim(textSC03.Text), 2) = "05" Or Left(Trim(textSC03.Text), 2) = "11" Then
   '   If (textSC05.Text = textSC05_1.Text) And (textSC06.Text = textSC06_1.Text) Then
   '      MsgBox "異動原因是05.晉升或11.歸籍時，新職稱或新職位至少一項與原資料不同！", vbExclamation
   '      textSC05.SetFocus
   '      Exit Function
   '   End If
   'End If
   ' 2008/12/22 END
   If Me.textSC04.Enabled = True Then
      Cancel = False
      textSC04_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   ' 2008/12/22 Add BY SINDY
   If textSC04.Text = "" Then
       MsgBox "新部門不可以空白！", vbExclamation
       textSC04.SetFocus
       Exit Function
   End If
   ' 2008/12/22 END
   If Me.textSC05.Enabled = True Then
      Cancel = False
      textSC05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '' 2008/12/22 Add BY SINDY
   'If textSC05.Text = "" Then
   '    MsgBox "新職稱不可以空白！", vbExclamation
   '    textSC05.SetFocus
   '    Exit Function
   'End If
   '' 2008/12/22 END
   If Me.textSC06.Enabled = True Then
      Cancel = False
      textSC06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   ' 2008/12/22 Add BY SINDY
   If textSC06.Text = "" And textSC05.Text = "" Then
       MsgBox "新職稱或職位至少要輸入一項！", vbExclamation
       textSC06.SetFocus
       Exit Function
   End If
   ' 2008/12/22 END
   If Me.textSC07.Enabled = True Then
      Cancel = False
      textSC07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2010/9/20
   If textSC14.Text = "" Then
       MsgBox "新所別不可以空白！", vbExclamation
       textSC14.SetFocus
       Exit Function
   End If
   If Me.textSC14.Enabled = True Then
      Cancel = False
      textSC14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2021/7/5
   If m_EditMode = 1 And Left(Trim(textSC03.Text), 2) = "01" Then '新進
      strSql = "SELECT * FROM staff_change " & _
               "WHERE SC01 = '" & textSC01 & "'  and SC03='" & Left(Trim(textSC03.Text), 2) & "'  "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp.RecordCount >= 1 Then
            MsgBox "新進不可以再輸第2次！", vbExclamation
            textSC03.SetFocus
            Exit Function
         End If
      End If
   End If
   '2021/7/5 END
   
   'Added by Morgan 2013/8/2
   '離職檢查是否有指紋檔並提醒會清除
   m_bolKillFingerPrinter = False
   'Modified by Morgan 2015/6/18 +08,09,10
   If InStr("03,08,09,10", Left(Trim(textSC03.Text), 2)) > 0 Then
      strExc(0) = "select * from staffcarddata where scd01='" & textSC01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modified by Morgan 2015/6/18
         'If MsgBox("員工離職將會清除指紋及卡片資料(含考勤機)，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
         If MsgBox("員工離職將會清除指紋資料(含考勤機)，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
            m_bolKillFingerPrinter = True
         Else
            Exit Function
         End If
      End If
   End If
   'end 2013/8/2
   
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
   For nIndex = 0 To tf_SC - 1 'edit by nickc 2006/10/24  MAX_FIELD - 1
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
   
   For nIndex = 0 To tf_SC - 1
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

'Modify By Sindy 2024/5/27 寫成共用函數
Private Function PressCardData() As String
   
   PressCardData = ""
   
   'Added Morgan 2013/8/2
   If m_bolKillFingerPrinter = True Then
      'Add By Sindy 2015/3/16 +if
      If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
         MsgBox "略過清除刷卡機資料程式,因會直接刪除「刷卡機」資料,請單獨測試"
      Else
      '2015/3/16 END
         If PUB_ClearCardData(textSC01.Text) = False Then
            'Add By Sindy 2024/5/2 刷卡資料有誤,無法刪除; 但人員急著建檔,先略過
            If MsgBox("清除刷卡機資料失敗，是否仍要繼續建檔？" & vbCrLf & _
                      "（人員急著建檔時，先略過！請通知電腦中心處理）", vbExclamation + vbYesNo + vbDefaultButton2, "提醒！") = vbNo Then
               PressCardData = "N"
            Else
               PressCardData = "Y"
            End If
            Exit Function
            '2024/5/2 END
         End If
      End If
   End If
   'end 2013/8/3
End Function

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
   Dim strSC02 As String
   Dim strSC01 As String
   Dim MyArr As Variant
   Dim oMailCount As String ' Add By Sindy 98/03/02
   Dim m_FnoMSG As String   '2009/12/10 add by sonia
   Dim rsTmp As New ADODB.Recordset
   Dim strContent As String
   Dim strTo As String, strCC As String 'Add By Sindy 2024/2/1
   
   AddRecord = False
   
   strSC01 = textSC01
   strSC02 = DBDATE(textSC02)

   ' 檢查記錄是否已存在
   If IsRecordExist(strSC01, strSC02) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      Exit Function
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO staff_change ("
   For nIndex = 0 To tf_SC - 1
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
   For nIndex = 0 To tf_SC - 1
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
   Screen.MousePointer = vbHourglass 'Add By Sindy 2014/3/24
   
   'Modify By Sindy 2024/5/27
   strExc(10) = PressCardData()
   If strExc(10) = "N" Then
      Screen.MousePointer = vbDefault
      Exit Function
   ElseIf strExc(10) = "Y" Then
      cnnConnection.BeginTrans
      Screen.MousePointer = vbHourglass
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
'   'Added Morgan 2013/8/2
'   If m_bolKillFingerPrinter = True Then
'      'Add By Sindy 2015/3/16 +if
'      If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
'         MsgBox "略過清除刷卡機資料程式,因會直接刪除「刷卡機」資料,請單獨測試"
'      Else
'      '2015/3/16 END
'         If PUB_ClearCardData(textSC01.Text) = False Then
''            Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
''            Exit Function
'            'Add By Sindy 2024/5/2 刷卡資料有誤,無法刪除; 但人員急著建檔,先略過
'            If MsgBox("清除刷卡機資料失敗，是否仍要繼續建檔？" & vbCrLf & _
'                      "（人員急著建檔時，先略過！請通知電腦中心處理）", vbExclamation + vbYesNo + vbDefaultButton2, "提醒！") = vbNo Then
'               Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
'               Exit Function
'            Else
'               cnnConnection.BeginTrans
'               Screen.MousePointer = vbHourglass
'            End If
'            '2024/5/2 END
'         End If
'      End If
'   End If
'   'end 2013/8/3
   '2024/5/27 END
   
   ' 2008/12/29 Add BY SINDY
   '增加判斷 異動日期<971102者都不要更新回員工基本檔.
   '因為 971102以前都是補資料.
   If Val(textSC02) >= 971102 Then
         '異動員工基本資料 ********************************
         strSql = ""
         '部門:
         'Modify By Sindy 2023/12/20
         If strSrvDate(1) >= 新部門啟用日 Then
            If textSC04.Text <> "" Then
                 MyArr = Split(textSC04, " ")
                 strSql = strSql & " st93='" & MyArr(0) & "' "
            Else
                 strSql = strSql & " st93=null "
            End If
         Else
         '2023/12/20 END
            If textSC04.Text <> "" Then
                 MyArr = Split(textSC04, " ")
                 strSql = strSql & " st03='" & MyArr(0) & "' "
            Else
                 strSql = strSql & " st03=null "
            End If
         End If
         '職稱:
         If textSC05.Text <> "" Then
              MyArr = Split(textSC05, " ")
              strSql = strSql & " ,st20='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st20=null "
         End If
         '職位:
         If textSC06.Text <> "" Then
              MyArr = Split(textSC06, " ")
              strSql = strSql & " ,st21='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st21=null "
         End If
         'Add By Sindy 2010/9/20
         '所別:
         If textSC14.Text <> "" Then
              MyArr = Split(textSC14, " ")
              strSql = strSql & " ,st06='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st06=null "
         End If
         '2010/9/20 End
         '職稱說明:
         If textSC07.Text <> "" Then
              strSql = strSql & " ,st49='" & ChgSQL(textSC07) & "' "
         Else
              strSql = strSql & " ,st49=null "
         End If
         '離職日:
         '異動原因為02復職時，更新員工檔之離職日為null；
         If Left(Trim(textSC03.Text), 2) = "02" Then
              strSql = strSql & " ,st51=null "
         End If
         '異動原因為03離職、04留職停薪、08退休、09撤職、10資遣，更新員工檔之離職日為異動日期。
         If Left(Trim(textSC03.Text), 2) = "03" Or _
            Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "08" Or _
            Left(Trim(textSC03.Text), 2) = "09" Or _
            Left(Trim(textSC03.Text), 2) = "10" Then
              strSql = strSql & " ,st51='" & ChangeTStringToWString(Trim(textSC02)) & "' "
         End If
         If Left(Trim(strSql), 1) = "," Then
              strSql = Mid(Trim(strSql), 2)
         End If
         Pub_SeekTbLog "update staff set " & strSql & " where st01='" & strSC01 & "' "
         cnnConnection.Execute "update staff set " & strSql & " where st01='" & strSC01 & "' "
         'Added by Lydia 2025/08/14 利益衝突案件：檢查利益衝突案件之權限，若人事有異動留下相關記錄
         Call PUB_SaveCUFA_Staff_Log(False, strSC01, Left(Trim(textSC03.Text), 2), Me.Name, pub_HostName)
         
         '異動員工基本資料 END ********************************
         
         '異動薪資基本檔 ********************************
         '異動原因04.留職停薪同時更新薪資基本檔SalaryData之sd02編制為’S’。
         'Modify By Sindy 2010/6/25 異動原因為02復職時，sd02編制為’R’。
         If Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "02" Then
            strSql = ""
            If Left(Trim(textSC03.Text), 2) = "04" Then
               strSql = strSql & " sd02='S' "
            ElseIf Left(Trim(textSC03.Text), 2) = "02" Then
               strSql = strSql & " sd02='R' "
            End If
            If Left(Trim(strSql), 1) = "," Then
                 strSql = Mid(Trim(strSql), 2)
            End If
            Pub_SeekTbLog "update SalaryData set " & strSql & " where sd01='" & strSC01 & "' "
            cnnConnection.Execute "update SalaryData set " & strSql & " where sd01='" & strSC01 & "' "
         End If
         
         '2010/11/11 ADD BY SONIA 異動原因為02復職時,且為北所SD05='2'有帳號SD06者,同時更新帳號通知銀行薪資年月為復職年月
         If Left(Trim(textSC03.Text), 2) = "02" Then
            strExc(0) = "select sd05,sd06 from salarydata where sd01='" & strSC01 & "' and sd05='2' and sd06 is not null "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Pub_SeekTbLog "update SalaryData set sd46=" & Val(Left(DBDATE(textSC02), 6)) & " where sd01='" & strSC01 & "' "
               cnnConnection.Execute "update SalaryData set sd46=" & Val(Left(DBDATE(textSC02), 6)) & " where sd01='" & strSC01 & "' "
            End If
         End If
         '2010/11/11 END
         
         'Add by Morgan 2010/7/14 若職位有變動時重算婚喪扣款並更新回薪資檔
         'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
         'If m_FieldList(5).fiOldData <> m_FieldList(5).fiNewData Then
         '   strExc(0) = "select max(sc02) from staff_change where sc01='" & textSC01 & "'"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 1 Then
         '      '最後異動才做
         '      If RsTemp(0) = DBDATE(textSC02) Then
         '         If PUB_GetHelpFee(textSC01, strExc(1)) = True Then
         '            strSql = "update salarydata set sd09=decode(nvl(sd09,0),0,sd09," & Val(strExc(1)) & "),sd10=decode(nvl(sd10,0),0,sd10," & Val(strExc(1)) & ") where sd01='" & textSC01 & "'"
         '            Pub_SeekTbLog strSql
         '            cnnConnection.Execute strSql
         '         End If
         '      End If
         '   End If
         'End If
         'end 2025/7/29
         
         '異動薪資基本檔 END ********************************
   End If
   
   'Add By Sindy 2019/8/12
   '異動原因為02復職時，更新留職停薪特休起算日
   If Left(Trim(textSC03.Text), 2) = "02" Then
      strTmp = Pub_BackTaieToDate(textSC01, "")
      If Val(strTmp) > 0 Then
         strSql = "UPDATE staff SET ST72=" & strTmp & " WHERE st01='" & textSC01 & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
   
   'Add By Sindy 2016/5/20
   '智權人員離職時,調整待會稿區正在送會中及會圖中的收受者
   '異動原因為03離職、04留職停薪、08退休、09撤職、10資遣，更新員工檔之離職日為異動日期。
   If Left(Trim(textSC03.Text), 2) = "03" Or _
      Left(Trim(textSC03.Text), 2) = "04" Or _
      Left(Trim(textSC03.Text), 2) = "08" Or _
      Left(Trim(textSC03.Text), 2) = "09" Or _
      Left(Trim(textSC03.Text), 2) = "10" Then
      Call PUB_SalseLeaveUpEEP05(textSC01)
   End If
   
   cnnConnection.CommitTrans
   
   If ((strSC01 & strSC02) < (m_FirstKEY(0) & m_FirstKEY(1))) Or ((strSC01 & strSC02) > (m_LastKEY(0) & m_LastKEY(1))) Then
      RefreshRange
   End If
   
   ' Add By Sindy 98/03/02
   ' 除了05.晉升06.真除外, 都需要發Mail通知
   '2010/5/4 MODIFY BY SONIA 加11.歸籍
   'modify by sonia 2017/5/2 再加12.任命 13,晉升、任命
   If Left(Trim(textSC03.Text), 2) <> "05" And Left(Trim(textSC03.Text), 2) <> "06" And Left(Trim(textSC03.Text), 2) <> "11" And Left(Trim(textSC03.Text), 2) <> "12" And Left(Trim(textSC03.Text), 2) <> "13" Then
      oMailCount = ""
      oMailCount = Pub_GetSpecMan("人事異動郵件通知")
      '2009/12/10 MODIFY BY SONIA 若有外翻編號則請財務改外翻編號之入帳類別,否則翻譯費付款方式會錯
      'PUB_SendMail strUserNum, oMailCount, "", "人事異動通知！", "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
      "異動原因：" + textSC03 & vbCrLf & _
      "異動日期：" + textSC02 & vbCrLf & _
      "部　門：" + textSC04_1 & vbCrLf & _
      "職　位：" + textSC06_1 & vbCrLf & _
      "新部門：" + textSC04 & vbCrLf & _
      "新職位：" + textSC06
      If Left(Trim(textSC03.Text), 2) = "01" Or Left(Trim(textSC03.Text), 2) = "07" Then
         'Modify By Sindy 2020/7/6 增加控管請假時,不發職代
         PUB_SendMail strUserNum, oMailCount, "", "人事異動通知！", "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
         "異動原因：" + textSC03 & vbCrLf & _
         "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf & _
         "部　門：" + textSC04_1 & vbCrLf & _
         "職　位：" + textSC06_1 & vbCrLf & _
         "職　稱：" + textSC05_1 & vbCrLf & _
         "職稱說明：" + textSC07_1 & vbCrLf & _
         "所　別：" + textSC14_1 & vbCrLf & vbCrLf & _
         "新部門：" + textSC04 & vbCrLf & _
         "新職位：" + textSC06 & vbCrLf & _
         "新職稱：" + textSC05 & vbCrLf & _
         "新職稱說明：" + textSC07 & vbCrLf & _
         "新所別：" + textSC14 & vbCrLf, , , , , , , , , , True
      Else
         'Modify By Sindy 2019/8/16 改成共用函數:人員刪除時要檢查的資料
         m_FnoMSG = StaffLeaveChkData(textSC01, textSC03.Text)
         '2019/8/16 END
         
         '2011/5/10 ADD BY SONIA 異動原因為03離職、04留職停薪、08退休、09撤職、10資遣時若為特殊人員或帶人主管
         If Left(Trim(textSC03.Text), 2) = "03" Or _
            Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "08" Or _
            Left(Trim(textSC03.Text), 2) = "09" Or _
            Left(Trim(textSC03.Text), 2) = "10" Then
            
            'Add By Sindy 2011/8/31
            'Modify By Sindy 2020/4/17 + "B0117='" & textSC01 & "' or " & _
                                         "B0119='" & textSC01 & "' or " & _
                                         "B0121='" & textSC01 & "' or " & _
                                         "B0123='" & textSC01 & "' or " & _
                                         "instr(B0124,'" & textSC01 & "')>0 "
            'Modify By Sindy 2023/10/19 + "B0128='" & textSC01 & "' or " & _
                                          "B0129='" & textSC01 & "' or " & _
                                          "B0130='" & textSC01 & "' or " & _
                                          "B0131='" & textSC01 & "' or " & _
                                          "B0132='" & textSC01 & "')"
            strSql = "select * from ABS001,staff where (B0102='" & textSC01 & "' or " & _
                                                "B0103='" & textSC01 & "' or " & _
                                                "B0104='" & textSC01 & "' or " & _
                                                "B0105='" & textSC01 & "' or " & _
                                                "B0106='" & textSC01 & "' or " & _
                                                "B0107='" & textSC01 & "' or " & _
                                                "B0108='" & textSC01 & "' or " & _
                                                "B0109='" & textSC01 & "' or " & _
                                                "B0110='" & textSC01 & "' or " & _
                                                "B0111='" & textSC01 & "' or " & _
                                                "B0117='" & textSC01 & "' or " & _
                                                "B0119='" & textSC01 & "' or " & _
                                                "B0121='" & textSC01 & "' or " & _
                                                "B0123='" & textSC01 & "' or " & _
                                                "instr(B0124,'" & textSC01 & "')>0 or " & _
                                                "B0128='" & textSC01 & "' or " & _
                                                "B0129='" & textSC01 & "' or " & _
                                                "B0130='" & textSC01 & "' or " & _
                                                "B0131='" & textSC01 & "' or " & _
                                                "B0132='" & textSC01 & "')" & _
                                                " and B0101=st01(+) and st04='1'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               MsgBox "須維護「人事職代及審核主管設定」資料！", vbExclamation
            End If
            strSql = "select * from ABS002,staff where (B0202='" & textSC01 & "' or " & _
                                                "B0203='" & textSC01 & "' or " & _
                                                "B0204='" & textSC01 & "' or " & _
                                                "B0205='" & textSC01 & "' or " & _
                                                "B0206='" & textSC01 & "' or " & _
                                                "B0207='" & textSC01 & "') " & _
                                                "and B0201=st01(+) and st04='1'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               MsgBox "須維護「簽核主管特殊身份職務代理人」資料！", vbExclamation
            End If
         End If
         '2011/5/10 END
         
         'Modify By Sindy 2012/4/27 +特休未休天數
         'Modify By Sindy 2020/7/6 增加控管請假時,不發職代
         PUB_SendMail strUserNum, oMailCount, "", "人事異動通知！", "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
         "異動原因：" + textSC03 & vbCrLf & _
         "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf & _
         "部　門：" + textSC04_1 & vbCrLf & _
         "職　位：" + textSC06_1 & vbCrLf & _
         "職　稱：" + textSC05_1 & vbCrLf & _
         "職稱說明：" + textSC07_1 & vbCrLf & _
         "所　別：" + textSC14_1 & vbCrLf & vbCrLf & _
         "新部門：" + textSC04 & vbCrLf & _
         "新職位：" + textSC06 & vbCrLf & _
         "新職稱：" + textSC05 & vbCrLf & _
         "新職稱說明：" + textSC07 & vbCrLf & _
         "新所別：" + textSC14 & vbCrLf & vbCrLf & _
         "特休未休天數：" + GetCurrSpecRestDay(textSC01, 2) & vbCrLf & vbCrLf & vbCrLf & _
         m_FnoMSG, , , , , , , , , , True
         
         'Add By Sindy 2021/2/18 同仁異動系統發通知給分機維護人員
         Call PUB_CallScMailTOM13(textSC01, textSC02)
         '2021/2/18 END
         
         'Add By Sindy 2011/10/26 已實習期滿者離職通知
         'Modify By Sindy 2014/6/4 +textSC01 <= "99053" 因舊員工都沒有試用期間的資料
         'modify by sonia 2015/5/20 加入L02法務部,因為P31,F31都改為L02
'         If Left(Trim(textSC03.Text), 2) = "03" And _
'            (Left(Trim(textSC04_1), 3) = "F21" Or _
'             Left(Trim(textSC04_1), 3) = "P11" Or _
'             Left(Trim(textSC04_1), 2) = "F1" Or _
'             Left(Trim(textSC04_1), 2) = "P2" Or _
'             Left(Trim(textSC04_1), 3) = "L01" Or _
'             Left(Trim(textSC04_1), 3) = "L02" Or _
'             Left(Trim(textSC04_1), 3) = "P31" Or _
'             Left(Trim(textSC04_1), 3) = "F31") And _
'            ((m_ST29 <> "" And m_ST29 < strSrvDate(1)) Or _
'             textSC01 <= "99053") Then
'            oMailCount = ""
'            oMailCount = Pub_GetSpecMan("試用期滿通知")
'            PUB_SendMail strUserNum, oMailCount, "", "已實習期滿者離職通知！", "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
'            "異動原因：" + textSC03 & vbCrLf & _
'            "異動日期：" + textSC02 & vbCrLf & vbCrLf & _
'            "部　門：" + textSC04_1 & vbCrLf & _
'            "職　位：" + textSC06_1 & vbCrLf & _
'            "職　稱：" + textSC05_1 & vbCrLf & _
'            "職稱說明：" + textSC07_1 & vbCrLf & _
'            "所　別：" + textSC14_1 & vbCrLf
'         End If
         'Modify By Sindy 2015/12/25 03離職,08退休,09撤職,10資遣
         'Modify By Sindy 2017/12/6  02復職亦也要通知
         If (Left(Trim(textSC03.Text), 2) = "03" Or _
             Left(Trim(textSC03.Text), 2) = "08" Or _
             Left(Trim(textSC03.Text), 2) = "09" Or _
             Left(Trim(textSC03.Text), 2) = "10" Or _
             Left(Trim(textSC03.Text), 2) = "02") And _
            ((m_ST29 <> "" And m_ST29 < strSrvDate(1)) Or _
             textSC01 <= "99053") Then
             
            'oMailCount = ""
            strTo = ""
            'Modify By Sindy 2023/12/20
            If strSrvDate(1) >= 新部門啟用日 Then
               strSql = "SELECT A0926 as A0912 FROM Acc090NEW WHERE A0921='" & Left(Trim(textSC04_1), 3) & "' and A0926 is not null"
            Else
            '2023/12/20 END
               strSql = "SELECT A0912 FROM Acc090 WHERE A0901='" & Left(Trim(textSC04_1), 3) & "' and A0912 is not null"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strTo = RsTemp.Fields("A0912")
               strCC = Pub_GetSpecMan("試用期滿通知")
            Else
               'Add By Sindy 2024/2/1 增加：員工個人人事資料中之「智權專業證照」有內容者,也要發通知信
               strSql = "select * from staff_specialty where ss01='" & textSC01 & "' and ss04 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strTo = Pub_GetSpecMan("試用期滿通知")
                  strCC = ""
               End If
               '2024/2/1 END
            End If
            'Modify By Sindy 2021/7/16 + 改設定Acc090.A0912欄位值,加入楊經理和Widen
'            'Add By Sindy 2021/3/30 於人事處登錄同仁離職時，由系統通知A4024.Widen
'            oMailCount = oMailCount & IIf(oMailCount <> "", ";", "") & "A4024"
'            '2021/3/30 END
            'If oMailCount <> "" Then
            If strTo <> "" Then
               strContent = "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
                            "異動原因：" + textSC03 & vbCrLf & _
                            "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf
               'Modify By Sindy 2022/5/13
               If Trim(textSC04_1) & Trim(textSC06_1) & Trim(textSC05_1) & Trim(textSC07_1) & Trim(textSC14_1) = Trim(textSC04) & Trim(textSC06) & Trim(textSC05) & Trim(textSC07) & Trim(textSC14) Then
                  strContent = strContent & _
                            "部　門：" + textSC04_1 & vbCrLf & _
                            "職　位：" + textSC06_1 & vbCrLf & _
                            "職　稱：" + textSC05_1 & vbCrLf & _
                            "職稱說明：" + textSC07_1 & vbCrLf & _
                            "所　別：" + textSC14_1 & vbCrLf
               Else
                  strContent = strContent & _
                            "部　門：" + textSC04 & vbCrLf & _
                            "職　位：" + textSC06 & vbCrLf & _
                            "職　稱：" + textSC05 & vbCrLf & _
                            "職稱說明：" + textSC07 & vbCrLf & _
                            "所　別：" + textSC14 & vbCrLf
               End If
               '2022/5/13 END
               'Modify By Sindy 2020/7/6 增加控管請假時,不發職代
               'Modify By Sindy 2022/11/30 + CC: Pub_GetSpecMan("試用期滿通知")
               PUB_SendMail strUserNum, strTo, "", IIf(Left(Trim(textSC03.Text), 2) = "02", "同仁復職通知，請注意台一網頁同仁資料是否需要調整。", "已實習期滿者離職通知，請撤下台一網頁離職同仁資料。"), strContent, , , , , , strCC, , , , True
            End If
         End If
         '2015/12/25 END
      End If
      '2009/12/10 end
   'Add By Sindy 2017/9/18
   Else
      '新增資料時, 若原部門或新部門為'S'字頭人員,
      '若有異動職稱欄時, 請發E-mail至財務信箱account@taie.com.tw
      If (Left(Trim(textSC04_1), 1) = "S" Or Left(Trim(textSC04), 1) = "S") And _
         (Trim(textSC05_1) <> Trim(textSC05)) Then
         'Modify By Sindy 2020/7/6 增加控管請假時,不發職代
         PUB_SendMail strUserNum, "account@taie.com.tw", "", "智權部同仁" & textSC03 & "通知！", _
            "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
            "異動原因：" + textSC03 & vbCrLf & _
            "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf & _
            "部　門：" + textSC04_1 & vbCrLf & _
            "職　位：" + textSC06_1 & vbCrLf & _
            "職　稱：" + textSC05_1 & vbCrLf & _
            "職稱說明：" + textSC07_1 & vbCrLf & _
            "所　別：" + textSC14_1 & vbCrLf & vbCrLf & _
            "新部門：" + textSC04 & vbCrLf & _
            "新職位：" + textSC06 & vbCrLf & _
            "新職稱：" + textSC05 & vbCrLf & _
            "新職稱說明：" + textSC07 & vbCrLf & _
            "新所別：" + textSC14 & vbCrLf, , , , , , , , , , True
      End If
      '2017/9/18 END
      'add by sonia 2019/6/25
      If Left(Trim(textSC03.Text), 2) <> "06" Then
         'Modify By Sindy 2020/7/6 增加控管請假時,不發職代
         PUB_SendMail strUserNum, GetDeptMan("M51") & ";" & Pub_GetSpecMan("程式管理人員"), "", "人事職稱異動通知！" & textSC03 & "通知！請調整Outlook人員職稱", _
            "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
            "異動原因：" + textSC03 & vbCrLf & _
            "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf & _
            "部　門：" + textSC04_1 & vbCrLf & _
            "職　位：" + textSC06_1 & vbCrLf & _
            "職　稱：" + textSC05_1 & vbCrLf & _
            "職稱說明：" + textSC07_1 & vbCrLf & _
            "所　別：" + textSC14_1 & vbCrLf & vbCrLf & _
            "新部門：" + textSC04 & vbCrLf & _
            "新職位：" + textSC06 & vbCrLf & _
            "新職稱：" + textSC05 & vbCrLf & _
            "新職稱說明：" + textSC07 & vbCrLf & _
            "新所別：" + textSC14 & vbCrLf, , , , , , , , , , True
      End If
      'end 2019/6/25
   End If
   '98/03/02 End
   
   'Add By Sindy 2021/7/16 劉經理認為只有6.真除及99其他可以不用給，其他都要給，否則有漏失就不好了。
   '另01新進、04留職停薪、05晉升、07 調職、11 歸籍、12任命、13 晉升、任命、14 真除、任命等8項請以系統通知99033、A4024、75033。
   If (Left(Trim(textSC03.Text), 2) = "01" Or _
       Left(Trim(textSC03.Text), 2) = "04" Or _
       Left(Trim(textSC03.Text), 2) = "05" Or _
       Left(Trim(textSC03.Text), 2) = "07" Or _
       Left(Trim(textSC03.Text), 2) = "11" Or _
       Left(Trim(textSC03.Text), 2) = "12" Or _
       Left(Trim(textSC03.Text), 2) = "13" Or _
       Left(Trim(textSC03.Text), 2) = "14") Then
      oMailCount = ""
      '有設定試用期滿通知人員時,才需要寄通知
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) >= 新部門啟用日 Then
         strSql = "SELECT A0926 as A0912 FROM Acc090NEW WHERE A0921='" & Left(Trim(textSC04_1), 3) & "' and A0926 is not null"
      Else
      '2023/12/20 END
         strSql = "SELECT A0912 FROM Acc090 WHERE A0901='" & Left(Trim(textSC04_1), 3) & "' and A0912 is not null"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modify By Sindy 2022/11/30 改抓特殊設定: Pub_GetSpecMan("試用期滿通知")
         'oMailCount = "99033;A4024;75033"
         oMailCount = Pub_GetSpecMan("試用期滿通知")
      Else
         'Add By Sindy 2024/2/1 增加：員工個人人事資料中之「智權專業證照」有內容者,也要發通知信
         strSql = "select * from staff_specialty where ss01='" & textSC01 & "' and ss04 is not null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            oMailCount = Pub_GetSpecMan("試用期滿通知")
         End If
         '2024/2/1 END
      End If
      If oMailCount <> "" Then
         strContent = "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
                      "異動原因：" + textSC03 & vbCrLf & _
                      "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf
         'Modify By Sindy 2022/5/13
         If Trim(textSC04_1) & Trim(textSC06_1) & Trim(textSC05_1) & Trim(textSC07_1) & Trim(textSC14_1) = Trim(textSC04) & Trim(textSC06) & Trim(textSC05) & Trim(textSC07) & Trim(textSC14) Then
            strContent = strContent & _
                      "部　門：" + textSC04_1 & vbCrLf & _
                      "職　位：" + textSC06_1 & vbCrLf & _
                      "職　稱：" + textSC05_1 & vbCrLf & _
                      "職稱說明：" + textSC07_1 & vbCrLf & _
                      "所　別：" + textSC14_1 & vbCrLf
         Else
            strContent = strContent & _
                      "部　門：" + textSC04 & vbCrLf & _
                      "職　位：" + textSC06 & vbCrLf & _
                      "職　稱：" + textSC05 & vbCrLf & _
                      "職稱說明：" + textSC07 & vbCrLf & _
                      "所　別：" + textSC14 & vbCrLf
         End If
         '2022/5/13 END
         '增加控管請假時,不發職代
         PUB_SendMail strUserNum, oMailCount, "", "請注意台一網頁同仁資料是否需要調整。", strContent, , , , , , , , , , True
      End If
   End If
   
   'Added by Lydia 2017/12/13 FCP案件命名電子化:命名人員離職(03離職、04留職停薪、08退休、09撤職、10資遣)時，若仍有命名未完成案件，系統要主動通知工程師主管。
   If strSrvDate(1) >= FCP案件命名啟用日 And Left(textSC04_1, 3) = "F21" _
        And (Left(Trim(textSC03.Text), 2) = "03" Or _
            Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "08" Or _
            Left(Trim(textSC03.Text), 2) = "09" Or _
            Left(Trim(textSC03.Text), 2) = "10") Then
        strSql = "select cp01,cp02,cp03,cp04,tct01,tct04,nvl(pa05,nvl(pa06,pa07)) pname from transcasetitle,caseprogress,patent " & _
                 "where tct10='" & textSC01 & "' and nvl(tct05,0)=0 and tct01=cp09(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) order by tct01 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        strExc(1) = "": strExc(2) = ""
        If intI = 1 Then
           RsTemp.MoveFirst
           strExc(2) = "" & RsTemp.Fields("tct04")
           Do While Not RsTemp.EOF
              strExc(1) = strExc(1) & RsTemp.Fields("cp01") & "-" & Val(RsTemp.Fields("cp02")) & IIf(RsTemp.Fields("cp03") & RsTemp.Fields("cp04") = "000", "", "-" & RsTemp.Fields("cp03") & "-" & Val(RsTemp.Fields("cp04"))) & vbTab & RsTemp.Fields("pname") & vbCrLf
              RsTemp.MoveNext
           Loop
           PUB_SendMail strUserNum, strExc(2), "", "人事異動:工程師" & textSC01_2 & "尚有新案命名未完成！", _
            "員工編號：" + textSC01 + " " + textSC01_2 & vbCrLf & _
            "異動原因：" + textSC03 & vbCrLf & _
            "異動日期：" + ChangeTStringToTDateString(textSC02) & vbCrLf & vbCrLf & _
            "未完成之新案命名：" & vbCrLf & strExc(1)
        End If
   End If
   'end 2017/12/13
   Call PUB_SendMailCache 'Added by Lydia 2025/08/14
   
   ShowCurrRecord strSC01, DBDATE(strSC02)
   Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
   AddRecord = True
   
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
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
   Dim strSC02 As String
   Dim strSC01 As String
   Dim MyArr As Variant
   
   ModRecord = False
   
   strSC01 = m_CurrKEY(0)
   strSC02 = m_CurrKEY(1)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE staff_change SET "

   bFirst = True
   bDifference = False
   For nIndex = 0 To tf_SC - 1
      strTmp = Empty
      '注意：此邊是由  0 開始   ==>nIndex=0 就是 sc01
      'If nIndex < 11 Or nIndex > 16 Then
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
                  "WHERE SC01 = '" & strSC01 & "' and SC02='" & strSC02 & "' ; end; "
On Error GoTo ErrHand

   cnnConnection.BeginTrans
   Screen.MousePointer = vbHourglass 'Add By Sindy 2014/3/24
   
   'Modify By Sindy 2024/5/27
   strExc(10) = PressCardData()
   If strExc(10) = "N" Then
      Screen.MousePointer = vbDefault
      Exit Function
   ElseIf strExc(10) = "Y" Then
      cnnConnection.BeginTrans
      Screen.MousePointer = vbHourglass
   End If
'      'Added Morgan 2013/8/2
'      If m_bolKillFingerPrinter = True Then
'         'Add By Sindy 2015/3/16 +if
'         If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
'            MsgBox "略過清除刷卡機資料程式,因會直接刪除「刷卡機」資料,請單獨測試"
'         Else
'         '2015/3/16 END
'            If PUB_ClearCardData(textSC01.Text) = False Then
'               Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
'               Exit Function
'            End If
'         End If
'      End If
'      'end 2013/8/3
   '2024/5/27 END
   
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      ' 2008/12/29 Add BY SINDY
      '增加判斷 異動日期<971102者都不要更新回員工基本檔.
      '因為 971102以前都是補資料.
      If Val(textSC02) >= 971102 Then
         '異動員工基本資料 ********************************
         strSql = ""
         '部門:
         'Modify By Sindy 2023/12/20
         If strSrvDate(1) >= 新部門啟用日 Then
            If textSC04.Text <> "" Then
                 MyArr = Split(textSC04, " ")
                 strSql = strSql & " st93='" & MyArr(0) & "' "
            Else
                 strSql = strSql & " st93=null "
            End If
         Else
         '2023/12/20 END
            If textSC04.Text <> "" Then
                 MyArr = Split(textSC04, " ")
                 strSql = strSql & " st03='" & MyArr(0) & "' "
            Else
                 strSql = strSql & " st03=null "
            End If
         End If
         '職稱:
         If textSC05.Text <> "" Then
              MyArr = Split(textSC05, " ")
              strSql = strSql & " ,st20='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st20=null "
         End If
         '職位:
         If textSC06.Text <> "" Then
              MyArr = Split(textSC06, " ")
              strSql = strSql & " ,st21='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st21=null "
         End If
         'Add By Sindy 2010/9/20
         '所別:
         If textSC14.Text <> "" Then
              MyArr = Split(textSC14, " ")
              strSql = strSql & " ,st06='" & MyArr(0) & "' "
         Else
              strSql = strSql & " ,st06=null "
         End If
         '2010/9/20 End
         '職稱說明:
         If textSC07.Text <> "" Then
              strSql = strSql & " ,st49='" & ChgSQL(textSC07) & "' "
         Else
              strSql = strSql & " ,st49=null "
         End If
         '離職日:
         '異動原因為02復職時，更新員工檔之離職日為null；
         If Left(Trim(textSC03.Text), 2) = "02" Then
              strSql = strSql & " ,st51=null "
         End If
         '異動原因為03離職、04留職停薪、08退休、09撤職、10資遣，更新員工檔之離職日為異動日期。
         If Left(Trim(textSC03.Text), 2) = "03" Or _
            Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "08" Or _
            Left(Trim(textSC03.Text), 2) = "09" Or _
            Left(Trim(textSC03.Text), 2) = "10" Then
              strSql = strSql & " ,st51='" & ChangeTStringToWString(Trim(textSC02)) & "' "
         End If
         If Left(Trim(strSql), 1) = "," Then
              strSql = Mid(Trim(strSql), 2)
         End If
         
         '2011/7/1 ADD BY SONIA 最後異動才做更新回員工檔
         strExc(0) = "select max(sc02) from staff_change where sc01='" & textSC01 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp(0) = DBDATE(textSC02) Then
         '2011/7/1 END
               Pub_SeekTbLog "update staff set " & strSql & " where st01='" & strSC01 & "' "
               cnnConnection.Execute "update staff set " & strSql & " where st01='" & strSC01 & "' "
               '異動員工基本資料 END ********************************
            End If
         End If
         '2011/7/1 ADD BY SONIA
         
         '異動薪資基本檔 ********************************
         '異動原因04.留職停薪同時更新薪資基本檔SalaryData之sd02編制為’S’。
         'Modify By Sindy 2010/6/25 異動原因為02復職時，sd02編制為’R’。
         If Left(Trim(textSC03.Text), 2) = "04" Or _
            Left(Trim(textSC03.Text), 2) = "02" Then
            strSql = ""
            If Left(Trim(textSC03.Text), 2) = "04" Then
               strSql = strSql & " sd02='S' "
            ElseIf Left(Trim(textSC03.Text), 2) = "02" Then
               strSql = strSql & " sd02='R' "
            End If
            If Left(Trim(strSql), 1) = "," Then
                 strSql = Mid(Trim(strSql), 2)
            End If
            Pub_SeekTbLog "update SalaryData set " & strSql & " where sd01='" & strSC01 & "' "
            cnnConnection.Execute "update SalaryData set " & strSql & " where sd01='" & strSC01 & "' "
         End If
         
         'Add by Morgan 2010/7/14 若職位有變動時重算婚喪扣款並更新回薪資檔
         'Removed by Morgan 2025/7/29 114/7/28起廢止婚喪互助辦法
         'If m_FieldList(5).fiOldData <> m_FieldList(5).fiNewData Then
         '   strExc(0) = "select max(sc02) from staff_change where sc01='" & textSC01 & "'"
         '   intI = 1
         '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         '   If intI = 1 Then
         '      '最後異動才做
         '      If RsTemp(0) = DBDATE(textSC02) Then
         '         If PUB_GetHelpFee(textSC01, strExc(1)) = True Then
         '            strSql = "update salarydata set sd09=decode(nvl(sd09,0),0,sd09," & Val(strExc(1)) & "),sd10=decode(nvl(sd10,0),0,sd10," & Val(strExc(1)) & ") where sd01='" & textSC01 & "'"
         '            Pub_SeekTbLog strSql
         '            cnnConnection.Execute strSql
         '         End If
         '      End If
         '   End If
         'End If
         'end 2025/7/29
         
         '異動薪資基本檔 END *****************************
      End If
   End If
   
   cnnConnection.CommitTrans

   ShowCurrRecord strSC01, DBDATE(strSC02)
   Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24

   ModRecord = True
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault 'Add By Sindy 2014/3/24
   cnnConnection.RollbackTrans
   MsgBox (Err.Description)
End Function

' 刪除記錄
Private Function DelRecord() As Boolean
   Dim strSql As String
   Dim strSC02 As String
   Dim strSC01 As String
   Dim MyArr As Variant
   
   DelRecord = False
   
On Error GoTo ErrHand
   
   cnnConnection.BeginTrans
   
   strSC01 = m_CurrKEY(0)
   strSC02 = m_CurrKEY(1)

   strSql = "DELETE FROM staff_change " & _
            "WHERE SC01 = '" & strSC01 & "'  and SC02='" & strSC02 & "' "
   
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   If (strSC01 = m_LastKEY(0) And strSC02 = m_LastKEY(1)) Or (strSC01 = m_FirstKEY(0) And strSC02 = m_FirstKEY(1)) Then
      RefreshRange
   End If
   ShowCurrRecord strSC01, DBDATE(strSC02)
   DelRecord = True
   cnnConnection.CommitTrans
   
   '2010/3/22 add by sonia
   MsgBox "員工檔資料無法自動還原，請人工修改或通知電腦中心更正！", vbExclamation
   '2010/3/22 end
   
   Exit Function
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "刪除失敗！" & vbCrLf & Err.Description
    
End Function

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strSC02 As String
   Dim strSC01 As String
   
   QueryRecord = False
   strSC02 = DBDATE(textSC02)
   strSC01 = textSC01
   If IsRecordExist(strSC01, strSC02) = True Then
      m_CurrKEY(0) = strSC01
      m_CurrKEY(1) = strSC02
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
         If textSC01 <> "" And textSC02 <> "" Then
            If QueryRecord = False Then
               strMsg = "無此資料"
               strTit = "查詢資料"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               UpdateCtrlData
            End If
         Else
            ' 2008/12/17 ADD BY SINDY
            If textSC01 = "" Or textSC02 = "" Then
               MsgBox "須輸入員工代號及異動日期才可進行查詢動作！", vbInformation
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
      Case 1: If Me.Visible = True Then textSC01.SetFocus
      Case 2: If Me.Visible = True Then textSC03.SetFocus
      Case 4: If Me.Visible = True Then textSC01.SetFocus
   End Select
End Sub

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String, ByVal strKEY02 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM staff_change " & _
            "WHERE SC01 = '" & strKEY01 & "'  and SC02='" & strKEY02 & "'  "
                  
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
      strSql = "SELECT SC01,SC02 FROM staff_change " & _
               "WHERE SC01 = '" & m_CurrKEY(0) & "' and SC02='" & m_CurrKEY(1) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
         If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT SC01,SC02 FROM staff_change " & _
               "WHERE SC02 = (SELECT MIN(SC02) FROM staff_change where SC01=(select min(SC01) from staff_change) ) and SC01=(select min(SC01) from staff_change) "
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
         If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
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
   
   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SC02 = (SELECT MAX(SC02) FROM staff_change " & _
                          "WHERE SC01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SC02 < '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = (SELECT MAX(SC01) FROM staff_change " & _
                           "WHERE SC01 < '" & m_CurrKEY(0) & "') AND " & _
                  "SC02 = (SELECT MAX(SC02) FROM staff_change " & _
                           "WHERE SC01 = (SELECT MAX(SC01) FROM staff_change " & _
                                          "WHERE SC01 < '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
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
   
   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = '" & m_CurrKEY(0) & "' AND " & _
                  "SC02 = (SELECT MIN(SC02) FROM staff_change " & _
                          "WHERE SC01 = '" & m_CurrKEY(0) & "' AND " & _
                                "SC02 > '" & m_CurrKEY(1) & "' )"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
      rsTmp.Close
      UpdateCtrlData
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = (SELECT MIN(SC01) FROM staff_change " & _
                           "WHERE SC01 > '" & m_CurrKEY(0) & "') AND " & _
                  "SC02 = (SELECT MIN(SC02) FROM staff_change " & _
                           "WHERE SC01 = (SELECT MIN(SC01) FROM staff_change " & _
                                          "WHERE SC01 > '" & m_CurrKEY(0) & "')) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_CurrKEY(1) = rsTmp.Fields("SC02")
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
   
   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = (SELECT MIN(SC01) FROM staff_change) AND " & _
                  "SC02 = (SELECT MIN(SC02) FROM staff_change " & _
                           "WHERE SC01 = (SELECT MIN(SC01) FROM staff_change)) "
                           
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_FirstKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_FirstKEY(1) = rsTmp.Fields("SC02")
   End If
   rsTmp.Close

   strSql = "SELECT SC01,SC02 FROM staff_change " & _
            "WHERE SC01 = (SELECT MAX(SC01) FROM staff_change) AND " & _
                  "SC02 = (SELECT MAX(SC02) FROM staff_change " & _
                           "WHERE SC01 = (SELECT MAX(SC01) FROM staff_change)) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("SC01")) = False Then: m_LastKEY(0) = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: m_LastKEY(1) = rsTmp.Fields("SC02")
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim i As Integer, j As Integer
   Dim strSC02Date As String
         
   strSql = "SELECT * FROM staff_change " & _
            "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 = '" & m_CurrKEY(1) & "'   "
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      ClearField
      If IsNull(rsTmp.Fields("SC01")) = False Then: textSC01 = rsTmp.Fields("SC01")
      If IsNull(rsTmp.Fields("SC02")) = False Then: textSC02 = TAIWANDATE(rsTmp.Fields("SC02"))
      If IsNull(rsTmp.Fields("sc03")) = False Then: textSC03 = rsTmp.Fields("sc03")
      If IsNull(rsTmp.Fields("sc04")) = False Then: textSC04 = rsTmp.Fields("sc04")
      If IsNull(rsTmp.Fields("sc05")) = False Then: textSC05 = rsTmp.Fields("sc05")
      If IsNull(rsTmp.Fields("sc06")) = False Then: textSC06 = rsTmp.Fields("sc06")
      If IsNull(rsTmp.Fields("sc07")) = False Then: textSC07 = rsTmp.Fields("sc07")
      If IsNull(rsTmp.Fields("sc14")) = False Then: textSC14 = rsTmp.Fields("sc14") 'Add By Sindy 2010/9/20
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
      
      textSC01_2 = GetStaffName(textSC01, True)
      textSC03_Validate False
      textSC04_Validate False
      textSC05_Validate False
      textSC06_Validate False
      textSC14_Validate False 'Add By Sindy 2010/9/20
      
      ' 2008/12/23 Add BY SINDY
      ' 查詢時舊欄位資料抓此員工編號的前一筆異動資料
'      strSql = "SELECT a1.a0901||' '||a1.a0902,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07,sc14 " & _
'                 "FROM staff_change,acc090 a1,allcode a2,allcode a3 " & _
'               "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 = " & _
'                     "(SELECT max(SC02) FROM staff_change " & _
'                     "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 < '" & m_CurrKEY(1) & "') " & _
'                     "and SC04=a1.a0901(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
      strSql = "SELECT SC02 " & _
                 "FROM staff_change,allcode a2,allcode a3 " & _
               "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 = " & _
                     "(SELECT max(SC02) FROM staff_change " & _
                     "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 < '" & m_CurrKEY(1) & "') " & _
                     "and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
      rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         'Modify By Sindy 2023/12/20
         strSC02Date = rsTmp.Fields(0) '異動日期
         If strSC02Date >= 新部門啟用日 Then
            strSql = "SELECT a1.a0921||' '||a1.a0922,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07,sc14 " & _
                       "FROM staff_change,acc090NEW a1,allcode a2,allcode a3 " & _
                     "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 = " & strSC02Date & _
                          " and SC04=a1.a0921(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
         Else
            strSql = "SELECT a1.a0901||' '||a1.a0902,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07,sc14 " & _
                       "FROM staff_change,acc090 a1,allcode a2,allcode a3 " & _
                     "WHERE SC01='" & m_CurrKEY(0) & "' and SC02 = " & strSC02Date & _
                          " and SC04=a1.a0901(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
         End If
         rsTmp.Close
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
         '2023/12/20 END
            If IsNull(rsTmp.Fields(0)) = False Then: textSC04_1 = rsTmp.Fields(0)
            If IsNull(rsTmp.Fields(1)) = False Then: textSC05_1 = rsTmp.Fields(1)
            If IsNull(rsTmp.Fields(2)) = False Then: textSC06_1 = rsTmp.Fields(2)
            If IsNull(rsTmp.Fields(3)) = False Then: textSC07_1 = rsTmp.Fields(3)
            If IsNull(rsTmp.Fields(4)) = False Then: textSC14_1 = rsTmp.Fields(4) 'Add By Sindy 2010/9/20
            textSC14_1_Validate False 'Add By Sindy 2023/12/22
         End If
      End If
      ' 2008/12/23 END
   End If
   
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Sub GetData()
Dim rsTmp As New ADODB.Recordset
Dim strSQL2 As String

strSql = ""
If txt1(0) <> "" Then
    strSql = strSql & " and SC01>='" & txt1(0) & "' "
End If
If txt1(1) <> "" Then
    strSql = strSql & " and SC01<='" & txt1(1) & "' "
End If
If txt1(2) <> "" Then
    strSql = strSql & " and SC02>='" & DBDATE(txt1(2)) & "' "
End If
If txt1(3) <> "" Then
    strSql = strSql & " and SC02<='" & DBDATE(txt1(3)) & "' "
End If
'Add By Sindy 2023/2/24
If textSC03_Q <> "" Then
    strSql = strSql & " and SC03='" & Left(textSC03_Q.Text, 2) & "' "
End If
'2023/2/24 END

'抓取資料
' 2008/12/22 Modify BY SINDY
'strSQL = "SELECT SC01,st02,sqldateT(SC02),a1.ac02||' '||a1.ac03,a2.a0901||' '||a2.a0902,sc07,a3.ac02||' '||a3.ac03,a4.ac02||' '||a4.ac03,a5.a0901||' '||a5.a0902,a6.ac02||' '||a6.ac03,a7.ac02||' '||a7.ac03 FROM staff_change,staff,allcode a1,acc090 a2,allcode a3,allcode a4,acc090 a5,allcode a6,allcode a7 where SC01=st01(+) and '05'=a1.ac01(+) and sc03=a1.ac02(+) and sc04=a2.a0901(+) and '01'=a3.ac01(+) and sc05=a3.ac02(+) and '02'=a4.ac01(+) and sc06=a4.ac02(+) and sc08=a5.a0901(+) and '01'=a6.ac01(+) and sc09=a6.ac02(+) and '02'=a7.ac01(+) and sc10=a7.ac02(+) " & strSQL & _
'        " order by SC01,SC02 "
'Modify By Sindy 2023/12/22
If strSrvDate(1) >= 新部門啟用日 Then
   strSQL2 = "SELECT SC01,st02,sqldateT(SC02),a4.ac02||' '||a4.ac03,a1.a0901||' '||a1.a0902,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07,SC02 " & _
            "FROM staff_change,staff,acc090 a1,allcode a2,allcode a3,allcode a4 " & _
            "WHERE SC01=st01(+) and '05'=a4.ac01(+) and SC03=a4.ac02(+) " & _
              "and SC04=a1.a0901(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) " & strSql & _
              "and SC02<" & 新部門啟用日 & " " & _
            "union all SELECT SC01,st02,sqldateT(SC02),a4.ac02||' '||a4.ac03,a1.a0921||' '||a1.a0922,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07,SC02 " & _
            "FROM staff_change,staff,acc090NEW a1,allcode a2,allcode a3,allcode a4 " & _
            "WHERE SC01=st01(+) and '05'=a4.ac01(+) and SC03=a4.ac02(+) " & _
              "and SC04=a1.a0921(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) " & strSql & _
              "and SC02>=" & 新部門啟用日 & " " & _
            " order by SC02,SC01 "
Else
'2023/12/22 END
   strSQL2 = "SELECT SC01,st02,sqldateT(SC02),a4.ac02||' '||a4.ac03,a1.a0901||' '||a1.a0902,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,sc07 " & _
            "FROM staff_change,staff,acc090 a1,allcode a2,allcode a3,allcode a4 " & _
            "WHERE SC01=st01(+) and '05'=a4.ac01(+) and SC03=a4.ac02(+) " & _
              "and SC04=a1.a0901(+) and '01'=a2.ac01(+) and SC05=a2.ac02(+) and '02'=a3.ac01(+) and SC06=a3.ac02(+) " & strSql & _
            " order by SC02,SC01 "
End If
' 2008/12/22 END
If rsTmp.State = 1 Then rsTmp.Close
rsTmp.CursorLocation = adUseClient
rsTmp.Open strSQL2, cnnConnection, adOpenStatic, adLockReadOnly
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
   textSC01_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC02_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC03_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC04_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC05_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC06_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   nResponse = False
   textSC07_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   'Add By Sindy 2010/9/20
   nResponse = False
   textSC14_Validate nResponse
   If nResponse = True Then GoTo EXITSUB
   '2010/9/20 End
   
   CheckDataValid = True
EXITSUB:
End Function

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textSC01.Locked = bEnable
   textSC02.Locked = bEnable
   If bEnable Then textSC01.BackColor = &H8000000F Else textSC01.BackColor = &H80000005
   If bEnable Then textSC02.BackColor = &H8000000F Else textSC02.BackColor = &H80000005
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
Dim i As Integer
   textSC01.Locked = bEnable
   textSC02.Locked = bEnable
   If bEnable Then textSC01.BackColor = &H8000000F Else textSC01.BackColor = &H80000005
   If bEnable Then textSC02.BackColor = &H8000000F Else textSC02.BackColor = &H80000005
   textSC03.Enabled = Not bEnable
   textSC04.Enabled = Not bEnable
   textSC05.Enabled = Not bEnable
   textSC06.Enabled = Not bEnable
   textSC07.Enabled = Not bEnable
   textSC14.Enabled = Not bEnable 'Add By Sindy 2010/9/20
   textSC04_1.Enabled = False
   textSC05_1.Enabled = False
   textSC06_1.Enabled = False
   textSC07_1.Enabled = False
   textSC14_1.Enabled = False 'Add By Sindy 2010/9/20
   If bEnable Then textSC04_1.BackColor = &H8000000F Else textSC04_1.BackColor = &H80000005
   If bEnable Then textSC05_1.BackColor = &H8000000F Else textSC05_1.BackColor = &H80000005
   If bEnable Then textSC06_1.BackColor = &H8000000F Else textSC06_1.BackColor = &H80000005
   If bEnable Then textSC07_1.BackColor = &H8000000F Else textSC07_1.BackColor = &H80000005
   If bEnable Then textSC14_1.BackColor = &H8000000F Else textSC14_1.BackColor = &H80000005 'Add By Sindy 2010/9/20
End Sub

Private Sub ClearField()
   Dim nIndex As Integer
   textSC01 = Empty
   textSC01_2 = Empty
   textSC02 = Empty
   textSC03 = Empty
   textSC04 = Empty
   textSC05 = Empty
   textSC06 = Empty
   textSC07 = Empty
   textSC14 = Empty 'Add By Sindy 2010/9/20
   textSC04_1 = Empty
   textSC05_1 = Empty
   textSC06_1 = Empty
   textSC07_1 = Empty
   textSC14_1 = Empty 'Add By Sindy 2010/9/20
   Label23 = Empty
   SetGrd
   For nIndex = 0 To tf_SC - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
End Sub

Private Sub UpdateFieldNewData()
    Dim MyArr As Variant
   '若新增資料
   If m_EditMode = 1 Then
      SetFieldNewData "SC01", textSC01
      SetFieldNewData "SC02", DBDATE(textSC02)
   End If
   If textSC03.Text <> "" Then
        MyArr = Split(textSC03, " ")
        SetFieldNewData "SC03", MyArr(0)
   Else
        SetFieldNewData "SC03", Empty
   End If
   If textSC04.Text <> "" Then
        MyArr = Split(textSC04, " ")
        SetFieldNewData "SC04", MyArr(0)
   Else
        SetFieldNewData "SC04", Empty
   End If
   If textSC05.Text <> "" Then
        MyArr = Split(textSC05, " ")
        SetFieldNewData "SC05", MyArr(0)
   Else
        SetFieldNewData "SC05", Empty
   End If
   If textSC06.Text <> "" Then
        MyArr = Split(textSC06, " ")
        SetFieldNewData "SC06", MyArr(0)
   Else
        SetFieldNewData "SC06", Empty
   End If
   SetFieldNewData "SC07", ChgSQL(textSC07)
   'Add By Sindy 2010/9/20
   If textSC14.Text <> "" Then
        MyArr = Split(textSC14, " ")
        SetFieldNewData "SC14", MyArr(0)
   Else
        SetFieldNewData "SC14", Empty
   End If
   '2010/9/20 End
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   ' 初始化欄位陣列
   For nIndex = 1 To tf_SC
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "SC" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 2:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

'帶預設資料
Private Sub InitialData()
Dim MyRs As New ADODB.Recordset

textSC03.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='05' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC03.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If
'Add By Sindy 2023/2/24
textSC03_Q.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='05' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC03_Q.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If
'2023/2/24 END

textSC04.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
'2009/3/2 modify by sonia
'strSQL = "select a0901||' '||a0902 from acc090 order by a0901"
'Modify By Sindy 2023/12/20
If strSrvDate(1) >= 新部門啟用日 Then
   strSql = "select a0921||' '||a0922 from acc090NEW order by a0921"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
      While Not MyRs.EOF
         textSC04.AddItem "" & MyRs.Fields(0).Value
         MyRs.MoveNext
      Wend
   End If
Else
'2023/12/20 END
   strSql = "select a0901||' '||a0902 from acc090 where a0904<>'Y' and a0901<>'CFL' order by a0901"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
      While Not MyRs.EOF
         textSC04.AddItem "" & MyRs.Fields(0).Value
         MyRs.MoveNext
      Wend
   End If
End If

textSC04_1.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
'2009/3/2 modify by sonia
'strSQL = "select a0901||' '||a0902 from acc090 order by a0901"
'Modify By Sindy 2023/12/20
If strSrvDate(1) >= 新部門啟用日 Then
   strSql = "select a0921||' '||a0922 from acc090NEW order by a0921"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
      While Not MyRs.EOF
         textSC04_1.AddItem "" & MyRs.Fields(0).Value
         MyRs.MoveNext
      Wend
   End If
Else
'2023/12/20 END
   strSql = "select a0901||' '||a0902 from acc090 where a0904<>'Y' and a0901<>'CFL' order by a0901"
   MyRs.CursorLocation = adUseClient
   MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If MyRs.RecordCount <> 0 Then
      While Not MyRs.EOF
         textSC04_1.AddItem "" & MyRs.Fields(0).Value
         MyRs.MoveNext
      Wend
   End If
End If

textSC05.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='01' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC05.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If

textSC05_1.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='01' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC05_1.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If

textSC06.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='02' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC06.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If

textSC06_1.Clear
Set MyRs = New ADODB.Recordset
If MyRs.State = 1 Then MyRs.Close
strSql = "select ac02||' '||ac03 from allcode where ac01='02' order by ac02"
MyRs.CursorLocation = adUseClient
MyRs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If MyRs.RecordCount <> 0 Then
    While Not MyRs.EOF
        textSC06_1.AddItem "" & MyRs.Fields(0).Value
        MyRs.MoveNext
    Wend
End If

SetGrd
End Sub

Private Sub textSC02_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC02
End If
End Sub

Private Sub textSC02_KeyPress(KeyAscii As Integer)
KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textSC02_Validate(Cancel As Boolean)
If m_EditMode = 1 And textSC02 <> "" Then
    If IsRecordExist(textSC01, DBDATE(textSC02)) = True And textSC02.Enabled = True And textSC02.Locked = False Then
        MsgBox "該員工當天已有資料，請修改！", vbInformation
        Call textSC02_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
    If CheckIsTaiwanDate(textSC02, False) = False Then
        Cancel = True
        MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
        Call textSC02_GotFocus ' 2008/12/23 Add BY SINDY
        Exit Sub
    End If
'    If ChkWorkDay(DBDATE(textSC02)) = False Then
'        Cancel = True
'        MsgBox "請輸入工作天！", vbInformation, "輸入日期錯誤"
'        Call textSC02_GotFocus ' 2008/12/23 Add BY SINDY
'        Exit Sub
'    End If
End If
End Sub

Private Sub textSC01_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC01
    CloseIme
End If
End Sub

Private Sub textSC01_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSC01_Validate(Cancel As Boolean)
If textSC01.Text = "" Then textSC01_2 = "" ' 2008/12/18 ADD BY SINDY

If m_EditMode = 1 And textSC01 <> "" Then
     textSC01_2 = GetStaffName(textSC01, True)
    If IsRecordExist(textSC01, DBDATE(textSC02)) = True And textSC01.Enabled = True And textSC01.Locked = False Then
        MsgBox "該員工當天已有資料，請修改！", vbInformation
        Call textSC01_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
    If textSC01_2 = "" Then
        MsgBox "員工編號錯誤！查無此員工！", vbInformation
        Call textSC01_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
    
    ' 2008/12/18 ADD BY SINDY
    ' 檢查員工編號規則
    If ChkStaffID(textSC01) Then
       Call textSC01_GotFocus
       Cancel = True
       Exit Sub
    End If
    ' 2008/12/18 END
   
   '當輸入員工代號後, 顯示員工基本資料 ************************
   'Modify By Sindy 2024/5/27 改判斷員工編號不同就要重新抓資料
'   If textSC04.Text = "" And textSC05.Text = "" And _
'         textSC06.Text = "" And textSC07.Text = "" And _
'         textSC14.Text = "" Then
   If textSC01.Text <> textSC01.Tag Then
      textSC01.Tag = textSC01.Text
   '2024/5/27 END
      'strSQL = "select * from staff where st01='" & textSC01 & "' "
      'Modify By Sindy 2023/12/20
      If strSrvDate(1) >= 新部門啟用日 Then
         strSql = "SELECT a1.a0921||' '||a1.a0922,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,st49,st06,st29 " & _
                          "FROM staff,acc090NEW a1,allcode a2,allcode a3 " & _
                 "WHERE ST01='" & textSC01 & "' " & _
                 "and st93=a1.a0921(+) and '01'=a2.ac01(+) and st20=a2.ac02(+) and '02'=a3.ac01(+) and st21=a3.ac02(+) "
      Else
      '2023/12/20 END
         strSql = "SELECT a1.a0901||' '||a1.a0902,a2.ac02||' '||a2.ac03,a3.ac02||' '||a3.ac03,st49,st06,st29 " & _
                          "FROM staff,acc090 a1,allcode a2,allcode a3 " & _
                 "WHERE ST01='" & textSC01 & "' " & _
                 "and st03=a1.a0901(+) and '01'=a2.ac01(+) and st20=a2.ac02(+) and '02'=a3.ac01(+) and st21=a3.ac02(+) "
      End If
      CheckOC3
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      m_ST29 = ""
      If AdoRecordSet3.RecordCount <> 0 Then
          textSC04.Text = CheckStr(AdoRecordSet3.Fields(0))
          textSC05.Text = CheckStr(AdoRecordSet3.Fields(1))
          textSC06.Text = CheckStr(AdoRecordSet3.Fields(2))
          textSC07.Text = CheckStr(AdoRecordSet3.Fields(3))
          textSC14.Text = CheckStr(AdoRecordSet3.Fields(4)) 'Add By Sindy 2010/9/20
          textSC04_1.Text = CheckStr(AdoRecordSet3.Fields(0))
          textSC05_1.Text = CheckStr(AdoRecordSet3.Fields(1))
          textSC06_1.Text = CheckStr(AdoRecordSet3.Fields(2))
          textSC07_1.Text = CheckStr(AdoRecordSet3.Fields(3))
          textSC14_1.Text = CheckStr(AdoRecordSet3.Fields(4)) 'Add By Sindy 2010/9/20
          m_ST29 = CheckStr(AdoRecordSet3.Fields("st29")) 'Add By Sindy 2011/10/26
      End If
      textSC04_Validate False
      textSC05_Validate False
      textSC06_Validate False
      textSC14_Validate False 'Add By Sindy 2010/9/20
      textSC14_1_Validate False 'Add By Sindy 2010/9/20
   End If
   '當輸入員工代號後, 顯示基本資料 END************************
End If
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   arrGridHeadText = Array("員工編號", "姓名", "異動日", "異動原因", "新部門", "新職稱", "新職位", "新職稱說明")
   arrGridHeadWidth = Array(800, 800, 800, 800, 1200, 1200, 1200, 1200)
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

Private Sub textSC03_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC03
End If
End Sub

Private Sub textSC03_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSC03_Validate(Cancel As Boolean)
If textSC03.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    
    MyArr = Split(textSC03, " ")
    For Myi = 0 To textSC03.ListCount - 1
        MyArr2 = Split(textSC03.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC03.Text = textSC03.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "異動原因代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC03_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textSC04_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC04
End If
End Sub

Private Sub textSC04_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSC04_Validate(Cancel As Boolean)
If textSC04.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(textSC04, " ")
    For Myi = 0 To textSC04.ListCount - 1
        MyArr2 = Split(textSC04.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC04.Text = textSC04.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "部門代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC04_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textSC05_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC05
End If
End Sub

Private Sub textSC05_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSC05_Validate(Cancel As Boolean)
If textSC05.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(textSC05, " ")
    For Myi = 0 To textSC05.ListCount - 1
        MyArr2 = Split(textSC05.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC05.Text = textSC05.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "職稱代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC05_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textSC06_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC06
End If
End Sub

Private Sub textSC06_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textSC06_Validate(Cancel As Boolean)
If textSC06.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(textSC06, " ")
    For Myi = 0 To textSC06.ListCount - 1
        MyArr2 = Split(textSC06.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC06.Text = textSC06.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "職位代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC06_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
End If
End Sub

Private Sub textSC07_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC07
    OpenIme
End If
End Sub

Private Sub textSC07_Validate(Cancel As Boolean)
If m_EditMode <> 0 And textSC07 <> "" Then
    If CheckLengthIsOK(textSC07, textSC07.MaxLength) = False Then
        Call textSC07_GotFocus ' 2008/12/23 Add BY SINDY
        Cancel = True
        Exit Sub
    End If
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

'Add By Sindy 2010/9/20
Private Sub textSC14_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC14
End If
End Sub

'Add By Sindy 2010/9/20
Private Sub textSC14_Validate(Cancel As Boolean)
If textSC14.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(textSC14, " ")
    For Myi = 0 To textSC14.ListCount - 1
        MyArr2 = Split(textSC14.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC14.Text = textSC14.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "所別代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC14_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
End Sub

'Add By Sindy 2010/9/20
Private Sub textSC14_1_GotFocus()
If m_EditMode <> 0 Then
    InverseTextBox textSC14_1
End If
End Sub

'Add By Sindy 2010/9/20
Private Sub textSC14_1_Validate(Cancel As Boolean)
If textSC14_1.Text <> "" Then
    Dim MyArr As Variant
    Dim MyArr2 As Variant
    Dim Myi As Integer
    MyArr = Split(textSC14_1, " ")
    For Myi = 0 To textSC14_1.ListCount - 1
        MyArr2 = Split(textSC14_1.List(Myi), " ")
        If MyArr(0) = MyArr2(0) Then
            textSC14_1.Text = textSC14_1.List(Myi)
            Exit Sub
        End If
    Next Myi
    If m_EditMode <> 0 Then
        MsgBox "所別代號輸入錯誤!!!", vbExclamation + vbOKOnly
        Call textSC14_1_GotFocus
        Cancel = True
        Exit Sub
    End If
End If
End Sub

