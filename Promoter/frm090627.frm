VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090627 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人特殊案件記錄維護"
   ClientHeight    =   5460
   ClientLeft      =   1752
   ClientTop       =   1860
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   9144
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8610
      Top             =   630
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
            Picture         =   "frm090627.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm090627.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4605
      Left            =   180
      TabIndex        =   12
      Top             =   720
      Width           =   8775
      _ExtentX        =   15473
      _ExtentY        =   8128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm090627.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(55)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblOurCaseNo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblCountry"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(4)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblCaseName"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblCaseProperty"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblPromoter"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSCR(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtSCR(1)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Check1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "多筆查詢"
      TabPicture(1)   =   "frm090627.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtCode(0)"
      Tab(1).Control(1)=   "txtCode(3)"
      Tab(1).Control(2)=   "txtCode(2)"
      Tab(1).Control(3)=   "txtCode(1)"
      Tab(1).Control(4)=   "cmdQuery(1)"
      Tab(1).Control(5)=   "txtSCR(13)"
      Tab(1).Control(6)=   "txtSCR(12)"
      Tab(1).Control(7)=   "txtSCR(11)"
      Tab(1).Control(8)=   "txtSCR(10)"
      Tab(1).Control(9)=   "cmdQuery(0)"
      Tab(1).Control(10)=   "grdList"
      Tab(1).Control(11)=   "Line1"
      Tab(1).Control(12)=   "Label1(12)"
      Tab(1).Control(13)=   "Label1(7)"
      Tab(1).Control(14)=   "Line3"
      Tab(1).Control(15)=   "Label1(5)"
      Tab(1).Control(16)=   "Line2"
      Tab(1).ControlCount=   17
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   -74010
         MaxLength       =   3
         TabIndex        =   7
         Top             =   810
         Width           =   585
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   3
         Left            =   -71940
         MaxLength       =   2
         TabIndex        =   10
         Top             =   810
         Width           =   492
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   -72330
         MaxLength       =   1
         TabIndex        =   9
         Top             =   810
         Width           =   315
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   -73350
         MaxLength       =   6
         TabIndex        =   8
         Top             =   810
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "列印(&P)"
         Height          =   400
         Index           =   1
         Left            =   -67470
         TabIndex        =   26
         Top             =   390
         Width           =   912
      End
      Begin VB.CheckBox Check1 
         Caption         =   "主管核可"
         Height          =   315
         Left            =   540
         TabIndex        =   2
         Top             =   3390
         Width           =   1785
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   1
         Left            =   1410
         MaxLength       =   4
         TabIndex        =   1
         Top             =   2850
         Width           =   945
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   0
         Left            =   1410
         MaxLength       =   9
         TabIndex        =   0
         Top             =   570
         Width           =   1125
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   13
         Left            =   -70050
         MaxLength       =   3
         TabIndex        =   6
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   12
         Left            =   -70680
         MaxLength       =   3
         TabIndex        =   5
         Top             =   450
         Width           =   525
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   11
         Left            =   -73050
         MaxLength       =   7
         TabIndex        =   4
         Top             =   450
         Width           =   945
      End
      Begin VB.TextBox txtSCR 
         Height          =   300
         Index           =   10
         Left            =   -74100
         MaxLength       =   7
         TabIndex        =   3
         Top             =   450
         Width           =   945
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "查詢(&Q)"
         Default         =   -1  'True
         Height          =   400
         Index           =   0
         Left            =   -68460
         TabIndex        =   16
         Top             =   390
         Width           =   912
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   3150
         Left            =   -74850
         TabIndex        =   32
         Top             =   1260
         Width           =   8490
         _ExtentX        =   14965
         _ExtentY        =   5546
         _Version        =   393216
         Cols            =   9
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
         _Band(0).Cols   =   9
      End
      Begin VB.Line Line1 
         X1              =   -73680
         X2              =   -71850
         Y1              =   930
         Y2              =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   12
         Left            =   -74820
         TabIndex        =   33
         Top             =   870
         Width           =   765
      End
      Begin MSForms.Label lblPromoter 
         Height          =   255
         Left            =   1440
         TabIndex        =   31
         Top             =   2250
         Width           =   1845
         VariousPropertyBits=   27
         Caption         =   "lblPromoter"
         Size            =   "3254;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseProperty 
         Height          =   255
         Left            =   1440
         TabIndex        =   30
         Top             =   1926
         Width           =   1845
         VariousPropertyBits=   27
         Caption         =   "lblCaseProperty"
         Size            =   "3254;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblCaseName 
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         Top             =   1282
         Width           =   6915
         VariousPropertyBits=   27
         Caption         =   "lblCaseName"
         Size            =   "12197;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label3 
         Height          =   255
         Left            =   510
         TabIndex        =   28
         Top             =   3900
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Create :"
         Size            =   "5821;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label4 
         Height          =   255
         Left            =   3990
         TabIndex        =   27
         Top             =   3900
         Width           =   3300
         VariousPropertyBits=   27
         Caption         =   "lblFM2-Update :"
         Size            =   "5821;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "完稿日:"
         Height          =   180
         Index           =   7
         Left            =   -74760
         TabIndex        =   25
         Top             =   510
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱:"
         Height          =   180
         Index           =   4
         Left            =   510
         TabIndex        =   24
         Top             =   1282
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人:"
         Height          =   180
         Index           =   6
         Left            =   510
         TabIndex        =   23
         Top             =   2250
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件性質:"
         Height          =   180
         Index           =   3
         Left            =   510
         TabIndex        =   22
         Top             =   1926
         Width           =   765
      End
      Begin VB.Label lblCountry 
         Caption         =   "lblCountry"
         Height          =   255
         Left            =   1440
         TabIndex        =   21
         Top             =   1604
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家:"
         Height          =   180
         Index           =   2
         Left            =   510
         TabIndex        =   20
         Top             =   1604
         Width           =   765
      End
      Begin VB.Label lblOurCaseNo 
         Caption         =   "lblOurCaseNo"
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   960
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號:"
         Height          =   180
         Index           =   0
         Left            =   510
         TabIndex        =   18
         Top             =   960
         Width           =   765
      End
      Begin VB.Line Line3 
         X1              =   -70380
         X2              =   -69810
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人部門:"
         Height          =   180
         Index           =   5
         Left            =   -71670
         TabIndex        =   17
         Top             =   480
         Width           =   945
      End
      Begin VB.Line Line2 
         X1              =   -73320
         X2              =   -72900
         Y1              =   570
         Y2              =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "收文號:"
         Height          =   180
         Index           =   55
         Left            =   510
         TabIndex        =   15
         Top             =   630
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "加值件數:"
         Height          =   180
         Index           =   1
         Left            =   510
         TabIndex        =   14
         Top             =   2910
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "( 0 < 加值件數 < 10) "
         Height          =   180
         Left            =   2430
         TabIndex        =   13
         Top             =   2910
         Width           =   1560
      End
   End
   Begin MSComctlLib.Toolbar TBar1 
      Align           =   1  '對齊表單上方
      Height          =   528
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
      _ExtentY        =   931
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
End
Attribute VB_Name = "frm090627"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Lydia 2022/01/03 改成Form2.0 ; grdList改字型=新細明體-ExtB(MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉)、Label3、Label4、lblCaseName、lblCaseProperty、lblPromoter; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

'Dim SCR(0 To 3) As String
Dim SCR(0 To 9) As String
Dim strRsStart1 As String, strRsStart2 As String, strRsStart3 As String, strRsStart4 As String, strRsEnd1 As String, strRsEnd2 As String, strRsEnd3 As String, strRsEnd4 As String
Dim rsDefineSize As New ADODB.Recordset
Dim intWhere As Integer
Dim ActionEdit As Integer
Dim intRow As Integer
'執行各項功能的權限
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_CurrSel As Integer
Dim PLeft(0 To 7) As Integer
Dim iPrint As Integer
Dim Page As Integer
Dim m_blnColOrderAsc As Boolean 'Added by Lydia 2016/08/11 欄位資料由小到大排序
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限

Public Sub SelectToolbarButtom()
Dim btn
    '設定為按下查詢鈕扭
    Set btn = Me.TBar1.Buttons(4)
    Tbar1_ButtonClick btn
End Sub

Private Sub Check1_Click()
    '若從個人進入, 若已核可的資料其他欄位不可更改
    If ProState = "1" Then
        If Me.Check1.Value = vbChecked Then
            Me.txtSCR(0).Enabled = False
            Me.txtSCR(1).Enabled = False
        Else
            Me.txtSCR(0).Enabled = True
            Me.txtSCR(1).Enabled = True
        End If
    End If
End Sub

Private Sub cmdQuery_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    
    If Me.SSTab1.Tab <> 1 Then Exit Sub
    If Me.txtSCR(10).Text = "" Then
        MsgBox "請輸入完稿起日!!!", vbExclamation + vbOKOnly
        Me.txtSCR(10).SetFocus
        TXTSCR_GotFocus 10
        Exit Sub
    End If
    If CheckIsTaiwanDate(Me.txtSCR(10).Text) = False Then
        Me.txtSCR(10).SetFocus
        TXTSCR_GotFocus 10
        Exit Sub
    End If
    If Me.txtSCR(11).Text = "" Then
        MsgBox "請輸入完稿止日!!!", vbExclamation + vbOKOnly
        Me.txtSCR(11).SetFocus
        TXTSCR_GotFocus 11
        Exit Sub
    End If
    If CheckIsTaiwanDate(Me.txtSCR(11).Text) = False Then
        Me.txtSCR(11).SetFocus
        Exit Sub
    End If
    If Val(Me.txtSCR(10).Text) > Val(Me.txtSCR(11).Text) Then
        MsgBox "完稿日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
        Me.txtSCR(10).SetFocus
        TXTSCR_GotFocus 10
        Exit Sub
    End If
    If Me.txtSCR(12).Text <> "" And Me.txtSCR(13).Text <> "" Then
        If Me.txtSCR(12).Text > Me.txtSCR(13).Text Then
            MsgBox "承辦人部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
            Me.txtSCR(12).SetFocus
            TXTSCR_GotFocus 12
            Exit Sub
        End If
    End If
    '查詢
    If Index = 0 Then
        m_blnColOrderAsc = True 'Added by Lydia 2016/08/11 欄位資料由小到大排序
        Screen.MousePointer = vbHourglass
        Me.grdList.MousePointer = flexHourglass
        If QueryData() = False Then
            strTit = "查詢資料"
            strMsg = "無資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        End If
        Me.grdList.MousePointer = flexDefault
        Screen.MousePointer = vbDefault
    '列印
    Else
        If Me.grdList.Rows > 1 Then
            Screen.MousePointer = vbHourglass
            PrintData
            ShowPrintOk
            Screen.MousePointer = vbDefault
        Else
            ShowNoData
        End If
    End If
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyF2
                  RsSitu 0
               Case vbKeyF3
                  RsSitu 1
               Case vbKeyF5
                  RsSitu 2
               Case vbKeyF4
                  RsSitu 5
            End Select
            KeyCode = 0
         End If
      Case vbKeyF9, vbKeyF10, vbKeyReturn
         If ActionEdit <> 3 Then
            Select Case KeyCode
               Case vbKeyF9, vbKeyReturn
                  RsSitu 3
               Case vbKeyF10
                  RsSitu 4
            End Select
            KeyCode = 0
         End If
      Case vbKeyHome, vbKeyPageUp, vbKeyPageDown, vbKeyEnd
         If ActionEdit = 3 Then
            Select Case KeyCode
               Case vbKeyHome
                  RsAction 0
               Case vbKeyPageUp
                  RsAction 1
               Case vbKeyPageDown
                  RsAction 2
               Case vbKeyEnd
                  RsAction 3
            End Select
            KeyCode = 0
         End If
    Case vbKeyEscape
        If MsgBox("是否確定結束?", vbYesNo + vbCritical) = vbYes Then Unload Me
    Case Else
        Exit Sub
    End Select
   If KeyCode <> vbKeyF2 And KeyCode <> vbKeyF3 And KeyCode <> vbKeyF4 And KeyCode <> vbKeyF5 And KeyCode <> vbKeyEscape Then
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
         'Added by Lydia 2022/05/18
         If m_bQuery Then
             TBar1.Buttons(4).Enabled = True
         Else
             TBar1.Buttons(4).Enabled = False
         End If
         'end 2022/05/18
   End If
End Sub

Private Sub Form_Load()
Dim i As Integer
   
   m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
    MoveFormToCenter Me
    '取得使用者執行各項功能的權限
    '由個人進入
    If ProState = "1" Then
        m_bInsert = IsUserHasRightOfFunction("frm090627P", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090627P", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090627P", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090627P", strFind, False)
        Me.Check1.Enabled = False
    '由管理進入
    Else
        m_bInsert = IsUserHasRightOfFunction("frm090627M", strAdd, False)
        m_bUpdate = IsUserHasRightOfFunction("frm090627M", strEdit, False)
        m_bDelete = IsUserHasRightOfFunction("frm090627M", strDel, False)
        m_bQuery = IsUserHasRightOfFunction("frm090627M", strFind, False)
        Me.Check1.Enabled = True
    End If
    strExc(0) = "SELECT * FROM SpecialCaseRecord WHERE ROWNUM<1"
    intI = 1
    Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   strRsStart1 = Empty
   strRsEnd1 = Empty
   strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord Order By SCR01 "
   intI = 1
   'edit by nickc 2007/02/05 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0), True)
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)
   If intI = 1 Then
        RsTemp.MoveFirst
      strRsStart1 = "" & RsTemp.Fields(0).Value
        RsTemp.MoveLast
      strRsEnd1 = "" & RsTemp.Fields(0).Value
      RsAction 0
    Else
    Me.lblCaseName.Caption = ""
    Me.lblCaseProperty.Caption = ""
    Me.lblCountry.Caption = ""
    Me.lblOurCaseNo.Caption = ""
    Me.lblPromoter.Caption = ""
   End If
   ActionEdit = 3
   CmdSitu True
    Me.txtSCR(0).Locked = True
    Me.txtSCR(1).Locked = True
    Me.Check1.Enabled = False
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
   'Added by Lydia 2022/05/18
   If m_bQuery Then
       TBar1.Buttons(4).Enabled = True
   Else
       TBar1.Buttons(4).Enabled = False
   End If
   'end 2022/05/18
         
   Me.SSTab1.Tab = 0 'Added by Lydia 2022/05/18
End Sub

Private Function ReadSpecialCaseRecord(ByRef tsTmp() As String) As Boolean
Dim i As Integer, j As Integer, Lbl As Label, txt As TextBox, strTmp As String
Dim strTxt(0 To 4) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
    strTxt(1) = tsTmp(1)
    SCR(1) = strTxt(1)
    For i = 0 To 2
        If i = 2 Then
            Me.Check1.Value = vbUnchecked
        Else
            Me.txtSCR(i).Text = ""
        End If
    Next i
    Me.lblCaseName.Caption = ""
    Me.lblCaseProperty.Caption = ""
    Me.lblCountry.Caption = ""
    Me.lblOurCaseNo.Caption = ""
    Me.lblPromoter.Caption = ""
    Me.Label3.Caption = "Create : "
    Me.Label4.Caption = "Update : "
   If SCR(1) = "" Then Exit Function
   StrSQLa = "Select * From SpecialCaseRecord Where SCR01='" & SCR(1) & "' Order By SCR01 "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
        SCR(1) = "" & rsA.Fields(0).Value
        SCR(2) = "" & rsA.Fields(1).Value
        SCR(3) = "" & rsA.Fields(2).Value
        SCR(4) = "" & rsA.Fields(3).Value
        SCR(5) = "" & rsA.Fields(4).Value
        SCR(6) = "" & rsA.Fields(5).Value
        SCR(7) = "" & rsA.Fields(6).Value
        SCR(8) = "" & rsA.Fields(7).Value
        SCR(9) = "" & rsA.Fields(8).Value
    Else
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        Exit Function
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   Me.txtSCR(0).Text = SCR(1)
   Me.txtSCR(1).Text = SCR(2)
    GetBasicData (Me.txtSCR(0).Text)
    Me.Check1.Value = IIf(SCR(3) <> "", vbChecked, vbUnchecked)
    'Add By Cheng 2004/03/04
    Me.txtSCR(1).Tag = Me.txtSCR(1).Text
    Me.Check1.Tag = Me.Check1.Value
    If SCR(4) <> "" Then
        Me.Label3.Caption = Me.Label3.Caption & GetStaffName(SCR(4))
    End If
    If SCR(5) <> "" Then
        Me.Label3.Caption = Me.Label3.Caption & " " & ChangeTStringToTDateString(Val(SCR(5)) - 19110000)
    End If
    If SCR(6) <> "" Then
        Me.Label3.Caption = Me.Label3.Caption & " " & Format(SCR(6), "##:##")
    End If
    If SCR(7) <> "" Then
        Me.Label4.Caption = Me.Label4.Caption & GetStaffName(SCR(7))
    End If
    If SCR(8) <> "" Then
        Me.Label4.Caption = Me.Label4.Caption & " " & ChangeTStringToTDateString(Val(SCR(8)) - 19110000)
    End If
    If SCR(9) <> "" Then
        Me.Label4.Caption = Me.Label4.Caption & " " & Format(SCR(9), "##:##")
    End If
    'End
End Function

Private Sub Form_Unload(Cancel As Integer)
   Set frm090627 = Nothing
End Sub

Private Sub RsSitu(ByVal Situ As Integer)
Dim i As Integer, St1 As String, St2 As String
Dim TBmk As Variant
Dim StrSQLa As String
 
 '911106 nick
 On Error GoTo CheckingErr
 
 Static TmpSCR(4) As String
   Select Case Situ
      Case 0 '按下新增add
        TmpSCR(1) = Me.txtSCR(0).Text
        Me.lblCaseName.Caption = ""
        Me.lblCaseProperty.Caption = ""
        Me.lblCountry.Caption = ""
        Me.lblOurCaseNo.Caption = ""
        Me.lblPromoter.Caption = ""
         CmdSitu False
         TxtLock 0
         ActionEdit = 0
         Me.txtSCR(0).SetFocus
        TXTSCR_GotFocus 0
      Case 1 '按下修改modi
         CmdSitu False
         TxtLock 1
         ActionEdit = 1
        TmpSCR(1) = Me.txtSCR(0).Text
      Case 2 '按下刪除delete
        If Me.txtSCR(0).Text = "" Then
            MsgBox "無資料可刪除!!!", vbExclamation + vbOKOnly
            Exit Sub
        End If
        If DelMsg Then
            StrSQLa = "Delete From SpecialCaseRecord Where SCR01='" & Me.txtSCR(0).Text & "' "
            cnnConnection.Execute StrSQLa
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01>='" & SCR(1) & "' Order By SCR01 "
             intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
               ReadSpecialCaseRecord strExc
            Else
                strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01<='" & SCR(1) & "' Order By SCR01 Desc "
                 intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   strExc(1) = "" & RsTemp.Fields(0).Value
                   ReadSpecialCaseRecord strExc
                Else
                   RsAction 0
                End If
            End If
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord Order By SCR01 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                 RsTemp.MoveFirst
               strRsStart1 = "" & RsTemp.Fields(0).Value
                 RsTemp.MoveLast
               strRsEnd1 = "" & RsTemp.Fields(0).Value
            End If
        End If
      Case 3 'update
         If ActionEdit = 0 Then '在新增狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Modify By Cheng 2004/02/23
            '若有勾選主管核可則用"V"存入資料庫
'            strSQLA = "Insert Into SpecialCaseRecord (SCR01, SCR02, SCR03) Values('" & Me.txtSCR(0).Text & "'," & Val(Me.txtSCR(1).Text) & ",'" & IIf(Me.Check1.Value = vbChecked, "Y", "") & "')"
            StrSQLa = "Insert Into SpecialCaseRecord (SCR01, SCR02, SCR03) Values('" & Me.txtSCR(0).Text & "'," & Val(Me.txtSCR(1).Text) & ",'" & IIf(Me.Check1.Value = vbChecked, "V", "") & "')"
            'End
            cnnConnection.Execute StrSQLa
            '寄E-Mail
            SendMail "新增"
            If Me.txtSCR(0).Text < strRsStart1 Then
                strRsStart1 = Me.txtSCR(0).Text
            End If
            If Me.txtSCR(0).Text > strRsEnd1 Then
                strRsEnd1 = Me.txtSCR(0).Text
            End If
            strExc(1) = Me.txtSCR(0).Text
            ReadSpecialCaseRecord strExc
         ElseIf ActionEdit = 1 Then '在修改狀態按Enter鍵
            If Not GetData Then Exit Sub
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'Modify By Cheng 2004/02/23
            '若有勾選主管核可則用"V"存入資料庫
'            strSQLA = "Update SpecialCaseRecord Set SCR02=" & Val(Me.txtSCR(1).Text) & ", SCR03='" & IIf(Me.Check1.Value = vbChecked, "Y", "") & "' " & _
'                            " Where SCR01='" & Me.txtSCR(0).Text & "' "
            StrSQLa = ""
            If Me.txtSCR(1).Text <> Me.txtSCR(1).Tag Then
                StrSQLa = StrSQLa & " SCR02=" & Val(Me.txtSCR(1).Text) & ","
            End If
            If Me.Check1.Value <> Me.Check1.Tag Then
                StrSQLa = StrSQLa & " SCR03='" & IIf(Me.Check1.Value = vbChecked, "V", "") & "',"
            End If
            If StrSQLa <> "" Then
                StrSQLa = Left(StrSQLa, Len(StrSQLa) - 1)
            Else
                GoTo NoUpdate
            End If
'            strSQLA = "Update SpecialCaseRecord Set SCR02=" & Val(Me.txtSCR(1).Text) & ", SCR03='" & IIf(Me.Check1.Value = vbChecked, "V", "") & "' " & _
'                            " Where SCR01='" & Me.txtSCR(0).Text & "' "
            StrSQLa = "Update SpecialCaseRecord Set " & StrSQLa & " " & _
                            " Where SCR01='" & Me.txtSCR(0).Text & "' "
            'End
            cnnConnection.Execute StrSQLa
            'Add By Cheng 2004/03/18
            '修改也要寄E-Mail
            If ProState <> "2" Then SendMail "修改"
            'End
NoUpdate:
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord Order By SCR01"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
                 RsTemp.MoveFirst
               strRsStart1 = "" & RsTemp.Fields(0).Value
                 RsTemp.MoveLast
               strRsEnd1 = "" & RsTemp.Fields(0).Value
            End If
            RefreshData Me.txtSCR(0).Text, "1"
            strExc(1) = Me.txtSCR(0).Text
            ReadSpecialCaseRecord strExc
         ElseIf ActionEdit = 2 Then '在查詢狀態按下Enter鍵
            If Me.txtSCR(0).Text = "" Then
               MsgBox "收文號不可空白，請重新輸入 !", vbCritical
               Me.txtSCR(0).SetFocus
               TXTSCR_GotFocus 0
               Exit Sub
            End If
            intI = 1
            strExc(0) = "SELECT COUNT(*) FROM SpecialCaseRecord WHERE SCR01='" & Me.txtSCR(0).Text & "' "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.Fields(0) = 0 Then
                  MsgBox "查無此特殊案件記錄 !", vbCritical
                    strExc(1) = TmpSCR(1)
               Else
                    strExc(1) = Me.txtSCR(0).Text
               End If
            End If
            ReadSpecialCaseRecord strExc
         End If
         CmdSitu True
         ActionEdit = 3
         TxtLock 3
      Case 4 'cancel
         If ActionEdit <> 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
         End If
         CmdSitu True
        If TmpSCR(1) = "" Then TmpSCR(1) = strRsStart1
        strExc(1) = TmpSCR(1)
         ActionEdit = 3
         ReadSpecialCaseRecord strExc
         TxtLock 3
      Case 5 'query
        TmpSCR(1) = Me.txtSCR(0).Text
         CmdSitu False
         TxtLock 2
         ActionEdit = 2
         Me.txtSCR(0).SetFocus
         TXTSCR_GotFocus 0
   End Select
   
   Exit Sub
CheckingErr:
    MsgBox Err.Description
End Sub

Private Sub RsAction(ByVal Sty As Integer)
 Dim i As Integer
On Error GoTo ErrHand
   Screen.MousePointer = vbHourglass
   intI = 1
   Select Case Sty
      Case 0 '第一筆
         strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01='" & strRsStart1 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields(0).Value
        Else
            strExc(0) = "SELECT SCR01 From SpecialCaseRecord WHERE SCR01>='" & strRsStart1 & "' Order By SCR01 "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strRsStart1 = strExc(1)
            End If
         End If
      Case 1 '前一筆
         If Me.txtSCR(0).Text = strRsStart1 Then
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 6
            Exit Sub
         Else
            intI = 1
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01<'" & Me.txtSCR(0).Text & "' Order By SCR01 Desc "
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
            End If
         End If
      Case 2 '後一筆
         If Me.txtSCR(0).Text = strRsEnd1 Then
            Beep
            Screen.MousePointer = vbDefault
            DataErrorMessage 7
            Exit Sub
         Else
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01>'" & Me.txtSCR(0).Text & "' Order By SCR01 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), True)  'edit by nickc 2007/02/05 不用 dll 了 = objLawDll.ReadRstMsg(intI, strExc(0), True)
            If intI = 1 Then
               strExc(1) = "" & RsTemp.Fields(0).Value
            End If
         End If
      Case 3 '最後筆
         strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01='" & strRsEnd1 & "' "
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strExc(1) = "" & RsTemp.Fields(0).Value
        Else
            strExc(0) = "SELECT SCR01 FROM SpecialCaseRecord WHERE SCR01<='" & strRsEnd1 & "' Order By SCR01 Desc "
             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                strExc(1) = "" & RsTemp.Fields(0).Value
                strRsEnd1 = strExc(1)
            End If
         End If
   End Select
   ReadSpecialCaseRecord strExc
   Screen.MousePointer = vbDefault
   Exit Sub
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Sub

Private Sub CmdSitu(ByVal TF As Boolean)
 Dim i As Integer, txt As TextBox
   If TF = True Then
'      TxtLock 0
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = True
         If Not IsEmptyText(strRsStart1) And Not IsEmptyText(strRsEnd1) Then
            TBar1.Buttons(i + 5).Enabled = True
         Else
            TBar1.Buttons(i + 5).Enabled = False
         End If
      Next
      TBar1.Buttons(11).Enabled = False
      TBar1.Buttons(12).Enabled = False
      TBar1.Buttons(14).Enabled = True
   Else
'      TxtLock 1
      For i = 1 To 4
         TBar1.Buttons(i).Enabled = False
         TBar1.Buttons(i + 5).Enabled = False
      Next
      TBar1.Buttons(11).Enabled = True
      TBar1.Buttons(12).Enabled = True
      TBar1.Buttons(14).Enabled = False
   End If
End Sub

Private Sub TxtLock(ByVal Lt As Integer)

Select Case Lt
Case 0 '新增
    Me.txtSCR(0).Locked = False
    Me.txtSCR(1).Locked = False
    Me.txtSCR(0).Text = ""
    Me.txtSCR(1).Text = ""
    Me.lblCaseName.Caption = ""
    Me.lblCaseProperty.Caption = ""
    Me.lblCountry.Caption = ""
    Me.lblOurCaseNo.Caption = ""
    Me.lblPromoter.Caption = ""
    If ProState = "2" Then
        Me.Check1.Enabled = True
        Me.Check1.Value = vbUnchecked
    Else
        Me.Check1.Enabled = False
        Me.Check1.Value = vbUnchecked
    End If
Case 1 '修改
    Me.txtSCR(0).Locked = False
    Me.txtSCR(1).Locked = False
    If ProState = "2" Then
        Me.Check1.Enabled = True
    Else
        Me.Check1.Enabled = False
    End If
Case 2 '查詢
    Me.txtSCR(0).Locked = False
    Me.txtSCR(1).Locked = False
    Me.txtSCR(0).Text = ""
    Me.txtSCR(1).Text = ""
    Me.lblCaseName.Caption = ""
    Me.lblCaseProperty.Caption = ""
    Me.lblCountry.Caption = ""
    Me.lblOurCaseNo.Caption = ""
    Me.lblPromoter.Caption = ""
    Me.Check1.Enabled = False
    Me.Check1.Value = vbUnchecked
Case 3 '按下取消後的狀態
    Me.txtSCR(0).Locked = True
    Me.txtSCR(1).Locked = True
    Me.Check1.Enabled = False
End Select
End Sub

Private Sub grdList_Click()
    grdList_ShowSelection
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
'Add by Morgan 2003/12/26
Private Sub grdList_DblClick()
   SSTab1.Tab = 0
End Sub

Private Sub grdList_SelChange()
   Dim nRow As Integer
   grdList_ShowSelection
   
   If grdList.row > 0 And grdList.row <= grdList.Rows - 1 Then
       nRow = grdList.row
       strExc(1) = Me.grdList.TextMatrix(nRow, 3)
       ReadSpecialCaseRecord strExc
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    
On Error Resume Next
    Select Case Me.SSTab1.Tab
    Case 0
        Me.txtSCR(0).SetFocus
        TXTSCR_GotFocus 0
        Me.cmdQuery(0).Default = False
    Case 1
        Me.txtSCR(10).SetFocus
        TXTSCR_GotFocus 10
        Me.cmdQuery(0).Default = True
    End Select
End Sub

Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHand
   SSTab1.Tab = 0 'Add by Morgan 2011/10/19
   Select Case Button.Index
      Case 1 '按下新增
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 0
      Case 2 '按下修改
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 1
      Case 3 '按下刪除
         RsSitu 2
      Case 4 '按下查詢
         SSTab1.TabEnabled(1) = False 'Add by Morgan 2011/10/19
         RsSitu 5
      Case 6 '第一筆
         RsAction 0
      Case 7 '前一筆
         RsAction 1
      Case 8 '後一筆
         RsAction 2
      Case 9 '最後筆
         RsAction 3
      Case 11 '按下確定
         RsSitu 3
      Case 12 '按下取消
         RsSitu 4
      Case 14 '結束
         Unload Me
         Exit Sub
   End Select

   If ActionEdit <> 3 Then Exit Sub 'Add by Morgan 2011/10/19
   SSTab1.TabEnabled(1) = True 'Add by Morgan 2011/10/19
   
   ' Ken 90.07.16 -- Start
   If Button.Index <> 14 And Button.Index <> 1 And Button.Index <> 2 And Button.Index <> 3 And Button.Index <> 4 Then
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
      'Added by Lydia 2022/05/18
      If m_bQuery Then
          TBar1.Buttons(4).Enabled = True
      Else
          TBar1.Buttons(4).Enabled = False
      End If
      'end 2022/05/18
   End If
   ' Ken 90.07.16 -- End
   Exit Sub
ErrHand:
   ShowMsg "錯誤 : " & Err.Description
End Sub

Private Function CheckRule() As Boolean
Dim i As Integer, bolChk As Boolean, j As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   CheckRule = False
   If Me.txtSCR(0).Text = "" Then
      MsgBox "收文號不可空白 !", vbCritical
      Me.txtSCR(0).SetFocus
      TXTSCR_GotFocus 0
      Exit Function
   End If
    GetBasicData (Me.txtSCR(0).Text)
    If ActionEdit = 0 Then
        If ChkDataRepeat(Me.txtSCR(0).Text) = True Then
            Me.txtSCR(0).SetFocus
            TXTSCR_GotFocus 0
            Exit Function
        End If
    End If
   If Me.txtSCR(1).Text = "" Then
      MsgBox "加值點數不可空白 !", vbCritical
      Me.txtSCR(1).SetFocus
      TXTSCR_GotFocus 1
      Exit Function
   End If
    If IsNumeric(Me.txtSCR(1).Text) = False Then
      MsgBox "加值點數輸入錯誤 !", vbCritical
      Me.txtSCR(1).SetFocus
      TXTSCR_GotFocus 1
      Exit Function
    End If
    If Val(Me.txtSCR(1).Text) <= 0 Or Val(Me.txtSCR(1).Text) >= 10 Then
      MsgBox "加值點數超過範圍 !", vbCritical
      Me.txtSCR(1).SetFocus
      TXTSCR_GotFocus 1
      Exit Function
    End If
   CheckRule = True
End Function

Private Function GetData() As Boolean
Dim i As Integer
    GetData = False
    If CheckRule = False Then Exit Function
    SCR(1) = Me.txtSCR(0).Text
    SCR(2) = Me.txtSCR(1).Text
    'Modify By Cheng 2004/02/23
'    SCR(3) = IIf(Me.Check1.Value = vbChecked, "Y", "")
    SCR(3) = IIf(Me.Check1.Value = vbChecked, "V", "")
    'End
    GetData = True
End Function

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
For Each objTxt In Me.txtSCR
    If objTxt.Enabled = True Then
       Cancel = False
       txtSCR_Validate objTxt.Index, Cancel
       If Cancel = True Then
          Exit Function
       End If
    End If
Next
TxtValidate = True
End Function

Private Sub TXTSCR_GotFocus(Index As Integer)
    TextInverse Me.txtSCR(Index)
End Sub

Private Sub TXTSCR_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
    Case 0, 12, 13 '收文號, 承辦人員部門別
        KeyAscii = UpperCase(KeyAscii)
    End Select
End Sub

Private Sub txtSCR_LostFocus(Index As Integer)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    Select Case Index
    Case 11 '收發文日期
        If Me.txtSCR(10).Text <> "" And Me.txtSCR(11).Text <> "" Then
            If Val(Me.txtSCR(10).Text) > Val(Me.txtSCR(11).Text) Then
                MsgBox "完稿日期範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtSCR(10).SetFocus
                TXTSCR_GotFocus 10
                Exit Sub
            End If
        End If
    Case 12 '承辦人部門
        If Me.txtSCR(12).Text <> "" And Me.txtSCR(13).Text <> "" Then
            If Me.txtSCR(12).Text > Me.txtSCR(13).Text Then
                MsgBox "承辦人部門範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                Me.txtSCR(12).SetFocus
                TXTSCR_GotFocus 12
                Exit Sub
            End If
        End If
    End Select
End Sub

Private Sub txtSCR_Validate(Index As Integer, Cancel As Boolean)
    If Me.txtSCR(Index).Text = "" Then Exit Sub
    If Me.txtSCR(Index).Locked = True Then Exit Sub
    Select Case Index
    Case 0 '收文號
        If GetBasicData(Me.txtSCR(Index).Text) = False Then
            Cancel = True
        ElseIf ActionEdit = 0 Then
            If ChkDataRepeat(Me.txtSCR(Index).Text) = True Then
                Cancel = True
            End If
        End If
    Case 1 '加值點數
        If IsNumeric(Me.txtSCR(Index).Text) = False Then
            MsgBox "加值點數輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
        ElseIf Val(Me.txtSCR(Index).Text) <= 0 Or Val(Me.txtSCR(Index).Text) >= 10 Then
            MsgBox "加值點數超過範圍!!!", vbExclamation + vbOKOnly
            Cancel = True
        End If
    Case 10, 11 '收發文日期區間
        If CheckIsTaiwanDate(Me.txtSCR(Index).Text) = False Then
            Cancel = True
        End If
    End Select
    If Cancel = True Then TXTSCR_GotFocus Index
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim nRow As Integer
   
    QueryData = False
    InitialGridList
    strSql = ""
    If Me.txtSCR(10).Text <> "" Then
        strSql = strSql & " And EP09>=" & DBDATE(Me.txtSCR(10).Text) & " "
    End If
    If Me.txtSCR(11).Text <> "" Then
        strSql = strSql & " And EP09<=" & DBDATE(Me.txtSCR(11).Text) & " "
    End If
    If Me.txtSCR(12).Text <> "" Then
        strSql = strSql & " And ST03>='" & ChgSQL(Me.txtSCR(12).Text) & "' "
    End If
    If Me.txtSCR(13).Text <> "" Then
        strSql = strSql & " And ST03<='" & ChgSQL(Me.txtSCR(13).Text) & "' "
    End If
    
   'Added by Lydia 2022/05/18 本所案號
   If Trim(txtCode(0)) <> "" Then
       strSql = strSql & " AND CP01='" & Trim(txtCode(0)) & "'"
   End If
   If Trim(txtCode(1)) <> "" Then
       strSql = strSql & " AND CP02='" & Trim(txtCode(1)) & "'"
   End If
   If Trim(txtCode(2)) <> "" Then
       strSql = strSql & " AND CP03='" & Trim(txtCode(2)) & "'"
   End If
   If Trim(txtCode(3)) <> "" Then
       strSql = strSql & " AND CP04='" & Trim(txtCode(3)) & "'"
   End If
   'end 2022/05/18
   
'    strSQL = "Select SCR01, SCR02, SCR03, PA01||'-'||PA02||'-'||PA03||'-'||PA04, EP09, NA03, Decode(PA09,'020',CPM04,CPM03), ST02 From SpecialCaseRecord, EngineerProgress, CaseProgress, Patent, Nation, Staff, CasePropertyMap Where SCR01=EP02 And SCR01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And PA09=NA01(+) And CP14=ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSQL & " Order By 5,4,1 "
    'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
    strSql = "Select EP09, ST02, SCR01, PA01||'-'||PA02||'-'||PA03||'-'||PA04, NA03, Decode(PA09,'000',CPM03,CPM04), SCR02, SCR03 From SpecialCaseRecord, EngineerProgress, CaseProgress, Patent, Nation, Staff, CasePropertyMap Where SCR01=EP02 And SCR01=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And PA09=NA01(+) And CP14=ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) " & strSql & " Order By 1, 4, 3 "
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount > 0 Then
        QueryData = True
        UpdateGridList rsTmp
        'Added by Lydia 2022/01/11 MSFlexGrid 不支援UniCode，以MSHFlexGrid換掉
        If grdList.Rows > 1 Then
           grdList.FixedRows = 1
        End If
    End If
    rsTmp.Close
    Set rsTmp = Nothing
End Function

' 初始化列表
Public Sub InitialGridList()
    grdList.Clear
    grdList.Rows = 1
    grdList.Cols = 9
    grdList.ColWidth(0) = 300
    grdList.row = 0
    grdList.col = 0
    grdList.ColAlignment(0) = flexAlignCenterCenter
    grdList.col = 1
    grdList.Text = "完稿日"
    grdList.ColWidth(1) = 800
    grdList.ColAlignment(1) = flexAlignLeftCenter
    grdList.col = 2
    grdList.Text = "承辦人"
    grdList.ColWidth(2) = 800
    grdList.ColAlignment(2) = flexAlignLeftCenter
    grdList.col = 3
    grdList.Text = "收文號"
    grdList.ColWidth(3) = 1000
    grdList.ColAlignment(3) = flexAlignLeftCenter
    grdList.col = 4
    grdList.Text = "本所案號"
    grdList.ColWidth(4) = 1500
    grdList.ColAlignment(4) = flexAlignLeftCenter
    grdList.col = 5
    grdList.Text = "申請國家"
    grdList.ColWidth(5) = 800
    grdList.ColAlignment(5) = flexAlignLeftCenter
    grdList.col = 6
    grdList.Text = "案件性質"
    grdList.ColWidth(6) = 1200
    grdList.ColAlignment(6) = flexAlignLeftCenter
    grdList.col = 7
    grdList.Text = "加值件數"
    grdList.ColWidth(7) = 1000
    grdList.ColAlignment(7) = flexAlignLeftCenter
    grdList.col = 8
    grdList.Text = "核可"
    grdList.ColWidth(8) = 600
    grdList.ColAlignment(8) = flexAlignLeftCenter
End Sub

Private Sub UpdateGridList(ByRef rsTmp As ADODB.Recordset)
Dim nRow As Integer
    rsTmp.MoveFirst
    Do While rsTmp.EOF = False
        grdList.Rows = grdList.Rows + 1
        nRow = grdList.Rows - 1
        grdList.TextMatrix(nRow, 1) = ChangeWStringToTString("" & rsTmp.Fields(0).Value)
        grdList.TextMatrix(nRow, 2) = "" & rsTmp.Fields(1).Value
        grdList.TextMatrix(nRow, 3) = "" & rsTmp.Fields(2).Value
        grdList.TextMatrix(nRow, 4) = "" & rsTmp.Fields(3).Value
        grdList.TextMatrix(nRow, 5) = "" & rsTmp.Fields(4).Value
        grdList.TextMatrix(nRow, 6) = "" & rsTmp.Fields(5).Value
        Me.grdList.row = nRow
        Me.grdList.col = 7
        Me.grdList.CellAlignment = flexAlignRightCenter
        grdList.TextMatrix(nRow, 7) = "" & rsTmp.Fields(6).Value
        grdList.TextMatrix(nRow, 8) = "" & rsTmp.Fields(7).Value
        Me.grdList.row = nRow
        rsTmp.MoveNext
    Loop
End Sub

Private Sub PrintData()
Dim ii As Integer
    
    Page = 1
    PrintTitle
    For ii = 1 To Me.grdList.Rows - 1
        '完稿日
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 1)
        '承辦人
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 2)
        '收文號
        Printer.CurrentX = PLeft(2)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 3)
        '本所案號
        Printer.CurrentX = PLeft(3)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 4)
        '申請國家
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 5)
        '案件性質
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 6)
        '加值件數
        Printer.CurrentX = PLeft(6) + Printer.TextWidth("加值件數") - Printer.TextWidth(Me.grdList.TextMatrix(ii, 7))
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 7)
        '核可
        Printer.CurrentX = PLeft(7)
        Printer.CurrentY = iPrint
        Printer.Print Me.grdList.TextMatrix(ii, 8)
        iPrint = iPrint + 300
        If iPrint > 10000 And ii <> Me.grdList.Rows - 1 Then
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print String(200, "-")
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
    Next ii
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    Printer.EndDoc

End Sub

Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "特殊案件記錄明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6000
Printer.CurrentY = iPrint
Printer.Print "完稿日：" & Format(ChangeTStringToTDateString(Me.txtSCR(10).Text) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(Me.txtSCR(11).Text)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "承辦人員部門別：" & Me.txtSCR(12).Text & " " & IIf(Me.txtSCR(12).Text <> "" Or Me.txtSCR(13).Text <> "", "－", "") & " " & Me.txtSCR(13).Text
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "加值件數"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "核可"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 1750
PLeft(2) = 3100
PLeft(3) = 4450
PLeft(4) = 6750
PLeft(5) = 8000
PLeft(6) = 9250 + 1250
PLeft(7) = 10500 + 1250
End Sub

Private Function GetBasicData(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetBasicData = False
'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
StrSQLa = "Select PA01||'-'||PA02||'-'||PA03||'-'||PA04, PA05||' '||PA06||' '||PA07, PA09||' '||NA03, CP10||' '||Decode(PA09,'000',CPM03,CPM04), CP14||' '||ST02 From CaseProgress, Patent, Nation, Staff, CasePropertyMap Where CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 And PA09=NA01(+) And CP14=ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) And CP09='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Me.lblOurCaseNo.Caption = "" & rsA.Fields(0).Value
    Me.lblCaseName.Caption = "" & rsA.Fields(1).Value
    Me.lblCountry.Caption = "" & rsA.Fields(2).Value
    Me.lblCaseProperty.Caption = "" & rsA.Fields(3).Value
    Me.lblPromoter.Caption = "" & rsA.Fields(4).Value
    GetBasicData = True
Else
    MsgBox "此收文號無相關資料!!!", vbExclamation + vbOKOnly
    Me.lblCaseName.Caption = ""
    Me.lblCaseProperty.Caption = ""
    Me.lblCountry.Caption = ""
    Me.lblOurCaseNo.Caption = ""
    Me.lblPromoter.Caption = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Function ChkDataRepeat(strSCR01 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkDataRepeat = False
StrSQLa = "Select * From SpecialCaseRecord Where SCR01='" & strSCR01 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    MsgBox "此收文號已輸入特殊案件記錄!!!", vbExclamation + vbOKOnly
    ChkDataRepeat = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Modify by Amy 2017/12/06 改Pub_SendMail發
Private Sub SendMail(strMod As String)
    Dim strSubject As String, strContent As String
    Dim strTo As String
    
    If strUserNum = "71011" Or strUserNum = "67002" Then
        strTo = "94007"
    Else
        Select Case Left(GetStaffDepartment(strUserNum), 2)
        Case "P1"
            'Added by Lydia 2023/04/24 修改王副總退休之相關控制
            If strSrvDate(1) >= "20230511" Then
                strTo = "99050"
            ElseIf strSrvDate(1) >= "20230501" Then
                strTo = "71011;99050"
            Else
            'end 2023/04/24
                strTo = "71011"
            End If 'Added by Lydia 2023/04/24
        Case "P2"
            'strTo = "67002"  'cancel by 2020/5/5
        Case Else
            Exit Sub
        End Select
    End If
   
    strSubject = "<<特殊案件>>" & strMod & "記錄通知"
    strContent = "收文號：" & Me.txtSCR(0).Text & vbCrLf & _
                 "本所案號：" & Me.lblOurCaseNo.Caption & vbCrLf & _
                 "案件名稱：" & Me.lblCaseName.Caption & vbCrLf & _
                 "案件性質：" & Me.lblCaseProperty.Caption & vbCrLf & _
                 "承辦人：" & Me.lblPromoter.Caption & vbCrLf & _
                 "加值件數：" & Me.txtSCR(1).Text & vbCrLf & _
                 "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
                 strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
    'Modified by Lydia 2022/05/30 傳入收文號
    'PUB_SendMail strUserNum, strTo, "", strSubject, strContent
    PUB_SendMail strUserNum, strTo, Me.txtSCR(0).Text, strSubject, strContent
End Sub

Private Sub SendMail_Old(strMod As String)

'    If strUserNum = "71011" Or strUserNum = "67002" Then
'        'modify by sonia 2014/9/9 改68001為94007
'        frm880005.txtEmail(0).Text = "94007"
'    Else
'        Select Case Left(GetStaffDepartment(strUserNum), 2)
'        Case "P1"
'            frm880005.txtEmail(0).Text = "71011"
'        Case "P2"
'            frm880005.txtEmail(0).Text = "67002"
'        Case Else
'            Exit Sub
'        End Select
'    End If
'    '若使用者為北所人員, 則E-Mail後面不加@taie.com.tw
'    If PUB_GetST06(strUserNum) = "1" Then
'        '無動作
'    '若使用者非北所人員, 則E-Mail後面加@taie.com.tw
'    Else
'        frm880005.txtEmail(0).Text = frm880005.txtEmail(0).Text & "@taie.com.tw"
'    End If
'    frm880005.txtEmail(1).Text = "<<特殊案件>>" & strMod & "記錄通知"
'    frm880005.txtEmail(2).Text = "收文號：" & Me.txtSCR(0).Text & vbCrLf & _
'                                                "本所案號：" & Me.lblOurCaseNo.Caption & vbCrLf & _
'                                                "案件名稱：" & Me.lblCaseName.Caption & vbCrLf & _
'                                                "案件性質：" & Me.lblCaseProperty.Caption & vbCrLf & _
'                                                "承辦人：" & Me.lblPromoter.Caption & vbCrLf & _
'                                                "加值件數：" & Me.txtSCR(1).Text & vbCrLf & _
'                                                "是否主管核可：" & IIf(Me.Check1.Value = vbChecked, "是", "否") & vbCrLf & vbCrLf & _
'                                                strMod & "資料人員：" & strUserNum & " " & GetStaffName(strUserNum)
'    frm880005.Form_Activate: DoEvents
'    frm880005.cmdok_Click 0: DoEvents
End Sub
'end 2017/12/06

Private Sub RefreshData(strSCR01 As String, strRefreshKind As String)
'strRefreshKind : 1 修改
Dim ii As Integer

With Me.grdList
    If .Rows > 1 And .TextMatrix(1, 1) <> "" Then
        For ii = 1 To .Rows - 1
            If .TextMatrix(ii, 3) = strSCR01 Then
                If strRefreshKind = "1" Then
                    .TextMatrix(ii, 7) = Me.txtSCR(1).Text
                    .TextMatrix(ii, 8) = IIf(Me.Check1.Value = vbChecked, "V", "")
                End If
            End If
        Next ii
    End If
End With
End Sub

'Added by Lydia 2016/08/11 點選欄位進行排序
Private Sub grdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long

   'Modified by Lydia 2022/01/11
   'Pub_MSFGrdColRow grdList, x, y, nCol, nRow
   getGrdColRow grdList, x, y, nCol, nRow
   
   If nCol < 0 Or nRow < 0 Then Exit Sub
   
   grdList.col = nCol
   grdList.row = nRow
   If Me.grdList.row < 1 Then
      If InStr("完稿日,加值件數", Me.grdList.Text) > 0 Then
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.grdList.Sort = 5 '字串昇冪
            
            m_blnColOrderAsc = False
         Else
            Me.grdList.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
'end 2016/08/11

'Added by Lydia 2022/05/18
Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

