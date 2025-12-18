VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210101_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶資料修改"
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8280
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8280
   Begin VB.CommandButton cmdContact 
      Caption         =   "接洽人資料"
      Height          =   330
      Left            =   6390
      TabIndex        =   71
      Top             =   630
      Width           =   1455
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7650
      Top             =   810
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
            Picture         =   "frm210101_1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm210101_1.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   8280
      _ExtentX        =   14605
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
            Object.Tag             =   "F2"
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
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4005
      Left            =   45
      TabIndex        =   36
      Top             =   1770
      Width           =   8205
      _ExtentX        =   14478
      _ExtentY        =   7070
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "聯絡資料"
      TabPicture(0)   =   "frm210101_1.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(6)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(9)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(10)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(11)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(13)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(16)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(18)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(19)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(20)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label6"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(3)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(24)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(25)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDisp(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(28)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cboContact"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtRead(3)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtRead(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtRead(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtRead(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtRead(9)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtRead(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtRead(10)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtRead(11)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtRead(0)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtEdit(9)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtEdit(10)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtEdit(5)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtEdit(4)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtEdit(3)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtEdit(2)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtEdit(1)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtEdit(8)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtEdit(6)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtEdit(13)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtRead(12)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label41(39)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "cmdSearchZip"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "cmdTW"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm210101_1.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblComp(5)"
      Tab(1).Control(1)=   "lblComp(4)"
      Tab(1).Control(2)=   "lblComp(3)"
      Tab(1).Control(3)=   "lblComp(2)"
      Tab(1).Control(4)=   "lblComp(1)"
      Tab(1).Control(5)=   "lblComp(0)"
      Tab(1).Control(6)=   "Label1(22)"
      Tab(1).Control(7)=   "Label1(21)"
      Tab(1).Control(8)=   "Label1(17)"
      Tab(1).Control(9)=   "Label1(26)"
      Tab(1).Control(10)=   "Label1(27)"
      Tab(1).Control(11)=   "txtEdit(12)"
      Tab(1).Control(12)=   "txtEdit(11)"
      Tab(1).Control(13)=   "txtEdit(7)"
      Tab(1).Control(14)=   "txtEdit(14)"
      Tab(1).Control(15)=   "Label1(23)"
      Tab(1).Control(16)=   "txtComp(5)"
      Tab(1).Control(17)=   "txtComp(4)"
      Tab(1).Control(18)=   "txtComp(3)"
      Tab(1).Control(19)=   "txtComp(2)"
      Tab(1).Control(20)=   "txtComp(1)"
      Tab(1).Control(21)=   "txtComp(0)"
      Tab(1).Control(22)=   "cboCU167"
      Tab(1).Control(23)=   "cboCU166"
      Tab(1).ControlCount=   24
      Begin VB.ComboBox cboCU166 
         Height          =   276
         Left            =   -73380
         Style           =   2  '單純下拉式
         TabIndex        =   82
         Top             =   3330
         Width           =   6450
      End
      Begin VB.ComboBox cboCU167 
         Height          =   276
         Left            =   -73380
         Style           =   2  '單純下拉式
         TabIndex        =   81
         Top             =   3630
         Width           =   2985
      End
      Begin VB.CommandButton cmdTW 
         Caption         =   "臺灣地址格式"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6930
         TabIndex        =   22
         Top             =   2625
         Width           =   1160
      End
      Begin VB.CommandButton cmdSearchZip 
         Caption         =   "Zip"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7610
         TabIndex        =   20
         Top             =   2295
         Width           =   480
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   0
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   41
         Top             =   1635
         Width           =   330
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   1
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   42
         Top             =   1920
         Width           =   330
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   2
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   43
         Top             =   2205
         Width           =   330
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   3
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   44
         Top             =   2520
         Width           =   330
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   4
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   45
         Top             =   2805
         Width           =   330
      End
      Begin VB.TextBox txtComp 
         Height          =   270
         Index           =   5
         Left            =   -72015
         MaxLength       =   1
         TabIndex        =   46
         Top             =   3075
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "顧問專用信箱："
         Height          =   180
         Index           =   23
         Left            =   -74880
         TabIndex        =   86
         Top             =   1300
         Width           =   1260
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   288
         Index           =   14
         Left            =   -73248
         TabIndex        =   40
         Top             =   1272
         Width           =   6264
         VariousPropertyBits=   679493659
         MaxLength       =   200
         Size            =   "11049;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "跨所同意主管："
         ForeColor       =   &H00FF0000&
         Height          =   180
         Index           =   39
         Left            =   4500
         TabIndex        =   85
         Top             =   2625
         Width           =   1260
      End
      Begin MSForms.TextBox txtRead 
         Height          =   300
         Index           =   12
         Left            =   5790
         TabIndex        =   84
         Top             =   2595
         Width           =   1095
         VariousPropertyBits=   679493651
         ForeColor       =   16711680
         Size            =   "1931;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   13
         Left            =   5115
         TabIndex        =   17
         Top             =   1650
         Width           =   2970
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   288
         Index           =   7
         Left            =   -73248
         TabIndex        =   37
         Top             =   408
         Width           =   336
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   288
         Index           =   11
         Left            =   -73248
         TabIndex        =   38
         Top             =   684
         Width           =   336
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   288
         Index           =   12
         Left            =   -73248
         TabIndex        =   39
         Top             =   984
         Width           =   336
         VariousPropertyBits=   679493659
         MaxLength       =   1
         Size            =   "582;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   405
         Index           =   6
         Left            =   1080
         TabIndex        =   27
         Top             =   3510
         Width           =   6900
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12171;714"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   8
         Left            =   5085
         TabIndex        =   24
         Top             =   2895
         Width           =   2895
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5115;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   1
         Left            =   1365
         TabIndex        =   23
         Top             =   2895
         Width           =   2895
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5115;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   2
         Left            =   1245
         TabIndex        =   14
         Top             =   1335
         Width           =   2970
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   3
         Left            =   5115
         TabIndex        =   15
         Top             =   1335
         Width           =   2970
         VariousPropertyBits=   679493659
         MaxLength       =   20
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   4
         Left            =   1245
         TabIndex        =   16
         Top             =   1650
         Width           =   2970
         VariousPropertyBits=   679493659
         MaxLength       =   15
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   5
         Left            =   1245
         TabIndex        =   11
         Top             =   750
         Width           =   2970
         VariousPropertyBits=   679493659
         MaxLength       =   35
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   10
         Left            =   5085
         TabIndex        =   26
         Top             =   3195
         Width           =   2895
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5115;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtEdit 
         Height          =   285
         Index           =   9
         Left            =   1365
         TabIndex        =   25
         Top             =   3195
         Width           =   2895
         VariousPropertyBits=   679493659
         MaxLength       =   50
         Size            =   "5115;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   300
         Index           =   0
         Left            =   7875
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   1665
         Width           =   300
         VariousPropertyBits=   679493661
         BackColor       =   -2147483633
         Size            =   "8555;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   11
         Left            =   5325
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   480
         Width           =   1995
         VariousPropertyBits=   679493661
         BackColor       =   -2147483633
         Size            =   "3519;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   10
         Left            =   1365
         TabIndex        =   21
         Top             =   2595
         Width           =   495
         VariousPropertyBits=   679493659
         MaxLength       =   3
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   300
         Index           =   5
         Left            =   1050
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   1950
         Width           =   6495
         VariousPropertyBits=   679493661
         BackColor       =   -2147483633
         Size            =   "11465;529"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   9
         Left            =   6885
         TabIndex        =   19
         Top             =   2295
         Width           =   705
         VariousPropertyBits=   679493659
         MaxLength       =   10
         Size            =   "1244;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   4
         Left            =   1035
         TabIndex        =   18
         Top             =   2295
         Width           =   5850
         VariousPropertyBits=   679493659
         Size            =   "10319;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   6
         Left            =   5430
         TabIndex        =   47
         TabStop         =   0   'False
         Top             =   750
         Width           =   2400
         VariousPropertyBits=   679493661
         BackColor       =   -2147483633
         Size            =   "4233;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   2
         Left            =   1245
         TabIndex        =   12
         Top             =   1035
         Width           =   2970
         VariousPropertyBits=   679493659
         Size            =   "5239;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtRead 
         Height          =   285
         Index           =   3
         Left            =   5115
         TabIndex        =   13
         Top             =   1035
         Width           =   2955
         VariousPropertyBits=   679493659
         Size            =   "5203;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   300
         Left            =   1245
         TabIndex        =   10
         Top             =   420
         Width           =   2985
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "5265;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "LINE ID："
         Height          =   180
         Index           =   28
         Left            =   4350
         TabIndex        =   83
         Top             =   1710
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   180
         Index           =   27
         Left            =   -74880
         TabIndex        =   80
         Top             =   3690
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   180
         Index           =   26
         Left            =   -74880
         TabIndex        =   79
         Top             =   3360
         Width           =   1440
      End
      Begin MSForms.Label lblDisp 
         Height          =   180
         Index           =   6
         Left            =   2115
         TabIndex        =   77
         Top             =   2505
         Width           =   45
         VariousPropertyBits=   27
         Size            =   "11721;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "客戶狀態："
         Height          =   180
         Index           =   25
         Left            =   4350
         TabIndex        =   75
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址國籍："
         Height          =   180
         Index           =   24
         Left            =   120
         TabIndex        =   74
         Top             =   2625
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請地址："
         Height          =   180
         Index           =   3
         Left            =   120
         TabIndex        =   73
         Top             =   1935
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄國內電子報：         （N:不寄）"
         Height          =   180
         Index           =   17
         Left            =   -74880
         TabIndex        =   70
         Top             =   444
         Width           =   2916
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄專利雙週報：      （N:不寄）"
         Height          =   180
         Index           =   21
         Left            =   -74880
         TabIndex        =   69
         Top             =   708
         Width           =   2772
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否寄顧問電子報：      （Y:寄/N:不寄）"
         Height          =   180
         Index           =   22
         Left            =   -74880
         TabIndex        =   68
         Top             =   1008
         Width           =   3168
      End
      Begin VB.Label Label1 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "負責人(中)："
         Height          =   255
         Index           =   14
         Left            =   4185
         TabIndex        =   67
         Top             =   780
         Width           =   1170
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "業務備註："
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   3540
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(其他3)："
         Height          =   180
         Index           =   20
         Left            =   4335
         TabIndex        =   65
         Top             =   3240
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "           (其他2)："
         Height          =   180
         Index           =   19
         Left            =   120
         TabIndex        =   64
         Top             =   3255
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(其他1)："
         Height          =   180
         Index           =   18
         Left            =   4335
         TabIndex        =   63
         Top             =   2970
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "負責人(英)："
         Height          =   180
         Index           =   16
         Left            =   120
         TabIndex        =   62
         Top             =   780
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話２："
         Height          =   180
         Index           =   13
         Left            =   4335
         TabIndex        =   61
         Top             =   1065
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "電話１："
         Height          =   180
         Index           =   12
         Left            =   120
         TabIndex        =   60
         Top             =   1065
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "手機："
         Height          =   180
         Index           =   11
         Left            =   120
         TabIndex        =   59
         Top             =   1680
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真２："
         Height          =   180
         Index           =   10
         Left            =   4335
         TabIndex        =   58
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "傳真１："
         Height          =   180
         Index           =   9
         Left            =   120
         TabIndex        =   57
         Top             =   1350
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預設接洽人："
         Height          =   180
         Index           =   7
         Left            =   120
         TabIndex        =   56
         Top             =   480
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "E-Mail  (代表)："
         Height          =   180
         Index           =   6
         Left            =   120
         TabIndex        =   55
         Top             =   2925
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "聯絡地址："
         Height          =   180
         Index           =   2
         Left            =   120
         TabIndex        =   54
         Top             =   2325
         Width           =   1230
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-台灣：                    (1：專利商標 2：專利法律)"
         Height          =   180
         Index           =   0
         Left            =   -74880
         TabIndex        =   53
         Top             =   1635
         Width           =   5445
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "專利案預設收據公司別-非台灣：                (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   1
         Left            =   -74880
         TabIndex        =   52
         Top             =   1920
         Width           =   6450
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-台灣：                    (1：專利商標 2：專利法律 )"
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   51
         Top             =   2205
         Width           =   5490
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "商標案預設收據公司別-非台灣：                (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   3
         Left            =   -74880
         TabIndex        =   50
         Top             =   2505
         Width           =   6450
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-台灣：                    (1：專利商標 2：專利法律)"
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   49
         Top             =   2805
         Width           =   5445
      End
      Begin VB.Label lblComp 
         AutoSize        =   -1  'True
         Caption         =   "其他案預設收據公司別-非台灣：                (1：專利商標 2：專利法律 J：台一智權)"
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   48
         Top             =   3075
         Width           =   6450
      End
   End
   Begin MSForms.TextBox txtRead 
      Height          =   285
      Index           =   8
      Left            =   5430
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1530
      Width           =   1275
      VariousPropertyBits=   679493661
      BackColor       =   -2147483633
      Size            =   "2249;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtRead 
      Height          =   285
      Index           =   7
      Left            =   1335
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1530
      Width           =   1275
      VariousPropertyBits=   679493661
      BackColor       =   -2147483633
      Size            =   "2249;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtRead 
      Height          =   270
      Index           =   1
      Left            =   1305
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1275
      Width           =   5985
      VariousPropertyBits=   679493661
      BackColor       =   -2147483633
      Size            =   "10557;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073750016
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtKey 
      Height          =   285
      Index           =   0
      Left            =   1305
      TabIndex        =   0
      Top             =   660
      Width           =   1005
      VariousPropertyBits=   679493659
      MaxLength       =   9
      Size            =   "1773;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtKey 
      Height          =   285
      Index           =   1
      Left            =   1290
      TabIndex        =   1
      Top             =   960
      Width           =   5985
      VariousPropertyBits=   679493659
      MaxLength       =   80
      Size            =   "10557;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   35
      Top             =   1550
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Index           =   5
      Left            =   4605
      TabIndex        =   34
      Top             =   1550
      Width           =   780
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   2
      Left            =   3465
      TabIndex        =   33
      Top             =   1550
      Width           =   1005
      ForeColor       =   255
      VariousPropertyBits=   27
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   1
      Left            =   6750
      TabIndex        =   32
      Top             =   1550
      Width           =   810
      VariousPropertyBits=   27
      Size            =   "1429;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   0
      Left            =   2655
      TabIndex        =   31
      Top             =   1550
      Width           =   810
      VariousPropertyBits=   27
      Size            =   "1429;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   3
      Left            =   3675
      TabIndex        =   9
      Top             =   735
      Width           =   810
      VariousPropertyBits=   27
      Size            =   "1429;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   4
      Left            =   4515
      TabIndex        =   8
      Top             =   735
      Width           =   810
      VariousPropertyBits=   27
      Size            =   "1429;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblDisp 
      Height          =   180
      Index           =   5
      Left            =   5400
      TabIndex        =   7
      Top             =   735
      Width           =   810
      VariousPropertyBits=   27
      Size            =   "1429;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Update："
      Height          =   180
      Index           =   15
      Left            =   2970
      TabIndex        =   5
      Top             =   735
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱(英)："
      Height          =   255
      Index           =   8
      Left            =   60
      TabIndex        =   4
      Top             =   1290
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶名稱(中)："
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   3
      Top             =   975
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   2
      Top             =   705
      Width           =   1230
   End
End
Attribute VB_Name = "frm210101_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/14 Form2.0已修改 txtKey()/lblDisp()/cboContact/txtRead()/txtEdit()
'Memo by Lydia 2019/07/01 表單名稱:個人客戶資料修改=>客戶資料修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
'Modify by Morgan 2008/8/11 接洽人(Text)改預設接洽人(Combo),新增"接洽人資料"按鍵
Option Explicit

'授權員工編號
Dim strSalesNo As String
'授權員工所屬部門代碼
Dim strDeptNo As String
'目前狀態
Dim iCurState As Integer
'前一客戶
Dim lst_CU01 As String, lst_CU02 As String
'目前客戶
Dim cur_CU01 As String, cur_CU02 As String
'呼叫表單
Dim frmCaller As Form
'智權人員
Dim strCU13 As String 'Added by Morgan 2016/12/13
'聯絡人編號
Dim strCU127 As String 'Add by Morgan 2008/7/31
'帶人主管權限
Dim bolPLimit As Boolean 'Add by Sindy 2010/7/28
'記錄 是否寄電子報值及專利雙週報值
Dim oldCU132 As String, oldCU145 As String 'Add by Amy 2014/03/06
'Add by Amy 2014/05/26
Dim bolSpecMan As Boolean  '是否為特殊設定檔人員
Dim strSpecCode As String '特殊設定檔設定代號
'Add by Amy 2016/05/16
Public bolBack As Boolean '是否為其他Form回來(跳板用)
Dim oldCU20 As String, strCU176 As String 'Add by Amy 2023/02/15 代表信箱,e化客戶指定信箱(正本)
Dim strCU10 As String, strCU79 As String, strCU112 As String, strAcrossAreaMail As String 'Add by Amy 2023/04/21
Dim bolAcrossArea As Boolean 'Add by Amy 2023/04/21
Dim oldCU199 As String 'Added by Lydia 2024/01/15 顧問專用信箱
Dim strECustMsg As String 'Added by Morgan 2024/6/3 全E化客戶提醒
Dim oldCU116 As String, oldCU117 As String, oldCU118 As String  'Added by Morgan 2024/6/3

Public Sub setCaller(frmFrom As Form)
   Set frmCaller = frmFrom
End Sub

Public Sub setSalesNo(stNo As String)
   strSalesNo = stNo
End Sub

Public Sub setDeptNo(stDept As String)
   strDeptNo = stDept
End Sub

Private Sub cboContact_Click()
   'Modify by Amy 2021/12/14 改成Form 2.0
   'strCU127 = Format(cboContact.ItemData(cboContact.ListIndex), "00")
   'Modified by Morgan 2022/5/17 設定選單時不要執行否則會變空白
   'strCU127 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
   If cboContact.Tag <> "" Then
      strCU127 = Format(PUB_GetItemData(cboContact.Tag, cboContact.ListIndex), "00")
   End If
   'end 2022/5/17
End Sub

Private Sub cboContact_LostFocus()
   If cboContact = "" And cboContact.ListCount > 0 Then cboContact.ListIndex = 0
End Sub

Private Sub cboCU166_Click()
   If cboCU166.ListIndex >= 0 And cboCU166.Tag <> "" & cboCU166.ListIndex Then
      cboCU167.Clear
      strExc(0) = "select cu127 from customer where " & ChgCustomer(Left(cboCU166, 9))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         PUB_AddContact Left(cboCU166, 8), cboCU167, "" & RsTemp("cu127")
      End If
      cboCU166.Tag = "" & cboCU166.ListIndex
   End If
End Sub

'Add by Morgan 2008/7/31
Private Sub cmdContact_Click()
   With frm210101_2
      'Added by Morgan 2015/8/24 修改接洽人前客戶資料先存檔
      If iCurState = 2 Then
         tlbar_ButtonClick tlbar.Buttons(11)
         If iCurState <> 0 Then Exit Sub
         tlbar_ButtonClick tlbar.Buttons(2)
         If iCurState <> 2 Then Exit Sub
      End If
      'end 2015/8/24
      If Me.txtRead(10) < "010" Then .bolCU87IsTW = True: .Check2.Value = 1
      .strCU127 = strCU127
      .txtCuNo = cur_CU01
      .lblCustName = txtKey(1)
      .lblCustAddress = txtRead(4)
      .OpenContactTable
      .SetState iCurState
      .Show vbModal
      If iCurState = 2 Then
         'Modify by Amy 2021/12/14 改成Form 2.0
         'PUB_AddContact cur_CU01, cboContact, strCU127
         'Modified by Morgan 2022/5/17
         'strExc(10) = cboContact.Tag
         cboContact.Tag = ""
         'end 2022/5/17
         PUB_AddContact cur_CU01, cboContact, strCU127, True, True, strExc(10)
         cboContact.Tag = strExc(10)
         'end 2021/12/14
      End If
   End With
End Sub

Private Sub cmdSearchZip_Click()
    Dim stBackField As String, stText As String
    
    If iCurState <> 2 Then Exit Sub
    
    stBackField = "txtRead(9)"
    stText = txtRead(4)
    Call frm100134.SetParent(Me)
    Me.Hide
    frm100134.BFormZip = stBackField
    If stText <> MsgText(601) Then
        frm100134.GetStreet stText, 2
        Call frm100134.QueryData
    End If
    frm100134.Show
End Sub

'Add by Amy 2016/05/16 for 接洽人按郵遞區號查詢返回用
Private Sub Form_Activate()
    If bolBack = True Then
        frm210101_2.Show vbModal
        bolBack = False
    End If
End Sub

'Modify by Amy 2021/12/14 KeyDown 搬過來
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
      Case vbKeyF3
      '修改
         If tlbar.Buttons(2).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(2))
         End If
      Case vbKeyF4
      '查詢
         If tlbar.Buttons(4).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(4))
         End If
      Case vbKeyHome
      '第一筆
         If tlbar.Buttons(6).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(6))
         End If
      Case vbKeyPageUp
      '上一筆
         If tlbar.Buttons(7).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(7))
         End If
      Case vbKeyPageDown
      '下一筆
         If tlbar.Buttons(8).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(8))
         End If
      Case vbKeyEnd
      '最後筆
         If tlbar.Buttons(9).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(9))
         End If
      Case vbKeyF9, vbKeyReturn
      '確定
         'Mark by Amy 2021/12/15 改完 Form2.0後Form_Load會觸發Enter
'         If tlbar.Buttons(11).Enabled = True Then
'            Call tlbar_ButtonClick(tlbar.Buttons(11))
'         End If
      Case vbKeyF10
      '取消
         If tlbar.Buttons(12).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(12))
         End If
      Case vbKeyEscape
      '結束
        If tlbar.Buttons(14).Enabled = True Then
            Call tlbar_ButtonClick(tlbar.Buttons(14))
         End If
    End Select
End Sub

Private Sub Form_Load()
   Dim objTxt As Object 'Add by Amy 2016/05/16
   
   MoveFormToCenter Me
   iCurState = 4 'Modify by Amy 2021/12/15 從下面搬上來
   Call SetToolBar(0)
   Call FormReset
   '預設為查詢
   Call SetToolBar(4)
   Call SetInputs(4)
   txtKey(0).Text = "X"
   
   'Add by Amy 2015/02/04 +總經理業務工作代理人員
   If CheckLevel(strUserNum, "總經理業務工作代理人員") = True Then
       bolSpecMan = True
       strSpecCode = "總經理業務工作代理人員"
   'Add  by Amy 2014/05/26 開放專利處部份智權同仁資料給彥葶代為處理
   ElseIf CheckLevel(strUserNum, "A8") = True Then
        bolSpecMan = True
        strSpecCode = "A8"
   End If
   'end 2014/05/26
   SSTab1.Tab = 0 'Add by Amy 2014/08/11
   If 案件預設收據公司別啟用日 >= Val(strSrvDate(1)) Then
        For Each objTxt In Me.lblComp
            objTxt.Visible = False
        Next
        For Each objTxt In Me.txtComp
            objTxt.Visible = False
        Next
   End If
   
   'Removed by Morgan 2012/1/2 恢復
   ''Add by Morgan 2011/4/8 取消專利雙週報欄位
   'Label1(21).Visible = False
   'txtEdit(11).Visible = False
   
   Call Pub_AddPersonRec("frm210101_1") 'Added by Lydia 2019/07/01 智權部-個人常用區
   
   'Added by Lydia 2020/03/31 事務所合併日起台灣案取消(1:專利商標 2:專利法律) 的標題，非台灣案改標題為(J:智權公司 空白:系統預設)。
   If strSrvDate(1) >= 事務所合併日 Then
       For intI = 0 To 5
          Select Case intI
              Case 0, 2, 4  '台灣案:CU160,CU162,CU164
                  'Modifed by Lydia 2021/07/13 debug-統一改標題為(J:智權公司 空白:系統預設)
                  'lblComp(intI).Visible = False
                  'lblCU16X(intI).Visible = False
                  lblComp(intI).Caption = Replace(lblComp(intI).Caption, "1：專利商標 2：專利法律", "J：智權公司 空白:系統預設")
                  'end 2021/07/13
              Case 1, 3, 5  '非台灣案:CU161,CU163,CU165
                  lblComp(intI).Caption = Replace(lblComp(intI).Caption, "1：專利商標 2：專利法律 J：台一智權", "J：智權公司 空白:系統預設")
          End Select
       Next
   End If
   'end 2020/03/30
End Sub

'工具列控制
Private Sub SetToolBar(iStatus As Integer)
   Dim i As Integer
   For i = 1 To 13
      tlbar.Buttons(i).Enabled = False
   Next
   tlbar.Buttons(14).Enabled = True
   
   Select Case iStatus
      Case 0
      '瀏覽
         tlbar.Buttons(2).Enabled = True
         tlbar.Buttons(4).Enabled = True
         'Modify By Sindy 2009/08/20
         '特定操作人員開放前後筆查詢
         '2011/4/13 MODIFY BY 取消C3改為C1
         'MODIFY BY SONIA 2015/6/1 分所權限改為部門為M71,另加F2
         'If Pub_StrUserSt03 = "M51" Or _
            strUserNum = "75033" Or _
            PUB_GetST05(strUserNum) = "C1" Or _
            PUB_GetST05(strUserNum) = "C2" Or _
            PUB_GetST05(strUserNum) = "F1" Or _
            PUB_GetST05(strUserNum) = "KM" Or _
            PUB_GetST05(strUserNum) = "K1" Or _
            PUB_GetST05(strUserNum) = "NM" Or _
            PUB_GetST05(strUserNum) = "N1" Then
         'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
         If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M71" Or _
            strUserNum = "74018" Or _
            PUB_GetST05(strUserNum) = "F1" Or _
            PUB_GetST05(strUserNum) = "F2" Then
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
         cmdContact.Enabled = True
      Case 1
      '新增
      Case 2
      '修改
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         cmdContact.Enabled = True
      Case 3
      '刪除
      Case 4
      '查詢
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         cmdContact.Enabled = False
      Case Else
   End Select
   'Add by Amy 2016/05/16 客戶狀態有值不可修改
   'modify by sonia 2022/3/7 +設為對造
   If iStatus = 0 And (txtRead(11) = "不再使用" Or txtRead(11) = "不得代理" Or txtRead(11) = "不得代理專利" Or txtRead(11) = "不得代理商標" Or txtRead(11) = "設為對造") Then
      tlbar.Buttons(2).Enabled = False
   End If
   
End Sub

Private Sub FormReset()
   Dim oText, oLabel  'Modify by Amy 2021/12/14 原:As TextBox/As LABEL
   
   For Each oText In txtKey
      oText.Text = ""
      oText.Locked = True
   Next
   
   For Each oText In txtRead
      oText.Text = ""
      'Add by Amy 2016/07/04
      If oText.Index = 4 Or oText.Index = 9 Or oText.Index = 10 Then
        oText.Tag = ""
      End If
      oText.Enabled = True
      oText.Locked = True
   Next
   
   For Each oText In txtEdit
      oText.Text = ""
      oText.Locked = True
      oText.Tag = "" 'Added by Morgan 2016/12/14
   Next
   'Added by Lydia 2020/03/31 收據公司別
   For Each oText In txtComp
      oText.Text = ""
      oText.Locked = True
      oText.Tag = ""
   Next
   'end 2020/03/31
   
   For Each oLabel In lblDisp
      oLabel.Caption = ""
   Next
   
   cboContact.Clear 'Add by Morgan 2008/7/31
   'Added by Morgan 2016/12/13
   'cboContact.Tag = "" 'Removed by Morgan 2022/5/17
   cboCU166.Clear
   cboCU167.Clear
   'end 2016/12/13
End Sub

Private Sub SetInputs(Optional ByVal iStatus As Integer = 0)

   Dim oText, oLabel  'Modify by Amy 2021/12/14 原:As TextBox/As LABEL
   
   'Add by Morgan 2008/7/31
   cboContact.Locked = True
   
   Select Case iStatus
      
      Case 2
      '修改
         For Each oText In txtKey
            oText.Locked = True
            oText.Enabled = False
         Next
         For Each oText In txtEdit
            oText.Locked = False
         Next
        'Added by Lydia 2020/03/31 收據公司別
        For Each oText In txtComp
           oText.Locked = False
        Next
        'end 2020/03/31
        
         'Add by Morgan 2008/7/31
         cboContact.Locked = False
         cboCU166.Locked = False 'Added by Morgan 2016/12/14
         cboCU167.Locked = False 'Added by Morgan 2016/12/14
         
         'Add by Amy 2016/05/16 開放修改電話1/2、國籍、聯絡地址、郵遞區號
         For Each oText In txtRead
            If (oText.Index >= 2 And oText.Index <= 4) Or (oText.Index >= 9 And oText.Index <= 10) Then
                 oText.Locked = False
            End If
         Next
      Case 4
      '查詢
         For Each oText In txtKey
            oText.Text = ""
            oText.Enabled = True
            oText.Locked = False
         Next
         For Each oText In txtRead
            oText.Text = ""
            'oText.Enabled = True
         Next
         For Each oText In txtEdit
            oText.Text = ""
            oText.Locked = True
            oText.Tag = "" 'Added by Morgan 2016/12/14
         Next
         'Added by Lydia 2020/03/31 收據公司別
         For Each oText In txtComp
            oText.Text = ""
            oText.Locked = True
            oText.Tag = ""
         Next
         'end 2020/03/31
         
         For Each oLabel In lblDisp
            oLabel.Caption = ""
         Next
         'Add by Morgan 2008/7/31
         cboContact.Clear
         'Added by Morgan 2016/12/13
         cboContact.Locked = True
         'cboContact.Tag = "" 'Removed by Morgan 2022/5/17
         cboCU166.Clear
         cboCU166.Locked = True
         cboCU167.Clear
         cboCU167.Locked = True
         'end 2016/12/13
      Case Else
      '其他
         For Each oText In txtKey
            oText.Enabled = True
            oText.Locked = True
         Next
         For Each oText In txtRead
            'oText.Enabled = True
            oText.Locked = True 'Added by Morgan 2016/12/13
         Next
         For Each oText In txtEdit
            oText.Locked = True
         Next
         'Added by Lydia 2020/03/31 收據公司別
         For Each oText In txtComp
            oText.Locked = True
         Next
         'end 2020/03/31
         
         'Added by Morgan 2016/12/14
         cboContact.Locked = True
         cboCU166.Locked = True
         cboCU167.Locked = True
         'end 2016/12/14
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2024/6/3
   'Mark by Amy 2021/12/15 按Esc會當掉
   'Set frm210101_1 = Nothing
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
   Select Case Button.Index
      Case 1
      '新增
      Case 2
      '修改
         'Modify By Sindy 2023/9/6 mark:查詢已檢查過權限了
'         'edit by nickc 2008/01/18  開放M51 可用
'         'If txtRead(7) <> strSalesNo And Val(txtRead(7)) > 63001 And lblDisp(2) <> "（離職）" Then
'         'Modify by Morgan 2008/12/22 開放夏慧珠不限制
'         '2009/7/2 MODIFY BY SONIA 開放有改客戶異動權限者也可以改
'         'If txtRead(7) <> strSalesNo And Val(txtRead(7)) > 63001 And lblDisp(2) <> "（離職）" And Pub_StrUserSt03 <> "M51" And strUserNum <> "75033" Then
'         'Modify by Amy 2014/05/26
'         'Modify by Amy 2020/09/23 +txtRead(7)<>登入者 ex:A2023 無法修改 X51896(自己的客戶)
'         If bolSpecMan = True And txtRead(7) <> strUserNum Then
'            'Add by Amy 2015/02/04 +總經理業務工作代理人員
'            If InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), txtRead(7)) = 0 Then
'                MsgBox "您無權限修改此客戶！", vbCritical
'                Exit Sub
'            '特殊人員判斷(for 開放專利處部份智權同仁資料給彥葶代為處理)
'            ElseIf InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), txtRead(7)) = 0 Then
'                MsgBox "您無權限修改此客戶！", vbCritical
'                Exit Sub
'            Else
'                Call SetToolBar(2)
'                Call SetInputs(2)
'                cboContact.SetFocus
'                iCurState = 2
'            End If
'         'end 2014/05/26
'         'MODIFY BY SONIA 2015/6/1 分所權限改為部門為M71,另加F2
'         'ElseIf txtRead(7) <> strSalesNo And Val(txtRead(7)) > 63001 And lblDisp(2) <> "（離職）" And _
'                Pub_StrUserSt03 <> "M51" And strUserNum <> "75033" And PUB_GetST05(strUserNum) <> "C2" And PUB_GetST05(strUserNum) <> "C1" And PUB_GetST05(strUserNum) <> "F1" And PUB_GetST05(strUserNum) <> "KM" And PUB_GetST05(strUserNum) <> "K1" And PUB_GetST05(strUserNum) <> "NM" And PUB_GetST05(strUserNum) <> "N1" Then
'         'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
'         ElseIf txtRead(7) <> strSalesNo And Val(txtRead(7)) > 63001 And lblDisp(2) <> "（離職）" And _
'                Pub_StrUserSt03 <> "M51" And Pub_StrUserSt03 <> "M71" And strUserNum <> "74018" And PUB_GetST05(strUserNum) <> "F1" And PUB_GetST05(strUserNum) <> "F2" Then
'            MsgBox "非所屬智權人員客戶，不可修改！", vbCritical
'            Exit Sub
'         Else
            Call SetToolBar(2)
            Call SetInputs(2)
            cboContact.SetFocus
            iCurState = 2
'         End If
      Case 3
      '刪除
      Case 4
      '查詢
         lst_CU01 = cur_CU01
         lst_CU02 = cur_CU02
         iCurState = 4
         Call SetToolBar(4)
         Call SetInputs(4)
         txtKey(0).Text = "X"
         txtKey(0).SetFocus
         txtkey_GotFocus 0
      Case 6
      '第一筆
         doQuery (6)
      Case 7
      '上一筆
         doQuery (7)
      Case 8
      '下一筆
         doQuery (8)
      Case 9
      '最後筆
         doQuery (9)
      Case 11
      '確定
         '查詢
         If iCurState = 4 Then
            If txtKey(0) = "" And Trim(txtKey(0)) = "X" And txtKey(1) = "" Then
               MsgBox "客戶編號不可空白！", vbCritical
               Exit Sub
            ' Add By Sindy 98/02/17 使用客戶名稱做查詢
            ElseIf (txtKey(0) = "" Or Trim(txtKey(0)) = "X") And txtKey(1) <> "" Then
               If doQuery2(4) = False Then Exit Sub
            End If
            ' 98/02/17 End
            txtKey(0) = UCase(Left(Me.txtKey(0).Text & "000000000", 9))
            cur_CU01 = Left(txtKey(0), 8)
            cur_CU02 = Right(txtKey(0), 1)
         '修改
         ElseIf iCurState = 2 Then
            
            PUB_FilterFormText Me 'Add by Morgan 2008/12/22 修正畫面所有含跳行符號的文字框
            
            'Add by Amy 2016/05/16 +依國別判斷臺灣地址
            If FormCheck() = False Then Exit Sub
        
            'Add by Morgan 2008/8/11
            If CheckContactAddr = False Then
               'Removed by Morgan 2022/5/17 不必特別改回原接洽人,且cboContact.Tag已改為記錄所有接洽人代碼
               'cboContact.ListIndex = Val(cboContact.Tag)
               'end 2022/5/17
               cboContact.SetFocus
               Exit Sub
            End If
            If UpdateData() = True Then
               'Add by Amy 2023/02/15 修改代表信箱,若為全E化客戶(cu176有值)時,提醒若要修改指定信箱需通知電腦中心處理-Morgan ex:X39187020
               'Modified by Morgan 2024/6/3 全E化客戶任一信箱異動時(含聯絡人)，發信通知智權人員(一併彈提醒)--杜協理
               'If txtEdit(1) <> oldCU20 And strCU176 <> MsgText(601) Then
               '      ShowMsg "此客戶為全E化客戶" & vbCrLf & _
               '                     "若要修改指定信箱需通知電腦中心處理"
               'End If
               If strECustMsg <> "" Then MsgBox strECustMsg, vbInformation
               'end 2024/6/3
               
               'Add by Amy 2023/04/21 申請[是]地跨所且聯絡地由[不是]跨所改跨所,需mail 通知電腦中心
               If bolAcrossArea = True And strAcrossAreaMail <> MsgText(61) Then
                    'Modify by Amy 2024/10/29 發信給[程式管理人員]放最後一個-秀玲
                    strExc(1) = ""
                    strExc(9) = Pub_GetSpecMan("程式管理人員")
                    Call GetFLOW001Person(txtRead(7), "3", , , , strExc(2))
                    'Add by Amy 2023/05/09 發給接洽單簽核人員
                    If strExc(2) <> MsgText(601) Then
                        strExc(1) = strExc(2)
                    'Add by Amy 2025/07/30 +else,智權人員-txtRead(7) [沒]接洽單簽核主管 者,抓操作者的接洽單簽核主管
                    '      ex:83001 操作 X74546000 聯絡地址改跨所 (才可寄給杜協理)
                    Else
                        Call GetFLOW001Person(strUserNum, "3", , , , strExc(2))
                        strExc(1) = strExc(2)
                    End If
                    '非智權人員自已修改,加發智權人員
                    If strUserNum <> Me.txtRead(7) Then
                        strExc(1) = strExc(1) & "," & Me.txtRead(7)
                    End If
                    'end 2023/05/09
                    If strExc(1) = MsgText(601) Then
                        strExc(1) = strExc(9)
                    Else
                        strExc(1) = strExc(1) & "," & strExc(9)
                    End If
                    'end 2024/10/29
                    PUB_SendMail strUserNum, strExc(1), "", "客戶編號：" & txtKey(0) & "，非跨所客戶改為跨所客戶通知，呈報結果請轉電腦中心做後續處理。", strAcrossAreaMail
               End If
               
               'MsgBox "修改成功", vbInformation
               'Add by Amy 2014/03/06 當研發處修改是否寄電子報/專利雙週報為N時發mail通知智權人員
               Dim strMail(5) As String
               Erase strMail
               If Pub_StrUserSt03 = "D01" Then
                    '組客戶E-mail
                    If txtEdit(1) <> MsgText(601) Then strMail(0) = txtEdit(1) & vbCrLf
                    For intI = 8 To 10
                        If Trim(txtEdit(intI)) <> MsgText(601) Then
                            strMail(0) = strMail(0) & txtEdit(intI) & vbCrLf
                        End If
                    Next intI
                    '顯示客戶名稱
                    If txtKey(1) <> MsgText(601) And txtRead(1) <> MsgText(601) Then
                        strMail(3) = txtKey(1) & "(中文)" & vbCrLf & String(5, "　") & txtRead(1) & "(英文)" & vbCrLf
                    ElseIf txtKey(1) <> MsgText(601) Then
                        strMail(3) = txtKey(1) & "(中文)" & vbCrLf
                    ElseIf txtRead(1) <> MsgText(601) Then
                        strMail(3) = txtRead(1) & "(英文)" & vbCrLf
                    End If
                    '組mail 內容
                    strMail(1) = "客戶編號：" & txtKey(0) & vbCrLf & _
                                      "客戶名稱：" & strMail(3) & _
                                      "E-mail  ：" & Replace(strMail(0), vbCrLf, vbCrLf & String(4, "　")) & vbCrLf & vbCrLf & _
                                      "研發處已將客戶資料的"
                                      
                    strMail(2) = " 欄改為 N," & vbCrLf & _
                                      "請詳查原因後, 自行修正客戶資料內容, 謝謝 !"
                    
                    If oldCU132 = MsgText(601) And txtEdit(7) = "N" Then strMail(4) = "'電子報'"
                    If oldCU145 = MsgText(601) And txtEdit(11) = "N" Then strMail(5) = "'專利雙週報'"
                    If strMail(4) <> MsgText(601) And strMail(5) <> MsgText(601) Then
                        strMail(1) = strMail(1) & strMail(4) & "及" & strMail(5) & strMail(2)
                        PUB_SendMail strUserNum, txtRead(7), "", "客戶E-mail 信箱於此次寄發" & strMail(4) & "及" & strMail(5) & "時遭退回通知 !", strMail(1)
                    ElseIf strMail(4) <> MsgText(601) Or strMail(5) <> MsgText(601) Then
                        strMail(1) = strMail(1) & IIf(strMail(4) = "", strMail(5), strMail(4)) & strMail(2)
                        PUB_SendMail strUserNum, txtRead(7), "", "客戶E-mail 信箱於此次寄發" & IIf(strMail(4) = "", strMail(5), strMail(4)) & "時遭退回通知 !", strMail(1)
                    End If
               End If
               'end 2014/03/06
               'Add by Amy 2022/05/23 修改時若有更名前資料,提醒使用者詢問智權人員
               If cur_CU02 = "0" Then
                    strExc(1) = "select * from customer where cu01='" & Left(txtKey(0) & "00000000", 8) & "' and cu02<>'0' "
                    CheckOC3
                    AdoRecordSet3.CursorLocation = adUseClient
                    AdoRecordSet3.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
                    If AdoRecordSet3.RecordCount <> 0 Then
                       ShowMsg "此客戶曾經變更名稱，請與智權人員確認是否修改更名前資料(除智權人員欄外) ! 若需要修改請查更名前編號再人工修改 !"
                    End If
               End If
               'end 2022/05/23
            Else
               Exit Sub
            End If
         End If
         
         If doQuery(4) = True Then
            Call SetToolBar(0)
            Call SetInputs
            iCurState = 0
         End If
         txtKey(0).SetFocus
         Call txtkey_GotFocus(0)
         
      Case 12
      '取消
         If iCurState = 4 Then
            cur_CU01 = lst_CU01
            cur_CU02 = lst_CU02
            If cur_CU01 = "" Then
               MsgBox "無前次查詢紀錄，不可取消！", vbCritical
               Exit Sub
'               cur_CU02 = "0"
'               Call doQuery(9)
            Else
               Call doQuery(4)
            End If
         ElseIf iCurState = 2 Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
               Exit Sub
            Else
               Call doQuery(4)
            End If
         End If
         Call SetToolBar(0)
         Call SetInputs
         iCurState = 0
      Case 14
      '結束
         'Modify by Morgan 2003/12/16
'         If PUB_CheckFormExist(frmCaller.Name) = True Then
'            frmCaller.Show
'         End If
         Unload Me
   End Select
End Sub

Private Function UpdateData() As Boolean
   Dim strSql As String, lngEffRec As Long, strSQL1 As String
   Dim strCU80 As String
   Dim strCU166 As String, strCU167 As String 'Added by Morgan 2016/12/14
   
   'Add by Morgan 2008/7/31 加聯絡人編號
   If Val(strCU127) > 0 Then
      strSQL1 = ",CU127='" & strCU127 & "'"
   Else
      strSQL1 = ",CU127=NULL"
   End If
   
   'Added by Morgan 2016/12/13
   'Modified by Morgan 2017/3/10 國內副本收件人改下拉選單
   If cboCU166.ListIndex > 0 Then
      strCU166 = Left(cboCU166.Text, 9)
      If cboCU167.ListIndex > 0 Then
         strCU167 = Format(cboCU167.ItemData(cboCU167.ListIndex), "00")
      Else
         strCU167 = ""
      End If
   Else
      strCU166 = ""
      strCU167 = ""
   End If
   
   strSQL1 = strSQL1 & ",CU166='" & strCU166 & "',CU167='" & strCU167 & "'"
   'end 2016/12/13
   
On Error GoTo ErrHand
   
   'Modify by Morgan 2008/12/22 +CU132 是否寄電子報
   'Add By Sindy 98/03/10 增加E-Mail:其他1,其他2,其他3
   'Add By Sindy 2011/1/14 增加是否寄發專利雙週報
   'Add By Sindy 2011/3/17 增加是否寄發顧問電子報
   'Modify by Morgan 2011/4/8 取消專利雙週報欄位  & ",CU145=" & CNULL(txtEdit(11))
   'Modified by Morgan 2012/1/2 恢復 CU145
   'Modify by Amy 2016/05/16 +CU16/17/31/30/80/87/CU160~165並整理欄位順序
'   strSql = "Update Customer Set CU20='" & txtEdit(1) & "',CU116='" & txtEdit(8) & "',CU117='" & txtEdit(9) & "',CU118='" & txtEdit(10) & "'," & _
'      "CU18='" & txtEdit(2) & "', CU19='" & txtEdit(3) & "', CU22='" & txtEdit(4) & "', CU103='" & txtEdit(5) & "'" & _
'      ",CU84='" & strSalesNo & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),cu125=" & CNULL(ChgSQL(txtEdit(6))) & strSQL1 & _
'      ",CU132=" & CNULL(txtEdit(7)) & ",CU153=" & CNULL(txtEdit(12)) & ",CU145=" & CNULL(txtEdit(11)) & _
'      ",CU16=" & CNULL(txtRead(2)) & ",CU17=" & CNULL(txtRead(3)) & ",CU31=" & CNULL(txtRead(4)) & ",CU30=" & CNULL(txtRead(9)) & ",CU87=" & CNULL(txtRead(10)) & _
'      ",CU160=" & CNULL(txtComp(0)) & ",CU161=" & CNULL(txtComp(1)) & ",CU162=" & CNULL(txtComp(2)) & ",CU163=" & CNULL(txtComp(3)) & ",CU164=" & CNULL(txtComp(4)) & ",CU165=" & CNULL(txtComp(5)) & _
'      " Where CU01='" & Left(txtKey(0), 8) & "' And CU02='" & Right(txtKey(0), 1) & "'"
   If txtRead(11) <> MsgText(601) Then
      'Modified by Morgan 2021/10/21 除狀態為"刪址","遷移不明"，其他改提醒但不可直接取消--秀玲 Ex:X38805030 不該清除業務自行處理
      'If MsgBox("是否要取消客戶狀態？", vbExclamation + vbYesNo) = vbYes Then
      '    strCU80 = ",CU80=Null"
      'End If
      If txtRead(11) = "刪址" Or txtRead(11) = "遷移不明" Then
         If MsgBox("請留意!! 此編號目前的 [客戶狀態] 為 [" & txtRead(11) & "] ，是否要取消此設定？", vbExclamation + vbYesNo) = vbYes Then
             strCU80 = ",CU80=Null"
         End If
      Else
         MsgBox "請留意!! 此編號目前的 [客戶狀態] 為 [" & txtRead(11) & "] ，若要 [取消] 請通知 [檔案室]。", vbExclamation
      End If
      'end 2021/10/21
   End If
   'Modified by Lydia 2019/04/12 拿掉UpdateID,Date,Time(CU84,CU85,CU86) ; 可能因為接洽人資料修改會call本程式,按下確定後會寫入無變更的log (ex.75033在108/4/12的維護記錄)
   'strSql = "Update Customer Set CU16=" & CNULL(ChgSQL(txtRead(2))) & ",CU17=" & CNULL(ChgSQL(txtRead(3))) & ",CU18=" & CNULL(ChgSQL(txtEdit(2))) & ", CU19=" & CNULL(ChgSQL(txtEdit(3))) & _
        ",CU20=" & CNULL(ChgSQL(txtEdit(1))) & ",CU22=" & CNULL(ChgSQL(txtEdit(4))) & ",CU30=" & CNULL(ChgSQL(txtRead(9))) & ",CU31=" & CNULL(ChgSQL(txtRead(4))) & _
        strCU80 & ",CU84='" & strSalesNo & "',CU85=to_number(to_char(sysdate,'YYYYMMDD')),CU86=to_number(to_char(sysdate,'HH24MI')),CU87=" & CNULL(ChgSQL(txtRead(10))) & _
        ",CU103=" & CNULL(ChgSQL(txtEdit(5))) & ",CU116=" & CNULL(ChgSQL(txtEdit(8))) & ",CU117=" & CNULL(ChgSQL(txtEdit(9))) & ",CU118=" & CNULL(ChgSQL(txtEdit(10))) & _
        ",cu125=" & CNULL(ChgSQL(txtEdit(6))) & strSQL1 & _
        ",CU132=" & CNULL(txtEdit(7)) & ",CU145=" & CNULL(txtEdit(11)) & ",CU153=" & CNULL(txtEdit(12))  & _
        ",CU160=" & CNULL(txtComp(0)) & ",CU161=" & CNULL(txtComp(1)) & ",CU162=" & CNULL(txtComp(2)) & ",CU163=" & CNULL(txtComp(3)) & ",CU164=" & CNULL(txtComp(4)) & ",CU165=" & CNULL(txtComp(5)) & _
        " Where CU01='" & Left(txtKey(0), 8) & "' And CU02='" & Right(txtKey(0), 1) & "'"
   'Modified by Lydia 2021/08/26 +LINE ID => CU21
   'Modified by Lydia 2024/01/15 +顧問專用信箱=>CU199
         strSql = "Update Customer Set CU16=" & CNULL(ChgSQL(txtRead(2))) & ",CU17=" & CNULL(ChgSQL(txtRead(3))) & ",CU18=" & CNULL(ChgSQL(txtEdit(2))) & ", CU19=" & CNULL(ChgSQL(txtEdit(3))) & _
        ",CU20=" & CNULL(ChgSQL(txtEdit(1))) & " ,CU21=" & CNULL(ChgSQL(txtEdit(13))) & " ,CU22=" & CNULL(ChgSQL(txtEdit(4))) & ",CU30=" & CNULL(ChgSQL(txtRead(9))) & ",CU31=" & CNULL(ChgSQL(txtRead(4))) & _
        strCU80 & ",CU87=" & CNULL(ChgSQL(txtRead(10))) & _
        ",CU103=" & CNULL(ChgSQL(txtEdit(5))) & ",CU116=" & CNULL(ChgSQL(txtEdit(8))) & ",CU117=" & CNULL(ChgSQL(txtEdit(9))) & ",CU118=" & CNULL(ChgSQL(txtEdit(10))) & _
        ",cu125=" & CNULL(ChgSQL(txtEdit(6))) & strSQL1 & _
        ",CU132=" & CNULL(txtEdit(7)) & ",CU145=" & CNULL(txtEdit(11)) & ",CU153=" & CNULL(txtEdit(12)) & _
        ",CU160=" & CNULL(txtComp(0)) & ",CU161=" & CNULL(txtComp(1)) & ",CU162=" & CNULL(txtComp(2)) & ",CU163=" & CNULL(txtComp(3)) & ",CU164=" & CNULL(txtComp(4)) & ",CU165=" & CNULL(txtComp(5)) & _
        ",CU199=" & CNULL(ChgSQL(txtEdit(14))) & " " & _
        " Where CU01='" & Left(txtKey(0), 8) & "' And CU02='" & Right(txtKey(0), 1) & "'"
        
   cnnConnection.BeginTrans
   'add by nickc 2007/03/05
   Pub_SeekTbLog strSql
   'Modified by Lydia 2019/04/23 觸發Trigger
   'cnnConnection.Execute strSql, lngEffRec
   cnnConnection.Execute "begin user_data.user_enabled:=1; " & strSql & " ; end; ", lngEffRec
   
   'Add by Morgan 2009/11/4 更名前的客戶編號的預設聯絡人也要同步
   If Right(txtKey(0), 1) = "0" Then
      'Modify by Amy 2022/08/30 是否寄電子報(CU132)也要同步
      strSql = "update customer a set (cu127,cu132)=(select cu127,cu132 from customer b where b.cu01=a.cu01 and b.cu02='0')" & _
         " where cu01='" & Left(txtKey(0), 8) & "' and cu02<>'0'"
      cnnConnection.Execute strSql, lngEffRec
   End If
   'end 2009/11/4
   
   'Added by Morgan 2024/6/3 全E化客戶任一信箱異動時(含聯絡人)，發信通知智權人員(一併彈提醒)--杜協理
   strECustMsg = ""
   If (txtEdit(1) <> oldCU20 Or txtEdit(8) <> oldCU116 Or txtEdit(9) <> oldCU117 Or txtEdit(10) <> oldCU118) And strCU176 <> MsgText(601) Then
      PUB_ECustEmailChangeInform txtKey(0), strECustMsg
   End If
   'end 2024/6/3
   
   cnnConnection.CommitTrans
   UpdateData = True
   Exit Function
   
ErrHand:

   cnnConnection.RollbackTrans
   MsgBox Err.Description, vbCritical
End Function

Private Function doQuery(ByVal iAct As Integer) As Boolean
   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery = False
   
   'Modify by Amy 2014/05/26 +CU13
   Select Case iAct
      Case 4
      '查詢
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU01='" & cur_CU01 & "' AND CU02='" & cur_CU02 & "'"
         stMessage = "無此記錄之資料！"
   
      Case 6
      '第一筆
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU01||CU02<'" & cur_CU01 & cur_CU02 & "'" & _
            " ORDER BY CU01, CU02"
         stMessage = "已是第一筆了！"

      Case 7
      '上一筆
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU01||CU02<'" & cur_CU01 & cur_CU02 & "'" & _
            " ORDER BY CU01 DESC, CU02 DESC"
         stMessage = "已是第一筆了！"

      Case 8
      '下一筆
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU01||CU02>'" & cur_CU01 & cur_CU02 & "'" & _
            " ORDER BY CU01, CU02"
         stMessage = "已是最後一筆了！"

      Case 9
      '最後筆
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU01||CU02>'" & cur_CU01 & cur_CU02 & "'" & _
            " ORDER BY CU01 DESC, CU02 DESC"
         stMessage = "已是最後一筆了！"
   End Select
   'end 2014/05/26
   
On Error GoTo ErrHand
   
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      'Modify By Sindy 2024/3/11 檔案室
      If PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Then
         lst_CU01 = cur_CU01
         cur_CU01 = "" & rsQuery.Fields(0).Value
         lst_CU02 = cur_CU02
         cur_CU02 = "" & rsQuery.Fields(1).Value
         If ReQuery() = True Then doQuery = True
      'add by sonia 2024/3/22  等級03電腦中心-非程式設計也開放修改
      ElseIf PUB_GetST05(strUserNum) = "03" Then
         lst_CU01 = cur_CU01
         cur_CU01 = "" & rsQuery.Fields(0).Value
         lst_CU02 = cur_CU02
         cur_CU02 = "" & rsQuery.Fields(1).Value
         If ReQuery() = True Then doQuery = True
      'end 2024/3/22
      Else
      '2024/3/11 END
         'Modify By Sindy 2023/9/6
         'Modify By Sindy 2023/12/22 +, , bolSpecMan, strSpecCode
         'Modify By Sindy 2025/4/2 +, , , , Me.Name
         If PUB_ChkSalePerLimit(rsQuery.Fields("cu13").Value, strSalesNo, , bolSpecMan, strSpecCode, , , , Me.Name) = True Then
            If ReQuery() = True Then doQuery = True
         End If
      End If
      
'      Dim strArea As String
'
'      strArea = "" & rsQuery.Fields(2).Value
'      Call GetPLimit(rsQuery.Fields(0).Value, rsQuery.Fields(1).Value) 'Add By Sindy 2010/7/28
'      'edit by nickc 2008/01/18  開放M51 可用
'      'If strArea = strDeptNo Then
'      'Modify by Morgan 2008/12/22 開放夏慧珠不限制
'      '2009/7/2 MODIFY BY SONIA 開放有改客戶異動權限者也可以改
'      'If strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Then
'      'Modify By Sindy 2010/7/28 開放帶人主管不限制
'      'If strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Or PUB_GetST05(strUserNum) = "C2" Or PUB_GetST05(strUserNum) = "C3" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "KM" Or PUB_GetST05(strUserNum) = "K1" Or PUB_GetST05(strUserNum) = "NM" Or PUB_GetST05(strUserNum) = "N1" Then
'      'Modify by Amy 2014/05/26 特殊人員判斷(for 開放專利處部份智權同仁資料給彥葶代為處理)
'      'Modiby by Amy 2015/02/04 +特殊人員(總經理業務工作代理人員)
'      'MODIFY BY SONIA 2015/6/1 分所權限改為部門為M71,另加F2
'      'If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) > 0) Or _
'         (bolSpecMan = True And InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), rsQuery.Fields("CU13")) > 0) Or _
'         bolPLimit = True Or strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Or PUB_GetST05(strUserNum) = "C2" Or _
'         PUB_GetST05(strUserNum) = "C1" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "KM" Or PUB_GetST05(strUserNum) = "K1" Or PUB_GetST05(strUserNum) = "NM" Or PUB_GetST05(strUserNum) = "N1" Then
'      'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
'      If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) > 0) Or _
'         (bolSpecMan = True And InStr(strSpecCode, "總經理業務工作代理人員") > 0 And InStr(Pub_GetSpecMan("總經理員工編號"), rsQuery.Fields("CU13")) > 0) Or _
'         bolPLimit = True Or strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M71" Or strUserNum = "74018" Or _
'         PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Then
'      '2015/6/1 END
'         lst_CU01 = cur_CU01
'         cur_CU01 = "" & rsQuery.Fields(0).Value
'         lst_CU02 = cur_CU02
'         cur_CU02 = "" & rsQuery.Fields(1).Value
'         If ReQuery() = True Then doQuery = True
'      ElseIf bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) = 0 Then
'        MsgBox "您無權限查此客戶！", vbCritical
'      'end 2015/02/04
'      ElseIf bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) = 0 Then
'        MsgBox "您無權限查此客戶！", vbCritical
'      'end 2014/05/26
'      Else
'         MsgBox "業務區別不同不可查詢！", vbCritical
'      End If
   Else
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
End Function

'Modify By Sindy 2023/9/6 mark
''Add By Sindy 2010/7/28 讀取是否為帶人主管權限
'Private Sub GetPLimit(strCU01 As String, strCU02 As String)
'   bolPLimit = False
'   strExc(0) = "select count(*) from staff " & _
'                     "where st01 in (select cu13 from customer where cu01='" & strCU01 & "' and cu02='" & strCU02 & "') " & _
'                     "and st04='2' and (st52='" & strUserNum & "' or st53='" & strUserNum & "' or st54='" & strUserNum & "' or st55='" & strUserNum & "') "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      If RsTemp.Fields(0) > 0 Then
'         bolPLimit = True
'      End If
'   End If
'End Sub

' Add By Sindy 98/02/17 使用客戶名稱做查詢
Private Function doQuery2(ByVal iAct As Integer) As Boolean

   Dim strSql As String, rsQuery As New ADODB.Recordset, stMessage As String
   Dim stIdList As String 'Add by Amy 2014/07/03
   
   rsQuery.MaxRecords = 2
   rsQuery.CursorLocation = adUseClient
   doQuery2 = False
   
   Select Case iAct
      Case 4
      '查詢
         'Modify by Amy 2014/05/26 +CU13
         strSql = "Select CU01, CU02, CU12, CU13 From Customer where CU04 like '%" & Trim(txtKey(1)) & "%' and CU02='0' "
         '2009/7/2 MODIFY BY SONIA 開放有改客戶異動權限者也可以改
         'If Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Then
         'Modify by Amy 特殊人員判斷(for 開放專利處部份智權同仁資料給彥葶代為處理)
         'MODIFY BY SONIA 2015/6/1 分所權限改為部門為M71,另加F2
         'If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0) Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Or PUB_GetST05(strUserNum) = "C2" Or PUB_GetST05(strUserNum) = "C1" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "KM" Or PUB_GetST05(strUserNum) = "K1" Or PUB_GetST05(strUserNum) = "NM" Or PUB_GetST05(strUserNum) = "N1" Then
         'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
         If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0) Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M71" Or strUserNum = "74018" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Then
            '電腦中心或夏慧珠或特殊人員 不限制智權人員代號
         Else
            'Modify by Amy 2014/07/03 簡協理查編號X41770可查,但名稱查不出
            'strSql = strSql & " and CU13='" & strUserNum & "' "
            stIdList = PUB_GetSalesList(strUserNum)
            If InStr(stIdList, ",") = 0 Then
               strSql = strSql & " and CU13=" & stIdList & " "
            Else
               strSql = strSql & " and CU13 in (" & stIdList & " ) "
            End If
            'end 2014/07/03
         End If
         stMessage = "無此客戶名稱之資料！"
   End Select
   
On Error GoTo ErrHand
   
   rsQuery.Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If rsQuery.RecordCount > 1 Then
      MsgBox "此客戶名稱資料重覆，請先至「以申請人查詢案件資料」做查詢！", vbCritical
   ElseIf rsQuery.RecordCount = 1 Then
      Dim strArea As String
      strArea = "" & rsQuery.Fields(2).Value
      txtKey(0).Text = UCase(Left(rsQuery.Fields(0).Value & "000000000", 9))
      cur_CU01 = Left(txtKey(0), 8)
      cur_CU02 = Right(txtKey(0), 1)
      
      'Modify By Sindy 2024/3/11 檔案室
      If PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Then
         lst_CU01 = cur_CU01
         cur_CU01 = "" & rsQuery.Fields(0).Value
         lst_CU02 = cur_CU02
         cur_CU02 = "" & rsQuery.Fields(1).Value
         If ReQuery() = True Then doQuery2 = True
      Else
      '2024/3/11 END
         'Modify By Sindy 2023/9/6
         'Modify By Sindy 2023/12/22 +, , bolSpecMan, strSpecCode
         'Modify By Sindy 2025/4/2 +, , , , Me.Name
         If PUB_ChkSalePerLimit(rsQuery.Fields("cu13").Value, strSalesNo, , bolSpecMan, strSpecCode, , , , Me.Name) = True Then
            If ReQuery() = True Then doQuery2 = True
         End If
      End If
      
'      Call GetPLimit(cur_CU01, cur_CU02) 'Add By Sindy 2010/7/28
'      'edit by nickc 2008/01/18  開放M51 可用
'      'If strArea = strDeptNo Then
'      'Modify by Morgan 2008/12/22 開放夏慧珠不限制
'      '2009/7/2 MODIFY BY SONIA 開放有改客戶異動權限者也可以改
'      'If strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Then
'      'Modify By Sindy 2010/7/28 開放帶人主管不限制
'      'If strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Or strUserNum = "75033" Or PUB_GetST05(strUserNum) = "C2" Or PUB_GetST05(strUserNum) = "C3" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "KM" Or PUB_GetST05(strUserNum) = "K1" Or PUB_GetST05(strUserNum) = "NM" Or PUB_GetST05(strUserNum) = "N1" Then
'      'Modify by Amy 2014/05/26 特殊人員判斷(for 開放專利處部份智權同仁資料給彥葶代為處理)
'      'MODIFY BY SONIA 2015/6/1 分所權限改為部門為M71,另加F2
'      'If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) > 0) Or _
'         bolPLimit = True Or strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or strUserNum = "75033" Or strUserNum = "75033" Or _
'         PUB_GetST05(strUserNum) = "C2" Or PUB_GetST05(strUserNum) = "C1" Or PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "KM" Or PUB_GetST05(strUserNum) = "K1" Or PUB_GetST05(strUserNum) = "NM" Or PUB_GetST05(strUserNum) = "N1" Then
'      'Modify By Sindy 2023/7/31 夏慧珠退休改先開放給杜協理
'      If (bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) > 0) Or _
'         bolPLimit = True Or strArea = strDeptNo Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "M71" Or strUserNum = "74018" Or _
'         PUB_GetST05(strUserNum) = "F1" Or PUB_GetST05(strUserNum) = "F2" Then
'      '2015/6/1 END
'         lst_CU01 = cur_CU01
'         cur_CU01 = "" & rsQuery.Fields(0).Value
'         lst_CU02 = cur_CU02
'         cur_CU02 = "" & rsQuery.Fields(1).Value
'         If ReQuery() = True Then doQuery2 = True
'      ElseIf bolSpecMan = True And InStr(strSpecCode, "A8") > 0 And InStr(Pub_GetSpecMan("A7"), rsQuery.Fields("CU13")) = 0 Then
'        MsgBox "您無權限查此客戶！", vbCritical
'      'end 2014/05/26
'      Else
'         txtKey(0).Text = "X"
'         MsgBox "業務區別不同不可查詢！", vbCritical
'      End If
   Else
      MsgBox stMessage, vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   Exit Function
   
ErrHand:
   
   MsgBox Err.Description, vbCritical
End Function

'完整資料查詢
Private Function ReQuery() As Boolean
   Dim strSql As String, rsQuery As New ADODB.Recordset, intI As Integer
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   
   ReQuery = False
   'Modify by Morgan 2008/12/22 +CU132
   'Add by Sindy 98/03/10 +CU116+CU117+CU118
   'Add By Sindy 2011/1/14 +CU145
   'Add By Sindy 2011/3/17 +CU153
   'Modify by Amy 2016/05/16原:RPAD( NVL(CU30,' '),11,' ')||CU31 R04拆開,並加CU160~165/CU87/CU80/NA03
   'Modified by Morgan 2016/12/13 +CU13,CU166,CU167
   'Modified by Lydia 2021/08/26 + CU21 as E13
   'Modify by Amy 2023/02/15 +CU176 e化客戶指定信箱(正本)
   'Modify by Amy 2023/04/21 +CU10 客戶國籍/CU112 中文地址郵區號/CU79 客戶備註
   'Modify by Amy 2023/05/09 +CU191跨所同意主管(中文字)
   'Modified by Lydia 2024/01/15 +CU199顧問專用信箱
   strSql = "SELECT CU04 R00, CU05||CU88||CU89||CU90 R01, CU16 R02, CU17 R03, CU31 R04" & _
         ", CU23 R05, CU07 R06, CU13 R07, CU12 R08, CU30 R09,CU87 R10,CU80 R11,CU191 R12" & _
         ", CU127, CU20 E01, CU18 E02, CU19 E03, CU22 E04, CU103 E05,CU125 E06,CU132 E07,CU116 E08,CU117 E09,CU118 E10,CU145 E11,CU153 E12,CU21 as E13,CU199 as E14" & _
         ", CU160 C00, CU161 C01, CU162 C02, CU163 C03, CU164 C04, CU165 C05" & _
         ", A.ST02 D00, A0902 D01, DECODE(A.ST04,'1','','（離職）') D02, B.ST02 D03" & _
         ", DECODE(CU85,NULL,NULL,(SUBSTRB(CU85,1,4)-1911)||'/'||SUBSTR(CU85,5,2)||'/'||SUBSTR(CU85,7,2)) D04" & _
         ", FLOOR(CU86/100)||':'||MOD(CU86,100) D05, NA03 D06,CU13,CU166,CU167,CU176,CU10,CU112,CU79" & _
         " From CUSTOMER, STAFF A, STAFF B, ACC090, Nation" & _
         " Where CU01='" & cur_CU01 & "' AND CU02='" & cur_CU02 & "'" & _
         " AND A.ST01(+) = CU13 AND B.ST01(+)=CU84 AND A0901(+) = CU12 And NA01(+)=CU87"

   rsQuery.CursorLocation = adUseClient
   rsQuery.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      txtKey(0) = cur_CU01 & cur_CU02
      'Modify by Amy 2016/05/16 +CU30/CU87/CU80
      'Modify by Amy 2023/05/09 +CU191
      For intI = 1 To 12
         txtRead(intI) = "" & rsQuery.Fields("R" & Format(intI, "00"))
         'Add by Amy 2016/07/04 +記錄地址資料
         If intI = 4 Or intI = 9 Or intI = 10 Then
            txtRead(intI).Tag = "" & rsQuery.Fields("R" & Format(intI, "00"))
         End If
         'end 2016/07/04
      Next intI
      If txtRead(10) <> MsgText(601) Then txtRead_LostFocus (10)
      '客戶狀態有值顯示紅字
      Label1(25).ForeColor = &H80000012: txtRead(11).ForeColor = &H80000012
      If txtRead(11) <> MsgText(601) Then
        Label1(25).ForeColor = &HFF&
        txtRead(11).ForeColor = &HFF&
      End If
      'end 2016/05/16
      
      'Modify By Sindy 98/03/10
      'For intI = 1 To 7
      'Modify By Sindy 2011/1/14
      'For intI = 1 To 10
      'Modify By Sindy 2011/3/17
      'For intI = 1 To 11
      'Modified by Morgan 2012/9/7
      'For intI = 1 To 12
      'Modified by Lydia 2021/08/26
      'For intI = 1 To 13
      'Modified by Lydia 2022/12/12 刪除原本txtEdit(13), 14=>13
      'For intI = 1 To 14
      'Modified by Lydia 2024/01/15 13>14
      For intI = 1 To 14
         txtEdit(intI) = "" & rsQuery.Fields("E" & Format(intI, "00"))
         If intI = 1 Then oldCU20 = txtEdit(intI) 'Add by Amy 2023/02/15 代表信箱
         'Add by Amy 2014/03/06
          If intI = 7 Then oldCU132 = txtEdit(intI)
          If intI = 11 Then oldCU145 = txtEdit(intI)
         'end 2014/03/06
          If intI = 14 Then oldCU199 = txtEdit(intI) 'Added by Lydia 2024/01/15 顧問專用信箱
      Next intI
      
      oldCU116 = txtEdit(8) 'Added by Morgan 2024/6/3
      oldCU117 = txtEdit(9) 'Added by Morgan 2024/6/3
      oldCU118 = txtEdit(10) 'Added by Morgan 2024/6/3
      
      'Add by Amy 2016/05/16
      For intI = 0 To 5
        txtComp(intI) = "" & rsQuery.Fields("C" & Format(intI, "00"))
      Next intI
      
      For intI = 0 To 5
         lblDisp(intI) = "" & rsQuery.Fields("D" & Format(intI, "00"))
      Next intI
      'end 2016/05/16
      'Add by Morgan 2008/7/31
      strCU127 = "" & rsQuery.Fields("CU127")
      txtKey(1) = "" & rsQuery.Fields("R00") ' Add By Sindy 98/02/17
      strCU176 = "" & rsQuery.Fields("CU176") 'Add by Amy 2023/02/15
      'Add by Amy 2023/04/21
      strCU10 = "" & rsQuery.Fields("CU10")
      strCU112 = "" & rsQuery.Fields("CU112")
      strCU79 = "" & rsQuery.Fields("CU79")
      'end 2023/04/21
      'Modify by Amy 2021/12/14 改成Form 2.0
'      PUB_AddContact cur_CU01, cboContact, "" & rsQuery.Fields("CU127")
'      cboContact.Tag = cboContact.ListIndex
      'Modified by Morgan 2022/5/17
      'strExc(10) = cboContact.Tag
      cboContact.Tag = ""
      'end 2022/5/17
       PUB_AddContact cur_CU01, cboContact, "" & rsQuery.Fields("CU127"), True, True, strExc(10)
       cboContact.Tag = strExc(10)
       'end 2021/12/14
      'Added by Morgan 2016/12/13
      'Modified by Morgan 2017/3/10 國內副本收件人改下拉選單
      strCU13 = "" & rsQuery.Fields("cu13")
      cboCU166.Clear
      cboCU166.AddItem "", 0
      '與文雄確認先從嚴控管以避免誤設而錯寄,遇特殊情形再例外設定
      '1.不可為相同客戶 2.只能是客戶的關係企業 3.必須是相同的智權人員
      strExc(0) = "select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where substr(cu01,1,6)='" & Left(txtKey(0), 6) & "' and cu02='0' and cu01<>'" & Left(txtKey(0), 8) & "' and cu13='" & strCU13 & "' order by 1"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = 0
         Do While Not RsTemp.EOF
            cboCU166.AddItem RsTemp("CNo") & " " & RsTemp("CName")
            If RsTemp("CNo") = "" & rsQuery.Fields("cu166") Then
               intI = cboCU166.ListCount - 1
            End If
            RsTemp.MoveNext
         Loop
         If intI > 0 Then
            cboCU166.Tag = "" & intI
            cboCU166.ListIndex = intI
         ElseIf Not IsNull(rsQuery.Fields("cu166")) Then
            strExc(0) = "select cu01||cu02 CNo,nvl(cu04,nvl(rtrim(cu05||' '||cu88||' '||cu89||' '||cu90),cu06)) CName from customer where cu01||cu02='" & rsQuery.Fields("cu166") & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               cboCU166.AddItem RsTemp("CNo") & " " & RsTemp("CName"), 1
               cboCU166.Tag = "1"
               cboCU166.ListIndex = 1
            End If
         End If
      End If
      
      cboCU167.Clear
      If cboCU166.ListIndex >= 0 Then
         PUB_AddContact Left(cboCU166, 8), cboCU167, "" & rsQuery.Fields("CU167")
      End If
      'end 2016/12/13
      ReQuery = True
   Else
      MsgBox "客戶〔" & cur_CU01 & cur_CU02 & "〕已被刪除！", vbCritical
   End If
   
   If rsQuery.State <> adStateClosed Then rsQuery.Close
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description, vbCritical
   Screen.MousePointer = vbDefault
End Function

Private Sub txtComp_GotFocus(Index As Integer)
    TextInverse txtComp(Index)
End Sub

Private Sub txtComp_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
   ' Added by Lydia 2020/03/31 事務所合併日起台灣案取消(1:專利商標 2:專利法律) 的標題，非台灣案改標題為(J:智權公司 空白:系統預設)。
    If txtComp(Index).Locked = False Then
        If strSrvDate(1) >= 事務所合併日 Then
            If Index = 1 Or Index = 3 Or Index = 5 Then
                 If KeyAscii <> 8 And KeyAscii <> Asc("J") Then
                    KeyAscii = 0
                    Beep
                End If
            End If
        Else
    'end 2020/03/31
            'Modify by Amy 2016/05/16 '預設收據公司台灣不可輸J
            If Index = 0 Or Index = 2 Or Index = 4 Then
                 If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
                    KeyAscii = 0
                    Beep
                End If
            Else
                If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("J") Then
                    KeyAscii = 0
                    Beep
                End If
            End If
    'Added by Lydia 2020/03/31
        End If
    End If
    'end 2020/03/31
    
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
   TextInverse txtEdit(Index)
   If txtEdit(Index).Locked = False Then
      Select Case Index
         'edit by nickc 2008/01/18
         'Case 0
         Case 6
         '接洽人
            'edit by nickc 2007/06/06 切換輸入法改用API
            'txtEdit(Index).IMEMode = 1
            OpenIme
         Case Else
            'edit by nickc 2007/06/06 切換輸入法改用API
            'txtEdit(Index).IMEMode = 2
            CloseIme
      End Select
   End If
End Sub

'Add by Morgan 2008/12/22
'Modify by Amy 2021/12/14 原:Integer
Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 7
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      'Add By Sindy 2011/1/14
      Case 11
         KeyAscii = UpperCase(KeyAscii)
         'Modified by Morgan 2012/1/2 改放 N
         'If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         If KeyAscii <> 8 And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
      'Add By Sindy 2011/3/17
      Case 12
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
   Dim stMsg As String 'Add by Amy 2024/04/10
   
   Select Case Index
      Case 1
      'E-Mail
         If CheckLengthIsOK(txtEdit(1), 50) = False Then
            Cancel = True
            Exit Sub
         End If
         'Add By Sindy 2012/8/28
         If txtEdit(1).Text = "" Or iCurState <> 2 Then Exit Sub
         Cancel = Not PUB_CheckMail(txtEdit(1).Text)
         Call txtEdit_GotFocus(Index)
         Exit Sub
         '2012/8/28 End
      Case 2
      '傳真1
         If CheckLengthIsOK(txtEdit(2), 20) = False Then
            Cancel = True
            Exit Sub
         End If
      Case 3
      '傳真2
         If CheckLengthIsOK(txtEdit(3), 20) = False Then
            Cancel = True
            Exit Sub
         End If
      Case 4
      '手機
         If CheckLengthIsOK(txtEdit(4), 15) = False Then
            Cancel = True
            Exit Sub
         End If
      'Add by Morgan 2004/4/21
      Case 5
      '負責人英文
         '接洽紀錄單只能印16個英文
         '2011/5/18 MODIFY BY SONIA X29285超過16碼不能輸,改為可完整輸,接洽單控制舊客戶帶出時若超過16碼則帶14碼+..,新客戶仍只可輸16碼
         'If CheckLengthIsOK(txtEdit(5), 16) = False Then
         If CheckLengthIsOK(txtEdit(5), txtEdit(5).MaxLength) = False Then
            Cancel = True
            Exit Sub
         End If
       'add by nickc 2008/01/18
       Case 6
         If CheckLengthIsOK(txtEdit(6), txtEdit(6).MaxLength) = False Then
            Cancel = True
            Exit Sub
         End If
      'Add by Amy 2024/04/10 只能輸N,小寫或全型都不允許
      'Modify by Amy 2024/05/14 顧問電子報可輸Y 拆開
      Case 7, 11
         If txtEdit(Index) <> "N" And txtEdit(Index) <> MsgText(601) Then
            If Index = 7 Then
               stMsg = "國內電子報"
            ElseIf Index = 11 Then
               stMsg = "專利雙週報"
            End If
            MsgBox stMsg & "只允許輸入N,不可輸小寫或全型..."
            Cancel = True
            Call txtEdit_GotFocus(Index)
            Exit Sub
         End If
'Add By Sindy 98/03/10
      Case 8
      'E-Mail(其他1)
         If CheckLengthIsOK(txtEdit(8), 50) = False Then
            Cancel = True
            Exit Sub
         End If
         'Add By Sindy 2012/8/28
         If txtEdit(8).Text = "" Or iCurState <> 2 Then Exit Sub
         Cancel = Not PUB_CheckMail(txtEdit(8).Text)
         Call txtEdit_GotFocus(Index)
         Exit Sub
         '2012/8/28 End
      Case 9
      'E-Mail(其他2)
         If CheckLengthIsOK(txtEdit(9), 50) = False Then
            Cancel = True
            Exit Sub
         End If
         'Add By Sindy 2012/8/28
         If txtEdit(9).Text = "" Or iCurState <> 2 Then Exit Sub
         Cancel = Not PUB_CheckMail(txtEdit(9).Text)
         Call txtEdit_GotFocus(Index)
         Exit Sub
         '2012/8/28 End
      Case 10
      'E-Mail(其他3)
         If CheckLengthIsOK(txtEdit(10), 50) = False Then
            Cancel = True
            Exit Sub
         End If
         'Add By Sindy 2012/8/28
         If txtEdit(10).Text = "" Or iCurState <> 2 Then Exit Sub
         Cancel = Not PUB_CheckMail(txtEdit(10).Text)
         Call txtEdit_GotFocus(Index)
         Exit Sub
         '2012/8/28 End
      Case 12 'Add by Amy 2024/05/14 顧問電子報
         If txtEdit(Index) <> "N" And txtEdit(Index) <> "Y" And txtEdit(Index) <> MsgText(601) Then
            MsgBox stMsg & "只允許輸入N或Y,不可輸小寫或全型..."
            Cancel = True
            Call txtEdit_GotFocus(Index)
            Exit Sub
         End If
'98/03/10 End
   End Select
End Sub

Private Sub txtkey_GotFocus(Index As Integer)
   If Index = 0 And txtKey(Index) <> "" And txtKey(Index).Locked = False Then
      txtKey(Index).SelStart = 1
      txtKey(Index).SelLength = Len(txtKey(Index)) - 1
   Else
      TextInverse txtKey(Index)
   End If
   Select Case Index
      Case 0
         If txtKey(Index).Enabled = True Then CloseIme
      Case 1
         If txtKey(Index).Enabled = True Then OpenIme
   End Select
End Sub

'Modify by Amy 2021/12/14 原:Integer
Private Sub txtKey_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TxtKey_LostFocus(Index As Integer)
   Select Case Index
      Case 0
         'Add by Amy 2021/12/15 改 Form2.0 後清資料會觸發此
         If iCurState = 4 And (Trim(Me.txtKey(0).Text) = "" Or Trim(Me.txtKey(0).Text) = "X") Then Exit Sub
         
         If Trim(Me.txtKey(0).Text) <> "" And Trim(Me.txtKey(0).Text) <> "X" Then
            txtKey(Index) = UCase(Left(Me.txtKey(0).Text & "000000000", 9))
         End If
   End Select
End Sub

Private Sub txtkey_Validate(Index As Integer, Cancel As Boolean)
   CloseIme
End Sub

Private Sub txtRead_GotFocus(Index As Integer)
    'Add by Amy 2016/05/16
    If Index = 10 Or Index = 9 Then
        CloseIme
    ElseIf Index = 4 Then
        OpenIme
    End If
    'end 2016/05/16
    TextInverse Me.txtRead(Index)
End Sub

'Add by Morgan 2008/8/11
'檢查預設接洽人聯絡地址是否與客戶聯絡地址一致
Private Function CheckContactAddr() As Boolean
         
   If Val(strCU127) = 0 Then
      CheckContactAddr = True
   Else
      strExc(0) = "select cu30,cu31,pcc21,pcc22 from customer,potcustcont" & _
         " where cu01='" & cur_CU01 & "' and cu02='0' and pcc01(+)=cu01 and pcc02(+)='" & strCU127 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With RsTemp
         If IsNull(.Fields("pcc21")) Then
            CheckContactAddr = True
         'Modify by Amy 2016/05/16 改比較畫面上欄位
         'ElseIf "" & .Fields("cu30") = "" & .Fields("pcc21") And "" & .Fields("cu31") = "" & .Fields("pcc22") Then
         'Modifiedd by Morgan 2022/9/2 為避免客戶聯絡地址變更後造成與預設接洽人不同改預設接洽人不可設聯絡地址 Ex:X44483
         'ElseIf Trim(txtRead(4)) = Trim("" & .Fields("pcc22")) And Trim(txtRead(9)) = Trim("" & .Fields("pcc21")) Then
         '   CheckContactAddr = True
         'Else
         '   MsgBox "預設接洽人的聯絡地址【" & .Fields("pcc21") & " " & .Fields("pcc22") & "】" & vbCrLf & "與客戶聯絡地址必須相同！"
         Else
            MsgBox "預設接洽人不可有【聯絡地址】！" & vbCrLf & "請先設定接洽人【" & cboContact & "】的【聯絡地址同客戶】！", vbExclamation
         'end 2022/9/2
         End If
         End With
      End If
   End If
End Function

'Add by Amy 2016/05/16 開放可修改電話及聯絡地址
'Modify by Amy 2021/12/14 原:Integer
Private Sub txtRead_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
    Dim strAddr As String, strNewArea As String, strZipCode As String, strCountry As String, strROC As String, strIndArea As String
    Dim intArea As Integer, intFocus As Integer
    Dim bolMany As Boolean

    Select Case Index
        '聯絡地址
        Case 4
            'Modify by Amy 2021/12/13 +txtRead(Index),輸數字轉全型時會帶出其他的字
            KeyAscii = ChangeZIP(KeyAscii, txtRead(Index))
            '臺灣地址判斷
            If LTrim(Me.txtRead(Index)) <> MsgText(601) And Me.txtRead(10) < "010" Then
                strROC = ""
                strAddr = Me.txtRead(Index)
                If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
                If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
                If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
                '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
                strIndArea = "True"
                strAddr = ReplaceIndArea(strAddr, strIndArea)
                If strIndArea = "True" Then strIndArea = MsgText(601)
                If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
                    strIndArea = "新竹" & strIndArea
                    strAddr = Mid(strAddr, 3)
                End If

                If Len(LTrim(strAddr)) > 4 And (Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣") Then
                    '輸到路/街/段.
                    If Asc("路") = KeyAscii Or Asc("街") = KeyAscii Or Asc("段") = KeyAscii Then
                        intFocus = Val(Me.txtRead(Index).SelStart) - Len(strROC) - Len(strIndArea)
                        strAddr = Mid(strAddr, 1, intFocus) & Chr(KeyAscii) & Mid(strAddr, intFocus + 1) 'KeyPress未完成時地址欄位尚未顯示目前字,故先加入當下的字查
                        '有鄉/鎮/市/區
                        'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉 X80024
                        If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
                            Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
                            Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                             strZipCode = GetZipCode_Tai(1, strAddr, , bolMany, , strCountry)
                            If strZipCode <> MsgText(601) Then
                                If bolMany = False Then
                                    Call ChkZipData(2, Me.txtRead(Index), strZipCode, , strCountry)
                                    Me.txtRead(Index).SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                    Me.txtRead(Index).SelLength = 0
                                Else
                                    '多筆以縣/市+鄉/鎮/市/區及路名查
                                    bolMany = False
                                    strZipCode = GetZipCode_Tai(3, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, , strCountry)
                                    If strZipCode <> MsgText(601) And bolMany = False Then
                                        Call ChkZipData(2, Me.txtRead(Index), strZipCode, , strCountry)
                                        Me.txtRead(Index).SelStart = intFocus + Len(strROC) + Len(strIndArea)
                                        Me.txtRead(Index).SelLength = 0
                                    End If
                                End If
                            End If
                        '沒鄉/鎮/市/區
                        Else
                            '取 段/路/街 查
                            strZipCode = GetZipCode_Tai(2, Mid(strAddr, 1, intFocus + 1), intArea, bolMany, strNewArea, strCountry)
                            If strZipCode <> MsgText(601) And bolMany = False Then
                                '補上查到的區,避免輸入兩個同樣的字(路/街/段)被取代,故不用Replace
                                Me.txtRead(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4, intArea - 4) & Mid(strAddr, intFocus + 2)
                                Call ChkZipData(2, Me.txtRead(Index), strZipCode, , strCountry)
                                Me.txtRead(Index).SelStart = intFocus + Len(strROC) + Len(strNewArea) + Len(strIndArea)
                                Me.txtRead(Index).SelLength = 0
                            End If
                        End If

                    End If
                End If
            End If
        '聯絡/申請郵遞區號
        Case 9
            'Modify by Amy 2021/12/13 +txtRead(Index),輸數字轉全型時會帶出其他的字
            KeyAscii = ChangeZIP(KeyAscii, txtRead(Index))
    End Select
End Sub

'Modify by Amy 2021/12/14 原:ByRef objTxt As TextBox
Private Sub ChkZipData(ByVal intChoose As Integer, ByRef objTxt As Control, Optional ByRef stZipCode As String = "", Optional ByRef intArea As Integer = 0, Optional ByRef stCountryCode As String = "")
    Dim intZipIdx As Integer, intCountryIdx As Integer '地址相對應Zip欄位/國籍欄位Index
    Dim strMsg As String, strAddr As String
    Dim intCount As Integer
    
     Select Case objTxt.Index
        Case 4
            strMsg = "聯絡地址"
            intZipIdx = 9
            intCountryIdx = 10
        Case Else
            intZipIdx = objTxt.Index
    End Select
    
    Select Case intChoose
        Case 1 'ZipCode多筆(同區/鄉 ZipCode不同)
            '且與畫面上欄位資料前3碼不同或空值,彈郵遞區號查詢畫面
            If InStr(stZipCode, Left(Trim(txtRead(intZipIdx)), 3)) = 0 Or Trim(txtRead(intZipIdx)) = MsgText(601) Then
                If Trim(txtRead(intZipIdx)) <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,請選擇正確郵遞區號！"
                Call frm100134.SetParent(Me)
                Me.Hide
                frm100134.BFormZip = "txtRead(" & intZipIdx & ")"
                frm100134.GetStreet objTxt.Text, 1, intArea, stZipCode
                Call frm100134.QueryData
                frm100134.Show
                Exit Sub
            End If
        Case 2 'ZipCode非多筆
            '判斷抓到的郵遞區號是否與畫面上欄位資料前3碼相同
            If Left(txtRead(intZipIdx), 3) <> stZipCode Then
                If txtRead(intZipIdx) <> MsgText(601) Then MsgBox strMsg & "郵遞區號有誤,系統將自動更正！", , MsgText(5)
                txtRead(intZipIdx) = stZipCode
                txtRead_GotFocus (objTxt.Index)
            End If
            If txtRead(intCountryIdx) <> MsgText(601) And stCountryCode <> txtRead(intCountryIdx) Then
                MsgBox strMsg & "國籍有誤,系統將自動更正！", , MsgText(5)
                txtRead(intCountryIdx) = stCountryCode
                txtRead_LostFocus (intCountryIdx)
            End If
            Exit Sub
        Case 3, 4, 5, 6 '抓不到ZipCode-3.區錯/4.只有路且郵遞區號為多筆/5.抓到2個字縣市,但多筆/6.舊客戶ZipCode多筆
            SSTab1.Tab = 0
            MsgBox strMsg & "無法解析郵遞區號，請由下一畫面選取！"
            Call frm100134.SetParent(Me)
            Me.Hide
            frm100134.BFormZip = "txtRead(" & intZipIdx & ")"
            frm100134.GetStreet objTxt.Text, IIf(intChoose = 6, 2, intChoose), intArea, stZipCode
            Call frm100134.QueryData
            frm100134.Show
            Exit Sub
        Case 9 '設定頁籤
            SSTab1.Tab = 0
            If txtRead(objTxt.Index).Enabled = True Then
                txtRead(objTxt.Index).SetFocus
                txtRead_GotFocus (objTxt.Index)
            End If
    End Select
End Sub

Public Sub txtRead_LostFocus(Index As Integer)
    Dim strZipCode As String, strAddr As String, strCountry As String, strCityN As String, strIndArea As String, strNewArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer

    If iCurState <> 2 Then Exit Sub
    If Index <> 4 And Index <> 9 And Index <> 10 Then Exit Sub
    
    If Me.txtRead(Index) = MsgText(601) Then
        If Index = 10 Then Me.lblDisp(6).Caption = ""
        Exit Sub
    End If
    
    'Memo by Amy 2021/12/14 地址、郵遞區號、國籍有修改需確認 接洽單/代理人/客戶檔是否改一致
    Select Case Index
        '聯絡地址
        Case 4
            If Me.txtRead(10) >= "010" Then Exit Sub
            'Add by Amy 2016/07/04 +判斷地址相關沒修改不檢查
            If Me.txtRead(Index) = Me.txtRead(Index).Tag And Me.txtRead(9) = Me.txtRead(9).Tag And Me.txtRead(10) = Me.txtRead(10).Tag Then
              Exit Sub
            End If
            Me.txtRead(Index) = ReplaceAddrTW(Me.txtRead(Index))
AgainCheck1:
            strROC = ""
            strAddr = Me.txtRead(Index)
            If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
            If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
            If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
            '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
            strIndArea = "True"
            strAddr = ReplaceIndArea(strAddr, strIndArea)
            If strIndArea = "True" Then strIndArea = MsgText(601)
            If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
                strIndArea = "新竹" & strIndArea
                strAddr = Mid(strAddr, 3)
            End If
            '第3個字是 縣 / 市
            If Mid(strAddr, 3, 1) = "市" Or Mid(strAddr, 3, 1) = "縣" Or Mid(strAddr, 1, 3) = "釣魚臺" Or Mid(strAddr, 1, 3) = "海南島" Then
                'Modify by Amy 2018/12/19 +判斷第七個字 ex:嘉義縣阿里山鄉樂野村 X80024
                If Mid(strAddr, 7, 1) = "市" Or Mid(strAddr, 7, 1) = "區" Or Mid(strAddr, 7, 1) = "鄉" Or Mid(strAddr, 7, 1) = "鎮" _
                    Or Mid(strAddr, 6, 1) = "市" Or Mid(strAddr, 6, 1) = "區" Or Mid(strAddr, 6, 1) = "鄉" Or Mid(strAddr, 6, 1) = "鎮" _
                    Or Mid(strAddr, 5, 1) = "市" Or Mid(strAddr, 5, 1) = "區" Or Mid(strAddr, 5, 1) = "鄉" Or Mid(strAddr, 5, 1) = "鎮" Then
                    'Modify by Amy 2021/12/14 判斷第七個字
                    '傳入地址前7個字抓到郵遞區號
                    intArea = 7
                    strZipCode = GetPostZip(Left(strAddr, 7), 7, , strCountry, bolMany)
                    '傳入地址前6個字抓到郵遞區號
                    If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany): intArea = 6
                    'end 201
                    intArea = 6
                    strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
                    '傳入地址前5個字取郵遞區號
                    If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
                    '抓到郵遞區號
                    If strZipCode <> MsgText(601) Then
                        If bolMany = True Then
                            '多筆以縣/市+鄉/鎮/市/區及路名查
                            bolMany = False
                            strZipCode = GetZipCode_Tai(3, strAddr, , bolMany, , strCountry)
                            If strZipCode <> MsgText(601) Then
                                '限制縣/市+鄉/鎮/市/區及路名查:一筆-直接帶/多筆-進查詢畫面
                                If bolMany = False Then
                                    Call ChkZipData(2, Me.txtRead(Index), strZipCode, , strCountry)
                                Else
                                    Call ChkZipData(1, Me.txtRead(Index), strZipCode, intArea, strCountry)
                                End If
                            End If
                        Else
                            '非多筆
                            Call ChkZipData(2, Me.txtRead(Index), strZipCode, intArea, strCountry)
                        End If
                    Else
                        '判斷是否有此區/鄉/鎮
                        strZipCode = GetPostZip(Mid(strAddr, 4, intArea - 3), intArea - 3, , strCountry, bolMany, "Pzd03")
                        If strZipCode <> MsgText(601) Then
                            '區別錯,進入查詢畫面
                            Call ChkZipData(3, Me.txtRead(Index), strZipCode, intArea, strCountry)
                        Else
                            '當作沒區只有路 ex:新竹縣or市園區二路
                            bolMany = False
                            strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
                            If strZipCode <> MsgText(601) Then
                                '以縣/市及路名查:一筆-直接帶/多筆-進查詢畫面
                                If bolMany = False Then
                                    Me.txtRead(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4)
                                    Call ChkZipData(2, Me.txtRead(Index), strZipCode, intArea, strCountry)
                                Else
                                    intArea = 0
                                    Call ChkZipData(4, Me.txtRead(Index), strZipCode, intArea, strCountry)
                                    Exit Sub
                                End If
                            End If
                        End If
                    End If
                '無鄉/鎮/市/區
                Else
                    '以路/街 抓是否有zip
                    strZipCode = GetZipCode_Tai(2, strAddr, intArea, bolMany, strNewArea, strCountry)
                    If strZipCode <> MsgText(601) Then
                        If bolMany = True Then
                            '多筆
                            intArea = 0
                            Call ChkZipData(4, Me.txtRead(Index), strZipCode, intArea, strCountry)
                        Else
                            '非多筆
                            Me.txtRead(Index) = strROC & Left(strAddr, 3) & strNewArea & strIndArea & Mid(strAddr, 4)
                            Call ChkZipData(2, Me.txtRead(Index), strZipCode, intArea, strCountry)
                        End If
                    End If
                End If
                '都抓不到ZipCode
                If strZipCode = MsgText(601) Then
                    If CheckTaiwanAddr_Tai(Me.txtRead(Index), Me.txtRead(Index - 1), "000", "", strZipCode, , False, Me.Name) = False Then
                        If strZipCode = "格式錯誤" Then
                            'Modify by Amy 2021/12/14 改Form2.0 時發現未輸入鄉、鎮、區,游標點至國籍會一直觸發,無法關閉frm100135,改同客戶檔維護
                            If PUB_CheckFormExist("frm100135") = False Then
                                frm100135.Show vbModal
                                Call ChkZipData(9, Me.txtRead(Index), strZipCode)
                            End If
                        Else
                            Call ChkZipData(3, Me.txtRead(Index), strZipCode)
                        End If
                        Exit Sub
                    End If
                        
                End If
            '第三3個字無 縣 / 市
            Else
                '傳入地址前2個字判斷是否有其縣/市
                strCityN = "Pzd02"
                strZipCode = GetPostZip(Left(strAddr, 2), 2, 1, strCountry, bolMany, "Pzd02", strCityN)
                If strZipCode <> MsgText(601) Then
                    If bolMany = False Then
                        '只有一筆
                        Me.txtRead(Index) = strROC & strCityN & strIndArea & Mid(strAddr, 3)
                        GoTo AgainCheck1
                    Else
                        '新竹、嘉義會有2筆
                        intArea = 0
                        Call ChkZipData(5, Me.txtRead(Index), strZipCode, intArea)
                    End If
                End If
            End If
        '聯絡地址國籍
        Case 10
            Me.lblDisp(6).Caption = ""
            If Me.txtRead(Index).Text = 台灣國家代號 Then
                ShowMsg "地址" & MsgText(9153)
                Me.txtRead(Index).SetFocus
                txtRead_GotFocus (Index)
                Exit Sub
            End If
            strExc(0) = ""
            If ClsPDGetNation(Me.txtRead(Index), strExc(0)) = True Then
                Me.lblDisp(6).Caption = strExc(0)
            Else
                txtRead_GotFocus (Index)
                Exit Sub
            End If
    End Select
End Sub

'Add by Amy 2016/05/16
Private Function FormCheck() As Boolean
    Dim strZipCode As String, strAddr As String, strCountry As String, strIndArea As String, strROC As String
    Dim bolMany As Boolean, intArea As Integer
    Dim strTpZip(2) As String  'Add by Amy 2023/04/21
    Dim bCancel As Boolean 'Add by Amy 2024/04/10
    
    FormCheck = False
    strAcrossAreaMail = "": bolAcrossArea = False 'Add by Amy 2023/04/21
    
    'Add by Amy 2021/12/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, True, True) = False Then
        Exit Function
    End If
    
    'Add by Amy 2025/02/10
    'TEL2[不是]空且TEL1[是]空彈訊息
    If Trim(txtRead(3)) <> MsgText(601) And Trim(txtRead(2)) = MsgText(601) Then
       MsgBox "電話２有值,電話１不可為空！"
       txtRead(3).SetFocus
       txtRead_GotFocus (3)
       SSTab1.Tab = 0
       Exit Function
    End If
    'FAX2[不是]空且FAX1[是]空彈訊息
    If Trim(txtEdit(3)) <> MsgText(601) And Trim(txtEdit(2)) = MsgText(601) Then
       MsgBox "傳真２有值,傳真１不可為空！"
       txtEdit(3).SetFocus
       txtEdit_GotFocus (3)
       SSTab1.Tab = 0
       Exit Function
    End If
    
    If Me.txtRead(10).Text = 台灣國家代號 Then
        ShowMsg "地址" & MsgText(9153)
        Me.txtRead(10).SetFocus
        txtRead_GotFocus (10)
        Exit Function
    End If
    strExc(0) = ""
    If ClsPDGetNation(Me.txtRead(10), strExc(0)) = False Then
        Me.txtRead(10).SetFocus
        Exit Function
    End If
    'end 2018/10/25
    
    'Add by Amy 2023/05/09 地址不可有刪址字樣
    If InStr(Me.txtRead(4), "刪址") > 0 Then
        MsgBox "聯絡地址不可有「刪址」字樣！"
        Me.txtRead(4).SetFocus
        Exit Function
    End If
    
    'Modify by Amy 2018/11/05 從上面搬下來 ex:X78656 地址改為020時台灣地址檢查要跳過
    'Modify by Amy 2018/10/25 非台灣國籍也要檢查
    If Me.txtRead(10) >= "010" Then FormCheck = True: Exit Function
    If Me.txtRead(4) = MsgText(601) Then
        MsgBox "國籍屬於臺灣聯絡地址不可為空！"
        Me.txtRead(4).SetFocus
        Exit Function
    End If
    
    If Me.txtRead(9) = MsgText(601) Then
        MsgBox "國籍屬於臺灣聯絡地址郵遞區號不可為空！"
        Me.txtRead(9).SetFocus
        Exit Function
    End If
    
    'Add by Amy 2016/07/04 +地址相關欄位有修改才檢查
    If Me.txtRead(4) <> Me.txtRead(4).Tag Or Me.txtRead(9) <> Me.txtRead(9).Tag Or Me.txtRead(10) <> Me.txtRead(10).Tag Then
        '聯絡地址判斷
        Me.txtRead(4) = ReplaceAddrTW(Me.txtRead(4))
        strROC = ""
        strAddr = Me.txtRead(4)
        If Left(strAddr, 4) = "中華民國" Then strROC = strROC & Left(strAddr, 4): strAddr = Mid(strAddr, 5)
        If Left(strAddr, 3) = "臺灣省" Or Left(strAddr, 3) = "台灣省" Then strROC = strROC & Left(strAddr, 3): strAddr = Mid(strAddr, 4)
        If Left(strAddr, 2) = "臺灣" Or Left(strAddr, 2) = "台灣" Then strROC = strROC & Left(strAddr, 2): strAddr = Mid(strAddr, 3)
        '去除xx工業區查(台中工業區/台塑工業園區不取代,可能抓錯zip)
        strIndArea = "True"
        strAddr = ReplaceIndArea(strAddr, strIndArea)
        If strIndArea = "True" Then strIndArea = MsgText(601)
        If Left(strAddr, 4) = "新竹新竹" And (strIndArea = "科學工業園區" Or strIndArea = "科學園區") Then
            strIndArea = "新竹" & strIndArea
            strAddr = Mid(strAddr, 3)
        End If
        intArea = 6
        strZipCode = GetPostZip(Left(strAddr, 6), 6, , strCountry, bolMany)
        '傳入地址前5個字取郵遞區號
        If strZipCode = MsgText(601) Then strZipCode = GetPostZip(Left(strAddr, 5), 5, , strCountry, bolMany): intArea = 5
        If InStr(strZipCode, Left(Me.txtRead(9), 3)) = 0 Then MsgBox "地址對應之郵遞區號有誤請確認！": Me.txtRead(4).SetFocus: Exit Function
        If Me.txtRead(10) <> strCountry Then MsgBox "地址對應之國籍有誤請確認！": Me.txtRead(4).SetFocus: Exit Function
        
        If CheckTaiwanAddr(Me.txtRead(4).Text, Me.txtRead(10).Text, "聯絡地址") = False Then
             Call ChkZipData(9, Me.txtRead(4))
             Exit Function
        End If
    End If
    
    'Add by Amy 2023/04/21 +修改的聯絡地址郵遞區號與智權人員的所別不同時,發mail 通知「程式管理人員」
    'Modify by Amy 2023/06/06 +if  有修改地址才檢查
    If Me.txtRead(4) <> Me.txtRead(4).Tag Then
         'Modify by Amy 2024/06/28 國籍改傳入ChkAcrossArea判斷(for 項2判斷),cu79 有[臺灣地址格式不檢查]則以[非]臺灣判斷
         '聯絡地址國籍(新)
         'If Me.txtRead(10) < "010" Then
             strTpZip(0) = Me.txtRead(9)
         'End If
         '申請地址國籍
         If strCU10 < "010" And InStr(strCU79, "臺灣地址格式不檢查") > 0 Then strCU10 = "999"
         'If strCU10 < "010" Then
             strTpZip(1) = strCU112
         'End If
         '聯絡地址國籍(舊)
         'If Me.txtRead(10).Tag < "010" Then
             strTpZip(2) = Me.txtRead(9).Tag
         'End If
                  
         '1.申請[是]地跨所且聯絡地由[不是]跨所改跨所,需mail 通知電腦中心
         '2.客戶國籍及地址國籍有一個不是台灣,若其中地址由 非跨所->跨所 ex:X08225050 客戶國籍大陸 聯絡地址 由 新北市->臺南市
         'Memo 舊資料 原跨所 未有 跨所同意主管 資料,改地址後仍「是」跨所,也「不需」補簽核-1120510 秀玲
         '             舊資料 原跨所 未有 跨所同意主管 資料,改地址後「不是」跨所,又再改回 跨所 就「需」補簽核
         bolAcrossArea = ChkAcrossArea(1, Me.Name, Me.txtRead(7), , strTpZip(0), strTpZip(1), strTpZip(2), strCU10, Me.txtRead(10), strCU10, Me.txtRead(10).Tag)
         'end 2024/06/28
         'Modify by Amy 2023/05/09 +跨所同意主管為空,才需mail
         If bolAcrossArea = True And Me.txtRead(12) = MsgText(601) Then
             'Add by Amy 2023/05/11
             If MsgBox("新地址為跨所請呈報主管，並將呈報結果寄給電腦中心做後續處理" & vbCrLf & _
                 "是：確定繼續操作　否：再修改地址？", vbExclamation + vbYesNo) = vbNo Then
                 Exit Function
             End If
             If txtKey(1) <> MsgText(601) Then
                 strAcrossAreaMail = strAcrossAreaMail & "　客戶名稱：" & txtKey(1) & vbCrLf
             Else
                 strAcrossAreaMail = strAcrossAreaMail & "　客戶名稱：" & txtRead(1) & vbCrLf
             End If
             strAcrossAreaMail = strAcrossAreaMail & "原聯絡地址：" & Me.txtRead(4).Tag & vbCrLf
             strAcrossAreaMail = strAcrossAreaMail & "新聯絡地址：" & Me.txtRead(4) & vbCrLf
             strAcrossAreaMail = strAcrossAreaMail & "　中文地址：" & Me.txtRead(5) & vbCrLf
             strAcrossAreaMail = strAcrossAreaMail & "　客戶狀態：" & Me.txtRead(11) & vbCrLf
             strAcrossAreaMail = strAcrossAreaMail & "　參考備註：" & strCU79 & vbCrLf
         End If
    End If
    'end 2023/04/21
    'Add by Amy 2024/04/10 電子報相關欄位檢查
    '國內電子報
    If txtEdit(9).Text <> MsgText(601) Then
      Call txtEdit_Validate(9, bCancel)
      If bCancel = True Then Exit Function
    End If
    '專利雙週報
    If txtEdit(11).Text <> MsgText(601) Then
      Call txtEdit_Validate(11, bCancel)
      If bCancel = True Then Exit Function
    End If
    '顧問電子報
    If txtEdit(12).Text <> MsgText(601) Then
      Call txtEdit_Validate(12, bCancel)
      If bCancel = True Then Exit Function
    End If
    'end 2024/04/10
    'Added by Lydia 2024/01/15 檢查顧問專用信箱
    If oldCU199 <> txtEdit(14).Text Then
       If Trim(txtEdit(14).Text) <> "" Then
          If Trim(txtEdit(14).Text) <> "無信箱" Then
             If PUB_CheckMail(txtEdit(14)) = False Then
                Exit Function
             End If
          Else
             txtEdit(12).Text = "N" '預設不寄顧問電子報
          End If
       End If
    End If
    'end 2024/01/15
    
    FormCheck = True
End Function
