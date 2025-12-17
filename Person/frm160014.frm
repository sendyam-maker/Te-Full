VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm160014 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工指紋卡片資料"
   ClientHeight    =   5952
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7512
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5952
   ScaleWidth      =   7512
   Begin VB.TextBox txtST04 
      Alignment       =   2  '置中對齊
      Enabled         =   0   'False
      Height          =   270
      Left            =   1032
      MaxLength       =   5
      TabIndex        =   50
      Top             =   1368
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刪除考勤機指紋卡片"
      Height          =   384
      Left            =   5544
      TabIndex        =   48
      Top             =   96
      Width           =   1896
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   768
      IntegralHeight  =   0   'False
      ItemData        =   "frm160014.frx":0000
      Left            =   72
      List            =   "frm160014.frx":0002
      TabIndex        =   45
      Top             =   5088
      Width           =   4872
   End
   Begin VB.CheckBox Check2 
      Caption         =   "全時段(開捲門)"
      Height          =   225
      Left            =   4080
      TabIndex        =   42
      Top             =   1056
      Width           =   1572
   End
   Begin VB.CheckBox Check1 
      Caption         =   "全時段(不開捲門)"
      Height          =   225
      Left            =   4080
      TabIndex        =   41
      Top             =   720
      Width           =   1740
   End
   Begin VB.CommandButton cmdChgTimeZone 
      Caption         =   "時段回寫門禁機"
      Height          =   408
      Left            =   5832
      TabIndex        =   40
      Top             =   1008
      Width           =   1536
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   180
      Left            =   3930
      TabIndex        =   38
      Text            =   "考勤機："
      Top             =   1668
      Width           =   765
   End
   Begin VB.ComboBox cboHtaIp 
      Height          =   276
      Left            =   4710
      TabIndex        =   37
      Text            =   "cboHtaIp"
      Top             =   1608
      Width           =   2685
   End
   Begin VB.TextBox txtST06 
      Alignment       =   2  '置中對齊
      Enabled         =   0   'False
      Height          =   270
      Left            =   1035
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1032
      Width           =   465
   End
   Begin VB.TextBox txtST01 
      Height          =   270
      Left            =   1035
      MaxLength       =   5
      TabIndex        =   0
      Top             =   690
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3396
      Left            =   96
      TabIndex        =   20
      Top             =   1680
      Width           =   7308
      _ExtentX        =   12891
      _ExtentY        =   5990
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "指紋"
      TabPicture(0)   =   "frm160014.frx":0004
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(39)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(38)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(37)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtFinger(3)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtFinger(2)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtFinger(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdFinger(5)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdFinger(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdFinger(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cmdFinger(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmdFinger(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtSrcCardNo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdFinger(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtSrcCardName"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdFinger(2)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).ControlCount=   17
      TabCaption(1)   =   "卡片"
      TabPicture(1)   =   "frm160014.frx":0020
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdCard(7)"
      Tab(1).Control(1)=   "cmdCard(6)"
      Tab(1).Control(2)=   "txtCardNo"
      Tab(1).Control(3)=   "cmdCard(0)"
      Tab(1).Control(4)=   "cmdCard(1)"
      Tab(1).Control(5)=   "cmdCard(2)"
      Tab(1).Control(6)=   "cmdCard(4)"
      Tab(1).Control(7)=   "cmdCard(3)"
      Tab(1).Control(8)=   "List1"
      Tab(1).Control(9)=   "txtCardMemo"
      Tab(1).Control(10)=   "Label1(35)"
      Tab(1).Control(11)=   "Label1(36)"
      Tab(1).ControlCount=   12
      Begin VB.CommandButton cmdFinger 
         Caption         =   "指紋檢查"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   6240
         TabIndex        =   47
         Top             =   420
         Width           =   888
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "卡片檢查"
         Enabled         =   0   'False
         Height          =   345
         Index           =   7
         Left            =   -69735
         TabIndex        =   39
         Top             =   990
         Width           =   1830
      End
      Begin VB.TextBox txtSrcCardName 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   35
         Top             =   1650
         Width           =   1365
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "卡片回寫考勤機"
         Enabled         =   0   'False
         Height          =   345
         Index           =   6
         Left            =   -69735
         TabIndex        =   34
         Top             =   540
         Width           =   1830
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "指紋回寫考勤機"
         Enabled         =   0   'False
         Height          =   372
         Index           =   6
         Left            =   4680
         TabIndex        =   32
         Top             =   406
         Width           =   1488
      End
      Begin VB.TextBox txtSrcCardNo 
         Alignment       =   2  '置中對齊
         Enabled         =   0   'False
         Height          =   270
         Left            =   5670
         MaxLength       =   5
         TabIndex        =   30
         Top             =   900
         Width           =   1455
      End
      Begin VB.TextBox txtCardNo 
         Height          =   315
         Left            =   -74190
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   1050
         Width           =   1635
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   -70875
         TabIndex        =   16
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "確定"
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   -71700
         TabIndex        =   15
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "新增"
         Enabled         =   0   'False
         Height          =   345
         Index           =   2
         Left            =   -74235
         TabIndex        =   12
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "刪除"
         Enabled         =   0   'False
         Height          =   345
         Index           =   4
         Left            =   -72540
         TabIndex        =   14
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdCard 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   -73395
         TabIndex        =   13
         Top             =   540
         Width           =   795
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "修改"
         Enabled         =   0   'False
         Height          =   345
         Index           =   3
         Left            =   855
         TabIndex        =   4
         Top             =   420
         Width           =   795
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "刪除"
         Enabled         =   0   'False
         Height          =   345
         Index           =   4
         Left            =   1710
         TabIndex        =   5
         Top             =   420
         Width           =   795
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "確定"
         Enabled         =   0   'False
         Height          =   345
         Index           =   1
         Left            =   2550
         TabIndex        =   6
         Top             =   420
         Width           =   795
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "取消"
         Enabled         =   0   'False
         Height          =   345
         Index           =   0
         Left            =   3375
         TabIndex        =   7
         Top             =   420
         Width           =   795
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   948
         ItemData        =   "frm160014.frx":003C
         Left            =   -74190
         List            =   "frm160014.frx":003E
         TabIndex        =   18
         Top             =   1350
         Width           =   1635
      End
      Begin VB.TextBox txtCardMemo 
         Height          =   1245
         Left            =   -71805
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   19
         Top             =   1050
         Width           =   1680
      End
      Begin VB.CommandButton cmdFinger 
         Caption         =   "從考勤機讀取指紋"
         Enabled         =   0   'False
         Height          =   345
         Index           =   5
         Left            =   5175
         TabIndex        =   11
         Top             =   1260
         Width           =   1965
      End
      Begin VB.TextBox txtFinger 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   1
         Left            =   855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   8
         Top             =   930
         Width           =   4155
      End
      Begin VB.TextBox txtFinger 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Index           =   2
         Left            =   855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   9
         Top             =   2130
         Width           =   4200
      End
      Begin VB.TextBox txtFinger 
         Height          =   924
         Index           =   3
         Left            =   5112
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   10
         Top             =   2400
         Width           =   2064
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "名稱"
         Height          =   180
         Left            =   5175
         TabIndex        =   36
         Top             =   1695
         Width           =   360
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "卡號"
         Height          =   180
         Left            =   5175
         TabIndex        =   31
         Top             =   945
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "卡號："
         Height          =   180
         Index           =   35
         Left            =   -74775
         TabIndex        =   26
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說明："
         Height          =   180
         Index           =   36
         Left            =   -72390
         TabIndex        =   25
         Top             =   1050
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指紋1："
         Height          =   180
         Index           =   37
         Left            =   225
         TabIndex        =   24
         Top             =   930
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "指紋2："
         Height          =   180
         Index           =   38
         Left            =   225
         TabIndex        =   23
         Top             =   2130
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "備註："
         Height          =   180
         Index           =   39
         Left            =   5160
         TabIndex        =   22
         Top             =   2136
         Width           =   540
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7560
      Top             =   300
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
            Picture         =   "frm160014.frx":0040
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":035C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":0678
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":0854
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":0B70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":0E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":11A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":14C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":17E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":1AFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm160014.frx":1E18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar TBar1 
      Height          =   576
      Left            =   -96
      TabIndex        =   28
      Top             =   24
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
            Enabled         =   0   'False
            Object.Visible         =   0   'False
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
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   216
      Left            =   4992
      TabIndex        =   43
      Top             =   5376
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   381
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   216
      Left            =   4992
      TabIndex        =   44
      Top             =   5112
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   381
      _Version        =   393216
      Appearance      =   1
      Max             =   20
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "在/離職              ( 1:在職 2:離職 )"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   49
      Top             =   1416
      Width           =   2508
   End
   Begin VB.Label lblNote 
      AutoSize        =   -1  'True
      Caption         =   "指紋尚未建檔!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   216
      Left            =   5856
      TabIndex        =   46
      Top             =   720
      Width           =   1452
   End
   Begin MSForms.TextBox txtST02 
      Height          =   285
      Left            =   2670
      TabIndex        =   1
      Top             =   690
      Width           =   1260
      VariousPropertyBits=   679495707
      MaxLength       =   12
      Size            =   "2222;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtDispName 
      Height          =   285
      Left            =   2670
      TabIndex        =   2
      Top             =   1032
      Width           =   1260
      VariousPropertyBits=   679495707
      MaxLength       =   12
      Size            =   "2222;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "顯示名稱"
      Height          =   180
      Left            =   1890
      TabIndex        =   33
      Top             =   1068
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "所別"
      Height          =   180
      Left            =   585
      TabIndex        =   29
      Top             =   1068
      Width           =   360
   End
   Begin VB.Label lblST02 
      AutoSize        =   -1  'True
      Caption         =   "中文名稱"
      Height          =   180
      Left            =   1890
      TabIndex        =   27
      Top             =   750
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號"
      Height          =   180
      Index           =   0
      Left            =   225
      TabIndex        =   21
      Top             =   720
      Width           =   720
   End
End
Attribute VB_Name = "frm160014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/13 Form2.0已修改
'Created by Morgan 2013/7/16
Option Explicit

Dim m_EditMode As Integer '1:新增 2:修改 3:刪除 4:查詢
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_bActived As Boolean
Dim arrCardMemo() As String
Dim m_stDomain As String
Dim arrHTAIP() As String
Dim arrIP() As String

Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      Check2.Value = vbUnchecked
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      Check1.Value = vbUnchecked
   End If
End Sub

Private Sub cmdCard_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   OnAction3 Index
   Screen.MousePointer = vbDefault
End Sub

Private Sub OnAction3(pIndex As Integer)
   Select Case pIndex
   Case 0 '取消
      If m_EditMode = 1 Then
         txtCardNo.Text = ""
         txtCardMemo.Text = ""
      End If
      m_EditMode = 0
   
   Case 1 '確定
      If m_EditMode = 1 Then
         If MsgBox("卡片將回寫『 " & txtST06 & "所 』所有考勤機，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問") = vbYes Then 'Added by Morgan 2019/1/15
            If AddCard() = False Then
               Exit Sub
            Else
               ReDim Preserve arrCardMemo(UBound(arrCardMemo) + 1) As String
               
               arrCardMemo(UBound(arrCardMemo)) = txtCardMemo
               List1.AddItem txtCardNo, 0
               List1.ItemData(0) = UBound(arrCardMemo)
            End If
         End If
      Else
         If UpdateCard() = False Then
            Exit Sub
         Else
            arrCardMemo(List1.ItemData(List1.ListIndex)) = txtCardMemo
         End If
      End If
      
      txtCardNo = ""
      txtCardMemo = ""
      List1.ListIndex = -1
      m_EditMode = 0
      
   Case 2 '新增
      m_EditMode = 1
      
   Case 3 '修改
      m_EditMode = 2
      
   Case 4 '刪除
      If MsgBox("是否要刪除卡片[ " & txtCardNo & " ]?", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
         If DeleteCard() = False Then
            Exit Sub
         Else
            MsgBox "卡片[ " & txtCardNo & " ]已刪除！", vbInformation 'Added by Morgan 2024/6/20
            
            txtCardNo = ""
            txtCardMemo = ""
            arrCardMemo(List1.ItemData(List1.ListIndex)) = ""
            List1.RemoveItem List1.ListIndex
            ShowRecord 'Added by Morgan 2024/3/14 刪卡片可能會自動新增空指紋所以要重讀指紋資料
         End If
      Else
         Exit Sub
      End If
      
   Case 6
      'Modified by Morgan 2020/5/14
      intI = 0
      If cboHtaIp <> "" Then
         intI = MsgBox("卡片將回寫至考勤機『 " & cboHtaIp & "』，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
      Else
         intI = MsgBox("卡片將回寫『 " & txtST06 & "所 』所有考勤機，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
      End If
      If intI = vbYes Then
      'end 2020/5/14
         Call UpdateCardNo
      End If
      
   Case 7
      If cboHtaIp <> "" And txtCardNo <> "" Then
         'Modified by Morgan 2024/3/5
         'HTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
         arrIP = Split(cboHtaIp, " ")
         HTAip = arrIP(0)
         'end 2023/3/5
         strExc(0) = ""
         If HTAqueryCard(txtCardNo, strExc(0), True, , intI) = True Then
            MsgBox "卡片 " & txtCardNo & "(" & strExc(0) & ") 檢查成功！", vbInformation
         Else
            If intI = 6 Then
               MsgBox "該考勤機無此卡片", vbExclamation
            Else
               MsgBox "卡片檢查成失敗", vbCritical
            End If
         End If
      'Modified by Morgan 2023/12/4
      'Else
      '   MsgBox "請先點選考勤機！", vbExclamation
      Else
         strExc(0) = "將對"
         If cboHtaIp = "" Then
            strExc(0) = strExc(0) & "『 " & txtST06 & "所 』所有考勤機"
         Else
            strExc(0) = strExc(0) & "考勤機(" & cboHtaIp & ")"
         End If
         strExc(0) = strExc(0) & "檢查"
         If txtCardNo = "" Then
            strExc(0) = strExc(0) & "所有卡片"
         End If
         strExc(0) = strExc(0) & "，是否確定要繼續？"
         
         If MsgBox(strExc(0), vbYesNo + vbExclamation + vbDefaultButton2, "詢問") = vbYes Then
            Call ChkCardNo
         End If
      'end 2023/12/4
      End If
      
   End Select
   
   UpdateToolbarState1
End Sub

Private Sub cmdChgTimeZone_Click()
   
   If cboHtaIp <> "" Then
      strExc(0) = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
      If CheckLevel(strExc(0), "門禁機IP") = False Then
         MsgBox "請選擇門禁機回寫！", vbExclamation
         intI = vbNo
      Else
         intI = MsgBox("時段將回寫至門禁機『 " & cboHtaIp & "』，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
      End If
   Else
      intI = MsgBox("時段將回寫『 " & txtST06 & "所 』所有門禁機，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
   End If
   If intI = vbYes Then
      If UpdateTimeZone() = True Then
         If UpdateST73() = True Then
            MsgBox "時段回寫成功！", vbInformation
         End If
      Else
         Exit Sub
      End If
   End If
End Sub

Private Sub cmdFinger_Click(Index As Integer)
   Screen.MousePointer = vbHourglass
   OnAction2 Index
   UpdateToolbarState
   Screen.MousePointer = vbDefault
End Sub

Private Sub OnAction2(pIndex As Integer)
   Select Case pIndex
   Case 0 '取消
      txtFinger(1) = txtFinger(1).Tag
      txtFinger(2) = txtFinger(2).Tag
      txtFinger(3) = txtFinger(3).Tag
      m_EditMode = 0
      
   Case 1 '確定
      If txtFinger(1) <> txtFinger(1).Tag Or txtFinger(2) <> txtFinger(2).Tag Or txtFinger(3) <> txtFinger(3).Tag Then
         If UpdateFinger() = True Then
            '指紋有變動更新考勤機
            If txtFinger(1) <> txtFinger(1).Tag Or txtFinger(2) <> txtFinger(2).Tag Then
               Call UpdateMachin
            Else
               MsgBox "備註已更新！", vbInformation
            End If
            m_EditMode = 0
         Else
            Exit Sub
         End If
      Else
         m_EditMode = 0
      End If
      
   'Added by Morgan 2023/12/4
   Case 2 '指紋檢查
      If cboHtaIp = "" Then
         If MsgBox("將檢查『 " & txtST06 & "所 』所有考勤機，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問") = vbNo Then
            Exit Sub
         End If
      End If
      Call ChkFinger
   'end 2023/12/4
   Case 3 '修改
      txtFinger(1).Tag = txtFinger(1)
      txtFinger(2).Tag = txtFinger(2)
      txtFinger(3).Tag = txtFinger(3)
      txtSrcCardNo = getSrcCardNo(txtST01)
      If InStr(cboHtaIp, cboHtaIp.Tag) <> 1 Then 'Added by Morgan 2024/3/5
         SetDefaultHtaIP 'Added by Morgan 2015/12/1
      End If
      m_EditMode = 2
   Case 4 '刪除
      If MsgBox("是否要刪除指紋資料?", vbYesNo + vbExclamation + vbDefaultButton2, "詢問") = vbYes Then
         If DeleteFinger() = False Then
            Exit Sub
         Else
            MsgBox "指紋已刪除！", vbInformation
            txtFinger(1) = ""
            txtFinger(2) = ""
            txtFinger(3) = ""
         End If
      Else
         Exit Sub
      End If
      
   Case 5 '從考勤機讀取指紋
      If GetFinger = True Then
         MsgBox "指紋讀取成功！", vbInformation
         'Mofieid by Morgan 2015/12/1
         'cboHtaIp.Tag = cboHtaIp
         'Modified by Morgan 2024/3/5
         'cboHtaIp.Tag = arrHTAIP(cboHtaIp.ListIndex)
         cboHtaIp.Tag = HTAip
         'end 2023/3/5
         'end 2015/12/1
         txtSrcCardNo.Tag = txtSrcCardNo
         'Added by Morgan 2024/3/5
         lblNote.Visible = False
         UpdateToolbarState2
         If cmdFinger(3).Enabled Then cmdFinger(3).Value = True
         'end 2024/3/5
      Else
         txtSrcCardNo.Tag = ""
         cboHtaIp.Tag = ""
         Exit Sub
      End If
      
   Case 6 '指紋回寫考勤機
      'Modified by Morgan 2019/12/9
      intI = 0
      If cboHtaIp <> "" Then
         intI = MsgBox("指紋將回寫至考勤機『 " & cboHtaIp & "』，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
      Else
         intI = MsgBox("指紋將回寫『 " & txtST06 & "所 』所有考勤機，是否確定要繼續？", vbYesNo + vbExclamation + vbDefaultButton2, "詢問")
      End If
      If intI = vbYes Then
      'end 2019/12/9
         Call UpdateMachin
      End If
   End Select
   UpdateToolbarState2
   
End Sub

Private Function getSrcCardNo(pNo As String) As String
   Dim stChar As String, stRtn As String
   
   stRtn = pNo
   If pNo <> "" Then
      stChar = Left(pNo, 1)
      If stChar >= "A" And stChar <= "Z" Then
         'Modified by Morgan 2018/5/16 改固定抓後4碼 -- 經理
         'stRtn = (Asc(stChar) - 64) & Mid(pNo, 2)
         stRtn = Mid(pNo, 2)
         'end 2018/5/16
      End If
   End If
   getSrcCardNo = stRtn
End Function
Private Function GetFinger() As Boolean
   Dim stFinger1 As String, stFinger2 As String
   
   txtSrcCardName = ""
   If cboHtaIp = "" Then
      MsgBox "請選取考勤機IP！", vbExclamation
      cboHtaIp.SetFocus
   Else
      'Modified by Morgan 2015/12/1
      'HTAip = cboHtaIp.Text
      'Modified by Morgan 2024/3/5
      'HTAip = arrHTAIP(cboHtaIp.ListIndex)
      arrIP = Split(cboHtaIp, " ")
      HTAip = arrIP(0)
      'end 2024/3/5
      'end 2015/12/1
      
      'Added by Morgan 2013/8/1
      '讀指紋指定在執行檔跑時若連線後馬上呼叫會錯,所有故意連線後斷線
      If HTAconnect() = True Then
         HTAclose
      'end 2013/8/1
         If HTAqueryFingerPrinter(IIf(txtSrcCardNo = "", txtST01, txtSrcCardNo), stFinger1, stFinger2) = True Then
            txtFinger(1) = stFinger1
            txtFinger(2) = stFinger2
            GetFinger = True
            'Added by Morgan 2013/10/23
            strExc(1) = ""
            If HTAqueryCard(IIf(txtSrcCardNo = "", txtST01, txtSrcCardNo), strExc(1)) = True Then
               txtSrcCardName = strExc(1)
            End If
            'end 2013/10/23
         End If
      End If
   End If
End Function

'Added by Morgan 2024/5/23
Private Sub Command1_Click()
   Dim bolTrans As Boolean
   Dim stMsg As String
   
   If txtST04 = "1" Then
      If MsgBox(stMsg & "【" & txtST01 & " " & txtST02 & "】目前仍在職，是否確定要繼續？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
      stMsg = "目前連線並非正式資料庫，"
   End If
   
   If MsgBox(stMsg & "是否確定要刪除所有考勤機/門禁機上" & vbCrLf & "【" & txtST01 & " " & txtST02 & "】的指紋及卡片資料？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
      Exit Sub
   End If
   
On Error GoTo ErrHand

   Screen.MousePointer = vbHourglass
   cnnConnection.BeginTrans
   bolTrans = True
   'Modified by Morgan 2024/9/9 +傳入 List2
   If PUB_ClearCardData(txtST01.Text, bolTrans, List2) = True Then
      cnnConnection.CommitTrans
      bolTrans = False
      MsgBox "作業完成！", vbInformation
      OnAction vbKeyF10
   End If
   
ErrHand:
   If bolTrans Then
      cnnConnection.RollbackTrans
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
   If m_bActived = False Then
      m_bActived = True
      SSTab1.Tab = 0
   End If
End Sub

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
         
      Case vbKeyEscape:
         If TypeName(Me.ActiveControl) <> "ComboBox" Then
            If m_EditMode <> 0 Then
               OnAction vbKeyF10
            Else
               OnAction KeyCode
            End If
         End If
         
      Case vbKeyReturn
         '做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
         KeyCode = 0
         If m_EditMode <> 0 Then
            OnAction vbKeyF9
         End If

   End Select
End Sub

Private Sub Form_Load()
   Dim arrIP() As String
   
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)

   MoveFormToCenter Me
   
   '讀取考勤機IP
   cboHtaIp.Clear
   cboHtaIp.AddItem "", 0 '第1筆空白表示不指定IP
   HTAips = GetHtaIP()
   If HTAips <> "" Then
      arrIP = Split(HTAips, ";")
      
      For intI = LBound(arrIP) To UBound(arrIP)
         If arrIP(intI) <> "" Then
            cboHtaIp.AddItem arrIP(intI)
         End If
      Next
      
      'Added by Morgan 2015/12/1
      'IP後面加說明
      If cboHtaIp.ListCount > 0 Then
         ReDim arrHTAIP(cboHtaIp.ListCount - 1) As String
         For intI = 0 To cboHtaIp.ListCount - 1
            arrHTAIP(intI) = cboHtaIp.List(intI)
            If arrHTAIP(intI) <> "" Then
               'Modified by Morgan 2020/1/7 改抓特殊設定
               'cboHtaIp.List(intI) = cboHtaIp.List(intI) & " " & GetIpName(cboHtaIp.List(intI))
               cboHtaIp.List(intI) = cboHtaIp.List(intI) & " " & Pub_GetSpecMan(cboHtaIp.List(intI))
            End If
         Next
      End If
      'end 2015/12/1
      
      'cboHtaIp.ListIndex = 0
   End If
   
   If ShowRecord(-2) = True Then
      m_EditMode = 0
   Else
      Form_KeyDown vbKeyF2, 0
   End If
   UpdateToolbarState
   
   'Added by Morgan 2020/8/27
   If m_bUpdate Then
      Check1.Enabled = True
      Check2.Enabled = True 'Added by Morgan 2023/10/6
      cmdChgTimeZone.Enabled = True
   Else
      Check1.Enabled = False
      Check2.Enabled = False 'Added by Morgan 2023/10/6
      cmdChgTimeZone.Enabled = False
   End If
   'end 2020/8/27
End Sub

'Added by Morgan 2015/12/1
'考勤機說明
'Private Function GetIpName(pIP As String) As String
'   Select Case pIP
'   Case "192.168.0.1": GetIpName = "北所-牆"
'   Case "192.168.0.2": GetIpName = "北所-柱"
'   Case "192.168.2.1": GetIpName = "中所"
'   Case "192.168.3.1": GetIpName = "南所"
'   Case "192.168.4.1": GetIpName = "高所"
'   End Select
'End Function

'Added by Morgan 2015/12/1
'預設考勤機
Private Sub SetDefaultHtaIP()
   Dim ii As Integer
   For ii = 0 To cboHtaIp.ListCount - 1
      If InStr(cboHtaIp.List(ii), m_stDomain) = 1 Then
         cboHtaIp.ListIndex = ii
         Exit For
      End If
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160014 = Nothing
End Sub

Private Sub List1_Click()
   
   If List1 <> "" Then
      txtCardNo = List1
      txtCardMemo = arrCardMemo(List1.ItemData(List1.ListIndex))
      If m_bUpdate = True Then
         cmdCard(3).Enabled = True
         cmdCard(6).Enabled = True
      End If
      cmdCard(7).Enabled = True 'Added by Morgan 2020/5/14
      
      If m_bDelete = True Then
         cmdCard(4).Enabled = True
      End If
   End If
End Sub


Private Sub Tbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
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

Private Sub UpdateToolbarState1()
   Dim oCmd As CommandButton
   
   For Each oCmd In cmdCard
      oCmd.Enabled = False
   Next
      
   List1.Enabled = False
   txtCardNo.Enabled = False
   txtCardMemo.Enabled = False
   
   Select Case m_EditMode
      Case 0
         If m_bInsert Then
            cmdCard(2).Enabled = True
         End If
         'Added by Morgan 2020/5/14
         If txtCardNo <> "" Then
            If m_bUpdate = True Then
               cmdCard(3).Enabled = True
               cmdCard(6).Enabled = True
            End If
            cmdCard(7).Enabled = True
            
            If m_bDelete = True Then
               cmdCard(4).Enabled = True
            End If
         'Added by Morgan 2023/12/4
         ElseIf List1.ListCount > 0 Then
            cmdCard(6).Enabled = True
            cmdCard(7).Enabled = True
         'end 2023/12/4
         End If
         'end 2020/514
         
         List1.Enabled = True
         TBar1.Enabled = True
         SSTab1.TabEnabled(0) = True
         
      Case 1
         SSTab1.TabEnabled(0) = False
         cmdCard(0).Enabled = True
         cmdCard(1).Enabled = True
         cmdCard(2).Enabled = False
         cmdCard(3).Enabled = False
         cmdCard(4).Enabled = False
         cmdCard(6).Enabled = False
         cmdCard(7).Enabled = False 'Added by Morgan 2020/5/14
         
         txtCardNo.Text = ""
         txtCardNo.Enabled = True
         txtCardMemo.Text = ""
         txtCardMemo.Enabled = True
         List1.ListIndex = -1
         txtCardNo.SetFocus
         TBar1.Enabled = False
         
      Case 2
         SSTab1.TabEnabled(0) = False
         cmdCard(0).Enabled = True
         cmdCard(1).Enabled = True
         cmdCard(2).Enabled = False
         cmdCard(3).Enabled = False
         cmdCard(4).Enabled = False
         cmdCard(6).Enabled = False
         cmdCard(7).Enabled = False 'Added by Morgan 2020/5/14
         txtCardMemo.Enabled = True
         txtCardMemo.SetFocus
         TBar1.Enabled = False
   End Select
   
   
End Sub

Private Sub UpdateToolbarState2()
   Dim oCmd As CommandButton
   
   For Each oCmd In cmdFinger
      oCmd.Enabled = False
   Next
   'txtFinger(1).Enabled = False
   'txtFinger(2).Enabled = False
   txtFinger(3).Enabled = False
   
   Select Case m_EditMode
      Case 0
         If m_bUpdate Then
            If lblNote.Visible = False Then
               cmdFinger(3).Caption = "修改"
               cmdFinger(3).Enabled = True
               cmdFinger(6).Enabled = True
               cmdFinger(2).Enabled = True
            Else
               cmdFinger(3).Caption = "新增"
               cmdFinger(3).Enabled = True
            End If
         End If
         If m_bDelete Then
            If lblNote.Visible = False Then
               cmdFinger(4).Enabled = True
            End If
         End If
         TBar1.Enabled = True
         SSTab1.TabEnabled(1) = True
         'Removed by Morgan 2023/12/26 不可清除，更換指紋是需要用來刪除暫存的新指紋
         'txtSrcCardNo = ""
         'txtSrcCardNo.Tag = ""
         'end 2023/12/26
         'Modified by Morgan 2023/7/5 查詢也開放讀取以便確認是否已建立指紋
         'txtSrcCardNo.Enabled = False
         txtSrcCardNo.Enabled = True
         cmdFinger(5).Enabled = True
         'end 2023/7/5
         txtSrcCardName = ""
         'cboHtaIp.ListIndex = -1
         'cboHtaIp.Tag = ""
         'cboHtaIp.Enabled = False 'Removed by Morgan 2019/12/9 開放可寫他所考勤機，因可能會在不同所別打卡
         
      Case 2
         TBar1.Enabled = False
         SSTab1.TabEnabled(1) = False
         cmdFinger(0).Enabled = True
         cmdFinger(1).Enabled = True
         cmdFinger(3).Enabled = False
         cmdFinger(4).Enabled = False
         cmdFinger(5).Enabled = True
         cmdFinger(6).Enabled = False
         cboHtaIp.Enabled = True
         txtFinger(3).Enabled = True
         txtFinger(3).SetFocus
         txtSrcCardNo.Enabled = True
   End Select
   
   
End Sub

'依照權限設定其工具列的按紐狀態
Private Sub UpdateToolbarState()
   Command1.Enabled = False 'Added by Morgan 2024/5/23
   Select Case m_EditMode
      Case 0 ' 無任何動作
         If m_bQuery Then
            TBar1.Buttons(4).Enabled = True
            If txtST01 <> "" Then Command1.Enabled = True  'Added by Morgan 2024/5/23
         Else
            TBar1.Buttons(4).Enabled = False
         End If
         If m_bQuery And txtST01 <> "" Then
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
         txtST01.Enabled = False
         txtST02.Enabled = False
         
      Case 4 '查詢
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = True
         TBar1.Buttons(12).Enabled = True
         TBar1.Buttons(14).Enabled = False
         txtST01.Enabled = True
         txtST02.Enabled = True
      'Added by Morgan 2013/10/23
      Case Else
         TBar1.Buttons(4).Enabled = False
         TBar1.Buttons(6).Enabled = False
         TBar1.Buttons(7).Enabled = False
         TBar1.Buttons(8).Enabled = False
         TBar1.Buttons(9).Enabled = False
         TBar1.Buttons(11).Enabled = False
         TBar1.Buttons(12).Enabled = False
         TBar1.Buttons(14).Enabled = False
         txtST01.Enabled = False
         txtST02.Enabled = False
   End Select
   UpdateToolbarState1
   UpdateToolbarState2
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim bCancel As Boolean
   
   Select Case KeyCode
      Case vbKeyF2 ' 新增
      Case vbKeyF3 ' 修改
      Case vbKeyF5 ' 刪除
      Case vbKeyF4 ' 查詢
         SSTab1.Tab = 0
         m_EditMode = 4
         ClearField
         UpdateToolbarState
         txtST01.SetFocus
         
      Case vbKeyHome ' 第一筆
         ShowRecord -2
      Case vbKeyPageUp ' 前一筆
         ShowRecord -1
      Case vbKeyPageDown ' 後一筆
         ShowRecord 1
      Case vbKeyEnd ' 最後一筆
         ShowRecord 2
      Case vbKeyF9 ' 確定
         If OnWork = True Then
            UpdateToolbarState
         Else
            Exit Sub
         End If
         
      Case vbKeyF10 ' 取消
         txtST01 = txtST01.Tag
         m_EditMode = 0
         ShowRecord
         UpdateToolbarState
         
      Case vbKeyEscape ' 離開
         Unload Me
   End Select
End Sub

Private Sub ClearField()
   m_stDomain = ""
   txtST01 = ""
   txtST02 = ""
   txtST06 = ""
   txtST04 = "" 'Added by Morgan 2024/5/27
   txtDispName = ""
   txtFinger(1) = ""
   txtFinger(2) = ""
   txtFinger(3) = ""
   List1.Clear
   txtCardNo = ""
   txtCardMemo = ""
   lblNote.Visible = True
   Check1.Value = vbUnchecked 'Added by Morgan 2020/8/28
   Check2.Value = vbUnchecked 'Added by Morgan 2023/10/6
   Erase arrCardMemo
End Sub

'p_iWay:0=尋找,-2=首筆,-1=前筆,+1=後筆,2=末筆
Private Function ShowRecord(Optional ByVal p_iWay As Integer = 0) As Boolean
   Dim iItemData As Integer
   Dim adoRst As New ADODB.Recordset
   
   'Modified by Morgan 2017/4/26 有用99996來賓,取消特殊編號限制(and substr(st01,-2)<'9')
   'Modified by Morgan 2017/6/8 +台一投資(R04),門禁要用
   'Modified by Morgan 2024/5/27 +ST04
   Select Case p_iWay
      Case 0
         strExc(0) = "SELECT ST01,ST02,decode(ST06,'1','北','2','中','3','南','4','高','其他') ST06C,ST06,ST60,ST73,ST04 FROM STAFF" & _
            " WHERE st01>'6' and st01<'F'   and substr(st01,-2)>'00'" & _
            " AND " & IIf(txtST01 <> "", "ST01 = '" & txtST01 & "'", "ST02 like '%" & txtST02 & "%'") & _
            " and (st04='1' or exists(select * from staffcarddata where scd01=st01))"
            
      Case -2
         strExc(0) = "SELECT ST01,ST02,decode(ST06,'1','北','2','中','3','南','4','高','其他') ST06C,ST06,ST60,ST73,ST04 FROM STAFF" & _
            " WHERE st01>'6' and st01<'F'   and substr(st01,-2)>'00'" & _
            " and (st04='1' or exists(select * from staffcarddata where scd01=st01)) order by 1 ASC"
            
      Case -1
         strExc(0) = "SELECT ST01,ST02,decode(ST06,'1','北','2','中','3','南','4','高','其他') ST06C,ST06,ST60,ST73,ST04 FROM STAFF" & _
            " WHERE st01>'6' and st01<'F'  and substr(st01,-2)>'00'" & _
            " AND ST01<'" & txtST01 & "'" & _
            " and (st04='1' or exists(select * from staffcarddata where scd01=st01)) order by 1 DESC"
            
      Case 1
         strExc(0) = "SELECT ST01,ST02,decode(ST06,'1','北','2','中','3','南','4','高','其他') ST06C,ST06,ST60,ST73,ST04 FROM STAFF" & _
            " WHERE st01>'6' and st01<'F'  and substr(st01,-2)>'00'" & _
            " AND ST01>'" & txtST01 & "'" & _
            " and (st04='1' or exists(select * from staffcarddata where scd01=st01)) order by 1 ASC"
            
      Case 2
         strExc(0) = "SELECT ST01,ST02,decode(ST06,'1','北','2','中','3','南','4','高','其他') ST06C,ST06,ST60,ST73,ST04 FROM STAFF" & _
            " WHERE st01>'6' and st01<'F'  and substr(st01,-2)>'00'" & _
            " and (st04='1' or exists(select * from staffcarddata where scd01=st01)) order by 1 DESC"
            
   End Select
   
   adoRst.MaxRecords = 1
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      ClearField
      txtST01 = adoRst("st01")
      txtST01.Tag = txtST01
      txtST02 = "" & adoRst("st02")
      txtST06 = "" & adoRst("ST06C")
      txtST04 = "" & adoRst("st04") 'Added by Morgan 2024/5/27
      
      'Added by Morgan 2020/8/27
      If adoRst("st73") = "Y" Then
         Check1.Value = vbChecked
         Check2.Value = vbUnchecked
      'Added by Morgan 2023/10/6
      ElseIf adoRst("st73") = "S" Then
         Check1.Value = vbUnchecked
         Check2.Value = vbChecked
      Else
         Check1.Value = vbUnchecked
         Check2.Value = vbUnchecked
      End If
      'end 2020/8/27
      '北
      If adoRst("ST06") = "1" Then
         m_stDomain = "192.168.0."
      '中
      ElseIf adoRst("ST06") = "2" Then
         m_stDomain = "192.168.2."
      '南
      ElseIf adoRst("ST06") = "3" Then
         m_stDomain = "192.168.3."
      '高
      ElseIf adoRst("ST06") = "4" Then
         m_stDomain = "192.168.4."
      End If
      
      txtDispName = "" & adoRst("ST60")
      
      strExc(0) = "select * from staffcarddata where scd01='" & txtST01 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         iItemData = 0
         ReDim arrCardMemo(RsTemp.RecordCount) As String
         Do While Not RsTemp.EOF
            If RsTemp("scd01") = RsTemp("scd02") Then
               txtFinger(1) = "" & RsTemp("scd03")
               txtFinger(2) = "" & RsTemp("scd04")
               txtFinger(3) = "" & RsTemp("scd05")
               lblNote.Visible = False
            Else
               iItemData = iItemData + 1
               ReDim Preserve arrCardMemo(iItemData) As String
               
               List1.AddItem RsTemp("scd02"), 0
               List1.ItemData(0) = iItemData
               arrCardMemo(iItemData) = "" & RsTemp("scd05")
            End If
            RsTemp.MoveNext
         Loop
      Else
         ReDim arrCardMemo(0) As String
      End If
      ShowRecord = True
   Else
      If p_iWay = -1 Then
         MsgBox "已經是第一筆！", vbInformation
      ElseIf p_iWay = 1 Then
         MsgBox "已經是最後筆！", vbInformation
      Else
         MsgBox "查無資料！", vbInformation
         ClearField
      End If
   End If
      
   UpdateToolbarState
   Set adoRst = Nothing
End Function

Private Function OnWork() As Boolean
   Select Case m_EditMode
      Case 1: '新增
      Case 2: '修改
      Case 3: '刪除
      Case 4: '查詢
         If txtST01 = "" And txtST02 = "" Then
            MsgBox "請輸入員工編號或中文名稱!!", vbInformation
         ElseIf ShowRecord = True Then
            OnWork = True
            m_EditMode = 0
         End If
   End Select
End Function

Private Function UpdateST73() As Boolean
On Error GoTo ErrHnd
   'Modified by Morgan 2023/10/6 +S:Check2
   strSql = "update staff set st73='" & IIf(Check1.Value = vbChecked, "Y", IIf(Check2.Value = vbChecked, "S", "")) & "' where st01='" & txtST01 & "'"
   cnnConnection.Execute strSql, intI
   UpdateST73 = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function UpdateCard() As Boolean
On Error GoTo ErrHnd
   strSql = "update staffcarddata set scd05='" & ChgSQL(txtCardMemo) & "' where scd02='" & txtCardNo & "'"
   cnnConnection.Execute strSql, intI
   UpdateCard = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Function AddCard() As Boolean
   Dim iRCode As Integer
   Dim ii As Integer
   
On Error GoTo ErrHnd
   
   If txtCardNo = "" Then
      MsgBox "請輸入卡號！", vbInformation
      txtCardNo.SetFocus
      Exit Function
   End If

   strExc(0) = "select scd01,st02 from staffcarddata,staff where scd02='" & txtCardNo & "' and st01(+)=scd01"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("scd01") = txtST01 Then
         MsgBox "卡號重複，請重新輸入!!"
      Else
         MsgBox "卡號已存在(卡片持有人：" & RsTemp("st02") & ")，請重新輸入!!"
         Exit Function
      End If
   ElseIf intI = 0 Then
   
      '新增考勤機卡號資料
      For ii = 0 To cboHtaIp.ListCount - 1
         'Modified by Morgan 2015/12/1
         'HTAip = cboHtaIp.List(ii)
         HTAip = arrHTAIP(ii)
         'end 2015/12/1
         If m_stDomain <> "" And InStr(HTAip, m_stDomain) = 1 Then 'Added by Morgan 2019/1/15
            If HTAqueryCard(txtCardNo, , True, , iRCode) = True Then
               If HTAdeleteCard(txtCardNo) = False Then
                  MsgBox "考勤機(" & HTAip & ")卡號( " & txtCardNo & " )刪除失敗，作業取消!!"
                  Exit Function
               End If
            ElseIf iRCode <> 6 Then
               MsgBox "考勤機(" & HTAip & ")卡號( " & txtCardNo & " )檢查失敗，作業取消!!"
               Exit Function
            End If
            
            'Modified by Morgan 2023/10/6
            'If HTAaddCard(txtCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , , IIf(Check1.Value = vbChecked, True, False)) = False Then
            If HTAaddCard(txtCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1))) = False Then
               MsgBox "考勤機(" & HTAip & ")卡號( " & txtCardNo & " )新增失敗，作業取消!!"
               Exit Function
            End If
         End If
      Next
      '新增資料庫卡號資料
      strSql = "insert into staffcarddata(scd01,scd02,scd05) values ('" & txtST01 & "','" & txtCardNo & "','" & ChgSQL(txtCardMemo) & "')"
      cnnConnection.Execute strSql, intI
      AddCard = True
   End If
   
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdatePollRecord(pCardNo As String, pUserNo As String) As Boolean
   Dim stSQL As String, intR As Integer
   
On Error GoTo ErrHnd
   cnnConnection.BeginTrans
   
   PUB_DelRepRec pUserNo 'Added by Morgan 2024/6/20
   
   stSQL = "update pollrecord set pr03='" & pUserNo & "' where pr03='" & pCardNo & "'"
   cnnConnection.Execute stSQL, intR
   
   'Added by Morgan 2024/3/14
   '檢查是否有指紋紀錄，若沒有時自動新增一筆空的以便查詢打卡紀錄
   If intR > 0 Then
      stSQL = "update staffcarddata set scd03=scd03 where scd01='" & pUserNo & "' and scd02='" & pUserNo & "'"
      cnnConnection.Execute stSQL, intR
      If intR = 0 Then
         stSQL = "insert into staffcarddata(scd01,scd02,scd05) values ('" & pUserNo & "','" & pUserNo & "','空指紋.查詢記錄用.勿刪')"
         cnnConnection.Execute stSQL, intR
      End If
   End If
   'end 2024/3/14
   
   cnnConnection.CommitTrans
   UpdatePollRecord = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical, "作業失敗！"
   
End Function

Private Function DeleteCard() As Boolean
   Dim ii As Integer, iRCode As Integer
   Dim idx As Integer 'Added by Morgan 2018/9/4
   
On Error GoTo ErrHnd

   strExc(0) = "select * from pollrecord where pr03='" & txtCardNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2020/4/27 改保留刷卡紀錄後刪除
      'MsgBox "卡片號( " & txtCardNo & " ) 已有刷卡紀錄不可刪除", vbExclamation
      If UpdatePollRecord(txtCardNo, txtST01) = False Then
         Exit Function
      End If
      'end 2020/4/27
   End If
   
   '刪除考勤機卡號資料
   For ii = 0 To cboHtaIp.ListCount - 1
      idx = ii 'Added by Morgan 2018/9/4 HTAqueryCard(hsHTA850QueryUserRecord) 若跨網段且查無卡號時 ii 會變 0
      'Modified by Morgan 2015/12/1
      'HTAip = cboHtaIp.List(ii)
      HTAip = arrHTAIP(ii)
      'end 2015/12/1
      If m_stDomain <> "" And InStr(HTAip, m_stDomain) = 1 Then 'Added by Morgan 2019/1/15
         If HTAqueryCard(txtCardNo, , True, , iRCode) = True Then
            If HTAdeleteCard(txtCardNo) = False Then
               MsgBox "考勤機(" & HTAip & ")卡號( " & txtCardNo & " )刪除失敗，作業取消!!"
               Exit Function
            End If
         ElseIf iRCode <> 6 Then
            MsgBox "考勤機(" & HTAip & ")卡號( " & txtCardNo & " )檢查失敗，作業取消!!"
            Exit Function
         End If
      End If
      ii = idx 'Added by Morgan 2018/9/4
   Next
   '刪除資料庫卡號資料
   strSql = "delete staffcarddata where scd02='" & txtCardNo & "'"
   cnnConnection.Execute strSql, intI
   
   
   DeleteCard = True

   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub txtCardMemo_GotFocus()
   TextInverse txtCardMemo
   OpenIme
End Sub

Private Sub txtFinger_GotFocus(Index As Integer)
   If Index > 2 Then
      TextInverse txtFinger(Index)
      OpenIme
   End If
End Sub

Private Function UpdateFinger() As Boolean
   
On Error GoTo ErrHnd
   
   strSql = "update staffcarddata set scd03='" & txtFinger(1).Text & "',scd04='" & txtFinger(2).Text & "',scd05='" & ChgSQL(txtFinger(3)) & "' where scd02='" & txtST01 & "'"
   cnnConnection.Execute strSql, intI
   
   'Added by Morgan 2025/6/23
   '新增指紋
   If intI = 0 Then
      strSql = "insert into staffcarddata(scd01,scd02,scd03,scd04,scd05) values('" & txtST01 & "','" & txtST01 & "','" & txtFinger(1).Text & "','" & txtFinger(2).Text & "','" & ChgSQL(txtFinger(3)) & "')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/6/23
   
   UpdateFinger = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateCardNo() As Boolean
   Dim iRCode As Integer, bolErr As Boolean
   Dim ii As Integer, jj As Integer
   Dim strHTAip As String 'Added by Morgan 2020/5/14
   Dim strCardNo As String
   Dim iErr As Integer
   
On Error GoTo ErrHnd

   'Added by Morgan 2020/5/14 可指定考勤機回寫
   'Modified by Morgan 2024/3/5
   'If cboHtaIp.ListIndex > 0 Then
   '   strHTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
   If cboHtaIp <> "" Then
      arrIP = Split(cboHtaIp, " ")
      strHTAip = arrIP(0)
   'end 2024/3/5
      ProgressBar1.max = 1
   Else
      strHTAip = m_stDomain
      jj = 0
      For ii = 0 To cboHtaIp.ListCount - 1
         If arrHTAIP(ii) <> "" And InStr(arrHTAIP(ii), strHTAip) = 1 Then
            jj = jj + 1
         End If
      Next
      If jj > 0 Then
         ProgressBar1.max = jj
      End If
   End If
   'end 2019/5/14
   
   If txtCardNo <> "" Then
      ProgressBar2.max = 1
   Else
      ProgressBar2.max = List1.ListCount
   End If
   
   List2.Clear
   iErr = 0
   ProgressBar2.Value = 0
   For jj = 0 To List1.ListCount - 1
      If txtCardNo <> "" Then
         strCardNo = txtCardNo
      Else
         strCardNo = List1.List(jj)
      End If
      ProgressBar1.Value = 0
      '新增考勤機卡號資料
      For ii = 0 To cboHtaIp.ListCount - 1
         'Modified by Morgan 2015/12/1
         'HTAip = cboHtaIp.List(ii)
         'Modified by Morgan 2024/3/25
         'HTAip = arrHTAIP(ii)
         If cboHtaIp <> "" Then
            HTAip = strHTAip
         Else
            HTAip = arrHTAIP(ii)
         End If
         'end 2015/12/1
         'end 2024/3/5
         If strHTAip <> "" And InStr(HTAip, strHTAip) = 1 Then
            bolErr = False
            If HTAqueryCard(strCardNo, , True, , iRCode) = True Then
               List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )檢查成功!!", 0
               If HTAdeleteCard(strCardNo) = True Then
                  List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )刪除成功!!", 0
               Else
                  List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )刪除失敗!!", 0
                  bolErr = True
               End If
            ElseIf iRCode = 6 Then
               List2.AddItem "考勤機( " & HTAip & " )無卡號( " & strCardNo & " )", 0
            Else
               List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )檢查失敗!!", 0
            End If
            
            If Not bolErr Then
               If HTAaddCard(strCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1))) = True Then
                  List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )新增成功!!", 0
               Else
                  List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )新增失敗!!", 0
               End If
            End If
            
            If bolErr Then iErr = iErr + 1
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            If cboHtaIp <> "" Then Exit For
         End If
      Next
      ProgressBar2.Value = ProgressBar2.Value + 1
      If txtCardNo <> "" Then Exit For
   Next
   
   UpdateCardNo = True
   MsgBox "卡片回寫完成！" & IIf(iErr > 0, "但有 " & iErr & " 筆失敗！", ""), IIf(iErr > 0, vbExclamation, vbInformation)
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Function UpdateMachin() As Boolean
   Dim iRCode As Integer, bolErr As Boolean
   Dim ii As Integer, jj As Integer
   Dim strHTAip As String 'Added by Morgan 2019/12/9
   Dim iErr As Integer
   
On Error GoTo ErrHnd

   'A以後的員工號修改指紋時改輸1####(B,2#### 類推),要同時刪除該筆卡號
   If cboHtaIp.Tag <> "" And txtSrcCardNo.Tag <> "" And txtST01 <> txtSrcCardNo.Tag Then
      '更新刷卡紀錄
      cnnConnection.Execute "update pollrecord set pr03='" & txtST01 & "' where pr03='" & txtSrcCardNo.Tag & "'", intI
      
      HTAip = cboHtaIp.Tag
      If HTAqueryCard(txtSrcCardNo.Tag, , True, , iRCode) = True Then
         If HTAdeleteCard(txtSrcCardNo.Tag) = False Then
            MsgBox "考勤機( " & HTAip & " )指紋( " & txtSrcCardNo.Tag & " ) 刪除失敗，作業取消!!"
            Exit Function
         Else
            List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtSrcCardNo.Tag & " ) 刪除成功!!", 0
         End If
      ElseIf iRCode <> 6 Then
         MsgBox "考勤機( " & HTAip & " )指紋( " & txtSrcCardNo.Tag & " ) 檢查失敗，作業取消!!"
         Exit Function
      End If
   End If
   
   'Added by Morgan 2019/12/9 可指定考勤機回寫
   'Modified by Morgan 2024/3/5
   'If cboHtaIp.ListIndex > 0 Then
   '   strHTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
   If cboHtaIp <> "" Then
      arrIP = Split(cboHtaIp, " ")
      strHTAip = arrIP(0)
   'end 2024/3/5
      ProgressBar1.max = 1
   Else
      strHTAip = m_stDomain
      jj = 0
      For ii = 0 To cboHtaIp.ListCount - 1
         If arrHTAIP(ii) <> "" And InStr(arrHTAIP(ii), strHTAip) = 1 Then
            jj = jj + 1
         End If
      Next
      If jj > 0 Then
         ProgressBar1.max = jj
      End If
   End If
   'end 2019/12/9
   ProgressBar2.max = 1
   List2.Clear
   iErr = 0
   ProgressBar2.Value = 0
   ProgressBar1.Value = 0
   
   '新增考勤機卡號資料
   For ii = 1 To cboHtaIp.ListCount - 1
      'Modified by Morgan 2024/3/5
      'HTAip = arrHTAIP(ii)
      If cboHtaIp <> "" Then
         HTAip = strHTAip
      Else
         HTAip = arrHTAIP(ii)
      End If
      'end 2024/3/5
   
      If HTAip <> "" And InStr(HTAip, strHTAip) = 1 Then
         bolErr = False
         If HTAqueryCard(txtST01, , True, , iRCode) = True Then
            List2.AddItem "考勤機( " & HTAip & " ) 指紋( " & txtST01 & " )檢查成功!!", 0
            If HTAdeleteCard(txtST01) = True Then
               List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtST01 & " ) 刪除成功!!", 0
            Else
               List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtST01 & " ) 刪除失敗!!", 0
               bolErr = True
            End If
         ElseIf iRCode = 6 Then
            List2.AddItem "考勤機( " & HTAip & " ) 無指紋 ( " & txtST01 & " )!!", 0
         Else
            List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtST01 & " ) 檢查失敗!!", 0
            bolErr = True
         End If
         
         If Not bolErr Then
            If txtFinger(1) & txtFinger(2) <> "" Then 'Added by Morgan 2021/6/9 空指紋不回寫(會錯)
               If HTAaddFingerPrinter(txtST01, IIf(txtDispName = "", txtST02, txtDispName), txtFinger(1), txtFinger(2), True, , , , IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1)), IIf(Check2.Value Or Check1.Value, True, False)) = True Then
                  List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtST01 & " ) 新增成功!!", 0
               Else
                  List2.AddItem "考勤機( " & HTAip & " ) 指紋 ( " & txtST01 & " ) 新增失敗!!", 0
                  bolErr = True
               End If
            End If
         End If
         If bolErr Then iErr = iErr + 1
         ProgressBar1.Value = ProgressBar1.Value + 1
         DoEvents
         If cboHtaIp <> "" Then Exit For
      End If
   Next
   ProgressBar2.Value = ProgressBar2.Value + 1
   
   UpdateMachin = True
   MsgBox "指紋回寫完成！" & IIf(iErr > 0, "但有 " & iErr & " 筆失敗！", ""), IIf(iErr > 0, vbExclamation, vbInformation)
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function
Private Function DeleteFinger() As Boolean
   Dim iRCode As Integer
   Dim ii As Integer
   
On Error GoTo ErrHnd

   strExc(0) = "select * from pollrecord where pr03='" & txtST01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("已有刷卡紀錄，是否確定要刪除？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   '刪除考勤機卡號資料
   For ii = 0 To cboHtaIp.ListCount - 1
      'Modified by Morgan 2015/12/1
      'HTAip = cboHtaIp.List(ii)
      HTAip = arrHTAIP(ii)
      'end 2015/12/1
      If HTAqueryCard(txtST01, , True, , iRCode) = True Then
         If HTAdeleteCard(txtST01) = False Then
            MsgBox "考勤機(" & HTAip & ")指紋刪除失敗，作業取消!!"
            Exit Function
         End If
      ElseIf iRCode <> 6 Then
         MsgBox "考勤機(" & HTAip & ")指紋檢查失敗，作業取消!!"
         Exit Function
      End If
   Next
   '刪除資料庫卡號資料
   strSql = "delete staffcarddata where scd02='" & txtST01 & "'"
   cnnConnection.Execute strSql, intI
   
   DeleteFinger = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

Private Sub txtSrcCardNo_GotFocus()
   CloseIme
End Sub

Private Sub txtSrcCardNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtST01_GotFocus()
   TextInverse txtST01
   CloseIme
End Sub

Private Sub txtST01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'
Private Function UpdateTimeZone() As Boolean
   Dim iRCode As Integer, bolDelete As Boolean
   Dim ii As Integer, jj As Integer
   Dim strHTAip As String, strCardNo As String
   Dim iTimeZone As Integer, iStatus As Integer
   Dim bNoCheck As Boolean 'Added by Morgan 2023/10/16
   
   iTimeZone = IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1)) 'Added by Morgan 2023/10/16
   bNoCheck = IIf(Check1.Value = vbChecked Or Check2.Value = vbChecked, True, False) 'Added by Morgan 2023/10/16
   'bNoCheck = False
   If cboHtaIp.ListIndex >= 0 Then
      strHTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
   Else
      strHTAip = m_stDomain
   End If
   
   For ii = 0 To cboHtaIp.ListCount - 1
      HTAip = arrHTAIP(ii)
      If strHTAip <> "" And InStr(HTAip, strHTAip) = 1 Then
         If CheckLevel(HTAip & ";", "門禁機IP") = True Then
            List2.AddItem "刪除門禁機( " & HTAip & " )指紋( " & txtST01 & " )...", 0 'Added by Morgan 2024/9/30
            DoEvents
            
            If HTAdeleteCard(txtST01) = False Then
               'Modified by Morgan 2024/9/30
               'MsgBox "門禁機( " & HTAip & " )卡號( " & txtST01 & " )更新(刪除)失敗，作業取消!!"
               List2.List(0) = List2.List(0) & "失敗!!"
               'end 2024/9/30
               Exit Function
            End If
            List2.List(0) = List2.List(0) & "成功!!" 'Added by Morgan 2024/9/30
            
            'iStatus:轉2進位 Bit0:0=正常卡,1=黑名單;Bit1:0=卡片,1=指紋;Bit2:0=假日不檢查,1=假日檢查;Bit3:0=時段不檢查,1=時段檢查
            If txtFinger(1) <> "" Or txtFinger(2) <> "" Then
               List2.AddItem "新增門禁機( " & HTAip & " )指紋( " & txtST01 & " )...", 0 'Added by Morgan 2024/9/30
               DoEvents
               
               'Modified by Morgan 2023/10/6
               'If HTAaddFingerPrinter(txtST01, IIf(txtDispName = "", txtST02, txtDispName), txtFinger(1), txtFinger(2), True, , , , , IIf(Check1.Value = vbChecked, True, False)) = False Then
               'Modified by Morgan 2023/10/16
               'If HTAaddFingerPrinter(txtST01, IIf(txtDispName = "", txtST02, txtDispName), txtFinger(1), txtFinger(2), True, , , , IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1))) = False Then
               If HTAaddFingerPrinter(txtST01, IIf(txtDispName = "", txtST02, txtDispName), txtFinger(1), txtFinger(2), True, , , , iTimeZone, bNoCheck) = False Then
                  'Modified by Morgan 2024/9/30
                  'MsgBox "門禁機( " & HTAip & " )卡號( " & txtST01 & " )更新失敗(已刪除)!!"
                  List2.List(0) = List2.List(0) & "失敗!!"
                  'end 2024/9/30
                  Exit Function
               End If
               List2.List(0) = List2.List(0) & "成功!!" 'Added by Morgan 2024/9/30
               
            End If
            For jj = 0 To List1.ListCount - 1
               strCardNo = List1.List(jj)
               If strCardNo <> "" Then
                  List2.AddItem "刪除門禁機( " & HTAip & " )卡號( " & strCardNo & " )...", 0 'Added by Morgan 2024/9/30
                  DoEvents
                  
                  If HTAdeleteCard(strCardNo) = False Then
                     'Modified by Morgan 2024/9/30
                     'MsgBox "門禁機( " & HTAip & " )卡號( " & strCardNo & " )更新(刪除)失敗，作業取消!!"
                     List2.List(0) = List2.List(0) & "失敗!!"
                     'end 2024/9/30
                     Exit Function
                  End If
                  List2.List(0) = List2.List(0) & "成功!!" 'Added by Morgan 2024/9/30
                  List2.AddItem "新增門禁機( " & HTAip & " )卡號( " & strCardNo & " )...", 0 'Added by Morgan 2024/9/30
                  DoEvents
                  
                  'Modified by Morgan 2023/10/6
                  'If HTAaddCard(strCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , , IIf(Check1.Value = vbChecked, True, False)) = False Then
                  'Modified by Morgan 2023/10/16
                  'If HTAaddCard(strCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , IIf(Check2.Value = vbChecked, 0, IIf(Check1.Value = vbChecked, 4, 1))) = False Then
                  If HTAaddCard(strCardNo, IIf(txtDispName = "", txtST02, txtDispName), , , , , iTimeZone, bNoCheck) = False Then
                     'Modified by Morgan 2024/9/30
                     'MsgBox "門禁機(" & HTAip & ")卡號( " & strCardNo & " )新增失敗，作業取消!!"
                     List2.List(0) = List2.List(0) & "失敗!!"
                     'end 2024/9/30
                     Exit Function
                  End If
                  List2.List(0) = List2.List(0) & "成功!!" 'Added by Morgan 2024/9/30
               End If
            Next
         End If
      End If
   Next
   UpdateTimeZone = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

'Added by Morgan 2023/12/4
Private Sub ChkCardNo()
   Dim iRCode As Integer
   Dim ii As Integer, jj As Integer
   Dim strHTAip As String
   Dim strCardNo As String
   Dim iErr As Integer
   
On Error GoTo ErrHnd
   
   'Modified by Morgan 2024/3/5
   'If cboHtaIp.ListIndex > 0 Then
   '   strHTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
   If cboHtaIp <> "" Then
      arrIP = Split(cboHtaIp, " ")
      strHTAip = arrIP(0)
   'end 2024/3/5
      ProgressBar1.max = 1
   Else
      strHTAip = m_stDomain
      jj = 0
      For ii = 0 To cboHtaIp.ListCount - 1
         If arrHTAIP(ii) <> "" And InStr(arrHTAIP(ii), strHTAip) = 1 Then
            jj = jj + 1
         End If
      Next
      If jj > 0 Then
         ProgressBar1.max = jj
      End If
   End If
   
   If txtCardNo <> "" Then
      ProgressBar2.max = 1
   Else
      ProgressBar2.max = List1.ListCount
   End If
   
   List2.Clear
   iErr = 0
   ProgressBar2.Value = 0
   For jj = 0 To List1.ListCount - 1
      If txtCardNo <> "" Then
         strCardNo = txtCardNo
      Else
         strCardNo = List1.List(jj)
      End If
      ProgressBar1.Value = 0
      '檢查考勤機卡號資料
      For ii = 0 To cboHtaIp.ListCount - 1
         'Modified by Morgan 2024/3/5
         'HTAip = arrHTAIP(ii)
         If cboHtaIp <> "" Then
            HTAip = strHTAip
         Else
            HTAip = arrHTAIP(ii)
         End If
         'end 2024/3/5
         If strHTAip <> "" And InStr(HTAip, strHTAip) = 1 Then
            If HTAqueryCard(strCardNo, , True, , iRCode) = True Then
                List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )檢查成功!!", 0
            Else
               iErr = iErr + 1
               If iRCode = 6 Then
                  List2.AddItem "考勤機(" & HTAip & ")無卡號( " & strCardNo & " )!!", 0
               Else
                  List2.AddItem "考勤機( " & HTAip & " )卡號( " & strCardNo & " )檢查失敗!!", 0
               End If
            End If
            ProgressBar1.Value = ProgressBar1.Value + 1
            DoEvents
            If cboHtaIp <> "" Then Exit For
         End If
      Next
      ProgressBar2.Value = ProgressBar2.Value + 1
      If txtCardNo <> "" Then Exit For
   Next
   
   MsgBox "卡片檢查完成！" & IIf(iErr > 0, "但有 " & iErr & " 筆失敗！", ""), IIf(iErr > 0, vbExclamation, vbInformation)
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub

Private Sub ChkFinger()
   Dim iRCode As Integer, bolDelete As Boolean
   Dim ii As Integer, jj
   Dim strHTAip As String
   Dim iErr As Integer
   
On Error GoTo ErrHnd
   
   'Modified by Morgan 2024/3/5
   'If cboHtaIp.ListIndex > 0 Then
   '   strHTAip = Trim(Left(cboHtaIp, InStr(cboHtaIp, " ")))
   If cboHtaIp <> "" Then
      arrIP = Split(cboHtaIp, " ")
      strHTAip = arrIP(0)
   'end 2024/3/5
      ProgressBar1.max = 1
   Else
      strHTAip = m_stDomain
      jj = 0
      For ii = 0 To cboHtaIp.ListCount - 1
         If arrHTAIP(ii) <> "" And InStr(arrHTAIP(ii), strHTAip) = 1 Then
            jj = jj + 1
         End If
      Next
      If jj > 0 Then
         ProgressBar1.max = jj
      End If
   End If
   ProgressBar2.max = 1
   List2.Clear
   iErr = 0
   ProgressBar2.Value = 0
   ProgressBar1.Value = 0
   
   '檢查考勤機指紋資料
   For ii = 1 To cboHtaIp.ListCount - 1
      'Modified by Morgan 2024/3/5
      'HTAip = arrHTAIP(ii)
      If cboHtaIp <> "" Then
         HTAip = strHTAip
      Else
         HTAip = arrHTAIP(ii)
      End If
      'end 2024/3/5
      If strHTAip <> "" And InStr(HTAip, strHTAip) = 1 Then
         If HTAqueryCard(txtST01, , True, , iRCode) = True Then
            List2.AddItem "考勤機( " & HTAip & " )" & txtST01 & "指紋檢查成功!!", 0
         Else
            iErr = iErr + 1
            If iRCode = 6 Then
               List2.AddItem "考勤機(" & HTAip & ")無" & txtST01 & "指紋!!", 0
            Else
               List2.AddItem "考勤機( " & HTAip & " )" & txtST01 & "指紋檢查失敗!!", 0
            End If
         End If
         ProgressBar1.Value = ProgressBar1.Value + 1
         DoEvents
         If cboHtaIp <> "" Then Exit For
      End If
   Next
   ProgressBar2.Value = ProgressBar2.Value + 1
   
   MsgBox "指紋檢查完成！" & IIf(iErr > 0, "但有 " & iErr & " 筆失敗！", ""), IIf(iErr > 0, vbExclamation, vbInformation)
   Exit Sub
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Sub
