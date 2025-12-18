VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040154 
   BorderStyle     =   1  '單線固定
   Caption         =   "不得代理案件之客戶或代理人資料維護"
   ClientHeight    =   6684
   ClientLeft      =   108
   ClientTop       =   936
   ClientWidth     =   9144
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6694.487
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9047.999
   Begin VB.TextBox textNT01 
      Height          =   264
      Left            =   1710
      MaxLength       =   3
      TabIndex        =   0
      Top             =   660
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8505
      Top             =   60
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
            Picture         =   "frm12040154.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":031C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":0638
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":0814
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":0B30
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":0E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":1168
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":1484
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":17A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":1ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm12040154.frx":1DD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   576
      Left            =   0
      TabIndex        =   35
      Top             =   0
      Width           =   9144
      _ExtentX        =   16129
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
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7620
         Top             =   120
         _ExtentX        =   593
         _ExtentY        =   593
         _Version        =   393216
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   8040
         Top             =   90
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5676
      Left            =   60
      TabIndex        =   37
      Top             =   936
      Width           =   9048
      _ExtentX        =   15960
      _ExtentY        =   10012
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm12040154.frx":20F4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label41(15)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label41(14)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label41(13)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label41(10)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label30(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label29"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label6"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "LabNT17_2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(6)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(7)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "LabNT18_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textNT02"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textNT07"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textNT20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textNT08_2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textNT06"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textNT05"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textNT04"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textNT03"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textNT08"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textNT21"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textNT18"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textNT17"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textNT19"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textNT22"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "cmdOpenAtt"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "cmdAddAtt"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "cmdRemAtt"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textNT30"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lstAtt"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textNT31"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm12040154.frx":2110
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(1)=   "textNT23"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textNT11"
      Tab(1).Control(3)=   "textNT12"
      Tab(1).Control(4)=   "textNT13"
      Tab(1).Control(5)=   "textNT14"
      Tab(1).Control(6)=   "textNT10"
      Tab(1).Control(7)=   "textNT15"
      Tab(1).Control(8)=   "lstUsers(0)"
      Tab(1).Control(9)=   "textNT16"
      Tab(1).Control(10)=   "textNT09"
      Tab(1).Control(11)=   "Label1(22)"
      Tab(1).Control(12)=   "Label18"
      Tab(1).Control(13)=   "Label16"
      Tab(1).Control(14)=   "Label13"
      Tab(1).Control(15)=   "Label41(32)"
      Tab(1).Control(16)=   "Label41(2)"
      Tab(1).Control(17)=   "Label41(3)"
      Tab(1).Control(18)=   "Label41(4)"
      Tab(1).Control(19)=   "Label41(5)"
      Tab(1).Control(20)=   "Label41(6)"
      Tab(1).ControlCount=   21
      Begin VB.Frame Frame1 
         Caption         =   "管制對象"
         Height          =   880
         Left            =   48
         TabIndex        =   71
         Top             =   2160
         Width           =   8964
         Begin VB.TextBox textNT35 
            Height          =   270
            Left            =   4776
            Locked          =   -1  'True
            MaxLength       =   40
            MultiLine       =   -1  'True
            TabIndex        =   77
            Top             =   480
            Visible         =   0   'False
            Width           =   3756
         End
         Begin VB.CommandButton cmdRemoveNT35 
            Caption         =   "移除 ->"
            Height          =   285
            Left            =   3840
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAddNT35 
            Caption         =   "<- 新增"
            Height          =   285
            Left            =   3840
            TabIndex        =   13
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox Text1 
            Height          =   300
            Left            =   4656
            MaxLength       =   8
            TabIndex        =   12
            Text            =   "Text1"
            Top             =   120
            Width           =   1000
         End
         Begin VB.CommandButton CmdClear 
            Caption         =   "清除"
            Height          =   300
            Left            =   72
            TabIndex        =   72
            Top             =   220
            Width           =   684
         End
         Begin MSForms.Label lblFM2 
            Height          =   276
            Left            =   5712
            TabIndex        =   76
            Top             =   144
            Width           =   3000
            BackColor       =   16777215
            Size            =   "5292;487"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.ListBox ListBox1 
            Height          =   336
            Left            =   1176
            TabIndex        =   75
            Top             =   108
            Width           =   2604
            ScrollBars      =   3
            DisplayStyle    =   2
            Size            =   "4593;593"
            MatchEntry      =   0
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label lblCnt 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   228
            Left            =   744
            TabIndex        =   74
            Top             =   600
            Width           =   204
         End
         Begin VB.Label Lbl1 
            Caption         =   "數量："
            ForeColor       =   &H00000000&
            Height          =   228
            Index           =   6
            Left            =   120
            TabIndex        =   73
            Top             =   600
            Width           =   564
         End
      End
      Begin VB.TextBox textNT31 
         Height          =   270
         Left            =   2280
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   4164
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   -72240
         TabIndex        =   66
         Top             =   1830
         Width           =   1815
         Begin VB.CommandButton cmdRemove 
            Caption         =   "移除 ->"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   420
            Width           =   735
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "<- 新增"
            Height          =   285
            Index           =   0
            Left            =   45
            TabIndex        =   33
            Top             =   120
            Width           =   735
         End
         Begin VB.TextBox txtUserNo 
            Height          =   264
            Index           =   0
            Left            =   810
            MaxLength       =   6
            TabIndex        =   32
            Top             =   120
            Width           =   945
         End
         Begin MSForms.TextBox lblName 
            Height          =   300
            Index           =   0
            Left            =   810
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   425
            Width           =   900
            VariousPropertyBits=   671105055
            Size            =   "1587;529"
            Value           =   "lblName"
            SpecialEffect   =   0
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox textNT23 
         Height          =   270
         Left            =   -74850
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   2220
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.ListBox lstAtt 
         BeginProperty Font 
            Name            =   "細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   768
         ItemData        =   "frm12040154.frx":212C
         Left            =   1245
         List            =   "frm12040154.frx":212E
         Sorted          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4764
         Width           =   6990
      End
      Begin VB.TextBox textNT30 
         Height          =   270
         Left            =   210
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   5064
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdRemAtt 
         Caption         =   "-> 移除"
         Height          =   255
         Left            =   8250
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   5364
         Width           =   735
      End
      Begin VB.CommandButton cmdAddAtt 
         Caption         =   "<- 新增"
         Height          =   285
         Left            =   8250
         TabIndex        =   21
         Top             =   5064
         Width           =   735
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   255
         Left            =   8250
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   4800
         Width           =   735
      End
      Begin VB.TextBox textNT22 
         Height          =   270
         Left            =   1245
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   4464
         Width           =   6060
      End
      Begin VB.TextBox textNT19 
         Height          =   270
         Left            =   1245
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3120
         Width           =   6060
      End
      Begin VB.TextBox textNT17 
         Height          =   270
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox textNT18 
         Height          =   270
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1830
         Width           =   855
      End
      Begin VB.TextBox textNT11 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   25
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textNT12 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   26
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textNT13 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   27
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textNT14 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   28
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textNT10 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   24
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textNT15 
         Height          =   270
         Left            =   -70005
         TabIndex        =   29
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textNT21 
         Height          =   270
         Left            =   1245
         MaxLength       =   7
         TabIndex        =   17
         Top             =   4164
         Width           =   975
      End
      Begin VB.TextBox textNT08 
         Height          =   270
         Left            =   1245
         MaxLength       =   4
         TabIndex        =   8
         Top             =   1530
         Width           =   612
      End
      Begin VB.TextBox textNT03 
         Height          =   270
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   3
         Top             =   630
         Width           =   3360
      End
      Begin VB.TextBox textNT04 
         Height          =   270
         Left            =   4995
         MaxLength       =   30
         TabIndex        =   4
         Top             =   630
         Width           =   3360
      End
      Begin VB.TextBox textNT05 
         Height          =   270
         Left            =   1245
         MaxLength       =   30
         TabIndex        =   5
         Top             =   930
         Width           =   3360
      End
      Begin VB.TextBox textNT06 
         Height          =   270
         Left            =   4995
         MaxLength       =   30
         TabIndex        =   6
         Top             =   930
         Width           =   3360
      End
      Begin VB.TextBox textNT08_2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '沒有框線
         Height          =   252
         Left            =   1890
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1530
         Width           =   1695
      End
      Begin MSForms.ListBox lstUsers 
         Height          =   2740
         Index           =   0
         Left            =   -73365
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1125
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "1984;4833"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT16 
         Height          =   300
         Left            =   -73740
         TabIndex        =   30
         Top             =   1584
         Width           =   7116
         VariousPropertyBits=   675299355
         MaxLength       =   70
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT09 
         Height          =   300
         Left            =   -73740
         TabIndex        =   23
         Top             =   360
         Width           =   7116
         VariousPropertyBits=   675299355
         MaxLength       =   70
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT20 
         Height          =   708
         Left            =   1245
         TabIndex        =   16
         Top             =   3432
         Width           =   7116
         VariousPropertyBits=   -1463795685
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "12552;1249"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT07 
         Height          =   300
         Left            =   1245
         TabIndex        =   7
         Top             =   1230
         Width           =   7116
         VariousPropertyBits=   675299355
         MaxLength       =   80
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT02 
         Height          =   300
         Left            =   1248
         TabIndex        =   2
         Top             =   276
         Width           =   7116
         VariousPropertyBits=   675299355
         MaxLength       =   80
         Size            =   "12559;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox LabNT18_2 
         Height          =   300
         Left            =   2160
         TabIndex        =   69
         TabStop         =   0   'False
         Top             =   1840
         Width           =   900
         VariousPropertyBits=   671105055
         Size            =   "1587;529"
         Value           =   "LabNT18_2"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "文件可查詢人員："
         Height          =   180
         Index           =   22
         Left            =   -74820
         TabIndex        =   64
         Top             =   2010
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "附件："
         Height          =   180
         Index           =   7
         Left            =   168
         TabIndex        =   63
         Top             =   4800
         Width           =   912
      End
      Begin VB.Label Label4 
         Caption         =   "撤銷原因："
         Height          =   252
         Left            =   168
         TabIndex        =   61
         Top             =   4500
         Width           =   912
      End
      Begin VB.Label Label3 
         Caption         =   "原因："
         Height          =   252
         Left            =   168
         TabIndex        =   60
         Top             =   3120
         Width           =   912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "負責同仁："
         Height          =   180
         Index           =   6
         Left            =   165
         TabIndex        =   59
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門別："
         Height          =   180
         Index           =   9
         Left            =   3630
         TabIndex        =   58
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label LabNT17_2 
         AutoSize        =   -1  'True
         Caption         =   "LabNT17_2"
         Height          =   180
         Left            =   5160
         TabIndex        =   57
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label18 
         Caption         =   "地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   56
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "地址(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   55
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "地址(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   54
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73845
         TabIndex        =   53
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   -73845
         TabIndex        =   52
         Top             =   1020
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   3
         Left            =   -73845
         TabIndex        =   51
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   4
         Left            =   -70125
         TabIndex        =   50
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   5
         Left            =   -70125
         TabIndex        =   49
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   6
         Left            =   -70125
         TabIndex        =   48
         Top             =   1290
         Width           =   90
      End
      Begin VB.Label Label6 
         Caption         =   "備註："
         Height          =   252
         Left            =   168
         TabIndex        =   47
         Top             =   3480
         Width           =   912
      End
      Begin VB.Label Label5 
         Caption         =   "撤銷日期："
         Height          =   252
         Left            =   168
         TabIndex        =   46
         Top             =   4200
         Width           =   912
      End
      Begin VB.Label Label2 
         Caption         =   "國籍："
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   45
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label27 
         Caption         =   "名稱(中)："
         Height          =   255
         Left            =   165
         TabIndex        =   44
         Top             =   330
         Width           =   915
      End
      Begin VB.Label Label29 
         Caption         =   "名稱(英)："
         Height          =   255
         Left            =   165
         TabIndex        =   43
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label30 
         Caption         =   "名稱(日)："
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   42
         Top             =   1230
         Width           =   915
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   10
         Left            =   1125
         TabIndex        =   41
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   13
         Left            =   4875
         TabIndex        =   40
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   14
         Left            =   1125
         TabIndex        =   39
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   15
         Left            =   4875
         TabIndex        =   38
         Top             =   990
         Width           =   90
      End
   End
   Begin MSForms.TextBox TextBox1 
      Height          =   300
      Left            =   0
      TabIndex        =   70
      Top             =   0
      Width           =   6735
      VariousPropertyBits=   671107099
      Size            =   "11880;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   3031
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   618
      Width           =   5960
      VariousPropertyBits=   671107103
      Size            =   "10513;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "編號："
      Height          =   255
      Index           =   0
      Left            =   1110
      TabIndex        =   36
      Top             =   660
      Width           =   555
   End
End
Attribute VB_Name = "frm12040154"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/13 Form2.0已修改(LabNT18_2,textNT02,textNT07,textNT09,textNT16,textNT20,lblName(0),lstUsers(0))
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Create By Sindy 2012/3/15
Option Explicit

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM

' 變數宣告區
Dim m_EditMode As Integer

' 第一筆資料的本所案號
Dim m_FirstKEY(1) As String
' 最後一筆資料的本所案號
Dim m_LastKEY(1) As String
' 目前正在顯示的本所案號
Dim m_CurrKEY(1) As String

'執行各項功能的權限
Dim m_bInsert As Boolean, m_bUpdate As Boolean, m_bDelete As Boolean, m_bQuery As Boolean
Dim m_Txt As TextBox
Dim TF_NT As Integer
'Add by Amy 2024/07/19
Dim bolAddFinish As Boolean
Public m_PrevForm As Form, m_RCL01 As String '前一畫面/風險檢查編號

Private Const cTableName As String = "NOTAGENT" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱

Private Sub RefreshRange()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT count(*) FROM NOTAGENT "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 And rsTmp.Fields(0) > 0 Then
      rsTmp.Close
      strSql = "SELECT MIN(NT01) FROM NOTAGENT "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_FirstKEY(0) = rsTmp.Fields(0)
      End If
      rsTmp.Close
      
      strSql = "SELECT MAX(NT01) FROM NOTAGENT "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields(0)) = False Then: m_LastKEY(0) = rsTmp.Fields(0)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

'開啟附件
Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09

   If lstAtt.Text = "" Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And textNT31 <> "" Then
         tmpArr = Empty
         tmpArr = Split(textNT31.Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, "(") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      Else
      'end 2017/08/09
         PUB_OpenFtpFile textNT01, lstAtt.Text, Winsock1, 3
      End If 'end 2017/08/09
      
   End If
End Sub

Private Sub Form_Initialize()
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from notagent where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_NT = AdoRecordSet3.Fields.Count
   CheckOC3
   ReDim m_FieldList(TF_NT) As FIELDITEM
End Sub

Private Sub Form_Load()
   '取得使用者執行各項功能的權限
   m_bInsert = IsUserHasRightOfFunction("frm12040154", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm12040154", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm12040154", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm12040154", strFind, False)
   
   MoveFormToCenter Me
   
   SSTab1.Tab = 0
   
   textNT08_2.BackColor = &H8000000F
   textCUID.BackColor = &H8000000F
   
   ListBox1.Height = 720 'Added by Lydia 2025/07/24
   
   'Modify by Amy 2024/07/19
   InitialField
   RefreshRange
   If m_PrevForm Is Nothing = False Then
      '風險檢查資料維護
      If UCase(m_PrevForm.Name) = "FRM12040163" Then
         OnAction vbKeyF2
      End If
   Else
      m_EditMode = 0
      ShowFirstRecord
      SetCtrlReadOnly True
   End If
   'end 2024/07/19
   UpdateToolbarState
   
   'Added by Lydia 2017/08/09
   'If Pub_StrUserSt03 <> "M51" Then cmd1.Visible = False 'Mark by Lydia 2025/07/23 程式碼不用了
    
End Sub

' 初始化欄位陣列
Private Sub InitialField()
   Dim nIndex As Integer
   Dim strTmp As String
   
   ' 初始化欄位陣列
   For nIndex = 1 To TF_NT
      strTmp = Format(nIndex, "00")
      m_FieldList(nIndex - 1).fiName = "NT" & strTmp
      m_FieldList(nIndex - 1).fiOldData = Empty
      m_FieldList(nIndex - 1).fiNewData = Empty
      m_FieldList(nIndex - 1).fiType = 0 '文字型態
      Select Case nIndex
         Case 21:
            m_FieldList(nIndex - 1).fiType = 1 '數值型態
      End Select
   Next nIndex
End Sub

' 設定欄位的內容
Private Sub SetFieldNewData(ByVal strName As String, Optional ByVal strData As String = "#==#")
   Dim nIndex As Integer
   For nIndex = 0 To TF_NT - 1
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

' 更新欄位的內容
Private Sub UpdateFieldNewData()
Dim strTmp  As String
   
   '若新增資料
   If m_EditMode = 1 Then
        '若未輸入編號
      If IsEmptyText(textNT01) = True Then
         textNT01 = Right("000" & CStr(Val(GetMaxNum) + 1), 3)
      End If
   End If
   
   '編號
   If IsEmptyText(textNT01) = False Then
      SetFieldNewData "NT01", Format(textNT01, "000")
   Else
      SetFieldNewData "NT01", textNT01
   End If
   SetFieldNewData "NT02", textNT02
   SetFieldNewData "NT03", textNT03
   SetFieldNewData "NT04", textNT04
   SetFieldNewData "NT05", textNT05
   SetFieldNewData "NT06", textNT06
   SetFieldNewData "NT07", textNT07
   SetFieldNewData "NT08", textNT08
   SetFieldNewData "NT09", textNT09
   SetFieldNewData "NT10", textNT10
   SetFieldNewData "NT11", textNT11
   SetFieldNewData "NT12", textNT12
   SetFieldNewData "NT13", textNT13
   SetFieldNewData "NT14", textNT14
   SetFieldNewData "NT15", textNT15
   SetFieldNewData "NT16", textNT16
   SetFieldNewData "NT17", textNT17
   SetFieldNewData "NT18", textNT18
   SetFieldNewData "NT19", textNT19
   SetFieldNewData "NT20", textNT20
   '撤銷日期
   If IsEmptyText(textNT21) = False Then
      SetFieldNewData "NT21", DBDATE(textNT21)
   Else
      SetFieldNewData "NT21", textNT21
   End If
   SetFieldNewData "NT22", textNT22
   'add by sonia 2022/1/17
   If Right(textNT23, 1) = "," Then textNT23 = Left(textNT23, Len(textNT23) - 1)
   'end 2022/1/17
   SetFieldNewData "NT23", textNT23
   SetFieldNewData "NT30", textNT30
   SetFieldNewData "NT31", textNT31 'Added by Lydia 2017/08/09
   SetFieldNewData "NT35", textNT35 'Added by Lydia 2025/07/24
End Sub

' 從記錄中更新欄位內容
Private Sub UpdateFieldOldData(ByRef rsTmp As ADODB.Recordset)
   Dim nIndex As Integer
   Dim strTmp As String
   
   For nIndex = 0 To TF_NT - 1
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

' 讀取資料庫所有的資料
Private Sub QueryDB()
   RefreshRange
End Sub

' 清除欄位內的資料內容
Private Sub ClearField()
   Dim nIndex As Integer
   
   textNT01 = Empty
   textNT02 = Empty
   textNT03 = Empty
   textNT04 = Empty
   textNT05 = Empty
   textNT06 = Empty
   textNT07 = Empty
   textNT08 = Empty
   textNT08_2 = Empty
   textNT09 = Empty
   textNT10 = Empty
   textNT11 = Empty
   textNT12 = Empty
   textNT13 = Empty
   textNT14 = Empty
   textNT15 = Empty
   textNT16 = Empty
   textNT17 = Empty
   LabNT17_2 = Empty
   textNT18 = Empty
   LabNT18_2 = Empty
   textNT19 = Empty
   textNT20 = Empty
   textNT21 = Empty
   textNT22 = Empty
   textNT23 = Empty
   textNT30 = Empty
   textNT31 = Empty: textNT31.Tag = textNT31.Text 'Added by Lydia 2017/08/09
   
   txtUserNo(0) = ""
   lblName(0) = ""
   lstUsers(0).Clear
   lstAtt.Clear
   
   For nIndex = 0 To TF_NT - 1
      m_FieldList(nIndex).fiOldData = Empty
      m_FieldList(nIndex).fiNewData = Empty
   Next nIndex
   textCUID = ""

   lstUsers(0).Tag = ""   'add by sonia 2021/12/24
   
   'Added by Lydia 2025/07/24
   Text1.Text = "": Text1.Tag = ""
   lblFM2.Caption = "": lblCnt.Caption = ""
   ListBox1.Clear
   textNT35 = Empty: textNT35.Tag = textNT35.Text
   'end 2025/07/24
End Sub

' 更新各控制項的狀態
Private Sub SetCtrlReadOnly(ByVal bEnable As Boolean)
   textNT01.Locked = bEnable
   textNT02.Locked = bEnable
   textNT03.Locked = bEnable
   textNT04.Locked = bEnable
   textNT05.Locked = bEnable
   textNT06.Locked = bEnable
   textNT07.Locked = bEnable
   textNT08.Locked = bEnable
   textNT09.Locked = bEnable
   textNT10.Locked = bEnable
   textNT11.Locked = bEnable
   textNT12.Locked = bEnable
   textNT13.Locked = bEnable
   textNT14.Locked = bEnable
   textNT15.Locked = bEnable
   textNT16.Locked = bEnable
   textNT17.Locked = bEnable
   textNT18.Locked = bEnable
   textNT19.Locked = bEnable
   textNT20.Locked = bEnable
   textNT21.Locked = bEnable
   txtUserNo(0).Locked = bEnable
   
   cmdOpenAtt.Enabled = bEnable
      
   cmdAddAtt.Enabled = Not bEnable
   cmdRemAtt.Enabled = Not bEnable
   
   'Added by Lydia 2025/07/24
   cmdAddNT35.Enabled = Not bEnable
   cmdRemoveNT35.Enabled = Not bEnable
   CmdClear.Enabled = Not bEnable
   'end 2025/07/24
End Sub

' 更新Key的狀態
Private Sub SetKeyReadOnly(ByVal bEnable As Boolean)
   textNT01.Locked = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim tmpArr As Variant 'Added by Lydia 2025/07/24
   
   strSql = "SELECT * FROM NOTAGENT " & _
            "WHERE NT01 = '" & m_CurrKEY(0) & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ClearField
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NT01")) = False Then: textNT01 = rsTmp.Fields("NT01")
      If IsNull(rsTmp.Fields("NT02")) = False Then: textNT02 = rsTmp.Fields("NT02")
      If IsNull(rsTmp.Fields("NT03")) = False Then: textNT03 = rsTmp.Fields("NT03")
      If IsNull(rsTmp.Fields("NT04")) = False Then: textNT04 = rsTmp.Fields("NT04")
      If IsNull(rsTmp.Fields("NT05")) = False Then: textNT05 = rsTmp.Fields("NT05")
      If IsNull(rsTmp.Fields("NT06")) = False Then: textNT06 = rsTmp.Fields("NT06")
      If IsNull(rsTmp.Fields("NT07")) = False Then: textNT07 = rsTmp.Fields("NT07")
      If IsNull(rsTmp.Fields("NT08")) = False Then: textNT08 = rsTmp.Fields("NT08"): textNT08_2 = GetNationName(textNT08, 0)
      If IsNull(rsTmp.Fields("NT09")) = False Then: textNT09 = rsTmp.Fields("NT09")
      If IsNull(rsTmp.Fields("NT10")) = False Then: textNT10 = rsTmp.Fields("NT10")
      If IsNull(rsTmp.Fields("NT11")) = False Then: textNT11 = rsTmp.Fields("NT11")
      If IsNull(rsTmp.Fields("NT12")) = False Then: textNT12 = rsTmp.Fields("NT12")
      If IsNull(rsTmp.Fields("NT13")) = False Then: textNT13 = rsTmp.Fields("NT13")
      If IsNull(rsTmp.Fields("NT14")) = False Then: textNT14 = rsTmp.Fields("NT14")
      If IsNull(rsTmp.Fields("NT15")) = False Then: textNT15 = rsTmp.Fields("NT15")
      If IsNull(rsTmp.Fields("NT16")) = False Then: textNT16 = rsTmp.Fields("NT16")
      If IsNull(rsTmp.Fields("NT17")) = False Then: textNT17 = rsTmp.Fields("NT17"): LabNT17_2 = GetDepartmentName(textNT17)
      If IsNull(rsTmp.Fields("NT18")) = False Then: textNT18 = rsTmp.Fields("NT18"): LabNT18_2 = GetPrjSalesNM(textNT18)
      If IsNull(rsTmp.Fields("NT19")) = False Then: textNT19 = rsTmp.Fields("NT19")
      If IsNull(rsTmp.Fields("NT20")) = False Then: textNT20 = rsTmp.Fields("NT20")
      '撤銷日期
      If IsNull(rsTmp.Fields("NT21")) = False Then
         If rsTmp.Fields("NT21") <> "0" Then
            textNT21 = TAIWANDATE(rsTmp.Fields("NT21"))
         End If
      End If
      If IsNull(rsTmp.Fields("NT22")) = False Then: textNT22 = rsTmp.Fields("NT22")
      If IsNull(rsTmp.Fields("NT23")) = False Then: textNT23 = rsTmp.Fields("NT23")
      If IsNull(rsTmp.Fields("NT30")) = False Then: textNT30 = rsTmp.Fields("NT30")
      'Added by Lydia 2017/08/09
      If IsNull(rsTmp.Fields("NT31")) = False Then: textNT31 = rsTmp.Fields("NT31")
      
      SetlstUsers 0, textNT23
      SetList lstAtt, textNT30
      
      'Added by Lydia 2025/07/24
      textNT35 = "" & rsTmp.Fields("NT35")
      ListBox1.Clear
      lblCnt.Caption = "0"
      If textNT35 <> "" Then
         tmpArr = Split(textNT35, ",")
         lblCnt.Caption = UBound(tmpArr) + 1
         For intI = UBound(tmpArr) To 0 Step -1
            If Trim(tmpArr(intI)) <> "" Then
               strExc(0) = ChangeCustomerL(tmpArr(intI))
               strExc(1) = ""
               If Left(tmpArr(intI), 1) = "Y" Then
                  strExc(1) = GetFAgentName(strExc(0))
               Else
                  strExc(1) = GetCustomerName(strExc(0), "1")
               End If
               ListBox1.AddItem Left(strExc(0), 8) & " " & strExc(1), 0
            End If
         Next intI
      End If
      'end 2025/07/24
      
      ' 更新CUID
      UpdateCUID rsTmp
      ' 更新暫存區的資料
      UpdateFieldOldData rsTmp
   End If
   rsTmp.Close
   
   textNT02.Tag = textNT02.Text
   textNT03.Tag = textNT03.Text
   textNT04.Tag = textNT04.Text
   textNT05.Tag = textNT05.Text
   textNT06.Tag = textNT06.Text
   textNT07.Tag = textNT07.Text
   textNT09.Tag = textNT09.Text
   textNT10.Tag = textNT10.Text
   textNT11.Tag = textNT11.Text
   textNT12.Tag = textNT12.Text
   textNT13.Tag = textNT13.Text
   textNT14.Tag = textNT14.Text
   textNT15.Tag = textNT15.Text
   textNT16.Tag = textNT16.Text
   textNT31.Tag = textNT31.Text 'Added by Lydia 2017/08/09
   textNT35.Tag = textNT35.Text 'Added by Lydia 2025/07/24
EXITSUB:
   Set rsTmp = Nothing
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
   
   If IsNull(rsSrcTmp.Fields("NT24")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT24")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("NT24"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NT25")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT25")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NT25"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NT26")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT26")) = False Then
         strTemp = rsSrcTmp.Fields("NT26")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NT27")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT27")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("NT27"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NT28")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT28")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("NT28"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("NT29")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("NT29")) = False Then
         strTemp = rsSrcTmp.Fields("NT29")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(6, " ") & _
              "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
End Sub

' 顯示資料
Private Sub ShowCurrRecord(ByVal strKEY01 As String)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   If IsRecordExist(strKEY01) = True Then
      m_CurrKEY(0) = strKEY01
   Else
      strSql = "SELECT NT01 FROM NOTAGENT " & _
               "WHERE NT01 = '" & m_CurrKEY(0) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
         rsTmp.Close
         UpdateCtrlData
         GoTo EXITSUB
      End If
      rsTmp.Close
      
      strSql = "SELECT NT01 FROM NOTAGENT " & _
               "WHERE NT01 = (SELECT MIN(NT01) FROM NOTAGENT " & _
                              "WHERE NT01 > '" & m_CurrKEY(0) & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
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
   
'   strSql = "SELECT NT01 FROM NOTAGENT " & _
'            "WHERE NT01 = '" & m_CurrKEY(0) & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT NT01 FROM NOTAGENT " & _
            "WHERE NT01 = (SELECT MAX(NT01) FROM NOTAGENT " & _
                           "WHERE NT01 < '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
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
   
'   strSql = "SELECT NT01 FROM NOTAGENT " & _
'            "WHERE NT01 = '" & m_CurrKEY(0) & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
'      rsTmp.Close
'      UpdateCtrlData
'      GoTo EXITSUB
'   End If
'   rsTmp.Close
   
   strSql = "SELECT NT01 FROM NOTAGENT " & _
            "WHERE NT01 = (SELECT MIN(NT01) FROM NOTAGENT " & _
                           "WHERE NT01 > '" & m_CurrKEY(0) & "') "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NT01")) = False Then: m_CurrKEY(0) = rsTmp.Fields("NT01")
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

' 更新toolbar上按紐的狀態
Private Sub UpdateToolbarState()
   Select Case m_EditMode
      ' 無任何動作
      Case 0:
         If m_bInsert Then
            tlbar.Buttons(1).Enabled = True
         Else
            tlbar.Buttons(1).Enabled = False
         End If
         If m_bUpdate Then
            tlbar.Buttons(2).Enabled = True
         Else
            tlbar.Buttons(2).Enabled = False
         End If
         If m_bDelete Then
            tlbar.Buttons(3).Enabled = True
         Else
            tlbar.Buttons(3).Enabled = False
         End If
         If m_bQuery Then
            tlbar.Buttons(4).Enabled = True
         Else
            tlbar.Buttons(4).Enabled = False
         End If
         If m_bQuery Then
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
         tlbar.Buttons(11).Enabled = False
         tlbar.Buttons(12).Enabled = False
         tlbar.Buttons(14).Enabled = True
         ' 新增
      Case 1, 2, 3, 4:
         tlbar.Buttons(1).Enabled = False
         tlbar.Buttons(2).Enabled = False
         tlbar.Buttons(3).Enabled = False
         tlbar.Buttons(4).Enabled = False
         tlbar.Buttons(6).Enabled = False
         tlbar.Buttons(7).Enabled = False
         tlbar.Buttons(8).Enabled = False
         tlbar.Buttons(9).Enabled = False
         tlbar.Buttons(11).Enabled = True
         tlbar.Buttons(12).Enabled = True
         tlbar.Buttons(14).Enabled = False
   End Select
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
      Case vbKeyEscape:
         If m_EditMode = 0 Then
            OnAction KeyCode
         Else
            OnAction vbKeyF10
         End If
   End Select
End Sub

'Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到Private Sub Form_KeyPress(KeyAscii As Integer)
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
      Case vbKeyReturn:
         If m_EditMode <> 0 Then
            KeyAscii = 0
            OnAction vbKeyF9
         End If
    End Select
End Sub

' 執行指令
Private Sub OnAction(ByVal KeyCode As Integer)
   Dim strTit As String, strMsg As String, nResponse
   
   Select Case KeyCode
      ' 新增
      Case vbKeyF2:
         m_EditMode = 1
         ClearField
         SetCtrlReadOnly False
         UpdateToolbarState
         If m_PrevForm Is Nothing Then SetInputEntry 'Modify by Amy 2024/07/17 +if
      ' 修改
      Case vbKeyF3:
         m_EditMode = 2
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
            OnWork
            UpdateToolbarState
         'Add by Amy 2024/07/19
         Else
            Exit Sub
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
         PUB_FilterFormText Me '修正畫面所有含跳行符號的文字框
         ' 將所有欄位的內容更新到欄位串列中的欄位內容項目
         If CheckDataValid() = True Then
            If m_EditMode = 1 Or m_EditMode = 2 Then
               '重新檢查欄位有效性
               If TxtValidate = False Then Exit Sub
            End If
            UpdateFieldNewData
            OnWork
            UpdateToolbarState
        'Add by Amy 2024/07/19
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
                  UpdateCtrlData
                  SetCtrlReadOnly True
                  UpdateToolbarState
               'Add by Amy 2024/07/19
               Else
                  Exit Sub
               End If
            Case Else
               m_EditMode = 0
               UpdateCtrlData
               SetCtrlReadOnly True
               UpdateToolbarState
         End Select
      ' 離開
      Case vbKeyEscape:
         Unload Me
   End Select
   If KeyCode <> vbKeyEscape And KeyCode <> vbKeyF3 Then
      'SSTab1.Tab = 0
   End If
   'Add by Amy 2024/07/19
   If m_PrevForm Is Nothing = False Then
      '按下[確定] or [取消] 且由[風險檢查資料維護],直接回[風險檢查資料維護]畫面
      If (KeyCode = vbKeyF9 Or KeyCode = vbKeyF10) And UCase(m_PrevForm.Name) = "FRM12040163" Then
         If KeyCode = vbKeyF9 And bolAddFinish = True Then
            MsgBox "風險檢查資料轉入完成" & vbCrLf & _
                            "不得代理編號：[" & textNT01 & "]", , "通知"
         End If
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Add by Amy 2024/07/19
   If m_PrevForm Is Nothing = False Then
      '風險檢查資料維護
      If UCase(m_PrevForm.Name) = "FRM12040163" And bolAddFinish = True Then
         Call m_PrevForm.AfterDelShowData(m_PrevForm.textRCL01)
      End If
      m_PrevForm.Show
   End If
   m_RCL01 = ""
   'end 2024/07/19
   
   PUB_KillTempFile "$$*.*" 'Added by Lydia 2017/08/09 清除暫存檔
   
   Set frm12040154 = Nothing
End Sub

Private Sub textNT01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'modify by sonia 2021/12/13
'Private Sub textNT02_KeyPress(KeyAscii As Integer)
Private Sub textNT02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'modify by sonia 2021/12/13
'Private Sub textNT07_KeyPress(KeyAscii As Integer)
Private Sub textNT07_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textNT08_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'modify by sonia 2021/12/13
'Private Sub textNT09_KeyPress(KeyAscii As Integer)
Private Sub textNT09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'日文地址要轉全形
'modify by sonia 2021/12/13
'Private Sub textNT16_KeyPress(KeyAscii As Integer)
Private Sub textNT16_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textNT17_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textNT18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'負責同仁
Private Sub textNT18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   Cancel = False
   LabNT18_2 = Empty
   textNT17 = Empty
   LabNT17_2 = Empty
   If IsEmptyText(textNT18) = False Then
      LabNT18_2 = GetPrjSalesNM(textNT18)
      If IsEmptyText(LabNT18_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "無此人員！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT18_GotFocus
      Else
         textNT17 = GetStaffDepartment(textNT18)
         LabNT17_2 = GetDepartmentName(textNT17)
      End If
   End If
End Sub

Private Sub textNT21_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

' 按下 ToolBar 的 Button
Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
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

' 檢查記錄是否已經存在
Private Function IsRecordExist(ByVal strKEY01 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   IsRecordExist = False
   strSql = "SELECT * FROM NOTAGENT " & _
            "WHERE NT01 = '" & strKEY01 & "' "
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

' 新增記錄
Private Sub AddRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strNT01 As String
   Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   strNT01 = Format(textNT01, "000")
   
   ' 檢查記錄是否已存在
   If IsRecordExist(strNT01) = True Then
      strTit = "新增資料"
      strMsg = "該筆記錄已存在"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      UpdateCtrlData
      GoTo ErrHand
   End If
   
   bFirst = True
   bDifference = False
   strSql = "INSERT INTO NOTAGENT ("
   For nIndex = 0 To TF_NT - 1
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
   'strSql = strSql & ",NT31" 'Added by Lydia 2023/08/18 FTP路徑
   
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   bFirst = True
   For nIndex = 0 To TF_NT - 1
      strTmp = Empty
      'Added by Lydia 2017/08/09 跳過FTP路徑
      If nIndex = 31 Then
        'strSql = strSql & ",NULL " 'Mark by Lydia 2023/08/18 debug
      Else
      'end 2017/08/09
        If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
           If m_FieldList(nIndex).fiType = 0 Then
              '字串中有單引號的處理
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
      End If 'end 2017/08/09
   Next nIndex
   strSql = strSql & ")"
   
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Added by Lydia 2017/08/09 判斷移檔日期
   If strSrvDate(1) >= CR_NewDate Then
      If UpdateAttFile(strNT01, iErr, sErrMsg) = False Then
         GoTo ErrHand
      Else
         If textNT31.Text <> textNT31.Tag Then
            strSql = "UPDATE NOTAGENT SET NT31='" & textNT31.Text & "' WHERE NT01='" & strNT01 & "' "
            cnnConnection.Execute strSql
         End If
      End If
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'   Else
'   'end 2017/08/09
'        '上傳附件檔
'        If UploadAtt(strNT01, iErr, sErrMsg) = False Then
'           GoTo ErrHand
'        End If
'end 2024/8/2
   End If 'end 2017/08/09
   
   'Add by Amy 2024/07/19 從[風險檢查資料維護]轉過來者,[刪除]風險檢查資料
   If m_PrevForm Is Nothing = False Then
      If UCase(m_PrevForm.Name) = "FRM12040163" Then
         '加入轉入之編號
         strSql = "Update RiskCheckList Set RCL23='" & strSrvDate(2) & "轉入不得代理名單中(編號:" & strNT01 & ");'||RCL23 " & _
                        "Where RCL01='" & m_RCL01 & "' "
         cnnConnection.Execute strSql
         '刪除[風險檢查資料維護]資料
         strSql = "Delete RiskCheckList Where RCL01='" & m_RCL01 & "' "
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
   
   cnnConnection.CommitTrans
   bolAddFinish = True 'Add by Amy 2024/07/19
   
   If (strNT01 < m_FirstKEY(0)) Or (strNT01 > m_LastKEY(0)) Then
      RefreshRange
   End If
   
   ShowCurrRecord strNT01
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

' 修改記錄
Private Sub ModRecord()
   Dim strSql As String
   Dim strTmp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strNT01 As String
   Dim iErr As Integer, sErrMsg As String
   Dim arrFile1
   Dim ii As Integer
   Dim bolRemove As Boolean
   
On Error GoTo ErrHand
   
   strNT01 = m_CurrKEY(0)
   
   strSql = "begin user_data.user_enabled:=1; UPDATE NOTAGENT SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To TF_NT - 1
      strTmp = Empty
      If nIndex < 45 Or nIndex > 50 Then
         'Added by Lydia 2017/08/09 跳過FTP路徑
         If nIndex = 30 Then
         Else
         'end 2017/08/09
            If m_FieldList(nIndex).fiOldData <> m_FieldList(nIndex).fiNewData Then
               If m_FieldList(nIndex).fiType = 0 Then
                  If m_FieldList(nIndex).fiNewData = Empty Then
                     strTmp = m_FieldList(nIndex).fiName & " = NULL "
                  Else
                     '字串中有單引號的處理
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
         End If 'end 2017/08/09
      End If
   Next nIndex
   strSql = strSql & " " & _
                  "WHERE NT01 = '" & strNT01 & "'; end; "

   If bDifference = True Then
      cnnConnection.BeginTrans
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
      
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate Then
         If UpdateAttFile(strNT01, iErr, sErrMsg) = False Then
            GoTo ErrHand
         Else
            If textNT31.Text <> textNT31.Tag Then
               strSql = "UPDATE NOTAGENT SET NT31='" & textNT31.Text & "' WHERE NT01='" & strNT01 & "' "
               cnnConnection.Execute strSql
            End If
         End If
         
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'      Else
'      'end 2017/08/09
'         '上傳附件檔
'         If UploadAtt(strNT01, iErr, sErrMsg) = False Then
'            GoTo ErrHand
'         End If
'        '檔案有異動時，移掉的要刪除
'         bolRemove = False
'         If m_FieldList(29).fiNewData <> m_FieldList(29).fiOldData Then
'            arrFile1 = Split(m_FieldList(29).fiOldData, ",")
'            For ii = LBound(arrFile1) To UBound(arrFile1)
'               If InStr(m_FieldList(29).fiNewData & ",", arrFile1(ii) & ",") > 0 Then
'                  arrFile1(ii) = ""
'               Else
'                  bolRemove = True
'               End If
'            Next
'            If bolRemove = True Then
'               If RemoveAtt(strNT01, Join(arrFile1, ","), iErr, sErrMsg) = False Then
'                  GoTo ErrHand
'               End If
'            End If
'         End If
'end 2024/8/2

      End If 'end 2017/08/09
      
      cnnConnection.CommitTrans
      ShowCurrRecord strNT01
   End If
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

' 刪除記錄
Private Sub DelRecord()
   Dim strSql As String
   Dim strNT01 As String
   Dim iErr As Integer, sErrMsg As String
   
On Error GoTo ErrHand
   
   strNT01 = m_CurrKEY(0)

   'Added by Lydia 2017/08/09 判斷移檔日期
   If m_FieldList(30).fiOldData <> "" And strSrvDate(1) >= CR_NewDate Then
      textNT31.Text = ""
      If UpdateAttFile(strNT01, iErr, sErrMsg) = False Then
         GoTo ErrHand
      End If
   End If
   'end 2017/08/09
      
   strSql = "DELETE FROM NOTAGENT " & _
            "WHERE NT01 = '" & strNT01 & "' "
   
   cnnConnection.BeginTrans
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   'Modifie by Lydia 2017/08/09 判斷移檔日期之前
   'If m_FieldList(29).fiOldData <> "" Then
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'   If m_FieldList(29).fiOldData <> "" And strSrvDate(1) < CR_NewDate Then
'      If RemoveAtt(strNT01, m_FieldList(29).fiOldData, iErr, sErrMsg) = False Then
'         GoTo ErrHand
'      End If
'   End If
'end 2024/8/2
   
   cnnConnection.CommitTrans
   ' 只有刪除的是最後一筆才須重新取的第一筆及最後一筆
   If (strNT01 = m_LastKEY(0)) Or (strNT01 = m_FirstKEY(0)) Then
      RefreshRange
   End If
   ShowCurrRecord strNT01
   Exit Sub
   
ErrHand:
   cnnConnection.RollbackTrans
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   ElseIf iErr <> 0 Then
      MsgBox sErrMsg, vbCritical
   End If
End Sub

' 查詢記錄
Private Function QueryRecord() As Boolean
   Dim strNT01 As String
   
   QueryRecord = False
   strNT01 = Format(textNT01, "000")
   
   textCUID = ""
   If IsRecordExist(strNT01) = True Then
      m_CurrKEY(0) = strNT01
      QueryRecord = True
      UpdateCtrlData
   Else
      QueryRecord = False
   End If

   UpdateToolbarState
End Function

' 使用者按下確定的按紐
Private Sub OnWork()
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Select Case m_EditMode
      Case 1: '新增
'         '重新檢查欄位有效性
'         If TxtValidate = False Then Exit Sub
         AddRecord
         RefreshRange
      Case 2: '修改
'         '重新檢查欄位有效性
'         If TxtValidate = False Then Exit Sub
         ModRecord
      Case 3: '刪除
         DelRecord
         RefreshRange
      Case 4: '查詢
         If QueryRecord = False Then
            strMsg = "無此資料"
            strTit = "查詢資料"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            UpdateCtrlData
         End If
   End Select
   m_EditMode = 0
   SetCtrlReadOnly True
EXITSUB:
End Sub

' 開始輸入資料
Private Sub SetInputEntry()
   Select Case m_EditMode
      Case 1: textNT01.SetFocus
      Case 2: textNT02.SetFocus
      Case 4: textNT01.SetFocus
   End Select
End Sub

'編號
Private Sub textNT01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim LongMaxNum As Long
   
   Cancel = False
   '若有輸入編號
   If IsEmptyText(textNT01) = False Then
      '補滿3碼
      textNT01 = Right("000" & textNT01, 3)
      '在新增時輸入的編號
      Select Case m_EditMode
         Case 1 '新增
            '不可大於目前檔案的最大號數
            LongMaxNum = Val(GetMaxNum)
            If Val(textNT01) > LongMaxNum Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "編號不可大於" & LongMaxNum
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNT01_GotFocus
               Exit Sub
            End If
            '不可已存在
            If IsRecordExist(textNT01) = True Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "該筆編號已存在"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNT01_GotFocus
               Exit Sub
            End If
      End Select
   End If
EXITSUB:
End Sub

'名稱(中)
Private Sub textNT02_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNT02) = False Then
      If StrLength(textNT02) > textNT02.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "名稱(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT02_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'名稱(日)
Private Sub textNT07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNT07) = False Then
      If StrLength(textNT07) > textNT07.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "名稱(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT07_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'國籍
Private Sub textNT08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   
   Cancel = False
   textNT08_2 = Empty
   If IsEmptyText(textNT08) = False Then
      textNT08_2 = GetNationName(textNT08, 0)
      If IsEmptyText(textNT08_2) = True Then
         Select Case m_EditMode
            Case 1, 2:
               Cancel = True
               strTit = "檢核資料"
               strMsg = "國籍不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNT08_GotFocus
            Case Else:
         End Select
      End If
   End If
End Sub

'地址(中)
Private Sub textNT09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNT09) = False Then
      If StrLength(textNT09) > textNT09.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "地址(中)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT09_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'地址(日)
Private Sub textNT16_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNT16) = False Then
      If StrLength(textNT16) > textNT16.MaxLength Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "地址(日)內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT16_GotFocus
      End If
   End If
   If Cancel = False Then CloseIme
End Sub

'撤銷日期
Private Sub textNT21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textNT21) = False Then
      If CheckIsTaiwanDate(textNT21, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "撤銷日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNT21_GotFocus
      End If
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim strTmp As String
   Dim nResponse
   CheckDataValid = False

   Select Case m_EditMode
      Case 4:
         ' 編號不可空白
         If IsEmptyText(textNT01) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入編號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textNT01.SetFocus
            GoTo EXITSUB
         End If
      Case Else:
   End Select
      
   Select Case m_EditMode
      Case 1, 2:
         '中文名稱, 英文名稱, 日文名稱不可全為空白
         If IsEmptyText(textNT02) = True And IsEmptyText(textNT03) = True And IsEmptyText(textNT07) = True Then
            strTit = "檢核資料"
            strMsg = "中文名稱, 英文名稱, 日文名稱不可全為空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textNT02.SetFocus
            GoTo EXITSUB
         End If
         '國籍
'cancel by sonia 2020/12/16 (2020/12外專提出資料有的無國籍)
'         If IsEmptyText(textNT08) = True Then
'            strTit = "檢核資料"
'            strMsg = "請輸入國籍！"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            SSTab1.Tab = 0
'            textNT08.SetFocus
'            GoTo EXITSUB
'         End If
'end 2020/12/16
         '負責同仁
         If IsEmptyText(textNT18) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入負責同仁！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textNT18.SetFocus
            GoTo EXITSUB
         End If
         '原因
         If IsEmptyText(textNT19) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入不得代理的原因！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textNT19.SetFocus
            GoTo EXITSUB
         End If
         '備註
         If IsEmptyText(textNT20) = True Then
            strTit = "檢核資料"
            strMsg = "備註不可空白！"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textNT20.SetFocus
            GoTo EXITSUB
         End If
         '文件可查詢人員
         If lstUsers(0).ListCount = 0 Then
            ShowMsg "文件可查詢人員不可空白！"
            SSTab1.Tab = 1
            txtUserNo(0).SetFocus
            txtUserNo_GotFocus 0
            Exit Function
         End If
         '撤銷日期或撤銷原因必須同時輸入或同時不輸
         If IsEmptyText(textNT21) = False Or IsEmptyText(textNT22) = False Then
            If IsEmptyText(textNT21) = True Then
               strTit = "檢核資料"
               strMsg = "有撤銷原因，撤銷日期不可空白！"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               SSTab1.Tab = 0
               textNT21.SetFocus
               GoTo EXITSUB
            End If
            If IsEmptyText(textNT22) = True Then
               strTit = "檢核資料"
               strMsg = "有撤銷日期，撤銷原因不可空白！"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               SSTab1.Tab = 0
               textNT22.SetFocus
               GoTo EXITSUB
            End If
         End If
      Case Else:
   End Select
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textNT01_GotFocus()
   InverseTextBox textNT01
End Sub

Private Sub textNT02_GotFocus()
   InverseTextBox textNT02
   OpenIme
End Sub

Private Sub textNT03_GotFocus()
   InverseTextBox textNT03
End Sub

Private Sub textNT04_GotFocus()
   InverseTextBox textNT04
End Sub

Private Sub textNT05_GotFocus()
   InverseTextBox textNT05
End Sub

Private Sub textNT06_GotFocus()
   InverseTextBox textNT06
End Sub

Private Sub textNT07_GotFocus()
   InverseTextBox textNT07
   OpenIme
End Sub

Private Sub textNT08_GotFocus()
   InverseTextBox textNT08
End Sub

Private Sub textNT09_GotFocus()
   InverseTextBox textNT09
   OpenIme
End Sub

Private Sub textNT10_GotFocus()
   InverseTextBox textNT10
End Sub

Private Sub textNT11_GotFocus()
   InverseTextBox textNT11
End Sub

Private Sub textNT12_GotFocus()
   InverseTextBox textNT12
End Sub

Private Sub textNT13_GotFocus()
   InverseTextBox textNT13
End Sub

Private Sub textNT14_GotFocus()
   InverseTextBox textNT14
End Sub

Private Sub textNT15_GotFocus()
   InverseTextBox textNT15
End Sub

Private Sub textNT16_GotFocus()
   InverseTextBox textNT16
   OpenIme
End Sub

Private Sub textNT17_GotFocus()
   InverseTextBox textNT17
End Sub

Private Sub textNT18_GotFocus()
   InverseTextBox textNT18
End Sub

Private Sub textNT19_GotFocus()
   InverseTextBox textNT19
End Sub

Private Sub textNT20_GotFocus()
   InverseTextBox textNT20
   OpenIme
End Sub

Private Sub textNT21_GotFocus()
   InverseTextBox textNT21
End Sub

Private Sub textNT22_GotFocus()
   InverseTextBox textNT22
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
Dim tmpArr1 As Variant, tmpArr2 As Variant 'Added by Lydia 2017/08/09

TxtValidate = False

If Me.textNT02.Enabled = True Then
   Cancel = False
   textNT02_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT07.Enabled = True Then
   Cancel = False
   textNT07_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT08.Enabled = True Then
   Cancel = False
   textNT08_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT09.Enabled = True Then
   Cancel = False
   textNT09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT16.Enabled = True Then
   Cancel = False
   textNT16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT18.Enabled = True Then
   Cancel = False
   textNT18_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If Me.textNT21.Enabled = True Then
   Cancel = False
   textNT21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Added by Lydia 2017/08/09 檢查長度
If CheckLengthIsOK(Me.textNT30, 600, False) = False Then
    MsgBox "全部的附件檔名超過最大長度600字元！" & vbCrLf & "(1個中文=2個字元)", vbCritical
    Exit Function
End If

'Added by Lydia 2017/08/09 檢查List和FTP檔名的數量是否一致
strExc(1) = "附件順序有誤，請全部移除後再新增附件！"
If (textNT30 = "" And textNT31 <> "") Or (textNT30 <> "" And textNT31 = "") Then
    ShowMsg strExc(1)
    Exit Function
End If

tmpArr1 = Empty: tmpArr2 = Empty
tmpArr1 = Split(textNT30, ",")
tmpArr2 = Split(textNT31, ",")
If UBound(tmpArr1) <> UBound(tmpArr2) Then
    ShowMsg strExc(1)
    Exit Function
End If

'預估一個ftp路徑約30字
If UBound(tmpArr2) > Format(600 / 30, "0") Then
   MsgBox "附件數量超過最大上限(" & Format(600 / 30, "0") & ")！", vbCritical
   Exit Function
End If
For intI = 0 To UBound(tmpArr1)
   If (Trim(tmpArr1(intI)) = "" And Trim(tmpArr2(intI)) <> "") Or (Trim(tmpArr1(intI)) <> "" And Trim(tmpArr2(intI)) = "") Then
      ShowMsg strExc(1)
      Exit Function
   End If
Next intI
'end 2017/08/09
      
'add by sonia 2021/12/24 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    Exit Function
End If
'end 2021/12/24

'Add by Amy 2024/07/19 由[風險檢查資料維護]進入者且無附件,彈提醒
If m_PrevForm Is Nothing = False Then
   If UCase(m_PrevForm.Name) = "FRM12040163" And Trim(textNT30) = MsgText(601) Then
      SSTab1.Tab = 0
      If MsgBox("[風險檢查資料]轉入者,需有上級同意之附件資料" & vbCrLf & _
                           "目前無任何附件,仍要繼續操作？" & vbCrLf & _
                           "是:存檔 否:回前畫面", vbYesNo + vbCritical, "無附件提醒！") = vbNo Then
         Exit Function
      End If
   End If
End If

TxtValidate = True
End Function

Private Function GetMaxNum() As String
   GetMaxNum = "0"
   strSql = "select count(*) from NOTAGENT "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         strSql = "select max(NT01) from NOTAGENT "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            GetMaxNum = RsTemp.Fields(0)
         End If
      End If
   End If
End Function

'新增接洽同仁
Private Sub cmdAdd_Click(Index As Integer)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   AddlstUsers Index
   'Modify By Sindy 2012/10/5
   'textNT23 = ComposeListX(Index)
   'modify by sonia 2022/1/ Form2.0 ITEMDATA不能用
   'If Trim(textNT23) <> "" Then textNT23 = Trim(textNT23) & ","
   'textNT23 = Trim(textNT23) & PUB_Num2Id(lstUsers(0).ITEMDATA(0))
   textNT23 = ComposeListX(Index)
   'end 2022/1/16
   '2012/10/5 End
   txtUserNo(Index).SetFocus
End Sub

'移除接洽同仁
Private Sub cmdRemove_Click(Index As Integer)
   If m_EditMode <> 1 And m_EditMode <> 2 Then Exit Sub
   RemovelstUsers Index
   textNT23 = ComposeListX(Index)
   txtUserNo(Index).SetFocus
End Sub

Private Sub AddlstUsers(p_idx As Integer)
   Dim idx As Integer, bFound As Boolean
   
   If txtUserNo(p_idx) <> "" And lblName(p_idx) <> "" Then
      '員工編號已可非數字需做轉換
      For idx = 0 To lstUsers(p_idx).ListCount - 1
         'modify by sonia 2021/12/24 Form 2.0不能用.ITEMDATA
         'If lstUsers(p_idx).ITEMDATA(idx) = PUB_Id2Num(txtUserNo(p_idx)) Then
         '   MsgBox "員工已存在於文件可查詢人員清單中！"
         '   txtUserNo(p_idx).SetFocus
         '   txtUserNo_GotFocus p_idx
         '   bFound = True
         '   Exit For
         'End If
         If InStr(lstUsers(p_idx).Tag, txtUserNo(p_idx)) > 0 Then
            MsgBox "員工已存在於開發人員清單中！"
            txtUserNo(p_idx).SetFocus
            txtUserNo_GotFocus p_idx
            bFound = True
         End If
         'end 2022/1/16
      Next
      If bFound = False Then
         lstUsers(p_idx).AddItem lblName(p_idx), 0
         'modify by sonia 2022/1/13 改Form 2.0不能用.ITEMDATA
         'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(txtUserNo(p_idx))
         lstUsers(p_idx).Tag = lstUsers(p_idx).Tag & txtUserNo(p_idx) & ","
         'end 2022/1/13
         txtUserNo(p_idx) = ""
         lblName(p_idx) = ""
      End If
   End If
End Sub

Private Sub txtUserNo_Change(Index As Integer)
   Dim strTempName As String
   If Len(txtUserNo(Index)) = 5 Then
      If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
         lblName(Index) = strTempName
      End If
   Else
      lblName(Index) = ""
   End If
End Sub

Private Sub txtUserNo_GotFocus(Index As Integer)
   TextInverse txtUserNo(Index)
End Sub

Private Sub txtUserNo_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtUserNo_Validate(Index As Integer, Cancel As Boolean)
   Dim strTempName As String
   If txtUserNo(Index).Visible = True Then
      If txtUserNo(Index) <> "" And lblName(Index) = "" Then
         If Len(txtUserNo(Index)) = 5 Then
            If ClsPDGetStaff(txtUserNo(Index), strTempName) = True Then
               lblName(Index) = strTempName
            End If
         End If
         If lblName(Index) = "" Then
            MsgBox "員工編號輸入錯誤！", vbExclamation
            Cancel = True
         End If
      End If
   End If
End Sub

Private Function ComposeListX(p_index As Integer) As String
   'modify by sonia 2022/1/16 改Form 2.0不能用.ITEMDATA
   'strExc(1) = ""
   'If lstUsers(p_index).ListCount > 0 Then
   '   strExc(1) = PUB_Num2Id(lstUsers(p_index).ITEMDATA(0))
   '   For intI = 1 To lstUsers(p_index).ListCount - 1
   '      strExc(1) = strExc(1) & "," & PUB_Num2Id(lstUsers(p_index).ITEMDATA(intI))
   '   Next
   'End If
   'ComposeListX = strExc(1)
   ComposeListX = lstUsers(p_index).Tag
   'end 2022/01/16
End Function

Private Sub RemovelstUsers(p_idx As Integer)
'Dim idx As Integer, ii As Integer
   'modify by sonia 2022/1/16
   'If lstUsers(p_idx).ListCount > 0 Then
   '   ii = 0
   '   For idx = 0 To lstUsers(p_idx).ListCount - 1
   '      If lstUsers(p_idx).Selected(ii) = True Then
   '         lstUsers(p_idx).RemoveItem ii
   '         'textNT23 = ComposeListX(Index)  'add by sonia 2022/1/13
   '         ii = ii - 1
   '      End If
   '      ii = ii + 1
   '   Next
   'End If
   lstUsers(p_idx).Tag = PUB_RemoveListBox2(lstUsers(p_idx), lstUsers(p_idx).Tag)
   'end 2022/01/10
End Sub

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
   lstUsers(p_idx).Tag = "" 'add by sonia 2022//16
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  'modify by sonia 2021/12/24 .Form2.0不能用.ITEMDATA
                  'lstUsers(p_idx).ITEMDATA(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub

Private Sub lstAtt_DblClick()
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub

'可多選,+顯示檔案大小
Private Sub cmdAddAtt_Click()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer
   Dim fs, f, s
   Dim strMid As String, strList As String 'Added by Lydia 2017/08/09
   
On Error GoTo ErrHnd
   
   stFileName = "*.*"
   strList = textNT31.Text  'Added by Lydia 2017/08/09
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = PUB_Getdesktop
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               Set fs = CreateObject("Scripting.FileSystemObject")
               Set f = fs.GetFile(stFileName)
               'Modified by Lydia 2017/08/09 存FTP檔名
               'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
               strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
               AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
               'end 2017/08/09
            Next
         Else
            stFileName = .FileName
            Set fs = CreateObject("Scripting.FileSystemObject")
            Set f = fs.GetFile(stFileName)
            'Modified by Lydia 2017/08/09 存FTP檔名
            'AddListX lstAtt, stFileName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
            strMid = PUB_GetNewFileNameSec(Mid(stFileName, InStrRev(stFileName, "\") + 1), , strList)
            AddListX lstAtt, PUB_StringFilter(stFileName) & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)", strMid
            'end 2017/08/09
         End If
         '改上傳到FTP,故只需留檔名
         'textnt30 = ComposeList(lstAtt)
         textNT30 = ComposeAttList(lstAtt)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdRemAtt_Click()
   If InStr(lstAtt, "\") = 0 And Pub_StrUserSt03 <> "M51" Then
         MsgBox "已上傳檔案不可移除！"
   ElseIf RemoveList(lstAtt) = True Then
      textNT30 = ComposeList(lstAtt)
      cmdAddAtt.SetFocus
   End If
End Sub

'上傳附件檔
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'Private Function UploadAtt(ByVal stKEY As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'   Dim hOpen As Long
'   Dim hConnection As Long
'   Dim hDir As Long
'   Dim bReturn As Boolean
'   Dim dwInternetFlags As Integer
'   Dim stDir As String
'   Dim stRemoteFile As String
'   Dim stLocalFile As String
'   Dim stItem As String
'   Dim idx As Integer
'   Dim iPos As Integer
'   Dim IsTimeOut As Boolean
'   Dim SeekTimer
'   Dim ACT_FTP_IP As String
'   Dim arrIP
'   Dim ii As Integer
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 不得代理案件存放路徑
'   If lstAtt.ListCount > 0 Then
'      For idx = 0 To lstAtt.ListCount - 1
'         stItem = lstAtt.List(idx)
'         iPos = InStr(stItem, "\")
'         If iPos > 0 Then
'            If InStrRev(stItem, " (") > 0 Then
'               stLocalFile = Left(stItem, InStrRev(stItem, " (") - 1)
'            Else
'               stLocalFile = stItem
'            End If
'            stRemoteFile = GetFileName(stLocalFile)
'
'            If hOpen = 0 Then
'               hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'               If hOpen = 0 Then
'                  iErrNo = 1
'                  stErrMsg = "網路錯誤！"
'                  GoTo OutPort
'               Else
'                  IsTimeOut = True
'                  If GOOD_FTP_IP <> "" Then
'                     arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'                  Else
'                     arrIP = Split(FTP_IP, ";")
'                  End If
'                  For ii = LBound(arrIP) To UBound(arrIP)
'                     ACT_FTP_IP = arrIP(ii)
'                     If ACT_FTP_IP <> "" Then
'                        '偵測 FTPServer 是否存在
'                        If Winsock1.State Then Winsock1.Close
'                        Winsock1.Connect ACT_FTP_IP, 21
'                        IsTimeOut = False
'                        SeekTimer = Timer
'                        Do While Winsock1.State = 6 And IsTimeOut = False
'                           DoEvents
'                           If Timer - SeekTimer > 1 Then
'                              IsTimeOut = True
'                           End If
'                        Loop
'                        If Winsock1.State Then Winsock1.Close
'                        If IsTimeOut = False Then
'                           Exit For
'                        End If
'                     End If
'                  Next
'
'                  '若是超過時間
'                  If IsTimeOut = True Then
'                     iErrNo = 2
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  Else
'                     GOOD_FTP_IP = ACT_FTP_IP
'                  End If
'
'                  hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                     "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'                  If hConnection = 0 Then
'                     iErrNo = 3
'                     stErrMsg = "無法與FTP Server建立連線！"
'                     GoTo OutPort
'                  ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                     iErrNo = 4
'                     stErrMsg = "切換至不得代理案件目錄失敗！"
'                     GoTo OutPort
'                  '切換至不得代理案件目錄
'                  ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                     hDir = FtpCreateDirectory(hConnection, stKEY)
'                     If hDir = 0 Then
'                        iErrNo = 5
'                        stErrMsg = "建立不得代理案件目錄失敗！"
'                        GoTo OutPort
'                     ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                        iErrNo = 6
'                        stErrMsg = "切換至不得代理案件目錄失敗！"
'                        GoTo OutPort
'                     End If
'                  End If
'               End If
'            End If
'
'            dwInternetFlags = FTP_TRANSFER_TYPE_BINARY
'            bReturn = FtpPutFile(hConnection, stLocalFile, stRemoteFile, dwInternetFlags, 0)
'            ' Upload successfully
'            If bReturn = False Then
'               iErrNo = 7
'               stErrMsg = "檔案上傳失敗！"
'               GoTo OutPort
'            End If
'         End If
'      Next
'   End If
'   UploadAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'
'End Function

'刪除附件檔
'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
'Private Function RemoveAtt(ByVal stKEY As String, stFiles As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
'   Dim hOpen As Long
'   Dim hConnection As Long
'   Dim bReturn As Boolean
'   Dim stDir As String
'   Dim IsTimeOut As Boolean
'   Dim SeekTimer
'   Dim ACT_FTP_IP As String
'   Dim arrIP
'   Dim ii As Integer, jj As Integer
'   Dim arrFile
'   Dim stRemoteFile As String
'
'   iErrNo = 0
'   stErrMsg = ""
'
'   stDir = 不得代理案件存放路徑
'   arrFile = Split(stFiles, ",")
'   For jj = LBound(arrFile) To UBound(arrFile)
'      If arrFile(jj) <> "" Then
'         If hOpen = 0 Then
'            hOpen = InternetOpen("Taie FTP", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
'            If hOpen = 0 Then
'               iErrNo = 1
'               stErrMsg = "網路錯誤！"
'               GoTo OutPort
'            Else
'               IsTimeOut = True
'               If GOOD_FTP_IP <> "" Then
'                  arrIP = Split(GOOD_FTP_IP & ";" & FTP_IP, ";")
'               Else
'                  arrIP = Split(FTP_IP, ";")
'               End If
'               For ii = LBound(arrIP) To UBound(arrIP)
'                  ACT_FTP_IP = arrIP(ii)
'                  If ACT_FTP_IP <> "" Then
'                     '偵測 FTPServer 是否存在
'                     If Winsock1.State Then Winsock1.Close
'                     Winsock1.Connect ACT_FTP_IP, 21
'                     IsTimeOut = False
'                     SeekTimer = Timer
'                     Do While Winsock1.State = 6 And IsTimeOut = False
'                        DoEvents
'                        If Timer - SeekTimer > 1 Then
'                           IsTimeOut = True
'                        End If
'                     Loop
'                     If Winsock1.State Then Winsock1.Close
'                     If IsTimeOut = False Then
'                        Exit For
'                     End If
'                  End If
'               Next
'
'               '若是超過時間
'               If IsTimeOut = True Then
'                  iErrNo = 2
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               Else
'                  GOOD_FTP_IP = ACT_FTP_IP
'               End If
'
'               hConnection = InternetConnect(hOpen, ACT_FTP_IP, FTP_Port, _
'                  "pgmid", "pgmpwd", INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
'               If hConnection = 0 Then
'                  iErrNo = 3
'                  stErrMsg = "無法與FTP Server建立連線！"
'                  GoTo OutPort
'               ElseIf FtpSetCurrentDirectory(hConnection, stDir) = False Then
'                  iErrNo = 4
'                  stErrMsg = "切換至不得代理案件目錄失敗！"
'                  GoTo OutPort
'               '切換至不得代理案件目錄
'               ElseIf FtpSetCurrentDirectory(hConnection, stKEY) = False Then
'                  '無法切換時當作已刪除
'                  'iErrNo = 6
'                  'stErrMsg = "切換至不得代理案件目錄失敗！"
'                  'GoTo OutPort
'                  Exit For
'               End If
'            End If
'         End If
'         If InStrRev(arrFile(jj), " (") > 0 Then
'            stRemoteFile = Left(arrFile(jj), InStrRev(arrFile(jj), " (") - 1)
'         Else
'            stRemoteFile = arrFile(jj)
'         End If
'         '刪除檔案不控制成功與否
'         bReturn = FtpDeleteFile(hConnection, stRemoteFile)
'      End If
'   Next
'
'   RemoveAtt = True
'
'OutPort:
'   If hOpen <> 0 Then InternetCloseHandle (hOpen)
'   If hConnection <> 0 Then InternetCloseHandle (hConnection)
'   If Winsock1.State Then Winsock1.Close
'
'End Function

Private Function GetFileName(ByVal FullPath As String) As String
   Dim stItem As String, iPos As Integer
   stItem = FullPath
   iPos = InStr(stItem, "\")
   Do While iPos > 0
      stItem = Mid(stItem, iPos + 1)
      iPos = InStr(stItem, "\")
   Loop
   GetFileName = stItem
End Function

Private Sub SetList(oList As ListBox, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

'Modified by Lydia 2017/08/09 +存FTP檔名 stFtpName
Private Function AddListX(oList As ListBox, stNewItem As String, stFtpName As String) As Boolean
   Dim idx As Integer, bFound As Boolean, stFileName As String
   If InStr(stNewItem, ",") > 0 Then
      MsgBox "逗號[,]為系統保留字，請重新命名！", vbExclamation
      cmdAddAtt.SetFocus
      Exit Function
   End If
   If stNewItem <> "" Then
      For idx = 0 To oList.ListCount - 1
         stFileName = GetFileName(oList.List(idx))
         If GetFileName(stNewItem) = stFileName Then
            MsgBox "附件[" & stFileName & "]已存在！"
            AddListX = False
            bFound = True
            Exit For
         End If
      Next
      If bFound = False Then
         oList.AddItem stNewItem, 0
         AddListX = True
         'Added by Lydia 2017/08/09 存FTP檔名 (堆疊)
         textNT31 = stFtpName & IIf(textNT31 <> "", ",", "") & textNT31
      End If
   End If
End Function

'附件
Private Function ComposeAttList(oList As ListBox) As String
   Dim iPos As Integer, stItem As String, stRtn As String, idx As Integer
   If oList.ListCount > 0 Then
      stItem = oList.List(0)
      stRtn = GetFileName(stItem)
      For idx = 1 To oList.ListCount - 1
         stItem = oList.List(idx)
         stRtn = stRtn & "," & GetFileName(stItem)
      Next
   End If
   ComposeAttList = stRtn
End Function

Private Function RemoveList(oList As ListBox) As Boolean
   Dim ii As Integer
   Dim tmpArr As Variant 'Added by Lydia 2017/08/09
   
   If oList.ListCount > 0 Then
      ii = 0
      Do While ii < oList.ListCount
         If oList.Selected(ii) = True Then
            RemoveList = True
            oList.RemoveItem ii
            'Added by Lydia 2017/08/09 移除FTP檔名(可複選)
            If textNT31 <> "" Then
               '重整FTP檔名
               textNT31 = Replace(textNT31, ",,", ",")
               If Left(textNT31, 1) = "," Then textNT31 = Mid(textNT31, 2)
               If Right(textNT31, 1) = "," Then textNT31 = Mid(textNT31, 1, Len(textNT31) - 1)
               tmpArr = Empty
               tmpArr = Split(textNT31, ",")
               If Trim(tmpArr(ii)) <> "" Then textNT31 = Replace(textNT31, Trim(tmpArr(ii)), "")
            End If
            'end 2017/08/09
            
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
      
      'Added by Lydia 2017/08/09 重整FTP檔名
      textNT31 = Replace(textNT31, ",,", ",")
      If Left(textNT31, 1) = "," Then textNT31 = Mid(textNT31, 2)
      If Right(textNT31, 1) = "," Then textNT31 = Mid(textNT31, 1, Len(textNT31) - 1)
      'end 2017/08/09
   End If
End Function

Private Function ComposeList(oList As ListBox, Optional p_iOpt As Integer = 0) As String
   Dim iPos As Integer, stItem As String
   strExc(1) = ""
   If oList.ListCount > 0 Then
      For intI = 0 To oList.ListCount - 1
         If p_iOpt = 0 Then
            iPos = InStr(oList.List(intI), Chr(1))
            If iPos > 0 Then
               stItem = Left(oList.List(intI), iPos - 1)
            Else
               stItem = oList.List(intI)
            End If
         Else
            stItem = Format(oList.ITEMDATA(intI), "00")
         End If
         stItem = GetFileName(stItem) 'Add By Sindy 2012/3/21
         If intI = 0 Then
            strExc(1) = stItem
         Else
            strExc(1) = strExc(1) & "," & stItem
         End If
      Next
   End If
   ComposeList = strExc(1)
End Function

'Added by Lydia 2017/08/09 新增／刪除附件
Private Function UpdateAttFile(ByVal stKey As String, Optional iErrNo As Integer, Optional stErrMsg As String) As Boolean
Dim arrTmp As Variant, arrOldTmp As Variant
Dim stFtpPath As String
Dim ii As Integer
Dim strMid As String
Dim stFileName As String

On Error GoTo OutPort
   
   iErrNo = 0
   stErrMsg = ""

   arrTmp = Empty: arrOldTmp = Empty
   arrTmp = Split(textNT31.Text, ",")
   arrOldTmp = Split(textNT31.Tag, ",")
   
   '先：刪除附件
   If textNT31.Tag <> "" Then
      For ii = 0 To UBound(arrOldTmp)
         If Trim(arrOldTmp(ii)) <> "" And InStr(textNT31.Text, Trim(arrOldTmp(ii))) = 0 Then
            If ChkNT31isDual(stKey, Trim(arrOldTmp(ii))) = False Then 'Added by Lydia 2025/07/23  檢查是否有相同檔案存在其他記錄
              If PUB_DelFtpFile2(stKey, Trim(arrOldTmp(ii)), cTableName) = False Then
                 GoTo OutPort
              End If
            End If 'Added by Lydia 2025/07/23
         End If
      Next ii
   End If
   
   '後：新增附件
   If textNT31.Text <> "" Then
      For ii = 0 To UBound(arrTmp)
         If Trim(arrTmp(ii)) <> "" And InStr(textNT31.Tag, Trim(arrTmp(ii))) = 0 Then
            'Added by Lydia 2025/07/23 檢查是否有相同檔案存在其他記錄
            If ChkNT31isDual(stKey, Trim(arrTmp(ii))) = True Then
               strMid = strMid & IIf(strMid <> "", ",", "") & Trim(arrTmp(ii))
            Else
            'end 2025/07/23
               stFileName = Trim(Mid(lstAtt.List(ii), 1, InStrRev(lstAtt.List(ii), "(") - 1))
               If PUB_PutFtpFile(stFileName, stKey, IIf(InStr(Trim(arrTmp(ii)), stKey & "_") = 0, stKey & "_", "") & Trim(arrTmp(ii)), stFtpPath, cTableName) = False Then
                  GoTo OutPort
               Else
                  strMid = strMid & IIf(strMid <> "", ",", "") & stFtpPath
               End If
            End If 'Added by Lydia 2025/07/23
         ElseIf Trim(arrTmp(ii)) <> "" Then
            strMid = strMid & IIf(strMid <> "", ",", "") & Trim(arrTmp(ii))
         End If
      Next ii
    textNT31.Text = strMid
   End If
   
   UpdateAttFile = True
   
   Exit Function
   
OutPort:
   iErrNo = Err.Number
   stErrMsg = Err.Description
   
End Function

'Added by Lydia 2017/08/09 搬檔
'Mark by Lydia 2025/07/23 程式碼不用了
'Private Sub Cmd1_Click()
'Dim stSQL As String, intR As Integer
'Dim rsQuery As ADODB.Recordset
'Dim stOldDir As String, stNewDir As String, stNewPath As String
'Dim oFileName As String, mFileName As String
'Dim strGrp As String, strList As String, strNameList As String
'Dim tmpArr As Variant
'Dim strTmpExc As String
'Dim stDownFile As String
'Dim strLost As String, strLostId As String
'
'   stOldDir = 不得代理案件存放路徑
'   stNewDir = PUB_GetFtpTableDir(stNewDir) & cTableName
'   stSQL = "select NT01,NT30 from NOTAGENT where NVL(NT30,'N') <> 'N' AND NVL(NT31,'N')='N' order by NT01 "
'   intR = 0
'   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
'   If intR = 1 Then
'      With rsQuery
'         .MoveFirst
'         MsgBox "開始工作，共" & .RecordCount & "筆記錄!"
'         Do While Not .EOF
'            '清除暫存檔
'            PUB_KillTempFile "$$*.*"
'
'            If strGrp <> "" & .Fields("NT01") Then
'               If strGrp <> "" Then
'                  strTmpExc = strTmpExc & "UPDATE NOTAGENT SET NT31='" & strList & "' WHERE NT01='" & strGrp & "' ;"
'               End If
'               strList = "": strNameList = ""
'               strGrp = "" & .Fields("NT01")
'               tmpArr = Empty
'               tmpArr = Split("" & .Fields("NT30"), ",")
'            End If
'
'            For intR = 0 To UBound(tmpArr)
'               If Trim(tmpArr(intR)) <> "" Then
'                   '先下載檔案
'                   stDownFile = ""
'                   '因為有附件檔名有包含刮號,直接到模組處理舊檔名
'                   strExc(1) = PUB_StringFilter(Trim(tmpArr(intR)))
'                   If InStr(strExc(1), "(") > 0 And InStr(strExc(1), " (") = 0 Then
'                      strExc(1) = Mid(strExc(1), 1, InStrRev(strExc(1), "(") - 1) & " " & Mid(strExc(1), InStrRev(strExc(1), "("))
'                   End If
'                   PUB_OpenFtpFile "" & .Fields("NT01"), strExc(1), Winsock1, "3", False, stDownFile
'
'                   If stDownFile = "" Then
'                       strLostId = strLostId & .Fields("CQ01") & "," & IIf(Len(strLostId) > 50, vbCrLf, "")
'                       strLost = strLost & .Fields("CQ01") & "_" & Trim(tmpArr(intR)) & vbCrLf
'                   Else
'                        oFileName = Trim(tmpArr(intR))
'                        oFileName = Trim(Mid(oFileName, 1, InStrRev(oFileName, "(") - 1))
'                        '新-FTP檔名(非中文)
'                        mFileName = PUB_GetNewFileNameSec(oFileName, "2", strNameList, "" & .Fields("NT01"))
'                        If PUB_PutFtpFile(stDownFile, "" & .Fields("NT01"), mFileName, stNewPath, cTableName) = True Then
'                           strList = strList & IIf(strList <> "", ",", "") & stNewPath
'                        Else
'                           MsgBox "Error !"
'                           Exit Sub
'                        End If
'                   End If
'               End If
'            Next intR
'            .MoveNext
'         Loop
'
'         '最後一筆
'         strTmpExc = strTmpExc & "UPDATE NOTAGENT SET NT31='" & strList & "' WHERE NT01='" & strGrp & "' ;"
'      End With
'
'      '清除暫存檔
'      PUB_KillTempFile "$$*.*"
'
'      If strTmpExc <> "" Then
'         tmpArr = Empty
'         tmpArr = Split(strTmpExc, ";")
'         cnnConnection.BeginTrans
'           For intR = 0 To UBound(tmpArr)
'              If Trim(tmpArr(intR)) <> "" Then
'                 cnnConnection.Execute Trim(tmpArr(intR)), intI
'              End If
'           Next intR
'         cnnConnection.CommitTrans
'         MsgBox "工作結束!"
'      End If
'   End If
'
'   If strLost <> "" Then
'      PUB_SendMail "QPGMR", "A3034", "", 不得代理案件存放路徑 & "在NT2缺少檔案", "資料夾:" & strLostId & vbCrLf & vbCrLf & "檔案名稱:" & strLost
'   End If
'
'   Set rsQuery = Nothing
'   Exit Sub
'
'ErrHandle:
'   cnnConnection.RollbackTrans
'
'OutPort:
'   Exit Sub
'
'End Sub
'end 2025/07/23

'Added by Lydia 2025/07/23 檢查是否有相同檔案存在其他記錄
Private Function ChkNT31isDual(ByVal pStKey As String, ByVal pStFileName As String) As Boolean
Dim intQ As Integer, strQuery As String
Dim rsQD As New ADODB.Recordset
   
   ChkNT31isDual = False
   If Trim(pStKey) = "" Or Trim(pStFileName) = "" Then Exit Function
   
   strQuery = "select nt01,nt31 from NotAgent where nt01<>'" & pStKey & "' and instr(upper(nt31),'" & UCase(pStFileName) & "') > 0 "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      ChkNT31isDual = True
   End If
   Set rsQD = Nothing
End Function

'Added by Lydia 2025/07/24
Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
Dim strTempA As String
   If Text1.Tag <> Text1.Text Then
      If Text1.Text = "" Then
         lblFM2.Caption = ""
      Else
         If Left(Text1, 1) <> "X" And Left(Text1, 1) <> "Y" Then
            MsgBox "請輸入代理人或申請人編號！", vbExclamation
            GoTo EXITSUB
         Else
            Text1 = ChangeCustomerL(Text1)
            If Left(Text1, 1) = "Y" Then
               strTempA = GetFAgentName(Text1)
            Else
               strTempA = GetCustomerName(Text1, "1")
            End If
            If strTempA <> "" Then
               lblFM2.Caption = strTempA
            Else
               lblFM2.Caption = ""
               MsgBox "資料庫無資料 !", vbInformation
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   Text1.Tag = Text1.Text
   
   Cancel = False
   Exit Sub
   
EXITSUB:
   Cancel = True
   Text1.SetFocus
   Text1_GotFocus
End Sub
'Modified by Lydia 2025/07/29 改為Public
Public Sub cmdAddNT35_Click()
   Call AddNT35No
   lblCnt = ListBox1.ListCount
End Sub

Private Sub cmdRemoveNT35_Click()
   Call RemoveNT35No
   lblCnt = ListBox1.ListCount
End Sub

Private Sub AddNT35No()
Dim bFound As Boolean
Dim intX As Integer
Dim strTempA As String

   bFound = True
   If Trim(Text1) = "" Then
      Exit Sub
   Else
      Call Text1_Validate(bFound)
      If bFound = True Then
         Exit Sub
      Else
         bFound = False
         strTempA = Mid(Trim(Text1) & String(8, "0"), 1, 8)
      End If
   End If
   
'------------------------------------------
   If strTempA <> "" And bFound = False Then
      If InStr(textNT35, strTempA) > 0 Then
         MsgBox strTempA & " " & lblFM2.Caption & "已存在於清單中！"
         cmdAddNT35.SetFocus
         bFound = True
      End If
      If bFound = False Then
         intX = ListBox1.ListCount
         ListBox1.AddItem strTempA & " " & lblFM2.Caption, intX
         textNT35 = textNT35 & "," & strTempA
         If Mid(textNT35, 1, 1) = "," Then textNT35 = Mid(textNT35, 2)
         '清除來源
         Text1 = ""
         Call Text1_Validate(bFound)
      End If
   End If
End Sub

Private Sub RemoveNT35No()
Dim strTempA As String, ii As Integer

   If ListBox1.ListCount > 0 Then  'ListBox1=>Form 2.0物件
      ii = 0
      Do While ii < ListBox1.ListCount
         If ListBox1.Selected(ii) = True Then
            strTempA = ListBox1.List(ii)
            strExc(1) = Trim(Left(strTempA, 8))
            textNT35 = Replace(textNT35, "," & strExc(1), "")
            textNT35 = Replace(textNT35, strExc(1) & ",", "")
            textNT35 = Replace(textNT35, strExc(1), "")
            
            ListBox1.RemoveItem ii
            '因若屬性為單選時會自動選取上一個項目會導致全部被刪除，故移除後需將索引設為-1(無勾選)
            If ListBox1.MultiSelect = 0 Then
               ListBox1.ListIndex = -1
               Exit Do
            End If
            ii = ii - 1
         End If
         ii = ii + 1
      Loop
      If textNT35 = "," Then textNT35 = ""
   End If
End Sub

Private Sub CmdClear_Click()
   If ListBox1.ListCount > 0 Then
      If MsgBox("是否要清除已輸入的代理人/申請人編號？", vbInformation + vbYesNo + vbSystemModal + vbDefaultButton2) = vbNo Then
         Exit Sub
      End If
   End If
   ListBox1.Clear
   lblCnt = ""
   textNT35 = ""
End Sub
'end 2025/07/24
