VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_25 
   BorderStyle     =   1  '單線固定
   Caption         =   "不得代理案件之客戶或代理人資料查詢"
   ClientHeight    =   5748
   ClientLeft      =   108
   ClientTop       =   936
   ClientWidth     =   9156
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleMode       =   0  '使用者自訂
   ScaleWidth      =   9047.999
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   8250
      TabIndex        =   59
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   7020
      TabIndex        =   58
      Top             =   60
      Width           =   1230
   End
   Begin VB.TextBox textNT01 
      Height          =   264
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4785
      Left            =   60
      TabIndex        =   28
      Top             =   780
      Width           =   9045
      _ExtentX        =   15939
      _ExtentY        =   8424
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_25.frx":0000
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
      Tab(0).Control(13)=   "Label3(0)"
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
      Tab(0).Control(20)=   "Label1(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(1)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "LabRCL18_2"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label3(1)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textNT08_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textNT06"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textNT05"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textNT04"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textNT03"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textNT08"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textNT21"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textNT18"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textNT17"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textNT19"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textNT22"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cmdOpenAtt"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textNT30"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "lstAtt"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textNT31"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textRCL17"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textRCL18"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textRCL20"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textRCL19"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "其他"
      TabPicture(1)   =   "frm100101_25.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label41(6)"
      Tab(1).Control(1)=   "Label41(5)"
      Tab(1).Control(2)=   "Label41(4)"
      Tab(1).Control(3)=   "Label41(3)"
      Tab(1).Control(4)=   "Label41(2)"
      Tab(1).Control(5)=   "Label41(32)"
      Tab(1).Control(6)=   "Label13"
      Tab(1).Control(7)=   "Label16"
      Tab(1).Control(8)=   "Label18"
      Tab(1).Control(9)=   "Label1(22)"
      Tab(1).Control(10)=   "lstUsers(0)"
      Tab(1).Control(11)=   "textNT09"
      Tab(1).Control(12)=   "textNT16"
      Tab(1).Control(13)=   "textNT15"
      Tab(1).Control(14)=   "textNT10"
      Tab(1).Control(15)=   "textNT14"
      Tab(1).Control(16)=   "textNT13"
      Tab(1).Control(17)=   "textNT12"
      Tab(1).Control(18)=   "textNT11"
      Tab(1).Control(19)=   "textNT23"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).ControlCount=   20
      Begin VB.TextBox textRCL19 
         Height          =   270
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   68
         Text            =   "textRCL"
         Top             =   2130
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox textRCL20 
         Height          =   270
         Left            =   8400
         MaxLength       =   2
         TabIndex        =   67
         Text            =   "CR"
         Top             =   2130
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox textRCL18 
         Height          =   270
         Left            =   6396
         MaxLength       =   8
         TabIndex        =   63
         Text            =   "textRCL"
         Top             =   1830
         Visible         =   0   'False
         Width           =   1000
      End
      Begin VB.TextBox textRCL17 
         Height          =   270
         Left            =   6060
         MaxLength       =   18
         TabIndex        =   61
         Text            =   "textRCL17"
         Top             =   1530
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.TextBox textNT31 
         Height          =   270
         Left            =   2280
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   4320
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.TextBox textNT23 
         Height          =   270
         Left            =   -74850
         TabIndex        =   57
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
         Left            =   1245
         Sorted          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3780
         Width           =   6990
      End
      Begin VB.TextBox textNT30 
         Height          =   270
         Left            =   210
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   4050
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdOpenAtt 
         Caption         =   "開啟"
         Height          =   255
         Left            =   8250
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3810
         Width           =   735
      End
      Begin VB.TextBox textNT22 
         Height          =   270
         Left            =   1245
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   3480
         Width           =   6060
      End
      Begin VB.TextBox textNT19 
         Height          =   270
         Left            =   1245
         MaxLength       =   40
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   2130
         Width           =   6060
      End
      Begin VB.TextBox textNT17 
         Height          =   270
         Left            =   3780
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1830
         Width           =   400
      End
      Begin VB.TextBox textNT18 
         Height          =   270
         Left            =   1245
         MaxLength       =   6
         TabIndex        =   10
         Top             =   1830
         Width           =   650
      End
      Begin VB.TextBox textNT11 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   20
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textNT12 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   21
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textNT13 
         Height          =   270
         Left            =   -70005
         MaxLength       =   30
         TabIndex        =   22
         Top             =   990
         Width           =   3360
      End
      Begin VB.TextBox textNT14 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   23
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textNT10 
         Height          =   270
         Left            =   -73740
         MaxLength       =   30
         TabIndex        =   19
         Top             =   690
         Width           =   3360
      End
      Begin VB.TextBox textNT15 
         Height          =   270
         Left            =   -70005
         TabIndex        =   24
         Top             =   1290
         Width           =   3360
      End
      Begin VB.TextBox textNT21 
         Height          =   270
         Left            =   1245
         MaxLength       =   7
         TabIndex        =   14
         Top             =   3180
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
      Begin VB.Label Label3 
         Caption         =   "延展次數："
         Height          =   252
         Index           =   1
         Left            =   7440
         TabIndex        =   66
         Top             =   2160
         Visible         =   0   'False
         Width           =   912
      End
      Begin MSForms.Label LabRCL18_2 
         Height          =   600
         Left            =   7440
         TabIndex        =   65
         Top             =   1860
         Visible         =   0   'False
         Width           =   1296
         Caption         =   "LabRCL18_2"
         Size            =   "2286;1058"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "要求檢查對象："
         Height          =   180
         Index           =   1
         Left            =   5160
         TabIndex        =   64
         Top             =   1860
         Visible         =   0   'False
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "身分證字號/統一編號："
         Height          =   180
         Index           =   10
         Left            =   4200
         TabIndex        =   62
         Top             =   1560
         Visible         =   0   'False
         Width           =   1848
      End
      Begin MSForms.TextBox textNT20 
         Height          =   705
         Left            =   1245
         TabIndex        =   13
         Top             =   2445
         Width           =   7116
         VariousPropertyBits=   -1463795685
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "12552;1244"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT16 
         Height          =   300
         Left            =   -73740
         TabIndex        =   25
         Top             =   1590
         Width           =   7095
         VariousPropertyBits=   675299355
         MaxLength       =   70
         Size            =   "12515;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textNT09 
         Height          =   300
         Left            =   -73740
         TabIndex        =   18
         Top             =   390
         Width           =   7095
         VariousPropertyBits=   675299355
         MaxLength       =   70
         Size            =   "12515;529"
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
         Left            =   1245
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
      Begin MSForms.ListBox lstUsers 
         Height          =   2740
         Index           =   0
         Left            =   -73365
         TabIndex        =   26
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
      Begin MSForms.Label LabNT18_2 
         Height          =   300
         Left            =   1920
         TabIndex        =   49
         Top             =   1860
         Width           =   864
         Caption         =   "LabNT18_2"
         Size            =   "1526;529"
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
         TabIndex        =   56
         Top             =   2010
         Width           =   1440
      End
      Begin VB.Label Label1 
         Caption         =   "附件："
         Height          =   180
         Index           =   7
         Left            =   90
         TabIndex        =   55
         Top             =   3816
         Width           =   912
      End
      Begin VB.Label Label4 
         Caption         =   "撤銷原因："
         Height          =   252
         Left            =   90
         TabIndex        =   53
         Top             =   3516
         Width           =   912
      End
      Begin VB.Label Label3 
         Caption         =   "原因："
         Height          =   252
         Index           =   0
         Left            =   90
         TabIndex        =   52
         Top             =   2160
         Width           =   1100
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "負責同仁："
         Height          =   180
         Index           =   6
         Left            =   90
         TabIndex        =   51
         Top             =   1860
         Width           =   912
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "部門別："
         Height          =   180
         Index           =   9
         Left            =   3036
         TabIndex        =   50
         Top             =   1860
         Width           =   720
      End
      Begin VB.Label LabNT17_2 
         AutoSize        =   -1  'True
         Caption         =   "LabNT17_2"
         Height          =   180
         Left            =   4200
         TabIndex        =   48
         Top             =   1860
         Width           =   852
      End
      Begin VB.Label Label18 
         Caption         =   "地址(中)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   47
         Top             =   420
         Width           =   1005
      End
      Begin VB.Label Label16 
         Caption         =   "地址(英)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   46
         Top             =   720
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "地址(日)："
         Height          =   255
         Left            =   -74820
         TabIndex        =   45
         Top             =   1620
         Width           =   1005
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   32
         Left            =   -73845
         TabIndex        =   44
         Top             =   720
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   2
         Left            =   -73845
         TabIndex        =   43
         Top             =   1020
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "5"
         Height          =   180
         Index           =   3
         Left            =   -73845
         TabIndex        =   42
         Top             =   1320
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   4
         Left            =   -70125
         TabIndex        =   41
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   5
         Left            =   -70125
         TabIndex        =   40
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "6"
         Height          =   180
         Index           =   6
         Left            =   -70125
         TabIndex        =   39
         Top             =   1290
         Width           =   90
      End
      Begin VB.Label Label6 
         Caption         =   "備註："
         Height          =   252
         Left            =   90
         TabIndex        =   38
         Top             =   2496
         Width           =   912
      End
      Begin VB.Label Label5 
         Caption         =   "撤銷日期："
         Height          =   252
         Left            =   90
         TabIndex        =   37
         Top             =   3216
         Width           =   912
      End
      Begin VB.Label Label2 
         Caption         =   "國籍："
         Height          =   252
         Index           =   0
         Left            =   90
         TabIndex        =   36
         Top             =   1560
         Width           =   912
      End
      Begin VB.Label Label27 
         Caption         =   "名稱(中)："
         Height          =   252
         Left            =   90
         TabIndex        =   35
         Top             =   336
         Width           =   912
      End
      Begin VB.Label Label29 
         Caption         =   "名稱(英)："
         Height          =   252
         Left            =   90
         TabIndex        =   34
         Top             =   660
         Width           =   912
      End
      Begin VB.Label Label30 
         Caption         =   "名稱(日)："
         Height          =   252
         Index           =   0
         Left            =   90
         TabIndex        =   33
         Top             =   1236
         Width           =   912
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "1"
         Height          =   180
         Index           =   10
         Left            =   1125
         TabIndex        =   32
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "2"
         Height          =   180
         Index           =   13
         Left            =   4875
         TabIndex        =   31
         Top             =   690
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "3"
         Height          =   180
         Index           =   14
         Left            =   1125
         TabIndex        =   30
         Top             =   990
         Width           =   90
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "4"
         Height          =   180
         Index           =   15
         Left            =   4875
         TabIndex        =   29
         Top             =   990
         Width           =   90
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   60
      Top             =   30
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2316
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   504
      Width           =   6072
      VariousPropertyBits=   671107103
      Size            =   "10710;529"
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
      Left            =   510
      TabIndex        =   27
      Top             =   510
      Width           =   555
   End
End
Attribute VB_Name = "frm100101_25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2022/1/24 Form2.0已修改(LabNT18_2,textNT02,textNT07,textNT09,textNT16,textNT20,lstUsers(0))
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2012/3/21
Option Explicit

Public cmdState As Integer
Private Const cTableName As String = "NOTAGENT" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱
Dim i As Integer, IsRiskCheck As Boolean 'Add by Amy 2023/12/12

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 0
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
   End Select
End Sub

Sub StrMenu()
   Dim strKey  As String, strKey1 As String
   Dim adoRst As New ADODB.Recordset
   
   'Modify by Amy 2023/12/12 +風險檢查對象
   ClearField
   IsRiskCheck = False
   'Modify by Amy 2024/12/31 +Me.Caption
   If Len(Me.Tag) = 5 Then
      Me.Caption = "風險檢查資料查詢"
      IsRiskCheck = True
      strKey = Me.Tag
      strExc(0) = "SELECT * FROM RiskCheckList WHERE RCL01='" & strKey & "' "
   Else
      Me.Caption = "不得代理案件之客戶或代理人資料查詢"
      strKey = Right("000" & Me.Tag, 3)
      strExc(0) = "SELECT * FROM NotAgent WHERE NT01='" & strKey & "' "
   End If
   
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If IsRiskCheck = False Then
         Call UpdateCtrlData(strKey)
      Else
         Call SetRishCheckData(adoRst)
      End If
   'end 2023/12/12
   Else
      If IsRiskCheck = True Then Call ShowField
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
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
      If strSrvDate(1) >= CR_NewDate And textNT31.Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(textNT31.Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, " (") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
      'Else
      ''end 2017/08/09
      '    PUB_OpenFtpFile textNT01, lstAtt.Text, Winsock1, 3
      'end 2024/8/2
      
      End If 'end 2017/08/09
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   cmdState = -1
   textNT08_2.BackColor = &H8000000F
   textCUID.BackColor = &H8000000F
   SSTab1.Tab = 0
   SetCtrlReadOnly True
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
   lstUsers(0).Clear
   lstAtt.Clear
   textCUID = ""
   'Add by Amy 2023/12/12 +風險檢查對象
   textRCL17 = Empty '身份證/統編
   textRCL18 = Empty '要求檢查對象
   LabRCL18_2.Caption = Empty: LabRCL18_2 = Empty
   textRCL19 = Empty '下次提醒日
   textRCL20 = Empty '延展次數
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
   
'   cmdOpenAtt.Enabled = bEnable
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub UpdateCtrlData(strKey As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM NOTAGENT " & _
            "WHERE NT01 = '" & strKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   'ClearField 'Mark by Amy 2023/12/12 搬至外層
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
      'Added by Lydia 2023/12/28
      If "" & rsTmp.Fields("NT25") >= 新部門啟用日 And textNT18 <> "" Then
         LabNT17_2 = GetDeptNameA0922(textNT18)
      End If
      'end 2023/12/28
      
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
      
      If InStr(textNT23, strUserNum) > 0 Or Pub_StrUserSt03 = "M51" Then
         cmdOpenAtt.Enabled = True
      Else
         cmdOpenAtt.Enabled = False
      End If
      
      SetlstUsers 0, textNT23
      SetList lstAtt, textNT30
      
      ' 更新CUID
      UpdateCUID rsTmp
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
      
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String, strCDate As String, strCTime As String
   Dim strUName As String, strUDate As String, strUTime As String
   
   If IsRiskCheck = True Then
      For i = 26 To 32
         strTemp = "" & rsSrcTmp.Fields("RCL" & Format(i, "00"))
         If strTemp <> MsgText(601) Then
            Select Case i
               Case 27, 30
                  strTemp = GetStaffName(strTemp, True)
                  If i = 27 Then
                     strCName = strTemp
                  Else
                     strUName = strTemp
                  End If
               Case 28, 31
                  strTemp = Format(TAIWANDATE(strTemp), "###/##/##")
                  If i = 28 Then
                     strCDate = strTemp
                  Else
                     strUDate = strTemp
                  End If
               Case 29, 32
                  strTemp = Format(strTemp, "0#:##") 'Modify by Amy 2024/12/31 原:##:## ex:00005-RCL32,只顯示:8
                  If i = 29 Then
                     strCTime = strTemp
                  Else
                     strUTime = strTemp
                  End If
            End Select
         End If
      Next i
   Else
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
   End If
   
   ' 設定CUID中的文字
   textCUID = "CREATE : " & strCName & " " & _
              " : " & strCDate & " " & _
              " : " & strCTime & String(6, " ")
   If strUName <> MsgText(601) Then
      textCUID = textCUID & "UPDATE : " & strUName & " " & _
              " : " & strUDate & " " & _
              " : " & strUTime
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

   PUB_KillTempFile "$$*.*" 'Added by Lydia 2017/08/09 清除暫存檔
   
   Set frm100101_25 = Nothing
End Sub

Private Sub textNT01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'modify by sonia 2022/1/24
'Private Sub textNT02_KeyPress(KeyAscii As Integer)
Private Sub textNT02_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'modify by sonia 2022/1/24
'Private Sub textNT07_KeyPress(KeyAscii As Integer)
Private Sub textNT07_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

Private Sub textNT08_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

'modify by sonia 2022/1/24
'Private Sub textNT09_KeyPress(KeyAscii As Integer)
Private Sub textNT09_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = ChangeZIP(KeyAscii)
End Sub

'日文地址要轉全形
'modify by sonia 2022/1/24
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

Private Sub textNT21_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

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

Private Sub SetlstUsers(p_idx As Integer, p_stNums As String)
   Dim arrID
   
   lstUsers(p_idx).Clear
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
                  lstUsers(p_idx).AddItem "" & .Fields(1), 0
                  'modify by sonia 2022/1/24 .Form2.0不能用.ITEMDATA
                  'lstUsers(p_idx).ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  lstUsers(p_idx).Tag = .Fields(0) & "," & lstUsers(p_idx).Tag
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
   Else
      MsgBox "無文件查詢權限!"
   End If
End Sub

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

'Add by Amy 2023/12/12
'風險檢查對象
Private Sub SetRishCheckData(ado As ADODB.Recordset)
   Call ShowField '設定顯示欄位
On Error GoTo ErrHand
   'Memo **代表 [不得代理資料表] 無此欄
   If IsNull(ado.Fields("RCL01")) = False Then: textNT01 = ado.Fields("RCL01") 'Key
   If IsNull(ado.Fields("RCL02")) = False Then: textNT02 = ado.Fields("RCL02") '中文名
   If IsNull(ado.Fields("RCL03")) = False Then: textNT03 = ado.Fields("RCL03") '英文名
   If IsNull(ado.Fields("RCL04")) = False Then: textNT04 = ado.Fields("RCL04")
   If IsNull(ado.Fields("RCL05")) = False Then: textNT05 = ado.Fields("RCL05")
   If IsNull(ado.Fields("RCL06")) = False Then: textNT06 = ado.Fields("RCL06")
   If IsNull(ado.Fields("RCL07")) = False Then: textNT07 = ado.Fields("RCL07") '日文名
   If IsNull(ado.Fields("RCL08")) = False Then: textNT08 = ado.Fields("RCL08"): textNT08_2 = GetNationName(textNT08, 0) '國籍
   If IsNull(ado.Fields("RCL09")) = False Then: textNT09 = ado.Fields("RCL09") '中文地址
   If IsNull(ado.Fields("RCL10")) = False Then: textNT10 = ado.Fields("RCL10") '英文地址
   If IsNull(ado.Fields("RCL11")) = False Then: textNT11 = ado.Fields("RCL11")
   If IsNull(ado.Fields("RCL12")) = False Then: textNT12 = ado.Fields("RCL12")
   If IsNull(ado.Fields("RCL13")) = False Then: textNT13 = ado.Fields("RCL13")
   If IsNull(ado.Fields("RCL14")) = False Then: textNT14 = ado.Fields("RCL14")
   If IsNull(ado.Fields("RCL15")) = False Then: textNT15 = ado.Fields("RCL15")
   If IsNull(ado.Fields("RCL16")) = False Then: textNT16 = ado.Fields("RCL16") '日文地址
   '身份證字號/統編 **
   If IsNull(ado.Fields("RCL17")) = False Then: textRCL17 = ado.Fields("RCL17")
   '要求檢查對象 **
   If IsNull(ado.Fields("RCL18")) = False Then: textRCL18 = ado.Fields("RCL18"): LabRCL18_2 = GetCustomerName(textRCL18 & "0")
   '下次提醒日 ** (不得代理-NT19-不得代理原因,不顯示)
   If IsNull(ado.Fields("RCL19")) = False Then: textRCL19 = TAIWANDATE(ado.Fields("RCL19")) 'Modify by Amy 2024/112/31 +TAIWANDATE
   '延展次數 **
   If IsNull(ado.Fields("RCL20")) = False Then: textRCL20 = ado.Fields("RCL20")
   '部門別 (不得代理-NT17)
   If IsNull(ado.Fields("RCL21")) = False Then: textNT17 = ado.Fields("RCL21"): LabNT17_2 = GetDepartmentName(textNT17)
   '負責同仁 (不得代理-NT18)
   If IsNull(ado.Fields("RCL22")) = False Then: textNT18 = ado.Fields("RCL22"): LabNT18_2 = GetPrjSalesNM(textNT18)
   'Added by Lydia 2023/12/28
   If "" & ado.Fields("RCL27") >= 新部門啟用日 And textNT18 <> "" Then
      LabNT17_2 = GetDeptNameA0922(textNT18)
   End If
   'end 2023/12/28
   
   '備註 (不得代理-NT20)
   If IsNull(ado.Fields("RCL23")) = False Then: textNT20 = ado.Fields("RCL23")
   '撤銷日期 (不得代理-NT21)
   If IsNull(ado.Fields("RCL24")) = False Then textNT21 = TAIWANDATE(ado.Fields("RCL24"))
   '撤銷原因 (不得代理-NT22)
   If IsNull(ado.Fields("RCL25")) = False Then: textNT22 = ado.Fields("RCL25")
   
   ' 更新CUID
   UpdateCUID ado
   ado.Close
      
ErrHand:
   Set ado = Nothing
End Sub

'顯示 風險檢查對象 欄位
Private Sub ShowField()
   textNT01.MaxLength = 5
   Label1(10).Visible = True: textRCL17.Visible = True '身份證/統編
   Label1(1).Visible = True: textRCL18.Visible = True: LabRCL18_2.Visible = True '要求檢查對象
   Label3(0).Caption = "下次提醒日：": textRCL19.Visible = True: textRCL19.Left = 1245
   Label3(1).Visible = True: Label3(1).Left = 3036: textRCL20.Visible = True: textRCL20.Left = 4000 '延展次數
   '不顯示欄位
   Label1(7).Visible = False: lstAtt.Visible = False: cmdOpenAtt.Visible = False '附件
   textNT19.Visible = False '不得代理原因
   Label1(22).Visible = False: lstUsers(0).Visible = False '其他頁籤-文件可查詢人員
End Sub

