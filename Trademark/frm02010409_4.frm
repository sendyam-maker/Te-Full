VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010409_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務結果輸入(網域變更, 著作權變更)"
   ClientHeight    =   5720
   ClientLeft      =   240
   ClientTop       =   960
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   9150
   Begin VB.TextBox textCP10_2 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1980
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   97
      TabStop         =   0   'False
      Top             =   1620
      Width           =   1812
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3432
      Left            =   120
      TabIndex        =   42
      Top             =   1920
      Width           =   8892
      _ExtentX        =   15699
      _ExtentY        =   6068
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm02010409_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label17"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label16"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label15"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label14"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label20"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label24"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label10"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label28"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label27"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCE10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCE04"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCE55"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCE63"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCE17"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCE23"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCE16"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCE02"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCE03"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCE09"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCE51"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCE52"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCE53"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCE54"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCE56"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCE65"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCE22"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCE38"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCE57"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCE58"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).ControlCount=   40
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm02010409_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCE66"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textCE67"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textCE62"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "textCE42"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textCE46"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textCE45"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textCE48"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textCE47"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "textCE50"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textCE49"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textCE40"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "textCE39"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textCE60"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textCE44"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCE59"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCE61"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textCE43"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCE41"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label25"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label11"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label43"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label44"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label35"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "Label36"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "Label37"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label38"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "Label39"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "Label40"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label41"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "Label42"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "Label29"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "Label30"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "Label31"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "Label32"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "Label33"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "Label34"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).ControlCount=   36
      Begin VB.TextBox textCE66 
         Height          =   264
         Left            =   -71760
         MaxLength       =   25
         TabIndex        =   18
         Top             =   3060
         Width           =   5472
      End
      Begin VB.TextBox textCE67 
         Height          =   264
         Left            =   -73680
         TabIndex        =   17
         Top             =   3060
         Width           =   372
      End
      Begin VB.TextBox textCE62 
         Height          =   264
         Left            =   -73680
         TabIndex        =   16
         Top             =   2760
         Width           =   372
      End
      Begin VB.TextBox textCE42 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5472
      End
      Begin VB.TextBox textCE46 
         Height          =   264
         Left            =   -73680
         TabIndex        =   13
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCE45 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   1860
         Width           =   5472
      End
      Begin VB.TextBox textCE48 
         Height          =   264
         Left            =   -73680
         TabIndex        =   14
         Top             =   2160
         Width           =   372
      End
      Begin VB.TextBox textCE47 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5472
      End
      Begin VB.TextBox textCE50 
         Height          =   264
         Left            =   -73680
         TabIndex        =   15
         Top             =   2460
         Width           =   372
      End
      Begin VB.TextBox textCE49 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   2460
         Width           =   5472
      End
      Begin VB.TextBox textCE40 
         Height          =   264
         Left            =   -73680
         TabIndex        =   10
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textCE39 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   360
         Width           =   5472
      End
      Begin VB.TextBox textCE60 
         Height          =   264
         Left            =   -73680
         TabIndex        =   11
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox textCE44 
         Height          =   264
         Left            =   -73680
         TabIndex        =   12
         Top             =   960
         Width           =   372
      End
      Begin VB.TextBox textCE59 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -71760
         Locked          =   -1  'True
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   660
         Width           =   5472
      End
      Begin VB.TextBox textCE58 
         Height          =   264
         Left            =   1320
         TabIndex        =   9
         Top             =   3060
         Width           =   372
      End
      Begin VB.TextBox textCE57 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3060
         Width           =   5352
      End
      Begin VB.TextBox textCE38 
         Height          =   264
         Left            =   1320
         TabIndex        =   8
         Top             =   2760
         Width           =   372
      End
      Begin VB.TextBox textCE22 
         Height          =   264
         Left            =   1320
         TabIndex        =   6
         Top             =   2160
         Width           =   372
      End
      Begin VB.TextBox textCE65 
         Height          =   264
         Left            =   1320
         TabIndex        =   7
         Top             =   2460
         Width           =   372
      End
      Begin VB.TextBox textCE56 
         Height          =   264
         Left            =   1320
         TabIndex        =   3
         Top             =   1260
         Width           =   372
      End
      Begin VB.TextBox textCE54 
         Height          =   264
         Left            =   1320
         TabIndex        =   4
         Top             =   1560
         Width           =   372
      End
      Begin VB.TextBox textCE53 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5352
      End
      Begin VB.TextBox textCE52 
         Height          =   264
         Left            =   1320
         TabIndex        =   5
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCE51 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1860
         Width           =   5352
      End
      Begin VB.TextBox textCE09 
         Height          =   264
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox textCE03 
         Height          =   264
         Left            =   1320
         TabIndex        =   1
         Top             =   660
         Width           =   372
      End
      Begin VB.TextBox textCE02 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   44
         TabStop         =   0   'False
         Top             =   660
         Width           =   5352
      End
      Begin VB.TextBox textCE16 
         Height          =   264
         Left            =   1320
         TabIndex        =   2
         Top             =   960
         Width           =   372
      End
      Begin MSForms.TextBox textCE61 
         Height          =   300
         Left            =   -71760
         TabIndex        =   24
         Top             =   2760
         Width           =   5472
         VariousPropertyBits=   679493661
         Size            =   "9652;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE43 
         Height          =   264
         Left            =   -71760
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1560
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE41 
         Height          =   264
         Left            =   -71760
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   960
         Width           =   5472
         VariousPropertyBits=   679493663
         Size            =   "9652;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE23 
         Height          =   264
         Left            =   3360
         TabIndex        =   67
         TabStop         =   0   'False
         Top             =   2760
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE17 
         Height          =   264
         Left            =   3360
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2160
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE63 
         Height          =   264
         Left            =   3360
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2460
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE55 
         Height          =   264
         Left            =   3360
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1260
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE04 
         Height          =   264
         Left            =   3360
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCE10 
         Height          =   264
         Left            =   3360
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   960
         Width           =   5352
         VariousPropertyBits=   679493663
         Size            =   "9440;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label25 
         Caption         =   "密碼 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   99
         Top             =   3060
         Width           =   612
      End
      Begin VB.Label Label11 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   98
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label43 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   96
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label Label44 
         Caption         =   "其它 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   95
         Top             =   2760
         Width           =   612
      End
      Begin VB.Label Label35 
         Caption         =   "案件名稱(英) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   94
         Top             =   1260
         Width           =   1212
      End
      Begin VB.Label Label36 
         Caption         =   "案件名稱(日) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   93
         Top             =   1560
         Width           =   1212
      End
      Begin VB.Label Label37 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   92
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label38 
         Caption         =   "縮減商品 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   91
         Top             =   1860
         Width           =   1212
      End
      Begin VB.Label Label39 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   90
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label40 
         Caption         =   "商品類別 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   89
         Top             =   2160
         Width           =   1212
      End
      Begin VB.Label Label41 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   88
         Top             =   2460
         Width           =   1092
      End
      Begin VB.Label Label42 
         Caption         =   "商品群組 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   87
         Top             =   2460
         Width           =   1212
      End
      Begin VB.Label Label29 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   81
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label30 
         Caption         =   "商標種類 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   80
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label31 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   79
         Top             =   660
         Width           =   1092
      End
      Begin VB.Label Label32 
         Caption         =   "圖樣 :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   78
         Top             =   660
         Width           =   612
      End
      Begin VB.Label Label33 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   77
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label34 
         Caption         =   "案件名稱(中) :"
         Height          =   252
         Left            =   -73080
         TabIndex        =   76
         Top             =   960
         Width           =   1212
      End
      Begin VB.Label Label27 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   72
         Top             =   3060
         Width           =   1092
      End
      Begin VB.Label Label28 
         Caption         =   "正商標號數 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   71
         Top             =   3060
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   69
         Top             =   2760
         Width           =   1092
      End
      Begin VB.Label Label24 
         Caption         =   "申請地址(中) :"
         Height          =   252
         Left            =   1920
         TabIndex        =   68
         Top             =   2760
         Width           =   1212
      End
      Begin VB.Label Label18 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   66
         Top             =   2160
         Width           =   1092
      End
      Begin VB.Label Label19 
         Caption         =   "申請人中譯文 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   65
         Top             =   2160
         Width           =   1332
      End
      Begin VB.Label Label20 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   64
         Top             =   2460
         Width           =   1092
      End
      Begin VB.Label Label9 
         Caption         =   "代表人中譯文 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   63
         Top             =   2460
         Width           =   1332
      End
      Begin VB.Label Label12 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   60
         Top             =   1260
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   59
         Top             =   1260
         Width           =   732
      End
      Begin VB.Label Label14 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   58
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label15 
         Caption         =   "代表人印鑑 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   57
         Top             =   1560
         Width           =   1092
      End
      Begin VB.Label Label16 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   56
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label17 
         Caption         =   "申請人印鑑 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   55
         Top             =   1860
         Width           =   1092
      End
      Begin VB.Label Label1 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Index           =   2
         Left            =   120
         TabIndex        =   51
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label8 
         Caption         =   "申請人 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   50
         Top             =   360
         Width           =   732
      End
      Begin VB.Label Label7 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   49
         Top             =   660
         Width           =   1092
      End
      Begin VB.Label Label4 
         Caption         =   "申請日 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   48
         Top             =   660
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "准(1) / 駁(2) :"
         Height          =   252
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label2 
         Caption         =   "代表人 :"
         Height          =   252
         Left            =   1920
         TabIndex        =   46
         Top             =   960
         Width           =   1092
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7044
      TabIndex        =   22
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6216
      TabIndex        =   21
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8268
      TabIndex        =   23
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textPrint 
      Height          =   264
      Left            =   900
      TabIndex        =   19
      Top             =   5370
      Width           =   315
   End
   Begin VB.TextBox textSPKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   420
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1620
      Width           =   612
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5700
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1260
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2532
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1260
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   690
      Width           =   7752
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13674;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5700
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1620
      Width           =   2532
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textSP08 
      Height          =   264
      Left            =   1260
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7752
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13674;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP64 
      Height          =   300
      Left            =   4920
      TabIndex        =   20
      Top             =   5370
      Width           =   4095
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "7223;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   240
      TabIndex        =   41
      Top             =   690
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4740
      TabIndex        =   39
      Top             =   1620
      Width           =   852
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
      Height          =   180
      Left            =   1260
      TabIndex        =   38
      Top             =   5400
      Width           =   2745
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "列印定稿 :"
      Height          =   180
      Left            =   60
      TabIndex        =   37
      Top             =   5400
      Width           =   810
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   240
      TabIndex        =   36
      Top             =   1020
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   240
      TabIndex        =   35
      Top             =   420
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   240
      TabIndex        =   34
      Top             =   1620
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   252
      Index           =   3
      Left            =   4740
      TabIndex        =   33
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   240
      TabIndex        =   32
      Top             =   1320
      Width           =   732
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "進度備註 :"
      Height          =   180
      Left            =   4080
      TabIndex        =   31
      Top             =   5400
      Width           =   810
   End
End
Attribute VB_Name = "frm02010409_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/29 Form2.0已修改 cmbTM05/textSP08/textCP13/textCE04/textCE10/textCE55/textCE17/textCE63/textCE23/textCE41/textCE43/textCE61/textCP64
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 本所案號
Dim m_SP01 As String
Dim m_SP02 As String
Dim m_SP03 As String
Dim m_SP04 As String
' 申請國家
Dim m_SP09 As String
' 來函收文日
Dim m_CP05 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
' 申請日
Dim m_CE02 As String
' 申請人
Dim m_CE04 As String
' 商品種類代碼
Dim m_CE39 As String
' 案件名稱
Dim m_SP06 As String
' 申請人
Dim m_TM23 As String
' 案件進度資料
Dim m_CP64 As String
'原承辦人  2015/1/14 add by sonia
Dim m_CP14 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 變更事項檔的欄位串列
Dim m_CEList() As FIELDITEM
Dim m_CECount As Integer
' 針對商標基本檔欄位所用的暫存陣列
Dim m_SPList() As FIELDITEM
Dim m_SPCount As Integer
'add by nickc 2006/11/21
Dim m_textPrint As String
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2019/12/19 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/19 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub ClearField()
   ClearCEFields
   ClearSPFields
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   ClearField
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010409_4 = Nothing
End Sub

Private Sub cmdCancel_Click()
   frm02010409_2.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010409_2
   Unload frm02010409_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
             'add by nickc 2005/04/22
          Pub_EndModCashMsg m_SP09
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      OnUpdateField
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010409_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010409_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
      '2019/5/22 END
      Else
         frm02010409_1.Show
      End If
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textSPKey.BackColor = &H8000000F
   textSP08.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP10_2.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   
   textCE04.BackColor = &H8000000F
   textCE02.BackColor = &H8000000F
   textCE10.BackColor = &H8000000F
   textCE55.BackColor = &H8000000F
   textCE53.BackColor = &H8000000F
   textCE51.BackColor = &H8000000F
   textCE17.BackColor = &H8000000F
   textCE63.BackColor = &H8000000F
   textCE23.BackColor = &H8000000F
   textCE57.BackColor = &H8000000F
   textCE39.BackColor = &H8000000F
   textCE59.BackColor = &H8000000F
   textCE41.BackColor = &H8000000F
   textCE42.BackColor = &H8000000F
   textCE43.BackColor = &H8000000F
   textCE45.BackColor = &H8000000F
   textCE47.BackColor = &H8000000F
   textCE49.BackColor = &H8000000F
   textCE61.BackColor = &H8000000F
   
   textCE10.MaxLength = Pub_MaxCEL10  'Added by Lydia 2016/09/10 設定代表人中文名稱長度
    
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010409_1.m_strIR01
   m_strIR02 = frm02010409_1.m_strIR02
   m_strIR03 = frm02010409_1.m_strIR03
   m_strIR04 = frm02010409_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_SP01 = Empty
      m_SP02 = Empty
      m_SP03 = Empty
      m_SP04 = Empty
      m_CP05 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_SP01 = strData
      ' 本所案號 欄位2
      Case 1: m_SP02 = strData
      ' 本所案號 欄位3
      Case 2: m_SP03 = strData
      ' 本所案號 欄位4
      Case 3: m_SP04 = strData
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
   End Select
End Sub

' 檢查該欄位是否存在
Private Function IsCEFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsCEFieldExist = False
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         IsCEFieldExist = True
         Exit For
      End If
   Next nIndex
End Function

' 新增一個欄位
Private Sub AddCEField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   If IsCEFieldExist(strField) = True Then
      GoTo EXITSUB
   End If
   ReDim Preserve m_CEList(m_CECount + 1)
   m_CEList(m_CECount).fiName = strField
   m_CEList(m_CECount).fiOldData = strOldData
   m_CEList(m_CECount).fiNewData = strOldData
   m_CEList(m_CECount).fiType = nType
   m_CECount = m_CECount + 1
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetCEFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         m_CEList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 取得欄位原值
Private Function GetCEFieldOldData(ByVal strField As String) As String
   Dim nIndex As Integer
   GetCEFieldOldData = Empty
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         GetCEFieldOldData = m_CEList(nIndex).fiOldData
         Exit For
      End If
   Next nIndex
End Function
' 取得欄位原值
Private Function GetCEFieldNewData(ByVal strField As String) As String
   Dim nIndex As Integer
   GetCEFieldNewData = Empty
   For nIndex = 0 To m_CECount - 1
      If m_CEList(nIndex).fiName = strField Then
         GetCEFieldNewData = m_CEList(nIndex).fiNewData
         Exit For
      End If
   Next nIndex
End Function
' 清除欄位串列
Private Sub ClearCEFields()
   Erase m_CEList
   m_CECount = 0
End Sub
' 檢查該商標基本檔的欄位是否存在
Private Function IsSPFieldExist(ByVal strField As String) As Boolean
   Dim nIndex As Integer
   IsSPFieldExist = False
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         IsSPFieldExist = True
         Exit For
      End If
   Next nIndex
End Function
' 新增一個欄位
Private Sub AddSPField(ByVal strField As String, ByVal strOldData As String, ByVal nType As Integer)
   Dim bFind As Boolean
   Dim nIndex As Integer
   bFind = False
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         bFind = True
         m_SPList(m_SPCount).fiOldData = strOldData
         m_SPList(m_SPCount).fiNewData = strOldData
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_SPList(m_SPCount + 1)
      m_SPList(m_SPCount).fiName = strField
      m_SPList(m_SPCount).fiOldData = strOldData
      m_SPList(m_SPCount).fiNewData = strOldData
      m_SPList(m_SPCount).fiType = nType
      m_SPCount = m_SPCount + 1
   End If
EXITSUB:
End Sub
' 設定欄位新值
Private Sub SetSPFieldNewData(ByVal strField As String, ByVal strNewData As String)
   Dim nIndex As Integer
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         m_SPList(nIndex).fiNewData = strNewData
         Exit For
      End If
   Next nIndex
End Sub
' 設定欄位新值
Private Function GetSPFieldOldData(ByVal strField As String) As String
   Dim nIndex As Integer
   GetSPFieldOldData = Empty
   For nIndex = 0 To m_SPCount - 1
      If m_SPList(nIndex).fiName = strField Then
         GetSPFieldOldData = m_SPList(nIndex).fiOldData
         Exit For
      End If
   Next nIndex
End Function

' 清除欄位串列
Private Sub ClearSPFields()
   Erase m_SPList
   m_SPCount = 0
End Sub

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTemp As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_SP01 & "' AND " & _
                  "SP02 = '" & m_SP02 & "' AND " & _
                  "SP03 = '" & m_SP03 & "' AND " & _
                  "SP04 = '" & m_SP04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/17
      m_SP09 = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_SP09 = rsTmp.Fields("SP09")
      End If
      ' 案件名稱(中)
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
         strTemp = rsTmp.Fields("SP05")
      End If
      AddSPField "SP05", strTemp, 0
      ' 案件名稱(英)
      strTemp = Empty
      m_SP06 = Empty
      If IsNull(rsTmp.Fields("SP06")) = False Then
         m_SP06 = rsTmp.Fields("SP06")
         strTemp = rsTmp.Fields("SP06")
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      AddSPField "SP06", strTemp, 0
      ' 案件名稱(日)
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP07")) = False Then
         strTemp = rsTmp.Fields("SP07")
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      AddSPField "SP07", strTemp, 0
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textSP08 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      AddSPField "SP08", m_TM23, 0
      
      'Add By Sindy 2019/12/19
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/19 END
      
      ' 申請日
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP10")) = False Then
         strTemp = rsTmp.Fields("SP10")
      End If
      AddSPField "SP10", strTemp, 0
      ' 代表人
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP42")) = False Then
         strTemp = rsTmp.Fields("SP42")
      End If
      AddSPField "SP42", strTemp, 0
      ' 網域密碼
      strTemp = Empty
      If IsNull(rsTmp.Fields("SP49")) = False Then
         strTemp = rsTmp.Fields("SP49")
      End If
      AddSPField "SP49", strTemp, 0
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      m_textPrint = textPrint
      AddSPField "SP72", textPrint, 0
   End If

   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取變更事項檔
Private Sub QueryChangeEvent()
   Dim strTemp As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' 清除欄位串列
   ClearCEFields
   
   ' 清除暫存變數
   m_CE02 = Empty
   m_CE04 = Empty
      
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請日
      strTemp = Empty
      If IsNull(rsTmp.Fields("CE02")) = False Then
         m_CE02 = rsTmp.Fields("CE02")
         strTemp = rsTmp.Fields("CE02")
         textCE02 = TAIWANDATE(rsTmp.Fields("CE02"))
      End If
      AddCEField "CE02", strTemp, 0
      If IsNull(rsTmp.Fields("CE03")) = False Then
         textCE03 = rsTmp.Fields("CE03")
      End If
      AddCEField "CE03", textCE03, 0
      ' 申請人
      strTemp = Empty
      If IsNull(rsTmp.Fields("CE04")) = False Then
         m_CE04 = rsTmp.Fields("CE04")
         strTemp = rsTmp.Fields("CE04")
         textCE04 = GetCustomerName(rsTmp.Fields("CE04"), 0)
      End If
      AddCEField "CE04", strTemp, 0
      If IsNull(rsTmp.Fields("CE09")) = False Then
         textCE09 = rsTmp.Fields("CE09")
      End If
      AddCEField "CE09", textCE09, 0
      ' 代表人
      strTemp = Empty
      If IsNull(rsTmp.Fields("CE10")) = False Then
         textCE10 = rsTmp.Fields("CE10")
         strTemp = rsTmp.Fields("CE10")
      End If
      AddCEField "CE10", strTemp, 0
      If IsNull(rsTmp.Fields("CE16")) = False Then
         textCE16 = rsTmp.Fields("CE16")
      End If
      AddCEField "CE16", textCE16, 0
      ' 申請人中譯文
      If IsNull(rsTmp.Fields("CE17")) = False Then
         textCE17 = rsTmp.Fields("CE17")
      End If
      If IsNull(rsTmp.Fields("CE22")) = False Then
         textCE22 = rsTmp.Fields("CE22")
      End If
      AddCEField "CE22", textCE22, 0
      ' 申請地址
      If IsNull(rsTmp.Fields("CE23")) = False Then
         textCE23 = rsTmp.Fields("CE23")
      End If
      If IsNull(rsTmp.Fields("CE38")) = False Then
         textCE38 = rsTmp.Fields("CE38")
      End If
      AddCEField "CE38", textCE38, 0
      ' 專利商標種類代號
      m_CE39 = Empty
      If IsNull(rsTmp.Fields("CE39")) = False Then
         m_CE39 = rsTmp.Fields("CE39")
         textCE39 = rsTmp.Fields("CE39")
      End If
      If IsNull(rsTmp.Fields("CE40")) = False Then
         textCE40 = rsTmp.Fields("CE40")
      End If
      AddCEField "CE40", textCE40, 0
      ' 案件名稱
      If IsNull(rsTmp.Fields("CE41")) = False Then
         textCE41 = rsTmp.Fields("CE41")
      End If
      AddCEField "CE41", textCE41, 0
      If IsNull(rsTmp.Fields("CE42")) = False Then
         textCE42 = rsTmp.Fields("CE42")
      End If
      AddCEField "CE42", textCE42, 0
      If IsNull(rsTmp.Fields("CE43")) = False Then
         textCE43 = rsTmp.Fields("CE43")
      End If
      AddCEField "CE43", textCE43, 0
      If IsNull(rsTmp.Fields("CE44")) = False Then
         textCE44 = rsTmp.Fields("CE44")
      End If
      AddCEField "CE44", textCE44, 0
      ' 縮減商品
      If IsNull(rsTmp.Fields("CE45")) = False Then
         textCE45 = rsTmp.Fields("CE45")
      End If
      If IsNull(rsTmp.Fields("CE46")) = False Then
         textCE46 = rsTmp.Fields("CE46")
      End If
      AddCEField "CE46", textCE46, 0
      ' 商品類別
      If IsNull(rsTmp.Fields("CE47")) = False Then
         textCE47 = rsTmp.Fields("CE47")
      End If
      If IsNull(rsTmp.Fields("CE48")) = False Then
         textCE48 = rsTmp.Fields("CE48")
      End If
      AddCEField "CE48", textCE48, 0
      ' 商品群組
      If IsNull(rsTmp.Fields("CE49")) = False Then
         textCE49 = rsTmp.Fields("CE49")
      End If
      If IsNull(rsTmp.Fields("CE50")) = False Then
         textCE50 = rsTmp.Fields("CE50")
      End If
      AddCEField "CE50", textCE50, 0
      ' 申請人印鑑
      If IsNull(rsTmp.Fields("CE51")) = False Then
         textCE51 = rsTmp.Fields("CE51")
      End If
      If IsNull(rsTmp.Fields("CE52")) = False Then
         textCE52 = rsTmp.Fields("CE52")
      End If
      AddCEField "CE52", textCE52, 0
      ' 代表人印鑑
      If IsNull(rsTmp.Fields("CE53")) = False Then
         textCE53 = rsTmp.Fields("CE53")
      End If
      If IsNull(rsTmp.Fields("CE54")) = False Then
         textCE54 = rsTmp.Fields("CE54")
      End If
      AddCEField "CE54", textCE54, 0
      ' 代理人
      If IsNull(rsTmp.Fields("CE55")) = False Then
         textCE55 = rsTmp.Fields("CE55")
      End If
      If IsNull(rsTmp.Fields("CE56")) = False Then
         textCE56 = rsTmp.Fields("CE56")
      End If
      AddCEField "CE56", textCE56, 0
      ' 正商標號數
      If IsNull(rsTmp.Fields("CE57")) = False Then
         textCE57 = rsTmp.Fields("CE57")
      End If
      If IsNull(rsTmp.Fields("CE58")) = False Then
         textCE58 = rsTmp.Fields("CE58")
      End If
      AddCEField "CE58", textCE58, 0
      ' 圖樣
      If IsNull(rsTmp.Fields("CE59")) = False Then
         textCE59 = rsTmp.Fields("CE59")
      End If
      If IsNull(rsTmp.Fields("CE60")) = False Then
         textCE60 = rsTmp.Fields("CE60")
      End If
      AddCEField "CE60", textCE60, 0
      ' 其它
      If IsNull(rsTmp.Fields("CE61")) = False Then
         textCE61 = rsTmp.Fields("CE61")
      End If
      If IsNull(rsTmp.Fields("CE62")) = False Then
         textCE62 = rsTmp.Fields("CE62")
      End If
      AddCEField "CE62", textCE62, 0
      ' 代表人譯文
      If IsNull(rsTmp.Fields("CE63")) = False Then
         textCE63 = rsTmp.Fields("CE63")
      End If
      If IsNull(rsTmp.Fields("CE65")) = False Then
         textCE65 = rsTmp.Fields("CE65")
      End If
      AddCEField "CE65", textCE65, 0
      ' 密碼
      If IsNull(rsTmp.Fields("CE66")) = False Then
         textCE66 = rsTmp.Fields("CE66")
      End If
      AddCEField "CE66", textCE66, 0
      If IsNull(rsTmp.Fields("CE67")) = False Then
         textCE67 = rsTmp.Fields("CE67")
      End If
      AddCEField "CE67", textCE67, 0
      
      OnUpdateCtrlState rsTmp
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

Private Sub OnUpdateCtrlState(ByRef rsTmp As ADODB.Recordset)
   '
   EnableTextBox textCE09, False
   If IsNull(rsTmp.Fields("CE04")) = False Then
      If IsEmptyText(rsTmp.Fields("CE04")) = False Then
         EnableTextBox textCE09, True
      End If
   End If
   '
   EnableTextBox textCE03, False
   If IsNull(rsTmp.Fields("CE02")) = False Then
      If IsEmptyText(rsTmp.Fields("CE02")) = False Then
         EnableTextBox textCE03, True
      End If
   End If
   '
   EnableTextBox textCE16, False
   If IsNull(rsTmp.Fields("CE10")) = False Then
      If IsEmptyText(rsTmp.Fields("CE10")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE11")) = False Then
      If IsEmptyText(rsTmp.Fields("CE11")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE12")) = False Then
      If IsEmptyText(rsTmp.Fields("CE12")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE13")) = False Then
      If IsEmptyText(rsTmp.Fields("CE13")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE14")) = False Then
      If IsEmptyText(rsTmp.Fields("CE14")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE15")) = False Then
      If IsEmptyText(rsTmp.Fields("CE15")) = False Then
         EnableTextBox textCE16, True
      End If
   End If
   '
   EnableTextBox textCE56, False
   If IsNull(rsTmp.Fields("CE55")) = False Then
      If IsEmptyText(rsTmp.Fields("CE55")) = False Then
         EnableTextBox textCE56, True
      End If
   End If
   '
   EnableTextBox textCE54, False
   If IsNull(rsTmp.Fields("CE53")) = False Then
      If IsEmptyText(rsTmp.Fields("CE53")) = False Then
         EnableTextBox textCE54, True
      End If
   End If
   '
   EnableTextBox textCE52, False
   If IsNull(rsTmp.Fields("CE51")) = False Then
      If IsEmptyText(rsTmp.Fields("CE51")) = False Then
         EnableTextBox textCE52, True
      End If
   End If
   '
   EnableTextBox textCE22, False
   If IsNull(rsTmp.Fields("CE17")) = False Then
      If IsEmptyText(rsTmp.Fields("CE17")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE18")) = False Then
      If IsEmptyText(rsTmp.Fields("CE18")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE19")) = False Then
      If IsEmptyText(rsTmp.Fields("CE19")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE20")) = False Then
      If IsEmptyText(rsTmp.Fields("CE20")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE21")) = False Then
      If IsEmptyText(rsTmp.Fields("CE21")) = False Then
         EnableTextBox textCE22, True
      End If
   End If
   '
   EnableTextBox textCE65, False
   If IsNull(rsTmp.Fields("CE63")) = False Then
      If IsEmptyText(rsTmp.Fields("CE63")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE64")) = False Then
      If IsEmptyText(rsTmp.Fields("CE64")) = False Then
         EnableTextBox textCE65, True
      End If
   End If
   '
   EnableTextBox textCE38, False
   If IsNull(rsTmp.Fields("CE23")) = False Then
      If IsEmptyText(rsTmp.Fields("CE23")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE24")) = False Then
      If IsEmptyText(rsTmp.Fields("CE24")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE25")) = False Then
      If IsEmptyText(rsTmp.Fields("CE25")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE26")) = False Then
      If IsEmptyText(rsTmp.Fields("CE26")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE27")) = False Then
      If IsEmptyText(rsTmp.Fields("CE27")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE28")) = False Then
      If IsEmptyText(rsTmp.Fields("CE28")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE29")) = False Then
      If IsEmptyText(rsTmp.Fields("CE29")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE30")) = False Then
      If IsEmptyText(rsTmp.Fields("CE30")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE31")) = False Then
      If IsEmptyText(rsTmp.Fields("CE31")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE32")) = False Then
      If IsEmptyText(rsTmp.Fields("CE32")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE33")) = False Then
      If IsEmptyText(rsTmp.Fields("CE33")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE34")) = False Then
      If IsEmptyText(rsTmp.Fields("CE34")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE35")) = False Then
      If IsEmptyText(rsTmp.Fields("CE35")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE36")) = False Then
      If IsEmptyText(rsTmp.Fields("CE36")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE37")) = False Then
      If IsEmptyText(rsTmp.Fields("CE37")) = False Then
         EnableTextBox textCE38, True
      End If
   End If
   '
   EnableTextBox textCE58, False
   If IsNull(rsTmp.Fields("CE57")) = False Then
      If IsEmptyText(rsTmp.Fields("CE57")) = False Then
         EnableTextBox textCE58, True
      End If
   End If
   '
   EnableTextBox textCE40, False
   If IsNull(rsTmp.Fields("CE39")) = False Then
      If IsEmptyText(rsTmp.Fields("CE39")) = False Then
         EnableTextBox textCE40, True
      End If
   End If
   '
   EnableTextBox textCE60, False
   If IsNull(rsTmp.Fields("CE59")) = False Then
      If IsEmptyText(rsTmp.Fields("CE59")) = False Then
         EnableTextBox textCE60, True
      End If
   End If
   '
   EnableTextBox textCE44, False
   If IsNull(rsTmp.Fields("CE41")) = False Then
      If IsEmptyText(rsTmp.Fields("CE41")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE42")) = False Then
      If IsEmptyText(rsTmp.Fields("CE42")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   If IsNull(rsTmp.Fields("CE43")) = False Then
      If IsEmptyText(rsTmp.Fields("CE43")) = False Then
         EnableTextBox textCE44, True
      End If
   End If
   '
   EnableTextBox textCE46, False
   If IsNull(rsTmp.Fields("CE45")) = False Then
      If IsEmptyText(rsTmp.Fields("CE45")) = False Then
         EnableTextBox textCE46, True
      End If
   End If
   '
   EnableTextBox textCE48, False
   If IsNull(rsTmp.Fields("CE47")) = False Then
      If IsEmptyText(rsTmp.Fields("CE47")) = False Then
         EnableTextBox textCE48, True
      End If
   End If
   '
   EnableTextBox textCE50, False
   If IsNull(rsTmp.Fields("CE49")) = False Then
      If IsEmptyText(rsTmp.Fields("CE49")) = False Then
         EnableTextBox textCE50, True
      End If
   End If
   '
   EnableTextBox textCE62, False
   If IsNull(rsTmp.Fields("CE61")) = False Then
      If IsEmptyText(rsTmp.Fields("CE61")) = False Then
         EnableTextBox textCE62, True
      End If
   End If

   ' 密碼欄位開放允許可輸入
   '92.10.8 還原 BY SONIA
   'EnableTextBox textCE66, True
   'EnableTextBox textCE67, True
   '
   EnableTextBox textCE66, False
   EnableTextBox textCE67, False
   If IsNull(rsTmp.Fields("CE66")) = False Then
      If IsEmptyText(rsTmp.Fields("CE66")) = False Then
         EnableTextBox textCE67, True
      End If
   End If
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   m_CP13 = Empty
   m_CP12 = Empty
   m_CP14 = Empty   '2015/1/14 add by sonia
   
   ' 本所案號
   textSPKey = m_SP01 & m_SP02 & m_SP03 & m_SP04
   ' 收文號
   textCP09 = m_CP09
   
   ' 讀取服務業務基本檔檔案
   QueryServicePractice
   ' 讀取變更事項檔檔案
   QueryChangeEvent
   
   ' 取得案件進度檔A類資料的最後一筆
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_SP01 & "' AND " & _
                  "CP02 = '" & m_SP02 & "' AND " & _
                  "CP03 = '" & m_SP03 & "' AND " & _
                  "CP04 = '" & m_SP04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
      End If
      ' 案件性質
      'Add By Cheng 2002/07/17
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         If m_SP09 < "010" Then
            textCP10_2 = GetCaseTypeName(m_SP01, m_CP10, 0)
         Else
            textCP10_2 = GetCaseTypeName(m_SP01, m_CP10, 1)
         End If
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"), True)
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 進度備註
      m_CP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         m_CP64 = rsTmp.Fields("CP64")
      End If
      m_CP14 = "" & rsTmp("CP14").Value  '2015/1/14 add by sonia
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_SP01, m_SP02, m_SP03, m_SP04)
   End If
   
End Sub

'Modify By Cheng 2002/11/07
'Public Sub OnUpdateServicePractice()
Public Function OnUpdateServicePractice() As Boolean
   Dim strSql As String
   Dim strTmp As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateServicePractice = True

   ' 更新服務業務基本檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_SPCount - 1
      strTmp = Empty
      If m_SPList(nIndex).fiOldData <> m_SPList(nIndex).fiNewData Then
         If m_SPList(nIndex).fiType = 0 Then
            If m_SPList(nIndex).fiNewData = Empty Then
               strTmp = m_SPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_SPList(nIndex).fiName & " = '" & m_SPList(nIndex).fiNewData & "'"
            End If
         Else
            If m_SPList(nIndex).fiNewData = Empty Then
               strTmp = m_SPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_SPList(nIndex).fiName & " = " & m_SPList(nIndex).fiNewData
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
   Next nIndex
   ' 組成SQL語法
   strSql = strSql & " " & _
                  "WHERE SP01 = '" & m_SP01 & "' AND " & _
                        "SP02 = '" & m_SP02 & "' AND " & _
                        "SP03 = '" & m_SP03 & "' AND " & _
                        "SP04 = '" & m_SP04 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

Private Sub OnUpdateField()
   Dim strTemp As String
   Dim strCP64 As String
   strCP64 = Empty
   If textCE09 = "1" Then
      SetSPFieldNewData "SP08", GetCEFieldOldData("CE04")
      strCP64 = strCP64 & "原申請人:" & GetSPFieldOldData("SP08") & " "
   End If
   If textCE03 = "1" Then
      SetSPFieldNewData "SP10", GetCEFieldOldData("CE02")
      strCP64 = strCP64 & "原申請日:" & GetSPFieldOldData("SP10") & " "
   End If
   If textCE16 = "1" Then
      SetSPFieldNewData "SP42", GetCEFieldOldData("CE10")
      strCP64 = strCP64 & "原代表人:" & GetSPFieldOldData("SP42") & " "
   End If
   If textCE44 = "1" Then
      strTemp = GetCEFieldOldData("CE41")
      If IsEmptyText(strTemp) = False Then
         SetSPFieldNewData "SP05", strTemp
         strCP64 = strCP64 & "原案件中文名稱:" & GetSPFieldOldData("SP05") & " "
      End If
      strTemp = GetCEFieldOldData("CE42")
      If IsEmptyText(strTemp) = False Then
         SetSPFieldNewData "SP06", strTemp
         strCP64 = strCP64 & "原案件英文名稱:" & GetSPFieldOldData("SP06") & " "
      End If
      strTemp = GetCEFieldOldData("CE43")
      If IsEmptyText(strTemp) = False Then
         SetSPFieldNewData "SP07", strTemp
         strCP64 = strCP64 & "原案件日文名稱:" & GetSPFieldOldData("SP07") & " "
      End If
   End If
   ' 密碼 (允許輸入, 故必須取得畫面上所輸入的新資料)
   If textCE67 = "1" Then
      SetSPFieldNewData "SP49", textCE66
      strCP64 = strCP64 & "原密碼:" & GetSPFieldOldData("SP49") & " "
   End If
   m_CP64 = m_CP64 & strCP64
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
      SetSPFieldNewData "SP72", textPrint
   Else
      SetSPFieldNewData "SP72", m_textPrint
   End If
   '92.10.8 ADD BY SONIA
   SetCEFieldNewData "CE03", textCE03: SetCEFieldNewData "CE09", textCE09: SetCEFieldNewData "CE16", textCE16: SetCEFieldNewData "CE22", textCE22: SetCEFieldNewData "CE38", textCE38
   SetCEFieldNewData "CE40", textCE40: SetCEFieldNewData "CE44", textCE44: SetCEFieldNewData "CE46", textCE46: SetCEFieldNewData "CE48", textCE48: SetCEFieldNewData "CE50", textCE50
   SetCEFieldNewData "CE52", textCE52: SetCEFieldNewData "CE54", textCE54: SetCEFieldNewData "CE56", textCE56: SetCEFieldNewData "CE58", textCE58: SetCEFieldNewData "CE60", textCE60
   SetCEFieldNewData "CE62", textCE62: SetCEFieldNewData "CE65", textCE65: SetCEFieldNewData "CE67", textCE67
   '92.10.8 END
End Sub

' 儲存資料
'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nIndex As Integer
   Dim bDifference As Boolean
   Dim bFirst As Boolean
   Dim strCP09 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strTmp As String
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質為服務業務結果
   strCP10 = "1801"
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   strCP27 = DBDATE(SystemDate())
   ' 91.03.25 modify by louis (單引號)
    '承辦人為使用者, 發文日為系統日
   '911216 NICK 要存入新的CP64
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
   '92.10.8 MODIFY BY SONIA CP64不必存入新的CP64
   'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
   '         "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
   '                 "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
   '                 "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64 & m_CP64) & "')"
    'Modify By Cheng 2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
'                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   '2015/1/14 modify by sonia 所有服務業務結果的承辦人改放原承辦人TM-000067(宋若蘭),否則期限表帶出之承辦人會是程序
   'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & strUserNum & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP64) " & _
            "VALUES ('" & m_SP01 & "','" & m_SP02 & "','" & m_SP03 & "','" & m_SP04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04)) & "','" & PUB_GetAKindSalesNo(m_SP01, m_SP02, m_SP03, m_SP04) & "','" & m_CP14 & "'," & _
                    "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & ChgSQL(textCP64) & "')"
    'End
   '92.10.8 END
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
   End If
   '2019/12/19 END
   
   'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
   Pub_UpdateFromMaxCP27 m_SP01, m_SP02, m_SP03, m_SP04
   
   ' 儲存變更事項檔
   strSql = "UPDATE ChangeEvent SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CECount - 1
      strTmp = Empty
      If m_CEList(nIndex).fiOldData <> m_CEList(nIndex).fiNewData Then
         If m_CEList(nIndex).fiType = 0 Then
            strTmp = m_CEList(nIndex).fiName & " = '" & m_CEList(nIndex).fiNewData & "'"
         Else
            If m_CEList(nIndex).fiNewData = Empty Then
               strTmp = m_CEList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_CEList(nIndex).fiName & " = " & m_CEList(nIndex).fiNewData
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
   Next nIndex
   
   strSql = strSql & " " & _
                  "WHERE CE01 = '" & m_CP09 & "'"
   
   If bDifference = True Then
      cnnConnection.Execute strSql
   End If
   
   '911204 nick 當cp01='TC' 時，下面 3 個動作不做
   If m_SP01 <> "TC" Then
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' 更新服務業務基本檔有變更的欄位
        'Modify By Cheng 2002/11/07
    '   OnUpdateServicePractice
       If OnUpdateServicePractice = False Then GoTo ErrorHandler
       
       ' 更新案件進度檔的進度備註欄位
       strSql = "UPDATE CaseProgress SET CP64 = '" & m_CP64 & "' " & _
                "WHERE CP09 = '" & m_CP09 & "' "
       cnnConnection.Execute strSql
   End If
       '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
       ' 更新案件進度檔所選取收文資料的實際結果為 1
       strSql = "UPDATE CaseProgress SET CP24 = '1' " & _
                "WHERE CP09 = '" & m_CP09 & "' "
       cnnConnection.Execute strSql
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '92.10.7 modify by sonia 改回仍列印
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
   If textPrint <> "N" Then
      PrintLetter
   End If
          'add by nickc 2005/04/22
          Pub_UpdateEndModCash m_SP01, m_SP02, m_SP03, m_SP04
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010409_1"
   End If
   '2019/5/22 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/11/21
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 檢查是否列印定稿欄位
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   UCase textPrint
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/11/21
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入 N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub
' 檢查該輸入的資料是否已完成
Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
   End If

   
   CheckDataValid = True
EXITSUB:
End Function

Private Function CheckIs1Or2(ByVal strData As String) As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckIs1Or2 = True
   If IsEmptyText(strData) = False Then
      Select Case strData
         Case "1", "2":
         Case Else
            CheckIs1Or2 = False
            strTit = "資料檢核"
            strMsg = "只可輸入1或2"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End Select
   End If
End Function

Private Sub textCE03_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE03) = False Then
      Cancel = True
      textCE03_GotFocus
   End If
End Sub

Private Sub textCE09_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE09) = False Then
      Cancel = True
      textCE09_GotFocus
   End If
End Sub

Private Sub textCE16_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE16) = False Then
      Cancel = True
      textCE16_GotFocus
   End If
End Sub

Private Sub textCE22_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE22) = False Then
      Cancel = True
      textCE22_GotFocus
   End If
End Sub

Private Sub textCE38_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE38) = False Then
      Cancel = True
      textCE38_GotFocus
   End If
End Sub

Private Sub textCE40_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE40) = False Then
      Cancel = True
      textCE40_GotFocus
   End If
End Sub

Private Sub textCE44_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE44) = False Then
      Cancel = True
      textCE44_GotFocus
   End If
End Sub

Private Sub textCE46_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE46) = False Then
      Cancel = True
      textCE46_GotFocus
   End If
End Sub

Private Sub textCE48_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE48) = False Then
      Cancel = True
      textCE48_GotFocus
   End If
End Sub

Private Sub textCE50_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE50) = False Then
      Cancel = True
      textCE50_GotFocus
   End If
End Sub

Private Sub textCE52_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE52) = False Then
      Cancel = True
      textCE52_GotFocus
   End If
End Sub

Private Sub textCE54_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE54) = False Then
      Cancel = True
      textCE54_GotFocus
   End If
End Sub

Private Sub textCE56_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE56) = False Then
      Cancel = True
      textCE56_GotFocus
   End If
End Sub

Private Sub textCE58_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE58) = False Then
      Cancel = True
      textCE58_GotFocus
   End If
End Sub

Private Sub textCE60_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE60) = False Then
      Cancel = True
      textCE60_GotFocus
   End If
End Sub

Private Sub textCE62_Validate(Cancel As Boolean)
   Cancel = False
   If CheckIs1Or2(textCE62) = False Then
      Cancel = True
      textCE62_GotFocus
   End If
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCE03_GotFocus()
   InverseTextBox textCE03
End Sub

Private Sub textCE09_GotFocus()
   InverseTextBox textCE09
End Sub

Private Sub textCE16_GotFocus()
   InverseTextBox textCE16
End Sub

Private Sub textCE22_GotFocus()
   InverseTextBox textCE22
End Sub

Private Sub textCE38_GotFocus()
   InverseTextBox textCE38
End Sub

Private Sub textCE40_GotFocus()
   InverseTextBox textCE40
End Sub

Private Sub textCE44_GotFocus()
   InverseTextBox textCE44
End Sub

Private Sub textCE46_GotFocus()
   InverseTextBox textCE46
End Sub

Private Sub textCE48_GotFocus()
   InverseTextBox textCE48
End Sub

Private Sub textCE50_GotFocus()
   InverseTextBox textCE50
End Sub

Private Sub textCE52_GotFocus()
   InverseTextBox textCE52
End Sub

Private Sub textCE54_GotFocus()
   InverseTextBox textCE54
End Sub

Private Sub textCE56_GotFocus()
   InverseTextBox textCE56
End Sub

Private Sub textCE58_GotFocus()
   InverseTextBox textCE58
End Sub

Private Sub textCE60_GotFocus()
   InverseTextBox textCE60
End Sub

Private Sub textCE62_GotFocus()
   InverseTextBox textCE62
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   Dim strTmp As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 系統別為TD
   If m_SP01 = "TD" Then
      ' 案件性質為變更
      If m_CP10 = "301" Then
         ' 清除定稿例外欄位檔原有資料
         EndLetter "06", m_CP09, "02", strUserNum
         ' 變更前名稱
         strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & "06" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                  "'" & "變更前名稱" & "','" & m_SP06 & "')"
         cnnConnection.Execute strSql
      End If
   End If
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "06"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/13 End
   
   Select Case m_SP01
      ' 系統別為TD
      Case "TD"
         ' 案件性質為變更
         If m_CP10 = "301" Then
            'add by nickc 2006/11/21
            If m_SP09 < "010" Then
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "06", "02", False, strUserNum, 0
                  ET03 = "02" 'Modify By Sindy 2012/1/13
                End If
            End If
         End If
      ' 系統別為TC
      Case "TC"
         ' 案件性質為變更
         If m_CP10 = "301" Then
            'add by nickc 2006/11/21
            If m_SP09 < "010" Then
                If textPrint = "1" Then
                    ' 列印定稿
'                    NowPrint m_CP09, "06", "06", False, strUserNum, 0
                  ET03 = "06" 'Modify By Sindy 2012/1/13
                End If
            End If
         End If
   End Select
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_SP01 & m_SP02 & m_SP03 & m_SP04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_SP01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/19 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/13 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textCE03.Enabled = True Then
   Cancel = False
   textCE03_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE09.Enabled = True Then
   Cancel = False
   textCE09_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE16.Enabled = True Then
   Cancel = False
   textCE16_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE22.Enabled = True Then
   Cancel = False
   textCE22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE38.Enabled = True Then
   Cancel = False
   textCE38_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE40.Enabled = True Then
   Cancel = False
   textCE40_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE44.Enabled = True Then
   Cancel = False
   textCE44_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE46.Enabled = True Then
   Cancel = False
   textCE46_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE48.Enabled = True Then
   Cancel = False
   textCE48_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE50.Enabled = True Then
   Cancel = False
   textCE50_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE52.Enabled = True Then
   Cancel = False
   textCE52_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE54.Enabled = True Then
   Cancel = False
   textCE54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE56.Enabled = True Then
   Cancel = False
   textCE56_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE58.Enabled = True Then
   Cancel = False
   textCE58_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE60.Enabled = True Then
   Cancel = False
   textCE60_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCE62.Enabled = True Then
   Cancel = False
   textCE62_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP64.Enabled = True Then
   Cancel = False
   textCP64_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

