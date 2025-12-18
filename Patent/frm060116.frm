VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060116 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件基本資料-業務承辦組"
   ClientHeight    =   7056
   ClientLeft      =   828
   ClientTop       =   972
   ClientWidth     =   8832
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7056
   ScaleWidth      =   8832
   Begin VB.CommandButton CmdPA174 
      BackColor       =   &H00C0FFFF&
      Caption         =   "特殊字"
      Height          =   280
      Left            =   60
      Style           =   1  '圖片外觀
      TabIndex        =   109
      Top             =   1562
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人參考資料(&A)"
      Height          =   375
      Index           =   3
      Left            =   1995
      TabIndex        =   20
      Top             =   73
      Width           =   1710
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人參考資料(&C)"
      Height          =   375
      Index           =   1
      Left            =   180
      TabIndex        =   19
      Top             =   73
      Width           =   1755
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   510
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7800
      TabIndex        =   34
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6915
      TabIndex        =   33
      Top             =   60
      Width           =   840
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5000
      Left            =   120
      TabIndex        =   41
      Top             =   1920
      Width           =   8505
      _ExtentX        =   15007
      _ExtentY        =   8805
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm060116.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "txtPA(75)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtPA(26)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblPA(75)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblPA(26)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblPA(77)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblPA(88)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblPA(48)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPA(133)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPA(134)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPA(159)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label27(26)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label27(75)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtPA(77)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtPA(159)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtPA(88)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label27(88)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtPA(133)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label27(133)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtPA(134)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label27(134)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtPA(48)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "聯絡人"
      TabPicture(1)   =   "frm060116.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblPA(51)"
      Tab(1).Control(1)=   "lblPA(52)"
      Tab(1).Control(2)=   "lblPA(53)"
      Tab(1).Control(3)=   "lblPA(54)"
      Tab(1).Control(4)=   "lblPA(55)"
      Tab(1).Control(5)=   "lblPA(56)"
      Tab(1).Control(6)=   "lblPA(98)"
      Tab(1).Control(7)=   "lblPA(99)"
      Tab(1).Control(8)=   "lblPA(100)"
      Tab(1).Control(9)=   "lblPA(139)"
      Tab(1).Control(10)=   "Combo1(0)"
      Tab(1).Control(11)=   "Combo1(1)"
      Tab(1).Control(12)=   "txtPA(52)"
      Tab(1).Control(13)=   "txtPA(51)"
      Tab(1).Control(14)=   "txtPA(53)"
      Tab(1).Control(15)=   "txtPA(55)"
      Tab(1).Control(16)=   "txtPA(54)"
      Tab(1).Control(17)=   "txtPA(56)"
      Tab(1).Control(18)=   "txtPA(99)"
      Tab(1).Control(19)=   "txtPA(98)"
      Tab(1).Control(20)=   "txtPA(100)"
      Tab(1).Control(21)=   "txtPA(139)"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "副本資料"
      TabPicture(2)   =   "frm060116.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblPA(86)"
      Tab(2).Control(1)=   "lblPA(87)"
      Tab(2).Control(2)=   "lblPA(101)"
      Tab(2).Control(3)=   "lblPA(102)"
      Tab(2).Control(4)=   "lblPA(103)"
      Tab(2).Control(5)=   "lblPA(104)"
      Tab(2).Control(6)=   "txtPA(86)"
      Tab(2).Control(7)=   "Label27(86)"
      Tab(2).Control(8)=   "txtPA(101)"
      Tab(2).Control(9)=   "Label27(101)"
      Tab(2).Control(10)=   "txtPA(87)"
      Tab(2).Control(11)=   "txtPA(102)"
      Tab(2).Control(12)=   "txtPA(103)"
      Tab(2).Control(13)=   "txtPA(104)"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "日文資料"
      TabPicture(3)   =   "frm060116.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lblPA(7)"
      Tab(3).Control(1)=   "lblPA(132)"
      Tab(3).Control(2)=   "lblPA(129)"
      Tab(3).Control(3)=   "lblPA(126)"
      Tab(3).Control(4)=   "lblPA(123)"
      Tab(3).Control(5)=   "lblPA(120)"
      Tab(3).Control(6)=   "lblPA(117)"
      Tab(3).Control(7)=   "lblPA(114)"
      Tab(3).Control(8)=   "lblPA(111)"
      Tab(3).Control(9)=   "lblPA(84)"
      Tab(3).Control(10)=   "lblPA(45)"
      Tab(3).Control(11)=   "lblPA(4)"
      Tab(3).Control(12)=   "lblPA(43)"
      Tab(3).Control(13)=   "lblPA(42)"
      Tab(3).Control(14)=   "lblPA(81)"
      Tab(3).Control(15)=   "lblPA(41)"
      Tab(3).Control(16)=   "txtPA(7)"
      Tab(3).Control(17)=   "txtPA(41)"
      Tab(3).Control(18)=   "txtPA(42)"
      Tab(3).Control(19)=   "txtPA(43)"
      Tab(3).Control(20)=   "txtPA(44)"
      Tab(3).Control(21)=   "txtPA(45)"
      Tab(3).Control(22)=   "txtPA(81)"
      Tab(3).Control(23)=   "txtPA(111)"
      Tab(3).Control(24)=   "txtPA(117)"
      Tab(3).Control(25)=   "txtPA(123)"
      Tab(3).Control(26)=   "txtPA(129)"
      Tab(3).Control(27)=   "txtPA(120)"
      Tab(3).Control(28)=   "txtPA(84)"
      Tab(3).Control(29)=   "txtPA(114)"
      Tab(3).Control(30)=   "txtPA(126)"
      Tab(3).Control(31)=   "txtPA(132)"
      Tab(3).ControlCount=   32
      TabCaption(4)   =   "備註"
      TabPicture(4)   =   "frm060116.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1(172)"
      Tab(4).Control(1)=   "txtPA(91)"
      Tab(4).ControlCount=   2
      Begin MSForms.TextBox txtPA 
         Height          =   4095
         Index           =   91
         Left            =   -74880
         TabIndex        =   60
         Top             =   480
         Width           =   8250
         VariousPropertyBits=   -1466906597
         ScrollBars      =   2
         Size            =   "14552;7223"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   132
         Left            =   -69720
         TabIndex        =   57
         Top             =   4080
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   126
         Left            =   -69720
         TabIndex        =   55
         Top             =   3720
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   114
         Left            =   -69720
         TabIndex        =   51
         Top             =   3000
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   84
         Left            =   -69720
         TabIndex        =   49
         Top             =   2640
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   120
         Left            =   -69720
         TabIndex        =   53
         Top             =   3360
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   129
         Left            =   -73880
         TabIndex        =   56
         Top             =   4080
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   123
         Left            =   -73880
         TabIndex        =   54
         Top             =   3720
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   117
         Left            =   -73880
         TabIndex        =   52
         Top             =   3360
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   111
         Left            =   -73880
         TabIndex        =   50
         Top             =   3000
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   81
         Left            =   -73880
         TabIndex        =   48
         Top             =   2640
         Width           =   3000
         VariousPropertyBits=   679495707
         Size            =   "5292;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   45
         Left            =   -73440
         TabIndex        =   47
         Top             =   2160
         Width           =   6000
         VariousPropertyBits=   679495707
         Size            =   "10583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   44
         Left            =   -73440
         TabIndex        =   46
         Top             =   1824
         Width           =   6000
         VariousPropertyBits=   679495707
         Size            =   "10583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   43
         Left            =   -73440
         TabIndex        =   45
         Top             =   1488
         Width           =   6000
         VariousPropertyBits=   679495707
         Size            =   "10583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   42
         Left            =   -73440
         TabIndex        =   44
         Top             =   1152
         Width           =   6000
         VariousPropertyBits=   679495707
         Size            =   "10583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   41
         Left            =   -73440
         TabIndex        =   43
         Top             =   816
         Width           =   6000
         VariousPropertyBits=   679495707
         Size            =   "10583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   7
         Left            =   -73440
         TabIndex        =   42
         Top             =   480
         Width           =   6720
         VariousPropertyBits=   679495707
         Size            =   "11853;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   525
         Index           =   104
         Left            =   -74880
         TabIndex        =   40
         Top             =   3120
         Width           =   8235
         VariousPropertyBits=   -1466939365
         Size            =   "14526;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   525
         Index           =   103
         Left            =   -74880
         TabIndex        =   39
         Top             =   2280
         Width           =   8235
         VariousPropertyBits=   -1466939365
         Size            =   "14526;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   102
         Left            =   -73500
         TabIndex        =   38
         Top             =   1560
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   87
         Left            =   -73500
         TabIndex        =   36
         Top             =   840
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   101
         Left            =   -72300
         TabIndex        =   108
         Top             =   1200
         Width           =   5205
         Caption         =   "Label27"
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   101
         Left            =   -73500
         TabIndex        =   37
         Top             =   1200
         Width           =   1095
         VariousPropertyBits=   679495707
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   86
         Left            =   -72300
         TabIndex        =   107
         Top             =   480
         Width           =   5205
         Caption         =   "Label27"
         Size            =   "9181;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   86
         Left            =   -73500
         TabIndex        =   35
         Top             =   480
         Width           =   1095
         VariousPropertyBits=   679495707
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   139
         Left            =   -73560
         TabIndex        =   29
         Top             =   3120
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11994;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   100
         Left            =   -73560
         TabIndex        =   32
         Top             =   4200
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11994;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   98
         Left            =   -73560
         TabIndex        =   30
         Top             =   3600
         Width           =   2400
         VariousPropertyBits=   679495707
         Size            =   "4233;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   99
         Left            =   -73560
         TabIndex        =   31
         Top             =   3900
         Width           =   5505
         VariousPropertyBits=   679495707
         Size            =   "9710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   56
         Left            =   -73560
         TabIndex        =   28
         Top             =   2812
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11986;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   54
         Left            =   -73560
         TabIndex        =   26
         Top             =   2200
         Width           =   5500
         VariousPropertyBits=   679495707
         Size            =   "9701;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   55
         Left            =   -73560
         TabIndex        =   27
         Top             =   2506
         Width           =   5500
         VariousPropertyBits=   679495707
         Size            =   "9701;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   53
         Left            =   -73560
         TabIndex        =   24
         Top             =   1440
         Width           =   6795
         VariousPropertyBits=   679495707
         Size            =   "11994;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   51
         Left            =   -73560
         TabIndex        =   22
         Top             =   840
         Width           =   5500
         VariousPropertyBits=   679495707
         Size            =   "9701;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   52
         Left            =   -73560
         TabIndex        =   23
         Top             =   1140
         Width           =   5500
         VariousPropertyBits=   679495707
         Size            =   "9701;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Index           =   1
         Left            =   -73560
         TabIndex        =   25
         Top             =   1860
         Width           =   6795
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "11986;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   330
         Index           =   0
         Left            =   -73560
         TabIndex        =   21
         Top             =   480
         Width           =   6795
         VariousPropertyBits=   746604571
         DisplayStyle    =   7
         Size            =   "11986;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   48
         Left            =   1920
         TabIndex        =   13
         Top             =   3240
         Width           =   6360
         VariousPropertyBits=   679495707
         Size            =   "11218;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   134
         Left            =   3060
         TabIndex        =   106
         Top             =   2880
         Width           =   5200
         Caption         =   "Label27"
         Size            =   "9172;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   134
         Left            =   1920
         TabIndex        =   12
         Top             =   2880
         Width           =   1095
         VariousPropertyBits=   679495707
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   133
         Left            =   3060
         TabIndex        =   105
         Top             =   2520
         Width           =   5200
         Caption         =   "Label27"
         Size            =   "9172;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   133
         Left            =   1920
         TabIndex        =   11
         Top             =   2520
         Width           =   1095
         VariousPropertyBits=   679495707
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   88
         Left            =   3060
         TabIndex        =   104
         Top             =   2160
         Width           =   5200
         Caption         =   "Label27"
         Size            =   "9172;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   88
         Left            =   1920
         TabIndex        =   10
         Top             =   2160
         Width           =   1095
         VariousPropertyBits=   679495707
         Size            =   "1940;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   159
         Left            =   1920
         TabIndex        =   9
         Top             =   1800
         Width           =   2100
         VariousPropertyBits=   679495707
         Size            =   "3704;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   77
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   4200
         VariousPropertyBits=   679495707
         Size            =   "7408;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   75
         Left            =   2040
         TabIndex        =   103
         Top             =   840
         Width           =   6255
         Caption         =   "Label27"
         Size            =   "11033;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   285
         Index           =   26
         Left            =   2040
         TabIndex        =   102
         Top             =   480
         Width           =   6375
         Caption         =   "Label27"
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID:"
         Height          =   225
         Index           =   159
         Left            =   120
         TabIndex        =   101
         Top             =   1770
         Width           =   1725
      End
      Begin VB.Label lblPA 
         Caption         =   "年費D/N列印對象:"
         Height          =   225
         Index           =   134
         Left            =   435
         TabIndex        =   100
         Top             =   2910
         Width           =   1485
      End
      Begin VB.Label lblPA 
         Caption         =   "D/N固定列印對象:"
         Height          =   225
         Index           =   133
         Left            =   440
         TabIndex        =   99
         Top             =   2550
         Width           =   1485
      End
      Begin VB.Label lblPA 
         Caption         =   "客戶案件案號:"
         Height          =   225
         Index           =   48
         Left            =   720
         TabIndex        =   98
         Top             =   3285
         Width           =   1215
      End
      Begin VB.Label lblPA 
         Caption         =   "固定請款對象:"
         Height          =   225
         Index           =   88
         Left            =   720
         TabIndex        =   97
         Top             =   2190
         Width           =   1215
      End
      Begin VB.Label lblPA 
         Caption         =   "彼所案號:"
         Height          =   225
         Index           =   77
         Left            =   1080
         TabIndex        =   96
         Top             =   1470
         Width           =   855
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日):"
         Height          =   180
         Index           =   139
         Left            =   -74880
         TabIndex        =   95
         Top             =   3172
         Width           =   1245
      End
      Begin VB.Label lblPA 
         Caption         =   "實體聯絡人(日):"
         Height          =   255
         Index           =   100
         Left            =   -74880
         TabIndex        =   94
         Top             =   4215
         Width           =   1335
      End
      Begin VB.Label lblPA 
         Caption         =   "實體聯絡人(英):"
         Height          =   255
         Index           =   99
         Left            =   -74880
         TabIndex        =   93
         Top             =   3915
         Width           =   1335
      End
      Begin VB.Label lblPA 
         Caption         =   "實體聯絡人(中):"
         Height          =   255
         Index           =   98
         Left            =   -74880
         TabIndex        =   92
         Top             =   3615
         Width           =   1335
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人2(日):"
         Height          =   255
         Index           =   56
         Left            =   -74880
         TabIndex        =   91
         Top             =   2827
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人2(英):"
         Height          =   255
         Index           =   55
         Left            =   -74880
         TabIndex        =   90
         Top             =   2521
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人2(中):"
         Height          =   255
         Index           =   54
         Left            =   -74880
         TabIndex        =   89
         Top             =   2215
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人1(日):"
         Height          =   255
         Index           =   53
         Left            =   -74880
         TabIndex        =   88
         Top             =   1455
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人1(英):"
         Height          =   255
         Index           =   52
         Left            =   -74880
         TabIndex        =   87
         Top             =   1155
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "聯絡人1(中):"
         Height          =   255
         Index           =   51
         Left            =   -74880
         TabIndex        =   86
         Top             =   855
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "實體副本收受人彼所案號2:"
         Height          =   255
         Index           =   104
         Left            =   -74880
         TabIndex        =   85
         Top             =   2900
         Width           =   2295
      End
      Begin VB.Label lblPA 
         Caption         =   "實體副本收受人彼所案號1:"
         Height          =   255
         Index           =   103
         Left            =   -74880
         TabIndex        =   84
         Top             =   2040
         Width           =   2295
      End
      Begin VB.Label lblPA 
         Caption         =   "實體副本聯絡人:"
         Height          =   225
         Index           =   102
         Left            =   -74880
         TabIndex        =   83
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblPA 
         Caption         =   "實體副本收受人:"
         Height          =   220
         Index           =   101
         Left            =   -74880
         TabIndex        =   82
         Top             =   1212
         Width           =   1335
      End
      Begin VB.Label lblPA 
         Caption         =   "副本聯絡人:"
         Height          =   225
         Index           =   87
         Left            =   -74880
         TabIndex        =   81
         Top             =   861
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "副本收受人:"
         Height          =   225
         Index           =   86
         Left            =   -74880
         TabIndex        =   80
         Top             =   510
         Width           =   1095
      End
      Begin VB.Label lblPA 
         Caption         =   "申請人1:"
         Height          =   225
         Index           =   26
         Left            =   120
         TabIndex        =   79
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lblPA 
         Caption         =   "代理人:"
         Height          =   225
         Index           =   75
         Left            =   240
         TabIndex        =   78
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "與他案合併計算結餘，請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   172
         Left            =   -74880
         TabIndex        =   77
         Top             =   4640
         Width           =   7665
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "申請人1地址(日):"
         Height          =   180
         Index           =   41
         Left            =   -74880
         TabIndex        =   76
         Top             =   885
         Width           =   1335
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人1(日):"
         Height          =   180
         Index           =   81
         Left            =   -74880
         TabIndex        =   75
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "申請人2地址(日):"
         Height          =   180
         Index           =   42
         Left            =   -74880
         TabIndex        =   74
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "申請人3地址(日):"
         Height          =   180
         Index           =   43
         Left            =   -74880
         TabIndex        =   73
         Top             =   1545
         Width           =   1335
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "申請人4地址(日):"
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   72
         Top             =   1875
         Width           =   1335
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "申請人5地址(日):"
         Height          =   180
         Index           =   45
         Left            =   -74880
         TabIndex        =   71
         Top             =   2205
         Width           =   1335
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人2(日):"
         Height          =   180
         Index           =   84
         Left            =   -70710
         TabIndex        =   70
         Top             =   2700
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人3(日):"
         Height          =   180
         Index           =   111
         Left            =   -74880
         TabIndex        =   69
         Top             =   3060
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人4(日):"
         Height          =   180
         Index           =   114
         Left            =   -70710
         TabIndex        =   68
         Top             =   3045
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人5(日):"
         Height          =   180
         Index           =   117
         Left            =   -74880
         TabIndex        =   67
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人6(日):"
         Height          =   180
         Index           =   120
         Left            =   -70710
         TabIndex        =   66
         Top             =   3420
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人7(日):"
         Height          =   180
         Index           =   123
         Left            =   -74880
         TabIndex        =   65
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人8(日):"
         Height          =   180
         Index           =   126
         Left            =   -70710
         TabIndex        =   64
         Top             =   3780
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人9(日):"
         Height          =   180
         Index           =   129
         Left            =   -74880
         TabIndex        =   63
         Top             =   4140
         Width           =   975
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "代表人10(日):"
         Height          =   180
         Index           =   132
         Left            =   -70800
         TabIndex        =   62
         Top             =   4140
         Width           =   1065
      End
      Begin VB.Label lblPA 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱      (日):"
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   61
         Top             =   525
         Width           =   1335
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   26
         Left            =   960
         TabIndex        =   59
         Top             =   480
         Width           =   960
         VariousPropertyBits=   679495707
         Size            =   "1693;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPA 
         Height          =   285
         Index           =   75
         Left            =   960
         TabIndex        =   58
         Top             =   810
         Width           =   960
         VariousPropertyBits=   679495707
         Size            =   "1693;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Label lblPA174 
      Caption         =   "有特殊字"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   120
      TabIndex        =   110
      Top             =   1290
      Width           =   765
   End
   Begin MSForms.TextBox txtPA07 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   1560
      Width           =   7200
      VariousPropertyBits=   679495707
      MaxLength       =   160
      Size            =   "12700;494"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   6
      Left            =   1440
      TabIndex        =   6
      Top             =   1260
      Width           =   7200
      VariousPropertyBits=   679495707
      MaxLength       =   250
      Size            =   "12700;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   5
      Left            =   1440
      TabIndex        =   5
      Top             =   960
      Width           =   7200
      VariousPropertyBits=   679495707
      MaxLength       =   160
      Size            =   "12700;494"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   4
      Left            =   3360
      TabIndex        =   3
      Top             =   555
      Width           =   495
      VariousPropertyBits=   679495707
      MaxLength       =   2
      Size            =   "873;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   3
      Left            =   2955
      TabIndex        =   2
      Top             =   555
      Width           =   375
      VariousPropertyBits=   679495707
      MaxLength       =   1
      Size            =   "661;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   2
      Left            =   2085
      TabIndex        =   1
      Top             =   555
      Width           =   855
      VariousPropertyBits=   679495707
      MaxLength       =   6
      Size            =   "1508;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtPA 
      Height          =   285
      Index           =   1
      Left            =   1440
      TabIndex        =   0
      Top             =   555
      Width           =   615
      VariousPropertyBits=   679495707
      MaxLength       =   3
      Size            =   "1085;503"
      Value           =   "FCP"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(外):"
      Height          =   180
      Index           =   0
      Left            =   960
      TabIndex        =   18
      Top             =   1612
      Width           =   345
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(英):"
      Height          =   180
      Left            =   960
      TabIndex        =   17
      Top             =   1312
      Width           =   345
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "(中):"
      Height          =   180
      Left            =   960
      TabIndex        =   16
      Top             =   1012
      Width           =   345
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱"
      Height          =   180
      Left            =   120
      TabIndex        =   15
      Top             =   1012
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   570
      Width           =   765
   End
End
Attribute VB_Name = "frm060116"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2018/11/19 改成Form2.0 (Label27、Combo1和txtPA)
'Create by Lydia 2015/12/03 案件基本資料-業務承辦組
Option Explicit

Dim pa() As String

Dim intWhere As Integer, bolExist As Boolean

'Modified by Lydia 2018/11/19 改成Form2.0
'Dim oText As TextBox
'Dim oLabel As Label
Dim oText As MSForms.TextBox
Dim oLabel As MSForms.LABEL
Dim bolMsgRight As Boolean 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效
Dim SyxMsg As String 'Added by Lydia 2018/11/21 Form 2.0表單是否彈過提示滑鼠右鍵無效(記錄前一位置)

Dim ii As Integer
Dim m_CP09 As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0 '確定
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         FormClear
         txtPA(2).SetFocus
         SSTab1.Tab = 0
      Case 2 '結束
         Unload Me
      Case 1 '申請人參考資料
         If txtPA(26) <> "" Then
            If mdiMain.mnuTitle(10).Enabled = True Then
               Me.Enabled = False
               frm100101_11.Show
               frm100101_11.Tag = ChangeCustomerL(txtPA(26)) '傳申請人代號
               frm100101_11.StrMenu
               Me.Enabled = True
            Else
               MsgBox "請先關閉共同查詢畫面！"
            End If
         End If
      Case 3 '代理人參考資料
         If txtPA(75) <> "" Then
            If mdiMain.mnuTitle(10).Enabled = True Then
               Me.Enabled = False
               frm100101_10.Show
               frm100101_10.Tag = ChangeCustomerL(txtPA(75)) '傳代理人代號
               frm100101_10.StrMenu
               Me.Enabled = True
            Else
               MsgBox "請先關閉共同查詢畫面！"
            End If
         End If
   End Select
End Sub

Private Sub Combo1_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   If Combo1(Index) = "" Then
      For i = 51 To 53
         txtPA(i + Index * 3) = ""
      Next i
      'Modified by Lydia 2017/05/02 無聯絡人資料才清空
      'txtPA(139) = ""
      If Trim(txtPA(51) & txtPA(52) & txtPA(53) & txtPA(54) & txtPA(55) & txtPA(56)) = "" Then
         txtPA(139) = ""
      End If
      'end 2017/05/02
      Exit Sub
   End If

   strTmp = Mid(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") + 1, 1)
   Select Case txtPA(1)
      Case "FCP", "P", "CFP"
         If pa(75) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09,FA78"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54,FA78"
            End Select
            
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60,CU114"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63,CU114"
            End Select
         End If

      Case "FG", "PS", "CPS"
         If pa(26) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09,FA78"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54,FA78"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60,CU114"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63,CU114"
            End Select
         End If
   End Select
   
   strExc(2) = ChgFagent(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   strExc(3) = ChgCustomer(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   Select Case txtPA(1)
      Case "FCP", "P", "CFP"
         If pa(75) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
      Case "FG", "PS", "CPS"
         If pa(26) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       For intI = 51 To 53
          txtPA(intI + Index * 3) = "" & RsTemp.Fields(intI - 51)
       Next intI
       'Modified by Lydia 2017/05/02 有值才代入,避免清掉聯絡人部門(日)
       'txtPA(139) = "" & RsTemp.Fields(3)
       If "" & RsTemp.Fields(3) <> "" Then txtPA(139) = RsTemp.Fields(3)
   End If
End Sub

Private Sub Combo1_GotFocus(Index As Integer)
   CloseIme
End Sub

Private Sub Command1_Click()
 Dim i As Integer
 
 SSTab1.Tab = 0

   If txtPA(1) = "" Or Len(txtPA(2)) <> 6 Then
      MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
      txtPA(1).SetFocus
      Exit Sub
   End If
   If txtPA(3) = "" Then txtPA(3) = "0"
   If txtPA(4) = "" Then txtPA(4) = "00"
   pa(1) = txtPA(1):    pa(2) = txtPA(2)
   pa(3) = txtPA(3):   pa(4) = txtPA(4)
   FormClear
   
   '限外專承辦收文案件
   'Modified by Lydia 2016/01/08 改為A類最後收文之部門
    'strExc(0) = "select cp12,cp09 from caseprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp31='Y' "
    strExc(0) = "select cp12,cp09 from caseprogress where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp09<'B' order by cp05 desc "
    intI = 1
    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        If Left(RsTemp("cp12"), 2) <> "F2" Then
            MsgBox "本案非外專承辦收文，請重新輸入 !", vbCritical
            txtPA(1).SetFocus
            Exit Sub
        End If
    Else
        MsgBox "本所案號輸入錯誤，請重新輸入 !", vbCritical
        txtPA(1).SetFocus
        Exit Sub
    End If
    
   m_CP09 = RsTemp.Fields("CP09")

   Select Case pa(1)
      Case "FCP", "P", "CFP"
         If ClsPDReadPatentDatabase(pa(), intWhere) Then
            PatentShow
         End If
      Case "FG", "PS", "CPS"
         If ClsPDReadServicePracticeDatabase(pa(), intWhere) Then
            ServiceShow
         End If
   End Select
   
   txtPA07 = txtPA(7)
   If Left(pa(26), 6) = "X27766" And txtPA(101) <> "" And txtPA(103) = "" And txtPA(104) = "" Then
      txtPA(103) = "*Murata's reference number for the U.S. Patent application is"
      txtPA(104) = "*Corresponding Japanese Patent Application number"
   End If
   
   'If txtPA(5) <> "" Then 'Removed by Morgan 2017/8/16 不必控制 P115159
      cmdOK(0).Enabled = True
   'End If
   
   If txtPA(26) <> "" Then
      cmdOK(1).Enabled = True
   End If
   If txtPA(75) <> "" Then
      cmdOK(3).Enabled = True
   End If
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   FormClear
   SSTab1.Tab = 0
   SendKeys "{Tab}"
   ReDim pa(TF_PA) '陣列大小改用全域變數
   
   'Added by Lydia 2018/11/20 模組-抓DB中的欄位實際長度
   For Each oText In txtPA
        If InStr("88,133,134,86,101", Format(oText.Index, "00")) > 0 Then '輸入8碼,後面存檔補0
             oText.MaxLength = 8
        'Added by Lydia 2021/07/30 因為不可維護，排除中文、英文、日文名稱,避免maxlength判斷
        ElseIf InStr("05,06,07", Format(oText.Index, "00")) > 0 Then
             oText.MaxLength = 0
        'end 2021/0730
        Else
             oText.MaxLength = PUB_GetFieldDefSize("PATENT", "PA" & Format(oText.Index, "00"))
        End If
   Next
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060116 = Nothing
End Sub

Private Sub PatentShow()

   For Each oText In txtPA
      oText.Text = pa(oText.Index)
      Select Case oText.Index
          'X,Y編號
          Case 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134
               oText.Text = ChangeCustomerS(oText.Text)
               txtPA_Validate oText.Index, False
               pa(oText.Index) = oText.Text
      End Select
   Next
   
   Call SetCombo1("P")
   SSTab1.TabVisible(2) = True
   SSTab1.TabVisible(3) = True
   
   'Added by Lydia 2020/02/17 預設「名稱有特殊字」
   If pa(174) = "Y" Then
        lblPA174.Visible = True
        CmdPA174.Visible = True
   End If
   'end 2020/02/17
End Sub

Private Sub SetCombo1(ByVal Typ As String)
Dim strAA As String
Dim j As Integer, i As Integer
   strAA = txtPA(75)
   If strAA <> "" Then
      Select Case strAA
         Case 1
            strExc(0) = "FA07,FA52"
         Case 2
            strExc(0) = "FA08,FA53"
         Case 3
            strExc(0) = "FA09,FA54"
         Case Else
            strExc(0) = "FA08,FA53"
      End Select
      
      strExc(0) = "SELECT " & strExc(0) & " FROM FAGENT WHERE " & ChgFagent(strAA)
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields(0)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(0)
         End If
         Combo1(0).AddItem strAA & "-1" & strExc(0)
         Combo1(1).AddItem strAA & "-1" & strExc(0)
         If IsNull(RsTemp.Fields(1)) Then
            strExc(0) = ""
         Else
            strExc(0) = "-" & RsTemp.Fields(1)
         End If
         Combo1(0).AddItem strAA & "-2" & strExc(0)
         Combo1(1).AddItem strAA & "-2" & strExc(0)
      End If
   Else
      If Typ = "P" Then
          For ii = 26 To 30
              If pa(ii) <> "" Then
                  Select Case pa(85)
                     Case 1
                        strExc(0) = "CU58,CU61"
                     Case 2
                        strExc(0) = "CU59,CU62"
                     Case 3
                        strExc(0) = "CU60,CU63"
                     Case Else
                        strExc(0) = "CU59,CU62"
                  End Select
                  strExc(0) = "SELECT " & strExc(0) & " FROM CUSTOMER WHERE " & ChgCustomer(pa(ii))
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     For j = 1 To 2
                        If IsNull(RsTemp.Fields(j - 1)) Then
                           strExc(0) = ""
                        Else
                           strExc(0) = "-" & RsTemp.Fields(j - 1)
                        End If
                        Combo1(0).AddItem pa(ii) & "-" & j & strExc(0)
                        Combo1(1).AddItem pa(ii) & "-" & j & strExc(0)
                     Next
                  End If
              End If
          Next
      Else
          For i = 8 To 66
              Select Case i
                  Case 8, 58, 59, 65, 66
                     If pa(i) <> "" Then
                        strExc(0) = "SELECT CU59,CU62 FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If IsNull(RsTemp.Fields(0)) Then
                              strExc(0) = ""
                           Else
                              strExc(0) = "-" & RsTemp.Fields(0)
                           End If
                           Combo1(0).AddItem pa(i) & "-1" & strExc(0)
                           If IsNull(RsTemp.Fields(1)) Then
                              strExc(0) = ""
                           Else
                              strExc(0) = "-" & RsTemp.Fields(1)
                           End If
                           Combo1(0).AddItem pa(i) & "-2" & strExc(0)
                        End If
                     End If
              End Select
          Next
      End If
      
   End If
   If Combo1(0).ListCount > 0 And txtPA(51) = "" And txtPA(52) = "" And txtPA(53) = "" Then Combo1(0).ListIndex = 0
   If Combo1(1).ListCount > 0 And txtPA(54) = "" And txtPA(55) = "" And txtPA(56) = "" Then Combo1(1).ListIndex = 0
   
End Sub
Private Sub ServiceShow()
 Dim i As Integer
   
   For Each oText In txtPA
      Select Case oText.Index
          Case 26
               oText.Tag = "8" 'sp(8)
          Case 75
               oText.Tag = "26" 'sp(26)
          Case 91
               oText.Tag = "18" 'sp(18)
          Case 77
               oText.Tag = "27" 'sp(27)
          Case 159
               oText.Tag = "84" 'sp(84)
          Case 88
               oText.Tag = "37" 'sp(37)
          Case 48
               oText.Tag = "29" 'sp(29)
          Case 133
               oText.Tag = "67" 'sp(67)
          Case 52
               oText.Tag = "30" 'sp(30)
          Case 55
               oText.Tag = "75" 'sp(75)
          Case 1, 2, 3, 4, 5, 6, 7 '同樣欄位
               oText.Tag = oText.Index
      End Select
      If oText.Tag <> "" Then
          oText.Text = pa(Val(oText.Tag))
          Select Case oText.Index
              'X,Y編號
              Case 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134
                   oText.Text = ChangeCustomerS(oText.Text)
                   txtPA_Validate oText.Index, False
                   pa(Val(oText.Tag)) = oText.Text
          End Select
      Else
          oText.Enabled = False
      End If
   Next
   
   Call SetCombo1("S")
   SSTab1.TabVisible(2) = False
   SSTab1.TabVisible(3) = False

End Sub

Private Function FormSave() As Boolean
   Dim stUpdates As String
 
On Error GoTo CheckingErr

   cnnConnection.BeginTrans
   
   stUpdates = ""
   Select Case txtPA(1)
      Case "FCP", "P", "CFP"
            For Each oText In txtPA
                If oText.Index > 6 And pa(oText.Index) <> oText.Text Then
                   pa(oText.Index) = oText.Text
                   If InStr("26,27,28,29,30,75,76,86,88,101,105,133,134", oText.Index) > 0 And oText.Index >= 26 Then
                      pa(oText.Index) = ChangeCustomerL(oText.Text)
                   End If
                   stUpdates = stUpdates & ",pa" & IIf(oText.Index < 100, Right("00" & Trim(oText.Index), 2), Trim(oText.Index)) & "=" & CNULL(ChgSQL(pa(oText.Index)))
                End If
            Next
            If stUpdates <> "" Then
               stUpdates = Mid(stUpdates, 2)
               strSql = "UPDATE PATENT SET " & stUpdates & " WHERE PA01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
               'Modified by Lydia 2019/10/04 記錄詳細=True
               'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
               'Pub_SeekTbLog strSql, , True
               'Modified by Lydia 2025/10/31 改用模組判斷
               'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
               Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql, intI
            End If
      Case "FG", "PS", "CPS"
            For Each oText In txtPA
                If oText.Tag <> "" Then
                    If oText.Index > 6 And pa(oText.Tag) <> oText.Text Then
                       pa(oText.Tag) = oText.Text
                       If InStr("26,27,28,29,30,75,76,86,88,101,105,133,134", oText.Index) > 0 And oText.Index >= 26 Then
                          pa(oText.Tag) = ChangeCustomerL(oText.Text)
                       End If
                       stUpdates = stUpdates & ",sp" & IIf(oText.Tag < 100, Right("00" & Trim(oText.Tag), 2), Trim(oText.Tag)) & "=" & CNULL(ChgSQL(pa(oText.Tag)))
                    End If
                End If
            Next
            If stUpdates <> "" Then
               stUpdates = Mid(stUpdates, 2)
               strSql = "UPDATE SERVICEPRACTICE SET " & stUpdates & " WHERE SP01='" & pa(1) & "' and SP02='" & pa(2) & "' and SP03='" & pa(3) & "' and SP04='" & pa(4) & "'"
               'Modified by Lydia 2019/10/04 記錄詳細=True
               'Modified by Lydia 2021/04/27 更新來源的表單名稱 ;
               'Pub_SeekTbLog strSql, , True
               'Modified by Lydia 2025/10/31 改用模組判斷
               'Pub_SeekTbLog strSql, , True, , Me.Caption & "(" & Me.Name & ")"
               Pub_SeekTbLog strSql, , PUB_FilterSeekSQL(strSql), , Me.Caption & "(" & Me.Name & ")"
               cnnConnection.Execute strSql, intI
            End If
   End Select
   
   cnnConnection.CommitTrans
   FormSave = True

CheckingErr:

   If Err.Number <> 0 Then
      cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
End Function

' 清除資料表
Private Sub FormClear()

   For Each oText In txtPA
      oText.Text = ""
      oText.Tag = ""
      oText.Enabled = True
   Next

   For Each oLabel In Label27
      oLabel.Caption = ""
   Next
   
   txtPA07 = "": txtPA07.Tag = ""
   txtPA(1) = strSysKind
   SSTab1.TabVisible(2) = True
   SSTab1.TabVisible(3) = True
   txtPA(5).Locked = True
   txtPA(6).Locked = True
   txtPA07.Locked = True
   
   Combo1(1).Enabled = True
   Combo1(0).Clear
   Combo1(1).Clear
   Combo1(0).AddItem ""
   Combo1(1).AddItem ""

   cmdOK(0).Enabled = False
   cmdOK(1).Enabled = False
   cmdOK(3).Enabled = False
   
   'Added by Lydia 2020/02/17 預設「名稱有特殊字」
   lblPA174.Visible = False
   CmdPA174.Visible = False
   
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

   TxtValidate = False
   For Each oText In txtPA
       If oText.Enabled = True Then
          Cancel = False
          txtPA_Validate oText.Index, Cancel
          If Cancel = True Then
             Exit Function
          End If
       End If
   Next
   
   If txtPA(1) = "" Or txtPA(2) = "" Or txtPA(3) = "" Or txtPA(4) = "" Then
      MsgBox "本所案號錯誤，儲存失敗 !", vbCritical
      Exit Function
   End If
    
    'Added by Lydia 2021/04/14 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True) = False Then
        Exit Function
    End If
    'end 2021/04/14
    
   TxtValidate = True
End Function

Private Sub txtPA_GotFocus(Index As Integer)
  'Modified by Lydia 2018/11/21 取消備註反白
  'TextInverse txtPA(Index)
  If Index <> 91 Then TextInverse txtPA(Index)
   Select Case Index
       Case 1, 2, 3, 4, 48, 52, 55, 99, 77, 159, 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134, 87, 102, 103, 104
            'CloseIme 'Remove by Lydia 2018/11/28 Lina反應還是有輸入法無效的狀況
       Case Else
            'OpenIme 'Memo by Lydia 2018/11/23 原本11/21取消,後來Lina先切換成新注音2010輸入法,查完案號後直接到案件備註時,輸入法無效正常使用
   End Select

End Sub

'Added by Lydia 2018/11/21
Private Sub txtPA_KeyDown(Index As Integer, KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    Call PUB_HandleForm2TextBox(Me.txtPA(Index), Command1, KeyCode, Shift)  '模組化-統一控制
    
End Sub

'Added by Lydia 2018/11/21
Private Sub txtPA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    If SyxMsg <> "txtPA_" & Format(Index, "00") Then '避免連續產生訊息
        bolMsgRight = False
        SyxMsg = "txtPA_" & Format(Index, "00")
    End If
    Call PUB_HandleForm2TextBoxR(Button, Shift, bolMsgRight) '模組化-統一控制
    
End Sub

'Modified by Lydia 2018/11/19 改成Form2.0
'Private Sub txtPA_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub txtPA_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
   
End Sub

Private Sub txtPA_Validate(Index As Integer, Cancel As Boolean)
Dim strTempName As String, strTxt As String
Dim j As Integer
    Select Case Index
        Case 1
            If txtPA(Index) <> "" Then
                If txtPA(Index) <> "FCP" And txtPA(Index) <> "FG" And txtPA(Index) <> "P" And txtPA(Index) <> "PS" And txtPA(Index) <> "CFP" And txtPA(Index) <> "CPS" Then
                   MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
                   txtPA(Index).SetFocus
                   Cancel = True
                End If
            End If
        'X,Y編號
        Case 26, 27, 28, 29, 30, 75, 76, 86, 88, 101, 105, 133, 134
            If txtPA(Index).Text <> "" Then
                strTxt = txtPA(Index).Text
                If ClsLawLawGetName(strTxt, strTempName) Then
                    Label27(Index) = strTempName
                    txtPA(Index).Text = strTxt
                Else
                    txtPA(Index).SetFocus
                    Cancel = True
                End If
            Else
                Label27(Index) = ""
            End If
        Case Else
    End Select
   If Not CheckLengthIsOK(txtPA(Index), txtPA(Index).MaxLength) Then
      txtPA(Index).SetFocus
      Cancel = True
   End If
   
   If Index = 7 Then txtPA07 = txtPA(Index)
End Sub

'Added by Lydia 2020/02/17 外專：案件名稱有特殊字，開啟FCP0xxxxx.新案性質.案件名稱.doc
Private Sub CmdPA174_Click()
    
    If pa(1) = "" Or pa(2) = "" Or pa(3) = "" Or pa(4) = "" Then Exit Sub
    If Pub_GetPA174toFile("0", pa(1), pa(2), pa(3), pa(4), Me, frm100101_M_1) = True Then
    End If
    
End Sub

