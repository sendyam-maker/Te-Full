VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110104_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "更換FC代理人作業－專利"
   ClientHeight    =   5892
   ClientLeft      =   348
   ClientTop       =   1440
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5892
   ScaleWidth      =   8952
   Begin TabDlg.SSTab SSTab1 
      Height          =   3948
      Left            =   12
      TabIndex        =   44
      Top             =   1920
      Width           =   8870
      _ExtentX        =   15642
      _ExtentY        =   6964
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm110104_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(62)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(65)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(163)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(162)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(55)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(32)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(34)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(45)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPA58_T"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblPA108_T"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblPA108"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblPA58"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(169)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(47)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(154)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(35)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(48)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(155)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(64)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(158)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(36)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(37)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblPA88"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblPA133"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblPA76"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblPA134"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text1(48)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text1(77)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text1(152)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text1(151)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text1(88)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text1(50)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text1(49)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text1(133)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text1(159)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text1(71)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text1(134)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text1(78)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Text1(70)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Text1(135)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Text1(76)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text1(143)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "Text1(156)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "Text1(89)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text1(90)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text1(146)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "代理人／聯絡人"
      TabPicture(1)   =   "frm110104_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "NewFagent"
      Tab(1).Control(1)=   "Text1(100)"
      Tab(1).Control(2)=   "Text1(98)"
      Tab(1).Control(3)=   "Text1(99)"
      Tab(1).Control(4)=   "Text1(53)"
      Tab(1).Control(5)=   "Text1(54)"
      Tab(1).Control(6)=   "Text1(55)"
      Tab(1).Control(7)=   "Text1(56)"
      Tab(1).Control(8)=   "Text1(51)"
      Tab(1).Control(9)=   "Text1(52)"
      Tab(1).Control(10)=   "Text1(139)"
      Tab(1).Control(11)=   "lblAgent"
      Tab(1).Control(12)=   "Label1(77)"
      Tab(1).Control(13)=   "Label1(78)"
      Tab(1).Control(14)=   "Label1(79)"
      Tab(1).Control(15)=   "Label7"
      Tab(1).Control(16)=   "Label6"
      Tab(1).Control(17)=   "Label49"
      Tab(1).Control(18)=   "Label51"
      Tab(1).Control(19)=   "Label55"
      Tab(1).Control(20)=   "Label57"
      Tab(1).Control(21)=   "Label59"
      Tab(1).Control(22)=   "Label4"
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "其他"
      TabPicture(2)   =   "frm110104_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(87)"
      Tab(2).Control(1)=   "Text1(86)"
      Tab(2).Control(2)=   "Text1(147)"
      Tab(2).Control(3)=   "Text1(153)"
      Tab(2).Control(4)=   "Text1(154)"
      Tab(2).Control(5)=   "Text1(155)"
      Tab(2).Control(6)=   "Text1(142)"
      Tab(2).Control(7)=   "Text1(107)"
      Tab(2).Control(8)=   "Text1(106)"
      Tab(2).Control(9)=   "Text1(105)"
      Tab(2).Control(10)=   "lblPA86"
      Tab(2).Control(11)=   "lblPA105"
      Tab(2).Control(12)=   "Label10"
      Tab(2).Control(13)=   "Label2"
      Tab(2).Control(14)=   "Label1(159)"
      Tab(2).Control(15)=   "Label1(164)"
      Tab(2).Control(16)=   "Label1(165)"
      Tab(2).Control(17)=   "Label1(166)"
      Tab(2).Control(18)=   "Label68"
      Tab(2).Control(19)=   "Label1(84)"
      Tab(2).Control(20)=   "Label1(85)"
      Tab(2).Control(21)=   "Label1(86)"
      Tab(2).ControlCount=   22
      TabCaption(3)   =   "參考備註"
      TabPicture(3)   =   "frm110104_3.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdIns"
      Tab(3).Control(1)=   "Text1(91)"
      Tab(3).ControlCount=   2
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   330
         Left            =   -74880
         TabIndex        =   47
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox NewFagent 
         Height          =   300
         Left            =   -73440
         MaxLength       =   9
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   146
         Left            =   5505
         TabIndex        =   12
         Top             =   2061
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   90
         Left            =   7410
         TabIndex        =   3
         Top             =   897
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   89
         Left            =   6096
         TabIndex        =   1
         Top             =   612
         Width           =   252
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   156
         Left            =   1740
         TabIndex        =   13
         Top             =   2352
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   87
         Left            =   -73200
         TabIndex        =   40
         Top             =   3240
         Width           =   5910
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "10425;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   86
         Left            =   -73200
         TabIndex        =   39
         Top             =   2928
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   1710
         Index           =   91
         Left            =   -74910
         TabIndex        =   49
         Top             =   780
         Width           =   8505
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "15002;3016"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   143
         Left            =   5505
         TabIndex        =   17
         Top             =   2934
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   76
         Left            =   1380
         TabIndex        =   16
         Top             =   2934
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   147
         Left            =   -73200
         TabIndex        =   34
         Top             =   1398
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   153
         Left            =   -73200
         TabIndex        =   35
         Top             =   1704
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   154
         Left            =   -73200
         TabIndex        =   36
         Top             =   2010
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   155
         Left            =   -73200
         TabIndex        =   38
         Top             =   2622
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   142
         Left            =   -73200
         TabIndex        =   37
         Top             =   2316
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   135
         Left            =   1380
         TabIndex        =   18
         Top             =   3225
         Width           =   3960
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "6985;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   107
         Left            =   -73200
         TabIndex        =   33
         Top             =   1092
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   106
         Left            =   -73200
         TabIndex        =   32
         Top             =   786
         Width           =   3795
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "6694;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   105
         Left            =   -73200
         TabIndex        =   31
         Top             =   480
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   70
         Left            =   1740
         TabIndex        =   11
         Top             =   2061
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   78
         Left            =   5505
         TabIndex        =   10
         Top             =   1770
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   134
         Left            =   5505
         TabIndex        =   15
         Top             =   2643
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   71
         Left            =   1740
         TabIndex        =   9
         Top             =   1770
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   159
         Left            =   1980
         TabIndex        =   19
         Top             =   3510
         Width           =   3360
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "5927;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   100
         Left            =   -68830
         TabIndex        =   29
         Top             =   3000
         Width           =   2400
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4233;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   98
         Left            =   -73440
         TabIndex        =   28
         Top             =   3000
         Width           =   2400
         VariousPropertyBits=   671105051
         MaxLength       =   10
         Size            =   "4233;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   99
         Left            =   -73440
         TabIndex        =   30
         Top             =   3330
         Width           =   4695
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "8281;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   133
         Left            =   1740
         TabIndex        =   14
         Top             =   2643
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   49
         Left            =   1380
         TabIndex        =   4
         Top             =   1188
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   50
         Left            =   3350
         TabIndex        =   5
         Top             =   1188
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   88
         Left            =   1380
         TabIndex        =   8
         Top             =   1479
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1720;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   151
         Left            =   5050
         TabIndex        =   6
         Top             =   1188
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   152
         Left            =   6760
         TabIndex        =   7
         Top             =   1188
         Width           =   300
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "529;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   77
         Left            =   1020
         TabIndex        =   0
         Top             =   612
         Width           =   3450
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "6085;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   48
         Left            =   1380
         TabIndex        =   2
         Top             =   897
         Width           =   3975
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "7011;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   53
         Left            =   -73440
         TabIndex        =   23
         Top             =   1350
         Width           =   7000
         VariousPropertyBits=   671105051
         Size            =   "12347;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   54
         Left            =   -73440
         TabIndex        =   24
         Top             =   1680
         Width           =   3600
         VariousPropertyBits=   671105051
         Size            =   "6350;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   55
         Left            =   -73440
         TabIndex        =   25
         Top             =   2010
         Width           =   4695
         VariousPropertyBits=   671105051
         Size            =   "8281;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   56
         Left            =   -73440
         TabIndex        =   26
         Top             =   2340
         Width           =   7000
         VariousPropertyBits=   671105051
         Size            =   "12347;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   51
         Left            =   -73440
         TabIndex        =   21
         Top             =   690
         Width           =   3600
         VariousPropertyBits=   671105051
         Size            =   "6350;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   52
         Left            =   -73440
         TabIndex        =   22
         Top             =   1020
         Width           =   4695
         VariousPropertyBits=   671105051
         Size            =   "8281;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   300
         Index           =   139
         Left            =   -73440
         TabIndex        =   27
         Top             =   2670
         Width           =   7000
         VariousPropertyBits=   671105051
         Size            =   "12347;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA134 
         Height          =   255
         Left            =   6540
         TabIndex        =   108
         Top             =   2670
         Width           =   2025
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA134"
         Size            =   "3572;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA86 
         Height          =   255
         Left            =   -72150
         TabIndex        =   107
         Top             =   2970
         Width           =   5760
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA86"
         Size            =   "10160;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA105 
         Height          =   255
         Left            =   -72150
         TabIndex        =   106
         Top             =   510
         Width           =   5760
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA105"
         Size            =   "10160;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAgent 
         Height          =   255
         Left            =   -72390
         TabIndex        =   105
         Top             =   390
         Width           =   5850
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Size            =   "10319;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA76 
         Height          =   255
         Left            =   2430
         TabIndex        =   104
         Top             =   2955
         Width           =   1215
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA76"
         Size            =   "2143;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA133 
         Height          =   255
         Left            =   2790
         TabIndex        =   103
         Top             =   2670
         Width           =   1035
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA133"
         Size            =   "1826;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblPA88 
         Height          =   255
         Left            =   2430
         TabIndex        =   102
         Top             =   1500
         Width           =   5985
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblPA88"
         Size            =   "10557;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "後續准駁簡單報告：      (Y:核准以及C類來函簡單報告)"
         Height          =   180
         Index           =   37
         Left            =   4536
         TabIndex        =   98
         Top             =   624
         Width           =   4284
      End
      Begin VB.Label Label1 
         Caption         =   "信函是否列印Title：         (Y:印)"
         Height          =   255
         Index           =   36
         Left            =   5850
         TabIndex        =   97
         Top             =   920
         Width           =   2505
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "C類收文是否請款：　         (N:否)"
         Height          =   255
         Index           =   158
         Left            =   3750
         TabIndex        =   96
         Top             =   2084
         Width           =   2610
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 年費特殊管制：          (Y:年費續辦:有別於Y / X設定  N:寄證書/二核後年費不續辦  空白:視Y / X設定)"
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   95
         Top             =   2376
         Width           =   8070
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   255
         Left            =   -74580
         TabIndex        =   94
         Top             =   3263
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   255
         Left            =   -74580
         TabIndex        =   93
         Top             =   2950
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費申請人是否出名：       (N:不出名)"
         Height          =   255
         Index           =   3
         Left            =   3765
         TabIndex        =   92
         Top             =   2957
         Width           =   2940
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費代理人："
         Height          =   255
         Index           =   64
         Left            =   240
         TabIndex        =   91
         Top             =   2952
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "帳單備註是否提醒：        (N:否)"
         Height          =   255
         Index           =   159
         Left            =   -74760
         TabIndex        =   90
         Top             =   1425
         Width           =   2445
      End
      Begin VB.Label Label1 
         Caption         =   "定稿份數："
         Height          =   255
         Index           =   164
         Left            =   -74760
         TabIndex        =   89
         Top             =   1730
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "請款單份數："
         Height          =   255
         Index           =   165
         Left            =   -74760
         TabIndex        =   88
         Top             =   2035
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 同時寄紙本：          (Y：是)"
         Height          =   255
         Index           =   166
         Left            =   -74760
         TabIndex        =   87
         Top             =   2645
         Width           =   2580
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "以 Email 通知：                 (Y：是   D：僅D/N）"
         Height          =   255
         Left            =   -74760
         TabIndex        =   86
         Top             =   2340
         Width           =   3600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費聯絡人："
         Height          =   255
         Index           =   155
         Left            =   240
         TabIndex        =   85
         Top             =   3240
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費請款對象："
         Height          =   255
         Index           =   84
         Left            =   -74760
         TabIndex        =   84
         Top             =   510
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費單筆不跑：                (Y:不跑)"
         Height          =   255
         Index           =   85
         Left            =   -74760
         TabIndex        =   83
         Top             =   1120
         Width           =   2625
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費彼所案號："
         Height          =   255
         Index           =   86
         Left            =   -74760
         TabIndex        =   82
         Top             =   815
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 年費自動代繳：          (Y:自動代繳)"
         Height          =   255
         Index           =   48
         Left            =   120
         TabIndex        =   81
         Top             =   2088
         Width           =   3060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCP 領證自動代繳：          (Y:自動代繳)"
         Height          =   255
         Index           =   35
         Left            =   120
         TabIndex        =   80
         Top             =   1800
         Width           =   3060
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費D/N列印對象："
         Height          =   255
         Index           =   154
         Left            =   3960
         TabIndex        =   79
         Top             =   2666
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：         (Y:是)"
         Height          =   255
         Index           =   47
         Left            =   3765
         TabIndex        =   78
         Top             =   1793
         Width           =   2595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   255
         Index           =   169
         Left            =   150
         TabIndex        =   77
         Top             =   3533
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(中)："
         Height          =   255
         Index           =   77
         Left            =   -74805
         TabIndex        =   76
         Top             =   3026
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(英)："
         Height          =   255
         Index           =   78
         Left            =   -74805
         TabIndex        =   75
         Top             =   3353
         Width           =   1380
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "實體聯絡人(日)："
         Height          =   255
         Index           =   79
         Left            =   -70200
         TabIndex        =   74
         Top             =   3023
         Width           =   1380
      End
      Begin VB.Label lblPA58 
         AutoSize        =   -1  'True
         Caption         =   "lblPA58"
         Height          =   252
         Left            =   1080
         TabIndex        =   73
         Top             =   360
         Width           =   576
      End
      Begin VB.Label lblPA108 
         AutoSize        =   -1  'True
         Caption         =   "lblPA108"
         Height          =   255
         Left            =   5400
         TabIndex        =   72
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "新FC代理人："
         Height          =   255
         Left            =   -74805
         TabIndex        =   70
         Top             =   410
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   69
         Top             =   1391
         Width           =   1110
      End
      Begin VB.Label lblPA108_T 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日期："
         Height          =   255
         Left            =   4080
         TabIndex        =   68
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label lblPA58_T 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   252
         Left            =   108
         TabIndex        =   67
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣：       %"
         Height          =   255
         Index           =   45
         Left            =   2085
         TabIndex        =   66
         Top             =   1211
         Width           =   1755
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   255
         Index           =   34
         Left            =   105
         TabIndex        =   65
         Top             =   1512
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣：         %"
         Height          =   255
         Index           =   32
         Left            =   465
         TabIndex        =   64
         Top             =   1224
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   255
         Index           =   55
         Left            =   165
         TabIndex        =   63
         Top             =   2664
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "領證折扣：       %"
         Height          =   255
         Index           =   162
         Left            =   4185
         TabIndex        =   62
         Top             =   1211
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "年費折扣：       %"
         Height          =   255
         Index           =   163
         Left            =   5895
         TabIndex        =   61
         Top             =   1211
         Width           =   1350
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   252
         Index           =   65
         Left            =   108
         TabIndex        =   60
         Top             =   648
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   255
         Index           =   62
         Left            =   105
         TabIndex        =   59
         Top             =   936
         Width           =   1260
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   58
         Top             =   737
         Width           =   1110
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   57
         Top             =   1064
         Width           =   1110
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   56
         Top             =   1718
         Width           =   1110
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   55
         Top             =   2045
         Width           =   1110
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   54
         Top             =   2372
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   255
         Left            =   -74805
         TabIndex        =   53
         Top             =   2699
         Width           =   1380
      End
   End
   Begin VB.TextBox textPA11 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   3900
      Locked          =   -1  'True
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   510
      Width           =   1770
   End
   Begin VB.TextBox textPAKey 
      BorderStyle     =   0  '沒有框線
      Height          =   300
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   510
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6660
      TabIndex        =   42
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5700
      TabIndex        =   41
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7860
      TabIndex        =   43
      Top             =   60
      Width           =   912
   End
   Begin MSForms.ComboBox cmbPA05 
      Height          =   300
      Left            =   1200
      TabIndex        =   101
      Top             =   1500
      Width           =   7500
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13229;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA26 
      Height          =   300
      Left            =   1200
      TabIndex        =   100
      TabStop         =   0   'False
      Top             =   1170
      Width           =   7500
      VariousPropertyBits=   671105055
      Size            =   "13229;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA75 
      Height          =   300
      Left            =   1200
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   840
      Width           =   7500
      VariousPropertyBits=   671105055
      Size            =   "13229;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   255
      Left            =   120
      TabIndex        =   71
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "原FC代理人："
      Height          =   255
      Left            =   120
      TabIndex        =   52
      Top             =   870
      Width           =   1110
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   255
      Left            =   2940
      TabIndex        =   51
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   50
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   1550
      Width           =   900
   End
End
Attribute VB_Name = "frm110104_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ;textPA75、textPA26、cmbPA05、lblPA88、lblPA133、lblPA76、lblAgent、lblPA105、lblPA86、lblPA134、Text1(index)
'2011/3/29 新增 BY SONIA
Option Explicit

' 本所案號
Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存基本檔檔案欄位的串列
Dim m_PASPList() As FIELDITEM
Dim m_PASPCount As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
Dim m_PA75 As String
Dim rsDefineSize As New ADODB.Recordset

Private Sub cmdCancel_Click()
   frm110104_1.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm110104_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   If CheckDataValid = True Then
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      'Added by Lydia 2022/07/04 FCP和FMP案之一案兩請僅其中一案更代時，彈提醒：
      If PUB_ChkFCforChange(m_PA01, m_PA02, m_PA03, m_PA04) = False Then
          bolLeave = True  '直接關閉
          GoTo JumpToNext
      End If
      'end 2022/07/04
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Added by Lydia 2019/12/19
      If Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51" Then
            strExc(1) = "Y"
            If Pub_StrUserSt03 = "M51" Then
                If MsgBox("是否發案件清單通知外專承辦和程序人員？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbNo Then
                    strExc(1) = ""
                End If
            End If
            If strExc(1) = "Y" Then
                'Modified by Lydia 2022/02/24 移到basFunction
                'If frm110104_2.PUB_ChgPA75List(Replace(textPAKey, "-", "")) = False Then
                strExc(2) = "智權人員：" & frm110104_1.txtCaseField(4) & " " & frm110104_1.lblSname & vbCrLf & _
                         "變更案件條件：代理人：" & frm110104_1.txtCaseField(1) & " " & frm110104_1.lblAgent & vbCrLf & _
                         "　　　　　　　申請人：" & frm110104_1.txtCaseField(2) & " " & frm110104_1.lblCustomer & vbCrLf & _
                         "新代理人：" & frm110104_1.txtCaseField(3) & " " & frm110104_1.NewAgent & vbCrLf & _
                         "　　　　　　　" & IIf(frm110104_1.Check1.Value = 1, "■", "□") & "含閉卷或銷卷案件　　　　　　" & IIf(frm110104_1.Check4.Value = 1, "■", "□") & "清除案件聯絡人資料" & vbCrLf & _
                         "　　　　　　　" & IIf(frm110104_1.Check2.Value = 1, "■", "□") & "彼所案號清除　　　　　　　　" & IIf(frm110104_1.Check3.Value = 1, "■", "□") & "案件聯絡人同時更改"
                If PUB_ChgPA75List(Replace(textPAKey, "-", ""), "0", "", strSrvDate(1), strExc(2)) = False Then
                'end 2022/02/24
                    MsgBox "發案件清單作業失敗！", vbCritical
                End If
            End If
      End If
      'end 2019/12/19
      
JumpToNext: 'Added by Lydia 2022/07/04
      frm110104_1.Show
      frm110104_1.Cleartxt
      Unload Me
   End If

   Exit Sub
  
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textPAKey.BackColor = &H8000000F
   textPA11.BackColor = &H8000000F
   textPA75.BackColor = &H8000000F
   textPA26.BackColor = &H8000000F
   MoveFormToCenter Me
   SSTab1.Tab = 0
   strExc(0) = "SELECT * FROM PATENT WHERE ROWNUM<1"
   intI = 1
   Set rsDefineSize = ClsLawReadRstMsg(intI, strExc(0))
         
   bolLeave = False
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      Text1(91).Top = 360
      Text1(91).Height = 2970
   End If
   'end 2020/05/05
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If bolLeave = False And NewFagent <> "" Then
      If MsgBox("你並未存檔，確定離開嗎?", vbYesNo + vbCritical + vbDefaultButton2) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2022/07/04
   
   Set frm110104_3 = Nothing
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_PA01 = Empty
      m_PA02 = Empty
      m_PA03 = Empty
      m_PA04 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_PA01 = strData
      ' 本所案號 欄位2
      Case 1: m_PA02 = strData
      ' 本所案號 欄位3
      Case 2: m_PA03 = strData
      ' 本所案號 欄位4
      Case 3: m_PA04 = strData
   End Select
End Sub

' 清除案件基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_PASPCount > 0 Then
      Erase m_PASPList
   End If
   m_PASPCount = 0
End Sub

' 設定案件基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
Dim nPos As Integer
Dim bFind As Boolean
   
   bFind = False
   For nPos = 0 To m_PASPCount - 1
      If m_PASPList(nPos).fiName = strFieldName Then
         bFind = True
         m_PASPList(nPos).fiOldData = strFieldData
         m_PASPList(nPos).fiNewData = strFieldData
         m_PASPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_PASPList(m_PASPCount + 1)
      m_PASPList(m_PASPCount).fiName = strFieldName
      m_PASPList(m_PASPCount).fiOldData = strFieldData
      m_PASPList(m_PASPCount).fiNewData = strFieldData
      m_PASPList(m_PASPCount).fiType = nFieldType '0.文字 1.數字
      m_PASPCount = m_PASPCount + 1
   End If
End Sub

' 設定案件基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
Dim nPos As Integer
Dim bFind As Boolean
   
   bFind = False
   For nPos = 0 To m_PASPCount - 1
      If m_PASPList(nPos).fiName = strFieldName Then
         bFind = True
         m_PASPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryPatent()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String 'Add By Sindy 2014/12/2
   
   strSql = "SELECT * FROM patent " & _
            "WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   cmbPA05.Clear
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then: textPA11 = rsTmp.Fields("PA11")
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("PA05")) = False Then: cmbPA05.AddItem rsTmp.Fields("PA05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("PA06")) = False Then: cmbPA05.AddItem rsTmp.Fields("PA06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("PA07")) = False Then: cmbPA05.AddItem rsTmp.Fields("PA07")
      ' 顯示案件名稱
      If cmbPA05.ListCount > 0 Then
         cmbPA05.ListIndex = 0
      End If
      ' 原FC代理人
      m_PA75 = ""
      If IsNull(rsTmp.Fields("PA75")) = False Then
         m_PA75 = rsTmp.Fields("PA75")
         'Add By Sindy 2014/12/2
         Call GetAgentAndState(m_PA75, strTemp, , , True)
         'textPA75 = rsTmp.Fields("PA75") & " " & GetFAgentName(rsTmp.Fields("PA75"))
         textPA75 = rsTmp.Fields("PA75") & " " & strTemp
         '2014/12/2 END
      End If
      SetTMSPFieldOldData "PA75", m_PA75, 0
      ' 申請人1
      If IsNull(rsTmp.Fields("PA26")) = False Then: textPA26 = rsTmp.Fields("PA26") & " " & GetAgentOrCustName(rsTmp.Fields("PA26"))
      ' 閉卷日期
      lblPA58 = ""
      If IsNull(rsTmp.Fields("PA58")) = False Then
         lblPA58 = ChangeWStringToTDateString(rsTmp.Fields("PA58"))
         lblPA58_T.ForeColor = &HC0&
         lblPA58.ForeColor = &HC0&
      End If
      ' 北所銷卷日期
      lblPA108 = ""
      If IsNull(rsTmp.Fields("PA108")) = False Then
         lblPA108 = ChangeWStringToTDateString(rsTmp.Fields("PA108"))
         lblPA108_T.ForeColor = &HC0&
         lblPA108.ForeColor = &HC0&
      End If
      ' 客戶案件案號
      If IsNull(rsTmp.Fields("PA48")) = False Then: Text1(48) = rsTmp.Fields("PA48")
      SetTMSPFieldOldData "PA48", Text1(48), 0
      ' 彼所案號
      If IsNull(rsTmp.Fields("PA77")) = False Then: Text1(77) = rsTmp.Fields("PA77")
      SetTMSPFieldOldData "PA77", Text1(77), 0
      ' 全部折扣
      If IsNull(rsTmp.Fields("PA49")) = False Then: Text1(49) = rsTmp.Fields("PA49")
      SetTMSPFieldOldData "PA49", Text1(49), 1
      ' 申請/翻譯折扣
      If IsNull(rsTmp.Fields("PA50")) = False Then: Text1(50) = rsTmp.Fields("PA50")
      SetTMSPFieldOldData "PA50", Text1(50), 1
      ' 領證折扣
      If IsNull(rsTmp.Fields("PA151")) = False Then: Text1(151) = rsTmp.Fields("PA151")
      SetTMSPFieldOldData "PA151", Text1(151), 1
      ' 年費折扣
      If IsNull(rsTmp.Fields("PA152")) = False Then: Text1(152) = rsTmp.Fields("PA152")
      SetTMSPFieldOldData "PA152", Text1(152), 1
      
      'Add By Sindy 2015/7/28
      ' 副本收受人
      lblPA86 = ""
      If IsNull(rsTmp.Fields("PA86")) = False Then: Text1(86) = rsTmp.Fields("PA86"): lblPA86 = GetAgentOrCustName(rsTmp.Fields("PA86"))
      SetTMSPFieldOldData "PA86", Text1(86), 0
      ' 副本聯絡人
      If IsNull(rsTmp.Fields("PA87")) = False Then: Text1(87) = rsTmp.Fields("PA87")
      SetTMSPFieldOldData "PA87", Text1(87), 0
      '2015/7/28 END
      
      ' 固定請款對象
      lblPA88 = ""
      If IsNull(rsTmp.Fields("PA88")) = False Then: Text1(88) = rsTmp.Fields("PA88"): lblPA88 = GetAgentOrCustName(rsTmp.Fields("PA88"))
      SetTMSPFieldOldData "PA88", Text1(88), 0
      ' FCP領證自動代繳
      If IsNull(rsTmp.Fields("PA71")) = False Then: Text1(71) = rsTmp.Fields("PA71")
      SetTMSPFieldOldData "PA71", Text1(71), 0
      ' FCP年費自動代繳
      If IsNull(rsTmp.Fields("PA70")) = False Then: Text1(70) = rsTmp.Fields("PA70")
      SetTMSPFieldOldData "PA70", Text1(70), 0
      'Added by Lydia 2019/11/27 FCP年費特殊管制
      If IsNull(rsTmp.Fields("PA156")) = False Then: Text1(156) = rsTmp.Fields("PA156")
      Text1(156).Tag = Text1(156).Text
      SetTMSPFieldOldData "PA156", Text1(156), 0
      'Added by Lydia 2019/12/19
      'Memo by Amy 2025/08/06  不續辦但准通知 改為 後續准駁簡單報告
      If IsNull(rsTmp.Fields("PA89")) = False Then: Text1(89) = rsTmp.Fields("PA89")
      Text1(89).Tag = Text1(89).Text
      SetTMSPFieldOldData "PA89", Text1(89), 0
      '信函是否列印Title
      If IsNull(rsTmp.Fields("PA90")) = False Then: Text1(90) = rsTmp.Fields("PA90")
      Text1(90).Tag = Text1(90).Text
      SetTMSPFieldOldData "PA90", Text1(90), 0
      'C類收文是否請款
      If IsNull(rsTmp.Fields("PA146")) = False Then: Text1(146) = rsTmp.Fields("PA146")
      Text1(146).Tag = Text1(146).Text
      SetTMSPFieldOldData "PA146", Text1(146), 0
      
      ' D/N是否列印申請人
      If IsNull(rsTmp.Fields("PA78")) = False Then: Text1(78) = rsTmp.Fields("PA78")
      SetTMSPFieldOldData "PA78", Text1(78), 0
      ' D/N固定列印對象
      lblPA133 = ""
      If IsNull(rsTmp.Fields("PA133")) = False Then: Text1(133) = rsTmp.Fields("PA133"): lblPA133 = GetAgentOrCustName(rsTmp.Fields("PA133"))
      SetTMSPFieldOldData "PA133", Text1(133), 0
      ' 年費D/N列印對象
      lblPA134 = ""
      If IsNull(rsTmp.Fields("PA134")) = False Then: Text1(134) = rsTmp.Fields("PA134"): lblPA134 = GetAgentOrCustName(rsTmp.Fields("PA134"))
      SetTMSPFieldOldData "PA134", Text1(134), 0
      ' 年費代理人
      lblPA76 = ""
      If IsNull(rsTmp.Fields("PA76")) = False Then: Text1(76) = rsTmp.Fields("PA76"): lblPA76 = GetAgentOrCustName(rsTmp.Fields("PA76"))
      SetTMSPFieldOldData "PA76", Text1(76), 0
      ' 年費聯絡人
      If IsNull(rsTmp.Fields("PA135")) = False Then: Text1(135) = rsTmp.Fields("PA135")
      SetTMSPFieldOldData "PA135", Text1(135), 0
      ' Client_Matter_id
      If IsNull(rsTmp.Fields("PA159")) = False Then: Text1(159) = rsTmp.Fields("PA159")
      SetTMSPFieldOldData "PA159", Text1(159), 0
      ' 聯絡人1(中)
      If IsNull(rsTmp.Fields("PA51")) = False Then: Text1(51) = rsTmp.Fields("PA51")
      SetTMSPFieldOldData "PA51", Text1(51), 0
      ' 聯絡人1(英)
      If IsNull(rsTmp.Fields("PA52")) = False Then: Text1(52) = rsTmp.Fields("PA52")
      SetTMSPFieldOldData "PA52", Text1(52), 0
      ' 聯絡人1(日)
      If IsNull(rsTmp.Fields("PA53")) = False Then: Text1(53) = rsTmp.Fields("PA53")
      SetTMSPFieldOldData "PA53", Text1(53), 0
      ' 聯絡人2(中)
      If IsNull(rsTmp.Fields("PA54")) = False Then: Text1(54) = rsTmp.Fields("PA54")
      SetTMSPFieldOldData "PA54", Text1(54), 0
      ' 聯絡人2(英)
      If IsNull(rsTmp.Fields("PA55")) = False Then: Text1(55) = rsTmp.Fields("PA55")
      SetTMSPFieldOldData "PA55", Text1(55), 0
      ' 聯絡人2(日)
      If IsNull(rsTmp.Fields("PA56")) = False Then: Text1(56) = rsTmp.Fields("PA56")
      SetTMSPFieldOldData "PA56", Text1(56), 0
      ' 聯絡人部門(日)
      If IsNull(rsTmp.Fields("PA139")) = False Then: Text1(139) = rsTmp.Fields("PA139")
      SetTMSPFieldOldData "PA139", Text1(139), 0
      ' 實體聯絡人(中)
      If IsNull(rsTmp.Fields("PA98")) = False Then: Text1(98) = rsTmp.Fields("PA98")
      SetTMSPFieldOldData "PA98", Text1(98), 0
      ' 實體聯絡人(英)
      If IsNull(rsTmp.Fields("PA99")) = False Then: Text1(99) = rsTmp.Fields("PA99")
      SetTMSPFieldOldData "PA99", Text1(99), 0
      ' 實體聯絡人(日)
      If IsNull(rsTmp.Fields("PA100")) = False Then: Text1(100) = rsTmp.Fields("PA100")
      SetTMSPFieldOldData "PA100", Text1(100), 0
      ' 年費請款對象
      lblPA105 = ""
      If IsNull(rsTmp.Fields("PA105")) = False Then: Text1(105) = rsTmp.Fields("PA105"): lblPA105 = GetAgentOrCustName(rsTmp.Fields("PA105"))
      SetTMSPFieldOldData "PA105", Text1(105), 0
      ' 彼所案號
      If IsNull(rsTmp.Fields("PA106")) = False Then: Text1(106) = rsTmp.Fields("PA106")
      SetTMSPFieldOldData "PA106", Text1(106), 0
      ' 年費單筆不跑
      If IsNull(rsTmp.Fields("PA107")) = False Then: Text1(107) = rsTmp.Fields("PA107")
      SetTMSPFieldOldData "PA107", Text1(107), 0
      ' 帳單備註是否提醒
      If IsNull(rsTmp.Fields("PA147")) = False Then: Text1(147) = rsTmp.Fields("PA147")
      SetTMSPFieldOldData "PA147", Text1(147), 0
      ' 定稿份數
      If IsNull(rsTmp.Fields("PA153")) = False Then: Text1(153) = rsTmp.Fields("PA153")
      SetTMSPFieldOldData "PA153", Text1(153), 1
      ' 請款單份數
      If IsNull(rsTmp.Fields("PA154")) = False Then: Text1(154) = rsTmp.Fields("PA154")
      SetTMSPFieldOldData "PA154", Text1(154), 1
      ' 以EMail通知
      If IsNull(rsTmp.Fields("PA142")) = False Then: Text1(142) = rsTmp.Fields("PA142")
      SetTMSPFieldOldData "PA142", Text1(142), 0
      ' EMail同時寄紙本
      If IsNull(rsTmp.Fields("PA155")) = False Then: Text1(155) = rsTmp.Fields("PA155")
      SetTMSPFieldOldData "PA155", Text1(155), 0
      ' 年費申請人是否出名
      If IsNull(rsTmp.Fields("PA143")) = False Then: Text1(143) = rsTmp.Fields("PA143")
      SetTMSPFieldOldData "PA143", Text1(143), 0
      ' 備註
      If IsNull(rsTmp.Fields("PA91")) = False Then: Text1(91) = rsTmp.Fields("PA91")
      SetTMSPFieldOldData "PA91", Text1(91), 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
Dim intCaseKind As Integer
   
   ' 先清除案件基本檔欄位串列
   ClearTMSPFieldList
   
   ' 本所案號
   textPAKey.Text = m_PA01 & "-" & m_PA02 & "-" & IIf(Len("" & m_PA03) <= 0, "0", m_PA03) & "-" & IIf(Len("" & m_PA04) <= 0, "00", m_PA04)
   
   ' 讀取基本檔
   QueryPatent
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
Dim intCaseKind As Integer
   
   NewFagent = NewFagent & String(9 - Len(NewFagent), "0")
   ' 更新基本檔
   ' FC代理人
   SetTMSPFieldNewData "PA75", NewFagent
   ' 客戶案件案號
   SetTMSPFieldNewData "PA48", Text1(48)
   ' 彼所案號
   SetTMSPFieldNewData "PA77", Text1(77)
   ' 全部折扣
   SetTMSPFieldNewData "PA49", Text1(49)
   ' 申請/翻譯折扣
   SetTMSPFieldNewData "PA50", Text1(50)
   ' 領證折扣
   SetTMSPFieldNewData "PA151", Text1(151)
   ' 年費折扣
   SetTMSPFieldNewData "PA152", Text1(152)
   'Add By Sindy 2015/7/28
   ' 副本收受人
   SetTMSPFieldNewData "PA86", Text1(86) & IIf(Text1(86) <> "", String(9 - Len(Text1(86)), "0"), "")
   ' 副本聯絡人
   SetTMSPFieldNewData "PA87", Text1(87)
   '2015/7/28 END
   ' 固定請款對象
   SetTMSPFieldNewData "PA88", Text1(88) & IIf(Text1(88) <> "", String(9 - Len(Text1(88)), "0"), "")
   ' FCP領證自動代繳
   SetTMSPFieldNewData "PA71", Text1(71)
   ' FCP年費自動代繳
   SetTMSPFieldNewData "PA70", Text1(70)
   'Added by Lydia 2019/11/27 FCP年費特殊管制
   SetTMSPFieldNewData "PA156", Text1(156)
   'Added by Lydia 2019/12/19
   'Memo by Amy 2025/08/06  不續辦但准通知 改為 後續准駁簡單報告
   SetTMSPFieldNewData "PA89", Text1(89)
   '信函是否列印Title
   SetTMSPFieldNewData "PA90", Text1(90)
   'C類收文是否請款
   SetTMSPFieldNewData "PA146", Text1(146)
      
   ' D/N是否列印申請人
   SetTMSPFieldNewData "PA78", Text1(78)
   ' D/N固定列印對象
   SetTMSPFieldNewData "PA133", Text1(133) & IIf(Text1(133) <> "", String(9 - Len(Text1(133)), "0"), "")
   ' 年費D/N列印對象
   SetTMSPFieldNewData "PA134", Text1(134) & IIf(Text1(134) <> "", String(9 - Len(Text1(134)), "0"), "")
   ' 年費代理人
   SetTMSPFieldNewData "PA76", Text1(76) & IIf(Text1(76) <> "", String(9 - Len(Text1(76)), "0"), "")
   ' 年費聯絡人
   SetTMSPFieldNewData "PA135", Text1(135)
   ' Client_Matter_id
   SetTMSPFieldNewData "PA159", Text1(159)
   ' 聯絡人1(中)
   SetTMSPFieldNewData "PA51", Text1(51)
   ' 聯絡人1(英)
   SetTMSPFieldNewData "PA52", Text1(52)
   ' 聯絡人1(日)
   SetTMSPFieldNewData "PA53", Text1(53)
   ' 聯絡人2(中)
   SetTMSPFieldNewData "PA54", Text1(54)
   ' 聯絡人2(英)
   SetTMSPFieldNewData "PA55", Text1(55)
   ' 聯絡人2(日)
   SetTMSPFieldNewData "PA56", Text1(56)
   ' 聯絡人部門(日)
   SetTMSPFieldNewData "PA139", Text1(139)
   ' 實體聯絡人(中)
   SetTMSPFieldNewData "PA98", Text1(98)
   ' 實體聯絡人(英)
   SetTMSPFieldNewData "PA99", Text1(99)
   ' 實體聯絡人(日)
   SetTMSPFieldNewData "PA100", Text1(100)
   ' 年費請款對象
   SetTMSPFieldNewData "PA105", Text1(105) & IIf(Text1(105) <> "", String(9 - Len(Text1(105)), "0"), "")
   ' 彼所案號
   SetTMSPFieldNewData "PA106", Text1(106)
   ' 年費單筆不跑
   SetTMSPFieldNewData "PA107", Text1(107)
   ' 帳單備註是否提醒
   SetTMSPFieldNewData "PA147", Text1(147)
   ' 定稿份數
   SetTMSPFieldNewData "PA153", Text1(153)
   ' 請款單份數
   SetTMSPFieldNewData "PA154", Text1(154)
   ' 以EMail通知
   SetTMSPFieldNewData "PA142", Text1(142)
   ' EMail同時寄紙本
   SetTMSPFieldNewData "PA155", Text1(155)
   ' 年費申請人是否出名
   SetTMSPFieldNewData "PA143", Text1(143)
   ' 備註
   'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
   'Modified by Lydia 2019/12/24 備註加「請留意最新指示及聯絡對象」
   Text1(91) = ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & m_PA75 & "/" & Mid(Trim(textPA75), 11)) & ";" & Trim(Text1(91))
   SetTMSPFieldNewData "PA91", Text1(91)
End Sub

Public Function OnSaveData() As Boolean
Dim intCaseKind As Integer
Dim bFirst As Boolean
Dim nIndex As Integer
Dim strTmp As String, strCP09 As String, strCP110 As String, strCP10 As String
Dim strErrMsg As String, strPassSql As String 'Added by Lydia 2020/03/17

On Error GoTo CheckingErr
   
   OnSaveData = True
   cnnConnection.BeginTrans
   
   ' 更新基本檔
   strCP10 = "937"
   strSql = "UPDATE Patent SET "
   bFirst = True
   For nIndex = 0 To m_PASPCount - 1
      strTmp = Empty
      If m_PASPList(nIndex).fiOldData <> m_PASPList(nIndex).fiNewData Then
         If m_PASPList(nIndex).fiType = 0 Then
            strTmp = m_PASPList(nIndex).fiName & " = '" & ChgSQL(m_PASPList(nIndex).fiNewData) & "'"
         Else
            If m_PASPList(nIndex).fiNewData = Empty Then
               strTmp = m_PASPList(nIndex).fiName & " = " & 0
            Else
               strTmp = m_PASPList(nIndex).fiName & " = " & m_PASPList(nIndex).fiNewData
            End If
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
   ' 設定SQL語法更新的條件
   strSql = strSql & _
                 " WHERE PA01 = '" & m_PA01 & "' AND " & _
                        "PA02 = '" & m_PA02 & "' AND " & _
                        "PA03 = '" & m_PA03 & "' AND " & _
                        "PA04 = '" & m_PA04 & "' "
   'Add By Sindy 2017/3/14 紀錄分析語法
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql

   'Added by Lydia 2019/11/27 FCP年費特殊管制PA165=N => 目前案件的年費期限自動上不續辦
   If Text1(156).Text = "N" And Text1(156).Tag <> Text1(156).Text Then
       'Modified by Lydia 2020/03/17 回傳FMP案範圍，發清單通知程序
       'Call Pub_AutoUpdFCP605(m_PA01 & m_PA02 & m_PA03 & m_PA04)
       If Pub_AutoUpdFCP605(m_PA01 & m_PA02 & m_PA03 & m_PA04, strPassSql, strErrMsg) = False Then
            GoTo CheckingErr
       End If
       'end 2020/03/17
   End If
   'end 2019/11/27
   
   '新增案件進度檔
   strCP09 = AutoNo("B", 6)
   '取得出名代理人
   strCP110 = ""
'CANCEL BY SONIA 2015/6/17 FCT-024182各式申請書抓最新A,B類發文之CP110會抓到此進度
'   strExc(0) = "select cp110 from caseprogress" & _
'               " where cp09=(select substr(max(cp27||cp09),9) from caseprogress" & _
'               " WHERE cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "'" & _
'               " and cp09<'C'" & _
'               " and cp110 is not null and cp27 is not null" & _
'               " group by cp01,cp02,cp03,cp04)"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      strCP110 = RsTemp.Fields(0)
'   End If
'END 2015/6/17
   'Modified by Morgan 2016/8/22 備註加 ChgSQL(代理人名稱可能有單引號)
   'Modified by Lydia 2019/12/24 備註加「請留意最新指示及聯絡對象」
   strSql = "INSERT INTO CASEPROGRESS(CP09,CP01,CP02,CP03,CP04,CP05,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP64,cp82,cp83,cp110)" & _
            " values(" & CNULL(strCP09) & "," & CNULL(m_PA01) & "," & CNULL(m_PA02) & "," & CNULL(m_PA03) & "," & CNULL(m_PA04) & _
            "," & strSrvDate(1) & "," & CNULL(strCP10) & ",'90'," & CNULL(PUB_GetStaffST15(frm110104_1.txtCaseField(4), 1)) & "," & CNULL(frm110104_1.txtCaseField(4)) & "," & CNULL(strUserNum) & ",'N','N'" & _
            "," & strSrvDate(1) & ",'" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & m_PA75 & "/" & Mid(Trim(textPA75), 11)) & ";'" & _
            ",substr(to_char(sysdate,'yyyymmddhh24mmss'),9)," & CNULL(strUserNum) & "," & CNULL(strCP110) & ")"
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   bolLeave = True
   
    'Added by Lydia 2020/03/17 FMP案件不自動上年費不續辦，改發清單給程序，由各區程序逐筆產生定稿通知大陸代理人
    If strPassSql <> "" Then
       If PUB_GetP605Email("1", strPassSql, strErrMsg) = False Then
          If strErrMsg <> "" Then
              MsgBox strErrMsg, vbCritical
          End If
       End If
    End If
    'end 2020/03/17
    
   Exit Function
   
CheckingErr:
   'Modified by Lydia 2020/03/17
   'MsgBox (Err.Description)
   MsgBox (Err.Description & vbCrLf & strErrMsg)
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   CheckDataValid = False
   
   If NewFagent = "" Then
      MsgBox "請輸入新FC代理人!!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1
      NewFagent.SetFocus
      Exit Function
   End If
   
   If m_PA75 = NewFagent Then
      MsgBox "代理人和新FC代理人不可相同 !!!", vbExclamation + vbOKOnly
      SSTab1.Tab = 1
      NewFagent.SetFocus
      Exit Function
   End If
   
   CheckDataValid = True
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean
   
   TxtValidate = False
   
   If Me.Text1(76).Enabled = True Then
      Cancel = False
      Call Text1_Validate(76, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(76).SetFocus
         Exit Function
      End If
   End If
   'Add By Sindy 2015/7/28
   If Me.Text1(86).Enabled = True Then
      Cancel = False
      Call Text1_Validate(86, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(86).SetFocus
         Exit Function
      End If
   End If
   '2015/7/28 END
   If Me.Text1(88).Enabled = True Then
      Cancel = False
      Call Text1_Validate(88, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(88).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(105).Enabled = True Then
      Cancel = False
      Call Text1_Validate(105, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(105).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(133).Enabled = True Then
      Cancel = False
      Call Text1_Validate(133, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(133).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(134).Enabled = True Then
      Cancel = False
      Call Text1_Validate(134, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(134).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(142).Enabled = True Then
      Cancel = False
      Call Text1_Validate(142, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(142).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(155).Enabled = True Then
      Cancel = False
      Call Text1_Validate(155, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 0
         Text1(155).SetFocus
         Exit Function
      End If
   End If
   If Me.Text1(91).Enabled = True Then
      Cancel = False
      Call Text1_Validate(91, Cancel)
      If Cancel = True Then
         SSTab1.Tab = 3
         Text1(91).SetFocus
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/09/22 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If

   TxtValidate = True
End Function

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
   Select Case Index
      Case 5, 7, 31, 33, 34, 36, 37, 39, 40, 42, 43, 45, 51, 53, 54, 56, 79, 81, 82, 84, 87, 91, 98, 100, 109, 111, 112, 114, 115, 117, 118, 120, 121, 123, 124, 126, 127, 129, 130, 132, 139
         OpenIme
      Case Else
         CloseIme
   End Select
End Sub

'Modified by Lydia 2021/09/22 改成Form 2.0
'Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Private Sub Text1_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
      Case 1, 26, 27, 28, 29, 30, 59, 75, 76, 86, 88, 101, 105, 133, 134, 160, 164
         KeyAscii = UpperCase(KeyAscii)
      Case 16
         If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 49, 50, 151, 152, 153, 154
         If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 23, 85
         If (KeyAscii < 49 Or KeyAscii > 51) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Modified by Lydia 2016/08/18 拿掉PA70
      'Case 18, 19, 46, 57, 70, 71, 78, 89, 107, 108
      'Modified by Lydia 2019/11/27 +FCP年費自動代繳PA70
      'Modified by Lydia 2019/12/19 +PA90
      Case 18, 19, 46, 57, 70, 71, 78, 89, 90, 107, 108
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then 'Y/null
            KeyAscii = 0
            Beep
         End If
      'Added by Lydia 2016/08/18 FCP年費自動代繳(Y)/寄證書後年費不續辦(N)
      'Modified by Lydia 2019/11/27 改成FCP年費特殊管制PA156: Y:年費續辦  N:寄證書/二核後年費不續辦  空白:視代理人/申請人設定
      'Case 70
      Case 156
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then 'Y/N/null
            KeyAscii = 0
            Beep
         End If
      'end 2016/08/18
      Case 141, 143, 146, 147
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then 'N/null
            KeyAscii = 0
            Beep
         End If
      Case 17, 157, 162, 163
         KeyAscii = UpperCase(KeyAscii)
         If Text1(1) = "P" And Index = 162 Then
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         Else
            If KeyAscii <> 89 And KeyAscii <> 78 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         End If
      Case 142, 155, 90 ', 161
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 161
         KeyAscii = UpperCase(KeyAscii)
         If strSrvDate(1) >= InvoiceStartDate Then
            'J,T,空白
            If KeyAscii <> 74 And KeyAscii <> 84 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         Else
            'Y,空白
            If KeyAscii <> 89 And KeyAscii <> 8 Then
               KeyAscii = 0
               Beep
            End If
         End If
   End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim strTmp As String, i As Integer, strTxt(10) As String
   
   If Text1(Index) = "" Then
      Select Case Index
         Case 76
            lblPA76 = ""
         'Add By Sindy 2015/7/28
         Case 86
            lblPA86 = ""
         '2015/7/28 END
         Case 88
            lblPA88 = ""
         Case 105
            lblPA105 = ""
         Case 133
            lblPA133 = ""
         Case 134
            lblPA134 = ""
      End Select
      Exit Sub
   '當有輸專用期時若准駁為空白則預設為1
   Else
      Select Case Index
         Case 142, 155
            If (Text1(142).Text = "" And Text1(155).Text = "Y") Then
               MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
               Cancel = True
               Exit Sub
            End If
      End Select
   End If
   
   '檢查中文欄位長度是否過長
   If CheckLengthIsOK(Text1(Index).Text, rsDefineSize.Fields(Index - 1).DefinedSize) Then
      Cancel = ChkKeyIn(Index)
   Else
      Cancel = True
   End If
   If Cancel = True Then TextInverse Text1(Index)
End Sub

Private Function ChkKeyIn(ByVal iSitu As Integer) As Boolean
Dim strTmp As String, strMain As String, bolChk As Boolean, i As Integer
   
   ChkKeyIn = False
   strMain = Text1(iSitu).Text
   Select Case iSitu
      Case 76
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA76 = strTmp
      Case 86
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA86 = strTmp
      Case 88
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA88 = strTmp
      Case 105
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA105 = strTmp
      Case 133
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA133 = strTmp
      Case 134
         ChkKeyIn = Not ClsLawLawGetName(strMain, strTmp)
         lblPA134 = strTmp
   End Select
End Function

Private Sub NewFagent_GotFocus()
   InverseTextBox NewFagent
End Sub

Private Sub NewFagent_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub NewFagent_Validate(Cancel As Boolean)
Dim strNo As String, strTemp As String
   
   lblAgent.Caption = ""
   If NewFagent <> "" Then
      strNo = NewFagent
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetAgentAndState(strNo, strTemp, , , True, m_PA01) Then
         NewFagent = ChangeCustomerL(strNo)
         lblAgent.Caption = strTemp
      Else
         NewFagent_GotFocus
         Cancel = True
         Exit Sub
      End If
      '若輸入9碼且最後一碼不為"0"
      If Len(NewFagent) = 9 And Right(NewFagent, 1) <> "0" Then
         MsgBox "此代理人已變更名稱，請使用新名稱之編號收文!!!", vbExclamation + vbOKOnly
         NewFagent_GotFocus
         Cancel = True
         Exit Sub
      End If
   Else
      MsgBox "請輸入新FC代理人!!!", vbExclamation + vbOKOnly
      NewFagent_GotFocus
      Cancel = True
      Exit Sub
   End If
End Sub

'Modify By Sindy 2014/12/2
' 取得客戶或是代理人名稱
Private Function GetAgentOrCustName(ByVal strData As String) As String
Dim strTemp As String
   
   GetAgentOrCustName = Empty
   If IsEmptyText(strData) = False Then
      Select Case UCase(Mid(strData, 1, 1))
         Case "X":
            'Modify By Sindy 2015/8/27 +m_PA01
            If GetCustomerAndState(strData, strTemp, , , True, m_PA01) Then
               GetAgentOrCustName = strTemp
            End If
         Case "Y":
            'Modify By Sindy 2015/8/27 +m_PA01
            If GetAgentAndState(strData, strTemp, , , True, m_PA01) Then
               GetAgentOrCustName = strTemp
            End If
      End Select
   End If
End Function
'' 取得客戶或是代理人名稱
'Private Function GetAgentOrCustName(ByVal strData As String) As String
'Dim rsTmp As ADODB.Recordset
'Dim strSql As String
'
'   GetAgentOrCustName = Empty
'   If IsEmptyText(strData) = False Then
'      ' 不滿8碼自動補0
'      If Len(strData) < 8 Then: strData = strData & String(8 - Len(strData), "0")
'      Select Case Mid(strData, 1, 1)
'      Case "X", "x":
'         Set rsTmp = New ADODB.Recordset
'         If Len(strData) > 8 Then
'            strSql = "SELECT * FROM Customer " & _
'                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
'                           "CU02 = '" & Mid(strData, 9, 1) & "'"
'         Else
'            strSql = "SELECT * FROM Customer " & _
'                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
'                           "CU02 = '0' "
'         End If
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp.RecordCount > 0 Then
'            rsTmp.MoveFirst
'            If IsNull(rsTmp.Fields("CU05")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("CU05")
'            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("CU04")
'            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("CU06")
'            End If
'         End If
'         rsTmp.Close
'      Case "Y", "y":
'         Set rsTmp = New ADODB.Recordset
'         If Len(strData) > 8 Then
'            strSql = "SELECT * FROM FAGENT " & _
'                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
'                           "FA02 = '" & Mid(strData, 9, 1) & "'"
'         Else
'            strSql = "SELECT * FROM FAGENT " & _
'                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
'                           "FA02 = '0' "
'         End If
'         rsTmp.CursorLocation = adUseClient
'         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsTmp.RecordCount > 0 Then
'            rsTmp.MoveFirst
'            If IsNull(rsTmp.Fields("FA05")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("FA05")
'            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("FA04")
'            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
'               GetAgentOrCustName = rsTmp.Fields("FA06")
'            End If
'         End If
'         rsTmp.Close
'      End Select
'   End If
'   Set rsTmp = Nothing
'End Function

'Added by Lydia 2016/11/23 各項指示
Private Sub cmdIns_Click()
   If textPAKey = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   frm12040159.SetParent "E", Trim(Replace(textPAKey, "-", "")), Me
   frm12040159.Show
    
End Sub
