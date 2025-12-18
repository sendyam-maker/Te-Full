VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_15 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(申請意見書, 補充理由, 訴願, 再訴願, 行政訴訟, 參加行政訴訟, 再審之訴)"
   ClientHeight    =   6168
   ClientLeft      =   5292
   ClientTop       =   2508
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6168
   ScaleWidth      =   9156
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5370
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   660
      Width           =   3675
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   660
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   384
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5370
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   930
      Width           =   3675
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   930
      Width           =   2532
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7044
      TabIndex        =   34
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6216
      TabIndex        =   33
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8268
      TabIndex        =   35
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Left            =   5016
      TabIndex        =   32
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3570
      Left            =   60
      TabIndex        =   66
      Top             =   2580
      Width           =   9045
      _ExtentX        =   15960
      _ExtentY        =   6287
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm020102_15.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblNameAgent"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label26"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label20"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label39"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(12)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label7"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label30"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label19"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label18"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label17"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label16"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label15"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label4"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(10)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label22"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label13"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label25"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label28"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label14"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label21"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label12"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label56"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label3"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lblPayToday"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lblCP113(18)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP14_R2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP44_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP64"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP64_R"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP49"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lstNameAgent"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM67"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtEditWord"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCP84"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textPrint"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCF09"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textNP09"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textNP08"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textNP07_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textNP07"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCP23_R"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textCP22"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textCP14_R"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textCP08"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textCP23"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textCP44"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textCP18"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textCP27"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textCP43"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textCP118"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textRecvDate"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtPayToday"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "txtCP113"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).ControlCount=   58
      TabCaption(1)   =   "對造資料"
      TabPicture(1)   =   "frm020102_15.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label54"
      Tab(1).Control(1)=   "Label32"
      Tab(1).Control(2)=   "Label52"
      Tab(1).Control(3)=   "Label53"
      Tab(1).Control(4)=   "Label33"
      Tab(1).Control(5)=   "Label49"
      Tab(1).Control(6)=   "Label50"
      Tab(1).Control(7)=   "Label51"
      Tab(1).Control(8)=   "Label55"
      Tab(1).Control(9)=   "textCP40"
      Tab(1).Control(10)=   "textCP42"
      Tab(1).Control(11)=   "textCP39"
      Tab(1).Control(12)=   "textCP37"
      Tab(1).Control(13)=   "textCP37_1"
      Tab(1).Control(14)=   "textCP36"
      Tab(1).Control(15)=   "textCP80"
      Tab(1).Control(16)=   "textCP41"
      Tab(1).Control(17)=   "textCP38"
      Tab(1).ControlCount=   18
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   5640
         MaxLength       =   4
         TabIndex        =   13
         Top             =   2067
         Width           =   540
      End
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   8235
         MaxLength       =   1
         TabIndex        =   15
         Top             =   2340
         Width           =   255
      End
      Begin VB.TextBox textRecvDate 
         Height          =   264
         Left            =   7830
         MaxLength       =   8
         TabIndex        =   5
         Top             =   870
         Width           =   1092
      End
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   3540
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2940
         Width           =   375
      End
      Begin VB.TextBox textCP43 
         Height          =   264
         Left            =   1380
         MaxLength       =   9
         TabIndex        =   4
         Top             =   870
         Width           =   1764
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   0
         Top             =   300
         Width           =   1092
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   8175
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   300
         Width           =   750
      End
      Begin VB.ComboBox textCP44 
         Height          =   300
         Left            =   1110
         TabIndex        =   3
         Top             =   570
         Width           =   2004
      End
      Begin VB.TextBox textCP23 
         Height          =   264
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   2
         Top             =   300
         Width           =   372
      End
      Begin VB.TextBox textCP08 
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   4020
         Locked          =   -1  'True
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   900
         Width           =   2535
      End
      Begin VB.TextBox textCP14_R 
         Height          =   264
         Left            =   1830
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1170
         Width           =   852
      End
      Begin VB.TextBox textCP22 
         Height          =   264
         Left            =   8490
         MaxLength       =   1
         TabIndex        =   20
         Top             =   2940
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textCP23_R 
         Height          =   264
         Left            =   6510
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1170
         Width           =   372
      End
      Begin VB.TextBox textNP07 
         Height          =   264
         Left            =   1110
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1770
         Width           =   732
      End
      Begin VB.TextBox textNP07_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   1890
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1770
         Width           =   1650
      End
      Begin VB.TextBox textNP08 
         Height          =   264
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   11
         Top             =   2070
         Width           =   1092
      End
      Begin VB.TextBox textNP09 
         Height          =   264
         Left            =   3270
         MaxLength       =   8
         TabIndex        =   12
         Top             =   2070
         Width           =   1092
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   660
         MaxLength       =   12
         TabIndex        =   17
         Top             =   2940
         Width           =   612
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   4470
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1770
         Width           =   372
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   3150
         TabIndex        =   1
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txtEditWord 
         Height          =   264
         Left            =   6420
         MaxLength       =   1
         TabIndex        =   19
         Top             =   2940
         Width           =   372
      End
      Begin VB.TextBox textCP38 
         Height          =   264
         Left            =   -73200
         MaxLength       =   100
         TabIndex        =   28
         Top             =   1812
         Width           =   6912
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73200
         MaxLength       =   600
         TabIndex        =   24
         Top             =   912
         Width           =   6912
      End
      Begin VB.TextBox Text14 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   82
         Top             =   912
         Width           =   372
      End
      Begin VB.TextBox Text13 
         Height          =   264
         Left            =   -73920
         MaxLength       =   4
         TabIndex        =   81
         Top             =   1812
         Width           =   732
      End
      Begin VB.TextBox Text12 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -73080
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   1812
         Width           =   1692
      End
      Begin VB.TextBox Text11 
         Height          =   264
         Left            =   -73896
         MaxLength       =   8
         TabIndex        =   79
         Top             =   2112
         Width           =   2532
      End
      Begin VB.TextBox Text10 
         Height          =   264
         Left            =   -69480
         MaxLength       =   8
         TabIndex        =   78
         Top             =   2112
         Width           =   2412
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   -69690
         MaxLength       =   12
         TabIndex        =   77
         Top             =   1245
         Width           =   612
      End
      Begin VB.TextBox Text8 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -69540
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   930
         Width           =   975
      End
      Begin VB.TextBox Text7 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   75
         Top             =   1212
         Width           =   372
      End
      Begin VB.TextBox Text6 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -72360
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   612
         Width           =   6012
      End
      Begin VB.TextBox Text5 
         Height          =   300
         Left            =   -73905
         MaxLength       =   2000
         TabIndex        =   73
         Top             =   2415
         Width           =   7632
      End
      Begin VB.TextBox Text4 
         Height          =   264
         Left            =   -73920
         MaxLength       =   300
         TabIndex        =   72
         Top             =   1512
         Width           =   7632
      End
      Begin VB.TextBox Text3 
         Height          =   264
         Left            =   -67860
         MaxLength       =   1
         TabIndex        =   71
         Top             =   312
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textCP80 
         Height          =   264
         Left            =   -73200
         MaxLength       =   39
         TabIndex        =   30
         Top             =   2400
         Width           =   6912
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   -73920
         TabIndex        =   70
         Top             =   612
         Width           =   1500
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   22
         Top             =   312
         Width           =   6912
      End
      Begin VB.TextBox Text2 
         Height          =   264
         Left            =   -73920
         MaxLength       =   8
         TabIndex        =   69
         Top             =   312
         Width           =   1092
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   -69555
         TabIndex        =   68
         Top             =   300
         Width           =   1425
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
         Height          =   48
         ItemData        =   "frm020102_15.frx":0038
         Left            =   -67515
         List            =   "frm020102_15.frx":0042
         Sorted          =   -1  'True
         Style           =   1  '項目包含核取方塊
         TabIndex        =   67
         Top             =   915
         Width           =   1260
      End
      Begin MSForms.TextBox textTM67 
         Height          =   270
         Left            =   1110
         TabIndex        =   14
         Top             =   2370
         Width           =   5100
         VariousPropertyBits=   671105051
         MaxLength       =   200
         Size            =   "8996;466"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   495
         Left            =   7620
         TabIndex        =   145
         Top             =   1800
         Width           =   1305
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2293;873"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   885
         Left            =   -73200
         TabIndex        =   26
         Top             =   1500
         Width           =   6912
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12192;1561"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP49 
         Height          =   300
         Left            =   1110
         TabIndex        =   16
         Top             =   2640
         Width           =   7752
         VariousPropertyBits=   -1467989989
         MaxLength       =   300
         Size            =   "13674;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64_R 
         Height          =   300
         Left            =   2040
         TabIndex        =   8
         Top             =   1470
         Width           =   6885
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12144;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   300
         Left            =   1110
         TabIndex        =   21
         Top             =   3210
         Width           =   7752
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13674;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73200
         TabIndex        =   27
         Top             =   1512
         Width           =   6912
         VariousPropertyBits=   -1467989989
         MaxLength       =   100
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73200
         TabIndex        =   29
         Top             =   2112
         Width           =   6912
         VariousPropertyBits=   -1467989989
         MaxLength       =   100
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73200
         TabIndex        =   25
         Top             =   1212
         Width           =   6912
         VariousPropertyBits=   -1467989989
         MaxLength       =   600
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73200
         TabIndex        =   23
         Top             =   612
         Width           =   6912
         VariousPropertyBits=   -1467989989
         MaxLength       =   600
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   250
         Left            =   3150
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   600
         Width           =   5760
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "10160;459"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_R2 
         Height          =   264
         Left            =   2700
         TabIndex        =   111
         TabStop         =   0   'False
         Top             =   1170
         Width           =   1875
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "11070;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   18
         Left            =   4830
         TabIndex        =   144
         Top             =   2115
         Width           =   765
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   6300
         TabIndex        =   143
         Top             =   2385
         Width           =   2655
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "收文指示日期 :"
         Height          =   180
         Left            =   6660
         TabIndex        =   142
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:          (Y: 是)"
         Height          =   180
         Left            =   2370
         TabIndex        =   141
         Top             =   2970
         Width           =   2085
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "(1:勝 2:敗 3:部分勝部分敗)"
         Height          =   180
         Left            =   5460
         TabIndex        =   140
         Top             =   330
         Width           =   2055
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "放棄專用權 :"
         Height          =   180
         Left            =   90
         TabIndex        =   139
         Top             =   2385
         Width           =   990
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號 :"
         Height          =   180
         Left            =   60
         TabIndex        =   138
         Top             =   900
         Width           =   1170
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "進度備註 :"
         Height          =   180
         Left            =   90
         TabIndex        =   137
         Top             =   3240
         Width           =   810
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "發文日 :"
         Height          =   180
         Left            =   90
         TabIndex        =   136
         Top             =   300
         Width           =   630
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "代理人 :"
         Height          =   180
         Left            =   90
         TabIndex        =   135
         Top             =   600
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "列印定稿 :"
         Height          =   180
         Left            =   3600
         TabIndex        =   134
         Top             =   1815
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "點數 :"
         Height          =   180
         Index           =   10
         Left            =   7680
         TabIndex        =   133
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "本所期限 :"
         Height          =   180
         Left            =   90
         TabIndex        =   132
         Top             =   2100
         Width           =   810
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "下一程序 :"
         Height          =   180
         Left            =   90
         TabIndex        =   131
         Top             =   1815
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "預估勝敗 :"
         Height          =   180
         Left            =   4200
         TabIndex        =   130
         Top             =   330
         Width           =   810
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "機關文號 :"
         Height          =   180
         Left            =   3180
         TabIndex        =   129
         Top             =   900
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號承辦人 :"
         Height          =   180
         Left            =   90
         TabIndex        =   128
         Top             =   1215
         Width           =   1710
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號預估勝敗 :"
         Height          =   180
         Left            =   4620
         TabIndex        =   127
         Top             =   1215
         Width           =   1890
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "(1:勝 2:敗 3:部分勝部分敗)"
         Height          =   180
         Left            =   6900
         TabIndex        =   126
         Top             =   1230
         Width           =   2055
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "相關總收文號進度備註 :"
         Height          =   180
         Left            =   90
         TabIndex        =   125
         Top             =   1500
         Width           =   1890
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "條款 :"
         Height          =   180
         Left            =   90
         TabIndex        =   124
         Top             =   2685
         Width           =   450
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "是否出名 :"
         Height          =   180
         Left            =   7680
         TabIndex        =   123
         Top             =   2970
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "法定期限 :"
         Height          =   180
         Left            =   2400
         TabIndex        =   122
         Top             =   2115
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "大約"
         Height          =   180
         Index           =   12
         Left            =   255
         TabIndex        =   121
         Top             =   2970
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "可接獲回音"
         Height          =   180
         Left            =   1320
         TabIndex        =   120
         Top             =   2970
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   4860
         TabIndex        =   119
         Top             =   1800
         Width           =   2745
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   2250
         TabIndex        =   118
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(Y:Word)"
         Height          =   180
         Left            =   6840
         TabIndex        =   117
         Top             =   2970
         Width           =   690
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "是否修改通知函內容 :"
         Height          =   180
         Left            =   4635
         TabIndex        =   116
         Top             =   2970
         Width           =   1710
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   6690
         TabIndex        =   115
         Top             =   2070
         Width           =   900
      End
      Begin VB.Label Label55 
         Caption         =   "對造號數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   109
         Top             =   312
         Width           =   972
      End
      Begin VB.Label Label51 
         Caption         =   "對造中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   105
         Top             =   612
         Width           =   1572
      End
      Begin VB.Label Label50 
         Caption         =   "對造英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   104
         Top             =   912
         Width           =   1572
      End
      Begin VB.Label Label49 
         Caption         =   "對造日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   103
         Top             =   1212
         Width           =   1572
      End
      Begin VB.Label Label48 
         Caption         =   "(1:勝 2:敗 3:部分勝部分敗)"
         Height          =   255
         Left            =   -73500
         TabIndex        =   102
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label47 
         Caption         =   "預估勝敗 :"
         Height          =   225
         Left            =   -74880
         TabIndex        =   101
         Top             =   930
         Width           =   975
      End
      Begin VB.Label Label46 
         Caption         =   "下一程序 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   100
         Top             =   1812
         Width           =   852
      End
      Begin VB.Label Label45 
         Caption         =   "本所期限 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   99
         Top             =   2112
         Width           =   852
      End
      Begin VB.Label Label44 
         Caption         =   "法定期限 :"
         Height          =   252
         Left            =   -70440
         TabIndex        =   98
         Top             =   2112
         Width           =   852
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   15
         Left            =   -70230
         TabIndex        =   97
         Top             =   1245
         Width           =   495
      End
      Begin VB.Label Label43 
         Caption         =   "可接獲回音"
         Height          =   255
         Left            =   -68970
         TabIndex        =   96
         Top             =   1245
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   255
         Index           =   5
         Left            =   -70110
         TabIndex        =   95
         Top             =   960
         Width           =   555
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   -73500
         TabIndex        =   94
         Top             =   1260
         Width           =   2745
      End
      Begin VB.Label Label41 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   93
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label40 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   92
         Top             =   612
         Width           =   972
      End
      Begin VB.Label Label38 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   91
         Top             =   312
         Width           =   852
      End
      Begin VB.Label Label37 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74865
         TabIndex        =   90
         Top             =   2415
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "條款 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   89
         Top             =   1512
         Width           =   852
      End
      Begin VB.Label Label35 
         Caption         =   "(N:不出名)"
         Height          =   255
         Left            =   -67380
         TabIndex        =   88
         Top             =   315
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "是否出名 :"
         Height          =   255
         Left            =   -68820
         TabIndex        =   87
         Top             =   315
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label33 
         Caption         =   "對造案件商品類別 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   86
         Top             =   2412
         Width           =   1572
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   -70485
         TabIndex        =   84
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   -68475
         TabIndex        =   83
         Top             =   945
         Width           =   900
      End
      Begin VB.Label Label53 
         Caption         =   "對造案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   107
         Top             =   1812
         Width           =   1572
      End
      Begin VB.Label Label52 
         Caption         =   "對造案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   106
         Top             =   2112
         Width           =   1572
      End
      Begin VB.Label Label32 
         Caption         =   "對造案件名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   85
         Top             =   1512
         Width           =   1572
      End
      Begin VB.Label Label54 
         Caption         =   "對造案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   108
         Top             =   1512
         Width           =   1572
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1170
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   2280
      Width           =   7812
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13779;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   264
      Left            =   1200
      TabIndex        =   65
      TabStop         =   0   'False
      Top             =   2010
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   264
      Left            =   5370
      TabIndex        =   64
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   264
      Left            =   1200
      TabIndex        =   63
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   264
      Left            =   5370
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3675
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1200
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1470
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5370
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   2010
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   264
      Left            =   5370
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   390
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5370
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3675
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   8
      Left            =   4500
      TabIndex        =   61
      Top             =   1515
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   60
      Top             =   1782
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   13
      Left            =   4500
      TabIndex        =   59
      Top             =   1785
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   58
      Top             =   2052
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   57
      Top             =   1512
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   4500
      TabIndex        =   56
      Top             =   2055
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人 :"
      Height          =   180
      Index           =   2
      Left            =   4500
      TabIndex        =   55
      Top             =   420
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發證日 :"
      Height          =   180
      Index           =   3
      Left            =   4500
      TabIndex        =   54
      Top             =   705
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   53
      Top             =   702
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   52
      Top             =   426
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   51
      Top             =   1242
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4500
      TabIndex        =   50
      Top             =   975
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4500
      TabIndex        =   49
      Top             =   1245
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   972
      Width           =   1575
   End
   Begin VB.Label Label31 
      Caption         =   "(N:不出名)"
      Height          =   255
      Left            =   9225
      TabIndex        =   37
      Top             =   5610
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   150
      TabIndex        =   36
      Top             =   2325
      Width           =   810
   End
End
Attribute VB_Name = "frm020102_15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/27 Form2.0已修改 textTM44/textCP13/textCP14/textCP44_2/textCP14_R2/textTM23(申請人名).../cmbTM05/textCP64/textCP64_R/textCP40(對造中/日名).../lstNameAgent/textTM67(111/8/8 Lydia)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 收文號
Dim m_CP09 As String
Dim m_CP31 As String 'Add By Sindy 2011/7/12
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
Dim m_CP07 As String 'Add By Sindy 2010/12/22
' 智權人員
Dim m_CP12 As String 'Add By Sindy 2010/10/12
Dim m_CP13 As String
' 承辦人
Dim m_CP14 As String
' 申請人
Dim m_TM23 As String
'Add By Sindy 2009/06/11
' 卷宗性質
Dim m_TM28 As String
'add by nickc 2007/02/01
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存商標基本檔或服務業務基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer

' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'add by nick 2004/08/12
Dim m_CP84 As String       '發文規費
'add by nick 2004/09/27
Public m_CU103 As String         '公司負責人英文名稱
'add by nick 2004/10/05
Public m_CU05 As String         '客戶英文名稱
Public m_CU88 As String         '客戶英文名稱
Public m_CU89 As String         '客戶英文名稱
Public m_CU90 As String         '客戶英文名稱
'add by nickc 2006/01/20
Public m_CU112 As String        '客戶中文地址郵遞區號
'Add By Sindy 2012/2/7
Public m_CU39 As String         '代表人1（中）
Public m_CU40 As String         '代表人1（英）
Public m_CU41 As String         '代表人1（日）
'2012/2/7 End

Dim m_TM24 As String
'add by nickc 2006/01/27
Dim m_CP110 As String
'add by nickc 2007/08/10
Dim SeekCu05(1 To 5) As String
Dim SeekCu88(1 To 5) As String
Dim SeekCu89(1 To 5) As String
Dim SeekCu90(1 To 5) As String
Dim SeekCu103(1 To 5) As String
Dim SeekCu112(1 To 5) As String
'Add By Sindy 2012/2/7
Dim SeekCu39(1 To 5) As String
Dim SeekCu40(1 To 5) As String
Dim SeekCu41(1 To 5) As String
'2012/2/7 End
'Add By Sindy 2012/10/31
Public m_CU10 As String
Dim SeekCu10(1 To 5) As String
'2012/10/31 End
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
Dim m_AgentName As String 'Add By Amy 2021/12/27
'Added by Lydia 2022/05/23
Dim m_LosCP84 As String ' 法律所案源之規費
Dim m_LOS15 As String '法律所案源單號
Dim m_LOS02 As String '法律所案源類別
Dim m_LosMemo As String  'email說明

Private Sub cmdCancel_Click()
   'Add By Sindy 2018/5/3
   If frm020102_01.bolIsEMPFlow = True Then
      frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
      frm090202_4.QueryData
   End If
   '2018/5/3 End
   frm020102_01.Show
   Unload Me
End Sub

Private Sub cmdCaseProgress_Click()
   frm020102_06.SetData 0, m_TM01, True
   frm020102_06.SetData 1, m_TM02, False
   frm020102_06.SetData 2, m_TM03, False
   frm020102_06.SetData 3, m_TM04, False
   frm020102_06.SetData 4, m_CP09, False
   frm020102_06.SetParent Me
   Me.Hide
   frm020102_06.Show
   frm020102_06.QueryData
End Sub

Private Sub cmdExit_Click()
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   ' 90.10.09 modify by louis
   'Add By Sindy 2018/5/3
   If frm020102_01.bolIsEMPFlow = True Then
      frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
      frm090202_4.QueryData
   End If
   '2018/5/3 End
   Unload frm020102_01
   'frm020102_01.Show
   Unload Me
End Sub

Private Sub cmdok_Click()
Dim strNewCP64 As String 'Add by Amy 2020/02/05 進度備註

   'Add By Cheng 2002/05/23
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   If CheckDataValid = True Then
      'add by nick 2004/09/27
      'edit by nick 2004/10/07
      'If m_TM01 <> "FCT" Then
      If m_TM01 <> "FCT" And m_TM01 <> "TB" And m_TM01 <> "TC" And m_TM01 <> "TD" And (m_TM01 = "T" And m_TM10 <> "020") Then
            'add by nickc 2007/08/10
            SeekCu05(1) = "": SeekCu05(2) = "": SeekCu05(3) = "": SeekCu05(4) = "": SeekCu05(5) = ""
            SeekCu88(1) = "": SeekCu88(2) = "": SeekCu88(3) = "": SeekCu88(4) = "": SeekCu88(5) = ""
            SeekCu89(1) = "": SeekCu89(2) = "": SeekCu89(3) = "": SeekCu89(4) = "": SeekCu89(5) = ""
            SeekCu90(1) = "": SeekCu90(2) = "": SeekCu90(3) = "": SeekCu90(4) = "": SeekCu90(5) = ""
            SeekCu103(1) = "": SeekCu103(2) = "": SeekCu103(3) = "": SeekCu103(4) = "": SeekCu103(5) = ""
            SeekCu112(1) = "": SeekCu112(2) = "": SeekCu112(3) = "": SeekCu112(4) = "": SeekCu112(5) = ""
            'Add By Sindy 2012/2/7
            SeekCu39(1) = "": SeekCu39(2) = "": SeekCu39(3) = "": SeekCu39(4) = "": SeekCu39(5) = ""
            SeekCu40(1) = "": SeekCu40(2) = "": SeekCu40(3) = "": SeekCu40(4) = "": SeekCu40(5) = ""
            SeekCu41(1) = "": SeekCu41(2) = "": SeekCu41(3) = "": SeekCu41(4) = "": SeekCu41(5) = ""
            '2012/2/7 End
            'Add By Sindy 2012/10/31
            SeekCu10(1) = "": SeekCu10(2) = "": SeekCu10(3) = "": SeekCu10(4) = "": SeekCu10(5) = ""
            '2012/10/31 End
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM23
            Call Pub_GetDataFrm020102(m_TM23, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            'edit by nickc 2006/01/20
            'If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Then
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, m_TM23)
                  frm020102_22.Label4.Caption = m_TM23 & " " & textTM23 'Add By Sindy 2014/7/30
                  frm020102_22.Show vbModal
                  'add by nickc 2007/08/10
                  SeekCu05(1) = m_CU05
                  SeekCu88(1) = m_CU88
                  SeekCu89(1) = m_CU89
                  SeekCu90(1) = m_CU90
                  SeekCu103(1) = m_CU103
                  SeekCu112(1) = m_CU112
                  'Add By Sindy 2012/2/27
                  SeekCu39(1) = m_CU39
                  SeekCu40(1) = m_CU40
                  SeekCu41(1) = m_CU41
                  '2012/2/27 End
                  'Add By Sindy 2012/10/31
                  SeekCu10(1) = m_CU10
                  '2012/10/31 End
            End If
            'add by nickc 2007/08/10 多申請人也要
            If m_TM78 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM78
            Call Pub_GetDataFrm020102(m_TM78, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, m_TM78)
                  frm020102_22.Label4.Caption = m_TM78 & " " & textTM78 'Add By Sindy 2014/7/30
                  frm020102_22.Show vbModal
                  SeekCu05(2) = m_CU05
                  SeekCu88(2) = m_CU88
                  SeekCu89(2) = m_CU89
                  SeekCu90(2) = m_CU90
                  SeekCu103(2) = m_CU103
                  SeekCu112(2) = m_CU112
                  'Add By Sindy 2012/2/27
                  SeekCu39(2) = m_CU39
                  SeekCu40(2) = m_CU40
                  SeekCu41(2) = m_CU41
                  '2012/2/27 End
                  'Add By Sindy 2012/10/31
                  SeekCu10(2) = m_CU10
                  '2012/10/31 End
            End If
            End If
            If m_TM79 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM79
            Call Pub_GetDataFrm020102(m_TM79, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)

            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, m_TM79)
                  frm020102_22.Label4.Caption = m_TM79 & " " & textTM79 'Add By Sindy 2014/7/30
                  frm020102_22.Show vbModal
                  SeekCu05(3) = m_CU05
                  SeekCu88(3) = m_CU88
                  SeekCu89(3) = m_CU89
                  SeekCu90(3) = m_CU90
                  SeekCu103(3) = m_CU103
                  SeekCu112(3) = m_CU112
                  'Add By Sindy 2012/2/27
                  SeekCu39(3) = m_CU39
                  SeekCu40(3) = m_CU40
                  SeekCu41(3) = m_CU41
                  '2012/2/27 End
                  'Add By Sindy 2012/10/31
                  SeekCu10(3) = m_CU10
                  '2012/10/31 End
            End If
            End If
            If m_TM80 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM80
            Call Pub_GetDataFrm020102(m_TM80, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)

            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, m_TM80)
                  frm020102_22.Label4.Caption = m_TM80 & " " & textTM80 'Add By Sindy 2014/7/30
                  frm020102_22.Show vbModal
                  SeekCu05(4) = m_CU05
                  SeekCu88(4) = m_CU88
                  SeekCu89(4) = m_CU89
                  SeekCu90(4) = m_CU90
                  SeekCu103(4) = m_CU103
                  SeekCu112(4) = m_CU112
                  'Add By Sindy 2012/2/27
                  SeekCu39(4) = m_CU39
                  SeekCu40(4) = m_CU40
                  SeekCu41(4) = m_CU41
                  '2012/2/27 End
                  'Add By Sindy 2012/10/31
                  SeekCu10(4) = m_CU10
                  '2012/10/31 End
            End If
            End If
            If m_TM81 <> "" Then    '2007/8/14 modify by sonia 加此條件判斷,有多個申請人才要做
            'Modified by Lydia 2024/07/03 改傳入變數;
            'GetCu103ByCustomer Me, m_TM81
            Call Pub_GetDataFrm020102(m_TM81, m_CU103, m_CU05, m_CU88, m_CU89, m_CU90, m_CU112, m_CU39, m_CU40, m_CU41, m_CU10)
            
            If m_CU103 = "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) = "" Or m_CU112 = "" Then
                  'Modified by Lydia 2024/07/03
                  'Set frm020102_22.oNextForm = Me
                  Call frm020102_22.SetParent(Me, m_TM81)
                  frm020102_22.Label4.Caption = m_TM81 & " " & textTM81 'Add By Sindy 2014/7/30
                  frm020102_22.Show vbModal
                  SeekCu05(5) = m_CU05
                  SeekCu88(5) = m_CU88
                  SeekCu89(5) = m_CU89
                  SeekCu90(5) = m_CU90
                  SeekCu103(5) = m_CU103
                  SeekCu112(5) = m_CU112
                  'Add By Sindy 2012/2/27
                  SeekCu39(5) = m_CU39
                  SeekCu40(5) = m_CU40
                  SeekCu41(5) = m_CU41
                  '2012/2/27 End
                  'Add By Sindy 2012/10/31
                  SeekCu10(5) = m_CU10
                  '2012/10/31 End
            End If
            End If
      End If
      
      strNewCP64 = textCP64 'Add by Amy 2020/02/05
      
      'Modify By Sindy 2011/3/9 若為電子送件則不經發文室
      'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
      If (textCP118.Visible = True And textCP118 <> "") Then
         'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
         If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
            Exit Sub
         End If
         'end 2016/5/16
         
         'Add by Amy 2020/02/05 +輸入收文文號
         If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
            'Add By Sindy 2020/8/12 主管機關為經濟部智慧財產局,才做自動扣款
            If m_CP130s = "經濟部智慧財產局" Then
            '2020/8/12 END
               'Add by Amy 2020/01/13
               'If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
                  'If textCP118 = "Y" And Val(textCP84) > 0 Then
                  If Val(textCP84) > 0 Then
                     If txtPayToday.Visible = True And txtPayToday = "" Then
                        MsgBox "電子送件請輸入是否當日扣款(Y/N)！", vbExclamation
                        txtPayToday.SetFocus
                        Exit Sub
                     End If
                     strExc(0) = InputBox("請輸入智慧局收文文號!!")
                     If strExc(0) = "" Then
                        Exit Sub
                     Else
                        strNewCP64 = "智慧局收文文號:" & strExc(0) & ";" & textCP64 '先保留進度備註，等檢查完後更新欄位
                     End If
                  End If
               'End If
               'end 2020/01/13
            'Add By Sindy 2020/8/12
            ElseIf txtPayToday.Visible = True And txtPayToday <> "" Then
               txtPayToday = ""
            End If
            '2020/8/12 END
        End If
        'end 2020/02/05
      Else
         'Add by Sindy 98/3/24
         If m_TM10 = "000" Then
            m_CP09s = m_CP09
            'Add by Sindy 2009/4/24
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
               Exit Sub
      '      Else
      '         m_CP123s = GetCPMSendYn(m_TM01, m_CP10, 1)
            End If
         End If
      End If
      
      textCP64 = strNewCP64 'Add by Amy 2020/02/05
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Add By Cheng 2002/11/08
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        'Add By Sindy 2021/2/25
        End If
        If textPrint = "N" Then
            If m_CP09 <> "" Then
               Call PUB_TCaseAskIsPost(m_CP09)
            End If
        '2021/2/25 END
        End If
       'Added by Lydia 2022/05/23 PT案(傳入收文號)取得法律案源之發文規費，並且有輸入發文規費才做檢查
       If textCP84.Enabled = True And m_TM10 = "000" And m_LosCP84 <> "0" Then
          If Val(Trim(textCP84.Text)) <> 0 And Val(m_LosCP84) <> Val(Trim(textCP84.Text)) Then
              PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", ""), m_LosMemo
          End If
       Else
       'end 2022/05/23
           '2012/7/23 add by sonia
           '台灣案發文規費與收文規費不符時,mail給智權人員
           If textCP84.Enabled = True And m_TM10 = "000" And Val(Me.textCP84.Text) <> Val(m_CP84) Then
              '2020/01/13 Modify by Amy +if 傳strCP118參數
              If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                  PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", "")
              Else
                  PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
              End If
           End If
           '2012/7/23 end
       End If 'Added by Lydia 2022/05/23
      'Add By Sindy 2018/5/3
      If frm020102_01.bolIsEMPFlow = True Then
         frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
         frm090202_4.QueryData
      End If
      '2018/5/3 End
      
      'Add By Sindy 2025/7/11 外商發文時,增加發Mail通知承辦人及副本給判發主管
      If Left(m_CP12, 1) = "F" Then
         Call PUB_FCTSendRecvMail(m_CP09)
      End If
      '2025/7/11 END
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      '********* 901123 nick   清畫面
      'frm020102_01.radio(0).Value = True
      'frm020102_01.textCP09.Enabled = True
      'frm020102_01.textCP09.Text = ""
      'frm020102_01.textTM01.Enabled = False
      'frm020102_01.textTM01.Text = "" modify by sonia
      'frm020102_01.textTM02.Enabled = False
      'frm020102_01.textTM02.Text = ""
      'frm020102_01.textTM02_2.Enabled = False
      'frm020102_01.textTM02_2.Text = ""
      'frm020102_01.textTM03.Enabled = False
      'frm020102_01.textTM03.Text = ""
      'frm020102_01.textTM04.Enabled = False
      'frm020102_01.textTM04.Text = ""
      'frm020102_01.grdList.Clear
      'frm020102_01.grdList.Rows = 2
      '*********************************
      'frm020102_01.RefreshData
      'Add By Cheng 2002/04/30
      '若有未發文資料顯示警告
      If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
         'Add By Sindy 2018/5/3
         If frm020102_01.bolIsEMPFlow = True Then
            Unload frm020102_01
            frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
            frm090202_4.Show
            Unload Me
            Exit Sub
         End If
         '2018/5/3 End
      End If
      
      frm020102_01.Show
      ' 90.12.07 modify by louis
'      frm020102_01.Clear
      
      'Add By Cheng 2002/01/10
      frm020102_01.Clear1
      
      Unload Me
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
   pub_ModifyCaseNum = ""
   QueryData
End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/02/01
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM45.BackColor = &H8000000F
   
   textCP08.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   textCP14_R2.BackColor = &H8000000F
   textNP07_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   'Add by nickc 2006/01/27
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Add by Amy 2021/12/27一開始將ListBox拉到需要的大小,字型會自動放大；所以畫面預設為一列高度,Form_Load才放大到需要的大小
    lstNameAgent.Height = 500
    lstNameAgent.Width = 1300

   SSTab1.Tab = 0 'Add By Sindy 2018/10/16
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 收文號
      Case 0: m_CP09 = strData
      ' 相關總收文號
      Case 99:
         textCP43 = strData
         textCP43_Validate False
         textCP43_LostFocus
   End Select
End Sub

Private Sub ClearAgentList()
   If m_AgentCount > 0 Then
      Erase m_AgentList
   End If
   m_AgentCount = 0
End Sub

Private Sub AddAgent(ByVal strAgentCode As String, ByVal strAgentName As String)
   Dim nIndex As Integer
   Dim bFind As Boolean
   bFind = False
   For nIndex = 0 To m_AgentCount - 1
      If m_AgentList(nIndex).aiCode = strAgentCode Then
         bFind = True
         Exit For
      End If
   Next nIndex
   If bFind = False Then
      ReDim Preserve m_AgentList(m_AgentCount + 1)
      m_AgentList(m_AgentCount).aiCode = strAgentCode
      m_AgentList(m_AgentCount).aiName = strAgentName
      m_AgentCount = m_AgentCount + 1
   End If
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiOldData = strFieldData
         m_TMSPList(nPos).fiNewData = strFieldData
         m_TMSPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_TMSPList(m_TMSPCount + 1)
      m_TMSPList(m_TMSPCount).fiName = strFieldName
      m_TMSPList(m_TMSPCount).fiOldData = strFieldData
      m_TMSPList(m_TMSPCount).fiNewData = strFieldData
      m_TMSPList(m_TMSPCount).fiType = nFieldType
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定商標基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetTMSPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_TMSPCount - 1
      If m_TMSPList(nPos).fiName = strFieldName Then
         bFind = True
         m_TMSPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 清除案件進度檔檔案欄位串列
Private Sub ClearCPFieldList()
   If m_CPCount > 0 Then
      Erase m_CPList
   End If
   m_CPCount = 0
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiOldData = strFieldData
         m_CPList(nPos).fiNewData = strFieldData
         m_CPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_CPList(m_CPCount + 1)
      m_CPList(m_CPCount).fiName = strFieldName
      m_CPList(m_CPCount).fiOldData = strFieldData
      m_CPList(m_CPCount).fiNewData = strFieldData
      m_CPList(m_CPCount).fiType = nFieldType
      m_CPCount = m_CPCount + 1
   End If
End Sub

' 設定案件進度檔欄位串列中的欄位內容
Private Sub SetCPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_CPCount - 1
      If m_CPList(nPos).fiName = strFieldName Then
         bFind = True
         m_CPList(nPos).fiNewData = strFieldData
         Exit For
      End If
   Next nPos
End Sub

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSubSQL As String
   Dim rsSubTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("TM44"))
      'Add By Sindy 2013/1/31
      If m_TM44 <> "" Then
         textTM44 = m_TM44 & "  " & GetPrjName1(m_TM44)
      Else
         textTM44 = ""
      End If
      '2013/1/31 End
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      'add by nickc 2007/02/01
      ElseIf IsNull(rsTmp.Fields("TM12")) = False Then
        textTM15 = rsTmp.Fields("TM12")
      End If
      ' 申請案號
'edit by nickc 2007/02/01
'      If IsNull(rsTmp.Fields("TM12")) = False Then
'         textTM12 = rsTmp.Fields("TM12")
'      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("TM07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 商標種類
      If IsNull(rsTmp.Fields("TM08")) = False Then
         'textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      End If
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      'Add By Sindy 2009/06/11
      ' 卷宗性質
      If IsNull(rsTmp.Fields("TM28")) = False Then
         m_TM28 = rsTmp.Fields("TM28")
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         textTM78 = GetCustomerName(rsTmp.Fields("TM78"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         textTM79 = GetCustomerName(rsTmp.Fields("TM79"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         textTM80 = GetCustomerName(rsTmp.Fields("TM80"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
         textTM81 = GetCustomerName(rsTmp.Fields("TM81"), 0)
      End If
      
      ' 放棄專用權
      textTM67 = Empty
      If IsNull(rsTmp.Fields("TM67")) = False Then
         textTM67 = rsTmp.Fields("TM67")
      End If
      'Add By Cheng 2002/12/31
      '內商T的案件(大陸->台), 要顯示基本檔的彼所案號
      If m_TM10 = 台灣國家代號 Then
        Me.textTM45.Text = "" & rsTmp("TM45").Value
      End If
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("tm24"))
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務頁務基本檔的欄位內容
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'add by nickc 2008/02/22
      m_TM44 = CheckStr(rsTmp.Fields("SP26"))
      'Add By Sindy 2013/1/31
      If m_TM44 <> "" Then
         textTM44 = m_TM44 & "  " & GetPrjName1(m_TM44)
      Else
         textTM44 = ""
      End If
      '2013/1/31 End
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("SP58")) = False Then
         m_TM78 = rsTmp.Fields("SP58")
         textTM78 = GetCustomerName(rsTmp.Fields("SP58"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("SP59")) = False Then
         m_TM79 = rsTmp.Fields("SP59")
         textTM79 = GetCustomerName(rsTmp.Fields("SP59"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("SP65")) = False Then
         m_TM80 = rsTmp.Fields("SP65")
         textTM80 = GetCustomerName(rsTmp.Fields("SP65"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("SP66")) = False Then
         m_TM81 = rsTmp.Fields("SP66")
         textTM81 = GetCustomerName(rsTmp.Fields("SP66"), 0)
      End If
      
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         'edit by nickc 2007/02/01
         'textTM12 = rsTmp.Fields("SP11")
         textTM15 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
      End If
      'Add By Cheng 2002/12/31
      '內商T的案件(大陸->台), 顯示彼所案號
      If m_TM10 = 台灣國家代號 Then
        Me.textTM45.Text = "" & rsTmp("SP27").Value
      End If
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("sp72"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgress()
Dim strSql As String
Dim strSubSQL As String
Dim rsTmp As New ADODB.Recordset
Dim rsSubTmp As New ADODB.Recordset
Dim strCP27 As String
Dim strCP43 As String
Dim strCP44 As String
'   Dim strCP45 As String
Dim nIndex As Integer
Dim bFind As Boolean
'Add By Cheng 2002/07/09
Dim strTempName As String
Dim m_Fee As String         '銷帳服務費 2012/8/3 add by sonia
Dim m_Official As String    '銷帳規費   2012/8/3 add by sonia
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      ' 案件性質
      'Add By Cheng 2002/07/17
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
      'Add By Sindy 2010/12/28 法定期限
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2010/12/28 End
      
      ' 業務區別
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         m_CP12 = rsTmp.Fields("CP12")
         'textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 承辦人員
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("CP14")) = False Then
         m_CP14 = rsTmp.Fields("CP14")
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      
      'Add By Sindy 2011/7/12
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      ' 是否出名
      textCP22 = Empty
      If IsNull(rsTmp.Fields("CP22")) = False Then
         textCP22 = rsTmp.Fields("CP22")
      End If
      SetCPFieldOldData "CP22", textCP22, 0
      
      ' 發文日(預設為系統日)
      textCP27 = TAIWANDATE(SystemDate())
      strCP27 = Empty
      If IsNull(rsTmp.Fields("CP27")) = False Then
         strCP27 = rsTmp.Fields("CP27")
      End If
      SetCPFieldOldData "CP27", strCP27, 1
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE("" & rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      ' 相關總收文號
      textCP43 = Empty
      If IsNull(rsTmp.Fields("CP43")) = False Then
         textCP43 = rsTmp.Fields("CP43")
      End If
      SetCPFieldOldData "CP43", textCP43, 0
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
'      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
         textTM45 = rsTmp.Fields("CP45")
'         strCP45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", textTM45, 0
'      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 預估結果(預估勝敗)
      textCP23 = Empty
      If IsNull(rsTmp.Fields("CP23")) = False Then
         textCP23 = rsTmp.Fields("CP23")
      End If
      SetCPFieldOldData "CP23", textCP23, 0
      ' 條款
      textCP49 = Empty
      If IsNull(rsTmp.Fields("CP49")) = False Then
         textCP49 = rsTmp.Fields("CP49")
      End If
      SetCPFieldOldData "CP49", textCP49, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
    'add by nick 2004/08/12 發文規費
     If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
         'edit by nick 2004/09/08
         'm_CP84 = CheckStr(rsTmp.Fields("CP17"))
         '2012/8/3 modify by sonia 若有銷帳則要扣除銷帳規費
         'm_CP84 = IIf(PUB_ChkDelay(m_CP09) = True, "0", CheckStr(rsTmp.Fields("CP17")))
         m_CP84 = CheckStr(rsTmp.Fields("CP17"))
         If Val("" & rsTmp.Fields("CP77")) <> 0 Then
            If GetCP77Detail(m_CP09, m_Fee, m_Official) = True Then
               m_CP84 = m_CP84 - m_Official
            End If
         End If
         If PUB_ChkDelay(m_CP09) = True Then m_CP84 = 0
         '2012/8/3 end
         textCP84.Text = m_CP84
     End If
     
     'Added by Morgan 2012/9/6 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/9/6

      'Added by Lydia 2022/05/23 法律所案源：取得案源類別、發文規費、email加註
      m_LOS15 = "" & rsTmp.Fields("cp162")
      '限制B2類發文規費; 只需考慮B2類, A類不會有PT案,B1類PT案不會繳規費, C類已取消(有也是照原來規則不必改)
      'Memo by Lydia 2022/09/16 模組回傳m_LOS02
      m_LosCP84 = PUB_GetLosCP84(m_LOS15, m_TM01, m_TM02, m_TM03, m_TM04, "B2", m_LOS02, m_LosMemo)
      'end 2022/05/23
      
      'Add by Amy 2020/01/13 電子送件一率自動扣款(A)若超過3點半發文則須人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
         txtPayToday = ""
         If textCP118 = "Y" Then
            'Modify by Amy 2020/08/11 發文日小於系統日,電子送件是否當日扣款設N;發文日為當天且3點半前才設Y(原只判斷3點半)
            If Val(textCP27) < strSrvDate(2) Then
               txtPayToday = "N"
            ElseIf Val(textCP27) = strSrvDate(2) And Val(ServerTime) <= 153000 Then
               txtPayToday = "Y"
            End If
            'end 2020/08/11
         End If
      End If
      'end 2020/01/13
      textCP27.Tag = textCP27.Text 'Add By Sindy 2020/8/12
      
      'add by nickc 2006/01/27
      'm_CP110 = CheckStr(rsTmp.Fields("cp110"))
      'SetCPFieldOldData "CP110", m_CP110, 0
      'Modify By Sindy 2010/9/20
      If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
      '2010/9/20 End
      
      'Add By Sindy 2009/06/11
      ' 對造號數
      textCP36 = Empty
      If IsNull(rsTmp.Fields("CP36")) = False Then
         textCP36 = rsTmp.Fields("CP36")
      End If
      SetCPFieldOldData "CP36", textCP36, 0
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            ' 對造案件名稱
            textCP37_1 = Empty
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37_1 = rsTmp.Fields("CP37")
            End If
            SetCPFieldOldData "CP37", textCP37_1, 0
        Case Else
            ' 對造案件名稱(中)
            textCP37 = Empty
            If IsNull(rsTmp.Fields("CP37")) = False Then
               textCP37 = rsTmp.Fields("CP37")
            End If
            SetCPFieldOldData "CP37", textCP37, 0
            ' 對造案件名稱(英)
            textCP38 = Empty
            If IsNull(rsTmp.Fields("CP38")) = False Then
               textCP38 = rsTmp.Fields("CP38")
            End If
            SetCPFieldOldData "CP38", textCP38, 0
            ' 對造案件名稱(日)
            textCP39 = Empty
            If IsNull(rsTmp.Fields("CP39")) = False Then
               textCP39 = rsTmp.Fields("CP39")
            End If
            SetCPFieldOldData "CP39", textCP39, 0
        End Select
      ' 對造名稱(中)
      textCP40 = Empty
      If IsNull(rsTmp.Fields("CP40")) = False Then
         textCP40 = rsTmp.Fields("CP40")
      End If
      SetCPFieldOldData "CP40", textCP40, 0
      ' 對造名稱(英)
      textCP41 = Empty
      If IsNull(rsTmp.Fields("CP41")) = False Then
         textCP41 = rsTmp.Fields("CP41")
      End If
      SetCPFieldOldData "CP41", textCP41, 0
      ' 對造名稱(日)
      textCP42 = Empty
      If IsNull(rsTmp.Fields("CP42")) = False Then
         textCP42 = rsTmp.Fields("CP42")
      End If
      SetCPFieldOldData "CP42", textCP42, 0
      ' 對造案件商品類別
      textCP80 = Empty
      If IsNull(rsTmp.Fields("CP80")) = False Then
         textCP80 = rsTmp.Fields("CP80")
      End If
      SetCPFieldOldData "CP80", textCP80, 0
      '2009/06/11 End
     
     
      ' 代理人
      ClearAgentList
      'add by nickc 2008/03/26 若是原先有，也要加入
      If textCP44.Text <> "" Then
            If PUB_GetAgentName(m_TM01, textCP44, strTempName) Then
               strCP44 = strTempName
            Else
               strCP44 = ""
            End If
            AddAgent textCP44, strCP44
      End If
        'Modify By Cheng 2004/02/20
'      strSubSQL = "SELECT DISTINCT CP44 FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' "
      strSubSQL = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null Group By CP44 Order By 2 Desc, 1 "
        'End
      rsSubTmp.CursorLocation = adUseClient
      rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSubTmp.RecordCount > 0 Then
         rsSubTmp.MoveFirst
         ' 依序將代理人加入到系統串列中
         Do While rsSubTmp.EOF = False
            If IsNull(rsSubTmp.Fields("CP44")) = False Then
               'Modify By Cheng 2002/07/09
'               strCP44 = GetFAgentName(rsSubTmp.Fields("CP44"))
'               AddAgent rsSubTmp.Fields("CP44"), GetFAgentName(rsSubTmp.Fields("CP44"))
               If PUB_GetAgentName(m_TM01, rsSubTmp.Fields("CP44"), strTempName) Then
                  strCP44 = strTempName
               Else
                  strCP44 = ""
               End If
               AddAgent rsSubTmp.Fields("CP44"), strTempName
            End If
            rsSubTmp.MoveNext
         Loop
      End If
      rsSubTmp.Close
    ' 從系統串列中取得所有代理人並放入Combo Box中
    For nIndex = 0 To m_AgentCount - 1
       'Modify By Cheng 2002/09/18
'            textCP44.AddItem m_AgentList(nIndex).aiName
       textCP44.AddItem m_AgentList(nIndex).aiCode
    Next nIndex
    ' 設定顯示為第一筆
    If textCP44.ListCount > 0 Then
       textCP44.ListIndex = 0
       textCP44_Validate False
    End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/27
   Dim tm(1 To 4) As String
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearTMSPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   ' 先取得本所案號
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 本所案號
      If IsNull(rsTmp.Fields("CP01")) = False Then: m_TM01 = rsTmp.Fields("CP01")
      If IsNull(rsTmp.Fields("CP02")) = False Then: m_TM02 = rsTmp.Fields("CP02")
      If IsNull(rsTmp.Fields("CP03")) = False Then: m_TM03 = rsTmp.Fields("CP03")
      If IsNull(rsTmp.Fields("CP04")) = False Then: m_TM04 = rsTmp.Fields("CP04")
   End If
   rsTmp.Close
   
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 本所案號
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04

   ' 收文號
   textCP09 = m_CP09
      
   ' 取得案件進度檔的欄位
   QueryCaseProgress

   'add by nickc 2006/01/27
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
      
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   'Add By Sindy 2020/8/12 取得主管機關, 只有主管機關是"經濟部智慧財產局"時, 才做自動扣款
   Call GetCaseFeeByNick(m_TM01, m_TM10, m_CP10, m_strCF10)
   If m_strCF10 <> "經濟部智慧財產局" And txtPayToday.Visible = True And txtPayToday <> "" Then
      txtPayToday = ""
   End If
   '2020/8/12 END
   
   'Add By Sindy 2021/1/15 T發文所有程式,台灣案鎖住畫面上之CP44,不可輸入
   If m_TM10 = "000" Then
      textCP44.Enabled = False
   End If
   '2021/1/15 END
   
   'Modify By Sindy 2012/7/26
   '台灣案才需顯示出名代理人
   lstNameAgent.Clear
   If m_TM10 = "000" Then
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      'Modify by Amy 2021/12/27 改Form2.0,bForm2設True
      PUB_SetOurAgent lstNameAgent, tm(), m_CP110, , True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2012/7/26 End
   
   ' 案件性質為申請意見書或補充理由時才讓使用者輸入以下欄位
   Select Case m_CP10
      '2011/11/4 modify by sonia 加210陳述意見書
      Case "202", "612", "210":
         EnableTextBox textCP14_R, False     'modify by sonia 2023/11/17 桂英同意阿蓮提出不更新相關總收文號承辦人請作,故改為False
         EnableTextBox textCP23_R, True
         EnableTextBox textCP64_R, True
         If IsEmptyText(textCP43) = False Then
            textCP43_LostFocus
         End If
      Case Else:
         EnableTextBox textCP14_R, False
         EnableTextBox textCP23_R, False
         EnableTextBox textCP64_R, False
   End Select
   
   'Add By Sindy 2009/06/11
   '2011/11/4 modify by sonia 卷宗性質非申請者改用210陳述意見書
   'If m_CP10 = "202" And m_TM28 <> "1" Then
   'Modify By Sindy 2020/11/13 + 214.陳述聲明
   If m_CP10 = "210" Or m_CP10 = "214" Then
      SSTab1.TabVisible(1) = True
      Select Case m_TM01
      Case "T", "FCT", "CFT", "TF"
          Me.Label54.Visible = False
          Me.Label53.Visible = False
          Me.Label52.Visible = False
          Me.textCP37.Visible = False
          Me.textCP37.Enabled = False
          Me.textCP38.Visible = False
          Me.textCP38.Enabled = False
          Me.textCP39.Visible = False
          Me.textCP39.Enabled = False
      Case Else
          Me.Label32.Visible = False
          Me.textCP37_1.Visible = False
          Me.textCP37_1.Enabled = False
      End Select
   Else
      SSTab1.TabVisible(1) = False
   End If
   '2009/06/11 End
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
   'Modify By Sindy 2014/5/1 從2014/05/01起,大陸商標異議(卷宗性質為2.異議)復審(401)時,不抓Casefee設定寫死在程式裡
   If Val(strSrvDate(1)) >= 20140501 And m_TM28 = "2" And m_TM10 = "020" And m_CP10 = "401" Then
      textCF09 = "15個月"
   Else
   '2014/4/2 END
      strSql = "SELECT * FROM CaseFee " & _
               "WHERE CF01 = '" & m_TM01 & "' AND " & _
                     "CF02 = '" & m_TM10 & "' AND " & _
                     "CF03 = '" & m_CP10 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If IsNull(rsTmp.Fields("CF09")) = False Then
            textCF09 = rsTmp.Fields("CF09")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   'Add By Sindy 2020/11/13 214.陳述聲明,不列印定稿
   If m_CP10 = "214" Then
      textPrint = "N" '不印定稿
   End If
   '2020/11/13 END
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N" '不印定稿
   End If
   '2025/8/11 END
   
   'Add By Sindy 2017/3/31 排除FCT案
   If textPrint = "3" And Left(GetST15(m_CP13), 2) = "P2" Then '英文
      textRecvDate.Enabled = True
      textRecvDate.BackColor = &H80000005 '白
   Else
      textRecvDate.Enabled = False
      textRecvDate.BackColor = &H8000000F '灰
   End If
   '2017/3/31 END
   
   'Add By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
   'Modified by Lydia 2017/10/24 外商爭議案403行政訴訟會由內商人員發文
   'If m_TM01 = "T" And m_TM10 = "000" Then
   'Modify by Amy 2020/01/23 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   If m_TM10 = "000" Then
      Label56.Visible = True
      textCP118.Visible = True
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
        lblPayToday.Visible = True
        txtPayToday.Visible = True
      End If
   'end 2020/01/13
   Else
      Label56.Visible = False
      textCP118.Visible = False
   End If
   '2011/10/28 End
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2022/05/23
   
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_15 = Nothing
End Sub

'add by nickc 2006/01/27
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modify by Amy 2021/12/27 改Form2.0,使用PUB_Num2Id會錯
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         m_AgentName = m_AgentName & "、" & lstNameAgent.List(ii)
         'end 2021/12/27

         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      textCP22 = ""
   Else
      textCP22 = "N"
   End If
   'Add By Sindy 2015/7/22
   If textCP118 = "Y" And textCP22 = "N" Then
      Cancel = True
      MsgBox "電子送件時不可為不出名!!!", vbExclamation, "資料檢核"
      Me.SSTab1.Tab = 0
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP14_R_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 相關總收文號承辦人
Private Sub textCP14_R_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textCP14_R2 = Empty
   If IsEmptyText(textCP14_R) = False Then
      textCP14_R2 = GetStaffName(textCP14_R, True)
      If IsEmptyText(textCP14_R2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP14_R_GotFocus
      End If
   End If
End Sub

' 是否出名
Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'edit by nickc 2006/01/27
' 是否出名
'Private Sub textCP22_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP22) = False Then
'      Select Case textCP22
'         Case " ", "N":
'         Case Else
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入空白或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP22_GotFocus
'      End Select
'   End If
'End Sub

' 相關總收文號預估勝敗
Private Sub textCP23_R_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP23_R) = False Then
      Select Case textCP23_R
         Case "1", "2", "3": 'Modify By Sindy 98/04/13 增加3
         Case Else
            strTit = "檢核資料"
            strMsg = "相關總收文號預估勝敗只可輸入1或2或3" 'Modify By Sindy 98/04/13 增加3
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23_R_GotFocus
      End Select
   End If
End Sub

' 預估勝敗
Private Sub textCP23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP23) = False Then
      Select Case textCP23
         Case "1", "2", "3": 'Modify By Sindy 98/04/13 增加3
         Case Else
            strTit = "檢核資料"
            strMsg = "預估勝敗只可輸入1或2或3" 'Modify By Sindy 98/04/13 增加3
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23_GotFocus
      End Select
   'Add by Morgan 2003/12/04
   '案件性質為4開頭且申請國家為台灣時
   ElseIf Left(m_CP10, 1) = "4" And m_TM10 = 台灣國家代號 Then
      If m_CP10 <= "417" Then  'add by sonia 2024/12/16
         MsgBox "預估勝敗不可空白！", vbCritical, strTit
         Cancel = True
      End If                   'add by sonia 2024/12/16
   End If
End Sub

'Add By Sindy 2009/06/11
Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
End Sub

Private Sub textCP37_GotFocus()
   InverseTextBox textCP37
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP37.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP38_GotFocus()
   InverseTextBox textCP38
End Sub

Private Sub textCP39_GotFocus()
   InverseTextBox textCP39
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP39.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP40.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP42.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP80_GotFocus()
   InverseTextBox textCP80
End Sub

' 對造號數
Private Sub textCP36_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modify by Amy 2022/09/29 原:20 放寬至200
   If CheckLengthIsOK(textCP36, textCP36.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造號數內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP36_GotFocus
   End If
   If GetTextLength(textCP36) > 20 And InStr(textCP36, ",") = 0 And InStr(textCP36, ";") = 0 Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "對造號數內容太長,無法寫入基本檔欄位中,請洽電腦中心"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textCP36_GotFocus
   End If
   'end 2022/09/29
End Sub

Private Sub textCP37_1_GotFocus()
    TextInverse textCP37_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
    
    Cancel = False
    If CheckLengthIsOK(textCP37_1, 140) = False Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "對造案件名稱內容太長"
        textCP37_1_GotFocus
    End If
    'edit by nickc 2007/06/06 切換輸入法改用API
    If Cancel = False Then CloseIme
End Sub

' 對造案件中文名稱
Private Sub textCP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP37, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP37_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP37.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造案件英文名稱
Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP38, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP38_GotFocus
   End If
End Sub

' 對造案件日文名稱
Private Sub textCP39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP39, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP39_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP39.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造中文名稱
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP40.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 對造英文名稱
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

' 對造日文名稱
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP42.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP80_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textCP80 = Replace(textCP80, " ", "")
End Sub
'2009/06/11 End

Private Sub textCP43_LostFocus()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCP14_R = Empty
   textCP14_R2 = Empty
   textCP23_R = Empty
   textCP64_R = Empty
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本案之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         Exit Sub
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         Exit Sub
      Else
         If IsNull(rsTmp.Fields("CP14")) = False Then
            textCP14_R = rsTmp.Fields("CP14")
            If IsEmptyText(textCP14_R) = False Then
               textCP14_R_Validate False
            End If
         End If
         If IsNull(rsTmp.Fields("CP23")) = False Then
            textCP23_R = rsTmp.Fields("CP23")
         End If
         If IsNull(rsTmp.Fields("CP64")) = False Then
            textCP64_R = rsTmp.Fields("CP64")
         End If
      End If
      rsTmp.Close
   End If
End Sub

' 相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   '92.10.29 cancel by sonia 轉至lostfocus
   'textCP14_R = Empty
   'textCP14_R2 = Empty
   'textCP23_R = Empty
   'textCP64_R = Empty
   '92.10.29 end
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本案之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_TM01 & "' AND " & _
                     "CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND " & _
                     "CP04 = '" & m_TM04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP43_GotFocus
         GoTo EXITSUB
      Else
         '92.10.29 cancel by sonia 轉至lostfocus
         'If IsNull(rsTmp.Fields("CP14")) = False Then
         '   textCP14_R = rsTmp.Fields("CP14")
         '   If IsEmptyText(textCP14_R) = False Then
         '      textCP14_R_Validate False
         '   End If
         'End If
         'If IsNull(rsTmp.Fields("CP23")) = False Then
         '   textCP23_R = rsTmp.Fields("CP23")
         'End If
         'If IsNull(rsTmp.Fields("CP64")) = False Then
         '   textCP64_R = rsTmp.Fields("CP64")
         'End If
         '92.10.29 end
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/03
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Modify by Amy 2021/12/27 原:Integer
Private Sub textCP49_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 條款
Private Sub textCP49_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim strTemp As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Cancel = False
   ' 無資料時不做任何檢查
   If IsEmptyText(textCP49) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textCP49)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textCP49, nIndex)
        'Modify By Cheng 2002/12/03
        '非大陸案
        If m_TM10 <> 大陸國家代號 Then
            'Modify By Cheng 2002/07/22
            '條款每項必須輸入4碼
            '      If Len(strTemp) > 4 Then
            'Modify By Sindy 2012/7/5
            'If Len(strTemp) <> 4 Then
            If Len(strTemp) <> 4 And Len(strTemp) <> 5 Then
            '2012/7/5 End
               Cancel = True
               strTit = "條款"
               strMsg = "條款內容<" & strTemp & ">不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP49_GotFocus
               GoTo EXITSUB
            End If
            ' 檢查主張內容分類表
            strSql = "SELECT * FROM ClaimContents " & _
                     "WHERE CC01 = '" & Right(strTemp, 1) & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
            If rsTmp.RecordCount <= 0 Then
               Cancel = True
               strTit = "條款"
               strMsg = "條款內容<" & strTemp & ">不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP49_GotFocus
               rsTmp.Close
               GoTo EXITSUB
            End If
            rsTmp.Close
        '大陸案
        Else
            'Add By Cheng 2002/12/03
            If Len(strTemp) <> 3 Then
               Cancel = True
               strTit = "條款"
               strMsg = "條款內容<" & strTemp & ">不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textCP49_GotFocus
               GoTo EXITSUB
            End If
        End If
        ' 檢查
        'Modify By Sindy 2012/7/5
'        strSql = "SELECT * FROM LAW " & _
'                 "WHERE LW01 = '" & Mid(strTemp, 1, 3) & "' "
        '非大陸案
        If m_TM10 <> 大陸國家代號 Then
            strSql = "SELECT * FROM LAW " & _
                     "WHERE LW01 = '" & Mid(strTemp, 1, Len(strTemp) - 1) & "' "
        '大陸案
        Else
            strSql = "SELECT * FROM LAW " & _
                     "WHERE LW01 = '" & Trim(strTemp) & "' "
        End If
        '2012/7/5 End
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
        If rsTmp.RecordCount <= 0 Then
           Cancel = True
           strTit = "條款"
           strMsg = "條款代號<" & strTemp & ">不存在"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP49_GotFocus
           rsTmp.Close
           GoTo EXITSUB
        End If
        rsTmp.Close
   Next nIndex
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 進度備註
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

Private Sub textCP84_GotFocus()
   InverseTextBox textCP84
End Sub

Private Sub textCP84_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
If IsEmptyText(textCP84) = False Then
    If IsNumeric(textCP84) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入數字"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP84_GotFocus
    Else
        textCP84.Text = Trim(Val(textCP84.Text))
    End If
End If
End Sub

' 下一程序
Private Sub textNP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP07) = False Then
      If m_TM10 = "000" Then
         textNP07_2 = GetCaseTypeName(m_TM01, textNP07, 0)
      Else
         textNP07_2 = GetCaseTypeName(m_TM01, textNP07, 1)
      End If
      If IsEmptyText(textNP07_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "下一程序代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP07_GotFocus
         GoTo EXITSUB
      End If
      
      EnableTextBox textNP08, True
      EnableTextBox textNP09, True
   Else
      EnableTextBox textNP08, False
      EnableTextBox textNP09, False
   End If
   
EXITSUB:
End Sub

' 發文日
Private Sub textCP27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nick 2006/06/22 系統日加一天
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nick 2006/06/22
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      'Add by Amy 2020/01/13 當發文日有改時,電子送件案要人工輸入是否當日扣款
      If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        If textCP27.Tag <> textCP27.Text Then
            textCP27.Tag = textCP27.Text
            If textCP118 = "Y" Then
                txtPayToday.Text = ""
            End If
        End If
      End If
      'end 2020/01/13
   End If
EXITSUB:
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

' 代理人
Private Sub textCP44_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   'Add By Cheng 2002/07/09
   Dim strTempName As String
   
   Cancel = False
   'Add By Cheng 2002/03/08
   If m_TM10 <> 台灣國家代號 Then
      If Len(Me.textCP44.Text) <= 0 Then
         MsgBox "當申請國家非台灣時, 代理人欄不可為空白!!!", vbExclamation
         Cancel = True
         Exit Sub
      End If
   End If
   
   If textCP44.ListIndex >= 0 Then
      textCP44 = m_AgentList(textCP44.ListIndex).aiCode
   End If
   'Add By Cheng 2002/12/03
   '若有輸入代理人則將代碼補滿9碼
   If Me.textCP44.Text <> "" Then Me.textCP44.Text = Left(Me.textCP44.Text & "000000000", 9)
   
   If IsEmptyText(textCP44) = False Then
      'Modify By Cheng 2002/07/09
'      textCP44_2 = GetFAgentName(textCP44)
      'edit by 2004/07/22 nick  檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If PUB_GetAgentName(m_TM01, Me.textCP44.Text, strTempName) Then
      If PUB_GetAgentNameAndState(m_TM01, Me.textCP44.Text, strTempName) Then
         textCP44_2 = strTempName
      Else
         textCP44_2 = ""
         If strTempName <> "" Then
                Cancel = True
                Exit Sub
         End If
      End If
      If IsEmptyText(textCP44_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "代理人不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44_GotFocus
      Else
         ' 依所選擇的代理人找出案件進度檔中其收文日最大的一筆其彼所案號更新到畫面上的彼所案號欄位
         strSql = "SELECT CP45 FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP44 = '" & textCP44 & "' AND " & _
                        "CP05 IN (SELECT MAX(CP05) FROM CASEPROGRESS " & _
                                 "WHERE CP01 = '" & m_TM01 & "' AND " & _
                                       "CP02 = '" & m_TM02 & "' AND " & _
                                       "CP03 = '" & m_TM03 & "' AND " & _
                                       "CP04 = '" & m_TM04 & "' AND " & _
                                       "CP44 = '" & textCP44 & "')"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CP45")) = False Then
               textTM45 = rsTmp.Fields("CP45")
            End If
         End If
         rsTmp.Close
      End If
   End If
   Set rsTmp = Nothing
End Sub

' 本所期限(起)
Private Sub textNP08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP08) = False Then
      ' 本所期限日期不正確
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08.SetFocus
         textNP08_GotFocus
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textNP08.Text = TransDate(PUB_GetWorkDay1(textNP08, True), 1)
      'end 2020/07/07
      End If
   End If
End Sub

' 本所期限(迄)
Private Sub textNP09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textNP09) = False Then
      ' 本所期限日期不正確
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
      End If
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         'edit by nickc 2006/06/29
         'Case " ", "N":
         Case "N", "1", "2", "3":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            'edit by nickc 2006/06/29
            'strMsg = "只可輸入空白或N"
            strMsg = "只可輸入N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
            Exit Sub
      End Select
      'Add By Sindy 2017/3/31 排除FCT案
      If textPrint = "3" And Left(GetST15(m_CP13), 2) = "P2" Then '英文
         textRecvDate.Enabled = True
         textRecvDate.BackColor = &H80000005 '白
      Else
         textRecvDate.Text = ""
         textRecvDate.Enabled = False
         textRecvDate.BackColor = &H8000000F '灰
      End If
      '2017/3/31 END
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim strCP64 As String
   
   ' 預估結果
   SetCPFieldNewData "CP23", textCP23
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 代理人代號
   If IsEmptyText(textCP44) = False Then
      SetCPFieldNewData "CP44", textCP44 & String(9 - Len(textCP44), "0")
      'add by nickc 2008/02/22
      m_CP44New = textCP44 & String(9 - Len(textCP44), "0")
   Else
      SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   End If
   ' 彼所案號
    'Modify By Cheng 2002/12/31
    '內商台對大陸的案件才要更新案件進度檔的彼所案號資料
'   SetCPFieldNewData "CP45", textTM45
    If m_TM10 <> 台灣國家代號 Then
       SetCPFieldNewData "CP45", textTM45
    End If
   ' 條款
   SetCPFieldNewData "CP49", textCP49
   ' 91.09.02 modify by louis
   ' 進度備註
   'SetCPFieldNewData "CP64", textCP64
   strCP64 = textCP64
'edit by nickc 2006/01/27
'   If IsEmptyText(textAgName) = False Then
'      strCP64 = strCP64 & "," & "本所出名代理人:" & textAgName
'   End If
   SetCPFieldNewData "CP64", strCP64
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   'add by nickc 2006/01/27
   SetCPFieldNewData "CP110", m_CP110
   
   'Add By Sindy 2009/06/11
   ' 對造號數
   SetCPFieldNewData "CP36", textCP36
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        ' 對造案件名稱
        SetCPFieldNewData "CP37", textCP37_1
    Case Else
        ' 對造案件名稱(中)
        SetCPFieldNewData "CP37", textCP37
        ' 對造案件名稱(英)
        SetCPFieldNewData "CP38", textCP38
        ' 對造案件名稱(日)
        SetCPFieldNewData "CP39", textCP39
    End Select
   ' 對造名稱(中)
   SetCPFieldNewData "CP40", textCP40
   ' 對造名稱(英)
   SetCPFieldNewData "CP41", textCP41
   ' 對造名稱(日)
   SetCPFieldNewData "CP42", textCP42
   ' 對造案件商品類別
   SetCPFieldNewData "CP80", textCP80
   '2009/06/11 End
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

' 更新案件進度檔
'Modify By Cheng 2002/11/07
'Private Sub OnUpdateCaseProgress()
Private Function OnUpdateCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateCaseProgress = True

   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
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
   ' 設定SQL語法更新的條件
   strSql = strSql & " " & _
                  "WHERE CP09 = '" & m_CP09 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/06
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP22 As String
   Dim objCopyCP As ClsCopyCP
   Dim strCP09 As String, strCP06 As String, strCP07 As String, strCP48 As String 'Add By Sindy 2010/10/12
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler

cnnConnection.BeginTrans
   
   'Add By Sindy 2010/12/28
   '非台灣案發文, 法定期限有值且為系統日或者過期時, 顯示訊息, 但仍可發文
   '上述情形的收達期限或提申期限都管制為系統日期
   bolSysDt = False
   If m_TM10 >= "010" Then
      If Trim(m_CP07) <> "" Then
         If Val(m_CP07) = Val(strSrvDate(1)) Then
            MsgBox "此案件已屆法定期限, 請注意！", vbExclamation + vbOKOnly
            bolSysDt = True
         ElseIf Val(m_CP07) < Val(strSrvDate(1)) Then
            MsgBox "此案件已逾法定期限, 請注意！", vbExclamation + vbOKOnly
            bolSysDt = True
         End If
      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
   'Modify By Cheng 2002/11/06
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   
    'add by nickc 2006/11/17
    Select Case m_TM01
    Case "T", "TF", "FCT"
        If textPrint <> "N" Then
            strSql = "Update TradeMark Set TM77='" & textPrint & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    Case Else
        If textPrint <> "N" Then
            strSql = "Update ServicePractice Set SP72='" & textPrint & "' " & _
                            " Where " & ChgService(m_TM01 & m_TM02 & m_TM03 & m_TM04)
            cnnConnection.Execute strSql
        End If
    End Select
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔的放棄專用權
   Select Case m_TM01
      ' 系統類別為CFT的為儲存商標基本檔
      Case "T", "TF", "FCT":
         ' 91.03.25 modify by louis (放棄專用權有單引號)
         ' strSQL = "UPDATE TradeMark SET TM67 = '" & chgsql(textTM67) & "' "
         strSql = "UPDATE TradeMark SET TM67 = '" & ChgSQL(textTM67) & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
         'add by ncick 2006/01/26
         If m_CU112 <> "" Then
            'Modify By Sindy 2011/2/22
'            strSql = "UPDATE TradeMark SET TM24 = '" & ChgSQL(Pub_RplCu112(m_TM24, m_CU112)) & "' " & _
'                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                           "TM02 = '" & m_TM02 & "' AND " & _
'                           "TM03 = '" & m_TM03 & "' AND " & _
'                           "TM04 = '" & m_TM04 & "' "
            strSql = "UPDATE TradeMark SET TM24 = '" & ChgSQL(Pub_RplCu112(m_TM24, m_CU112, m_TM23)) & "' " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "TM04 = '" & m_TM04 & "' "
            cnnConnection.Execute strSql
         End If
      Case Else:
   End Select
   
   
   'Add By Sindy 2009/06/11
    '若案件性質為"申請意見書"
    '2011/11/4 modify by sonia 卷宗性質非申請者改用210陳述意見書
    'If m_CP10 = "202" And m_TM28 <> "1" Then
    If m_CP10 = "210" Then
        '更新商標基本檔的案件中英日文名稱,申請案號,商品類別
        'Modify by Amy 2022/09/29 +GetCP36,避免cp36欄位放寬,導致寫入其他欄位錯誤
        strSql = "Update Trademark Set TM05='" & ChgSQL(Me.textCP37_1.Text) & "',TM12='" & GetCP36(Me.textCP36.Text) & "',TM09='" & Me.textCP80.Text & "' " & _
                        "Where TM01='" & m_TM01 & "' And TM02='" & m_TM02 & "' And TM03='" & m_TM03 & "' And TM04='" & m_TM04 & "'"
        cnnConnection.Execute strSql
    End If
   '2009/06/11 End
   
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 案件性質為申請意見書或補充理由時更新相關總收文號的記錄其承辦人代號, 預估勝敗, 進度備註
   Select Case m_CP10
      '2011/11/4 modify by sonia 加210陳述意見書
      Case "202", "612", "210":
         If IsEmptyText(textCP43) = False Then
            strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14_R & "', " & _
                                             "CP23 = '" & textCP23_R & "', " & _
                                             "CP49 = '" & textCP49 & "', " & _
                                             "CP64 = '" & textCP64_R & "' " & _
                     "WHERE CP09 = '" & textCP43 & "' "
            cnnConnection.Execute strSql
         End If
      Case Else:
   End Select
      
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下一程序時, 新增一筆資料到下一程序檔
   ' 收文號
   If IsEmptyText(textNP07) = False Then
      strNP22 = GetNextProgressNo()
        'Modify By Cheng 2003/11/24
        '重抓智權人員
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textNP07 & "," & _
'                          DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & m_CP13 & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textNP07 & "," & _
                          DBDATE(textNP08) & "," & DBDATE(textNP09) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case textNP07
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & m_CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF23")) = False Then
         strNP07 = "997"
         'Add By Sindy 2010/12/28
         '非台灣案發文, 法定期限有值且為系統日或者過期時, 收達期限或提申期限都管制為系統日期
         If bolSysDt = True Then
            strNP08 = strSrvDate(1)
         Else
         '2010/12/28 End
           'Modify By Cheng 2003/09/01
   '         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
            strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(textCP27))))
            'Add By Sindy 2019/6/11 檢查期限是否正確
            strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
            '2019/6/11 END
         End If
         strNP22 = GetNextProgressNo()
         'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
         
         If Not (m_TM01 = "T" And m_TM10 = "020" And (m_CP10 = "401" Or m_CP10 = "403")) Then
            'Add By Sindy 2022/6/7 判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限
            If IsNull(rsTmp.Fields("CF11")) = False Then
               strNP07 = "998"
               '非台灣案發文, 法定期限有值且為系統日或者過期時, 收達期限或提申期限都管制為系統日期
               If bolSysDt = True Then
                  strNP08 = strSrvDate(1)
               Else
                  strNP08 = DBDATE(textCP27)
                  strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF11")), ChangeWStringToWDateString(DBDATE(strNP08))))
                  '檢查期限是否正確
                  strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
               End If
               strNP22 = GetNextProgressNo()
               '本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                  PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
               cnnConnection.Execute strSql
            End If
            '2022/6/7 END
         End If
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
         '92.6.8 SONIA 加 言詞辯論, 準備程序
         Select Case strNP07
'            Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
            Case "102", "105", "702", "708", "305", "998", "997"
            Case Else:
               ' 列印國內案件接洽及結案記錄單
'               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
                'Add By Cheng 2004/04/08
                '新增列印接洽結案單資料
                pub_AddressListSN = pub_AddressListSN + 1
                PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
         End Select
      End If
      'Add By Sindy 2010/3/31
      ' 若有審查天數, 新增一筆催審期限的記錄到下一程序檔
      If IsNull(rsTmp.Fields("CF05")) = False Then
         strNP07 = "305"
         'Modify By Sindy 2014/5/1 從2014/05/01起,大陸商標異議(卷宗性質為2.異議)復審(401)時,不抓Casefee設定寫死在程式裡
         If Val(strSrvDate(1)) >= 20140501 And m_TM28 = "2" And m_TM10 = "020" And m_CP10 = "401" Then
            strNP08 = DBDATE(DateAdd("d", 460, ChangeWStringToWDateString(DBDATE(textCP27))))
         Else
         '2014/4/2 END
            strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
         End If
         strNP22 = GetNextProgressNo()
         'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      '2010/3/31 End
   End If
   rsTmp.Close
   
   'Add By Sindy 2010/12/22
   '內商大陸案案件性質401,403,408發文時,
   '請管制提申期限為法定期限(cp07)-2天,若法定期限(cp07)-2天<系統日管制提申期限為系統日
   If m_TM01 = "T" And m_TM10 = "020" And (m_CP10 = "401" Or m_CP10 = "403") Then
      strNP07 = "998"
      'Add By Sindy 2010/12/28
      '非台灣案發文, 法定期限有值且為系統日或者過期時, 收達期限或提申期限都管制為系統日期
      If bolSysDt = True Then
         strNP08 = strSrvDate(1)
      Else
      '2010/12/28 End
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(m_CP07)))
'         If Val(strNP08) < Val(strSrvDate(1)) Then
'            strNP08 = strSrvDate(1)
'         End If
         'Add By Sindy 2019/6/11 檢查期限是否正確
         strNP08 = PUB_T997998LimitDate(strNP08, m_CP07, 1)
         '2019/6/11 END
      End If
      strNP22 = GetNextProgressNo()
      'Modified by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                         strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                         PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      
      'Add By Sindy 2010/10/12
      'Modify By Sindy 2011/1/25
      'If m_TM10 = "020" And m_CP10 = "401" Then
      If MsgBox("是否管制7個工作天的寄委任狀期限？", vbYesNo) = vbYes Then
         strCP09 = AutoNo("B", 6)
         '本所期限=發文日+7個工作天
         strCP06 = CompWorkDay(8, DBDATE(textCP27), 0)
         '法定期限=發文日+7個工作天
         strCP07 = CompWorkDay(8, DBDATE(textCP27), 0)
         '承辦期限=發文日+7個工作天
         strCP48 = CompWorkDay(8, DBDATE(textCP27), 0)
         strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP26,CP32,CP43,CP48,CP64,CP20) " & _
                      "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strSrvDate(1) & "," & strCP06 & "," & strCP07 & "," & _
                      "'" & strCP09 & "','706','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & _
                      "'N','N','" & m_CP09 & "'," & strCP48 & ",'寄委任狀','N') "
         cnnConnection.Execute strSql
      End If
      
      '2012/9/19 add by sonia 大陸部分核駁之復審發文時要將原申請案之催審上N T-174043
      If m_CP10 = "401" Then
         strSql = "UPDATE NEXTPROGRESS SET NP06='N' WHERE NP02='" & m_TM01 & "' AND NP03='" & m_TM02 & "' AND NP04='" & m_TM03 & "' AND NP05='" & m_TM04 & "'" & _
                  " AND NP01 IN (SELECT CP43 FROM CASEPROGRESS WHERE CP10='1205' AND CP09 IN (SELECT CP43 FROM CASEPROGRESS WHERE CP09='" & m_CP09 & "')) AND NP07=305 AND NP06 IS NULL"
         cnnConnection.Execute strSql
      End If
      '2012/9/19 end
   End If
   '2010/12/22 End
   
    'add by nick 2004/08/12 更新實際發文規費
    If textCP84.Enabled = True Then
         strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
         cnnConnection.Execute strSql
    End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
        'Modify By Cheng 2002/11/07
'      objCopyCP.CopyCaseProgress m_CP09
      If objCopyCP.CopyCaseProgress(m_CP09) = False Then GoTo ErrorHandler
      Set objCopyCP = Nothing
   End If
   'add by nick 2004/09/27 存公司負責人英文名稱
   'edit by nick 2004/10/07
   'If m_CU103 <> "" And m_TM01 <> "FCT" Then
   'edit by nickc 2006/01/20
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "") And m_TM01 <> "FCT" Then
   'edit by nickc 2007/08/10
   'If (m_CU103 <> "" Or (m_CU05 & m_CU88 & m_CU89 & m_CU90) <> "" Or m_CU112 <> "") And m_TM01 <> "FCT" Then
   'Modify By Sindy 2012/10/31 +SeekCu10(1),SeekCu10(2),SeekCu10(3),SeekCu10(4),SeekCu10(5)
   If (SeekCu103(1) <> "" Or (SeekCu05(1) & SeekCu88(1) & SeekCu89(1) & SeekCu90(1)) <> "" Or SeekCu112(1) <> "" Or (SeekCu39(1) & SeekCu40(1) & SeekCu41(1)) <> "" Or SeekCu10(1) <> "") And m_TM01 <> "FCT" Then
            'edit by nickc 2006/01/20
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            'edit by nickc 2007/08/10
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'  Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(1)) & "',cu05='" & ChgSQL(SeekCu05(1)) & "',cu88='" & ChgSQL(SeekCu88(1)) & "',cu89='" & ChgSQL(SeekCu89(1)) & "',cu90='" & ChgSQL(SeekCu90(1)) & "',cu112='" & ChgSQL(SeekCu112(1)) & "',cu39='" & ChgSQL(SeekCu39(1)) & "',cu40='" & ChgSQL(SeekCu40(1)) & "',cu41='" & ChgSQL(SeekCu41(1)) & "',cu10='" & ChgSQL(SeekCu10(1)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(1)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   'add by nickc 2007/08/10 加多申請人也要
   If (SeekCu103(2) <> "" Or (SeekCu05(2) & SeekCu88(2) & SeekCu89(2) & SeekCu90(2)) <> "" Or SeekCu112(2) <> "" Or (SeekCu39(2) & SeekCu40(2) & SeekCu41(2)) <> "" Or SeekCu10(2) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(2)) & "',cu05='" & ChgSQL(SeekCu05(2)) & "',cu88='" & ChgSQL(SeekCu88(2)) & "',cu89='" & ChgSQL(SeekCu89(2)) & "',cu90='" & ChgSQL(SeekCu90(2)) & "',cu112='" & ChgSQL(SeekCu112(2)) & "',cu39='" & ChgSQL(SeekCu39(2)) & "',cu40='" & ChgSQL(SeekCu40(2)) & "',cu41='" & ChgSQL(SeekCu41(2)) & "',cu10='" & ChgSQL(SeekCu10(2)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM78), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM78), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(2)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM78), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM78), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(3) <> "" Or (SeekCu05(3) & SeekCu88(3) & SeekCu89(3) & SeekCu90(3)) <> "" Or SeekCu112(3) <> "" Or (SeekCu39(3) & SeekCu40(3) & SeekCu41(3)) <> "" Or SeekCu10(3) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(3)) & "',cu05='" & ChgSQL(SeekCu05(3)) & "',cu88='" & ChgSQL(SeekCu88(3)) & "',cu89='" & ChgSQL(SeekCu89(3)) & "',cu90='" & ChgSQL(SeekCu90(3)) & "',cu112='" & ChgSQL(SeekCu112(3)) & "',cu39='" & ChgSQL(SeekCu39(3)) & "',cu40='" & ChgSQL(SeekCu40(3)) & "',cu41='" & ChgSQL(SeekCu41(3)) & "',cu10='" & ChgSQL(SeekCu10(3)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM79), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM79), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(3)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM79), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM79), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(4) <> "" Or (SeekCu05(4) & SeekCu88(4) & SeekCu89(4) & SeekCu90(4)) <> "" Or SeekCu112(4) <> "" Or (SeekCu39(4) & SeekCu40(4) & SeekCu41(4)) <> "" Or SeekCu10(4) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(4)) & "',cu05='" & ChgSQL(SeekCu05(4)) & "',cu88='" & ChgSQL(SeekCu88(4)) & "',cu89='" & ChgSQL(SeekCu89(4)) & "',cu90='" & ChgSQL(SeekCu90(4)) & "',cu112='" & ChgSQL(SeekCu112(4)) & "',cu39='" & ChgSQL(SeekCu39(4)) & "',cu40='" & ChgSQL(SeekCu40(4)) & "',cu41='" & ChgSQL(SeekCu41(4)) & "',cu10='" & ChgSQL(SeekCu10(4)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM80), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM80), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(4)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM80), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM80), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   If (SeekCu103(5) <> "" Or (SeekCu05(5) & SeekCu88(5) & SeekCu89(5) & SeekCu90(5)) <> "" Or SeekCu112(5) <> "" Or (SeekCu39(5) & SeekCu40(5) & SeekCu41(5)) <> "" Or SeekCu10(5) <> "") And m_TM01 <> "FCT" Then
            strSql = "Update customer Set CU103='" & ChgSQL(SeekCu103(5)) & "',cu05='" & ChgSQL(SeekCu05(5)) & "',cu88='" & ChgSQL(SeekCu88(5)) & "',cu89='" & ChgSQL(SeekCu89(5)) & "',cu90='" & ChgSQL(SeekCu90(5)) & "',cu112='" & ChgSQL(SeekCu112(5)) & "',cu39='" & ChgSQL(SeekCu39(5)) & "',cu40='" & ChgSQL(SeekCu40(5)) & "',cu41='" & ChgSQL(SeekCu41(5)) & "',cu10='" & ChgSQL(SeekCu10(5)) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM81), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM81), 9, 1) & "' "
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            'Add By Sindy 2013/11/15
            'Modify By Sindy 2025/6/11 排除個人客戶不可更新負責人 => + and CU15<>'0'
            strSql = "Update customer Set CU07='" & Left(ChgSQL(SeekCu39(5)), 30) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM81), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM81), 9, 1) & "' and CU07 is null and CU15<>'0'"
            Pub_SeekTbLog strSql 'Add By Sindy 2021/5/27
            cnnConnection.Execute strSql
            '2013/11/15 END
   End If
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   Set rsTmp = Nothing
   
   'Add By Sindy 2012/3/23
   Call PUB_T020InsB301(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP44, m_TM10, m_CP10, textCP27, m_CP12, m_CP13, textTM45)
   
   'Add by Amy 2020/01/13
   If strSrvDate(1) >= T商標電子送件扣款啟用日 And textCP118.Visible = True Then
        strSql = ""
        If textCP118 = "Y" And Val(textCP84) > 0 Then
           If txtPayToday <> "" Then
              strSql = ",CP118 = 'A' "
              If txtPayToday = "Y" Then
                  strSql = strSql & ",CP152 = " & CompWorkDay(2, DBDATE(textCP27))
              Else
                  strSql = strSql & ",CP152 =" & CompWorkDay(3, DBDATE(textCP27))
              End If
              strSql = "Update CaseProgress Set " & Mid(strSql, 2) & " Where CP09 = '" & m_CP09 & "' "
              cnnConnection.Execute strSql
           End If
        End If
   End If
   'end 2020/01/13
   
   'Add By Sindy 2011/3/9 若為電子送件則自動設定為不經發文室
   '以防動作為重新發文, 所以一併把發文室相關欄位清空
   If textCP118.Visible = True And textCP118 = "Y" Then
      strSql = "Update CaseProgress Set CP123=null" & _
                                                          ",CP124=null" & _
                                                          ",CP125=null" & _
                                                          ",CP28=null" & _
                                                          ",CP131=null" & _
                                                          ",CP132=null" & _
                   " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Added by Lydia 2016/03/15 台灣案T案 申請意見書202發文存檔時,將該案號申請101那一道的下一程序催審305期限 (NP06 Is Null)更新 NP10 (催審人員) 為申請意見書之CP14.
   If m_TM01 = "T" And m_TM10 = "000" And m_CP10 = "202" Then
      strSql = "UPDATE NEXTPROGRESS SET NP10=" & CNULL(m_CP14) & _
               " WHERE (NP01,NP22) IN (SELECT NP01,NP22 FROM NEXTPROGRESS,CASEPROGRESS C1" & _
               " WHERE NP02='" & m_TM01 & "' AND NP03='" & m_TM02 & "' AND NP04='" & m_TM03 & "' AND NP05='" & m_TM04 & "'" & _
               " AND NP07='305' AND NP06 IS NULL AND NP01=C1.CP09(+) AND C1.CP10='101')"
      cnnConnection.Execute strSql
   End If
   'end 2016/03/15
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   'Add by Amy 2019/12/04
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
        'If textPrint <> "N" Then
            'Sindy 2020/2/6 大至台不用判發
            If Not (m_TM10 = "000" And textPrint = "2") Then
            '2020/2/6 END
               strExc(1) = Pub_GetSpecMan("內商程序客戶函發後補看人員")
            Else
               strExc(1) = ""
            End If
            PUB_AddLetterProgress m_CP09, 0, True, , False, m_TM23, m_CP10, m_TM44, , , , , strExc(1)
        'End If
   End If
   'end 2019/12/04
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Added by Lydia 2022/05/23 法律所案源：法務案無發文日則更新發文日為系統日，但法務案不更新CP84。另加發EMAIL給法務案承辦人，提供案號及案件性質、總收文號，提醒他去案件進度檔補輸工作時數及工作點數分配。
   'Modified by Lydia 2022/09/16 限制B2案源
   If m_LOS15 <> "" And m_LOS02 = "B2" Then
       Call PUB_UpdateLosCP27(m_LOS15)
   End If
   'end 2022/05/23
   
'Add By Cheng 2002/11/06
cnnConnection.CommitTrans

     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44New, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22

OnSaveData = True
Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 檢查欄位是否都已輸入或是輸入的值是否正確
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/27檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   'add by nickc 2008/05/01
   If IsDebt(m_TM10, textCP09) Then
        strTit = "警告！禁止發文！"
        strMsg = "未收款且無 預定收款日 請轉告智權同仁！！"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        GoTo EXITSUB
   End If
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 申請國家非台灣時代理人不可空白
   If m_TM10 >= "010" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2017/3/31
   If m_CP10 = "202" Then '申請意見書
      If textRecvDate.Enabled = True Then
         If IsEmptyText(textRecvDate) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入收文指示日期"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textRecvDate.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   '2017/3/31 END
   
   'Modify By Cheng 2002/06/14
   '若案件性質為"補充理由"(612)或申請國家為"大陸"時, 預估勝敗可為空白
   If m_CP10 = "612" Or m_TM10 = 大陸國家代號 Then
      '無動作
   Else
      'Modify By Sindy 2020/11/13 + 214.陳述聲明 不須填預估勝敗
      If m_CP10 <> "214" Then
      '2020/11/13 END
         ' 預估勝敗不可為空白
         If IsEmptyText(textCP23) = True Then
            strTit = "檢核資料"
            strMsg = "請輸入預估勝敗"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP23.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   ' 有輸入下一程序, 本所期限及法定期限不可為空白
   If IsEmptyText(textNP07) = False Then
      If IsEmptyText(textNP08) = True Or IsEmptyText(textNP09) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入本所期限及法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP08.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2009/06/11
    '若案件性質為申請意見書(202)
    '2011/11/4 modify by sonia 卷宗性質非申請者改用210陳述意見書
    'If m_CP10 = "202" And m_TM28 <> "1" Then
    If m_CP10 = "210" Then
        If Me.textCP36.Text = "" Then
            Me.SSTab1.Tab = 1
            MsgBox "請輸入對造號數!!!", vbExclamation + vbOKOnly
            textCP36.SetFocus
            GoTo EXITSUB
        End If
    End If
    '2009/06/11 End
   
   'Add By Sindy 2011/01/06
   '內商(TS)申請人1或FC代理人至少要輸入一個
   '其他的一定要輸入申請人1
   If m_TM01 = "TS" Then
        If textTM23 = "" And m_TM44 = "" Then
            MsgBox "申請人1或FC代理人至少要輸入一個!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   Else
        If textTM23 = "" Then
            MsgBox "申請人1不可空白!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
        End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
End Sub

'Add By Sindy 2017/3/31
Private Sub textRecvDate_GotFocus()
   InverseTextBox textRecvDate
End Sub
' 收文指示日期
Private Sub textRecvDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textRecvDate) = False Then
      ' 收文指示日期不正確
      If CheckIsTaiwanDate(textRecvDate, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的收文指示日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRecvDate_GotFocus
         GoTo EXITSUB
      End If
   End If
EXITSUB:
End Sub
'2017/3/31 END

Private Sub textTM67_GotFocus()
   InverseTextBox textTM67
End Sub

Private Sub textCP14_R_GotFocus()
   InverseTextBox textCP14_R
End Sub

Private Sub textCP23_R_GotFocus()
   InverseTextBox textCP23_R
End Sub

Private Sub textCP64_R_GotFocus()
   InverseTextBox textCP64_R
End Sub

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
End Sub

Private Sub textCP23_GotFocus()
   InverseTextBox textCP23
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textNP07_GotFocus()
   InverseTextBox textNP07
End Sub

Private Sub textNP08_GotFocus()
   InverseTextBox textNP08
End Sub

Private Sub textNP09_GotFocus()
   InverseTextBox textNP09
End Sub

'Add by Amy 2020/01/13
Private Sub txtPayToday_GotFocus()
    TextInverse txtPayToday
    CloseIme
End Sub

Private Sub txtPayToday_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
        KeyAscii = 0
        Beep
    End If
End Sub
'end 2020/01/13

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   Select Case m_CP10
      ' 申請意見書
      Case "202":
         'add by nickc 2005/07/15
         If m_TM01 = "TF" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "23", strUserNum
            End If
         Else
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/29
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "17", strUserNum
                  ' 商標狀況
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                           "'" & "商標狀況" & "','" & "申請" & "')"
                  cnnConnection.Execute strSql
                  ' 回音
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                           "'" & "回音" & "','" & textCF09 & "')"
                  cnnConnection.Execute strSql
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/29
               'Else
               ElseIf textPrint = "2" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "20", strUserNum
               'Add By Sindy 2017/3/31 + 英文定稿 排除FCT案
               ElseIf textPrint = "3" And Left(GetST15(m_CP13), 2) = "P2" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "21", strUserNum
                  ' 收文指示日期
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "21" & "','" & strUserNum & "'," & _
                           "'" & "收文指示日期" & "','" & DBDATE(textRecvDate) & "')"
                  cnnConnection.Execute strSql
               '2017/3/31 END
               End If
            End If
         End If
      ' 訴願
      Case "401":
         ' 系統類別為TF
         If m_TM01 = "TF" Then
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "23", strUserNum
            'Add By Sindy 2010/5/18
            ElseIf textPrint = "2" Then '外至台
                EndLetter "01", m_CP09, "24", strUserNum
            End If
         ' 系統類別非TF
         Else
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/29
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "17", strUserNum
                  ' 商標狀況
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                           "'" & "商標狀況" & "','" & "審定" & "')"
                  cnnConnection.Execute strSql
                  ' 回音
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                           "'" & "回音" & "','" & textCF09 & "')"
                  cnnConnection.Execute strSql
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/29
               'Else
               ElseIf textPrint = "2" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "01", m_CP09, "21", strUserNum
                  ' 列印備註
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "21" & "','" & strUserNum & "'," & _
                           "'" & "列印備註" & "','" & textCP64 & "')"
                  cnnConnection.Execute strSql
               End If
            ' 申請國家為大陸
            ElseIf m_TM10 = "020" Then
               'add by nickc 2006/06/29
               If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "22", strUserNum
               End If
            End If
         End If
      ' 再訴願
      Case "402":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "17", strUserNum
               ' 商標狀況
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                        "'" & "商標狀況" & "','" & "註冊" & "')"
               cnnConnection.Execute strSql
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                        "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "21", strUserNum
               ' 列印備註
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "21" & "','" & strUserNum & "'," & _
                        "'" & "列印備註" & "','" & textCP64 & "')"
               cnnConnection.Execute strSql
            End If
         End If
      ' 行政訴訟
      Case "403":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "17", strUserNum
               ' 商標狀況
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                        "'" & "商標狀況" & "','" & "核駁" & "')"
               cnnConnection.Execute strSql
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                        "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "21", strUserNum
               ' 列印備註
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "21" & "','" & strUserNum & "'," & _
                        "'" & "列印備註" & "','" & textCP64 & "')"
               cnnConnection.Execute strSql
            End If
         End If
      'Add By Cheng 2003/01/23
      '復審答辯
      Case "405"
         'add by nickc 2006/06/22 加入馬德里復審答辯定稿
         If m_TM01 = "TF" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "01", strUserNum
            End If
         Else
            '若申請國家為大陸
            If m_TM10 = 大陸國家代號 Then
                ' 清除定稿例外欄位檔原有資料
                'add by nickc 2006/06/29
                If textPrint = "1" Then
                    EndLetter "01", m_CP09, "22", strUserNum
                End If
            End If
         End If
      'Add By Cheng 2003/01/02
      ' 補充理由
      Case "612":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, "00", strUserNum
                ' 商標狀況
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                         "'" & "商標狀況" & "','" & "註冊" & "')"
                cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, "03", strUserNum
            End If
         End If
      '2005/4/19 CANCEL BY SONIA 補充答辯進入frm020102_16,不會進此FORM
      '' 補充答辯
      'Case "613":
      '   ' 申請國家為台灣
      '   If m_TM10 < "010" Then
      '      ' 清除定稿例外欄位檔原有資料
      '      EndLetter "01", m_CP09, "00", strUserNum
      '      ' 商標狀況
      '      StrSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      '               "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
      '               "'" & "商標狀況" & "','" & "註冊" & "')"
      '      cnnConnection.Execute StrSql
      '   End If
      '2005/4/19 END
      ' 陳情  92.7.11 ADD BY SONIA
      Case "622":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "17", strUserNum
               ' 商標狀況
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                        "'" & "商標狀況" & "','" & "審定" & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "21", strUserNum
               ' 列印備註
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "21" & "','" & strUserNum & "'," & _
                        "'" & "列印備註" & "','" & textCP64 & "')"
               cnnConnection.Execute strSql
            End If
         End If
   End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET02 = m_CP09
   bolEdit = IIf(Trim(txtEditWord) = "Y", True, False)
   '2012/1/12 End
   
   Select Case m_CP10
      ' 申請意見書
      Case "202":
        'add by nickc 2005/07/15
        If m_TM01 = "TF" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "23", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "23" 'Modify By Sindy 2012/1/12
            End If
        Else
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/29
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 列印定稿
                  'edit by nickc 2005/07/15
                  'NowPrint m_CP09, "01", "17", False, strUserNum, 0
'                     NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "17" 'Modify By Sindy 2012/1/12
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/29
               'Else
               ElseIf textPrint = "2" Then
                  ' 列印定稿
                  'edit by nickc 2005/07/15
                  'NowPrint m_CP09, "01", "20", False, strUserNum, 0
'                     NowPrint m_CP09, "01", "20", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "20" 'Modify By Sindy 2012/1/12
               'Add By Sindy 2017/3/31 + 英文定稿  排除FCT案
               ElseIf textPrint = "3" And Left(GetST15(m_CP13), 2) = "P2" Then
                  ET03 = "21"
               '2017/3/31 END
               End If
            End If
         End If
      '2011/11/4 add by sonia
      ' 陳述意見書
      Case "210":
         ' 申請人國籍為台灣
         If textPrint = "1" Then
'            NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
            ET03 = "17" 'Modify By Sindy 2012/1/12
         ' 申請人國籍非台灣
         ElseIf textPrint = "2" Then
'            NowPrint m_CP09, "01", "20", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
            ET03 = "20" 'Modify By Sindy 2012/1/12
         End If
      '2011/11/4 end
      ' 訴願
      Case "401":
         ' 系統類別為TF
         If m_TM01 = "TF" Then
            ' 列印定稿
            'edit by nickc 2005/07/15
            'NowPrint m_CP09, "01", "23", False, strUserNum, 0
            'add by nickc 2006/06/29
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "23", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "23" 'Modify By Sindy 2012/1/12
            'Add By Sindy 2010/5/18
            ElseIf textPrint = "2" Then '外至台
'                NowPrint m_CP09, "01", "24", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "24" 'Modify By Sindy 2012/1/12
            End If
         ' 系統類別非TF
         Else
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               ' 申請人國籍為台灣
               'edit by nickc 2006/06/29
               'If strTM23Nation < "010" Then
               If textPrint = "1" Then
                  ' 列印定稿
                  'edit by nickc 2005/07/15
                  'NowPrint m_CP09, "01", "17", False, strUserNum, 0
'                  NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "17" 'Modify By Sindy 2012/1/12
               ' 申請人國籍非台灣
               'edit by nickc 2006/06/29
               'Else
               ElseIf textPrint = "2" Then
                  ' 列印定稿
                  'edit by nickc 2005/07/15
                  'NowPrint m_CP09, "01", "21", False, strUserNum, 0
'                  NowPrint m_CP09, "01", "21", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "21" 'Modify By Sindy 2012/1/12
               End If
            ' 申請國家為大陸
            ElseIf m_TM10 = "020" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "22", False, strUserNum, 0
               'add by nickc 2006/06/29
               If textPrint = "1" Then
'                    NowPrint m_CP09, "01", "22", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "22" 'Modify By Sindy 2012/1/12
               End If
            End If
         End If
      ' 再訴願
      Case "402":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "17", False, strUserNum, 0
'               NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "17" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "21", False, strUserNum, 0
'               NowPrint m_CP09, "01", "21", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "21" 'Modify By Sindy 2012/1/12
            End If
         End If
      ' 行政訴訟
      Case "403":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "17", False, strUserNum, 0
'               NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "17" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "21", False, strUserNum, 0
'               NowPrint m_CP09, "01", "21", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "21" 'Modify By Sindy 2012/1/12
            End If
         '2007/11/19 add by sonia
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 列印定稿
            If textPrint = "1" Then
'               NowPrint m_CP09, "01", "22", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "22" 'Modify By Sindy 2012/1/12
            End If
         '2007/11/19 end
         End If
      '復審答辯
      Case "405"
         'add by nickc 2006/06/22 加入馬德里復審答辯定稿
         If m_TM01 = "TF" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "01", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "01" 'Modify By Sindy 2012/1/12
            End If
         Else
            '若申請國家為大陸
            If m_TM10 = 大陸國家代號 Then
                ' 列印定稿
                'edit by nickc 2005/07/15
                'NowPrint m_CP09, "01", "22", False, strUserNum, 0
                'add by nickc 2006/06/29
                If textPrint = "1" Then
'                    NowPrint m_CP09, "01", "22", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
                  ET03 = "22" 'Modify By Sindy 2012/1/12
                End If
            End If
         End If
        'Add By Cheng 2003/01/02
      ' 補充理由
      Case "612":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                ' 列印定稿
                'edit by nickc 2005/07/15
                'NowPrint m_CP09, "01", "00", False, strUserNum, 0
'                NowPrint m_CP09, "01", "00", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "00" 'Modify By Sindy 2012/1/12
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
                ' 列印定稿
                'edit by nickc 2005/07/15
                'NowPrint m_CP09, "01", "03", False, strUserNum, 0
'                NowPrint m_CP09, "01", "03", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "03" 'Modify By Sindy 2012/1/12
            End If
         '92.11.6 ADD BY SONIA
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 列印定稿
            'edit by nickc 2005/07/15
            'NowPrint m_CP09, "01", "39", False, strUserNum, 0
            'add by nickc 2006/06/29
            If textPrint = "1" Then
'                NowPrint m_CP09, "01", "39", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "39" 'Modify By Sindy 2012/1/12
            End If
         '92.11.6 END
         End If
      '2005/4/19 CANCEL BY SONIA 補充答辯進入frm020102_16,不會進此FORM
      '' 補充答辯
      'Case "613":
      '   ' 申請國家為台灣
      '   If m_TM10 < "010" Then
      '      ' 列印定稿
      '      NowPrint m_CP09, "01", "00", False, strUserNum, 0
      '   End If
      '2005/4/19 END
      ' 陳情  92.7.11 ADD BY SONIA
      Case "622":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "17", False, strUserNum, 0
'               NowPrint m_CP09, "01", "17", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "17" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               'edit by nickc 2005/07/15
               'NowPrint m_CP09, "01", "21", False, strUserNum, 0
'               NowPrint m_CP09, "01", "21", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "21" 'Modify By Sindy 2012/1/12
            End If
         End If
      '2011/12/6 add by sonia T-157258再審之訴
      Case Else
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            If textPrint = "1" Then
'               NowPrint m_CP09, "01", "30", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "30" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            ElseIf textPrint = "2" Then
'               NowPrint m_CP09, "01", "21", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "21" 'Modify By Sindy 2012/1/12
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            If textPrint = "1" Then
'               NowPrint m_CP09, "01", "22", IIf(Trim(txtEditWord) = "Y", True, False), strUserNum, 0
               ET03 = "22" 'Modify By Sindy 2012/1/12
            End If
         End If
      '2011/12/6 end
   End Select
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Modify by Amy 2019/12/04 +信函收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True, , , , , IIf(strSrvDate(1) >= T商標電子化第2階段啟用日, m_CP09, "")
         MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
      Else
         'Modify by Amy 2019/12/04 +信函收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , IIf(strSrvDate(1) >= T商標電子化第2階段啟用日, m_CP09, "")
      End If
'   'Add By Sindy 2021/1/5 沒有系統產出的定稿
'   Else
'      If m_CP09 <> "" Then
'         'Modify By Sindy 2025/8/15
'         'Call PUB_TCaseAskIsPost(m_CP09)
'         textPrint = "N"
'         '2025/8/15 END
'      End If
'   '2021/1/5 EMD
   End If
   '2012/1/12 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   If Me.textCP84.Enabled = True Then
      Cancel = False
      textCP84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textCP84.Enabled = True And m_TM10 = "000" Then
       'Added by Lydia 2022/05/23 PT案(傳入收文號)取得法律案源之發文規費，並且有輸入發文規費才做檢查
       If m_LosCP84 <> "0" Then
          If Val(m_LosCP84) <> Val(Trim(textCP84.Text)) Then
              If MsgBox("法律所收文規費[" & Trim(Val(m_LosCP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
                  textCP84_GotFocus
                  Exit Function
              End If
          End If
       Else
       'end 2022/05/23
          If Val(textCP84.Text) <> Val(m_CP84) Then
              If MsgBox("收文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
                  textCP84_GotFocus
                  Exit Function
              End If
          End If
       End If 'Added by Lydia 2022/05/23
   End If
   'edit by nickc 2006/01/27
   'If Me.textCP22.Enabled = True Then
   '   Cancel = False
   '   textCP22_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP23.Enabled = True Then
      Cancel = False
      textCP23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP23_R.Enabled = True Then
      Cancel = False
      textCP23_R_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'Add By Sindy 2009/06/11
   If Me.textCP36.Enabled = True Then
      Cancel = False
      textCP36_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP37.Enabled = True Then
      Cancel = False
      textCP37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP37_1.Enabled = True Then
      Cancel = False
      textCP37_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP38.Enabled = True Then
      Cancel = False
      textCP38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP39.Enabled = True Then
      Cancel = False
      textCP39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP40.Enabled = True Then
      Cancel = False
      textCP40_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP41.Enabled = True Then
      Cancel = False
      textCP41_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP42.Enabled = True Then
      Cancel = False
      textCP42_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2009/06/11 End
   
   If Me.textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP49.Enabled = True Then
      Cancel = False
      textCP49_Validate Cancel
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
   
   If Me.textNP07.Enabled = True Then
      Cancel = False
      textNP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP08.Enabled = True Then
      Cancel = False
      textNP08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textNP09.Enabled = True Then
      Cancel = False
      textNP09_Validate Cancel
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
   
   'add by nickc 2006/03/07
   'Modify By Sindy 2012/7/26
   'If lstNameAgent.Enabled = True Then
   If lstNameAgent.Visible = True Then
   '2012/7/26 End
       Cancel = False
       lstNameAgent_Validate Cancel
       If Cancel = True Then
           Exit Function
       End If
   End If
   
   'Add By Sindy 2016/12/20
   '檢查有設定副本收受人需提醒並新增信函副本B類收文
   m_990CP09 = ""
   If textPrint = "N" Then '不印定稿
      If PUB_ChkCC(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_990CP09) = False Then
         SSTab1.Tab = 0
         Exit Function
      End If
   End If
   '2016/12/20 END
   
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        SSTab1.Tab = 0
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   TxtValidate = True
End Function

' 91.09.02 modify by louis
'edit by nickc 2006/01/27
'Private Sub textAgName_GotFocus()
'   InverseTextBox textAgName
'   textAgName.IMEMode = 1
'End Sub
'
'' 91.09.02 modify by louis
'' 本所出名代理人
'Private Sub textAgName_Validate(Cancel As Boolean)
'   Cancel = False
'   If CheckLengthIsOK(textAgName, 10) = False Then
'      Cancel = True
'   End If
'   If Cancel = False Then: textAgName.IMEMode = 2
'End Sub

Private Sub txtEditWord_GotFocus()
   InverseTextBox txtEditWord
End Sub

Private Sub txtEditWord_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtEditWord_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(txtEditWord) = False Then
      Select Case txtEditWord
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            txtEditWord_GotFocus
      End Select
   End If
End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_GotFocus()
   TextInverse txtCP113
End Sub

'Added by Lydia 2021/06/04
Private Sub txtCP113_Validate(Cancel As Boolean)
   If txtCP113 <> "" Then
      If Not IsNumeric(txtCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         txtCP113.SetFocus
         txtCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
End Sub
