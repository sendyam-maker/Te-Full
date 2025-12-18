VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010302_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "各式申請書-讓與"
   ClientHeight    =   7236
   ClientLeft      =   -2532
   ClientTop       =   1416
   ClientWidth     =   9192
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7236
   ScaleWidth      =   9192
   Begin VB.TextBox TextPA178 
      Height          =   270
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   8
      Top             =   5880
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '平面
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   6150
      TabIndex        =   147
      Top             =   4920
      Width           =   2805
      Begin VB.CheckBox chkAtt 
         Caption         =   "稽徵機關核發證明書"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   158
         Tag             =   ".TAX.pdf"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "印鑑切結書"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   157
         Tag             =   ".ATT.pdf"
         Top             =   1410
         Width           =   2145
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "委任書"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   156
         Tag             =   ".POA.pdf"
         Top             =   1620
         Width           =   2145
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "其他共有人之同意書"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   155
         Tag             =   "."
         Top             =   990
         Width           =   2145
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "基本資料表"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   153
         Tag             =   ".CONTACT.pdf"
         Top             =   375
         Width           =   1425
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "文件檔名"
         Enabled         =   0   'False
         Height          =   195
         Index           =   2
         Left            =   990
         TabIndex        =   152
         Tag             =   ".ATT.pdf"
         Top             =   2040
         Width           =   1200
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "文件描述"
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   990
         TabIndex        =   151
         Tag             =   ".簽章切結書"
         Top             =   1830
         Width           =   1200
      End
      Begin VB.CheckBox chkDoc 
         Caption         =   "其他"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   150
         Top             =   1830
         Width           =   660
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "讓與契約書或讓與證明文件"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   149
         Tag             =   ".assignment.pdf"
         Top             =   570
         Width           =   2490
      End
      Begin VB.CheckBox chkAtt 
         Caption         =   "公司併購證明文件"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   148
         Tag             =   "."
         Top             =   780
         Width           =   2145
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "附送書件"
         Height          =   180
         Left            =   90
         TabIndex        =   154
         Top             =   150
         Width           =   720
      End
   End
   Begin VB.TextBox txtCP84 
      Height          =   270
      Left            =   4740
      MaxLength       =   7
      TabIndex        =   3
      Top             =   5250
      Width           =   1095
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Index           =   3
      Left            =   1770
      MaxLength       =   2
      TabIndex        =   6
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Index           =   2
      Left            =   930
      MaxLength       =   2
      TabIndex        =   5
      Top             =   5280
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Index           =   1
      Left            =   2250
      MaxLength       =   2
      TabIndex        =   2
      Top             =   4950
      Width           =   495
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Index           =   0
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   1
      Top             =   4950
      Width           =   495
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   4740
      MaxLength       =   7
      TabIndex        =   0
      Top             =   4920
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2700
      MaxLength       =   2
      TabIndex        =   15
      Top             =   240
      Width           =   375
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2460
      MaxLength       =   1
      TabIndex        =   14
      Top             =   240
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1620
      MaxLength       =   6
      TabIndex        =   13
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
   Begin VB.TextBox Text14 
      Height          =   270
      Left            =   2010
      MaxLength       =   1
      TabIndex        =   7
      Top             =   5610
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Index           =   1
      Left            =   7935
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   7110
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "變更事項(&R)"
      Height          =   405
      Index           =   3
      Left            =   5880
      TabIndex        =   9
      Top             =   70
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   2835
      Left            =   120
      TabIndex        =   44
      Top             =   2040
      Width           =   8985
      _ExtentX        =   15854
      _ExtentY        =   4995
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "受讓人1"
      TabPicture(0)   =   "frm04010302_1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label14(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(4)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5(5)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(6)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5(7)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5(8)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Combo2(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Combo3(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Combo3(1)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtCaseField(39)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtCaseField(40)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtCaseField(41)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtCaseField(42)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtCaseField(43)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtCaseField(44)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Text6(0)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "受讓人2"
      TabPicture(1)   =   "frm04010302_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text6(1)"
      Tab(1).Control(1)=   "txtCaseField(50)"
      Tab(1).Control(2)=   "txtCaseField(49)"
      Tab(1).Control(3)=   "txtCaseField(48)"
      Tab(1).Control(4)=   "txtCaseField(47)"
      Tab(1).Control(5)=   "txtCaseField(46)"
      Tab(1).Control(6)=   "txtCaseField(45)"
      Tab(1).Control(7)=   "Combo3(3)"
      Tab(1).Control(8)=   "Combo3(2)"
      Tab(1).Control(9)=   "Combo2(1)"
      Tab(1).Control(10)=   "Label5(24)"
      Tab(1).Control(11)=   "Label5(25)"
      Tab(1).Control(12)=   "Label5(26)"
      Tab(1).Control(13)=   "Label5(27)"
      Tab(1).Control(14)=   "Label5(28)"
      Tab(1).Control(15)=   "Label5(29)"
      Tab(1).Control(16)=   "Label14(2)"
      Tab(1).Control(17)=   "Label18(1)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "受讓人3"
      TabPicture(2)   =   "frm04010302_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text6(2)"
      Tab(2).Control(1)=   "txtCaseField(56)"
      Tab(2).Control(2)=   "txtCaseField(55)"
      Tab(2).Control(3)=   "txtCaseField(54)"
      Tab(2).Control(4)=   "txtCaseField(53)"
      Tab(2).Control(5)=   "txtCaseField(52)"
      Tab(2).Control(6)=   "txtCaseField(51)"
      Tab(2).Control(7)=   "Combo3(5)"
      Tab(2).Control(8)=   "Combo3(4)"
      Tab(2).Control(9)=   "Combo2(2)"
      Tab(2).Control(10)=   "Label5(18)"
      Tab(2).Control(11)=   "Label5(17)"
      Tab(2).Control(12)=   "Label5(16)"
      Tab(2).Control(13)=   "Label14(6)"
      Tab(2).Control(14)=   "Label5(33)"
      Tab(2).Control(15)=   "Label5(34)"
      Tab(2).Control(16)=   "Label5(35)"
      Tab(2).Control(17)=   "Label14(3)"
      Tab(2).ControlCount=   18
      TabCaption(3)   =   "受讓人4"
      TabPicture(3)   =   "frm04010302_1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Text6(3)"
      Tab(3).Control(1)=   "txtCaseField(62)"
      Tab(3).Control(2)=   "txtCaseField(61)"
      Tab(3).Control(3)=   "txtCaseField(60)"
      Tab(3).Control(4)=   "txtCaseField(59)"
      Tab(3).Control(5)=   "txtCaseField(58)"
      Tab(3).Control(6)=   "txtCaseField(57)"
      Tab(3).Control(7)=   "Combo3(7)"
      Tab(3).Control(8)=   "Combo3(6)"
      Tab(3).Control(9)=   "Combo2(3)"
      Tab(3).Control(10)=   "Label5(21)"
      Tab(3).Control(11)=   "Label5(20)"
      Tab(3).Control(12)=   "Label5(19)"
      Tab(3).Control(13)=   "Label18(4)"
      Tab(3).Control(14)=   "Label5(12)"
      Tab(3).Control(15)=   "Label5(11)"
      Tab(3).Control(16)=   "Label5(10)"
      Tab(3).Control(17)=   "Label14(5)"
      Tab(3).ControlCount=   18
      TabCaption(4)   =   "受讓人5"
      TabPicture(4)   =   "frm04010302_1.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Text6(4)"
      Tab(4).Control(1)=   "txtCaseField(68)"
      Tab(4).Control(2)=   "txtCaseField(67)"
      Tab(4).Control(3)=   "txtCaseField(66)"
      Tab(4).Control(4)=   "txtCaseField(65)"
      Tab(4).Control(5)=   "txtCaseField(64)"
      Tab(4).Control(6)=   "txtCaseField(63)"
      Tab(4).Control(7)=   "Combo3(9)"
      Tab(4).Control(8)=   "Combo3(8)"
      Tab(4).Control(9)=   "Combo2(4)"
      Tab(4).Control(10)=   "Label5(15)"
      Tab(4).Control(11)=   "Label5(14)"
      Tab(4).Control(12)=   "Label5(13)"
      Tab(4).Control(13)=   "Label18(3)"
      Tab(4).Control(14)=   "Label5(9)"
      Tab(4).Control(15)=   "Label5(2)"
      Tab(4).Control(16)=   "Label5(1)"
      Tab(4).Control(17)=   "Label14(4)"
      Tab(4).ControlCount=   18
      Begin VB.TextBox Text6 
         Height          =   300
         Index           =   4
         Left            =   -74580
         MaxLength       =   9
         TabIndex        =   139
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Index           =   3
         Left            =   -74580
         MaxLength       =   9
         TabIndex        =   137
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Index           =   2
         Left            =   -74580
         MaxLength       =   9
         TabIndex        =   135
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Index           =   1
         Left            =   -74580
         MaxLength       =   9
         TabIndex        =   133
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Index           =   0
         Left            =   420
         MaxLength       =   9
         TabIndex        =   131
         Top             =   390
         Width           =   1095
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   68
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   77
         Top             =   2445
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   67
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   76
         Top             =   2220
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   66
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   75
         Top             =   1995
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   65
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   74
         Top             =   1455
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   64
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   73
         Top             =   1230
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   63
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   72
         Top             =   1005
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   62
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   71
         Top             =   2475
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   61
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   70
         Top             =   2235
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   60
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   69
         Top             =   2010
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   59
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   68
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   58
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   67
         Top             =   1215
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   57
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   66
         Top             =   975
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   56
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   65
         Top             =   2505
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   55
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   64
         Top             =   2280
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   54
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   63
         Top             =   2055
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   53
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   61
         Top             =   1470
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   52
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   60
         Top             =   1230
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   51
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   59
         Top             =   1005
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   50
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   57
         Top             =   2490
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   49
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   56
         Top             =   2265
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   48
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   55
         Top             =   2040
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   47
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   54
         Top             =   1470
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   46
         Left            =   -73485
         MaxLength       =   60
         TabIndex        =   53
         Top             =   1230
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   45
         Left            =   -73485
         MaxLength       =   40
         TabIndex        =   52
         Top             =   1005
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   44
         Left            =   1515
         MaxLength       =   40
         TabIndex        =   50
         Top             =   2490
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   43
         Left            =   1515
         MaxLength       =   60
         TabIndex        =   49
         Top             =   2265
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   42
         Left            =   1515
         MaxLength       =   40
         TabIndex        =   48
         Top             =   2025
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   41
         Left            =   1515
         MaxLength       =   40
         TabIndex        =   47
         Top             =   1485
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   40
         Left            =   1515
         MaxLength       =   60
         TabIndex        =   46
         Top             =   1260
         Width           =   6135
      End
      Begin VB.TextBox txtCaseField 
         BackColor       =   &H80000004&
         Height          =   270
         Index           =   39
         Left            =   1515
         MaxLength       =   40
         TabIndex        =   45
         Top             =   1035
         Width           =   6135
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   5
         Left            =   -73485
         TabIndex        =   146
         Top             =   1755
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   4
         Left            =   -73485
         TabIndex        =   145
         Top             =   720
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   3
         Left            =   -73485
         TabIndex        =   144
         Top             =   1740
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   2
         Left            =   -73485
         TabIndex        =   143
         Top             =   720
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   1
         Left            =   1515
         TabIndex        =   142
         Top             =   1740
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   0
         Left            =   1515
         TabIndex        =   141
         Top             =   750
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   7
         Left            =   -73485
         TabIndex        =   140
         Top             =   1710
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   9
         Left            =   -73485
         TabIndex        =   62
         Top             =   1710
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   8
         Left            =   -73485
         TabIndex        =   58
         Top             =   720
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo3 
         Height          =   300
         Index           =   6
         Left            =   -73485
         TabIndex        =   51
         Top             =   720
         Width           =   6135
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "10821;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   4
         Left            =   -73470
         TabIndex        =   138
         TabStop         =   0   'False
         Top             =   390
         Width           =   6840
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "12065;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   3
         Left            =   -73470
         TabIndex        =   136
         TabStop         =   0   'False
         Top             =   390
         Width           =   6840
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "12065;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   2
         Left            =   -73470
         TabIndex        =   134
         TabStop         =   0   'False
         Top             =   390
         Width           =   6840
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "12065;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   1
         Left            =   -73470
         TabIndex        =   132
         TabStop         =   0   'False
         Top             =   390
         Width           =   6840
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "12065;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   300
         Index           =   0
         Left            =   1530
         TabIndex        =   130
         TabStop         =   0   'False
         Top             =   390
         Width           =   6840
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "12065;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   15
         Left            =   -73980
         TabIndex        =   129
         Top             =   1500
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   14
         Left            =   -73980
         TabIndex        =   128
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   13
         Left            =   -73980
         TabIndex        =   127
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人9"
         Height          =   180
         Index           =   3
         Left            =   -74340
         TabIndex        =   126
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   9
         Left            =   -73980
         TabIndex        =   125
         Top             =   2460
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   2
         Left            =   -73980
         TabIndex        =   124
         Top             =   2220
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   1
         Left            =   -73980
         TabIndex        =   123
         Top             =   1980
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人10"
         Height          =   180
         Index           =   4
         Left            =   -74340
         TabIndex        =   122
         Top             =   1740
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   21
         Left            =   -73980
         TabIndex        =   121
         Top             =   1470
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   20
         Left            =   -73980
         TabIndex        =   120
         Top             =   1230
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   19
         Left            =   -73980
         TabIndex        =   119
         Top             =   990
         Width           =   345
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人7"
         Height          =   180
         Index           =   4
         Left            =   -74340
         TabIndex        =   118
         Top             =   750
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   12
         Left            =   -73980
         TabIndex        =   117
         Top             =   2430
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   11
         Left            =   -73980
         TabIndex        =   116
         Top             =   2190
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   10
         Left            =   -73950
         TabIndex        =   115
         Top             =   1950
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人8"
         Height          =   180
         Index           =   5
         Left            =   -74340
         TabIndex        =   114
         Top             =   1710
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   18
         Left            =   -73980
         TabIndex        =   113
         Top             =   2490
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   17
         Left            =   -73980
         TabIndex        =   112
         Top             =   2250
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   16
         Left            =   -73980
         TabIndex        =   111
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人6"
         Height          =   180
         Index           =   6
         Left            =   -74340
         TabIndex        =   110
         Top             =   1800
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   33
         Left            =   -73980
         TabIndex        =   109
         Top             =   1500
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   34
         Left            =   -73980
         TabIndex        =   108
         Top             =   1260
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   35
         Left            =   -73980
         TabIndex        =   107
         Top             =   1020
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人5"
         Height          =   180
         Index           =   3
         Left            =   -74340
         TabIndex        =   106
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   24
         Left            =   -74010
         TabIndex        =   105
         Top             =   2535
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   25
         Left            =   -74010
         TabIndex        =   104
         Top             =   2295
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   26
         Left            =   -74010
         TabIndex        =   103
         Top             =   2055
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   27
         Left            =   -74010
         TabIndex        =   102
         Top             =   1515
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   28
         Left            =   -74010
         TabIndex        =   101
         Top             =   1275
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   29
         Left            =   -74010
         TabIndex        =   100
         Top             =   1035
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人3"
         Height          =   180
         Index           =   2
         Left            =   -74370
         TabIndex        =   99
         Top             =   795
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人4"
         Height          =   180
         Index           =   1
         Left            =   -74370
         TabIndex        =   98
         Top             =   1785
         Width           =   630
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   8
         Left            =   1020
         TabIndex        =   97
         Top             =   2490
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   7
         Left            =   1020
         TabIndex        =   96
         Top             =   2250
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   6
         Left            =   1020
         TabIndex        =   95
         Top             =   2010
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(日):"
         Height          =   180
         Index           =   5
         Left            =   1020
         TabIndex        =   94
         Top             =   1530
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(英):"
         Height          =   180
         Index           =   4
         Left            =   1020
         TabIndex        =   93
         Top             =   1290
         Width           =   345
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "(中):"
         Height          =   180
         Index           =   3
         Left            =   1020
         TabIndex        =   92
         Top             =   1050
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "代表人1"
         Height          =   180
         Index           =   1
         Left            =   660
         TabIndex        =   91
         Top             =   840
         Width           =   630
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "代表人2"
         Height          =   180
         Index           =   2
         Left            =   660
         TabIndex        =   90
         Top             =   1770
         Width           =   630
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   89
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Index           =   5
         Left            =   -73170
         TabIndex        =   88
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   87
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   86
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   85
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73140
         TabIndex        =   84
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   83
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   82
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   81
         Top             =   360
         Width           =   225
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "日:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   80
         Top             =   840
         Width           =   225
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "英:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   79
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "中:"
         Height          =   180
         Left            =   -73170
         TabIndex        =   78
         Top             =   360
         Width           =   225
      End
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "證書形式:   　      (1:電子 2:紙本)"
      Height          =   180
      Left            =   150
      TabIndex        =   159
      Top             =   5940
      Width           =   2505
   End
   Begin MSForms.ListBox lstNameAgent 
      Height          =   810
      Left            =   7440
      TabIndex        =   4
      Top             =   1080
      Width           =   1500
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2646;1429"
      MatchEntry      =   0
      ListStyle       =   1
      MultiSelect     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   1140
      TabIndex        =   16
      Top             =   750
      Width           =   8010
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "14129;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "繳費金額："
      Height          =   180
      Left            =   3795
      TabIndex        =   43
      Top             =   5295
      Width           =   900
   End
   Begin VB.Label lblNameAgent 
      AutoSize        =   -1  'True
      Caption         =   "出名代理人"
      Height          =   180
      Left            =   6495
      TabIndex        =   42
      Top             =   1110
      Width           =   900
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   7
      Left            =   4020
      TabIndex        =   41
      Top             =   1590
      Width           =   3360
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "5927;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   6
      Left            =   1140
      TabIndex        =   40
      Top             =   1590
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "3281;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   5
      Left            =   4020
      TabIndex        =   39
      Top             =   1350
      Width           =   2490
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "4392;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   4
      Left            =   1140
      TabIndex        =   38
      Top             =   1350
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "3281;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   3
      Left            =   1140
      TabIndex        =   37
      Top             =   1110
      Width           =   1860
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "3281;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   1
      Left            =   4020
      TabIndex        =   36
      Top             =   540
      Width           =   3090
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "5450;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label4 
      Height          =   180
      Index           =   0
      Left            =   1140
      TabIndex        =   35
      Top             =   540
      Width           =   1950
      VariousPropertyBits=   27
      Caption         =   "Label4"
      Size            =   "3440;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   180
      TabIndex        =   34
      Top             =   780
      Width           =   765
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "專利號數:"
      Height          =   180
      Index           =   0
      Left            =   3180
      TabIndex        =   33
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   32
      Top             =   540
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "申請書日期:"
      Height          =   180
      Left            =   3750
      TabIndex        =   31
      Top             =   4965
      Width           =   945
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   180
      TabIndex        =   30
      Top             =   270
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "承辦人:"
      Height          =   180
      Left            =   180
      TabIndex        =   29
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "智權人員:"
      Height          =   180
      Left            =   3180
      TabIndex        =   28
      Top             =   1350
      Width           =   765
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   180
      TabIndex        =   27
      Top             =   1590
      Width           =   945
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "機關文號:"
      Height          =   180
      Left            =   3180
      TabIndex        =   26
      Top             =   1590
      Width           =   765
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "領證及繳納第"
      Height          =   180
      Left            =   150
      TabIndex        =   25
      Top             =   4950
      Width           =   1080
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "至"
      Height          =   180
      Left            =   2010
      TabIndex        =   24
      Top             =   4950
      Width           =   180
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "年年費"
      Height          =   180
      Left            =   2850
      TabIndex        =   23
      Top             =   4950
      Width           =   540
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "繳納第"
      Height          =   180
      Left            =   150
      TabIndex        =   22
      Top             =   5280
      Width           =   540
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "至"
      Height          =   180
      Left            =   1530
      TabIndex        =   21
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "年年費"
      Height          =   180
      Left            =   2370
      TabIndex        =   20
      Top             =   5280
      Width           =   540
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "是否修改申請書內容"
      Height          =   180
      Left            =   150
      TabIndex        =   19
      Top             =   5610
      Width           =   1620
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "(Y:WORD)"
      Height          =   180
      Left            =   2610
      TabIndex        =   18
      Top             =   5610
      Width           =   810
   End
   Begin VB.Label Label29 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   180
      TabIndex        =   17
      Top             =   1110
      Width           =   765
   End
End
Attribute VB_Name = "frm04010302_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/10 改成Form2.0 (Combo1,Combo2,lstNameAgent,Label4...
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
'整理 by Morgan 2005/7/29
Option Explicit

Dim strReceiveNo As String
'Modify by Morgan 2005/7/29 改用動態陣列
'Dim pa(1 To T_PA) As String
Dim pa() As String
Dim cp() As String 'Add By Sindy 2019/1/17
Dim m_CP110 As String, m_CP110_2 As String
Dim m_CP22 As String

Dim intWhere As Integer
' 90.07.05 modify by louis 儲存應繳年費的資料
Dim m_CaseFee(1 To 2) As String
Dim m_OldCaseFee As String
'Add By Cheng 2003/03/07
Dim m_CP55 As String '原讓與人
'Add by Morgan 2010/5/24
Dim m_CP(93 To 96) As String
Dim m_CP10 As String '案件性質 Add by Amy 2014/08/14
Dim m_CaseNo As String
Public m_CP118isY As Boolean '是否為電子送件申請書:True.是


Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
 Dim strTxt(1 To 10) As String, strTmp(1 To 2) As String, i As Integer
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   If Text10(0) <> "" And Text10(1) <> "" Then
      strTmp(1) = Text10(0).Text
      strTmp(2) = Text10(1).Text
   Else
      strTmp(1) = Text10(2).Text
      strTmp(2) = Text10(3).Text
   End If
   For i = 1 To 2
      Select Case strTmp(i)
         Case "1"
            strTmp(i) = "一"
         Case "2"
            strTmp(i) = "二"
         Case "3"
            strTmp(i) = "三"
         Case "4"
            strTmp(i) = "四"
         Case "5"
            strTmp(i) = "五"
         Case "6"
            strTmp(i) = "六"
         Case "7"
            strTmp(i) = "七"
         Case "8"
            strTmp(i) = "八"
         Case "9"
            strTmp(i) = "九"
         Case "10"
            strTmp(i) = "十"
         Case "11"
            strTmp(i) = "十一"
         Case "12"
            strTmp(i) = "十二"
         Case "13"
            strTmp(i) = "十三"
         Case "14"
            strTmp(i) = "十四"
         Case "15"
            strTmp(i) = "十五"
         Case "16"
            strTmp(i) = "十六"
         Case "17"
            strTmp(i) = "十七"
         Case "18"
            strTmp(i) = "十八"
         Case "19"
            strTmp(i) = "十九"
         Case "20"
            strTmp(i) = "二十"
      End Select
   Next
   If strTmp(1) = strTmp(2) Then
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','第" & strTmp(1) & "年年費')"
   Else
      strTxt(1) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
         "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','第幾年至幾年費','第" & strTmp(1) & "年至第" & strTmp(2) & "年年費')"
   End If
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(1, strTxt) Then
   If Not ClsLawExecSQL(1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add By Cheng 2003/01/19
'附件的例外欄位
Private Sub StartLetter1(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 10) As String, strTmp(1 To 2) As String, i As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer
    ii = 0
    EndLetter ET01, strReceiveNo, ET03, strUserNum
    StrSQLa = "Select CU07 From CaseProgress, Customer Where SUBSTR(CP55,1,8)=CU01 AND  SUBSTR(CP55,9,1)=CU02 AND CP09 ='" & strReceiveNo & "' And CU07 IS NOT NULL And CU15 <>'0' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與代表人','代表人：" & rsA.Fields(0).Value & "')"
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    StrSQLa = "Select CU07 From CaseProgress, Customer Where SUBSTR(CP56,1,8)=CU01 AND SUBSTR(CP56,9,1)=CU02 AND CP09 ='" & strReceiveNo & "' And CU07 IS NOT NULL And CU15 <>'0' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        ii = ii + 1
        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
             "('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓代表人','代表人：" & rsA.Fields(0).Value & "')"
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    If ii > 0 Then
        'edit by nickc 2007/02/05 不用 dll 了
        'If Not objLawDll.ExecSQL(ii, strTxt) Then
        If Not ClsLawExecSQL(ii, strTxt) Then
           MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
        End If
    End If
End Sub

Private Sub chkDoc_Click(Index As Integer)
   If Index = 0 Then
      If chkDoc(0).Value = 0 Then
         chkDoc(1).Enabled = False
         chkDoc(1).Value = 0
         chkDoc(2).Enabled = False
         chkDoc(2).Value = 0
      Else
         chkDoc(1).Enabled = True
         chkDoc(2).Enabled = True
      End If
   End If
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim bolChk As Boolean, strTmp As String
Dim strFolder As String
Dim strFileName As String
   
   Select Case Index
      Case 0
         If Text5 = "" Then
            MsgBox "申請書日期不可空白 !", vbCritical
            Text5.SetFocus
            Exit Sub
         End If
         If Text6(0) = "" Then
            MsgBox "受讓人不可空白 !", vbCritical
            Text6(0).SetFocus
            Text6_GotFocus 0
            Exit Sub
         End If
         If Text10(0) <> "" And Text10(1) <> "" And Text10(2) <> "" And Text10(3) <> "" Then
            MsgBox "領證及繳納年費與繳納年費不可同時輸入 !", vbCritical
            Text10(0).SetFocus
            Exit Sub
         End If
         'Add By Sindy 2022/12/28
         If TextPA178.Visible = True Then
            If TextPA178 = "" Then
               MsgBox "證書形式不可空白！", vbExclamation + vbOKOnly
               Me.TextPA178.SetFocus
               Exit Sub
            End If
         End If
         '2022/12/28 END
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         If TxtValidate = False Then Exit Sub
         
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         'Add By Sindy 2019/1/4 + 產生申請書
         If m_CP118isY = True Then '電子送件
            m_CaseNo = PUB_FCPCaseNo2FileName(pa(1), pa(2), pa(3), pa(4))
            strFolder = PUB_Getdesktop
            strFolder = strFolder & "\" & m_CaseNo
            If Dir(strFolder, vbDirectory) = "" Then
               MkDir strFolder
            End If
            
            '2.申請書
            If Trim(pa(22)) = "" Then '判斷專利號數
               If StartLetter2("01", "04") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "04", False, strUserNum, , , True, strExc(9)
               'strFileName = strFolder & "\" & "申請權讓與登記申請書"
               'Modify By Sindy 2020/1/6
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               'Call PUB_MakeDoc(strExc(9), strFileName)
            Else
               If StartLetter2("01", "05") = False Then Exit Sub
               NowPrint strReceiveNo, "01", "05", False, strUserNum, , , True, strExc(9)
               'strFileName = strFolder & "\" & "專利權讓與登記申請書"
               'Modify By Sindy 2020/1/6
               strFileName = strFolder & "\" & m_CaseNo & ".data"
               'Call PUB_MakeDoc(strExc(9), strFileName)
            End If
            '1.基本資料
            If StartLetter3("01", "07") = False Then Exit Sub
            NowPrint strReceiveNo, "01", "07", False, strUserNum, , , True, strExc(10)
            'strFileName = strFolder & "\" & m_CaseNo & ".Contact"
            'Call PUB_MakeDoc(strExc(10), strFileName)
            Call PUB_MakeDoc(strExc(9) & Chr(12) & strExc(10), strFileName, False)
            
         '2019/1/4 END
         '紙本申請書
         Else
            If Text14 = "Y" Then
               bolChk = True
            Else
               bolChk = False
            End If
            
            If Text6(0) <> "" And Text10(0) <> "" Then
               '讓與 + 領證    01
               strTmp = "01"
            ElseIf Text6(0) <> "" And Text10(2) <> "" Then
               '讓與 + 繳年費  02
               strTmp = "02"
            ElseIf Text6(0) <> "" Then
               '一般 0
               strTmp = "00"
            End If
            StartLetter "01", strTmp
            strLetterDate = Text5.Text
            'Mark by Amy 2014/08/14 沒使用-玲玲
            'NowPrint strReceiveNo, "01", strTmp, bolChk, strUserNum, 0, , , , , , , , , , , , strReceiveNo
            'Add By Cheng 2003/01/19
            '申請書附件
            StartLetter1 "01", "03"
            'Modify by Amy 2014/08/14 +傳strLetterRecNo,修改改frm1105_1開 (系統先上傳此版，User會再上傳給智慧局的版本-玲玲)
            NowPrint strReceiveNo, "01", "03", True, strUserNum, 0, , , , , , , , , , , , strReceiveNo
            
            'Modify By Sindy 2019/1/4 沒有要上傳卷宗區,Mark
   '         If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
   '         If bolChk = True Then
   '             frm1105_1.m_RecNo = strReceiveNo
   '             frm1105_1.m_PdfName = Text1 & Text2 & IIf(Text3 & Text4 = "000", "", "-" & Text3 & "-" & Text4) & "." & m_CP10 & ".DATA.PDF"
   '             frm1105_1.Show
   '         End If
   '         End If
            'end 2014/08/14
         End If
         
         frm040103_1.Show
         ' 90.08.27 modify by louis
         frm040103_1.ClearForm
         Unload Me
      Case 1
         frm040103_1.Show
         Unload Me
      Case 3
         Set frm06010303_1.oParent = Me 'Add by Morgan 2011/10/5
         frm06010303_1.LoadMe strReceiveNo, pa(1), pa(2), pa(3), pa(4), 42
         Me.Hide
   End Select
End Sub

'基本資料表
Private Function StartLetter3(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(110) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim strInventor As String
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   '受讓人--畫面上
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), Text6(0), Text6(1), Text6(2), Text6(3), Text6(4), , , , cp(10), txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67), "E", ET01, ET03, strReceiveNo)
   '讓與人--CP讓與人或發文前申請人
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , , , , cp(10), , , , , , , , , , , , , , , , , , , , , "E", ET01, ET03, strReceiveNo)
   
   '受讓人之代理人(出名代理人)
'   strExc(0) = "select oa08,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-證書字號','" & .Fields("oa08") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-ID','" & .Fields("ST26") & "')"
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/10 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 2, cp(110), ET01, strReceiveNo, ET03, ii, strTxt(), "受讓人")
   '讓與人之代理人
   Call GetCP110_2
   If m_CP110_2 <> "" Then
'      strExc(0) = "select oa08,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110_2 & "',oa02)>0 and st01(+)=oa02 order by OA03"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         With RsTemp
'         jj = 1
'         Do While Not .EOF
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-證書字號','" & .Fields("oa08") & "')"
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-ID','" & .Fields("ST26") & "')"
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'            jj = jj + 1
'            .MoveNext
'         Loop
'         End With
'      End If
      'Modify By Sindy 2020/4/10 申請書:出名代理人
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 2, m_CP110_2, ET01, strReceiveNo, ET03, ii, strTxt(), "讓與人")
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter3 = True
   End If
End Function

'申請書
Private Function StartLetter2(ByVal ET01 As String, ByVal ET03 As String) As Boolean
Dim strTxt(200) As String, strTmp As String
Dim ii As Integer, jj As Integer
Dim chk As CheckBox, strTmp1 As String
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','本所案號','" & m_CaseNo & "')"
   
   'Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), IIf(chkAtt(26).Value = 1, False, True))
   '受讓人--畫面上
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), Text6(0), Text6(1), Text6(2), Text6(3), Text6(4), , , , cp(10), txtCaseField(39), txtCaseField(40), txtCaseField(42), txtCaseField(43), txtCaseField(45), txtCaseField(46), txtCaseField(48), txtCaseField(49), txtCaseField(51), txtCaseField(52), txtCaseField(54), txtCaseField(55), txtCaseField(57), txtCaseField(58), txtCaseField(60), txtCaseField(61), txtCaseField(63), txtCaseField(64), txtCaseField(66), txtCaseField(67), "E", ET01, ET03, strReceiveNo)
   '讓與人--CP讓與人或發文前申請人
   Call PUB_GetApplData(pa(), pa(1), pa(2), pa(3), pa(4), , , , , , , , , cp(10), , , , , , , , , , , , , , , , , , , , , "E", ET01, ET03, strReceiveNo)
   
   '受讓人之代理人(出名代理人)
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','受讓人之代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/10 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, strReceiveNo, ET03, ii, strTxt(), "受讓人之")
   '讓與人之代理人
   Call GetCP110_2
   If m_CP110_2 <> "" Then
'      strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & m_CP110_2 & "',oa02)>0 and st01(+)=oa02 order by OA03"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         With RsTemp
'         jj = 1
'         Do While Not .EOF
'            ii = ii + 1
'            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','讓與人之代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'            jj = jj + 1
'            .MoveNext
'         Loop
'         End With
'      End If
      'Modify By Sindy 2020/4/10 申請書:出名代理人
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, m_CP110_2, ET01, strReceiveNo, ET03, ii, strTxt(), "讓與人之")
   End If
   
   'Add By Sindy 2019/5/14
   If ET03 = "05" Then '專利權讓與
      strTmp = ""
      If Text6(0) <> "" Then
         strTmp = strTmp & "、「" & GetPrjPeople1(ChangeCustomerL(pa(26))) & "」讓與「" & Mid(Combo2(0).Text, 3) & "」"
      End If
      If Text6(1) <> "" Then
         strTmp = strTmp & "、「" & GetPrjPeople1(ChangeCustomerL(pa(27))) & "」讓與「" & Mid(Combo2(1).Text, 3) & "」"
      End If
      If Text6(2) <> "" Then
         strTmp = strTmp & "、「" & GetPrjPeople1(ChangeCustomerL(pa(28))) & "」讓與「" & Mid(Combo2(2).Text, 3) & "」"
      End If
      If Text6(3) <> "" Then
         strTmp = strTmp & "、「" & GetPrjPeople1(ChangeCustomerL(pa(29))) & "」讓與「" & Mid(Combo2(3).Text, 3) & "」"
      End If
      If Text6(4) <> "" Then
         strTmp = strTmp & "、「" & GetPrjPeople1(ChangeCustomerL(pa(30))) & "」讓與「" & Mid(Combo2(4).Text, 3) & "」"
      End If
      If strTmp = "" Then
         strTmp = "由「」讓與「」。"
      Else
         strTmp = Mid(strTmp, 2)
         strTmp = "由" & strTmp & "。"
      End If
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請內容','" & strTmp & "')"
   End If
   
   'Add By Sindy 2022/12/27 + 申請證書形式
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請證書形式','" & IIf(TextPA178 = "1", "電子證書", IIf(TextPA178 = "2", "紙本證書", "電子證書/紙本證書")) & "')"
   '2022/12/27 END
   
   '*******************************************************************************
   '附送書件
   '*******************************************************************************
   strTmp = ""
   For Each chk In chkAtt
      If chk.Value = 1 Then
         strTmp1 = ""
         strTmp1 = chk.Caption
         If strTmp = "" Then
            strTmp1 = "　【" & strTmp1 & "】"
            If Len(strTmp1) < 16 Then
               strTmp1 = strTmp1 & String(16 - Len(strTmp1), "　")
            End If
         Else
            strTmp1 = "　　【" & strTmp1 & "】"
            If Len(strTmp1) < 17 Then
               strTmp1 = strTmp1 & String(17 - Len(strTmp1), "　")
            End If
         End If
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & m_CaseNo & Trim(chk.Tag)
      End If
   Next
   '其他
   If chkDoc(1).Value = 1 Or chkDoc(2).Value = 1 Then
      strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & "　　【" & chkDoc(0).Caption & "】"
      If chkDoc(1).Value = 1 Then
         strTmp1 = "　　　【" & chkDoc(1).Caption & "】"
         If Len(strTmp1) < 17 Then
            strTmp1 = strTmp1 & String(17 - Len(strTmp1), "　")
         End If
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & m_CaseNo & Trim(chkDoc(1).Tag)
      End If
      If chkDoc(2).Value = 1 Then
         strTmp1 = "　　　【" & chkDoc(2).Caption & "】"
         If Len(strTmp1) < 17 Then
            strTmp1 = strTmp1 & String(17 - Len(strTmp1), "　")
         End If
         strTmp = IIf(strTmp <> "", strTmp & vbCrLf, "") & strTmp1 & m_CaseNo & Trim(chkDoc(2).Tag)
      End If
   End If
   If strTmp <> "" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','附送書件','　" & strTmp & "')"
   End If
   '*******************************************************************************
   
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','繳費金額','" & Val(txtCP84) & "')"
   'Modify By Sindy 2020/4/9 有繳費金額就要帶出收據抬頭
   If Val(txtCP84) > 0 Then
      Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 3, , ET01, strReceiveNo, ET03, ii, strTxt())
   End If
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetter2 = True
   End If
End Function

'Add By Sindy 2019/1/17
'讓與人之代理人
Private Function GetCP110_2() As String
   m_CP110_2 = ""
   GetCP110_2 = ""
   '最近一筆A,B類收文已發文,有主管機關者
   strExc(0) = "select cp09,cp110" & _
               " from caseprogress" & _
               " where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
               " and cp110 is not null and cp130 is not null and cp27 is not null" & _
               " and cp57 is null" & _
               " order by cp27 desc,cp09 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      m_CP110_2 = "" & RsTemp.Fields("cp110")
      GetCP110_2 = PUB_GetAgentCP110(RsTemp.Fields("cp09"))
   Else
      m_CP110_2 = ""
      GetCP110_2 = PUB_GetAgentCP110("")
   End If
End Function

Private Function FormSave() As Boolean
Dim stUpdate As String
   
On Error GoTo ErrorHandler
   
   cnnConnection.BeginTrans
   
   stUpdate = ""
   '若有輸入讓與申請人
   If Me.Text6(0).Text <> "" Then
      stUpdate = stUpdate & ",CP56=" & CNULL(ChangeCustomerL(Text6(0).Text))
      
      '若原讓與人與原申請人不同
      If ChangeCustomerL(m_CP55) <> ChangeCustomerL(pa(26)) Then
         stUpdate = stUpdate & ",CP55=" & CNULL(ChangeCustomerL(pa(26)))
      End If
   End If
   
   'Add by Morgan 2010/5/24
   For intI = 0 To 3
      If Me.Text6(intI + 1).Text <> "" Then
         stUpdate = stUpdate & ",CP" & (89 + intI) & "=" & CNULL(ChangeCustomerL(Text6(intI + 1).Text))
         '若原讓與人與原申請人不同
         If ChangeCustomerL(m_CP(93 + intI)) <> ChangeCustomerL(pa(27 + intI)) Then
            stUpdate = stUpdate & ",CP" & (93 + intI) & "=" & CNULL(ChangeCustomerL(pa(27 + intI)))
         End If
      End If
   Next
   
   If lstNameAgent.Visible = True Then
      cp(110) = m_CP110 'Add By Sindy 2019/2/27
      stUpdate = stUpdate & ",CP22=" & CNULL(m_CP22) & ",cp110=" & CNULL(m_CP110)
   End If
   
   'Add By Sindy 2019/1/17
'   cp(84) = Val(txtCP84)
'   stUpdate = stUpdate & ",cp84=" & cp(84)
   If m_CP118isY = True Then
      cp(118) = "A"
   Else
      cp(118) = ""
   End If
   stUpdate = stUpdate & ",cp118=" & CNULL(cp(118))
   
   If stUpdate <> "" Then
      stUpdate = Mid(stUpdate, 2)
      strSql = " UPDATE CASEPROGRESS SET " & stUpdate & " WHERE CP09='" & strReceiveNo & "' and cp158=0 and cp159=0"
      cnnConnection.Execute strSql, intI
   End If
   
   'Modify By Sindy 2019/1/4 沒有要上傳卷宗區,Mark
'   'Add by Amy 2014/08/14 P台灣案電子化
'   If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then
'   If ExistCheck("AppForm", "AF01", strReceiveNo, "", False) = False Then
'        '新增申請書轉檔記錄
'        PUB_AddAppForm strReceiveNo
'   End If
'   End If
'   'end 2014/08/14
   'Add By Sindy 2022/12/28
   If TextPA178.Visible = True And TextPA178.Tag <> TextPA178.Text Then
      strSql = " UPDATE patent SET pa178='" & TextPA178 & "' WHERE PA01='" & pa(1) & "' and PA02='" & pa(2) & "' and PA03='" & pa(3) & "' and PA04='" & pa(4) & "'"
      cnnConnection.Execute strSql, intI
   End If
   '2022/12/28 END
   
   cnnConnection.CommitTrans
   FormSave = True
   
ErrorHandler:
   If Err.NUMBER <> 0 Then
      cnnConnection.RollbackTrans
   End If
   
End Function

'Add By Sindy 2019/1/17
Private Sub Combo3_Click(Index As Integer)
Dim i As Integer, strTmp As String
   
   If Combo3(Index) = "" Then
      For i = 0 To 2
         txtCaseField(i + (Index + 1) * 3 + 36) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo3(Index).Text, InStr(Combo3(Index).Text, "-") + 1, 1)
   strExc(1) = "CU" & 39 + (Val(strTmp) - 1) * 3 & ",CU" & 40 + (Val(strTmp) - 1) * 3 & ",CU" & 41 + (Val(strTmp) - 1) * 3
   strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & ChgCustomer(Left(Combo3(Index).Text, InStr(Combo3(Index).Text, "-") - 1))
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 0 To 2
         If Not IsNull(RsTemp.Fields(i)) Then
            txtCaseField(i + (Index + 1) * 3 + 36) = RsTemp.Fields(i)
         Else
            txtCaseField(i + (Index + 1) * 3 + 36) = ""
         End If
      Next
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   With frm040103_1
      Text1 = .Text1
      Text2 = .Text2
      Text3 = .Text3
      Text4 = .Text4
      strReceiveNo = .Tag
   End With
   'Add by Morgan 2005/7/29
   ReDim pa(TF_PA)
   ReDim cp(TF_CP)
   ReadPatent
   
   'Add by Morgan 2005/7/29
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   If pa(9) = "000" Then
      PUB_SetOurAgent lstNameAgent, pa(), m_CP110, , True 'Modified by Morgan 2021/12/10 +傳入bForm2=True
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2005/7/29 END
   Combo1.ListIndex = 0
   Text5.Text = strSrvDate(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010302_1 = Nothing
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim rsTemp1 As New ADODB.Recordset, Lbl As Object, i As Integer, j As Integer
 Dim strKey(0 To 5) As String
 Dim nIndex As Integer
 
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4
   For Each Lbl In Label4
      Lbl.Caption = ""
   Next
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label4(0) = pa(11)
      Label4(1) = pa(22)
      AddCboName Combo1, pa(5), pa(6), pa(7)
      
      j = 0
      For i = 26 To 30
         If pa(i) <> "" Then j = j + 1
      Next
      If pa(79) <> "" Then j = j + 1
      If pa(82) <> "" Then j = j + 1
      If j > 1 Then Text14 = "Y"
   End If
   
   'Add By Sindy 2019/1/17
   cp(9) = strReceiveNo
   If PUB_ReadCaseProgressDatabase(cp(), intWhere) Then
      txtCP84.Tag = cp(17)
      txtCP84.Text = txtCP84.Tag
   End If
      
   ' 原已繳費的年度
   m_OldCaseFee = pa(72)
   ' 設定本所案號
   For nIndex = 1 To 4
      strKey(nIndex) = pa(nIndex)
   Next
   ' 取得繳年費的資料
   If GetMoneyDate(pa(8), pa(9), strKey, m_CaseFee(1), m_CaseFee(2)) = True Then
   End If
   'Modify by Amy 2014/08/14 +CP10
   strExc(0) = "select cpm03,staff.st02 as st1,staff1.st02 as st2,CP43,CP56,CP55,CP110,CP89,CP90,CP91,CP92,CP93,CP94,CP95,CP96,CP10" & _
      " from caseprogress,casepropertymap,staff,staff staff1 where " & _
      "cp09='" & strReceiveNo & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp14=staff.st01(+) and cp13=staff1.st01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         m_CP110 = "" & .Fields("CP110")
         m_CP10 = .Fields("CP10") 'Add by Amy 2014/08/14
         If Not IsNull(.Fields(0)) Then Label4(3) = .Fields(0)
         If Not IsNull(.Fields(1)) Then Label4(4) = .Fields(1)
         If Not IsNull(.Fields(2)) Then Label4(5) = .Fields(2)
         If Not IsNull(.Fields(3)) Then
            strExc(0) = "SELECT CP05,CP08 FROM CASEPROGRESS WHERE CP09='" & .Fields(3) & "'"
            intI = 1
            Set rsTemp1 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Not IsNull(rsTemp1.Fields(0)) Then Label4(6) = TransDate(rsTemp1.Fields(0), 1)
               If Not IsNull(rsTemp1.Fields(1)) Then Label4(7) = rsTemp1.Fields(1)
            End If
         End If
         If Not IsNull(.Fields(4)) Then Text6(0).Text = ChangeCustomerS(.Fields(4)): Text6_Validate 0, False
         '記錄讓與人
         m_CP55 = "" & .Fields("CP55").Value
         'Add by Morgan 2010/5/24 +受讓人,讓與人 2~5
         For i = 0 To 3
            If Not IsNull(.Fields("CP" & (89 + i))) Then Text6(i + 1).Text = ChangeCustomerS(.Fields("CP" & (89 + i))): Text6_Validate i + 1, False
            m_CP(93 + i) = "" & .Fields("CP" & (93 + i))
         Next
         'END 2010/5/24
         'Add By Sindy 2019/1/17
         For i = 1 To 5 '讓與申請人代表
            Call SetCombo3(i)
         Next
         '2019/1/17 END
      End If
   End With
   
   'Add By Sindy 2022/12/28
   If cp(10) = 專利權讓與 And strSrvDate(1) >= "20230101" Then
      Label10.Visible = True
      TextPA178.Visible = True
      TextPA178.Text = pa(178)
      TextPA178.Tag = pa(178)
   Else
      Label10.Visible = False
      TextPA178.Visible = False
   End If
   '2022/12/28 END
End Sub

'Add By Sindy 2019/1/17
Private Sub SetCombo3(Index As Integer)
Dim i As Integer, j As Integer
   
   Combo3((Index - 1) * 2).Clear
   Combo3((Index - 1) * 2).AddItem ""
   Combo3((Index - 1) * 2 + 1).Clear
   Combo3((Index - 1) * 2 + 1).AddItem ""
   
   If Text6(Index - 1) <> "" Then
      'strExc(0) = "SELECT CU40,CU43,CU46,CU49,CU52,CU55 FROM CUSTOMER WHERE " & ChgCustomer(Text6(Index - 1))
      strExc(0) = "SELECT CU39,CU42,CU45,CU48,CU51,CU54 FROM CUSTOMER WHERE " & ChgCustomer(Text6(Index - 1))
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         For j = 1 To 6
            If IsNull(RsTemp.Fields(j - 1)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(j - 1)
            End If
            Combo3((Index - 1) * 2).AddItem Text6(Index - 1) & "-" & j & strExc(0)
            Combo3((Index - 1) * 2 + 1).AddItem Text6(Index - 1) & "-" & j & strExc(0)
         Next
      End If
   End If
End Sub



Private Sub Text10_GotFocus(Index As Integer)
  TextInverse Text10(Index)
End Sub

Private Sub Text10_LostFocus(Index As Integer)
 Dim i As Integer, bolChk As Boolean
   Select Case Index
      Case 1
        
      Case 3
         If IsEmptyText(Text10(0)) = False And IsEmptyText(Text10(1)) = True Then
            MsgBox "領證及繳納年費年度輸入不完整，請查明後再輸入 !", vbCritical
            Text10(0).SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(Text10(0)) = True And IsEmptyText(Text10(1)) = False Then
            MsgBox "領證及繳納年費年度輸入不完整，請查明後再輸入 !", vbCritical
            Text10(0).SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(Text10(2)) = False And IsEmptyText(Text10(3)) = True Then
            MsgBox "繳納年費年度輸入不完整，請查明後再輸入 !", vbCritical
            Text10(2).SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(Text10(2)) = True And IsEmptyText(Text10(3)) = False Then
            MsgBox "繳納年費年度輸入不完整，請查明後再輸入 !", vbCritical
            Text10(2).SetFocus
            GoTo EXITSUB
         End If
         If IsEmptyText(Text10(0)) = False And IsEmptyText(Text10(2)) = False Then
            MsgBox "領證及繳納年費不可同時輸入 !", vbCritical
            Text10(0).SetFocus
            GoTo EXITSUB
         End If
   End Select
EXITSUB:
End Sub

Private Sub Text10_Validate(Index As Integer, Cancel As Boolean)
   Dim nPos As Integer
   Dim nCurrPos As Integer
   Dim aryCaseFee As Variant
   Dim aryCurrFee As Variant
   Dim bFind As Boolean
 
   Cancel = False
   Select Case Index
      Case 0:
         If IsEmptyText(Text10(0)) = False Then
            If Text10(0) <> "1" Then
               MsgBox "領證必須從第一年開始繳 !", vbCritical
               Cancel = True
               GoTo EXITSUB
            End If
         End If
      Case 1:
         If IsEmptyText(Text10(1)) = False Then
            aryCaseFee = Split(m_CaseFee(2), ",")
            aryCurrFee = Split(m_OldCaseFee, ",")
            ' 找尋繳費年度迄在繳費年度串列中的位置(是否存在?)
            bFind = False
            For nCurrPos = 0 To UBound(aryCaseFee) - 1
               If Text10(1) = aryCaseFee(nCurrPos) Then
                  bFind = True
                  Exit For
               End If
            Next nCurrPos
            If bFind = False Then
               MsgBox "輸入之年費年度(迄)不正確, 請查明後再輸入!", vbCritical
               Cancel = True
               GoTo EXITSUB
            End If
            ' 輸入的迄值必須不在已繳過的年度中
            For nCurrPos = 0 To UBound(aryCurrFee) - 1
               If Text10(1) = aryCurrFee(nCurrPos) Then
                  MsgBox "輸入之年費年度(迄)不正確, 請查明後再輸入!", vbCritical
                  Cancel = True
                  GoTo EXITSUB
                  Exit For
               End If
            Next nCurrPos
         End If
      Case 2:
         If IsEmptyText(Text10(2)) = False Then
            If IsNumeric(Text10(2)) = False Then
               Cancel = True
               MsgBox "請輸入正確的數值 !", vbCritical
               GoTo EXITSUB
            End If
            aryCaseFee = Split(m_CaseFee(2), ",")
            aryCurrFee = Split(m_OldCaseFee, ",")
            ' 找尋已繳年度串列中空白的位置
            For nPos = 0 To UBound(aryCurrFee) - 1
               If IsEmptyText(aryCurrFee(nPos)) = True Then
                  Exit For
               End If
            Next nPos
            If nPos > UBound(aryCaseFee) - 1 Then
               MsgBox "無繳年費年度，請查明後再輸入 !", vbCritical
               Cancel = True
               GoTo EXITSUB
            Else
               If Text10(2) <> aryCaseFee(nPos) Then
                  MsgBox "起始繳費年度錯誤，請查明後再輸入 !", vbCritical
                  Cancel = True
                  GoTo EXITSUB
               Else
                  Cancel = False
               End If
            End If
            Erase aryCurrFee
            Erase aryCaseFee
         End If
      Case 3:
         If IsEmptyText(Text10(3)) = False Then
            If IsNumeric(Text10(3)) = False Then
               Cancel = True
               MsgBox "請輸入正確的數值 !", vbCritical
               GoTo EXITSUB
            End If
            aryCaseFee = Split(m_CaseFee(2), ",")
            bFind = False
            ' 找尋繳費年度迄在繳費年度串列中的位置(是否存在?)
            For nCurrPos = 0 To UBound(aryCaseFee) - 1
               If Text10(3) = aryCaseFee(nCurrPos) Then
                  bFind = True
                  Exit For
               End If
            Next nCurrPos
            ' 數入的年度不在繳費年度串列中
            If bFind = False Then
               MsgBox "繳費年度迄輸入錯誤，請查明後再輸入 !", vbCritical
               Cancel = True
               GoTo EXITSUB
            Else
               ' 找尋繳年度起在繳費年度串列中的位置
               bFind = False
               For nPos = 0 To UBound(aryCaseFee) - 1
                  If Text10(2) = aryCaseFee(nPos) Then
                     bFind = True
                     Exit For
                  End If
               Next nPos
               ' 繳費年度起及迄的範圍不正確
               If nPos > nCurrPos Then
                  MsgBox "繳費年度範圍輸入錯誤，請查明後再輸入 !", vbCritical
                  Cancel = True
                  GoTo EXITSUB
               Else
                  Cancel = False
               End If
            End If
            Erase aryCaseFee
         End If
   End Select
EXITSUB:
End Sub

Private Sub Text14_GotFocus()
  TextInverse Text14
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 89 And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text5_GotFocus()
  TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   Cancel = Not ChkLetterDate(Text5.Text)
   If Cancel = True Then TextInverse Text5
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.Text5.Enabled = True Then
   Cancel = False
   Text5_Validate Cancel
   If Cancel = True Then
      Me.Text5.SetFocus
      Text5_GotFocus
      Exit Function
   End If
End If

For ii = 0 To 4
   If Me.Text6(ii).Enabled = True Then
      Cancel = False
      Text6_Validate ii, Cancel
      If Cancel = True Then
         Me.Text6(ii).SetFocus
         Text6_GotFocus ii
         Exit Function
      End If
   End If
Next

For Each objTxt In Text10
   If objTxt.Enabled = True Then
      Cancel = False
      Text10_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Me.Text10(objTxt.Index).SetFocus
         Text10_GotFocus objTxt.Index
         Exit Function
      End If
   End If
Next

   'Add by Morgan 2005/7/29
   If lstNameAgent.Visible = True Then
      Cancel = False
      lstNameAgent_Validate Cancel
      If Cancel = True Then
         lstNameAgent.SetFocus
         Exit Function
      End If
   End If

TxtValidate = True
End Function

'Add by Morgan 2005/7/29
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/7 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Morgan 2021/12/10 Forms2.0 改用模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         'end 2021/12/10
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      m_CP22 = ""
   Else
      m_CP22 = "N"
      If MsgBox("未勾選代理人，確定不出名？", vbYesNo + vbDefaultButton2 + vbInformation) = vbNo Then
         Cancel = True
      End If
   End If
End Sub

Private Sub Text6_GotFocus(Index As Integer)
   TextInverse Text6(Index)
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text6_Validate(Index As Integer, Cancel As Boolean)
   
   If Text6(Index).Tag <> Text6(Index) Then
      Combo2(Index).Clear
      If Text6(Index) <> "" Then
         If ClsLawGetCusCAJnam(Text6(Index).Text, strExc(1), strExc(2), strExc(3)) Then
            Combo2(Index).AddItem "中:" & strExc(1)
            Combo2(Index).AddItem "英:" & strExc(2)
            Combo2(Index).AddItem "日:" & strExc(3)
            Combo2(Index).ListIndex = 0
            Call SetCombo3(Index + 1) 'Add By Sindy 2019/1/17
         Else
            Cancel = True
         End If
      End If
   End If
   If Cancel = True Then
      TextInverse Text6(Index)
   Else
      Text6(Index).Tag = Text6(Index)
   End If
End Sub

'Add By Sindy 2022/12/28
Private Sub TextPA178_KeyPress(KeyAscii As Integer)
   If (KeyAscii < 49 Or KeyAscii > 50) And KeyAscii <> 8 Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtCP84_GotFocus()
   TextInverse txtCP84
   CloseIme
End Sub

Private Sub txtCP84_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Private Sub txtCP84_Validate(Cancel As Boolean)
'   If pa(9) = "000" Then
'      If Val(txtCP84.Text) <> Val(cp(17)) And Val(txtCP84.Text) <> Val(txtCP84.Tag) Then
'         If MsgBox("發文規費【" & txtCP84.Text & "】與收文規費【" & cp(17) & "】不同，確定要繼續！", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
'            txtCP84.Tag = txtCP84.Text
'         Else
'            txtCP84_GotFocus
'            Cancel = True
'         End If
'      End If
'   End If
'End Sub
