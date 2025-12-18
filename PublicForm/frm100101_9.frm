VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_9 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料（網域）"
   ClientHeight    =   6320
   ClientLeft      =   0
   ClientTop       =   1000
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6320
   ScaleWidth      =   9310
   Begin VB.CommandButton cmdok 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   6
      Left            =   4062
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "各項指示"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   5
      Left            =   2055
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   940
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   4
      Left            =   3011
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   5640
      Left            =   0
      TabIndex        =   15
      Top             =   435
      Width           =   9270
      _ExtentX        =   16334
      _ExtentY        =   9948
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_9.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label92"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label91(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label84"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label89"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label88"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label11"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label19"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label8"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label3"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label10"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lbl1(1)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl1(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "lbl1(4)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl1(5)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lbl1(6)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbl1(8)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl1(9)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label91(1)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label91(2)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label12(0)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label12(1)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label4"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl1(25)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl1(85)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Label113"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Label112"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl1(87)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Label14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl1(86)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label21"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txt1(4)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txt1(1)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt1(0)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt1(3)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt1(5)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt1(2)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt1(6)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt1(8)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt1(9)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Line1"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "代理人相關資料"
      TabPicture(1)   =   "frm100101_9.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt1(7)"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "lbl1(84)"
      Tab(1).Control(3)=   "lbl1(21)"
      Tab(1).Control(4)=   "Label13"
      Tab(1).Control(5)=   "lbl1(17)"
      Tab(1).Control(6)=   "lbl1(16)"
      Tab(1).Control(7)=   "lbl1(15)"
      Tab(1).Control(8)=   "lbl1(14)"
      Tab(1).Control(9)=   "lbl1(13)"
      Tab(1).Control(10)=   "lbl1(12)"
      Tab(1).Control(11)=   "lbl1(11)"
      Tab(1).Control(12)=   "lbl1(10)"
      Tab(1).Control(13)=   "Label31"
      Tab(1).Control(14)=   "Label28"
      Tab(1).Control(15)=   "Label24"
      Tab(1).Control(16)=   "Label20"
      Tab(1).Control(17)=   "Label18"
      Tab(1).Control(18)=   "Label16"
      Tab(1).Control(19)=   "Label15"
      Tab(1).Control(20)=   "Label29"
      Tab(1).Control(21)=   "Label7"
      Tab(1).ControlCount=   22
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm100101_9.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl1(24)"
      Tab(2).Control(1)=   "lbl1(23)"
      Tab(2).Control(2)=   "lbl1(22)"
      Tab(2).Control(3)=   "lbl1(20)"
      Tab(2).Control(4)=   "Label78"
      Tab(2).Control(5)=   "Label79"
      Tab(2).Control(6)=   "Label80"
      Tab(2).Control(7)=   "Label81"
      Tab(2).ControlCount=   8
      Begin VB.Line Line1 
         X1              =   2100
         X2              =   2190
         Y1              =   2490
         Y2              =   2490
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   9
         Left            =   5430
         TabIndex        =   81
         Top             =   1681
         Width           =   2205
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3889;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   8
         Left            =   1125
         TabIndex        =   80
         Top             =   330
         Width           =   2205
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3889;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   960
         Index           =   7
         Left            =   -74880
         TabIndex        =   38
         Top             =   3930
         Width           =   9060
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "15981;1693"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   6
         Left            =   1155
         TabIndex        =   12
         Top             =   4782
         Width           =   8025
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14155;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   2
         Left            =   1404
         TabIndex        =   9
         Top             =   1353
         Width           =   7740
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13652;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   5
         Left            =   1155
         TabIndex        =   11
         Top             =   4293
         Width           =   8025
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14155;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   3
         Left            =   1155
         TabIndex        =   10
         Top             =   1980
         Width           =   8010
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "14129;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   0
         Left            =   1404
         TabIndex        =   7
         Top             =   651
         Width           =   7740
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13652;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   1
         Left            =   1404
         TabIndex        =   8
         Top             =   1002
         Width           =   7740
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13652;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   5490
         TabIndex        =   13
         Top             =   2610
         Width           =   3690
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "6509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   255
         Left            =   150
         TabIndex        =   88
         Top             =   3750
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   86
         Left            =   1680
         TabIndex        =   87
         Top             =   3741
         Width           =   4140
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "7302;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   255
         Left            =   5970
         TabIndex        =   86
         Top             =   3750
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   87
         Left            =   7455
         TabIndex        =   85
         Top             =   3750
         Width           =   1710
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3016;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   255
         Left            =   1710
         TabIndex        =   84
         Top             =   5310
         Width           =   2115
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   255
         Left            =   150
         TabIndex        =   83
         Top             =   5310
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   85
         Left            =   1410
         TabIndex        =   82
         Top             =   5310
         Width           =   270
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "476;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   960
         Width           =   1860
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   84
         Left            =   -72960
         TabIndex        =   78
         Top             =   960
         Width           =   3780
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6667;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   6750
         TabIndex        =   77
         Top             =   3465
         Width           =   2385
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4207;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   255
         Left            =   5940
         TabIndex        =   76
         Top             =   3465
         Width           =   720
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -73590
         TabIndex        =   75
         Top             =   1230
         Width           =   4725
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "8334;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73590
         TabIndex        =   74
         Top             =   930
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73590
         TabIndex        =   73
         Top             =   660
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   -73590
         TabIndex        =   72
         Top             =   390
         Width           =   1000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1764;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   255
         Left            =   -74880
         TabIndex        =   71
         Top             =   1230
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   255
         Left            =   -74880
         TabIndex        =   70
         Top             =   930
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   255
         Left            =   -74880
         TabIndex        =   69
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   255
         Left            =   -74880
         TabIndex        =   68
         Top             =   390
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   -73290
         TabIndex        =   66
         Top             =   3240
         Width           =   7275
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12832;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   255
         Left            =   -74880
         TabIndex        =   67
         Top             =   3240
         Width           =   1545
      End
      Begin MSForms.Label Label12 
         Height          =   255
         Index           =   1
         Left            =   5640
         TabIndex        =   65
         Top             =   4017
         Width           =   3195
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5636;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label12 
         Height          =   255
         Index           =   0
         Left            =   1155
         TabIndex        =   64
         Top             =   4017
         Width           =   2835
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5001;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "申請者代號："
         Height          =   255
         Index           =   2
         Left            =   4530
         TabIndex        =   63
         Top             =   4017
         Width           =   1080
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "主管機關："
         Height          =   255
         Index           =   1
         Left            =   150
         TabIndex        =   62
         Top             =   4017
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   -73800
         TabIndex        =   60
         Top             =   2880
         Width           =   7815
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13785;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   1410
         TabIndex        =   59
         Top             =   3465
         Width           =   4320
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "7620;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73800
         TabIndex        =   57
         Top             =   2520
         Width           =   7785
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13732;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   -73200
         TabIndex        =   56
         Top             =   2160
         Width           =   495
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   -73560
         TabIndex        =   55
         Top             =   1860
         Width           =   7650
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13494;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -73800
         TabIndex        =   54
         Top             =   1560
         Width           =   1935
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3413;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73800
         TabIndex        =   53
         Top             =   1260
         Width           =   8000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -73800
         TabIndex        =   52
         Top             =   660
         Width           =   6570
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11589;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   -73800
         TabIndex        =   51
         Top             =   360
         Width           =   8000
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "14111;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   5565
         TabIndex        =   50
         Top             =   2361
         Width           =   645
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1138;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   1155
         TabIndex        =   49
         Top             =   3189
         Width           =   2820
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4974;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   1155
         TabIndex        =   48
         Top             =   2913
         Width           =   495
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   1155
         TabIndex        =   47
         Top             =   2637
         Width           =   2715
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4789;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   2250
         TabIndex        =   46
         Top             =   2361
         Width           =   1125
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1984;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   1155
         TabIndex        =   45
         Top             =   2361
         Width           =   915
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "1614;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1155
         TabIndex        =   44
         Top             =   1704
         Width           =   3120
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y：閉卷）"
         Height          =   255
         Left            =   6360
         TabIndex        =   43
         Top             =   2361
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷："
         Height          =   255
         Left            =   4530
         TabIndex        =   42
         Top             =   2361
         Width           =   900
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   39
         Top             =   3720
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   37
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：            （Y：印）"
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   2160
         Width           =   3105
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   1860
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   33
         Top             =   1260
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "折扣："
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1560
         Width           =   540
      End
      Begin VB.Label Label3 
         Caption         =   "客戶備註："
         Height          =   195
         Left            =   150
         TabIndex        =   31
         Top             =   4830
         Width           =   975
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請日："
         Height          =   255
         Left            =   4560
         TabIndex        =   30
         Top             =   1704
         Width           =   720
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "使用期限：           "
         Height          =   255
         Left            =   150
         TabIndex        =   29
         Top             =   2361
         Width           =   1395
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   255
         Left            =   150
         TabIndex        =   28
         Top             =   1704
         Width           =   720
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "密碼："
         Height          =   180
         Left            =   150
         TabIndex        =   27
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         Height          =   180
         Left            =   150
         TabIndex        =   26
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日)："
         Height          =   180
         Left            =   150
         TabIndex        =   25
         Top             =   1428
         Width           =   1200
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英)："
         Height          =   180
         Left            =   150
         TabIndex        =   24
         Top             =   1077
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中)："
         Height          =   180
         Left            =   150
         TabIndex        =   23
         Top             =   726
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   21
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   255
         Left            =   -74880
         TabIndex        =   20
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   255
         Left            =   150
         TabIndex        =   19
         Top             =   2637
         Width           =   900
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因："
         Height          =   255
         Left            =   4500
         TabIndex        =   18
         Top             =   2637
         Width           =   900
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定搞語文：                 (１.中文　２.英文 ３.日文)"
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   2913
         Width           =   3765
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   3189
         Width           =   900
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   150
         TabIndex        =   14
         Top             =   3465
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "代理人資料"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   4878
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7510
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "申請人資料"
      Height          =   400
      Index           =   0
      Left            =   6194
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8430
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   800
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   19
      Left            =   5760
      TabIndex        =   61
      Top             =   6075
      Width           =   3525
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "6218;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   1050
      TabIndex        =   58
      Top             =   6075
      Width           =   3615
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "6376;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   255
      Left            =   4770
      TabIndex        =   41
      Top             =   6075
      Width           =   945
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   255
      Left            =   90
      TabIndex        =   40
      Top             =   6075
      Width           =   855
   End
End
Attribute VB_Name = "frm100101_9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/20 改成Form2.0 ; lbl1(index)、txt1(index)、Label12(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
Dim m_pub_QL05 As String 'Add By Sindy 2025/8/28 只記錄於此Form


'92.04.16 nick
Public Sub PubShowNextData()
Select Case cmdState
Case 0
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag1
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 2
     fnCloseAllFrm100
Case 3
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_10.Show
     frm100101_10.Tag = StrTag
     frm100101_10.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_10.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'add by nickc 2005/05/31
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(8).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Added by Lydia 2016/11/23
Case 5 '各項指示
    'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
    If PUB_CheckFormExist("frm12040159") Then
        MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
        Exit Sub
    End If
    'end 2020/05/05
    
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm12040159.SetParent "Q", Trim(Replace(txt1(8), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 6 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(8)
   frm100101_2.StrMenu
   Screen.MousePointer = vbDefault
   Me.Enabled = True
Case Else
End Select
End Sub

Private Sub cmdok_Click(Index As Integer)
'92.04.16 nick 紀錄作用按鍵
cmdState = Index
PubShowNextData
End Sub

Sub StrMenu()
Dim strSql  As String
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
'edit by nickc 2006/07/12
'Dim strArr(T_SP) As String, i As Integer, StrOk(21) As String, StrOkTxt(7) As String
Dim strArr() As String, i As Integer, StrOk(21) As String, StrOkTxt(7) As String
ReDim strArr(tf_SP) As String

'Add By Cheng 2002/07/08
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSK03 As String
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String

'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'End by Toni 20080926

Str01 = ""
Str02 = ""
Str03 = ""
Str04 = ""
If Left(Me.Tag, 1) = "N" Then
   strSql = Right(Me.Tag, Len(Me.Tag) - 1)
Else
   strSql = Me.Tag
End If
Str01 = SystemNumber(strSql, 1)
Str02 = SystemNumber(strSql, 2)
Str03 = SystemNumber(strSql, 3)
Str04 = SystemNumber(strSql, 4)

'add by Toni 20080926 控制跨部門權限 for 網域
'2008/10/2 modify by sonia
'If IsUserHasRightOfSystem(strUserNum, Str01) = False Then
'   If IsUserHasRightOfFunction("frm100101_1", strCrossDept, False) = False Then
'      strTit = "檢核資料"
'      strMsg = "您沒有使用該系統類別的權限"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      tmpBol = fnCancelNowFormAndShowParentForm(Me)
'      Exit Sub
'   End If
'End If
If CheckSR09(strUserNum, Str01, "Y", , Str01, Str02, Str03, Str04) = False Then
   tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
'2008/10/2 end
'End 20080926

pub_QL05 = ";本所案號：" & Str01 & "-" & Str02 & "-" & Str03 & "-" & Str04 & _
           "(基本資料)" 'Add By Sindy 2025/8/7

'Add By Cheng 2002/07/08
strSK03 = ""
StrSQLa = "Select SK03 From SystemKind Where SK01='" & Str01 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic
If rsA.RecordCount > 0 Then
   strSK03 = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

'欲搜尋的SQL字串
strSql = "SELECT * FROM SERVICEPRACTICE WHERE SP01='" & Str01 & "' AND SP02='" & Str02 & "' AND SP03='" & Str03 & "' AND SP04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/7
'   For i = 0 To 61
   For i = 0 To (tf_SP - 1) 'edit by nickc 2006/07/12 T_SP - 1)
      Select Case i
      Case 9, 11, 15, 19, 20, 30, 38, 39, 52, 53, 55, 56
           If IsNull(adoRecordset.Fields(i)) Then
               strArr(i + 1) = ""
           Else
               strArr(i + 1) = str(adoRecordset.Fields(i))
           End If
      Case Else
           If IsNull(adoRecordset.Fields(i)) Then
                strArr(i + 1) = ""
           Else
                strArr(i + 1) = adoRecordset.Fields(i)
           End If
      End Select
      DoEvents
   Next i
   If IsNull(adoRecordset.Fields("SP51")) Then
      Label12(0) = ""
   Else
      Label12(0) = adoRecordset.Fields("SP51")
   End If
   If IsNull(adoRecordset.Fields("SP64")) Then
      Label12(1) = ""
   Else
      Label12(1) = adoRecordset.Fields("SP64")
   End If
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
   ShowNoData
   Screen.MousePointer = vbDefault
       '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 62
'For i = 0 To 62
For i = 1 To tf_SP 'edit by nickc 2006/07/12 T_SP
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         txt1(8) = StrOk(0) 'Add By Sindy 2013/1/31
    Case 5
         StrOkTxt(0) = strArr(i)
    Case 6
         StrOkTxt(1) = strArr(i)
    Case 7
         StrOkTxt(2) = strArr(i)
    Case 8
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
              strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(1) = strArr(i) + ""
'                     Else
'                          StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(1) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
            If IsNull(adoRecordset.Fields("CU04")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU04")
            ElseIf IsNull(adoRecordset.Fields("CU05")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU05")
            ElseIf IsNull(adoRecordset.Fields("CU06")) = False Then
               StrOk(1) = strArr(i) + "  " + adoRecordset.Fields("CU06")
            End If
             If IsNull(adoRecordset.Fields(3)) Then
                  StrOkTxt(6) = ""
             Else
                  StrOkTxt(6) = adoRecordset.Fields(3)
             End If
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbBlack
         Else
            StrOk(1) = ""
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbRed
             StrOk(1) = strArr(i)
             StrOkTxt(6) = ""
         End If
         CheckOC
    Case 20
         StrOk(2) = strArr(i)
    Case 21
         StrOk(3) = strArr(i)
    Case 16
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(4) = ""
         Else
             StrOk(4) = ChangeWStringToTString(strArr(i))
         End If
    Case 34
         StrOk(5) = strArr(i)
    Case 28
         StrOk(6) = strArr(i)
    Case 10
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(7) = ""
         Else
             StrOk(7) = ChangeWStringToTString(strArr(i))
         End If
         txt1(9) = StrOk(7) 'Add By Sindy 2013/1/31
    Case 15
         StrOk(8) = strArr(i)
    Case 29
         StrOk(9) = strArr(i)
    Case 26
         If Len(strArr(i)) = 9 Then
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
              strSql = "SELECT FA05,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
         Else
              'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
              'strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
              strSql = "SELECT FA05,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
         End If
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            '2005/9/15 MODIFY BY SONIA
            'If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            If Trim(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) = "" Then
'            '2005/9/15 END
'               'Modify By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(10) = strArr(i) + ""
'                   Else
'                         StrOk(10) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                  'Modify By Cheng 2002/07/08
''                   StrOk(10) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   'edit by nickc 2006/10/14 修正 null 的錯誤
'                   'StrOk(10) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))
'                   StrOk(10) = strArr(i) + "  " + CheckStr(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1)))
'               End If
'            Else
'               'Modify By Cheng 2002/07/08
''               StrOk(10) = StrArr(i) + "  " + adoRecordset.Fields(0)
'               'edit by nickc 2006/10/14 修正 null 的錯誤
'               'StrOk(10) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))
'               StrOk(10) = strArr(i) + "  " + CheckStr(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0)))
'
'            End If
            If IsNull(adoRecordset.Fields("FA05")) = False Then
               StrOk(10) = strArr(i) + "  " + adoRecordset.Fields("FA05")
            ElseIf IsNull(adoRecordset.Fields("FA04")) = False Then
               StrOk(10) = strArr(i) + "  " + adoRecordset.Fields("FA04")
            ElseIf IsNull(adoRecordset.Fields("FA06")) = False Then
               StrOk(10) = strArr(i) + "  " + adoRecordset.Fields("FA06")
            End If
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(7) = ""
            Else
                StrOkTxt(7) = adoRecordset.Fields(3)
            End If
         Else
            StrOk(10) = ""
            StrOkTxt(7) = ""
         End If
         CheckOC
    Case 27
         StrOk(11) = strArr(i)
    Case 30
         StrOk(12) = strArr(i)
    Case 31
         StrOk(13) = strArr(i)
    Case 37
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(14) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
               StrOk(14) = strArr(i) + "  " + tmp02
            Else
               StrOk(14) = strArr(i)
            End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(14) = strArr(i) + ""
'                    Else
'                        StrOk(14) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                     'Modify By Cheng 2002/07/08
''                    StrOk(14) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(14) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(14) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(14) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(14) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(14).ForeColor = vbBlack
         Else
            StrOk(14) = ""
            'Add by Morgan 2004/1/16
            Lbl1(14).ForeColor = vbRed
            StrOk(14) = strArr(i)
         End If
         CheckOC
    Case 33
         StrOk(15) = strArr(i)
    Case 35
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(16) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                StrOk(16) = strArr(i) + "  " + tmp02
             Else
                StrOk(16) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Modify By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Modify By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(16) = strArr(i) + ""
'                    Else
'                        StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Modify By Cheng 2002/07/08
''                    StrOk(16) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Modify By Cheng 2002/07/08
''                StrOk(16) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(16) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(16) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(16).ForeColor = vbBlack
         Else
            StrOk(16) = ""
            'Add by Morgan 2004/1/6
            Lbl1(16).ForeColor = vbRed
            StrOk(16) = strArr(i)
         End If
         CheckOC
    Case 36
         StrOk(17) = strArr(i)
    Case 52
         'edit by nick 2004/10/05
         'StrOk(18) = strArr(i)
         StrOk(18) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(53))) & " " & Format(strArr(54), "##:##")
    Case 55
         'edit by nick 2004/10/05
         'StrOk(19) = strArr(i)
         StrOk(19) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(56))) & " " & Format(strArr(57), "##:##")
    Case 61
         'edit by nickc 2006/07/12
         'StrOk(20) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(20) = ""
         Else
             StrOk(20) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 18
         StrOkTxt(5) = strArr(i)
    Case 49
         StrOkTxt(3) = strArr(i)
    Case 17
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOkTxt(4) = ""
             Else
                     StrOkTxt(4) = adoRecordset.Fields(0)
             End If
         Else
             StrOkTxt(4) = ""
         End If
         CheckOC
    Case 67 'D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
              StrOk(21) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
                If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                   StrOk(21) = strArr(i) + "  " + tmp02
                Else
                   StrOk(21) = strArr(i)
                End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(21) = strArr(i) + ""
'                    Else
'                        StrOk(21) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(21) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(21) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
'         Else
'            StrOk(21) = ""
'         End If
'         CheckOC
    'add by nickc 2006/07/12
    Case 68
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(22) = ""
         Else
             Lbl1(22) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 69
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(23) = strArr(i) + ""
            Else
               Lbl1(23) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(23) = ""
         End If
         CheckOC
    Case 70
         Lbl1(24) = strArr(i)
    'Add by Morgan 2008/8/5
    Case 78
         Lbl1(25) = PUB_GetContact(strArr(8), strArr(i))
    Case 84 'Add by Morgan 2010/11/8
         Lbl1(i) = strArr(i)
    Case 85 'Add by Sindy 2014/2/10
         Lbl1(i) = strArr(i)
    'Added by Morgan 2016/12/8
    Case 86 '國內副本收件人
         Lbl1(i) = strArr(i)
         If strArr(i) <> "" Then
            If ClsLawLawGetName(strArr(i), strExc(9)) = True Then
               Lbl1(i) = Lbl1(i) + "  " + strExc(9)
            End If
         End If
    Case 87 '國內副本接洽人
         If strArr(86) <> "" And strArr(i) <> "" Then
            Lbl1(i) = PUB_GetContact(strArr(86), strArr(i))
         Else
            Lbl1(i) = ""
         End If
    'end 2016/12/8
    Case Else
    End Select
    DoEvents
Next i
For i = 0 To 21               '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
   If i <> 0 And i <> 7 Then 'Add By Sindy 2013/1/31
      Lbl1(i) = StrOk(i)        '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
   End If
Next i
For i = 0 To 7
   txt1(i) = StrOkTxt(i)
Next i
'傳入參數     代理人
StrTag = strArr(26)
'傳入參數     申請人
StrTag1 = strArr(8)
cmdok(4).Visible = ChkDataByCR(txt1(8).Text)
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

'edit by nickc 2005/05/31
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(8), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
Dim Lbl As Object

For Each Lbl In Me.Lbl1
    Lbl.BackColor = &H8000000F
Next
For Each Lbl In Me.Label12
    Lbl.BackColor = &H8000000F
Next
bolToEndByNick = False

SSTab3.Tab = 0 'Added by Lydia 2016/11/23

   MoveFormToCenter Me
   If bolFNation = False Then
        SSTab3.TabVisible(1) = False
        cmdok(3).Visible = False
   End If
   'Add By Cheng 2002/04/24
   'modify by sonia 2016/7/27 以M51及P20控制即可
   'If strUserNum = "67002" Then
   '   Me.Label22.Visible = True
   '   Me.txt1(3).Visible = True
   'ElseIf GetUserDept(strUserNum) = "M51" Or GetUserDept(strUserNum) = "P20" Then
   'Modified by Lydia 2023/12/20 改共用變數
   'If GetUserDept(strUserNum) = "M51" Or GetUserDept(strUserNum) = "P20" Then
   ''end 2016/7/27
   If Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "P20" Then
   'end 2023/12/20
      Me.Label22.Visible = True
      Me.txt1(3).Visible = True
   Else
      Me.Label22.Visible = False
      Me.txt1(3).Visible = False
   End If
'92.04.16 nick
cmdState = -1

'Added by Lydia 2020/05/05 各項指示：顯示按鈕
If strSrvDate(1) >= 各項指示啟用日 Then
   cmdok(5).Visible = True
Else
   cmdok(5).Visible = False
End If
'end 2020/05/05
End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_9 = Nothing
End Sub

'Mark by Lydia 2023/12/20 改共用變數
'Private Function GetUserDept(strUserNum) As String
'Dim Rs As New ADODB.Recordset
'If Rs.State <> adStateClosed Then Rs.Close
'Set Rs = Nothing
'Rs.CursorLocation = adUseClient
'Rs.Open "SELECT ST03 FROM STAFF WHERE ST01='" & strUserNum & "'", cnnConnection, adOpenStatic, adLockReadOnly
'If Rs.RecordCount > 0 Then
'   GetUserDept = "" & Rs.Fields(0).Value
'Else
'   GetUserDept = ""
'End If
'If Rs.State <> adStateClosed Then Rs.Close
'Set Rs = Nothing
'
'End Function
'end 2023/12/20

'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
' 取得客戶或是代理人名稱
Private Function GetAgentOrCustName(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   GetAgentOrCustName = Empty
   If IsEmptyText(strData) = False Then
      ' 不滿8碼自動補0
      If Len(strData) < 8 Then: strData = strData & String(8 - Len(strData), "0")
      Select Case Mid(strData, 1, 1)
      Case "X", "x":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM Customer " & _
                     "WHERE CU01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "CU02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("CU05")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetAgentOrCustName = rsTmp.Fields("CU06")
            End If
         End If
         rsTmp.Close
      Case "Y", "y":
         Set rsTmp = New ADODB.Recordset
         If Len(strData) > 8 Then
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '" & Mid(strData, 9, 1) & "'"
         Else
            strSql = "SELECT * FROM FAGENT " & _
                     "WHERE FA01 = '" & Mid(strData, 1, 8) & "' AND " & _
                           "FA02 = '0' "
         End If
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA05")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA05")
            ElseIf IsNull(rsTmp.Fields("FA04")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA04")
            ElseIf IsNull(rsTmp.Fields("FA06")) = False Then
               GetAgentOrCustName = rsTmp.Fields("FA06")
            End If
         End If
         rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Function

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
