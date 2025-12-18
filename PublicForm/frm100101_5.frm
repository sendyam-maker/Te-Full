VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_5 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件資料及案件進度查詢（法務基本資料）"
   ClientHeight    =   5750
   ClientLeft      =   150
   ClientTop       =   990
   ClientWidth     =   8960
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8960
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   10
      Left            =   4110
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   45
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項指示"
      Height          =   375
      Index           =   9
      Left            =   405
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   45
      Visible         =   0   'False
      Width           =   940
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4752
      Left            =   0
      TabIndex        =   15
      Top             =   468
      Width           =   8928
      _ExtentX        =   15752
      _ExtentY        =   8378
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_5.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label2(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl1(28)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label13"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl1(30)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label27"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label14"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label24"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label6(2)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label8(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl1(14)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label16"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label26"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl1(15)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label23"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label19"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label12"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lbl1(27)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label6(0)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label8(4)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lbl1(18)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label15"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label28"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "lbl1(16)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "lbl1(13)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl1(29)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "lbl1(31)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "lbl1(1)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "lbl1(2)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "lbl1(3)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "lbl1(32)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl1(33)"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Label8(1)"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Label6(1)"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "lbl1(52)"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "Label8(5)"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Label6(3)"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt1(0)"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txt1(1)"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txt1(2)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt1(4)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "FC資料"
      TabPicture(1)   =   "frm100101_5.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo3(1)"
      Tab(1).Control(1)=   "txt1(3)"
      Tab(1).Control(2)=   "Label1(0)"
      Tab(1).Control(3)=   "lbl1(51)"
      Tab(1).Control(4)=   "Label80(26)"
      Tab(1).Control(5)=   "lbl1(49)"
      Tab(1).Control(6)=   "Label80(29)"
      Tab(1).Control(7)=   "Label22"
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(9)=   "lbl1(23)"
      Tab(1).Control(10)=   "lbl1(12)"
      Tab(1).Control(11)=   "lbl1(11)"
      Tab(1).Control(12)=   "lbl1(10)"
      Tab(1).Control(13)=   "Label20(2)"
      Tab(1).Control(14)=   "Label20(1)"
      Tab(1).Control(15)=   "Label20(3)"
      Tab(1).Control(16)=   "lbl1(9)"
      Tab(1).Control(17)=   "Label21"
      Tab(1).Control(18)=   "Label8(0)"
      Tab(1).Control(19)=   "lbl1(19)"
      Tab(1).Control(20)=   "Label8(2)"
      Tab(1).Control(21)=   "lbl1(8)"
      Tab(1).Control(22)=   "Label9"
      Tab(1).Control(23)=   "Label11"
      Tab(1).Control(24)=   "lbl1(20)"
      Tab(1).Control(25)=   "lbl1(7)"
      Tab(1).Control(26)=   "Label7"
      Tab(1).Control(27)=   "Label17"
      Tab(1).Control(28)=   "lbl1(17)"
      Tab(1).Control(29)=   "lbl1(4)"
      Tab(1).Control(30)=   "Label18(1)"
      Tab(1).ControlCount=   31
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm100101_5.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl1(5)"
      Tab(2).Control(1)=   "Label81"
      Tab(2).Control(2)=   "Label80(0)"
      Tab(2).Control(3)=   "Label79"
      Tab(2).Control(4)=   "Label78"
      Tab(2).Control(5)=   "lbl1(24)"
      Tab(2).Control(6)=   "lbl1(25)"
      Tab(2).Control(7)=   "lbl1(26)"
      Tab(2).ControlCount=   8
      Begin VB.ComboBox Combo3 
         Height          =   300
         Index           =   1
         ItemData        =   "frm100101_5.frx":0054
         Left            =   -71490
         List            =   "frm100101_5.frx":0067
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   92
         Top             =   3150
         Width           =   1470
      End
      Begin MSForms.TextBox txt1 
         Height          =   288
         Index           =   4
         Left            =   1050
         TabIndex        =   88
         Top             =   420
         Width           =   2295
         VariousPropertyBits=   671105055
         Size            =   "4048;508"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   2
         Left            =   1590
         TabIndex        =   59
         Top             =   2220
         Width           =   7200
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "12700;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   1
         Left            =   1590
         TabIndex        =   58
         Top             =   1840
         Width           =   7200
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "12700;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   360
         Index           =   0
         Left            =   1590
         TabIndex        =   57
         Top             =   1460
         Width           =   7200
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "12700;635"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   660
         Index           =   3
         Left            =   -73950
         TabIndex        =   55
         Top             =   3480
         Width           =   7650
         VariousPropertyBits=   -1466941409
         ScrollBars      =   2
         Size            =   "13494;1164"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "專案服務案："
         Height          =   250
         Index           =   3
         Left            =   4170
         TabIndex        =   100
         Top             =   1232
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   250
         Index           =   5
         Left            =   6120
         TabIndex        =   99
         Top             =   1232
         Width           =   465
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   52
         Left            =   5400
         TabIndex        =   98
         Top             =   1232
         Width           =   495
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "873;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   97
         Top             =   4410
         Width           =   1860
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   51
         Left            =   -72960
         TabIndex        =   96
         Top             =   4410
         Width           =   2475
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "4360;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   255
         Index           =   26
         Left            =   -74880
         TabIndex        =   95
         Top             =   3150
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   49
         Left            =   -73950
         TabIndex        =   94
         Top             =   3150
         Width           =   435
         VariousPropertyBits=   27
         Caption         =   "TEXT"
         Size            =   "767;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   255
         Index           =   29
         Left            =   -73290
         TabIndex        =   93
         Top             =   3150
         Width           =   1800
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   91
         Top             =   4290
         Width           =   1260
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   90
         Top             =   4290
         Width           =   2115
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   1500
         TabIndex        =   89
         Top             =   4290
         Width           =   495
         VariousPropertyBits=   27
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   555
         Index           =   32
         Left            =   1095
         TabIndex        =   87
         Top             =   4020
         Width           =   7575
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13361;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   3
         Left            =   1095
         TabIndex        =   86
         Top             =   3210
         Width           =   2985
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5265;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   1095
         TabIndex        =   85
         Top             =   2940
         Width           =   960
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1693;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   84
         Top             =   2670
         Width           =   285
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "503;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   31
         Left            =   945
         TabIndex        =   83
         Top             =   1232
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5953;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   29
         Left            =   945
         TabIndex        =   82
         Top             =   975
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5953;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   13
         Left            =   945
         TabIndex        =   81
         Top             =   720
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5953;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   1440
         TabIndex        =   80
         Top             =   3480
         Width           =   2655
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "4683;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "案件屬性："
         Height          =   255
         Left            =   120
         TabIndex        =   79
         Top             =   4020
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "署名人："
         Height          =   255
         Left            =   4440
         TabIndex        =   78
         Top             =   3750
         Width           =   720
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   5430
         TabIndex        =   77
         Top             =   3750
         Width           =   3240
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5715;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   2040
         TabIndex        =   76
         Top             =   3750
         Width           =   495
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "873;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   75
         Top             =   3750
         Width           =   465
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "是否為智慧財產權案："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   74
         Top             =   3750
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   5430
         TabIndex        =   73
         Top             =   3480
         Width           =   3135
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5530;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   255
         Left            =   4440
         TabIndex        =   72
         Top             =   3480
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   3480
         Width           =   1260
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "相關國家："
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   3210
         Width           =   900
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因："
         Height          =   255
         Left            =   4440
         TabIndex        =   69
         Top             =   2940
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   5430
         TabIndex        =   68
         Top             =   2940
         Width           =   3105
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5477;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   255
         Left            =   120
         TabIndex        =   67
         Top             =   2940
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   255
         Left            =   4440
         TabIndex        =   66
         Top             =   2670
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   5430
         TabIndex        =   65
         Top             =   2670
         Width           =   3120
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5503;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Index           =   3
         Left            =   1620
         TabIndex        =   64
         Top             =   2670
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷："
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   63
         Top             =   2670
         Width           =   900
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱（中）："
         Height          =   180
         Left            =   120
         TabIndex        =   62
         Top             =   1550
         Width           =   1440
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱（英）："
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   1930
         Width           =   1440
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱（日）："
         Height          =   180
         Left            =   120
         TabIndex        =   60
         Top             =   2310
         Width           =   1440
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   56
         Top             =   3480
         Width           =   900
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   250
         Left            =   -74880
         TabIndex        =   54
         Top             =   2790
         Width           =   1545
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73380
         TabIndex        =   53
         Top             =   2790
         Width           =   7005
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12356;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   12
         Left            =   -73560
         TabIndex        =   52
         Top             =   2528
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;441"
         BorderColor     =   -2147483648
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -73560
         TabIndex        =   51
         Top             =   2270
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;450"
         BorderColor     =   -2147483648
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   -73560
         TabIndex        =   50
         Top             =   2014
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;450"
         BorderColor     =   -2147483648
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人（中）："
         Height          =   250
         Index           =   2
         Left            =   -74880
         TabIndex        =   49
         Top             =   2016
         Width           =   1260
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人（日）："
         Height          =   250
         Index           =   1
         Left            =   -74880
         TabIndex        =   48
         Top             =   2528
         Width           =   1260
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人（英）："
         Height          =   250
         Index           =   3
         Left            =   -74880
         TabIndex        =   47
         Top             =   2272
         Width           =   1260
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "當事人5："
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1230
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   30
         Left            =   5430
         TabIndex        =   45
         Top             =   975
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5953;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "當事人4："
         Height          =   250
         Left            =   4440
         TabIndex        =   44
         Top             =   975
         Width           =   810
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "當事人3："
         Height          =   250
         Left            =   120
         TabIndex        =   43
         Top             =   975
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   28
         Left            =   5430
         TabIndex        =   42
         Top             =   720
         Width           =   3375
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5953;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "當事人2："
         Height          =   250
         Left            =   4440
         TabIndex        =   41
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "當事人1："
         Height          =   250
         Index           =   1
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   810
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Index           =   1
         Left            =   120
         TabIndex        =   39
         Top             =   480
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   9
         Left            =   -73800
         TabIndex        =   38
         Top             =   1760
         Width           =   7425
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "13097;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   250
         Left            =   -74880
         TabIndex        =   37
         Top             =   1760
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人："
         Height          =   250
         Index           =   0
         Left            =   -71070
         TabIndex        =   36
         Top             =   1504
         Width           =   1725
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   19
         Left            =   -69285
         TabIndex        =   35
         Top             =   1504
         Width           =   1095
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1931;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   250
         Index           =   2
         Left            =   -68025
         TabIndex        =   34
         Top             =   1504
         Width           =   465
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   8
         Left            =   -74220
         TabIndex        =   33
         Top             =   1504
         Width           =   2895
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5106;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "折扣："
         Height          =   250
         Left            =   -74880
         TabIndex        =   32
         Top             =   1504
         Width           =   540
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人：  "
         Height          =   250
         Left            =   -74880
         TabIndex        =   31
         Top             =   1248
         Width           =   1170
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   20
         Left            =   -73575
         TabIndex        =   30
         Top             =   1248
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   7
         Left            =   -73575
         TabIndex        =   29
         Top             =   992
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   250
         Left            =   -74880
         TabIndex        =   28
         Top             =   992
         Width           =   1260
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   250
         Left            =   -74880
         TabIndex        =   27
         Top             =   736
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   17
         Left            =   -73875
         TabIndex        =   26
         Top             =   736
         Width           =   3105
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "5477;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   4
         Left            =   -73860
         TabIndex        =   25
         Top             =   480
         Width           =   7260
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12806;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人："
         Height          =   250
         Index           =   1
         Left            =   -74880
         TabIndex        =   24
         Top             =   480
         Width           =   930
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   5
         Left            =   -73770
         TabIndex        =   23
         Top             =   480
         Width           =   1290
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "2275;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   250
         Left            =   -74880
         TabIndex        =   22
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   250
         Index           =   0
         Left            =   -72150
         TabIndex        =   21
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   250
         Left            =   -69150
         TabIndex        =   20
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   255
         Left            =   -74880
         TabIndex        =   19
         Top             =   810
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   24
         Left            =   -71040
         TabIndex        =   18
         Top             =   480
         Width           =   1095
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "1931;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   250
         Index           =   25
         Left            =   -68040
         TabIndex        =   17
         Top             =   480
         Width           =   1410
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "2487;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73560
         TabIndex        =   16
         Top             =   810
         Width           =   7335
         VariousPropertyBits=   27
         Caption         =   " "
         Size            =   "12938;441"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當2"
      Height          =   375
      Index           =   5
      Left            =   1380
      TabIndex        =   9
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當3"
      Height          =   375
      Index           =   6
      Left            =   1800
      TabIndex        =   8
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當4"
      Height          =   375
      Index           =   7
      Left            =   2220
      TabIndex        =   7
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "當5"
      Height          =   375
      Index           =   8
      Left            =   2640
      TabIndex        =   6
      Top             =   45
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   375
      Index           =   4
      Left            =   3060
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   45
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   3
      Left            =   8100
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   45
      Width           =   765
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人"
      Height          =   375
      Index           =   1
      Left            =   6030
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   45
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人"
      Height          =   375
      Index           =   0
      Left            =   4890
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   45
      Width           =   1125
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   375
      Index           =   2
      Left            =   7170
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   45
      Width           =   900
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   180
      Index           =   1
      Left            =   4800
      TabIndex        =   14
      Top             =   5280
      Width           =   972
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   255
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   5280
      Width           =   852
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   21
      Left            =   960
      TabIndex        =   12
      Top             =   5280
      Width           =   3768
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "6646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   22
      Left            =   5820
      TabIndex        =   11
      Top             =   5280
      Width           =   3072
      VariousPropertyBits=   27
      Caption         =   " "
      Size            =   "5419;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm100101_5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/15 改成Form2.0 ; lbl1(index)、txt1(index)
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
     frm100101_10.Show
     frm100101_10.Tag = StrTag ' StrTag  傳代理人代號
     frm100101_10.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_10.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 1
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
Case 2
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
Case 3
     fnCloseAllFrm100
'add by nickc 2005/05/30
Case 4
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100108_3.Show
     frm100108_3.Tag = txt1(4).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add By Sindy 2011/1/17
Case 5
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(Lbl1(28).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(Lbl1(28).Caption, 9) '當事人2
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 6
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(Lbl1(29).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(Lbl1(29).Caption, 9) '當事人3
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 7
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(Lbl1(30).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(Lbl1(30).Caption, 9) '當事人4
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
Case 8
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     If Trim(Lbl1(31).Caption) <> "" Then
         frm100101_11.Show
         frm100101_11.Tag = Left(Lbl1(31).Caption, 9) '當事人5
         frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
         frm100101_11.StrMenu
     End If
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'2011/1/17 End
'Added by Lydia 2016/11/23
Case 9 '各項指示
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
     frm12040159.SetParent "Q", Trim(Replace(txt1(4), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 10 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(4)
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
Dim strArr() As String, i As Integer, StrOkTxt(3) As String
'Modify By Sindy 2011/6/3
'Dim StrOk(23) As String
Dim StrOk(25) As String
'add by nickc 2006/07/12
ReDim strArr(TF_LC) As String
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
Dim tmp01 As String, tmp02 As String
'add by Toni 20080926 控制跨部門權限訊息
Dim strTit As String
Dim strMsg As String
Dim nResponse
'end by Toni 20080926

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

' 使用者沒有權限
'add by Toni 20080926 控制跨部門權限訊息 for 法務基本資料查詢
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

'Add By Sindy 2011/1/17
cmdok(5).Visible = False
cmdok(6).Visible = False
cmdok(7).Visible = False
cmdok(8).Visible = False
'2011/1/17 End

'欲搜尋的SQL字串
strSql = "SELECT * FROM LAWCASE WHERE LC01='" & Str01 & "' AND LC02='" & Str02 & "' AND LC03='" & Str03 & "' AND LC04='" & Str04 & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28 記錄此Form的查詢條件
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   If pub_QL04 <> "" Then InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2025/8/7
   'For i = 0 To 33
   For i = 0 To (TF_LC - 1) 'edit by nickc 2006/07/12 (T_LC - 1)
      Select Case i
      Case 8, 23, 28, 29, 31, 32
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
   Lbl1(51) = "" & adoRecordset.Fields("lc51") 'Added by Morgan 2018/4/11
   Lbl1(52) = "" & adoRecordset.Fields("lc52") 'Add by Amy 2018/08/15
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
intK = 34
'For i = 0 To 34
For i = 1 To TF_LC 'edit by nickc 2006/07/12 T_LC
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         txt1(4) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4) 'Add By Sindy 2013/1/31
    Case 8
         StrOk(1) = strArr(i)
    Case 9
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(2) = ""
         Else
             StrOk(2) = ChangeWStringToTString(strArr(i))
         End If

    Case 15
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(3) = strArr(i) + ""
             Else
                  StrOk(3) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
            'Add by Morgan 2004/1/14
            Lbl1(3).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/14
            'StrOk(3) = ""
            Lbl1(3).ForeColor = vbRed
             StrOk(3) = strArr(i)
         End If
         CheckOC
    Case 22
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Len(strArr(i)) = 9 Then
'              strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'         Else
'              strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'               If IsNull(adoRecordset.Fields(1)) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(4) = strArr(i) + ""
'                   Else
'                         StrOk(4) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                   StrOk(4) = strArr(i) + "  " + adoRecordset.Fields(1)
'               End If
'            Else
'               StrOk(4) = strArr(i) + "  " + adoRecordset.Fields(0)
'
'            End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetAgent Trim(strArr(i)), tmp02
         End If
         If tmp02 <> "" Then
            StrOk(4) = strArr(i) + "  " + tmp02
            
            'Add by Morgan 2004/1/14
            Lbl1(4).ForeColor = vbBlack
         Else
            StrOk(4) = ""
            'Add by Morgan 2004/1/14
            Lbl1(4).ForeColor = vbRed
            StrOk(4) = strArr(i)
         End If
         CheckOC
    Case 34
         'edit by nickc 2006/07/12
         'StrOk(5) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(5) = ""
         Else
             StrOk(5) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 13
         StrOk(6) = strArr(i)
    Case 26
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             tmp02 = ""
             If Trim(strArr(i)) <> "" Then
                ClsPDGetCustomer Trim(strArr(i)), tmp02
             End If
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             tmp02 = ""
             If Trim(strArr(i)) <> "" Then
                ClsPDGetAgent Trim(strArr(i)), tmp02
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                If IsNull(adoRecordset.Fields(1)) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(7) = strArr(i) + ""
'                    Else
'                        StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(1)
'                End If
'            Else
'                StrOk(7) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
         If tmp02 <> "" Then
            StrOk(7) = strArr(i) + "  " + tmp02
            'Add by Morgan 2004/1/14
            Lbl1(7).ForeColor = vbBlack
         Else
            StrOk(7) = ""
            'Add by Morgan 2004/1/14
            Lbl1(7).ForeColor = vbRed
            StrOk(7) = strArr(i)
         End If
         CheckOC
    Case 24
         StrOk(8) = strArr(i)
    Case 21
         StrOk(9) = strArr(i)
    Case 18
         StrOk(10) = strArr(i)
    Case 19
         StrOk(11) = strArr(i)
    Case 20
         StrOk(12) = strArr(i)
    Case 11
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         If Len(strArr(i)) = 9 Then
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'         Else
'              strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'         End If
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'             If IsNull(adoRecordset.Fields(0)) Then
'                  If IsNull(adoRecordset.Fields(1)) Then
'                     If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(13) = strArr(i) + ""
'                     Else
'                          StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(2)
'                     End If
'                  Else
'                     StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(1)
'                  End If
'             Else
'                  StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(0)
'             End If
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
        End If
        If tmp02 <> "" Then
            StrOk(13) = strArr(i) + "  " + tmp02
            'Add by Morgan 2004/1/14
            Lbl1(13).ForeColor = vbBlack
         Else
            'Modify by Morgan 2004/1/14
            'StrOk(13) = ""
            Lbl1(13).ForeColor = vbRed
             StrOk(13) = strArr(i)
         End If
         CheckOC
    Case 16
         StrOk(14) = strArr(i)
    Case 10
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOk(15) = strArr(i) + ""
             Else
                     StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(0)
             End If
         Else
             StrOk(15) = ""
         End If
         CheckOC
    Case 17
         StrOk(16) = strArr(i)
    Case 23
         StrOk(17) = strArr(i)
    Case 14
         StrOk(18) = strArr(i)
    Case 25
         StrOk(19) = strArr(i)
    Case 12
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             tmp02 = ""
             If Trim(strArr(i)) <> "" Then
                ClsPDGetCustomer Trim(strArr(i)), tmp02
             End If
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             tmp02 = ""
             If Trim(strArr(i)) <> "" Then
                ClsPDGetAgent Trim(strArr(i)), tmp02
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                If IsNull(adoRecordset.Fields(1)) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(20) = strArr(i) + ""
'                    Else
'                        StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(1)
'                End If
'            Else
'                StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
         If tmp02 <> "" Then
            StrOk(20) = strArr(i) + "  " + tmp02
            'Add by Morgan 2004/1/14
            Lbl1(20).ForeColor = vbBlack
         Else
            'Add by Morgan 2004/1/14
            'StrOk(20) = ""
            Lbl1(20).ForeColor = vbRed
            StrOk(20) = strArr(i)
         End If
         CheckOC
    Case 5
         StrOkTxt(0) = strArr(i)
    Case 6
         StrOkTxt(1) = strArr(i)
    Case 7
         StrOkTxt(2) = strArr(i)
    Case 27
         StrOkTxt(3) = strArr(i)
    Case 28
         'edit by nick 2004/10/05
         'StrOk(21) = GetPrjSalesNM(strArr(i)) & " " & strArr(29) & " " & strArr(30)
         StrOk(21) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(29))) & " " & Format(strArr(30), "##:##")
    Case 31
         'edit by nick 2004/10/05
         'StrOk(22) = GetPrjSalesNM(strArr(i)) & " " & strArr(32) & " " & strArr(33)
         StrOk(22) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(32))) & " " & Format(strArr(33), "##:##")
    Case 35 'D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
            tmp02 = ""
            If Trim(strArr(i)) <> "" Then
                ClsPDGetCustomer Trim(strArr(i)), tmp02
            End If
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            tmp02 = ""
            If Trim(strArr(i)) <> "" Then
                ClsPDGetAgent Trim(strArr(i)), tmp02
            End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'                If IsNull(adoRecordset.Fields(1)) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(23) = strArr(i) + ""
'                    Else
'                        StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(1)
'                End If
'            Else
'                StrOk(23) = strArr(i) + "  " + adoRecordset.Fields(0)
'            End If
        If tmp02 <> "" Then
            StrOk(23) = strArr(i) + "  " + tmp02
            'Add by Morgan 2004/1/14
            Lbl1(23).ForeColor = vbBlack
         Else
            StrOk(23) = ""
            'Add by Morgan 2004/1/14
            Lbl1(23).ForeColor = vbRed
            StrOk(23) = strArr(i)
         End If
         CheckOC
    'add by nickc 2006/07/12
    Case 36
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(24) = ""
         Else
             Lbl1(24) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 37
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(25) = strArr(i) + ""
            Else
               Lbl1(25) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(25) = ""
         End If
         CheckOC
    Case 38
         Lbl1(26) = strArr(i)
    'Add by Morgan 2008/8/4
    Case 42
         Lbl1(27) = PUB_GetContact(strArr(11), strArr(i))
    'Add By Sindy 2011/1/17
    Case 43
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdok(5).Visible = True
         End If
         If tmp02 <> "" Then
            Lbl1(28) = strArr(i) + "  " + tmp02
            Lbl1(28).ForeColor = vbBlack
         Else
            Lbl1(28) = strArr(i)
            Lbl1(28).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 44
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdok(6).Visible = True
         End If
         If tmp02 <> "" Then
            Lbl1(29) = strArr(i) + "  " + tmp02
            Lbl1(29).ForeColor = vbBlack
         Else
            Lbl1(29) = strArr(i)
            Lbl1(29).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 45
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdok(7).Visible = True
         End If
         If tmp02 <> "" Then
            Lbl1(30) = strArr(i) + "  " + tmp02
            Lbl1(30).ForeColor = vbBlack
         Else
            Lbl1(30) = strArr(i)
            Lbl1(30).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/1/17
    Case 46
         tmp02 = ""
         If Trim(strArr(i)) <> "" Then
            ClsPDGetCustomer Trim(strArr(i)), tmp02
            cmdok(8).Visible = True
         End If
         If tmp02 <> "" Then
            Lbl1(31) = strArr(i) + "  " + tmp02
            Lbl1(31).ForeColor = vbBlack
         Else
            Lbl1(31) = strArr(i)
            Lbl1(31).ForeColor = vbRed
         End If
         CheckOC
    'Add By Sindy 2011/6/3
    Case 47
         StrOk(24) = strArr(i)
    'Add By Sindy 2013/12/13
    Case 48
         StrOk(25) = strArr(i)
    'Add By Sindy 2016/11/24
    Case 49
      Lbl1(49) = strArr(i)
    Case 50
      Combo3(1).ListIndex = Val(strArr(i))
    '2016/11/24 END
    Case Else
    End Select
    DoEvents
Next i
For i = 0 To 25 '23
   If i = 24 Then 'Add By Sindy 2011/6/3
      Lbl1(32) = StrOk(i)
   ElseIf i = 25 Then 'Add By Sindy 2013/12/13
      Lbl1(33) = StrOk(i)
   Else
      If i <> 0 Then 'Add By Sindy 2013/1/31 +if
         Lbl1(i) = StrOk(i)
      End If
   End If
Next i
txt1(0) = StrOkTxt(0)
txt1(1) = StrOkTxt(1)
txt1(2) = StrOkTxt(2)
txt1(3) = StrOkTxt(3)
'傳參數　　　代理人，申請人
'StrToSystem(ArrTemp) = StrArr(22) + "=" + StrArr(11) + "=    "
StrTag = strArr(22)
StrTag1 = strArr(11)
'add by nickc 2005/05/30  檢查有無分割或相關卷號
     cmdok(4).Visible = ChkDataByCR(txt1(4).Text)

   'Added by Lydia 2020/03/30 事務所合併日起取消( J:智權公司 空白:系統預設)的標題
   'Moidified by Lydia 2020/05/29 從Form_Load移過來; +非ACS案才取消。
   If strSrvDate(1) >= 事務所合併日 And Str01 <> "ACS" Then
       Label6(1).Visible = False
       Label8(1).Visible = False
       Lbl1(33).Visible = False
   End If
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

'edit by nickc 2005/05/30 改成與我們現在的共用相同
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(4), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
bolToEndByNick = False

SSTab1.Tab = 0 'Added by Lydia 2016/11/23

   MoveFormToCenter Me
If bolFNation = False Then
    Label18(1).Visible = False
    Lbl1(4).Visible = False
    cmdok(0).Value = False
    Label20(2).Visible = False
    Lbl1(10).Visible = False
    Label20(3).Visible = False
    Lbl1(11).Visible = False
    Label20(1).Visible = False
    Lbl1(12).Visible = False
    Label21.Visible = False
    Lbl1(9).Visible = False
    Label9.Visible = False
    Lbl1(8).Visible = False
    Label8(0).Visible = False
    Lbl1(19).Visible = False
    Label8(2).Visible = False
    Label11.Visible = False
    Lbl1(20).Visible = False
    Label7.Visible = False
    Lbl1(7).Visible = False
End If
'92.04.16 nick
cmdState = -1
   
    'Added by Lydia 2020/05/05 各項指示：顯示按鈕
    If strSrvDate(1) >= 各項指示啟用日 Then
       cmdok(9).Visible = True
    Else
       cmdok(9).Visible = False
    End If
    'end 2020/05/05
End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_5 = Nothing
End Sub

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub

