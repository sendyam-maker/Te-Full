VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_A 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料（著作權）"
   ClientHeight    =   6090
   ClientLeft      =   470
   ClientTop       =   980
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   11
      Left            =   4790
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "各項備註"
      Height          =   350
      Index           =   10
      Left            =   120
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "已設定代表圖"
      Height          =   350
      Index           =   9
      Left            =   1010
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   30
      Width           =   1245
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申5"
      Height          =   350
      Index           =   8
      Left            =   3490
      TabIndex        =   6
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申4"
      Height          =   350
      Index           =   7
      Left            =   3080
      TabIndex        =   7
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申3"
      Height          =   350
      Index           =   6
      Left            =   2670
      TabIndex        =   8
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申2"
      Height          =   350
      Index           =   5
      Left            =   2260
      TabIndex        =   9
      Top             =   30
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   350
      Index           =   4
      Left            =   3900
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   30
      Visible         =   0   'False
      Width           =   885
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   5385
      Left            =   60
      TabIndex        =   12
      Top             =   420
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   9507
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料1"
      TabPicture(0)   =   "frm100101_A.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txt1(15)"
      Tab(0).Control(1)=   "txt1(16)"
      Tab(0).Control(2)=   "txt1(14)"
      Tab(0).Control(3)=   "txt1(11)"
      Tab(0).Control(4)=   "txt1(0)"
      Tab(0).Control(5)=   "txt1(1)"
      Tab(0).Control(6)=   "txt1(2)"
      Tab(0).Control(7)=   "txt1(3)"
      Tab(0).Control(8)=   "txt1(4)"
      Tab(0).Control(9)=   "Label112"
      Tab(0).Control(10)=   "Label113"
      Tab(0).Control(11)=   "lbl1(85)"
      Tab(0).Control(12)=   "lbl1(11)"
      Tab(0).Control(13)=   "Label89"
      Tab(0).Control(14)=   "Label10"
      Tab(0).Control(15)=   "Label66"
      Tab(0).Control(16)=   "Label69"
      Tab(0).Control(17)=   "Label60"
      Tab(0).Control(18)=   "Label72"
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(20)=   "lbl1(34)"
      Tab(0).Control(21)=   "lbl1(33)"
      Tab(0).Control(22)=   "Label12"
      Tab(0).Control(23)=   "lbl1(6)"
      Tab(0).Control(24)=   "lbl1(8)"
      Tab(0).Control(25)=   "lbl1(17)"
      Tab(0).Control(26)=   "lbl1(16)"
      Tab(0).Control(27)=   "lbl1(13)"
      Tab(0).Control(28)=   "lbl1(12)"
      Tab(0).Control(29)=   "lbl1(10)"
      Tab(0).Control(30)=   "lbl1(9)"
      Tab(0).Control(31)=   "lbl1(7)"
      Tab(0).Control(32)=   "lbl1(5)"
      Tab(0).Control(33)=   "lbl1(4)"
      Tab(0).Control(34)=   "lbl1(1)"
      Tab(0).Control(35)=   "Label73"
      Tab(0).Control(36)=   "Label71"
      Tab(0).Control(37)=   "Label63"
      Tab(0).Control(38)=   "Label56"
      Tab(0).Control(39)=   "Label53"
      Tab(0).Control(40)=   "Label68"
      Tab(0).Control(41)=   "Label67"
      Tab(0).Control(42)=   "Label65"
      Tab(0).Control(43)=   "Label64"
      Tab(0).Control(44)=   "Label59"
      Tab(0).Control(45)=   "Label52"
      Tab(0).Control(46)=   "Label48"
      Tab(0).Control(47)=   "Label47"
      Tab(0).Control(48)=   "Label46"
      Tab(0).Control(49)=   "Label45"
      Tab(0).Control(50)=   "Label76"
      Tab(0).Control(51)=   "Label17"
      Tab(0).Control(52)=   "Label88"
      Tab(0).Control(53)=   "Label13"
      Tab(0).ControlCount=   54
      TabCaption(1)   =   "基本資料2"
      TabPicture(1)   =   "frm100101_A.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label110"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label102"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label22"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label92"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label91"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label84"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label50"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label62"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label75"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label77"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label5"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "lbl1(19)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lbl1(20)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "lbl1(21)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "Label9"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "lbl1(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lbl1(2)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label4(0)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Label2(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Label4(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "Label2(1)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "Label3"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Label6"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "lbl1(39)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "lbl1(18)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "Label1"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "lbl1(87)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "lbl1(86)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).Control(28)=   "Label23"
      Tab(1).Control(28).Enabled=   0   'False
      Tab(1).Control(29)=   "lbl1(83)"
      Tab(1).Control(29).Enabled=   0   'False
      Tab(1).Control(30)=   "lbl1(80)"
      Tab(1).Control(30).Enabled=   0   'False
      Tab(1).Control(31)=   "txt1(10)"
      Tab(1).Control(31).Enabled=   0   'False
      Tab(1).Control(32)=   "txt1(5)"
      Tab(1).Control(32).Enabled=   0   'False
      Tab(1).Control(33)=   "txt1(6)"
      Tab(1).Control(33).Enabled=   0   'False
      Tab(1).Control(34)=   "txt1(7)"
      Tab(1).Control(34).Enabled=   0   'False
      Tab(1).Control(35)=   "txt1(8)"
      Tab(1).Control(35).Enabled=   0   'False
      Tab(1).Control(36)=   "txt1(9)"
      Tab(1).Control(36).Enabled=   0   'False
      Tab(1).Control(37)=   "txt1(12)"
      Tab(1).Control(37).Enabled=   0   'False
      Tab(1).Control(38)=   "Label14(0)"
      Tab(1).Control(38).Enabled=   0   'False
      Tab(1).Control(39)=   "Label14(1)"
      Tab(1).Control(39).Enabled=   0   'False
      Tab(1).ControlCount=   40
      TabCaption(2)   =   "代理人相關資料"
      TabPicture(2)   =   "frm100101_A.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txt1(13)"
      Tab(2).Control(1)=   "lbl1(40)"
      Tab(2).Control(2)=   "lbl1(25)"
      Tab(2).Control(3)=   "lbl1(84)"
      Tab(2).Control(4)=   "lbl1(35)"
      Tab(2).Control(5)=   "lbl1(29)"
      Tab(2).Control(6)=   "lbl1(28)"
      Tab(2).Control(7)=   "lbl1(27)"
      Tab(2).Control(8)=   "lbl1(26)"
      Tab(2).Control(9)=   "lbl1(24)"
      Tab(2).Control(10)=   "lbl1(23)"
      Tab(2).Control(11)=   "lbl1(22)"
      Tab(2).Control(12)=   "Label8"
      Tab(2).Control(13)=   "Label21"
      Tab(2).Control(14)=   "Label19"
      Tab(2).Control(15)=   "Label29"
      Tab(2).Control(16)=   "Label7"
      Tab(2).Control(17)=   "Label15"
      Tab(2).Control(18)=   "Label16"
      Tab(2).Control(19)=   "Label18"
      Tab(2).Control(20)=   "Label20"
      Tab(2).Control(21)=   "Label24"
      Tab(2).Control(22)=   "Label28"
      Tab(2).Control(23)=   "Label31"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "銷卷資料"
      TabPicture(3)   =   "frm100101_A.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "lbl1(38)"
      Tab(3).Control(1)=   "lbl1(37)"
      Tab(3).Control(2)=   "lbl1(36)"
      Tab(3).Control(3)=   "lbl1(32)"
      Tab(3).Control(4)=   "Label78"
      Tab(3).Control(5)=   "Label79"
      Tab(3).Control(6)=   "Label80"
      Tab(3).Control(7)=   "Label81"
      Tab(3).ControlCount=   8
      Begin MSForms.Label Label14 
         Height          =   255
         Index           =   1
         Left            =   930
         TabIndex        =   142
         Top             =   1131
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
      Begin MSForms.Label Label14 
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   141
         Top             =   864
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
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   15
         Left            =   -74010
         TabIndex        =   129
         Top             =   1990
         Width           =   1635
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2884;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   16
         Left            =   -70410
         TabIndex        =   128
         Top             =   1990
         Width           =   1635
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2884;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   14
         Left            =   -73896
         TabIndex        =   127
         Top             =   330
         Width           =   1635
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "2884;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   12
         Left            =   1035
         TabIndex        =   101
         Top             =   4815
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   480
         Index           =   11
         Left            =   -73950
         TabIndex        =   99
         Top             =   4560
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;847"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   9
         Left            =   1035
         TabIndex        =   22
         Top             =   3885
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   405
         Index           =   8
         Left            =   6570
         TabIndex        =   21
         Top             =   4560
         Visible         =   0   'False
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12197;714"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   7
         Left            =   1035
         TabIndex        =   20
         Top             =   3450
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   6
         Left            =   1035
         TabIndex        =   19
         Top             =   2970
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   5
         Left            =   1035
         TabIndex        =   18
         Top             =   2490
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   0
         Left            =   -73560
         TabIndex        =   13
         Top             =   653
         Width           =   6480
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11430;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   1
         Left            =   -73560
         TabIndex        =   14
         Top             =   1006
         Width           =   6480
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11430;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   330
         Index           =   2
         Left            =   -73560
         TabIndex        =   15
         Top             =   1359
         Width           =   6480
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11430;582"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   3
         Left            =   -73800
         TabIndex        =   16
         Top             =   2869
         Width           =   2430
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "4286;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   10
         Left            =   1035
         TabIndex        =   23
         Top             =   4380
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;811"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   4
         Left            =   -70320
         TabIndex        =   17
         Top             =   4282
         Width           =   3000
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5292;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1005
         Index           =   13
         Left            =   -74880
         TabIndex        =   24
         Top             =   3930
         Width           =   7800
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13758;1773"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   80
         Left            =   4500
         TabIndex        =   138
         Top             =   1665
         Width           =   255
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "450;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   83
         Left            =   8025
         TabIndex        =   137
         Top             =   1665
         Width           =   240
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "423;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   255
         Left            =   120
         TabIndex        =   136
         Top             =   2205
         Width           =   1470
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   86
         Left            =   1650
         TabIndex        =   135
         Top             =   2205
         Width           =   3360
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5927;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   87
         Left            =   6585
         TabIndex        =   133
         Top             =   2205
         Width           =   2070
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3651;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   255
         Left            =   -73260
         TabIndex        =   132
         Top             =   5070
         Width           =   2115
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   255
         Left            =   -74880
         TabIndex        =   131
         Top             =   5070
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   85
         Left            =   -73590
         TabIndex        =   130
         Top             =   5070
         Width           =   225
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "397;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   40
         Left            =   -73590
         TabIndex        =   125
         Top             =   1710
         Width           =   7245
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12779;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -71580
         TabIndex        =   81
         Top             =   690
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   84
         Left            =   -72930
         TabIndex        =   116
         Top             =   1020
         Width           =   6525
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11509;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   35
         Left            =   -73290
         TabIndex        =   103
         Top             =   3210
         Width           =   6945
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12250;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   29
         Left            =   -73800
         TabIndex        =   85
         Top             =   2880
         Width           =   7455
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13150;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   28
         Left            =   -73800
         TabIndex        =   84
         Top             =   2520
         Width           =   7485
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13203;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   27
         Left            =   -68640
         TabIndex        =   83
         Top             =   690
         Width           =   400
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "706;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73590
         TabIndex        =   82
         Top             =   2130
         Width           =   7245
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12779;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -73920
         TabIndex        =   80
         Top             =   1350
         Width           =   7575
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13361;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73920
         TabIndex        =   79
         Top             =   690
         Width           =   1665
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2937;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73920
         TabIndex        =   78
         Top             =   360
         Width           =   7425
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "13097;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門："
         Height          =   255
         Left            =   -74880
         TabIndex        =   126
         Top             =   1710
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -70320
         TabIndex        =   70
         Top             =   4026
         Width           =   390
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "688;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label89 
         AutoSize        =   -1  'True
         Caption         =   "閉卷原因："
         Height          =   255
         Left            =   -71280
         TabIndex        =   124
         Top             =   4305
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷：           （Y：閉卷）"
         Height          =   255
         Left            =   -71280
         TabIndex        =   123
         Top             =   4026
         Width           =   2415
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "發證日："
         Height          =   255
         Left            =   -71340
         TabIndex        =   122
         Top             =   2313
         Width           =   1005
      End
      Begin VB.Label Label69 
         AutoSize        =   -1  'True
         Caption         =   "申請案號："
         Height          =   180
         Left            =   -71280
         TabIndex        =   121
         Top             =   2050
         Width           =   900
      End
      Begin VB.Label Label60 
         AutoSize        =   -1  'True
         Caption         =   "首次發表日："
         Height          =   180
         Left            =   -71280
         TabIndex        =   120
         Top             =   2628
         Width           =   1080
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "擁有狀態："
         Height          =   255
         Left            =   -71280
         TabIndex        =   119
         Top             =   3470
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "主管機關："
         Height          =   255
         Left            =   120
         TabIndex        =   118
         Top             =   1398
         Width           =   930
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   18
         Left            =   1110
         TabIndex        =   117
         Top             =   1398
         Width           =   2865
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5054;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   255
         Left            =   -74880
         TabIndex        =   115
         Top             =   1020
         Width           =   1860
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   39
         Left            =   5850
         TabIndex        =   114
         Top             =   1932
         Width           =   2760
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4868;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   255
         Left            =   5085
         TabIndex        =   113
         Top             =   1932
         Width           =   750
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   38
         Left            =   -73590
         TabIndex        =   112
         Top             =   1200
         Width           =   5355
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "9446;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   37
         Left            =   -73590
         TabIndex        =   111
         Top             =   900
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
         Index           =   36
         Left            =   -73590
         TabIndex        =   110
         Top             =   630
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
         Index           =   32
         Left            =   -73590
         TabIndex        =   109
         Top             =   360
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
         Left            =   -74850
         TabIndex        =   108
         Top             =   1200
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   255
         Left            =   -74850
         TabIndex        =   107
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   255
         Left            =   -74850
         TabIndex        =   106
         Top             =   630
         Width           =   1080
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   255
         Left            =   -74850
         TabIndex        =   105
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   255
         Left            =   -74880
         TabIndex        =   104
         Top             =   3210
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註："
         Height          =   180
         Left            =   120
         TabIndex        =   102
         Top             =   4815
         Width           =   900
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   100
         Top             =   4590
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "申請人4："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   98
         Top             =   864
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人5："
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   1131
         Width           =   840
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "申請人2："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   96
         Top             =   330
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "申請人3："
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   95
         Top             =   597
         Width           =   840
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   94
         Top             =   330
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
         Index           =   3
         Left            =   960
         TabIndex        =   93
         Top             =   597
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
         Index           =   34
         Left            =   -72750
         TabIndex        =   91
         Top             =   4026
         Width           =   795
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "1402;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   33
         Left            =   -73800
         TabIndex        =   90
         Top             =   4026
         Width           =   795
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "中-lblFM2"
         Size            =   "1402;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   3
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "專用期間："
         Height          =   255
         Left            =   -74880
         TabIndex        =   89
         Top             =   4026
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   6
         Left            =   -73800
         TabIndex        =   65
         Top             =   3192
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   -73800
         TabIndex        =   67
         Top             =   3748
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
         Caption         =   "登記項目："
         Height          =   180
         Left            =   120
         TabIndex        =   88
         Top             =   4095
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   1455
         TabIndex        =   77
         Top             =   1932
         Width           =   3450
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6085;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   1110
         TabIndex        =   76
         Top             =   1665
         Width           =   1785
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3149;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   5100
         TabIndex        =   75
         Top             =   1398
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
         ParagraphAlign  =   2
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   17
         Left            =   -70125
         TabIndex        =   74
         Top             =   2591
         Width           =   1500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2646;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -70230
         TabIndex        =   73
         Top             =   2313
         Width           =   1500
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2646;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -70290
         TabIndex        =   72
         Top             =   353
         Width           =   1800
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3175;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73800
         TabIndex        =   71
         Top             =   4305
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   -68625
         TabIndex        =   69
         Top             =   3192
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   -70320
         TabIndex        =   68
         Top             =   3470
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
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   -73800
         TabIndex        =   66
         Top             =   3465
         Width           =   1380
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2434;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   -73320
         TabIndex        =   64
         Top             =   2591
         Width           =   1845
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3254;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   -74040
         TabIndex        =   63
         Top             =   2313
         Width           =   2535
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4471;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   -73980
         TabIndex        =   62
         Top             =   1710
         Width           =   6930
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12224;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "著作權"
         Height          =   180
         Left            =   120
         TabIndex        =   61
         Top             =   3885
         Width           =   540
      End
      Begin VB.Label Label77 
         AutoSize        =   -1  'True
         Caption         =   "軟件說明："
         Height          =   180
         Left            =   120
         TabIndex        =   60
         Top             =   4380
         Width           =   900
      End
      Begin VB.Label Label75 
         AutoSize        =   -1  'True
         Caption         =   "地址："
         Height          =   180
         Left            =   120
         TabIndex        =   59
         Top             =   3450
         Width           =   540
      End
      Begin VB.Label Label62 
         AutoSize        =   -1  'True
         Caption         =   "著作人："
         Height          =   180
         Left            =   120
         TabIndex        =   58
         Top             =   2490
         Width           =   720
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "代表人："
         Height          =   180
         Left            =   120
         TabIndex        =   57
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "（1.單獨擁有  2.共同擁有）"
         Height          =   255
         Left            =   -69840
         TabIndex        =   56
         Top             =   3470
         Width           =   2160
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "（1.單獨開發  2.合作開發  3.委託開發  4.下達任務開發）"
         Height          =   255
         Left            =   -73395
         TabIndex        =   55
         Top             =   3748
         Width           =   4410
      End
      Begin VB.Label Label63 
         AutoSize        =   -1  'True
         Caption         =   "開發形式："
         Height          =   255
         Left            =   -74880
         TabIndex        =   54
         Top             =   3748
         Width           =   900
      End
      Begin VB.Label Label56 
         AutoSize        =   -1  'True
         Caption         =   "著作完成日："
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   3465
         Width           =   1080
      End
      Begin VB.Label Label53 
         Caption         =   "（1.原創軟件  2.修改本  3.合成軟件  4.翻譯本）"
         Height          =   255
         Left            =   -73455
         TabIndex        =   52
         Top             =   3192
         Width           =   3735
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "登記號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   51
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Label67 
         AutoSize        =   -1  'True
         Caption         =   "申請日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   50
         Top             =   2050
         Width           =   720
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "作品類型："
         Height          =   255
         Left            =   -74880
         TabIndex        =   49
         Top             =   3192
         Width           =   900
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "申請人1："
         Height          =   180
         Left            =   -74880
         TabIndex        =   48
         Top             =   1749
         Width           =   810
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         Height          =   180
         Left            =   -71280
         TabIndex        =   47
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label52 
         AutoSize        =   -1  'True
         Caption         =   "註冊號數/證書號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   46
         Top             =   2628
         Width           =   1485
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   45
         Top             =   1359
         Width           =   1200
      End
      Begin VB.Label Label47 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   44
         Top             =   1006
         Width           =   1200
      End
      Begin VB.Label Label46 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   43
         Top             =   653
         Width           =   1200
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   42
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label76 
         AutoSize        =   -1  'True
         Caption         =   "作品種類："
         Height          =   180
         Left            =   -74880
         TabIndex        =   41
         Top             =   2929
         Width           =   900
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：           （1.中文  2.英文  3.日文）"
         Height          =   255
         Left            =   4140
         TabIndex        =   40
         Top             =   1398
         Width           =   3450
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   255
         Left            =   120
         TabIndex        =   39
         Top             =   1665
         Width           =   930
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   1932
         Width           =   1290
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "是否發行：           （Y：發行）"
         Height          =   255
         Left            =   -69555
         TabIndex        =   37
         Top             =   3192
         Width           =   2415
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   255
         Left            =   -74880
         TabIndex        =   36
         Top             =   4305
         Width           =   900
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   35
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   255
         Left            =   -74880
         TabIndex        =   34
         Top             =   690
         Width           =   900
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "折扣：            %"
         Height          =   255
         Left            =   -72210
         TabIndex        =   33
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   32
         Top             =   1350
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   255
         Left            =   -74880
         TabIndex        =   31
         Top             =   2130
         Width           =   1260
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：            （Y：印）"
         Height          =   255
         Left            =   -70440
         TabIndex        =   30
         Top             =   690
         Width           =   3105
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   29
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   255
         Left            =   -74880
         TabIndex        =   28
         Top             =   2880
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   27
         Top             =   3660
         Width           =   1080
      End
      Begin VB.Label Label13 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         Caption         =   "－"
         Height          =   165
         Left            =   -72960
         TabIndex        =   92
         Top             =   4050
         Width           =   90
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   255
         Left            =   5100
         TabIndex        =   134
         Top             =   2205
         Width           =   1470
      End
      Begin VB.Label Label102 
         AutoSize        =   -1  'True
         Caption         =   "以電子郵件通知：        (Y：是  D：僅D/N）"
         Height          =   255
         Left            =   3030
         TabIndex        =   139
         Top             =   1665
         Width           =   3405
      End
      Begin VB.Label Label110 
         AutoSize        =   -1  'True
         Caption         =   "EMail同時寄紙本：       (Y:是)"
         Height          =   255
         Left            =   6495
         TabIndex        =   140
         Top             =   1665
         Width           =   2310
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   3
      Left            =   5595
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   30
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   7285
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   30
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人"
      Height          =   350
      Index           =   0
      Left            =   6440
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   30
      Width           =   840
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   2
      Left            =   8130
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   31
      Left            =   4650
      TabIndex        =   87
      Top             =   5850
      Width           =   3315
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "5847;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   30
      Left            =   1005
      TabIndex        =   86
      Top             =   5850
      Width           =   2625
      BackColor       =   16777215
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "4630;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   255
      Left            =   3690
      TabIndex        =   26
      Top             =   5850
      Width           =   900
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   255
      Left            =   45
      TabIndex        =   25
      Top             =   5850
      Width           =   900
   End
End
Attribute VB_Name = "frm100101_A"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/20 改成Form2.0 ; lbl1(index)、txt1(index)、Label14(index)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/26 日期欄已修改
Option Explicit

Dim StrTag As String, StrTag1 As String, intK As Integer
'92.04.16 nick 紀錄作用按鍵
Public cmdState As Integer
'Add By Sindy 2010/02/04
Dim StrTag2 As String, StrTag3 As String, StrTag4 As String, StrTag5 As String
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
     frm100108_3.Tag = txt1(14).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add By Sindy 2010/02/04
Case 5
     cmdState = -1
     Me.Enabled = False
     If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
     End If
     Screen.MousePointer = vbHourglass
     frm100101_11.Show
     frm100101_11.Tag = StrTag2
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
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
     frm100101_11.Show
     frm100101_11.Tag = StrTag3
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
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
     frm100101_11.Show
     frm100101_11.Tag = StrTag4
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
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
     frm100101_11.Show
     frm100101_11.Tag = StrTag5
     frm100101_11.m_pub_QL05 = pub_QL05 'Add By Sindy 2025/8/28
     frm100101_11.StrMenu
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'2010/02/04 End
'Add By Sindy 103/2/18
Case 9
    frmPic001.oCP01 = SystemNumber(txt1(14), 1)
    frmPic001.oCP02 = SystemNumber(txt1(14), 2)
    frmPic001.oCP03 = SystemNumber(txt1(14), 3)
    frmPic001.oCP04 = SystemNumber(txt1(14), 4)
    frmPic001.StrMenu
    frmPic001.CanScan
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
    frmPic001.Show vbModal
    'Modify by Amy 2018/07/16  改寫至function
'    strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(txt1(14), 1) & "' and ibf02='" & SystemNumber(txt1(14), 2) & "' and ibf03='" & SystemNumber(txt1(14), 3) & "' and ibf04='" & SystemNumber(txt1(14), 4) & "' and ibf05='1'"
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    If ChkImgByteFile(SystemNumber(txt1(14), 1), SystemNumber(txt1(14), 2), SystemNumber(txt1(14), 3), SystemNumber(txt1(14), 4)) = True Then
        'Modified by Lydia 2021/12/20 拿掉快速鍵
        cmdok(9).Caption = "已設定代表圖"
        cmdok(9).BackColor = &HC0FFC0
    Else
        'Modified by Lydia 2021/12/20 拿掉快速鍵
        cmdok(9).Caption = "未設定代表圖"
        cmdok(9).BackColor = &HC0C0FF
    End If
'    CheckOC2
    'end 2018/07/16
'103/2/18 End
'Added by Lydia 2016/11/23
Case 10 '各項指示
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
     frm12040159.SetParent "Q", Trim(Replace(txt1(14), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 11 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(14)
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
'Modify By Cheng 2002/04/22
'Dim StrArr(62) As String, i As Integer, StrOk(32) As String, StrOkTxt(13) As String
'edit by nickc 2006/07/12
'Dim strArr(T_SP) As String, i As Integer, StrOk(35) As String, StrOkTxt(13) As String
Dim strArr() As String, i As Integer, StrOk(35) As String, StrOkTxt(13) As String
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

'add by Toni 20080926 控制跨部門權限 for 著作權
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
      'DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
   Next i
Else
   If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/7
   ShowNoData
   Screen.MousePointer = vbDefault
       '920416 nick
     'Me.Hide
     tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Exit Sub
End If

Label14(0).Caption = ""
If Not IsNull(adoRecordset.Fields("SP65")) Then
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.GetCusCAJnam(adoRecordset.Fields("SP65"), strExc(1), strExc(2), strExc(3)) Then
   If ClsLawGetCusCAJnam(adoRecordset.Fields("SP65"), strExc(1), strExc(2), strExc(3)) Then
      Label14(0).Caption = adoRecordset.Fields("SP65") & " "
      If strExc(1) = "" Then
         If strExc(2) = "" Then
            Label14(0).Caption = Label14(0).Caption & strExc(3)
         Else
            Label14(0).Caption = Label14(0).Caption & strExc(2)
         End If
      Else
         Label14(0).Caption = Label14(0).Caption & strExc(1)
      End If
   End If
End If

Label14(1).Caption = ""
If Not IsNull(adoRecordset.Fields("SP66")) Then
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.GetCusCAJnam(adoRecordset.Fields("SP66"), strExc(1), strExc(2), strExc(3)) Then
   If ClsLawGetCusCAJnam(adoRecordset.Fields("SP66"), strExc(1), strExc(2), strExc(3)) Then
      Label14(1).Caption = adoRecordset.Fields("SP66") & " "
      If strExc(1) = "" Then
         If strExc(2) = "" Then
            Label14(1).Caption = Label14(1).Caption & strExc(3)
         Else
            Label14(1).Caption = Label14(1).Caption & strExc(2)
         End If
      Else
         Label14(1).Caption = Label14(1).Caption & strExc(1)
      End If
   End If
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
         txt1(14) = StrOk(0) 'Add By Sindy 2013/1/31
         
         'Add By Sindy 103/2/18 檢查有無代表圖
         'Modify by Amy 2018/07/16  改寫至function
'         strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & strArr(1) & "' and ibf02='" & strArr(2) & "' and ibf03='" & strArr(3) & "' and ibf04='" & strArr(4) & "' and ibf05='1'"
'         CheckOC2
'         adoRecordset1.CursorLocation = adUseClient
'         adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
         If ChkImgByteFile(strArr(1), strArr(2), strArr(3), strArr(4)) = True Then
             'Modified by Lydia 2021/12/20 拿掉快速鍵
             cmdok(9).Caption = "已設定代表圖"
             cmdok(9).BackColor = &HC0FFC0
         Else
             'Modified by Lydia 2021/12/20 拿掉快速鍵
             cmdok(9).Caption = "未設定代表圖"
             cmdok(9).BackColor = &HC0C0FF
         End If
'         CheckOC2
         'end 2016/07/16
         '103/2/18 End
    Case 8
         'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
         'strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(strArr(i))
         strSql = "SELECT CU04,cu05,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(strArr(i))
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
                  StrOkTxt(12) = ""
             Else
                  StrOkTxt(12) = adoRecordset.Fields(3)
             End If
             'Add by Morgan 2004/1/6
             Lbl1(1).ForeColor = vbBlack
         Else
             StrOk(1) = ""
             'Add by Morgan 2004/1/16
             Lbl1(1).ForeColor = vbRed
             StrOk(1) = strArr(i)
             
             StrOkTxt(12) = ""
         End If
         CheckOC
    Case 58
         If strArr(i) <> "" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(strArr(i))
'            CheckOC
'            adoRecordset.CursorLocation = adUseClient
'            adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                If IsNull(adoRecordset.Fields(0)) Then
'                     If IsNull(adoRecordset.Fields(1)) Then
'                        If IsNull(adoRecordset.Fields(2)) Then
'                             StrOk(2) = strArr(i) + ""
'                        Else
'                             StrOk(2) = strArr(i) + "  " + adoRecordset.Fields(2)
'                        End If
'                     Else
'                        StrOk(2) = strArr(i) + "  " + adoRecordset.Fields(1)
'                     End If
'                Else
'                     StrOk(2) = strArr(i) + "  " + adoRecordset.Fields(0)
'                End If
            StrOk(2) = strArr(i) + "  " + GetCustName(strArr(i))
            If StrOk(2) <> strArr(i) Then
                'Add by Morgan 2004/1/16
                Lbl1(2).ForeColor = vbBlack
            Else
                StrOk(2) = ""
                'Add by Morgan 2004/1/16
                Lbl1(2).ForeColor = vbBlack
                StrOk(2) = strArr(i)
            End If
            CheckOC
         End If
    Case 59
         If strArr(i) <> "" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'            strSQL = "SELECT CU04,cu05||' '||cu88||' '||cu89||' '||cu90,CU06,CU79 FROM CUSTOMER WHERE " & ChgCustomer(strArr(i))
'            CheckOC
'            adoRecordset.CursorLocation = adUseClient
'            adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                If IsNull(adoRecordset.Fields(0)) Then
'                     If IsNull(adoRecordset.Fields(1)) Then
'                        If IsNull(adoRecordset.Fields(2)) Then
'                             StrOk(3) = strArr(i) + ""
'                        Else
'                             StrOk(3) = strArr(i) + "  " + adoRecordset.Fields(2)
'                        End If
'                     Else
'                        StrOk(3) = strArr(i) + "  " + adoRecordset.Fields(1)
'                     End If
'                Else
'                     StrOk(3) = strArr(i) + "  " + adoRecordset.Fields(0)
'                End If
            StrOk(3) = strArr(i) + "  " + GetCustName(strArr(i))
            If StrOk(3) <> strArr(i) Then
                'Add by Morgan 2004/1/16
                Lbl1(3).ForeColor = vbBlack
            Else
                StrOk(3) = ""
                'Add by Morgan 2004/1/16
                Lbl1(3).ForeColor = vbRed
                StrOk(3) = strArr(i)
            End If
            CheckOC
         End If
    Case 13
         StrOk(4) = strArr(i)
    Case 14
         StrOk(5) = strArr(i)
    Case 38
         StrOk(6) = strArr(i)
    Case 39
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(7) = ""
         Else
            'Modify By Sindy 2012/10/1
            If Right(strArr(i), 2) = "00" Then
               StrOk(7) = Val(Left(strArr(i), Len(strArr(i)) - 2)) - 191100
            Else
            '2012/10/1 End
               StrOk(7) = ChangeWStringToTString(strArr(i))
            End If
         End If
    Case 63
         StrOk(8) = strArr(i)
    Case 47
         StrOk(9) = strArr(i)
    Case 48
         StrOk(10) = strArr(i)
    Case 15
         StrOk(11) = strArr(i)
    Case 16
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(12) = ""
         Else
             StrOk(12) = ChangeWStringToTString(strArr(i))
         End If
    Case 9
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(13) = strArr(i) + ""
              Else
                  StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
         Else
              StrOk(13) = ""
         End If
         CheckOC
    Case 10
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(14) = ""
         Else
             StrOk(14) = ChangeWStringToTString(strArr(i))
         End If
         txt1(15) = StrOk(14) 'Add By Sindy 2013/1/31
    Case 11
         StrOk(15) = strArr(i)
         txt1(16) = StrOk(15) 'Add By Sindy 2013/1/31
    Case 12
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(16) = ""
         Else
             StrOk(16) = ChangeWStringToTString(strArr(i))
         End If
    Case 40
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(17) = ""
         Else
             StrOk(17) = ChangeWStringToTString(strArr(i))
         End If
    Case 51
         StrOk(18) = strArr(i)
    Case 34
         StrOk(19) = strArr(i)
    Case 28
         StrOk(20) = strArr(i)
    Case 29
         StrOk(21) = strArr(i)
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
'               'Add By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(22) = strArr(i) + ""
'                   Else
'                         StrOk(22) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                  'Add By Cheng 2002/07/08
''                   StrOk(22) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   StrOk(22) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))
'               End If
'            Else
'               'Add By Cheng 2002/07/08
''               StrOk(22) = StrArr(i) + "  " + adoRecordset.Fields(0)
'               StrOk(22) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))
'
'            End If
            If IsNull(adoRecordset.Fields("FA05")) = False Then
               StrOk(22) = strArr(i) + "  " + adoRecordset.Fields("FA05")
            ElseIf IsNull(adoRecordset.Fields("FA04")) = False Then
               StrOk(22) = strArr(i) + "  " + adoRecordset.Fields("FA04")
            ElseIf IsNull(adoRecordset.Fields("FA06")) = False Then
               StrOk(22) = strArr(i) + "  " + adoRecordset.Fields("FA06")
            End If
                  
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(13) = ""
            Else
                StrOkTxt(13) = adoRecordset.Fields(3)
            End If
            'Add by Morgan 2004/1/6
            Lbl1(22).ForeColor = vbBlack
         Else
            StrOk(22) = ""
            'Add by Morgan 2004/1/6
            Lbl1(22).ForeColor = vbRed
            StrOk(22) = strArr(i)
            
            StrOkTxt(13) = ""
         End If
         CheckOC
    Case 27
         StrOk(23) = strArr(i)
    Case 30
         StrOk(24) = strArr(i)
    Case 31
         StrOk(25) = strArr(i)
    Case 37
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(26) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                 StrOk(26) = strArr(i) + "  " + tmp02
             Else
                StrOk(26) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Add By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Add By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(26) = strArr(i) + ""
'                    Else
'                        StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Add By Cheng 2002/07/08
''                    StrOk(26) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Add By Cheng 2002/07/08
''                StrOk(26) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(26) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(26) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(26).ForeColor = vbBlack
         Else
            StrOk(26) = ""
            'Add by Morgan 2004/1/16
            Lbl1(26).ForeColor = vbBlack
            StrOk(26) = strArr(i)
         End If
         CheckOC
    Case 33
         StrOk(27) = strArr(i)
    Case 35
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(28) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                  StrOk(28) = strArr(i) + "  " + tmp02
             Else
                  StrOk(28) = strArr(i)
             End If
         End If
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            'Add By Cheng 2002/07/08
''            If IsNull(adoRecordset.Fields(0)) Then
'            If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))) Then
'               'Add By Cheng 2002/07/08
''                If IsNull(adoRecordset.Fields(1)) Then
'                If IsNull(adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))) Then
'                    If IsNull(adoRecordset.Fields(2)) Then
'                        StrOk(28) = strArr(i) + ""
'                    Else
'                        StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Add By Cheng 2002/07/08
''                    StrOk(28) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Add By Cheng 2002/07/08
''                StrOk(28) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(28) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(28) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(28).ForeColor = vbBlack
         Else
            StrOk(28) = ""
            'Add by Morgan 2004/1/16
            Lbl1(28).ForeColor = vbRed
            StrOk(28) = strArr(i)
         End If
         CheckOC
    Case 36
         StrOk(29) = strArr(i)
    Case 52
         'edit by nick 2004/10/05
         'StrOk(30) = GetPrjSalesNM(strArr(i)) & " " & strArr(53) & " " & strArr(54)
         StrOk(30) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(53))) & " " & Format(strArr(54), "##:##")
    Case 55
         'edit by nick 2004/10/05
         'StrOk(31) = GetPrjSalesNM(strArr(i)) & " " & strArr(56) & " " & strArr(57)
         StrOk(31) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(56))) & " " & Format(strArr(57), "##:##")
    Case 61
         'edit by nickc 2006/07/12
         'StrOk(32) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(32) = ""
         Else
             StrOk(32) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 5
         StrOkTxt(0) = strArr(i)
    Case 6
         StrOkTxt(1) = strArr(i)
    Case 7
         StrOkTxt(2) = strArr(i)
    Case 46
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
    Case 41
         StrOkTxt(5) = strArr(i)
    Case 42
         StrOkTxt(6) = strArr(i)
    Case 43
         StrOkTxt(7) = strArr(i)
    Case 44
         StrOkTxt(8) = strArr(i)
    Case 62
         StrOkTxt(9) = ""
         strExc(0) = "SELECT PTM02,NVL(PTM03,NVL(PTM04,NVL(PTM05,PTM06))) FROM COPYRIGHTITEM,PATENTTRADEMARKMAP " & _
            "WHERE PTM01='3' AND PTM02=CRI02 AND CRI01 IN " & _
            "(SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & Str01 & "' AND " & _
               "CP02='" & Str02 & "' AND CP03='" & Str03 & "' AND CP04='" & Str04 & "' AND CP09<'C')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               StrOkTxt(9) = StrOkTxt(9) & RsTemp.Fields(1) & ";"
               RsTemp.MoveNext
            Loop
            If Right(StrOkTxt(9), 1) = ";" Then StrOkTxt(9) = Left(StrOkTxt(9), Len(StrOkTxt(9)) - 1)
         End If
    Case 45
         StrOkTxt(10) = strArr(i)
    Case 18
         StrOkTxt(11) = strArr(i)
    'Add By Cheng 2002/04/22
    Case 20
         StrOk(33) = strArr(i)
    Case 21
         StrOk(34) = strArr(i)
    Case 67 'D/N固定列印對象
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(35) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
                If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                   StrOk(35) = strArr(i) + "  " + tmp02
                Else
                   StrOk(35) = strArr(i)
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
'                        StrOk(35) = strArr(i) + ""
'                    Else
'                        StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(35) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(35) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(35).ForeColor = vbBlack
         Else
            StrOk(35) = ""
            'Add by Morgan 2004/1/16
            Lbl1(35).ForeColor = vbRed
            StrOk(35) = strArr(i)
         End If
         CheckOC
    'add by nickc 2006/07/12
    Case 68
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(36) = ""
         Else
             Lbl1(36) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 69
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(37) = strArr(i) + ""
            Else
               Lbl1(37) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(37) = ""
         End If
         CheckOC
    Case 70
         Lbl1(38) = strArr(i)
    'Add By Sindy 2012/10/1
    Case 71
         Lbl1(40) = strArr(i)
    '2012/10/1 End
    'Add by Morgan 2008/8/5
    Case 78
         Lbl1(39) = PUB_GetContact(strArr(8), strArr(i))
    'Add by Sindy 2017/11/9 +80,83
    Case 80, 83
         Lbl1(i) = strArr(i)
    '2017/11/9 END
    Case 84 'Add by Morgan 2010/11/8
         Lbl1(84) = strArr(i)
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
    'DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
Next i
'Modify By Cheng 2002/04/22          '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
'For i = 0 To 32                     '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
For i = 0 To 35
   If i <> 0 And i <> 14 And i <> 15 Then 'Add By Sindy 2013/1/31
      Lbl1(i) = StrOk(i)
   End If
Next i
'txt1(53) = StrOkTxt(53)
For i = 0 To 13
   txt1(i) = StrOkTxt(i)
Next i
'傳入參數     代理人
StrTag = strArr(26)
'傳入參數     申請人
StrTag1 = strArr(8)
'Add By Sindy 2010/02/04
cmdok(5).Visible = False
cmdok(6).Visible = False
cmdok(7).Visible = False
cmdok(8).Visible = False
StrTag2 = strArr(58)
StrTag3 = strArr(59)
StrTag4 = strArr(65)
StrTag5 = strArr(66)
If Trim(StrTag2) <> "" Then cmdok(5).Visible = True
If Trim(StrTag3) <> "" Then cmdok(6).Visible = True
If Trim(StrTag4) <> "" Then cmdok(7).Visible = True
If Trim(StrTag5) <> "" Then cmdok(8).Visible = True
'2010/02/04 End
'add by nickc 2005/05/31  檢查有無分割或相關卷號
     cmdok(4).Visible = ChkDataByCR(txt1(14).Text)
End Sub
'edit by nickc 2005/05/31
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(14), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

Private Sub Form_Load()
'Added by Lydia 2021/12/20
Dim Lbl As Object

For Each Lbl In Me.Lbl1
    Lbl.BackColor = &H8000000F
Next
For Each Lbl In Me.Label14
    Lbl.BackColor = &H8000000F
Next
'end 2021/12/20

bolToEndByNick = False
   MoveFormToCenter Me
   If bolFNation = False Then
        SSTab2.TabVisible(2) = False
        cmdok(3).Visible = False
   End If
'92.04.16 nick
cmdState = -1

'Added by Lydia 2020/05/05 各項指示：顯示按鈕
If strSrvDate(1) >= 各項指示啟用日 Then
   cmdok(10).Visible = True
Else
   cmdok(10).Visible = False
End If
'end 2020/05/05

SSTab2.Tab = 0 'Added by Lydia 2021/12/20
End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_A = Nothing
End Sub
'add by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
' 取得客戶名稱
Private Function GetCustName(ByVal strData As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   GetCustName = Empty
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
            If IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustName = rsTmp.Fields("CU06")
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
         ' 檢查讀取的資料筆數
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("FA03")) = False Then
               strKey = rsTmp.Fields("FA03")
               rsTmp.Close
               If Len(strKey) > 8 Then
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                              "CU02 = '" & Mid(strKey, 9, 1) & "'"
               Else
                  strSql = "SELECT * FROM Customer " & _
                        "WHERE CU01 = '" & Mid(strKey, 1, 8) & "' AND " & _
                              "CU02 = '0' "
               End If
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
               If rsTmp.RecordCount > 0 Then
                  rsTmp.MoveFirst
                  If IsNull(rsTmp.Fields("CU04")) = False Then
                     GetCustName = rsTmp.Fields("CU04")
                  ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
                     GetCustName = rsTmp.Fields("CU05")
                  ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
                     GetCustName = rsTmp.Fields("CU06")
                  End If
               End If
            End If
         End If
         rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Function
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


'Added by Lydia 2016/10/26 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
