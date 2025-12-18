VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_B 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料（其他業務）"
   ClientHeight    =   6090
   ClientLeft      =   360
   ClientTop       =   980
   ClientWidth     =   8120
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   8120
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   7
      Left            =   2896
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "已設定代表圖"
      Height          =   400
      Index           =   5
      Left            =   450
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   0
      Width           =   1395
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H008080FF&
      Caption         =   "相關卷號"
      Height          =   400
      Index           =   4
      Left            =   1853
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "代理人資料"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   3
      Left            =   3704
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6320
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "申請人資料"
      Height          =   400
      Index           =   0
      Left            =   5012
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   1300
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7230
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   0
      Width           =   800
   End
   Begin TabDlg.SSTab SSTab3 
      Height          =   5340
      Left            =   30
      TabIndex        =   7
      Top             =   435
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   9419
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   6
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm100101_B.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(1)=   "Label4"
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(5)=   "Label19"
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(7)=   "Label6"
      Tab(0).Control(8)=   "Label5"
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(10)=   "Label88"
      Tab(0).Control(11)=   "Label89"
      Tab(0).Control(12)=   "Label84"
      Tab(0).Control(13)=   "Label91"
      Tab(0).Control(14)=   "Label92"
      Tab(0).Control(15)=   "lbl1(1)"
      Tab(0).Control(16)=   "lbl1(2)"
      Tab(0).Control(17)=   "lbl1(3)"
      Tab(0).Control(18)=   "lbl1(4)"
      Tab(0).Control(19)=   "lbl1(5)"
      Tab(0).Control(20)=   "lbl1(8)"
      Tab(0).Control(21)=   "lbl1(19)"
      Tab(0).Control(22)=   "Label14"
      Tab(0).Control(23)=   "Label9"
      Tab(0).Control(24)=   "Label12"
      Tab(0).Control(25)=   "Label17"
      Tab(0).Control(26)=   "lbl1(24)"
      Tab(0).Control(27)=   "Label26(1)"
      Tab(0).Control(28)=   "lbl1(25)"
      Tab(0).Control(29)=   "lbl1(85)"
      Tab(0).Control(30)=   "Label113"
      Tab(0).Control(31)=   "Label112"
      Tab(0).Control(32)=   "lbl1(87)"
      Tab(0).Control(33)=   "Label25"
      Tab(0).Control(34)=   "lbl1(86)"
      Tab(0).Control(35)=   "Label27"
      Tab(0).Control(36)=   "txt1(5)"
      Tab(0).Control(37)=   "txt1(2)"
      Tab(0).Control(38)=   "txt1(0)"
      Tab(0).Control(39)=   "txt1(1)"
      Tab(0).Control(40)=   "txt1(3)"
      Tab(0).Control(41)=   "txt1(7)"
      Tab(0).Control(42)=   "txt1(8)"
      Tab(0).Control(43)=   "txt1(9)"
      Tab(0).Control(44)=   "txt1(10)"
      Tab(0).Control(45)=   "txt1(11)"
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "代理人"
      TabPicture(1)   =   "frm100101_B.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo3(1)"
      Tab(1).Control(1)=   "txt1(6)"
      Tab(1).Control(2)=   "Label80(29)"
      Tab(1).Control(3)=   "lbl1(88)"
      Tab(1).Control(4)=   "Label80(26)"
      Tab(1).Control(5)=   "lbl1(26)"
      Tab(1).Control(6)=   "Label11"
      Tab(1).Control(7)=   "lbl1(84)"
      Tab(1).Control(8)=   "Label21"
      Tab(1).Control(9)=   "lbl1(16)"
      Tab(1).Control(10)=   "lbl1(20)"
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(12)=   "lbl1(15)"
      Tab(1).Control(13)=   "lbl1(14)"
      Tab(1).Control(14)=   "Label31"
      Tab(1).Control(15)=   "Label28"
      Tab(1).Control(16)=   "Label24"
      Tab(1).Control(17)=   "Label18"
      Tab(1).Control(18)=   "Label16"
      Tab(1).Control(19)=   "Label15"
      Tab(1).Control(20)=   "Label29"
      Tab(1).Control(21)=   "Label7"
      Tab(1).Control(22)=   "lbl1(9)"
      Tab(1).Control(23)=   "lbl1(10)"
      Tab(1).Control(24)=   "lbl1(11)"
      Tab(1).Control(25)=   "lbl1(12)"
      Tab(1).Control(26)=   "lbl1(13)"
      Tab(1).Control(27)=   "Label20"
      Tab(1).ControlCount=   28
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm100101_B.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lbl1(23)"
      Tab(2).Control(1)=   "lbl1(22)"
      Tab(2).Control(2)=   "lbl1(21)"
      Tab(2).Control(3)=   "lbl1(7)"
      Tab(2).Control(4)=   "Label78"
      Tab(2).Control(5)=   "Label79"
      Tab(2).Control(6)=   "Label80(0)"
      Tab(2).Control(7)=   "Label81"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "其他"
      TabPicture(3)   =   "frm100101_B.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label1(164)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(165)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(166)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label22"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "lbl1(80)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "lbl1(81)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "lbl1(82)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "lbl1(83)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frame1K"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "參考備註"
      TabPicture(4)   =   "frm100101_B.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdOK(6)"
      Tab(4).Control(1)=   "txt1(4)"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame1K 
         Enabled         =   0   'False
         Height          =   280
         Left            =   30
         TabIndex        =   104
         Top             =   1560
         Width           =   4870
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   107
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   106
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   105
            Top             =   60
            Width           =   910
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   34
            Left            =   150
            TabIndex        =   108
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Index           =   1
         ItemData        =   "frm100101_B.frx":008C
         Left            =   -69090
         List            =   "frm100101_B.frx":009F
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   96
         Top             =   3450
         Width           =   1470
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "各項指示"
         Height          =   330
         Index           =   6
         Left            =   -74880
         TabIndex        =   95
         Top             =   330
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSForms.TextBox txt1 
         Height          =   4200
         Index           =   4
         Left            =   -74880
         TabIndex        =   92
         Top             =   720
         Width           =   7785
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13732;7408"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   11
         Left            =   -73980
         TabIndex        =   80
         Top             =   2850
         Width           =   1935
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3413;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   10
         Left            =   -73860
         TabIndex        =   79
         Top             =   330
         Width           =   2145
         VariousPropertyBits=   671105055
         BackColor       =   16777215
         Size            =   "3784;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   465
         Index           =   9
         Left            =   -74010
         TabIndex        =   71
         Top             =   2100
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
         Index           =   8
         Left            =   -74010
         TabIndex        =   69
         Top             =   1620
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
         Height          =   945
         Index           =   7
         Left            =   -73650
         TabIndex        =   60
         Top             =   690
         Width           =   6555
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11562;1658"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   1005
         Index           =   6
         Left            =   -74880
         TabIndex        =   39
         Top             =   3780
         Width           =   7830
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "13822;1764"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   390
         Index           =   3
         Left            =   -70110
         TabIndex        =   11
         Top             =   3135
         Width           =   2895
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "5106;688"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   1
         Left            =   -73650
         TabIndex        =   9
         Top             =   1005
         Width           =   6555
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11557;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   0
         Left            =   -73650
         TabIndex        =   8
         Top             =   690
         Width           =   6555
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11557;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   300
         Index           =   2
         Left            =   -73650
         TabIndex        =   10
         Top             =   1320
         Width           =   6555
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "11557;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txt1 
         Height          =   495
         Index           =   5
         Left            =   -73950
         TabIndex        =   12
         Top             =   4500
         Width           =   6915
         VariousPropertyBits=   -1467989985
         BackColor       =   16777215
         ScrollBars      =   2
         Size            =   "12192;882"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "國內副本收件人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   103
         Top             =   4230
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   86
         Left            =   -73365
         TabIndex        =   102
         Top             =   4230
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
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "國內副本接洽人："
         Height          =   180
         Left            =   -70335
         TabIndex        =   101
         Top             =   4230
         Width           =   1440
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   87
         Left            =   -68865
         TabIndex        =   100
         Top             =   4230
         Width           =   1695
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2990;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Index           =   29
         Left            =   -70890
         TabIndex        =   99
         Top             =   3510
         Width           =   1800
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   88
         Left            =   -71460
         TabIndex        =   98
         Top             =   3510
         Width           =   435
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "767;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   26
         Left            =   -72390
         TabIndex        =   97
         Top             =   3510
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   26
         Left            =   -73920
         TabIndex        =   94
         Top             =   1490
         Width           =   6850
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12083;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2："
         Height          =   180
         Left            =   -74880
         TabIndex        =   93
         Top             =   1490
         Width           =   810
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   83
         Left            =   1665
         TabIndex        =   91
         Top             =   1290
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   82
         Left            =   1665
         TabIndex        =   90
         Top             =   990
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   81
         Left            =   1665
         TabIndex        =   89
         Top             =   690
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   80
         Left            =   1665
         TabIndex        =   88
         Top             =   390
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "是否以Email通知:          (Y:是   D:僅D/N）"
         Height          =   180
         Left            =   180
         TabIndex        =   87
         Top             =   390
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email同時寄紙本:          (Y:是)"
         Height          =   180
         Index           =   166
         Left            =   180
         TabIndex        =   86
         Top             =   1290
         Width           =   2265
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款單份數:"
         Height          =   180
         Index           =   165
         Left            =   180
         TabIndex        =   85
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿份數:"
         Height          =   180
         Index           =   164
         Left            =   180
         TabIndex        =   84
         Top             =   690
         Width           =   765
      End
      Begin VB.Label Label112 
         AutoSize        =   -1  'True
         Caption         =   "(J:智權公司 空白:系統預設)"
         Height          =   180
         Left            =   -73260
         TabIndex        =   83
         Top             =   5040
         Width           =   2115
      End
      Begin VB.Label Label113 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司："
         Height          =   180
         Left            =   -74880
         TabIndex        =   82
         Top             =   5040
         Width           =   1260
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   85
         Left            =   -73605
         TabIndex        =   81
         Top             =   5040
         Width           =   300
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "529;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   84
         Left            =   -72900
         TabIndex        =   78
         Top             =   940
         Width           =   3225
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5689;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   180
         Left            =   -74880
         TabIndex        =   77
         Top             =   940
         Width           =   1860
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   25
         Left            =   -69960
         TabIndex        =   76
         Top             =   3690
         Width           =   2145
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3784;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         Caption         =   "工程師組別："
         Height          =   180
         Index           =   1
         Left            =   -71085
         TabIndex        =   75
         Top             =   3705
         Width           =   1080
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   24
         Left            =   -69555
         TabIndex        =   74
         Top             =   3960
         Width           =   2370
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4180;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "接洽人："
         Height          =   180
         Left            =   -70350
         TabIndex        =   73
         Top             =   3960
         Width           =   720
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "商品組群："
         Height          =   180
         Left            =   -74880
         TabIndex        =   72
         Top             =   2130
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "商品類別："
         Height          =   180
         Left            =   -74880
         TabIndex        =   70
         Top             =   1635
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   23
         Left            =   -73530
         TabIndex        =   68
         Top             =   1260
         Width           =   3660
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "6456;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   22
         Left            =   -73710
         TabIndex        =   67
         Top             =   960
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   21
         Left            =   -73770
         TabIndex        =   66
         Top             =   660
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   7
         Left            =   -73770
         TabIndex        =   65
         Top             =   360
         Width           =   1200
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2117;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74850
         TabIndex        =   64
         Top             =   1260
         Width           =   1260
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74850
         TabIndex        =   63
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Index           =   0
         Left            =   -74850
         TabIndex        =   62
         Top             =   660
         Width           =   1080
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74850
         TabIndex        =   61
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱："
         Height          =   180
         Left            =   -74880
         TabIndex        =   59
         Top             =   690
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   16
         Left            =   -73800
         TabIndex        =   40
         Top             =   2895
         Width           =   6645
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11721;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   20
         Left            =   -73260
         TabIndex        =   58
         Top             =   3180
         Width           =   6105
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "10769;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   180
         Left            =   -74880
         TabIndex        =   57
         Top             =   3180
         Width           =   1545
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   15
         Left            =   -73800
         TabIndex        =   41
         Top             =   2625
         Width           =   6645
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11721;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   14
         Left            =   -73110
         TabIndex        =   42
         Top             =   2325
         Width           =   435
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "767;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   19
         Left            =   -70050
         TabIndex        =   37
         Top             =   345
         Width           =   2745
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4842;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "代理人備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   56
         Top             =   3510
         Width           =   1080
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "副本聯絡人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   55
         Top             =   2925
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "副本收受人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   54
         Top             =   2625
         Width           =   1080
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   180
         Left            =   -74880
         TabIndex        =   52
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   51
         Top             =   1215
         Width           =   720
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "折扣："
         Height          =   180
         Left            =   -74880
         TabIndex        =   50
         Top             =   1765
         Width           =   540
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "FC代理人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   49
         Top             =   390
         Width           =   930
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   48
         Top             =   665
         Width           =   900
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   9
         Left            =   -73920
         TabIndex        =   47
         Top             =   390
         Width           =   6795
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11986;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   10
         Left            =   -73950
         TabIndex        =   46
         Top             =   665
         Width           =   4695
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "8281;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   11
         Left            =   -73920
         TabIndex        =   45
         Top             =   1215
         Width           =   6850
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "12083;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   12
         Left            =   -73920
         TabIndex        =   44
         Top             =   1765
         Width           =   1875
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "3307;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   13
         Left            =   -73605
         TabIndex        =   43
         Top             =   2040
         Width           =   6390
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "11271;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   8
         Left            =   -73605
         TabIndex        =   34
         Top             =   3960
         Width           =   3060
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "5397;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   5
         Left            =   -73980
         TabIndex        =   33
         Top             =   3690
         Width           =   2730
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "4815;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   4
         Left            =   -73980
         TabIndex        =   32
         Top             =   3420
         Width           =   435
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "767;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   3
         Left            =   -73980
         TabIndex        =   31
         Top             =   3165
         Width           =   1545
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "2725;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   2
         Left            =   -70110
         TabIndex        =   30
         Top             =   2873
         Width           =   405
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "lblFM2"
         Size            =   "714;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   1
         Left            =   -73980
         TabIndex        =   29
         Top             =   2580
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
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   28
         Top             =   3960
         Width           =   1260
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "分所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   27
         Top             =   3690
         Width           =   900
      End
      Begin VB.Label Label84 
         AutoSize        =   -1  'True
         Caption         =   "定稿語文：           （1.中文  2.英文  3.日文）"
         Height          =   180
         Left            =   -74880
         TabIndex        =   26
         Top             =   3457
         Width           =   3420
      End
      Begin VB.Label Label89 
         Caption         =   "閉卷原因："
         Height          =   180
         Left            =   -71040
         TabIndex        =   25
         Top             =   3202
         Width           =   975
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         Height          =   180
         Left            =   -74880
         TabIndex        =   24
         Top             =   3165
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         Height          =   180
         Left            =   -74880
         TabIndex        =   23
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label5 
         Caption         =   "案件名稱(中)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   22
         Top             =   690
         Visible         =   0   'False
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(英)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   21
         Top             =   1005
         Width           =   1200
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱(日)："
         Height          =   180
         Left            =   -74880
         TabIndex        =   20
         Top             =   1320
         Width           =   1200
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         Height          =   180
         Left            =   -74880
         TabIndex        =   19
         Top             =   2580
         Width           =   720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "申請日："
         Height          =   180
         Left            =   -74880
         TabIndex        =   18
         Top             =   2895
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "客戶備註："
         Height          =   180
         Left            =   -74880
         TabIndex        =   17
         Top             =   4530
         Width           =   900
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "是否閉卷： "
         Height          =   180
         Left            =   -71070
         TabIndex        =   16
         Top             =   2895
         Width           =   945
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "(Y：閉卷)"
         Height          =   180
         Left            =   -69555
         TabIndex        =   15
         Top             =   2895
         Width           =   780
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         Height          =   180
         Index           =   0
         Left            =   -71040
         TabIndex        =   38
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：            （Y：印）"
         Height          =   180
         Left            =   -74880
         TabIndex        =   53
         Top             =   2340
         Width           =   3105
      End
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   18
      Left            =   4920
      TabIndex        =   36
      Top             =   5805
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
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   17
      Left            =   1065
      TabIndex        =   35
      Top             =   5805
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
   Begin VB.Label Label51 
      Caption         =   "Update ID："
      Height          =   180
      Left            =   3960
      TabIndex        =   14
      Top             =   5805
      Width           =   975
   End
   Begin VB.Label Label49 
      Caption         =   "Create ID："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   5805
      Width           =   855
   End
End
Attribute VB_Name = "frm100101_B"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/12/13 改成Form2.0 ; lbl1(index)、txt1(index)
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
     frm100108_3.Tag = txt1(10).Text
     frm100108_3.Caption = "相關卷號"
     frm100108_3.StrMenu2
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'Add by Amy 2016/08/26 代表圖
Case 5
    frmPic001.oCP01 = SystemNumber(txt1(10), 1)
    frmPic001.oCP02 = SystemNumber(txt1(10), 2)
    frmPic001.oCP03 = SystemNumber(txt1(10), 3)
    frmPic001.oCP04 = SystemNumber(txt1(10), 4)
    frmPic001.StrMenu
    frmPic001.CanScan
    frmPic001.SetSeekCmdok 'Add by Amy 2018/07/16
    frmPic001.Show vbModal
    Call ReadPic
'Added by Lydia 2016/11/23
Case 6 '各項指示
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
     frm12040159.SetParent "Q", Trim(Replace(txt1(10), "-", "")), Me
     frm12040159.Show
     Screen.MousePointer = vbDefault
     Me.Enabled = True
'end 2016/11/23
'Add By Sindy 2020/7/15
Case 7 '進度
   cmdState = -1
   Me.Enabled = False
   If fnSaveParentForm(Me) = False Then
      Me.Enabled = True
      Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   frm100101_2.Show
   frm100101_2.Tag = txt1(10)
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
'Dim strArr(T_SP) As String, i As Integer, StrOk(20) As String, StrOkTxt(6) As String
Dim strArr() As String, i As Integer, StrOk(21) As String, StrOkTxt(6) As String
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
Dim arrID 'Add By Sindy 2025/1/7

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

'add by Toni 20080926 控制跨部門權限
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
      Case 9, 11, 15, 19, 20, 30, 38, 39, 52, 53, 55, 56, 75
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
CheckOC
Dim strTemp As String    '暫存
Dim strTemp1 As Variant, strTemp2 As Variant, strTemp3 As Variant
Dim j As Integer
intK = 62
For i = 1 To tf_SP 'edit by nickc 2006/07/12 T_SP
    Select Case i
    Case 1
         StrOk(0) = strArr(1) + "-" + strArr(2) + "-" + strArr(3) + "-" + strArr(4)
         txt1(10) = StrOk(0) 'Add By Sindy 2013/1/31
         'Add by Amy 2016/08/26
         Call ReadPic '檢查有無代表圖
         cmdok(5).Visible = False
         If Str01 = "TS" Then
            If cmdok(4).Visible = False Then cmdok(5).Left = 2350
            cmdok(5).Visible = True
         End If
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
             If IsNull(adoRecordset.Fields(1)) Then
                  StrOkTxt(5) = ""
             Else
                  StrOkTxt(5) = adoRecordset.Fields(1)
             End If
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbBlack
         Else
            StrOk(1) = ""
            'Add by Morgan 2004/1/14
            Lbl1(1).ForeColor = vbRed
             StrOk(1) = strArr(i)
             
             StrOkTxt(5) = ""
         End If
         CheckOC
     Case 15
         StrOk(2) = strArr(i)
     Case 16
         StrOk(3) = strArr(i)
    Case 34
         StrOk(4) = strArr(i)
    Case 28
         StrOk(5) = strArr(i)
    Case 10
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(6) = ""
         Else
             StrOk(6) = ChangeWStringToTString(strArr(i))
         End If
         txt1(11) = StrOk(6) 'Add By Sindy 2013/1/31
    Case 61
         'edit by nickc 2006/07/12
         'StrOk(7) = strArr(i)
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             StrOk(7) = ""
         Else
             StrOk(7) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 29
         StrOk(8) = strArr(i)
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
            'Add By Cheng 2002/07/08
'            If IsNull(adoRecordset.Fields(0)) Then
            '2005/9/14 MODIFY BY SONIA
            'If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) Then
'
'            If Trim(adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))) = "" Then
'            '2005/9/14 END
'               'Add By Cheng 2002/07/08
''               If IsNull(adoRecordset.Fields(1)) Then
'               If IsNull(adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))) Then
'                   If IsNull(adoRecordset.Fields(2)) Then
'                          StrOk(9) = strArr(i) + ""
'                   Else
'                         StrOk(9) = strArr(i) + "  " + adoRecordset.Fields(2)
'                   End If
'               Else
'                  'Add By Cheng 2002/07/08
''                   StrOk(9) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                   StrOk(9) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 0, 1))
'               End If
'            Else
'               'Add By Cheng 2002/07/08
''               StrOk(9) = StrArr(i) + "  " + adoRecordset.Fields(0)
'               StrOk(9) = strArr(i) + "  " + adoRecordset.Fields(IIf(strSK03 = "0", 1, 0))
'
'            End If
            If IsNull(adoRecordset.Fields("FA05")) = False Then
               StrOk(9) = strArr(i) + "  " + adoRecordset.Fields("FA05")
            ElseIf IsNull(adoRecordset.Fields("FA04")) = False Then
               StrOk(9) = strArr(i) + "  " + adoRecordset.Fields("FA04")
            ElseIf IsNull(adoRecordset.Fields("FA06")) = False Then
               StrOk(9) = strArr(i) + "  " + adoRecordset.Fields("FA06")
            End If
            If IsNull(adoRecordset.Fields(3)) Then
                StrOkTxt(6) = ""
            Else
                StrOkTxt(6) = adoRecordset.Fields(3)
            End If
            'Add by Morgan 2004/1/16
            Lbl1(9).ForeColor = vbBlack
         Else
            StrOk(9) = ""
            'Add by Morgan 2004/1/16
            Lbl1(9).ForeColor = vbRed
            StrOk(9) = strArr(i)
            
            StrOkTxt(6) = ""
         End If
         CheckOC
    Case 27
         StrOk(10) = strArr(i)
    Case 30
         StrOk(11) = strArr(i)
    Case 31
         StrOk(12) = strArr(i)
    Case 37
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(13) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
            If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
               StrOk(13) = strArr(i) + "  " + tmp02
            Else
               StrOk(13) = strArr(i)
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
'                        StrOk(13) = strArr(i) + ""
'                    Else
'                        StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Add By Cheng 2002/07/08
''                    StrOk(13) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Add By Cheng 2002/07/08
''                StrOk(13) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(13) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(13) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(13).ForeColor = vbBlack
         Else
            StrOk(13) = ""
            'Add by Morgan 2004/1/16
            Lbl1(13).ForeColor = vbRed
            StrOk(13) = strArr(i)
         End If
         CheckOC
    Case 33
         StrOk(14) = strArr(i)
    Case 35
         If Left$(strArr(i), 1) = "X" Then
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='" & Right$(strArr(i), 1) & "'"
'             Else
'                  strSQL = "SELECT cu05||' '||cu88||' '||cu89||' '||cu90,CU04,CU06 FROM CUSTOMER WHERE CU01='" & Left$(strArr(i), 8) & "' AND CU02='0'"
'             End If
             StrOk(15) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
             If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                StrOk(15) = strArr(i) + "  " + tmp02
             Else
                StrOk(15) = strArr(i)
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
'                        StrOk(15) = strArr(i) + ""
'                    Else
'                        StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                  'Add By Cheng 2002/07/08
''                    StrOk(15) = StrArr(i) + "  " + adoRecordset.Fields(1)
'                    StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'               'Add By Cheng 2002/07/08
''                StrOk(15) = StrArr(i) + "  " + adoRecordset.Fields(0)
'                StrOk(15) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(15) <> strArr(i) Then
            'Add by Morgan 2004/1/16
            Lbl1(15).ForeColor = vbBlack
         Else
            StrOk(15) = ""
            'Add by Morgan 2004/1/16
            Lbl1(15).ForeColor = vbRed
            StrOk(15) = strArr(i)
         End If
         CheckOC
    Case 36
         StrOk(16) = strArr(i)
    Case 52
         'edit by nick 2004/10/05
         'StrOk(17) = strArr(i)
         StrOk(17) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(53))) & " " & Format(strArr(54), "##:##")
    Case 55
         'edit by nick 2004/10/05
         'StrOk(18) = strArr(i)
         StrOk(18) = GetPrjSalesNM(strArr(i)) & " " & ChangeTStringToTDateString(ChangeWStringToTString(strArr(56))) & " " & Format(strArr(57), "##:##")
    Case 5
         StrOkTxt(0) = strArr(i)
    Case 6
         StrOkTxt(1) = strArr(i)
    Case 7
         StrOkTxt(2) = strArr(i)
    Case 17
         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
             If IsNull(adoRecordset.Fields(0)) Then
                     StrOkTxt(3) = ""
             Else
                     StrOkTxt(3) = adoRecordset.Fields(0)
             End If
         Else
             StrOkTxt(3) = ""
         End If
         CheckOC
    Case 18
         StrOkTxt(4) = strArr(i)
    Case 9
         strSql = "SELECT NA03 FROM NATION WHERE NA01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
              If IsNull(adoRecordset.Fields(0)) Then
                  StrOk(19) = strArr(i) + ""
              Else
                  StrOk(19) = strArr(i) + "  " + adoRecordset.Fields(0)
              End If
         Else
              StrOk(19) = ""
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
             StrOk(20) = strArr(i) + "  " + GetAgentOrCustName(Trim(strArr(i)))
         Else
'edit by nickc 2007/08/27 秀玲說基本檔怎麼抓，就跟基本檔一樣，不用統一中英日或是英中日
'             If Len(strArr(i)) = 9 Then
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='" & Right$(Trim(strArr(i)), 1) & "'"
'             Else
'                  strSQL = "SELECT FA05||' '||FA63||' '||FA64||' '||FA65,FA04,FA06,FA29 FROM FAGENT WHERE FA01='" & Left$(Trim(strArr(i)), 8) & "' AND FA02='0'"
'             End If
                If PUB_GetAgentName(Str01, Trim(strArr(i)), tmp02) Then
                   StrOk(20) = strArr(i) + "  " + tmp02
                Else
                   StrOk(20) = strArr(i)
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
'                        StrOk(20) = strArr(i) + ""
'                    Else
'                        StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(2)
'                    End If
'                Else
'                    StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 0, 1)))
'                End If
'            Else
'                StrOk(20) = strArr(i) + "  " + adoRecordset.Fields(IIf(Left$(strArr(i), 1) = "X", 0, IIf(strSK03 = "0", 1, 0)))
'            End If
         If StrOk(20) <> strArr(i) Then
            'Add by Morgan 2004/1/6
            Lbl1(20).ForeColor = vbBlack
         Else
            StrOk(20) = ""
            'Add by Morgan 2004/1/6
            Lbl1(20).ForeColor = vbRed
            StrOk(20) = strArr(i)
         End If
         CheckOC
    'add by nickc 2006/07/12
    Case 68
         If Len(Trim(strArr(i))) = 0 Or strArr(i) = "0" Then
             Lbl1(21) = ""
         Else
             Lbl1(21) = ChangeTStringToTDateString(ChangeWStringToTString(strArr(i)))
         End If
    Case 69
         strSql = "SELECT nvl(ST02,'" & strArr(i) & "') FROM STAFF WHERE ST01='" & strArr(i) & "'"
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            If IsNull(adoRecordset.Fields(0)) Then
               Lbl1(22) = strArr(i) + ""
            Else
               Lbl1(22) = strArr(i) + "  " + adoRecordset.Fields(0)
            End If
         Else
            Lbl1(22) = ""
         End If
         CheckOC
    Case 70
         Lbl1(23) = strArr(i)
    'add by nickc 2006/12/07
    Case 73
         txt1(8) = strArr(i)
    Case 74
         txt1(9) = strArr(i)
    'Add By Sindy 2016/11/4 聯絡人2
    Case 75
         Lbl1(26) = strArr(i) 'StrOk(21) = strArr(i)
    '2016/11/4 END
    'Add by Morgan 2008/8/5
    Case 78
         Lbl1(24) = PUB_GetContact(strArr(8), strArr(i))
    'Add by Moran 2009/9/8
    Case 79
         If strArr(i) = "" Then
            Lbl1(25) = ""
         Else
            Lbl1(25) = strArr(i) + "." + PUB_GetFCPGrpName(strArr(i))
         End If
    
    'Add by Morgan 2010/11/5
    'Modified by Morgan 2014/6/5 +80,81,82,83
    Case 80, 81, 82, 83, 84
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
    'Add By Sindy 2016/11/24
    Case 88
      Lbl1(88) = strArr(i)
    Case 89
      Combo3(1).ListIndex = Val(strArr(i))
    '2016/11/24 END
    'Add by Sindy 2025/1/7
    Case 90
         If Trim(strArr(i)) <> "" Then
            arrID = Split(strArr(i), ",")
            For intI = UBound(arrID) To LBound(arrID) Step -1
               Chk1K(Val(arrID(intI)) - 1).Value = 1
            Next intI
         End If
    '2025/1/7 END
    Case Else
    End Select
    'DoEvents Add By Sindy 2019/1/4 Mark,因為會和視窗的function(MenuForFormControl)有ErrCode互影響
Next i
For i = 0 To 21 '20            '2006/07/12 加備註，以後新增欄位，直接在上面修改，此2段迴圈
   If i <> 0 And i <> 6 Then 'Add By Sindy 2013/1/31
      Lbl1(i) = StrOk(i)      '不可修改，不然會影響資料顯現，而且陣列的宣告也不用一直的修改
   End If
Next i
'txt1(53) = StrOkTxt(53)
For i = 0 To 6
   txt1(i) = StrOkTxt(i)
Next i
'傳入參數     代理人
StrTag = strArr(26)
'傳入參數     申請人
StrTag1 = strArr(8)
'Add By Cheng 2004/02/25
Select Case strArr(1)
Case "TS", "S"
    Me.Label14.Visible = True
    Me.txt1(7).Visible = True
    Me.txt1(7).Enabled = True
    Me.Label5.Visible = False
    Me.txt1(0).Visible = False
    Me.txt1(0).Enabled = False
    Me.Label6.Visible = False
    Me.txt1(1).Visible = False
    Me.txt1(1).Enabled = False
    Me.Label23.Visible = False
    Me.txt1(2).Visible = False
    Me.txt1(2).Enabled = False
    Me.txt1(7).Text = StrOkTxt(0)
Case Else
    Me.Label14.Visible = False
    Me.txt1(7).Visible = False
    Me.txt1(7).Enabled = False
    Me.Label5.Visible = True
    Me.txt1(0).Visible = True
    Me.txt1(0).Enabled = True
    Me.Label6.Visible = True
    Me.txt1(1).Visible = True
    Me.txt1(1).Enabled = True
    Me.Label23.Visible = True
    Me.txt1(2).Visible = True
    Me.txt1(2).Enabled = True
End Select
'End

'add by nickc 2005/05/31  檢查有無分割或相關卷號
     cmdok(4).Visible = ChkDataByCR(txt1(10).Text)
End Sub

Private Sub Form_Activate()
   pub_QL05 = m_pub_QL05 'Add By Sindy 2025/8/28 還原此Form的查詢條件記錄
End Sub

'edit by nickc 2005/05/31
'Private Sub cmdRef_Click()
'    Dim stTmp As String
'    stTmp = Right(Space(2) & txt1(10), 15)
'    Where1103ComeFrom Me, Trim(Left(stTmp, 3)), Mid(stTmp, 5, 6), Mid(stTmp, 12, 1), Mid(stTmp, 14, 2)
'End Sub

Private Sub Form_Load()
Dim Lbl As Object

   For Each Lbl In Me.Lbl1
       Lbl.BackColor = &H8000000F
   Next
   bolToEndByNick = False
   
   SSTab3.Tab = 0 'Added by Lydia 2016/11/23
   
      MoveFormToCenter Me
   If bolFNation = False Then
       SSTab3.TabVisible(1) = False
       cmdok(3).Visible = False
   End If
   '92.04.16 nick
   cmdState = -1
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdok(6).Visible = True
   Else
      cmdok(6).Visible = False
      txt1(4).Top = 360
      txt1(4).Height = 4570
   End If
   'end 2020/05/05
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

Private Sub Form_Unload(Cancel As Integer)
pub_QL05 = m_pub_QL05 'Add By Sindy 2025/9/12 還原此Form的查詢條件記錄 (多筆查詢有影響)
Set frm100101_B = Nothing
End Sub

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

'Add by Amy 2016/08/26
Private Sub ReadPic()
    'Modify by Amy 2018/07/16  改寫至function
'    strSql = "SELECT ibf01,ibf02 FROM imgbytefile WHERE ibf01='" & SystemNumber(txt1(10), 1) & "' and ibf02='" & SystemNumber(txt1(10), 2) & "' and ibf03='" & SystemNumber(txt1(10), 3) & "' and ibf04='" & SystemNumber(txt1(10), 4) & "' and ibf05='1'"
'    CheckOC2
'    adoRecordset1.CursorLocation = adUseClient
'    adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    If ChkImgByteFile(SystemNumber(txt1(10), 1), SystemNumber(txt1(10), 2), SystemNumber(txt1(10), 3), SystemNumber(txt1(10), 4)) = True Then
        'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
        cmdok(5).Caption = "已設定代表圖"
        cmdok(5).BackColor = &HC0FFC0
    Else
        'Modified by Lydia 2021/12/16 拿掉快速鍵(&I)
        cmdok(5).Caption = "未設定代表圖"
        cmdok(5).BackColor = &HC0C0FF
    End If
'    CheckOC2
    'end 2018/07/16
End Sub

'Added by Lydia 2016/10/27 修正Win7 輸入法問題
Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index) 'Added by Lydia 2016/12/6
   OpenIme
End Sub
