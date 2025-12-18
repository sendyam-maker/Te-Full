VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm110104_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "更換FC代理人作業－非專利"
   ClientHeight    =   5510
   ClientLeft      =   350
   ClientTop       =   1440
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5510
   ScaleWidth      =   8950
   Begin TabDlg.SSTab SSTab1 
      Height          =   3650
      Left            =   120
      TabIndex        =   34
      Top             =   1830
      Width           =   8700
      _ExtentX        =   15346
      _ExtentY        =   6438
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm110104_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(62)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(55)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(32)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(34)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(45)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblTM30_T"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblTM57_T"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblTM57"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblTM30"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(47)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(154)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(48)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(155)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(169)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(64)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(86)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(84)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblTM129"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(4)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(164)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(165)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(166)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label68"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lblTM70"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lblTM69"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblTM56"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lblTM66"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "lblTM33"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM71"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Label1(1)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Label1(2)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textTM35"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM56"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textTM37"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM36"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM69"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM70"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM46"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM122"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM127"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTM65"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM66"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textTM129"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textTM68"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM124"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textTM125"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textTM126"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textTM121"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textTM33"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textTM141"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textTM140"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).ControlCount=   51
      TabCaption(1)   =   "代理人／聯絡人"
      TabPicture(1)   =   "frm110104_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM45"
      Tab(1).Control(1)=   "NewFagent"
      Tab(1).Control(2)=   "textTM40"
      Tab(1).Control(3)=   "textTM41"
      Tab(1).Control(4)=   "textTM42"
      Tab(1).Control(5)=   "textTM43"
      Tab(1).Control(6)=   "textTM38"
      Tab(1).Control(7)=   "textTM39"
      Tab(1).Control(8)=   "textTM76"
      Tab(1).Control(9)=   "lblAgent"
      Tab(1).Control(10)=   "Label1(65)"
      Tab(1).Control(11)=   "Label7"
      Tab(1).Control(12)=   "Label6"
      Tab(1).Control(13)=   "Label49"
      Tab(1).Control(14)=   "Label51"
      Tab(1).Control(15)=   "Label55"
      Tab(1).Control(16)=   "Label57"
      Tab(1).Control(17)=   "Label59"
      Tab(1).Control(18)=   "Label4"
      Tab(1).ControlCount=   19
      TabCaption(2)   =   "參考備註"
      TabPicture(2)   =   "frm110104_4.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdIns"
      Tab(2).Control(1)=   "textTM58"
      Tab(2).ControlCount=   2
      Begin VB.TextBox textTM140 
         Height          =   300
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   6
         Top             =   1170
         Width           =   300
      End
      Begin VB.TextBox textTM141 
         Height          =   300
         Left            =   3860
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1170
         Width           =   300
      End
      Begin VB.TextBox textTM45 
         Height          =   300
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   22
         Top             =   697
         Width           =   2772
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   330
         Left            =   -74880
         TabIndex        =   30
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox textTM33 
         Height          =   300
         Left            =   1695
         MaxLength       =   8
         TabIndex        =   14
         Top             =   2370
         Width           =   975
      End
      Begin VB.TextBox textTM121 
         Height          =   300
         Left            =   6750
         MaxLength       =   1
         TabIndex        =   18
         Top             =   2990
         Width           =   255
      End
      Begin VB.TextBox textTM126 
         Height          =   300
         Left            =   6870
         MaxLength       =   1
         TabIndex        =   20
         Top             =   3300
         Width           =   255
      End
      Begin VB.TextBox textTM125 
         Height          =   300
         Left            =   7620
         MaxLength       =   1
         TabIndex        =   5
         Top             =   880
         Width           =   255
      End
      Begin VB.TextBox textTM124 
         Height          =   300
         Left            =   5790
         MaxLength       =   1
         TabIndex        =   4
         Top             =   880
         Width           =   255
      End
      Begin VB.TextBox textTM68 
         Height          =   300
         Left            =   5100
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2050
         Width           =   255
      End
      Begin VB.TextBox textTM129 
         Height          =   300
         Left            =   7320
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2050
         Width           =   255
      End
      Begin VB.TextBox textTM66 
         Height          =   300
         Left            =   1695
         MaxLength       =   8
         TabIndex        =   17
         Top             =   2990
         Width           =   975
      End
      Begin VB.TextBox textTM65 
         Height          =   300
         Left            =   1695
         MaxLength       =   30
         TabIndex        =   15
         Top             =   2680
         Width           =   2772
      End
      Begin VB.TextBox textTM127 
         Height          =   300
         Left            =   6060
         MaxLength       =   20
         TabIndex        =   1
         Top             =   570
         Width           =   2500
      End
      Begin VB.TextBox textTM122 
         Height          =   300
         Left            =   1875
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2050
         Width           =   255
      End
      Begin VB.TextBox textTM46 
         Height          =   300
         Left            =   1875
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1430
         Width           =   255
      End
      Begin VB.TextBox textTM70 
         Height          =   300
         Left            =   1695
         MaxLength       =   8
         TabIndex        =   19
         Top             =   3300
         Width           =   975
      End
      Begin VB.TextBox NewFagent 
         Height          =   300
         Left            =   -73440
         MaxLength       =   9
         TabIndex        =   21
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox textTM69 
         Height          =   300
         Left            =   4890
         MaxLength       =   8
         TabIndex        =   9
         Top             =   1430
         Width           =   975
      End
      Begin VB.TextBox textTM36 
         Height          =   300
         Left            =   1410
         MaxLength       =   2
         TabIndex        =   2
         Top             =   880
         Width           =   300
      End
      Begin VB.TextBox textTM37 
         Height          =   300
         Left            =   3855
         MaxLength       =   2
         TabIndex        =   3
         Top             =   880
         Width           =   300
      End
      Begin VB.TextBox textTM56 
         Height          =   300
         Left            =   1410
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1740
         Width           =   975
      End
      Begin VB.TextBox textTM35 
         Height          =   300
         Left            =   1410
         MaxLength       =   40
         TabIndex        =   0
         Top             =   570
         Width           =   2772
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "繳註冊費折扣：         %"
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   87
         Top             =   1230
         Width           =   1850
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展折扣：       %"
         Height          =   180
         Index           =   1
         Left            =   2970
         TabIndex        =   86
         Top             =   1230
         Width           =   1390
      End
      Begin MSForms.TextBox textTM71 
         Height          =   300
         Left            =   5790
         TabIndex        =   16
         Top             =   2680
         Width           =   2780
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4895;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM40 
         Height          =   300
         Left            =   -73440
         TabIndex        =   25
         Top             =   1708
         Width           =   7000
         VariousPropertyBits=   671105051
         Size            =   "12347;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM41 
         Height          =   300
         Left            =   -73440
         TabIndex        =   26
         Top             =   2045
         Width           =   3600
         VariousPropertyBits=   671105051
         Size            =   "6350;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM42 
         Height          =   300
         Left            =   -73440
         TabIndex        =   27
         Top             =   2382
         Width           =   4065
         VariousPropertyBits=   671105051
         Size            =   "7170;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM43 
         Height          =   300
         Left            =   -73440
         TabIndex        =   28
         Top             =   2719
         Width           =   7000
         VariousPropertyBits=   671105051
         Size            =   "12347;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM38 
         Height          =   300
         Left            =   -73440
         TabIndex        =   23
         Top             =   1034
         Width           =   3600
         VariousPropertyBits=   671105051
         Size            =   "6350;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM39 
         Height          =   300
         Left            =   -73440
         TabIndex        =   24
         Top             =   1371
         Width           =   4065
         VariousPropertyBits=   671105051
         Size            =   "7170;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM76 
         Height          =   300
         Left            =   -73440
         TabIndex        =   29
         Top             =   3060
         Width           =   7005
         VariousPropertyBits=   671105051
         Size            =   "12356;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM58 
         Height          =   1710
         Left            =   -74910
         TabIndex        =   85
         Top             =   810
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
      Begin MSForms.Label lblTM33 
         Height          =   260
         Left            =   2750
         TabIndex        =   84
         Top             =   2390
         Width           =   5870
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblTM33"
         Size            =   "10345;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblTM66 
         Height          =   260
         Left            =   2750
         TabIndex        =   83
         Top             =   3010
         Width           =   2410
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblTM66"
         Size            =   "4251;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblTM56 
         Height          =   260
         Left            =   2460
         TabIndex        =   82
         Top             =   1770
         Width           =   6110
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblTM56"
         Size            =   "10769;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblTM69 
         Height          =   260
         Left            =   5940
         TabIndex        =   81
         Top             =   1460
         Width           =   2620
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblTM69"
         Size            =   "4621;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblTM70 
         Height          =   260
         Left            =   2750
         TabIndex        =   80
         Top             =   3320
         Width           =   2410
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Caption         =   "lblTM70"
         Size            =   "4251;459"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblAgent 
         Height          =   255
         Left            =   -72270
         TabIndex        =   79
         Top             =   390
         Width           =   5880
         BackColor       =   -2147483644
         VariousPropertyBits=   27
         Size            =   "10372;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "彼所案號："
         Height          =   180
         Index           =   65
         Left            =   -74805
         TabIndex        =   75
         Top             =   748
         Width           =   900
      End
      Begin VB.Label Label68 
         AutoSize        =   -1  'True
         Caption         =   "以EMail通知：                (Y：是 D：僅D/N）"
         Height          =   180
         Left            =   5210
         TabIndex        =   74
         Top             =   3050
         Width           =   3390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email 同時寄紙本：        (Y：是)"
         Height          =   180
         Index           =   166
         Left            =   5190
         TabIndex        =   73
         Top             =   3360
         Width           =   2490
      End
      Begin VB.Label Label1 
         Caption         =   "請款單份數："
         Height          =   180
         Index           =   165
         Left            =   6510
         TabIndex        =   72
         Top             =   940
         Width           =   1080
      End
      Begin VB.Label Label1 
         Caption         =   "定稿份數："
         Height          =   180
         Index           =   164
         Left            =   4830
         TabIndex        =   71
         Top             =   940
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展單筆不跑 :           Y:不跑"
         Height          =   180
         Index           =   4
         Left            =   3810
         TabIndex        =   70
         Top             =   2110
         Width           =   2190
      End
      Begin VB.Label lblTM129 
         AutoSize        =   -1  'True
         Caption         =   "不催延展：        (Y:不催)"
         Height          =   180
         Left            =   6420
         TabIndex        =   69
         Top             =   2110
         Width           =   1910
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展請款對象："
         Height          =   180
         Index           =   84
         Left            =   150
         TabIndex        =   68
         Top             =   3050
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展彼所案號："
         Height          =   180
         Index           =   86
         Left            =   150
         TabIndex        =   67
         Top             =   2740
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展代理人："
         Height          =   180
         Index           =   64
         Left            =   290
         TabIndex        =   66
         Top             =   2430
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID："
         Height          =   180
         Index           =   169
         Left            =   4230
         TabIndex        =   65
         Top             =   630
         Width           =   1860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展聯絡人："
         Height          =   180
         Index           =   155
         Left            =   4680
         TabIndex        =   64
         Top             =   2740
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FCT註冊費自動代繳：        (Y:自動代繳)"
         Height          =   180
         Index           =   48
         Left            =   140
         TabIndex        =   63
         Top             =   2110
         Width           =   3120
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "延展D/N列印對象："
         Height          =   180
         Index           =   154
         Left            =   150
         TabIndex        =   62
         Top             =   3270
         Width           =   1550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N是否列印申請人：           (Y:是)"
         Height          =   180
         Index           =   47
         Left            =   80
         TabIndex        =   61
         Top             =   1490
         Width           =   2780
      End
      Begin VB.Label lblTM30 
         AutoSize        =   -1  'True
         Caption         =   "lblTM30"
         Height          =   180
         Left            =   1440
         TabIndex        =   60
         Top             =   330
         Width           =   620
      End
      Begin VB.Label lblTM57 
         AutoSize        =   -1  'True
         Caption         =   "lblTM57"
         Height          =   180
         Left            =   5520
         TabIndex        =   59
         Top             =   330
         Width           =   620
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "新FC代理人："
         Height          =   180
         Left            =   -74805
         TabIndex        =   57
         Top             =   420
         Width           =   1110
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(日)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   56
         Top             =   1762
         Width           =   1110
      End
      Begin VB.Label lblTM57_T 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日期："
         Height          =   180
         Left            =   4200
         TabIndex        =   55
         Top             =   330
         Width           =   1260
      End
      Begin VB.Label lblTM30_T 
         AutoSize        =   -1  'True
         Caption         =   "閉卷日期："
         ForeColor       =   &H00000000&
         Height          =   180
         Left            =   470
         TabIndex        =   54
         Top             =   330
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請/翻譯折扣：       %"
         Height          =   180
         Index           =   45
         Left            =   2550
         TabIndex        =   53
         Top             =   940
         Width           =   1760
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "固定請款對象："
         Height          =   180
         Index           =   34
         Left            =   140
         TabIndex        =   52
         Top             =   1800
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "全部折扣：         %"
         Height          =   180
         Index           =   32
         Left            =   500
         TabIndex        =   51
         Top             =   940
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "D/N固定列印對象："
         Height          =   180
         Index           =   55
         Left            =   3320
         TabIndex        =   50
         Top             =   1490
         Width           =   1550
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號："
         Height          =   180
         Index           =   62
         Left            =   140
         TabIndex        =   49
         Top             =   630
         Width           =   1260
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(中)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   48
         Top             =   1086
         Width           =   1110
      End
      Begin VB.Label Label51 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人1(英)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   47
         Top             =   1424
         Width           =   1110
      End
      Begin VB.Label Label55 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(中)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   46
         Top             =   2100
         Width           =   1110
      End
      Begin VB.Label Label57 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(英)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   45
         Top             =   2438
         Width           =   1110
      End
      Begin VB.Label Label59 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人2(日)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   44
         Top             =   2779
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門(日)："
         Height          =   180
         Left            =   -74805
         TabIndex        =   43
         Top             =   3120
         Width           =   1380
      End
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   240
      Left            =   3900
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   540
      Width           =   1770
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   240
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   540
      Width           =   1335
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   240
      Left            =   7080
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   540
      Width           =   1600
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6660
      TabIndex        =   32
      Top             =   60
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5700
      TabIndex        =   31
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7860
      TabIndex        =   33
      Top             =   60
      Width           =   912
   End
   Begin MSForms.TextBox textTM44 
      Height          =   300
      Left            =   1200
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   810
      Width           =   7500
      VariousPropertyBits=   671105055
      Size            =   "13229;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   300
      Left            =   1200
      TabIndex        =   77
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1200
      TabIndex        =   76
      Top             =   1490
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
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   180
      Left            =   120
      TabIndex        =   58
      Top             =   1230
      Width           =   900
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "原FC代理人："
      Height          =   180
      Left            =   120
      TabIndex        =   42
      Top             =   870
      Width           =   1110
   End
   Begin VB.Label lblTM12 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Left            =   2940
      TabIndex        =   41
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   540
      Width           =   900
   End
   Begin VB.Label lblTM15 
      AutoSize        =   -1  'True
      Caption         =   "審定號數："
      Height          =   180
      Left            =   6120
      TabIndex        =   39
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱："
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   1550
      Width           =   900
   End
End
Attribute VB_Name = "frm110104_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; textTM40~TM43、textTM38、textTM39、textTM76、textTM58、lblTM33、lblTM66、lblTM56、lblTM69、lblTM70、lblAgent、textTM44、textTM23、cmbTM05
'2011/3/29 新增 BY SONIA
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type
' 儲存基本檔檔案欄位的串列
Dim m_TMSPList() As FIELDITEM
Dim m_TMSPCount As Integer
'bolLeave判斷離開時，是否要彈出詢問視窗，回答Yes後改為True 跳下一畫面
Dim bolLeave As Boolean
Dim m_TM44 As String


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
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      frm110104_1.Show
      frm110104_1.Cleartxt
      Unload Me
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   MoveFormToCenter Me
   SSTab1.Tab = 0
   bolLeave = False
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      textTM58.Top = 360
      textTM58.Height = 2970
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
   Set frm110104_4 = Nothing
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
   End If
   
   Select Case nType
      ' 本所案號 欄位1
      Case 0: m_TM01 = strData
      ' 本所案號 欄位2
      Case 1: m_TM02 = strData
      ' 本所案號 欄位3
      Case 2: m_TM03 = strData
      ' 本所案號 欄位4
      Case 3: m_TM04 = strData
   End Select
End Sub

' 清除案件基本檔檔案欄位串列
Private Sub ClearTMSPFieldList()
   If m_TMSPCount > 0 Then
      Erase m_TMSPList
   End If
   m_TMSPCount = 0
End Sub

' 設定案件基本檔欄位串列中的欄位內容
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
      m_TMSPList(m_TMSPCount).fiType = nFieldType '0.文字 1.數字
      m_TMSPCount = m_TMSPCount + 1
   End If
End Sub

' 設定案件基本檔欄位串列中的欄位內容
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

' 取得商標基本檔的欄位內容
Private Sub QueryTradeMark()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String 'Add By Sindy 2014/12/2
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   cmbTM05.Clear
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then: textTM12 = rsTmp.Fields("TM12")
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then: cmbTM05.AddItem rsTmp.Fields("TM07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 原FC代理人
      m_TM44 = ""
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
         'Add By Sindy 2014/12/2
         Call GetAgentAndState(m_TM44, strTemp, , , True)
         'textTM44 = rsTmp.Fields("TM44") & " " & GetFAgentName(rsTmp.Fields("TM44"))
         textTM44 = rsTmp.Fields("TM44") & " " & strTemp
         '2014/12/2 END
      End If
      SetTMSPFieldOldData "TM44", m_TM44, 0
      ' 申請人1
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = rsTmp.Fields("TM23") & " " & GetAgentOrCustName(rsTmp.Fields("TM23"))
      ' 閉卷日期
      lblTM30 = ""
      If IsNull(rsTmp.Fields("TM30")) = False Then
         lblTM30 = ChangeWStringToTDateString(rsTmp.Fields("TM30"))
         lblTM30_T.ForeColor = &HC0&
         lblTM30.ForeColor = &HC0&
      End If
      ' 北所銷卷日期
      lblTM57 = ""
      If IsNull(rsTmp.Fields("TM57")) = False Then
         lblTM57 = ChangeWStringToTDateString(rsTmp.Fields("TM57"))
         lblTM57_T.ForeColor = &HC0&
         lblTM57.ForeColor = &HC0&
      End If
      
      'Modify By Sindy 2016/8/23 修改操作人員為'F1'字頭外商人員時,
      '                          除案件備註欄以外, 其他欄位都不要帶出來
      If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
         ' 客戶案件案號
         If IsNull(rsTmp.Fields("TM35")) = False Then
            SetTMSPFieldOldData "TM35", rsTmp.Fields("TM35"), 0
         Else
            SetTMSPFieldOldData "TM35", textTM35, 0
         End If
         ' 彼所案號
         If IsNull(rsTmp.Fields("TM45")) = False Then
            SetTMSPFieldOldData "TM45", rsTmp.Fields("TM45"), 0
         Else
            SetTMSPFieldOldData "TM45", textTM45, 0
         End If
         ' Client_Matter_id
         If IsNull(rsTmp.Fields("TM127")) = False Then
            SetTMSPFieldOldData "TM127", rsTmp.Fields("TM127"), 0
         Else
            SetTMSPFieldOldData "TM127", textTM127, 0
         End If
         ' 全部折扣
         If IsNull(rsTmp.Fields("TM36")) = False Then
            SetTMSPFieldOldData "TM36", rsTmp.Fields("TM36"), 1
         Else
            SetTMSPFieldOldData "TM36", textTM36, 1
         End If
         ' 申請/翻譯折扣
         If IsNull(rsTmp.Fields("TM37")) = False Then
            SetTMSPFieldOldData "TM37", rsTmp.Fields("TM37"), 1
         Else
            SetTMSPFieldOldData "TM37", textTM37, 1
         End If
         'Add By Sindy 2025/3/10
         ' 繳註冊費折扣
         If IsNull(rsTmp.Fields("TM140")) = False Then
            SetTMSPFieldOldData "TM140", rsTmp.Fields("TM140"), 1
         Else
            SetTMSPFieldOldData "TM140", textTM140, 1
         End If
         ' 延展折扣
         If IsNull(rsTmp.Fields("TM141")) = False Then
            SetTMSPFieldOldData "TM141", rsTmp.Fields("TM141"), 1
         Else
            SetTMSPFieldOldData "TM141", textTM141, 1
         End If
         '2025/3/10 END
         ' 定稿份數
         If IsNull(rsTmp.Fields("TM124")) = False Then
            SetTMSPFieldOldData "TM124", rsTmp.Fields("TM124"), 1
         Else
            SetTMSPFieldOldData "TM124", textTM124, 1
         End If
         ' 請款單份數
         If IsNull(rsTmp.Fields("TM125")) = False Then
            SetTMSPFieldOldData "TM125", rsTmp.Fields("TM125"), 1
         Else
            SetTMSPFieldOldData "TM125", textTM125, 1
         End If
         ' D/N是否列印申請人
         If IsNull(rsTmp.Fields("TM46")) = False Then
            SetTMSPFieldOldData "TM46", rsTmp.Fields("TM46"), 0
         Else
            SetTMSPFieldOldData "TM46", textTM46, 0
         End If
         ' D/N固定列印對象
         lblTM69 = ""
         If IsNull(rsTmp.Fields("TM69")) = False Then
            SetTMSPFieldOldData "TM69", rsTmp.Fields("TM69"), 0
         Else
            SetTMSPFieldOldData "TM69", textTM69, 0
         End If
         ' 固定請款對象
         lblTM56 = ""
         If IsNull(rsTmp.Fields("TM56")) = False Then
            SetTMSPFieldOldData "TM56", rsTmp.Fields("TM56"), 0
         Else
            SetTMSPFieldOldData "TM56", textTM56, 0
         End If
         ' FCT註冊費自動代繳
         If IsNull(rsTmp.Fields("TM122")) = False Then
            SetTMSPFieldOldData "TM122", rsTmp.Fields("TM122"), 0
         Else
            SetTMSPFieldOldData "TM122", textTM122, 0
         End If
         ' 延展單不跑
         If IsNull(rsTmp.Fields("TM68")) = False Then
            SetTMSPFieldOldData "TM68", rsTmp.Fields("TM68"), 0
         Else
            SetTMSPFieldOldData "TM68", textTM68, 0
         End If
         ' 不催延展
         If IsNull(rsTmp.Fields("TM129")) = False Then
            SetTMSPFieldOldData "TM129", rsTmp.Fields("TM129"), 0
         Else
            SetTMSPFieldOldData "TM129", textTM129, 0
         End If
         ' 延展代理人
         lblTM33 = ""
         If IsNull(rsTmp.Fields("TM33")) = False Then
            SetTMSPFieldOldData "TM33", rsTmp.Fields("TM33"), 0
         Else
            SetTMSPFieldOldData "TM33", textTM33, 0
         End If
         ' 延展彼所案號
         If IsNull(rsTmp.Fields("TM65")) = False Then
            SetTMSPFieldOldData "TM65", rsTmp.Fields("TM65"), 0
         Else
            SetTMSPFieldOldData "TM65", textTM65, 0
         End If
         ' 延展聯絡人
         If IsNull(rsTmp.Fields("TM71")) = False Then
            SetTMSPFieldOldData "TM71", rsTmp.Fields("TM71"), 0
         Else
            SetTMSPFieldOldData "TM71", textTM71, 0
         End If
         ' 延展請款對象
         lblTM66 = ""
         If IsNull(rsTmp.Fields("TM66")) = False Then
            SetTMSPFieldOldData "TM66", rsTmp.Fields("TM66"), 0
         Else
            SetTMSPFieldOldData "TM66", textTM66, 0
         End If
         ' 以EMail通知
         If IsNull(rsTmp.Fields("TM121")) = False Then
            SetTMSPFieldOldData "TM121", rsTmp.Fields("TM121"), 0
         Else
            SetTMSPFieldOldData "TM121", textTM121, 0
         End If
         ' 延展D/N列印對象
         lblTM70 = ""
         If IsNull(rsTmp.Fields("TM70")) = False Then
            SetTMSPFieldOldData "TM70", rsTmp.Fields("TM70"), 0
         Else
            SetTMSPFieldOldData "TM70", textTM70, 0
         End If
         ' EMail同時寄紙本
         If IsNull(rsTmp.Fields("TM126")) = False Then
            SetTMSPFieldOldData "TM126", rsTmp.Fields("TM126"), 0
         Else
            SetTMSPFieldOldData "TM126", textTM126, 0
         End If
         ' 聯絡人1(中)
         If IsNull(rsTmp.Fields("TM38")) = False Then
            SetTMSPFieldOldData "TM38", rsTmp.Fields("TM38"), 0
         Else
            SetTMSPFieldOldData "TM38", textTM38, 0
         End If
         ' 聯絡人1(英)
         If IsNull(rsTmp.Fields("TM39")) = False Then
            SetTMSPFieldOldData "TM39", rsTmp.Fields("TM39"), 0
         Else
            SetTMSPFieldOldData "TM39", textTM39, 0
         End If
         ' 聯絡人1(日)
         If IsNull(rsTmp.Fields("TM40")) = False Then
            SetTMSPFieldOldData "TM40", rsTmp.Fields("TM40"), 0
         Else
            SetTMSPFieldOldData "TM40", textTM40, 0
         End If
         ' 聯絡人2(中)
         If IsNull(rsTmp.Fields("TM41")) = False Then
            SetTMSPFieldOldData "TM41", rsTmp.Fields("TM41"), 0
         Else
            SetTMSPFieldOldData "TM41", textTM41, 0
         End If
         ' 聯絡人2(英)
         If IsNull(rsTmp.Fields("TM42")) = False Then
            SetTMSPFieldOldData "TM42", rsTmp.Fields("TM42"), 0
         Else
            SetTMSPFieldOldData "TM42", textTM42, 0
         End If
         ' 聯絡人2(日)
         If IsNull(rsTmp.Fields("TM43")) = False Then
            SetTMSPFieldOldData "TM43", rsTmp.Fields("TM43"), 0
         Else
            SetTMSPFieldOldData "TM43", textTM43, 0
         End If
         ' 聯絡人部門(日)
         If IsNull(rsTmp.Fields("TM76")) = False Then
            SetTMSPFieldOldData "TM76", rsTmp.Fields("TM76"), 0
         Else
            SetTMSPFieldOldData "TM76", textTM76, 0
         End If
      Else
      '2016/8/23 END
         ' 客戶案件案號
         If IsNull(rsTmp.Fields("TM35")) = False Then: textTM35 = rsTmp.Fields("TM35")
         SetTMSPFieldOldData "TM35", textTM35, 0
         ' 彼所案號
         If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
         SetTMSPFieldOldData "TM45", textTM45, 0
         ' Client_Matter_id
         If IsNull(rsTmp.Fields("TM127")) = False Then: textTM127 = rsTmp.Fields("TM127")
         SetTMSPFieldOldData "TM127", textTM127, 0
         ' 全部折扣
         If IsNull(rsTmp.Fields("TM36")) = False Then: textTM36 = rsTmp.Fields("TM36")
         SetTMSPFieldOldData "TM36", textTM36, 1
         ' 申請/翻譯折扣
         If IsNull(rsTmp.Fields("TM37")) = False Then: textTM37 = rsTmp.Fields("TM37")
         SetTMSPFieldOldData "TM37", textTM37, 1
         'Add By Sindy 2025/3/10
         ' 繳註冊費折扣
         If IsNull(rsTmp.Fields("TM140")) = False Then: textTM140 = rsTmp.Fields("TM140")
         SetTMSPFieldOldData "TM140", textTM140, 1
         ' 延展折扣
         If IsNull(rsTmp.Fields("TM141")) = False Then: textTM141 = rsTmp.Fields("TM141")
         SetTMSPFieldOldData "TM141", textTM141, 1
         '2025/3/10 END
         ' 定稿份數
         If IsNull(rsTmp.Fields("TM124")) = False Then: textTM124 = rsTmp.Fields("TM124")
         SetTMSPFieldOldData "TM124", textTM124, 1
         ' 請款單份數
         If IsNull(rsTmp.Fields("TM125")) = False Then: textTM125 = rsTmp.Fields("TM125")
         SetTMSPFieldOldData "TM125", textTM125, 1
         ' D/N是否列印申請人
         If IsNull(rsTmp.Fields("TM46")) = False Then: textTM46 = rsTmp.Fields("TM46")
         SetTMSPFieldOldData "TM46", textTM46, 0
         ' D/N固定列印對象
         lblTM69 = ""
         If IsNull(rsTmp.Fields("TM69")) = False Then: textTM69 = rsTmp.Fields("TM69"): lblTM69 = GetAgentOrCustName(rsTmp.Fields("TM69"))
         SetTMSPFieldOldData "TM69", textTM69, 0
         ' 固定請款對象
         lblTM56 = ""
         If IsNull(rsTmp.Fields("TM56")) = False Then: textTM56 = rsTmp.Fields("TM56"): lblTM56 = GetAgentOrCustName(rsTmp.Fields("TM56"))
         SetTMSPFieldOldData "TM56", textTM56, 0
         ' FCT註冊費自動代繳
         If IsNull(rsTmp.Fields("TM122")) = False Then: textTM122 = rsTmp.Fields("TM122")
         SetTMSPFieldOldData "TM122", textTM122, 0
         ' 延展單不跑
         If IsNull(rsTmp.Fields("TM68")) = False Then: textTM68 = rsTmp.Fields("TM68")
         SetTMSPFieldOldData "TM68", textTM68, 0
         ' 不催延展
         If IsNull(rsTmp.Fields("TM129")) = False Then: textTM129 = rsTmp.Fields("TM129")
         SetTMSPFieldOldData "TM129", textTM129, 0
         ' 延展代理人
         lblTM33 = ""
         If IsNull(rsTmp.Fields("TM33")) = False Then: textTM33 = rsTmp.Fields("TM33"): lblTM33 = GetAgentOrCustName(rsTmp.Fields("TM33"))
         SetTMSPFieldOldData "TM33", textTM33, 0
         ' 延展彼所案號
         If IsNull(rsTmp.Fields("TM65")) = False Then: textTM65 = rsTmp.Fields("TM65")
         SetTMSPFieldOldData "TM65", textTM65, 0
         ' 延展聯絡人
         If IsNull(rsTmp.Fields("TM71")) = False Then: textTM71 = rsTmp.Fields("TM71")
         SetTMSPFieldOldData "TM71", textTM71, 0
         ' 延展請款對象
         lblTM66 = ""
         If IsNull(rsTmp.Fields("TM66")) = False Then: textTM66 = rsTmp.Fields("TM66"): lblTM66 = GetAgentOrCustName(rsTmp.Fields("TM66"))
         SetTMSPFieldOldData "TM66", textTM66, 0
         ' 以EMail通知
         If IsNull(rsTmp.Fields("TM121")) = False Then: textTM121 = rsTmp.Fields("TM121")
         SetTMSPFieldOldData "TM121", textTM121, 0
         ' 延展D/N列印對象
         lblTM70 = ""
         If IsNull(rsTmp.Fields("TM70")) = False Then: textTM70 = rsTmp.Fields("TM70"): lblTM70 = GetAgentOrCustName(rsTmp.Fields("TM70"))
         SetTMSPFieldOldData "TM70", textTM70, 0
         ' EMail同時寄紙本
         If IsNull(rsTmp.Fields("TM126")) = False Then: textTM126 = rsTmp.Fields("TM126")
         SetTMSPFieldOldData "TM126", textTM126, 0
         ' 聯絡人1(中)
         If IsNull(rsTmp.Fields("TM38")) = False Then: textTM38 = rsTmp.Fields("TM38")
         SetTMSPFieldOldData "TM38", textTM38, 0
         ' 聯絡人1(英)
         If IsNull(rsTmp.Fields("TM39")) = False Then: textTM39 = rsTmp.Fields("TM39")
         SetTMSPFieldOldData "TM39", textTM39, 0
         ' 聯絡人1(日)
         If IsNull(rsTmp.Fields("TM40")) = False Then: textTM40 = rsTmp.Fields("TM40")
         SetTMSPFieldOldData "TM40", textTM40, 0
         ' 聯絡人2(中)
         If IsNull(rsTmp.Fields("TM41")) = False Then: textTM41 = rsTmp.Fields("TM41")
         SetTMSPFieldOldData "TM41", textTM41, 0
         ' 聯絡人2(英)
         If IsNull(rsTmp.Fields("TM42")) = False Then: textTM42 = rsTmp.Fields("TM42")
         SetTMSPFieldOldData "TM42", textTM42, 0
         ' 聯絡人2(日)
         If IsNull(rsTmp.Fields("TM43")) = False Then: textTM43 = rsTmp.Fields("TM43")
         SetTMSPFieldOldData "TM43", textTM43, 0
         ' 聯絡人部門(日)
         If IsNull(rsTmp.Fields("TM76")) = False Then: textTM76 = rsTmp.Fields("TM76")
         SetTMSPFieldOldData "TM76", textTM76, 0
      End If
      
      ' 備註
      If IsNull(rsTmp.Fields("TM58")) = False Then: textTM58 = rsTmp.Fields("TM58")
      SetTMSPFieldOldData "TM58", textTM58, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得法務基本檔的欄位內容
Private Sub QueryLawCase()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String 'Add By Sindy 2014/12/2
   
   strSql = "SELECT * FROM LawCase " & _
            "WHERE LC01 = '" & m_TM01 & "' AND " & _
                  "LC02 = '" & m_TM02 & "' AND " & _
                  "LC03 = '" & m_TM03 & "' AND " & _
                  "LC04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   cmbTM05.Clear
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請案號
      lblTM12.Visible = False
      textTM12.Visible = False
      ' 審定號數
      lblTM15.Visible = False
      textTM15.Visible = False
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("LC05")) = False Then: cmbTM05.AddItem rsTmp.Fields("LC05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("LC06")) = False Then: cmbTM05.AddItem rsTmp.Fields("LC06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("LC07")) = False Then: cmbTM05.AddItem rsTmp.Fields("LC07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 原FC代理人
      m_TM44 = ""
      If IsNull(rsTmp.Fields("LC22")) = False Then
         m_TM44 = rsTmp.Fields("LC22")
         'Add By Sindy 2014/12/2
         Call GetAgentAndState(m_TM44, strTemp, , , True)
         'textTM44 = rsTmp.Fields("LC22") & " " & GetFAgentName(rsTmp.Fields("LC22"))
         textTM44 = rsTmp.Fields("LC22") & " " & strTemp
         '2014/12/2 END
      End If
      SetTMSPFieldOldData "LC22", m_TM44, 0
      ' 申請人1
      If IsNull(rsTmp.Fields("LC11")) = False Then: textTM23 = rsTmp.Fields("LC11") & " " & GetAgentOrCustName(rsTmp.Fields("LC11"))
      ' 閉卷日期
      lblTM30 = ""
      If IsNull(rsTmp.Fields("LC09")) = False Then
         lblTM30 = ChangeWStringToTDateString(rsTmp.Fields("LC09"))
         lblTM30_T.ForeColor = "&H000000C0&"
         lblTM30.ForeColor = "&H000000C0&"
      End If
      ' 北所銷卷日期
      lblTM57 = ""
      If IsNull(rsTmp.Fields("LC34")) = False Then
         lblTM57 = ChangeWStringToTDateString(rsTmp.Fields("LC34"))
         lblTM57_T.ForeColor = "&H000000C0&"
         lblTM57.ForeColor = "&H000000C0&"
      End If
      ' 客戶案件案號
      If IsNull(rsTmp.Fields("LC17")) = False Then: textTM35 = rsTmp.Fields("LC17")
      SetTMSPFieldOldData "LC17", textTM35, 0
      ' 彼所案號
      If IsNull(rsTmp.Fields("LC23")) = False Then: textTM45 = rsTmp.Fields("LC23")
      SetTMSPFieldOldData "LC23", textTM45, 0
      ' Client_Matter_id
      textTM127.Enabled = False
      ' 全部折扣
      If IsNull(rsTmp.Fields("LC24")) = False Then: textTM36 = rsTmp.Fields("LC24")
      SetTMSPFieldOldData "LC24", textTM36, 1
      ' 申請/翻譯折扣
      textTM37.Enabled = False
      'Add By Sindy 2025/3/10
      ' 繳註冊費折扣
      textTM140.Enabled = False
      ' 延展折扣
      textTM141.Enabled = False
      '2025/3/10 END
      ' 定稿份數
      textTM124.Enabled = False
      ' 請款單份數
      textTM125.Enabled = False
      ' D/N是否列印申請人
      If IsNull(rsTmp.Fields("LC25")) = False Then: textTM46 = rsTmp.Fields("LC25")
      SetTMSPFieldOldData "LC25", textTM46, 0
      ' D/N固定列印對象
      lblTM69 = ""
      If IsNull(rsTmp.Fields("LC35")) = False Then: textTM69 = rsTmp.Fields("LC35"): lblTM69 = GetAgentOrCustName(rsTmp.Fields("LC35"))
      SetTMSPFieldOldData "LC35", textTM69, 0
      ' 固定請款對象
      lblTM56 = ""
      If IsNull(rsTmp.Fields("LC26")) = False Then: textTM56 = rsTmp.Fields("LC26"): lblTM56 = GetAgentOrCustName(rsTmp.Fields("LC26"))
      SetTMSPFieldOldData "LC26", textTM56, 0
      ' FCT註冊費自動代繳
      textTM122.Enabled = False
      ' 延展單不跑
      textTM68.Enabled = False
      ' 不催延展
      textTM129.Enabled = False
      ' 延展代理人
      lblTM33 = ""
      textTM33.Enabled = False
      ' 延展彼所案號
      textTM65.Enabled = False
      ' 延展聯絡人
      textTM71.Enabled = False
      ' 延展請款對象
      lblTM66 = ""
      textTM66.Enabled = False
      ' 以EMail通知
      textTM121.Enabled = False
      ' 延展D/N列印對象
      lblTM70 = ""
      textTM70.Enabled = False
      ' EMail同時寄紙本
      textTM126.Enabled = False
      ' 聯絡人1(中)
      If IsNull(rsTmp.Fields("LC18")) = False Then: textTM38 = rsTmp.Fields("LC18")
      SetTMSPFieldOldData "LC18", textTM38, 0
      ' 聯絡人1(英)
      If IsNull(rsTmp.Fields("LC19")) = False Then: textTM39 = rsTmp.Fields("LC19")
      SetTMSPFieldOldData "LC19", textTM39, 0
      ' 聯絡人1(日)
      If IsNull(rsTmp.Fields("LC20")) = False Then: textTM40 = rsTmp.Fields("LC20")
      SetTMSPFieldOldData "LC20", textTM40, 0
      ' 聯絡人2(中)
      textTM41.Enabled = False
      ' 聯絡人2(英)
      textTM42.Enabled = False
      ' 聯絡人2(日)
      textTM43.Enabled = False
      ' 聯絡人部門(日)
      If IsNull(rsTmp.Fields("LC39")) = False Then: textTM76 = rsTmp.Fields("LC39")
      SetTMSPFieldOldData "LC39", textTM76, 0
      ' 備註
      If IsNull(rsTmp.Fields("LC27")) = False Then: textTM58 = rsTmp.Fields("LC27")
      SetTMSPFieldOldData "LC27", textTM58, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務基本檔的欄位內容
Private Sub QueryServicePractice()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String 'Add By Sindy 2014/12/2
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   cmbTM05.Clear
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then: textTM12 = rsTmp.Fields("SP11")
      ' 審定號數
      lblTM15.Visible = False
      textTM15.Visible = False
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP05")
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP06")
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then: cmbTM05.AddItem rsTmp.Fields("SP07")
      ' 顯示案件名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 原FC代理人
      m_TM44 = ""
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
         'Add By Sindy 2014/12/2
         Call GetAgentAndState(m_TM44, strTemp, , , True)
         'textTM44 = rsTmp.Fields("SP26") & " " & GetFAgentName(rsTmp.Fields("SP26"))
         textTM44 = rsTmp.Fields("SP26") & " " & strTemp
         '2014/12/2 END
      End If
      SetTMSPFieldOldData "SP26", m_TM44, 0
      ' 申請人1
      If IsNull(rsTmp.Fields("SP08")) = False Then: textTM23 = rsTmp.Fields("SP08") & " " & GetAgentOrCustName(rsTmp.Fields("SP08"))
      ' 閉卷日期
      lblTM30 = ""
      If IsNull(rsTmp.Fields("SP16")) = False Then
         lblTM30 = ChangeWStringToTDateString(rsTmp.Fields("SP16"))
         lblTM30_T.ForeColor = "&H000000C0&"
         lblTM30.ForeColor = "&H000000C0&"
      End If
      ' 北所銷卷日期
      lblTM57 = ""
      If IsNull(rsTmp.Fields("SP61")) = False Then
         lblTM57 = ChangeWStringToTDateString(rsTmp.Fields("SP61"))
         lblTM57_T.ForeColor = "&H000000C0&"
         lblTM57.ForeColor = "&H000000C0&"
      End If
      
      'Modify By Sindy 2016/10/20 修改操作人員為'F1'字頭外商人員時,
      '                          除案件備註欄以外, 其他欄位都不要帶出來
      If Left(Trim(Pub_StrUserSt03), 2) = "F1" Then
         ' 客戶案件案號
         If IsNull(rsTmp.Fields("SP29")) = False Then
            SetTMSPFieldOldData "SP29", rsTmp.Fields("SP29"), 0
         Else
            SetTMSPFieldOldData "SP29", textTM35, 0
         End If
         ' 彼所案號
         If IsNull(rsTmp.Fields("SP27")) = False Then
            SetTMSPFieldOldData "SP27", rsTmp.Fields("SP27"), 0
         Else
            SetTMSPFieldOldData "SP27", textTM45, 0
         End If
         ' Client_Matter_id
         If IsNull(rsTmp.Fields("SP84")) = False Then
            SetTMSPFieldOldData "SP84", rsTmp.Fields("SP84"), 0
         Else
            SetTMSPFieldOldData "SP84", textTM127, 0
         End If
         ' 全部折扣
         If IsNull(rsTmp.Fields("SP31")) = False Then
            SetTMSPFieldOldData "SP31", rsTmp.Fields("SP31"), 1
         Else
            SetTMSPFieldOldData "SP31", textTM36, 1
         End If
         ' 申請/翻譯折扣
         textTM37.Enabled = False
         'Add By Sindy 2025/3/10
         ' 繳註冊費折扣
         textTM140.Enabled = False
         ' 延展折扣
         textTM141.Enabled = False
         '2025/3/10 END
         ' 定稿份數
         If IsNull(rsTmp.Fields("SP81")) = False Then
            SetTMSPFieldOldData "SP81", rsTmp.Fields("SP81"), 1
         Else
            SetTMSPFieldOldData "SP81", textTM124, 1
         End If
         ' 請款單份數
         If IsNull(rsTmp.Fields("SP82")) = False Then
            SetTMSPFieldOldData "SP82", rsTmp.Fields("SP82"), 1
         Else
            SetTMSPFieldOldData "SP82", textTM125, 1
         End If
         ' D/N是否列印申請人
         If IsNull(rsTmp.Fields("SP33")) = False Then
            SetTMSPFieldOldData "SP33", rsTmp.Fields("SP33"), 0
         Else
            SetTMSPFieldOldData "SP33", textTM46, 0
         End If
         ' D/N固定列印對象
         lblTM69 = ""
         If IsNull(rsTmp.Fields("SP67")) = False Then
            SetTMSPFieldOldData "SP67", rsTmp.Fields("SP67"), 0
         Else
            SetTMSPFieldOldData "SP67", textTM69, 0
         End If
         ' 固定請款對象
         lblTM56 = ""
         If IsNull(rsTmp.Fields("SP37")) = False Then
            SetTMSPFieldOldData "SP37", rsTmp.Fields("SP37"), 0
         Else
            SetTMSPFieldOldData "SP37", textTM56, 0
         End If
         ' FCT註冊費自動代繳
         textTM122.Enabled = False
         ' 延展單不跑
         textTM68.Enabled = False
         ' 不催延展
         textTM129.Enabled = False
         ' 延展代理人
         lblTM33 = ""
         textTM33.Enabled = False
         ' 延展彼所案號
         textTM65.Enabled = False
         ' 延展聯絡人
         textTM71.Enabled = False
         ' 延展請款對象
         lblTM66 = ""
         textTM66.Enabled = False
         ' 以EMail通知
         If IsNull(rsTmp.Fields("SP80")) = False Then
            SetTMSPFieldOldData "SP80", rsTmp.Fields("SP80"), 0
         Else
            SetTMSPFieldOldData "SP80", textTM121, 0
         End If
         ' 延展D/N列印對象
         lblTM70 = ""
         textTM70.Enabled = False
         ' EMail同時寄紙本
         If IsNull(rsTmp.Fields("SP83")) = False Then
            SetTMSPFieldOldData "SP83", rsTmp.Fields("SP83"), 0
         Else
            SetTMSPFieldOldData "SP83", textTM126, 0
         End If
         ' 聯絡人1(中)
         If IsNull(rsTmp.Fields("SP30")) = False Then
            SetTMSPFieldOldData "SP30", rsTmp.Fields("SP30"), 0
         Else
            SetTMSPFieldOldData "SP30", textTM38, 0
         End If
         ' 聯絡人1(英)
         textTM39.Enabled = False
         ' 聯絡人1(日)
         textTM40.Enabled = False
         ' 聯絡人2(中)
         If IsNull(rsTmp.Fields("SP75")) = False Then
            SetTMSPFieldOldData "SP75", rsTmp.Fields("SP75"), 0
         Else
            SetTMSPFieldOldData "SP75", textTM41, 0
         End If
         ' 聯絡人2(英)
         textTM42.Enabled = False
         ' 聯絡人2(日)
         textTM43.Enabled = False
         ' 聯絡人部門(日)
         If IsNull(rsTmp.Fields("SP71")) = False Then
            SetTMSPFieldOldData "SP71", rsTmp.Fields("SP71"), 0
         Else
            SetTMSPFieldOldData "SP71", textTM76, 0
         End If
      Else
      '2016/10/20 END
         ' 客戶案件案號
         If IsNull(rsTmp.Fields("SP29")) = False Then: textTM35 = rsTmp.Fields("SP29")
         SetTMSPFieldOldData "SP29", textTM35, 0
         ' 彼所案號
         If IsNull(rsTmp.Fields("SP27")) = False Then: textTM45 = rsTmp.Fields("SP27")
         SetTMSPFieldOldData "SP27", textTM45, 0
         ' Client_Matter_id
         If IsNull(rsTmp.Fields("SP84")) = False Then: textTM127 = rsTmp.Fields("SP84")
         SetTMSPFieldOldData "SP84", textTM127, 0
         ' 全部折扣
         If IsNull(rsTmp.Fields("SP31")) = False Then: textTM36 = rsTmp.Fields("SP31")
         SetTMSPFieldOldData "SP31", textTM36, 1
         ' 申請/翻譯折扣
         textTM37.Enabled = False
         'Add By Sindy 2025/3/10
         ' 繳註冊費折扣
         textTM140.Enabled = False
         ' 延展折扣
         textTM141.Enabled = False
         '2025/3/10 END
         ' 定稿份數
         If IsNull(rsTmp.Fields("SP81")) = False Then: textTM124 = rsTmp.Fields("SP81")
         SetTMSPFieldOldData "SP81", textTM124, 1
         ' 請款單份數
         If IsNull(rsTmp.Fields("SP82")) = False Then: textTM125 = rsTmp.Fields("SP82")
         SetTMSPFieldOldData "SP82", textTM125, 1
         ' D/N是否列印申請人
         If IsNull(rsTmp.Fields("SP33")) = False Then: textTM46 = rsTmp.Fields("SP33")
         SetTMSPFieldOldData "SP33", textTM46, 0
         ' D/N固定列印對象
         lblTM69 = ""
         If IsNull(rsTmp.Fields("SP67")) = False Then: textTM69 = rsTmp.Fields("SP67"): lblTM69 = GetAgentOrCustName(rsTmp.Fields("SP67"))
         SetTMSPFieldOldData "SP67", textTM69, 0
         ' 固定請款對象
         lblTM56 = ""
         If IsNull(rsTmp.Fields("SP37")) = False Then: textTM56 = rsTmp.Fields("SP37"): lblTM56 = GetAgentOrCustName(rsTmp.Fields("SP37"))
         SetTMSPFieldOldData "SP37", textTM56, 0
         ' FCT註冊費自動代繳
         textTM122.Enabled = False
         ' 延展單不跑
         textTM68.Enabled = False
         ' 不催延展
         textTM129.Enabled = False
         ' 延展代理人
         lblTM33 = ""
         textTM33.Enabled = False
         ' 延展彼所案號
         textTM65.Enabled = False
         ' 延展聯絡人
         textTM71.Enabled = False
         ' 延展請款對象
         lblTM66 = ""
         textTM66.Enabled = False
         ' 以EMail通知
         If IsNull(rsTmp.Fields("SP80")) = False Then: textTM121 = rsTmp.Fields("SP80")
         SetTMSPFieldOldData "SP80", textTM121, 0
         ' 延展D/N列印對象
         lblTM70 = ""
         textTM70.Enabled = False
         ' EMail同時寄紙本
         If IsNull(rsTmp.Fields("SP83")) = False Then: textTM126 = rsTmp.Fields("SP83")
         SetTMSPFieldOldData "SP83", textTM126, 0
         ' 聯絡人1(中)
         If IsNull(rsTmp.Fields("SP30")) = False Then: textTM38 = rsTmp.Fields("SP30")
         SetTMSPFieldOldData "SP30", textTM38, 0
         ' 聯絡人1(英)
         textTM39.Enabled = False
         ' 聯絡人1(日)
         textTM40.Enabled = False
         ' 聯絡人2(中)
         If IsNull(rsTmp.Fields("SP75")) = False Then: textTM41 = rsTmp.Fields("SP75")
         SetTMSPFieldOldData "SP75", textTM41, 0
         ' 聯絡人2(英)
         textTM42.Enabled = False
         ' 聯絡人2(日)
         textTM43.Enabled = False
         ' 聯絡人部門(日)
         If IsNull(rsTmp.Fields("SP71")) = False Then: textTM76 = rsTmp.Fields("SP71")
         SetTMSPFieldOldData "SP71", textTM76, 0
      End If
      
      ' 備註
      If IsNull(rsTmp.Fields("SP18")) = False Then: textTM58 = rsTmp.Fields("SP18")
      SetTMSPFieldOldData "SP18", textTM58, 0
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
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   ' 讀取基本檔
   If ClsPDGetSystemKind(m_TM01, intCaseKind) Then
      Select Case intCaseKind
         Case 商標
            QueryTradeMark
         Case 法務
            QueryLawCase
         Case Else '服務
            QueryServicePractice
      End Select
   End If
   SSTab1.Tab = 1 'Add By Sindy 2019/12/11
   NewFagent.SetFocus 'Add By Sindy 2019/12/11
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
Dim intCaseKind As Integer
   
   NewFagent = NewFagent & String(9 - Len(NewFagent), "0")
   'Modified by Sindy 2018/1/24 備註加 ChgSQL(代理人名稱可能有單引號)
   textTM58 = ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & m_TM44 & "/" & Mid(Trim(textTM44), 11)) & ";" & textTM58
   ' 更新基本檔
   If ClsPDGetSystemKind(m_TM01, intCaseKind) Then
      Select Case intCaseKind
         Case 商標
            ' FC代理人
            SetTMSPFieldNewData "TM44", NewFagent
            ' 客戶案件案號
            SetTMSPFieldNewData "TM35", textTM35
            ' 彼所案號
            SetTMSPFieldNewData "TM45", textTM45
            ' Client_Matter_id
            SetTMSPFieldNewData "TM127", textTM127
            ' 全部折扣
            SetTMSPFieldNewData "TM36", textTM36
            ' 申請/翻譯折扣
            SetTMSPFieldNewData "TM37", textTM37
            'Add By Sindy 2025/3/10
            ' 繳註冊費折扣
            SetTMSPFieldNewData "TM140", textTM140
            ' 延展折扣
            SetTMSPFieldNewData "TM141", textTM141
            '2025/3/10 END
            ' 定稿份數
            SetTMSPFieldNewData "TM124", textTM124
            ' 請款單份數
            SetTMSPFieldNewData "TM125", textTM125
            ' D/N是否列印申請人
            SetTMSPFieldNewData "TM46", textTM46
            ' D/N固定列印對象
            SetTMSPFieldNewData "TM69", textTM69 & IIf(textTM69 <> "", String(9 - Len(textTM69), "0"), "")
            ' 固定請款對象
            SetTMSPFieldNewData "TM56", textTM56 & IIf(textTM56 <> "", String(9 - Len(textTM56), "0"), "")
            ' FCT註冊費自動代繳
            SetTMSPFieldNewData "TM122", textTM122
            ' 延展單不跑
            SetTMSPFieldNewData "TM68", textTM68
            ' 不催延展
            SetTMSPFieldNewData "TM129", textTM129
            ' 延展代理人
            SetTMSPFieldNewData "TM33", textTM33 & IIf(textTM33 <> "", String(9 - Len(textTM33), "0"), "")
            ' 延展彼所案號
            SetTMSPFieldNewData "TM65", textTM65
            ' 延展聯絡人
            SetTMSPFieldNewData "TM71", textTM71
            ' 延展請款對象
            SetTMSPFieldNewData "TM66", textTM66 & IIf(textTM66 <> "", String(9 - Len(textTM66), "0"), "")
            ' 以EMail通知
            SetTMSPFieldNewData "TM121", textTM121
            ' 延展D/N列印對象
            SetTMSPFieldNewData "TM70", textTM70 & IIf(textTM70 <> "", String(9 - Len(textTM70), "0"), "")
            ' EMail同時寄紙本
            SetTMSPFieldNewData "TM126", textTM126
            ' 聯絡人1(中)
            SetTMSPFieldNewData "TM38", textTM38
            ' 聯絡人1(英)
            SetTMSPFieldNewData "TM39", textTM39
            ' 聯絡人1(日)
            SetTMSPFieldNewData "TM40", textTM40
            ' 聯絡人2(中)
            SetTMSPFieldNewData "TM41", textTM41
            ' 聯絡人2(英)
            SetTMSPFieldNewData "TM42", textTM42
            ' 聯絡人2(日)
            SetTMSPFieldNewData "TM43", textTM43
            ' 聯絡人部門(日)
            SetTMSPFieldNewData "TM76", textTM76
            ' 備註
            SetTMSPFieldNewData "TM58", textTM58
         Case 法務
            ' FC代理人
            SetTMSPFieldNewData "LC22", NewFagent
            ' 客戶案件案號
            SetTMSPFieldNewData "LC17", textTM35
            ' 彼所案號
            SetTMSPFieldNewData "LC23", textTM45
            ' 全部折扣
            SetTMSPFieldNewData "LC24", textTM36
            ' D/N是否列印申請人
            SetTMSPFieldNewData "LC25", textTM46
            ' D/N固定列印對象
            SetTMSPFieldNewData "LC35", textTM69 & IIf(textTM69 <> "", String(9 - Len(textTM69), "0"), "")
            ' 固定請款對象
            SetTMSPFieldNewData "LC26", textTM56 & IIf(textTM56 <> "", String(9 - Len(textTM56), "0"), "")
            ' 聯絡人1(中)
            SetTMSPFieldNewData "LC18", textTM38
            ' 聯絡人1(英)
            SetTMSPFieldNewData "LC19", textTM39
            ' 聯絡人1(日)
            SetTMSPFieldNewData "LC20", textTM40
            ' 聯絡人部門(日)
            SetTMSPFieldNewData "LC39", textTM76
            ' 備註
            SetTMSPFieldNewData "LC27", textTM58
         Case Else '服務
            ' FC代理人
            SetTMSPFieldNewData "SP26", NewFagent
            ' 客戶案件案號
            SetTMSPFieldNewData "SP29", textTM35
            ' 彼所案號
            SetTMSPFieldNewData "SP27", textTM45
            ' Client_Matter_id
            SetTMSPFieldNewData "SP84", textTM127
            ' 全部折扣
            SetTMSPFieldNewData "SP31", textTM36
            ' 定稿份數
            SetTMSPFieldNewData "SP81", textTM124
            ' 請款單份數
            SetTMSPFieldNewData "SP82", textTM125
            ' D/N是否列印申請人
            SetTMSPFieldNewData "SP33", textTM46
            ' D/N固定列印對象
            SetTMSPFieldNewData "SP67", textTM69 & IIf(textTM69 <> "", String(9 - Len(textTM69), "0"), "")
            ' 固定請款對象
            SetTMSPFieldNewData "SP37", textTM56 & IIf(textTM56 <> "", String(9 - Len(textTM56), "0"), "")
            ' 以EMail通知
            SetTMSPFieldNewData "SP80", textTM121
            ' EMail同時寄紙本
            SetTMSPFieldNewData "SP83", textTM126
            ' 聯絡人1(中)
            SetTMSPFieldNewData "SP30", textTM38
            ' 聯絡人2(中)
            SetTMSPFieldNewData "SP75", textTM41
            ' 聯絡人部門(日)
            SetTMSPFieldNewData "SP71", textTM76
            ' 備註
            SetTMSPFieldNewData "SP18", textTM58
      End Select
   End If
End Sub

Public Function OnSaveData() As Boolean
Dim intCaseKind As Integer
Dim bFirst As Boolean
Dim nIndex As Integer
Dim strTmp As String, strCP09 As String, strCP110 As String, strCP10 As String
Dim strMCTF(0) As String, stMsg As String 'Add by Amy 2019/06/26
   
On Error GoTo CheckingErr
   
   OnSaveData = True
   'Add by Amy 2019/06/26 取得新代理人之控管智權人員
   strExc(0) = GetCusORFagentData(ChangeCustomerL(NewFagent), "FA120", strMCTF())
   cnnConnection.BeginTrans
   
   ' 更新基本檔
   If ClsPDGetSystemKind(m_TM01, intCaseKind) Then
      Select Case intCaseKind
         Case 商標
            strCP10 = "726"
            strSql = "UPDATE TradeMark SET "
            bFirst = True
            For nIndex = 0 To m_TMSPCount - 1
               strTmp = Empty
               If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
                  If m_TMSPList(nIndex).fiType = 0 Then
                     strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
                  Else
                     If m_TMSPList(nIndex).fiNewData = Empty Then
                        strTmp = m_TMSPList(nIndex).fiName & " = " & 0
                     Else
                        strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                          " WHERE TM01 = '" & m_TM01 & "' AND " & _
                                 "TM02 = '" & m_TM02 & "' AND " & _
                                 "TM03 = '" & m_TM03 & "' AND " & _
                                 "TM04 = '" & m_TM04 & "' "
         Case 法務
            strCP10 = "994"
            strSql = "UPDATE LawCase SET "
            bFirst = True
            For nIndex = 0 To m_TMSPCount - 1
               strTmp = Empty
               If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
                  If m_TMSPList(nIndex).fiType = 0 Then
                     strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
                  Else
                     If m_TMSPList(nIndex).fiNewData = Empty Then
                        strTmp = m_TMSPList(nIndex).fiName & " = " & 0
                     Else
                        strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                          " WHERE LC01 = '" & m_TM01 & "' AND " & _
                                 "LC02 = '" & m_TM02 & "' AND " & _
                                 "LC03 = '" & m_TM03 & "' AND " & _
                                 "LC04 = '" & m_TM04 & "' "
         Case Else '服務
            If m_TM01 = "FG" Or _
               m_TM01 = "PS" Or _
               m_TM01 = "CPS" Then
               strCP10 = "937"
            Else
               strCP10 = "726"
            End If
            strSql = "UPDATE ServicePractice SET "
            bFirst = True
            For nIndex = 0 To m_TMSPCount - 1
               strTmp = Empty
               If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
                  If m_TMSPList(nIndex).fiType = 0 Then
                     strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
                  Else
                     If m_TMSPList(nIndex).fiNewData = Empty Then
                        strTmp = m_TMSPList(nIndex).fiName & " = " & 0
                     Else
                        strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                          " WHERE SP01 = '" & m_TM01 & "' AND " & _
                                 "SP02 = '" & m_TM02 & "' AND " & _
                                 "SP03 = '" & m_TM03 & "' AND " & _
                                 "SP04 = '" & m_TM04 & "' "
      End Select
   End If
   'Add By Sindy 2017/3/14 紀錄分析語法
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   
   '新增案件進度檔
   strCP09 = AutoNo("B", 6)
   '取得出名代理人
   strCP110 = ""
'CANCEL BY SONIA 2015/6/17 FCT-024182各式申請書抓最新A,B類發文之CP110會抓到此進度
'   strExc(0) = "select cp110 from caseprogress" & _
'               " where cp09=(select substr(max(cp27||cp09),9) from caseprogress" & _
'               " WHERE cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'" & _
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
   strSql = "INSERT INTO CASEPROGRESS(CP09,CP01,CP02,CP03,CP04,CP05,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP27,CP64,cp82,cp83,cp110)" & _
            " values(" & CNULL(strCP09) & "," & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & "," & CNULL(m_TM04) & _
            "," & strSrvDate(1) & "," & CNULL(strCP10) & ",'90'," & CNULL(PUB_GetStaffST15(frm110104_1.txtCaseField(4), 1)) & "," & CNULL(frm110104_1.txtCaseField(4)) & "," & CNULL(strUserNum) & ",'N','N'" & _
            "," & strSrvDate(1) & ",'" & ChgSQL(ChangeTStringToTDateString(strSrvDate(2)) & "換FC代理人,請留意最新指示及聯絡對象,原FC代理人" & m_TM44 & "/" & ChgSQL(Mid(Trim(textTM44), 11))) & ";'" & _
            ",substr(to_char(sysdate,'yyyymmddhh24mmss'),9)," & CNULL(strUserNum) & "," & CNULL(strCP110) & ")"
   cnnConnection.Execute strSql
   
   'Add by Amy 2019/06/26 新代理人之管控智權人員為MCTF,依案號更新當日AB類收文之 收文MCTF組別
   If InStr(strMCTF(0), "MCTF") > 0 Then
        stMsg = "Y"
        If UpdCP161(m_TM01 & ";" & m_TM02 & ";" & m_TM03 & ";" & m_TM04, strMCTF(0), stMsg) = False Then GoTo CheckingErr
   End If
   cnnConnection.CommitTrans
   bolLeave = True
   Exit Function
   
CheckingErr:
   'Modify by Amy 2019/06/26
   'MsgBox (Err.Description)
   If stMsg = MsgText(601) Then stMsg = Err.Description
   MsgBox stMsg
   'end 2019/06/
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
   
   If m_TM44 = NewFagent Then
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
   
   If Me.textTM46.Enabled = True Then
      Cancel = False
      textTM46_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM46.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM69.Enabled = True Then
      Cancel = False
      textTM69_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM69.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM56.Enabled = True Then
      Cancel = False
      textTM56_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM56.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM122.Enabled = True Then
      Cancel = False
      textTM122_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM122.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM68.Enabled = True Then
      Cancel = False
      textTM68_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM68.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM129.Enabled = True Then
      Cancel = False
      textTM129_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM129.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM33.Enabled = True Then
      Cancel = False
      textTM33_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM33.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM66.Enabled = True Then
      Cancel = False
      textTM66_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM66.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM121.Enabled = True Then
      Cancel = False
      textTM121_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM121.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM70.Enabled = True Then
      Cancel = False
      textTM70_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM70.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM126.Enabled = True Then
      Cancel = False
      textTM126_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textTM126.SetFocus
         Exit Function
      End If
   End If
   If Me.NewFagent.Enabled = True Then
      Cancel = False
      NewFagent_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         NewFagent.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM38.Enabled = True Then
      Cancel = False
      textTM38_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         textTM38.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM40.Enabled = True Then
      Cancel = False
      textTM40_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         textTM40.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM41.Enabled = True Then
      Cancel = False
      textTM41_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         textTM41.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM43.Enabled = True Then
      Cancel = False
      textTM43_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         textTM43.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM76.Enabled = True Then
      Cancel = False
      textTM76_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 1
         textTM76.SetFocus
         Exit Function
      End If
   End If
   If Me.textTM58.Enabled = True Then
      Cancel = False
      textTM58_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 2
         textTM58.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Lydia 2021/09/22 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub textTM35_GotFocus()
   InverseTextBox textTM35
End Sub

Private Sub textTM45_GotFocus()
   InverseTextBox textTM45
End Sub

Private Sub textTM127_GotFocus()
   TextInverse textTM127
   CloseIme
End Sub

Private Sub textTM36_GotFocus()
   InverseTextBox textTM36
End Sub

Private Sub textTM37_GotFocus()
   InverseTextBox textTM37
End Sub

'Add By Sindy 2025/3/10
Private Sub textTM140_GotFocus()
   InverseTextBox textTM140
End Sub
Private Sub textTM141_GotFocus()
   InverseTextBox textTM141
End Sub
'2025/3/10 END

Private Sub textTM124_GotFocus()
   InverseTextBox textTM124
End Sub
Private Sub textTM125_GotFocus()
   InverseTextBox textTM125
End Sub

Private Sub textTM124_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub
Private Sub textTM125_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM46_GotFocus()
   InverseTextBox textTM46
End Sub

Private Sub textTM46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' D/N是否列印申請人
Private Sub textTM46_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM46) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM46_GotFocus
   End If
End Sub

Private Sub textTM58_GotFocus()
   OpenIme
   textTM58.SetFocus
End Sub

' 案件備註
Private Sub textTM58_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If CheckLengthIsOK(textTM58, textTM58.MaxLength) = False Then
      Cancel = True
      textTM58_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM69_GotFocus()
   TextInverse Me.textTM69
End Sub

Private Sub textTM69_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM69_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM69) = False Then
      strTemp = textTM69
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM69.Text, 1) = "X" Then
         lblTM69 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(m_TM01, strTemp, strTempName) Then
            lblTM69 = strTempName
         Else
            lblTM69 = ""
         End If
      End If
      If IsEmptyText(lblTM69) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "D/N固定列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM69_GotFocus
      End If
   End If
End Sub

Private Sub textTM56_GotFocus()
   InverseTextBox textTM56
End Sub

Private Sub textTM56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 固定請款對象
Private Sub textTM56_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM56) = False Then
      strTemp = textTM56
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM56.Text, 1) = "X" Then
         lblTM56 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(m_TM01, strTemp, strTempName) Then
            lblTM56 = strTempName
         Else
            lblTM56 = ""
         End If
      End If
      
      If IsEmptyText(lblTM56) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "固定請款對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM56_GotFocus
      End If
   End If
End Sub

Private Sub textTM122_GotFocus()
   InverseTextBox textTM122
End Sub

Private Sub textTM122_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM122_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsYesOrSpace(textTM122) = False Or textTM122 = " " Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入Y,不可輸入空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM122_GotFocus
   End If
End Sub

Private Sub textTM68_GotFocus()
   InverseTextBox textTM68
End Sub

Private Sub textTM68_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 延展單筆不跑
Private Sub textTM68_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM68) = False Then
      If IsYesOrSpace(textTM68) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展單筆不跑請輸入Y或空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM68_GotFocus
      End If
   End If
End Sub

Private Sub textTM129_GotFocus()
   InverseTextBox textTM129
End Sub

Private Sub textTM129_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM129_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If textTM129 <> "Y" And textTM129 <> "" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "不催延展只可輸入Y或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM129_GotFocus
   End If
End Sub

Private Sub textTM33_GotFocus()
   InverseTextBox textTM33
End Sub

Private Sub textTM33_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 延展通知人
Private Sub textTM33_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM33) = False Then
      strTemp = textTM33
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM33.Text, 1) = "X" Then
         lblTM33 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(m_TM01, strTemp, strTempName) Then
            lblTM33 = strTempName
         Else
            lblTM33 = ""
         End If
      End If
      If IsEmptyText(lblTM33) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展代理人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM33_GotFocus
      End If
   End If
End Sub

Private Sub textTM65_GotFocus()
   InverseTextBox textTM65
End Sub

Private Sub textTM71_GotFocus()
    TextInverse Me.textTM71
End Sub

Private Sub textTM66_GotFocus()
   InverseTextBox textTM66
End Sub

Private Sub textTM66_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 延展請款對象
Private Sub textTM66_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM66) = False Then
      strTemp = textTM66
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM66.Text, 1) = "X" Then
         lblTM66 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(m_TM01, strTemp, strTempName) Then
            lblTM66 = strTempName
         Else
            lblTM66 = ""
         End If
      End If
      
      If IsEmptyText(lblTM66) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展請款對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM66_GotFocus
      End If
   End If
End Sub

Private Sub textTM121_GotFocus()
   CloseIme
   TextInverse textTM121
End Sub

Private Sub textTM121_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("D") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub textTM121_Validate(Cancel As Boolean)
   If (textTM121 = "" And textTM126 = "Y") Then
      MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
      Cancel = True
   End If
End Sub

Private Sub textTM70_GotFocus()
   TextInverse Me.textTM70
End Sub

Private Sub textTM70_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM70_Validate(Cancel As Boolean)
Dim strTemp As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim strTempName As String
   
   Cancel = False
   If IsEmptyText(textTM70) = False Then
      strTemp = textTM70
      ' 不滿八碼補0
      If Len(strTemp) < 8 Then: strTemp = strTemp & String(8 - Len(strTemp), "0")
      If Left(Me.textTM70.Text, 1) = "X" Then
         lblTM70 = GetAgentOrCustName(strTemp)
      Else
         If PUB_GetAgentName(m_TM01, strTemp, strTempName) Then
            lblTM70 = strTempName
         Else
            lblTM70 = ""
         End If
      End If
      If IsEmptyText(lblTM70) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "延展D/N列印對象代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM70_GotFocus
      End If
   End If
End Sub

Private Sub textTM126_GotFocus()
   InverseTextBox textTM126
End Sub

Private Sub textTM126_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM126_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM126) = False Then
      If IsYesOrSpace(textTM126) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "EMail同時寄紙本請輸入Y或空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM126_GotFocus
      End If
   End If
   If Cancel = False Then
      If (textTM121 = "" And textTM126 = "Y") Then
         MsgBox "【EMail 同時寄紙本】為 Y 時，【以EMail 通知】欄位也必須為 Y！"
         Cancel = True
      End If
   End If
End Sub

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
      'Modify By Sindy 2015/8/27 +m_TM01
      If GetAgentAndState(strNo, strTemp, , , True, m_TM01) Then
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

Private Sub textTM38_GotFocus()
   OpenIme
   InverseTextBox textTM38
End Sub

Private Sub textTM39_GotFocus()
   InverseTextBox textTM39
End Sub

Private Sub textTM40_GotFocus()
   OpenIme
   InverseTextBox textTM40
End Sub

Private Sub textTM41_GotFocus()
   OpenIme
   InverseTextBox textTM41
End Sub

Private Sub textTM42_GotFocus()
   InverseTextBox textTM42
End Sub

Private Sub textTM43_GotFocus()
   OpenIme
   InverseTextBox textTM43
End Sub

Private Sub textTM76_GotFocus()
   OpenIme
   InverseTextBox textTM76
End Sub

' 聯絡人1(中)
Private Sub textTM38_Validate(Cancel As Boolean)
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
    'If CheckLengthIsOK(textTM38, textTM38.MaxLength) = False Then
    If CheckLengthIsOK(textTM38, IIf((m_TM01 = "T" Or m_TM01 = "TF" Or m_TM01 = "FCT"), 30, 60)) = False Then
      Cancel = True
      textTM38_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人1(日)
Private Sub textTM40_Validate(Cancel As Boolean)
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM40, textTM40.MaxLength) = False Then
   If CheckLengthIsOK(textTM40, 60) = False Then
      Cancel = True
      textTM40_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(中)
Private Sub textTM41_Validate(Cancel As Boolean)
   Cancel = False
   'Modified by Lydia 2017/06/14 聯絡人(中)改為30字
   'If CheckLengthIsOK(textTM41, textTM41.MaxLength) = False Then
   If CheckLengthIsOK(textTM41, 30) = False Then
      Cancel = True
      textTM41_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

' 聯絡人2(日)
Private Sub textTM43_Validate(Cancel As Boolean)
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM43, textTM43.MaxLength) = False Then
   If CheckLengthIsOK(textTM43, 60) = False Then
      Cancel = True
      textTM43_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

Private Sub textTM76_Validate(Cancel As Boolean)
   Cancel = False
   'Modified by Lydia 2017/06/14
   'If CheckLengthIsOK(textTM76, textTM76.MaxLength) = False Then
   If CheckLengthIsOK(textTM76, 60) = False Then
      Cancel = True
      textTM76_GotFocus
   End If
   If Cancel = False Then CloseIme
End Sub

'Modify By Sindy 2014/12/2
' 取得客戶或是代理人名稱
Private Function GetAgentOrCustName(ByVal strData As String) As String
Dim strTemp As String
   
   GetAgentOrCustName = Empty
   If IsEmptyText(strData) = False Then
      Select Case UCase(Mid(strData, 1, 1))
         Case "X":
            'Modify By Sindy 2015/8/27 +m_TM01
            If GetCustomerAndState(strData, strTemp, , , True, m_TM01) Then
               GetAgentOrCustName = strTemp
            End If
         Case "Y":
            'Modify By Sindy 2015/8/27 +m_TM01
            If GetAgentAndState(strData, strTemp, , , True, m_TM01) Then
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

' 檢查是否為Y或空白
Private Function IsYesOrSpace(ByVal strData As String) As Boolean
   IsYesOrSpace = False
   Select Case strData
      Case "", "Y", " ":
         IsYesOrSpace = True
      Case Else:
         IsYesOrSpace = False
   End Select
End Function

'Added by Lydia 2016/11/23 各項指示
Private Sub cmdIns_Click()
   If textTMKey = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(Replace(textTMKey, "-", "")), Me
   frm12040159.Show
End Sub

'Added by Lydia 2017/06/14
Private Sub textTM39_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM39, 35) = False Then
      Cancel = True
      textTM39_GotFocus
   End If
End Sub
Private Sub textTM42_Validate(Cancel As Boolean)
   Cancel = False
    If CheckLengthIsOK(textTM42, 35) = False Then
      Cancel = True
      textTM42_GotFocus
   End If
End Sub

