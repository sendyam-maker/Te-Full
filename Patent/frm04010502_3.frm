VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010502_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准函輸入"
   ClientHeight    =   5412
   ClientLeft      =   1176
   ClientTop       =   1848
   ClientWidth     =   8772
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5412
   ScaleWidth      =   8772
   Begin VB.CommandButton Command2 
      Caption         =   "優先權資料"
      Height          =   375
      Left            =   2448
      TabIndex        =   79
      Top             =   70
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3630
      Left            =   0
      TabIndex        =   29
      Top             =   1800
      Width           =   8745
      _ExtentX        =   15431
      _ExtentY        =   6414
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   " 核准資料"
      TabPicture(0)   =   "frm04010502_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label11"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label12"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label14"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label15(0)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label16"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label18"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label27(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label28(0)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label29"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label3"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label15(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label28(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lbl606Fee"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbl606Year"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl412"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblDispDate"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label25(5)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "LabPriNo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label48"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label8"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label5"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text5(0)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text5(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text5(2)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text5(3)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text5(4)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text5(5)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text5(8)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text5(10)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text5(11)"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text5(12)"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text5(14)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text5(15)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text5(16)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "lbl415Date"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Combo2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text9"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Text10"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt412"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtDispDate"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtPriNo"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "Text6"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtIDSPt(2)"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtIDSFee(2)"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtIDSPt(1)"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtIDSFee(1)"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt415Date"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "本所分析"
      TabPicture(1)   =   "frm04010502_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text19"
      Tab(1).Control(1)=   "Label31"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "帳單"
      TabPicture(2)   =   "frm04010502_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text14"
      Tab(2).Control(1)=   "Text13"
      Tab(2).Control(2)=   "Text12"
      Tab(2).Control(3)=   "Combo3"
      Tab(2).Control(4)=   "Label25(3)"
      Tab(2).Control(5)=   "Label25(2)"
      Tab(2).Control(6)=   "Label25(1)"
      Tab(2).Control(7)=   "Label25(4)"
      Tab(2).ControlCount=   8
      Begin VB.TextBox txt415Date 
         Height          =   300
         Left            =   1680
         MaxLength       =   8
         TabIndex        =   10
         Top             =   1830
         Width           =   1095
      End
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   21
         Top             =   2970
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   1
         Left            =   3180
         MaxLength       =   3
         TabIndex        =   22
         Top             =   2970
         Width           =   375
      End
      Begin VB.TextBox txtIDSFee 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   2250
         MaxLength       =   6
         TabIndex        =   23
         Top             =   3270
         Width           =   765
      End
      Begin VB.TextBox txtIDSPt 
         Alignment       =   2  '置中對齊
         Height          =   270
         Index           =   2
         Left            =   3180
         MaxLength       =   3
         TabIndex        =   24
         Top             =   3270
         Width           =   375
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   7470
         MaxLength       =   2
         TabIndex        =   20
         Top             =   2670
         Width           =   375
      End
      Begin VB.TextBox txtPriNo 
         Height          =   270
         Left            =   4344
         MaxLength       =   4
         TabIndex        =   19
         Top             =   2670
         Width           =   1095
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Left            =   -73425
         TabIndex        =   69
         Top             =   780
         Width           =   1095
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   -69315
         MaxLength       =   8
         TabIndex        =   68
         Top             =   450
         Width           =   1095
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   -73425
         MaxLength       =   15
         TabIndex        =   67
         Top             =   450
         Width           =   2505
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   -71445
         TabIndex        =   66
         Top             =   780
         Width           =   945
      End
      Begin VB.TextBox txtDispDate 
         Height          =   270
         Left            =   7515
         MaxLength       =   8
         TabIndex        =   2
         Top             =   390
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txt412 
         Height          =   270
         Left            =   7605
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1512
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   7320
         MaxLength       =   2
         TabIndex        =   12
         Top             =   1830
         Width           =   735
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   4365
         MaxLength       =   2
         TabIndex        =   11
         Top             =   1830
         Width           =   735
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   4350
         TabIndex        =   7
         Top             =   1215
         Width           =   2430
      End
      Begin VB.Label lbl415Date 
         AutoSize        =   -1  'True
         Caption         =   "專利權期間延長至                           止"
         Height          =   180
         Left            =   180
         TabIndex        =   78
         Top             =   1890
         Width           =   2835
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   16
         Left            =   6870
         TabIndex        =   5
         Top             =   945
         Width           =   1725
         VariousPropertyBits=   671107099
         MaxLength       =   15
         Size            =   "3043;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   15
         Left            =   7320
         TabIndex        =   14
         Top             =   2100
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   14
         Left            =   1680
         TabIndex        =   8
         Top             =   1512
         Width           =   4596
         VariousPropertyBits=   671107097
         MaxLength       =   20
         Size            =   "8107;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   12
         Left            =   1680
         TabIndex        =   18
         Top             =   2670
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   11
         Left            =   7320
         TabIndex        =   17
         Top             =   2385
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   10
         Left            =   4350
         TabIndex        =   16
         Top             =   2385
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   8
         Left            =   1680
         TabIndex        =   15
         Top             =   2385
         Width           =   1095
         VariousPropertyBits=   671107097
         MaxLength       =   7
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   5
         Left            =   1680
         TabIndex        =   13
         Top             =   2115
         Width           =   255
         VariousPropertyBits=   671107103
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   4
         Left            =   1680
         TabIndex        =   6
         Top             =   1230
         Width           =   255
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   3
         Left            =   1680
         TabIndex        =   4
         Top             =   945
         Width           =   3870
         VariousPropertyBits=   671107099
         MaxLength       =   32
         Size            =   "6826;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   2
         Left            =   1680
         TabIndex        =   3
         Top             =   675
         Width           =   6255
         VariousPropertyBits=   671107099
         MaxLength       =   50
         Size            =   "11033;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   1
         Left            =   4680
         TabIndex        =   1
         Top             =   390
         Width           =   255
         VariousPropertyBits=   671107099
         MaxLength       =   1
         Size            =   "450;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text5 
         Height          =   300
         Index           =   0
         Left            =   1680
         TabIndex        =   0
         Top             =   390
         Width           =   1095
         VariousPropertyBits=   671107099
         MaxLength       =   8
         Size            =   "1931;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text19 
         Height          =   2775
         Left            =   -74760
         TabIndex        =   25
         Top             =   720
         Width           =   8055
         VariousPropertyBits=   -1467987941
         Size            =   "14208;4895"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ＩＤＳ報價:  1. 第一階段                    (           P)"
         Height          =   180
         Left            =   240
         TabIndex        =   77
         Top             =   3015
         Width           =   3540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "2. 第二階段                    (           P)"
         Height          =   180
         Left            =   1275
         TabIndex        =   76
         Top             =   3315
         Width           =   2505
      End
      Begin VB.Label Label48 
         AutoSize        =   -1  'True
         Caption         =   "引證前案檔案數量:"
         Height          =   180
         Left            =   5940
         TabIndex        =   75
         Top             =   2715
         Width           =   1485
      End
      Begin VB.Label LabPriNo 
         AutoSize        =   -1  'True
         Caption         =   "優先權存取碼:"
         Height          =   180
         Left            =   3165
         TabIndex        =   74
         Top             =   2730
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "帳單金額:"
         Height          =   180
         Index           =   3
         Left            =   -74775
         TabIndex        =   73
         Top             =   780
         Width           =   765
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "帳單日期:"
         Height          =   180
         Index           =   2
         Left            =   -70185
         TabIndex        =   72
         Top             =   480
         Width           =   765
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "代理人D/N NO:"
         Height          =   180
         Index           =   1
         Left            =   -74775
         TabIndex        =   71
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label25 
         Caption         =   "幣別:"
         Height          =   180
         Index           =   4
         Left            =   -71985
         TabIndex        =   70
         Top             =   840
         Width           =   405
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "審查委員編號:"
         Height          =   180
         Index           =   5
         Left            =   5625
         TabIndex        =   65
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label lblDispDate 
         AutoSize        =   -1  'True
         Caption         =   "機關發文日:"
         Height          =   180
         Left            =   6525
         TabIndex        =   64
         Top             =   435
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lbl412 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "延緩公告日:"
         Height          =   180
         Left            =   6570
         TabIndex        =   63
         Top             =   1560
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label lbl606Year 
         AutoSize        =   -1  'True
         Caption         =   "維持費起始年度:"
         Height          =   180
         Left            =   5925
         TabIndex        =   62
         Top             =   1875
         Width           =   1305
      End
      Begin VB.Label lbl606Fee 
         AutoSize        =   -1  'True
         Caption         =   "大陸維持費:"
         Height          =   180
         Left            =   5925
         TabIndex        =   61
         Top             =   2145
         Width           =   945
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "大陸年度:"
         Height          =   180
         Index           =   2
         Left            =   3525
         TabIndex        =   60
         Top             =   1875
         Width           =   765
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Label15"
         Height          =   180
         Index           =   2
         Left            =   6840
         TabIndex        =   59
         Top             =   1275
         Width           =   555
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "主管機關:"
         Height          =   180
         Index           =   1
         Left            =   3480
         TabIndex        =   58
         Top             =   1275
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "主管機關文書:"
         Height          =   180
         Left            =   240
         TabIndex        =   57
         Top             =   1557
         Width           =   1128
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "下次繳費日:"
         Height          =   180
         Left            =   240
         TabIndex        =   40
         Top             =   2730
         Width           =   945
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "領證法定期限:"
         Height          =   180
         Index           =   0
         Left            =   5925
         TabIndex        =   39
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "領證本所期限:"
         Height          =   180
         Index           =   0
         Left            =   3165
         TabIndex        =   38
         Top             =   2430
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "大陸領證費:"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   37
         Top             =   2430
         Width           =   945
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "專利權是否存在:          (Y/N)"
         Height          =   180
         Left            =   240
         TabIndex        =   36
         Top             =   2160
         Width           =   2190
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "列印客戶通知函:          (N:不印)"
         Height          =   180
         Left            =   240
         TabIndex        =   35
         Top             =   1275
         Width           =   2430
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "審查委員名稱:"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   34
         Top             =   993
         Width           =   1125
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "機關文號:"
         Height          =   180
         Left            =   240
         TabIndex        =   33
         Top             =   717
         Width           =   768
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "案件目前准駁:         (1:准 , 2:駁)"
         Height          =   180
         Left            =   3480
         TabIndex        =   32
         Top             =   435
         Width           =   2415
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "申請案核准日:"
         Height          =   180
         Left            =   240
         TabIndex        =   31
         Top             =   435
         Width           =   1125
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "本所分析"
         Height          =   180
         Left            =   -74760
         TabIndex        =   30
         Top             =   480
         Width           =   720
      End
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   960
      MaxLength       =   3
      TabIndex        =   44
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   43
      Top             =   660
      Width           =   855
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   42
      Top             =   660
      Width           =   255
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   270
      Left            =   2520
      MaxLength       =   2
      TabIndex        =   41
      Top             =   660
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7788
      TabIndex        =   28
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5736
      TabIndex        =   26
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6564
      TabIndex        =   27
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   960
      TabIndex        =   45
      Top             =   960
      Width           =   7635
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13467;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Enabled         =   0   'False
      Height          =   180
      Index           =   4
      Left            =   8010
      TabIndex        =   56
      Top             =   660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   1200
      TabIndex        =   55
      Top             =   1560
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   5340
      TabIndex        =   54
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   1200
      TabIndex        =   53
      Top             =   1320
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   0
      Left            =   5340
      TabIndex        =   52
      Top             =   660
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   960
      Width           =   768
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "申請案號:"
      Height          =   180
      Left            =   4380
      TabIndex        =   50
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   49
      Top             =   660
      Width           =   768
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質:"
      Height          =   180
      Left            =   120
      TabIndex        =   48
      Top             =   1320
      Width           =   765
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "收文號:"
      Height          =   180
      Left            =   4380
      TabIndex        =   47
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   1560
      Width           =   945
   End
End
Attribute VB_Name = "frm04010502_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (Text5,Text19,Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'2006/2/7整理
Option Explicit

Dim intWhere As Integer
'edit by nickc 2007/02/02 改動態
'Dim cp(1 To T_CP) As String
'Dim pA(1 To T_PA) As String
Dim cp() As String
Dim pa() As String

Dim strReceiveNo As String
Dim strYear As String, strSales As String
'Add By Cheng 2002/06/21
Public m_strCP10 As String '記錄上個畫面所點選案件的案件性質代號
'Add By Cheng 2003/01/07
Dim m_CP08 As String '機關文號
'92.1.14 add by sonia
Dim m_Year As String '是否通知年費期限
'Add by Morgan 2004/6/16
'新申請案領證期限
Dim stNP07 As String, stNP08 As String, stNP09 As String

'add by nickc 2005/06/17 'Memo by Lydia 2021/11/10 大陸發明案的衍生香港案
Dim m_HaveHK As Boolean
Dim m_HK_CP01 As String
Dim m_HK_CP02 As String
Dim m_HK_CP03 As String
Dim m_HK_CP04 As String
'Added by Lydia 2021/11/10
Dim m_Have044 As Boolean, m_CPto044(1 To 4) As String '大陸發明案的衍生澳門案
Dim m_bolFMP2 As Boolean '是否為寰華
'end 2021/11/10

'Remove by Morgan 2009/11/27
'Dim m_HaveHKInCP As String
'Dim m_HaveHKInNP As String
'Dim m_HKMailID As String
'Dim m_SendHKMail As Boolean
'end 2009/11/27
Dim cm(7) As String
Dim i As Integer
Dim m_strRetSheetNP07 As String '回覆單案件性質

'2006/2/7 ADD BY SONIA 香港公布選擇定稿別 1 記錄請求公佈通知書 2 政府憲報
Dim m_LetterType As VbMsgBoxResult
'Add by Morgan 2006/10/12
Dim m_strPA14 As String '預定公告日
Dim bolCancelClose As Boolean 'Add by Morgan 2007/5/7 是否取消閉卷
'Add by Morgan 2007/9/17
Dim strDiscCase As String '年費是否可抵減
Dim bolHave421 As Boolean '是否已收文技術報告
Dim bolHave121 As Boolean '是否被主張國內優先權
'Add by Morgan 2007/10/25
Dim m_strDualAppNP22 As String '一案兩請新型NP22
Dim m_strDualAppNP07 As String '一案兩請新型NP07
Dim m_strDualAppNo As String '一案兩請新型本所號
Dim m_dbl601OfficialFee As Double '大陸領證規費
Dim m_bolSaveCheck As Boolean '是否為存檔前檢查
Dim m_strHK202CP09 As String '香港補文件收文號
Dim m_bolFMP As Boolean '是否FMP案
Dim NewReceiveNo As String     '2010/7/7 modify by sonia 由FormSave移過來(因要印FMP通函定稿)
Dim m_bolTw307Chk As Boolean '台灣發明初審核准定稿是否提醒分割案期限 Added by Morgan 2012/9/19
Dim m_bolTw601Chk As Boolean '台灣領證期限是否適用新法 Added by Morgan 2012/9/19
Public str941CP14 As String   '內部收文941收文號及承辦人
'Added by Morgan 2014/1/14
Public m_DocNo As String
Public m_AppNo As String
'end 2014/1/14
'Added by Morgan 2014/4/17
Public m_DocWord As String
'end 2014/4/17
Dim m_str1914CP09 As String 'Added by Morgan 2014/7/15 通知放棄專利權收文號
Dim stCP10 As String 'Add by Morgan 2007/10/25 來函案件性質 'Add by Lydia 2014/11/18 改全域變數

Dim m_CP44 As String 'CF代理人 Added by Morgan 2014/11/24
Dim stCP12 As String, stCP13 As String 'Added by Morgan 2015/11/5
Dim m_bolEngCase As Boolean 'Added by Morgan 2016/3/18 臺灣舉發、舉發答辯審定來函是否工程師承辦
Dim m_Subject As String 'Added by Morgan 2016/6/8
'Add By Sindy 2016/10/5
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/5 END
Dim m_USCaseNo As String 'Added by Morgan 2019/5/24 相關美國案本所案號(提IDS用)
Dim m_bolNoCP27 As Boolean '不上發文 Added by Morgan 2020/1/16
Dim m_bolNewMedInform As Boolean, m_PA176 As String 'Added by Morgan 2021/6/29 大陸新藥是否通知專利權期限延長
Dim m_Close413 As String 'Added by Morgan 2023/3/3 自請撤回413詢問是否閉卷
'Added by Lydia 2023/06/15
Dim bolChk414for106 As Boolean, strFirstPriDate As String '寰華案:是否為「414恢復權利-主張優先權106」、最早優先權日
Dim strPriority(1 To 5) As String '優先權資料
Dim m_bolRePriDate As Boolean '優先權資料需重新輸入
Dim bolChgRlt As Boolean 'Added by Morgan 2024/6/5 是否為申請案核准(基本檔上准)
Dim bolCN445 As Boolean, bol615NP As Boolean 'Added by Morgan 2024/9/25
Dim m_bolReKeyInOK As Boolean 'Added by Morgan 2024/9/27 是否與2次確認期限一致
Dim m_bolAddB908 As Boolean 'Added by Morgan 2025/3/7 是否內部收文代辦退費

'Add by Morgan 2007/10/25
Private Sub StartLetter1(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)

   Dim strTxt(1 To 20) As String, intStep As Integer


   EndLetter ET01, ET02, ET03, strUserNum
   intStep = 1
   
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','發明案申請號','" & pa(11) & "')"
      
   intStep = intStep + 1
   
   'Added by Morgan 2013/11/19
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','發明案本所案號','" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & "')"
      
   intStep = intStep + 1
   'end 2013/11/19
   
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','發明案核准發文日','" & strSrvDate(1) & "')"
      
   intStep = intStep + 1
   
      If m_strDualAppNP07 <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','下一程序','" & m_strDualAppNP07 & "')"
         intStep = intStep + 1
      End If
   
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
   
End Sub

'Added by Morgan 2022/10/18
Private Sub StartLetter3(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 20) As String, intStep As Integer
   
   EndLetter ET01, ET02, ET03, strUserNum
   
   intStep = 0
   
   strExc(0) = ChgEngDate(DBDATE(pa(10)))
   strExc(0) = Left(strExc(0), Len(strExc(0)) - 6)
   intStep = intStep + 1
   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
      "','申請月日','" & strExc(0) & "')"
   
   If pa(22) <> "" Then
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','有專用期才印','♀')"
   End If
   
   intI = 1
   strExc(0) = " select NP09 FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "'" & _
         " AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
         " AND NP06 IS NULL AND NP07='605' order by np09 desc"
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','年費法定期限','" & RsTemp(0) & "')"
   End If
   
   'Added by Morgan 2025/3/10
   If bolCN445 Then
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','補償天數','" & txt412.Text & "')"
      
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','原專用期限止日','" & DBDATE(pa(25)) & "')"
         
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','原專利權期滿終止日','" & CompDate(2, 1, pa(25)) & "')"
      
      intStep = intStep + 1
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
         "','新專用期限止日','" & CompDate(2, -1, txt415Date.Text) & "')"
         
      
      If bol615NP Then
         intStep = intStep + 1
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','有補償期年費才印','♀')"
         
         intStep = intStep + 1
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & _
            "','有補償期年費不印','♀')"
      End If
   End If
   'end 2025/3/10
                  
   If Not ClsLawExecSQL(intStep, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add by Morgan 2009/11/27
Private Sub StartLetter2(ByVal ET01 As String, ByVal ET03 As String)
   Dim strTxt() As String, i As Integer
   Dim bolDualCaseUtility As Boolean, strInventionCaseNo As String, strInventionPA11 As String, strInventionPA77 As String 'Added by Morgan 2017/9/20 是否一案兩請新型案,發明案本所案號,發明案申請號,發明案彼所案號
   Dim strMemo As String 'Added by Morgan 2022/10/18
   
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   i = 0
   If pa(46) = "Y" Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','PCT案','♀')"
   Else
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','非PCT案','♀')"
   End If
   
   If Val(Text9) > 0 Then
      i = i + 1
      ReDim Preserve strTxt(i)
   
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','大陸年費年度','" & TranslateKeyWord(incCNV_ENGLISH_FREQUENCY, Val(Text9), "") & "')"
   
'Remove by Morgan 2010/1/25
'      If Val(Text10) = Val(Text9) - 1 Then
'         i = i + 1
'         ReDim Preserve strTxt(i)
'
'         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'            "','第幾年至幾年','" & TranslateKeyWord(11, Val(Text10), "") & "')"
'
'      ElseIf Val(Text10) < Val(Text9) - 1 Then
'         i = i + 1
'         ReDim Preserve strTxt(i)
'
'         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'            "','第幾年至幾年','" & TranslateKeyWord(11, Val(Text10), "") & " to " & TranslateKeyWord(11, Val(Text9) - 1, "") & "')"
'      End If

   End If
   
   If Text5(10) <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
   
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','本所期限','" & DBDATE(Text5(10)) & "')"
   End If
   
   If Text5(11) <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
   
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','法定期限','" & DBDATE(Text5(11)) & "')"
   End If
'Modify by Morgan 2010/1/25
'   If Text5(8) <> "" Or Text5(15) <> "" Then
'      i = i + 1
'      ReDim Preserve strTxt(i)
'
'      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'         "','台幣報價','" & (Val(Text5(8)) + Val(Text5(15))) & "')"
'
'      i = i + 1
'      ReDim Preserve strTxt(i)
'      strExc(1) = PUB_GetUSXRate
'      If Val(strExc(1)) <> 0 Then
'         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'            "','美金報價','" & Fix((Val(Text5(8)) + Val(Text5(15))) / Val(strExc(1))) & "')"
'      End If
'   End If
   If Text5(8) <> "" Then
      i = i + 1
      ReDim Preserve strTxt(i)
   
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','台幣報價','" & Val(Text5(8)) & "')"
   
      i = i + 1
      ReDim Preserve strTxt(i)
      strExc(1) = PUB_GetUSXRate
      If Val(strExc(1)) <> 0 Then
         strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','美金報價','" & Fix((Val(Text5(8))) / Val(strExc(1))) & "')"
      End If
   End If
'end 2010/1/25
   
   'Added by Morgan 2017/1/4
   '領證自動代繳
   'Modified by Morgan 2023/5/2 bug修正 NVL(NVL(PA71,FA42),CU72)->NVL(NVL(PA71,FA42),CU75)
   strExc(0) = "Select PA71, FA42, CU75" & _
      " From PATENT, FAGENT, CUSTOMER" & _
      " WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
      " AND CU01(+)=SUBSTR(PA26,1,8) AND CU02(+)=SUBSTR(PA26,9,1) AND NVL(NVL(PA71,FA42),CU75)='Y'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','領證自動代繳','♀')"
      
      'Added by Morgan 2021/2/1
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','領證自動代繳不印','♀')"
      'end 2021/2/1
   End If
   'end 2017/1/4
   
   'Added by Morgan 2017/9/22 大陸新型案核准英加一案兩請的段落--David
   If pa(9) = "020" And pa(8) = "2" Then
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) C1,pa16,pa14,pa11,pa77" & _
         " from (select cm05,cm06,cm07,cm08 from casemap where cm10='3' and cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "'" & _
         " union select cm01,cm02,cm03,cm04 from casemap where cm10='3' and cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "') X" & _
         ",patent where pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08 AND pa57 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '發明案未核准公告
         If Not ("" & RsTemp("pa16") = "1" And Not IsNull(RsTemp("pa14"))) Then
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','一案兩請新型案要印','♀')"
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','發明案本所案號','" & RsTemp("C1") & "')"
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','發明案申請號','" & ChgSQL(RsTemp("pa11")) & "')"
            i = i + 1
            ReDim Preserve strTxt(i)
            strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','發明案彼所案號','" & ChgSQL(RsTemp("pa77")) & "')"
               
         End If
      End If
   End If
   'end 2017/9/22
   
   'Added by Morgan 2022/10/18
   If pa(8) <> "3" Then
      'Modified by Lydia 2023/03/22 +pKind=1
      If PUB_GetApprovalPS("1", pa(1) & pa(2) & pa(3) & pa(4), pa(75), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), strMemo) = True Then
         If strMemo <> "" Then
             i = i + 1
             ReDim Preserve strTxt(i)
             strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','額外段落','" & strMemo & "')"
         End If
      End If
   End If
   'end 2022/10/18
   
   'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管(英文定稿)
   If m_bolNewMedInform Then
      i = i + 1
      ReDim Preserve strTxt(i)
      strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','新藥通知專利權期限延長','♀')"
   End If
   'end 2023/03/10
      
   If Not ClsLawExecSQL(i, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET03 As String)
Dim strTxt(1 To 22) As String, intStep As Integer, lTmp As Long, lTmp1 As Long
Dim strTmp As String
Dim dblTotalFee As Double
'Add by Morgan 2007/10/30
Dim bolUsIdsNotice As Boolean '美國IDS提醒
Dim dblPoint As Double '點數

   EndLetter ET01, strReceiveNo, ET03, strUserNum
   intStep = 1
   
   'Add by Morgan 2007/9/11
   bolUsIdsNotice = False
   '台灣之發明及設計案件,若有同時辦美國案(主案,未閉卷,未核准),則於通知核准或核駁定稿中加入一段美國提IDS之提醒字眼
   'Modify by Morgan 2011/7/18 +控制美國需為發明案(設計不用)--郭
   'Modified by Morgan 2019/5/27 需輸入報價，改存檔前檢查
   'If (PA(9) = "000" And (PA(8) = "1" Or PA(8) = "3")) Or cp(10) = "421" Then
   '   strExc(1) = PUB_GetUSCaseNo(PA(1), PA(2), PA(3), PA(4))
   '   If strExc(1) <> "" Then
      If m_USCaseNo <> "" Then
         bolUsIdsNotice = True
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','美國案本所案號','" & m_USCaseNo & "')"
         intStep = intStep + 1
         
         'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫，定稿要控制不出該報價文字
         If Val(txtIDSFee(1)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','IDS報價1','" & txtIDSFee(1) & "')"
            intStep = intStep + 1
         End If
         
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','IDS報價2','" & txtIDSFee(2) & "')"
         intStep = intStep + 1
         
      End If
   'End If
   'end 2019/5/27
   
   'Add by Morgan 2007/9/17
   'add by Toni 2008/09/11 加判斷ET03="36"
   'Modified by Lydia 2025/11/03 +FMP案定稿63
   If ET03 = "35" Or ET03 = "36" Or ET03 = "63" Then
      'Added by Lydia 2023/10/02
      Dim Str601SFee As String '領證服務費(取代strExc(1))
      Dim Str601Ann As String '領證費(取代strExc(2))
      Dim str605ANN As String '1~3年年費規費(取代strExc(3))
      'end 2023/10/02
      Str601SFee = "": Str601Ann = "": str605ANN = ""
      'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
      'Modified by Lydia 2015/04/13 call共用模組
'      strExc(0) = "Select YF06,YF07 From PatentYearFee Where YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='601' AND YF05=1"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Str601SFee = "" & RsTemp("YF06")
'         Str601Ann = "" & RsTemp("YF07")
'      End If
       strExc(3) = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "1", "1", Str601SFee, Str601Ann)
       'end 2015/04/13
      If Str601SFee <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','領證服務費','" & Str601SFee & "')"
         intStep = intStep + 1
      End If
      
      'Added by Lydia 2023/10/02 因為定稿分別顯示規費，所以抓標準規費;ex.亞浩電子X41570080
      Str601Ann = PUB_GetYF07(pa(9), pa(8), "Y00000001", "601", "1", "1", "1")
      'end 2023/10/02
      If Str601Ann <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','領證規費','" & Str601Ann & "')"
         intStep = intStep + 1
      End If
      
       'Modified by Lydia 2015/04/13 call共用模組
'      strExc(0) = "Select YF07 From PatentYearFee Where YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='605' AND YF05=1"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         str605ANN = "" & RsTemp("YF07")
'      End If
      str605ANN = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(pa(26)), "605", "1", "1", "1") 'Memo by Lydia 2023/10/02 1~3年年費原本就抓標準規費=>PAgent=1
      'end 2015/04/13
      If str605ANN <> "" Then
         If strDiscCase = "Y" Then
            str605ANN = Val(str605ANN) - 800
         End If
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','年費規費','" & str605ANN & "')"
         intStep = intStep + 1
      End If

      If ET03 = "36" Or Val(str605ANN) > 0 Then 'Add by Morgan 2011/7/6 若年費為 0 時前兩年的報價不印(台灣設計可減免1-3年免年費)
         '規費1
         strExc(4) = Val(Str601Ann) + Val(str605ANN)
         If Val(strExc(4)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','規費1','" & strExc(4) & "')"
            intStep = intStep + 1
         End If
         
         '費用1
         strExc(5) = Val(Str601SFee) + Val(strExc(4))
         If Val(strExc(5)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','費用1','" & strExc(5) & "')"
            intStep = intStep + 1
         End If
         
         '規費2
         strExc(4) = Val(Str601Ann) + Val(str605ANN) * 2
         If Val(strExc(4)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','規費2','" & strExc(4) & "')"
            intStep = intStep + 1
         End If
         
         '費用2
         strExc(3) = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "2", "1", Str601SFee) 'Added by Lydia 2024/04/30 重抓服務費;ex.亞浩電子X41570080
         strExc(5) = Val(Str601SFee) + Val(strExc(4))
         If Val(strExc(5)) > 0 Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','費用2','" & strExc(5) & "')"
            intStep = intStep + 1
         End If
      End If 'Add by Morgan 2011/7/6 若年費為 0 時前兩年的報價不印
      
      '規費3
      strExc(4) = Val(Str601Ann) + Val(str605ANN) * 3
      If Val(strExc(4)) > 0 Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','規費3','" & strExc(4) & "')"
         intStep = intStep + 1
      End If
      
      '費用3
      strExc(3) = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "3", "1", Str601SFee) 'Added by Lydia 2024/04/30 重抓服務費;ex.亞浩電子X41570080
      strExc(5) = Val(Str601SFee) + Val(strExc(4))
      If Val(strExc(5)) > 0 Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','費用3','" & strExc(5) & "')"
         intStep = intStep + 1
      End If
      'Modified by Morgan 2022/12/7 定稿插入項次2
      'strExc(6) = "3"
      strExc(6) = "4"
      'end 2022/12/7
      'Added by Morgan 2012/9/19
      '台灣領證期限跨102/1/1時適用102新法可補繳
      If m_bolTw601Chk = True Then
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','102新法不印','♀')"
            intStep = intStep + 1
      End If
      
      '台灣發明初審核准分割期限提醒
      If m_bolTw307Chk = True Then
         'Modified by Morgan 2019/7/26 108/8/1
         '配合108/11/1新法施行 108/8/1(含)分割敘述加"現行",108/10/1(含)改用新法
         If DBDATE(Label2(3)) >= 20191001 Then
            '改用新法
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x.3','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
            
         ElseIf DBDATE(Label2(3)) >= 20190801 Then
            '加"現行"
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x.1','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
            
            '預告新法
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x.2','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
            
         Else
            '舊法
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
         End If
         
      'Added by Morgan 2019/7/26
      '配合108/11/1新法施行，10/1(含)以後收文核准函發明(不論初審或再審)及新型都有3個月的分割期限
      'Modified by Lydia 2025/11/03 +FMP案定稿63
      'ElseIf ET03 = "35" And (pa(8) = "1" Or pa(8) = "2") Then
      ElseIf (ET03 = "35" Or ET03 = "63") And (pa(8) = "1" Or pa(8) = "2") Then
         If DBDATE(Label2(3)) >= 20191101 Then
            '改用新法
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x.3','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
            
         ElseIf DBDATE(Label2(3)) >= 20190801 Then
            '預告新法
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','項次3x.2','" & strExc(6) & "')"
            intStep = intStep + 1
            strExc(6) = Val(strExc(6)) + 1
         End If
      'end 2019/7/26
      End If
      'end 2012/9/19
      '是否有被主張國內優先權
      If bolHave121 = True Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','項次3','" & strExc(6) & "')"
         intStep = intStep + 1
         strExc(6) = Val(strExc(6)) + 1
      End If
      
      '是否有收文技術報告
      If pa(8) = "2" And bolHave421 = False Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','項次4','" & strExc(6) & "')"
         intStep = intStep + 1
         strExc(6) = Val(strExc(6)) + 1
      End If
      
      '美國IDS提醒
      If bolUsIdsNotice = True Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','項次5','" & strExc(6) & "')"
         intStep = intStep + 1
         strExc(6) = Val(strExc(6)) + 1
      End If
      
      'Added by Morgan 2020/10/19
      '設計案核准增加衍生設計提醒
      If pa(8) = "3" And pa(3) = "0" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','項次6','" & strExc(6) & "')"
         intStep = intStep + 1
         strExc(6) = Val(strExc(6)) + 1
      End If
      
      'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
      'strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
      '   " SELECT '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "'" & _
      '   ",'年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
      'intStep = intStep + 1
      'end 2011/7/6
      
      If stNP07 <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','本所期限','" & stNP08 & "')"
         intStep = intStep + 1
            
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','法定期限','" & stNP09 & "')"
         intStep = intStep + 1
         
      End If
      '2012/7/18 add by sonia
      If Text5(14) <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','來函主管機關文書','" & Text5(14) & "')"
         intStep = intStep + 1
      End If
      '2012/7/18 end
   
   Else
   
      'Modify by Morgan 2005/2/23 區分申請國家的例外欄位
      Select Case pa(9)
         Case "000"
         
            'Add by Morgan 2004/6/15
            '收文日＞＝93/7/1核准案件掛三個月的領證期限
            If stNP07 <> "" Then
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','本所期限','" & stNP08 & "')"
               intStep = intStep + 1
                  
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','法定期限','" & stNP09 & "')"
               intStep = intStep + 1
               
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "SELECT '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "'" & _
                  ",'下一程序名稱',CPM03 FROM CASEPROPERTYMAP WHERE CPM01='" & pa(1) & "' AND CPM02='" & stNP07 & "' AND ROWNUM<2"
               intStep = intStep + 1
               
               'Remove by Morgan 2011/7/6 定稿已改用共用文字欄位
               'If (ET03 <> "28" And ET03 <> "30") Then
               '   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06)" & _
               '      " SELECT '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "'" & _
               '      ",'年費收費標準',FTM05 FROM FINALTEXTMAP WHERE FTM01='P' AND FTM02='21' AND FTM03='000' AND FTM04='02'"
               '   intStep = intStep + 1
               'End If
               'end 2011/7/6
               
               'Add by Morgan 2004/7/5
               '若發明, 新型申請有被主張國內優先權則註明15個月將被撤回
               If cp(10) <> 設計申請 Then
                  If bolHave121 = True Then
                     strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                        "','撤回註明','本案曾被主張國內優先權，依專利法規定，本案將自申請日起滿十五個月後，視為撤回，謹提醒　" & IIf(IsCustomerIndividual(pa(26)), "台端", "貴公司") & "。" & IIf(ET03 = "25", Chr(13), "") & "')"
                     intStep = intStep + 1
                  End If
               End If
            End If
            
            'Add by Morgan 2004/9/9 台灣聯合追加案用
            '2009/8/27 MODIFY BY SONIA 加 "37"
            If ET03 = "32" Or ET03 = "37" Then
               'Modify by Morgan 2010/12/24 申請號改碼數
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','加註種類','" & IIf(Mid(pa(11), 10, 1) = "U", "聯合", "追加") & "')"
               intStep = intStep + 1
            End If
            
            'Add by Morgan 2004/9/10 台灣延緩公告用
            If txt412.Visible = True Then
               If m_strCP10 = "412" Then
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','延緩公告日','" & txt412.Text & "')"
                  intStep = intStep + 1
               
               'Added by Morgan 2024/6/20
               ElseIf m_strCP10 = "245" Then
                  'Modified by Lydia 2025/02/12 續行審查日>>延緩審查日
                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                     "','延緩審查日','" & txt412.Text & "')"
                  intStep = intStep + 1
               'end 2024/6/20
               End If
            End If
            
         Case "020"
         
            'Added by Morgan 2024/9/25
            If bolCN445 Then
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','補償天數','" & txt412.Text & "')"
               intStep = intStep + 1
               
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','原專利權期滿終止日','" & CompDate(2, 1, pa(25)) & "')"
               intStep = intStep + 1
               
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','現專利權期滿終止日','" & txt415Date.Text & "')"
               intStep = intStep + 1
            End If
            'end2024/9/25
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514'
'舊程式已刪除
'end 2009/11/27
         
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','本所期限','" & Text5(10) & "')"
            intStep = intStep + 1
               
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','法定期限','" & Text5(11) & "')"
            intStep = intStep + 1
               
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','大陸領證費','" & Text5(8) & "')"
            intStep = intStep + 1
            strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               "','年費本所期限','" & Text5(12) & "')"
            intStep = intStep + 1

'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'舊程式已刪除
'end 2009/11/27
         
            'Modify by Morgan 2004/4/15加 23 24 定稿
            'Modified by Lydia 2025/11/03 +FMP案定稿61,62
            If ET03 = "03" Or ET03 = "04" Or ET03 = "23" Or ET03 = "24" Or ET03 = "61" Or ET03 = "62" Then
               'Removed by Morgan 2021/1/25 改定稿內容，不分專利種類都附通知書--郭
               'If pa(8) = "1" Then
               '   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               '      "','列印備註','附送本案之授予專利權通知書乙份，敬請查收備存。')"
               'Else
               '   strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '      "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
               '      "','列印備註','函請查照。')"
               'End If
               'intStep = intStep + 1
               'end 2021/1/25

               '92.10.17 add by sonia
               
'Remove by Morgan 2010/1/25
'               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'                  "','大陸維持費','" & Text5(15) & "')"
'               intStep = intStep + 1
               
               strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
                  "','條碼年費年度','" & Text9 & "')"
               intStep = intStep + 1
               
'Remove by Morgan 2010/1/25
'               If Val(Text10) = Val(Text9) - 1 Then
'                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'                     "','第幾年至幾年費','第" & Val(Text10) & "年')"
'               'Modify by Morgan 2006/8/22 有維持費才帶--玲玲
'               'Else
'               ElseIf Val(Text9) > Val(Text10) Then
'                  strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
'                     "','第幾年至幾年費','第" & Val(Text10) & "年至第" & Val(Text9) - 1 & "年')"
'               End If
'               '92.10.17 end
'               intStep = intStep + 1

            End If
            
         '2006/2/7 ADD BY SONIA
         Case "013"
         '2006/2/7 END
      End Select
   
      'Add by Morgan 2003/12/22
      'Modify by Morgan 2010/1/25
      'dblTotalFee = Val(Text5(8)) + Val(Text5(15))
      dblTotalFee = Val(Text5(8))
      
      'Add end
      'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
      If m_Year = "Y" Then
         'Modified by Morgan 2014/12/9 年度不可固定抓2
         'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & ChangeCustomerL(cp(44)) & "' AND YF04='605' AND YF05=2"
         
          'Modified by Lydia 2015/04/13 call共用模組
'         strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & ChangeCustomerL(cp(44)) & "' AND YF04='605' AND YF05=" & Val(Text9) + 1
'         intI = 1
'         lTmp = 0
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            lTmp = Val(RsTemp.Fields(0))
'            dblPoint = dblPoint + Val(Format(Val(RsTemp.Fields(1)) / 1000, "0.0")) 'Add by Morgan 2008/6/16
         strExc(3) = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(cp(44)), "605", str(Val(Text9) + 1), str(Val(Text9) + 1), "1", strExc(1), strExc(2))
         If strExc(3) > 0 Then
            lTmp = Val(strExc(3))
            dblPoint = dblPoint + Val(Format(Val(strExc(1)) / 1000, "0.0")) 'YF06
'         Else
'            '內專抓代理人Y00000001
'            'Modified by Morgan 2014/12/9 年度不可固定抓2
'            'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='605' AND YF05=2"
'            strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0),NVL(YF06,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='605' AND YF05=" & Val(Text9) + 1
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               lTmp = Val(RsTemp.Fields(0))
'               dblPoint = dblPoint + Val(Format(Val(RsTemp.Fields(1)) / 1000, "0.0")) 'Add by Morgan 2008/6/16
'            End If
         End If
         'end 2015/04/13
         
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','年費金額','" & lTmp & "')"
         intStep = intStep + 1
         'Add by Morgan 2003/12/22
         dblTotalFee = dblTotalFee + lTmp
         'Add end
      End If
      
      'Add by Morgan 2003/12/22
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','總費用','" & Format(dblTotalFee) & "')"
      intStep = intStep + 1
      'Add end 2003/12/22
      
      'Add by Morgan 2008/6/16
      dblPoint = dblPoint + Val(Format((Val(Text5(8)) - m_dbl601OfficialFee) / 1000, "0.0"))
      If Text5(8) <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','點數','" & dblPoint & "')"
         intStep = intStep + 1
         
         'Added by Morgan 2016/6/13 非台灣信函進度要存報價
         PUB_UpdateLP2930 NewReceiveNo, Format(dblTotalFee), Format(dblPoint)
         'end 2016/6/13
         
      End If
      
      strTmp = Text19.Text
      strTmp = Replace(strTmp, Label15(2).Caption, Combo2.Text)
      Text19.Text = strTmp
      'Add By Cheng 2002/12/29
      '若收文日為111111
      'Modified by Lydia 2016/06/08
      'If cp(5) = 19221111 Then
      If Val(cp(5)) = 19221111 Or Val(cp(5)) = 111111 Then
           strTmp = Replace(strTmp, "本案獲勝係主管機關採信我方之理由。", "")
      End If
      
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','本所分析'," & CNULL(strTmp) & ")"
      intStep = intStep + 1
      
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','來函主管機關','" & Combo2.Text & "')"
      intStep = intStep + 1
      
      'Add By Cheng 2002/06/21
      If Text5(14) <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','來函主管機關文書','" & Text5(14) & "')"
         intStep = intStep + 1
      End If
      
      If m_strRetSheetNP07 <> "" Then
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','下一程序','" & m_strRetSheetNP07 & "')"
         intStep = intStep + 1
      End If
   End If
   'END 2007/9/17
   
   'Add by Morgan 2010//1/18
   If m_bolFMP Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " select '" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','約定期限',NP23 FROM NEXTPROGRESS WHERE NP02='" & pa(1) & "' AND NP03='" & pa(2) & "'" & _
         " AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "'" & _
         " AND NP06 IS NULL AND NP07='601' AND ROWNUM<2"
      intStep = intStep + 1
   'Move by Lydia 2025/11/03 內專P案增加約定期限；修改定稿P-05-000-03,P-05-000-04,P-05-000-35,P-05-307-04
        '原程式在End If下方，改移到這裡
   Else
      'Added by Lydia 2025/10/29
      'Modified by Lydia 2025/11/11 +有法限才要產生; Ex.P-136275核發-申請優先權證明書(11/10)
      'If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
      If m_bolFMP = False And ((pa(9) = "000" And Val(stNP09) > 0) Or Text5(11) <> "") Then
         If pa(9) = "000" And Val(stNP09) > 0 Then   'debug:2025/10/28 大陸案用輸入的領證期限
            strExc(1) = PUB_GetPOurDeadline(DBDATE(stNP09), pa(9), strSql, pa(1), m_strCP10)
         Else
            strExc(1) = PUB_GetPOurDeadline(DBDATE(Text5(11)), pa(9), strSql, pa(1), m_strCP10)
         End If
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','約定期限'," & CNULL(strSql, True) & ")"
         intStep = intStep + 1
      'Modified by Lydia 2025/11/11 +有法限才要產生
      'Else
      ElseIf (pa(9) = "000" And Val(stNP08) > 0) Or Text5(10) <> "" Then
         '使用協理提供的定稿設定<本所期限>-5日
         If pa(9) = "000" And Val(stNP08) > 0 Then   'debug:2025/10/28 大陸案用輸入的領證期限
             strExc(1) = CompWorkDay(5, CompDate(2, -1, DBDATE(stNP08)), 1)
         Else
             strExc(1) = CompWorkDay(5, CompDate(2, -1, DBDATE(Text5(10))), 1)
         End If
         strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
            "','約定期限','" & strExc(1) & "')"
         intStep = intStep + 1
      End If
      'end 2025/10/29
   'end 2025/11/03
   End If

            
   'Added by Morgan 2021/6/29 中文定稿
   If m_bolNewMedInform Then
      strTxt(intStep) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         "VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & _
         "','新藥通知專利權期限延長','♀')"
      intStep = intStep + 1
   End If
   'end 2021/6/29
   
   'edit by nickc 2007/02/05 不用 dll 了
   'If Not objLawDll.ExecSQL(intStep - 1, strTxt) Then
   If Not ClsLawExecSQL(intStep - 1, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If

End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim strTmp As String
   Dim strTmp1 As String
   Dim strTmp2 As String
   Dim tmpArr As Variant 'Added by Lydia 2023/06/15
   
   Select Case Index
      Case 0 '確定
         'Add by Morgan 2004/5/6將輸入的機關文號帶到本所分析
         Text5_LostFocus 2
         'Add end
         
         'Add by Morgan 2007/6/12
         If txtDispDate.Visible = True Then
            If txtDispDate = "" Then
               MsgBox "機關發文日不可空白！", vbCritical
            ElseIf ChkDate(txtDispDate) = False Then
               txtDispDate.SetFocus
               Exit Sub
            End If
         End If
         'end 2007/6/12
         
         'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
         'If IsEmptyText(Text5(6)) = True Then
            If (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "104" Or cp(10) = "105" Or cp(10) = "307" Or Left(cp(10), 1) = "3") And Text5(0).Text = "" Then
               MsgBox "申請案核准日不可空白 !", vbCritical
               Exit Sub
            End If
            
            If pa(9) <> "000" And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "104" Or cp(10) = "105" Or cp(10) = "307") Then
               If IsEmptyText(Text5(10)) = True Then
                  MsgBox "請輸入領證本所期限", vbCritical
                  Text5(10).SetFocus
                  Exit Sub
               End If
               If IsEmptyText(Text5(11)) = True Then
                  MsgBox "請輸入領證法定期限", vbCritical
                  Text5(11).SetFocus
                  Exit Sub
               End If
            End If
         'End If
         
         'If Text5(1).Text = "" Then MsgBox "是否更新目前准/駁不可空白", vbInformation
         
         ' 90.07.05 modify by louis (專利權是否存在不輸入)
         'If Text5(5).Text = "" Then MsgBox "專利權是否存在不可空白", vbInformation: Text5(5).SetFocus: Exit Sub
         
         ' 90.07.05 modify by louis
         ' 案件性質為爭議案(8字頭),
         If Mid(cp(10), 1, 1) = "8" Then
            If pa(9) < "010" Then
               If IsEmptyText(Text19) = True Then
                  MsgBox "案件性質為爭議案且申請國家為台灣, 本所分析不可為空白 !", vbCritical
                  Exit Sub
               End If
            End If
         End If
         
         '檢查下次繳費日
         With Me.Text5(12)
            If .Text <> "" Then
               If Len(.Text) = 8 Then
                  If Val(.Text) < strSrvDate(1) Then
                     MsgBox "下次繳費日不可小於系統日!!!", vbExclamation
                     .SetFocus
                     Exit Sub
                  End If
               ElseIf Len(.Text) = 7 Or Len(.Text) = 6 Then
                  If Val(.Text) + 19110000 < strSrvDate(1) Then
                     MsgBox "下次繳費日不可小於系統日!!!", vbExclamation
                     .SetFocus
                     Exit Sub
                  End If
               End If
            End If
         End With
         
         m_Year = ""
         '92.1.14 add by sonia檢查是否通知下一年年費期限
         'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
         'If IsEmptyText(Text5(6)) = True And pa(9) = "020" Then
         If pa(9) = "020" Then
            CheckYear
         End If
         '92.1.14 end
        
         'Add By Cheng 2003/03/26檢查機關文號
         If pa(9) = 台灣國家代號 Then
            If cp(10) <> 申請優先權證明 Then 'Add by Morgan 2011/6/24
               If Me.Text5(2).Tag = Me.Text5(2).Text Then
                  MsgBox "請輸入機關文號!!!", vbExclamation + vbOKOnly
                  Me.Text5(2).SetFocus
                  Text5_GotFocus 2
                  Exit Sub
               End If
            End If 'Add by Morgan 2011/6/24
         End If
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'         'Add By Cheng 2003/07/18大陸案
'         If pa(9) <> 台灣國家代號 Then
'            If Me.Text5(6).Text = "2" Or Me.Text5(6).Text = "4" Then
'               If Me.Text7.Text = "" Then
'                  MsgBox "請輸入公開日!!!", vbExclamation + vbOKOnly
'                  Me.Text7.SetFocus
'                  Text7_GotFocus
'                  Exit Sub
'               End If
'               If Me.Text8.Text = "" Then
'                  MsgBox "請輸入公開號!!!", vbExclamation + vbOKOnly
'                  Me.Text8.SetFocus
'                  Text8_GotFocus
'                  Exit Sub
'               End If
'            End If
'         End If
'
'         '2006/2/7 ADD BY SONIA香港案公告
'         If pa(9) = "013" And Me.Text5(6).Text = "6" Then
'            If Me.Text5(13).Text = "" Then
'               MsgBox "請輸入香港公告日!!!", vbExclamation + vbOKOnly
'               Me.Text5(13).SetFocus
'               Text7_GotFocus
'               Exit Sub
'            End If
'            If pa(8) = "2" And Me.Text11.Text = "" Then   '短期再檢查公告號
'               MsgBox "請輸入香港公告號!!!", vbExclamation + vbOKOnly
'               Me.Text11.SetFocus
'               Text8_GotFocus
'               Exit Sub
'            End If
'         End If
'
'         'Add By Morgan 2006/10/14  澳門公告
'         If pa(9) = "044" And Me.Text5(6).Text = "6" Then
'            If Me.Text5(13).Text = "" Then
'               MsgBox "請輸入澳門公告日!!!", vbExclamation + vbOKOnly
'               Me.Text5(13).SetFocus
'               Text7_GotFocus
'               Exit Sub
'            End If
'         End If
'end 2009/11/27
         
         'Add by Morgan 2008/5/20
         If Text5(3) = "" Then
            'Modified by Morgan 2023/7/25 排除新型的分割
            'If pa(9) = 台灣國家代號 And InStr("101,103,104,107,301,303,304,306,307,803,804", cp(10)) > 0 Then
            If pa(9) = 台灣國家代號 And InStr("101,103,104,107,301,303,304,306,307,803,804", cp(10)) > 0 And Not (pa(8) = "2" And cp(10) = "307") Then
            'end 2023/7/25
               MsgBox "案件性質為【" & Label2(1) & "】時，審查委員不可空白！"
               Text5(3).SetFocus
               Exit Sub
            End If
         End If
         'end 2008/5/20
         
         'Add by Amy 2014/03/24
         'Modify by Amy 2014/07/09 玲玲只需控制申請國為台灣才需輸優先權存取碼P-104461
         'If Text1 = "P" And (pa(9) = 台灣國家代號 Or pa(9) = 大陸國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
         'Modified by Morgan 2017/3/8 申請優先權存取碼 436 都要輸存取碼
         'If Text1 = "P" And (pa(9) = 台灣國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
         If Text1 = "P" And (cp(10) = "436" Or (pa(9) = 台灣國家代號 And (pa(8) = "1" Or pa(8) = "2") And cp(10) = "405")) Then
            If Trim(txtPriNo) = "" Then
                MsgBox "案件性質為【" & Label2(1) & "】時，優先權存取碼不可空白！"
                txtPriNo.SetFocus
                Exit Sub
            End If
            If Len(Trim(txtPriNo)) <> 4 Then
                MsgBox "優先權存取碼輸入錯誤，需輸入4碼！"
                txtPriNo.SetFocus
                Exit Sub
            End If
         End If
         'end 2014/03/24
         
         'Add By Cheng 2002/05/22
         '重新檢查欄位有效性
         m_bolSaveCheck = True
         If TxtValidate = False Then
            m_bolSaveCheck = False
            Exit Sub
         End If
         
         'Added by Lydia 2023/06/15 寰華案:「414恢復權利-主張優先權106」操作人員檢查並且維護優先權資料，確定優先權作業後，系統檢查最早優先權日是否一致。
         If bolChk414for106 = True Then
            If m_bolRePriDate = False Then
               MsgBox "請檢查優先權資料！", vbCritical
               Exit Sub
            Else
               tmpArr = Split(strPriority(2), "，")
               strExc(5) = tmpArr(0)
               For intI = 0 To UBound(tmpArr)
                  If Trim("" & tmpArr(intI)) <> "" Then
                     If Val(strExc(5)) > Val(Trim("" & tmpArr(intI))) Then
                        strExc(5) = Trim("" & tmpArr(intI))
                     End If
                  End If
               Next intI
               If strExc(5) <> strFirstPriDate Then
                  If MsgBox("優先權資料與前次輸入不一致，請問是否存檔？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Sub
                  End If
               End If
            End If
         End If
         'end 2023/06/15
         
         'Add by Morgan 2005/5/20
         '非台灣 宣告無效答辯 詢問是否計算結餘
         If cp(10) = "804" Then
            'Modified by Lydia 2015/03/03 +pa01,pa02,pa03,pa04
            Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
         End If
         
         'add by sonia 2025/3/31 案件僅變更401、讓與701及708，於核准時詢問是否計算結餘
         If (cp(10) = "401" Or cp(10) = "701" Or cp(10) = "708") And pa(16) = "1" Then
            strExc(0) = "Select * From Caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' And Cp09<>'" & cp(9) & "' and cp09<'B' and cp60 is not null and Cp59 Is Null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               Pub_EndModCashMsg pa(9), pa(1), pa(2), pa(3), pa(4)
            End If
         End If
         'end 2025/3/31
         
         'add by nickc 2005/06/17 取有關大陸的香港關聯判斷
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'         m_HaveHKInCP = ""
'         m_HaveHKInNP = ""
'         m_SendHKMail = False
'         m_HKMailID = ""
'
'end 2009/11/27

         m_HaveHK = False
         m_Have044 = False 'Added by Lydia 2021/11/10
         
         'Modified by Morgan 2014/9/23 大陸發明才要--郭
         'If pa(9) = "020" Then
         If pa(8) = "1" And pa(9) = "020" Then
            'edit by nickc 2006/05/05
            'm_HaveHK = ChkCMIsExist013(pA(1), pA(2), pA(3), pA(4))
            'Modified by Morgan 2014/9/23 香港標準專利(發明)才要--郭
            m_HaveHK = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_HK_CP01, m_HK_CP02, m_HK_CP03, m_HK_CP04, , "1")
            'Added by Lydia 2021/11/10 澳門案
            m_Have044 = ChkCMIsExist013(pa(1), pa(2), pa(3), pa(4), m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4), , "1", "5")
            
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'            If m_HaveHK = True Then
''edit by nickc 2006/05/05
''               For i = 1 To 4
''                  cm(i - 1) = pA(i)
''               Next
''               If obj003.GetCaseMap(cm, 4) = True Then
''                  m_HK_CP01 = cm(4)
''                  m_HK_CP02 = cm(5)
''                  m_HK_CP03 = cm(6)
''                  m_HK_CP04 = cm(7)
''               End If
''               m_HaveHKInCP = Chk013Have111(pa(1), pa(2), pa(3), pa(4), m_HKMailID)
''               m_HaveHKInNP = Chk013Have111(pa(1), pa(2), pa(3), pa(4), m_HKMailID, "NP")
'               m_HaveHKInCP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID)
'               m_HaveHKInNP = Chk013Have110(pa(1), pa(2), pa(3), pa(4), m_HKMailID, "NP")
'            End If
'en 2009/11/27
         End If

         
         'Add by Morgan 2006/5/16 大陸新案沒通知書且要通知下一年年費期限的才要提醒
         Select Case cp(10)
            Case 發明申請, 新型申請, 設計申請, 改請發明, 改請新型, 改請設計, "107", "307", 追加申請, 聯合申請, "110", "111", "112", "109"
               If pa(9) = "020" Then
                  'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
                  'If Text5(6) = "" And m_year = "Y" Then
                  If m_Year = "Y" Then
                     '從Text5(12)的Validate移過來
                     If strYear <> "" And Text5(12) <> strYear Then
                        If MsgBox("是否確定修改 ?", vbYesNo, "下次繳費日") = vbNo Then
                           Text5(12) = strYear
                           Exit Sub
                        End If
                     End If
                  End If
               End If
         End Select
         'end 2006/5/16
         
         'Add by Morgan 2007/5/4 若來函有期限但已閉卷
         bolCancelClose = False
         'modify by sonia 2015/11/6 +803,804 (P-094520)
         If pa(57) = "Y" And (Text5(11) <> "" Or (pa(9) = 台灣國家代號 And Val(Label2(3)) >= 930701 And _
            InStr("101,102,103,107,301,302,303,306,307,803,804", m_strCP10) > 0)) Then
            If MsgBox("本案目前為閉卷狀態，為管制期限將於存檔時取消閉卷，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
               Exit Sub
            End If
            bolCancelClose = True
         End If
         'end 2007/5/4
         
         'Added by Morgan 2024/11/20 一案兩請新型案的代辦退費908核准不可取消閉卷--玲玲
         intI = 0
         If pa(9) = 台灣國家代號 And cp(10) = "908" Then
            If PUB_ChkDualCase(pa) Then
               intI = 1
            End If
         End If
         If intI = 0 Then
         'end 2024/11/20
         
            'Added by Lydia 2015/12/17 對於已經閉卷的案件,後續若有官方來函是無期限的,全部都詢問user是否要取消閉卷,由user來判斷
            If pa(57) = "Y" And Text5(11) = "" And bolCancelClose = False Then
               If MsgBox("本案目前為閉卷狀態，您輸入的是無期限的來函，是否要取消閉卷？", vbYesNo + vbDefaultButton1) = vbYes Then
                  bolCancelClose = True
               End If
            End If
            'end 2015/12/17
            
         End If 'Added by Morgan 2024/11/20
         
         '2008/3/5 ADD BY SONIA 放棄專利權429核准且未閉卷時提醒會閉卷 P-068288
         If cp(10) = "429" And pa(57) = "" Then
            If MsgBox("此為放棄專利權之核准, 存檔後此案自動閉卷, 請確認？", vbYesNo + vbDefaultButton1) = vbNo Then
               Exit Sub
            End If
         End If
         '2008/3/5 END
         
         'Added by Morgan 2023/3/3 自請撤回413詢問是否閉卷
         m_Close413 = "N"
         If cp(10) = "413" And pa(57) = "" Then
            '新申請案不必詢問直接閉卷
            strExc(0) = "select cp27 from caseprogress where cp09='" & cp(43) & "' and cp10 in (" & NewCasePtyList & ") "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_Close413 = "Y"
            Else
               If MsgBox("此為自請撤回之核准, 請問是否要閉卷？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                  m_Close413 = "Y"
               End If
            End If
         End If
         'end 2023/3/3
         
         'Added by Morgan 2021/6/29
         '通知新藥發明專利權期限補償(大陸2021.6.1新法)
         m_bolNewMedInform = False
         m_PA176 = pa(176)
         'Modified by Morgan 2021/7/13 +判斷發明申請
         'Modified by Lydia 2023/03/07 +107再審申請
         If pa(9) = "020" And pa(8) = "1" And pa(158) = "3" And DBDATE(Text5(0)) >= "20210601" And (cp(10) = "101" Or cp(10) = "307" Or cp(10) = "107") Then
            '若尚未設定是否新藥
            If m_PA176 = "" Then
               intI = MsgBox("本案是否為新藥專利？", vbYesNoCancel + vbDefaultButton3 + vbQuestion)
               If intI = vbYes Then
                  m_PA176 = "Y"
               ElseIf intI = vbNo Then
                  m_PA176 = "N"
               Else
                  Exit Sub
               End If
            End If
            
            If m_PA176 = "Y" Then
               'Modified by Morgan 2022/3/2 有設定過都要通知,不必再確認 -- 郭
               'intI = MsgBox("本案為新藥專利，請先交承辦工程師確認是否通知專利權期限延長，確認後請選擇是或否！", vbYesNoCancel + vbDefaultButton3 + vbQuestion, "是否通知專利權期限延長？")
               'If intI = vbYes Then
               '   m_bolNewMedInform = True
               'ElseIf intI = vbNo Then
               '   m_bolNewMedInform = False
               'Else
               '   Exit Sub
               'End If
               m_bolNewMedInform = True
               'end 2022/3/2
            End If
         End If
         'end 2021/6/29
         'Added by Lydia 2023/03/10 FMP大陸新藥發明專利權期限補償控管
         If m_bolFMP = True And pa(9) = "020" And pa(8) = "1" And pa(150) = "2" And (cp(10) = "101" Or cp(10) = "307" Or cp(10) = "107") Then
             If m_PA176 = "Y" Then  '外專命名記錄設定(frm090902_2)
                m_bolNewMedInform = True
             End If
         End If
         'end 2023/03/10

         
         'Add By Sindy 2022/7/1
         'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
         'If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
         '      Exit Sub
         '   End If
         'End If
         ''2022/7/1 END
         'end 2023/05/17
         
         'Added by Morgan 2024/9/30
'         If Pub_B911NotPay(pa(1), pa(2), pa(3), pa(4)) = True Then
'           MsgBox "此案有未收款！", vbExclamation
'         End If
         'end 2024/9/30
         
         'Added by Morgan 2024/10/4
         If pa(1) = "P" And pa(9) <> "000" And m_bolFMP = False Then
            If Pub_B911NotPay(pa(1), pa(2), pa(3), pa(4)) = True Then
                MsgBox "此案有未收款！", vbExclamation
            End If
         End If
         'end 2024/10/4
            
         If FormSave = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
         
         '請連同新型卷一併交工程師確認是否為一案兩請
         
         strTmp1 = ""
         If Text5(4) <> "N" Then '通知函
            Select Case cp(10)
               '2006/2/6 MODIFY BY SONIA 加香港之110,111,112及ＰＣＴ之109
               'Modify by Morgan 2007/9/21 加改請獨立
               '2009/8/29 MODIFY BY SONIA 加改請聯合
               'Modified by Morgan 2012/12/19 +衍生設計125,改請衍生設計308
               Case 發明申請, 新型申請, 設計申請, 改請發明, 改請新型, 改請設計, "107", "307", 追加申請, 聯合申請, "110", "111", "112", "109", 改請獨立, 改請聯合, 衍生設計, 改請衍生設計
                  Select Case pa(9)
                     Case 台灣國家代號
                        
                        'strTmp = "02" 'Mark by Lydid 2025/06/18 整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿
                       
                        'Add by Morgan 2004/6/15
                        '新法要掛3個月的領證期限
                        If Val(Text5(0)) >= 930701 Then
                           'Add by Morgan 2007/9/17
                           bolHave121 = PUB_ChkPriDate(pa(11), , False)
                           strDiscCase = PUB_GetCaseDiscStat(pa(1) & pa(2) & pa(3) & pa(4))
                           'end 2007/9/17
                           'Modify by Morgan 2004/7/13
                           '聯合申請 出不同定稿
                           'Modify by Morgan 2004/9/9 改定稿內容-追加聯合共用
                           'Modify by Morgan 2010/12/24 申請號改碼數
                           'If Mid(pa(11), 9, 1) <> "" Then
                           'Modified by Morgan 2012/12/19 +衍生設計也有單獨的證書
                           'If Mid(pa(11), 10, 1) <> "" Then
                           'Mark by Lydid 2025/06/18 整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿；已確定無追加聯合設計
                           'If Mid(pa(11), 10, 1) <> "" And Mid(pa(11), 10, 1) <> "D" Then
                           '   '2009/8/27 ADD BY SONIA
                           '   If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then     '大-->台 定稿
                           '      strTmp = "37"
                           '   Else
                           '   '2009/8/27 END
                           '      strTmp = "32"
                           '   End If
                           'Else
                           'end 2025/06/18
                              'Modify by Morgan 2004/7/6
                              '改判斷專利種類
                              'If cp(10) = 新型申請 Then
                              If pa(8) = "2" Then
                                 'Add by Morgan 2007/9/17
                                 bolHave421 = PUB_ChkCPExist(pa, "421")
                                 'end 2007/9/17
                                 
                              'Removed by Morgan 2008/1/7
                              '   '規費可抵減
                              '   If Val(pa(10)) < 930701 Then
                              '      '符合年費減免
                              '      If strDiscCase = "Y" Then
                              '         strTmp = "30"
                              '      '不符合年費減免
                              '      Else
                              '         strTmp = "28"
                              '      End If
                              '   '規費不可抵減
                              '   Else
                              '      strTmp = "26"
                              '   End If
                              '   '是否已收文技術報告 27,29,31
                              '   If bolHave421 Then
                              '      strTmp = Format(Val(strTmp) + 1)
                              '   End If
                              'Else
                              '   strTmp = "25"
                              'end 2008/1/7
                              
                              End If
                               '2008/09/10 add by Toni
                               
                              If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then     '大-->台 定稿
                                 strTmp = "36"
                              'Add by Morgan 2008/1/7 一率改用新定稿
                              Else
                                 'Added by Lydia 2025/11/03 內專P案增加約定期限修改定稿，而FMP案定稿不動；
                                 If m_bolFMP = True Then
                                     strTmp = "63"
                                 Else
                                 'end 2025/11/03
                                     strTmp = "35"
                                 End If 'Added by Lydia 2025/11/03
                                 'Added by Morgan 2021/10/15 寶齡富錦 Y55435 案件
                                 If ChangeCustomerS(pa(75)) = "Y55435" Then
                                    strTmp = "99"
                                 End If
                                 'end 2021/10/15
                              'end 2008/1/7
                              End If
                              
                           'End If 'Mark by Lydia 2025/06/18
                        End If
                        'end
                        
                     Case 大陸國家代號
                        'Added by Morgan 2012/2/16
                        If cp(10) = "107" Then
                           strTmp = "38"
                        Else
                        'end 2012/2/16
                        
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514

                              '將屆
                              If m_Year = "Y" Then
                                  

                                  'Modify by Morgan 2004/4/15
                                  '加 23 24 定稿
                                  'strTmp = "03"
                                  'Modify by Morgan 2006/5/1 改判斷發明且維持費起始年度>大陸年度
                                  'If Text9 = "3" And Text10 = "3" Then
                                  'Modify by Morgan 2010/1/25 發明案取消維持費定稿
                                  'If pa(8) = "1" And Val(Text10) > 0 And Val(Text10) > Val(Text9) Then
                                  'Modified by Moragn 2012/6/28 改用統一定稿 --郭
                                  'If pa(8) = "1" Then
                                  '  strTmp = "23"
                                  'Else
                                    'Added by Lydia 2025/11/03 內專P案增加約定期限修改定稿，而FMP案定稿不動；
                                    If m_bolFMP = True Then
                                        strTmp = "61"
                                    Else
                                    'end 2025/11/03
                                        strTmp = "03"
                                    End If 'Added by Lydia 2025/11/03
                                  'End If
                                  'Modify end
                              
                              '未屆
                              Else
                              
                                  'Modify by Morgan 2004/4/15
                                  '加 23 24 定稿
                                  'strTmp = "04"
                                  'Modify by Morgan 2006/5/1 改判斷發明且維持費起始年度>大陸年度
                                  'If Text9 = "3" And Text10 = "3" Then
                                  'Modify by Morgan 2010/1/25 發明案取消維持費定稿
                                  'If pa(8) = "1" And Val(Text10) > 0 And Val(Text10) > Val(Text9) Then
                                  '  strTmp = "24"
                                  'Else
                                    'Added by Lydia 2025/11/03 內專P案增加約定期限修改定稿，而FMP案定稿不動；
                                    If m_bolFMP = True Then
                                        strTmp = "62"
                                    Else
                                    'end 2025/11/03
                                        strTmp = "04"
                                    End If 'Added by Lydia 2025/11/03
                                  'End If
                                  'Modify end
                                  
                              End If
                              
                              'Add by Morgan 2009/11/27
                              If m_bolFMP Then
                                 'If cp(10) = 發明申請 Then 'Remove by Morgan 2010/9/8 不必限制
                                    '付款後辦案
                                    'Modified by Morgan 2022/7/15 定稿合併
                                    'If CU72FA39("", pa(75)) Then
                                    '   strTmp2 = "55"
                                    'Else
                                    '   strTmp2 = "54"
                                    'End If
                                    strTmp2 = "54"
                                    'end 2022/7/15
                                 'End If
                              End If
                              
                           End If

                              
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514


                  End Select
               Case 異議_專, 舉發
                  If pa(9) = 台灣國家代號 Then '台灣 15
                     strTmp = "15"
                  Else                         '非台灣 16
                     strTmp = "16"
                  End If
               Case 異議答辯, 舉發答辯
                  If pa(9) = 台灣國家代號 Then '台灣 17
                     strTmp = "17"
                  Else                         '非台灣 18
                     strTmp = "18"
                  End If
               'Add By Cheng 2002/12/29
'2010/11/12 CANCEL BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
'               Case 訴願, 行政訴訟
'                  If pa(9) = 台灣國家代號 Then '台灣 08
'                     strTmp = "08"
'                  Else                         '大陸 01
'                     strTmp = "01"
'                  End If
               Case 申請英文證明
                  If pa(9) = 台灣國家代號 Then '台灣 09
                     strTmp = "09"
                  Else                         '大陸 01
                     strTmp = "01"
                  End If
               
               Case 申請優先權證明
                  strTmp = "19"
                  
               'Add by Morgan 2005/12/26 台灣技術報告改開窗
               'Modify by Morgan 2007/8/31 加807
               Case "421", "807"
                  If pa(9) = 台灣國家代號 Then
                     strTmp = "00"
                     '有客戶案件案號
                     If pa(48) <> "" Then
                        strTmp = "50"
                     End If
                  'Mark by Lydia 2025/06/24 整理所有國內對客戶的通知函定稿：協理確定要刪除的定稿P-05-421-01，大陸案沒有807，所以Mark程式碼
                  '相關備註有:Add by Morgan 2009/10/6：Form_Load: 檢索報告預設不出定稿、FormSave：大陸檢索報告421,新穎性調查426的核准預設原承辦不上發文日,後面加上文件齊備日。不上發文日表示由工程師出撰寫信函。
                  'Else
                  '   strTmp = "01"
                  End If
                  
               Case Else
                  '96/9/10P-078895授權案取消定稿內之影本乙份(改寄正本),因前讓與,申請權讓與已改為獨立定稿一併取消
                  If pa(9) = 台灣國家代號 Then '台灣 00
                     '2008/12/25 MODIFY BY SONIA
                     'strTmp = "00"
                     If PUB_CheckCuNation(pa(26), Text1, Text2, Text3, Text4) = "1" Then     '大-->台 定稿
                        strTmp = "12"
                     Else
                        strTmp = "00"
                     End If
                     '2008/12/25 END
                     'Added by Lydia 2016/06/08 非本所辦理的讓與701和專利權讓與708通知書
                     If (Val(cp(5)) = 19221111 Or Val(cp(5)) = 111111) And InStr("701,708", cp(10)) > 0 Then
                        strTmp = "57"
                     End If
                     'end 2016/06/08
                  Else                         '大陸 01
                     strTmp = "01"
                     'Add by Morgan 2006/11/28 變更,讓與,申請權讓與,授權
                     If InStr("401,701,708,704", cp(10)) > 0 Then
                        strTmp = "10"
                        'Added by Morgan 2021/3/29 寶齡富錦 Y55435 案件
                        If pa(75) = "Y55435" Then
                           strTmp = "08"
                        End If
                     End If
                     
                  End If
            End Select
            
            StartLetter "05", strTmp
            'Add by Morgan 2009/11/23
            If m_bolFMP Then
               'Modify by Morgan 2010/7/21 要報價改回原定稿 --敏惠(美珍要求)
               ''2010/7/7 MODIFY BY SONIA 改用通函
               ''NowPrint strReceiveNo, "05", strTmp, False, strUserNum, 0, , , , 1
               'NowPrint NewReceiveNo, "07", "99", False, strUserNum, 0, , , , 1
               ''2010/7/7 END
               'Modified by Morgan 2023/4/10 FMP不必再印紙本定稿
               NowPrint strReceiveNo, "05", strTmp, False, strUserNum, 0, , , , 1, , , , , , , , NewReceiveNo, , , , , True
               'end 2010/7/21
               If strTmp2 <> "" Then
                  strUserNum = strFMPNum
                  StartLetter2 "05", strTmp2
                  'Modified by Morgan 2016/5/30 不可傳LD18否則FCP承辦執行定維護時會開E化的畫面
                  'NowPrint strReceiveNo, "05", strTmp2, False, strUserNum, 0, , , , , , , , , , , , NewReceiveNo
                  NowPrint strReceiveNo, "05", strTmp2, False, strUserNum
                  strUserNum = strUser1Num
                  
               'Added by Morgan 2022/10/18 變更,讓與,合併,繼承,授權...核准函
               ElseIf stCP10 = "1001" Then
                  'Added by Morgan 2025/3/10
                  If cp(10) = "445" Then
                     strTmp2 = "58"
                  Else
                  'end 2025/3/10
                     strTmp2 = "51"
                  End If
                  strUserNum = strFMPNum
                  StartLetter3 "05", NewReceiveNo, strTmp2
                  NowPrint NewReceiveNo, "05", strTmp2, False, strUserNum
                  strUserNum = strUser1Num
               'end 2022/10/18
               End If
            Else
            'end 2009/11/23
               NowPrint strReceiveNo, "05", strTmp, False, strUserNum, 0, , , , , , , , , , , , NewReceiveNo
            End If
                     
            'Added by Lydia 2023/06/15 寰華案:「414恢復權利-主張優先權106」更新實審期限
            If bolChk414for106 = True Then
               '請彈跳提醒視窗：已更新實體審查期限為:
               strSql = "select '1' as ord1, cp06 from caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10='416' and cp158=0 " & _
                        "union select '2' as ord1, np08 from nextprogress WHERE np02='" & pa(1) & "' AND np03='" & pa(2) & "' AND np04='" & pa(3) & "' AND np05='" & pa(4) & "' and np07='416' and np06 is null " & _
                        "order by ord1"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  MsgBox IIf(Left(strFirstPriDate, 1) = "Y", "已更新", "") & "實體審查期限為: " & ChangeWStringToTDateString("" & RsTemp.Fields("cp06")) & "，請通知承辦報告客戶。 ", vbExclamation + vbOKOnly
               End If
            End If
            'end 2023/06/15
            
            'Add by Morgan 2007/10/24 +新型放棄專利權
            If m_strDualAppNo <> "" Then
               strTmp = ""
               If pa(9) = 大陸國家代號 Then
                  strTmp = "11"
               'Added by Morgan 2012/8/15 102新法
               ElseIf pa(9) = "000" Then
                  strTmp = "13"
               End If
               If strTmp <> "" Then
                  StartLetter1 "05", m_strDualAppNo & "&000", strTmp
                  NowPrint m_strDualAppNo & "&000", "05", strTmp, False, strUserNum, 0, , , , IIf(m_bolFMP, 1, 0), , , , , , , , m_str1914CP09
               End If
            End If
            
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'            'add by nickc 2005/06/17 發 mail
'            If m_SendHKMail = True And m_HKMailID <> "" And m_HaveHKInCP <> "" Then
'               Call PUB_SendMail(strUserNum, m_HKMailID, m_HaveHKInCP, "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "大陸案(" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & ")已公布，香港案(" & m_HK_CP01 & "-" & m_HK_CP02 & "-" & m_HK_CP03 & "-" & m_HK_CP04 & ")的[批准紀錄請求]可以處理！", "")
'            End If
'
'end 2009/11/27

         End If
                     
         'Add by Morgan 2009/10/2 香港技術報告
         If pa(9) = "013" And cp(10) = "421" Then
            If m_strHK202CP09 <> "" Then
               NowPrint cp(9), "13", "04", False, strUserNum, , , , , IIf(m_bolFMP, 1, 0), , , , , , , , m_strHK202CP09
               'Modified by Morgan 2016/6/8 指示信電子化
               'NowPrint m_strHK202CP09, "02", "40", False, strUserNum, , , , , , , , , , , , , NewReceiveNo
               'PUB_PrintLetter m_strHK202CP09
               If Left(Pub_StrUserSt03, 1) = "F" Then
                  NowPrint m_strHK202CP09, "02", "40", False, strUserNum
                  PUB_PrintLetter m_strHK202CP09
               Else
                  NowPrint m_strHK202CP09, "02", "40", True, strUserNum, , , , , , , , , , , , , m_strHK202CP09
                  frm1105_1.m_RecNo = m_strHK202CP09
                  frm1105_1.m_PdfName = PUB_CaseNo2FileName(Text1, Text2, Text3, Text4) & ".202.DATA.PDF"
                  frm1105_1.m_Subject = m_Subject
                  frm1105_1.Show
               End If
               'end 2016/6/8
            End If
         End If
         'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
         'Modified by Lydia 2022/08/15 開放P大陸案
         'If pa(9) = "000" And pa(1) = "P" Then
         'Modified by Lydia 2022/10/11 經查此設定並不適用於外專及日專，故請協助排除FMP案
         'If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" Then
         If (pa(9) = "000" Or pa(9) = "020") And pa(1) = "P" And m_bolFMP = False Then
            'Modified by Morgan 2016/3/18 工程師承辦的都通知
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), stCP10, NewReceiveNo
            'Modified by Lydia 2017/03/29 模組已改成，已收文未發文的承辦人全部發mail通知
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), stCP10, IIf(m_bolEngCase, "", NewReceiveNo)
            'Modified by Lydia 2022/08/16 +申請國家
            'PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), stCP10, NewReceiveNo
            PUB_TaiwanCInputMsg pa(1), pa(2), pa(3), pa(4), stCP10, pa(9), NewReceiveNo
            'end 2016/3/18
         End If
   
         'Add By Sindy 2016/10/5
         If Me.m_strIR01 <> "" Then
            Unload frm04010502_1
            Unload frm04010502_2
            Unload Me
            'Modify By Sindy 2022/5/20
            'frm04010519.GoNext
            Forms(0).Tmpfrm04010519.GoNext
            Set Forms(0).Tmpfrm04010519 = Nothing
            '2022/5/20 END
         ElseIf Me.m_DocNo <> "" Then
         'Added by Morgan 2014/1/14
         'If Me.m_DocNo <> "" Then
         '2016/10/5 END
            Unload frm04010502_1
            Unload frm04010502_2
            Unload Me
            frm04010516.GoNext
         Else
         'end 2014/1/14
         
            strKey1 = "1"
            Unload frm04010502_2
            Unload Me
            frm04010502_1.Show
            frm04010502_1.Clear
            
         End If 'Added by Morgan 2014/1/14
         
      Case 1
         frm04010502_2.Show
         Unload Me
      Case 2
         Unload frm04010502_1
         Unload frm04010502_2
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer, intStep As Integer, strTxt(1 To 20) As String, j As Integer
   Dim strCe(99) As String, bolChk As Boolean, lMax As Long
   Dim strTmp(1 To 5) As String
   Dim strDate As String
   '************** 90.11.13   nick
   Dim StrlMAXbyNick As String
   Dim VarlMaxByNick As Variant
   Dim jjjbyNick As Integer
   '*******************************
   
   'Add by Morgan 2004/2/9
   'Dim stCP12 As String, stCP13 As String 'Removed by Morgan 2015/11/5 改全域變數
   Dim stCP115 As String 'Add by Morgan 2007/6/12
   'Dim stCP10 As String 'Add by Morgan 2007/10/25 來函案件性質
   Dim st307Msg As String '分割案提醒訊息
   Dim bUpdate411 As Boolean 'Add by Morgan 2009/9/1 是否取消催審期限
   
   'Add by Morgan 2009/10/2
   Dim strBillNo As String '帳單編號
   Dim stErrMsg As String '錯誤訊息
   Dim stCP14  As String, stCP27  As String, stCP48 As String '承辦人,發文日,承辦期限
   Dim m_HK_111Cp06 As String, m_HK_111Cp07 As String
   Dim stCP133 As String, stCP134 As String, stNP23 As String
   Dim strMsg As String
   Dim stCP53 As String, stCP54 As String 'Add by Morgan 2010/6/22 年費年度起迄
   Dim stEP06, stCP06 As String      '文件齊備日,本所期限   '2010/7/6 add by sonia
   Dim strDualAppMsg As String 'Added by Morgan 2013/1/11 一案兩請發明核准提醒
   Dim str941ReceiveNo As String, str941CP06 As String 'Add By Sindy 2013/4/2
   Dim bolRegMail As Boolean '是否掛號'Added by Morgan 2014/4/17
   'Dim bolAdd941 As Boolean '是否內部收文分析 Added by Morgan 2014/12/8 'Removed by Morgan 2016/3/18
   Dim strCP20 As String 'Added by Morgan 2019/8/8
   'Added by Lydia 2020/04/06
   'Dim bolTmp As Boolean, aKind As String 'Remove by Lydia 2020/12/16 取消產生C類接洽記錄單
   'Dim m_strMemo As String 'C類接洽單備註 'Remove by Lydia 2020/12/16 取消產生C類接洽記錄單
   Dim bolDNUPL As Boolean 'Added by Lydia 2023/10/31
   Dim strChk As String 'Added by Lydia 2023/11/16
   Dim stCP07 As String, stCP71 As String, stCP142 As String, stNP15 As String 'Added by Morgan 2024/9/25
   Dim bolDualAppConfirmMail As Boolean 'Added by Morgan 2025/9/26 是否EMail工程師確認一案兩請
   Dim strLetterJudge As String 'Added by Morgan 2025/9/26 信函判發人
   
bUpdate411 = False

'Add by Morgan 2007/10/25
'清除一案兩請資料
m_strDualAppNo = ""
m_strDualAppNP07 = ""
m_strDualAppNP22 = ""
m_str1914CP09 = "" 'Added by Morgan 2014/7/15

    'Added by Lydia 2020/04/06 核准函前一畫面的結果
    'Remove by Lydia 2020/12/16 取消產生C類接洽記錄單
    'If frm04010502_2.Text6 = "1" Then
    '   '部份案件性質之核准1001改為核發1008
    '   If InStr(Patent1001Display, cp(10)) > 0 Then
    '       aKind = "1008"
    '   Else
    '       aKind = "1001"  '核准
    '   End If
    'End If
    'end 2020/04/06
    'end 2020/12/16
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
   FormSave = False
   cnnConnection.BeginTrans

   ' 90.07.10 modify by louis (申請案核准日)
   strDate = Empty
   '************** 90.11.13    nick
   StrlMAXbyNick = ""
   
   If IsEmptyText(Text5(0)) = False Then
     strDate = DBDATE(Text5(0))
   End If
 
   intStep = 1

   strExc(0) = Empty
   strExc(0) = strExc(0) & "PA17='" & Me.Text5(5).Text & "',"
   'Modified by Morgan 2012/2/16 大陸復審核准不要更新基本檔准駁--玲玲 P-084471
   'If (cp(10) >= "101" And cp(10) <= "105") Or cp(10) = "107" Or cp(10) = "503" Or cp(10) = "504" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "802" Or cp(10) = "804" Then
   'Modified by Morgan 2012/12/19 +衍生設計125,改請衍生設計308
   '2013/10/24 MODIFY BY SONIA 再加入卷宗性質判斷pa(23) = "1",P-083407的503不可更新,否則後續改變原處分也不會更新
   'Modified by Morgan 2020/12/18 改寫函數判斷以便共用及修改
   'If pa(23) = "1" And ((cp(10) >= "101" And cp(10) <= "105") Or (pa(9) <> "020" And cp(10) = "107") Or cp(10) = "125" Or cp(10) = "308" Or cp(10) = "503" Or cp(10) = "504" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "802" Or cp(10) = "804") Then
   'Modified by Morgan 2024/9/18 判斷及變數設定改在Form_Load
   'If pa(23) = "1" And PUB_ChkIsRltPty(cp(1), cp(10), pa(9)) = True Then
   '   bolChgRlt = True 'Added by Morgan 2024/6/4
   If bolChgRlt Then
   'end 2024/9/18
      'Add by Morgan 2007/8/15 台灣案或沒有選通知書時才更新
      'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
      'If pa(9) = 台灣國家代號 Or Text5(6) = "" Then
         strExc(0) = strExc(0) & "PA16='" & Me.Text5(1).Text & "',"
         'Modify by Morgan 2004/11/30 爭議程序不更新基本檔准駁日
         'If IsEmptyText(Me.Text5(0).Text) = False Then
         'Modify by Morgan 2006/5/22 申請案才回寫
         'If IsEmptyText(Me.Text5(0).Text) = False And Not (Val(cp(10)) >= 801 And Val(cp(10)) <= 805) Then
         If IsEmptyText(Me.Text5(0).Text) = False And pa(23) = "1" Then
            strExc(0) = strExc(0) & "PA20=" & CNULL(TransDate(Text5(0), 2)) & ","
         End If
         
         'Add by Morgan 2009/7/1
         '大陸核准日+5個月作為預定公告更新到多國案期限
         If pa(9) = "020" And Text5(0) <> "" Then
            strExc(1) = CompDate(1, 5, Text5(0))
            PUB_UpdCP07byPA14 pa, strExc(1), strMsg
            'Memo by Lydia 2015/09/09 大陸發明案之核准同時更新大陸案之領證期限至澳門發明進度,寫在下面寫入領證期限之後 (PUB_UpdCP07by020)
         End If
      'End If
      'end 2009/11/27
      'end 2007/8/15
      'Memo by Lydia 2015/09/09 大陸發明案之核准，若該案有澳門案則同時將大陸案之領證期限更新至澳門發明進度之法定期限 (PUB_UpdCP07by020)
   End If
   
   '2008/3/5 ADD BY SONIA 放棄專利權429核准且未閉卷時閉卷 P-068288
   If cp(10) = "429" And pa(57) = "" Then
      strExc(0) = strExc(0) & "PA57='Y',PA59='99',PA58=" & strSrvDate(1) & ","
   End If
   'end 2007/5/4
   
   'Added by Morgan 2023/3/3
   '自請撤回413核准閉卷
   If m_Close413 = "Y" And pa(57) = "" Then
      strExc(0) = strExc(0) & "PA57='Y',PA59='09',PA58=" & strSrvDate(1) & ","
      
      strExc(2) = AutoNo("B", 6) 'B類總收文號
      strSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,CP44,cp45,cp46,cp57,cp58,cp116) values " & _
         " ('" & cp(1) & "','" & cp(2) & "','" & cp(3) & "','" & cp(4) & "'," & strSrvDate(1) & _
         ",'" & strExc(2) & "','913','" & stCP12 & "','" & stCP13 & "','" & strUserNum & "','N','N'," & strSrvDate(1) & _
         ",'N','" & cp(9) & "','" & cp(44) & "','" & cp(45) & "','" & cp(46) & "'," & strSrvDate(1) & ",'09','" & cp(116) & "')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2023/3/3
   
   'Added by Morgan 2023/2/23
   If pa(9) = "000" And m_strCP10 = "415" And txt415Date <> "" Then
      strExc(0) = strExc(0) & "PA25=" & DBDATE(txt415Date) & ","
   End If
   'end 2023/2/23
   
   'Added by Morgan 2024/9/25
   stCP07 = "NULL": stCP142 = "NULL"
   If bolCN445 Then
      stCP142 = DBDATE(txt415Date)
      strExc(1) = CompDate(2, -1, txt415Date)
      strExc(0) = strExc(0) & "PA25=" & strExc(1) & ","
      '大陸專利權補償天數滿一年以上者要管制補償期年費(加上補償天數後的專利權期滿終止日>=原專利權期滿終止日+1年; 原專利權期滿終止日=原專用期止日+1天))
      If DBDATE(txt415Date) >= CompDate(0, 1, CompDate(2, 1, pa(25))) Then
         'Modified by Morgan 2024/10/23
         'stCP07 = CompDate(2, 1, pa(25)) '補償期年費法限=原專用期止日-1
         stCP07 = DBDATE(pa(25)) '補償期年費法限=原專用期止日
      End If
   End If
   'end 2024/9/25
   
   'Modify by Amy 2014/03/24
   '因應台日優先權證明文件,P案申請國為台灣或大陸,種類為發明或新型,案件性質為405 or 436 則將優先權存取碼 存於專利基本檔優先權存取碼及進度檔cp10="1001"的備註欄
   strSql = ""
   'Modify by Amy 2014/07/09 有輸優先權存取碼才存
   'If Text1 = "P" And (pa(9) = 台灣國家代號 Or pa(9) = 大陸國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
   If Trim(txtPriNo) <> "" Then
       strExc(0) = strExc(0) & "PA164='" & txtPriNo & "',"
   End If
   
   'Added by Morgan 2021/6/29 大陸發明生化生醫案設定是否新藥專利
   If pa(9) = "020" And pa(8) = "1" And pa(158) = "3" And m_PA176 <> pa(176) Then
      strExc(0) = strExc(0) & "PA176='" & m_PA176 & "',"
   End If
   'end 2021/6/29
   
   If strExc(0) <> "" Then
      If Right(strExc(0), 1) = "," Then strExc(0) = Left(strExc(0), Len(strExc(0)) - 1)
      strTxt(intStep) = "UPDATE PATENT SET " & strExc(0) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
        'Add By Cheng 2002/11/06
        cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
     
   End If
   'end 2014/03/24
   
   '1
   If frm04010502_2.Text6 = "1" Then
      If IsEmptyText(strDate) = False Then
         If Left(cp(10), 1) = "1" Or Left(cp(10), 1) = "3" Then
            'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & TransDate(Text5(5), 2) & ",CP35=" & CNULL(Text5(3)) & " WHERE "
            'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
            'If Text5(6) = "" Then
               '2005/10/19 MODIFY BY SONIA 不判斷 CP25
               'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & TransDate(Text5(0), 2) & ",CP08=" & CNULL(Text5(2)) & ",CP35=" & CNULL(Text5(3)) & " WHERE " & _
               '   "CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
               'Modify by Morgan 2008/5/13 +CP117
               strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & TransDate(Text5(0), 2) & ",CP08=" & CNULL(Text5(2)) & ",CP35=" & CNULL(Text5(3)) & ",CP117=" & CNULL(Text5(16)) & " WHERE " & _
                  "CP09='" & strReceiveNo & "' AND CP24 IS NULL"
               '2005/10/19 END
               'Add By Cheng 2002/11/06
               cnnConnection.Execute strTxt(intStep)
               intStep = intStep + 1
               
               bUpdate411 = True
            'End If
            'end 2009/11/27
         End If
      End If
      
      If Left(cp(10), 1) <> "1" And Left(cp(10), 1) <> "3" Then
         '2005/10/19 MODIFY BY SONIA 不判斷 CP25
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & TransDate(Label2(3), 2) & ",CP08=" & CNULL(Text5(2)) & ",CP35=" & CNULL(Text5(3)) & " WHERE " & _
         '   "CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
         'Modify by Morgan 2008/5/13 +CP117
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & TransDate(Label2(3), 2) & ",CP08=" & CNULL(Text5(2)) & ",CP35=" & CNULL(Text5(3)) & ",CP117=" & CNULL(Text5(16)) & " WHERE " & _
            "CP09='" & strReceiveNo & "' AND CP24 IS NULL"
         '2005/10/19 END
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
         
         bUpdate411 = True
         
         If cp(10) = "701" Then
            strTxt(intStep) = "UPDATE PATENT SET PA23=1 WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
         'Add By Cheng 2002/01/11
         If cp(10) = 專利權讓與 Then
            strTxt(intStep) = "UPDATE PATENT SET PA23=1 WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            'Add By Cheng 2002/11/06
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
         End If
      End If
      
      'Modify by Morgan 2004/9/10
      If cp(10) = 延緩公告 Then
         stCP10 = 准予延緩公告
         
      'Added by Morgan 2024/6/20
      ElseIf cp(10) = "245" Then
         stCP10 = 1924 '准予延緩審查
      'end 2024/6/20
      
      'Modified by Lydia 2015/10/02 部份案件性質之核准1001改為核發1008
      ElseIf InStr(Patent1001Display, cp(10)) > 0 Then
         stCP10 = "1008"
      'end 2015/10/02
      Else
         stCP10 = 核准
      End If
      
   Else
      '2006/8/15 MODIFY BY SONIA 加入大陸案P-65644,但大陸仍存核准
      'stCP10 = 改變原處分
      If pa(9) = 大陸國家代號 Then
         stCP10 = 核准
      Else
         stCP10 = 改變原處分
      End If
      '2006/8/15 END
   End If
    
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'   'Add By Cheng 2003/04/16
'   '若申請國家非台灣
'   If pa(9) <> 台灣國家代號 Then
'      '判斷大陸香港發明通知書欄位
'      Select Case Me.Text5(6).Text
'         Case "1"
'            stCP10 = "1213"
'         Case "2"
'            stCP10 = "1207"
'         Case "3"
'            stCP10 = "1214"
'         Case "4"
'            stCP10 = "1215"
'         Case "5"
'            stCP10 = "1204"
'         '2006/2/7 ADD BY SONIA香港公告
'         Case "6"
'            stCP10 = "1208"
'      End Select
'   End If
'
'   'MODIFY BY SONIA 90.10.21
'   If Text5(6) = "1" Then
'      strTmp(1) = "初步審查合格通知書"
'   End If
'   If Text5(6) = "2" Then
'      strTmp(1) = "公布通知書"
'   End If
'   'Add By Cheng 2002/06/21
'   If Me.Text5(6).Text = "3" Then
'      strTmp(1) = "初步審查合格及進入實質審查程序通知書"
'   End If
'   If Me.Text5(6).Text = "4" Then
'      strTmp(1) = "公布及進入實質審查程序通知書"
'   End If
'   '93.1.6 ADD BY SONIA 93.1.6
'   If Text5(6) = "5" Then
'      strTmp(1) = "進入實質審查程序通知書"
'   End If
'   '93.1.6 END
'   '2006/2/7 ADD BY SONIA 93.1.6
'   If Text5(6) = "6" Then
'      'Modify by Morgan 2006/10/14 加澳門
'      'strTmp(1) = "香港公告"
'      If pa(9) = "013" Then
'         strTmp(1) = "香港公告"
'      ElseIf pa(9) = "044" Then
'         strTmp(1) = "澳門公告"
'      Else
'         strTmp(1) = "公告"
'      End If
'   End If
'   '2006/2/7 END
'
'end 2009/11/27
   
   stNP15 = "" 'Added by Morgan 2024/9/25
   'Add by Morgan 2004/9/10
   If txt412.Visible = True Then
      'Added by Morgan 2024/9/25
      If bolCN445 Then
         strTmp(1) = lbl412 & txt412.Text & "天"
         stNP15 = strTmp(1)
      Else
      'end 2024/9/25
         strTmp(1) = lbl412 & txt412.Text
      End If
   End If
   
   '3
   NewReceiveNo = AutoNo("C", 6)
   'MODIFY BY SONIA 90.11.27 承辦人應為操作者,智權人員存最近收文A類接洽記錄單的智權人員
   'Modify by Morgan 2004/2/9
   'strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
     "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP64) VALUES ('" & Text1 & "','" & Text2 & "','" & _
     Text3 & "','" & Text4 & "'," & TransDate(Label2(3), 2) & "," & _
     CNULL(Text5(2)) & ",'" & NewReceiveNo & "','" & stCP10 & "'," & CNULL(cp(12)) & "," & _
     CNULL(PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)) & "," & strUserNum & ",'','N','N'," & strSrvDate(1) & ",'" & strReceiveNo & "'," & CNULL(ChgSQL(strTmp(1))) & ")"
   
   'Removed by Morgan 2015/11/5 改在ReadPatent
   'stCP13 = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
   'stCP12 = GetSalesArea(stCP13)
   
   'Modify by Morgan 2007/6/12 加CP115
   If txtDispDate.Visible = True Then
      stCP115 = DBDATE(txtDispDate)
   Else
      stCP115 = "NULL"
   End If
   
   'Add by Morgan 2009/12/1 +CP133,CP134
   If pa(9) <> "000" And Text5(0) <> "" Then
      stCP133 = DBDATE(Text5(0))
      stCP134 = 2
   Else
      stCP133 = "NULL"
      stCP134 = "NULL"
   End If
      
   'Add by Morgan 2009/10/2
   '銷催審期限(由 Trigger 處理)
   '大陸檢索報告421,新穎性調查426的核准預設原承辦不上發文日,後面加上文件齊備日
   '2010/12/1 modify by sonia 加423申請專利權評價報告P-092713
   stCP48 = "NULL"    '2009/11/9 ADD BY SONIA
   stEP06 = "NULL": stCP06 = "NULL" '2010/7/6 add by sonia P-083646
   If pa(9) = 大陸國家代號 And (cp(10) = "421" Or cp(10) = "426" Or cp(10) = "423") Then
      stCP14 = cp(14)
      'Added by Morgan 2025/5/19 若原承辦為程序時改抓新申請案承辦人，若也是程序則再改抓國內案工程師，若離職則再改掛李柏翰
      If PUB_GetST03(stCP14) = "P12" Then
         strExc(0) = "select cp14 from caseprogress,staff where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
            " and cp10 in (" & NewCasePtyList & ") and st01(+)=cp14 and st03<>'P12'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            stCP14 = RsTemp(0)
         Else
            stCP14 = PUB_GetInCaseCP14(cp(1), cp(2), cp(3), cp(4))
         End If
      End If
      If GetStaffName(stCP14) = "" Then
         stCP14 = "99050"
         
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc13)" & _
            " values('" & strUserNum & "','" & stCP14 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'分案通知','" & NewReceiveNo & "')"
         cnnConnection.Execute strSql, intI
      End If
      'end 2025/5/19
      
      stCP27 = "NULL"
      'Modify by Morgan 2010/10/19
      'stCP48 = CompDate(2, 7, strSrvDate(1)) '2009/11/9 ADD BY SONIA 系統日加7天為承辦期限
      'stCP06 = stCP48: stEP06 = strSrvDate(1) '2010/7/6 ADD BY SONIA  P-083646
      stCP06 = CompDate(2, 7, strSrvDate(1))
      stCP06 = PUB_GetWorkDay1(stCP06, False) 'Add by Morgan 2011/3/29 要抓工作天
      If m_bolFMP Or PUB_IfSetCP48() Then
         stCP48 = stCP06
      End If
      stEP06 = strSrvDate(1)
      'end 2010/10/19
      
   Else
      '香港檢索報告
      If pa(9) = "013" And cp(10) = "421" Then
         '更新B類收文補文件且相關收文號為申請檢索報告的發文日
         strExc(0) = "select cp09 from caseprogress where cp43='" & cp(9) & "' and cp10='202' and cp09>'B' and cp57 is null and cp27 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_strHK202CP09 = RsTemp.Fields(0)
            strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_strHK202CP09 & "'"
            cnnConnection.Execute strSql, intI
            
            'Added by Morgan 2016/6/8
            If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
               'Modified by Morgan 2018/8/1
               'strExc(1) = PUB_GetLetterJudge(pa(1), "202", , pa(9), pa(1), pa(2), pa(3), pa(4))
               strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "202", pa(9), , , m_bolFMP)
               '附件:香港短期專利申請檢索報告
               PUB_AddLetterProgress m_strHK202CP09, 1, True, strExc(1), False, pa(26), "202", pa(75)
            End If
            
            If Left(Pub_StrUserSt03, 1) <> "F" Then
               strExc(1) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
               m_Subject = "委託 " & strExc(1) & " 案提交香港短期專利申請檢索報告"
               If ExistCheck("AppForm", "AF01", m_strHK202CP09, "", False) = False Then
                  'Modified by Morgan 2018/7/30 指示信判發人改抓設定檔
                  strExc(2) = PUB_GetLetterJudgeNew("2", pa(1), "202", pa(9), "421")
                  PUB_AddAppForm m_strHK202CP09, True, strExc(2), m_Subject
               End If
            End If
            'end 2016/6/8
         End If
      End If
   'end 2009/10/2
   
      stCP14 = strUserNum
      
      'Added by Morgan 2014/12/8 P台灣案舉發及舉發答辯審定書自動內部收文分析
      'Modified by Morgan 2015/6/25 +501訴願,505參加訴願
      'Modified by Morgan 2016/3/18
      'If pa(9) = 台灣國家代號 And (cp(10) = "803" Or cp(10) = "804" Or cp(10) = "501" Or cp(10) = "505") Then
      '   bolAdd941 = True
      'Memo by Morgan 2018/10/4 已改非台灣案也適用
      If m_bolEngCase Then
         If GetStaffName(cp(14)) <> "" Then
            stCP14 = cp(14)
            'add by sonia 2024/7/15 A7010柯昱安調離也要改為游經理73022
            If GetStaffDepartment(stCP14) >= "P10" And GetStaffDepartment(stCP14) <= "P11" Then
            Else
               'Modified by Morgan 2025/2/21 73022->left(pub_PMan,5)
               pub_PMan = Pub_GetSpecMan("專利處特定編號")
               stCP14 = Left(pub_PMan, 5)
               'end 2025/2/21
            End If
            'end 2024/7/15
         Else
            'Modified by Morgan 2025/2/21 73022->left(pub_PMan,5)
            pub_PMan = Pub_GetSpecMan("專利處特定編號")
            stCP14 = Left(pub_PMan, 5)
            'end 2025/2/21
         End If
      'end 2016/3/18
         stCP27 = "NULL"
      'Added by Morgan 2020/1/17
      ElseIf m_bolNoCP27 = True Then
         stCP27 = "NULL"
      'end 2020/1/17
      Else
         stCP27 = strSrvDate(1)
      End If
      'end 2014/12/8
      
   End If
   
   'Modify by Morgan 2008/5/20 +CP35,CP117
   '2009/11/9 MODIFY BY SONIA 加CP48承辦期限
   'Modify by Morgan 2009/11/27 +CP133官方發文日,CP134官方期限月數
   'Modify by Morgan 2010/6/22 +大陸年費報價年度起迄 CP53,CP54
   stCP53 = "": stCP54 = ""
   If InStr(NewCasePtyList, cp(10)) > 0 Then     '2010/7/6 ADD BY SONIA 新案案件性質才要放CP53,CP54
      If Val(Text9) > 0 Then
         stCP53 = Val(Text9)
         If m_Year = "Y" Then
            stCP54 = Val(stCP53) + 1
         Else
            stCP54 = stCP53
         End If
      End If
   End If        '2010/7/6 ADD BY SONIA
   
   'Add by Amy 2014/03/24
   'Modify by Amy 2014/07/09 有輸優先權存取碼才存
   'If Text1 = "P" And (pa(9) = 台灣國家代號 Or pa(9) = 大陸國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
   If Trim(txtPriNo) <> "" Then
        strTmp(1) = strTmp(1) & Format(Label2(3), "###/##/##") & "核准優先權存取碼:" & txtPriNo & ";"
   End If
   'end 2014/03/24
   
   'Added by Morgan 2019/5/27 備註＋IDS報價
   If m_USCaseNo <> "" Then
      'Modified by Morgan 2019/6/3 第１階段報價金額大於０才寫
      'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
      If Val(txtIDSFee(1)) > 0 Then
         strTmp(1) = "IDS報價:1.第一階段 " & txtIDSFee(1) & "(" & txtIDSPt(1) & "P), 2.第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strTmp(1)
      Else
         strTmp(1) = "IDS報價:第二階段 " & txtIDSFee(2) & "(" & txtIDSPt(2) & "P);" & strTmp(1)
      End If
   End If
   'end 2019/5/27
   
   'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
   strCP20 = ""
   If m_bolFMP Then
      strCP20 = PUB_GetCP20(pa(1), stCP10, , pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   End If
   
   'Added by Morgan 2024/9/25
   If bolCN445 Then
      stCP71 = "'" & txt412 & "'"
      If Val(stCP07) > 0 Then
         'Memo by Lydia 2025/10/29 在下面套新規則
         stCP06 = PUB_GetWorkDay1(CompDate(2, -10, stCP07), True)
         If m_bolFMP Then
            stNP23 = stCP06
         Else
            stNP23 = "NULL"
         End If
      End If
   Else
      stCP71 = "NULL"
   End If
   'end 2024/9/25
   'Added by Lydia 2025/10/29
   If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
      If Val(stCP07) > 0 Then
         stCP06 = PUB_GetPOurDeadline(stCP07, pa(9), stNP23, pa(1), "615")  '案件性質用615補償期年費---參考後面:大陸專利權補償天數滿一年以上者要管制補償期年費
      End If
   End If
   'end 2025/10/29
   
   '2010/7/6 MODIFY BY SONIA 加CP06
   'Modified by Morgan 2012/4/30 +cp119=櫃檯收文日
   'Modified by Morgan 2024/9/25 +cp07,cp71(補償天數),cp142
   strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08,CP09,CP10," & _
     "CP12,CP13,CP14,CP20,CP26,CP32,CP27,CP43,CP64,CP71,CP115,CP35,CP117,CP48,CP133,CP134" & _
     ",CP53,CP54,cp119,CP142) VALUES ('" & Text1 & "','" & Text2 & "','" & _
     Text3 & "','" & Text4 & "'," & TransDate(Label2(3), 2) & "," & stCP06 & "," & stCP07 & "," & _
     CNULL(Text5(2)) & ",'" & NewReceiveNo & "','" & stCP10 & "'," & CNULL(stCP12) & "," & _
     CNULL(stCP13) & ",'" & stCP14 & "','" & strCP20 & "','N','N'," & stCP27 & ",'" & strReceiveNo & "'" & _
     "," & CNULL(ChgSQL(strTmp(1))) & "," & stCP71 & "," & stCP115 & "," & CNULL(Text5(3)) & _
     "," & CNULL(Text5(16)) & "," & stCP48 & "," & stCP133 & "," & stCP134 & ",'" & stCP53 & "','" & stCP54 & "'," & DBDATE(Label2(3)) & "," & stCP142 & ")"
   'end 2010/6/22
   cnnConnection.Execute strTxt(intStep), intI
   intStep = intStep + 1
   
   'Add By Sindy 2013/4/2 台灣舉發及答辯自動產生一道分析(順德案件亦同),原工程師離職掛王副總
   'Modified by Morgan 2014/12/8 判斷移到前面以便控制來函發文日
   'If pa(9) = 台灣國家代號 And (cp(10) = "803" Or cp(10) = "804") Then
   
'Modified by Morgan 2016/3/18
'   If bolAdd941 = True Then
'   'end 2014/12/8
'      str941CP14 = cp(14)
'      strExc(0) = "SELECT ST04,DECODE(ST04,'1',ST06,'1') ST06 FROM STAFF WHERE ST01='" & str941CP14 & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If "" & RsTemp(0).Value <> "1" Then str941CP14 = "71011" '原工程師離職掛王副總
'      End If
'
'      str941ReceiveNo = AutoNo("B", 6)
'      '本所期限為系統日+3個工作天
'      str941CP06 = CompWorkDay(3, strSrvDate(1), 0)
'      'Modified by Morgan 2014/6/24 承辦人加單引號(原員工號都是數字所以才沒有錯誤)
'      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06," & _
'         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48) VALUES " & _
'         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & str941CP06 & _
'         ",'" & str941ReceiveNo & "','941','90'," & CNULL(GetSalesArea(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4)))) & "," & CNULL(PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))) & _
'         ",'" & str941CP14 & "','N','N','N','" & NewReceiveNo & "'," & str941CP06 & ") "
'      cnnConnection.Execute strSql
'
''      If "" & RsTemp(1).Value <> "1" Then '分所
''         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & CompWorkDay(2, strSrvDate(1), 0) & " WHERE EP02='" & str941ReceiveNo & "'"
''      Else
'         strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & str941ReceiveNo & "'"
''      End If
'      cnnConnection.Execute strSql
'
'      '更新承辦期限=本所期限,因ENGINEERPROGRESS_BEFORE5及CASEPROGRESS_AFTER6會造成承辦期限=本所期限-1天
'      strSql = "UPDATE CASEPROGRESS SET CP48=CP06 WHERE CP09='" & str941ReceiveNo & "'"
'      cnnConnection.Execute strSql
   If m_bolEngCase Then
      '不會稿,判發人73022
      'Modified by Morgan 2025/2/21 73022->left(pub_PMan,5)
      pub_PMan = Pub_GetSpecMan("專利處特定編號")
      strSql = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & ",EP34='N',EP40='" & Left(pub_PMan, 5) & "' WHERE EP02='" & NewReceiveNo & "'"
      'end 2025/2/21
      cnnConnection.Execute strSql
      
      '承辦期限=系統日+3個工作天(沿用原規則)
      str941CP06 = CompWorkDay(3, strSrvDate(1), 0)
      strSql = "UPDATE CASEPROGRESS SET CP48=" & str941CP06 & " WHERE CP09='" & NewReceiveNo & "'"
      cnnConnection.Execute strSql
'end 2016/3/18
   End If
   '2013/4/2 End
   
   '2010/7/6 ADD BY SONIA 大陸檢索報告421,新穎性調查426的核准不上發文日,加文件齊備日
   '2010/12/1 modify by sonia 加423申請專利權評價報告P-092713
   If pa(9) = 大陸國家代號 And (cp(10) = "421" Or cp(10) = "426" Or cp(10) = "423") Then
      strTxt(intStep) = "update engineerprogress set ep06=" & strSrvDate(1) & " where ep02='" & NewReceiveNo & "' "
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
   End If
   '2010/7/6 END
   
   '92.11.18 ADD BY SONIA
   If stCP10 = 改變原處分 Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1' WHERE CP09='" & NewReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      
      bUpdate411 = True
   End If
   '92.11.18 END
   '4
   
'Remove by Morgan 2009/10/2 已改由 Trigger 處理
'
'   'Modify by Morgan 2009/9/1 有更新進度檔准駁的才要取消催審期限
'   'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07=" & 催審 & " "
'   'cnnConnection.Execute strTxt(intStep)
'   'intStep = intStep + 1
'   If bUpdate411 = True Then
'      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07=" & 催審 & " and np06 is null"
'      cnnConnection.Execute strTxt(intStep)
'      intStep = intStep + 1
'   End If
'   'end 2009/9/1
'
'end 2009/10/2
   
   '5
   If frm04010502_2.Text6 = "1" And pa(9) = 大陸國家代號 And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "104" Or cp(10) = "105" Or cp(10) = "307") Then
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07=" & 維持費 & " "
      'Add By Cheng 2002/11/06
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   '6
   If frm04010502_2.Text6 = "2" Then
      'Modify by Morgan 2011/10/12 改以本所案號更新(同外專)
      'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07=" & 改變原處分 & " "
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP02='" & pa(1) & "' and NP03='" & pa(2) & "' and NP04='" & pa(3) & "' and NP05='" & cp(4) & "' and NP06 is null AND NP07=" & 改變原處分 & " "
      'Add By Cheng 2002/11/06
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   'Added by Morgan 2024/9/25
   '大陸專利權補償天數滿一年以上者要管制補償期年費
   bol615NP = False 'Added by Morgan 2025/3/10
   If bolCN445 And Val(stCP07) > 0 Then
      lMax = GetNextProgressNo
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP15,NP22,NP23) " & _
         "VALUES ('" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "','615'," & _
         stCP06 & "," & stCP07 & ",'" & stCP13 & "'," & CNULL(Text5(2)) & ",'" & stNP15 & "'," & _
         lMax & "," & stNP23 & ")"
      cnnConnection.Execute strTxt(intStep), intI
      intStep = intStep + 1
      bol615NP = True 'Added by Morgan 2025/3/10
   End If
   'end 2024/9/25
   
   m_strRetSheetNP07 = "" 'Add by Morgan 2005/11/16
   '7
   If IsEmptyText(Text5(10)) = False Then
      If pa(9) = 大陸國家代號 Then
         'Modify by Morgan 2009/11/27 非台灣通知書移到 frm04010514
         'If pa(8) = "1" And Text5(6) = "" Then
         If pa(8) = "1" Then
            strTxt(intStep) = "DELETE FROM CASEMAP WHERE " & ChgCaseMap(pa(1) & pa(2) & pa(3) & pa(4), 0, 1) & " AND CM10='2'"
            cnnConnection.Execute strTxt(intStep)
            intStep = intStep + 1
            '92.11.25 CANCEL BY SONIA
            'strTmp(1) = CompDate(1, 3, TransDate(Label2(3), 2))
            'strTxt(intStep) = "UPDATE PATENT SET PA21=" & strTmp(1) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            ''Add By Cheng 2002/11/06
            'cnnConnection.Execute strTxt(intStep)
            'intStep = intStep + 1
            '92.11.25 END
         End If
      End If
      
      ' 90.12.19 modify by louis
      Select Case cp(10)
         Case "101", "102", "103", "104", "105", "307"
            If pa(9) = 大陸國家代號 Then
               'Add by Morgan 2009/11/27
               'FMP 大陸核准更新或新增香港案的批准紀錄請求期限,法限=核准日+6個月,所限=法限-1個月5天
               'Modified by Morgan 2013/12/19 考慮分割案核准,改判斷專利種類為發明--玲玲 Ex.P-94572
               'If m_bolFMP And cp(10) = "101" And m_HaveHK Then
               If m_bolFMP And pa(8) = "1" And m_HaveHK Then
                  m_HK_111Cp07 = CompDate(1, 6, Text5(0).Text)
                  m_HK_111Cp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, m_HK_111Cp07)), True)
                  strSql = "update caseprogress set cp06=" & m_HK_111Cp06 & ",cp07=" & m_HK_111Cp07 & _
                     " where cp01='" & m_HK_CP01 & "' and cp02='" & m_HK_CP02 & "'" & _
                     " and cp03='" & m_HK_CP03 & "' and cp04='" & m_HK_CP04 & "' and cp10='111' and cp27 is null"
                  cnnConnection.Execute strSql, intI
                  If intI = 0 Then
                     strSql = "update nextprogress set np08=" & m_HK_111Cp06 & ",np09=" & m_HK_111Cp07 & _
                        " where np02='" & m_HK_CP01 & "' and np03='" & m_HK_CP02 & "'" & _
                        " and np04='" & m_HK_CP03 & "' and np05='" & m_HK_CP04 & "' and np06||np07='111'"
                     cnnConnection.Execute strSql, intI
                     If intI = 0 Then
                        strSql = "INSERT INTO NEXTPROGRESS(NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
                           " SELECT '" & NewReceiveNo & "','" & m_HK_CP01 & "','" & m_HK_CP02 & "','" & m_HK_CP03 & "','" & m_HK_CP04 & "','111'" & _
                           "," & m_HK_111Cp06 & "," & m_HK_111Cp07 & ",'" & stCP13 & "'" & _
                           ",NP22 FROM (SELECT MAX(NP22)+1 NP22 FROM NEXTPROGRESS) X"
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
               End If
               If m_bolFMP Then
                  stNP23 = DBDATE(Text5(10))
               Else
                  'Added by Lydia 2025/10/29
                  If strSrvDate(1) >= 內專本所約定期限啟用日 Then
                     strSql = PUB_GetPOurDeadline(DBDATE(Text5(11)), pa(1), stNP23, pa(1), 領證及繳年費)
                  Else
                  'end 2025/10/29
                     stNP23 = "NULL"
                  End If 'Added by Lydia 2025/10/29
               End If
               
               lMax = GetNextProgressNo
               m_strRetSheetNP07 = 領證及繳年費 'Add by Morgan 2005/11/16
               '智權人員存最近收文A類接洽記錄單的智權人員
               strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP22,NP23) " & _
                  "VALUES ('" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & m_strRetSheetNP07 & "," & _
                  TransDate(Text5(10), 2) & "," & TransDate(Text5(11), 2) & ",'" & stCP13 & "'," & CNULL(Text5(2)) & "," & _
                  lMax & "," & stNP23 & ")"
               cnnConnection.Execute strTxt(intStep), intI
               intStep = intStep + 1
               
               '2010/12/20 ADD BY SONIA 陳玲玲要求同時更新核准來函期限
               strTxt(intStep) = "UPDATE CASEPROGRESS SET CP06=" & TransDate(Text5(10), 2) & ",CP07=" & TransDate(Text5(11), 2) & " WHERE CP09='" & NewReceiveNo & "'"
               cnnConnection.Execute strTxt(intStep), intI
               intStep = intStep + 1
               '2010/12/20 END
               StrlMAXbyNick = StrlMAXbyNick & lMax & ","
               
               bolRegMail = True 'Added by Morgan 2014/4/17
            End If
            'Added by Lydia 2015/09/09 大陸發明案之核准，若該案有澳門案則同時將大陸案之領證期限更新至澳門發明進度之法定期限
            If pa(9) = 大陸國家代號 And pa(8) = "1" Then
               Call PUB_UpdCP07by020(pa, m_bolFMP, "5")
            End If
            'end 2015/09/09

            
            'Add by Morgan 2007/10/25 大陸發明核准且為一案兩請若新型案已發證則新增新型案下一程序放棄專利權(429)
            If pa(9) = 大陸國家代號 Then
               If pa(8) = "1" Then
                  strSql = "select cm01,cm02,cm03,cm04" & _
                     " from (select cm01,cm02,cm03,cm04 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'" & _
                     " union select cm05,cm06,cm07,cm08 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3') X,patent" & _
                     " where pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa08='2' and pa21>0 and pa57 is null" & _
                     " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='429' and cp57 is null)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     'Modified by Morgan 2017/1/20 改通知放棄專利權(1914)要先新增(cp30不必寫),放棄專利權(429)期限的相關收文號改放通知放棄專利權的收文號(當刪除1914CP時才會一併刪除429NP,否則會殘留)
                     m_strDualAppNP22 = ""
                     
                     'Added by Morgan 2014/7/15
                     '新增通知放棄專利權
                     m_str1914CP09 = AutoNo("C", 6)
                     strExc(2) = PUB_GetAKindSalesNo(RsTemp(0), RsTemp(1), RsTemp(2), RsTemp(3))
                     strExc(3) = GetSalesArea(strExc(2))
                     
                     'Modified by Morgan 2019/8/8 FMP案的CP20要抓設定
                     strCP20 = "N"
                     If m_bolFMP Then
                        strCP20 = PUB_GetCP20(pa(1), "1914", , pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
                     End If
   
                     strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                        ",cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp32,cp43 ) values ('" & RsTemp(0) & "'" & _
                        ",'" & RsTemp(1) & "','" & RsTemp(2) & "','" & RsTemp(3) & "'," & strSrvDate(1) & _
                        "," & DBDATE(Text5(10)) & "," & DBDATE(Text5(11)) & ",'" & m_str1914CP09 & "','1914','" & strExc(3) & "'" & _
                        ",'" & strExc(2) & "','" & strUserNum & "','" & strCP20 & "','N'," & strSrvDate(1) & ",'" & m_strDualAppNP22 & "','N','" & NewReceiveNo & "')"
                     cnnConnection.Execute strSql, intI
                     'end 2014/7/15
                     
                     lMax = GetNextProgressNo
                     m_strDualAppNP07 = "429"
                     m_strDualAppNP22 = lMax
                     m_strDualAppNo = RsTemp(0) & RsTemp(1) & RsTemp(2) & RsTemp(3)
                     'Added by Lydia 2025/10/29
                     stNP23 = ""
                     If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                        strSql = PUB_GetPOurDeadline(DBDATE(Text5(11)), pa(1), stNP23, pa(1), m_strDualAppNP07)
                     End If
                     'end 2025/10/29
                     
                     'Modified by Lydia 2025/10/29 +NP23
                     strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22,np23) values (" & _
                        " '" & m_str1914CP09 & "','" & RsTemp(0) & "','" & RsTemp(1) & "','" & RsTemp(2) & "','" & RsTemp(3) & "'" & _
                        ",'" & m_strDualAppNP07 & "'," & DBDATE(Text5(10)) & "," & DBDATE(Text5(11)) & _
                        ",'" & strExc(2) & "'," & m_strDualAppNP22 & ", " & CNULL(stNP23, True) & ")"
                     cnnConnection.Execute strSql, intI
                     'end 2017/1/20
                     
                     'Added by Morgan 2016/6/7
                     If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
                        'Modified by Morgan 2018/8/1
                        'strExc(1) = PUB_GetLetterJudge(RsTemp(0), "1914", , "020", RsTemp(0), RsTemp(1), RsTemp(2), RsTemp(3))
                        strExc(1) = PUB_GetLetterJudgeNew("1", RsTemp(0), "1914", "020", , , m_bolFMP)
                        PUB_AddLetterProgress m_str1914CP09, 0, True, strExc(1), True, pa(26), "1914", pa(75)
                     End If
                     'end 2016/6/7
                     
                     'Added by Morgan 2020/2/24
                      If Left(Pub_StrUserSt03, 1) = "F" Then
                         strDualAppMsg = "請和工程師確認本案是否有聲明放棄新型專利權！"
                         '行事曆
                         '管制期限=系統日+2個工作天
                         strExc(1) = CompWorkDay(2, strSrvDate(1))
                         '管制人
                         strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
                         strExc(4) = "工程師是否已回覆程序人員一案二請之結果"
                         '工程師
                         strExc(8) = PUB_GetFCPPromoterNo(cp(9), "1001", cp(14))
                         'end 2019/7/17
                         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(3) & "," & strExc(8), strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)
                         
                         'EMail 承辦工程師,副本:工程師主管、程序管制人員、程序主管
                         '主旨
                         strExc(4) = "請工程師確認" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "(發明案)一案二請事宜"
                         '內文
                         strExc(0) = "工程師：請確認本案是否已有聲明放棄新型專利權(請用此Email 回覆程序人員)" & vbCrLf & _
                                     "　　　　□是: 維持一案兩請之關聯" & vbCrLf & _
                                     "　　　　□否: 本案已非屬相同創作，故可解除一案兩請之設定" & vbCrLf & vbCrLf & _
                                     "程序人員：若工程師回覆結果為" & vbCrLf & _
                                     "　　　　　是: 調新型案的卷宗進行""通知放棄專利權""之程序並解除行事曆" & vbCrLf & _
                                     "　　　　　否: 1. 請解除一案兩請之關聯(請由內專系統→資料處理→關聯案件資料" & vbCrLf & _
                                     "　　　　　　　　 維護→一案兩申請案件資料維護解除關聯)及行事曆期限之設定" & vbCrLf & _
                                     "　　　　　　　2. 請電腦中心刪除""新型案""進度檔之""通知放棄專利權""及下一" & vbCrLf & _
                                     "　　　　　　　　 程序""放棄專利權"""
                            
                         '工程師主管
                         strExc(5) = PUB_GetFCPEngSup(strExc(8))
                         '程序主管
                         strExc(6) = PUB_GetFCPProSup(strExc(3))
                         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                            " values( '" & strUserNum & "','" & strExc(8) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                            ",'" & strExc(4) & "','" & strExc(0) & "','" & strExc(3) & ";" & strExc(5) & ";" & strExc(6) & "')"
                         cnnConnection.Execute strSql, intI
                      'Added by Morgan 2025/9/26
                      Else
                        bolDualAppConfirmMail = True
                      'end 2025/9/26
                      End If
                      'end 2020/2/24
                  End If
               End If
            End If
            
            'Added by Lydia 2021/11/10 大陸案Key核准時，請更新衍生的澳門案之Title(請抓大陸案最新之Title)，並且同時發email通知FMP或寰華案相關人員，兩者Email設定不同。
            If pa(9) = 大陸國家代號 And pa(8) = "1" And (m_bolFMP = True Or m_bolFMP2 = True) Then
               If m_Have044 = True And m_CPto044(1) <> "" Then '澳門案
                     strSql = "Update Patent Set PA05=" & CNULL(ChgSQL(pa(5))) & ",PA06=" & CNULL(ChgSQL(pa(6))) & ", PA07=" & CNULL(ChgSQL(pa(7))) & " " & _
                                 "where pa01='" & m_CPto044(1) & "' and pa02='" & m_CPto044(2) & "' and pa03='" & m_CPto044(3) & "' and pa04='" & m_CPto044(4) & "' "
                     Pub_SeekTbLog strSql
                     cnnConnection.Execute strSql, intI
                     If m_bolFMP2 = True Then  '寰華案
                        '收件者:智權人員;
                        strExc(1) = PUB_GetFCPSalesNo(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                        '副本收受者:智權人員之主管; 程序人員;backup
                        strExc(2) = PUB_GetFCPProSup(strExc(1))
                        strExc(3) = PUB_GetFCPHandler(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                     ElseIf m_bolFMP = True Then
                        '收件者:外專智權人員;
                        strExc(1) = PUB_GetFCPSalesNo(m_CPto044(1), m_CPto044(2), m_CPto044(3), m_CPto044(4))
                        '副本收受者:外專智權人員之主管;內專人員(暫時預設98012品薇);backup
                        strExc(2) = PUB_GetFCPProSup(strExc(1))
                        'Modified by Morgan 2025/1/21
                        'strExc(1) = "98012"
                        If strSrvDate(1) >= P業務區劃分啟用日 Then
                          strExc(1) = PUB_GetPHandler(m_CPto044(1) & m_CPto044(2) & m_CPto044(3) & m_CPto044(4))
                        Else
                          strExc(1) = Pub_GetSpecMan("PS2")
                        End If
                        'end 2025/1/21
                     End If
                     '主旨
                     strExc(4) = "大陸案" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "已核准，" & _
                                       "澳門案" & m_CPto044(1) & "-" & m_CPto044(2) & IIf(m_CPto044(3) & m_CPto044(4) <> "000", "-" & m_CPto044(3) & "-" & m_CPto044(4), "") & _
                                       "請向代理人索取提申文件 Our Ref: " & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "[INCOM. 1001]"
                     '內文
                     strExc(5) = "請向代理人報告並索取公證的委托書以利提申。"
                     If strExc(1) <> "" Then
                         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                   " values( '" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                                   ",'" & strExc(4) & "','" & strExc(5) & "','" & strExc(2) & ";" & strExc(3) & ";backup')"
                         cnnConnection.Execute strSql, intI
                     End If
               End If
            End If
            'end 2021/11/10
      End Select
   End If
   
   'Add by Morgan 2004/6/15收文日＞＝93/7/1核准案件掛三個月的領證期限
   stNP07 = "": stNP08 = "": stNP09 = ""
   m_bolTw601Chk = False 'Added by Morgan 2012/9/19
   If pa(9) = 台灣國家代號 And Val(Label2(3)) >= 930701 Then
      'Modify by Morgan 2004/7/5 加再審申請 107,聯合案改用申請案號第9碼判斷
      'Modify by Morgan 2004/9/9 聯合追加不再掛期限-郭雅娟
      'If InStr("101,102,103,104,105, 107,301,302,303,304,305,306,307", m_strCP10) > 0 Then
      'Modified by Morgan 2012/12/19 +衍生設計125,改請衍生設計308
      If InStr("101,102,103,107,125,301,302,303,306,307,308", m_strCP10) > 0 Then
      'end 2004/9/9
      
         'Modify by Morgan 2007/6/12 若有輸機關發文日時改用該日計算期限
         'stNP09 = Format(Val(Label2(3)) + 19110000)
         If txtDispDate.Visible = True Then
            stNP09 = DBDATE(txtDispDate)
         Else
            stNP09 = Format(Val(Label2(3)) + 19110000)
         End If
         'end 2007/6/12
         'Modify by Morgan 2010/12/24 申請號改碼數
         'If Mid(pa(11), 9, 1) <> "" Then
         'Modified by Morgan 2012/12/19 +衍生設計也有單獨的證書
         'If Mid(pa(11), 10, 1) <> "" Then
         If Mid(pa(11), 10, 1) <> "" And Mid(pa(11), 10, 1) <> "D" Then
            stNP07 = 加註聯合 '603
            '法定期限=收文日+30天
            stNP09 = CompDate(2, 30, stNP09)
            'Added by Morgan 2014/10/9
            'Memo by Lydia 2025/10/29 內專本所約定期限啟用日放在下方執行
            If strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               stNP08 = PUB_GetOurDeadline(stNP09)
            Else
            'end 2014/10/9
               '本所期限=法定-2天
               stNP08 = PUB_GetWorkDay1(CompDate(2, -2, stNP09), True)
            End If 'Added by Morgan 2014/10/9
         Else
            'Added by Morgan 2015/9/16
            '台灣案核准時,若多國案仍未發文,且未設定期限,則請預設自核准日起算1個月為該多國案之本所期限,並發MAIL告知工程師
            'Modified by Morgan 2015/12/1 改原來無所限或較晚時更新，若法限晚於核准日+7個月時一併清除
            'Modified by Morgan 2015/12/21 改呼叫共用函數
'            strExc(4) = CompDate(1, 7, Text5(0)) '最晚預定公告日=核准日+7個月(3個月領證+3個月延緩+1個月預估)
'            strExc(1) = PUB_GetWorkDay1(CompDate(1, 1, Text5(0)), True)
'            strExc(2) = PUB_GetRefCaseMapSQL(pa)
'            strExc(3) = "期限來源:" & Right("  " & pa(1), 3) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "(所限=核准日" & ChangeWStringToWDateString(DBDATE(Text5(0))) & "+1個月);"
'            strExc(0) = "SELECT CP09,CP01||'-'||CP02||decode(CP03||CP04,'000','','-'||CP03||'-'||CP04) RefNo,CP14,cp07" & _
'               " FROM CASEPROGRESS C1 WHERE (CP01,CP02,CP03,CP04) IN (" & strExc(2) & ") AND CP01<>'FCP'" & _
'               " AND CP10 IN (" & SameCaseProperty4Update & ") AND CP27||CP57 IS NULL and (cp06 is null or cp06>" & strExc(1) & ")" & _
'               " AND NOT EXISTS(SELECT * FROM patent WHERE PA01=C1.CP01 AND PA02=C1.CP02 AND PA03=C1.CP03 AND PA04=C1.CP04 AND PA46='Y')" & _
'               " AND NOT EXISTS(SELECT * FROM CASEPROGRESS C2 WHERE C2.CP01=C1.CP01 AND C2.CP02=C1.CP02 AND C2.CP03=C1.CP03 AND C2.CP04=C1.CP04 AND C2.CP10='106' AND C2.CP57 IS NULL)"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               With RsTemp
'               Do While Not .EOF
'                  If "" & RsTemp("cp07") > strExc(4) Then
'                     strSql = "update caseprogress set cp06=" & strExc(1) & ",cp07=null,CP64=SUBSTR(CP64,1,INSTR(CP64,'期限來源:')-1)||'" & strExc(3) & "'||SUBSTR(CP64,INSTR(CP64,';',instr(CP64,'期限來源:'))+1) where cp09='" & .Fields("cp09") & "'"
'                  Else
'                     strSql = "update caseprogress set cp06=" & strExc(1) & ",CP64='" & strExc(3) & "'||CP64 where cp09='" & .Fields("cp09") & "'"
'                  End If
'                  cnnConnection.Execute strSql, intI
'
'                  strExc(2) = .Fields("RefNo") & "之相對應" & cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "己核准,預設本所期限為" & TranslateKeyWord(incCNV_CHINESE_MINKO, strExc(1), "") & ",請儘速作業!!"
'                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                     " values( '" & strUserNum & "','" & .Fields("CP14") & "',to_char(sysdate,'yyyymmdd')" & _
'                     ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
'                  cnnConnection.Execute strSql, intI
'                  .MoveNext
'               Loop
'               End With
'            End If
            PUB_UpdCP06byTwPA20 pa, Text5(0)
            'end 2015/9/16
            
            stNP07 = 領證及繳年費 '601
            '法定期限=收文日+3個月
            'Modified by Morgan 2012/7/12 文到次日起算改(+1日+3月-1日)方式計算(Ex.4/30-->7/31)
            'stNP09 = CompDate(1, 3, stNP09)
            stNP09 = CompDate(2, 1, stNP09)
            stNP09 = CompDate(1, 3, stNP09)
            stNP09 = CompDate(2, -1, stNP09)
            If stNP09 >= 20130101 Then m_bolTw601Chk = True 'Added by Morgan 2012/9/19
            'end 2012/7/12
            
            'Added by Morgan 2014/10/9
            If strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               stNP08 = PUB_GetOurDeadline(stNP09)
            Else
            'end 2014/10/9
               '本所期限=法定-4天
               stNP08 = PUB_GetWorkDay1(CompDate(2, -4, stNP09), True)
            End If 'Added by Morgan 2014/10/9
         End If
         'Added by Lydia 2025/10/29
         stNP23 = ""
         If strSrvDate(1) >= 內專本所約定期限啟用日 Then
            stNP08 = PUB_GetPOurDeadline(stNP09, pa(9), stNP23, pa(1), stNP07)
         End If
         'end 2025/10/29
         lMax = GetNextProgressNo
         'Modified by Lydia 2025/10/29 +NP23
         strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP22,NP23) " & _
            "VALUES ('" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & stNP07 & "," & _
            stNP08 & "," & stNP09 & ",'" & stCP13 & "'," & CNULL(Text5(2)) & "," & _
            lMax & "," & CNULL(stNP23, True) & " )"
         StrlMAXbyNick = StrlMAXbyNick & lMax & ","
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
         '2010/12/20 ADD BY SONIA 陳玲玲要求同時更新核准來函期限
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP06=" & stNP08 & ",CP07=" & stNP09 & " WHERE CP09='" & NewReceiveNo & "'"
         cnnConnection.Execute strTxt(intStep), intI
         intStep = intStep + 1
         '2010/12/20 END
         m_strRetSheetNP07 = stNP07 'Add by Morgan 2005/11/16
         
         bolRegMail = True 'Added by Morgan 2014/4/17
         
         'Modified by Morgan 2012/8/15
         '發明核准且為一案兩請若新型案已發證則新增新型案下一程序放棄專利權(429)
         'Modified by Morgan 2014/7/29
         '若輸審查意見時已選擇放棄新型則此處不必再確認
         'If pa(8) = "1" Then
         'Modified by Morgan 2015/7/9 一案兩請是否放棄新型改放PA60
         'If pa(8) = "1" And pa(162) <> "Y" Then
         If pa(8) = "1" And pa(60) <> "Y" Then
         'end 2014/7/29
            'Modified by Morgan 2018/8/31
            '取消新型未閉卷條件,因有可能年費不辦閉卷(自己繳) Ex: P-115648, P-115649 --陳玲玲
            '取消有發證日條件改判斷已核准(可能領證不辦但還是要提醒) Ex:P-118076, P-117954 --韻丞
            strSql = "select cm01,cm02,cm03,cm04" & _
               " from (select cm01,cm02,cm03,cm04 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'" & _
               " union select cm05,cm06,cm07,cm08 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3') X,patent" & _
               " where pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa08='2' and pa16='1'" & _
               " and not exists(select * from caseprogress where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='429' and cp57 is null)"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               'Modified by Morgan 2017/1/20 改通知放棄專利權(1914)要先新增(cp30不必寫),放棄專利權(429)期限的相關收文號改放通知放棄專利權的收文號(當刪除1914CP時才會一併刪除429NP,否則會殘留)
               m_strDualAppNP22 = ""
               
               'Added by Morgan 2014/7/15
               '新增通知放棄專利權
               m_str1914CP09 = AutoNo("C", 6)
               strExc(2) = PUB_GetAKindSalesNo(RsTemp(0), RsTemp(1), RsTemp(2), RsTemp(3))
               strExc(3) = GetSalesArea(strExc(2))
               'modify by sonia 2019/1/4 雅娟通知台灣案不掛期限(因為定稿內容並無期限)
               'strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10" & _
                  ",cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp32,cp43 ) values ('" & RsTemp(0) & "'" & _
                  ",'" & RsTemp(1) & "','" & RsTemp(2) & "','" & RsTemp(3) & "'," & strSrvDate(1) & _
                  "," & stNP08 & "," & stNP09 & ",'" & m_str1914CP09 & "','1914','" & strExc(3) & "'" & _
                  ",'" & strExc(2) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'" & m_strDualAppNP22 & "','N','" & NewReceiveNo & "')"
               strSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                  ",cp12,cp13,cp14,cp20,cp26,cp27,cp30,cp32,cp43 ) values ('" & RsTemp(0) & "'" & _
                  ",'" & RsTemp(1) & "','" & RsTemp(2) & "','" & RsTemp(3) & "'," & strSrvDate(1) & _
                  ",'" & m_str1914CP09 & "','1914','" & strExc(3) & "'" & _
                  ",'" & strExc(2) & "','" & strUserNum & "','N','N'," & strSrvDate(1) & ",'" & m_strDualAppNP22 & "','N','" & NewReceiveNo & "')"
               cnnConnection.Execute strSql, intI
               '台灣案新增信函進度(判發人為游經理)
               'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
               'Modified by Morgan 2016/8/4 修正系統別傳錯抓不到判發人問題 Ex.P-103491
               'Modified by Morgan 2018/8/1
               'strExc(1) = PUB_GetLetterJudge(RsTemp(0), "1914", , , RsTemp(0), RsTemp(1), RsTemp(2), RsTemp(3))
               strExc(1) = PUB_GetLetterJudgeNew("1", RsTemp(0), "1914")
               PUB_AddLetterProgress m_str1914CP09, 0, True, strExc(1), True, pa(26), "1914", pa(75)
               'end 2014/7/15
            
'cancel by sonia  2019/1/4 雅娟通知台灣案不掛期限(因為定稿內容並無期限)
'               lMax = GetNextProgressNo
'               m_strDualAppNP07 = "429"
'               m_strDualAppNP22 = lMax
               m_strDualAppNo = RsTemp(0) & RsTemp(1) & RsTemp(2) & RsTemp(3)
'               strSql = "insert into nextprogress (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22) values (" & _
'                  " '" & m_str1914CP09 & "','" & RsTemp(0) & "','" & RsTemp(1) & "','" & RsTemp(2) & "','" & RsTemp(3) & "'" & _
'                  ",'" & m_strDualAppNP07 & "'," & stNP08 & "," & stNP09 & _
'                  ",'" & PUB_GetAKindSalesNo(RsTemp(0), RsTemp(1), RsTemp(2), RsTemp(3)) & "'," & m_strDualAppNP22 & ")"
'               cnnConnection.Execute strSql, intI
'end 2019/1/4
               
               'Modified by Morgan 2025/9/26 台灣及大陸案改統一在下面用相同提醒並EMail給工程師
               'strDualAppMsg = "請連同新型( " & m_strDualAppNo & " )卷一併交工程師確認是否為一案兩請!!" 'Added by Morgan 2013/1/11
               bolDualAppConfirmMail = True
            End If
         End If
      
         'Added by Morgan 2012/12/27
         'Modified by Morgan 2014/12/17 不必限制已送件
         If PUB_ChkPriDate(pa(11), strExc(3), False, False) = True Then
            '若本案被主張國內優先權時以E-MAIL告知智權同仁
            strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
            strExc(2) = strExc(1) & " 已核准!!但該案已被 " & strExc(3) & " 主張國內優先權，故應無須提出領證，請再與客戶作確認。"
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values( '" & strUserNum & "','" & stCP13 & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
            cnnConnection.Execute strSql, intI
            
            '若主張國內優先權案件未發文時以E-MAIL告知承辦工程師
            strExc(0) = "select cp14,cp09 from patent,caseprogress where " & ChgPatent(Replace(strExc(3), "-", "")) & _
               " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10 in ('101','102','103') and cp27 is null and cp57 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = "期限來源:" & pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4) & "(國內優先權案領證期限);"
               strSql = "update caseprogress set cp06=" & stNP08 & ",cp07=" & stNP09 & ",cp64='" & strExc(1) & "'||cp64 where cp09='" & RsTemp("cp09") & "' and (cp07 is null or cp07>" & stNP09 & ")"
               cnnConnection.Execute strSql, intI
               If Not IsNull(RsTemp("cp14")) Then
                  strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
                  strExc(2) = strExc(3) & " 所主張的國內優先權案 " & strExc(1) & " 已核准!!"
                  If intI = 1 Then strExc(2) = strExc(2) & "本案期限已設定為該核准案的領證期限(所限:" & ChangeWStringToTDateString(stNP08) & ",法限:" & ChangeWStringToTDateString(stNP09) & ")。"
                  'Modified by Morgan 2021/9/22 應該要寄給承辦人
                  'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                     " values( '" & strUserNum & "','" & stCP13 & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
                  strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                     " values( '" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd')" & _
                     ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','如旨')"
                  'end 2021/9/22
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
         'end 2012/12/27
         
      End If
   End If
   
   'MODIFY BY SONIA 90.11.15向香港提新案期限只在定稿上通知,不管制期限

   '8
   'Modify By Cheng 2006/03/18
   '變更事項要更新基本檔或進度檔的動作, 移至變更發文時做
   strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 99
            If IsNull(.Fields(i - 1)) Then
               strCe(i) = ""
            Else
               strCe(i) = .Fields(i - 1)
            End If
         Next
      End With
      strExc(1) = ""
      strExc(2) = ""
      strExc(3) = ""

      '申請日 10
      If strCe(2) <> "" Then
         strExc(1) = strExc(1) & "申請日 : " & strCe(2) & ","
         strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
         strExc(3) = strExc(3) & "CE03='1',"
      End If
      
      '申請人 26-30
      bolChk = False
      For i = 4 To 8
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strExc(1) = strExc(1) & "申請人 : "
         For i = 4 To 8
            If strCe(i) <> "" Then
               strExc(1) = strExc(1) & strCe(i) & ","
               'edit by nickc 2007/02/02 不用 dll 了
               'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
               If ClsPDGetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                  strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmp(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmp(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmp(3))) & ","
               End If
            End If
            'Modify By Cheng 2003/05/13
            'strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(strCe(i)) & ","
            strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
         Next
         strExc(3) = strExc(3) & "CE09='1',"
      Else
         '申請地址 31-45
         bolChk = False
         For i = 23 To 37
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strExc(1) = strExc(1) & "申請地址 : "
            For i = 23 To 37
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
               End If
               strExc(2) = strExc(2) & "PA" & i + 8 & "=" & CNULL(strCe(i)) & ","
            Next
            strExc(3) = strExc(3) & "CE38='1',"
         End If
      End If
      
      '專利商標種類代號 08
      If strCe(39) <> "" Then
         strExc(1) = strExc(1) & "專利商標種類代號 : " & strCe(39) & ","
         strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
         strExc(3) = strExc(3) & "CE40='1',"
      End If

      '案件名稱 05-07
      bolChk = False
      For i = 41 To 43
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strExc(1) = strExc(1) & "案件名稱 : "
         For i = 41 To 43
            If strCe(i) <> "" Then
               strExc(1) = strExc(1) & strCe(i) & ","
            End If
            strExc(2) = strExc(2) & "PA" & i - 36 & "=" & CNULL(strCe(i)) & ","
         Next
         strExc(3) = strExc(3) & "CE44='1',"
      End If

      '代表人 79-84
      bolChk = False
      For i = 10 To 15
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 68 To 91
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If

      If bolChk Then
         strExc(1) = strExc(1) & "代表人 : "
         For i = 10 To 15
            If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
         Next
         For i = 68 To 91
            If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
         Next
         strExc(3) = strExc(3) & "CE16='1',"
      End If
      
      '代表人中譯文
      If Not bolChk Then
         bolChk = False
         For i = 63 To 64
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If Not bolChk Then
            For i = 92 To 99
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
         End If
         If bolChk Then
            strExc(1) = strExc(1) & "代表人中譯文 : "
            strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
               "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
               "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
               "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
            For i = 63 To 64
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            Next
            For i = 92 To 99
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            Next
            strExc(3) = strExc(3) & "CE65='1',"
         End If
      End If
      If strExc(1) <> "" Then
         For i = 2 To 3
            If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
         Next
         intStep = intStep + 1
         strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(3) & " WHERE CE01='" & strReceiveNo & "'"
         'Add By Cheng 2002/11/06
         cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
      End If
   End If

   '2003/03/04讓與,專利權讓與改在發文時更新
   '92.4.23 cancel by sonia授權,設定質權取消終止日管制
   '92.3.15 CANCEL BY SONIA 此時不可產生年費期限, 於領證時產生
   
'Remove by Morgan 2009/11/26 已改至發證輸入
'   '910705 Sieg 405
'   strTxt(intStep) = "UPDATE PATENT SET PA14=" & CNULL(TransDate(Text5(9), 2)) & ",PA15=" & CNULL(ChgSQL(Text6)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
'   'Add By Cheng 2002/11/06
'   cnnConnection.Execute strTxt(intStep)
'   intStep = intStep + 1
   
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'   '2006/2/7 MODIFY BY SONIA
'   'If pa(8) = "1" And pa(9) = "020" Then
'   If pa(8) = "1" And pa(9) <> "000" Then
'      strTxt(intStep) = "UPDATE PATENT SET PA12=" & CNULL(TransDate(Text7, 2)) & ",PA13=" & CNULL(Text8) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
'      'Add By Cheng 2002/11/06
'      cnnConnection.Execute strTxt(intStep)
'      intStep = intStep + 1
'   End If
'
'   '2006/2/7 ADD BY SONIA 香港公告
'   'Modify by Morgan 2006/10/14 加澳門
'   'If pA(9) = "013" And Text5(6) = "6" Then
'   If (pa(9) = "013" Or pa(9) = "044") And Text5(6) = "6" Then
'      strTxt(intStep) = "UPDATE PATENT SET PA14=" & CNULL(TransDate(Text5(13), 2)) & ",PA15=" & CNULL(ChgSQL(Text11)) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
'      cnnConnection.Execute strTxt(intStep)
'      intStep = intStep + 1
'   End If
'   '2006/2/7 END
'
'   'add by nickc 2005/06/17 大陸香港關聯案
'   'edit by nickc 2006/08/18 公布才作
'   'If pA(9) = "020" Then
'   'Modify by Morgan 2007/10/25 應判斷來函的案件性質
'   'If pa(9) = "020" And cp(10) = "1207" Then
'   If pa(9) = "020" And stCP10 = "1207" Then
'   'end 2007/10/25
'   Dim tmpCp06 As String
'   Dim tmpCp07 As String
'   '檢查有無香港
'      If m_HaveHK = True And Text7.Text <> "" Then
'         tmpCp07 = PUB_GetWorkDay1(CompDate(1, 6, Text7.Text), True)
'         tmpCp06 = PUB_GetWorkDay1(CompDate(2, -5, CompDate(1, -1, tmpCp07)), True)
'         '檢查有無收香港的 110
'         If m_HaveHKInCP <> "" Then
'            '更新期限，上發 mail tag
'            strSQL = "Update CaseProgress Set CP06=" & tmpCp06 & ",CP07=" & tmpCp07 & " Where CP09='" & m_HaveHKInCP & "' "
'            cnnConnection.Execute strSQL
'            '更新齊備日
'            strSQL = "update engineerprogress set ep06=" & ServerDate & " where ep02='" & m_HaveHKInCP & "' "
'            cnnConnection.Execute strSQL
'            m_SendHKMail = True
'         End If
'      End If
'   End If
'
'en 2009/11/27

   'Add by Morgan 2005/5/20
   '非台灣 宣告無效答辯 更新結餘
   'modify by sonia 2025/3/31 再加案件僅變更401、讓與701及708，於核准時詢問是否計算結餘
   If cp(10) = "804" Or cp(10) = "401" Or cp(10) = "701" Or cp(10) = "708" Then
      Pub_UpdateEndModCash cp(1), cp(2), cp(3), cp(4)
   End If
   
   'Add by Morgan 2007/5/4
   If bolCancelClose = True Then
      strSql = "UPDATE PATENT SET PA57=NULL,PA58=NULL,PA59=NULL" & _
         " WHERE PA01 = '" & pa(1) & "' AND PA02 = '" & pa(2) & "'" & _
         " AND PA03 = '" & pa(3) & "' AND PA04 = '" & pa(4) & "' "
      cnnConnection.Execute strSql
   End If
   'end 2007/5/4
         
   'Add by Morgan 2009/7/16 大陸有領證或陳述意見期限時檢查是否有分割案期限需更新
   st307Msg = ""
   If IsEmptyText(Text5(10)) = False And pa(9) = 大陸國家代號 Then
      strSql = "select cp09 from divisioncase,caseprogress" & _
         " where dc05='" & cp(1) & "' and dc06='" & cp(2) & "'" & _
         " and dc07='" & cp(3) & "' and dc08='" & cp(4) & "'" & _
         " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp57||cp27 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '會有多個分割案
         Do While Not RsTemp.EOF
            strExc(1) = PUB_Update307Ref(RsTemp(0))
            If strExc(1) <> "" Then
               st307Msg = st307Msg & strExc(1) & vbCrLf
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   
   m_bolTw307Chk = False 'Added by Morgan 2012/9/19
   'Added by Morgan 2012/8/14 102新法
   '台灣母案初審核准必須更新分割案期限
   If pa(9) = "000" And cp(10) = "101" Then
      'Added by Morgan 2012/12/21 改判斷 20121202 以後核准的
      If DBDATE(Text5(0)) >= 20121202 Then
         m_bolTw307Chk = True
         
         strSql = "select cp09,cp27 from divisioncase,caseprogress" & _
            " where dc05='" & cp(1) & "' and dc06='" & cp(2) & "'" & _
            " and dc07='" & cp(3) & "' and dc08='" & cp(4) & "'" & _
            " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               If IsNull(RsTemp("cp27")) Then
                  strExc(1) = PUB_Update307RefTw(RsTemp(0))
                  If strExc(1) <> "" Then
                     st307Msg = st307Msg & strExc(1) & vbCrLf
                  End If
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
   End If
   'end 2012/8/14
   
   'Added by Lydia 2023/06/15 寰華案:「414恢復權利-主張優先權106」更新實審期限
   If bolChk414for106 = True Then
      If Not ClsPDSavePriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
           GoTo ErrorHandler
      End If
      '更新公開和實審期限
      strExc(5) = PUB_GetFirstPriDate(cp())
      strExc(9) = ""
      If strExc(5) <> "" And strExc(5) <> strFirstPriDate Then
         '參考一般來函輸入frm04010504_3:「視為未主張」1918=>公開或實審期限的相關總收文號用申請程序的收文號
         strSql = "select cp09 from caseprogress WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and instr('" & NewCasePtyList & "',cp10)>0 and cp159=0 order by cp05 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            strExc(9) = "" & RsTemp(0)
         End If
         '模組沒有寫備註
         strSql = "Update CaseProgress Set CP64=sqldatet(to_char(sysdate,'yyyymmdd'))||'更新期限：原所限'||sqldatet(CP06)||'，原法限'||sqldatet(CP07)||';'||CP64 where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' and cp10='416' and cp158=0 "
         cnnConnection.Execute strSql
         strSql = "Update NextProgress Set NP15=sqldatet(to_char(sysdate,'yyyymmdd'))||'更新期限：原所限'||sqldatet(NP08)||'，原法限'||sqldatet(NP09)||';'||NP15 where np02='" & pa(1) & "' AND np03='" & pa(2) & "' AND np04='" & pa(3) & "' AND np05='" & pa(4) & "' and np07='416' and np06 is null "
         cnnConnection.Execute strSql

         PUB_UpdCfpDate2 pa(1), pa(2), pa(3), pa(4), strExc(5), strExc(9)
         '請彈跳提醒視窗：已更新實體審查期限為: XXX/XX/XX=>移到存檔完成
         strFirstPriDate = "Y" & strFirstPriDate
      End If
   End If
   'end 2023/06/15
   
   'Add by Morgan 2009/10/2 大陸,香港檢索報告需可輸入帳單(原來在提申輸)
   '若有輸入代理人D/N No, 帳單日期 及 帳單金額, 則新增國外帳單資料
   strBillNo = "": stErrMsg = ""
   If Me.Text12.Text <> "" And Me.Text13.Text <> "" And Me.Text14.Text <> "" Then
      If PUB_AddNewFBillData(cp(9), Me.Text12.Text, Me.Text13.Text, Me.Text14.Text, strBillNo, Combo3.Text) = False Then
         stErrMsg = "新增國外帳單資料作業失敗!!!"
         GoTo ErrorHandler
      End If
   End If
   'end 2009/10/2
         
   'Add by Morgan 2010/5/17
   '台灣案加速審查通知
   'Modified by Morgan 2014/1/9 排除107復審
   If pa(9) = "020" And Text5(1).Text = "1" And Text5(1).Text <> Text5(1).Tag And cp(10) <> "107" Then
      strExc(0) = "select na28,na29 from nation where na01='" & pa(9) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '發明或新型有訂實審期限的國家
         If pa(8) = "1" Or (pa(8) = "2" And Not IsNull(RsTemp("na28")) And RsTemp("na29") > 0) Then
            '台灣案已收文通知實審日且相關收文號尚無結果(不必管是否曾收文加速審查)
            '承辦人相同
            'Modify by Morgan 2010/5/25 承辦人離職不用
            'Modified by Morgan 2022/7/21 排除台灣案已收到審查意見通知函者--陳玲玲
            strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo" & _
               " from casemap,patent,caseprogress a" & _
               " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "'" & _
               " and cm03='" & cp(3) & "' and cm04='" & cp(4) & "' and cm05='P'" & _
               " and pa01(+)=cm05 and pa02(+)=cm06 and pa03(+)=cm07 and pa04(+)=cm08" & _
               " and (pa16 is null or pa16='2') and pa09='000' and pa08='1'" & _
               " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='1204'" & _
               " and exists(select * from caseprogress b,staff where b.cp09=a.cp43 and b.cp24 is null" & _
               " and b.cp14='" & cp(14) & "' and st01(+)=b.cp14 and st04='1')" & _
               " and not exists(select * from caseprogress b where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and cp10='1202')"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
               strExc(2) = strExc(1) & " 已核准，台灣發明案 " & RsTemp(0) & " 符合提出加速審查之條件.."
               strExc(3) = "台灣發明案 " & RsTemp(0) & " 仍在審查中，惟其相對應 " & strExc(1) & " 已核准" & _
                  "，故台灣發明案符合提出加速審查之條件，若欲辦理，請洽承辦工程師 '||st02||'" & _
                  "('||st01||') 研討內容及評估費用。"
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                  " select '" & strUserNum & "','" & stCP13 & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & strExc(2) & "','" & strExc(3) & "',st01" & _
                  " from staff where st01='" & cp(14) & "'"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   End If
   'end 2010/5/17
   
   'Added by Morgan 2017/3/8
   '核發申請優先權證明書或申請優先權存取碼時發E-MAIL通知主張本案的大陸案(通知98012)或日本案的程序(通知承辦人)--陳玲玲
   '核發非臺灣案的申請優先權存取碼時所有主張案都要通知Ex.P-114082,CFP-29151--郭雅娟
   'Modified by Morgan 2017/6/30 +補優先權證明232未發文的也要通知--陳玲玲 Ex.P117318(台),P117608(陸)
   'Modified by Morgan 2017/12/21 台灣案也改都要通知--玲玲 Ex:P-118984
   'Modified by Morgan 2021/10/29 改不限制 106,232是否已發文都通知(CFP案會先發文掛回覆委任代理人) Ex:P-127982,CFP-32636 --有跟郭確認
   'Modified by Morgan 2021/11/17 修正不限制未發文有補優先權證明232會重複通知問題 Ex:P-128322
   If cp(10) = "405" Or cp(10) = "436" Then
      strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,a.cp05,nvl(b.cp14,a.cp14) cp14,pa09,a.cp09,pa01" & _
         " from pridate,patent,caseprogress a,caseprogress b where pd06='" & pa(11) & "' and pa01(+)=pd01 and pa02(+)=pd02 and pa03(+)=pd03 and pa04(+)=pd04" & _
         IIf(pa(9) <> "000" And cp(10) = "436", "", " and pa09 in ('020','011','000')") & _
         " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp10='106' and a.cp57 is null" & _
         " and b.cp01(+)=pa01 and b.cp02(+)=pa02 and b.cp03(+)=pa03 and b.cp04(+)=pa04 and b.cp10(+)='232' and b.cp57(+) is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            
            '大陸案都要通知(設計案無存取碼但大陸可用電子檔所以不論是否有存取碼都要通知)
            If RsTemp("pa09") = "020" Then
               'Added by Morgan 2025/1/21
               If strSrvDate(1) >= P業務區劃分啟用日 Then
                  strExc(1) = PUB_GetPHandler(RsTemp("CaseNo"))
               Else
               'end 2025/1/21
               
                  strExc(1) = Pub_GetSpecMan("PS2")
                  
               End If 'Added by Morgan 2025/1/21
               
            'Added by Morgan 2017/12/21 台灣案也改都要通知--玲玲
            ElseIf RsTemp("pa09") = "000" Then
               strExc(1) = "" & RsTemp("cp14")
            '非大陸案有存取碼才通知
            ElseIf Trim(txtPriNo) <> "" Then
               strExc(1) = "" & RsTemp("cp14")
               'Added by Morgan 2021/10/14 CFP改抓管制人
               If RsTemp("pa01") = "CFP" Then
                  strExc(1) = PUB_GetCFPHandler(RsTemp("CaseNo"))
               End If
               'end 2021/10/14
            Else
               strExc(1) = ""
            End If
            If strExc(1) <> "" Then
               strExc(2) = cp(1) & "-" & cp(2) & IIf(cp(3) & cp(4) = "000", "", "-" & cp(3) & "-" & cp(4))
               'Modified by Morgan 2018/5/24 一案可能會收多道主張優先權導致dup
               'strExc(3) = RsTemp("CaseNo") & " 有主張 " & strExc(2) & " 優先權，" & IIf(txtPriNo <> "", "已取得存取碼(" & txtPriNo & ")", "優先權證明書已核發") & "！"
               strExc(3) = RsTemp("CaseNo") & "(" & RsTemp("CP09") & ") 有主張 " & strExc(2) & " 優先權，" & IIf(txtPriNo <> "", "已取得存取碼(" & txtPriNo & ")", "優先權證明書已核發") & "！"
               strExc(4) = "主張案：" & RsTemp.Fields("CaseNo") & vbCrLf & _
                           "主張優先權收文日：" & ChangeTStringToTDateString(ChangeWStringToTString(RsTemp.Fields("cp05"))) & vbCrLf & vbCrLf & _
                           "優先權案：" & strExc(2) & vbCrLf & _
                           "優先權號：" & pa(11) & vbCrLf & _
                           "存取碼：" & txtPriNo & vbCrLf
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " values('" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(3)) & "','" & ChgSQL(strExc(4)) & "')"
               cnnConnection.Execute strSql, intI
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   'end 2017/3/8
   
   If m_USCaseNo <> "" Then PUB_SetUsIDS pa(1), pa(2), pa(3), pa(4), NewReceiveNo, Text5(0).Text, , , , True    'Added by Morgan 2020/12/18 美國IDS期限管制

   'Add by Morgan 2010/6/2
   '保密審查核准通知相關案可發文
   If cp(10) = "430" Then
      PUB_430OkInform pa 'Add by Morgan 2010/6/2
   End If
   
   'Added by Morgan 2014/1/14
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, NewReceiveNo, pa(1), pa(2), pa(3), pa(4), stCP10
   End If
   'end 2014/1/14
   
   'Add by Sindy 2016/10/5
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", NewReceiveNo, "")
      'Modified by Lydia 2023/05/18 +不開啟附件, , , False
      'Modified by Morgan 2024/9/25 大陸專利權補償審批決定也要2次確認
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm04010502_1", IIf(Pub_StrUserSt03 = "F22" Or bolCN445, NewReceiveNo, ""), m_bolReKeyInOK, , False
   End If
   '2016/10/5 END
   
   'Added by Morgan 2018/11/7 舉發,舉發答辯,核准通知設定為直寄--玲玲
   If (cp(10) = "803" Or cp(10) = "804") Then
      bolRegMail = True
   End If
   'end 2018/11/7
   
   'Added by Morgan 2014/4/14 電子化-新增信函進度檔
   If pa(9) = "000" Then
      'Modified by Morgan 2014/12/8 改都要抓判發人(舉發及舉發答辯的分析由工程師撰寫但發文後來函改通知客戶且待判發)
      'strExc(1) = ""
      'If Text5(4) <> "N" Then
         'Added by Morgan 2014/7/15
         '發明案核淮函及新型案通知放棄專利權函, 設定判發人為游經理
         If m_str1914CP09 <> "" Then
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), "1914", , , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1914")
         'Added by Morgan 2016/3/22
         '工程師承辦的來函不必判發(歷程判)
         ElseIf m_bolEngCase Then
            strExc(1) = ""
         'end 2016/3/22
         Else
         'end 2014/7/15
            'Modified by Morgan 2018/8/1
            'strExc(1) = PUB_GetLetterJudge(pa(1), stCP10, cp(10), , pa(1), pa(2), pa(3), pa(4))
            strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), stCP10, , cp(10))
         End If 'Added by Morgan 2014/7/15
      'End If
      'end 2014/12/8
      'Modified by Morgan 2014/7/22 +傳FC代理人(pa75)
      PUB_AddLetterProgress NewReceiveNo, 1 + Val(Text6), IIf(Text5(4) <> "N", True, False), strExc(1), bolRegMail, pa(26), stCP10, pa(75)
      
      strLetterJudge = strExc(1) 'Added by Morgan 2025/9/26
      
   'Added by Morgan 2016/6/7  非臺灣案電子化
   ElseIf 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
      If m_str1914CP09 <> "" Then
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), "1914", , pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), "1914", pa(9), , , m_bolFMP)
      ElseIf m_bolEngCase Then
         strExc(1) = ""
      Else
         'Modified by Morgan 2018/8/1
         'strExc(1) = PUB_GetLetterJudge(pa(1), stCP10, cp(10), pa(9), pa(1), pa(2), pa(3), pa(4))
         strExc(1) = PUB_GetLetterJudgeNew("1", pa(1), stCP10, pa(9), cp(10), , m_bolFMP)
      End If
      'Modified by Morgan 2016/7/6 核發(1008)可不必存ALTR
      PUB_AddLetterProgress NewReceiveNo, IIf(stCP10 = "1008", 1, 2) + Val(Text6), IIf(Text5(4) <> "N", True, False), strExc(1), bolRegMail, pa(26), stCP10, pa(75)
   'end 2016/6/7
   
      strLetterJudge = strExc(1) 'Added by Morgan 2025/9/26
   End If
   'end 2014/4/14
   
   'Added by Lydia 2023/02/02 寰華案Key核准(相關收文號是掛新案101,102,103,307,107)確定後，判斷是否已經請款
   'Modified by Lydia 2023/11/16
   'strExc(1) = ""
   strChk = ""
   'Modified by Lydia 2023/09/21 排除107復審申請
   If stCP10 = 核准 And pa(9) = 大陸國家代號 And m_bolFMP2 = True And InStr("101,102,103,307", cp(10)) > 0 Then
      'Modified by Lydia 2023/10/31 改成模組
'      strExc(0) = "select cp09,cp60,cp14 from caseprogress where cp09= (" & _
'                        "select max(cp09) mno from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 " & _
'                        "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' " & _
'                        "and cp05 = (select max(cp05) mdate from caseprogress, staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 " & _
'                        "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' and nvl(cp43,'N') <> '" & NewReceiveNo & "' )) "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strExc(9) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) '程序
'         If "" & RsTemp.Fields("CP60") = "" Then
'             '1.上一道工程師案件性質未有請款單號，則自動發Mail
'             '收件者: 工程師   副本收受者: 工程師之主管;程序管制人員(Key來函人員不是管制人員也列入收件者);backup
'             '主旨: 本案已核准，請工程師儘速處理請款，以利後續告准流程Our Ref: P-060000 [INCOM.1001]
'             strExc(2) = PUB_GetFCPEngSup(RsTemp.Fields("CP14"))
'             '主旨
'             strExc(4) = "本案已核准，請工程師儘速處理請款，以利後續告准流程Our Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & " [INCOM." & 核准 & "]"
'             strExc(6) = strExc(2) & ";" & strExc(9) & IIf(strExc(9) <> strUserNum, ";" & strUserNum, "") & ";backup"
'             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                    " values( '" & strUserNum & "','" & RsTemp.Fields("CP14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                     ",'" & strExc(4) & "','如旨','" & strExc(6) & "')"
'             cnnConnection.Execute strSql, intI
'             strExc(1) = "Y1"
'         Else
'             '2.上一道工程師案件性質已有請款單號 , 但卷宗區無REPDN(寄請款函) Or DNUPL(請款單上傳)(有一項就不發email), 則自動發Mail:
'             '收件者: 程序管制人員 (Key來函人員不是管制人員也列入收件者) 副本收受者: 程序管制人員主管; backup
'             '主旨: 本案已核准，請程序儘速處理請款，以利後續告准流程Our Ref: P-060000 [INCOM.1001]
'             strExc(0) = "SELECT CPP01, CPP02 FROM CASEPAPERPDF B " & _
'                               "WHERE CPP01 in (select cp09 from caseprogress where cp60='" & RsTemp.Fields("CP60") & "')  AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.REPDN.%' OR UPPER(CPP02) LIKE '%.DNUPL.%' ) "
'             intI = 1
'             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'             If intI = 0 Then
'                 strExc(2) = PUB_GetFCPProSup(strExc(9))
'                 '主旨
'                 strExc(4) = "本案已核准，請程序儘速處理請款，以利後續告准流程Our Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & " [INCOM." & 核准 & "]"
'                 'CC
'                 strExc(6) = strExc(2) & ";backup"
'                 strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                        " values( '" & strUserNum & "','" & strExc(9) & IIf(strExc(9) <> strUserNum, ";" & strUserNum, "") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                         ",'" & strExc(4) & "','如旨','" & strExc(6) & "')"
'                 cnnConnection.Execute strSql, intI
'                 strExc(1) = "Y2"
'             End If
'         End If
'         If strExc(1) <> "" Then '另外通知
'             Sleep 100
'             strExc(8) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4)) '承辦
'             '主旨: ◎【FMP(寰華)案核准通知】 Our Ref: P-129322 [INCOM.1001] ，請通知代理人！
'             strExc(4) = "【FMP(寰華)案核准通知】Our Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4)) & " [INCOM." & 核准 & "]，請通知代理人！"
'             '內文:
'             'To:ＸＸＸ（請抓智權管制人員）
'             '本案尚有請款程序未完成，待程序E出請款後，再報告核准，謝謝。
'             'To:ＸＸＸ（請抓程序管制人員）
'             '  請儘速完成請款程序，待E出請款後，通知承辦人員報告核准，謝謝。
'             strExc(0) = "To " & PUB_ReadUserData(strExc(8)) & ": " & vbCrLf & _
'                              "本案尚有請款程序未完成，待程序E出請款後，再報告核准，謝謝。" & vbCrLf & vbCrLf & _
'                              "To " & PUB_ReadUserData(strExc(9)) & ": " & vbCrLf & _
'                              "請儘速完成請款程序，待E出請款後，通知承辦人員報告核准，謝謝。"
'             'Modified by Lydia 2023/06/14 +CC: backup(mc09)
'             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                        " values( '" & strUserNum & "','" & strExc(8) & ";" & strExc(9) & IIf(strExc(9) <> strUserNum, ";" & strUserNum, "") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                         ",'" & strExc(4) & "','" & ChgSQL(strExc(0)) & "','backup' )"
'             cnnConnection.Execute strSql, intI
'         End If
'      End If
      bolDNUPL = PUB_ChkFCPtoDNUPL(pa(1), pa(2), pa(3), pa(4), stCP10, NewReceiveNo)
      If bolDNUPL = True Then strChk = "Y" 'Added by Lydia 2023/11/16
      'end 2023/10/31
   End If
   'end 2023/02/02
   
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail--核准-PPH
   'Modified by Lydia 2023/05/26 已閉卷不通知
   'Move by Lydia 2023/05/26 從commit上方移過來,
   Dim bolFMP2mail As Boolean  'Added by Lydia 2023/05/26
   'Modified by Lydia 2023/06/02 傳入相關收文號案件性質,在模組內判斷
   'If m_bolFMP = True And m_bolFMP2 = True And cp(10) = "431" And pa(57) = "" Then
   '   bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), stCP10, cp(14))  '傳入相關收文的承辦人
   'Modified by Lydia 2023/10/31 +未寄發 And bolDNUPL = False
   If m_bolFMP = True And m_bolFMP2 = True And pa(57) = "" And bolDNUPL = False Then
      bolFMP2mail = Pub_SetFMP2toCMail(pa(1), pa(2), pa(3), pa(4), stCP10, cp(14), NewReceiveNo, cp(10)) '傳入相關收文的承辦人
   End If
   'end 2023/05/17
   
   'Added by Morgan 2020/4/10
   'FMP有期限之案件EMAIL通知
   'Move by Lydia 2023/03/13 從「寰華案Key核准」上方移下來
   'Modified by Lydia 2023/03/13 判斷有「寰華案Key核准」不用再通知
   'If m_bolFMP = True Then
   'Modified by Lydia 2023/05/26 排除-寰華案無期限之官方來函，系統自動發Mail => And bolFMP2mail = False
   'Modified by Lydia 2023/11/16 strExc(1)=>strChk
   If m_bolFMP = True And strChk = "" And bolFMP2mail = False Then
      'Modified by Morgan 2020/9/15 +寰華案,改通知智權人員
      'If Left(Pub_StrUserSt03, 1) <> "F" Then
      '   PUB_FMPCaseInform NewReceiveNo
      'End If
      'Modified by Morgan 2023/5/25 FMP電子化所有來函應該都要EMail通知
      'PUB_FMPCaseInform NewReceiveNo, , True, Left(Pub_StrUserSt03, 1) = "F"
      PUB_FMPCaseInform NewReceiveNo, False, True, Left(Pub_StrUserSt03, 1) = "F", , bolChgRlt
      'end 2023/5/25
      'end 2020/9/15
   End If
   'end 2020/4/10
   
   'Added by Morgan 2024/6/20
   '准予延緩審查更申請/改請程序的催審期限=續行審查日(延緩審查日)+催審天數
   If cp(10) = "245" Then
      strExc(0) = "select np01,np22,cf05 from caseprogress,nextprogress,casefee" & _
         " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10 in (" & NewCasePtyList & ")" & _
         " and np01(+)=cp09 and np07='411' and np06 is null" & _
         " and cf01(+)=cp01 and cf02='" & pa(9) & "' and cf03(+)=cp10 and cf05>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = CompDate(2, RsTemp("cf05"), DBDATE(txt412))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np07='411' and np22=" & RsTemp("np22")
         cnnConnection.Execute strSql, intI
      End If
   End If
   'end 2024/6/20
   
   'Added by Morgan 2025/3/7 面詢未辦理，向官方辦理退費控管--玲玲
   If m_bolAddB908 Then
      strExc(9) = AutoNo("B", 6)
      strExc(1) = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05," & _
         "CP09,CP10,CP11,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP64) VALUES " & _
         "('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & _
         ",'" & strExc(9) & "','908','90','" & stCP12 & "','" & stCP13 & "'" & _
         ",'" & strExc(1) & "','N','N','N','" & NewReceiveNo & "','退請求面詢規費') "
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/3/7
   
   If stCP10 = 核准 And pa(9) = "000" Then PUB_ChkTW413 cp(9), True 'Added by Morgan 2025/3/13
   
   'Added by Morgan 2025/9/26
   '一案兩請
   If bolDualAppConfirmMail Then
      strDualAppMsg = "請連同新型( " & m_strDualAppNo & " )卷一併交工程師確認是否為一案兩請!!"
      
      '收件人:發明案工程師
      If m_bolFMP Then
         strExc(8) = PUB_GetFCPPromoterNo(cp(9), "1001", cp(14))
      Else
         strExc(8) = PUB_GetPPromoter(pa(1) & pa(2) & pa(3) & pa(4))
      End If
      '副本:判發主管及程序人員
      strExc(3) = PUB_GetPHandler(pa(1) & pa(2) & pa(3) & pa(4))
      If strLetterJudge <> "" Then
         strExc(3) = strLetterJudge & ";" & strExc(3)
      End If
      '主旨
      strExc(4) = pa(1) & pa(2) & pa(3) & pa(4) & " 與 " & m_strDualAppNo & "，請確認是否仍為一案兩請，若是請一併註明相對應的請求項，謝謝!"
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
               " values( '" & strUserNum & "','" & strExc(8) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & strExc(4) & "','如旨','" & strExc(3) & "')"
      cnnConnection.Execute strSql, intI
                         
   End If
   'end 2025/9/26
   
   cnnConnection.CommitTrans
   FormSave = True
   
   If strMsg <> "" Then MsgBox strMsg, vbInformation 'Add by Morgan 20104/23
   
   If st307Msg <> "" Then MsgBox st307Msg, vbInformation 'Add by Morgan 2009/7/16
   
   If strDualAppMsg <> "" Then MsgBox strDualAppMsg, vbInformation 'Added by Morgan 2013/1/11
   
   If strBillNo <> "" Then MsgBox "已新增帳單【 " & strBillNo & " 】。", vbInformation 'Add by Morgan 2009/10/2
   
'Removed by Morgan 2016/3/18 改都通知
'   'Modify By Sindy 2013/4/2 若承辦人是王協理且未發文則要發EMail通知
'   If pa(9) = 台灣國家代號 And (cp(10) = "803" Or cp(10) = "804") And str941CP14 = "71011" Then
'      Call PUB_SendMail(strUserNum, "71011", str941ReceiveNo, "分案通知")
'   End If
'   '2013/4/2 End
'end 2016/3/18
   
'Removed by Morgan 2017/3/8 併入上面大陸案的通知(設計案無存取碼但大陸可用電子檔所以不論是否有存取碼都要通知)--陳玲玲
'   'Add by Amy 2014/03/24 因應台日優先權證明文件,為P案申請國為台灣或大陸,種類為發明或新型,案件性質405 or 436需發mail通知日本案的程序人員
'    Dim strSubject As String, strContent As String
'    strSubject = "": strContent = ""
'    'Modify by Amy 2014/07/09 有輸優先權存取碼才判斷是否發mail
'    'If Text1 = "P" And (pa(9) = 台灣國家代號 Or pa(9) = 大陸國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
'    If Trim(txtPriNo) <> "" Then
'        strExc(0) = "Select cp14,cm05||'-'||cm06||'-'||cm07||'-'||cm08 as TWNo,cm01||'-'||cm02||'-'||cm03||'-'||cm04 as JPNo,cp05 From CaseMap,Patent,PriDate,CaseProgress Where cm05='" & Text1 & "' And cm06='" & Text2 & "' And cm07='" & Text3 & "' And cm08='" & Text4 & "' " & _
'                         "And cm10='0' And pa09='011' And pd07='000' And pd06='" & pa(11) & "' And cp10='106' And cm01=pa01(+) And cm02=pa02(+) And cm03=pa03(+) And cm04=pa04(+) " & _
'                         "And cm01=pd01(+) And cm02=pd02(+) and cm03=pd03(+)and cm04=pd04(+) And cm01=cp01(+) And cm02=cp02(+) And cm03=cp03(+) And cm04=cp04(+)"
'
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         strSubject = RsTemp.Fields("JPNo") & " 主張台灣優先權 " & RsTemp.Fields("TWNo") & "，已取得存取碼(" & txtPriNo & ")"
'         strContent = "台灣案號：" & RsTemp.Fields("TWNo") & vbCrLf & _
'                           "申請案號：" & pa(11) & vbCrLf & _
'                           "優先權存取碼：" & txtPriNo & vbCrLf & vbCrLf & _
'                           "日本案號：" & RsTemp.Fields("JPNo") & vbCrLf & _
'                           "主張優先權收文日：" & ChangeWStringToTString(RsTemp.Fields("cp05"))
'
'         Call PUB_SendMail(strUserNum, RsTemp.Fields("cp14"), "", strSubject, strContent)
'      End If
'    End If
'   'end 2014/03/24
'end 2017/3/8
   
   '**********  90.11.13    nickc
   If FormSave = True Then
      VarlMaxByNick = Split(StrlMAXbyNick, ",")
      For jjjbyNick = 0 To UBound(VarlMaxByNick)
         If Val(VarlMaxByNick(jjjbyNick)) <> 0 Then
            g_PrtForm001.PrintForm Trim(Val(VarlMaxByNick(jjjbyNick))), pa(1), pa(2), pa(3), pa(4) 'Memo by Lydia 2020/04/06  99/5/1起不再列印
         End If
      Next jjjbyNick
      
      'Add by Morgan 2005/3/16 若發明核准且為一案兩請則列印新型案自請撤回接洽單
      If m_strDualAppNP22 <> "" Then
         g_PrtForm001.PrintForm m_strDualAppNP22
      End If
      
'Removed by Morgan 2015/11/17 配合無紙化取消列印--郭
'      'Add By Sindy 2013/4/2 台灣舉發及答辯同時跑一張B類接洽記錄單
'      If pa(9) = 台灣國家代號 And (cp(10) = "803" Or cp(10) = "804") Then
'         g_PrtForm001.PrintCForm str941ReceiveNo
'      End If
'      '2013/4/2 End
'end 2015/11/17
      
      'Added by Lydia 2020/04/06 因應防疫在家上班作業，請將FMP案key來函產生的C類接洽記錄單回存到卷宗區
      '                                         比照FCP案C類接洽單同時列印並且上傳到卷宗區frm06010602_3: 1.爭議程序(8開頭) 2.結果為”改變原處分”3.存檔前-不需彈特殊備註
      'Remove by Lydia 2020/12/16 取消產生C類接洽記錄單
'      If m_bolFMP = True And NewReceiveNo <> "" Then
'          If Left(cp(10), 1) = "8" Or frm04010502_2.Text6 = "2" Then
'                strExc(1) = "": strExc(2) = ""
'                For intI = 0 To 4
'                    If pa(26 + intI) <> "" Then
'                        strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)), "1", bolTmp)
'                        If strExc(1) <> "" Then
'                           If bolTmp = True Then '個案備註
'                              m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
'                              strExc(2) = strExc(2) & strExc(1)
'                              Exit For
'                           ElseIf strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
'                              If m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0) Then
'                                   m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
'                              End If
'                              strExc(2) = strExc(2) & strExc(1) & "||" '判斷是否有重複備註 (一般核准的檢查)
'                           End If
'                        End If
'                    End If
'                Next intI
'                g_PrtForm001.PrintCFormNew NewReceiveNo, m_strMemo, , True
'          End If
'      End If
      'end 2020/12/16
      'end 2020/04/06
      
      'add by sonia 2016/9/6 台灣新型只要有主動修正收文都提醒,不管是否發文P-111945
      If pa(9) = 台灣國家代號 And (cp(10) = "102" Or cp(10) = "302") Then
         strExc(0) = "select count(*) from caseprogress " & _
            " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "'" & _
            " and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='203' and cp159=0 and cp05>=" & cp(5) + 19110000
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val(RsTemp.Fields(0)) > 0 Then MsgBox "本案曾提過主動修正,請確認處分書上是否有載明此事!!", vbExclamation
         End If
      End If
      'end 2016/9/6
      
   End If

   '****************************************
   Exit Function

ErrorHandler:
   If FormSave = False Then
      cnnConnection.RollbackTrans
      If stErrMsg <> "" Then MsgBox stErrMsg, vbCritical
   End If
End Function

Private Sub Form_Initialize()
   'add by nickc 2007/02/02
   ReDim cp(1 To TF_CP) As String
   ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
   Dim strTmp As String, bolChk As Boolean
   Dim bPaper As Boolean
   
   MoveFormToCenter Me
   SSTab1.Tab = 0
   intWhere = 國內
   
   With frm04010502_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      strSales = strExc(5)
      ReadPatent
   End With
   Label2(2) = strReceiveNo
   Label2(3) = frm04010502_1.Text5
   
   'Add By Sindy 2017/12/27
   m_strIR01 = frm04010502_2.m_strIR01
   m_strIR02 = frm04010502_2.m_strIR02
   m_strIR03 = frm04010502_2.m_strIR03
   m_strIR04 = frm04010502_2.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/27 END
   
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'   '2006/5/19 ADD BY SONIA 發明案件性質101,109,110才可選擇非台灣案通知書
'   Text5(6).Locked = True
'   Text5(6).TabStop = False
'
'end 2009/11/27
   'Add by Amy 2014/03/24 因應台日優先權證明文件P案申請國為台灣或大陸,種類為發明或新型,案件性質405 or 436 要輸優先權存取碼
   'Modified by Morgan 2022/1/21 此處改控制案件性質就好(申請優先權存取碼 436都要輸存取碼)
   'If Text1 = "P" And (pa(9) = 台灣國家代號 Or pa(9) = 大陸國家代號) And (pa(8) = "1" Or pa(8) = "2") And InStr("405,436", cp(10)) > 0 Then
   If Text1 = "P" And InStr("405,436", cp(10)) > 0 Then
   'end 2022/1/21
      LabPriNo.Visible = True: txtPriNo.Visible = True
   Else
      LabPriNo.Visible = False: txtPriNo.Visible = False
   End If
   'end 2014/03/24
   
   m_bolEngCase = False 'Added by Morgan 2016/3/18
   If pa(9) = 台灣國家代號 Then
      Text5(0) = TAIWANDATE(frm04010502_1.Text5)
      '2007/7/13 ADD BY SONIA
      If cp(10) = "928" Then Text5(4) = "N"
      '2007/7/13 END
      'Modified by Morgan 2015/6/25 +501訴願,505參加訴願
      'Modified by Morgan 2016/3/18
      '臺灣案舉發或舉發答辯的審定來函預設原工程師承辦,若離職則改游經理(73022),但假發文的則由程序承辦並修改定稿
      'If cp(10) = "803" Or cp(10) = "804" Or cp(10) = "501" Or cp(10) = "505" Then Text5(4) = "N" 'Add By Sindy 2013/4/2 舉發及舉發答辯不出通知函
      'modify by sonia 2018/11/9 +行政訴訟503,參加訴訟506
      If (cp(10) = "803" Or cp(10) = "804" Or cp(10) = "501" Or cp(10) = "505" Or cp(10) = "503" Or cp(10) = "506") And DBDATE(cp(27)) <> "19221111" Then
         m_bolEngCase = True
         Text5(4) = "N"
         Text5(4).Enabled = False
      End If
      'end 2016/3/18
   Else
      Text5(0) = ""
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      '2006/5/19 ADD BY SONIA 發明案件性質101,109,110才可選擇非台灣案通知書
'      'Modify by Morgan 2006/6/29 加分割307也要可以
'      'Modify y Morgan 2009/5/14 +澳門設計
'      If cp(10) = "101" Or cp(10) = "109" Or cp(10) = "110" Or (pa(8) = "1" And cp(10) = "307") Or (pa(9) = "044" And cp(10) = "103") Then
'         Text5(6).Locked = False
'         Text5(6).TabStop = True
'      End If
'      '2006/5/19 END
'
'end 2009/11/27

      'Add by Morgan 2009/10/6
      '檢索報告預設不出定稿
      '2010/12/1 modify by sonia 加423申請專利權評價報告P-092713
      'Modified by Morgan 2016/9/1 +426新穎性調查
      If cp(10) = "421" Or cp(10) = "423" Or cp(10) = "426" Then Text5(4) = "N"
      
      'Added by Morgan 2018/10/4
      '大陸案無效宣告,無效宣告答辯,行政訴訟,核准通知請改設定為"直寄",與台灣案處理模式相同--玲玲
      'modify by sonia 2018/11/9 +訴願501,參加訴訟506
      If (cp(10) = "803" Or cp(10) = "804" Or cp(10) = "503" Or cp(10) = "501" Or cp(10) = "506") And DBDATE(cp(27)) <> "19221111" And m_bolFMP <> True Then
         m_bolEngCase = True
         Text5(4) = "N"
         Text5(4).Enabled = False
      End If
      'end 2018/10/4
      
   End If
   
'Remove by Morgan 2009/10/29 沒有作用了(非台灣改為核准日起算,台灣則存檔時計算)
'   '92.1.20 ADD BY SONIA 92.10.7 加分割
'   If cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "104" Or cp(10) = "105" Or cp(10) = "307" Then
'      strTmp = CompDate(1, 2, TransDate(Label2(3), 2))
'      Text5(11) = TransDate(strTmp, 1)
'      Text5(10) = TransDate(CompDate(2, -5, strTmp), 1)
'   End If
'   '92.1.20 END
   
   strExc(0) = "SELECT CF15,CF10,CF24 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      
      '下一救濟程序名稱
      If Not IsNull(RsTemp.Fields(0)) Then
         strExc(0) = RsTemp.Fields(0)
         If pa(9) = 台灣國家代號 Then
            bolChk = False
         Else
            bolChk = True
         End If
         'edit by nickc 2007/02/02 不用 dll 了
         'objPublicData.GetCaseProperty pA(1), strExc(0), strTmp, bolChk
         ClsPDGetCaseProperty pa(1), strExc(0), strTmp, bolChk
      Else
         strTmp = ""
      End If
      strExc(0) = strTmp
      
      '主管機關
      If Not IsNull(RsTemp.Fields(1)) Then
         strExc(1) = RsTemp.Fields(1)
      Else
         strExc(1) = ""
      End If
      
      '主管機關文書
      If Not IsNull(RsTemp.Fields(2)) Then
         strExc(2) = RsTemp.Fields(2)
      Else
         strExc(2) = ""
      End If
   End If
   
   strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & pa(1) & "' AND CPM02='" & cp(10) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
   
      '來函期限
      If Not IsNull(RsTemp.Fields(0)) Then
         strExc(3) = RsTemp.Fields(0)
      Else
         strExc(3) = ""
      End If
      
      '來函期限天數或月數
      If Not IsNull(RsTemp.Fields(1)) Then
         strExc(4) = RsTemp.Fields(1) & "天"
      ElseIf Not IsNull(RsTemp.Fields(2)) Then
         strExc(4) = RsTemp.Fields(2) & "月"
      Else
         strExc(4) = ""
      End If
   End If
   
   'Add By Cheng 2002/06/21
   '若申請國家為台灣, 則帶出主管機關文書
   If pa(9) = 台灣國家代號 Then
      'Add by Morgan 2004/7/5
      If pa(8) = "2" And m_strCP10 = "107" Then
        Me.Text5(14).Text = "處分書"
      End If
      'Add by Morgan 2004/9/10
      If m_strCP10 = "412" Then
         lbl412.Visible = True
         txt412.Visible = True
         
      'Added by Morgan 2024/6/20 延緩審查
      'Modofied by Lydia 2025/02/12 臺灣：智慧局會單獨發准予延緩審查的函，請在輸入核准時讓user再次輸入准予延緩的日期以便系統比對日期是否相符(如延緩公告之操作模式)
      'ElseIf m_strCP10 = "245" Then
      '   lbl412.Caption = "續行審查日:"
      ElseIf m_strCP10 = "245" And pa(1) = "P" And pa(9) = "000" Then
         lbl412.Caption = "延緩審查日期:"
         lbl412.Left = 6400
      'end 2025/02/12
         lbl412.Visible = True
         txt412.Visible = True
         
      End If
      '2004/9/10 end
    
      strExc(0) = "SELECT CF24 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & Me.m_strCP10 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Me.Text5(14).Enabled = True
         Me.Text5(14).Text = "" & RsTemp(0).Value
      End If
      'Add By Cheng 2002/07/19
      '游標預設在機關文號
      SendKeys "{Tab}"
   End If
    '記錄機關文號
    m_CP08 = "" & Me.Text5(2).Text

   'Add by Morgan 2009/10/2
   '帳單幣別
   If cp(44) <> "" Then
      PUB_Add2Combo Combo3, cp(44)
   End If
  
  'Add by Morgan 2010/1/25
  '99/2/1起大陸取消維持費
  lbl606Year.Visible = False
  Text10.Visible = False
  lbl606Fee.Visible = False
  Text5(15).Visible = False
  'end 2010/1/25
  
  ClsPDGetCasePreAgent pa(), m_CP44, False 'Added by Morgan 2014/11/24
  
   'Added by Morgan 2015/6/23
   '1001,1002,1202,1209,1802,1807,1809,1810 E化提醒
   If PUB_GetEMailFlag(pa(1) & pa(2) & pa(3) & pa(4), , , bPaper) = True And bPaper = False Then
      MsgBox "E化案件，不印前案!!", vbExclamation
   End If
   'end 2015/6/23
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Add by Morgan 2010/5/18
   'Set frm04010502_3 = Nothing 'Removed by Morgan 2021/12/16 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

'************************************************
' 取回專利基本資料及收文資料
'
'************************************************
Private Sub ReadPatent()
 Dim Lbl As LABEL, i As Integer, strTmp As String, bolChk As Boolean, strTemp As String, strTemp1 As String
   For Each Lbl In Label2
      Lbl = ""
   Next
   If ClsPDReadPatentDatabase(pa(), intWhere) Then 'edit by nickc 2007/02/02 不用 dll 了 If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
      Label2(0) = pa(11): Label2(4) = pa(10)
      AddCboName Combo1, pa(5), pa(6), pa(7)
      Text5(5) = pa(17)
      
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      Text5(9) = pa(14)
'      Text6 = pa(15)
'
'      'Modify by Morgan 2007/11/22 顯示也要改西元否則檢查時會錯
'      'Text7 = pa(12)
'      Text7 = DBDATE(pa(12))
'      'end 2007/11/22
'      Text8 = pa(13)
'      '2006/2/8 ADD BY SONIA
'      If pa(9) = "013" Then
'         Text5(13) = pa(14)
'         Text11 = pa(15)
'      End If
'      '2006/2/8 END
'
'end 2009/11/27

   End If
   
   cp(9) = strReceiveNo
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.ReadCaseProgressDatabase(cp, intWhere) Then
   If ClsPDReadCaseProgressDatabase(cp, intWhere) Then
      If pa(9) = 台灣國家代號 Then
         bolChk = False
      Else
         bolChk = True
      End If
      'Add by Morgan 2008/5/13
      Text5(3) = cp(35)
      Text5(16) = cp(117)
      'end 2008/5/13
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseProperty(pA(1), cp(10), strTmp, bolChk) Then Label2(1) = strTmp
      If ClsPDGetCaseProperty(pa(1), cp(10), strTmp, bolChk) Then Label2(1) = strTmp
   End If
   
   Text1 = pa(1)
   Text2 = pa(2)
   Text3 = pa(3)
   Text4 = pa(4)
   
   strTemp = ""
   strTemp1 = ""
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetNationTaxEx(Val(pA(8)), pA(9), strTemp, strTemp1, , , False) = 0 Then
   'Modified by Morgan 2013/10/23
   'If ClsPDGetNationTaxEx(Val(pa(8)), pa(9), strTemp, strTemp1, , , False) = 0 Then
   If ClsPDGetNationTaxEx(Val(pa(8)), pa(9), strTemp, strTemp1, , , False, , pa(10), pa(21), pa(72)) = 0 Then
      If Val(strTemp) = 申請日 Then
         strExc(0) = "SELECT NP09 FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND np06 IS null AND NP07 IN (" & 年費 & "," & 維持費 & ")"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 And Not IsNull(RsTemp.Fields("np09")) Then
            Text5(12) = TransDate(RsTemp.Fields("np09"), 1)
         Else
            GetNextDate
         End If
      End If
   End If
   strYear = Text5(12).Text
   
   '92.10.7 ADD BY SONIA
   'Modify by Morgan 2010/8/11 百年蟲
   'If Text5(12).Text < GetTaiwanTodayDate Then Text5(12).Text = ""
   If Val(Text5(12)) < Val(strSrvDate(2)) Then Text5(12).Text = ""
   '92.10.7 END
    
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'    Me.Text5(7).Text = ""
'   If pa(9) = 大陸國家代號 Then
'      intI = 1
'      '抓本案是否有收文(不論是否發文)資料
'      'Modify By Cheng 2004/02/17
'      'strExc(0) = "SELECT COUNT(*) FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='" & 實體審查 & "' AND CP57 IS NULL"
'      strExc(0) = "SELECT * FROM CASEPROGRESS WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP10='" & 實體審查 & "' AND CP57 IS NULL"
'      'End
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'      'Modify By  Cheng 2002/12/29
'      '若無收文資料時, 大陸實審收文預設為"N"
'      'If intI = 1 Then Text5(7) = "N"
'      If intI = 0 Then Text5(7) = "N"
'   End If
'
'end 2009/11/27
   
   ' 90.07.05 modify by louis (機關文號設預設內容)
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If pa(9) = 台灣國家代號 Then
      Select Case cp(10)
         'Modified by Morgan 2012/12/19 +衍生設計
         'Modified by Morgan 2013/7/31--玲玲
         'Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, 分割, 衍生設計
         '   Text5(2).Text = "（" & strTmp & "）智專一（三）字第號"
         Case 發明申請, 分割
            Text5(2).Text = "（" & strTmp & "）智專一（五）字第號"
         Case 新型申請
            Text5(2).Text = "（" & strTmp & "）智專一（四）字第號"
         Case 設計申請, 追加申請, 聯合申請, 衍生設計
            Text5(2).Text = "（" & strTmp & "）智專一（三）字第號"
         'end  2013/7/31
         
         'Modify by Morgan 2011/6/24 申請優先權證明 不再有文號
         'Case 申請優先權證明, 變更, 讓與, 專利權讓與
         Case 申請優先權證明
         Case 變更, 讓與, 專利權讓與
         'end 2011/6/24
            Text5(2).Text = "（" & strTmp & "）智專一（一）字第號"
         'Modify by Morgan 2006/8/14
         Case "505" '參加訴願
            Text5(2).Text = "經訴字第號"
'2010/11/12 CANCEL BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
'         Case 訴願
'            Text5(2).Text = "經（" & strTmp & "）字第號"
'         Case 異議_專, 舉發
            Text5(2).Text = "（" & strTmp & "）智專三（三）字第號"
         '2007/7/16加928
         Case "928"
            '2007/7/31 MODIFY BY SONIA 依有無專利權預設不同
            'Text5(2).Text = "（" & strTmp & "）智專一（一）字第號"
            If pa(25) <> "" Then
               Text5(2).Text = "（" & strTmp & "）智專一（一）15106字第號"
            Else
               Text5(2).Text = "（" & strTmp & "）智專一（一）字第號"
            End If
            '2007/7/31 END
         
         'Added by Morgan 2012/6/29
         '尚未確定先空白
         Case "108"
         
         '2007/7/16 END
         Case Else
            Text5(2).Text = "（" & strTmp & "）智專一（二）字第號"
      End Select
   
      '記錄機關文號預設值
      Me.Text5(2).Tag = Me.Text5(2).Text
   End If
   
   'Added by Morgan 2014/1/14
   'Modified by Morgan 2014/4/17 +發文字
   If m_DocWord <> "" Then
     Text5(2) = m_DocWord & "字第" & m_DocNo & "號"
   ElseIf m_DocNo <> "" Then
      Text5(2) = Replace(Text5(2), "第號", "第" & m_DocNo & "號")
   End If
   'end 2014/1/14
   
   'Added by Morgan 2015/1/20
   '電子公文帶入審查委員及國際分類
   If m_DocNo <> "" Then
      If PUB_GetEDocData(m_DocNo, strExc(1)) Then
         Text5(3) = strExc(1)
      End If
   End If
   'end 2015/1/20
   
   ' 90.07.05 modify by louis (Disable專利權是否存在,是否更新基本檔准駁)
   EnableTextBox Text5(5), False
   EnableTextBox Text5(1), False
   'Add By Cheng 2002/07/23
   '顯示目前准駁
   Me.Text5(1).Text = "" & pa(16)
   Text5(1).Tag = Text5(1).Text 'Add by Morgan 2010/5/17
   '顯示專利權是否存在
   Me.Text5(5).Text = "" & pa(17)
   'MODIFY BY SONIA 90.10.21,90.11.4
   strTmp = ""
   Select Case cp(10)
      Case 發明申請, 新型申請, 設計申請, 追加申請, 聯合申請, 答辯
         'Modify By Chneg 2002/07/23
         'Text5(1).Text = "Y"
      Case 改請發明, 改請新型, 改請設計, 改請追加, 改請聯合, 改請獨立, 分割
         'Text5(1).Text = "Y"
      Case 異議_專, 舉發
         'Text5(1).Text = "Y"
      Case 異議答辯, 舉發答辯
         'Text5(1).Text = "Y"
         strTmp = "不"
   End Select
   
   '92.4.11 MODIFY BY SONIA 參考核駁函設定
   'If (cp(10) >= "101" And cp(10) <= "105") Or cp(10) = "107" Or cp(10) = "503" Or cp(10) = "504" Or (cp(10) >= "301" And cp(10) <= "307") Or cp(10) = "802" Or cp(10) = "804" Then
   'Modified by Morgan 2012/12/19 +衍生設計125,改請衍生設計308
   If (Val(cp(10)) >= 101 And Val(cp(10)) <= 105) Or Val(cp(10)) = 107 Or Val(cp(10)) = 125 Or Val(cp(10)) = 308 Or Val(cp(10)) = 503 Or Val(cp(10)) = 504 Or _
        (Val(cp(10)) >= 301 And Val(cp(10)) <= 307) Or (Val(cp(10)) >= 801 And Val(cp(10)) <= 805) Then
   '92.4.11 END
      Me.Text5(1).Text = "1"
   End If
   If cp(10) = "804" Then
      Me.Text5(5).Text = "Y"
   End If
   
   ' 90.07.05 modify by louis
   Label2(3) = frm04010502_1.Text5
   RefreshSpecData
   
   'Add by Morgan 2006/10/12 計算預定公告日
   If cp(10) = "412" And pa(9) = "000" Then
      m_strPA14 = PUB_GetPrePA14(cp)
   End If
   
   'Add by Morgan 2009/11/26
   'Modified by Morgan 2015/11/5
   'If Left(cp(12), 1) = "F" And pa(10) <> "000" Then
   '   m_bolFMP = True
   'Else
   '   m_bolFMP = False
   'End If
   stCP13 = PUB_GetAKindSalesNo(Me.Text1.Text, Me.Text2.Text, Me.Text3.Text, Me.Text4.Text)
   stCP12 = GetSalesArea(stCP13)
   'Modified by Lydia 2023/06/20 pa(10)=> pa(9)
   If Left(stCP12, 1) = "F" And pa(9) <> "000" Then
      m_bolFMP = True
   Else
      m_bolFMP = False
   End If
   'end 2015/11/5
   'Added by Lydia 2021/11/10 判斷寰華案
   m_bolFMP2 = False
   If m_bolFMP = True Then
      m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pa(1), pa(2), pa(3), pa(4))
   End If
   'end 2021/11/10
   
   'Added by Lydia 2023/06/15 寰華案:是否為「414恢復權利-主張優先權106」
   bolChk414for106 = False
   strFirstPriDate = ""
   Command2.Visible = False

   If m_bolFMP2 = True And frm04010502_2.Text6 = "1" And cp(10) = "414" And cp(43) <> "" Then
      strExc(0) = "select c1.cp09 as cp09_1,c1.cp10 as cp10_1,c2.cp09 as cp09_2,c2.cp10 as cp10_2 from caseprogress c1, caseprogress c2 where c1.cp09='" & cp(43) & "' and c1.cp43=c2.cp09(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If "" & RsTemp.Fields("cp10_1") = "106" Or "" & RsTemp.Fields("cp10_2") = "106" Then
            bolChk414for106 = True
            strFirstPriDate = PUB_GetFirstPriDate(cp())
            Command2.Visible = True
         End If
      End If
   End If
   If Not ClsPDReadPriority(pa, strPriority(1), strPriority(2), strPriority(3), strPriority(4), strPriority(5)) Then
      '抓目前優先權資料
   End If
   'end 2023/06/15
   
   'Add by Morgan 2007/6/12 檢查65002是否為最後的代理人
   lblDispDate.Visible = False
   txtDispDate.Visible = False
   txtDispDate = ""
   If pa(9) = "000" Then
      If PUB_IsLatestAgent(pa(1), pa(2), pa(3), pa(4)) = True Then
         lblDispDate.Visible = True
         txtDispDate.Visible = True
         txtDispDate.MaxLength = 7
      End If
   End If
   'end 2007/6/12
   
   bolCN445 = False 'Added by Morgan 2024/9/25
   
   'Added by Morgan 2023/2/23
   '台灣專利權延長
   If pa(9) = "000" And m_strCP10 = "415" Then
      lbl415Date.Visible = True
      txt415Date.Visible = True
   
   'Added by Morgan 2024/9/18
   '大陸專利權期限補償
   ElseIf pa(9) = "020" And m_strCP10 = "445" Then
      lbl412 = "補償天數:"
      lbl412.Visible = True
      txt412.Visible = True
      lbl415Date = "專利權期滿終止日:"
      lbl415Date.Visible = True
      txt415Date.Visible = True
      bolCN445 = True
   'end 2024/9/18
   
   Else
      lbl415Date.Visible = False
      txt415Date.Visible = False
   End If
   'end 2023/2/23
   
   'Added by Morgan 2024/9/18 從FormSave移來並增加445專利權期限補償控制
   If pa(23) = "1" And PUB_ChkIsRltPty(cp(1), cp(10), pa(9)) = True Then
      bolChgRlt = True
   Else
      bolChgRlt = False
      Label11 = "核准日:"
   End If
   'end 2024/9/18
End Sub

'計算下次繳費日
Private Sub GetNextDate()
   
   Dim strTemp1 As String, strTemp2 As String, strCaseProperty As String, strTemp As String
   Dim dobDateAdd As Double, strStartDate As String
   'edit by nickc 2007/02/02 不用 dll 了
   'If objPublicData.GetNationTax(Val(pA(8)), pA(9), strTemp, strTemp1, strTemp2, strCaseProperty) Then
   If ClsPDGetNationTax(Val(pa(8)), pa(9), strTemp, strTemp1, strTemp2, strCaseProperty) Then
      Dim strNext As String
      Dim strLawDate As String, strOurDate As String
      
      '起算日
      strStartDate = GetStartDate(strTemp, cp(), pa())
      'Add by Morgan 2006/3/23 大陸下次繳費日=起算日+大陸年度
      If pa(9) = "020" Then
         dobDateAdd = Val(Text9)
         If dobDateAdd > 0 Then
            '法定期限
            strLawDate = CompDate(0, dobDateAdd, strStartDate)
            
            '本所期限=法定期限-1個月又5天
            'Modify by Morgan 2014/12/9 FMP案所限改法限-10天--郭雅娟
            'strOurDate = CompDate(1, -1, strLawDate)
            'strOurDate = CompDate(2, -5, strOurDate)
            'Modified by Morgan 2019/7/30 非FMP也改10天--玲玲(Ex:P120303,應該是107100402請作單遺漏部分)
            'If m_bolFMP Then
            '   strOurDate = CompDate(2, -10, strLawDate)
            'Else
            '   strOurDate = CompDate(1, -1, strLawDate)
            '   strOurDate = CompDate(2, -5, strOurDate)
            'End If
            'end 2014/12/9
            strOurDate = CompDate(2, -10, strLawDate)
            strOurDate = PUB_GetWorkDay1(strOurDate, True) '改與年費發文規則一致
            'end 2019/7/30
            Text5(12) = ChangeWStringToTString(strOurDate)
         Else
            Text5(12) = ""
         End If
      Else
      '2006/3/23 end
         strNext = GetNext(strTemp1, 1)
         If Not IsEmptyText(strNext) Then
            dobDateAdd = CDbl(strNext)
            If strStartDate <> "" Then
               strStartDate = CompDate(0, (dobDateAdd - 1), strStartDate)
               '起算日減1天再減1個月又5天=起算日減1個月又6天
               strStartDate = CompDate(1, -1, strStartDate)
               strStartDate = CompDate(2, -6, strStartDate)
               Text5(12) = ChangeWStringToTString(strStartDate)
            End If
         End If
      End If
   End If
End Sub

Private Function GetLast(ByVal strValue As String) As String
Dim strTemp As String
Dim nIndex As Integer
Dim aryDate
   
   strTemp = Empty
   aryDate = Split(strValue, ",")
   For nIndex = 0 To UBound(aryDate)
      If Not IsEmptyText(aryDate(nIndex)) Then
         strTemp = aryDate(nIndex)
      End If
   Next nIndex
   GetLast = strTemp
End Function

Private Function GetNext(ByVal strValue As String, ByVal strLast As String) As String
Dim strTemp As String
Dim nIndex As Integer
Dim aryDate
   
   strTemp = Empty
   aryDate = Split(strValue, ",")
   For nIndex = 0 To UBound(aryDate)
      If Not IsEmptyText(aryDate(nIndex)) Then
         If Val(aryDate(nIndex)) > Val(strLast) Then
            strTemp = aryDate(nIndex)
            Exit For
         End If
      End If
   Next nIndex
   GetNext = strTemp
End Function

'Remove by Morgan 2010/1/25
'Private Sub Text10_GotFocus()
'  TextInverse Text10
'End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text13_GotFocus()
   TextInverse Text12
   CloseIme
End Sub

Private Sub Text13_Validate(Cancel As Boolean)
   If Me.Text13.Text <> "" Then
      If CheckIsDate(Text13.Text) = False Then
         Cancel = True
         TextInverse Text13
      ElseIf Val(Text13) > Val(strSrvDate(1)) Then
         MsgBox "帳單日期不可大於系統日！", vbExclamation
         Cancel = True
      End If
   End If
End Sub

Private Sub Text14_GotFocus()
   TextInverse Text14
   CloseIme
End Sub

Private Sub Text14_Validate(Cancel As Boolean)
    If Text14.Text <> "" Then
        If IsNumeric(Text14.Text) = False Then
            MsgBox "帳單金額輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            TextInverse Text14
        ElseIf Val(Text14) <> 0 Then
            If cp(44) = "" Then
                MsgBox "該筆進度資料無代理人，不可輸入帳單!!!", vbExclamation + vbOKOnly
                Cancel = True
                TextInverse Text14
            End If
        End If
    End If
End Sub

Private Sub Text19_Change()
   PUB_RefreshText Text19 'Added by Morgan 2022/3/22
End Sub

Private Sub Text19_GotFocus()
  TextInverse Text19
End Sub

Private Sub Text5_Change(Index As Integer)
   If Index = 0 Then Text9 = "" 'Added by Morgan 2016/3/17 核准日改變時年度要清除會殘留前次計算 Ex.P106308
End Sub

Private Sub Text5_GotFocus(Index As Integer)
Dim intPos As Integer
   
   'Modify By Cheng 2002/04/22
   '將游標設定在機關文號欄的"專"的後面
   If Index <> 2 Then
      InverseTextBox Text5(Index)
   Else
      With Me.Text5(Index)
         'Modify By Cheng 2003/03/11
         Select Case cp(10)
         Case 申請優先權證明 '游標停在"字"的前面
            If Len("" & .Text) > 0 Then
               intPos = InStr("" & .Text, "字")
               If intPos - 1 >= 0 Then
                   .SelStart = intPos - 1
                   .SelLength = 0
               End If
            End If
            
         'Add by Morgan 2006/8/14
         Case "505" '參加訴願--游標停在"第"的後面
            If Len("" & .Text) > 0 Then
               intPos = InStr("" & .Text, "第")
               If intPos > 0 Then
                   .SelStart = intPos
                   .SelLength = 0
               End If
            End If
         
         '2007/7/18 ADD by sonia 重新委任928,游標停在"字"的前面
         Case "928"
            If Len("" & .Text) > 0 Then
               '2007/7/31 MODIFY BY SONIA 依有無專利權預設不同位置
               'intPos = InStr("" & .Text, "字")
               If pa(25) <> "" Then
                  intPos = InStr("" & .Text, "號")
               Else
                  intPos = InStr("" & .Text, "一")
               End If
               '2007/7/31 END
               If intPos - 1 >= 0 Then
                   .SelStart = intPos - 1
                   .SelLength = 0
               End If
            End If
         '2007/7/18 END
         
         Case Else
            If Len("" & .Text) > 0 Then
               intPos = InStr("" & .Text, "專")
               If intPos > 0 Then
                   .SelStart = intPos
                   .SelLength = 0
               End If
            End If
         End Select
      End With
   End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
   Select Case Index
   
'Remove by Morgan 2010/1/25
'      Case 15:
'         KeyAscii = UpperCase(KeyAscii)
'         If KeyAscii <> 89 And KeyAscii <> 8 Then
'            KeyAscii = 0
'            Beep
'         End If

      Case 4, 7
         KeyAscii = UpperCase(KeyAscii)
         If KeyAscii <> 78 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      Case 6 '大陸/香港發明通知書
'         'Modify By Cheng 2002/10/11
'         'If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
'         '2006/2/27 MODIFY BY SONIA 加6香港公告
'         If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 53 And KeyAscii <> 54 And KeyAscii <> 8 Then
'            KeyAscii = 0
'            Beep
'         End If
'
'end 2009/11/27
   End Select
End Sub

Private Sub Text5_LostFocus(Index As Integer)
   Select Case Index
   'Add By Cheng 2003/01/07
   '本所分析的機關文號與機關文號欄要一致
      Case 2 '機關文號
         Me.Text19.Text = Replace(Me.Text19.Text, m_CP08, Me.Text5(2).Text)
         m_CP08 = "" & Me.Text5(2).Text
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      Case 6
'         'Add By Cheng 2002/12/02
'         '若有選大陸/香港發明通知書, 則不用帶出期限及費用
'         If Me.Text5(Index).Text <> "" Then
'             Me.Text5(8).Text = ""
'             Me.Text5(10).Text = ""
'             Me.Text5(11).Text = ""
'             Me.Text5(12).Text = ""
'             Me.Text5(15).Text = ""
'             Me.Text10 = ""
'         End If
'         '2006/2/7 ADD BY SONIA
'         If Me.Text5(Index).Text = "2" And pa(9) = "013" Then
'            m_LetterType = MsgBox("香港公布通知請選擇定稿別？" & vbCrLf & vbCrLf & "公佈通知書請選 YES, 政府憲報請選 NO", vbYesNo)
'         End If
'         '2006/2/7 END
'
'end 2009/11/27

   End Select
End Sub

Private Sub Text5_Validate(Index As Integer, Cancel As Boolean)
Dim lTmp As Long, i As Integer
Dim strTmp As String
   
   Select Case Index
      Case 0 '申請案核准日
         'Add By Cheng 2003/11/21若卷宗性質不為申請, 則直接跳開本段
         If pa(23) <> "1" Then Exit Sub
        
         If IsEmptyText(Text5(Index)) = False Then
            If ChkDate(Text5(Index)) Then
               'Modify by Morgan 2005/5/3
               'If Val(Text5(Index)) > Val(strSrvDate(2)) Then
               If Val(TransDate(Text5(Index), 2)) > Val(strSrvDate(1)) Then
                  MsgBox Replace(Label11, ":", "") & "不可大於系統日 !", vbCritical
                  Cancel = True
               End If
               
               'Add by Morgan 2005/5/3 大陸控制西元年
               If pa(9) <> "000" Then
                  If Not CheckIsDate(Text5(Index)) Then Cancel = True
               End If
               '2005/5/3 end
               
               If Cancel = False Then
               
                  If bolCN445 Then Exit Sub 'Added by Morgan 2024/9/26 專利權補償的核准不用計算年費期限
               
                  'Remove by Morgan 2008/5/28 改以共用函數設定
                  'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & ChangeCustomerL(cp(44)) & "' AND YF04='601' AND YF05=1"
                  'intI = 1
                  'lTmp = 0
                  'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  'If intI = 1 Then
                  '   lTmp = Val(RsTemp.Fields(0))
                  'Else
                  '   'Modify By Cheng 2003/01/14內專抓代理人Y00000001
                  '   'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000000' AND YF04='601' AND YF05=1"
                  '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='Y00000001' AND YF04='601' AND YF05=1"
                  '   intI = 1
                  '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
                  '   If intI = 1 Then lTmp = Val(RsTemp.Fields(0))
                  'End If
                  'end 2008/5/28
                  
                  'lTmp = ""    92.7.15 cancel by sonia
                  
                  '92.11.19 MODIFY BY SONIA
                  If pa(9) = "020" And Text9 = "" Then '大陸年度=""
                     '92.12.1 MODIFY BY SONIA
                     'i = Int(DateDiff("d", ChangeWStringToWDateString(TransDate(pa(10), 2)), ChangeWStringToWDateString(TransDate(Text5(0), 2))) / 365) + 1
                     '2007/11/9 MODIFY BY SONIA 因為會有潤年問題(P-082880)故改判斷年份及月日
                     'i = Int(DateDiff("d", ChangeWStringToWDateString(TransDate(pa(10), 2)), ChangeWStringToWDateString(TransDate(CompDate(1, 5, Text5(0)), 2))) / 365) + 1
                     'Memo by Morgan 2009/10/20 核准日+5個月為預定公告日--敏惠
                     'Modified by Morgan 2014/12/24 改呼叫公用函數
                     'i = Int((TransDate(CompDate(1, 5, Text5(0)), 2) - TransDate(pa(10), 2)) / 10000) + 1
                     ''2007/11/9 END
                     ''92.12.1 END
                     'Text9 = i
                     Text9 = PUB_GetChina605StartYear(Text5(0), pa(10))
                     'end 2014/12/24
                     
'Remove by Morgan 2010/1/23
'                     If pa(8) = "1" Then
'                        Text10 = 3
'                        'Remove by Morgan 2008/5/28 改以共用函數設定
'                        'strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='" & ChangeCustomerL(cp(44)) & "' AND YF04='605' AND YF05=" & i
'                        'intI = 1
'                        'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                        'If intI = 1 Then
'                        '   lTmp = lTmp + Val(RsTemp.Fields(0))
'                        'Else
'                        '   strExc(0) = "SELECT NVL(YF06,0)+NVL(YF07,0) FROM PATENTYEARFEE WHERE YF01='" & pa(9) & "' AND YF02='" & pa(8) & "' AND YF03='000000000' AND YF04='605' AND YF05=" & i
'                        '   intI = 1
'                        '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
'                        '   If intI = 1 Then lTmp = lTmp + Val(RsTemp.Fields(0))
'                        'End If
'                     End If
                     
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'                     'Modify by Morgan 2006/5/29 加控制非台灣案通知書空白才設定下次繳費日
'                     If Text5(6) = "" Then
                        GetNextDate 'Add by Morgan 2006/3/23
'                     Else
'                        Text5(12) = ""
'                     End If
'end 2009/11/27
                  End If
                  
                  'Modify by Morgan 2008/5/28 改以共用函數設定
                  'If Text5(8) = "" Then Text5(8) = Format(lTmp)
                  If Text5(8) = "" Then Text5(8) = GetFee
                  'End 2008/5/28
                  
                  '92.8.6 CANCEL BY SONIA 因為 P-067102案此處會與 CHECKYEAR 不符
                  'If ChangeTStringToWString(Text5(12)) < CompDate(1, 5, Text5(0)) Then
                  '   Text5(12) = ""
                  'End If
                  '92.8.6 END
                  
                  '92.5.8 ADD BY SONIA
                  If pa(9) <> 台灣國家代號 And (cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Or cp(10) = "104" Or cp(10) = "105" Or cp(10) = "307") Then
                     '領證期限=申請核准日+2個月
                     strTmp = CompDate(1, 2, TransDate(Text5(0), 2))
                     Text5(11) = TransDate(strTmp, 1)
                     'Added by Lydia 2025/10/29
                     If m_bolFMP = False And strSrvDate(1) >= 內專本所約定期限啟用日 Then
                        Text5(10) = TransDate(PUB_GetPOurDeadline(strTmp, pa(9)), 1)
                     Else
                     'end 2025/10/29
                        'Add by Morgan 2009/12/2 FMP案 所限=法限-10天
                        If m_bolFMP Then
                           Text5(10) = TransDate(PUB_GetWorkDay1(CompDate(2, -10, strTmp), True), 1)
                        Else
                        'end 2009/12/2
                           Text5(10) = TransDate(PUB_GetWorkDay1(CompDate(2, -5, strTmp), True), 1)
                        End If
                     End If 'Added by Lydia 2025/10/29
                  Else
                     Text5(11) = ""
                     Text5(10) = ""
                     Text5(8) = ""
                     'Remove by Mogan 2010/1/25
                     'Text5(15) = ""
                     'Text10 = ""
                     
                  End If
                  '92.5.8 END
               End If
            Else
               Cancel = True
            End If
         End If
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      Case 6:
'         '2005/6/9 MODIFY BY SONIA
'         'If Text5(Index) <> "" Then
'         If Text5(Index) <> "" And pa(9) <> 台灣國家代號 Then
'            Text5(1) = ""
'            Text5(10) = ""
'            Text5(11) = ""
'            '92.10.17 add by sonia
'            Text9 = ""
'            Text10 = ""
'            '92.10.17 end
'         End If
'         '2006/2/7 ADD BY SONIA
'         'Modify by Morgan 2006/10/14
'         'If Me.Text5(Index).Text = "6" And pA(9) <> "013" Then
'         If Me.Text5(Index).Text = "6" And pa(9) <> "013" And pa(9) <> "044" Then
'            MsgBox "申請國家非香港案時不可選擇香港/澳門公告通知函 !", vbCritical
'            Cancel = True
'         End If
'         '2006/2/7 END
'
'end 2009/11/27
         
      Case 2 '機關文號
         'Modify by Morgan 2011/1/3 機關文號欄位改長度(百年問題)改抓MaxLength屬性控制
         If CheckLengthIsOK(Text5(Index), Text5(Index).MaxLength) = False Then
            Cancel = True
         End If
         'Modify by Morgan 2011/6/24
         'If Text5(index) = "" And pa(9) = 台灣國家代號 Then
         If Text5(Index) = "" And pa(9) = 台灣國家代號 And cp(10) <> 申請優先權證明 Then
            MsgBox "申請國家為台灣時不可空白，請重新輸入 !", vbCritical
            Cancel = True
         End If
         
'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
'
'      Case 9 '公告日
'         If Text5(Index) <> "" Then
'            If CheckIsDate(Text5(Index)) = False Then
'               Cancel = True
'            End If
'         End If
'
'end 2009/11/27

      Case 13
         If Text5(Index) <> "" Then
            'Modify by Morgan 2005/5/3 大陸控制西元年
            'If ChkDate(Text5(Index)) Then
            If CheckIsDate(Text5(Index)) Then
               If Val(TransDate(Text5(Index), 2)) > Val(strSrvDate(1)) Then
                  MsgBox "不可大於系統日 !", vbCritical
                  Cancel = True
               End If
            Else
               Cancel = True
            End If
         End If
      
      Case 10, 11 '領證本所期限, 領證法定期限
         If Text5(Index) <> "" Then
            If ChkDate(Text5(Index)) Then
               If Index = 11 Then Cancel = Not ChkRange(Text5(10), Text5(11), "日期")
            Else
               Cancel = True
            End If
            'Add By Cheng 2002/03/11
            '若有輸入領證本所期限, 則領證本所期限不可小於系統日
            If Index = 10 Then
               If Val(Me.Text5(Index).Text + 19110000) < strSrvDate(1) Then
                  MsgBox "領證本所期限不可小於系統日!!!", vbExclamation
                  Cancel = True
               End If
               'Add By Cheng 2003/12/08若領證本所期限非工作天則直接調整至最近的工作天
               If Cancel = False Then
                   Me.Text5(10).Text = TransDate(PUB_GetWorkDay1(Me.Text5(10).Text, True), 1)
               End If
               'End
            End If
         End If
         
      Case 12 '下次繳費日
         If Text5(Index) <> "" Then
            If ChkDate(Text5(Index)) = False Then
               Cancel = True
            End If
                     
            'Add By Cheng 2002/03/11
            '若有輸入下次繳費日, 則下次繳費日不可小於系統日
            '西元年
            If Len(Me.Text5(Index).Text) = 8 Then
               If Me.Text5(Index).Text < strSrvDate(1) Then
                  MsgBox "下次繳費日不可小於系統日!!!", vbExclamation
                  Cancel = True
               End If
            '民國年
            ElseIf Len(Me.Text5(Index).Text) = 7 Or Len(Me.Text5(Index).Text) = 6 Then
               If Val(Me.Text5(Index).Text + 19110000) < strSrvDate(1) Then
                  MsgBox "下次繳費日不可小於系統日!!!", vbExclamation
                  Cancel = True
               End If
            End If
         End If
   End Select
   If Cancel Then
      TextInverse Text5(Index)
      Text5(Index).SetFocus 'Added by Morgan 2021/12/16
   End If
End Sub
'92.1.14 add by sonia
Private Sub CheckYear()
'檢查是否通知下一年年費期限
'1.第二年  : (申請日+1年)之月份 - (核准日+5個月)之月份 <= 4
'2.第三年起: (申請日+N年)之月份 - (核准日+5個月)之月份 <= 2
Dim m_PA10 As String, m_Text5 As String, iYear As Integer
Dim m_PA10MM As String, m_Text5MM As String
Dim m_PA10DD As String, m_Text5DD As String
Dim m_DiffMM As String
   
   m_Year = "": iYear = 1
   m_PA10 = CompDate(0, iYear, pa(10))
   m_PA10MM = Val(Mid(m_PA10, 5, 2))
   m_PA10DD = Val(Mid(m_PA10, 7, 2))
   m_Text5 = CompDate(1, 5, Text5(0))
   m_Text5MM = Val(Mid(m_Text5, 5, 2))
   m_Text5DD = Val(Mid(m_Text5, 7, 2))
   
   If Val(m_PA10MM) < Val(m_Text5MM) Then m_PA10MM = m_PA10MM + 12
   m_DiffMM = m_PA10MM - m_Text5MM
   If m_DiffMM = 0 Then
      If Val(m_Text5DD) >= Val(m_PA10DD) Then m_DiffMM = 5
   End If
   
   '1
   'Modified by Morgan 2014/11/25 FMP 都抓4個月--玲玲
   'If m_PA10 >= m_Text5 Then
   If m_bolFMP Or m_PA10 >= m_Text5 Then
   'end 2014/11/25
      'Modified by Morgan 2014/11/24 +發明--玲玲
      'If pa(8) <> "1" And Val(m_DiffMM) <= 4 Then
      If Val(m_DiffMM) <= 4 Then
      'end 2014/11/24
         m_Year = "Y"
      End If
      Exit Sub
   Else
   '2
      'If pa(8) <> "1" Then 'Removed by Morgan 2014/11/24 +發明--玲玲
         If Val(m_DiffMM) < 2 Then
            m_Year = "Y"
         ElseIf m_DiffMM = 2 Then
            If Val(m_Text5DD) < Val(m_PA10DD) Then
               m_Year = "Y"
            End If
         End If
      'End If 'Removed by Morgan 2014/11/24 +發明--玲玲
      Exit Sub
   End If
End Sub
'92.1.14 end

Private Sub RefreshSpecData()
Dim strSql As String
Dim rsTmp As ADODB.Recordset
Dim strDate As String
Dim strCF10 As String
Dim strCF15 As String
Dim strCF24 As String
Dim strCP10 As String
Dim strExp As String
Dim strDayOrMon As String
Dim strIssue As String
Dim strNextOrg As String '下一救濟程序主管機關
   
   Label15(2).Caption = ""
   
   If pa(9) < "010" Then
      strCF10 = Empty
      strCF15 = Empty
      strCF24 = Empty
      strExp = Empty
      strDayOrMon = Empty
      
      strSql = "SELECT * FROM CASEFEE " & _
               "WHERE CF01 = '" & pa(1) & "' AND " & _
                     "CF02 = '" & pa(9) & "' AND " & _
                     "CF03 = '" & cp(10) & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         ' 主管機關
         If IsNull(rsTmp.Fields("CF10")) = False Then
            strCF10 = rsTmp.Fields("CF10")
         End If
         If IsNull(rsTmp.Fields("CF15")) = False Then
            ' 取得下一救濟程序名稱
            If pa(9) < "010" Then
               strCF15 = GetCaseTypeName(pa(1), rsTmp.Fields("CF15"), 0)
            Else
               strCF15 = GetCaseTypeName(pa(1), rsTmp.Fields("CF15"), 1)
            End If
            'Modify by Morgan 2005/11/21 抓下一救濟程序主管機關
            strExc(0) = "SELECT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF03='" & rsTmp.Fields("CF15") & "'"
            intI = 1
            Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strNextOrg = "" & AdoRecordSet3.Fields(0)
            End If
            '2005/11/21
         End If
            
         ' 主管機關文書
         If IsNull(rsTmp.Fields("CF24")) = False Then
            strCF24 = rsTmp.Fields("CF24")
         End If
      End If
      rsTmp.Close
      strSql = "SELECT * FROM CASEPROPERTYMAP " & _
               "WHERE CPM01 = '" & pa(1) & "' AND " & _
                     "CPM02 = '" & cp(10) & "' "
      Set rsTmp = New ADODB.Recordset
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         ' 來函期限
         If IsNull(rsTmp.Fields("CPM07")) = False Then
            Select Case rsTmp.Fields("CPM07")
               Case "1": strExp = "文到當日"
               Case "2": strExp = "文到次日"
            End Select
         End If
         ' 期限天數
         If IsNull(rsTmp.Fields("CPM08")) = False Then
            If IsEmptyText(rsTmp.Fields("CPM08")) = False Then
               strDayOrMon = rsTmp.Fields("CPM08") & "日"
            End If
         End If
         ' 期限月數
         If IsEmptyText(strDayOrMon) = True Then
            If IsNull(rsTmp.Fields("CPM09")) = False Then
               If IsEmptyText(rsTmp.Fields("CPM09")) = False Then
                  strDayOrMon = rsTmp.Fields("CPM09") & "個月"
               End If
            End If
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
      ' 取得案件名稱
      If pa(9) < "010" Then
         strCP10 = GetCaseTypeName(pa(1), cp(10), 0)
      Else
         strCP10 = GetCaseTypeName(pa(1), cp(10), 1)
      End If
      strIssue = "成立"
      ' 成立或不成立
      Select Case cp(10)
         Case "802", "804":   strIssue = "不成立"
      End Select
      ' 來函收文日
        'Modify By Cheng 2003/01/07
        If pa(9) = 台灣國家代號 Then
            strDate = Label2(3)
            strDate = ChangeTStringToWString(strDate)
            strDate = Left(strDate, 4) - 1911 & "年" & Mid(strDate, 5, 2) & "月" & Right(strDate, 2) & "日"
        Else
            strDate = Label2(3)
            strDate = ChangeTStringToWString(strDate)
            strDate = Left(strDate, 4) & "年" & Mid(strDate, 5, 2) & "月" & Right(strDate, 2) & "日"
        End If
      If pa(9) = 台灣國家代號 Then
         Select Case cp(10)
           'Modify By Cheng 2002/12/03
           Case 異議答辯, 舉發答辯
               'Modify By Cheng 2003/01/07
   '            Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
   '                Text5(2) & StrCp10 & strCF24 & "通知書(如附件)" & "," & "本案" & Left(StrCp10, 2) & strIssue & "。" & _
   '                "惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向經濟部提出" & _
   '                strCF15 & "，倘逾期仍未提起則本案即告確定。"
   
               'Modify by Morgan 2005/11/21 改 "經濟部" --> " & strNextOrg & "
               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
                   Text5(2) & Left(strCP10, 2) & strCF24 & "通知（如附件）" & "，" & "本案" & Left(strCP10, 2) & strIssue & "。" & _
                   "惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向" & strNextOrg & "提出" & _
                   strCF15 & "，倘逾期仍未提起則本案即告確定。"
                   
'2010/11/12 CANCEL BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
'           '92.9.23 ADD BY SONIA
'           Case 行政訴訟
'               'Modify by Morgan 2005/11/21
''               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
''                   Text5(2) & StrCp10 & strCF24 & "通知（如附件）" & "，" & "本案訴願決定及原處分均撤銷。" & _
''                   "本案獲勝係行政法院採信我方之理由。惟依法對方如不服此判決，可於" & strExp & "起" & strDayOrMon & "內" & _
''                   "提起" & strCF15 & "，倘逾期仍未提起則本案判決即告確定。"
'
'               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
'                   Text5(2) & strCP10 & strCF24 & "通知（如附件）" & "，" & "本案訴願決定及原處分均撤銷。" & _
'                   "本案獲勝係" & strCF10 & "採信我方之理由。惟依法對方如不服此判決，可於" & strCF24 & "送達後" & strDayOrMon & "內" & _
'                   "向" & strNextOrg & "提起" & strCF15 & "，倘逾期仍未提起則本案判決即告確定。"
'2010/11/12 END
           'Add by Morgan 2006/8/14
           Case "505" '參加訴願
               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & Text5(2) & _
                   strCF24 & "通知（如附件），本案訴願駁回。惟依法對方如不服此判決" & _
                   "，可於" & strCF24 & "送達後" & strDayOrMon & "內向" & strNextOrg & "提起" & strCF15 & "。"
                                      
           '92.9.23 END
           '93/2/5 ADD By SONIA
           Case 參加訴訟
               'Modify by Morgan 2005/11/21
'               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & Text5(2) & _
'                   strCF24 & "通知（如附件），本案原告之訴駁回。惟依法對方如不服此判決" & _
'                   "，應於" & strCF24 & "送達後二十日內提起上訴，倘逾期仍未提起則本案即告確定。"

               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & Text5(2) & _
                   strCF24 & "通知（如附件），本案原告之訴駁回。惟依法對方如不服此判決" & _
                   "，可於" & strCF24 & "送達後" & strDayOrMon & "內向" & strNextOrg & "提起" & strCF15 & "。"
                   
'2010/11/12 MODIFY BY SONIA 訴願或行政訴訟或上訴的核准請改輸  一般來函的撤銷原處分
'           'Add by Morgan 2005/2/4
'           Case 行政訴訟上訴
'               'Modify by Morgan 2005/11/21
''               Text19.Text = "本案獲" & strCF10 & "於" & strDate & "以" & _
''                   Text5(2) & StrCp10 & strCF24 & "通知（如附件）" & "，" & "本案原判決廢棄，發回台北高等行政法院。本案獲勝係主管機關採信我方之理由。"
'               Text19.Text = "本案獲" & strCF10 & "於" & strDate & "以" & _
'                   Text5(2) & strCP10 & strCF24 & "通知（如附件）" & "，" & "本案原判決廢棄，發回台北高等行政法院。本案獲勝係" & strCF10 & "採信我方之理由。"
'
'           '93/2/5 END
'           'Add by Morgan 2005/3/28
'           Case 訴願
'               'Modify by Morgan 2005/10/7 改固定用'原處分撤銷'
'               '本案" & Left(StrCp10, 2) & strIssue & "。
'               'Modify by Morgan 2005/11/17 改內容--郭,游經理
''               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
''                   Text5(2) & StrCp10 & strCF24 & "通知（如附件）" & "，" & "本案原處分撤銷並由智慧局另為處分的決定。" & _
''                   "本案獲勝係主管機關採信我方之理由。惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向台北高等行政法院提出" & _
''                   strCF15 & "，倘逾期仍未提起則本案即告確定。"
'               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
'                   Text5(2) & strCP10 & strCF24 & "通知（如附件）" & "，" & "本案原處分撤銷並由智慧局另為處分的決定。" & _
'                   "本案獲勝係" & strCF10 & "採信我方之理由。惟依法對方如不服此決定，可於" & strExp & "起" & strDayOrMon & "內向" & strNextOrg & "提出" & _
'                   strCF15 & "。"
'2010/11/12 END
           Case Else
               'Modify By Cheng 2003/01/07
   '            Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
   '                Text5(2) & StrCp10 & strCF24 & "通知書(如附件)" & "," & "本案" & Left(StrCp10, 2) & strIssue & "。" & _
   '                "本案獲勝係主管機關採信我方之理由。惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向經濟部提出" & _
   '                strCF15 & "，倘逾期仍未提起則本案即告確定。"
               'Modify by Morgan 2005/11/21
'               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
'                   Text5(2) & StrCp10 & strCF24 & "通知（如附件）" & "，" & "本案" & Left(StrCp10, 2) & strIssue & "。" & _
'                   "本案獲勝係主管機關採信我方之理由。惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向經濟部提出" & _
'                   strCF15 & "，倘逾期仍未提起則本案即告確定。"
                   
               Text19.Text = "本案已獲" & strCF10 & "於" & strDate & "以" & _
                   Text5(2) & strCP10 & strCF24 & "通知（如附件）" & "，" & "本案" & Left(strCP10, 2) & strIssue & "。" & _
                   "本案獲勝係" & strCF10 & "採信我方之理由。惟依法對方如不服此審定，可於" & strExp & "起" & strDayOrMon & "內向" & strNextOrg & "提出" & _
                   strCF15 & "，倘逾期仍未提起則本案即告確定。"
                   
         End Select
      End If
      '910704 Sieg 405
      Label15(2).Caption = strCF10
      Combo2.Clear
      strExc(0) = "SELECT DISTINCT CF10 FROM CASEFEE WHERE CF01='" & pa(1) & "' AND CF02='" & pa(9) & "' AND CF10 IS NOT NULL"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      Do While Not RsTemp.EOF
         Combo2.AddItem RsTemp.Fields("CF10")
         RsTemp.MoveNext
      Loop
      If Combo2.ListCount > 0 Then Combo2.Text = Label15(2)
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim arrCaseNo() As String 'Added by Morgan 2021/2/25

   TxtValidate = False
   
   'Added by Morgan 2021/12/16 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/16
   
   For Each objTxt In Text5
      If objTxt.Enabled = True Then
         Cancel = False
         Text5_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Me.Text5(objTxt.Index).SetFocus
            Text5_GotFocus objTxt.Index
            Exit Function
         End If
      End If
   Next
   '92.10.17 ADD BY SONIA
   
   Text9_Validate Cancel
   If Cancel = True Then
      Text9.SetFocus
      Exit Function
   End If
   
'Remove by Mogan 2010/1/25
'   Text10_Validate Cancel
'   If Cancel = True Then
'      Text10.SetFocus
'      Exit Function
'   End If

   '92.10.17 END
   
   'Add by Morgan 2004/9/10
   If txt412.Visible = True Then
      txt412_Validate Cancel
      If Cancel = True Then
         txt412.SetFocus
         Exit Function
      End If
   End If
   
   'Added by Morgan 2022/5/12 Ex:P-110285 --陳玲玲
   If pa(9) = "000" And cp(10) = "908" Then
      If MsgBox("請確認退費後繳費年度及下一程序的期限管制是否正確！" & vbCrLf & vbCrLf & "【 是:繼續存檔   否:回畫面 】", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   'end 2022/5/12

   'Add by Morgan 2008/6/16
   If Text5(8) <> "" Then
      If Val(Text5(8)) < m_dbl601OfficialFee Then
         MsgBox "領證費輸入錯誤，點數不可小於0！"
         Text5(8).SetFocus
         Exit Function
      End If
   End If
   'end 2008/6/16
   
   
   'Added by Morgan 2014/11/24
   '帳單檢查(參考frm04010505_2)
   '專利處只須以代理人+代理人D/N No.做重覆檢核
   If Text1 = "P" And Left(Pub_StrUserSt03, 2) = "P1" Then
      '若有輸入代理人D/N No.
      If Me.Text12.Text <> "" Then
         If PUB_ChkDNDup("", ChangeCustomerL(m_CP44), Text12.Text) = True Then
            Text12.SetFocus
            Text12_GotFocus
            Exit Function
         End If
      End If
   Else
      '若有輸入代理人D/N No.或帳單日期
      If Me.Text12.Text <> "" Or Me.Text13.Text <> "" Then
         If PUB_ChkDNDup(Text13.Text, ChangeCustomerL(m_CP44), Text12.Text) = True Then
            Text12.SetFocus
            Text12_GotFocus
            Exit Function
         End If
      End If
   End If
   'end 2014/11/24

   'Added by Morgan 2014/4/25
   'Modified by Morgan 2016/7/5 非臺灣案電子化
   'If pa(9) = "000" And Text6 = "" And (cp(10) = "421" Or cp(10) = "807") Then
   If Left(Pub_StrUserSt03, 1) <> "F" And Text6 = "" And (cp(10) = "421" Or cp(10) = "807") Then
   'end 2016/7/5
      MsgBox "請輸入引證前案檔案數量!!", vbExclamation
      Text6.SetFocus
      Exit Function
   End If
   'end 2014/4/25

   'Added by Morgan 2014/5/15 電子化-檢查pdf檔
   If pa(9) = "000" Then
      If PUB_CheckPDF(pa(1), pa(2), pa(3), pa(4), 1 + Val(Text6), m_DocNo) = False Then
         Exit Function
      End If
   End If
   'end 2014/5/15
   
   'Added by Morgan 2023/2/23
   If pa(9) = "000" And m_strCP10 = "415" Then
      If txt415Date = "" Then
         MsgBox "專利權期間延長後日期！", vbCritical
         txt415Date.SetFocus
         Exit Function
      Else
         Cancel = False
         Call txt415Date_Validate(Cancel)
         If Cancel = True Then
            txt415Date_GotFocus
            Exit Function
         End If
      End If
   End If
   'end 2023/2/23
   
   'Added by Morgan 2024/9/25
   If bolCN445 Then
      If txt415Date = "" Then
         MsgBox "專利權期滿終止日不可空白！", vbCritical
         txt415Date.SetFocus
         Exit Function
      Else
         Cancel = False
         Call txt415Date_Validate(Cancel)
         If Cancel = True Then
            txt415Date_GotFocus
            Exit Function
         End If
      End If
   
      If m_strIR01 <> "" Then
         If PUB_ChkReKeyInOk(m_strIR01, m_strIR02, m_strIR03, m_strIR04, txt415Date, m_bolReKeyInOK) = False Then
            txt415Date.SetFocus
            Exit Function
         End If
      End If
   End If
   'end 2024/9/25
   
   'Added by Morgan 2019/5/24 從寫定稿例外欄位移來此處先作檢查
   'IDS報價檢查
   m_USCaseNo = ""
   'Modified by Morgan 2019/6/14 申請案核准要排除收文審查意見,檢索報告要控制發明案--郭雅娟
   'If (pa(9) = "000" And (pa(8) = "1" Or pa(8) = "3") And InStr("101,103,301,303,107,307", cp(10)) > 0) Or cp(10) = "421" Then
   '   m_USCaseNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
   'End If
   If (pa(9) = "000" And (pa(8) = "1" Or pa(8) = "3") And InStr("101,103,301,303,107,307", cp(10)) > 0) Then
      'Modified by Morgan 2023/12/7 +判斷來函的相關收文號為點選的收文號(因再審要判斷再審的審查意見通知,發文申請的不算)--郭
      'If PUB_ChkCPExist(cp, "1202") = False Then
      If PUB_ChkCPExist(cp, "1202", , , , , , cp(9)) = False Then
      'end 2023/12/7
         m_USCaseNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
         'Added by Morgan 2021/6/3
         '若有通知擇一申復(1232)時詢問
         If m_USCaseNo <> "" Then
            If PUB_ChkCPExist(cp, "1232") = True Then
               If MsgBox("本案曾有通知擇一申復來函，是否要通知ＩＤＳ報價？", vbYesNo + vbDefaultButton1 + vbExclamation, "ＩＤＳ報價") = vbNo Then
                  m_USCaseNo = ""
               End If
            'Added by Morgan 2024/4/10 再審准也要問
            ElseIf cp(10) = "107" Then
               strExc(0) = "1.請確認引證前案是否與前一次審查意見通知書相同。" & vbCrLf & _
                           "2.請確認美國案 " & m_USCaseNo & " 是否已提出相同引證前案的IDS。" & vbCrLf & _
                           "若二者均相同可不必通知IDS報價"
               strExc(0) = strExc(0) & vbCrLf & vbCrLf & "【是】:要通知    【否】:不通知    【取消】:回畫面" 'Added by Morgan 2020/12/18
               intI = MsgBox(strExc(0), vbYesNoCancel + vbInformation + vbDefaultButton3, "是否通知IDS報價？")
               If intI = vbCancel Then
                  Exit Function
               ElseIf intI = vbNo Then
                  m_USCaseNo = ""
               End If
            'end 2024/4/10
            End If
         End If
         'end 2021/6/3
      End If
   'Removed by Morgan 2019/10/9 P案檢索報告不用--郭雅娟
   'ElseIf (pa(8) = "1" And cp(10) = "421") Then
   '   m_USCaseNo = PUB_GetUSCaseNo(pa(1), pa(2), pa(3), pa(4))
   'end 2019/10/9
   End If
   'end 2019/6/14
   
   If m_USCaseNo <> "" Then
      If txtIDSFee(1) = "" Or txtIDSFee(2) = "" Or txtIDSPt(1) = "" Or txtIDSPt(2) = "" Then
         If MsgBox("尚未輸入ＩＤＳ報價，是否 EMail 通知 CFP 程序人員報價？", vbYesNo + vbDefaultButton2 + vbExclamation, "ＩＤＳ報價") = vbYes Then
            'Modified by Morgan 2019/10/9 P案檢索報告不用--郭雅娟
            'strExc(0) = ""
            'If cp(10) = "421" Then
            '   If pa(9) = "000" Then
            '      strExc(0) = "技術報告"
            '   Else
            '      strExc(0) = "檢索報告"
            '   End If
            'Else
            '   strExc(0) = "核准函"
            'End If
            strExc(0) = "核准函"
            
            'Added by Morgan 2023/5/24
            If Text6.Text = "" Then
               Do
                  Text6.Text = InputBox("請輸入引證前案檔案數量：")
                  If Val(Text6.Text) > 0 Then
                     Exit Do
                  ElseIf Text6.Text = "" Then
                     MsgBox "未輸入引證前案檔案數量，取消 EMail 通知！", vbExclamation
                     Exit Function
                  Else
                     MsgBox "引證前案檔案數量必須大於 0，請重新輸入！", vbExclamation
                  End If
               Loop
            End If
            'end 2023/5/24
            
            'end 2019/10/9
            strExc(2) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & pa(3) & "-" & pa(4))
            'Modified by Morgan 2021/2/25 考慮會有多個美國案
            'strExc(1) = PUB_GetCFPHandler(m_USCaseNo)
            'strExc(4) = strExc(2) & " 案已收到" & strExc(0) & "，請提供相關美國案( " & m_USCaseNo & " )的IDS報價！"
            arrCaseNo = Split(m_USCaseNo, "、")
            For ii = LBound(arrCaseNo) To UBound(arrCaseNo)
               strExc(1) = PUB_GetCFPHandler(arrCaseNo(ii))
               strExc(4) = strExc(2) & " 案已收到" & strExc(0) & "，請提供相關美國案( " & arrCaseNo(ii) & " )的IDS報價！"
            'end 2021/1/25
               If strExc(1) <> "" Then
                  'Modified by Morgan 2019/9/9 調整報價欄位名及定稿內容--郭
                  'Modified by Morgan 2023/5/24 +引證前案檔案數量
                  strExc(3) = "引證前案共: " & Text6.Text & " 件" & vbCrLf & _
                              "IDS報價:" & vbCrLf & _
                              "　1.第一階段　　　(　P)" & vbCrLf & _
                              "　2.第二階段　　　(　P)" & vbCrLf & vbCrLf & _
                              "**　若該案已是第二階段，則第一階段請輸　0　**"
   
                  PUB_SendMail strUserNum, strExc(1), "", strExc(4), strExc(3)
               End If
            Next 'Added by Morgan 2021/2/25
            
         ElseIf txtIDSFee(1) = "" Then
            txtIDSFee(1).SetFocus
         ElseIf txtIDSPt(1) = "" Then
            txtIDSPt(1).SetFocus
         ElseIf txtIDSFee(2) = "" Then
            txtIDSFee(2).SetFocus
         ElseIf txtIDSPt(2) = "" Then
            txtIDSPt(2).SetFocus
         End If
         Exit Function
      End If
   End If
   'end 2019/5/24

   'Added by Morgan 2020/1/16
   '大陸案,有通知函,程序承辦,非掛號(無期限)
   m_bolNoCP27 = False
   'Removed by Morgan 2024/1/30 取消--郭
   'If pa(9) = "020" And Text5(4) <> "N" And m_bolEngCase = False And Text5(11) = "" Then
   '   If PUB_GetCustomerValue(pa(26), "CU182") = "Y" Then
   '      If MsgBox("請確認是否已收到公文正本？", vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then
   '         m_bolNoCP27 = True
   '      End If
   '   End If
   'End If
   'end 2020/1/16
   
   'Added by Morgan 2025/3/7 面詢未辦理，向官方辦理退費控管--玲玲
   m_bolAddB908 = False
   If bolChgRlt And pa(9) = "000" Then
      strExc(0) = "select * from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='407' and cp27>19221111" & _
         " and not exists(select * from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp10='408' and cp27>a.cp27)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         intI = MsgBox("本案曾申請【請求面詢】但未辦理，請問是否自動收文【代辦退費】？", vbQuestion + vbYesNoCancel + vbDefaultButton3)
         If intI = vbYes Then
            m_bolAddB908 = True
         ElseIf intI = vbCancel Then
            Exit Function
         End If
      End If
   End If
   'end 2025/3/7
   
   TxtValidate = True
End Function

Private Sub Text6_GotFocus()
   CloseIme
   TextInverse Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
   End If
End Sub

Private Sub Text9_GotFocus()
  TextInverse Text9
End Sub

'92.10.17 add by sonia
Private Sub Text9_Validate(Cancel As Boolean)
   If Check606(0) = False Then
      Cancel = True
   End If
End Sub

'Remove by Mogan 2010/1/25
'Private Sub Text10_Validate(Cancel As Boolean)
'   If Check606(1) = False Then
'      Cancel = True
'   End If
'End Sub

'92.10.17 end

Private Function Check606(iOpt As Integer) As Boolean
   'Modify by Morgan 2008/10/2 +分割(因分割不一定是發明故需控制專利種類)
   'If pa(9) = 大陸國家代號 And Text5(6) = "" And cp(10) = "101" Then
   
   'Remove by Morgan 2009/11/27 非台灣通知書移到 frm04010514
   'If pa(9) = 大陸國家代號 And pa(8) = "1" And Text5(6) = "" And (cp(10) = "101" Or cp(10) = "307") Then
   If pa(9) = 大陸國家代號 And pa(8) = "1" And (cp(10) = "101" Or cp(10) = "307") Then
      If iOpt = 0 Then
         If Text9 = "" Then
            MsgBox "申請國家為大陸時, 大陸年度不可空白，請重新輸入 !", vbCritical
            Exit Function
         Else
            If Val(Text9) < 1 Or Val(Text9) > 20 Then
               MsgBox "大陸年度錯誤，請重新輸入 !", vbCritical
               Exit Function
            End If
            'Modify by Morgan 2008/6/16 若未設定領證費或不是存檔前檢查時才預設以免改過的金額又被蓋掉
            'Modified by Morgan 2013/6/26 領證費改只顯示不可修改,存檔前一率重算
            'If Text5(8) = "" Or m_bolSaveCheck = False Then
               Text5(8) = GetFee
            'End If
            'end 2013/6/26
            'end 2008/5/28
            GetNextDate 'Add by Morgan 2006/3/20
         End If
         
'Remove by Morgan 2010/1/25
'      '維持費起始年度
'      Else
'         If Text10 = "" Then
'            MsgBox "申請國家為大陸時, 維持費起始年度不可空白，請重新輸入 !", vbCritical
'            Exit Function
'         Else
'            If Val(Text10) < 1 Or Val(Text10) > Val(Text9) Then
'               MsgBox "維持費起始年度錯誤，請重新輸入 !", vbCritical
'               Exit Function
'            End If
'         End If

      End If
      
'Remove by Morgan 2010/1/25
'      '計算大陸維持費
'      '維持費(第3年起每年3500)
'      If Val(Text9) > 0 And Val(Text10) > 0 Then
'         Text5(15) = Format(3500 * (Val(Text9) - Val(Text10)))
'      Else
'         Text5(15) = ""
'      End If
'   Else
'      Text5(15) = ""

   End If
   Check606 = True
End Function

Private Sub txt412_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txt412.IMEMode = 2
   CloseIme
   TextInverse txt412
End Sub

Private Sub txt412_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not (KeyAscii >= 48 And KeyAscii <= 57) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txt412_Validate(Cancel As Boolean)
      
   If IsEmptyText(txt412) = True Then
      If m_strCP10 = "412" Then
         MsgBox "延緩公告日不可空白 !", vbCritical
      'Added by Morgan 2024/6/20
      ElseIf m_strCP10 = "245" Then
         'Modified by Lydia 2025/02/12
         'MsgBox "續行審查日不可空白 !", vbCritical
         MsgBox "延緩審查日期不可空白 !", vbCritical
      'end 2024/6/20
      
      'Added by Morgan 2024/9/18
      ElseIf bolCN445 Then
         MsgBox "補償天數不可空白 !", vbCritical
      'end 2024/918
      
      End If
      Cancel = True
   
   'Added by Morgan 2024/9/18
   ElseIf bolCN445 Then
      Exit Sub
   'end 2024/9/18
   'Modified by Lydia 2025/02/12
   'ElseIf ChkDate(txt412) = True Then
   '   If m_strCP10 = "412" Then
   ElseIf m_strCP10 = "412" Or m_strCP10 = "245" Then
      If m_strCP10 = "412" Then
         If ChkDate(txt412) = True Then
   'end 2025/02/12
            'Modify by Morgan 2006/10/12
            'If Val(txt412) < Val(strSrvDate(2)) Then
            '   MsgBox "延緩公告日不可小於系統日 !", vbCritical
            If DBDATE(txt412) <> DBDATE(m_strPA14) Then
               MsgBox "核准延緩公告日不可與原申請延緩公告日不同!", vbCritical
            'end 2006/10/12
               txt412_GotFocus
               Cancel = True
            End If
         End If
      End If
      'Added by Lydia 2025/02/12
      If m_strCP10 = "245" Then
         If CheckIsTaiwanDate(txt412) = False Then
            txt412_GotFocus
            Cancel = True
         Else
            If cp(71) <> "" And DBDATE(txt412) <> DBDATE(cp(71)) Then
               MsgBox "延緩審查日期與發文輸入延緩審查日期不同！", vbCritical
               txt412_GotFocus
               Cancel = True
            End If
         End If
      End If
      'end 2025/02/12
   Else
      txt412_GotFocus
      Cancel = True
   End If
End Sub
'Add by Morgan 2008/5/28 原輸核准日及大陸年度時都會做改統一在此設定
'Modify by Morgan 2008/6/16 加抓規費 m_dbl601OfficialFee 以便計算點數
'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
'設定領證費
Private Function GetFee() As Long
'Modified by Morgan2014/12/24 改呼叫共用函數
'Added by Lydia 2015/03/27 +PA01
   
   'Modified by Morgan 2018/2/27 取消 PUB_GetFee 改用 PUB_Get020601Fee(與接洽單一致)
   'GetFee = PUB_GetFee(pa(1), pa(9), pa(8), pa(26), Val(Text9), m_dbl601OfficialFee, cp(44))
   Dim dblSFee As Double
   m_dbl601OfficialFee = PUB_Get020601Fee(pa, cp(44), Val(Text9), Val(Text9), dblSFee)
   GetFee = dblSFee + m_dbl601OfficialFee
   'end 2018/2/27
   
'   Dim lTmp As Long, lBase As Long, lPlus As Long
'   '領證費
'   'Modify by Morgan 2010/12/9 客戶只抓服務費,規費改抓代理人設定,否則調價後會不同步
'   'lTmp = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "1")
'   lTmp = PUB_GetYF06(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "1")
'   'end 2010/12/9
'   If lTmp = 0 Then
'      If cp(44) <> "" Then
'         lTmp = PUB_GetYF0607(pa(9), pa(8), ChangeCustomerL(cp(44)), "601", "1", "1")
'      End If
'      If lTmp = 0 Then
'         lTmp = PUB_GetYF0607(pa(9), pa(8), "Y00000001", "601", "1", "1")
'         m_dbl601OfficialFee = PUB_GetYF07(pa(9), pa(8), "Y00000001", "601", "1", "1")
'      Else
'         m_dbl601OfficialFee = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(cp(44)), "601", "1", "1")
'      End If
'   Else
'      'Modify by Morgan 2010/12/9 客戶只抓服務費,規費改抓代理人設定,否則調價後會不同步
'      'm_dbl601OfficialFee = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(pa(26)), "601", "1", "1")
'      If cp(44) <> "" Then
'         m_dbl601OfficialFee = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(cp(44)), "601", "1", "1")
'      End If
'      If m_dbl601OfficialFee = 0 Then
'         m_dbl601OfficialFee = PUB_GetYF07(pa(9), pa(8), "Y00000001", "601", "1", "1")
'      End If
'      lTmp = lTmp + m_dbl601OfficialFee
'      'end 2010/12/9
'   End If
'
'   If Val(Text9) > 0 And Text9 <> "3" Then
'      '年費(先抓1-3年的Base金額,再抓輸入大陸年度的金額, 計算出差額, 再加上上面之領證則為大陸領證費)
'      lBase = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(pa(26)), "605", "3", "3")
'      If lBase = 0 Then
'         If cp(44) <> "" Then
'            lBase = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(cp(44)), "605", "3", "3")
'         End If
'         If lBase = 0 Then
'            lBase = PUB_GetYF07(pa(9), pa(8), "Y00000001", "605", "3", "3")
'         End If
'      End If
'      lPlus = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(pa(26)), "605", Text9, Text9)
'      If lPlus = 0 Then
'         If cp(44) <> "" Then
'            lPlus = PUB_GetYF07(pa(9), pa(8), ChangeCustomerL(cp(44)), "605", Text9, Text9)
'         End If
'         If lPlus = 0 Then
'            lPlus = PUB_GetYF07(pa(9), pa(8), "Y00000001", "605", Text9, Text9)
'         End If
'      End If
'   End If
'   GetFee = lTmp + lPlus - lBase
'   m_dbl601OfficialFee = m_dbl601OfficialFee + lPlus - lBase
'end 2014/12/24
End Function

Private Sub Combo3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo3_Validate(Cancel As Boolean)
   If Combo3 = MsgText(601) Then
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo3, Label25(4)) = False Then
      Cancel = True
      Combo3.SetFocus
   End If
End Sub

Private Sub txt415Date_GotFocus()
   TextInverse txt415Date
End Sub

Private Sub txt415Date_Validate(Cancel As Boolean)
   If txt415Date <> "" Then
      Cancel = Not ChkDate(txt415Date)
      If Cancel = False Then
         If DBDATE(txt415Date) <= DBDATE(pa(25)) Then
            MsgBox "延長後專用期必須大於目前專用期！", vbCritical, "專利權期滿終止日錯誤"
            Cancel = True
         End If
      End If
      
      'Added by Morgan 2024/9/25
      '專利權期滿終止日=原專利權期滿終止日(原專用期止日+1天)+補償天數
      If bolCN445 And Cancel = False Then
         strExc(0) = CompDate(2, Val(txt412) + 1, pa(25))
         If DBDATE(txt415Date) <> strExc(0) Then
            MsgBox "專利權期滿終止日錯誤，應該為 " & strExc(0) & "，請再確認！", vbCritical
            Cancel = True
         End If
      End If
      'end 2024/9/25
   End If
End Sub

Private Sub txtDispDate_GotFocus()
   TextInverse txtDispDate
End Sub

Private Sub txtIDSFee_GotFocus(Index As Integer)
   TextInverse txtIDSFee(Index)
End Sub

Private Sub txtIDSFee_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtIDSPt_GotFocus(Index As Integer)
   TextInverse txtIDSPt(Index)
End Sub

Private Sub txtIDSPt_KeyPress(Index As Integer, KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub txtPriNo_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2023/06/15
Private Sub Command2_Click()

   m_bolRePriDate = True
   ModifyPriority strPriority(1), strPriority(2), strPriority(3), pa(8), , pa(1) & pa(2) & pa(3) & pa(4), pa(9), True, strPriority(4), strPriority(5)
   
End Sub
