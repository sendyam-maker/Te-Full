VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010401_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核准輸入"
   ClientHeight    =   5748
   ClientLeft      =   1416
   ClientTop       =   2040
   ClientWidth     =   8952
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8952
   Begin VB.CommandButton Command2 
      Caption         =   "商品及服務資料輸入(&I)"
      Height          =   375
      Left            =   2820
      TabIndex        =   23
      Top             =   30
      Width           =   1965
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "變更事項(R)"
      Height          =   375
      Left            =   4815
      TabIndex        =   24
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Left            =   6870
      TabIndex        =   26
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   6030
      TabIndex        =   25
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   375
      Left            =   8130
      TabIndex        =   27
      Top             =   30
      Width           =   800
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   52
      TabStop         =   0   'False
      Top             =   2565
      Width           =   2412
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5970
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   50
      TabStop         =   0   'False
      Top             =   2265
      Width           =   2412
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   2265
      Width           =   2412
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5970
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2412
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1965
      Width           =   2412
   End
   Begin VB.TextBox textTM22S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   6690
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1665
      Width           =   1692
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1665
      Width           =   2412
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5970
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1365
      Width           =   2412
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5970
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   435
      Width           =   2412
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1740
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1365
      Width           =   2412
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   435
      Width           =   2412
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2835
      Left            =   30
      TabIndex        =   56
      Top             =   2880
      Width           =   8895
      _ExtentX        =   15685
      _ExtentY        =   4995
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm02010401_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label24"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label21"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label20"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label19"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label18"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label17"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Line1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label13"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label12"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label11"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label10"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label9"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label8"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label7"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(19)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textPS"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP35"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtPrtM"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtPrtY"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtNote"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP54"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP53"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textPrint"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textTM17"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM16S"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM29"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTM22"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM21"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTMBM07_2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textTMBM07_1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM14"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP08"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP25"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textTM15"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCP64"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "馬德里領証期限"
      TabPicture(1)   =   "frm02010401_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label26"
      Tab(1).Control(1)=   "Label32"
      Tab(1).Control(2)=   "Label33"
      Tab(1).Control(3)=   "textNP08"
      Tab(1).Control(4)=   "textNP09"
      Tab(1).Control(5)=   "textFee"
      Tab(1).ControlCount=   6
      Begin VB.TextBox textCP64 
         Height          =   270
         Left            =   5352
         MaxLength       =   2000
         TabIndex        =   19
         Top             =   2475
         Width           =   2290
      End
      Begin VB.TextBox textFee 
         Height          =   270
         Left            =   -73530
         TabIndex        =   22
         Top             =   990
         Width           =   1125
      End
      Begin VB.TextBox textNP09 
         Height          =   270
         Left            =   -70950
         MaxLength       =   7
         TabIndex        =   21
         Top             =   660
         Width           =   1125
      End
      Begin VB.TextBox textNP08 
         Height          =   270
         Left            =   -73530
         MaxLength       =   7
         TabIndex        =   20
         Top             =   660
         Width           =   1125
      End
      Begin VB.TextBox textTM15 
         Height          =   270
         Left            =   1032
         MaxLength       =   20
         TabIndex        =   0
         Top             =   390
         Width           =   2412
      End
      Begin VB.TextBox textCP25 
         Height          =   270
         Left            =   5352
         MaxLength       =   7
         TabIndex        =   1
         Top             =   390
         Width           =   2412
      End
      Begin VB.TextBox textCP08 
         Height          =   270
         Left            =   1032
         MaxLength       =   40
         TabIndex        =   2
         Top             =   660
         Width           =   2412
      End
      Begin VB.TextBox textTM14 
         Height          =   270
         Left            =   1032
         MaxLength       =   7
         TabIndex        =   4
         Top             =   990
         Width           =   2412
      End
      Begin VB.TextBox textTMBM07_1 
         Height          =   270
         Left            =   5352
         MaxLength       =   2
         TabIndex        =   5
         Top             =   990
         Width           =   732
      End
      Begin VB.TextBox textTMBM07_2 
         Height          =   270
         Left            =   6672
         MaxLength       =   4
         TabIndex        =   6
         Top             =   990
         Width           =   732
      End
      Begin VB.TextBox textTM21 
         Height          =   270
         Left            =   1032
         MaxLength       =   7
         TabIndex        =   7
         Top             =   1290
         Width           =   972
      End
      Begin VB.TextBox textTM22 
         Height          =   270
         Left            =   2352
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1290
         Width           =   1092
      End
      Begin VB.TextBox textTM29 
         Height          =   270
         Left            =   5352
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1290
         Width           =   732
      End
      Begin VB.TextBox textTM16S 
         Height          =   270
         Left            =   1395
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1590
         Width           =   405
      End
      Begin VB.TextBox textTM17 
         Height          =   270
         Left            =   5652
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1590
         Width           =   435
      End
      Begin VB.TextBox textPrint 
         Height          =   270
         Left            =   1032
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1890
         Width           =   732
      End
      Begin VB.TextBox textCP53 
         Height          =   270
         Left            =   5955
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1890
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox textCP54 
         Height          =   270
         Left            =   7530
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1890
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.TextBox txtNote 
         Enabled         =   0   'False
         Height          =   270
         Left            =   6420
         TabIndex        =   16
         Top             =   2175
         Visible         =   0   'False
         Width           =   2000
      End
      Begin VB.TextBox txtPrtY 
         Height          =   270
         Left            =   1032
         MaxLength       =   10
         TabIndex        =   17
         Top             =   2475
         Width           =   690
      End
      Begin VB.TextBox txtPrtM 
         Height          =   270
         Left            =   2148
         MaxLength       =   10
         TabIndex        =   18
         Top             =   2475
         Width           =   750
      End
      Begin MSForms.TextBox textCP35 
         Height          =   300
         Left            =   5355
         TabIndex        =   3
         Top             =   660
         Width           =   2415
         VariousPropertyBits=   679493659
         MaxLength       =   32
         Size            =   "4254;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPS 
         Height          =   300
         Left            =   1032
         TabIndex        =   15
         Top             =   2175
         Width           =   4560
         VariousPropertyBits=   -1467989989
         MaxLength       =   128
         ScrollBars      =   2
         Size            =   "8043;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "大陸案受理函發文日 :                                                     (西元)"
         Height          =   180
         Index           =   19
         Left            =   3516
         TabIndex        =   91
         Top             =   2508
         Width           =   4740
      End
      Begin VB.Label Label33 
         Caption         =   "領証費 :"
         Height          =   255
         Left            =   -74790
         TabIndex        =   90
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label32 
         Caption         =   "領証法定期限 :"
         Height          =   255
         Left            =   -72210
         TabIndex        =   89
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label26 
         Caption         =   "領証本所期限 :"
         Height          =   255
         Left            =   -74790
         TabIndex        =   88
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "審定號 :"
         Height          =   255
         Left            =   75
         TabIndex        =   87
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "核准通知日 :"
         Height          =   252
         Left            =   4152
         TabIndex        =   86
         Top             =   396
         Width           =   1332
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   75
         TabIndex        =   85
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "審查委員 :"
         Height          =   255
         Left            =   4155
         TabIndex        =   84
         Top             =   660
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "公告日 :"
         Height          =   255
         Left            =   75
         TabIndex        =   83
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label11 
         Caption         =   "公報卷期 :"
         Height          =   252
         Left            =   4152
         TabIndex        =   82
         Top             =   996
         Width           =   1092
      End
      Begin VB.Label Label12 
         Caption         =   "卷"
         Height          =   252
         Left            =   6192
         TabIndex        =   81
         Top             =   996
         Width           =   252
      End
      Begin VB.Label Label13 
         Caption         =   "期"
         Height          =   252
         Left            =   7572
         TabIndex        =   80
         Top             =   996
         Width           =   252
      End
      Begin VB.Label Label14 
         Caption         =   "專用期限 :"
         Height          =   255
         Left            =   75
         TabIndex        =   79
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   2112
         X2              =   2232
         Y1              =   1416
         Y2              =   1416
      End
      Begin VB.Label Label15 
         Caption         =   "是否閉卷 :"
         Height          =   252
         Left            =   4152
         TabIndex        =   78
         Top             =   1296
         Width           =   1212
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:閉卷)"
         Height          =   252
         Left            =   6192
         TabIndex        =   77
         Top             =   1296
         Width           =   1692
      End
      Begin VB.Label Label17 
         Caption         =   "案件目前准駁 :"
         Height          =   255
         Left            =   75
         TabIndex        =   76
         Top             =   1590
         Width           =   1275
      End
      Begin VB.Label Label18 
         Caption         =   "(1:准 , 2:駁)"
         Height          =   255
         Left            =   1830
         TabIndex        =   75
         Top             =   1590
         Width           =   2235
      End
      Begin VB.Label Label19 
         Caption         =   "專用權是否存在 :"
         Height          =   180
         Left            =   4152
         TabIndex        =   74
         Top             =   1596
         Width           =   1500
      End
      Begin VB.Label Label20 
         Caption         =   "(Y / N)"
         Height          =   252
         Left            =   6192
         TabIndex        =   73
         Top             =   1596
         Width           =   612
      End
      Begin VB.Label Label21 
         Caption         =   "列印備註 :"
         Height          =   255
         Left            =   75
         TabIndex        =   72
         Top             =   2190
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   75
         TabIndex        =   71
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1995
         TabIndex        =   70
         Top             =   1935
         Width           =   2745
      End
      Begin VB.Label Label4 
         Caption         =   "質權設定期間 :"
         Height          =   255
         Index           =   0
         Left            =   4770
         TabIndex        =   69
         Top             =   1890
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "－"
         Height          =   180
         Index           =   1
         Left            =   7320
         TabIndex        =   68
         Top             =   1950
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label4 
         Caption         =   "證書號 :"
         Height          =   252
         Index           =   2
         Left            =   5700
         TabIndex        =   67
         Top             =   2220
         Visible         =   0   'False
         Width           =   732
      End
      Begin VB.Label Label5 
         Caption         =   "期"
         Height          =   252
         Left            =   2952
         TabIndex        =   66
         Top             =   2508
         Width           =   252
      End
      Begin VB.Label Label24 
         Caption         =   "年"
         Height          =   252
         Left            =   1788
         TabIndex        =   65
         Top             =   2508
         Width           =   252
      End
      Begin VB.Label Label25 
         Caption         =   "定稿用 :"
         Height          =   225
         Left            =   75
         TabIndex        =   64
         Top             =   2505
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   63
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label28 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   62
         Top             =   2100
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   61
         Top             =   1860
         Width           =   1575
      End
      Begin VB.Label Label31 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   60
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label34 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   59
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label35 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   58
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label27 
         Caption         =   "對造商品類別 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   57
         Top             =   2370
         Width           =   1575
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1290
      TabIndex        =   55
      TabStop         =   0   'False
      Top             =   735
      Width           =   7605
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13414;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5970
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   2565
      Width           =   2412
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "4254;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1290
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1065
      Width           =   7605
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13414;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數/申請案號 :"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   34
      Top             =   1365
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4770
      TabIndex        =   53
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   51
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   4770
      TabIndex        =   49
      Top             =   2265
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   90
      TabIndex        =   47
      Top             =   2265
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4770
      TabIndex        =   45
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   43
      Top             =   1965
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "正商標專用期止日 :"
      Height          =   255
      Index           =   5
      Left            =   4770
      TabIndex        =   41
      Top             =   1665
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數 :"
      Height          =   255
      Index           =   4
      Left            =   90
      TabIndex        =   39
      Top             =   1665
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4770
      TabIndex        =   38
      Top             =   1365
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   255
      Index           =   2
      Left            =   4770
      TabIndex        =   35
      Top             =   435
      Width           =   915
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   32
      Top             =   1065
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   31
      Top             =   735
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   30
      Top             =   435
      Width           =   1215
   End
End
Attribute VB_Name = "frm02010401_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/28 Form2.0已修改 cmbTM05/textTM23/textCP13/textPS/textCP35
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
'2005/7/19整理
Option Explicit

' 本所案號
Public m_TM01 As String
Public m_TM02 As String
Public m_TM03 As String
Public m_TM04 As String
' 來函收文日
Public m_CP05 As String
' 收文號
Public m_CP09 As String
' 原案件性質
Public m_CP10 As String
' 原智權人員代號
Dim m_CP13 As String
Dim m_CP12 As String
' 後金
'91.6.12 CANCEL
'Dim m_CP19 As String
' 申請案公告日
Dim m_TM14 As String
'91.6.12 END
' 原移轉申請人代號
Dim m_CP56 As String
'Add By Sindy 2013/1/11
Dim m_CP89 As String
Dim m_CP90 As String
Dim m_CP91 As String
Dim m_CP92 As String
'2013/1/11 End
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Public m_TM10 As String
' 發證日
Public m_TM20 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 原申請人代號
Public m_TM23 As String
'Add By Sindy 2013/1/11
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String
'2013/1/11 End
' 申請國家的延展年度
Dim m_NA14 As Integer

'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String

Public m_txtTM14 As String
'edit by nick 2004/10/29
'Dim m_txtTMBM07_1 As String
'Dim m_txtTMBM07_2 As String
Public m_txtTMBM07_1 As String
Public m_txtTMBM07_2 As String
Dim m_txtTM16S As String
Dim m_txtTM17 As String
Public m_blnNotFirst As Boolean
'Add By Cheng 2002/12/11
'Dim m_blnClkChgButton As Boolean '是否有按變更事項鈕
Public m_blnClkChgButton As Boolean '是否有按變更事項鈕 'Modify By Sindy 2012/2/6 Dim->Public
'Add By Cheng 2003/11/18
Public m_TM11 As String '申請日
'Add By Cheng 2003/12/08
Dim m_blnReceiveFirst As Boolean  '是否已收第一期註冊費
'add by nick 2004/10/19
Public m_CP14 As String
'2005/11/9 ADD BY SONIA
Dim m_strLanguage As String '定稿語文
'add by nickc 2008/04/07
Dim m_TM44 As String
' 原收文號法定期限
Dim m_CP07 As String    '2010/3/2 ADD BY SONIA
'Add By Sindy 2010/12/28 檢查是否已經有商品及服務
Public ChkTG As Boolean
Dim BolPrintCaseCheck As Boolean 'Add By Sindy 2012/4/16
Dim BolPrintLetterDemand As Boolean 'Add By Sindy 2012/4/16
Dim m_NP08 As String '下一程序本所期限 'Add By Sindy 2012/4/20
Dim m_CP47 As String 'Add By Sindy 2012/10/18
Dim bolMod As Boolean 'Added by Lydia 2016/07/19 是否有變更事項
'Added by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
Dim bolA1kdataMail As Boolean '發催款函(Outlook)
Dim m_ULD02 As String   '更新定稿日期
'Modified by Lydia 2017/04/06 請款單的請款對象,可能和代理不一致,改設變數
'Dim m_AC2470 As String  '定稿加印催款單PDF
Dim m_rA1k28 As String  '請款單的請款對象
Dim m_rSpec As String  '特定代理人的mail內文不同
'end 2017/04/06
Dim strCP09 As String   '新增的C類收文號
Dim strCP10 As String   '新增的C類收文案件性質

'Added by Morgan 2017/4/12 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/12
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Public strLD18 As String 'Add By Sindy 2019/12/19 信函總收文號
Dim m_CP16 As String 'Added by Lydia 2020/03/19 原收文之費用
Dim m_CP43 As String 'Added by Lydia 2024/04/01 原收文之相關收文號

'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm02010401_3.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm02010401_3
   Unload frm02010401_2
   Unload frm02010401_1
   Unload Me
End Sub

Private Sub cmdMod_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
    'Add By Cheng 2002/12/11
    'Modify By Sindy 2012/2/6 Mark
'    m_blnClkChgButton = True

   bolMod = False 'Added by Lydia 2016/07/19
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   'edit by nickc 200/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "無變更事項記錄"
      strTit = "資料檢核"
      'Modified by Lydia 2016/07/19 +判斷
      If cmdMod.Visible = True Then
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      End If
      
      GoTo EXITSUB
   End If
   
   bolMod = True 'Added by Lydia 2016/07/19
   rsTmp.Close
   DisplayNextForm
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdok_Click()
Dim rsA As New ADODB.Recordset
   
   'Add by Morgan 2003/11/21
   BolPrintCaseCheck = CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   '---end
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          '2011/11/8 modify by sonia TF子案不可結餘故加傳本所案號
          'Pub_EndModCashMsg m_TM10
          Pub_EndModCashMsg m_TM10, m_TM01, m_TM02, m_TM03, m_TM04
          
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      'Modify By Cheng 2002/11/07
'      'OnSaveData
      'Added by Lydia 2016/07/19 延展核准在存檔時,直接將變更事項確定全部核准
      'Modified by Lydia 2017/07/28 +301變更核准,比照延展核准辦理
      If m_CP10 = "102" Or m_CP10 = "301" Then
         Call cmdMod_Click
         If bolMod Then '有變更事項
            If frm02010401_5.Get102_Approve = False Then
               Screen.MousePointer = vbDefault: Exit Sub
            End If
         End If
      End If
      'end 2016/07/19
      
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      'Add By Cheng 2002/11/08
      ' 列印定稿
      If textPrint <> "N" Then
          BolPrintLetterDemand = True 'Add By Sindy 2012/4/16
          PrintLetter
      Else
          BolPrintLetterDemand = False 'Add By Sindy 2012/4/16
      End If
      
      'Add By Sindy 2012/4/16 列印帳款未結清案件資料
      If BolPrintCaseCheck = True And BolPrintLetterDemand = False Then
          Call GetPrintCaseCheck(m_CP09)
      End If
      '2012/4/16 End
      Call PUB_ChkTemporaryReceipts(m_TM01, m_TM02, m_TM03, m_TM04) 'Add By Sindy 2014/5/28 檢查是否有暫收款
      
      'Added by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
      '為了延緩出定稿,更新定稿日期
      If m_ULD02 <> "" Then
         'Modified by Lydia 2017/04/24 改成Function
         'Call PUB_UpdateET07LD0216("1", m_CP09, m_TM01, m_TM02, m_TM03, m_TM04, "03", m_ULD02)
         If PUB_UpdateET07LD0216("1", m_CP09, m_TM01, m_TM02, m_TM03, m_TM04, "03", m_ULD02) = False Then
         End If
         'end 2017/04/24
      End If
      '發催款函
      If bolA1kdataMail = True Then
         'Modified by Lydia 2017/02/18 預設都附催款,並區分是否為特定客戶(寄紙本)
         'Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strCP09, strCP10, m_AC2470)
         'Modified by Lydia 2017/04/06 區分請款對象
         'Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strCP09, strCP10, m_TM44, IIf(m_AC2470 <> "", "Y", "N"))
         'Added by Lydia 2017/11/01 因為郵件預設收件人為基本檔之代理人,若欠款之對象與TM44不同時,彈訊息提醒即可
                                   'ex. T-156008現在TM44=Y5338100,106/10/24 核准-延展CA6066488,判斷同案件98年有Y51318000的欠款(催款單的請款對象),所以產生D類收款寄證和發MAIL; 發MAIL套用模組預設抓TM44為收件人,然後發信Y5338100造成對方的疑問。
         If m_rA1k28 <> m_TM44 Then
            MsgBox "欠款請款單之請款對象與現在FC代理人不同, 請自行注意欲催款對象！！", vbCritical, "收款寄證"
         End If
         'end 2017/11/01
         'Added by Lydia 2023/03/23
         If Dir(GetMyDocPath & "\收款寄證", vbDirectory) = "" Then
            MkDir GetMyDocPath & "\收款寄證"
         End If
         'end 2023/03/23
         Call PUB_SendA1kdataMail(Me, m_TM01, m_TM02, m_TM03, m_TM04, strCP09, strCP10, m_rA1k28, m_rSpec)
      End If
      'end 2016/12/22
              
      'Add By Sindy 2023/3/8 T091286 移轉(501)或變更(301)申請人自請撤回須還原申請人
      '於自請撤回核准輸入時彈提醒修改申請人資料
      If m_CP10 = "306" Then
         strSql = "Select CP09,CP10 From CaseProgress Where CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "')"
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If rsA.Fields("CP10") = "501" Or rsA.Fields("CP10") = "301" Then
               strExc(10) = rsA.Fields("CP09")
               If rsA.Fields("CP10") = "301" Then
                  '有變更申請人
                  strSql = "Select CE01 From ChangeEvent Where CE01='" & strExc(10) & "'" & _
                           " AND CE04||CE05||CE06||CE07||CE08 is not null"
                  rsA.CursorLocation = adUseClient
                  rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount = 0 Then
                     strExc(10) = ""
                  End If
               End If
               If strExc(10) <> "" Then
                  MsgBox "請注意，修改申請人資料！", vbCritical, "自請撤回核准"
               End If
            End If
         End If
      End If
      '2023/3/8 END
      
      'Add By Cheng 2002/01/15
      m_txtTM14 = Me.textTM14.Text
      m_txtTMBM07_1 = Me.textTMBM07_1.Text
      m_txtTMBM07_2 = Me.textTMBM07_2.Text
      'add by nick 2004/10/29
      frm02010401_1.m_txtTMBM07_1 = m_txtTMBM07_1
      frm02010401_1.m_txtTMBM07_2 = m_txtTMBM07_2
      frm02010401_1.m_txtTM14 = m_txtTM14
      frm02010401_1.m_blnNotFirst = True
      'Modify By Cheng 2002/07/22
      '取消記錄此次輸入的資料
'      m_txtTM16S = Me.textTM16S.Text
'      m_txtTM17 = Me.textTM17.Text
      m_blnNotFirst = True
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Unload Me
      Unload frm02010401_3
      Unload frm02010401_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010401_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/20 電子公文
      'frm02010401_1.Show
      ElseIf m_DocNo <> "" Then
         'Unload frm02010401_1 要保留公告日,公報卷期
         'Added by Morgan 2022/3/18 電子公文變數要清除，否則會誤判
         frm02010401_1.m_DocNo = ""
         frm02010401_1.m_AppNo = ""
         frm02010401_1.m_RegNo = ""
         frm02010401_1.textTM15 = ""
         frm02010401_1.textTM12 = ""
         'end 2022/3/18
         Unload Me
         frm02010412.GoNext
      Else
         frm02010401_1.Show
         Unload Me
      End If
      'end 2017/4/20
   End If
End Sub

Private Sub Command2_Click()
   frm03010303_04.Hide
   Set frm03010303_04.UpForm = Me
   frm03010303_04.TGKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   frm03010303_04.AllClass = textTM09.Text
   frm03010303_04.cmdok(2).Visible = True
   
'   If m_EditMode <> 1 And m_EditMode <> 2 Then
'       frm03010303_04.Label2.Visible = False
'       frm03010303_04.cmdok(0).Visible = False
'       frm03010303_04.cmdok(2).Visible = False
'       frm03010303_04.cmd.Visible = False
'       frm03010303_04.cmd2.Visible = False
'       frm03010303_04.txt2(0).Visible = False
'       frm03010303_04.txt2(1).Visible = False
'       frm03010303_04.txt2(2).Visible = False
'       frm03010303_04.txt2(3).Visible = False
'       frm03010303_04.Line1.Visible = False
'   End If
   If textTM09 <> "" Then  '2010/5/10 MODIFY BY SONIA 有商品類別才可進入 T-113511團體標章
      Me.Hide
      frm03010303_04.QueryData
      frm03010303_04.Show vbModal 'Modify By Sindy 2009/09/17 改為強制回應表單
   '2010/5/10 ADD BY SONIA
   Else
      MsgBox ("無商品類別，不可使用此按鈕 !")
   End If
   '2010/5/10 END
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
   textTM22S.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
  
   MoveFormToCenter Me
   'Add By Cheng 2002/12/11
   'm_blnClkChgButton = False
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010401_3.m_strIR01
   m_strIR02 = frm02010401_3.m_strIR02
   m_strIR03 = frm02010401_3.m_strIR03
   m_strIR04 = frm02010401_3.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
      m_CP09 = Empty
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
      ' 來函收文日
      Case 4: m_CP05 = strData
      ' 收文號
      Case 5: m_CP09 = strData
             'add by nickc 2005/08/04
            strSql = "SELECT * FROM ChangeEvent " & _
                     "WHERE CE01 = '" & m_CP09 & "' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount <= 0 Then
               m_blnClkChgButton = True
            Else
               m_blnClkChgButton = False
            End If
            rsTmp.Close
      'add by nick 2004/10/29
      Case 6: m_txtTMBM07_1 = strData
      Case 7: m_txtTMBM07_2 = strData
      Case 8: m_txtTM14 = strData
   End Select
End Sub

Public Sub QueryData()
Dim strSql As String
Dim strSub As String
Dim rsTmp As New ADODB.Recordset
Dim rsSub As ADODB.Recordset
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   m_CP10 = Empty
   m_CP56 = Empty
   'Add By Sindy 2013/1/11
   m_CP89 = Empty
   m_CP90 = Empty
   m_CP91 = Empty
   m_CP92 = Empty
   '2013/1/11 End
   m_TM08 = Empty
   m_TM10 = Empty
   m_TM11 = Empty
   m_TM14 = Empty
   m_TM21 = Empty
   m_TM22 = Empty
   m_TM23 = Empty
   m_CP07 = Empty   '2010/3/2 ADD BY SONIA
   'Add By Sindy 2013/1/11
   m_TM78 = Empty
   m_TM79 = Empty
   m_TM80 = Empty
   m_TM81 = Empty
   '2013/1/11 End
    m_blnReceiveFirst = False
   ' 來函收文日
   textCP05S = m_CP05
    'Modify By Cheng 2004/02/04
    Select Case m_TM01
    Case "T", "FCT", "TF"
        ' 取得商標基本檔的相關項目
        strSql = "SELECT * FROM TradeMark " & _
                 "WHERE TM01 = '" & m_TM01 & "' AND " & _
                       "TM02 = '" & m_TM02 & "' AND " & _
                       "TM03 = '" & m_TM03 & "' AND " & _
                       "TM04 = '" & m_TM04 & "'"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
        If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
           'add by nickc 2008/04/07
           m_TM44 = CheckStr(rsTmp.Fields("TM44"))
           
           textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
           ' 申請國家
           'Add By Cheng 2002/07/17
           m_NA14 = Empty
           If IsNull(rsTmp.Fields("TM10")) = False Then
              m_TM10 = rsTmp.Fields("TM10")
              textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
              m_NA14 = GetNationExtentYear(rsTmp.Fields("TM10"))
           End If
           
            'Add By Sindy 2012/6/18 TF案申請國家為英國201時,卷期鎖住
            textTMBM07_1.Enabled = True
            If m_TM01 = "TF" And m_TM10 = "201" Then
                textTMBM07_1.Enabled = False
            End If
            '2012/6/18 End
            
           ' 90.07.19 modify 專用期限以國別判斷使用民國或西元日期
           If m_TM10 < "010" Then
              textTM21.MaxLength = 7
              textTM22.MaxLength = 7
              'Add By Cheng 2002/06/12
              Me.textCP53.MaxLength = 7
              Me.textCP54.MaxLength = 7
           Else
              textTM21.MaxLength = 8
              textTM22.MaxLength = 8
              'Add By Cheng 2002/06/12
              Me.textCP53.MaxLength = 8
              Me.textCP54.MaxLength = 8
           End If
             '申請日
             m_TM11 = "" & rsTmp.Fields("TM11").Value
           ' 審定號
           If IsNull(rsTmp.Fields("TM15")) = False Then
              textTM15 = rsTmp.Fields("TM15")
           End If
           'Add By Sindy 2010/12/31
           ' 審定號數
           If IsNull(rsTmp.Fields("TM15")) = False Then
               textTM12 = rsTmp.Fields("TM15")
           Else
               ' 申請案號
               If IsNull(rsTmp.Fields("TM12")) = False Then
                  textTM12 = rsTmp.Fields("TM12")
               End If
           End If
           '2010/12/31 End
           ' 商標名稱(中)
           If IsNull(rsTmp.Fields("TM05")) = False Then
              cmbTM05.AddItem rsTmp.Fields("TM05")
           End If
           ' 商標名稱(英)
           If IsNull(rsTmp.Fields("TM06")) = False Then
              cmbTM05.AddItem rsTmp.Fields("TM06")
           End If
           ' 商標名稱(日)
           If IsNull(rsTmp.Fields("TM07")) = False Then
              cmbTM05.AddItem rsTmp.Fields("TM07")
           End If
           ' 顯示商標名稱
           If cmbTM05.ListCount > 0 Then
              cmbTM05.ListIndex = 0
           End If
           ' 商標種類
           If IsNull(rsTmp.Fields("TM08")) = False Then
              m_TM08 = rsTmp.Fields("TM08")
              If m_TM10 < "010" Then
                 textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
              Else
                 textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
              End If
           End If
           ' 商品類別
           If IsNull(rsTmp.Fields("TM09")) = False Then
              textTM09 = rsTmp.Fields("TM09")
           End If
           ' 發證日
           m_TM20 = Empty
           If IsNull(rsTmp.Fields("TM20")) = False Then
              If rsTmp.Fields("TM20") <> "0" Then
                 m_TM20 = rsTmp.Fields("TM20")
              End If
           End If
           ' 申請人
           If IsNull(rsTmp.Fields("TM23")) = False Then
              m_TM23 = rsTmp.Fields("TM23")
              textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
           End If
            'Add By Sindy 2013/1/11
            If IsNull(rsTmp.Fields("TM78")) = False Then
               m_TM78 = rsTmp.Fields("TM78")
            End If
            If IsNull(rsTmp.Fields("TM79")) = False Then
               m_TM79 = rsTmp.Fields("TM79")
            End If
            If IsNull(rsTmp.Fields("TM80")) = False Then
               m_TM80 = rsTmp.Fields("TM80")
            End If
            If IsNull(rsTmp.Fields("TM81")) = False Then
               m_TM81 = rsTmp.Fields("TM81")
            End If
            '2013/1/11 End
           ' 正商標號數
           If IsNull(rsTmp.Fields("TM27")) = False Then
              textTM27 = rsTmp.Fields("TM27")
           End If
           ' 彼所案號
           If IsNull(rsTmp.Fields("TM45")) = False Then
              textTM45 = rsTmp.Fields("TM45")
           End If
           'Add By Cheng 2002/07/22
           '顯示目前准駁
           Me.textTM16S.Text = "" & rsTmp.Fields("TM16").Value
           
           'add by nickc 2006/11/17
           textPrint = CheckStr(rsTmp.Fields("tm77"))
           ' 正商標專用期止日
           Set rsSub = New ADODB.Recordset
           strSub = "SELECT * FROM TradeMark " & _
                    "WHERE TM15 = '" & textTM27 & "' AND " & _
                          "TM10 = '" & m_TM10 & "' "
           rsSub.CursorLocation = adUseClient
           rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
           If rsSub.RecordCount > 0 Then
              rsSub.MoveFirst
              If IsNull(rsSub.Fields("TM22")) = False Then
                 textTM22S = rsSub.Fields("TM22")
              End If
           End If
           rsSub.Close
           Set rsSub = Nothing
           ' 審定號
           If IsNull(rsTmp.Fields("TM15")) = False Then
              textTM15 = rsTmp.Fields("TM15")
           End If
           ' 公告日  90/06/20不抓基本檔公告日
           'If IsNull(rsTmp.Fields("TM14")) = False Then
           '   textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
           'End If
           '91.6.12 MODIFY BY SONIA
           If IsNull(rsTmp.Fields("TM14")) = False Then
              m_TM14 = TAIWANDATE(rsTmp.Fields("TM14"))
           End If
           '91.6.12 END
           ' 專用權是否存在
           If IsNull(rsTmp.Fields("TM17")) = False Then
              textTM17 = rsTmp.Fields("TM17")
           End If
           ' 專用期限 (起)
           If IsNull(rsTmp.Fields("TM21")) = False Then
              m_TM21 = rsTmp.Fields("TM21")
              ' 90.07.19 modify (專用期限以國別判斷使用民國或西元日期)
              If m_TM10 < "010" Then
                 textTM21 = TAIWANDATE(rsTmp.Fields("TM21"))
              Else
                 textTM21 = DBDATE(rsTmp.Fields("TM21"))
              End If
           End If
           ' 專用期限 (止)
           If IsNull(rsTmp.Fields("TM22")) = False Then
              m_TM22 = rsTmp.Fields("TM22")
              ' 90.07.19 modify (專用期限以國別判斷使用民國或西元日期)
              If m_TM10 < "010" Then
                 textTM22 = TAIWANDATE(rsTmp.Fields("TM22"))
              Else
                 textTM22 = DBDATE(rsTmp.Fields("TM22"))
              End If
           End If
           ' 是否閉卷
           If IsNull(rsTmp.Fields("TM29")) = False Then
              textTM29 = rsTmp.Fields("TM29")
           End If
        End If
        rsTmp.Close
    Case "TC"
        ' 取得服務業務基本檔的相關項目
        strSql = "SELECT * FROM Servicepractice " & _
                 "WHERE SP01 = '" & m_TM01 & "' AND " & _
                       "SP02 = '" & m_TM02 & "' AND " & _
                       "SP03 = '" & m_TM03 & "' AND " & _
                       "SP04 = '" & m_TM04 & "'"
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
        If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
           textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
           ' 申請國家
           'Add By Cheng 2002/07/17
           m_NA14 = Empty
           If IsNull(rsTmp.Fields("SP09")) = False Then
              m_TM10 = rsTmp.Fields("SP09")
              textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
              m_NA14 = GetNationExtentYear(rsTmp.Fields("SP09"))
           End If
           ' 90.07.19 modify 專用期限以國別判斷使用民國或西元日期
           If m_TM10 < "010" Then
              textTM21.MaxLength = 7
              textTM22.MaxLength = 7
              'Add By Cheng 2002/06/12
              Me.textCP53.MaxLength = 7
              Me.textCP54.MaxLength = 7
           Else
              textTM21.MaxLength = 8
              textTM22.MaxLength = 8
              'Add By Cheng 2002/06/12
              Me.textCP53.MaxLength = 8
              Me.textCP54.MaxLength = 8
           End If
             '申請日
             m_TM11 = "" & rsTmp.Fields("SP10").Value
'           ' 審定號
'           If IsNull(rsTmp.Fields("TM15")) = False Then
'              textTM15 = rsTmp.Fields("TM15")
'           End If
           'Add By Sindy 2010/12/31
           ' 申請案號
           If IsNull(rsTmp.Fields("SP11")) = False Then
              textTM12 = rsTmp.Fields("SP11")
           End If
           '2010/12/31 End
           
           'Add By Sindy 2019/12/25
            ' FC代理人
            m_TM44 = Empty
            If IsNull(rsTmp.Fields("SP26")) = False Then
               m_TM44 = rsTmp.Fields("SP26")
            End If
            '2019/12/25 END
            
           ' 商標名稱(中)
           If IsNull(rsTmp.Fields("SP05")) = False Then
              cmbTM05.AddItem rsTmp.Fields("SP05")
           End If
           ' 商標名稱(英)
           If IsNull(rsTmp.Fields("SP06")) = False Then
              cmbTM05.AddItem rsTmp.Fields("SP06")
           End If
           ' 商標名稱(日)
           If IsNull(rsTmp.Fields("SP07")) = False Then
              cmbTM05.AddItem rsTmp.Fields("SP07")
           End If
           ' 顯示商標名稱
           If cmbTM05.ListCount > 0 Then
              cmbTM05.ListIndex = 0
           End If
           ' 申請人
           If IsNull(rsTmp.Fields("SP08")) = False Then
              m_TM23 = rsTmp.Fields("SP08")
              textTM23 = GetCustomerName(rsTmp.Fields("SP08"))
           End If
           ' 是否閉卷
           If IsNull(rsTmp.Fields("SP15")) = False Then
              textTM29 = rsTmp.Fields("SP15")
           End If
            'add by nickc 2006/11/17
            textPrint = CheckStr(rsTmp.Fields("SP72"))
        End If
        rsTmp.Close
        Me.Label4(2).Visible = True
        Me.txtNote.Visible = True
        Me.txtNote.Enabled = True
    End Select
   
   ' 取得案件進度檔檔案中欄位
   '2010/9/28 MODIFY BY SONIA 判斷原承辦人若離職改抓P2001
   'strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   strSql = "SELECT * FROM CaseProgress,STAFF WHERE CP09 = '" & m_CP09 & "' AND CP14=ST01(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
      End If
'      ' 收文號
'      If IsNull(rsTmp.Fields("CP09")) = False Then
'         textCP09 = rsTmp.Fields("CP09")
'      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
         'Add By Cheng 2002/06/13
         '若案件性質為授權
         If m_CP10 = "502" Then
            Me.Label4(0).Visible = True
            Me.Label4(1).Visible = True
            Me.Label4(0).Caption = "授權期間："
            Me.textCP53.Visible = True
            Me.textCP54.Visible = True
            If m_TM10 < "010" Then
               Me.textCP53.MaxLength = 7
               Me.textCP54.MaxLength = 7
               Me.textCP53.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP53"))
               Me.textCP54.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP54"))
            Else
               Me.textCP53.MaxLength = 8
               Me.textCP54.MaxLength = 8
               Me.textCP53.Text = "" & DBDATE("" & rsTmp.Fields("CP53"))
               Me.textCP54.Text = "" & DBDATE("" & rsTmp.Fields("CP54"))
            End If
         '若案件性質為設定質權時
         ElseIf m_CP10 = "506" Then
            Me.Label4(0).Visible = True
            Me.Label4(1).Visible = True
            Me.Label4(0).Caption = "質權設定期間："
            Me.textCP53.Visible = True
            Me.textCP54.Visible = True
            If m_TM10 < "010" Then
               Me.textCP53.MaxLength = 7
               Me.textCP54.MaxLength = 7
               Me.textCP53.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP53"))
               Me.textCP54.Text = "" & TAIWANDATE("" & rsTmp.Fields("CP54"))
            Else
               Me.textCP53.MaxLength = 8
               Me.textCP54.MaxLength = 8
               Me.textCP53.Text = "" & DBDATE("" & rsTmp.Fields("CP53"))
               Me.textCP54.Text = "" & DBDATE("" & rsTmp.Fields("CP54"))
            End If
         'add by sonia 2019/1/4 若案件性質為723出具同意書時 T-214606, 預設對方案件號數CP30
         ElseIf m_CP10 = "723" Then
            Label1(19).Caption = "　　　　　進度備註 :"
            textCP64 = "對方以「對方商標名稱」（第對方商品類別類）商標申請註冊（申請案號：" & rsTmp.Fields("CP30") & "）"
            textCP64.Width = 3300
            textCP64.MaxLength = 2000
         Else
            textCP64.Width = 2290
            textCP64.MaxLength = 8
         'end 2019/1/4
         End If
      End If
      ' 智權人員
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區   91.08.22 nickc
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
          m_CP12 = rsTmp.Fields("cp12")
      End If
      'add by nick 2004/10/19
      '承辦人
      m_CP14 = Empty
      If IsNull(rsTmp.Fields("cp14")) = False Then
         m_CP14 = "" & rsTmp.Fields("cp14")
         '2010/9/28 ADD BY SONIA 判斷原承辦人若離職改抓P2001
         If "" & rsTmp.Fields("ST04") = "2" Then m_CP14 = "P2001"
         '2010/9/28 END
      End If
      
      m_CP16 = "" & rsTmp.Fields("CP16") 'Added by Lydia 2020/03/19 原收文之費用
      m_CP43 = "" & rsTmp.Fields("CP43") 'Added by Lydia 2024/04/01 原收文之相關收文號
      
      ' 後金
      '91.6.12 MODIFY BY SONIA
      'm_CP19 = Empty
      'If IsNull(rsTmp.Fields("CP19")) = False Then
      '   m_CP19 = rsTmp.Fields("CP19")
      'End If
      '91.6.12 END
      ' 核准通知日
      'Add/Modify By Cheng 2002/04/30
      '若申請國家為台灣時, 核准通知日隱藏且Disable; 其他則不隱藏且Enable, 且不預設任何值
      If m_TM10 = 台灣國家代號 Then
         Me.textCP25.Text = frm02010401_1.textCP05.Text
         Me.textCP25.Visible = False
         Me.Label7.Visible = False
      '91.6.12 MODIFY BY SONIA
      Else
         If IsNull(rsTmp.Fields("CP24")) = False And rsTmp.Fields("CP24") = "1" And IsNull(rsTmp.Fields("CP25")) = False Then
            textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
         End If
      '91.6.12 END
      End If
      '申請案同時帶出公告日
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If rsTmp.Fields("CP10") = "101" And IsNull(m_TM14) = False Then
      'modify by sonia 2024/9/4 +104領土延伸
      If (rsTmp.Fields("CP10") = "101" Or rsTmp.Fields("CP10") = "308" Or rsTmp.Fields("CP10") = "104") And IsNull(m_TM14) = False Then
         textTM14 = TAIWANDATE(m_TM14)
      End If
      'add by sonia 2016/11/17 台灣申請案不顯示公告日以免誤輸影響發證時之專用期間T-203519(桂英)
      If m_TM10 = 台灣國家代號 And rsTmp.Fields("CP10") = "101" Then
         textTM14 = "": m_txtTM14 = ""
         Me.textTM14.Visible = False
         Me.Label10.Visible = False
      End If
         'end 2016/11/17
'      If IsNull(rsTmp.Fields("CP25")) = False Then
'         textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
'      End If
      
      ' 審查委員
      If IsNull(rsTmp.Fields("CP35")) = False Then
         textCP35 = rsTmp.Fields("CP35")
      End If
      ' 移轉申請人代號
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      'Add By Sindy 2013/1/11
      If IsNull(rsTmp.Fields("CP89")) = False Then
         m_CP89 = rsTmp.Fields("CP89")
      End If
      If IsNull(rsTmp.Fields("CP90")) = False Then
         m_CP90 = rsTmp.Fields("CP90")
      End If
      If IsNull(rsTmp.Fields("CP91")) = False Then
         m_CP91 = rsTmp.Fields("CP91")
      End If
      If IsNull(rsTmp.Fields("CP92")) = False Then
         m_CP92 = rsTmp.Fields("CP92")
      End If
      '2013/1/11 End
      ' 案件性質為延展者, 帶出授權期間於專用期限欄位
      If m_CP10 = "102" Then
         If IsNull(rsTmp.Fields("CP53")) = False Then
            ' 90.07.19 modify (專用期限以國別判斷使用民國或西元日期)
            If m_TM10 < "010" Then
               textTM21 = TAIWANDATE(rsTmp.Fields("CP53"))
            Else
               textTM21 = DBDATE(rsTmp.Fields("CP53"))
            End If
         End If
         If IsNull(rsTmp.Fields("CP54")) = False Then
            ' 90.07.19 modify (專用期限以國別判斷使用民國或西元日期)
            If m_TM10 < "010" Then
               textTM22 = TAIWANDATE(rsTmp.Fields("CP54"))
            Else
               textTM22 = DBDATE(rsTmp.Fields("CP54"))
            End If
         End If
      End If
      '2010/3/2 ADD BY SONIA
      If m_CP10 = "109" Then
         If IsNull(rsTmp.Fields("CP07")) = False Then
            m_CP07 = rsTmp.Fields("CP07")
         End If
      End If
      '2010/3/2 END
      'Add By Sindy 2012/10/18 代理人提申日
      m_CP47 = ""
      If IsNull(rsTmp.Fields("CP47")) = False Then
         m_CP47 = rsTmp.Fields("CP47")
      End If
      '2012/10/18 End
   End If
   rsTmp.Close
   
   ' 案件性質為申請, 申請國家為台灣時, 以審定號數+商標種類代號抓商標公報檔, 帶出卷期
   'If m_CP10 = "101" And m_TM10 < "010" Then
   '   strSQL = "SELECT * FROM TMBULLETIN " & _
   '            "WHERE TMBM01 = '" & textTM15 & "' AND " & _
   '                  "TMBM02 = '" & m_TM08 & "' "
   '   rsTmp.CursorLocation = adUseClient
   '   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   '   If rsTmp.RecordCount > 0 Then
   '      rsTmp.MoveFirst
   '      If IsNull(rsTmp.Fields("TMBM07")) = False Then
   '         textTMBM07_1 = Mid(rsTmp.Fields("TMBM07"), 1, 2)
   '         textTMBM07_2 = Mid(rsTmp.Fields("TMBM07"), 3, 3)
   '      End If
   '   End If
   '   rsTmp.Close
   'End If
   
   '2010/12/22 ADD BY SONIA 台灣案案件性質為申請且為改變原處分時, 清除商標基本檔的審定號,否則在發證前會以為是核准審定號
   If m_CP10 = "101" And m_TM10 = "000" And frm02010401_3.GetSelectResult() = "2" Then
      textTM15 = ""
   End If
   '2010/12/22 END
   ' 案件性質為延展時, 才可輸入專用期限
   'Modified by Lydia 2017/07/28 +301變更核准,比照延展核准辦理
   If m_CP10 = "102" Or m_CP10 = "301" Then
      EnableTextBox textTM21, True
      EnableTextBox textTM22, True
      cmdMod.Visible = False 'Added by Lydia 2016/07/19
   Else
      EnableTextBox textTM21, False
      EnableTextBox textTM22, False
      cmdMod.Visible = True  'Added by Lydia 2016/07/19
   End If
   ' 案件性質為自請撤回或自請撤銷時, 才可輸入是否閉卷的欄位
   'Modify By Sindy 2016/8/10 + 626.註銷
   If m_CP10 = "306" Or m_CP10 = "307" Or m_CP10 = "626" Then
      EnableTextBox textTM29, True
      '93.9.3 ADD BY SONIA
      'Modify By Sindy 2016/8/10 + 626.註銷
      If m_CP10 = "307" Or m_CP10 = "626" Then
         textTM29 = "Y"
      End If
      '93.9.3 END
      '2011/10/4 ADD BY SONIA 申請程序的自請撤回也要閉卷 T-169065(自撤發文後才改相關總收文號)
      If m_CP10 = "306" Then
         '2011/11/7 MODIFY BY SONIA 再加601異議,603評定,605廢止
         '2012/12/19 MODIFY BY SONIA 601異議,603評定,605廢止要再判斷卷宗性質 FCT-032571,並取消805
         'strSql = "SELECT C2.CP09 FROM CASEPROGRESS C1,CASEPROGRESS C2 " & _
                  "WHERE C1.CP09 = '" & m_CP09 & "' AND C1.CP43=C2.CP09(+) AND C2.CP10 IN ('101','805','601','603','605') AND C2.CP09 IS NOT NULL "
         'modify by sonia 2017/8/30 +623
         strSql = "SELECT C2.CP09,C2.CP10,TM28 FROM CASEPROGRESS C1,CASEPROGRESS C2,TRADEMARK " & _
                  "WHERE C1.CP09 = '" & m_CP09 & "' AND C1.CP43=C2.CP09(+) AND C2.CP10 IN ('101','601','603','605','623') AND C2.CP09 IS NOT NULL " & _
                  "AND C1.CP01=TM01(+) AND C1.CP02=TM02(+) AND C1.CP03=TM03(+) AND C1.CP04=TM04(+) "
         '2012/12/19 END
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            '2012/12/19 MODIFY BY SONIA FCT-032571
            'textTM29 = "Y"
            Select Case rsTmp.Fields(1)
               Case "101"
                  textTM29 = "Y"
               Case Else
                  If rsTmp.Fields(2) <> "1" Then
                     textTM29 = "Y"
                  Else
                     EnableTextBox textTM29, False
                  End If
            End Select
            '2012/12/19 END
         End If
         rsTmp.Close
      End If
      '2011/10/4 END
   Else
      EnableTextBox textTM29, False
   End If

   Set rsTmp = Nothing
   
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   
   If m_TM10 < "010" Then
      
      If textCP08 = "" Then
         textCP08 = "（" & strTmp & "）智商字第號"
      End If
      m_strNumBegin = "商"
      m_strNumEnd = "字"
      
      'Added by Morgan 2017/4/13 電子公文
      If m_DocNo <> "" Then
         If m_DocWord <> "" Then
            textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
            '審查委員
            textCP08_LostFocus
         Else
            textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
         End If
      End If
      'end 2017/4/13
   End If
      
   'Add By Cheng 2002/01/15
   If m_blnNotFirst Then
      If Me.textTM14.Text = "" Then Me.textTM14.Text = m_txtTM14
      Me.textTMBM07_1.Text = m_txtTMBM07_1
      Me.textTMBM07_2.Text = m_txtTMBM07_2
      'Modify By Cheng 2002/07/22
      '取消顯示前次輸入的資料
'      Me.textTM16S.Text = m_txtTM16S
      'Modify By Cheng 2002/04/29
      '避免與讀取資料顯示欄位值相衝突
'      Me.textTM17.Text = m_txtTM17
   End If
   'Modify By Cheng 2002/07/22
'   'Add By Cheng 2002/07/11
'   '若案件性質為申請 "101"
'   If m_CP10 = "101" Then
'      '是否更新基本檔目前准駁預設為"Y"
'      Me.textTM16S.Text = "Y"
'   '其他案件性質
'   Else
'      '是否更新基本檔目前准駁預設為"N"
'      Me.textTM16S.Text = "N"
'   End If
'   'Add By Cheng 2002/07/16
'   '若申請國家為大陸, 且前一畫面的結果欄為"1"
'   If m_TM01 = "020" And frm02010401_3.textResult.Text = "1" Then
'      Me.textTM16S.Text = "N"
'      Me.textTM16S.Enabled = False
'   End If
'   '若前一畫面的結果欄為"3"
'   If frm02010401_3.textResult.Text = "3" Then
'      Me.textTM16S.Text = "Y"
'   End If
   'Add By Cheng 2002/07/22
   '若為商申案, 則預設為"1"(准)
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If (m_CP10 = "101" Or m_CP10 = "308") Then
      Me.textTM16S.Text = "1"
   End If
    Select Case m_TM01
    Case "T", "FCT", "TF"
        'Add By SONIA 91.11.25
        '若為國內案, 則預設列印備註
        'edit by nick 2004/12/23 分割與申請做相同的事情
        'If m_TM10 = "000" And m_CP10 <> "101" Then
        If m_TM10 = "000" And (m_CP10 <> "101" And m_CP10 <> "308") Then
            'Modify By Cheng 2004/03/15
'            Me.textPS.Text = "附件：註冊證正本。"
            Me.textPS.Text = "附件：智慧財產局函乙紙。"
            'End
            '依不同案件性質預設列印備註
            Select Case m_CP10
'            'Add By Cheng 2004/02/09
'            Case "102", "501" '延展, 移轉
'                Me.textPS.Text = "附件：經濟部智慧財產局函乙紙。"
'            'End
            'Add By Cheng 2004/03/18
            Case "103" '補換發證書(補證)
                Me.textPS.Text = "附件：註冊證正本乙紙。"
            'Add By Cheng 2003/05/30
            Case "304" '申請英文證明
                Me.textPS.Text = "附件：英文註冊證明書正本。"
            Case "309" '申請中文證明
                Me.textPS.Text = "附件：註冊證影本一份。"
'            Case "716" '第二期註冊費
'                Me.textPS.Text = "附件：智慧財產局函乙紙。"
            End Select
        End If
        If m_TM10 = "020" And m_CP10 = "103" Then
            Me.textPS.Text = "附件：註冊證正本。"
        'edit by nick 2004/12/23 分割與申請做相同的事情
        'ElseIf m_TM10 = "020" And m_CP10 <> "101" Then
        ElseIf m_TM10 = "020" And m_CP10 <> "101" And m_CP10 <> "308" Then
            Me.textPS.Text = "附件：核准" & textCP10 & "證明。"
        End If
        '91.11.25 end
        'Add By Cheng 2003/06/19
        '設定定稿使用的列印備註
        Select Case m_TM10
        Case "020" '大陸
            Select Case m_CP10
            Case "502" '許可合同
                Me.textPS.Text = "附件：許可合同備案通知書正本。"
            End Select
        End Select
        'Add By Cheng 2003/12/08
        '若核准申請, 判斷本案是否已收第一期註冊費
        'edit by nick 2004/12/23 分割與申請做相同的事情
        'If m_CP10 = "101" Then
        If m_CP10 = "101" Or m_CP10 = "308" Then
            '2005/7/19 MODIFY BY SONIA
            'strSQLA = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 in ('715','717')  "
            StrSQLa = "Select * From Caseprogress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 in ('715','717')  AND CP57 IS NULL"
            '2005/7/19 END
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                '已收第一期註冊費
                m_blnReceiveFirst = True
            Else
                '未收第一期註冊費
                m_blnReceiveFirst = False
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
        'End
    End Select
    
   'Add By Sindy 2009/07/17
   '若為TF案申請國家為日本011或古巴135且案件性質為申請101時, 預設領証費
   'Modify By Sindy 2022/9/22 桂英說,依林經理最後詢問代理人得知古巴已不需再繳註冊費，
   '故古巴無須控管領證本所期限、法定期限及領證費 ==> 取消 Or m_TM10 = "135"
   textFee = ""
   If m_TM01 = "TF" And (m_CP10 = "101" Or m_CP10 = "104") Then
      If m_TM10 = "011" Then
         textFee = "31000"
      'Modify By Sindy 2023/9/1 巴西117領證費預設：12000；
      ElseIf m_TM10 = "117" Then
         textFee = "12000"
      End If
   End If
   '2009/07/17 End
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'Add By Sindy 2013/1/11
   '若該筆移轉或讓與的受讓人(5個),與基本檔不符時,顯示訊息且不可輸入核准函
   cmdok.Enabled = True
   If m_CP10 = "501" Then
      If m_TM23 <> m_CP56 Or m_TM78 <> m_CP89 Or m_TM79 <> m_CP90 Or m_TM80 <> m_CP91 Or m_TM81 <> m_CP92 Then
         MsgBox "此案基本檔申請人與此程序受讓人不同，請確認資料！"
         cmdok.Enabled = False
      End If
   End If
   '2013/1/11 End
   
   'Modify By Sindy 2013/12/30 Mark
'   'Modify By Sindy 2013/12/19 台灣案的申請則鎖住
'   If m_TM10 = "000" And m_CP10 = "101" Then
'      textTM14.Enabled = False
'      textTMBM07_1.Enabled = False
'      textTMBM07_2.Enabled = False
'   End If
'   '2013/12/19 END
End Sub

Private Sub DisplayNextForm()
   frm02010401_5.SetData 0, m_TM01, True
   frm02010401_5.SetData 1, m_TM02, False
   frm02010401_5.SetData 2, m_TM03, False
   frm02010401_5.SetData 3, m_TM04, False
   frm02010401_5.SetData 5, m_CP09, False
   Me.Hide
   frm02010401_5.Show
   frm02010401_5.QueryData
End Sub

'Modify By Cheng 2002/11/07
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim bUpdate As Boolean
Dim strSubTMSQL As String
Dim strSubCPSQL As String
'Dim strCP09 As String 'Move by Lydia 2016/12/22 改成共用變數
'Dim strCP10 As String 'Move by Lydia 2016/12/22 改成共用變數
Dim strCP12 As String
Dim strCP27 As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP09 As String
Dim strNP14 As String
Dim strNP15 As String
Dim strNP22 As String
'Add By Cheng 2002/11/27
Dim strCP64 As String '進度備註
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP06 As String
Dim strCP07 As String
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount <= 0 Then
       GoTo EXITSUB
   End If
   
   'Added by Lydia 2017/01/18 T-175064的移轉(AA5051203)不知為何輸入到核准通知
   If m_TM10 < "010" And frm02010401_3.GetSelectResult() = "3" Then
      MsgBox "申請國家為台灣的不可選擇核准通知!", vbExclamation
      GoTo EXITSUB
   End If
   'end 2017/01/18
   
   rsTmp.MoveFirst
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
    Select Case m_TM01
    Case "T", "FCT", "TF"
        ' 設定SQL中Update TradeMark的語法
        strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                             "TM02 = '" & m_TM02 & "' AND " & _
                             "TM03 = '" & m_TM03 & "' AND " & _
                             "TM04 = '" & m_TM04 & "' "
    Case "TC"
        ' 設定SQL中Update Servicepractice的語法
        strSubTMSQL = "WHERE SP01 = '" & m_TM01 & "' AND " & _
                             "SP02 = '" & m_TM02 & "' AND " & _
                             "SP03 = '" & m_TM03 & "' AND " & _
                             "SP04 = '" & m_TM04 & "' "
    End Select
   ' 設定SQL中CaseProgress的語法
   strSubCPSQL = "WHERE CP09 = '" & m_CP09 & "' "
   
   ' 原實際結果或准駁日無資料時需 Update 實際結果及准駁日的欄位
   '核准
   'If frm02010401_3.GetSelectResult() = "1" Then
   'Modify By Sindy 2010/4/13
   If frm02010401_3.GetSelectResult() = "1" Or frm02010401_3.GetSelectResult() = "3" Then
      bUpdate = False
      If IsNull(rsTmp.Fields("CP24")) = False Then
         If IsEmptyText(rsTmp.Fields("CP24")) = True Then
            bUpdate = True
         End If
      Else
         bUpdate = True
      End If
      If IsNull(rsTmp.Fields("CP25")) = False Then
         If IsEmptyText(rsTmp.Fields("CP25")) = True Then
            bUpdate = True
         End If
      Else
         bUpdate = True
      End If
      
      If bUpdate = True Then
         '91.4.29 modify by sonia
         'strSQL = "UPDATE CaseProgress SET CP24 = '1', " & _
         '                                 "CP25=" & DBNullNumeric(textCP25) & ", " & _
         '                                 "CP35=" & DBNullString(textCP35) & " "
         strSql = "UPDATE CaseProgress SET CP24 = '1', " & _
                                          "CP25=" & DBNullDate(textCP25) & ", " & _
                                          "CP35=" & DBNullString(textCP35) & " "
         '91.4.29 end
         strSql = strSql & strSubCPSQL
         cnnConnection.Execute strSql
      End If
   End If
   
   'Add By Cheng 2002/06/13
   '若案件性質為授權或設定質權時, 更新授權(設定質權)期間
   If m_CP10 = "502" Or m_CP10 = "506" Then
      strSql = "Update Caseprogress Set cp53=" & DBDATE(Me.textCP53.Text) & ",CP54=" & DBDATE(Me.textCP54.Text) & " WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
    'Add By Cheng 2002/12/30
    strSql = "Update Caseprogress Set CP08='" & ChgSQL(Me.textCP08.Text) & "' WHERE CP09 = '" & m_CP09 & "' "
    cnnConnection.Execute strSql
    'Add By Cheng 2004/02/05
    '若系統類別為著作權(TC), 將證書號寫入進度備註
    If m_TM01 = "TC" Then
        If Me.txtNote.Text <> "" Then
            strSql = "Update Caseprogress Set CP64=CP64||Decode(CP64, Null, '', ',')||'證字第：'||'" & ChgSQL(Me.txtNote.Text) & "號' WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
        End If
    End If
    'End
    
   'Add By Sindy 2012/1/18
   '存檔時(商標局發函日)此欄有輸入值則存入cp64
   If (m_CP10 = "102" Or m_CP10 = "501") And Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" And textCP64.Visible = True Then
      strSql = "UPDATE caseprogress SET CP64 = decode(CP64,null,'主管機關受理函發文日：" & Trim(textCP64.Text) & "',CP64||" & "';'" & "||'主管機關受理函發文日：" & Trim(textCP64.Text) & "'),CP133= " & Trim(textCP64.Text) & _
               " WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 當案件性質為延展時, 更新商標基本檔專用期限
   If m_CP10 = "102" Then
      strSql = "UPDATE TradeMark SET TM21=" & DBDATE(textTM21) & "," & _
                                    "TM22=" & DBDATE(textTM22) & " "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
      'add by sonia 2015/12/14 同時更新TF未閉卷所有子案之專用期間TF-00049
      If m_TM01 = "TF" Then
         strSql = "UPDATE TradeMark SET TM21=" & DBDATE(textTM21) & "," & _
                                    "TM22=" & DBDATE(textTM22) & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "SUBSTR(TM02,1,5) = '" & Left(m_TM02, 5) & "' AND " & _
                  "TM29 IS NULL AND TM16='1'"
         cnnConnection.Execute strSql
      End If
      'end 2015/12/14
   End If
   ' 更新商標基本檔的是否閉卷
   '2011/10/4 MODIFY BY SONIA 閉卷才做,同時更新閉卷日期及原因,再加同時更新服務業務基本檔
   'If m_CP10 = "306" Or m_CP10 = "307" Then
   '   strSql = "UPDATE TradeMark SET TM29=" & DBNullString(textTM29) & " "
   '   strSql = strSql & strSubTMSQL
   '   cnnConnection.Execute strSql
   'End If
   'Modify By Sindy 2016/8/10 + 626.註銷
   If (m_CP10 = "306" Or m_CP10 = "307" Or m_CP10 = "626") And textTM29 = "Y" Then
      If m_TM01 = "TC" Then
         strSql = "UPDATE SERVICEPRACTICE SET SP15='Y',SP17='09',SP16=" & strSrvDate(1) & " "
      Else
         strSql = "UPDATE TradeMark SET TM29='Y',TM31='09',TM30=" & strSrvDate(1) & " "
      End If
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
   '2011/10/4 END
   '2006/6/1 ADD BY SONIA 原未撰寫此段更新
   ' 若案件性質為移轉時, 更新商標基本檔之卷宗性質
   'Modify By Sindy 2019/12/20 + And (m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "TF")
   If m_CP10 = "501" And _
      (m_TM01 = "T" Or m_TM01 = "FCT" Or m_TM01 = "TF") Then
      strSql = "UPDATE TradeMark SET TM28='1' "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
   '2006/6/1 END
   'add by nickc 2006/06/07  更新商標基本檔的專用權是否存在
   'Modify By Sindy 2016/8/10 + 626.註銷
   If m_CP10 = "307" Or m_CP10 = "626" Then
      strSql = "UPDATE TradeMark SET TM17='N' "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
   'add by nickc 2006/11/17
   If textPrint <> "N" Then
        Select Case m_TM01
        Case "T", "FCT", "TF"
            strSql = "UPDATE TradeMark SET TM77='" & textPrint & "' "
            strSql = strSql & strSubTMSQL
            cnnConnection.Execute strSql
        Case Else
            strSql = "UPDATE Servicepractice SET SP72='" & textPrint & "' "
            strSql = strSql & strSubTMSQL
            cnnConnection.Execute strSql
        End Select
   End If
   'Modify By Cheng 2002/07/22
   '取消更新"專用權是否存在"欄位
'   ' 更新商標基本檔的專用權是否存在
'   strSQL = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' "
'   strSQL = strSQL & strSubTMSQL
'   cnnConnection.Execute strSQL
   ' 案件性質為申請時
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   'modify by sonia 2024/9/4 +柬埔寨046(TF-00083-3-1-03之領土延伸)
   If m_CP10 = "101" Or m_CP10 = "308" Or m_TM10 = "046" Then
      ' 更新審定號, 來函收文日, 公告日
      'strSQL = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
      '                              "TM13 = " & DBDATE(textCP05S) & "," & _
      '                              "TM14 = " & DBDATE(textTM14) & " "
      strSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                    "TM13 = " & DBNullDate(textCP05S) & "," & _
                                    "TM14 = " & DBNullDate(textTM14) & " "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
      'Modify By Cheng 2002/07/22
      '當案件性質為商申案時(101), 才更新目前准/駁及審定來函日兩個欄位
'      ' 當使用者輸入要更新基本檔之准/駁時, 更新目前准/駁及審定來函日兩個欄位
'      If textTM16S = "Y" Then
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" Then
      'If m_CP10 = "101" Or m_CP10 = "308" Then
         strSql = "UPDATE TradeMark SET TM16='1'," & _
                                       "TM13=" & DBNullDate(textCP25) & " "
         'If IsEmptyText(textCP25) = True Then
         '   strSQL = "UPDATE TradeMark SET TM16='1'," & _
         '                                 "TM13=" & "NULL" & " "
         'Else
         '   strSQL = "UPDATE TradeMark SET TM16='1'," & _
         '                                 "TM13=" & DBDATE(textCP25) & " "
         'End If
         strSql = strSql & strSubTMSQL
         cnnConnection.Execute strSql
      'End If
      ' 當系統別為"TF"而此筆資料是母案時, 同時更新子案的資料
      If m_TM01 = "TF" And m_TM04 = "00" Then
         If IsEmptyText(textCP25) = True Then
            strSql = "UPDATE TradeMark SET TM16='1'," & _
                                          "TM13=" & "NULL" & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "(TM16 IS NULL OR TM16 = '' OR TM16 = ' ')"
         Else
            strSql = "UPDATE TradeMark SET TM16='1'," & _
                                          "TM13=" & DBDATE(textCP25) & " " & _
                     "WHERE TM01 = '" & m_TM01 & "' AND " & _
                           "TM02 = '" & m_TM02 & "' AND " & _
                           "TM03 = '" & m_TM03 & "' AND " & _
                           "(TM16 IS NULL OR TM16 = '' OR TM16 = ' ')"
         End If
         cnnConnection.Execute strSql
      End If
   End If
    
    'Add By Cheng 2002/11/27
    strCP64 = ""
'    '補核准通知
'    If frm02010401_3.GetSelectResult() = "3" Then
'        StrSQLa = "Select CP27 From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And  CP10='1001' And CP43 ='" & m_CP09 & "' And CP27 Is Not Null Order By CP27 "
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'        rsA.CursorLocation = adUseClient
'        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'        While Not rsA.EOF
'            strCP64 = strCP64 & rsA.Fields(0).Value & ","
'            rsA.MoveNext
'        Wend
'        If rsA.State <> adStateClosed Then rsA.Close
'        Set rsA = Nothing
'        If strCP64 <> "" Then
'            strCP64 = Left(strCP64, Len(strCP64) - 1)
'            strCP64 = "核准通知日期：" & strCP64
'        End If
'        strSql = "Delete From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And  CP10='1001' And CP43 ='" & m_CP09 & "'"
'        cnnConnection.Execute strSql
'    End If
   
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質
   strCP10 = "1001"
   Select Case frm02010401_3.GetSelectResult
      'Case "1", "2":    '2006/11/1 MODIFY BY SONIA 改變原處分也要掛期限
      Case "1", "2", "3":    'Modify By Sindy 2010/4/13
         Select Case frm02010401_3.GetSelectResult
            Case "1": strCP10 = "1001" '核准
            Case "2": strCP10 = "1403" '重為處分
            Case "3": strCP10 = "1102" '核准通知
         End Select
         'edit by nick 2004/12/23 分割與申請做相同的事情
         'If m_CP10 = "101" Then
         If m_CP10 = "101" Or m_CP10 = "308" Then
            'Add By Cheng 2003/11/18
            '若為商申案且本案申請日為921128(含)以後者
            If m_TM10 = "000" And Val(m_TM11) >= 20031128 Then
                '法定期限
                strCP07 = DBDATE(DateAdd("m", 2, ChangeWStringToWDateString(DBDATE(Me.textCP25.Text))))
                '本所期限
                'edit by nick 2004/07/28 改為減 4 天
                'strCP06 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strCP07))))
                'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                   strCP06 = PUB_GetOurDeadline(DBDATE(strCP07))
                Else
                '2014/10/6 END
                   strCP06 = DBDATE(DateAdd("d", -4, ChangeWStringToWDateString(DBDATE(strCP07))))
                End If
                strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                'Add By Sindy 2012/4/20
                m_NP08 = strCP06
                '2012/4/20 End
            End If
         End If
'      Case "3":
'         StrCP10 = "1102"
   End Select
    'Modify By Cheng 2002/11/22
    '若案件性質為自請撤回或自請撤銷照舊
    'Modify By Sindy 2016/8/10 + 626.註銷
    If rsTmp("CP10") = "306" Or rsTmp("CP10") = "307" Or rsTmp("CP10") = "626" Then
'add by nick 2004/10/19
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP64) " & _
'                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                         "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                         "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & strCP64 & "' ) "
        'Modify By Sindy 2010/7/12 承辦人改掛操作人員 old:m_CP14
        '2010/9/28 MODIFY BY SONIA 宋若蘭說因第一期註冊費的期限管制表承辦人會帶成程序故改為仍掛原承辦人,離職掛P2001商標處,於上方控制m_CP14
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP64) " & _
                 "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                         "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & _
                         "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & strCP64 & "' ) "

        'End
        cnnConnection.Execute strSql
        
        'Add By Sindy 2021/5/20
        '註銷的核准固定上發文日
        If rsTmp("CP10") = "626" Then
            strSql = "Update CaseProgress Set CP27=" & strSrvDate(1) & " Where CP09='" & strCP09 & "' "
            cnnConnection.Execute strSql
        End If
        
        '判斷其相關總收文號的案件性質屬於商申案
        StrSQLa = "Select CP10 From CaseProgress Where CP09=(Select CP43 From CaseProgress Where CP09='" & m_CP09 & "' " & " )"
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            '若為商申案
            If Left("" & rsA.Fields(0).Value, 1) <> "4" And Left("" & rsA.Fields(0).Value, 1) <> "6" And "" & rsA.Fields(0).Value <> "202" And "" & rsA.Fields(0).Value <> "204" And "" & rsA.Fields(0).Value <> "205" And "" & rsA.Fields(0).Value <> "207" Then
                '上發文日
                strSql = "Update CaseProgress Set CP27=" & strSrvDate(1) & " Where CP09='" & strCP09 & "' "
                cnnConnection.Execute strSql
            '93.9.3 ADD BY SONIA
            Else
                '非商申案承辦人抓原承辦人
                strSql = "Update CaseProgress Set CP14='" & rsTmp("CP14") & "' Where CP09='" & strCP09 & "' "
                cnnConnection.Execute strSql
            '93.9.3 END
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    '其他案件性質
    Else
      ' 台灣案件性質為申請時,同時存第一期註冊費期限
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_TM10 = "000" And m_CP10 = "101" And Val(m_TM11) >= 20031128 Then
      If m_TM10 = "000" And (m_CP10 = "101" Or m_CP10 = "308") And Val(m_TM11) >= 20031128 Then
'edit by nick 2004/10/19
'           strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64,CP06,CP07) " & _
'                    "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                            "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                            "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "'," & Val(ServerDate) & ",'" & strCP64 & "'," & DBDATE(strCP06) & "," & DBDATE(strCP07) & " ) "
'           cnnConnection.Execute strSQL
           strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64,CP06,CP07) " & _
                    "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                            "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & _
                            "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "'," & Val(ServerDate) & ",'" & strCP64 & "'," & DBDATE(strCP06) & "," & DBDATE(strCP07) & " ) "
           cnnConnection.Execute strSql
            '2005/7/19 modify by sonia
            'strSQLA = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 IN ('715','717') "
            StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(m_TM01 & m_TM02 & m_TM03 & m_TM04) & " And CP10 IN ('715','717') AND CP57 IS NULL"
            '2005/7/19 end
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            '若有收文第一期註冊費, 更新進度檔
            If rsA.RecordCount > 0 Then
                '93.7.22 ADD BY SONIA 若已收文本所期限改為來函日+7天
                strCP06 = DBDATE(DateAdd("d", 7, ChangeWStringToWDateString(DBDATE(textCP25))))
                strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                '93.7.22 END
                StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
                cnnConnection.Execute StrSQLa
            '若未收文第一期註冊費, 新增下一程序檔
            Else
                'Modify By Sindy 2012/6/27 商標修法
'                If Val(DBDATE(m_CP05)) >= 20120701 Then
                  strNP07 = "717"
'                Else
'                '2012/6/27 End
'                  strNP07 = "715"
'                End If
                strNP22 = GetNextProgressNo()
                strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                                DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
                cnnConnection.Execute strSql
                 
                 'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
                 Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
                
                 ' 93.6.8 add by sonia 加印接洽結案單
                 pub_AddressListSN = pub_AddressListSN + 1
                 PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
                 '93.6.8 end
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
      Else
'2004/10/19
'           strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64) " & _
'                    "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                            "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
'                            "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "'," & Val(ServerDate) & ",'" & strCP64 & "' ) "
'           cnnConnection.Execute strSQL
           
            'Add By Sindy 2009/07/17
            'TF案申請國家為日本011或古巴135且案件性質為申請101時, 新增下一程序701期限,故進度檔也要放期限
            'Modify By Sindy 2022/9/22 桂英說,古巴無須控管領證本所期限、法定期限及領證費 + And Val(textNP09) > 0
            'Modify By Sindy 2023/9/1 And (m_TM10 = "011" Or m_TM10 = "135") 改判斷 And Val(textFee.Text) > 0
            If m_TM01 = "TF" And Val(textFee.Text) > 0 And (m_CP10 = "101" Or m_CP10 = "104") And Val(textNP09) > 0 Then
               '法定期限
               strCP07 = DBDATE(textNP09)
               '本所期限
               strCP06 = DBDATE(textNP08)
               strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64,CP06,CP07) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                                "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & _
                                "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "'," & Val(ServerDate) & ",'" & strCP64 & "'," & DBDATE(strCP06) & "," & DBDATE(strCP07) & ") "
               cnnConnection.Execute strSql
            Else
               'add by sonia 2019/1/4
               If m_CP10 = "723" Then
                  strCP64 = Trim(textCP64.Text)
               End If
               'end 2019/1/4
               strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP27,CP64) " & _
                        "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                                "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "'," & _
                                "'" & "N" & "','" & "N" & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "'," & Val(ServerDate) & ",'" & strCP64 & "' ) "
               cnnConnection.Execute strSql
            End If
            '2009/07/17 End
      End If
      '93.6.3 END
      '92.11.19 ADD BY SONIA
      If strCP10 = "1403" Then
          strSql = "Update CaseProgress Set CP24='1' Where CP09='" & strCP09 & "' "
          cnnConnection.Execute strSql
      End If
      '92.11.19 END
    End If
    
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      strExc(1) = ""
      If m_CP10 <> "101" And m_CP10 <> "308" Then '非申請案
         If m_TM10 <> "000" Then '為台->大
            strExc(1) = Pub_GetSpecMan("內商程序客戶函發後補看人員")
         End If
      End If
      If Val(strCP07) > 0 Then '有期限
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), , True, m_TM23, strCP10, m_TM44, , , , , strExc(1)
      Else
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), , False, m_TM23, strCP10, m_TM44, , , , , strExc(1)
      End If
   End If
   '2019/12/19 END
    
   ' 更新下一程序檔案件性質為催審的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "305"
   cnnConnection.Execute strSql
   '2007/8/7 ADD BY SONIA更新下一程序檔案件性質為收達及提申的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 IN (997,998) AND NP06 IS NULL"
   cnnConnection.Execute strSql
   '2007/8/7 END
   ' 當使用者在前畫面選取2時, 更新下一程序檔案件性質為改變原處份的資料
   If frm02010401_3.GetSelectResult() = "2" Then
      'modify by sonia 2022/6/24 因為下一程序1403期限都是因撤銷原處分來函而產生的，所以此處不可限制NP01 = '" & m_CP09 & "', 改為 AND NP06 IS NULL條件
      'strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "1403"
      strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = 1403 AND NP06 IS NULL"
      'end 2022/6/24
      cnnConnection.Execute strSql
   End If
   ' 依案件性質來決定是否要新增一筆資料到下一程序檔
   'Modify By Sindy 2009/06/16 增加109.被異議續展
   Select Case m_CP10
      'Case "102":
      Case "102", "109":
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
         '2010/3/2 MODIFY BY SONIA
         'strNP09 = DBDATE(textTM22)
         If m_CP10 = "102" Then
            strNP09 = DBDATE(textTM22)
         Else
            'Modified by Lydia 2019/11/13 改用共用模組, 第1次專用期間=公告日+10年-1天，之後延展102沒有減１天；與專利不一樣
            'strNP09 = DBDATE(DateAdd("M", 120, ChangeWStringToWDateString(DBDATE(m_CP07))))
            'Modify By Sindy 2022/3/7 + m_TM10 : 延展後之專用期限年度倘有2月29日時，專用期限止日應為2月29日，而非以加10年之方式計算為2月28日
            strNP09 = PUB_GetEndDate(DBDATE(m_CP07), Val(m_NA14), "N", m_TM10)
         End If
         '2010/3/2 END
        'Modify By Cheng 2003/09/01
'         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP09)), Val(DBMONTH(strNP09)), Val(DBDAY(strNP09)) - 2)))
         'edit by nickc 2007/06/13  TF 改成一個月
         If m_TM01 = "TF" Then
            strNP08 = DBDATE(DateAdd("m", -1, ChangeWStringToWDateString(DBDATE(strNP09))))
         Else
            'Modify By Sindy 2014/10/6 台灣案之本所期限設定
            If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
               strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
            Else
            '2014/10/6 END
               'modify by sonia 2023/3/7 大陸案也改為2個工作天
               'strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
               strNP08 = PUB_GetOurDeadline(DBDATE(strNP09))
            End If
         End If
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         '2010/2/23 modify by sonia 沒有才新增,因為已提申時已先掛T-110343
         Set rsA = New ADODB.Recordset
         If rsA.State = 1 Then rsA.Close
         strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='" & m_CP10 & "' AND NP06 IS NULL"
         rsA.CursorLocation = adUseClient
         rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount <> 0 Then
            strSql = "UPDATE NEXTPROGRESS SET NP01='" & strCP09 & "',NP08=" & strNP08 & " where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP09=" & strNP09 & " AND NP07='" & m_CP10 & "' AND NP06 IS NULL"
         Else
            strNP22 = GetNextProgressNo()
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                     "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & m_CP10 & "," & _
                             strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         End If
         '2010/2/23 END
         cnnConnection.Execute strSql
   End Select
   
   'Added by Lydia 2016/09/12 大陸案審定核准輸入時,同時管制催註冊証時間8個月。
   '  因為大陸案核准由國方系統給資料，所以是在內商\資料處理\商申機關來函\大陸商標審定公告及通知續展匯入作業frm020320呼叫存檔程式
   'Modify By Sindy 2022/3/2 + Or strCP10 = "1102"):代理人通知核准,程序人員輸入核准通知時也要一併掛下一程序[註冊證]期限,管制證書催審!!
   If m_TM01 = "T" And m_TM10 = "020" And (strCP10 = "1001" Or strCP10 = "1102") And (m_CP10 = "101" Or m_CP10 = "308") Then
      '申請(101)及分割(308)核准函掛下一程序1701期限NP08=NP09=系統日+8個月,NP10=申請或分割之CP14
      'modify by sonia 2024/9/12 8個月改6個月，本所期限改為工作日
      strNP08 = CompDate(1, 6, strSrvDate(1))
      strNP09 = strNP08
      strNP08 = PUB_GetWorkDay1(strNP08, True) 'add by sonia 2024/9/24 若本所期限非工作天則直接調整至最近的工作天
      strSql = "select * from NEXTPROGRESS where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP07='1701' AND NP06 IS NULL"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '若重復匯入則更新期限
         strSql = "update nextprogress set np01='" & strCP09 & "', np08='" & strNP08 & "', np09='" & strNP09 & "' where NP02='" & m_TM01 & "' and NP03='" & m_TM02 & "' and NP04='" & m_TM03 & "' and NP05='" & m_TM04 & "' AND NP22='" & "" & rsA.Fields("NP22") & "' AND NP07='1701' AND NP06 IS NULL"
      Else
         'Modify By Sindy 2022/3/3 大陸分割案之核准通知，下一程序掛的註冊證期限要掛在分割案的承辦人
         If m_CP10 = "308" Then
            strSql = "SELECT '1' ord ,CP09,CP14,ST04 FROM CaseProgress,STAFF WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                     "AND CP10='308' AND CP14=ST01(+) "
         Else
         '2022/3/3 END
            'Added by Lydia 2016/09/23 抓申請或分割的承辦人
            'Modified by Lydia 2016/10/18 只抓申請,若申請為B類收文改抓CP31='Y'的A類收文承辦人
            'strSql = "SELECT CP14,ST04 FROM CaseProgress,STAFF WHERE CP09 = '" & m_CP09 & "' AND CP14=ST01(+) "
            '
            strSql = "SELECT '1' ord ,CP09,CP14,ST04 FROM CaseProgress,STAFF WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                     "AND CP10='101' AND SUBSTR(CP09,1,1)='A' AND CP14=ST01(+) " & _
                     "UNION SELECT '2' ord ,CP09,CP14,ST04 FROM CaseProgress,STAFF WHERE CP01='" & m_TM01 & "' AND CP02='" & m_TM02 & "' AND CP03='" & m_TM03 & "' AND CP04='" & m_TM04 & "' " & _
                     "AND CP31='Y' AND SUBSTR(CP09,1,1)='A' AND CP14=ST01(+) ORDER BY 1 "
         End If
         intI = 1
         strExc(1) = "P2001"  '原承辦人若離職改抓P2001
         Set rsA = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & rsA.Fields("ST04") <> "2" Then strExc(1) = "" & rsA.Fields("CP14")
         End If
         
         strNP22 = GetNextProgressNo()
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "', 1701 ," & _
                          strNP08 & "," & strNP09 & ",'" & strExc(1) & "'," & strNP22 & ")"
      End If
      cnnConnection.Execute strSql
   End If
   'end 2016/09/12
   
   'Added by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
   bolA1kdataMail = False
   m_ULD02 = "": ': m_AC2470 = "" 'Remove by Lydia 2017/04/06
   m_rA1k28 = "": m_rSpec = ""  'Added by Lydia 2017/04/06
   'Modified by Lydia 2020/03/19 該核准的相關收文號，若無收取任何費用(規費或服務費)就無需產生 收款寄證。
   'If strCP10 = "1001" Then
   'Modified by Lydia 2021/05/20 收款寄證-限MCT的案件,所以必須申請國家是台灣; ex.T-166495
   If strCP10 = "1001" And Val(m_CP16) > 0 And m_TM10 = 台灣國家代號 Then
      'Modified by Lydia 2017/03/14 抓最新的智權人員
      'bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_CP09, m_CP10, m_CP13, strCP09, m_ULD02, m_AC2470)
      'Modified by Lydia 2017/04/06 區分請款對象
      'bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_CP09, m_CP10, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), strCP09, m_ULD02, m_AC2470)
      'Modifeid by Lydia 2023/04/11 +申請人1~5 +m_TM23 & "," & m_TM78 & "," & m_TM79 & "," & m_TM80 & "," & m_TM81
      bolA1kdataMail = PUB_CheckA1kdataMail(m_TM01, m_TM02, m_TM03, m_TM04, m_TM44, m_TM23 & "," & m_TM78 & "," & m_TM79 & "," & m_TM80 & "," & m_TM81, m_CP09, m_CP10, PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), strCP09, m_ULD02, m_rA1k28, m_rSpec)
   End If
   'end 2016/12/22
   
   'Add By Sindy 2009/07/17
   'TF案申請國家為日本011或古巴135且案件性質為申請101時, 新增下一程序領証701期限
   'Modify By Sindy 2022/9/22 桂英說,古巴無須控管領證本所期限、法定期限及領證費 + And Val(textNP09) > 0
   'Modify By Sindy 2023/9/1 And (m_TM10 = "011" Or m_TM10 = "135") 改判斷 And Val(textFee.Text) > 0
   If m_TM01 = "TF" And Val(textFee.Text) > 0 And (m_CP10 = "101" Or m_CP10 = "104") And Val(textNP09) > 0 Then
      strNP09 = DBDATE(textNP09)
      strNP08 = DBDATE(textNP08)
      strNP22 = GetNextProgressNo()
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',701," & _
                        strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 加印接洽結案單
      pub_AddressListSN = pub_AddressListSN + 1
      PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
   End If
   '2009/07/17 End
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If

   Dim SeekMonTM01 As String
   Dim SeekMonTM02 As String
   Dim SeekMonTM03 As String
   Dim SeekMonTM04 As String
   'ADD BY nickc 2006/09/27 若是B類申請案，則代表是分割產生，要檢查分割的相關子案是否有准駁，若全都有，則將母案上閉卷
   'MODIFY BY SONIA 2015/6/25 大陸案母案不可閉卷(分割案為母案抽部分出來)分割案T-196252母案T-190094
   'If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" Then
   If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" And m_TM10 = "000" Then
       Set rsA = New ADODB.Recordset
       If rsA.State = 1 Then rsA.Close
       strSql = "select * from divisioncase where dc01='" & m_TM01 & "' and dc02='" & m_TM02 & "' and dc03='" & m_TM03 & "' and dc04='" & m_TM04 & "' "
       rsA.CursorLocation = adUseClient
       rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsA.RecordCount <> 0 Then
            SeekMonTM01 = CheckStr(rsA.Fields("dc05"))
            SeekMonTM02 = CheckStr(rsA.Fields("dc06"))
            SeekMonTM03 = CheckStr(rsA.Fields("dc07"))
            SeekMonTM04 = CheckStr(rsA.Fields("dc08"))
            Set rsA = New ADODB.Recordset
            If rsA.State = 1 Then rsA.Close
            strSql = "select * from divisioncase,trademark where dc05='" & SeekMonTM01 & "' and dc06='" & SeekMonTM02 & "' and dc07='" & SeekMonTM03 & "' and dc08='" & SeekMonTM04 & "' and dc01=tm01(+) and dc02=tm02(+) and dc03=tm03(+) and dc04=tm04(+) and (tm16 is null or tm16='') "
            rsA.CursorLocation = adUseClient
            rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount = 0 Then
                strSql = "update trademark set tm29='Y',tm30=to_number(to_char(sysdate,'YYYYMMDD')),tm31='87' where tm01='" & SeekMonTM01 & "' and tm02='" & SeekMonTM02 & "' and tm03='" & SeekMonTM03 & "' and tm04='" & SeekMonTM04 & "' and (tm29 is null or tm29='') "
                cnnConnection.Execute strSql
            End If
       End If
   End If
   
   'Added by Lydia 2016/10/20 台-大分割案的母案結果於子案申請核准輸入時,同時把子案及母案分割下結果;PS：大陸分割是將沒有問題的部分分割出來，所以分割案一定會核准，不必考慮核駁。
   If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" And m_TM10 = "020" Then
       strSql = "select c1.cp09 CCP09,c2.cp09 MCP09 From divisioncase, caseprogress c1,caseprogress c2 " & _
               "where dc01='" & m_TM01 & "' and dc02='" & m_TM02 & "' and dc03='" & m_TM03 & "' and dc04='" & m_TM04 & "' " & _
               "and dc01=c1.cp01(+) and dc02=c1.cp02(+) and dc03=c1.cp03(+) and dc04=c1.cp04(+) and c1.cp10(+)='308' " & _
               "and dc05=c2.cp01(+) and dc06=c2.cp02(+) and dc07=c2.cp03(+) and dc08=c2.cp04(+) and c2.cp10(+)='308' "
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         '更新子案的分割結果
         If "" & rsA.Fields("CCP09") <> "" Then
            strSql = "UPDATE CaseProgress SET CP24='1',CP25=" & DBNullDate(textCP25) & " Where CP09='" & rsA.Fields("CCP09") & "' "
            cnnConnection.Execute strSql
         End If
         '更新母案的分割結果
         If "" & rsA.Fields("MCP09") <> "" Then
            strSql = "UPDATE CaseProgress SET CP24='1',CP25=" & DBNullDate(textCP25) & " Where CP09='" & rsA.Fields("MCP09") & "' "
            cnnConnection.Execute strSql
         End If
      End If
   End If
   
'cancel by sonia 2020/12/28 TF專用10年延展10年,在每10年內只要提一次使用宣誓,故不再抓NA39改抓NA78
'2011/9/23 CANCEL BY SONIA 因為現行3個國家都是專用10年延展10年,在每10年內只要提一次使用宣誓,在發註冊證掛期限即可
'modify by sonia 2014/10/31 原抓na38改為抓na39,有值時每次延展後都要管制使用宣誓
Dim strNA78 As String
Dim strNA38 As String   'add by sonia 2024/9/4
Dim MyTFrs As New ADODB.Recordset
   'add by nickc 2007/11/13 TF 掛使用宣誓期限
   'modify by sonia 2016/12/6 母案或是領土延伸案才要做 TF-00073-0-1-01不必,故加m_TM03="0"
   'modify by sonia 2020/12/28 僅延展才要做
   If m_TM01 = "TF" And m_TM03 = "0" And m_CP10 = "102" Then
      Set MyTFrs = New ADODB.Recordset
      If MyTFrs.State = 1 Then MyTFrs.Close
      MyTFrs.CursorLocation = adUseClient
      'add by nickc 2007/11/13 第2碼的第6個字是0的只要判斷前 5 字相同，且第4碼<>"00"
      '若第2碼的第6個字<>0的只要判斷前 6 字相同，且第4碼<>"00"
      MyTFrs.Open "select * from trademark where tm01='" & m_TM01 & "' and tm04<>'00' " & IIf(Mid(m_TM02, 6, 1) = "0", " and substr(tm02,1,5)='" & Mid(m_TM02, 1, 5) & "' ", " and tm02='" & m_TM02 & "' "), cnnConnection, adOpenStatic, adLockReadOnly
      If MyTFrs.RecordCount <> 0 Then
          'edit by nickc 2007/11/13 原只有美國有掛使用宣誓，現改成國家檔有掛就要掛
          MyTFrs.MoveFirst
          Do While Not MyTFrs.EOF
              ' 取得使用宣誓年度
              strNA78 = 0
              Set rsA = New ADODB.Recordset
              Set rsA = Nothing
              StrSQLa = "SELECT * FROM Nation WHERE NA01 = '" & CheckStr(MyTFrs.Fields("tm10")) & "' AND na78 IS NOT NULL "
              rsA.CursorLocation = adUseClient
              rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
              'edit by nickc 2007/11/13
              'If rsA.RecordCount > 0 Then strna39 = rsA.Fields("na39")
              If rsA.RecordCount > 0 Then
                  strNA78 = rsA.Fields("na78")
                  If rsA.State <> adStateClosed Then rsA.Close
                  '法定期限  '2007/11/13 註解  秀玲說下一程序，新增子案資料，期限由母案或是領土延伸本案來計算
                  strCP07 = DBDATE(DateAdd("yyyy", Val(strNA78), ChangeWStringToWDateString(DBDATE(textTM21))))
                  '本所期限
                  'modify by sonia 2014/10/28 改同CFT,業務說改成本所=法定-2個月 不管任何國家
                  'strCP06 = DBDATE(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(strCP07))))
                  strCP06 = CompDate(1, -2, strCP07)
                  strCP06 = PUB_GetWorkDay1(strCP06, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  '先檢查是否已收文 105
                  StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & " And CP10='105' AND CP27 IS NULL AND CP57 IS NULL"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                  '若有收文使用宣誓, 更新進度檔
                  If rsA.RecordCount > 0 Then
                      StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
                      cnnConnection.Execute StrSQLa
                  '若未收文使用宣誓, 新增下一程序檔
                  Else
                      ' 檢查下一程序有無使用宣誓
                      Set rsA = New ADODB.Recordset
                      StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & " And np07=105 AND NP06 IS NULL"
                      rsA.CursorLocation = adUseClient
                      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
                      If rsA.RecordCount > 0 Then
                          strSql = "update NextProgress set np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(CheckStr(MyTFrs.Fields("tm01")) & CheckStr(MyTFrs.Fields("tm02")) & CheckStr(MyTFrs.Fields("tm03")) & CheckStr(MyTFrs.Fields("tm04"))) & " And np07=105 "
                          cnnConnection.Execute strSql
                      Else
                          '2007/11/13 註解  秀玲說下一程序，新增子案資料，智權人員掛母案或是領土延伸本案收文號也掛母案那道
                          strNP07 = "105"
                          strNP22 = GetNextProgressNo()
                          strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                                  "VALUES ('" & strCP09 & "','" & CheckStr(MyTFrs.Fields("tm01")) & "','" & CheckStr(MyTFrs.Fields("tm02")) & "','" & CheckStr(MyTFrs.Fields("tm03")) & "','" & CheckStr(MyTFrs.Fields("tm04")) & "'," & strNP07 & "," & _
                                  DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & strNP22 & ")"
                          cnnConnection.Execute strSql
                      End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
              End If
              MyTFrs.MoveNext
          Loop
      End If
   'add by sonia 2024/9/4 柬埔寨046申請101或領土延伸104之核准，以畫面上柬埔寨公告日計算使用宣誓期限更新(TF-00083-3-1-03)
   ElseIf m_TM01 = "TF" And m_TM10 = "046" And (m_CP10 = "101" Or m_CP10 = "104") Then
      strNA38 = 0
      Set rsA = New ADODB.Recordset
      Set rsA = Nothing
      StrSQLa = "SELECT * FROM Nation WHERE NA01 = '" & m_TM10 & "' AND na38 IS NOT NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         strNA38 = rsA.Fields("na38")
         If rsA.State <> adStateClosed Then rsA.Close
         strCP07 = DBDATE(DateAdd("yyyy", Val(strNA38), ChangeWStringToWDateString(DBDATE(textTM14))))
         strCP06 = CompDate(1, -2, strCP07)
         strCP06 = PUB_GetWorkDay1(strCP06, True)
         '先檢查是否已收文 105
         StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(CheckStr(m_TM01) & CheckStr(m_TM02) & CheckStr(m_TM03) & CheckStr(m_TM04)) & " And CP10='105' AND CP27 IS NULL AND CP57 IS NULL"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         '若有收文使用宣誓, 更新進度檔
         If rsA.RecordCount > 0 Then
            StrSQLa = "Update CaseProgress Set CP06=" & strCP06 & ", CP07=" & strCP07 & " Where CP09='" & rsA("CP09").Value & "' "
            cnnConnection.Execute StrSQLa
         '若未收文使用宣誓, 新增下一程序檔
         Else
            ' 檢查下一程序有無使用宣誓
            Set rsA = New ADODB.Recordset
            StrSQLa = "Select * From NextProgress Where " & ChgNextProgress(CheckStr(m_TM01) & CheckStr(m_TM02) & CheckStr(m_TM03) & CheckStr(m_TM04)) & " And np07=105 AND NP06 IS NULL"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                strSql = "update NextProgress set np08=" & DBDATE(strCP06) & ",np09=" & DBDATE(strCP07) & " where " & ChgNextProgress(CheckStr(m_TM01) & CheckStr(m_TM02) & CheckStr(m_TM03) & CheckStr(m_TM04)) & " And np07=105 "
                cnnConnection.Execute strSql
            Else
                strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & strCP09 & "','" & CheckStr(m_TM01) & "','" & CheckStr(m_TM02) & "','" & CheckStr(m_TM03) & "','" & CheckStr(m_TM04) & "'," & "105" & "," & _
                        DBDATE(strCP06) & "," & DBDATE(strCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
                cnnConnection.Execute strSql
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
   'end 2024/9/4
   End If
   
   'Add By Sindy 2009/09/24
   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
   Dim strCP48 As String, strCP09B As String
   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And _
      ((m_TM01 = "T" And m_TM10 = "020") Or (m_TM01 = "FCT" And m_TM10 = "000")) Then
      strCP09B = AutoNo("B", 6)
      '承辦期限為系統日加4個工作天
      strCP48 = DBDATE(Pub_GetHandleDay(m_TM01, m_TM10, "722", strSrvDate(1), , m_CP09))
      '2011/4/28 modify by sonia 智權人員原抓點選收文號之智權人員,改抓該案最後收文在職智權人員
      strSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp48,cp20,cp26,cp32,cp43) " & _
                     "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                     "," & CNULL(m_TM04) & "," & CNULL(strSrvDate(1)) & "," & CNULL(strCP09B) & ",722," & _
                     CNULL(GetSalesArea(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "," & CNULL(strCP48) & ",'N','N','N'," & CNULL(strCP09) & ")"
      cnnConnection.Execute strSql
   End If
   '2009/09/24 End
   
   'add by nickc 2005/04/22
   Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04
   
   'Added by Morgan 2017/4/13 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/13
   
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010401_1", strCP09
   End If
   '2019/5/10 END
   
   'Add By Cheng 2002/11/07
   cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
EXITSUB:
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
'edit by nickc 2008/04/25 改整批
'    'add by nick 2004/10/27
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
       'Add By Cheng 2002/07/18
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   
   Set frm02010401_4 = Nothing
End Sub

Private Sub textCP08_LostFocus()
On Error GoTo ErrorHandler

'Add By Cheng 2002/01/15
If Len(Me.textCP08.Text) > 0 Then
   m_intNumBegin = InStr(Me.textCP08.Text, m_strNumBegin)
   m_intNumEnd = InStr(Me.textCP08.Text, m_strNumEnd)
Else
   m_intNumBegin = 0
   m_intNumEnd = 0
End If
If m_intNumBegin < m_intNumEnd Then
   Me.textCP35.Text = Mid(Me.textCP08.Text, m_intNumBegin + 1, (m_intNumEnd - m_intNumBegin - 1))
End If

Exit Sub

ErrorHandler:
   m_intNumBegin = 0
   m_intNumEnd = 0
End Sub

Private Sub textCP53_GotFocus()
InverseTextBox textCP53
End Sub

Private Sub textCP53_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

    '申請國家為台灣
   If m_TM10 < "010" Then
        ' 檢核是否為民國日期
        If CheckIsTaiwanDate(Me.textCP53, False) = False Then
           Cancel = True
           strTit = "資料檢核"
           strMsg = "請輸入正確的日期"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCP53.SetFocus
           textCP53_GotFocus
           Exit Sub
        End If
   '申請國家非台灣
   Else
      ' 檢核是否為西元日期
        If CheckIsDate(Me.textCP53, False) = False Then
           Cancel = True
           strTit = "資料檢核"
           strMsg = "請輸入正確的日期"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCP53.SetFocus
           textCP53_GotFocus
           Exit Sub
        End If
   End If

If Val(Me.textCP53.Text) < Val(Me.textTM21.Text) Or Val(Me.textCP53.Text) > Val(Me.textTM22.Text) Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = Replace(Me.Label4(0).Caption, "：", "") & "與專用期間不符, 是否重新輸入???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "專用期間：" & Me.textTM21.Text & "－" & Me.textTM22.Text & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "－" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP53.SetFocus
      textCP53_GotFocus
      Exit Sub
   End If
   Cancel = False
End If

End Sub

Private Sub textCP54_GotFocus()
InverseTextBox textCP54
End Sub

Private Sub textCP54_lostfocus()
If Me.textCP53.Visible And Me.textCP54.Visible Then
   If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
      MsgBox Replace(Me.Label4(0).Caption, "：", "") & "輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.textCP53.SetFocus
      textCP53_GotFocus
   End If
End If
End Sub

Private Sub textCP54_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

    '申請國家為台灣
    If m_TM10 < "010" Then
        ' 檢核是否為民國日期
        If CheckIsTaiwanDate(Me.textCP54, False) = False Then
           Cancel = True
           strTit = "資料檢核"
           strMsg = "請輸入正確的日期"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCP54.SetFocus
           textCP54_GotFocus
           Exit Sub
        End If
    '申請國家非台灣
    Else
        ' 檢核是否為西元日期
        If CheckIsDate(Me.textCP54, False) = False Then
           Cancel = True
           strTit = "資料檢核"
           strMsg = "請輸入正確的日期"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCP54.SetFocus
           textCP54_GotFocus
           Exit Sub
        End If
    End If
If Val(Me.textCP54.Text) < Val(Me.textTM21.Text) Or Val(Me.textCP54.Text) > Val(Me.textTM22.Text) Then
   Cancel = True
   strTit = "資料檢核"
   strMsg = Replace(Me.Label4(0).Caption, "：", "") & "與專用期間不符, 是否重新輸入???" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "專用期間：" & Me.textTM21.Text & "－" & Me.textTM22.Text & Chr(10) & Chr(13) & Me.Label4(0).Caption & Me.textCP53.Text & "－" & Me.textCP54.Text
   nResponse = MsgBox(strMsg, vbYesNo, strTit)
   If nResponse = vbYes Then
      Me.textCP54.SetFocus
      textCP54_GotFocus
      Exit Sub
   End If
   Cancel = False
End If
End Sub

'Add By Sindy 2012/1/18
Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Sindy 2012/1/18
'商標局發函日
Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP64) = False Then
      If textCP64.MaxLength = 8 Then  'add by sonia 2019/1/4
         If CheckIsDate(textCP64, False) = False Then
            Cancel = True
            strTit = "檢核資料"
            strMsg = "大陸案發函日不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP64_GotFocus
            GoTo EXITSUB
         End If
      End If                          'add by sonia 2019/1/4
   End If
EXITSUB:
End Sub

'Add By Sindy 2009/07/17
Private Sub textFee_GotFocus()
   InverseTextBox textFee
End Sub
'領証費
Private Sub textFee_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Cancel = False
   If IsEmptyText(textFee) = False Then
      If IsNumeric(textFee) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入數字"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textFee_GotFocus
      End If
   End If
End Sub
Private Sub textNP08_GotFocus()
   InverseTextBox textNP08
End Sub
' 領証本所期限
Private Sub textNP08_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Cancel = False
   If IsEmptyText(textNP08) = False Then
      ' 檢查是否為民國年
      If CheckIsTaiwanDate(textNP08, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的領証本所期限"
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
Private Sub textNP09_GotFocus()
   InverseTextBox textNP09
End Sub
' 領証法定期限
Private Sub textNP09_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Cancel = False
   If IsEmptyText(textNP09) = False Then
      ' 檢查是否為民國年
      If CheckIsTaiwanDate(textNP09, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的領証法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNP09_GotFocus
      End If
   End If
End Sub
'2009/07/17 End

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
            strMsg = "只可輸入 N 或 1-3"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
      
      'Add By Sindy 2012/1/18
      '列印定稿為3(英文)時, 申請國家為大陸(020)者, 商標局發函日一定要輸入
      If (m_CP10 = "102" Or m_CP10 = "501") And Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" Then
         Label1(19).Visible = True
         textCP64.MaxLength = 8   'add by sonia 2019/1/4
         textCP64.Width = 2290    'add by sonia 2019/1/4
         textCP64.Visible = True
         If m_CP10 <> "101" Then '爭議案
            Label1(19).Caption = "大陸案受理函發文日 :                                                     (西元)"
         Else '申請案
            Label1(19).Caption = "　　　大陸案發函日 :                                                     (西元)"
         End If
      'add by sonia 2019/1/4  723出具同意書 T-214606
      ElseIf m_CP10 = "723" Then
         Label1(19).Visible = True
         Label1(19).Caption = "　　　　　進度備註 :"
         textCP64.MaxLength = 2000
         textCP64.Width = 3300
         textCP64.Visible = True
      'end 2019/1/4
      Else
         Label1(19).Visible = False
         textCP64.Visible = False
      End If
   End If
End Sub

' 列印備註
Private Sub textPS_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPS, 128) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPS_GotFocus
   End If
End Sub

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      ' 檢查是否為民國年
      If CheckIsTaiwanDate(textTM14, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
   End If
End Sub

'Add By Sindy 2010/8/31
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM15) = False Then
      '檢查審定號所輸入的長度是否正確
      '2012/3/28 modify by sonia 自撤核准不檢查 T-174298核駁訴願自撤
      'If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10) = False Then
      If m_CP10 <> "306" Then
         'Add By Sindy 2017/5/17 + strRetrunText
         If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
            Cancel = True
            textTM15_GotFocus
            Exit Sub
         'Add By Sindy 2017/5/17
         Else
            textTM15 = strRetrunText
         '2017/5/17 END
         End If
      End If
   End If
End Sub

' 專用期限起日
Private Sub textTM21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCorrDate As String
   Dim strDate As String
   
   Cancel = False
   ' 原專用期限止日
   If IsEmptyText(m_TM22) = True Then
      GoTo EXITSUB
   End If
   ' 未輸入專用期限起日
   If IsEmptyText(textTM21) = True Then
      GoTo EXITSUB
   End If
   ' 案件性質非延展
   If m_CP10 <> "102" Then
      GoTo EXITSUB
   End If
   
   If m_TM10 < "010" Then
      ' 檢核是否為民國日期
      If CheckIsTaiwanDate(textTM21, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
         GoTo EXITSUB
      End If
   Else
      ' 檢核是否為西元日期
      If CheckIsDate(textTM21, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM21_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   'Modify By Sindy 2012/10/18 非台灣案且有代理人提申日時則不必檢查,因在提申時已有修改TM22
   If m_TM10 <> "000" And m_CP47 <> "" Then
   Else
   '2012/10/18 End
       'Modify By Cheng 2003/09/01
   '   strCorrDate = DBDATE(Format(DateSerial(Val(DBYEAR(m_TM22)), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)) + 1)))
      'modify by sonia 2014/10/31
      'strCorrDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
      If m_TM01 <> "TF" Then
         strCorrDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(DBDATE(m_TM22))))
      Else
         strCorrDate = DBDATE(m_TM22)
      End If
      'end 2014/10/31
      strDate = DBDATE(textTM21)
      'strDate = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + 1)))
      
      If m_TM01 <> "TF" Then     'add by sonia 2014/10/31
         If strCorrDate <> strDate Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "專用期限起日必須為原專用期限止日的後一天"
              'Modify By Cheng 2002/11/08
              '若按確定, 仍可作業
      '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            nResponse = MsgBox(strMsg, vbOKCancel, strTit)
            If nResponse = vbOK Then Cancel = False: Exit Sub
            textTM21_GotFocus
         End If
      'add by sonia 2014/10/31
      Else
         If strCorrDate <> strDate Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "專用期限起日必須為原專用期限止日"
            nResponse = MsgBox(strMsg, vbOKCancel, strTit)
            If nResponse = vbOK Then Cancel = False: Exit Sub
            textTM21_GotFocus
         End If
      'end 2014/10/31
      End If
   End If
   
EXITSUB:
End Sub

' 專用期限止日
Private Sub textTM22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCorrDate As String
   Dim strDate As String
   Cancel = False
   
   ' 原專用期限止日
   If IsEmptyText(m_TM22) = True Then
      GoTo EXITSUB
   End If
   ' 未輸入專用期限起日
   If IsEmptyText(textTM22) = True Then
      GoTo EXITSUB
   End If
   ' 案件性質非延展
   If m_CP10 <> "102" Then
      GoTo EXITSUB
   End If
   
   If m_TM10 < "010" Then
      ' 檢核是否為民國日期
      If CheckIsTaiwanDate(textTM22, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22_GotFocus
         GoTo EXITSUB
      End If
   Else
      ' 檢核是否為西元日期
      If CheckIsDate(textTM22, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM22_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   'Modify By Sindy 2012/10/18 非台灣案且有代理人提申日時則不必檢查,因在提申時已有修改TM22
   If m_TM10 <> "000" And m_CP47 <> "" Then
   Else
   '2012/10/18 End
      strDate = DBDATE(textTM22)
      Select Case m_TM08
         Case "1", "4", "7", "8":
           'Modify By Cheng 2003/09/01
   '         strCorrDate = DBDATE(Format(DateSerial(Val(DBYEAR(m_TM22)) + Val(m_NA14), Val(DBMONTH(m_TM22)), Val(DBDAY(m_TM22)))))
            'Modified by Lydia 2019/11/13  改用共用模組, 並且因應商標案的算法,不抓NA85直接設「計算商標專用期是否減1天」=N
            'strCorrDate = DBDATE(DateAdd("yyyy", Val(m_NA14), ChangeWStringToWDateString(DBDATE(m_TM22))))
            ''2013/10/24 add by sonia T-134070 遇到 2/28 時，檢查 2/29 有的話已 2/29 為準
            'If Mid(DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strCorrDate))), 5) = "0229" Then
            '   strCorrDate = DBDATE(DateAdd("d", 1, ChangeWStringToWDateString(strCorrDate)))
            'End If
            ''2013/10/24 end
            'Modify By Sindy 2022/3/7 + m_TM10 : 延展後之專用期限年度倘有2月29日時，專用期限止日應為2月29日，而非以加10年之方式計算為2月28日
            strCorrDate = PUB_GetEndDate(DBDATE(m_TM22), Val(m_NA14), "N", m_TM10)
            'end 2019/11/13
         Case Else:
            strCorrDate = DBDATE(textTM22S)
      End Select
      If strDate <> strCorrDate Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "專用期限止日不正確"
           'Modify By Cheng 2002/11/08
           '若按確定, 仍可作業
   '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         nResponse = MsgBox(strMsg, vbOKCancel, strTit)
         If nResponse = vbOK Then Cancel = False: Exit Sub
         
         textTM22_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub textTM29_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否閉卷
Private Sub textTM29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textTM29) = False Then
      Select Case textTM29
         Case "", " ":
         Case "Y":
            strTit = "閉卷"
            strMsg = "請確認是否閉卷"
            nResponse = MsgBox(strMsg, vbYesNo, strTit)
            If nResponse = vbNo Then
               textTM29 = Empty
            End If
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入Y或空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM29_GotFocus
      End Select
   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   Me.SSTab1.Tab = 0 'Add By Sindy 2023/9/1
   'Add by Amy 2021/12/28檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

    'Add By Cheng 2002/12/11
'edit by nickc 2005/08/04
'    If m_CP10 = "301" Then
        'Modified by Lydia 2016/07/19  +判斷
        'If m_blnClkChgButton = False Then
        If m_blnClkChgButton = False And Me.cmdMod.Visible = True Then
            MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
            Me.cmdMod.SetFocus
            GoTo EXITSUB
        End If
'    End If
   
   ' 專用期限的起日不可超過迄日
   If Val(textTM21) > Val(textTM22) Then
      strTit = "資料檢核"
      strMsg = "專用期限的起日不可超過迄日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTM21.SetFocus
      GoTo EXITSUB
   End If
   ' 申請國家為大陸時, 核准通知日才可不輸
   If IsEmptyText(textCP25) = True Then
      If m_TM10 <> "020" Then
         strTit = "資料檢核"
         strMsg = "申請國家非大陸, 一定要輸入核准通知日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2012/6/18
   'TF案申請國家為美國時, 審定號不可空白
   If IsEmptyText(textTM15) = True Then
      If m_TM01 = "TF" And m_TM10 = "101" Then
         strTit = "資料檢核"
         strMsg = "審定號不可空白！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM15.SetFocus
         GoTo EXITSUB
      End If
   End If
   'TF案申請國家為歐盟,英國時, 公告日不可空白
   'modify by sonia 2024/9/4 +柬埔寨046(TF-00083-3-1-03)
   If IsEmptyText(textTM14) = True Then
      If m_TM01 = "TF" And (m_TM10 = "239" Or m_TM10 = "201" Or m_TM10 = "046") Then
         strTit = "資料檢核"
         strMsg = "公告日不可空白！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14.SetFocus
         GoTo EXITSUB
      End If
   End If
   'TF案申請國家為英國時, 期數不可空白
   If IsEmptyText(textTMBM07_2) = True Then
      If m_TM01 = "TF" And m_TM10 = "201" Then
         strTit = "資料檢核"
         strMsg = "期數不可空白！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_2.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2012/6/18 End
   
   'Modify By Cheng 2002/07/22
   '取消檢查
'   ' 是否更新基本檔目前准駁
'   If IsEmptyText(textTM16S) = True Then
'      strTit = "資料檢核"
'      strMsg = "請輸入是否更新基本檔目前准駁"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM16S.SetFocus
'      GoTo EXITSUB
'   End If
'   ' 專用權是否存在
'   If IsEmptyText(textTM17) = True Then
'      strTit = "資料檢核"
'      strMsg = "請輸入專用權是否存在"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM17.SetFocus
'      GoTo EXITSUB
'   End If
   ' 機關文號(申請國家為台灣時不可為空白)
   If IsEmptyText(textCP08) = True Then
      If m_TM10 < "010" Then
         strTit = "資料檢核"
         strMsg = "申請國家為台灣時機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2009/07/17
   '若為TF案申請國家為日本011或古巴135且案件性質為申請101時, 檢查之
   'Modify By Sindy 2022/9/22 桂英說,依林經理最後詢問代理人得知古巴已不需再繳註冊費，只有一筆100瑞郎的費用
   '古巴無須控管領證本所期限、法定期限及領證費 ==> 取消 Or m_TM10 = "135"
   'Modify By Sindy 2023/9/1 增加巴西117一定要輸領證期限及領證費
   If m_TM01 = "TF" And (m_TM10 = "011" Or m_TM10 = "117") And (m_CP10 = "101" Or m_CP10 = "104") Then
      If IsEmptyText(textNP08) = True Then
         strTit = "資料檢核"
         strMsg = "領証本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1 'Add By Sindy 2023/9/1
         textNP08.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textNP09) = True Then
         strTit = "資料檢核"
         strMsg = "領証法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1 'Add By Sindy 2023/9/1
         textNP09.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textFee) = True Then
         strTit = "資料檢核"
         strMsg = "領証費不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Me.SSTab1.Tab = 1 'Add By Sindy 2023/9/1
         textFee.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2009/07/17 End
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPS_GotFocus()
   InverseTextBox textPS
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Sub textTM14_GotFocus()
   InverseTextBox textTM14
End Sub

Private Sub textTM15_GotFocus()
   InverseTextBox textTM15
End Sub

Private Sub textTM16S_GotFocus()
   'Modify By Cheng 2002/07/22
'   InverseTextBox textTM16S
End Sub

Private Sub textTM17_GotFocus()
   'Modify By Cheng 2002/07/22
'   InverseTextBox textTM17
End Sub

Private Sub textTM21_GotFocus()
   InverseTextBox textTM21
End Sub

Private Sub textTM22_GotFocus()
   InverseTextBox textTM22
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textCP08_GotFocus()
   'Modify By Cheng 2002/04/22
   '將游標停在"字"的前面
'   InverseTextBox textCP08
Dim intPos As Integer
With Me.textCP08
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "字")
      If intPos - 1 >= 0 Then
         .SelStart = intPos - 1
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub textCP25_GotFocus()
   InverseTextBox textCP25
End Sub

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strTmp As String
'Add By Cheng 2003/07/18
Dim strSalesNo As String '智權人員
'add by nickc 2006/06/14 加入欠款資料
Dim A1kData As String
'2008/11/24 ADD BY SONIA 加入註冊費及服務費
Dim oRate As Double   '匯率
Dim o71713 As Double  '點數即服務費
Dim o71708 As Double  '規費
Dim o71613 As Double
Dim o71608 As Double
Dim o71513 As Double
Dim o71508 As Double
Dim intCnt As Integer  '2009/3/4 ADD BY SONIA 商品類別數
Dim strCurrTypeNM As String 'Add By Sindy 2010/11/18
Dim dbl_usxr02 As Double, intUsAmt As Integer 'Add By Sindy 2011/8/3
'Add By Sindy 2012/6/18
Dim rsTmp As New ADODB.Recordset
Dim strNA1 As String, strNA2 As String
'2012/6/18 End
Dim stA1k10 As Double 'Add By Sindy 2012/12/24
Dim dbl_OneAndTwoCase As Double 'Add By Sindy 2018/4/17
Dim strET03 As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   ' 商標狀況 -- 有發證日為註冊, 無發證日有審定號為審定, 均無為申請
   If IsEmptyText(m_TM20) = False Then
      strTmp = "註冊"
   ElseIf IsEmptyText(textTM15) = False Then
      strTmp = "審定"
   Else
      strTmp = "申請"
   End If
   
   'Add By Sindy 2012/6/18
   '取得領土延申指定國家及馬德里指定國家
   If m_TM01 = "TF" Then
      ' 取領土延伸指定國家
      strSql = "SELECT DISTINCT(TM10) FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM04 <> '00' AND (TM16 IS NULL OR TM16<>'2') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("TM10")) = False Then
               strTmp = GetNationName(rsTmp.Fields("TM10"), 0)
               If IsEmptyText(strTmp) = False Then
                  If strNA1 <> Empty Then: strNA1 = strNA1 & ","
                  strNA1 = strNA1 & strTmp
               End If
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
      ' 取馬德里指定國家
      strSql = "SELECT DISTINCT(TM10) FROM TradeMark " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "SUBSTR(TM02,1,5) = '" & Mid(m_TM02, 1, 5) & "' AND " & _
                     "TM04 <> '00' AND (TM16 IS NULL OR TM16<>'2') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            If IsNull(rsTmp.Fields("TM10")) = False Then
               strTmp = GetNationName(rsTmp.Fields("TM10"), 0)
               If IsEmptyText(strTmp) = False Then
                  If strNA2 <> Empty Then: strNA2 = strNA2 & ","
                  strNA2 = strNA2 & strTmp
               End If
            End If
            rsTmp.MoveNext
         Loop
      End If
      rsTmp.Close
   End If
   '2012/6/18 End
       
    Select Case m_TM01
    Case "T", "FCT", "TF"
       Select Case m_CP10
          ' 申請
          'edit by nick 2004/12/23 分割與申請做相同的事情
          'Case "101":
          Case "101", "308", "104":
             '2009/3/4 ADD BY SONIA 抓商品類別數
            intCnt = GetTMKindCnt(m_TM01, m_TM02, m_TM03, m_TM04)
            '2009/3/4 END
            ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                    'Modify By Cheng 2003/11/18
                    '若申請日小於921128
                    If Val(m_TM11) < 20031128 Then
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "03", m_CP09, "01", strUserNum
                        ' 卷數
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                                 "'" & "卷數" & "','" & textTMBM07_1 & "')"
                        cnnConnection.Execute strSql
                        ' 期數
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                                 "'" & "期數" & "','" & textTMBM07_2 & "')"
                        cnnConnection.Execute strSql
                        '91.6.12 modify by sonia
                        '' 列印備註
                        'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        '         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                        '         "'" & "列印備註" & "','" & textPS & "')"
                        'cnnConnection.Execute strSQL
                        '91.6.12 end
                    '若申請日大於等於921128
                    Else
                        'Modify By Sindy 2012/7/3 移上來共用
                        If m_blnReceiveFirst = False Then
                           '2008/11/24 ADD BY SONIA 定稿加註冊費及服務費
                           '外->台先不考慮匯率
                           'CheckOC3
                           'strSQL = "select * from usxrate where USXR01 in (select max(USXR01) from usxrate where USXR01<=to_number(to_char(sysdate, 'YYYYMMDD'))) "
                           'AdoRecordSet3.CursorLocation = adUseClient
                           'AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
                           'If AdoRecordSet3.RecordCount <> 0 Then
                           '   oRate = AdoRecordSet3.Fields("USXR02").Value
                           'End If
                           CheckOC3
                           'Modify By Sindy 2024/5/23 + 加抓 patentyearfee 設定
                           'Modify By Sindy 2024/7/31 + and yf15='新台幣'
                           strSql = "select cf03,cf08,cf13,2 as sort from casefee where cf01='" & m_TM01 & "' and cf02='" & m_TM10 & "'" & _
                                    " and cf03 in ('715','716','717')" & _
                                    " union select yf04,yf07,yf06/1000,1 as sort from patentyearfee where yf01='" & m_TM10 & "' and yf02='" & m_TM01 & "'" & _
                                    " and substr(yf03,1,8)='" & Left(m_TM23, 8) & "' and yf04='717' and yf15='新台幣'" & _
                                    " order by sort asc,cf03 asc "
                           AdoRecordSet3.CursorLocation = adUseClient
                           AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                           If AdoRecordSet3.RecordCount <> 0 Then
                              AdoRecordSet3.MoveFirst
                              Do While Not AdoRecordSet3.EOF
                                 '2009/3/4 MODIFY BY SONIA CF08規費要*商品類別數
                                 Select Case AdoRecordSet3.Fields("cf03").Value
                                 Case "715"
                                    o71508 = AdoRecordSet3.Fields("cf08").Value * intCnt
                                    o71513 = AdoRecordSet3.Fields("cf13").Value
                                 Case "716"
                                    o71608 = AdoRecordSet3.Fields("cf08").Value * intCnt
                                    o71613 = AdoRecordSet3.Fields("cf13").Value
                                 Case "717"
                                    o71708 = AdoRecordSet3.Fields("cf08").Value * intCnt
                                    o71713 = AdoRecordSet3.Fields("cf13").Value
                                    'Add By Sindy 2024/9/26
                                    If AdoRecordSet3.Fields("sort").Value = 1 Then '已抓到設定檔離開
                                       Exit Do
                                    End If
                                    '2024/9/26 END
                                 Case Else
                                 End Select
                                 AdoRecordSet3.MoveNext
                              Loop
                           End If
                           'Modify By Sindy 2024/5/23 Mark,請人員建檔 patentyearfee
'                           'Add By Sindy 2010/02/11
'                           If Trim(m_TM23) = "X18321010" Then
'                              o71513 = 2.5
'                              o71613 = 2.5
'                              o71713 = 5
'                           End If
'                           '2010/02/11 End
'                           'add by sonia 2016/4/7
'                           If Left(m_TM23, 6) = "X73867" Then
'                              o71713 = 1
'                           End If
'                           'end 2016/4/7
'                           'modify by sonia 2019/1/25 +X13175010財團法人工業技術研究院的關係企業,且智權人員為吳碧梧70005者才改
'                           'modify by sonia 2023/1/30 吳碧梧70005改為薛力銘B0006
'                           If Left(m_TM23, 6) = "X13175" And PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) = "B0006" Then
'                              o71713 = 1
'                           End If
'                           'end 2019/1/25
'                            'add by sonia 2022/9/20 加X85860
'                           If Left(m_TM23, 6) = "X85860" Then
'                              o71713 = 2.5
'                           End If
'                           'end 2022/9/20
                           '2024/5/23 END
                           
                           CheckOC3
                        End If
                        '2012/7/3 End
                        'Add By Sindy 2012/8/22 加註 frm210138 也有此費用的計算,若有異動時,須一併改寫
                        '若為商標
                        If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                            '若已收第一期註冊費
                            If m_blnReceiveFirst = True Then
                                ' 清除定稿例外欄位檔原有資料
                                EndLetter "03", m_CP09, "04", strUserNum
                                ' 卷數
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                                         "'" & "卷數" & "','" & textTMBM07_1 & "')"
                                cnnConnection.Execute strSql
                                ' 期數
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                                         "'" & "期數" & "','" & textTMBM07_2 & "')"
                                cnnConnection.Execute strSql
                            '若未收第一期註冊費
                            Else
                              'Modify By Sindy 2012/5/29 商標修法
'                              If Val(DBDATE(m_CP05)) >= 20120701 Then
                                 ' 清除定稿例外欄位檔原有資料
                                 EndLetter "03", m_CP09, "16", strUserNum
                                 ' 案件回覆單
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & "'," & _
                                          "'下一程序','717')"
                                 cnnConnection.Execute strSql
                                 ' 卷數
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & "'," & _
                                          "'" & "卷數" & "','" & textTMBM07_1 & "')"
                                 cnnConnection.Execute strSql
                                 ' 期數
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & "'," & _
                                          "'" & "期數" & "','" & textTMBM07_2 & "')"
                                 cnnConnection.Execute strSql
                                 ' 加入欠款資料
                                 A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                                 If A1kData <> "" Then
                                     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf  '& vbCrLf
                                     '巨京商標(96030)的客戶不出款項
                                     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                                 End If
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & "'," & _
                                          "'" & "欠款資料" & "','" & A1kData & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & _
                                          "','本所期限','" & m_NP08 & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & _
                                          "','錢717','" & Format(o71713 * 1000 + o71708, "###,###,##0") & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "16" & "','" & strUserNum & _
                                          "','錢71708','" & Format(o71708, "###,###,##0") & "')"
                                 cnnConnection.Execute strSql
'                              Else
'                              '2012/5/29 End
'                                 ' 清除定稿例外欄位檔原有資料
'                                 'Modify By Sindy 2009/04/17 定稿別03處理狀況07改13
'                                 ' 清除定稿例外欄位檔原有資料
'                                 EndLetter "03", m_CP09, "13", strUserNum
'                                 'add by nickc 2008/04/25  案件回覆單
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                          "'下一程序','715')"
'                                 cnnConnection.Execute strSql
'                                 ' 卷數
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                          "'" & "卷數" & "','" & textTMBM07_1 & "')"
'                                 cnnConnection.Execute strSql
'                                 ' 期數
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                          "'" & "期數" & "','" & textTMBM07_2 & "')"
'                                 cnnConnection.Execute strSql
'                                 'add by nickc 2006/06/14 加入欠款資料
'                                 A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                                 If A1kData <> "" Then
'                                     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf & vbCrLf
'                                     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
'                                     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                                 End If
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & "'," & _
'                                          "'" & "欠款資料" & "','" & A1kData & "')"
'                                 cnnConnection.Execute strSql
'                                 '2009/04/17 End
'                                 'Add By Sindy 2012/4/20
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','本所期限','" & m_NP08 & "')"
'                                 cnnConnection.Execute strSql
'                                 '2012/4/20 End
'                                 'Modify By Sindy 2009/04/17 定稿別03處理狀況07改13
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢715','" & Format(o71513 * 1000 + o71508, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢716','" & Format(o71613 * 1000 + o71608, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢717','" & Format(o71713 * 1000 + o71708, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 'Add By Sindy 2009/04/17
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢71508','" & Format(o71508, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢71608','" & Format(o71608, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & _
'                                          "','錢71708','" & Format(o71708, "###,###,##0") & "')"
'                                 cnnConnection.Execute strSql
'                                 '2009/04/17 End
'                                 '2008/11/24 END
'                              End If
                           End If
                        '若為標章
                        Else
                            '若已收第一期註冊費
                            If m_blnReceiveFirst = True Then
                                ' 清除定稿例外欄位檔原有資料
                                EndLetter "03", m_CP09, "05", strUserNum
                                ' 卷數
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                                         "'" & "卷數" & "','" & textTMBM07_1 & "')"
                                cnnConnection.Execute strSql
                                ' 期數
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                                         "'" & "期數" & "','" & textTMBM07_2 & "')"
                                cnnConnection.Execute strSql
                            '若未收第一期註冊費
                            Else
                              'Modify By Sindy 2012/7/3 商標修法
'                              If Val(DBDATE(m_CP05)) >= 20120701 Then
                                 ' 清除定稿例外欄位檔原有資料
                                 EndLetter "03", m_CP09, "18", strUserNum
                                 ' 案件回覆單
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & "'," & _
                                          "'下一程序','717')"
                                 cnnConnection.Execute strSql
                                 ' 卷數
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & "'," & _
                                          "'" & "卷數" & "','" & textTMBM07_1 & "')"
                                 cnnConnection.Execute strSql
                                 ' 期數
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & "'," & _
                                          "'" & "期數" & "','" & textTMBM07_2 & "')"
                                 cnnConnection.Execute strSql
                                 ' 加入欠款資料
                                 A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                                 If A1kData <> "" Then
                                     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                                     '巨京商標(96030)的客戶不出款項
                                     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                                 End If
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & "'," & _
                                          "'" & "欠款資料" & "','" & A1kData & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
                                          "','本所期限','" & m_NP08 & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
                                          "','錢717','" & Format(o71713 * 1000 + o71708, "###,###,##0") & "')"
                                 cnnConnection.Execute strSql
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "18" & "','" & strUserNum & _
                                          "','錢71708','" & Format(o71708, "###,###,##0") & "')"
                                 cnnConnection.Execute strSql
'                              Else
'                              '2012/7/3 End
'                                 ' 清除定稿例外欄位檔原有資料
'                                 EndLetter "03", m_CP09, "08", strUserNum
'                                 'add by nickc 2008/04/25  案件回覆單
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                          "'下一程序','715')"
'                                 cnnConnection.Execute strSql
'                                 ' 卷數
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                          "'" & "卷數" & "','" & textTMBM07_1 & "')"
'                                 cnnConnection.Execute strSql
'                                 ' 期數
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                          "'" & "期數" & "','" & textTMBM07_2 & "')"
'                                 cnnConnection.Execute strSql
'                                 'add by nickc 2006/06/14 加入欠款資料
'                                 A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                                 If A1kData <> "" Then
'                                     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf & vbCrLf
'                                     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
'                                     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                                 End If
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
'                                          "'" & "欠款資料" & "','" & A1kData & "')"
'                                 cnnConnection.Execute strSql
'
''                                 '2008/11/24 ADD BY SONIA 定稿加註冊費及服務費
''                                 '外->台先不考慮匯率
''                                  'CheckOC3
''                                  'strSQL = "select * from usxrate where USXR01 in (select max(USXR01) from usxrate where USXR01<=to_number(to_char(sysdate, 'YYYYMMDD'))) "
''                                  'AdoRecordSet3.CursorLocation = adUseClient
''                                  'AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
''                                  'If AdoRecordSet3.RecordCount <> 0 Then
''                                  '    oRate = AdoRecordSet3.Fields("USXR02").Value
''                                  'End If
''                                  'CheckOC3
''                                  strSql = "select * from casefee where cf01='" & m_TM01 & "' and cf02='" & m_TM10 & "' and cf03 in ('715','716','717') order by cf03 "
''                                  AdoRecordSet3.CursorLocation = adUseClient
''                                  AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
''                                  If AdoRecordSet3.RecordCount <> 0 Then
''                                      AdoRecordSet3.MoveFirst
''                                      Do While Not AdoRecordSet3.EOF
''                                          '2009/3/4 MODIFY BY SONIA CF08規費要*商品類別數
''                                          Select Case AdoRecordSet3.Fields("cf03").Value
''                                          Case "715"
''                                              o71508 = AdoRecordSet3.Fields("cf08").Value * intCnt
''                                              o71513 = AdoRecordSet3.Fields("cf13").Value
''                                          Case "716"
''                                              o71608 = AdoRecordSet3.Fields("cf08").Value * intCnt
''                                              o71613 = AdoRecordSet3.Fields("cf13").Value
''                                          Case "717"
''                                              o71708 = AdoRecordSet3.Fields("cf08").Value * intCnt
''                                              o71713 = AdoRecordSet3.Fields("cf13").Value
''                                          Case Else
''                                          End Select
''                                          AdoRecordSet3.MoveNext
''                                      Loop
''                                  End If
''                                  CheckOC3
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢715','" & Format(o71513 * 1000 + o71508, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢716','" & Format(o71613 * 1000 + o71608, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢717','" & Format(o71713 * 1000 + o71708, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  'Add By Sindy 2009/04/17
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢71508','" & Format(o71508, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢71608','" & Format(o71608, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','錢71708','" & Format(o71708, "###,###,##0") & "')"
'                                  cnnConnection.Execute strSql
'                                  '2009/04/17 End
'                                  '2008/11/24 END
'                                  'Add By Sindy 2012/4/20
'                                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & _
'                                           "','本所期限','" & m_NP08 & "')"
'                                  cnnConnection.Execute strSql
'                                  '2012/4/20 End
'                              End If
                            End If
                        End If
                    End If
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                    'Modify By Cheng 2003/11/18
                    '若申請日小於921128
                    If Val(m_TM11) < 20031128 Then
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "03", m_CP09, "02", strUserNum
                        'add by nickc 2006/06/14 加入欠款資料
                        A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                        If A1kData <> "" Then
                           A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                           'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                           If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                        End If
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                                 "'" & "欠款資料" & "','" & A1kData & "')"
                        cnnConnection.Execute strSql
                    '若申請日大於等於921128
                    Else
                        '若已收第一期註冊費
                        If m_blnReceiveFirst = True Then
                            ' 清除定稿例外欄位檔原有資料
                            EndLetter "03", m_CP09, "06", strUserNum
                            'add by nickc 2006/06/14 加入欠款資料
                            A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                            If A1kData <> "" Then
                                 A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                                 'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                                 If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                            End If
                            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                                     "'" & "欠款資料" & "','" & A1kData & "')"
                            cnnConnection.Execute strSql
                        '若未收第一期註冊費
                        Else
                           'Modify By Sindy 2012/5/29 商標修法
'                           If Val(DBDATE(m_CP05)) >= 20120701 Then
                              ' 清除定稿例外欄位檔原有資料
                              EndLetter "03", m_CP09, "17", strUserNum
'                              If Left(Trim(m_TM44), 6) = "Y51566" Then
'                                 strCurrTypeNM = "新台幣"
'                              Else
'                                 strCurrTypeNM = "人民幣"
'                              End If
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
'                                       "'幣別種類','" & strCurrTypeNM & "')"
'                              cnnConnection.Execute strSql
                              ' 加入欠款資料
                              A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                              If A1kData <> "" Then
                                 'Modified by Morgan 2022/12/21 取消跳行符號
                                 A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" '& vbCrLf '& vbCrLf
                                 ' 巨京商標(96030)的客戶不出款項
                                 If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                              End If
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                       "'" & "欠款資料" & "','" & A1kData & "')"
                              cnnConnection.Execute strSql
                              
                              'Add By Sindy 2024/7/31
                              CheckOC3
                              strSql = "select yf03,yf06,yf15 from patentyearfee" & _
                                       " where yf01='" & m_TM10 & "' and yf02='" & m_TM01 & "'" & _
                                       " and (substr(yf03,1,8)='" & Left(m_TM44, 8) & "' or substr(yf03,1,8)='" & Left(m_TM23, 8) & "')" & _
                                       " and yf04='717' and yf15<>'新台幣'" & _
                                       " order by yf03 desc"
                              AdoRecordSet3.CursorLocation = adUseClient
                              AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                              If AdoRecordSet3.RecordCount <> 0 Then
                                 '每類 XXX 元*類別數
                                 dbl_OneAndTwoCase = AdoRecordSet3.Fields("yf06").Value * intCnt
                                 strCurrTypeNM = "" & AdoRecordSet3.Fields("yf15")
                              Else
                                 dbl_OneAndTwoCase = intCnt * 1000# 'Add By Sindy 2018/4/17
                                 strCurrTypeNM = "人民幣"
                              End If
                              'Add By Sindy 2024/9/26 因金額設定為0時,代表取消定稿上之註冊費用
                              If dbl_OneAndTwoCase > 0 Then '檢查有金額,才列印註冊費此段文字
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                          "'" & "註冊費用列印" & "','♀')"
                                 cnnConnection.Execute strSql
                              End If
                              '2024/9/26 END
                              
                              'UBound(tmptextTM09)改抓intCnt
                              'Modify By Sindy 2013/1/22
'                              If Left(Trim(m_TM44), 6) = "Y51566" Then
                                 'modify by sonia 2018/8/28 改每類人民幣800元,依此類推T-212285
                                 'If intCnt > 1 Then '1類以上
                                 '   '服務費2000規費2500*類別數
                                 '   dbl_OneAndTwoCase = (intCnt * 2500#) + 2000 'Add By Sindy 2018/4/17
                                 '   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 '            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                 '            "'" & "OneAndTwoCase" & "','新台幣" & Trim(dbl_OneAndTwoCase) & "')"
                                 '   cnnConnection.Execute strSql
                                 'Else
                                 '   '服務費1000規費2500
                                 '   dbl_OneAndTwoCase = 1000 + 2500# 'Add By Sindy 2018/4/17
                                 '   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 '            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                 '            "'" & "OneAndTwoCase" & "','新台幣" & Trim(dbl_OneAndTwoCase) & "')"
                                 '   cnnConnection.Execute strSql
                                 'End If
'                                 '每類人民幣800元*類別數
'                                 dbl_OneAndTwoCase = intCnt * 800#
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                          "'" & "OneAndTwoCase" & "','" & strCurrTypeNM & Trim(dbl_OneAndTwoCase) & "')"
                                 cnnConnection.Execute strSql
                                 'end 2018/8/28
'                              Else
''                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
''                                          "'" & "OneCase" & "','人民幣" & Trim(intCnt * 650#) & "')"
''                                 cnnConnection.Execute strSql
''                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
''                                          "'" & "TwoCase" & "','人民幣" & Trim(intCnt * 775#) & "')"
''                                 cnnConnection.Execute strSql
'                                 dbl_OneAndTwoCase = intCnt * 1000# 'Add By Sindy 2018/4/17
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
'                                          "'" & "OneAndTwoCase" & "','人民幣" & Trim(dbl_OneAndTwoCase) & "')"
'                                 cnnConnection.Execute strSql
'                              End If
                              '2013/1/22 End
                              ' 案件回覆單
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                       "'下一程序','717')"
                              cnnConnection.Execute strSql
                              If Left(m_TM44, 6) = "Y46083" Or Left(m_TM44, 6) = "Y27598" Then
'                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
'                                          "'代繳註冊費','本所將代為繳納全期註冊費人民幣1000元。')"
                                 'Modify By Sindy 2012/6/27 商標修法全期字樣拿掉
                                 'Modify By Sindy 2018/4/17 1000 ==> Trim(dbl_OneAndTwoCase)
                                 strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "17" & "','" & strUserNum & "'," & _
                                          "'代繳註冊費','本所將代為繳納註冊費" & strCurrTypeNM & Trim(dbl_OneAndTwoCase) & "元。')"
                                 '2012/6/27 End
                                 cnnConnection.Execute strSql
                              End If
                              '2024/7/31 END
'                           Else
'                           '2012/5/29 End
'                            ' 清除定稿例外欄位檔原有資料
'                            'edit by nickc 2005/03/31 若是英文，抓另一定稿
'                            'EndLetter "03", m_CP09, "09", strUserNum
''edit by nickc 2006/12/15
''                            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
''                            Case "2"
''                                    EndLetter "03", m_CP09, "10", strUserNum
''                                    'add by nickc 2006/06/14 加入欠款資料
''                                    A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
''                                    If A1kData <> "" Then
''                                        A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf & vbCrLf
''                                    End If
''                                    strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                             "'" & "欠款資料" & "','" & A1kData & "')"
''                                    cnnConnection.Execute strSQL
''                            Case Else
'                                    ' 清除定稿例外欄位檔原有資料
'                                    EndLetter "03", m_CP09, "14", strUserNum
''                                    'Add By Sindy 2010/11/18
''                                    If Left(Trim(m_TM44), 6) = "Y51566" Then
''                                       strCurrTypeNM = "新台幣"
''                                    Else
''                                       strCurrTypeNM = "人民幣"
''                                    End If
''                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
''                                             "'幣別種類','" & strCurrTypeNM & "')"
''                                    cnnConnection.Execute strSql
'                                    '2009/04/10 End
'
'                                    'Add By Sindy 2009/04/10  案件回覆單
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                             "'下一程序','715')"
'                                    cnnConnection.Execute strSql
'                                    '2009/04/10 End
'
'                                    'Add By Sindy 2010/11/15
'                                    If Left(m_TM44, 6) = "Y46083" Or Left(m_TM44, 6) = "Y27598" Then
''                                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
''                                                "'代繳註冊費','本所將代為繳納全期註冊費人民幣1000元。')"
'                                       'Modify By Sindy 2012/6/27 商標修法全期字樣拿掉
'                                       strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                "'代繳註冊費','本所將代為繳納註冊費人民幣1000元。')"
'                                       '2012/6/27 End
'                                       cnnConnection.Execute strSql
'                                    End If
'                                    '2010/11/15
'
'                                    'add by nickc 2006/06/14 加入欠款資料
'                                    'Modify By Sindy 2009/04/17 定稿別03處理狀況09改14
'                                    A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                                    If A1kData <> "" Then
'                                       A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf & vbCrLf
'                                       'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
'                                       If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                                    End If
'                                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                             "'" & "欠款資料" & "','" & A1kData & "')"
'                                    cnnConnection.Execute strSql
'                                    'Dim tmptextTM09 As Variant             '2009/9/10 CANCEL BY SONIA 加抓intCnt
'                                    'tmptextTM09 = Split(textTM09, ",")     '2009/9/10 CANCEL BY SONIA 加抓intCnt
'                                    '2008/4/1 MODIFY BY SONIA 每一類美金第一期$82,第二期$100,全期$132,改為人民幣第一期$650,第二期$775,全期$1025
''                                    'edit by nickc 2008/04/07 依幣別
''                                    Select Case pub_GetCurrency(m_TM44)
''                                    Case "N"
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "OneCase" & "','NTD$" & Trim((UBound(tmptextTM09) + 1) * 2500#) & "')"
''                                            cnnConnection.Execute strSQL
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "TwoCase" & "','NTD$" & Trim((UBound(tmptextTM09) + 1) * 3000#) & "')"
''                                            cnnConnection.Execute strSQL
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "OneAndTwoCase" & "','NTD$" & Trim((UBound(tmptextTM09) + 1) * 4000#) & "')"
''                                            cnnConnection.Execute strSQL
''                                    Case "R"
'                                            '2009/9/10 MODIFY BY SONIA UBound(tmptextTM09)改抓intCnt
'                                             'Modify By Sindy 2013/1/22
'                                             If Left(Trim(m_TM44), 6) = "Y51566" Then
'                                                If intCnt > 1 Then '1類以上
'                                                   '服務費2000規費2500*類別數
'                                                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                            "'" & "OneAndTwoCase" & "','新台幣" & Trim((intCnt * 2500#) + 2000) & "')"
'                                                   cnnConnection.Execute strSql
'                                                Else
'                                                   '服務費1000規費2500
'                                                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                            "'" & "OneAndTwoCase" & "','新台幣" & Trim(1000 + 2500#) & "')"
'                                                   cnnConnection.Execute strSql
'                                                End If
'                                             Else
'                                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                         "'" & "OneCase" & "','人民幣" & Trim(intCnt * 650#) & "')"
'                                                cnnConnection.Execute strSql
'                                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                         "'" & "TwoCase" & "','人民幣" & Trim(intCnt * 775#) & "')"
'                                                cnnConnection.Execute strSql
'                                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                                         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
'                                                         "'" & "OneAndTwoCase" & "','人民幣" & Trim(intCnt * 1000#) & "')"
'                                                cnnConnection.Execute strSql
'                                             End If
'                                             '2013/1/22 End
''                                    Case "U"
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "OneCase" & "','USD$" & Trim((UBound(tmptextTM09) + 1) * 82#) & "')"
''                                            cnnConnection.Execute strSQL
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "TwoCase" & "','USD$" & Trim((UBound(tmptextTM09) + 1) * 100#) & "')"
''                                            cnnConnection.Execute strSQL
''                                            strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                                     "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
''                                                     "'" & "OneAndTwoCase" & "','USD$" & Trim((UBound(tmptextTM09) + 1) * 132#) & "')"
''                                            cnnConnection.Execute strSQL
''                                    End Select
'                                    '2008/4/1 END
''                            End Select
'                           End If
                        End If
                    End If
                   '91.6.12 modify by sonia
                   '' 列印備註
                   'strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                   '         "'" & "列印備註" & "','" & textPS & "')"
                   'cnnConnection.Execute strSQL
                   '91.6.12 end
                'add by nickc 2006/12/15
                ElseIf textPrint = "3" Then
                     'Add By Sindy 2012/12/24
                     '人民幣匯率
                     stA1k10 = PUB_GetUSXRate_1(strSrvDate(2), "RMB")
                     '2012/12/24 End
                     'Modify By Sindy 2012/9/3 商標修法
'                     If Val(DBDATE(m_CP05)) >= 20120701 Then
                        EndLetter "03", m_CP09, "19", strUserNum
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & _
                                 "','本所期限','" & m_NP08 & "')"
                        cnnConnection.Execute strSql
'2015/8/6 cancel by sonia 葉特助說金杜改回NTD及美金T-195507
'                        'TM44=Y51817000時改用人民幣,其他代理人仍用美金
'                        If Left(Trim(m_TM44), 6) = "Y51817" Then
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
'                                    "'" & "金額1" & "','RMB" & Format(Fix(5500 / stA1k10), "##,##0") & " per')" '原1000 Modify By Sindy 2012/12/24
'                           cnnConnection.Execute strSql
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
'                                    "'" & "金額2" & "','RMB" & Format(Fix(2500 / stA1k10), "##,##0") & " per')" '原625 Modify By Sindy 2012/12/24
'                           cnnConnection.Execute strSql
'                           'Add by Sindy 2015/6/22
'                           '有跨類時要列出
'                           If InStr(Trim(textTM09), ",") > 0 Then
'                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
'                                       "'" & "金額3" & "','　 　              RMB" & Format(Fix(3000 / stA1k10), "##,##0") & " per for each additional class.')"
'                              cnnConnection.Execute strSql
'                           End If
'                           '2015/6/22 End
'                        Else
                           strSql = "select usxr02 from usxrate where usxr01=(select max(usxr01) from usxrate where usxr01<=" & strSrvDate(2) & ")"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           dbl_usxr02 = 0: intUsAmt = 0
                           If intI = 1 Then
                              If Not IsNull(RsTemp.Fields(0)) Then
                                 dbl_usxr02 = RsTemp.Fields(0)
                              End If
                           End If
                           'Modify By Sindy 2015/8/21 T-197029 的請款單算出來US185,定稿是US186,因此做下列調整
                           'If dbl_usxr02 > 0 Then intUsAmt = Format(5500 / dbl_usxr02, "##,##0") '原4000 Modify By Sindy 2012/12/24
                           If dbl_usxr02 > 0 Then intUsAmt = Format(Int(3000 / dbl_usxr02) + Int(2500 / dbl_usxr02), "##,##0") '原4000 Modify By Sindy 2012/12/24
                           '2015/8/21 END
                           'Modify by Sindy 2013/5/22
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
'                                    "'" & "金額1" & "','US$" & intUsAmt & " for one')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                                    "'" & "金額1" & "','NT$5,500 (about US$" & intUsAmt & ")')"
                           '2013/5/22 End
                           cnnConnection.Execute strSql
                           If dbl_usxr02 > 0 Then intUsAmt = Format(2500 / dbl_usxr02, "##,##0")
                           'Modify by Sindy 2013/5/22
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
'                                    "'" & "金額2" & "','US$" & intUsAmt & " for each one')"
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                                    "'" & "金額2" & "','NT$2,500 (about US$" & intUsAmt & ")')"
                           '2013/5/22 End
                           cnnConnection.Execute strSql
                           'Add by Sindy 2013/5/22
                           '有跨類時要列出
                           If InStr(Trim(textTM09), ",") > 0 Then
                              If dbl_usxr02 > 0 Then intUsAmt = Format(3000 / dbl_usxr02, "##,##0")
                              strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                       "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                                       "'" & "金額3" & "','　 　              NT$3,000 (about US$" & intUsAmt & ") for each additional class.')"
                              cnnConnection.Execute strSql
                           End If
                           '2013/5/22 End
                           
                           'Added by Morgan 2023/3/6 Disbursement雜費
                           If dbl_usxr02 > 0 Then intUsAmt = Format(500 / dbl_usxr02, "##,##0")
                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                                 "'" & "金額4" & "','NT$500 (about US$" & intUsAmt & ")')"
                           cnnConnection.Execute strSql
                           'end 2023/3/6
                                    
'                        End If
'                     Else
'                     '2012/9/3 End
'                        EndLetter "03", m_CP09, "10", strUserNum
'                        'add by nickc 2006/06/14 加入欠款資料
'                        A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
'                        If A1kData <> "" Then
'                           A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf & vbCrLf
'                           'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
'                           If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
'                        End If
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                 "'" & "欠款資料" & "','" & A1kData & "')"
'                        cnnConnection.Execute strSql
'                        'Add By Sindy 2012/4/20
'                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & _
'                                 "','本所期限','" & m_NP08 & "')"
'                        cnnConnection.Execute strSql
'                        '2012/4/20 End
''2015/8/6 cancel by sonia 葉特助說金杜改回NTD及美金T-195507
''                        'Add By Sindy 2011/5/27 TM44=Y51817000時改用人民幣,其他代理人仍用美金
''                        If Left(Trim(m_TM44), 6) = "Y51817" Then
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額1" & "','RMB" & Format(Fix(5500 / stA1k10), "##,##0") & " per')" '原1000 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額2" & "','RMB" & Format(Fix(2500 / stA1k10), "##,##0") & " per')" '原625 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額3" & "','RMB" & Format(Fix(5500 / stA1k10), "##,##0") & " per')" '原1000 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額4" & "','RMB" & Format(Fix(1000 / stA1k10), "##,##0") & " per')" '原250 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額5" & "','RMB" & Format(Fix(3500 / stA1k10), "##,##0") & " per')" '原1000 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
''                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
''                                    "'" & "金額6" & "','RMB" & Format(Fix(1500 / stA1k10), "##,##0") & " per')" '原375 Modify By Sindy 2012/12/24
''                           cnnConnection.Execute strSql
''                        Else
'                           'Add By Sindy 2011/8/3
'                           strSql = "select usxr02 from usxrate where usxr01=(select max(usxr01) from usxrate where usxr01<=" & strSrvDate(2) & ")"
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'                           dbl_usxr02 = 0: intUsAmt = 0
'                           If intI = 1 Then
'                              If Not IsNull(RsTemp.Fields(0)) Then
'                                 dbl_usxr02 = RsTemp.Fields(0)
'                              End If
'                           End If
'                           '2011/8/3 End
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(5500 / dbl_usxr02, "##,##0") '原4000 Modify By Sindy 2012/12/24
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額1" & "','US$" & intUsAmt & " for one')"
'                           cnnConnection.Execute strSql
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(2500 / dbl_usxr02, "##,##0")
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額2" & "','US$" & intUsAmt & " for each one')"
'                           cnnConnection.Execute strSql
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(5500 / dbl_usxr02, "##,##0") '原4000 Modify By Sindy 2012/12/24
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額3" & "','US$" & intUsAmt & " for one')"
'                           cnnConnection.Execute strSql
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(1000 / dbl_usxr02, "##,##0")
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額4" & "','US$" & intUsAmt & " for each one')"
'                           cnnConnection.Execute strSql
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(3500 / dbl_usxr02, "##,##0") '原4000 Modify By Sindy 2012/12/24
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額5" & "','US$" & intUsAmt & " for one')"
'                           cnnConnection.Execute strSql
'                           If dbl_usxr02 > 0 Then intUsAmt = Format(1500 / dbl_usxr02, "##,##0")
'                           strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                                    "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
'                                    "'" & "金額6" & "','US$" & intUsAmt & " for each')"
'                           cnnConnection.Execute strSql
''                        End If
'                     End If
                End If
             ' 申請國家為大陸
             ElseIf m_TM10 = "020" Then
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", m_CP09, "03", strUserNum
                    '91.6.12 modify by sonia
                    'If IsEmptyText(m_CP19) = False Then
                    '   ' 後金備註
                    '   strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    '            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                    '            "'" & "後金備註" & "','" & "請先通知下次欲收之領證金額" & "')"
                    '   cnnConnection.Execute strSQL
                    'End If
                    'Modify By Cheng 2003/07/18
                    '取得消智權人員
                    strSalesNo = PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
                    '若為莊宏宇, 蘇玉峰, 余文達則要列印加註領證費
                    If strSalesNo = "80010" Or strSalesNo = "69010" Or strSalesNo = "76051" Then
                        ' 加註領證費
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                                 "'" & "加註領證費" & "','" & "請先通知下次欲收之領證金額" & "')"
                        cnnConnection.Execute strSql
                    End If
                    ' 期數
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                             "'" & "期數" & "','" & textTMBM07_2 & "')"
                    cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                    A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                    If A1kData <> "" Then
                        A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                        'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                        If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                    End If
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                             "'" & "欠款資料" & "','" & A1kData & "')"
                    cnnConnection.Execute strSql
                    '91.6.12 end
                'add by nickc 2007/07/24 加入英文
                ElseIf textPrint = "3" Then
                    EndLetter "03", m_CP09, "12", strUserNum
                    'Add By Sindy 2010/3/2
                    If m_CP10 = "101" Then
                        EndLetter "03", m_CP09, "15", strUserNum
                    End If
                    '2010/3/2 End
                End If
             
             'Add By Sindy 2009/07/17
             '申請國家為日本011或巴西117(巴西定稿同日本)
             ElseIf m_TM10 = "011" Or m_TM10 = "117" Then
                  strET03 = "01"
                  If m_TM01 = "TF" And (m_CP10 = "101" Or m_CP10 = "104") Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, strET03, strUserNum
                     ' 領證費
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                              "'" & "領證費" & "','" & Format(Val(textFee.Text), "##,##0") & "')"
                     cnnConnection.Execute strSql
                     ' 本所期限
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                              "'" & "本所期限" & "','" & DBDATE(textNP08) & "')"
                     cnnConnection.Execute strSql
                     ' 法定期限
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                              "'" & "法定期限" & "','" & DBDATE(textNP09) & "')"
                     cnnConnection.Execute strSql
                     ' 案件回覆單
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                              "'下一程序','701')"
                     cnnConnection.Execute strSql
                  End If
             '2009/07/17 End
             'Add By Sindy 2022/9/22 日本和古巴定稿要分開
             '申請國家為古巴135
             ElseIf m_TM10 = "135" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, "06", strUserNum
                  End If
             'Add By Sindy 2012/6/18
             '申請國家為歐盟239
             ElseIf m_TM10 = "239" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, "02", strUserNum
                  End If
             '申請國家為美國101
             ElseIf m_TM10 = "101" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, "03", strUserNum
                  End If
             '申請國家為韓國012
             ElseIf m_TM10 = "012" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, "04", strUserNum
                  End If
             '申請國家為英國201
             ElseIf m_TM10 = "201" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "03", m_CP09, "05", strUserNum
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                              "'" & "期數" & "','" & textTMBM07_2 & "')"
                     cnnConnection.Execute strSql
                  End If
             '2012/6/18 End
             End If
            'Modify By Cheng 2003/01/14
            '延展(102), 變更(301)
    '      ' 變更  2007/6/7 加減縮商品
    '      Case "301":
          'Modify By Sindy 2009/06/16 增加109.被異議續展
          'Case "102", "301", "313":
          Case "102", "301", "313", "109":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "04", strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2002/06/14
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "07", strUserNum
                    'Add By Cheng 2003/02/21
                    '加例外欄位--卷數及期數
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   ' '2009/04/17 End
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                   
                'add by sonia 2014/6/4
                ElseIf textPrint = "3" Then
                     EndLetter "03", m_CP09, "02", strUserNum
                     ' 其他公告日
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & _
                              "','其他公告日','" & DBDATE(Me.textTM14.Text) & "')"
                     cnnConnection.Execute strSql
                     EndLetter "03", m_CP09, "03", strUserNum
                     ' 其他公告日
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                              "','其他公告日','" & DBDATE(Me.textTM14.Text) & "')"
                     cnnConnection.Execute strSql
                     If m_TM08 = "7" Then '證明標章
                        ' 商標種類內文一
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                                 "','商標種類內文一','contents of certification')"
                        cnnConnection.Execute strSql
                     Else
                        ' 商標種類內文一
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & _
                                 "','商標種類內文一','specification of good/services')"
                        cnnConnection.Execute strSql
                     End If
                     'Added by Lydia 2017/04/21 增加定稿發函日期
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                              "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                     cnnConnection.Execute strSql
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                              "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                     cnnConnection.Execute strSql
                     'end 2017/04/21
               End If
             'Add By Cheng 2003/01/23
             '申請國家為大陸
             Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    'add by nickc 2006/09/15 加入續展
                    If m_TM01 = "TF" Then
                        EndLetter "03", m_CP09, "01", strUserNum
                        Dim otmpCountry As String
'edit by nickc 2007/02/15
'                        strSQL = "select * from caseprogress,trademark ,nation where tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm03<>'0' and tm04<>'00' and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp10='102' and tm10=na01(+) "
'                        CheckOC3
'                        AdoRecordSet3.CursorLocation = adUseClient
'                        AdoRecordSet3.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                        otmpCountry = ""
'                        If AdoRecordSet3.RecordCount <> 0 Then
'                            AdoRecordSet3.MoveFirst
'                            Do While Not AdoRecordSet3.EOF
'                                If otmpCountry <> "" Then
'                                    otmpCountry = otmpCountry & "、"
'                                End If
'                                otmpCountry = otmpCountry & CheckStr(AdoRecordSet3.Fields("na03"))
'                                AdoRecordSet3.MoveNext
'                            Loop
'                        End If
                    Dim otmpTm09 As Variant
                    Dim oII As Integer
                    otmpCountry = ""
                    otmpTm09 = Split(textTM09, ",")
                    For oII = 0 To UBound(otmpTm09)
                        '2009/4/21 MODIFY BY SONIA 領土延伸案的子案也要抓TF-000420
                        'strSQL = "select distinct tm03 ,na03 from nation,trademark,caseprogress where tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10=na01(+) and tm04<>'00' and tm03='" & Trim(oII + 1) & "' and tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "' and tm16='1' and tm29 is null   order by tm03 "
                        strSql = "select distinct tm03 ,na03 from nation,trademark,caseprogress where tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and tm10=na01(+) and tm04<>'00' and tm03='" & Trim(oII + 1) & "' and tm01='" & m_TM01 & "' and substr(tm02,1,5)='" & Mid(m_TM02, 1, 5) & "' and tm16='1' and tm29 is null   order by tm03 "
                        CheckOC3
                        AdoRecordSet3.CursorLocation = adUseClient
                        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                        If AdoRecordSet3.RecordCount <> 0 Then
                            otmpCountry = otmpCountry & "　　第 " & otmpTm09(oII) & "類："
                            AdoRecordSet3.MoveFirst
                            Do While Not AdoRecordSet3.EOF
                                If AdoRecordSet3.AbsolutePosition > 1 Then
                                    otmpCountry = otmpCountry & "、"
                                End If
                                otmpCountry = otmpCountry & CheckStr(AdoRecordSet3.Fields("na03"))
                                AdoRecordSet3.MoveNext
                            Loop
                            otmpCountry = otmpCountry & "。" & vbCrLf
                        End If
                    Next oII
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & _
                              "','例外國家欄位','" & otmpCountry & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                                 "'公告年','" & txtPrtY & "')"
                        cnnConnection.Execute strSql
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                                 "'公告期','" & txtPrtM & "')"
                        cnnConnection.Execute strSql
                    Else
                        ' 清除定稿例外欄位檔原有資料
                        EndLetter "03", m_CP09, "12", strUserNum
                        ' 列印備註
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                 "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & "'," & _
                                 "'" & "列印備註" & "','" & textPS & "')"
                        cnnConnection.Execute strSql
                    End If
                
                'Add By Sindy 2012/1/18
                ElseIf textPrint = "3" And m_TM10 = "020" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "00", strUserNum
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "01", strUserNum
                  ' 商標局發函日
                  If Trim(textCP64) <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','01','" & strUserNum & _
                              "','商標局發函日','" & ChgSQL(textCP64) & "')"
                     cnnConnection.Execute strSql
                  End If
                  'Added by Lydia 2017/04/21 增加定稿發函日期
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                           "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                           "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                  cnnConnection.Execute strSql
                  'end 2017/04/21
                End If
             End If
            
          'Add By Cheng 2003/01/01
          ' 補換發證書
          Case "103":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "00", strUserNum
                   ' 列印備註
                   
                   'Removed by Morgan 2022/12/29 已改寫在定稿內
                   'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   'cnnConnection.Execute strSql
                   'end 2022/12/29
                   
                   ' 附件(案件性質名稱)
                    'Modify By Cheng 2003/01/15
                    '案件性質為補換發證書的情形只有補發沒有換發
    '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
    '                        "'" & "附件" & "','" & IIf(Me.Check1(0).Value = vbChecked, "補發證書", "換發證書") & "')"
                   'Removed by Morgan 2022/12/29 已改寫在定稿內
                   'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                            "'" & "附件" & "','補發證書')"
                   'cnnConnection.Execute strSql
                   'end 2022/12/29
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "11", strUserNum
                   'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", m_CP09, "01", strUserNum
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                             "'" & "列印備註" & "','" & textPS & "')"
                    cnnConnection.Execute strSql
                End If
             End If
          ' 補正
          Case "201":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "10", strUserNum
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                   
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "11", strUserNum
                    'add by nickc 2006/06/14 加入欠款資料
                    'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                    'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                    'If A1kData <> "" Then
                    '    A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                    '    'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                    '    If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                    'End If
                    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    '         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                    '         "'" & "欠款資料" & "','" & A1kData & "')"
                    'cnnConnection.Execute strSql
                    'end 2016/12/22
                    
                    'Added by Lydia 2017/04/21 增加定稿發函日期
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                             "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
                    'end 2017/04/21
                End If
             End If
          ' 更正
          Case "302":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "05", strUserNum
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   ' 商標狀況
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "05" & "','" & strUserNum & "'," & _
                            "'" & "商標狀況" & "','" & strTmp & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "07", strUserNum
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2002/12/30
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                    '2009/04/17 End
                    
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             End If
          ' 英文證明, 中文證明
            'Modify By Cheng 2003/03/05
            '加中文證明
    '      Case "304":
          Case "304", "309":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "10", strUserNum
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "10" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "11", strUserNum
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "11" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             End If
          ' 自請撤回, 自請撤銷, 註銷
          'Case "306", "307":
          'Modify By Sindy 2016/8/10 + 626.註銷
          Case "306", "307", "626":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "06", strUserNum
                   ' 商標狀況
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                            "'" & "商標狀況" & "','" & strTmp & "')"
                   cnnConnection.Execute strSql
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                    'Modify By Cheng 2003/02/27
                    '改出定稿
    '               EndLetter "03", m_CP09, "07", strUserNum
                   EndLetter "03", m_CP09, "03", strUserNum
                   ' 機關文號
                    'Modify By Cheng 2003/02/27
    '               strSQL = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
    '                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "07" & "','" & strUserNum & "'," & _
    '                        "'" & "機關文號" & "','" & textCP08 & "')"
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "03" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
            'Add By Cheng 2003/12/22
            '申請國家為大陸
            Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", m_CP09, "06", strUserNum
                    ' 商標狀況
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                             "'" & "商標狀況" & "','" & strTmp & "')"
                    cnnConnection.Execute strSql
                    ' 機關文號
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                             "'" & "機關文號" & "','" & textCP08 & "')"
                    cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                    'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                    'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                    'If A1kData <> "" Then
                    '    A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                    '    'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                    '    If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                    'End If
                    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    '         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                    '         "'" & "欠款資料" & "','" & A1kData & "')"
                    'cnnConnection.Execute strSql
                    'end 2016/12/22
                    
                    'Added by Lydia 2017/04/21 增加定稿發函日期
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "06" & "','" & strUserNum & "'," & _
                             "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                    cnnConnection.Execute strSql
                    'end 2017/04/21
                End If
             End If
          ' 移轉
          Case "501":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "04", strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2002/06/14
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "08", strUserNum
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2003/01/02
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             'Add By Cheng 2003/01/23
             '申請國家為大陸
             Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", m_CP09, "12", strUserNum
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & "'," & _
                             "'" & "列印備註" & "','" & textPS & "')"
                    cnnConnection.Execute strSql
                
                'Add By Sindy 2012/1/18
                ElseIf textPrint = "3" And m_TM10 = "020" Then
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "00", strUserNum
                  ' 清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "01", strUserNum
                  ' 商標局發函日
                  If Trim(textCP64) <> "" Then
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "03" & "','" & m_CP09 & "','01','" & strUserNum & _
                              "','商標局發函日','" & ChgSQL(textCP64) & "')"
                     cnnConnection.Execute strSql
                  End If
                  'Added by Lydia 2017/04/21 增加定稿發函日期
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                           "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                  cnnConnection.Execute strSql
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                           "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                  cnnConnection.Execute strSql
                  'end 2017/04/21
                End If
             End If
          ' 授權
          Case "502":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "04", strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2002/06/14
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "08", strUserNum
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             'Add By Cheng 2003/06/19
             '申請國家為大陸
             Else
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "03", m_CP09, "00", strUserNum
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                             "'" & "案件性質" & "','許可合同備案')"
                    cnnConnection.Execute strSql
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                             "'" & "列印備註" & "','" & textPS & "')"
                    cnnConnection.Execute strSql
                End If
             End If
          'Add By Sindy 2012/5/17
          ' 廢止授權
          Case "503":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "04", strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "04" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "08", strUserNum
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                    '加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     '巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             End If
             '2012/5/17 End
          ' 再授權, 終止授權, 設定質權, 撤銷設定質權
          Case "504", "505", "506", "507":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "09", strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   'Add By Cheng 2002/06/14
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                   'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                   ' A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                   ' If A1kData <> "" Then
                   '     A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                   '     'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                   '     If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                   ' End If
                   ' strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   '          "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                   '          "'" & "欠款資料" & "','" & A1kData & "')"
                   ' cnnConnection.Execute strSql
                   'end 2016/12/22
                   
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "09" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
             End If
            'Add By Cheng 2004/02/09
          ' 第一期註冊費, 第二期註冊費
          Case "715", "716":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "13", strUserNum
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "13" & "','" & strUserNum & "'," & _
                             "'" & "列印備註" & "','" & textPS & "')"
                    cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "14", strUserNum
                    ' 列印備註
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                             "'" & "列印備註" & "','" & textPS & "')"
                    cnnConnection.Execute strSql
                    'add by nickc 2006/06/14 加入欠款資料
                    'Remove by Lydia 2016/12/22 內商之非爭議案核准或註冊證輸入,原本定稿的欠款資料改成D類收文控制
                    'A1kData = GetT_020_a1k_data(m_TM01, m_TM02, m_TM03, m_TM04)
                    'If A1kData <> "" Then
                    '    A1kData = "　本所迄今尚未收到本件商標" & A1kData & "，煩請儘速將上述款項擲寄本所，是祈！" & vbCrLf '& vbCrLf
                    '    'Modify By Sindy 2009/10/21 巨京商標(96030)的客戶不出款項
                    '    If m_CP13 = "96030" Then A1kData = "|\" & A1kData & "\|"
                    'End If
                    'strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                    '         "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                    '         "'" & "欠款資料" & "','" & A1kData & "')"
                    'cnnConnection.Execute strSql
                    'end 2016/12/22
                    
                   'Added by Lydia 2017/04/21 增加定稿發函日期
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                            "'" & "定稿發函日期" & "','" & strSrvDate(1) & "')"
                   cnnConnection.Execute strSql
                   'end 2017/04/21
                End If
            End If
            
          'Add By Sindy 2011/3/24
          Case "105"
            'TF宣誓核准定稿
            If m_TM01 = "TF" Then
               '清除定稿例外欄位檔原有資料
               EndLetter "03", m_CP09, "01", strUserNum
            End If
            
          'Add By Sindy 2012/6/18
          Case "104"
            'TF領土延伸公告定稿
            If m_TM01 = "TF" Then
               '清除定稿例外欄位檔原有資料
               EndLetter "03", m_CP09, "02", strUserNum
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "03" & "','" & m_CP09 & "','" & "02" & "','" & strUserNum & "'," & _
                        "'" & "領土延伸指定國家" & "','" & strNA1 & "')"
               cnnConnection.Execute strSql
            End If
            
          'add by sonia 2019/1/4 723出具同意書
          Case "723"
             If m_TM10 < "010" Then
                If textPrint = "1" Then
                  '清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "00", strUserNum
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                           "'" & "附註" & "','" & textCP64 & "')"
                  cnnConnection.Execute strSql
                End If
             End If
             
          'Added by Lydia 2024/04/01 代辦退費
          Case "725"
             If m_TM10 < "010" Then
                If textPrint = "1" Then
                  '清除定稿例外欄位檔原有資料
                  EndLetter "03", m_CP09, "00", strUserNum
                  strExc(1) = ""
                  If Mid(m_CP43, 1, 1) = "A" Then
                     strExc(1) = Pub_GetNoToCPM("1", m_CP43)
                  ElseIf m_CP43 <> "" Then
                     strExc(1) = Pub_GetNoToCPM("2", m_CP43, , "A")
                  End If
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "03" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
                           "'" & "來源案件性質" & "','" & IIf(strExc(1) <> "", strExc(1), "♀") & "')"
                  cnnConnection.Execute strSql
                End If
             End If
             
          'Add By Sindy 2021/8/25 729.復權
          Case "729"
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                If textPrint = "1" Then
                  'Modified by Morgan 2023/1/4 +標章
                  If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                     strET03 = "01"
                  Else
                     strET03 = "02"
                  End If
                  'end 2023/1/4
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, strET03, strUserNum
                   ' 卷數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                            "'" & "卷數" & "','" & textTMBM07_1 & "')"
                   cnnConnection.Execute strSql
                   ' 期數
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                            "'" & "期數" & "','" & textTMBM07_2 & "')"
                   cnnConnection.Execute strSql
                   ' 其他公告日
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                            "'" & "其他公告日" & "','" & DBDATE(Me.textTM14.Text) & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & strET03 & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                ' 申請人國籍非台灣
                ElseIf textPrint = "2" Then
                   ' 清除定稿例外欄位檔原有資料
                   EndLetter "03", m_CP09, "08", strUserNum
                   ' 機關文號
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "機關文號" & "','" & textCP08 & "')"
                   cnnConnection.Execute strSql
                   ' 列印備註
                   strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                            "VALUES ('" & "03" & "','" & m_CP09 & "','" & "08" & "','" & strUserNum & "'," & _
                            "'" & "列印備註" & "','" & textPS & "')"
                   cnnConnection.Execute strSql
                End If
             End If
       End Select
    Case "TC"
        'add by nickc 2006/06/30
        If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "03", m_CP09, "00", strUserNum
        End If
    End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
'2003/03/07--內商案--除非畫面上有"是否修改定稿欄", 否則一律不開Word
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean, ET03_1 As String
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   '2005/11/9 ADD BY SONIA
   '取得定稿語文
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "03"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/13 End
   
   Select Case m_TM01
    Case "T", "FCT", "TF"
       Select Case m_CP10
          ' 申請
          'edit by nick 2004/12/23 分割與申請做相同的事情
          'Case "101":
          Case "101", "308", "104":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'editby nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                    '若申請日小於921128
                    If Val(m_TM11) < 20031128 Then
                        ' 列印定稿
                        ET03 = "01" 'Modify By Sindy 2012/1/13
                    '若申請日大於等於921128
                    Else
                        '若為商標
                        If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                            '若已收第一期註冊費
                            If m_blnReceiveFirst = True Then
                                ' 列印定稿
                                 ET03 = "04" 'Modify By Sindy 2012/1/13
                            '若未收第一期註冊費
                            Else
                                ' 列印定稿
                                '2005/11/9 MODIFY BY SONIA 加入定稿語文判斷
'edit by nickc 2007/07/11 改成依照畫面上的出 且 11 與 10 重複
'                                 Select Case m_strLanguage
'                                 Case "1"  '中文
                                    'Modify By Sindy 2009/04/17 定稿別03處理狀況07改13
                                 'Modify By Sindy 2012/5/29 商標修法
'                                 If Val(DBDATE(m_CP05)) >= 20120701 Then
                                    'Memo by Lydia 2019/08/13 臺灣商申核准案改雙面列印
                                    'Added by Morgan 2022/12/30 註冊後分割
                                    If m_TM20 <> "" Then
                                       ET03 = "02"
                                    Else
                                    'end 2022/12/30
                                       ET03 = "16"
                                    End If
'                                 Else
'                                 '2012/5/29 End
'                                    ET03 = "13" 'Modify By Sindy 2012/1/13
'                                 End If
'                                 Case "2"  '英文
'                                    NowPrint m_CP09, "03", "11", False, strUserNum, 0
'                                 End Select
                                 '2005/11/9 END
                            End If
                        Else
                            '若已收第一期註冊費
                            If m_blnReceiveFirst = True Then
                                ' 列印定稿
                                 ET03 = "05" 'Modify By Sindy 2012/1/13
                            '若未收第一期註冊費
                            Else
                              'Modify By Sindy 2012/7/3 商標修法
'                              If Val(DBDATE(m_CP05)) >= 20120701 Then
                                 'Added by Morgan 2022/12/30 註冊後分割(證明標章)
                                 If m_TM20 <> "" Then
                                    ET03 = "06"
                                 Else
                                 'end 2022/12/30
                                    ET03 = "18" 'Memo by Lydia 2021/01/18 證明標章改雙面列印，第一面不用列印商品類別;ex.T-228294為證明標章，目前為單面會忘記去列印定稿。
                                 End If
'                              Else
'                              '2012/7/3 End
'                                ' 列印定稿
'                                 ET03 = "08" 'Modify By Sindy 2012/1/13
'                              End If
                            End If
                        End If
                    End If
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
'Removed by Morgan 2022/12/14 定稿已刪除 02,06,09,14
'                    '若申請日小於921128
'                    If Val(m_TM11) < 20031128 Then
'                        ' 列印定稿
'                        ET03 = "02" 'Modify By Sindy 2012/1/13
'                    '若申請日大於等於921128
'                    Else
'                        '若已收第一期註冊費
'                        If m_blnReceiveFirst = True Then
'                            ' 列印定稿
'                              ET03 = "06" 'Modify By Sindy 2012/1/13
'                        '若未收第一期註冊費
'                        Else
'end 2022/12/14
                            ' 列印定稿
                            'edit by nickc 2005/03/31 若是英文，改印另一定稿
'edit by nickc 2006/12/15
'                            Select Case GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
'                            Case "2"
'                                    NowPrint m_CP09, "03", "10", False, strUserNum, 0
'                            Case Else
                                    'Modify By Sindy 2009/04/17 定稿別03處理狀況09改14
                                    'NowPrint m_CP09, "03", "09", False, strUserNum, 0
'                                    NowPrint m_CP09, "03", "14", False, strUserNum, 0
                              'Modify By Sindy 2012/5/29 商標修法
'                              If Val(DBDATE(m_CP05)) >= 20120701 Then
                                 ET03 = "17"
'                              Else
'                              '2012/5/29 End
'                                 ET03 = "14" 'Modify By Sindy 2012/1/13
'                              End If
'                            End Select
'                        End If
'                    End If
                'add by nickc 2006/12/15
                ElseIf textPrint = "3" Then
                     'Modify By Sindy 2012/9/3 商標修法
'                     If Val(DBDATE(m_CP05)) >= 20120701 Then
                        ET03 = "19"
'                     Else
'                     '2012/9/3 End
'                        ET03 = "10" 'Modify By Sindy 2012/1/13
'                     End If
                End If
             ' 申請國家為大陸
             '92.6.23 modify by sonia 大陸案不管結果為何, 全部都要出定稿
             'ElseIf m_TM10 = "020" And frm02010401_3.textResult <> "2" Then
             ElseIf m_TM10 = "020" Then
             '92.6.23 end
             '92.1.29 end
                ' 列印定稿
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                     ET03 = "03" 'Modify By Sindy 2012/1/13
                'add by nickc 2007/07/24 加入英文
                ElseIf textPrint = "3" Then
                     ET03 = "12" 'Modify By Sindy 2012/1/13
                    'Add By Sindy 2010/3/2
                    If m_CP10 = "101" Then
                        ET03_1 = "15" 'Modify By Sindy 2012/1/13
                    End If
                    '2010/3/2 End
                End If
             
             'Modify By Sindy 2022/9/22 日本和古巴定稿要分開
             'Add By Sindy 2009/07/17
             '申請國家為日本011或古巴135
             '申請國家為日本011
             'Modify By Sindy 2023/9/1 巴西117(定稿同日本)
             ElseIf m_TM10 = "011" Or m_TM10 = "117" Then
                  If m_TM01 = "TF" And (m_CP10 = "101" Or m_CP10 = "104") Then
                     ET03 = "01" 'Modify By Sindy 2012/1/13
                  End If
             '申請國家為古巴135
             ElseIf m_TM10 = "135" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ET03 = "06"
                  End If
             '2022/9/22 END
             '2009/07/17 End
             'Add By Sindy 2012/6/18
             '申請國家為歐盟239
             ElseIf m_TM10 = "239" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ET03 = "02"
                  End If
             '申請國家為美國101
             ElseIf m_TM10 = "101" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ET03 = "03"
                  End If
             '申請國家為韓國012
             ElseIf m_TM10 = "012" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ET03 = "04"
                  End If
             '申請國家為英國201
             ElseIf m_TM10 = "201" Then
                  If m_TM01 = "TF" And m_CP10 = "101" Then
                     ET03 = "05"
                  End If
             '2012/6/18 End
             End If
            'Modify By Cheng 2003/01/14
            '延展(102), 變更(301)
    '      ' 變更 2007/6/7 加減縮商品
    '      Case "301":
          'Modify By Sindy 2009/06/16 增加109.被異議續展
          'Case "102", "301", "313":
          Case "102", "301", "313", "109":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "04" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "07" 'Modify By Sindy 2012/1/13
                'add by sonia 2014/6/4
                ElseIf textPrint = "3" Then
                     ET03 = "02"
                     ET03_1 = "03" '英譯本
                'end 2014/6/4
                End If
             'Add By Cheng 2003/01/23
             '申請國家為大陸
             Else
                ' 列印定稿
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                    'add by nickc 2006/09/15
                    If m_TM01 = "TF" Then
                        ET03 = "01" 'Modify By Sindy 2012/1/13
                    Else
                        ET03 = "12" 'Modify By Sindy 2012/1/13
                    End If
                    
                'Add By Sindy 2012/1/18
                ElseIf textPrint = "3" And m_TM10 = "020" Then
                   ET03 = "00"
                   ET03_1 = "01" '英譯本
                End If
             End If
            'Add By Cheng 2003/01/01
          ' 補證
          Case "103":
                ' 申請國家為台灣
                If m_TM10 < "010" Then
                   ' 申請人國籍為台灣
                   'edit by nickc 2006/06/30
                   'If strTM23Nation < "010" Then
                   If textPrint = "1" Then
                      ' 列印定稿
                        'Modified by Morgan 2022/12/30 +02
                        If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                           ET03 = "00" 'Modify By Sindy 2012/1/13
                        Else
                           ET03 = "02"
                        End If
                        'end 2022/12/30
                   ' 申請人國籍非台灣
                   'edit by nickc 2006/06/30
                   'Else
                   ElseIf textPrint = "2" Then
                      ' 列印定稿
                        ET03 = "11" 'Modify By Sindy 2012/1/13
                   End If
                'Add By Cheng 2003/06/12
                '申請國家為大陸
                Else
                    ' 列印定稿
                    'add by nickc 2006/06/30
                    If textPrint = "1" Then
                        ET03 = "01" 'Modify By Sindy 2012/1/13
                    End If
                End If
                
          'Added by Morgan 2023/1/3
          Case "314" '申請註冊證副本
            ' 申請國家為台灣
            If m_TM10 < "010" Then
               If textPrint = "1" Then
                  If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                     ET03 = "00"
                  Else
                     ET03 = "02"
                  End If
               'Add By Sindy 2024/10/15
               ElseIf textPrint = "2" Then
                  ET03 = "01"
               '2024/10/15 END
               End If
            End If
          'end 2023/1/3
          
          ' 更正
          Case "302":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
    '                'Modify By Cheng 2003/01/08
    '                '開啟定稿, 讓使用者修改
                     ET03 = "05" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "07" 'Modify By Sindy 2012/1/13
                End If
             End If
          ' 英文證明, 中文證明
            'Modify By Cheng 2003/03/06
            '加中文證明
    '      Case "304":
          Case "304", "309":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "10" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "11" 'Modify By Sindy 2012/1/13
                End If
             End If
          ' 自請撤回, 自請撤銷, 註銷
          'Case "306", "307":
          'Modify By Sindy 2016/8/10 + 626.註銷
          Case "306", "307", "626":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "06" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                    'Modify By Cheng 2003/02/27
                    '改出定稿
                     ET03 = "03" 'Modify By Sindy 2012/1/13
                End If
            'Add By Cheng 2003/12/22
            '申請國家為大陸
            Else
                ' 列印定稿
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                     ET03 = "06" 'Modify By Sindy 2012/1/13
                End If
             End If
          ' 移轉
          Case "501":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "04" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "08" 'Modify By Sindy 2012/1/13
                End If
             'Add By Cheng 2003/03/04
             '申請國家為大陸
             Else
                ' 列印定稿
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                     ET03 = "12" 'Modify By Sindy 2012/1/13
                     
                'Add By Sindy 2012/1/18
                ElseIf textPrint = "3" And m_TM10 = "020" Then
                   ET03 = "00"
                   ET03_1 = "01" '英譯本
                End If
             End If
          ' 授權
          Case "502":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "04" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "08" 'Modify By Sindy 2012/1/13
                End If
            'Add By Cheng 2003/06/19
            '申請國家為大陸
            Else
                ' 列印定稿
                'add by nickc 2006/06/30
                If textPrint = "1" Then
                     ET03 = "00" 'Modify By Sindy 2012/1/13
                End If
             End If
          'Add By Sindy 2012/5/17
          ' 廢止授權
          Case "503":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "04"
                ' 申請人國籍非台灣
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "08"
                End If
             End If
             '2012/5/17 End
          ' 再授權, 終止授權, 設定質權, 撤銷設定質權
          'edit by nickc 2005/03/31 打錯了，應該為 505 終止再授權
          'Case "504", "504", "506", "507":
          Case "504", "505", "506", "507":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "09" 'Modify By Sindy 2012/1/13
                End If
             End If
          'Add By Cheng 2004/02/09
          '第一期註冊費, 第二期註冊費
          Case "715", "716":
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                'edit by nickc 2006/06/30
                'If strTM23Nation < "010" Then
                If textPrint = "1" Then
                   ' 列印定稿
                     ET03 = "13" 'Modify By Sindy 2012/1/13
                ' 申請人國籍非台灣
                'edit by nickc 2006/06/30
                'Else
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "14" 'Modify By Sindy 2012/1/13
                End If
             End If
            'End
         
          'Add By Sindy 2011/3/24
          Case "105"
            'TF宣誓核准定稿
            If m_TM01 = "TF" Then
               '列印定稿
               '2014/1/28 modify by sonia
               'ET03 = "01" 'Modify By Sindy 2012/1/13
               If m_TM10 = "101" Then
                  ET03 = "02"
               Else
                  ET03 = "01"
               End If
               '2014/1/28 end
            End If
          '2011/3/24 End
          'Add By Sindy 2012/6/18
          Case "104"
            'TF領土延伸公告定稿
            If m_TM01 = "TF" Then
               '列印定稿
               ET03 = "02"
            End If
          '2012/6/18 End
          'add by sonia 2019/1/4 出具同意書 T-214606
          Case "723"
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                If textPrint = "1" Then
                   ET03 = "00"
                End If
             End If
          'Added by Lydia 2024/04/01 代辦退費
          Case "725"
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                If textPrint = "1" Then
                   ET03 = "00"
                End If
             End If
          'Add By Sindy 2021/8/25 729.復權
          Case "729"
             ' 申請國家為台灣
             If m_TM10 < "010" Then
                ' 申請人國籍為台灣
                If textPrint = "1" Then
                   ' 列印定稿
                     'Modified by Morgan 2023/1/4 +標章
                     If m_TM08 <> "6" And m_TM08 <> "7" And m_TM08 <> "8" Then
                        ET03 = "01"
                     Else
                        ET03 = "02"
                     End If
                     'end 2023/1/4
                ' 申請人國籍非台灣
                ElseIf textPrint = "2" Then
                   ' 列印定稿
                     ET03 = "08"
                End If
             End If
       End Select
    Case "TC"
         'add by nickc 2006/06/30
         If textPrint = "1" Then
            ET03 = "00" 'Modify By Sindy 2012/1/13
         End If
   End Select
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
            End If
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            If ET03_1 <> "" Then
               NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            End If
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/19 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         If ET03_1 <> "" Then
            'Add By Sindy 2019/12/19 + strLD18.信函總收文號
            NowPrint ET02, ET01, ET03_1, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
      End If
   Else
      'Add By Sindy 2021/1/5 沒有系統產出的定稿
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
      '2021/1/5 EMD
      
      BolPrintLetterDemand = False 'Add By Sindy 2012/4/16
      'Added by Lydia 2016/12/22 不出定稿,取消D類收文控制
      m_ULD02 = ""
      bolA1kdataMail = False
      'Modified by Lydia 2017/04/06
      'm_AC2470 = ""
      m_rA1k28 = ""
      m_rSpec = ""
      'end 2017/04/06
      'end 2016/12/22
   End If
   '2012/1/13 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim bPaper As Boolean

TxtValidate = False

'Add By Sindy 2010/12/24
If Me.textTM15.Enabled = True Then
   Cancel = False
   textTM15_Validate Cancel
   If Cancel = True Then
      textTM15.SetFocus
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

'Add By Sindy 2012/1/18
If Me.textCP64.Enabled = True Then
   Cancel = False
   textCP64_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   '列印定稿為3(英文)時, 申請國家為大陸(020)者, 商標局發函日一定要輸入
   If (m_CP10 = "102" Or m_CP10 = "501") And Trim(m_TM10) = "020" And Trim(textPrint.Text) = "3" Then
      If Trim(textCP64) = "" Then
         MsgBox "大陸案發函日不可空白！", vbExclamation, "資料檢核"
         textCP64.SetFocus
         Exit Function
      End If
   'add by sonia 2019/1/4
   ElseIf m_CP10 = "723" Then
      If Trim(textCP64) = "" Then
         MsgBox "出具同意書之核准，進度備註不可空白！", vbExclamation, "資料檢核"
         textCP64.SetFocus
         Exit Function
      End If
   'end 2019/1/4
   End If
End If

If Me.textPS.Enabled = True Then
   Cancel = False
   textPS_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM14.Enabled = True Then
   Cancel = False
   textTM14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM21.Enabled = True Then
   Cancel = False
   textTM21_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM22.Enabled = True Then
   Cancel = False
   textTM22_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textTM29.Enabled = True Then
   Cancel = False
   textTM29_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP53.Enabled = True And Me.textCP53.Visible = True Then
   Cancel = False
   textCP53_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP54.Enabled = True And Me.textCP54.Visible = True Then
   Cancel = False
   textCP54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
   If Me.textCP53.Visible And Me.textCP54.Visible Then
      If Val(Me.textCP53.Text) > Val(Me.textCP54.Text) Then
         MsgBox "日期區間輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.textCP53.SetFocus
         textCP53_GotFocus
         Exit Function
      End If
   End If
End If

'Added by Morgan 2022/1/3 E化案件提醒--桂英
If m_TM44 = "" Then
   If PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bPaper) = True And bPaper = False Then
      MsgBox "E化案件!!", vbExclamation
   End If
End If
'end 2022/1/3
   
TxtValidate = True
End Function

Private Sub txtNote_GotFocus()
    TextInverse Me.txtNote
End Sub
