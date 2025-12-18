VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_20 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(著作權)"
   ClientHeight    =   5760
   ClientLeft      =   5712
   ClientTop       =   1872
   ClientWidth     =   9156
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9156
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   6996
      TabIndex        =   27
      Top             =   10
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6168
      TabIndex        =   26
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8220
      TabIndex        =   28
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   350
      Left            =   4944
      TabIndex        =   25
      Top             =   10
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3660
      Left            =   144
      TabIndex        =   52
      Top             =   2064
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   6456
      _Version        =   393216
      Tab             =   1
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "第一頁"
      TabPicture(0)   =   "frm020102_20.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label25"
      Tab(0).Control(1)=   "Label31"
      Tab(0).Control(2)=   "Label30"
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(4)=   "Label7"
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(7)=   "Label10"
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(9)=   "Label12"
      Tab(0).Control(10)=   "Label13"
      Tab(0).Control(11)=   "Label14"
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(13)=   "Label16"
      Tab(0).Control(14)=   "Label17"
      Tab(0).Control(15)=   "Label18"
      Tab(0).Control(16)=   "Label19"
      Tab(0).Control(17)=   "Label20"
      Tab(0).Control(18)=   "Label1(10)"
      Tab(0).Control(19)=   "Label21"
      Tab(0).Control(20)=   "Label33"
      Tab(0).Control(21)=   "Label34"
      Tab(0).Control(22)=   "Label35"
      Tab(0).Control(23)=   "Label36"
      Tab(0).Control(24)=   "Label27"
      Tab(0).Control(25)=   "Label39"
      Tab(0).Control(26)=   "lblCP113(18)"
      Tab(0).Control(27)=   "textCP44_2"
      Tab(0).Control(28)=   "textSP51"
      Tab(0).Control(29)=   "textSP07"
      Tab(0).Control(30)=   "textSP05"
      Tab(0).Control(31)=   "textSP46"
      Tab(0).Control(32)=   "textSP44"
      Tab(0).Control(33)=   "textCP22"
      Tab(0).Control(34)=   "textCP44"
      Tab(0).Control(35)=   "textSP06"
      Tab(0).Control(36)=   "textCP27"
      Tab(0).Control(37)=   "textSP38"
      Tab(0).Control(38)=   "textSP39"
      Tab(0).Control(39)=   "textSP40"
      Tab(0).Control(40)=   "textSP63"
      Tab(0).Control(41)=   "textSP47"
      Tab(0).Control(42)=   "textSP48"
      Tab(0).Control(43)=   "textCP18"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textPrint"
      Tab(0).Control(45)=   "textWord"
      Tab(0).Control(46)=   "textCP84"
      Tab(0).Control(47)=   "txtCP113"
      Tab(0).ControlCount=   48
      TabCaption(1)   =   "第二頁"
      TabPicture(1)   =   "frm020102_20.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label22"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label23"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label28"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label29"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label32"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "textSP41"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "textSP42"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "textSP43"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "textSP45"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textCP64"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "textSP18"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "第三頁"
      TabPicture(2)   =   "frm020102_20.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(4)"
      Tab(2).Control(1)=   "grdList"
      Tab(2).ControlCount=   2
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   2856
         Left            =   -74880
         TabIndex        =   90
         Top             =   600
         Width           =   8652
         _ExtentX        =   15261
         _ExtentY        =   5038
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   -66720
         MaxLength       =   4
         TabIndex        =   4
         Top             =   552
         Width           =   540
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   264
         Left            =   -71790
         TabIndex        =   1
         Top             =   288
         Width           =   1425
      End
      Begin VB.TextBox textWord 
         Height          =   264
         Left            =   -68976
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1104
         Width           =   372
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1104
         Width           =   372
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   -67245
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   288
         Width           =   1080
      End
      Begin VB.TextBox textSP48 
         Height          =   264
         Left            =   -69480
         MaxLength       =   1
         TabIndex        =   16
         Top             =   3084
         Width           =   372
      End
      Begin VB.TextBox textSP47 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   15
         Top             =   3084
         Width           =   372
      End
      Begin VB.TextBox textSP63 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   14
         Top             =   2796
         Width           =   372
      End
      Begin VB.TextBox textSP40 
         Height          =   264
         Left            =   -69480
         MaxLength       =   8
         TabIndex        =   13
         Top             =   2508
         Width           =   1092
      End
      Begin VB.TextBox textSP39 
         Height          =   264
         Left            =   -73680
         MaxLength       =   8
         TabIndex        =   12
         Top             =   2508
         Width           =   1092
      End
      Begin VB.TextBox textSP38 
         Height          =   264
         Left            =   -73920
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2220
         Width           =   372
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   -73920
         MaxLength       =   8
         TabIndex        =   0
         Top             =   288
         Width           =   1092
      End
      Begin VB.TextBox textSP06 
         Height          =   270
         Left            =   -73560
         MaxLength       =   60
         TabIndex        =   9
         Top             =   1650
         Width           =   7215
      End
      Begin VB.ComboBox textCP44 
         Height          =   276
         Left            =   -73920
         TabIndex        =   3
         Top             =   552
         Width           =   1500
      End
      Begin VB.TextBox textCP22 
         Height          =   264
         Left            =   -69285
         MaxLength       =   1
         TabIndex        =   2
         Top             =   288
         Width           =   372
      End
      Begin MSForms.TextBox textSP44 
         Height          =   285
         Left            =   -69360
         TabIndex        =   18
         Top             =   3330
         Width           =   3165
         VariousPropertyBits=   671105051
         MaxLength       =   120
         Size            =   "5583;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP46 
         Height          =   285
         Left            =   -73920
         TabIndex        =   17
         Top             =   3330
         Width           =   2625
         VariousPropertyBits=   671105051
         MaxLength       =   20
         Size            =   "4630;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP18 
         Height          =   900
         Left            =   1080
         TabIndex        =   24
         Top             =   2520
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13520;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   465
         Left            =   1080
         TabIndex        =   23
         Top             =   2085
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13520;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP45 
         Height          =   465
         Left            =   1080
         TabIndex        =   22
         Top             =   1650
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   500
         ScrollBars      =   2
         Size            =   "13520;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP43 
         Height          =   465
         Left            =   1080
         TabIndex        =   21
         Top             =   1230
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   350
         ScrollBars      =   2
         Size            =   "13520;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP42 
         Height          =   465
         Left            =   1080
         TabIndex        =   20
         Top             =   795
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   120
         ScrollBars      =   2
         Size            =   "13520;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP41 
         Height          =   465
         Left            =   1080
         TabIndex        =   19
         Top             =   360
         Width           =   7665
         VariousPropertyBits=   -1466941413
         MaxLength       =   120
         ScrollBars      =   2
         Size            =   "13520;820"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP05 
         Height          =   285
         Left            =   -73560
         TabIndex        =   8
         Top             =   1380
         Width           =   7215
         VariousPropertyBits=   671105051
         MaxLength       =   140
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP07 
         Height          =   285
         Left            =   -73560
         TabIndex        =   10
         Top             =   1905
         Width           =   7215
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textSP51 
         Height          =   285
         Left            =   -73920
         TabIndex        =   5
         Top             =   825
         Width           =   6375
         VariousPropertyBits=   671105051
         MaxLength       =   30
         Size            =   "11245;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   285
         Left            =   -72420
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   555
         Width           =   4875
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "8599;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblCP113 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   18
         Left            =   -67530
         TabIndex        =   88
         Top             =   552
         Width           =   765
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   -72720
         TabIndex        =   87
         Top             =   300
         Width           =   900
      End
      Begin VB.Label Label27 
         Caption         =   "著作財產權人 :"
         Height          =   255
         Left            =   -70656
         TabIndex        =   86
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label36 
         Caption         =   "是否修改定稿內容 :"
         Height          =   252
         Left            =   -70656
         TabIndex        =   85
         Top             =   1128
         Width           =   1572
      End
      Begin VB.Label Label35 
         Caption         =   "(Y:修改)"
         Height          =   252
         Left            =   -68496
         TabIndex        =   84
         Top             =   1128
         Width           =   1332
      End
      Begin VB.Label Label34 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   83
         Top             =   1128
         Width           =   972
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   -73500
         TabIndex        =   82
         Top             =   1125
         Width           =   2745
      End
      Begin VB.Label Label21 
         Caption         =   "作品種類 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   72
         Top             =   3360
         Width           =   852
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "點數 :"
         Height          =   180
         Index           =   10
         Left            =   -67740
         TabIndex        =   81
         Top             =   300
         Width           =   450
      End
      Begin VB.Label Label1 
         Caption         =   "登記項目 :"
         Height          =   252
         Index           =   4
         Left            =   -74880
         TabIndex        =   79
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label32 
         Caption         =   "案件備註 :"
         Height          =   252
         Left            =   120
         TabIndex        =   78
         Top             =   2592
         Width           =   972
      End
      Begin VB.Label Label29 
         Caption         =   "登記原因 :"
         Height          =   252
         Left            =   120
         TabIndex        =   77
         Top             =   2208
         Width           =   972
      End
      Begin VB.Label Label28 
         Caption         =   "軟件說明 :"
         Height          =   252
         Left            =   120
         TabIndex        =   76
         Top             =   1752
         Width           =   852
      End
      Begin VB.Label Label26 
         Caption         =   "地址 :"
         Height          =   252
         Left            =   240
         TabIndex        =   75
         Top             =   1320
         Width           =   852
      End
      Begin VB.Label Label23 
         Caption         =   "代表人 :"
         Height          =   252
         Left            =   144
         TabIndex        =   74
         Top             =   888
         Width           =   852
      End
      Begin VB.Label Label22 
         Caption         =   "著作人 :"
         Height          =   252
         Left            =   144
         TabIndex        =   73
         Top             =   456
         Width           =   852
      End
      Begin VB.Label Label20 
         Caption         =   "(Y:發行)"
         Height          =   252
         Left            =   -69000
         TabIndex        =   71
         Top             =   3108
         Width           =   1092
      End
      Begin VB.Label Label19 
         Caption         =   "是否發行 :"
         Height          =   252
         Left            =   -70656
         TabIndex        =   70
         Top             =   3108
         Width           =   852
      End
      Begin VB.Label Label18 
         Caption         =   "(1:獨立狀態 2:共同狀態)"
         Height          =   252
         Left            =   -73440
         TabIndex        =   69
         Top             =   3132
         Width           =   2052
      End
      Begin VB.Label Label17 
         Caption         =   "擁有狀態 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   68
         Top             =   3132
         Width           =   852
      End
      Begin VB.Label Label16 
         Caption         =   "(1:單獨開發 2:合作開發 3:委託開發 4:下達任務開發)"
         Height          =   252
         Left            =   -73440
         TabIndex        =   67
         Top             =   2832
         Width           =   5892
      End
      Begin VB.Label Label15 
         Caption         =   "開發型式 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   66
         Top             =   2832
         Width           =   852
      End
      Begin VB.Label Label14 
         Caption         =   "首次發表日 :"
         Height          =   252
         Left            =   -70680
         TabIndex        =   65
         Top             =   2508
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "著作完成日 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   2532
         Width           =   1212
      End
      Begin VB.Label Label12 
         Caption         =   "(1:原創軟件 2:修改本 3:合成軟件 4:翻譯本)"
         Height          =   252
         Left            =   -73464
         TabIndex        =   63
         Top             =   2280
         Width           =   5892
      End
      Begin VB.Label Label11 
         Caption         =   "作品類型 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   62
         Top             =   2232
         Width           =   852
      End
      Begin VB.Label Label10 
         Caption         =   "案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   61
         Top             =   1428
         Width           =   1332
      End
      Begin VB.Label Label9 
         Caption         =   "案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   60
         Top             =   1680
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   59
         Top             =   1944
         Width           =   1452
      End
      Begin VB.Label Label7 
         Caption         =   "主管機關 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   58
         Top             =   864
         Width           =   972
      End
      Begin VB.Label Label5 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   57
         Top             =   564
         Width           =   972
      End
      Begin VB.Label Label30 
         Caption         =   "是否出名 :"
         Height          =   255
         Left            =   -70170
         TabIndex        =   55
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "(N:不出名)"
         Height          =   255
         Left            =   -68880
         TabIndex        =   54
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   53
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4950
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   1740
      Width           =   4125
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4950
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   660
      Width           =   4125
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4950
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   936
      Width           =   4125
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   34
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
      TabIndex        =   33
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
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1185
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   4950
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1185
      Width           =   4125
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   936
      Width           =   2532
   End
   Begin MSForms.ComboBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   89
      Top             =   1740
      Width           =   2550
      VariousPropertyBits=   679495707
      DisplayStyle    =   7
      Size            =   "4498;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   1200
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2535
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   4950
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   384
      Width           =   4125
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   4950
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1455
      Width           =   4125
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "7223;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "申請國家 :"
      Height          =   255
      Left            =   3930
      TabIndex        =   50
      Top             =   1770
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   49
      Top             =   1764
      Width           =   852
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   48
      Top             =   1488
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "FC代理人 :"
      Height          =   255
      Index           =   2
      Left            =   3930
      TabIndex        =   47
      Top             =   390
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   255
      Index           =   3
      Left            =   3930
      TabIndex        =   46
      Top             =   660
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   255
      Left            =   3930
      TabIndex        =   45
      Top             =   930
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   44
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   43
      Top             =   384
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   120
      TabIndex        =   42
      Top             =   1212
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   255
      Index           =   9
      Left            =   3930
      TabIndex        =   41
      Top             =   1215
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   3930
      TabIndex        =   40
      Top             =   1485
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   252
      Left            =   120
      TabIndex        =   39
      Top             =   936
      Width           =   732
   End
End
Attribute VB_Name = "frm020102_20"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo By Sindy 2022/2/21 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
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
' 申請人
Dim m_TM23 As String
'add by nickc 2007/02/01
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String

'Add By Sindy 2009/04/30
Dim m_CP84 As String       '發文規費

' 案件性質代號
Dim m_CP10 As String
' 登記項目
'Dim m_SP62 As String

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
'
Dim m_CurrSel As Integer
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
'add by nickc 2006/11/17
Dim m_textPrint As String
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
Dim m_CP07 As String 'Add By Sindy 2010/12/28 法定期限
Dim m_CP14 As String 'Add By Sindy 2012/9/10
Dim m_CP13 As String 'Add By Sindy 2012/10/4
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim strLD18 As String 'Add By Sindy 2019/12/25 信函總收文號


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

Private Sub cmdExit_Click()
'edit by nickc 2008/04/25 改整批印
'    'Add By Cheng 2004/04/08
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
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
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
                  'Add By Sindy 2014/7/30
                  If m_TM23 <> "" Then
                     frm020102_22.Label4.Caption = m_TM23 & " " & textTM23.List(0)
                  Else
                     frm020102_22.Label4.Caption = ""
                  End If
                  '2014/7/30 END
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
                  'Add By Sindy 2014/7/30
                  If m_TM78 <> "" Then
                     frm020102_22.Label4.Caption = m_TM78 & " " & textTM23.List(1)
                  Else
                     frm020102_22.Label4.Caption = ""
                  End If
                  '2014/7/30 END
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
                  'Add By Sindy 2014/7/30
                  If m_TM79 <> "" Then
                     frm020102_22.Label4.Caption = m_TM79 & " " & textTM23.List(2)
                  Else
                     frm020102_22.Label4.Caption = ""
                  End If
                  '2014/7/30 END
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
                  'Add By Sindy 2014/7/30
                  If m_TM80 <> "" Then
                     frm020102_22.Label4.Caption = m_TM80 & " " & textTM23.List(3)
                  Else
                     frm020102_22.Label4.Caption = ""
                  End If
                  '2014/7/30 END
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
                  'Add By Sindy 2014/7/30
                  If m_TM81 <> "" Then
                     frm020102_22.Label4.Caption = m_TM81 & " " & textTM23.List(4)
                  Else
                     frm020102_22.Label4.Caption = ""
                  End If
                  '2014/7/30 END
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
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
        'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
       'Add By Cheng 2002/11/08
       If Me.textPrint.Text <> "N" Then
          PrintLetter
       'Add By Sindy 2021/3/31
       End If
       If textPrint = "N" Then
         If strLD18 <> "" Then
            Call PUB_TCaseAskIsPost(strLD18)
         End If
       '2021/3/31 END
       End If
      
      '2012/7/23 add by sonia
      '台灣案發文規費與收文規費不符時,mail給智權人員
      If textCP84.Enabled = True And m_TM10 = "000" And Val(Me.textCP84.Text) <> Val(m_CP84) Then
        'Add by Lydia 2014/10/13 內商服務業務(TC)之台灣案發文-規費與收文規費不符時,請加同時發給特殊設定人員"財務處總帳人員"
        If m_QSP = True Then
          PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, "A"
        Else
          PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
        End If
      End If
      '2012/7/23 end
      
      'Add By Sindy 2018/5/3
      If frm020102_01.bolIsEMPFlow = True Then
         frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
         frm090202_4.QueryData
      End If
      '2018/5/3 End
      
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
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
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
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         textSP05 = rsTmp.Fields("TM05")
      End If
      SetTMSPFieldOldData "TM05", textSP05, 0
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("TM06")) = False Then
         textSP06 = rsTmp.Fields("TM06")
      End If
      SetTMSPFieldOldData "TM06", textSP06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("TM07")) = False Then
         textSP07 = rsTmp.Fields("TM07")
      End If
      SetTMSPFieldOldData "TM07", textSP07, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
      End If
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         'edit by nickc 2007/02/01
         'textTM23.AddItem rsTmp.Fields("TM23")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("TM23"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("TM23"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("TM78"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("TM78"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("TM79"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("TM79"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("TM80"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("TM80"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("TM81"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("TM81"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 顯示申請人
      If textTM23.ListCount > 0 Then
         textTM23.ListIndex = 0
      End If
      
      ' 案件備註
      If IsNull(rsTmp.Fields("TM58")) = False Then
         textSP18 = rsTmp.Fields("TM58")
      End If
      SetTMSPFieldOldData "TM58", textSP18, 0
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("tm24"))
      SetTMSPFieldOldData "TM24", m_TM24, 0
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "TM77", textPrint, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
End Sub

' 取得服務業務基本檔的欄位內容
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
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_TM23 = rsTmp.Fields("SP08")
      End If
      'If IsNull(rsTmp.Fields("SP08")) = False Then
      '   textTM23.AddItem rsTmp.Fields("SP08")
      'End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      m_TM79 = Empty
      m_TM80 = Empty
      m_TM81 = Empty
      '910709 Sieg
      ' 申請人1
      If IsNull(rsTmp.Fields("SP08")) = False Then
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("SP08"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("SP08"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 申請人2
      If IsNull(rsTmp.Fields("SP58")) = False Then
         'add by nickc 2007/02/01
         m_TM78 = rsTmp.Fields("SP58")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("SP58"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("SP58"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 申請人3
      If IsNull(rsTmp.Fields("SP59")) = False Then
         'add by nickc 2007/02/01
         m_TM79 = rsTmp.Fields("SP59")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("SP59"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("SP59"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 申請人4
      If IsNull(rsTmp.Fields("SP65")) = False Then
         'add by nickc 2007/02/01
         m_TM80 = rsTmp.Fields("SP65")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("SP65"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("SP65"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 申請人5
      If IsNull(rsTmp.Fields("SP66")) = False Then
         'add by nickc 2007/02/01
         m_TM81 = rsTmp.Fields("SP66")
         'edit by nickc 2007/02/06 不用 dll 了
         'objLawDll.LawGetName rsTmp.Fields("SP66"), strExc(0)
         ClsLawLawGetName rsTmp.Fields("SP66"), strExc(0)
         textTM23.AddItem strExc(0)
      End If
      ' 顯示申請人
      If textTM23.ListCount > 0 Then
         textTM23.ListIndex = 0
      End If
      ' 案件中文名稱
      If IsNull(rsTmp.Fields("SP05")) = False Then
         textSP05 = rsTmp.Fields("SP05")
      End If
      SetTMSPFieldOldData "SP05", textSP05, 0
      ' 案件英文名稱
      If IsNull(rsTmp.Fields("SP06")) = False Then
         textSP06 = rsTmp.Fields("SP06")
      End If
      SetTMSPFieldOldData "SP06", textSP06, 0
      ' 案件日文名稱
      If IsNull(rsTmp.Fields("SP07")) = False Then
         textSP07 = rsTmp.Fields("SP07")
      End If
      SetTMSPFieldOldData "SP07", textSP07, 0
      ' 申請國家
      If IsNull(rsTmp.Fields("SP09")) = False Then
         '910709 Sieg
         'textTM10 = rsTmp.Fields("SP09")
         textTM10 = GetNationName(rsTmp.Fields("SP09"), 0)
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         textTM12 = rsTmp.Fields("SP11")
      End If
      ' 發證日
      If IsNull(rsTmp.Fields("SP12")) = False Then
         textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
      End If
      ' 案件備註
      If IsNull(rsTmp.Fields("SP18")) = False Then
         textSP18 = rsTmp.Fields("SP18")
      End If
      SetTMSPFieldOldData "SP18", textSP18, 0
      ' 主管機關
      If IsNull(rsTmp.Fields("SP51")) = False Then
         textSP51 = rsTmp.Fields("SP51")
      End If
      SetTMSPFieldOldData "SP51", textSP51, 0
        'Add By Cheng 2003/07/14
        If Me.textSP51.Text = "" Then
            'edit by nickc 2006/10/23 桂英說的
            'Me.textSP51.Text = "財團法人台灣經濟發展研究院"
            'modify by sonia 2015/11/5 TC-010797
            'Me.textSP51.Text = "財團法人台灣經濟科技發展研究院"
            'Modified by Lydia 2025/01/15 改用CaseFee
            'If m_TM10 = "000" Then
            '   Me.textSP51.Text = "財團法人台灣經濟科技發展研究院"
            'Else
            '   Me.textSP51.Text = "中國版權局"
            'End If
            ''end 2015/11/5
            strExc(0) = "SELECT Distinct(CF10) FROM CaseFee WHERE CF01='" & m_TM01 & "' AND CF02='" & m_TM10 & "' AND length(CF03)=3 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Me.textSP51.Text = "" & RsTemp.Fields(0)
            End If
            'end 2025/01/15
        End If
        Me.textSP51.Locked = True 'Added by Lydia 2025/01/15 鎖定發文時之主管機關欄位---桂英
        
      ' 作品類型
      If IsNull(rsTmp.Fields("SP38")) = False Then
         textSP38 = rsTmp.Fields("SP38")
      End If
      SetTMSPFieldOldData "SP38", textSP38, 0
      ' 開發完成日
      If IsNull(rsTmp.Fields("SP39")) = False Then
         'Modify By Sindy 2012/10/1
         'textSP39 = TAIWANDATE(rsTmp.Fields("SP39"))
         If Right(rsTmp.Fields("SP39"), 2) = "00" Then
            textSP39 = Val(Left(rsTmp.Fields("SP39"), Len(rsTmp.Fields("SP39")) - 2)) - 191100
         Else
         '2012/10/1 End
            textSP39 = TAIWANDATE(rsTmp.Fields("SP39"))
         End If
      End If
      SetTMSPFieldOldData "SP39", textSP39, 1
      ' 首次發表日
      If IsNull(rsTmp.Fields("SP40")) = False Then
         textSP40 = TAIWANDATE(rsTmp.Fields("SP40"))
      End If
      SetTMSPFieldOldData "SP40", textSP40, 1
      ' 開發型式
      If IsNull(rsTmp.Fields("SP63")) = False Then
         textSP63 = rsTmp.Fields("SP63")
      End If
      SetTMSPFieldOldData "SP63", textSP63, 0
      ' 著作人
      If IsNull(rsTmp.Fields("SP41")) = False Then
         textSP41 = rsTmp.Fields("SP41")
      End If
      SetTMSPFieldOldData "SP41", textSP41, 0
      ' 代表人
      If IsNull(rsTmp.Fields("SP42")) = False Then
         textSP42 = rsTmp.Fields("SP42")
      End If
      SetTMSPFieldOldData "SP42", textSP42, 0
      ' 地址
      If IsNull(rsTmp.Fields("SP43")) = False Then
         textSP43 = rsTmp.Fields("SP43")
        'Add By Cheng 2003/07/14
        Else
            textSP43 = PUB_GetCustEachAdd("" & rsTmp.Fields("SP08").Value, "1")
      End If
      SetTMSPFieldOldData "SP43", textSP43, 0
        'Modify By Cheng 2004/02/03
        'Cancel Marked
'       91.09.02 marked by louis
       '著作權人
      If IsNull(rsTmp.Fields("SP44")) = False Then
         textSP44 = rsTmp.Fields("SP44")
      '2009/5/15 MODIFY BY SONIA 自著作人SP41移過來(宋若蘭)
      Else
         textSP44.Text = GetCustomerName("" & rsTmp.Fields("SP08"), "0")
      '2009/5/15 END
      End If
      SetTMSPFieldOldData "SP44", textSP44, 0
        'End
      ' 軟件說明
      If IsNull(rsTmp.Fields("SP45")) = False Then
         textSP45 = rsTmp.Fields("SP45")
      End If
      SetTMSPFieldOldData "SP45", textSP45, 0
      ' 作品種類
      If IsNull(rsTmp.Fields("SP46")) = False Then
         textSP46 = rsTmp.Fields("SP46")
      End If
      SetTMSPFieldOldData "SP46", textSP46, 0
      ' 擁有狀態
      If IsNull(rsTmp.Fields("SP47")) = False Then
         textSP47 = rsTmp.Fields("SP47")
      End If
      SetTMSPFieldOldData "SP47", textSP47, 0
      ' 是否發行
      If IsNull(rsTmp.Fields("SP48")) = False Then
         textSP48 = rsTmp.Fields("SP48")
      End If
      SetTMSPFieldOldData "SP48", textSP48, 0
      ' 91.09.02 modify by louis
      ' 登記項目
      'm_SP62 = Empty
      'If IsNull(rsTmp.Fields("SP62")) = False Then
      '   m_SP62 = rsTmp.Fields("SP62")
      'End If
      'SetTMSPFieldOldData "SP62", m_SP62, 0
      'Add By Cheng 2002/07/22
      SetTMSPFieldOldData "SP10", "" & rsTmp.Fields("SP10"), 0
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      m_textPrint = textPrint
      SetTMSPFieldOldData "SP72", textPrint, 0
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
Dim strCP45 As String
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
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         'textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      m_CP13 = "" 'Add By Sindy 2012/10/4
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
         m_CP13 = rsTmp.Fields("CP13") 'Add By Sindy 2012/10/4
      End If
      ' 承辦人員
      m_CP14 = "" 'Add By Sindy 2012/9/10
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
         m_CP14 = rsTmp.Fields("CP14") 'Add By Sindy 2012/9/10
      End If
      
      'Add By Sindy 2010/12/28 法定期限
      m_CP07 = ""
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      '2010/12/28 End
      
      'Add By Sindy 2011/7/12
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      
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
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then: textCP44 = rsTmp.Fields("CP44")
      SetCPFieldOldData "CP44", textCP44, 0
      ' 彼所案號
      strCP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then: strCP45 = rsTmp.Fields("CP45")
      SetCPFieldOldData "CP45", strCP45, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      'Add By Sindy 2009/04/30 發文規費
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
         m_CP84 = CheckStr(rsTmp.Fields("CP17"))
         '2012/8/3 add by sonia 若有銷帳則要扣除銷帳規費
         If Val("" & rsTmp.Fields("CP77")) <> 0 Then
            If GetCP77Detail(m_CP09, m_Fee, m_Official) = True Then
               m_CP84 = m_CP84 - m_Official
            End If
         End If
         '2012/8/3 end
         textCP84.Text = m_CP84
      End If
      
      'Added by Morgan 2012/9/6 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/9/6
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
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
       'Modify By Cheng 2002/09/19
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
   Dim nIndex As Integer
   Dim nRow As Integer
   Dim strTemp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
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
   
   'Add By Cheng 2002/06/18
   ' 取得國家代碼
   m_TM10 = GetNationNo(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 本所案號
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04

   ' 收文號
   textCP09 = m_CP09
      
   ' 取得案件進度檔的欄位
   QueryCaseProgress
   
   'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
   m_QSP = False
   
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
        'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
         m_QSP = True
   End Select
   
   'Add By Sindy 2021/1/15 T發文所有程式,台灣案鎖住畫面上之CP44,不可輸入
   If m_TM10 = "000" Then
      textCP44.Enabled = False
   End If
   '2021/1/15 END
   
   ' 登記項目
   InitialGrdList
   strSql = "SELECT * FROM PatentTrademarkMap " & _
            "WHERE PTM01 = '3' " & _
            "ORDER BY PTM02 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         If IsNull(rsTmp.Fields("PTM02")) = False Then
            grdList.TextMatrix(grdList.row, 1) = rsTmp.Fields("PTM02")
         End If
         If IsNull(rsTmp.Fields("PTM03")) = False Then
            grdList.TextMatrix(grdList.row, 2) = rsTmp.Fields("PTM03")
         End If
         rsTmp.MoveNext
      Loop
      grdList.FixedRows = 1 'Added by Lydia 2023/10/13
   End If
   rsTmp.Close
   'Set rsTmp = Nothing
   ' 91.09.02 modify by louis
   ' 顯示檔案中所包含的項目
   'For nIndex = 1 To GetSubStringCount(m_SP62)
   '   strTemp = GetSubString(m_SP62, nIndex)
   '   For nRow = 1 To grdList.Rows - 1
   '      If strTemp = grdList.TextMatrix(nRow, 2) Then
   '         grdList.TextMatrix(nRow, 0) = "V"
   '      End If
   '   Next nRow
   'Next nIndex
   ' 從登記項目檔中取出所勾選的部份
   strSql = "SELECT * FROM COPYRIGHTITEM " & _
            "WHERE CRI01 = '" & m_CP09 & "' " & _
            "ORDER BY CRI02 ASC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Dim strItem As String
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         strItem = rsTmp.Fields("CRI02")
         For nRow = 1 To grdList.Rows - 1
            If grdList.TextMatrix(nRow, 1) = strItem Then
               grdList.TextMatrix(nRow, 0) = "V"
               Exit For
            End If
         Next nRow
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'Add By Sindy 2021/3/31 案件性質為706(其它),定稿列印請自動上 "N"
   If m_CP10 = "706" Then
      textPrint = "N"
   End If
   '2021/3/31 END
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 3
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "種類"
   grdList.ColWidth(1) = 800
   grdList.col = 2
   grdList.Text = "登記項目"
   grdList.ColWidth(2) = 3000
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm020102_20 = Nothing
End Sub

Private Sub grdList_Click()
   If grdList.row > 0 Then
      grdList.col = 0
      If grdList.Text = "V" Then
         grdList.Text = Empty
      Else
         grdList.Text = "V"
      End If
   End If
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 是否出名
Private Sub textCP22_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否出名
Private Sub textCP22_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP22) = False Then
      Select Case textCP22
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP22_GotFocus
      End Select
   End If
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
   End If
EXITSUB:
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/03
    KeyAscii = UpperCase(KeyAscii)
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

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

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
   End If
End Sub

' 案件中文名稱
Private Sub textSP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP05, 140) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP05_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP05.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 案件英文名稱
Private Sub textSP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP06, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP06_GotFocus
   End If
End Sub

' 案件日文名稱
Private Sub textSP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP07, 40) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP07_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textSP07.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub
        
' 案件備註
Private Sub textSP18_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textSP18, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP18_GotFocus
   End If
End Sub

'Add By Sindy 2009/04/30
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
'2009/04/30 End

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

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
'edit by nickc 2006/06/29
'   If KeyAscii <> 78 And KeyAscii <> 8 Then
'      KeyAscii = 0
'   End If
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim nIndex As Integer
   'Dim strSP62 As String
   
   ' 是否出名
   SetCPFieldNewData "CP22", textCP22
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 代理人代號
   SetCPFieldNewData "CP44", textCP44
      'add by nickc 2008/02/22
      m_CP44New = textCP44
   ' 彼所案號
   SetCPFieldNewData "CP45", textTM45
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   Select Case m_TM01
      Case "T", "TF", "FCT":
         ' 案件名稱(中)
         SetTMSPFieldNewData "TM05", textSP05
         ' 案件名稱(英)
         SetTMSPFieldNewData "TM06", textSP06
         ' 案件名稱(日)
         SetTMSPFieldNewData "TM07", textSP07
         ' 案件備註
         SetTMSPFieldNewData "TM58", textSP18
         'add by nickc 2006/01/26
         If m_CU112 <> "" Then
            'Modify By Sindy 2011/2/22
            'SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112)
            SetTMSPFieldNewData "TM24", Pub_RplCu112(m_TM24, m_CU112, m_TM23)
         Else
            SetTMSPFieldNewData "TM24", m_TM24
         End If
         'add by nickc 2006/11/17
         If textPrint <> "N" Then
            SetTMSPFieldNewData "TM77", textPrint
         Else
            SetTMSPFieldNewData "TM77", m_textPrint
         End If
      Case Else:
         ' 案件名稱(中)
         SetTMSPFieldNewData "SP05", textSP05
         ' 案件名稱(英)
         SetTMSPFieldNewData "SP06", textSP06
         ' 案件名稱(日)
         SetTMSPFieldNewData "SP07", textSP07
         'Add By Cheng 2002/07/22
         If m_CP10 = "806" Then
            ' 申請日
            SetTMSPFieldNewData "SP10", DBDATE(Me.textCP27.Text)
         End If
         ' 案件備註
         SetTMSPFieldNewData "SP18", textSP18
         ' 作品類型
         SetTMSPFieldNewData "SP38", textSP38
         ' 開發完成日
         'Modify By Sindy 2012/10/1
         'SetTMSPFieldNewData "SP39", DBDATE(textSP39)
         If Len(Trim(textSP39)) >= 6 Then
            SetTMSPFieldNewData "SP39", DBDATE(textSP39)
         Else
            SetTMSPFieldNewData "SP39", (Val(textSP39) + 191100) & "00"
         End If
         '2012/10/1 End
         ' 首次發表日
         SetTMSPFieldNewData "SP40", DBDATE(textSP40)
         ' 著作人
         SetTMSPFieldNewData "SP41", textSP41
         ' 代表人
         SetTMSPFieldNewData "SP42", textSP42
         ' 地址
         SetTMSPFieldNewData "SP43", textSP43
         'Modify By Cheng 2004/02/03
         'Cancel Marked
         ' 91.09.02 marked by louis
         ' 著作權人
         SetTMSPFieldNewData "SP44", textSP44
        'End
         ' 軟件說明
         SetTMSPFieldNewData "SP45", textSP45
         ' 作品種類
         SetTMSPFieldNewData "SP46", textSP46
         ' 擁有狀態
         SetTMSPFieldNewData "SP47", textSP47
         ' 是否發行
         SetTMSPFieldNewData "SP48", textSP48
         ' 主管機關
         SetTMSPFieldNewData "SP51", textSP51
         ' 開發型式
         SetTMSPFieldNewData "SP63", textSP63
         
         ' 91.09.02 modify by louis
         ' 登記項目
         'strSP62 = Empty
         'For nIndex = 1 To grdList.Rows - 1
         '   If grdList.TextMatrix(nIndex, 0) = "V" Then
         '      If strSP62 <> Empty Then
         '         strSP62 = strSP62 & ","
         '      End If
         '      strSP62 = strSP62 & grdList.TextMatrix(nIndex, 2)
         '   End If
         'Next nIndex
         ' 開發型式
         'SetTMSPFieldNewData "SP62", strSP62
         'add by nickc 2006/11/17
         If textPrint <> "N" Then
            SetTMSPFieldNewData "SP72", textPrint
         Else
            SetTMSPFieldNewData "SP72", m_textPrint
         End If
         
   End Select
   
End Sub

' 更新商標基本檔的相關欄位
'Modify By Cheng 2002/11/07
'Private Sub OnUpdateTradeMark()
Private Function OnUpdateTradeMark() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateTradeMark = True

   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateTradeMark = False
End Function

' 更新服務業務基本檔的相關欄位
'Modify By Cheng 2002/11/07
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnUpdateServicePractice = True

   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               ' 91.03.25 modify by louis
               'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
               strTmp = m_TMSPList(nIndex).fiName & " = '" & ChgSQL(m_TMSPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_TMSPList(nIndex).fiNewData = Empty Then
               strTmp = m_TMSPList(nIndex).fiName & " = NULL "
            Else
               strTmp = m_TMSPList(nIndex).fiName & " = " & m_TMSPList(nIndex).fiNewData
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
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
   ' 執行SQL指令
   'Modified by Lydia 2025/01/15 增加log
   'If bDifference = True Then: cnnConnection.Execute strSql
   If bDifference = True Then
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   End If
   'end 2025/01/15
   
'Add By Cheng 2002/11/07
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

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
               ' 91.03.25 modify by louis
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

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim nIndex As Integer
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   
'Add By Cheng 2002/11/07
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
    'Modify By Cheng 2002/11/07
'   OnUpdateCaseProgress
   If OnUpdateCaseProgress = False Then GoTo ErrorHandler
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔
   Select Case m_TM01
      Case "T", "TF", "FCT":
        'Modify By Cheng 2002/11/07
'         OnUpdateTradeMark
         If OnUpdateTradeMark = False Then GoTo ErrorHandler
      Case Else:
        'Modify By Cheng 2002/11/07
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select
      
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
            strNP08 = DBDATE(textCP27)
           'Modify By Cheng 2003/09/01
   '         strNP08 = DBDATE(Format(DateSerial(Val(DBYEAR(strNP08)), Val(DBMONTH(strNP08)), Val(DBDAY(strNP08)) + Val(rsTmp.Fields("CF23")))))
            strNP08 = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF23")), ChangeWStringToWDateString(DBDATE(strNP08))))
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
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'         '92.6.8 SONIA 加 言詞辯論, 準備程序
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
      'Add By Sindy 2012/9/10
      If IsNull(rsTmp.Fields("CF05")) = False Then
         strNP07 = "305"
         strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
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
      '2012/9/10 End
   End If
   
   'Added by Lydia 2016/09/07 台至大著作權(TC)案,送交代理提申列印申請書(管制代理人10天回覆-期限掛承辦人-列申請書)
                            '大陸案件發文存檔時，新增下一程序期限、NP07=994陸代申請書、NP08=NP09=發文日+10日曆天、NP10=CP14。
   If m_TM10 = "020" Then
        strNP07 = "994"
        strNP08 = CompDate(2, 10, DBDATE(textCP27))
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
   'end 2016/09/07
   
   rsTmp.Close
   
   ' 91.09.02 modify by louis
   ' 登記項目
   ' 先刪除舊的資料
   strSql = "DELETE FROM COPYRIGHTITEM " & _
            "WHERE CRI01 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
   ' 增加新的項目
   strSql = Empty
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strSql = "INSERT INTO COPYRIGHTITEM (CRI01, CRI02) " & _
                        "VALUES ('" & m_CP09 & "','" & grdList.TextMatrix(nIndex, 1) & "') "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
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
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
            'edit by nickc 2007/08/10
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "' Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
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
    'Modify By Cheng 2002/11/08
'   'Modify By Cheng 2002/06/14
'   If Me.textPrint.Text <> "N" Then
''   ' 直接列印定稿
'      PrintLetter
'   End If
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add By Sindy 2009/04/30 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2019/12/25 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = m_CP09
      PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, m_CP10, m_TM44
   End If
   '2019/12/25 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
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
   ' 案件名稱不可同時空白
   If IsEmptyText("textSP05") = True And IsEmptyText("textSP06") = True And IsEmptyText("textSP07") = True Then
      strTit = "檢核資料"
      strMsg = "請輸入案件名稱"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP05.SetFocus
      GoTo EXITSUB
   End If
   
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

Private Sub textCP22_GotFocus()
   InverseTextBox textCP22
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textSP05_GotFocus()
   InverseTextBox textSP05
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP05.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP06_GotFocus()
   InverseTextBox textSP06
End Sub

Private Sub textSP07_GotFocus()
   InverseTextBox textSP07
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textSP07.IMEMode = 1
   OpenIme
End Sub

Private Sub textSP18_GotFocus()
   InverseTextBox textSP18
End Sub

Private Sub textSP38_GotFocus()
   InverseTextBox textSP38
End Sub

Private Sub textSP38_KeyPress(KeyAscii As Integer)
If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textSP39_GotFocus()
   InverseTextBox textSP39
End Sub

Private Sub textSP39_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse

'Add By Cheng 2002/06/18
If Len(Me.textSP39.Text) > 0 Then
   If Len(Me.textSP39.Text) = 8 Then '西元日期
      If CheckIsDate(Me.textSP39.Text) = False Then
         Cancel = True
         Me.textSP39.SetFocus
         TextInverse Me.textSP39
         Exit Sub
      End If
   'Modify By Sindy 2012/10/1
   ElseIf Len(Me.textSP39.Text) = 7 Or Len(Me.textSP39.Text) = 6 Then '民國日期
      If CheckIsTaiwanDate(Me.textSP39.Text) = False Then
         Cancel = True
         Me.textSP39.SetFocus
         TextInverse Me.textSP39
         Exit Sub
      End If
   Else '民國年月
      If CheckIsTaiwanDate(Me.textSP39.Text & "01", False) = False Then
         strMsg = "民國年月不正確, 請重新輸入"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Cancel = True
         Me.textSP39.SetFocus
         TextInverse Me.textSP39
         Exit Sub
      End If
   End If
   '2012/10/1 End
End If
End Sub

Private Sub textSP40_GotFocus()
   InverseTextBox textSP40
End Sub

Private Sub textSP40_Validate(Cancel As Boolean)
'Add By Cheng 2002/06/18
If Len(Me.textSP40.Text) > 0 Then
   If Len(Me.textSP40.Text) = 8 Then
      If CheckIsDate(Me.textSP40.Text) = False Then
         Cancel = True
         Me.textSP40.SetFocus
         TextInverse Me.textSP40
         Exit Sub
      End If
   Else
      If CheckIsTaiwanDate(Me.textSP40.Text) = False Then
         Cancel = True
         Me.textSP40.SetFocus
         TextInverse Me.textSP40
         Exit Sub
      End If
   End If
End If
End Sub

Private Sub textSP41_GotFocus()
   InverseTextBox textSP41
End Sub

Private Sub textSP41_LostFocus()
   If Not CheckLengthIsOK(textSP41, 120) Then
      'MsgBox "著作人內容太長 !", vbCritical
      textSP41.SetFocus
   End If
End Sub

Private Sub textSP42_LostFocus()
   If Not CheckLengthIsOK(textSP42, 120) Then
      'MsgBox "代表人內容太長 !", vbCritical
      textSP42.SetFocus
   End If
End Sub

Private Sub textSP43_LostFocus()
   If Not CheckLengthIsOK(textSP43, 350) Then
      'MsgBox "地址內容太長 !", vbCritical
      textSP43.SetFocus
   End If
End Sub

'Modify By Cheng 2004/02/03
'Cancel Marked
' 91.09.02 marked by louis
Private Sub textSP44_LostFocus()
    If Not CheckLengthIsOK(textSP44, 120) Then
        'MsgBox "著作權人內容太長 !", vbCritical
        textSP44.SetFocus
    End If
End Sub

Private Sub textSP45_LostFocus()
   If Not CheckLengthIsOK(textSP45, 500) Then
      'MsgBox "軟件說明內容太長 !", vbCritical
      textSP45.SetFocus
   End If
End Sub

Private Sub textSP46_LostFocus()
   If Not CheckLengthIsOK(textSP46, 20) Then
      'MsgBox "作品種類內容太長 !", vbCritical
      textSP46.SetFocus
   End If

   '若申請國家為大陸且作品種類為美術著作
   '2009/5/14 modify by sonia
   'If m_TM10 = 大陸國家代號 And Me.textSP46.Text = "美術著作" Then
   If m_TM10 = 大陸國家代號 And Me.textSP46.Text = "計算機軟件" Then
      If Len(Me.textSP38.Text) <= 0 Then
         MsgBox "請輸入作品類型!!!", vbExclamation + vbOKOnly
         Me.textSP38.SetFocus
         TextInverse Me.textSP38
         Exit Sub
      End If
      If Len(Me.textSP63.Text) <= 0 Then
         MsgBox "請輸入開發型式!!!", vbExclamation + vbOKOnly
         Me.textSP63.SetFocus
         TextInverse Me.textSP38
         Exit Sub
      End If
      If Len(Me.textSP47.Text) <= 0 Then
         MsgBox "請輸入開發型式!!!", vbExclamation + vbOKOnly
         Me.textSP47.SetFocus
         TextInverse Me.textSP38
         Exit Sub
      End If
   End If
   
End Sub

Private Sub textSP42_GotFocus()
   InverseTextBox textSP42
End Sub

Private Sub textSP43_GotFocus()
   InverseTextBox textSP43
End Sub

'Modify By Cheng 2004/02/03
'Cancel Marked
' 91.09.02 marked by louis
Private Sub textSP44_GotFocus()
    InverseTextBox textSP44
End Sub

Private Sub textSP45_GotFocus()
   InverseTextBox textSP45
End Sub

Private Sub textSP46_GotFocus()
   InverseTextBox textSP46
End Sub

Private Sub textSP47_GotFocus()
   InverseTextBox textSP47
End Sub

Private Sub textSP47_KeyPress(KeyAscii As Integer)
If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textSP48_GotFocus()
   InverseTextBox textSP48
End Sub

Private Sub textSP51_GotFocus()
   InverseTextBox textSP51
End Sub

Private Sub textSP51_LostFocus()
   If Not CheckLengthIsOK(textSP51, 30) Then
      'MsgBox "主管機關內容太長 !", vbCritical
      textSP51.SetFocus
   End If
End Sub

Private Sub textSP63_GotFocus()
   InverseTextBox textSP63
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 案件性質為著作權申請
   'edit by nickc 2006/06/29 之前有漏掉
   'If m_CP10 = "806" Then
   If m_CP10 = "806" Or m_CP10 = "503" Then
      ' 申請國家為大陸
      If m_TM10 = "020" Then
         'add by nickc 2006/06/29
         If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "30", strUserNum
         End If
      ' 申請國家為台灣
      ElseIf m_TM10 < "010" Then
         'add by nickc 2006/06/29
         If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "32", strUserNum
         'Add By Sindy 2010/01/12 大->台
         ElseIf textPrint = "2" And m_CP10 = "806" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "01", m_CP09, "33", strUserNum
         '2010/01/12 End
         End If
      End If
   End If
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
   bolEdit = IIf(Me.textWord.Text = "Y", True, False)
   '2012/1/12 End
   
'   ' 案件性質為著作權申請
   'edit by nickc 2006/06/29
   'If m_CP10 = "806" Then
   If m_CP10 = "806" Or m_CP10 = "501" Then
      ' 申請國家為大陸
      If m_TM10 = "020" Then
         'add by nickc 2006/06/29
         If textPrint = "1" Then
         ' 列印定稿
         'Modify By Cheng 2002/06/14
'         NowPrint m_CP09, "01", "30", False, strUserNum, 0
'            NowPrint m_CP09, "01", "30", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
            ET03 = "30" 'Modify By Sindy 2012/1/12
         End If
      ' 申請國家為台灣
      ElseIf m_TM10 < "010" Then
         'add by nickc 2006/06/29
         If textPrint = "1" Then
         ' 列印定稿
         'Modify By Cheng 2002/06/14
'         NowPrint m_CP09, "01", "32", False, strUserNum, 0
'            NowPrint m_CP09, "01", "32", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
            ET03 = "32" 'Modify By Sindy 2012/1/12
          'Add By Sindy 2010/01/12 大->台
          ElseIf textPrint = "2" And m_CP10 = "806" Then
'            NowPrint m_CP09, "01", "33", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
            ET03 = "33" 'Modify By Sindy 2012/1/12
          '2010/01/12 End
          End If
      End If
   End If
'    'Add By Cheng 2003/03/18
'    '移轉(讓與)
'    If m_CP10 = "501" Then
'        ' 申請國家為大陸
'        If m_TM10 = "020" Then
'           ' 列印定稿
'           NowPrint m_CP09, "01", "30", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
'        ' 申請國家為台灣
'        ElseIf m_TM10 < "010" Then
'           ' 列印定稿
'           NowPrint m_CP09, "01", "32", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
'        End If
'    End If
   
   'Add By Sindy 2012/1/12
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , , , , , , , strLD18
         Else
         '2020/1/7 END
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , iCopy, , True, True
            MsgBox "電子檔已存於 [ " & PUB_GetEFilePath(m_TM01) & " ]！"
         End If
      Else
         'Add By Sindy 2019/12/25 + strLD18.信函總收文號
         NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
      End If
   'Add By Sindy 2021/1/5 沒有系統產出的定稿
   Else
      If strLD18 <> "" Then
         'Modify By Sindy 2025/8/15
         'Call PUB_TCaseAskIsPost(strLD18)
         textPrint = "N"
         '2025/8/15 END
      End If
   '2021/1/5 EMD
   End If
   '2012/1/12 End
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   
   'Add By Sindy 2009/04/30
   If Me.textCP84.Enabled = True Then
      Cancel = False
      textCP84_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textCP84.Enabled = True And m_TM10 = "000" Then
       If Val(textCP84.Text) <> Val(m_CP84) Then
           If MsgBox("收文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
               textCP84_GotFocus
               Exit Function
           End If
       End If
   End If
   '2009/04/30 End
   
   '910723 Sieg
   If Not CheckLengthIsOK(textSP51, 30) Then
      'MsgBox "主管機關內容太長 !", vbCritical
      textSP51.SetFocus
      Exit Function
   End If
   
   If Not CheckLengthIsOK(textSP46, 20) Then
      'MsgBox "作品種類內容太長 !", vbCritical
      textSP46.SetFocus
      Exit Function
   End If
   
   If Not CheckLengthIsOK(textSP41, 120) Then
      'MsgBox "著作人內容太長 !", vbCritical
      textSP41.SetFocus
      Exit Function
   End If
   
   If Not CheckLengthIsOK(textSP42, 120) Then
      'MsgBox "代表人內容太長 !", vbCritical
      textSP42.SetFocus
      Exit Function
   End If
   
   If Not CheckLengthIsOK(textSP43, 350) Then
      'MsgBox "地址內容太長 !", vbCritical
      textSP43.SetFocus
      Exit Function
   End If
   
   'Modify By Cheng 2004/02/03
   'Cancel Marked
   ' 91.09.02 marked by louis
   If Not CheckLengthIsOK(textSP44, 120) Then
   '   MsgBox "著作權人內容太長 !", vbCritical
      textSP44.SetFocus
      Exit Function
   End If
   
   If Not CheckLengthIsOK(textSP45, 500) Then
      'MsgBox "軟件說明內容太長 !", vbCritical
      textSP45.SetFocus
      Exit Function
   End If
   
   If Me.textCP22.Enabled = True Then
      Cancel = False
      textCP22_Validate Cancel
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
   
   If Me.textCP44.Enabled = True Then
      Cancel = False
      textCP44_Validate Cancel
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
   
   If Me.textSP05.Enabled = True Then
      Cancel = False
      textSP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP06.Enabled = True Then
      Cancel = False
      textSP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP07.Enabled = True Then
      Cancel = False
      textSP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP18.Enabled = True Then
      Cancel = False
      textSP18_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP39.Enabled = True Then
      Cancel = False
      textSP39_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textSP40.Enabled = True Then
      Cancel = False
      textSP40_Validate Cancel
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
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   TxtValidate = True
End Function

Private Sub textSP63_KeyPress(KeyAscii As Integer)
If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 52 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textWord_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 89 And KeyAscii <> 8 Then
   KeyAscii = 0
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
