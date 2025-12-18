VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_16 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(異議答辯, 評定答辯, 廢止答辯, 補充答辯, 參加被評定, 撤銷禁止處分, 修正, 註冊費, 其它)"
   ClientHeight    =   6084
   ClientLeft      =   4776
   ClientTop       =   2196
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6084
   ScaleWidth      =   9132
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   325
      Index           =   1
      Left            =   3870
      TabIndex        =   77
      Top             =   0
      Width           =   1200
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   350
      Left            =   5100
      TabIndex        =   67
      Top             =   0
      Width           =   1200
   End
   Begin VB.TextBox textCP08 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5190
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Top             =   1980
      Width           =   3885
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   630
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   360
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5190
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   630
      Width           =   3885
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   900
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8316
      TabIndex        =   23
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   6300
      TabIndex        =   21
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7092
      TabIndex        =   22
      Top             =   0
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3492
      Left            =   120
      TabIndex        =   24
      Top             =   2556
      Width           =   8892
      _ExtentX        =   15685
      _ExtentY        =   6160
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "發文資料"
      TabPicture(0)   =   "frm020102_16.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label16"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label15"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label11"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(12)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(5)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(7)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(8)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(13)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblRCP14"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label39"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label43"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label3"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lblPayToday"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblCP113(18)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textRCP64"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP44_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "grdList"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP30"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textRCP14"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textTM29"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textPrint"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP18"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP44"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCF09"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP23"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP49"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCP43"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textRCP23"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP118"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtPayToday"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCP84"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP27"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtCP113"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame1"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "相關人"
      TabPicture(1)   =   "frm020102_16.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label28"
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label14"
      Tab(1).Control(4)=   "lblNameAgent"
      Tab(1).Control(5)=   "Label31"
      Tab(1).Control(6)=   "Label30"
      Tab(1).Control(7)=   "textCP64"
      Tab(1).Control(8)=   "textCP42"
      Tab(1).Control(9)=   "textCP40"
      Tab(1).Control(10)=   "lstNameAgent"
      Tab(1).Control(11)=   "textCP41"
      Tab(1).Control(12)=   "textCP22"
      Tab(1).ControlCount=   13
      Begin VB.Frame Frame1 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   6060
         TabIndex        =   91
         Top             =   300
         Visible         =   0   'False
         Width           =   2780
         Begin VB.TextBox textTM136 
            Height          =   264
            Left            =   1140
            MaxLength       =   1
            TabIndex        =   3
            Top             =   0
            Width           =   345
         End
         Begin VB.Label Label71 
            AutoSize        =   -1  'True
            Caption         =   "註冊證型式:            (1:電子2:紙本)"
            Height          =   180
            Left            =   60
            TabIndex        =   92
            Top             =   30
            Width           =   2660
            WordWrap        =   -1  'True
         End
      End
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   5400
         MaxLength       =   4
         TabIndex        =   2
         Top             =   285
         Width           =   540
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1080
         MaxLength       =   8
         TabIndex        =   0
         Top             =   285
         Width           =   1092
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   3330
         TabIndex        =   1
         Top             =   270
         Width           =   1092
      End
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   8115
         MaxLength       =   1
         TabIndex        =   12
         Top             =   1680
         Width           =   255
      End
      Begin VB.TextBox textCP118 
         Height          =   264
         Left            =   7920
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1944
         Width           =   375
      End
      Begin VB.TextBox textCP22 
         Height          =   264
         Left            =   -71490
         MaxLength       =   1
         TabIndex        =   74
         Top             =   1410
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textRCP23 
         Enabled         =   0   'False
         Height          =   264
         Left            =   6390
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1125
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textCP43 
         Height          =   264
         Left            =   5700
         MaxLength       =   12
         TabIndex        =   6
         Top             =   840
         Width           =   1485
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73800
         MaxLength       =   600
         TabIndex        =   18
         Top             =   660
         Width           =   7392
      End
      Begin VB.TextBox textCP49 
         Height          =   264
         Left            =   1080
         MaxLength       =   300
         TabIndex        =   16
         Top             =   2196
         Width           =   7572
      End
      Begin VB.TextBox textCP23 
         Height          =   264
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1692
         Width           =   372
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   4860
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1944
         Width           =   825
      End
      Begin VB.ComboBox textCP44 
         Height          =   276
         Left            =   1080
         TabIndex        =   4
         Top             =   564
         Width           =   1476
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   240
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   870
         Width           =   1395
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   13
         Top             =   1944
         Width           =   372
      End
      Begin VB.TextBox textTM29 
         Height          =   264
         Left            =   4590
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1680
         Width           =   375
      End
      Begin VB.TextBox textRCP14 
         Enabled         =   0   'False
         Height          =   264
         Left            =   1944
         MaxLength       =   6
         TabIndex        =   7
         Top             =   1128
         Visible         =   0   'False
         Width           =   1092
      End
      Begin VB.TextBox textCP30 
         Height          =   264
         Left            =   3720
         MaxLength       =   20
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   900
         Left            =   1080
         TabIndex        =   93
         Top             =   2496
         Width           =   7740
         _ExtentX        =   13653
         _ExtentY        =   1588
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
      Begin MSForms.ListBox lstNameAgent 
         Height          =   495
         Left            =   -73830
         TabIndex        =   90
         Top             =   1290
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
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73800
         TabIndex        =   17
         Top             =   360
         Width           =   7392
         VariousPropertyBits=   -1467989989
         MaxLength       =   600
         ScrollBars      =   2
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73800
         TabIndex        =   19
         Top             =   960
         Width           =   7392
         VariousPropertyBits=   -1467989989
         MaxLength       =   600
         ScrollBars      =   2
         Size            =   "13039;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   732
         Left            =   -73830
         TabIndex        =   20
         Top             =   1800
         Width           =   7392
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13039;1291"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   264
         Left            =   2568
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   564
         Width           =   6084
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         MaxLength       =   20
         Size            =   "6482;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textRCP64 
         Height          =   300
         Left            =   2115
         TabIndex        =   9
         Top             =   1410
         Visible         =   0   'False
         Width           =   6495
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11465;529"
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
         Left            =   4590
         TabIndex        =   89
         Top             =   330
         Width           =   765
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   6180
         TabIndex        =   88
         Top             =   1710
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "對方案件號數:"
         Height          =   255
         Left            =   2640
         TabIndex        =   87
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:          (Y: 是)"
         Height          =   180
         Left            =   6750
         TabIndex        =   86
         Top             =   1980
         Width           =   2085
      End
      Begin VB.Label Label30 
         Caption         =   "是否出名 :"
         Height          =   255
         Left            =   -72450
         TabIndex        =   76
         Top             =   1410
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label31 
         Caption         =   "(N:不出名)"
         Height          =   255
         Left            =   -71010
         TabIndex        =   75
         Top             =   1410
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   -74856
         TabIndex        =   73
         Top             =   1308
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   2385
         TabIndex        =   72
         Top             =   330
         Width           =   900
      End
      Begin MSForms.Label lblRCP14 
         Height          =   255
         Left            =   3120
         TabIndex        =   71
         Top             =   1155
         Visible         =   0   'False
         Width           =   1305
         VariousPropertyBits=   27
         Size            =   "11721;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "相關總收文號進度備註 :"
         Height          =   252
         Index           =   13
         Left            =   120
         TabIndex        =   70
         Top             =   1416
         Visible         =   0   'False
         Width           =   2064
      End
      Begin VB.Label Label1 
         Caption         =   "相關總收文號預估勝敗 :         (1:勝 2:敗 3:部分勝部分敗)"
         Height          =   255
         Index           =   8
         Left            =   4470
         TabIndex        =   69
         Top             =   1140
         Visible         =   0   'False
         Width           =   4350
      End
      Begin VB.Label Label1 
         Caption         =   "相關總收文號承辦人 :"
         Height          =   252
         Index           =   7
         Left            =   120
         TabIndex        =   68
         Top             =   1152
         Visible         =   0   'False
         Width           =   1848
      End
      Begin VB.Label Label1 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Index           =   5
         Left            =   4470
         TabIndex        =   66
         Top             =   870
         Width           =   1275
      End
      Begin VB.Label Label14 
         Caption         =   "關係人(日) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   65
         Top             =   960
         Width           =   1092
      End
      Begin VB.Label Label13 
         Caption         =   "關係人(英) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   660
         Width           =   972
      End
      Begin VB.Label Label7 
         Caption         =   "關係人(中) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   63
         Top             =   360
         Width           =   1092
      End
      Begin VB.Label Label26 
         Caption         =   "條款 :"
         Height          =   252
         Left            =   120
         TabIndex        =   62
         Top             =   2196
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74910
         TabIndex        =   61
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label12 
         Caption         =   "(1:勝 2:敗 3:部分勝部分敗)"
         Height          =   255
         Left            =   1530
         TabIndex        =   38
         Top             =   1710
         Width           =   2205
      End
      Begin VB.Label Label10 
         Caption         =   "預估勝敗 :"
         Height          =   252
         Left            =   120
         TabIndex        =   37
         Top             =   1704
         Width           =   972
      End
      Begin VB.Label Label1 
         Caption         =   "大約"
         Height          =   255
         Index           =   12
         Left            =   4470
         TabIndex        =   36
         Top             =   1980
         Width           =   495
      End
      Begin VB.Label Label11 
         Caption         =   "可接獲回音"
         Height          =   255
         Left            =   5730
         TabIndex        =   35
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "點數 :"
         Height          =   252
         Index           =   10
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   732
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1530
         TabIndex        =   33
         Top             =   1980
         Width           =   2745
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1980
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   252
         Left            =   120
         TabIndex        =   31
         Top             =   564
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   285
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "本案期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   29
         Top             =   2496
         Width           =   852
      End
      Begin VB.Label Label15 
         Caption         =   "是否閉卷 :"
         Height          =   255
         Left            =   3720
         TabIndex        =   28
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Left            =   5100
         TabIndex        =   27
         Top             =   1710
         Width           =   1095
      End
   End
   Begin MSForms.TextBox textTM81 
      Height          =   264
      Left            =   1200
      TabIndex        =   85
      TabStop         =   0   'False
      Top             =   1980
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   264
      Left            =   5190
      TabIndex        =   84
      TabStop         =   0   'False
      Top             =   1710
      Width           =   3885
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
      TabIndex        =   83
      TabStop         =   0   'False
      Top             =   1710
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   264
      Left            =   5190
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3885
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
      TabIndex        =   48
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5190
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1170
      Width           =   3885
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
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
      Left            =   5190
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   360
      Width           =   3885
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
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
      Left            =   5190
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   900
      Width           =   3885
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "6482;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1200
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   2250
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   17
      Left            =   4320
      TabIndex        =   81
      Top             =   1485
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   16
      Left            =   120
      TabIndex        =   80
      Top             =   1752
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   15
      Left            =   4320
      TabIndex        =   79
      Top             =   1755
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   120
      TabIndex        =   78
      Top             =   2022
      Width           =   720
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "機關文號 :"
      Height          =   180
      Left            =   4320
      TabIndex        =   60
      Top             =   2025
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   59
      Top             =   1482
      Width           =   720
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   4320
      TabIndex        =   58
      Top             =   1215
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人 :"
      Height          =   180
      Index           =   2
      Left            =   4320
      TabIndex        =   57
      Top             =   405
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   56
      Top             =   672
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   55
      Top             =   402
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   54
      Top             =   1212
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4320
      TabIndex        =   53
      Top             =   675
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4320
      TabIndex        =   52
      Top             =   945
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   51
      Top             =   942
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   120
      TabIndex        =   50
      Top             =   2310
      Width           =   810
   End
End
Attribute VB_Name = "frm020102_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/27 Form2.0已修改 textTM44/textCP13/textCP14/textCP44_2/lblCP14/textTM23(申請人名).../cmbTM05/textRCP64/textCP64/textCP40/textCP42/grdList/lblNameAgent
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
'2005/11/9 整理
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
'add by sonia 2021/9/22
Dim m_TM11 As String
Dim m_TM20 As String
Dim m_TM21 As String
Dim m_TM22 As String
'end 2021/9/22
'add by nickc 2007/02/01
Dim m_TM78 As String
Dim m_TM79 As String
Dim m_TM80 As String
Dim m_TM81 As String

' 案件性質代號
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
' 承辦人
Dim m_CP14 As String
' 相關總收文號
Dim m_CP43 As String
'2005/11/9 ADD BY SONIA
Dim m_strLanguage As String '定稿語文

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
Dim m_blnDelay As Boolean '判斷是否延期  2009/6/30 add by sonia

' 宣告代理人內容結構
Private Type AGENTITEM
   aiCode As String
   aiName As String
End Type
Dim m_AgentList() As AGENTITEM
Dim m_AgentCount As Integer
'
Dim m_CurrSel As Integer
'Add By Cheng 2003/11/19
Dim m_TM14 As String '註冊公告日
'add by nick  是否要同時發文第二期註冊費
Dim is715And716 As Boolean
'add by nick 2004/08/12
Dim m_CP84 As String       '發文規費
'add by nick 2004/09/27
Public m_CU103 As String         '公司負責人英文名稱
'add by nick 2004/10/05
Public m_CU05 As String         '客戶英文名稱
Public m_CU88 As String         '客戶英文名稱
Public m_CU89 As String         '客戶英文名稱
Public m_CU90 As String         '客戶英文名稱
' 申請人國籍 add by nick 2004/10/29
Public m_CU10 As String
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
'add by nickc 2006/09/11
Dim m_CP12 As String
Dim m_CP06 As String
Dim m_CP07 As String
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
Dim SeekCu10(1 To 5) As String
'2012/10/31 End
'add by nickc 2008/02/22
Dim m_CP44New As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_204or205 As Boolean                 '2009/4/8 add by sonia 是否繼續管制準備程序204言詞辯論205期限
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
Dim m_bolWebApp As Boolean 'Add by Sindy 2011/3/9 是否電子送件案
Dim m_QSP As Boolean 'Add by Lydia 2014/10/13 內商服務業務之台灣案發文
Dim m_CP16 As String 'Add By Sindy 2016/5/30 費用
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
Dim m_IPONumber As String 'Add By Sindy 2021/6/10
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
    'Add By Cheng 2002/12/03
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

Private Sub cmdok_Click(Index As Integer)
Dim strNewCP64 As String 'Add by Amy 2020/02/05 進度備註

   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
         If CheckDataValid = True Then
            'Add By Cheng 2002/05/23
            '重新檢查欄位有效性
            If TxtValidate = False Then Exit Sub
            'add by nick 2004/09/27 'edit by nick 2004/10/07
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
                  'GetCu103ByCustomer020102_16, m_TM23
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
                  'GetCu103ByCustomer020102_16, m_TM78
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
                  'GetCu103ByCustomer020102_16, m_TM79
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
                  'GetCu103ByCustomer020102_16, m_TM80
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
                  'GetCu103ByCustomer020102_16, m_TM81
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
            
            '2009/4/8 add by sonia 詢問是否繼續管制準備程序204言詞辯論205期限T-145587
            m_204or205 = False
            If m_CP09 < "C" And (m_CP10 = "204" Or m_CP10 = "205") Then
               Dim nResponse
               nResponse = MsgBox("是否繼續管制準備程序或言詞辯論期限？", vbYesNo + vbCritical + vbDefaultButton2, "發文")
               If nResponse = vbYes Then
                  m_204or205 = True
               End If
            End If
            '2009/4/8 end
            
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
                           'Modify By Sindy 2021/6/10 超項費預設同申請案的智慧局收文文號
                           strExc(0) = InputBox("請輸入智慧局收文文號!!", , m_IPONumber)
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
            
            'Added by Lydia 2017/06/29 出具同意書時，於點數欄下一行增加'對方案件號數'，不可空白。存檔時存入進度檔的CP30，並於進度備註加註'對方案件號數：……..'。例T-170200
            If m_CP10 = "723" And textCP30.Visible = True And Trim(textCP30.Text) <> "" Then
               Me.textCP64.Text = Me.textCP64.Text & IIf(Trim(Me.textCP64.Text) <> "", ";", "") & "對方案件號數：" & Trim(textCP30.Text) & ";"
            End If
            'end 2017/06/29
            
            ' 更新欄位輸入的內容
            OnUpdateField
            ' 存檔
            'Modify By Cheng 2002/11/07
            If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
            'Add By Cheng 2002/11/08
            ' 列印定稿
            If textPrint <> "N" Then
               PrintLetter
            'Add By Sindy 2021/3/31
            End If
            If textPrint = "N" Then
               If m_CP09 <> "" Then
                  Call PUB_TCaseAskIsPost(m_CP09)
               End If
            '2021/3/31 END
            End If
            
            'Add By Sindy 2018/12/4 註冊費發文時, 將繳費單電子檔歸檔於該「註冊費」程序
            If m_CP10 = "717" Then
               If PUB_TCheckCppPDF(textCP09, 1, True, , m_CP10) = False Then
                  If Val(m_CP07) >= 20190201 Then
                     MsgBox "沒有繳費單電子檔(.DATA.PDF)！", vbInformation
                  End If
               End If
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
                    'Add by Lydia 2014/10/13 內商服務業務(TC)之台灣案發文-規費與收文規費不符時,請加同時發給特殊設定人員"財務處總帳人員"
                    If m_QSP = True Then
                      PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, "A"
                    Else
                      '2020/01/13 Modify by Amy +if 傳strCP118參數
                      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                        PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", "")
                      Else
                         PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
                      End If
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
            If Index = 0 Then '確定鍵
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
               'Add By Cheng 2002/01/10
               frm020102_01.Clear1
               
               Unload Me
            ElseIf Index = 1 Then '同時發文鍵
               ' 呼叫第一個畫面
               frm020102_01.SetData 0, m_TM01, True
               frm020102_01.SetData 1, m_TM02, False
               frm020102_01.SetData 2, m_TM03, False
               frm020102_01.SetData 3, m_TM04, False
               frm020102_01.SetQueryFromTM
               Unload Me
               frm020102_01.Show
               frm020102_01.radio(1).Value = True
               frm020102_01.radio_Click 1
               frm020102_01.QueryData
            End If
         End If
      Case Else
   End Select
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

'Private Sub cmdTogether_Click()
'   If CheckDataValid = True Then
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
'
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'        ' 列印定稿
'        If textPrint <> "N" Then
'           PrintLetter
'        End If
'
'      ' 設定滑鼠游標為預設
'      Screen.MousePointer = vbDefault
'
'      ' 呼叫第一個畫面
'      frm020102_01.SetData 0, m_TM01, True
'      frm020102_01.SetData 1, m_TM02, False
'      frm020102_01.SetData 2, m_TM03, False
'      frm020102_01.SetData 3, m_TM04, False
'      frm020102_01.SetQueryFromTM
'      Unload Me
'      frm020102_01.Show
'      frm020102_01.radio(1).Value = True
'      frm020102_01.radio_Click 1
'      frm020102_01.QueryData
'   End If
'End Sub

Private Sub Form_Activate()
'add by nickc 2005/08/23
If (pub_ModifyCaseNum = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 And pub_ModifyCaseNum <> "") Then
   pub_ModifyCaseNum = ""
   QueryData
End If
End Sub

Private Sub Form_Load()

   is715And716 = False
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
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
      'Add By Cheng 2002/12/06
      Case 99: m_CP43 = strData
            Me.textCP43.Text = m_CP43
            ShowCP43Data m_CP43
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
'      If IsNull(rsTmp.Fields("TM12")) = False Then
'         textTM12 = rsTmp.Fields("TM12")
'      End If
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then
         'textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
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
        'Add By Cheng 2003/11/19
        '註冊公告日
        m_TM14 = "" & rsTmp.Fields("TM14").Value
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = "" & rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'add by sonia 2021/9/22 專用期間
      m_TM21 = Empty
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = "" & rsTmp.Fields("TM21")
      End If
      m_TM22 = Empty
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = "" & rsTmp.Fields("TM22")
      End If
      '申請日
      m_TM11 = Empty
      If IsNull(rsTmp.Fields("TM11")) = False Then
         m_TM11 = "" & rsTmp.Fields("TM11")
      End If
      '發證日
      m_TM20 = Empty
      If IsNull(rsTmp.Fields("TM20")) = False Then
         m_TM20 = "" & rsTmp.Fields("TM20")
      End If
      'end 2021/9/22
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = "" & rsTmp.Fields("TM78")
         textTM78 = GetCustomerName(rsTmp.Fields("TM78"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = "" & rsTmp.Fields("TM79")
         textTM79 = GetCustomerName(rsTmp.Fields("TM79"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = "" & rsTmp.Fields("TM80")
         textTM80 = GetCustomerName(rsTmp.Fields("TM80"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = "" & rsTmp.Fields("TM81")
         textTM81 = GetCustomerName(rsTmp.Fields("TM81"), 0)
      End If
      
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("tm24"))
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
      
      textTM136.Tag = "" & rsTmp.Fields("tm136") 'Added by Morgan 2022/12/15
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
         'textTM20 = TAIWANDATE(rsTmp.Fields("SP12"))
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
'add by nickc 2007/05/11 第二期發文時，直接上核准，且核准日為發文日
Dim strCP24 As String
Dim strCP25 As String
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
      
      'Add By Sindy 2016/5/30 費用
      m_CP16 = Empty
      If IsNull(rsTmp.Fields("CP16")) = False Then
         m_CP16 = rsTmp.Fields("CP16")
      End If
      '2016/5/30 END
      
      'Add By Cheng 2002/07/16
      '若案件性質為"補充答辯"(613)
      If m_CP10 = "613" Then
         Label1(5).Visible = True
         Me.textCP43.Visible = True
      End If
      
      'add by nickc 2007/05/11 第二期發文時，直接上核准，且核准日為發文日
      If m_CP10 = "716" Then
            strCP24 = Empty
            If IsNull(rsTmp.Fields("CP24")) = False Then
               strCP24 = rsTmp.Fields("CP24")
            End If
            SetCPFieldOldData "CP24", strCP24, 0
            strCP25 = Empty
            If IsNull(rsTmp.Fields("CP25")) = False Then
               strCP25 = rsTmp.Fields("CP25")
            End If
            SetCPFieldOldData "CP25", strCP25, 1
      End If
      
      ' 業務區別
      'add by nickc 2006/09/11
      m_CP06 = CheckStr(rsTmp.Fields("cp06"))
      m_CP07 = CheckStr(rsTmp.Fields("cp07"))
      m_CP12 = Empty
      
      If IsNull(rsTmp.Fields("CP12")) = False Then
         'add by nickc 2006/09/11
         m_CP12 = rsTmp.Fields("cp12")
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
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
      SetFrame1 'Added by Morgan 2022/12/15
      
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE("" & rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
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
      ' 相關總收文號
      m_CP43 = Empty
      If IsNull(rsTmp.Fields("CP43")) = False Then
         m_CP43 = rsTmp.Fields("CP43")
         '2005/11/17 ADD BY SONIA
         textCP43 = rsTmp.Fields("CP43")
      End If
      'add by nickc 2007/11/06 補收款要強制輸入相關收文號，商爭統計要
      If m_CP10 = "705" Then
          SetCPFieldOldData "CP43", textCP43, 0
      End If
      'Added by Lydia 2017/06/29 出具同意書 增加'對方案件號數'
      If m_CP10 = "723" Then
         SetCPFieldOldData "CP30", textCP30, 0
      End If
      'end 2017/06/23
      ' 條款
      textCP49 = Empty
      If IsNull(rsTmp.Fields("CP49")) = False Then
         textCP49 = rsTmp.Fields("CP49")
      End If
      SetCPFieldOldData "CP49", textCP49, 0
      ' 關係人(中)
      textCP40 = Empty
      If IsNull(rsTmp.Fields("CP40")) = False Then
         textCP40 = rsTmp.Fields("CP40")
      End If
      SetCPFieldOldData "CP40", textCP40, 0
      ' 關係人(英)
      textCP41 = Empty
      If IsNull(rsTmp.Fields("CP41")) = False Then
         textCP41 = rsTmp.Fields("CP41")
      End If
      SetCPFieldOldData "CP41", textCP41, 0
      ' 關係人(日)
      textCP42 = Empty
      If IsNull(rsTmp.Fields("CP42")) = False Then
         textCP42 = rsTmp.Fields("CP42")
      End If
      SetCPFieldOldData "CP42", textCP42, 0
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
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         ' 是否續辦欄位必須為空白
         If IsNull(rsTmp.Fields("NP06")) = False Then
            If IsEmptyText(rsTmp.Fields("NP06")) = False Then
               GoTo NextRecord
            End If
         End If
         
         grdList.Rows = grdList.Rows + 1
         grdList.row = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(grdList.row, 7) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            If m_TM10 = "000" Then
               grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 0)
            Else
               grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"), 1)
            End If
            grdList.TextMatrix(grdList.row, 8) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(grdList.row, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(grdList.row, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(grdList.row, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(grdList.row, 5) = rsTmp.Fields("NP14")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(grdList.row, 6) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(grdList.row, 9) = rsTmp.Fields("NP22")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/13
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/13
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
   Dim intLen As Integer, strTempCP64
   
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
   
   'Modify By Sindy 2012/7/26
   '台灣案才需顯示出名代理人
   lstNameAgent.Clear
   'Modify by Amy 2018/10/12 +len(m_CP10)
   If m_TM10 = "000" And Len(m_CP10) <> 4 Then
      lstNameAgent.Visible = True
      lblNameAgent.Visible = True
      'Modify by Amy 2021/12/27 改Form2.0,bForm2設True
      PUB_SetOurAgent lstNameAgent, tm(), m_CP110, , True
   Else
      lstNameAgent.Visible = False
      lblNameAgent.Visible = False
   End If
   '2012/7/26 End
   
   ' 大約?可接獲回音(欄位)
   textCF09 = Empty
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
   
    'Add By Cheng 2002/12/06
    ShowCP43Data m_CP43

   'add by nick 2004/07/01 檢查如果案件性質是第一註冊費的話，
   '                                    再檢查是否第二註冊費也存在，
   '                                    若存在，詢問是否要合併發文。
   If m_CP10 = "715" Then
        Dim nResponse
        strSql = "select * from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='716' and cp27 is null and cp57 is null "
        rsTmp.CursorLocation = adUseClient
        rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If rsTmp.RecordCount > 0 Then
            nResponse = MsgBox("是否同時發文第二期註冊費？", vbYesNo, "發文")
            If nResponse = vbYes Then
                 is715And716 = True
            End If
        End If
   End If
    Set rsTmp = Nothing
    
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
       textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'Add by Sindy 2011/3/9 檢查是否為電子送件案
   '2012/11/30 modify by sonia 加發文日為同日者
   'strExc(0) = "select 1 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('101') and cp118='Y' and cp27>0 and cp57 is null"
   'Modify By Sindy 2015/3/2 此sql run很久改語法
   'Modify By Sindy 2021/6/10 + or cp10||cp118='101A')
   'strExc(0) = "select 1 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10 in ('101') and cp118='Y' and cp27=" & DBDATE(textCP27) & " and cp57 is null"
   strExc(0) = "select cp64,cp118,cp85 from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "'" & _
               " and (cp10||cp118='101Y' or cp10||cp118='101A') and cp27" & IIf(textCP27 <> "", "=" & DBDATE(textCP27), ">0") & " and cp57 is null"
   '2015/3/2 END
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'Modify by Amy 2020/01/23 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   m_IPONumber = ""
   If intI = 1 Then
      m_bolWebApp = True
      
      'Modify By Sindy 2021/6/10 714.超項費預設同申請案的智慧局收文文號
      If Trim("" & RsTemp.Fields("CP64")) <> "" And m_CP10 = "714" Then
         If InStr(RsTemp.Fields("CP64"), "智慧局收文文號:") > 0 Then
            intLen = InStr(RsTemp.Fields("CP64"), "智慧局收文文號:")
            strTempCP64 = Mid(RsTemp.Fields("CP64"), intLen + Len("智慧局收文文號:"))
            m_IPONumber = Mid(strTempCP64, 1, InStr(strTempCP64, ";") - 1)
         End If
      End If
      '2021/6/10 END
      
      Label43.Visible = True
      textCP118.Visible = True
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
        lblPayToday.Visible = True
        txtPayToday.Visible = True
      End If
   Else
      m_bolWebApp = False
      'Modify By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
      'Modify By Sindy 2021/6/10 + FCT,也要電子送件
      If (m_TM01 = "T" Or m_TM01 = "FCT") And m_TM10 = "000" Then
         Label43.Visible = True
         textCP118.Visible = True
         If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
              lblPayToday.Visible = True
              txtPayToday.Visible = True
        End If
   'end 2020/01/13
      '2011/10/28 End
      Else
         Label43.Visible = False
         textCP118.Visible = False
      End If
   End If
   '714.超項費
   If m_bolWebApp = True And (m_CP10 = "714") And textCP118 <> "Y" Then
      MsgBox "本案應以電子送件方式呈送!!", vbExclamation
   End If
   '2011/3/9 End
   
   '2011/3/30 超項費預設不印定稿
   'Modified by Lydia 2016/12/22 收款寄證預設不印定稿 + Or m_CP10 = "1728"
   'Modify By Sindy 2021/3/31 + 案件性質為706(其它),定稿列印請自動上 "N"
   If m_CP10 = "714" Or m_CP10 = "1728" Or m_CP10 = "706" Then textPrint = "N"
   '2011/3/30 end
   'ADD BY SONIA 2016/3/29 跨類107,主張優先權108,文件公／簽證711若申請101案發文日為系統日時預設不印定稿 T-203184
   If m_CP10 = "107" Or m_CP10 = "108" Or m_CP10 = "711" Then
      strSql = "SELECT CP27 FROM CASEPROGRESS WHERE " & _
                     "CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "' AND " & _
                     "CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "' AND " & _
                     "CP10='101' AND CP09<'B'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         'Modify By Sindy 2020/9/7 ex:T-229684(9/4發文)
         'If "" & rsTmp.Fields(0) = strSrvDate(1) Then
         If "" & rsTmp.Fields(0) = DBDATE(textCP27) Then
         '2020/9/7 END
            textPrint = "N"
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   'END 2016/3/29
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
   
   'Added by Lydia 2017/06/29 出具同意書發文輸入對方案件號數
   If m_CP10 = "723" Then
      Label3.Visible = True
      textCP30.Visible = True
      Label3.Left = Label1(10).Left
      Label3.Top = Label1(7).Top
      textCP30.Left = 1400
      textCP30.Top = Label3.Top
      textCP30.Width = 2630
   Else
      Label3.Visible = False
      textCP30.Visible = False
   End If
   'end 2017/06/29
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2022/05/23
   
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    'Add By Cheng 2002/07/18
   Set frm020102_16 = Nothing
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
      Me.SSTab1.Tab = 1
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub

'Add By Sindy 2011/3/9
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/3/9
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
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

Private Sub textCP43_GotFocus()
TextInverse Me.textCP43
End Sub

Private Sub textCP43_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/06
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP43_LostFocus()
    'Add By Cheng 2002/12/06
    If Me.textCP43.Visible And m_CP10 = "613" Then
        ShowCP43Data Me.textCP43.Text
    End If
End Sub

Private Sub textCP43_Validate(Cancel As Boolean)
'Add By Cheng 2002/07/16
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

'若案件性質為"補充答辯"(613)
'Modify By Sindy 2021/5/20 626.註銷  2022/10/12取消626註銷,因為已改為307
If Me.textCP43.Visible And m_CP10 = "613" Then
   If Me.textCP43.Text = "" Then
      MsgBox "請輸入相關總收文號!!!", vbExclamation + vbOKOnly
      Cancel = True
   Else
      StrSQLa = "Select CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP09='" & Me.textCP43.Text & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If rsA.Fields("CP01").Value <> m_TM01 Or rsA.Fields("CP02").Value <> m_TM02 Or _
            rsA.Fields("CP03").Value <> m_TM03 Or rsA.Fields("CP04").Value <> m_TM04 Then
            MsgBox "相關總收文號的本所案號與畫面上的本所案號不同, 請重新輸入!!!", vbExclamation + vbOKOnly
            Cancel = True
         End If
      Else
         MsgBox "無此相關總收文號的本所案號資料, 請重新輸入!!!", vbExclamation + vbOKOnly
         Cancel = True
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
End If
If Cancel = True Then textCP43_GotFocus
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/03
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP49_KeyPress(KeyAscii As Integer)
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
        If m_TM10 <> 大陸國家代號 Then
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

' 關係人(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "關係人(中)內容太長"
      textCP40_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP40.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 關係人(英)
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "關係人(英)內容太長"
      textCP41_GotFocus
   End If
End Sub

' 關係人(日)
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 60) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "關係人(日)內容太長"
      textCP42_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then: textCP42.IMEMode = 2
   If Cancel = False Then CloseIme
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

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

Private Sub textRCP14_GotFocus()
    'Add By Cheng 2002/12/06
    TextInverse Me.textRCP14
End Sub

'Add By Sindy 2010/11/26
Private Sub textRCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textRCP14_Validate(Cancel As Boolean)
    If Me.textRCP14.Text <> "" Then
        Me.lblRCP14.Caption = "" & GetStaffName(Me.textRCP14.Text)
        If Me.lblRCP14.Caption = "" Then
            MsgBox "相關總收文號承辦人欄位輸入錯誤!!!", vbExclamation + vbOKOnly
            Cancel = True
            Me.textRCP14.SetFocus
            textRCP14_GotFocus
        End If
    End If
End Sub

Private Sub textRCP23_GotFocus()
    'Add By Cheng 2002/12/06
    TextInverse Me.textRCP23
End Sub

Private Sub textRCP23_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2002/12/06
    Select Case KeyAscii
    Case 8, 49, 50
        '無動作
    Case Else
        KeyAscii = 0
    End Select
End Sub

Private Sub textTM136_GotFocus()
   TextInverse textTM136
End Sub

Private Sub textTM136_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      KeyAscii = 0
   End If
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
      
      'Add By Sindy 2025/8/11 分析發文時發文日不可輸11/11/11
      If m_TM10 = 台灣國家代號 And m_CP10 = "727" And DBDATE(textCP27.Text) = "19221111" Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "台灣分析案發文日不可輸入11/11/11"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      '2025/8/11 END
   End If
   SetFrame1 'Added by Morgan 2022/12/15
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
   'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
   If m_TM10 <> 台灣國家代號 And m_TM01 <> "TD" Then
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
         'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
         If m_TM01 <> "TD" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "代理人不存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP44_GotFocus
         End If
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
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   Dim strCP64 As String
   
   ' 預估結果
   SetCPFieldNewData "CP23", textCP23
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
   
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
   'add by nickc 2007/05/11 若是第二期發文時，直接上核准，且核准日為發文日(簡易連絡單--宋若蘭)
   If m_CP10 = "716" Then
        SetCPFieldNewData "CP24", "1"
        SetCPFieldNewData "CP25", DBDATE(textCP27)
   End If
   
   ' 關係人(中)
   SetCPFieldNewData "CP40", textCP40
   ' 關係人(英)
   SetCPFieldNewData "CP41", textCP41
   ' 關係人(日)
   SetCPFieldNewData "CP42", textCP42
   'add by nickc 2007/11/06 補收款要強制輸入相關收文號，商爭統計要
   If m_CP10 = "705" Then
        SetCPFieldNewData "CP43", textCP43
   End If
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
   SetCPFieldNewData "CP45", textTM45
   ' 條款
   SetCPFieldNewData "CP49", textCP49
   ' 進度備註
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
      
   'Added by Lydia 2017/06/29 出具同意書 增加'對方案件號數'
   If m_CP10 = "723" Then
      SetCPFieldNewData "CP30", textCP30
   End If
   'end 2017/06/23
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

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim nIndex As Integer
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim strNP07 As String
Dim strNP08 As String
Dim strNP09 As String  'add by sonia 2021/9/22
Dim strNP22 As String
Dim objCopyCP As ClsCopyCP
Dim strCP06 As String
Dim strCP07 As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
Dim strCP09 As String, strCP48 As String 'Add By Sindy 2011/1/25
'Dim iErrNumber As Integer, iErrDescript As String 'Added by Lydia 2017/04/24
Dim bolRegMail As Boolean

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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/07
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
    
   'add by nick 2004/07/01 更新第二期註冊發文日
   If is715And716 = True Then
         strSql = "update caseprogress set cp27=" & ChangeTStringToWString(textCP27) & _
                        " where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & _
                        "' and cp03='" & m_TM03 & "' and cp04='" & m_TM04 & "' and cp10='716'"
         cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔的是否閉卷欄位
   Select Case m_TM01
      ' 系統類別為CFT的為儲存商標基本檔
      Case "T", "TF", "FCT":
         strSql = "UPDATE TradeMark SET TM29 = '" & textTM29 & "' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
         'add by nickc 2006/01/26
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
         strSql = "UPDATE ServicePractice SET SP15 = '" & textTM29 & "' " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
   End Select
   
    'Added by Morgan 2022/12/15
    '註冊證形式
    If textTM136.Visible And textTM136.Tag <> textTM136 Then
      strSql = "Update trademark Set tm136='" & textTM136 & "' " & _
                  "WHERE tm01='" & m_TM01 & "' and tm02='" & m_TM02 & "'" & _
                   " and tm03='" & m_TM03 & "' and tm04='" & m_TM04 & "'"
      cnnConnection.Execute strSql, intI
    End If
    'end 2022/12/15
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新相關總收文號之案件進度檔的條款
   If IsEmptyText(m_CP43) = False Then
        strSql = "UPDATE CaseProgress SET CP49 = '" & textCP49 & "' " & _
                     "WHERE CP09 = '" & m_CP43 & "' "
        cnnConnection.Execute strSql
   End If
    'Add By Cheng 2002/12/06
    '若案件性質為補充答辯, 更新相關總收文號之案件進度檔
    If IsEmptyText(m_CP43) = False And m_CP10 = "613" Then
        strSql = "UPDATE CaseProgress SET CP14 = '" & Me.textRCP14.Text & "', CP23='" & Me.textRCP23.Text & "',CP64='" & Me.textRCP64.Text & "'  " & _
                     "WHERE CP09 = '" & Me.textCP43.Text & "' "
        cnnConnection.Execute strSql
        'Add By Cheng 2004/02/11
        '更新發文資料的相關總收文號
        strSql = "UPDATE CaseProgress SET CP43 = '" & Me.textCP43.Text & "' WHERE CP09 = '" & m_CP09 & "' "
        cnnConnection.Execute strSql
        'End
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
         
         If Not (m_TM01 = "T" And m_TM10 = "020" And m_CP10 = "408") Then
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
      'Add By Sindy 2010/3/31
      ' 若有審查天數, 新增一筆催審期限的記錄到下一程序檔
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
      '2010/3/31 End
   End If
   rsTmp.Close
   '93.6.8 CANCEL BY SONIA 改在發證時才掛第二期註冊費期限
   
   'Add By Sindy 2010/12/22
   '內商大陸案案件性質401,403,408發文時,
   '請管制提申期限為法定期限(cp07)-2天,若法定期限(cp07)-2天<系統日管制提申期限為系統日
   If m_TM01 = "T" And m_TM10 = "020" And m_CP10 = "408" Then
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
      
      'Add By Sindy 2011/1/25
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
      '2011/1/25 End
   End If
   '2010/12/22 End
   
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 8)
         strNP22 = grdList.TextMatrix(nIndex, 9)
         strSql = "UPDATE NextProgress SET NP06 = 'Y' " & _
                  "WHERE NP02 = '" & m_TM01 & "' AND " & _
                        "NP03 = '" & m_TM02 & "' AND " & _
                        "NP04 = '" & m_TM03 & "' AND " & _
                        "NP05 = '" & m_TM04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   'add by nickc 2006/09/11 204、205要產生B 類資料 同 frm040104_3
   '2009/4/8 add by sonia 依使用者決定是否繼續管制準備程序204言詞辯論205期限
   'If m_CP09 < "C" And (m_CP10 = "204" Or m_CP10 = "205") Then
   If m_204or205 = True Then
   '2009/4/8 end
      strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP09," & _
         "CP10,CP20,CP26,CP32,CP11,CP05,CP06,CP07,CP12,CP13,CP14,CP43) VALUES ('" & m_TM01 & "','" & m_TM02 & _
         "','" & m_TM03 & "','" & m_TM04 & "','" & AutoNo("B", 6) & "','" & m_CP10 & _
         "','N','N','N','90'," & strSrvDate(1) & "," & CNULL(PUB_GetWorkDay1(m_CP06, True)) & "," & CNULL(m_CP07) & ",'" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & _
         "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & m_CP14 & "','" & m_CP09 & "')"
      cnnConnection.Execute strSql
   End If
   
   'add by nick 2004/08/12 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
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
   
   Set rsTmp = Nothing
   
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(m_CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2012/7/4
   '若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
      objCopyCP.CopyCaseProgress m_CP09, "", m_CP10, m_CP06
      Set objCopyCP = Nothing
   End If
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
   
   'Added by Lydia 2016/12/22 D類收款寄證(1728) 更新定稿資料,可以出定稿
   If Mid(m_CP09, 1, 1) = "D" And m_CP10 = "1728" Then
      'Modified by Lydia 2017/04/24 改成Function
      'Call PUB_UpdateET07LD0216("2", m_CP43, m_TM01, m_TM02, m_TM03, m_TM04, "00", strSrvDate(1))
      If PUB_UpdateET07LD0216("2", m_CP43, m_TM01, m_TM02, m_TM03, m_TM04, "00", strSrvDate(1)) = False Then
         GoTo ErrorHandler
      End If
      'end 2017/04/24
   End If
   'end 2016/12/22
   'Add by Amy 2019/12/04
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
        'If textPrint <> "N" Then
            'Sindy 2020/2/6 大至台不用判發
            '排除717.註冊費
            If Not (m_TM10 = "000" And textPrint = "2") And m_CP10 <> "717" Then
            '2020/2/6 END
               strExc(1) = Pub_GetSpecMan("內商程序客戶函發後補看人員")
            Else
               strExc(1) = ""
            End If
            'Add By Sindy 2025/7/21 ex:T252486~87 727=分析 為掛號直寄
            If m_TM01 = "T" And m_CP10 = "727" Then
               'Add By Sindy 2025/8/15 T分析案件,相關總收文號來函性質發文日若為11/11/11，則為掛號
               bolRegMail = False
               If m_CP43 <> "" Then
                  strSql = "SELECT * FROM caseprogress WHERE cp09='" & m_CP43 & "' AND CP27=19221111"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     bolRegMail = True
                  End If
               Else
                  bolRegMail = True
               End If
               '2025/8/15 END
               PUB_AddLetterProgress m_CP09, 0, True, , IIf(bolRegMail = True, True, False), m_TM23, m_CP10, m_TM44, , , , , strExc(1)
            Else
            '2025/7/21END
               PUB_AddLetterProgress m_CP09, 0, True, , False, m_TM23, m_CP10, m_TM44, , , , , strExc(1)
            End If
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
   
   'add by sonia 2021/9/22 若案件性質為使用宣誓且有專用期限
   If m_CP10 = "105" And Val("" & m_TM21) <> "0" And Val("" & m_TM22) <> "0" Then
      '若為菲律賓的申請日+3年使用宣誓期限則不必掛下一次
      If m_TM10 = "030" Then
         If m_TM11 + 30000 = m_CP07 Then
            GoTo Nextstep
         End If
      End If
   
      Dim ii As Integer
      Dim dblNewDate As Double
      ii = 0
      StrSQLa = "Select NA38,NA39 From Nation Where NA01='" & m_TM10 & "' AND NA39 IS NOT NULL "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
ReDo:
         ii = ii + 1
         '以發證日計算,無發證日才抓專用期起日
         If m_TM20 > 0 Then
            dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM20))))
         Else
            dblNewDate = DBDATE(DateAdd("yyyy", Val(rsA.Fields("NA38").Value) + Val(rsA.Fields("NA39").Value * ii), ChangeWStringToWDateString(DBDATE(m_TM21))))
         End If
         '若大於專用期止日無動作 2021/9/17 葉易雲於2021/8/27提出菲律賓2017/8/1新法延展核准後一年方再提出「延展使用宣誓」，故菲律賓不檢查專用期止日
         If dblNewDate > m_TM22 And m_TM10 <> "030" Then
         '若小於等於專用期止日
         Else
            '若小於此筆資料的法定期限+2年,再計算下一次
            If dblNewDate < DBDATE(DateAdd("yyyy", 2, ChangeWStringToWDateString(DBDATE(m_CP07)))) Then
               GoTo ReDo
            '若大於等於此筆資料的法定期限+2年
            Else
               '法定期限
               strNP09 = dblNewDate
               '本所期限 業務說改成本所=法定-2個月 不管任何國家
               strNP08 = DBDATE(DateAdd("m", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
               strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
               strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                        "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "',105," & _
                                    strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & GetNextProgressNo() & ")"
               cnnConnection.Execute strSql
            End If
         End If
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   'end 2021/9/22
   
Nextstep:
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

'    'Added by Lydia 2017/04/24 更新定稿日期失敗
'    If iErrNumber <> 0 Then
'       MsgBox iErrDescript
'    End If
'    'end 2017/04/24
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   grdList.Cols = 10
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "下一程序"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "本所期限"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "法定期限"
   grdList.ColWidth(3) = 1000
   grdList.col = 4
   grdList.Text = "機關文號"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "相關人"
   grdList.ColWidth(5) = 1200
   grdList.col = 6
   grdList.Text = "備註"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "收文號"
   grdList.ColWidth(7) = 0
   grdList.col = 8
   grdList.Text = "下一程序代號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "序號"
   grdList.ColWidth(9) = 0
End Sub

Private Sub grdList_Click()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            grdList.TextMatrix(grdList.row, 0) = "V"
         End If
      End If
   End If
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

' 檢查欄位是否都已輸入或是輸入的值是否正確
Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim ii As Integer   'add by sonia 2016/11/17
   
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
   'Modify By Sindy 2012/3/29 TD申請時皆在台灣申請不須控管CF代理人
   If m_TM10 >= "010" And m_TM01 <> "TD" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Modify By Cheng 2002/06/14
   '若案件性質為"其他"(706)或申請國家非台灣時, 預估勝敗欄可不輸入
   If m_CP10 = "706" Or m_TM10 > "010" Then
      '無動作
   Else
        'Modify By Cheng 2003/04/11
        '可不輸入
'      ' 預估勝敗不可為空白
'      If IsEmptyText(textCP23) = True Then
'         strTit = "檢核資料"
'         strMsg = "請輸入預估勝敗"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP23.SetFocus
'         GoTo EXITSUB
'      End If
   End If
   'add by nickc 2007/11/06 補收款要強制輸入相關收文號，商爭統計要
   If m_CP10 = "705" Then
        If IsEmptyText(textCP43) = True Then
           strTit = "檢核資料"
           strMsg = "請輸入相關總收文號"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP43.SetFocus
           GoTo EXITSUB
        End If
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
   
   'add by sonia 2016/11/17 +催審305(T-201684)
   If m_CP10 = "305" Then
      For ii = 1 To Me.grdList.Rows - 1
         If Me.grdList.TextMatrix(ii, 0) <> "" Then
            MsgBox "案件性質為<催審>，不可點選下一程序期限資料!!!", vbExclamation + vbOKOnly
            GoTo EXITSUB
         End If
      Next ii
   End If
   'end 2016/11/17
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
End Sub

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
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

Private Sub textCP44_GotFocus()
   InverseTextBox textCP44
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
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

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
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
    
    '2009/6/30 add by sonia判斷是否延期
    Select Case m_CP10
      Case "625"
          m_blnDelay = CheckDelay(m_CP09)
      Case Else
          m_blnDelay = False
    End Select
    '2009/6/30 End
   
   Select Case m_CP10
        'Add By Cheng 2003/01/10
        '查名(將來新系統正式上時是使用TS的)
        Case "001"
         ' 申請國家為大陸
         If m_TM10 = "020" Then
            'add by nickc 2006/06/30
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "00", strUserNum
               ' 案件性質分類
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                     "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
               cnnConnection.Execute strSql
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                     "','回音','" & "大約" & textCF09 & "後可接獲回音。')"
               '      "','回音','" & textCF09 & "')"
               cnnConnection.Execute strSql
             End If
         End If
        'Add By Cheng 2003/03/28
      ' 參加訴願, 參加訴訟, 行政訴訟上訴, 上訴答辯
'      Case "406", "407":
      Case "406", "407", "408", "410":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "02", strUserNum
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "03", strUserNum
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
               EndLetter "01", m_CP09, "00", strUserNum
               ' 案件性質分類
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "01" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & _
                     "','案件性質分類','" & GetCaseTypeName(m_TM01, m_CP10, 1) & "')"
               cnnConnection.Execute strSql
            End If
         End If
      ' 異議答辯
      Case "602":
         '2011/6/10 ADD BY SONIA 加入TD,TM
         Select Case m_TM01
            Case "TM"
            Case "TD"
            Case Else
         '2011/6/10 END
               ' 申請國家為台灣
               If m_TM10 < "010" Then
                  ' 申請人國籍為台灣
                  'edit by nickc 2006/06/29
                  'If strTM23Nation < "010" Then
                  If textPrint = "1" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "01", m_CP09, "24", strUserNum
                     ' 案件性質分類
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                              "'" & "商標狀況" & "','" & "審定" & "')"
                     cnnConnection.Execute strSql
                     ' 回音
                     strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                              "'" & "回音" & "','" & textCF09 & "')"
                     cnnConnection.Execute strSql
                  ' 申請人國籍非台灣
                  'edit by nickc 2006/06/29
                  'Else
                  ElseIf textPrint = "2" Then
                     ' 清除定稿例外欄位檔原有資料
                     EndLetter "01", m_CP09, "25", strUserNum
                  End If
               ' 申請國家為大陸
               ElseIf m_TM10 = "020" Then
                  ' 清除定稿例外欄位檔原有資料
                  'add by nickc 2006/06/29
                  If textPrint = "1" Then
                      EndLetter "01", m_CP09, "22", strUserNum
                  End If
               End If
         End Select   '2011/6/10 ADD BY SONIA
      ' 評定答辯
      Case "604":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "24", strUserNum
               ' 案件性質分類
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                        "'" & "商標狀況" & "','" & "註冊" & "')"
               cnnConnection.Execute strSql
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                        "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "25", strUserNum
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "22", strUserNum
            End If
         End If
      ' 廢止答辯
      Case "606":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "24", strUserNum
               ' 案件性質分類
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                        "'" & "商標狀況" & "','" & "註冊" & "')"
               cnnConnection.Execute strSql
               ' 回音
               strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        "VALUES ('" & "01" & "','" & m_CP09 & "','" & "24" & "','" & strUserNum & "'," & _
                        "'" & "回音" & "','" & textCF09 & "')"
               cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "25", strUserNum
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "22", strUserNum
            End If
         End If
        'Add By Cheng 2003/03/17
        '補充理由
        Case "612"
            ' 清除定稿例外欄位檔原有資料
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                EndLetter "01", m_CP09, "00", strUserNum
            End If
        '補充答辯
        Case "613"
            ' 清除定稿例外欄位檔原有資料
            'edit by nick 2004/10/29
            'If m_CU10 = "020" Then
            If textPrint = "2" Then
                EndLetter "01", m_CP09, "01", strUserNum
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "1" Then
                EndLetter "01", m_CP09, "00", strUserNum
            End If
        'Add By Cheng 2003/04/18
        '證據確認
        Case "621"
            '若申請國家為大陸
            If m_TM10 = 大陸國家代號 Then
                'add by nickc 2006/06/29
                If textPrint = "1" Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "22", strUserNum
                End If
            End If
    'Add By Cheng 2004/01/19
    '第一期註冊費, 第二期註冊費
    Case "715", "716", "717"
            Select Case textPrint
            Case "1", "2" '中文
               'add by nick  2004/07/01
               'edit by nick 2004/08/17
               'If is715And716 = True Then
               If is715And716 = True Or m_CP10 = "717" Then
                   ' 申請國家為台灣
                   If m_TM10 < "010" Then
                       ' 申請人國籍為台灣
                       'edit by nickc 2006/06/29
                       'If strTM23Nation < "010" Then
                       If textPrint = "1" Then
                           ' 列印定稿
                           EndLetter "01", m_CP09, "43", strUserNum
                           'add by nickc 2007/05/11 加入第二期才出現的字
                           If m_CP10 = "716" Then
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "43" & "','" & strUserNum & "'," & _
                                         "'第二期專用文字','" & vbCrLf & "　　附上智慧局的規費收據，證明第二期註冊費已繳納。繳第二期註冊費，智慧局現已不會再發給繳納第二期註冊費收訖書函。特此告知。" & vbCrLf & "')"
                                cnnConnection.Execute strSql
                            End If
                       ' 申請人國籍非台灣
                       'edit by nickc 2006/06/29
                       'Else
                       ElseIf textPrint = "2" Then
                           ' 列印定稿
                           EndLetter "01", m_CP09, "44", strUserNum
                           'add by nickc 2007/05/11 加入第二期才出現的字
                           If m_CP10 = "716" Then
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "44" & "','" & strUserNum & "'," & _
                                         "'第二期專用文字','" & vbCrLf & "　　附上智慧局的規費收據，證明第二期註冊費已繳納。繳第二期註冊費，智慧局現已不會再發給繳納第二期註冊費收訖書函。特此告知。" & vbCrLf & "')"
                                cnnConnection.Execute strSql
                            End If
                       End If
                   End If
               Else
                   ' 申請國家為台灣
                   If m_TM10 < "010" Then
                       ' 申請人國籍為台灣
                       'edit by nickc 2006/06/29
                       'If strTM23Nation < "010" Then
                       If textPrint = "1" Then
                           ' 列印定稿
                           EndLetter "01", m_CP09, "41", strUserNum
                           'add by nickc 2007/05/11 加入第二期才出現的字
                           If m_CP10 = "716" Then
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "41" & "','" & strUserNum & "'," & _
                                         "'第二期專用文字','" & vbCrLf & "　　附上智慧局的規費收據，證明第二期註冊費已繳納。繳第二期註冊費，智慧局現已不會再發給繳納第二期註冊費收訖書函。特此告知。" & vbCrLf & "')"
                                cnnConnection.Execute strSql
                            End If
                       ' 申請人國籍非台灣
                       'edit by nickc 2006/06/29
                       'Else
                       ElseIf textPrint = "2" Then
                           ' 列印定稿
                           EndLetter "01", m_CP09, "42", strUserNum
                           'add by nickc 2007/05/11 加入第二期才出現的字
                           If m_CP10 = "716" Then
                                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & "42" & "','" & strUserNum & "'," & _
                                         "'第二期專用文字','" & vbCrLf & "　　附上智慧局的規費收據，證明第二期註冊費已繳納。繳第二期註冊費，智慧局現已不會再發給繳納第二期註冊費收訖書函。特此告知。" & vbCrLf & "')"
                                cnnConnection.Execute strSql
                            End If
                       End If
                   End If
               End If
            Case "3"  '英文
               If is715And716 = True Or m_CP10 = "717" Then
                   ' 申請國家為台灣
                   If m_TM10 < "010" Then
                        ' 列印定稿
                        EndLetter "01", m_CP09, "45", strUserNum
                   End If
               End If
            '2005/11/9 END
            End Select
    'End
      '2009/6/30 add by sonia 參加異議625
      Case "625":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            If textPrint = "1" Then
                '本進度未延期
                If m_blnDelay = False Then
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
                '本進度延期
                Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "27", strUserNum
                End If
            ' 申請人國籍非台灣
            ElseIf textPrint = "2" Then
                '本進度未延期
                If m_blnDelay = False Then
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "23", strUserNum
                '本進度延期
                Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "26", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "26" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & GetDelayIssueDate(m_CP09) & "')"
                    cnnConnection.Execute strSql
                End If
            End If
         ' 申請國家非台灣
         Else
            If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, "18", strUserNum
            End If
         End If
        
      'Add By Sindy 2015/8/5
      ' 補優先權證明
      Case "208":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍非台灣 : 2.外->台
            If textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "01", strUserNum
               'Add By Sindy 2016/5/30 有費用
               If Val(m_CP16) > 0 Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                           "'" & "有費用" & "','及本所收費通知各乙紙')"
                  cnnConnection.Execute strSql
               End If
               '2016/5/30 END
            End If
         End If
      ' 暫緩審理
      Case "310":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍非台灣 : 2.外->台
            If textPrint = "2" Then
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "01", strUserNum
               'Add By Sindy 2016/5/30 有費用
               If Val(m_CP16) > 0 Then
                  strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           "VALUES ('" & "01" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                           "'" & "有費用" & "','及本所收費通知各乙紙')"
                  cnnConnection.Execute strSql
               End If
               '2016/5/30 END
            End If
         End If
      '2015/8/5 END
      
      'Add By Sindy 2024/5/21
      ' 加速審查
      Case "311":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            If textPrint = "1" Then '申請人國籍=台灣
               ' 清除定稿例外欄位檔原有資料
               EndLetter "01", m_CP09, "01", strUserNum
            End If
         End If
      '2024/5/21 END
      
      'Add By Sindy 2011/2/16
      Case "105"
          If m_TM01 = "TF" Then
             ' 清除定稿例外欄位檔原有資料
             EndLetter "01", m_CP09, "01", strUserNum
          End If
      '2011/2/16 End
      
      'Add By Sindy 2016/10/17
      ' 文件公／簽證
      Case "711":
         If textPrint = "1" Then '1.台->各國
            ' 清除定稿例外欄位檔原有資料
             EndLetter "01", m_CP09, "01", strUserNum
         ElseIf textPrint = "2" Then '2.外->台
            ' 清除定稿例外欄位檔原有資料
             EndLetter "01", m_CP09, "02", strUserNum
         End If
      '2016/10/17 END
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
   
   '2005/11/9 ADD BY SONIA
   '取得定稿語文
   m_strLanguage = GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/12
   ET01 = "01"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/12 End
   
   Select Case m_CP10
        'Add By Cheng 2003/01/10
        '查名(將來新系統正式上時是使用TS的)
        Case "001"
         ' 申請國家為大陸
         If m_TM10 = "020" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
               ET03 = "00" 'Modify By Sindy 2012/1/12
            End If
         End If
         
        'Add By Cheng 2003/03/28
      ' 參加訴願, 參加訴訟, 行政訴訟上訴, 上訴答辯
'      Case "406", "407":
      Case "406", "407", "408", "410":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "02", False, strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "03", False, strUserNum, 0
               ET03 = "03" 'Modify By Sindy 2012/1/12
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
               ET03 = "00" 'Modify By Sindy 2012/1/12
            End If
         End If
      ' 異議答辯
      Case "602":
         '2011/6/10 ADD BY SONIA 加入TD,TM
         Select Case m_TM01
            Case "TM"
               If m_TM10 <> "000" Then
                  If textPrint = "1" Then
                     ' 列印定稿
'                     NowPrint m_CP09, "01", "29", False, strUserNum, 0
                     ET03 = "29" 'Modify By Sindy 2012/1/12
                  End If
               End If
            Case "TD"
               If textPrint = "1" Then
'                  NowPrint m_CP09, "01", "00", False, strUserNum, 0
                  ET03 = "00" 'Modify By Sindy 2012/1/12
               End If
            Case Else
         '2011/6/10 END
               ' 申請國家為台灣
               If m_TM10 < "010" Then
                  ' 申請人國籍為台灣
                  'edit by nickc 2006/06/29
                  'If strTM23Nation < "010" Then
                  If textPrint = "1" Then
                     ' 列印定稿
'                     NowPrint m_CP09, "01", "24", False, strUserNum, 0
                     ET03 = "24" 'Modify By Sindy 2012/1/12
                  ' 申請人國籍非台灣
                  'edit by nickc 2006/06/29
                  'Else
                  ElseIf textPrint = "2" Then
                     ' 列印定稿
'                     NowPrint m_CP09, "01", "25", False, strUserNum, 0
                     ET03 = "25" 'Modify By Sindy 2012/1/12
                  End If
               ' 申請國家為大陸
               ElseIf m_TM10 = "020" Then
                  'add by nickc 2006/06/29
                  If textPrint = "1" Then
                      ' 列印定稿
'                      NowPrint m_CP09, "01", "22", False, strUserNum, 0
                     ET03 = "22" 'Modify By Sindy 2012/1/12
                  End If
               End If
         End Select  '2011/6/10 ADD BY SONIA
      ' 評定答辯
      Case "604":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "24", False, strUserNum, 0
               ET03 = "24" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "25", False, strUserNum, 0
               ET03 = "25" 'Modify By Sindy 2012/1/12
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "01", "22", False, strUserNum, 0
               ET03 = "22" 'Modify By Sindy 2012/1/12
            End If
         End If
      ' 廢止答辯
      Case "606":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by  nickc 2006/06/29
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "24", False, strUserNum, 0
               ET03 = "24" 'Modify By Sindy 2012/1/12
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/29
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
'               NowPrint m_CP09, "01", "25", False, strUserNum, 0
               ET03 = "25" 'Modify By Sindy 2012/1/12
            End If
         ' 申請國家為大陸
         ElseIf m_TM10 = "020" Then
            'add by nickc 2006/06/29
            If textPrint = "1" Then
                ' 列印定稿
'                NowPrint m_CP09, "01", "22", False, strUserNum, 0
               ET03 = "22" 'Modify By Sindy 2012/1/12
            End If
         End If
        'Add By Cheng 2003/03/17
      '補充理由
      Case "612"
          'add by nickc 2006/06/29
          If textPrint = "1" Then
              ' 列印定稿
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
             ET03 = "00" 'Modify By Sindy 2012/1/12
          End If
      '補充答辯
      Case "613"
          ' 列印定稿
          'edit by nick 2004/10/29
          'edit by nickc 2006/06/29
          'If m_CU10 = "020" Then
          If textPrint = "2" Then
'                NowPrint m_CP09, "01", "01", False, strUserNum, 0
             ET03 = "01" 'Modify By Sindy 2012/1/12
          'edit by nickc 2006/06/29
          'Else
          ElseIf textPrint = "1" Then
'                NowPrint m_CP09, "01", "00", False, strUserNum, 0
             ET03 = "00" 'Modify By Sindy 2012/1/12
          End If
      'Add By Cheng 2003/04/18
      '證據確認
      Case "621"
          ' 申請國家為大陸
          If m_TM10 = "020" Then
             'add by nickc 2006/06/29
             If textPrint = "1" Then
                  ' 列印定稿
'                    NowPrint m_CP09, "01", "22", False, strUserNum, 0
                ET03 = "22" 'Modify By Sindy 2012/1/12
             End If
          End If
      'Add By Cheng 2004/01/19
      '第一期註冊費, 第二期註冊費
      'edit by nick 2004/08/17
      'Case "715", "716"
      Case "715", "716", "717"
          '2005/11/9 MODIFY BY SONIA 加入定稿語文判斷
          'edit by nickc 2006/06/29
          'Select Case m_strLanguage
          Select Case textPrint
          Case "1", "2" '中文
             'add by nick  2004/07/01
             'edit by nick 2004/08/17
             'If is715And716 = True Then
             If is715And716 = True Or m_CP10 = "717" Then
                 ' 申請國家為台灣
                 If m_TM10 < "010" Then
                     ' 申請人國籍為台灣
                     'edit by nickc 2006/06/29
                     'If strTM23Nation < "010" Then
                     If textPrint = "1" Then
                         ' 列印定稿
'                           NowPrint m_CP09, "01", "43", False, strUserNum, 0
                         ET03 = "43" 'Modify By Sindy 2012/1/12
                     ' 申請人國籍非台灣
                     'edit by nickc 2006/06/29
                     'Else
                     ElseIf textPrint = "2" Then
                         ' 列印定稿
'                           NowPrint m_CP09, "01", "44", False, strUserNum, 0
                         ET03 = "44" 'Modify By Sindy 2012/1/12
                     End If
                 End If
             Else
                 ' 申請國家為台灣
                 If m_TM10 < "010" Then
                     ' 申請人國籍為台灣
                     'edit by nickc 2006/06/29
                     'If strTM23Nation < "010" Then
                     If textPrint = "1" Then
                         ' 列印定稿
'                           NowPrint m_CP09, "01", "41", False, strUserNum, 0
                         ET03 = "41" 'Modify By Sindy 2012/1/12
                     ' 申請人國籍非台灣
                     'edit by nickc 2006/06/29
                     'Else
                     ElseIf textPrint = "2" Then
                         ' 列印定稿
'                           NowPrint m_CP09, "01", "42", False, strUserNum, 0
                         ET03 = "42" 'Modify By Sindy 2012/1/12
                     End If
                 End If
             End If
          '2005/11/9 ADD BY SONIA
          Case "3"  '英文
             If is715And716 = True Or m_CP10 = "717" Then
                 ' 申請國家為台灣
                 If m_TM10 < "010" Then
                      ' 列印定稿
'                        NowPrint m_CP09, "01", "45", False, strUserNum, 0
                   ET03 = "45" 'Modify By Sindy 2012/1/12
                 End If
             Else
                 ' 申請國家為台灣
                 If m_TM10 < "010" Then
                      ' 列印定稿
                      'NowPrint m_CP09, "01", "??", False, strUserNum, 0
                 End If
             End If
          '2005/11/9 END
          End Select
       '2009/6/30 add by sonia 625參加異議
       Case "625"
          ' 申請國家為台灣
          If m_TM10 < "010" Then
             ' 申請人國籍為台灣
             If textPrint = "1" Then
                 '本進度未延期
                 If m_blnDelay = False Then
'                       NowPrint m_CP09, "01", "17", False, strUserNum, 0
                   ET03 = "17" 'Modify By Sindy 2012/1/12
                 '本進度延期
                 Else
'                       NowPrint m_CP09, "01", "27", False, strUserNum, 0
                   ET03 = "27" 'Modify By Sindy 2012/1/12
                 End If
             ' 申請人國籍非台灣
             ElseIf textPrint = "2" Then
                 '本進度未延期
                 If m_blnDelay = False Then
'                       NowPrint m_CP09, "01", "23", False, strUserNum, 0
                   ET03 = "23" 'Modify By Sindy 2012/1/12
                 '本進度延期
                 Else
                     ' 列印定稿
'                       NowPrint m_CP09, "01", "26", False, strUserNum, 0
                   ET03 = "26" 'Modify By Sindy 2012/1/12
                 End If
             End If
          '申請國家非台灣
          Else
             If textPrint = "1" Then
'                   NowPrint m_CP09, "01", "18", False, strUserNum, 0
                ET03 = "18" 'Modify By Sindy 2012/1/12
             End If
          End If
       '2009/6/30 end
      'Add By Sindy 2011/2/16
      Case "105"
          If m_TM01 = "TF" Then
'               NowPrint m_CP09, "01", "01", False, strUserNum, 0
             ET03 = "01" 'Modify By Sindy 2012/1/12
          End If
      '2011/2/16 End
      
      'Add By Sindy 2015/8/5
      ' 補優先權證明
      Case "208":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍非台灣 : 2.外->台
            If textPrint = "2" Then
               ET03 = "01"
            End If
         End If
      ' 暫緩審理
      Case "310":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            ' 申請人國籍非台灣 : 2.外->台
            If textPrint = "2" Then
               ET03 = "01"
            'Added by Lydia 2020/03/06 申請人國籍=台灣
            ElseIf textPrint = "1" Then
               ET03 = "02"
            End If
         End If
      '2015/8/5 END
      'Add By Sindy 2024/5/21
      ' 加速審查
      Case "311":
         ' 申請國家為台灣
         If m_TM10 < "010" Then
            If textPrint = "1" Then '申請人國籍=台灣
               ET03 = "01"
            End If
         End If
      '2024/5/21 END
      'Add By Sindy 2016/10/17
      ' 文件公／簽證
      Case "711":
         If textPrint = "1" Then '1.台->各國
            ET03 = "01"
         ElseIf textPrint = "2" Then '2.外->台
            ET03 = "02"
         End If
      '2016/10/17 END
      '2011/3/4 add by sonia 加通用定稿
      Case Else
          ' 申請國家為台灣
          If m_TM10 < "010" Then
             ' 申請人國籍為台灣
             If textPrint = "1" Then
'                  NowPrint m_CP09, "01", "30", False, strUserNum, 0
                ET03 = "30" 'Modify By Sindy 2012/1/12
             ' 申請人國籍非台灣
             ElseIf textPrint = "2" Then
'                  NowPrint m_CP09, "01", "31", False, strUserNum, 0
                ET03 = "31" 'Modify By Sindy 2012/1/12
             End If
          '申請國家非台灣
          Else
             If textPrint = "1" Then
'                  NowPrint m_CP09, "01", "32", False, strUserNum, 0
                ET03 = "32" 'Modify By Sindy 2012/1/12
             End If
          End If
         '2011/3/4 end
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
   
   'add by nick 2004/08/12 發文規費，申請國家台灣才檢查
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
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
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
   
   If Me.textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
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
   
   If Me.textTM29.Enabled = True Then
      Cancel = False
      textTM29_Validate Cancel
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
   
   'Add by Sindy 2011/3/9 714.超項費
   If m_bolWebApp = True And (m_CP10 = "714") And textCP118 <> "Y" Then
      MsgBox "本案應以電子送件方式呈送!!", vbExclamation
      textCP118.SetFocus
      Exit Function
   End If
   '2011/3/9 End
   
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
   
   'Added by Lydia 2017/06/29
   If textCP30.Visible = True Then
      textCP30_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'end 2017/06/29
   
    'Added by Lydia 2021/06/04 ACS智財顧問專業分配比例管制：有相關卷號(CaseRelation1)為ACS且曾有收文智財顧問112
    If Pub_ChkACS112isNull(m_TM01, m_TM02, m_TM03, m_TM04, txtCP113) = True Then
        SSTab1.Tab = 0
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   'Added by Morgan 2022/12/15
   If textTM136.Visible And textTM136.Enabled Then
      If textTM136 = "" Then
         MsgBox "請輸入註冊證形式！", vbExclamation
         textTM136.SetFocus
         Exit Function
      ElseIf textTM136.Tag <> "" And textTM136 <> textTM136.Tag Then
         If MsgBox("您輸入的註冊證形式為【" & IIf(textTM136 = "1", "電子", "紙本") & "】與分案設定【" & IIf(textTM136.Tag = "1", "電子", "紙本") & "】不同是否確定要繼續？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
            textTM136.SetFocus
            textTM136_GotFocus
            Exit Function
         End If
      End If
   End If
   'end 2022/12/15
    
   TxtValidate = True
End Function

'Added by Morgan 2022/12/15
'台灣112年以後繳註冊費需輸入形式
Private Sub SetFrame1()
   Frame1.Visible = False
   If m_TM10 = "000" And Len(m_CP10) = 3 Then
      If PUB_TWCertPty(m_TM01, m_CP10, m_TM02, m_TM03, m_TM04) = True Then
         Frame1.Visible = True
         If DBDATE(textCP27) > "20230000" Then
            textTM136.Enabled = True
         Else
            textTM136 = ""
            textTM136.Enabled = False
         End If
      End If
   End If
End Sub

'edit by nickc 2006/01/27
' 91.09.02 modify by louis
'Private Sub textAgName_GotFocus()
'   InverseTextBox textAgName
'   textAgName.IMEMode = 1
'End Sub

' 91.09.02 modify by louis
' 本所出名代理人
'edit by nickc 2006/01/27
'Private Sub textAgName_Validate(Cancel As Boolean)
'   Cancel = False
'   If CheckLengthIsOK(textAgName, 10) = False Then
'      Cancel = True
'   End If
'   If Cancel = False Then: textAgName.IMEMode = 2
'End Sub

Private Sub ShowCP43Data(strCP43 As String)
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
'Add By Cheng 2004/02/12
Dim arrCP64
Dim ii As Integer
'End
    'Add By Cheng 2002/12/06
    '若案件性質為補充答辯, 則抓相關總收文號的承辦人, 預估勝敗, 進度備註
    If m_CP10 = "613" Then
        Me.Label1(7).Visible = True
        Me.Label1(8).Visible = True
        Me.Label1(13).Visible = True
        Me.lblRCP14.Visible = True
        Me.textRCP14.Visible = True
        Me.textRCP23.Visible = True
        Me.textRCP64.Visible = True
        Me.textRCP14.Enabled = True
        Me.textRCP23.Enabled = True
        Me.textRCP64.Enabled = True
        StrSQLa = "Select CP14, ST02, CP23, CP64 From CaseProgress,Staff Where CP14=ST01(+) And CP09='" & strCP43 & "'"
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            Me.textCP43.Text = strCP43
            m_CP43 = Me.textCP43.Text
            Me.textRCP14.Text = "" & rsA.Fields(0).Value
            Me.lblRCP14.Caption = "" & rsA.Fields(1).Value
            Me.textRCP23.Text = "" & rsA.Fields(2).Value
            Me.textRCP64.Text = "" & rsA.Fields(3).Value
            'Add By Cheng 2004/02/12
            '預設本所出名代理人
'edit nickc 2006/01/27
'            If "" & rsA.Fields(3).Value <> "" Then
'                arrCP64 = Split("" & rsA.Fields(3).Value, ",")
'                For ii = LBound(arrCP64) To UBound(arrCP64)
'                    If InStr(arrCP64(ii), "本所出名代理人:") > 0 Then
'                        Me.textAgName.Text = Replace(arrCP64(ii), "本所出名代理人:", "")
'                        Exit For
'                    End If
'                Next ii
'            End If
            'End
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
End Sub
'add by nick 2004/09/27
'傳入客戶編號
'回傳公司負責人英文名稱
'Mark by Lydia 2024/07/03
' Sub GetCu103ByCustomer020102_16(oForm As Form, ByVal oCu As String)
'   CheckOC3
'oCu = oCu & "00000000"
'strSql = "SELECT * FROM Customer " & _
'         "WHERE CU01 = '" & Mid(oCu, 1, 8) & "' AND " & _
'               "CU02 = '" & Mid(oCu, 9, 1) & "'"
'   AdoRecordSet3.CursorLocation = adUseClient
'   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If AdoRecordSet3.RecordCount > 0 Then
'      AdoRecordSet3.MoveFirst
'      oForm.m_CU103 = CheckStr(AdoRecordSet3.Fields("CU103").Value)
'      oForm.m_CU05 = CheckStr(AdoRecordSet3.Fields("CU05").Value)
'      oForm.m_CU88 = CheckStr(AdoRecordSet3.Fields("CU88").Value)
'      oForm.m_CU89 = CheckStr(AdoRecordSet3.Fields("CU89").Value)
'      oForm.m_CU90 = CheckStr(AdoRecordSet3.Fields("CU90").Value)
'      oForm.m_CU10 = CheckStr(AdoRecordSet3.Fields("CU10").Value)
'      'add by nickc 2006/01/20
'      oForm.m_CU112 = CheckStr(AdoRecordSet3.Fields("CU112").Value)
'      'edit by nickc 2007/08/10
'      'Add By Sindy 2012/2/8
'      oForm.m_CU39 = CheckStr(AdoRecordSet3.Fields("CU39").Value)
'      oForm.m_CU40 = CheckStr(AdoRecordSet3.Fields("CU40").Value)
'      oForm.m_CU41 = CheckStr(AdoRecordSet3.Fields("CU41").Value)
'      '2012/2/8 End
'    Else
'        oForm.m_CU103 = ""
'        oForm.m_CU05 = ""
'        oForm.m_CU88 = ""
'        oForm.m_CU89 = ""
'        oForm.m_CU90 = ""
'        oForm.m_CU10 = ""
'        oForm.m_CU112 = ""
'        'Add By Sindy 2012/2/8
'        oForm.m_CU39 = ""
'        oForm.m_CU40 = ""
'        oForm.m_CU41 = ""
'        '2012/2/8 End
'    End If
'CheckOC3
'End Sub
'end 2024/07/03

'2009/6/30 add by sonia COPY FROM frm020102_14
'判斷是否延期(303)
Private Function CheckDelay(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   StrSQLa = "Select * From CaseProgress where CP43='" & m_CP09 & "' And CP10='303' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      CheckDelay = True
   Else
      CheckDelay = False
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Function

'取得延期的發文日
Private Function GetDelayIssueDate(strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "Select CP27 From Caseprogress Where CP43='" & strCP09 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetDelayIssueDate = "" & rsA.Fields(0).Value
Else
    GetDelayIssueDate = ""
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function
'2009/6/30 end

'Added by Lydia 2017/06/29
Private Sub textCP30_GotFocus()
TextInverse Me.textCP30
End Sub

Private Sub textCP30_Validate(Cancel As Boolean)

If textCP30.Visible = True Then
   If Trim(textCP30) = "" Then
       MsgBox "請輸入對方案件號數!", vbCritical
       textCP30_GotFocus
       Cancel = True
   Else
       If CheckLengthIsOK(textCP30, textCP30.MaxLength) = False Then
          textCP30_GotFocus
          Cancel = True
       End If
   End If
End If

End Sub
'end 2017/06/29

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
