VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm020102_11 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(延期)"
   ClientHeight    =   5880
   ClientLeft      =   3768
   ClientTop       =   1932
   ClientWidth     =   9132
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   9132
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   972
      Width           =   2745
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   660
      Width           =   3495
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1230
      Width           =   3252
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   390
      Width           =   3252
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   660
      Width           =   3252
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6996
      TabIndex        =   24
      Top             =   -15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6180
      TabIndex        =   23
      Top             =   -15
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8208
      TabIndex        =   25
      Top             =   -15
      Width           =   800
   End
   Begin VB.TextBox textDL02 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   2010
      Width           =   3495
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3252
      Left            =   60
      TabIndex        =   48
      Top             =   2580
      Width           =   9048
      _ExtentX        =   15960
      _ExtentY        =   5736
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm020102_11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(12)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(10)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label25"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label28"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label10"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label39"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblNameAgent"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label12"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label43"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblPayToday"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblCP113(18)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textCP44_2"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP64"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "lstNameAgent"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "grdList"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCF09"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP22"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP44"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP18"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textPrint"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP06"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP07"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt(0)"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP27"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP84"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text9"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP118"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtPayToday"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCP113"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "對造資料"
      TabPicture(1)   =   "frm020102_11.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP36"
      Tab(1).Control(1)=   "textCP38"
      Tab(1).Control(2)=   "textCP41"
      Tab(1).Control(3)=   "textCP80"
      Tab(1).Control(4)=   "textCP37_1"
      Tab(1).Control(5)=   "textCP40"
      Tab(1).Control(6)=   "textCP37"
      Tab(1).Control(7)=   "textCP39"
      Tab(1).Control(8)=   "textCP42"
      Tab(1).Control(9)=   "Label15"
      Tab(1).Control(10)=   "Label14"
      Tab(1).Control(11)=   "Label13"
      Tab(1).Control(12)=   "Label17"
      Tab(1).Control(13)=   "Label18"
      Tab(1).Control(14)=   "Label19"
      Tab(1).Control(15)=   "Label20"
      Tab(1).Control(16)=   "Label21"
      Tab(1).Control(17)=   "Label16"
      Tab(1).ControlCount=   18
      Begin VB.TextBox txtCP113 
         Height          =   270
         Left            =   6990
         MaxLength       =   4
         TabIndex        =   12
         Top             =   1440
         Width           =   540
      End
      Begin VB.TextBox txtPayToday 
         Height          =   264
         Left            =   4335
         MaxLength       =   1
         TabIndex        =   11
         Top             =   1470
         Width           =   255
      End
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   1260
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1470
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   264
         Left            =   930
         MaxLength       =   2
         TabIndex        =   5
         Top             =   900
         Width           =   372
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   3405
         TabIndex        =   1
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   810
         MaxLength       =   7
         TabIndex        =   0
         Top             =   330
         Width           =   1092
      End
      Begin VB.TextBox textCP36 
         Height          =   330
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   14
         Top             =   420
         Width           =   6912
      End
      Begin VB.TextBox textCP38 
         Height          =   264
         Left            =   -73200
         MaxLength       =   100
         TabIndex        =   20
         Top             =   2424
         Width           =   6912
      End
      Begin VB.TextBox textCP41 
         Height          =   330
         Left            =   -73200
         MaxLength       =   600
         TabIndex        =   16
         Top             =   1152
         Width           =   6912
      End
      Begin VB.TextBox textCP80 
         Height          =   330
         Left            =   -73200
         MaxLength       =   39
         TabIndex        =   22
         Top             =   2760
         Width           =   6912
      End
      Begin VB.TextBox txt 
         Height          =   264
         Index           =   0
         Left            =   6300
         MaxLength       =   5
         TabIndex        =   2
         Top             =   330
         Width           =   1092
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5370
         MaxLength       =   8
         TabIndex        =   7
         Top             =   900
         Width           =   1155
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   2760
         MaxLength       =   8
         TabIndex        =   6
         Top             =   900
         Width           =   1155
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   930
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1185
         Width           =   372
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   7200
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   615
         Width           =   1428
      End
      Begin VB.ComboBox textCP44 
         Height          =   276
         Left            =   930
         TabIndex        =   4
         Top             =   615
         Width           =   1500
      End
      Begin VB.TextBox textCP22 
         Height          =   264
         Left            =   8595
         MaxLength       =   1
         TabIndex        =   3
         Top             =   330
         Visible         =   0   'False
         Width           =   372
      End
      Begin VB.TextBox textCF09 
         Height          =   264
         Left            =   5130
         MaxLength       =   12
         TabIndex        =   9
         Top             =   1185
         Width           =   612
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   924
         Left            =   1200
         TabIndex        =   88
         Top             =   1776
         Width           =   7764
         _ExtentX        =   13695
         _ExtentY        =   1630
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
         Left            =   7470
         TabIndex        =   87
         Top             =   900
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;882"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   924
         Left            =   -73200
         TabIndex        =   18
         Top             =   1824
         Width           =   6912
         VariousPropertyBits=   679493659
         ScrollBars      =   2
         Size            =   "12192;1630"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   336
         Left            =   -73200
         TabIndex        =   15
         Top             =   792
         Width           =   6912
         VariousPropertyBits=   679493659
         MaxLength       =   600
         Size            =   "12192;593"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73200
         TabIndex        =   19
         Top             =   2136
         Width           =   6912
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73200
         TabIndex        =   21
         Top             =   2220
         Width           =   6912
         VariousPropertyBits=   679493659
         MaxLength       =   100
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73200
         TabIndex        =   17
         Top             =   1512
         Width           =   6912
         VariousPropertyBits=   679493659
         MaxLength       =   600
         Size            =   "12192;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   480
         Left            =   1176
         TabIndex        =   13
         Top             =   2736
         Width           =   7812
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13779;847"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP44_2 
         Height          =   270
         Left            =   2445
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   615
         Width           =   3945
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "6964;466"
         SpecialEffect   =   0
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
         Left            =   6090
         TabIndex        =   86
         Top             =   1470
         Width           =   765
      End
      Begin VB.Label lblPayToday 
         AutoSize        =   -1  'True
         Caption         =   "電子送件是否當日扣款:         (Y/N)"
         Height          =   180
         Left            =   2400
         TabIndex        =   85
         Top             =   1500
         Width           =   2655
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:          (Y: 是)"
         Height          =   180
         Left            =   90
         TabIndex        =   84
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Label Label12 
         Caption         =   "延期月數 :"
         Height          =   180
         Left            =   90
         TabIndex        =   83
         Top             =   930
         Width           =   810
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   6570
         TabIndex        =   74
         Top             =   930
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   2475
         TabIndex        =   73
         Top             =   345
         Width           =   900
      End
      Begin VB.Label Label15 
         Caption         =   "對造案件名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   72
         Top             =   1848
         Width           =   1572
      End
      Begin VB.Label Label14 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   71
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   70
         Top             =   1872
         Width           =   1572
      End
      Begin VB.Label Label17 
         Caption         =   "對造案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   69
         Top             =   2136
         Width           =   1572
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   68
         Top             =   2472
         Width           =   1572
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   67
         Top             =   840
         Width           =   1572
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   66
         Top             =   1176
         Width           =   1572
      End
      Begin VB.Label Label21 
         Caption         =   "對造日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   65
         Top             =   1536
         Width           =   1572
      End
      Begin VB.Label Label16 
         Caption         =   "對造案件商品類別 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   2760
         Width           =   1572
      End
      Begin VB.Label Label10 
         Caption         =   "延期後法定期限 :"
         Height          =   255
         Left            =   3960
         TabIndex        =   63
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "延期後本所期限 :"
         Height          =   255
         Left            =   1350
         TabIndex        =   62
         Top             =   930
         Width           =   1455
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   90
         TabIndex        =   61
         Top             =   2760
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   255
         Left            =   90
         TabIndex        =   60
         Top             =   345
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "代理人 :"
         Height          =   255
         Left            =   90
         TabIndex        =   59
         Top             =   645
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   90
         TabIndex        =   58
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1350
         TabIndex        =   57
         Top             =   1230
         Width           =   2745
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "點數 :"
         Height          =   180
         Index           =   10
         Left            =   6690
         TabIndex        =   56
         Top             =   645
         Width           =   450
      End
      Begin VB.Label Label9 
         Caption         =   "欲延期期限 :"
         Height          =   228
         Left            =   90
         TabIndex        =   55
         Top             =   1800
         Width           =   996
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "大約"
         Height          =   180
         Index           =   12
         Left            =   4710
         TabIndex        =   54
         Top             =   1230
         Width           =   360
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "可接獲回音"
         Height          =   180
         Left            =   5730
         TabIndex        =   53
         Top             =   1230
         Width           =   900
      End
      Begin MSForms.Label lbl 
         Height          =   255
         Index           =   0
         Left            =   5235
         TabIndex        =   52
         Top             =   345
         Width           =   1095
         VariousPropertyBits=   27
         Caption         =   "延期承辦人 :"
         Size            =   "11721;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lbl 
         Height          =   255
         Index           =   1
         Left            =   7425
         TabIndex        =   51
         Top             =   345
         Width           =   1125
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "1984;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1200
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   2280
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
   Begin MSForms.TextBox textTM81 
      Height          =   264
      Left            =   1200
      TabIndex        =   82
      TabStop         =   0   'False
      Top             =   2010
      Width           =   3252
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "5736;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   264
      Left            =   5610
      TabIndex        =   81
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3495
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6165;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   264
      Left            =   1200
      TabIndex        =   80
      TabStop         =   0   'False
      Top             =   1740
      Width           =   3252
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "5736;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   264
      Left            =   5610
      TabIndex        =   79
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3495
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "6165;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5610
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   930
      Width           =   3495
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6165;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   264
      Left            =   5610
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   390
      Width           =   3495
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6165;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   264
      Left            =   5610
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1200
      Width           =   3495
      VariousPropertyBits=   679493663
      MaxLength       =   20
      Size            =   "6165;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1200
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1470
      Width           =   3252
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "5736;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Index           =   5
      Left            =   4590
      TabIndex        =   78
      Top             =   1512
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Index           =   7
      Left            =   90
      TabIndex        =   77
      Top             =   1782
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Index           =   13
      Left            =   4590
      TabIndex        =   76
      Top             =   1782
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Index           =   14
      Left            =   90
      TabIndex        =   75
      Top             =   2052
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   120
      TabIndex        =   47
      Top             =   2340
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數/申請案號 :"
      Height          =   180
      Left            =   120
      TabIndex        =   46
      Top             =   972
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   4590
      TabIndex        =   45
      Top             =   972
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4590
      TabIndex        =   44
      Top             =   702
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   43
      Top             =   1242
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   42
      Top             =   432
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   41
      Top             =   702
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FC代理人 :"
      Height          =   180
      Index           =   2
      Left            =   4590
      TabIndex        =   40
      Top             =   435
      Width           =   840
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   4590
      TabIndex        =   39
      Top             =   1242
      Width           =   630
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   120
      TabIndex        =   38
      Top             =   1512
      Width           =   720
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "上次延期日 :"
      Height          =   180
      Left            =   4590
      TabIndex        =   37
      Top             =   2055
      Width           =   990
   End
End
Attribute VB_Name = "frm020102_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/27 Form2.0已修改 textTM44/textCP13/textCP14/textCP44_2/textTM23(申請人名)/cmbTM05/lbl()/textCP64/textCP40/textCP42/textCP37_1/textCP37/textCP39/lstNameAgent
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

' 案件性質代號
Dim m_CP10 As String
' 本所期限
Dim m_CP06 As String
' 法定期限
Dim m_CP07 As String
' 智權人員代號
Dim m_CP13 As String
' 業務區
Dim m_CP12 As String
'相關總文號
Dim m_CP43 As String
' 代理人
Dim m_CP44 As String
'彼所案號
Dim m_CP45 As String
Dim strCP64 As String
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
'Add By Cheng 2002/01/31
Public bln_From_Frm020102_01_BtnExt As Boolean
'Add By Cheng 2002/08/19
Dim m_strDL05 As String
'Add By Cheng 2003/08/05
Dim m_strNPCode As String '下一程序代號
Dim m_strNPName As String '下一程序名稱
'add by nick 2004/08/12
Dim m_CP84 As String       '發文規費
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
'add by nickc 2006/01/27
Dim m_CP110 As String
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
'Add By Cheng 2002/12/12
Dim m_CP14 As String '原承辦人
Dim m_EP06 As String '文件齊備日 Add By Sindy 2015/4/13
Dim m_990CP09 As String 'Add By Sindy 2016/12/20
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_strCF10 As String 'Add By Sindy 2020/8/12 取得主管機關
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

Private Sub cmdok_Click()
Dim ii As Integer
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      If Len(Trim(Me.txt(0).Text)) <= 0 And Me.txt(0).Visible = True Then
         MsgBox "請輸入延期承辦人!!!", vbExclamation
         Me.txt(0).SetFocus
         txt_GotFocus 0
         Exit Sub
      End If
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
            'GetCu103ByCustomer Me, m_TM78
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
            'GetCu103ByCustomer Me, m_TM79
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
            'GetCu103ByCustomer Me, m_TM80
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
            'GetCu103ByCustomer Me, m_TM81
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
      If CheckKeyIn(0) = 0 Then
         ' 設定滑鼠游標為等待狀態
         Screen.MousePointer = vbHourglass
         ' 更新欄位輸入的內容
         OnUpdateField
         ' 存檔
        'Modify By Cheng 2002/11/06
    '      'OnSaveData
          If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Add By Cheng 2002/11/08
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        'Add By Sindy 2021/2/25
        End If
        If textPrint = "N" Then
            If strLD18 <> "" Then
               Call PUB_TCaseAskIsPost(strLD18)
            End If
        '2021/2/25 END
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
                '2020/01/13 Modify by Amy +if 傳strCP118參數
                If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
                   PUB_ChkOfficialFee m_CP09, Me.textCP84.Text, IIf(textCP118 = "Y", "A", "")
                Else
                    PUB_ChkOfficialFee m_CP09, Me.textCP84.Text
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
         If m_CP10 = "303" Then
            If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
               Me.Tag = 0
            End If
         Else
            Me.Tag = 0
         End If
         'Add By Sindy 2018/5/3
         If Val(Me.Tag) = 0 Then
            If frm020102_01.bolIsEMPFlow = True Then
               Unload frm020102_01
               frm090202_4.m_ProState = "T" 'Add By Sindy 2021/1/29
               frm090202_4.Show
               Unload Me
               Exit Sub
            End If
         End If
         '2018/5/3 End
         
         frm020102_01.Show
         ' 90.12.07 modify by louis
   '      frm020102_01.Clear
         
         'Add By Cheng 2002/01/10
         frm020102_01.Clear1
         
         cmdOK.SetFocus
         Unload Me
      Else
         Me.txt(0).SetFocus
         txt_GotFocus 0
      End If
   End If
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
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/02/01
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM45.BackColor = &H8000000F
   
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textTM44.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   textCP44_2.BackColor = &H8000000F
   
   textDL02.BackColor = &H8000000F
   lbl(1).BackColor = Me.BackColor
   MoveFormToCenter Me
   'Add by nickc 2006/01/27
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Add by Amy 2021/12/27一開始將ListBox拉到需要的大小,字型會自動放大,所以畫面預設為一列高度，Form_Load才放大到需要的大小
    lstNameAgent.Height = 500
    lstNameAgent.Width = 1300
   
   'Add By Cheng 2002/01/31
   '按下確定按鈕後進入
   If Me.bln_From_Frm020102_01_BtnExt = False Then
      Me.lbl(0).Visible = False
      Me.lbl(1).Visible = False
      Me.txt(0).Visible = False
   '按下延期按鈕後進入
   Else
      Me.lbl(0).Visible = True
      Me.lbl(1).Visible = True
      Me.txt(0).Visible = True
   End If
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
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
      Else
         'Add By Sindy 2009/06/29
         ' 申請案號
         If IsNull(rsTmp.Fields("TM12")) = False Then
            textTM15 = rsTmp.Fields("TM12")
         End If
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         'textTM12 = rsTmp.Fields("TM12")
      End If
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
      ' 申請人
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'add by nickc 2007/02/01
      m_TM78 = Empty
      If IsNull(rsTmp.Fields("TM78")) = False Then
         m_TM78 = rsTmp.Fields("TM78")
         textTM78 = GetCustomerName(rsTmp.Fields("TM78"), 0)
      End If
      m_TM79 = Empty
      If IsNull(rsTmp.Fields("TM79")) = False Then
         m_TM79 = rsTmp.Fields("TM79")
         textTM79 = GetCustomerName(rsTmp.Fields("TM79"), 0)
      End If
      m_TM80 = Empty
      If IsNull(rsTmp.Fields("TM80")) = False Then
         m_TM80 = rsTmp.Fields("TM80")
         textTM80 = GetCustomerName(rsTmp.Fields("TM80"), 0)
      End If
      m_TM81 = Empty
      If IsNull(rsTmp.Fields("TM81")) = False Then
         m_TM81 = rsTmp.Fields("TM81")
         textTM81 = GetCustomerName(rsTmp.Fields("TM81"), 0)
      End If
      
      'add by nickc 2006/01/26
      m_TM24 = CheckStr(rsTmp.Fields("TM24"))
      'add by nickc 2006/11/17
      textPrint = CheckStr(rsTmp.Fields("tm77"))
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
         'textTM12 = rsTmp.Fields("SP11")
         'Add By Sindy 2009/06/29
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
Dim strCP44 As String
'Dim strCP45 As String
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
      
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         '案件性質
         textCP10 = textCP10 & PUB_GetRelateCasePropertyName(m_CP09, "1")
      End If
      '2010/12/27 End
      
      ' 本所期限
      m_CP06 = Empty
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      If m_CP10 <> "303" Then
         SetCPFieldOldData "CP06", textCP06, 1
      End If
      ' 法定期限
      m_CP07 = Empty
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      If m_CP10 <> "303" Then
         SetCPFieldOldData "CP07", textCP07, 1
      End If
      ' 業務區別
      If IsNull(rsTmp.Fields("CP12")) = False Then
         '91.6.11 MODIFY BY SONIA
         'textCP12 = GetStaffDepartment(rsTmp.Fields("CP12"))
         m_CP12 = rsTmp.Fields("CP12")
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
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      'Add By Cheng 2002/12/12
      m_CP14 = "" & rsTmp.Fields("CP14")
      
      'Add By Sindy 2011/7/12
      m_CP31 = Empty
      If IsNull(rsTmp.Fields("CP31")) = False Then
         m_CP31 = rsTmp.Fields("CP31")
      End If
      'Add By Sindy 2011/3/9
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
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
      'ADD BY SONIA 2014/11/6 電子送件案預設發文日為承辦人發文日CP85
      If textCP118 = "Y" Then
         textCP27 = TAIWANDATE("" & rsTmp.Fields("CP85"))
      End If
      'END  2014/11/6
      
      'Added by Lydia 2021/06/04 工作時數
       txtCP113 = "" & rsTmp.Fields("CP113")
       SetCPFieldOldData "CP113", txtCP113, 1
      'end 2021/06/04
      
      'Modify By Sindy 2009/05/04
      'CaculateNP08NP09
      Call CaculateNP08NP09(m_CP10)
      '2009/05/04 End
      
      ' 相關總收文號
      m_CP43 = Empty
      If IsNull(rsTmp.Fields("CP43")) = False Then
         m_CP43 = rsTmp.Fields("CP43")
      End If
      SetCPFieldOldData "CP43", m_CP43, 0
      ' 代理人
      textCP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         textCP44 = rsTmp.Fields("CP44")
      End If
      SetCPFieldOldData "CP44", textCP44, 0
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then
         textCP18 = rsTmp.Fields("CP18")
      End If
      ' 代理人
      m_CP44 = Empty
      If IsNull(rsTmp.Fields("CP44")) = False Then
         m_CP44 = rsTmp.Fields("CP44")
      End If
      ' 彼所案號
'      strCP45 = Empty
      m_CP45 = Empty
      If IsNull(rsTmp.Fields("CP45")) = False Then
'         strCP45 = rsTmp.Fields("CP45")
         m_CP45 = rsTmp.Fields("CP45")
         textTM45 = rsTmp.Fields("CP45")
      End If
      SetCPFieldOldData "CP45", textTM45, 0
'      SetCPFieldOldData "CP45", strCP45, 0
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      
    'add by nick 2004/08/12 發文規費
     If IsNull(rsTmp.Fields("CP17")) = False Then
         'Modify by Morgan 2010/5/27 判斷是否已有延期發文且有發文規費
         'm_CP84 = CheckStr(rsTmp.Fields("CP17"))
         If m_CP10 = "303" Then
            strExc(1) = m_CP43
         Else
            strExc(1) = m_CP09
         End If
         strExc(0) = "select * from caseprogress where cp43='" & strExc(1) & "' and cp10='303' and cp27>0 and cp84>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_CP84 = 0
         Else
            m_CP84 = CheckStr(rsTmp.Fields("CP17"))
            '2012/8/3 add by sonia 若有銷帳則要扣除銷帳規費
            If Val("" & rsTmp.Fields("CP77")) <> 0 Then
               If GetCP77Detail(strExc(1), m_Fee, m_Official) = True Then
                  m_CP84 = m_CP84 - m_Official
               End If
            End If
            '2012/8/3 end
         End If
         'End 2010/5/27
         textCP84.Text = m_CP84
     End If
     
     'Added by Morgan 2012/9/6 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/9/6
      
      'Added by Lydia 2022/05/23 法律所案源：取得案源類別、發文規費、email加註
      If m_CP10 = "303" Then
          strExc(0) = "select cp162 from caseprogress where cp09='" & m_CP43 & "' "
      Else
          strExc(0) = "select cp162 from caseprogress where cp09='" & m_CP09 & "' "
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
          m_LOS15 = "" & RsTemp.Fields("cp162")
          '限制B2類發文規費; 只需考慮B2類, A類不會有PT案,B1類PT案不會繳規費, C類已取消(有也是照原來規則不必改)
          'Memo by Lydia 2022/09/16 模組回傳m_LOS02
          m_LosCP84 = PUB_GetLosCP84(m_LOS15, m_TM01, m_TM02, m_TM03, m_TM04, "B2", m_LOS02, m_LosMemo)
      End If
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
      
      'add by nick 2004/10/08對照號數頁簽
      textCP36 = Empty
      If IsNull(rsTmp.Fields("CP36")) = False Then
         textCP36 = rsTmp.Fields("CP36")
      End If
      textCP40 = Empty
      If IsNull(rsTmp.Fields("CP40")) = False Then
         textCP40 = rsTmp.Fields("CP40")
      End If
      textCP41 = Empty
      If IsNull(rsTmp.Fields("CP41")) = False Then
         textCP41 = rsTmp.Fields("CP41")
      End If
      textCP42 = Empty
      If IsNull(rsTmp.Fields("CP42")) = False Then
         textCP42 = rsTmp.Fields("CP42")
      End If
      textCP37 = Empty
      textCP37_1 = Empty
      If IsNull(rsTmp.Fields("CP37")) = False Then
         textCP37 = rsTmp.Fields("CP37")
         textCP37_1 = textCP37
      End If
      textCP38 = Empty
      If IsNull(rsTmp.Fields("CP38")) = False Then
         textCP38 = rsTmp.Fields("CP38")
      End If
      textCP39 = Empty
      If IsNull(rsTmp.Fields("CP39")) = False Then
         textCP39 = rsTmp.Fields("CP39")
      End If
      textCP80 = Empty
      If IsNull(rsTmp.Fields("CP80")) = False Then
         textCP80 = rsTmp.Fields("CP80")
      End If
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
   
   ' 未收文期限
   InitialGrdList
   ' 案件性質為延期時才讀取下一程序檔的資料
   If m_CP10 = "303" Then
      ' 取得下一程序檔案中的資料列表在 Grid List 中
      'Modify by Morgan 2009/12/29 下一程序要排除程序管制的案件性質
      '2010/3/23 MODIFY BY SONIA 剔除下一程序非智權人員掌控之案件性質改以strNpSqlOfNoSalesDuty控制
      strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND NP06 IS NULL " & strNpSqlOfNoSalesDuty
   
      'Add by Morgan 2009/12/29 延期+AB類未發文未取消收文的程序
      strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
         " FROM CASEPROGRESS WHERE CP01 = '" & m_TM01 & "' AND CP02 = '" & m_TM02 & "'" & _
         " AND CP03 = '" & m_TM03 & "' AND CP04 = '" & m_TM04 & "'" & _
         " AND CP09<'C' and cp10<>'303' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
            'Remove by Morgan 2009/12/29 改語法加條件控制
            '' 是否續辦欄位必須為空白
            'If IsNull(rsTmp.Fields("NP06")) = False Then
            '   If IsEmptyText(rsTmp.Fields("NP06")) = False Then
            '      GoTo NextRecord
            '   End If
            'End If
         
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
                  grdList.TextMatrix(grdList.row, 2) = TAIWANDATE(rsTmp.Fields("NP08"))
               End If
            End If
            ' 法定期限
            If IsNull(rsTmp.Fields("NP09")) = False Then
               If IsEmptyText(rsTmp.Fields("NP09")) = False Then
                  grdList.TextMatrix(grdList.row, 3) = TAIWANDATE(rsTmp.Fields("NP09"))
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
   End If
   
   ' 上次延期日(取最後一筆)
   strSql = "SELECT * FROM DateLimit " & _
            "WHERE DL01 = '" & m_CP09 & "' AND " & _
                  "DL02 IN (SELECT MAX(DL02) FROM DateLimit " & _
                           "WHERE DL01 = '" & m_CP09 & "') "
            
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("DL02")) = False Then
         If rsTmp.Fields("DL02") <> "0" Then
            textDL02 = rsTmp.Fields("DL02")
         End If
      End If
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
    'Add By Cheng 2003/11/10
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        Me.Label13.Visible = False
        Me.Label17.Visible = False
        Me.Label18.Visible = False
        Me.textCP37.Visible = False
        Me.textCP37.Enabled = False
        Me.textCP38.Visible = False
        Me.textCP38.Enabled = False
        Me.textCP39.Visible = False
        Me.textCP39.Enabled = False
    Case Else
        Me.Label15.Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
    End Select
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
     
   Select Case m_TM01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "T", "TF", "FCT":
         QueryTradeMark
      Case Else:
         QueryServicePractice
   End Select
   
   'Add By Sindy 2021/1/15 T發文所有程式,台灣案鎖住畫面上之CP44,不可輸入
   If m_TM10 = "000" Then
      textCP44.Enabled = False
   End If
   '2021/1/15 END
   
   'Modify By Sindy 2012/7/26
   '台灣案才需顯示出名代理人
   lstNameAgent.Clear
   If m_TM10 = "000" Then
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
   Set rsTmp = Nothing
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/17 若已經從基本檔抓出來，就不重抓
   If Trim(textPrint) = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   'Add By Sindy 2025/8/11 檢查卷宗區是否已有承辦放入之CUS,若有,系統不產出定稿
   If PUB_CPPChkFileExists(m_CP09, "cus") = True Then
      textPrint = "N"
   End If
   '2025/8/11 END
   
   'Add By Sindy 2011/10/28 T內商000台灣案所有案件性質加電子送件功能
   'Modified by Lydia 2017/10/24 外商爭議案403行政訴訟-延期由內商人員新增
   'If m_TM01 = "T" And m_TM10 = "000" Then
   'Modify by Amy 2020/01/23 +是否電子送件
   lblPayToday.Visible = False
   txtPayToday.Visible = False
   If m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
        lblPayToday.Visible = True
        txtPayToday.Visible = True
      End If
   'end 2020/01/13
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2011/10/28 End
   
   Call PUB_TCaseEFeeRemind(m_CP09) 'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2022/05/23
   
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    bln_From_Frm020102_01_BtnExt = False
   'Add By Cheng 2002/07/18
   Set frm020102_11 = Nothing
End Sub

Private Sub grdList_Click()
Dim i As Integer
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            'Add By Sindy 2009/05/04
            For i = 1 To grdList.Rows - 1
               grdList.TextMatrix(i, 0) = Empty
            Next i
            '2009/05/04 End
            grdList.TextMatrix(grdList.row, 0) = "V"
         End If
         
         'Add by Morgan 2009/12/29 加期限檢查
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            strExc(2) = Replace(grdList.TextMatrix(grdList.row, 3), "/", "")
            If DBDATE(strExc(2)) <> m_CP07 Then
               MsgBox "所點選案件性質的法定期限與延期程序不同，不可點選！"
               grdList.TextMatrix(grdList.row, 0) = ""
            Else
         'end 2009/12/29
         
               'Add By Sindy 2009/05/04
               Call CaculateNP08NP09(grdList.TextMatrix(grdList.row, 8))
            End If
         End If

      End If
   End If
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
      Me.SSTab1.Tab = 0
      lstNameAgent.SetFocus
   End If
   '2015/7/22 END
End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If Me.SSTab1.Tab = 0 Then
        Me.textCP27.SetFocus
        textCP27_GotFocus
    Else
        Me.textCP36.SetFocus
        textCP36_GotFocus
    End If
End Sub
'Add by Morgan 2009/12/29
Private Sub Text9_GotFocus()
   TextInverse Text9
   CloseIme
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text9_Validate(Cancel As Boolean)
   If Val(Text9) > 0 Then
      strExc(1) = m_TM01
      strExc(2) = m_TM10
      strExc(3) = CompDate("1", Val(Text9), m_CP07)
      textCP07 = TransDate(strExc(3), 1)
      'Modify By Sindy 2014/10/6 台灣案之本所期限設定
      If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
         textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
      Else
      '2014/10/6 END
         GetCtrlDT strExc '由法定期限計算本所期限
         textCP06 = TransDate(PUB_GetWorkDay1(strExc(0), True), 1)
      End If
   End If
End Sub
'end 2009/12/29

' 延期後本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的延期後本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCP06
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
   End If
End Sub

' 延期後法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的延期後法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCP07
      End If
   End If
End Sub
'
'Private Sub textCP09S_GotFocus()
'    TextInverse Me.textCP09S
'End Sub
'
'Private Sub textCP09S_Validate(Cancel As Boolean)
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSQL As String
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP09S) = False Then
'      If textCP09S = m_CP09 Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "商業司查詢總收文號不可為本案之收文號"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'
'      strSQL = "SELECT * FROM CaseProgress " & _
'               "WHERE CP01 = 'TR' AND " & _
'                     "CP09 = '" & textCP09S & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount <= 0 Then
'         rsTmp.Close
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "商業司查詢總收文號資料不存在"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP09S_GotFocus
'         GoTo EXITSUB
'      End If
'      rsTmp.Close
'   End If
'EXITSUB:
'   Set rsTmp = Nothing
'End Sub

'Add By Sindy 2011/10/28
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2011/10/28
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
'' 是否出名
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

Private Sub textCP27_LostFocus()
'Add By Cheng 2002/07/16
'Modify By Sindy 2009/05/04
'CaculateNP08NP09
'Add By Sindy 2020/3/3 桂英說當期限調整完,到第2頁籤再回來時,會再重算延期的期限,所以要擋住
If Val(textCP06) = 0 Or Val(textCP07) = 0 Then
'2020/3/3 END
   Call CaculateNP08NP09(m_CP10)
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
      'Modify By Cheng 2002/07/16
      '將計算期限的動作移至textCP27_LostFocus
'      CaculateNP08NP09
   End If
EXITSUB:
End Sub

' 計算本所期限及法定期限
Private Sub CaculateNP08NP09(strCF03 As String)
   If IsEmptyText(textCP27) = False Then
      strExc(0) = TransDate(textCP27.Text, 2)
      'edit by nickc 2007/02/06 不用 dll 了
      'If objLawDll.GetCaseFeeDelay(m_TM01, m_TM10, m_CP10, strExc) Then
      'Modify By Sindy 2009/05/04
      'If ClsLawGetCaseFeeDelay(m_TM01, m_TM10, m_CP10, strExc) Then
      If ClsLawGetCaseFeeDelay(m_TM01, m_TM10, strCF03, strExc) Then
         textCP07 = TransDate(strExc(1), 1)
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(strExc(2), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

Private Sub textCP36_GotFocus()
    TextInverse Me.textCP36
End Sub

Private Sub textCP36_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   'Modify by Amy 2022/09/29 原20 放寬至200
   If CheckLengthIsOK(textCP36, textCP36.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造號數內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP36_GotFocus
   End If
   If GetTextLength(textCP36) > 20 And InStr(textCP36, ",") = 0 And InStr(textCP36, ";") = 0 Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "對造號數內容太長,無法寫入基本檔欄位中,請洽電腦中心"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        textCP36_GotFocus
   End If
End Sub

Private Sub textCP37_1_GotFocus()
    TextInverse Me.textCP37_1
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
    Cancel = False
    If CheckLengthIsOK(textCP37_1, 140) = False Then
        Cancel = True
        strTit = "檢核資料"
        strMsg = "對造案件名稱內容太長"
        textCP37_1_GotFocus
    End If
    'edit by nickc 2007/06/06 切換輸入法改用API
    If Cancel = False Then CloseIme
End Sub

Private Sub textCP37_GotFocus()
    TextInverse Me.textCP37
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP37, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP37_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP38_GotFocus()
    TextInverse Me.textCP38
End Sub

Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP38, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP38_GotFocus
   End If
End Sub

Private Sub textCP39_GotFocus()
    TextInverse Me.textCP39
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP39, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP39_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP40_GotFocus()
    TextInverse Me.textCP40
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造中文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP41_GotFocus()
    TextInverse Me.textCP41
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造英文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

Private Sub textCP42_GotFocus()
    TextInverse Me.textCP42
    'edit by nickc 2007/06/06 切換輸入法改用API
    OpenIme
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造日文名稱內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCP44_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCP44.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

Private Sub textCP44_KeyPress(KeyAscii As Integer)
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

Private Sub textCP80_GotFocus()
    TextInverse Me.textCP80
End Sub

Private Sub textCP80_Validate(Cancel As Boolean)
'add by nickc 2005/06/03
textCP80 = Replace(textCP80, " ", "")
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

   Dim nIndex As Integer
   ' 更新案件進度檔的欄位
   If m_CP10 = "303" Then
      ' 是否出名
      SetCPFieldNewData "CP22", textCP22
      ' 更新發文日為延期日
      SetCPFieldNewData "CP27", DBDATE(textCP27)
      ' 代理人
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
      ' 91.09.02 modify by louis
      ' 進度備註
      'SetCPFieldNewData "CP64", textCP64
      strCP64 = textCP64
'edit by nickc 2006/01/27
'      If IsEmptyText(textAgName) = False Then
'         strCP64 = strCP64 & "," & "本所出名代理人:" & textAgName
'      End If
      SetCPFieldNewData "CP64", strCP64
      'add by nickc 2006/01/27
      SetCPFieldNewData "CP110", m_CP110
      
      '910722 Sieg
      If m_CP43 = "" Then
         For nIndex = 1 To grdList.Rows - 1
            grdList.row = nIndex
            ' 判斷該列是否有被選取
            If grdList.TextMatrix(nIndex, 0) = "V" Then
               SetCPFieldNewData "CP43", grdList.TextMatrix(grdList.row, 7)
               Exit For
            End If
         Next nIndex
      End If
   Else
      ' 本所期限
      SetCPFieldNewData "CP06", DBDATE(textCP06)
      ' 法定期限
      SetCPFieldNewData "CP07", DBDATE(textCP07)
   End If
   'Add By Cheng 2002/11/19
   '承辦期限
''''edit by nickc 2007/10/12 改抓有時效性的
''''    If textCP27.Text <> "" Then SetCPFieldNewData "CP48", DBDATE(DateAdd("d", GetCF04(m_TM01, m_TM10, m_CP10), ChangeTStringToWDateString(textCP27.Text)))
   If textCP27.Text <> "" Then SetCPFieldNewData "CP48", Pub_GetHandleDay(m_TM01, m_TM10, m_CP10, DBDATE(textCP27), DBDATE(textCP06), textCP09)
   'Add By Sindy 2011/3/9
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
   
   'Added by Lydia 2021/06/04 工作時數
   SetCPFieldNewData "CP113", txtCP113
   
End Sub

' 更新案件進度檔
'Private Sub OnUpdateCaseProgress()
Private Function OnUpdateCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
'Add By Cheng 2002/11/06
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
   If bDifference = True Then cnnConnection.Execute strSql Else Exit Function
   '910918 nick 更新彼所案號
   strSql = "update caseprogress set cp45=(select min(cp45) from caseprogress where cp01='" & m_TM01 & "' and cp02='" & m_TM02 & "' and cp03='" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "' and cp04='" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04) & "' and cp44='" & textCP44.Text & "') where cp09='" & m_CP09 & "' "
   cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateCaseProgress = False
End Function

'Modify By Cheng 2002/11/06
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strCP05 As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP01 As String
   Dim strNP08 As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strDL01 As String
   Dim strDL03 As String
   Dim strDL04 As String
   Dim objCopyCP As ClsCopyCP
   Dim bolSysDt As Boolean 'Add By Sindy 2010/12/28
   Dim strCP30 As String 'Add by Morgan 2011/4/22
   Dim str303CP10 As String, str303CP09 As String 'Add By Sindy 2012/3/23
   
'Add By Cheng 2002/11/06
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
   
   ' 新增資料到延期記錄檔
   strDL01 = Empty
   strDL03 = Empty
   strDL04 = Empty
   '按確定按鈕進入
   If m_CP10 = "303" Then
      ' 案件性質為延期時, 總收文號, 本所期限及法定期限為未收文期限所選取的收文資料
      For nIndex = 1 To grdList.Rows - 1
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strDL01 = grdList.TextMatrix(nIndex, 7)
            strDL03 = DBDATE(grdList.TextMatrix(grdList.row, 2))
            strDL04 = DBDATE(grdList.TextMatrix(grdList.row, 3))
            'Add By Cheng 2002/08/19
            strNP22 = grdList.TextMatrix(grdList.row, 9)
            
            ' 若原延期記錄檔已存在時則先刪除
            strSql = "DELETE FROM DateLimit WHERE DL01 = '" & strDL01 & "' "
            cnnConnection.Execute strSql
            ' 新增一筆資料到延期記錄檔
            'Modify By Cheng 2002/08/19
'            'Modify By Cheng 2002/06/20
'            strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'                     "VALUES ('" & strDL01 & "'," & _
'                              DBDATE(textCP27) & "," & _
'                              strDL03 & "," & strDL04 & ")"
'            strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
'                     "VALUES ('" & strDL01 & "'," & _
'                              DBDATE(textCP27) & "," & _
'                              strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
            m_strDL05 = IIf(m_CP10 = "303", "2", "1")
            strSql = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05, DL06) " & _
                     "VALUES ('" & strDL01 & "'," & _
                              DBDATE(textCP27) & "," & _
                              strDL03 & "," & strDL04 & ",'" & m_strDL05 & "','" & IIf(m_strDL05 = "1", "", strNP22) & "' )"
            cnnConnection.Execute strSql
         End If
      Next nIndex
   '按延期按鈕進入
   Else
      ' 案件性質不為延期時, 總收文號, 本所期限及法定期限為該案本身
      strDL01 = m_CP09
      strDL03 = m_CP06
      strDL04 = m_CP07
      'Add By Cheng 2002/08/19
      strNP22 = ""
      
      ' 若原延期記錄檔已存在時則先刪除
      strSql = "DELETE FROM DateLimit WHERE DL01 = '" & strDL01 & "' "
      cnnConnection.Execute strSql
      ' 新增一筆資料到延期記錄檔
      'Modify By Cheng 2002/08/19
'      'Modify By Cheng 2002/06/20
'      strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'               "VALUES ('" & strDL01 & "'," & _
'                        DBDATE(textCP27) & "," & _
'                        strDL03 & "," & strDL04 & ")"
'      strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
'               "VALUES ('" & strDL01 & "'," & _
'                        DBDATE(textCP27) & "," & _
'                        strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
      m_strDL05 = IIf(m_CP10 = "303", "2", "1")
      strSql = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05, DL06) " & _
               "VALUES ('" & strDL01 & "'," & _
                        DBDATE(textCP27) & "," & _
                        strDL03 & "," & strDL04 & ",'" & m_strDL05 & "','" & IIf(m_strDL05 = "1", "", strNP22) & "' )"
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '911011 nick 因為下面 function 已經有程式更新了，所以不用作
   '911015 nick  確認後還是要在做
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
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
    
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質非延期時時新增一筆資料到案件進度檔中
   If m_CP10 <> "303" Then
      ' 收文號
      strCP09 = Empty
      strCP09 = AutoNo("B", 6)
            
      ' 收文日
      strCP05 = DBDATE(SystemDate())
      ' 案件性質
      strCP10 = "303"
      ' 業務區別 MODIFY BY SONIA 91.8.24
      'strCP12 = GetST15(m_CP13)
      ' 發文日
      strCP27 = DBDATE(textCP27)
      'Modify By Cheng 2002/02/01
      '若從延期按鈕進來的, 以延期承辦人寫入承辦人欄位
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP22,CP27,CP43,CP44,CP45) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
'                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & strCP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                       "'" & textCP22 & "'," & strCP27 & ",'" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "') "
      '910918 nick 要更新是否向客戶收款  是否算案件數  是否開電腦收據
      '***** start
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP22,CP27,CP43,CP44,CP45) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & IIf(Me.bln_From_Frm020102_01_BtnExt, Me.txt(0).Text, strUserNum) & "'," & _
                       "'" & textCP22 & "'," & strCP27 & ",'" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "') "
      '911015 nick 更新 cp64
      '***** start
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP22,CP27,CP43,CP44,CP45,cp20,cp26,cp32) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & IIf(Me.bln_From_Frm020102_01_BtnExt, Me.txt(0).Text, strUserNum) & "'," & _
                       "'" & textCP22 & "'," & strCP27 & ",'" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','N','N','N') "
'edit by nickc 2006/11/30 補出名代理人
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP22,CP27,CP43,CP44,CP45,cp20,cp26,cp32,cp64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & IIf(Me.bln_From_Frm020102_01_BtnExt, Me.txt(0).Text, strUserNum) & "'," & _
                       "'" & textCP22 & "'," & strCP27 & ",'" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','N','N','N','" & strCP64 & "') "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP22,CP27,CP43,CP44,CP45,cp20,cp26,cp32,cp64,cp110) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & IIf(Me.bln_From_Frm020102_01_BtnExt, Me.txt(0).Text, strUserNum) & "'," & _
                       "'" & textCP22 & "'," & strCP27 & ",'" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','N','N','N','" & strCP64 & "','" & m_CP110 & "') "
      '***** end
      '***** end
      cnnConnection.Execute strSql
        'Add By Cheng 2003/10/29
        '更新此新增B類進度的對造資料
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            strSql = "Update CaseProgress Set CP36='" & Me.textCP36.Text & "', CP37='" & ChgSQL(Me.textCP37_1.Text) & "', CP40='" & ChgSQL(Me.textCP40.Text) & "', CP41='" & ChgSQL(Me.textCP41.Text) & "', CP42='" & ChgSQL(Me.textCP42.Text) & "', CP80='" & Me.textCP80.Text & "' " & _
                        " Where CP09='" & strCP09 & "' "
        Case Else
            strSql = "Update CaseProgress Set CP36='" & Me.textCP36.Text & "', CP37='" & ChgSQL(Me.textCP37.Text) & "', CP38='" & ChgSQL(Me.textCP38.Text) & "', CP39='" & ChgSQL(Me.textCP39.Text) & "', CP40='" & ChgSQL(Me.textCP40.Text) & "', CP41='" & ChgSQL(Me.textCP41.Text) & "', CP42='" & ChgSQL(Me.textCP42.Text) & "', CP80='" & Me.textCP80.Text & "' " & _
                        " Where CP09='" & strCP09 & "' "
        End Select
        cnnConnection.Execute strSql
   End If
    'Add By Cheng 2003/10/29
    '更新點選進入的進度對造資料
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            strSql = "Update CaseProgress Set CP36='" & Me.textCP36.Text & "', CP37='" & ChgSQL(Me.textCP37_1.Text) & "', CP40='" & ChgSQL(Me.textCP40.Text) & "', CP41='" & ChgSQL(Me.textCP41.Text) & "', CP42='" & ChgSQL(Me.textCP42.Text) & "', CP80='" & Me.textCP80.Text & "' " & _
                        " Where CP09='" & m_CP09 & "' "
        Case Else
            strSql = "Update CaseProgress Set CP36='" & Me.textCP36.Text & "', CP37='" & ChgSQL(Me.textCP37.Text) & "', CP38='" & ChgSQL(Me.textCP38.Text) & "', CP39='" & ChgSQL(Me.textCP39.Text) & "', CP40='" & ChgSQL(Me.textCP40.Text) & "', CP41='" & ChgSQL(Me.textCP41.Text) & "', CP42='" & ChgSQL(Me.textCP42.Text) & "', CP80='" & Me.textCP80.Text & "' " & _
                        " Where CP09='" & m_CP09 & "' "
        End Select
    cnnConnection.Execute strSql
    'Add By Cheng 2004/02/24
    '若點選的案件性質為異議, 評定, 廢止, 則對造資料更新至基本檔
    If m_CP10 = "601" Or m_CP10 = "603" Or m_CP10 = "605" Then
        'Modify by Amy 2022/09/29 +GetCP36,避免cp36欄位放寬,導致寫入其他欄位錯誤
        strSql = "Update Trademark Set TM05='" & ChgSQL(Me.textCP37_1.Text) & "', TM15='" & GetCP36(Me.textCP36.Text) & "', TM09='" & Me.textCP80.Text & "' " & _
                    " Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        cnnConnection.Execute strSql
    End If
    'End
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的未收文期限資料
   'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
   Dim TheFirstV As String
   TheFirstV = ""
   If m_CP10 = "303" Then
      For nIndex = 1 To grdList.Rows - 1
         grdList.row = nIndex
         ' 判斷該列是否有被選取
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strNP01 = grdList.TextMatrix(grdList.row, 7)
            'Add by Morgan 2011/4/22
            If TheFirstV = "" Then
               If grdList.TextMatrix(grdList.row, 9) = "0" Then
                  strCP30 = ""
               Else
                  strCP30 = grdList.TextMatrix(grdList.row, 9)
               End If
            End If
            
            'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
            If TheFirstV = "" Then TheFirstV = strNP01
            strNP07 = grdList.TextMatrix(grdList.row, 8)
            strNP22 = grdList.TextMatrix(grdList.row, 9)
            'Modify by Morgan 2009/12/29 +更新CP
            If Val(strNP22) > 0 Then
               strSql = "UPDATE NextProgress SET NP08 = " & ChangeTStringToWString(textCP06) & "," & _
                                             "NP09 = " & ChangeTStringToWString(textCP07) & " " & _
                     "WHERE NP01 = '" & strNP01 & "' AND " & _
                           "NP07 = " & strNP07 & " AND " & _
                           "NP22 = " & strNP22 & " "
            Else
               strSql = "UPDATE CaseProgress SET CP06 = " & ChangeTStringToWString(textCP06) & "," & _
                                             "CP07 = " & ChangeTStringToWString(textCP07) & " " & _
                     "WHERE CP09 = '" & strNP01 & "'"
            End If
            
            cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
            cnnConnection.Execute strSql, intI
            cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2010/7/13 +控制來函期限通知的 Trigger 不被觸發
         End If
      Next nIndex
      'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
      If TheFirstV <> "" Then
            'Modify by Morgan 2011/4/22 +CP30
            strSql = " update caseprogress set cp43='" & TheFirstV & "',cp30='" & strCP30 & "' Where CP09='" & m_CP09 & "' "
            cnnConnection.Execute strSql
      End If
   End If

   'Add By Sindy 2012/5/8
   '台灣商標Ｔ,FCT案若收文爭議案件性質時
   '其點選之案件進度已有齊備日時，於延期發文時重新計算承辦期限更新進度檔
   If (m_TM01 = "T" Or m_TM01 = "FCT") And m_TM10 = "000" Then
      'Add By Sindy 2013/7/3 補加程式段
      'A類收延期
      If m_CP10 = "303" Then
         'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
         If InStr(TMdebate, strNP07) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, strNP07) > 0) And TheFirstV <> "" And _
            (Left(Trim(TheFirstV), 1) = "A" Or Left(Trim(TheFirstV), 1) = "B") Then
            strSql = "SELECT ep06,ep34 FROM engineerprogress WHERE ep02='" & TheFirstV & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_EP06 = "" & RsTemp.Fields("ep06") '文件齊備日 Add By Sindy 2015/4/13
               If Not IsNull(RsTemp.Fields(0)) Then
                  If RsTemp.Fields(0) > 0 Then
                     '以延期發文次日(遇假日順延) 起7個工作天計算。此時不再考慮急件
                     strDate = CompWorkDay(8, DBDATE(textCP27), 0)
                     strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                              "WHERE CP09 = '" & TheFirstV & "' "
                     cnnConnection.Execute strSql
                     '並新增商標案管制日期異動記錄檔，異動類別=3，輸入日期=系統日，異動動作='延期發文'
                     strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                        " values('" & TheFirstV & "','3','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'延期發文')"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
         End If
      '2013/7/3 END
      '點延期按鈕進來時
      Else
         'Modify By Sindy 2025/7/30 FCT案727分析不屬於爭議案 + FCT_NotTMdebate
         If InStr(TMdebate, m_CP10) > 0 And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, m_CP10) > 0) Then
            strSql = "SELECT ep06,ep34 FROM engineerprogress WHERE ep02='" & m_CP09 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               m_EP06 = "" & RsTemp.Fields("ep06") '文件齊備日 Add By Sindy 2015/4/13
               If Not IsNull(RsTemp.Fields(0)) Then
                  If RsTemp.Fields(0) > 0 Then
                     '以延期發文次日(遇假日順延) 起7個工作天計算。此時不再考慮急件
                     strDate = CompWorkDay(8, DBDATE(textCP27), 0)
                     strSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                              "WHERE CP09 = '" & m_CP09 & "' "
                     cnnConnection.Execute strSql
                     '並新增商標案管制日期異動記錄檔，異動類別=3，輸入日期=系統日，異動動作='延期發文'
                     strSql = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                        " values('" & m_CP09 & "','3','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'延期發文')"
                     cnnConnection.Execute strSql
                  End If
               End If
            End If
         End If
      End If
   End If
   '2012/5/8 End

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '若有定義代理人收達天數時, 新增一筆收達的記錄到下一程序檔
   'Modify By Sindy 2012/3/23 收達期限的m_CP10及m_CP09要依延期的狀況帶入不同資料
   '按確定按鈕進入
   If m_CP10 = "303" Then
      str303CP10 = strNP07 '點未收文期限的那一筆案件性質
      str303CP09 = m_CP09
   '按延期按鈕進入
   Else
      str303CP10 = m_CP10
      str303CP09 = strCP09 'B類收文
   End If
   '2012/3/23 End
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & str303CP10 & "' "
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
                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                            strNP08 & "," & strNP08 & ",'" & strUserNum & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
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
         
         ' 延展, 使用宣誓, 刊登廣告, 繳年費, 收達不印接洽結案單
'         '92.6.8 SONIA 加 言詞辯論, 準備程序
         Select Case strNP07
'            Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
            Case "102", "105", "702", "708", "305", "998", "997"
            Case Else:
               ' 列印國內案件接洽及結案記錄單
               g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
         End Select
      End If
      'Modify By Sindy 2013/3/25 內外商的延期發文取消掛催審期限
'      'Add By Sindy 2012/9/10
'      If IsNull(rsTmp.Fields("CF05")) = False Then
'         strNP07 = "305"
'         'strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
'         strNP08 = GetUrgeDate(m_TM01, m_TM10, str303CP10, textCP27)
'         strNP22 = GetNextProgressNo()
'         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
'         cnnConnection.Execute strSql
'      End If
'      '2012/9/10 End
   End If
   rsTmp.Close
   
   'add by nick 2004/08/12 更新實際發文規費
   'edit by nick 2004/08/31
   If m_CP10 = "303" Then
      'Modified by Lydia 2022/05/23 更新法律所案源單號
      'strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      'modify by sonia 2023/11/20 同時更新為要請款CP20=NULL
      strSql = "Update CaseProgress Set CP20=NULL,CP84=" & Trim(Val(textCP84.Text)) & IIf(m_LOS15 <> "", ", CP162=" & CNULL(m_LOS15), "") & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   Else
      'Modified by Lydia 2022/05/23 更新法律所案源單號
      'strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & strCP09 & "' "
      'modify by sonia 2023/11/20 同時更新為要請款CP20=NULL
      strSql = "Update CaseProgress Set CP20=NULL,CP84=" & Trim(Val(textCP84.Text)) & IIf(m_LOS15 <> "", ", CP162=" & CNULL(m_LOS15), "") & " Where CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add by Sindy 2012/10/4 外->台,智權人員是葉雪貞及巨京,發文規費和收文規費不相同時,系統自動更改進度檔內規費費用及計算點數
   'Modified by Lydia 2015/10/16 + m_CP84
   Call PUB_TSendUpdateCP1718(str303CP09, textCP84, textPrint, m_TM10, m_CP13, m_CP84)
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若該筆記錄是母案時, 同時對所有的子案做新增案件進度檔的工作
   If m_TM01 = "TF" And m_TM03 = "0" And m_TM04 = "00" Then
      Set objCopyCP = New ClsCopyCP
    'Modify By Cheng 2002/11/06
'      objCopyCP.CopyCaseProgress m_CP09
      If objCopyCP.CopyCaseProgress(m_CP09) = False Then GoTo ErrorHandler
      Set objCopyCP = Nothing
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then
'      PrintLetter
'   End If
   
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
            'strSQL = "Update customer Set CU103='" & ChgSQL(m_CU103) & "',cu05='" & ChgSQL(m_CU05) & "',cu88='" & ChgSQL(m_CU88) & "',cu89='" & ChgSQL(m_CU89) & "',cu90='" & ChgSQL(m_CU90) & "',cu112='" & ChgSQL(m_CU112) & "'   Where Cu01 = '" & Mid(ChangeCustomerL(m_TM23), 1, 8) & "' and cu02='" & Mid(ChangeCustomerL(m_TM23), 9, 1) & "' "
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
   
   'add by nickc 2006/01/26
   'edit by nickc 2007/08/10
   'If m_CU112 <> "" Then
   If SeekCu112(1) <> "" Then
        'EDIT BY NICKc 2007/08/10
        'strSQL = "Update Trademark Set TM24='" & ChgSQL(Pub_RplCu112(m_TM24, m_CU112)) & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        'Modify By Sindy 2011/2/22
        'strSql = "Update Trademark Set TM24='" & ChgSQL(Pub_RplCu112(m_TM24, SeekCu112(1))) & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        strSql = "Update Trademark Set TM24='" & ChgSQL(Pub_RplCu112(m_TM24, SeekCu112(1), m_TM23)) & "' Where " & ChgTradeMark(m_TM01 & m_TM02 & m_TM03 & m_TM04)
        cnnConnection.Execute strSql
    End If
    
   Set rsTmp = Nothing
   
   'Add by Sindy 98/3/24
   ' 設定滑鼠游標為預設
   Screen.MousePointer = vbDefault
   
   'Modify By Sindy 2011/3/9 若為電子送件則不經發文室
   'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
   If (textCP118.Visible = True And textCP118 <> "") Then
      'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
      'Modified by Lydia 2017/10/24 抓現在收文號
      'If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
      'Debug by Lydia 2017/10/26 第一個收文號傳現在要發,第2個傳本案今日發文的收文號
      'If ModifyDispatchCp130(textCP09, str303CP09, m_CP123s, m_CP130s, textCP27, , True) = False Then
      If ModifyDispatchCp130(IIf(strCP09 <> "", strCP09, textCP09), m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
         'Exit Function
         GoTo ErrorHandler
      End If
      'end 2016/5/16
      
      'Add by Amy 2020/02/05 +輸入收文文號
      If strSrvDate(1) >= T商標電子送件扣款啟用日 Then
      
         'Add By Sindy 2020/11/23 取得主管機關, 只有主管機關是"經濟部智慧財產局"時, 才做自動扣款
         'ex:T225478:408 行政訴訟上訴 403 行政訴訟
         If m_CP10 = "303" Then
            strExc(10) = ""
            If TheFirstV <> "" Then
               strSql = "SELECT cp10 FROM caseprogress WHERE cp09='" & TheFirstV & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strExc(10) = RsTemp.Fields("cp10")
               End If
            End If
            If strExc(10) = "" And m_CP43 <> "" Then
               strSql = "SELECT cp10 FROM caseprogress WHERE cp09='" & m_CP43 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strExc(10) = RsTemp.Fields("cp10")
               End If
            End If
            Call GetCaseFeeByNick(m_TM01, m_TM10, strExc(10), m_strCF10)
            If m_strCF10 <> "經濟部智慧財產局" And m_strCF10 <> "" And txtPayToday.Visible = True Then
               txtPayToday = ""
               m_CP130s = m_strCF10
            End If
         End If
         '2020/11/23 END
      
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
                     'Exit Function
                     GoTo ErrorHandler
                  End If
                  strExc(0) = InputBox("請輸入智慧局收文文號!!")
                  If strExc(0) = "" Then
                     'Exit Function
                     GoTo ErrorHandler
                  Else
                     textCP64 = "智慧局收文文號:" & strExc(0) & ";" & textCP64
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
      
      'Added by Lydia 2017/10/24 更新B類收文的電子送件
      If m_CP10 <> "303" Then
        strSql = "Update CaseProgress Set CP118='Y' Where CP09 = '" & strCP09 & "' "
        cnnConnection.Execute strSql
        PUB_UpdateDispatch strCP09, m_CP123s, m_CP130s 'Added by Lydia 2017/10/26
      End If
      'end 2017/10/24
   Else
      If m_TM10 = "000" Then
         'Add by Sindy 2009/4/24
         'Modified by Lydia 2017/10/24 抓現在收文號
         'If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
         'Debug by Lydia 2017/10/26 第一個收文號傳現在要發,第2個傳本案今日發文的收文號
         'If ModifyDispatchCp130(textCP09, str303CP09, m_CP123s, m_CP130s, textCP27) = False Then
         If ModifyDispatchCp130(IIf(strCP09 <> "", strCP09, textCP09), m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
            'Exit Function
            GoTo ErrorHandler
   '      Else
   '         m_CP123s = GetCPMSendYn(m_TM01, m_CP10, 1)
         End If
         'Modify By Sindy 2009/04/24
         
         'Add by Sindy 98/07/01 B類延期要補收文號
         If m_CP10 <> "303" And m_TM10 = "000" Then
            m_CP09s = strCP09
         End If
         
         'PUB_UpdateDispatch m_CP09s, m_CP123s
         PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
      End If
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
                  strSql = strSql & ",CP152 = " & CompWorkDay(2, DBDATE(textCP27)) & ",CP64='" & textCP64 & "'"
              Else
                  strSql = strSql & ",CP152 =" & CompWorkDay(3, DBDATE(textCP27)) & ",CP64='" & textCP64 & "'"
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
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = str303CP09
      PUB_AddLetterProgress strLD18, 0, IIf(textPrint = "N", False, True), "", False, m_TM23, "303", m_TM44
   End If
   '2019/12/20 END
   Call PUB_UpdateLP19_T(m_TM01, m_TM02, m_TM03, m_TM04, str303CP09, textCP27) 'Add by Sindy 2020/2/12 收據/回執設定
   
   'Add By Sindy 2016/12/20
   If m_990CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_990CP09 & "' and cp27 is null"
      cnnConnection.Execute strSql
   End If
   '2016/12/20 END
    
   'Added by Lydia 2023/03/17 若延期的相關總收文號為B2案源時，同時新增法務案之內部收文39延期
   If m_LOS15 <> "" And m_LOS02 = "B2" Then
       Call PUB_InsertLosBCP(m_LOS15, DBDATE(textCP27), DBDATE(textCP06), DBDATE(textCP07))
   End If
   'end 2023/03/17
   
   ' 設定滑鼠游標為等待狀態
   Screen.MousePointer = vbHourglass
   
'Add By Cheng 2002/11/06
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

Private Function CheckDataValid() As Boolean
   Dim nIndex As Integer
   Dim bFind As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
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
   If m_TM10 >= "010" Then
      If IsEmptyText(textCP44) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入代理人"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP44.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 延期後本所期限
   If IsEmptyText(textCP06) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入延期後本所期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   ' 延期後法定期限
   If IsEmptyText(textCP07) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入延期後法定期限"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP07.SetFocus
      GoTo EXITSUB
   End If
   ' 延期後本所期限範圍
   If Val(textCP06) > Val(textCP07) Then
      strTit = "檢核資料"
      strMsg = "延期後本所期限範圍不正確"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   ' 當案件性質為延期時, 未收文期限至少要選取一筆
   If m_CP10 = "303" Then
      If grdList.Rows <= 1 Then
         strTit = "檢核資料"
         strMsg = "未收文期限無資料, 無法執行延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         GoTo EXITSUB
      End If
      
      bFind = False
      For nIndex = 1 To grdList.Rows - 1
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            '記錄下一程序
            m_strNPCode = Me.grdList.TextMatrix(nIndex, 8)
            m_strNPName = Me.grdList.TextMatrix(nIndex, 1)
            bFind = True
            Exit For
         End If
      Next nIndex
      If bFind = False Then
         strTit = "檢核資料"
         strMsg = "請先選取未收文期限的資料來做延期的處理"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
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
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

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

Private Sub textCF09_GotFocus()
   InverseTextBox textCF09
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
Dim strTmp As String   '2008/9/8 ADD BY SONIA
    
    strTM23Nation = Empty
    If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
    ' 系統類別為TC
    If m_TM01 = "TC" Then
        ' 清除定稿例外欄位檔原有資料
        'add by nickc 2006/06/30
        If textPrint = "1" Then
            EndLetter "01", m_CP09, "14", strUserNum
        End If
    ' 系統類別非TC
    Else
        ' 案件性質非延期
        If m_CP10 <> "303" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                Select Case m_CP10
                Case "602", "604", "606", "613" '所有答辯
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "12", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "12" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & DBDATE(textCP27) & "')"
                    cnnConnection.Execute strSql
                'Add By Sindy 2013/12/30
                Case "210" '陳述意見
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "46", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "46" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & DBDATE(textCP27) & "')"
                    cnnConnection.Execute strSql
                Case Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "37", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "37" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & DBDATE(textCP27) & "')"
                    cnnConnection.Execute strSql
                End Select
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
                '2008/9/8 ADD BY SONIA
                Select Case m_strNPCode
                   Case "201" '補正
                      strTmp = "29"
                   Case Else
                      strTmp = "13"
                End Select
                '2008/9/8 end
                ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, strTmp, strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "列印備註" & "','" & textCP64 & "')"
                ' 法定期限
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "法定期限" & "','" & DBDATE(textCP07) & "')"
                cnnConnection.Execute strSql
                ' 本所期限
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "本所期限" & "','" & DBDATE(textCP06) & "')"
                cnnConnection.Execute strSql
                'Add By Sindy 2015/4/13 未齊備要印
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "未齊備要印" & "','" & IIf(Val(m_EP06) = 0, "♀", "") & "')"
                cnnConnection.Execute strSql
                '2015/4/13 END
            End If
        'Add By Cheng 2003/08/05
        '案件性質為延期
        Else
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                Select Case m_strNPCode
                Case "602", "604", "606", "613" '所有答辯
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "14", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & DBDATE(textCP27) & "')"
                    cnnConnection.Execute strSql
                    '下一程序名稱
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                             "'" & "下一程序名稱" & "','" & m_strNPName & "')"
                    cnnConnection.Execute strSql
                    '主管機關
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "14" & "','" & strUserNum & "'," & _
                             "'" & "主管機關" & "','" & GetCF10(m_TM01, m_TM10, m_strNPCode) & "')"
                    cnnConnection.Execute strSql
                Case Else
                    ' 清除定稿例外欄位檔原有資料
                    EndLetter "01", m_CP09, "19", strUserNum
                    ' 延期發文日
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                             "'" & "延期發文日" & "','" & DBDATE(textCP27) & "')"
                    cnnConnection.Execute strSql
                    '下一程序名稱
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                             "'" & "下一程序名稱" & "','" & m_strNPName & "')"
                    cnnConnection.Execute strSql
                    '主管機關
                    strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                             "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                             "'" & "主管機關" & "','" & GetCF10(m_TM01, m_TM10, m_strNPCode) & "')"
                    cnnConnection.Execute strSql
                    'add by sonia 2023/3/13 訴願二十日(定稿內容調整),行政訴訟上訴二十日
                    If m_strNPCode = "401" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                              "'" & "附註" & "','依訴願法第62條規定，經濟部將發函請訴願人於二十日內補正理由書，屆期未補正，經濟部訴願會將會為訴願不受理之決定')"
                    ElseIf m_strNPCode = "408" Then
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                              "'" & "附註" & "','此類補呈理由書通常須於二十日內提出')"
                    Else
                        strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              "VALUES ('" & "01" & "','" & m_CP09 & "','" & "19" & "','" & strUserNum & "'," & _
                              "'" & "附註" & "','此類補呈理由書通常須於三十日內提出')"
                    End If
                    cnnConnection.Execute strSql
                    'end 2023/3/13
                End Select
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
                '2008/9/8 ADD BY SONIA
                Select Case m_strNPCode
                   Case "201" '補正
                      strTmp = "28"
                   Case Else
                      strTmp = "16"
                End Select
                '2008/9/8 end
                ' 清除定稿例外欄位檔原有資料
                EndLetter "01", m_CP09, strTmp, strUserNum
                ' 列印備註
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "列印備註" & "','" & textCP64 & "')"
                ' 法定期限
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "法定期限" & "','" & DBDATE(textCP07) & "')"
                cnnConnection.Execute strSql
                ' 本所期限
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "本所期限" & "','" & DBDATE(textCP06) & "')"
                cnnConnection.Execute strSql
                '下一程序名稱
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "下一程序名稱" & "','" & m_strNPName & "')"
                cnnConnection.Execute strSql
                '主管機關
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "01" & "','" & m_CP09 & "','" & strTmp & "','" & strUserNum & "'," & _
                         "'" & "主管機關" & "','" & GetCF10(m_TM01, m_TM10, m_strNPCode) & "')"
                cnnConnection.Execute strSql
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
   bolEdit = False
   '2012/1/12 End
    
    ' 系統類別為TC
    If m_TM01 = "TC" Then
        ' 列印定稿
        'add by nickc 2006/06/30
        If textPrint = "1" Then
            'NowPrint m_CP09, "01", "14", False, strUserNum, 0
            ET03 = "14" 'Modify By Sindy 2012/1/12
        End If
    ' 系統類別非TC
    Else
        ' 案件性質非延期
        If m_CP10 <> "303" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               Select Case m_CP10
               Case "602", "604", "606", "613" '所有答辯
                   ' 列印定稿
                   'NowPrint m_CP09, "01", "12", False, strUserNum, 0
                   ET03 = "12" 'Modify By Sindy 2012/1/12
               'Add By Sindy 2013/12/30
               Case "210" '陳述意見
                   ET03 = "46"
               Case Else
                   ' 列印定稿
                   'NowPrint m_CP09, "01", "37", False, strUserNum, 0
                   ET03 = "37" 'Modify By Sindy 2012/1/12
               End Select
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               '2008/9/5 modify by sonia
               'NowPrint m_CP09, "01", "13", False, strUserNum, 0
               Select Case m_strNPCode
                  Case "201" '補正
                     'NowPrint m_CP09, "01", "29", False, strUserNum, 0
                     ET03 = "29" 'Modify By Sindy 2012/1/12
                  Case Else
                     'NowPrint m_CP09, "01", "13", False, strUserNum, 0
                     ET03 = "13" 'Modify By Sindy 2012/1/12
               End Select
               '2008/9/5 end
            End If
        'Add By Cheng 2003/08/05
        '案件性質為延期
        Else
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
               Select Case m_strNPCode
               Case "602", "604", "606", "613" '所有答辯
                   ' 列印定稿
                   'NowPrint m_CP09, "01", "14", False, strUserNum, 0
                   ET03 = "14" 'Modify By Sindy 2012/1/12
               Case Else
                   ' 列印定稿
                   'NowPrint m_CP09, "01", "19", False, strUserNum, 0
                   ET03 = "19" 'Modify By Sindy 2012/1/12
               End Select
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
               ' 列印定稿
               '2008/9/5 modify by sonia
               'NowPrint m_CP09, "01", "16", False, strUserNum, 0
               Select Case m_strNPCode
                  Case "201" '補正
                     'NowPrint m_CP09, "01", "28", False, strUserNum, 0
                     ET03 = "28" 'Modify By Sindy 2012/1/12
                  Case Else
                     'NowPrint m_CP09, "01", "16", False, strUserNum, 0
                     ET03 = "16" 'Modify By Sindy 2012/1/12
               End Select
               '2008/9/5 end
            End If
        End If
    End If
    
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
         'Add By Sindy 2019/12/20 + strLD18.信函總收文號
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

Private Sub txt_GotFocus(Index As Integer)
With Me.txt(Index)
   .SelStart = 0
   .SelLength = Len(.Text)
End With
End Sub

'Add By Sindy 2010/11/26
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
Cancel = False
If Len(Me.txt(0).Text) > 0 Then
   If CheckKeyIn(0) <> 0 Then
      Cancel = True
      txt_GotFocus 0
   End If
End If
End Sub
Private Function CheckKeyIn(Index) As Integer
Dim rs1 As New ADODB.Recordset

CheckKeyIn = -1
Select Case Index
Case 0 '延期承辦人
   If Me.txt(0).Visible = False Then CheckKeyIn = 0: Exit Function
   If rs1.State <> adStateClosed Then rs1.Close
   Set rs1 = Nothing
   rs1.CursorLocation = adUseClient
   rs1.Open " Select ST02 From Staff Where ST01 = '" & Me.txt(Index).Text & "' And ST04 = '1' ", cnnConnection, adOpenStatic, adLockReadOnly
   If rs1.RecordCount > 0 Then
      Me.lbl(1).Caption = "" & rs1.Fields(0).Value
      CheckKeyIn = 0
   Else
      MsgBox "延期承辦人輸入錯誤!!!", vbExclamation
      Me.lbl(1).Caption = ""
   End If
   rs1.Close
   Set rs1 = Nothing
Case Else
   CheckKeyIn = 0
End Select
End Function

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
       'Added by Lydia 2022/05/23 法律所案源：B2類PT案取得法律案源之發文規費，並且有輸入發文規費才做檢查
       If m_LosCP84 <> "0" Then
          If Val(m_LosCP84) <> Val(Trim(textCP84.Text)) Then
              If MsgBox("法律所收文規費[" & Trim(Val(m_LosCP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
                  textCP84_GotFocus
                  Exit Function
              End If
          End If
       Else
       'end 2022/05/23
          If Val(textCP84.Text) <> Val(m_CP84) Then
              If MsgBox("收文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", vbOKCancel) = vbCancel Then
                  textCP84_GotFocus
                  Exit Function
              End If
          End If
       End If 'Added by Lydia 2022/05/23
   End If
      
   If Me.textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
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
   
   If Me.textCP27.Enabled = True Then
      Cancel = False
      textCP27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   'If Me.textCP09S.Enabled = True Then
   '   Cancel = False
   '   textCP09S_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If Me.textCP36.Enabled = True Then
      Cancel = False
      textCP36_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP37.Enabled = True Then
      Cancel = False
      textCP37_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP37_1.Enabled = True Then
      Cancel = False
      textCP37_1_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textCP38.Enabled = True Then
      Cancel = False
      textCP38_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.textCP39.Enabled = True Then
      Cancel = False
      textCP39_Validate Cancel
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
   
   If Me.textPrint.Enabled = True Then
      Cancel = False
      textPrint_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   For Each objTxt In txt
      If objTxt.Enabled = True And objTxt.Visible = True Then
         Cancel = False
         txt_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
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
        SSTab1.Tab = 0
        txtCP113.SetFocus
        txtCP113_GotFocus
        Exit Function
    End If
    'end 2021/06/04
    
   TxtValidate = True
End Function

'edit by nickc 2006/01/27
' 91.09.02 modify by louis
'Private Sub textAgName_GotFocus()
'   InverseTextBox textAgName
'   textAgName.IMEMode = 1
'End Sub
'
'' 91.09.02 modify by louis
'' 本所出名代理人
'Private Sub textAgName_Validate(Cancel As Boolean)
'   Cancel = False
'   If CheckLengthIsOK(textAgName, 10) = False Then
'      Cancel = True
'   End If
'   If Cancel = False Then: textAgName.IMEMode = 2
'End Sub

'Add By Cheng 2002/11/19
'取得案件收費表的工作天數
''''edit by nickc 2007/10/12 取消不用，改抓有時效性的
''''Private Function GetCF04(strCF01 As String, strCF02 As String, strCF03 As String) As String
''''Dim rsA As New ADODB.Recordset
''''Dim StrSQLa As String
''''
''''GetCF04 = "0"
''''If rsA.State <> adStateClosed Then rsA.Close
''''Set rsA = Nothing
''''StrSQLa = "Select CF04 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF04 IS NOT NULL"
''''rsA.CursorLocation = adUseClient
''''rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
''''If rsA.RecordCount > 0 Then
''''   GetCF04 = "" & rsA.Fields(0).Value
''''End If
''''If rsA.State <> adStateClosed Then rsA.Close
''''Set rsA = Nothing
''''
''''End Function

'Add By Cheng 2003/08/05
'取得主管機關
Private Function GetCF10(strCF01 As String, strCF02 As String, strCF03 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetCF10 = ""
StrSQLa = "Select CF10 From CaseFee Where CF01='" & strCF01 & "' And CF02='" & strCF02 & "' And CF03='" & strCF03 & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    GetCF10 = "" & rsA.Fields(0).Value
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

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

