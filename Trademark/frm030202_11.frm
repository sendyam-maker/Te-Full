VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_11 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(延期)"
   ClientHeight    =   6600
   ClientLeft      =   710
   ClientTop       =   2210
   ClientWidth     =   9140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   9140
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   150
      TabIndex        =   46
      Top             =   3060
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   6156
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm030202_11.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label37"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label36"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label28"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label25"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label10"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label11"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label39"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblNameAgent"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(12)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label16"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label43"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lstNameAgent"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP64"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "grdList"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textPetition"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP18"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textDN"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP27"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textPrint"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP07"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP06"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP84"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP113"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "Text9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP118"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "代表人"
      TabPicture(1)   =   "frm030202_11.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textTM50"
      Tab(1).Control(1)=   "textTM51"
      Tab(1).Control(2)=   "textTM52"
      Tab(1).Control(3)=   "textTM47"
      Tab(1).Control(4)=   "textTM48"
      Tab(1).Control(5)=   "textTM49"
      Tab(1).Control(6)=   "Label30"
      Tab(1).Control(7)=   "Label31"
      Tab(1).Control(8)=   "Label32"
      Tab(1).Control(9)=   "Label33"
      Tab(1).Control(10)=   "Label34"
      Tab(1).Control(11)=   "Label35"
      Tab(1).ControlCount=   12
      Begin VB.TextBox textCP118 
         Height          =   285
         Left            =   4305
         MaxLength       =   1
         TabIndex        =   10
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   4
         Top             =   645
         Width           =   372
      End
      Begin VB.TextBox textCP113 
         Height          =   285
         Left            =   6150
         MaxLength       =   4
         TabIndex        =   2
         Top             =   360
         Width           =   600
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   5820
         MaxLength       =   1
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1305
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   3810
         TabIndex        =   1
         Top             =   360
         Width           =   1425
      End
      Begin VB.TextBox textCP06 
         Height          =   285
         Left            =   4305
         MaxLength       =   7
         TabIndex        =   5
         Top             =   660
         Width           =   1212
      End
      Begin VB.TextBox textCP07 
         Height          =   285
         Left            =   7605
         MaxLength       =   7
         TabIndex        =   6
         Top             =   660
         Width           =   1212
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   9
         Top             =   1260
         Width           =   492
      End
      Begin VB.TextBox textCP27 
         Height          =   285
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   1212
      End
      Begin VB.TextBox textDN 
         Height          =   285
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   7
         Top             =   960
         Width           =   492
      End
      Begin VB.TextBox textCP18 
         Height          =   285
         Left            =   7710
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox textPetition 
         Height          =   285
         Left            =   4305
         MaxLength       =   1
         TabIndex        =   8
         Top             =   960
         Width           =   492
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1092
         Left            =   1200
         TabIndex        =   85
         Top             =   1776
         Width           =   7515
         _ExtentX        =   13247
         _ExtentY        =   1923
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
      Begin MSForms.TextBox textTM50 
         Height          =   285
         Left            =   -73770
         TabIndex        =   16
         Top             =   1320
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM51 
         Height          =   285
         Left            =   -73770
         TabIndex        =   17
         Top             =   1635
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM52 
         Height          =   285
         Left            =   -73770
         TabIndex        =   18
         Top             =   1965
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM47 
         Height          =   285
         Left            =   -73770
         TabIndex        =   13
         Top             =   360
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM48 
         Height          =   285
         Left            =   -73770
         TabIndex        =   14
         Top             =   675
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   80
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textTM49 
         Height          =   285
         Left            =   -73770
         TabIndex        =   15
         Top             =   1005
         Width           =   7500
         VariousPropertyBits=   671105051
         MaxLength       =   40
         Size            =   "13229;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   1200
         TabIndex        =   12
         Top             =   2910
         Width           =   7515
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13250;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   7140
         TabIndex        =   11
         Top             =   960
         Width           =   1500
         VariousPropertyBits=   746586139
         ScrollBars      =   2
         DisplayStyle    =   2
         Size            =   "2646;556"
         MatchEntry      =   0
         ListStyle       =   1
         MultiSelect     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:                     (Y: 是)"
         Height          =   180
         Left            =   2865
         TabIndex        =   75
         Top             =   1320
         Width           =   2580
      End
      Begin VB.Label Label16 
         Caption         =   "延期月數 :"
         Height          =   180
         Left            =   120
         TabIndex        =   74
         Top             =   700
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   12
         Left            =   5340
         TabIndex        =   73
         Top             =   420
         Width           =   765
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   6165
         TabIndex        =   68
         Top             =   1320
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   2880
         TabIndex        =   67
         Top             =   420
         Width           =   900
      End
      Begin VB.Label Label30 
         Caption         =   "代表人1(中) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   64
         Top             =   360
         Width           =   1212
      End
      Begin VB.Label Label31 
         Caption         =   "代表人1(英) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   63
         Top             =   684
         Width           =   1212
      End
      Begin VB.Label Label32 
         Caption         =   "代表人1(日) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   62
         Top             =   1008
         Width           =   1212
      End
      Begin VB.Label Label33 
         Caption         =   "代表人2(中) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   61
         Top             =   1332
         Width           =   1212
      End
      Begin VB.Label Label34 
         Caption         =   "代表人2(英) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   60
         Top             =   1656
         Width           =   1212
      End
      Begin VB.Label Label35 
         Caption         =   "代表人2(日) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   59
         Top             =   1981
         Width           =   1212
      End
      Begin VB.Label Label11 
         Caption         =   "延期後法定期限 :"
         Height          =   255
         Left            =   6195
         TabIndex        =   58
         Top             =   700
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "延期後本所期限 :"
         Height          =   255
         Left            =   2865
         TabIndex        =   57
         Top             =   700
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "延期點數 :"
         Height          =   255
         Index           =   10
         Left            =   6850
         TabIndex        =   56
         Top             =   375
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   255
         Left            =   1920
         TabIndex        =   55
         Top             =   1260
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   54
         Top             =   1260
         Width           =   972
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   120
         TabIndex        =   53
         Top             =   420
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2910
         Width           =   975
      End
      Begin VB.Label Label36 
         Caption         =   "是否輸入D/N :"
         Height          =   252
         Left            =   120
         TabIndex        =   51
         Top             =   1000
         Width           =   1212
      End
      Begin VB.Label Label37 
         Caption         =   "(Y:輸入)"
         Height          =   255
         Left            =   1920
         TabIndex        =   50
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "是否列印申請書 :"
         Height          =   255
         Left            =   2865
         TabIndex        =   49
         Top             =   1000
         Width           =   1455
      End
      Begin VB.Label Label7 
         Caption         =   "(Y:印)"
         Height          =   255
         Left            =   4935
         TabIndex        =   48
         Top             =   1000
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "欲延期期限 :"
         Height          =   180
         Left            =   120
         TabIndex        =   47
         Top             =   1770
         Width           =   990
      End
   End
   Begin VB.TextBox textDL02 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   66
      TabStop         =   0   'False
      Top             =   2445
      Width           =   1545
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8070
      TabIndex        =   22
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5910
      TabIndex        =   20
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6870
      TabIndex        =   21
      Top             =   0
      Width           =   1152
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5730
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1020
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1305
      Width           =   1545
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   450
      Width           =   1545
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   735
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   735
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   7230
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   450
      Width           =   1545
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   4050
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   450
      Width           =   1545
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5730
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   2160
      Width           =   2532
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   960
      TabIndex        =   84
      Top             =   2445
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   960
      TabIndex        =   83
      Top             =   1875
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM80 
      Height          =   285
      Left            =   5730
      TabIndex        =   82
      Top             =   1875
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM81 
      Height          =   285
      Left            =   960
      TabIndex        =   81
      Top             =   2160
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   960
      TabIndex        =   80
      Top             =   1590
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM78 
      Height          =   285
      Left            =   5730
      TabIndex        =   79
      Top             =   1590
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP14 
      Height          =   285
      Left            =   7230
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   1305
      Width           =   1545
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "2725;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   4050
      TabIndex        =   77
      Top             =   1260
      Width           =   1545
      VariousPropertyBits=   671105055
      Size            =   "2725;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   960
      TabIndex        =   76
      Top             =   2730
      Width           =   8055
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14208;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "申請人5 :"
      Height          =   180
      Left            =   90
      TabIndex        =   72
      Top             =   2208
      Width           =   720
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "申請人4 :"
      Height          =   180
      Left            =   4695
      TabIndex        =   71
      Top             =   1922
      Width           =   720
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人3 :"
      Height          =   180
      Left            =   90
      TabIndex        =   70
      Top             =   1922
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "申請人2 :"
      Height          =   180
      Left            =   4695
      TabIndex        =   69
      Top             =   1636
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "上次延期日 :"
      Height          =   180
      Index           =   5
      Left            =   4695
      TabIndex        =   65
      Top             =   2494
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "案件名稱 :"
      Height          =   180
      Left            =   120
      TabIndex        =   45
      Top             =   2782
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "審定號數 :"
      Height          =   180
      Left            =   90
      TabIndex        =   44
      Top             =   1064
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "智權人員 :"
      Height          =   180
      Index           =   11
      Left            =   3015
      TabIndex        =   43
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "彼所案號 :"
      Height          =   180
      Index           =   9
      Left            =   4695
      TabIndex        =   42
      Top             =   1064
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "案件性質 :"
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   41
      Top             =   1350
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "收文號 :"
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   40
      Top             =   492
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號 :"
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   39
      Top             =   778
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   4695
      TabIndex        =   38
      Top             =   778
      Width           =   810
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "發證日 :"
      Height          =   180
      Index           =   3
      Left            =   6450
      TabIndex        =   37
      Top             =   492
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "業務區別 :"
      Height          =   180
      Index           =   2
      Left            =   2985
      TabIndex        =   36
      Top             =   492
      Width           =   810
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "承辦人 :"
      Height          =   180
      Left            =   6450
      TabIndex        =   35
      Top             =   1350
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "商標種類 :"
      Height          =   180
      Index           =   4
      Left            =   4695
      TabIndex        =   34
      Top             =   2208
      Width           =   810
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "申請人1 :"
      Height          =   180
      Left            =   90
      TabIndex        =   33
      Top             =   1636
      Width           =   720
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   90
      TabIndex        =   32
      Top             =   2494
      Width           =   630
   End
End
Attribute VB_Name = "frm030202_11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/17 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/09 改成Form2.0 ;cmbTM05、textCP13、textCP14、textCP64、textTM44、textTM23、textTM78~81、lstNameAgent、textTM47~52；grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 原本所期限
Dim m_CP06 As String
' 原法定期限
Dim m_CP07 As String
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 業務區
Dim m_CP12 As String
' 智權人員
Dim m_CP13 As String
'承辦人 Add By Sindy 98/03/11
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 發文時間
' 代理人
Dim m_CP44 As String
' 彼所案號
Dim m_CP45 As String

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
'
Dim m_CurrSel As Integer
'Add By Cheng 2002/08/19
Dim m_strDL05 As String
'add by nick 2004/08/13
Dim m_CP84 As String       '發文規費
'add by nickc 2006/01/26
Dim m_CP110 As String
'add by nickc 2008/02/22
Dim m_CP44no As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
Dim m_CP43 As String 'Add By Sindy 2012/5/4 相關總收文號
'Added by Lydia 2023/03/17
Dim m_LOS15 As String '法律所案源單號
Dim m_LOS02 As String '法律所案源類別

Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      
      'Add by Sindy 98/3/24 設定是否算發文室案件
      If m_TM10 = "000" Then
         'Modify By Sindy 2012/12/20 若為電子送件則不經發文室
         'Modify By Sindy 2023/8/1 電子送件欄位值不是空白者,即為電子送件
         If (textCP118.Visible = True And textCP118 <> "") Then
            'Added by Morgan 2016/5/16 電子送件也要記錄主管機關
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27, , True) = False Then
               Exit Sub
            End If
            'end 2016/5/16
            'add by sonia 2016/3/31
            strExc(0) = Trim(InputBox("請輸入智慧局收文文號!!"))
            If strExc(0) = "" Then
               Exit Sub
            Else
               textCP64 = "智慧局收文文號:" & strExc(0) & ";" & Trim(textCP64)
            End If
            'end 2016/3/31
         Else
            'Add by Sindy 2009/4/24
            If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
               Exit Sub
            Else
               If m_CP123s = "Y" Then
                  'modify by sonia 2014/6/23 加傳發文規費, P-108903
                  If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP84, textCP27, IIf(m_CP10 <> "303", True, False)) = False Then
                      Exit Sub
                  End If
               End If
            End If
         End If '2012/12/20 End
      End If
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 更新欄位輸入的內容
      OnUpdateField
      ' 存檔
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Add By Sindy 2012/4/5 CFT,FCT所有案件性質發文時,檢查代表圖是否存在
      'Mark by Amy 2018/07/31 因ChkIsExistImg不使用,與Sindy確認FCT不彈Msg故拿掉
      'Call ChkIsExistImg(m_TM01, m_TM02, m_TM03, m_TM04)
        
        'Added by Lyddia 2018/08/10 增加重新發文判斷
        strExc(1) = m_CP82
        If Val(m_CP82) > 0 Then
             If MsgBox("重新發文是否上傳檔案到卷宗區？", vbInformation + vbYesNo + vbDefaultButton1) = vbYes Then
                 strExc(1) = ""
             End If
        End If
        If Val(strExc(1)) = 0 Then
        'end 2018/08/10
            'Added by Lydia 2018/07/19 FCT發文自動將下載的PDF檔,上傳到卷宗區
            If Pub_AutoSavePdf_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_CP09, m_CP10) = False Then
            End If
            'end 2018/07/19
        End If 'end 2018/08/10
        
      '************   90.11.23 nick   清畫面
      'frm030202_01.radio(0).Value = True
      'frm030202_01.textCP09.Enabled = True
      'frm030202_01.textCP09.Text = ""
      'frm030202_01.textTM01.Enabled = False
      'frm030202_01.textTM01.Text = ""
      'frm030202_01.textTM02.Enabled = False
      'frm030202_01.textTM02.Text = ""
      'frm030202_01.textTM02_2.Enabled = False
      'frm030202_01.textTM02_2.Text = ""
      'frm030202_01.textTM03.Enabled = False
      'frm030202_01.textTM03.Text = ""
      'frm030202_01.textTM04.Enabled = False
      'frm030202_01.textTM04.Text = ""
      'frm030202_01.grdList.Clear
      'frm030202_01.grdList.Rows = 2
      'frm030202_01.QueryData
      'frm030202_01.Show
      '*************************************
      
      Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 外商發文時,增加發Mail通知承辦人及副本給判發主管
      'Add By Sindy 2024/8/19
      If frm030202_01.bolIsEMPFlow = True Then
         frm090202_4.QueryData
      End If
      '2024/8/19 End
      'Ken 91.04.09 -- Start
      If textDN = "Y" Then
        'Add By Cheng 2003/03/19
        '新增地址條列表資料
'edit by nick 2004/11/17  因為請款已經有產生了
'        pub_AddressListSN = pub_AddressListSN + 1
'        PUB_AddNewAddressList strUserNum, m_TM01, m_TM02, m_TM03, m_TM04, "" & pub_AddressListSN, "0"
         Screen.MousePointer = vbHourglass
         Frmacc21h0.Show
         mdiMain.ToolShow
         mdiMain.tool1_enabled
         Screen.MousePointer = vbDefault
         Set Frmacc21h0.frmlink = frm030202_01
         'add by nick 2004/11/24
         Frmacc21h0.IsPrintAddress = False
            'Add By Cheng 2003/04/24
            Me.Visible = False
            strFormName = Frmacc21h0.Name
            Do While PUB_CheckFormExist(strFormName)
                DoEvents
            Loop
            If m_CP10 = "303" Then PUB_GetCPunIssueDatas "" & Me.textTMKey.Text
            frm030202_01.Show
            frm030202_01.Clear1
      Else
         'Add By Cheng 2002/04/30
         '若有未發文資料顯示警告
         If m_CP10 = "303" Then
            If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = False Then
               'Add By Sindy 2024/8/19
               If frm030202_01.bolIsEMPFlow = True Then
                  Unload frm030202_01
                  frm090202_4.Show
                  Unload Me
                  Exit Sub
               End If
               '2024/8/19 End
            End If
         End If
         frm030202_01.Show
         ' 90.12.07 modify by louis
         frm030202_01.Clear1
      End If
      'Ken 91.04.09 -- End
      Unload Me
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
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/01/30
   textTM78.BackColor = &H8000000F
   textTM79.BackColor = &H8000000F
   textTM80.BackColor = &H8000000F
   textTM81.BackColor = &H8000000F
   
   textTM44.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
      
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP12.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14.BackColor = &H8000000F
   textCP18.BackColor = &H8000000F
   
   textDL02.BackColor = &H8000000F
   
   'Added by Lydia 2016/09/10 設定代表人中文名稱和英文名稱長度
    textTM47.MaxLength = Pub_MaxCEL10
    textTM48.MaxLength = Pub_MaxCEL11
    textTM50.MaxLength = Pub_MaxCEL10
    textTM51.MaxLength = Pub_MaxCEL11
   'end 2016/09/10
   
   MoveFormToCenter Me
   'Add by nickc 2006/01/26
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   'Added by Lydia 2021/09/09 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 765
   lstNameAgent.Width = 1500

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
      m_TM119 = CheckStr(rsTmp.Fields("TM119"))
      m_TM120 = CheckStr(rsTmp.Fields("TM120"))
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then: textTM15 = rsTmp.Fields("TM15")
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then: textTM12 = rsTmp.Fields("TM12")
      ' 發證日
      If IsNull(rsTmp.Fields("TM20")) = False Then: textTM20 = TAIWANDATE(rsTmp.Fields("TM20"))
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
      ' 商標種類
'      textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      textTM08 = GetTradeMarkName("" & rsTmp.Fields("TM08"), 0)
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      'add by nickc 2007/01/30
      If IsNull(rsTmp.Fields("TM78")) = False Then: textTM78 = GetCustomerName("" & rsTmp.Fields("TM78"), 0)
      If IsNull(rsTmp.Fields("TM79")) = False Then: textTM79 = GetCustomerName("" & rsTmp.Fields("TM79"), 0)
      If IsNull(rsTmp.Fields("TM80")) = False Then: textTM80 = GetCustomerName("" & rsTmp.Fields("TM80"), 0)
      If IsNull(rsTmp.Fields("TM81")) = False Then: textTM81 = GetCustomerName("" & rsTmp.Fields("TM81"), 0)
      
      ' FC代理人
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
      ' 代表人1(中)
      If IsNull(rsTmp.Fields("TM47")) = False Then: textTM47 = rsTmp.Fields("TM47")
      SetTMSPFieldOldData "TM47", textTM47, 0
      ' 代表人1(英)
      If IsNull(rsTmp.Fields("TM48")) = False Then: textTM48 = rsTmp.Fields("TM48")
      SetTMSPFieldOldData "TM48", textTM48, 0
      ' 代表人1(日)
      If IsNull(rsTmp.Fields("TM49")) = False Then: textTM49 = rsTmp.Fields("TM49")
      SetTMSPFieldOldData "TM49", textTM49, 0
      ' 代表人2(中)
      If IsNull(rsTmp.Fields("TM50")) = False Then: textTM50 = rsTmp.Fields("TM50")
      SetTMSPFieldOldData "TM50", textTM50, 0
      ' 代表人2(英)
      If IsNull(rsTmp.Fields("TM51")) = False Then: textTM51 = rsTmp.Fields("TM51")
      SetTMSPFieldOldData "TM51", textTM51, 0
      ' 代表人2(日)
      If IsNull(rsTmp.Fields("TM52")) = False Then: textTM52 = rsTmp.Fields("TM52")
      SetTMSPFieldOldData "TM52", textTM52, 0
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
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   ' 系統日
   strDate = TAIWANDATE(SystemDate())
   ' 收文號
   textCP09 = m_CP09
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'add by nickc 2008/02/22
      m_CP116 = CheckStr(rsTmp.Fields("CP116"))
      m_CP44no = CheckStr(rsTmp.Fields("CP44"))
      m_CP82 = "" & rsTmp.Fields("CP82")  'Added by Lydia 2018/08/10 發文時間
      ' 本所期限
      m_CP06 = "0"
      If IsNull(rsTmp.Fields("CP06")) = False Then
         m_CP06 = rsTmp.Fields("CP06")
      End If
      ' 法定期限
      m_CP07 = "0"
      If IsNull(rsTmp.Fields("CP07")) = False Then
         m_CP07 = rsTmp.Fields("CP07")
      End If
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      m_CP43 = Empty 'Add By Sindy 2012/5/4
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         m_CP43 = rsTmp.Fields("CP43") 'Add By Sindy 2012/5/4
         '案件性質
         textCP10 = textCP10 & PUB_GetRelateCasePropertyName(m_CP09, "1")
      End If
      '2010/12/27 End
      
      ' 業務區別
      '910718 Sieg
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      
      'Add By Sindy 98/03/11
      '工作時數
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      '承辦人
      m_CP14 = "" & rsTmp.Fields("CP14")
      '98/03/11 End
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 發文日(預設為系統日)
      strCP27 = Empty
      textCP27 = TAIWANDATE(SystemDate())
      If IsNull(rsTmp.Fields("CP27")) = False Then: strCP27 = rsTmp.Fields("CP27")
      
      'Modify By Sindy 2009/05/01
      'CaculateNP08NP09 ()
      Call CaculateNP08NP09(m_CP07, m_CP10)
      
      SetCPFieldOldData "CP27", strCP27, 1
      ' 點數
      textCP18 = Empty
      'Add By Sindy 2014/10/31
      If m_CP10 = "201" Then '補正案件性質按延期按鈕時,不預設點數及開放可以輸入
         textCP18.Locked = False
         textCP18.BackColor = &H80000005
      Else
      '2014/10/31 END
         textCP18.Locked = True
         textCP18.BackColor = &H8000000F
         If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      End If
      ' 代理人
      If IsNull(rsTmp.Fields("CP44")) = False Then: m_CP44 = GetFAgentName(rsTmp.Fields("CP44"))
      ' 彼所案號
      If IsNull(rsTmp.Fields("CP45")) = False Then: m_CP45 = rsTmp.Fields("CP45")
      ' 進度備註
      textCP64 = Empty
      If IsNull(rsTmp.Fields("CP64")) = False Then: textCP64 = rsTmp.Fields("CP64")
      SetCPFieldOldData "CP64", textCP64, 0
      
      'Add By Sindy 2012/12/20
      ' 是否電子送件
      textCP118 = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         textCP118 = rsTmp.Fields("CP118")
      End If
      SetCPFieldOldData "CP118", textCP118, 0
      
      'add by nick 2004/08/13 發文規費
      If IsNull(rsTmp.Fields("CP17")) = False And textCP84.Enabled = True Then
          m_CP84 = CheckStr(rsTmp.Fields("CP17"))
      End If
      'Add By Sindy 2012/12/20 電子送件發文規費預設為承辦人已輸入的金額
      If rsTmp.Fields("cp118") = "Y" Then
         textCP84 = Val("" & rsTmp.Fields("cp84"))
      End If
      'end 2012/12/20
      
      'add by nickc 2006/02/10
      Text7 = CheckStr(rsTmp.Fields("CP22"))
      SetCPFieldOldData "CP22", Text7, 0
   End If
   'add by nickc 2006/01/26
   'SetCPFieldOldData "CP110", m_CP110, 0
   'Modify By Sindy 2010/9/20
   If m_CP110 = "" Then m_CP110 = CheckStr(rsTmp.Fields("cp110"))
   SetCPFieldOldData "CP110", CheckStr(rsTmp.Fields("cp110")), 0
   '2010/9/20 End
   
   rsTmp.Close
   Set rsTmp = Nothing
   Set rsSubTmp = Nothing
   
   'Add By Sindy 2012/5/4 預設延期月數
   If m_CP43 <> "" Then
      Text9 = GetDelayMonth(m_CP43)
      If Text9 = "" Then Text9.Text = "1"
      Dim Cancel As Boolean
      Call Text9_Validate(Cancel)
   End If
   '2012/5/4 End
   
   'Added by Lydia 2023/03/17 法律所案源：取得案源類別
   If m_CP10 = "303" Then
       strExc(0) = "select cp162,los02 from caseprogress,lawofficesource where cp09='" & m_CP43 & "' and cp162=los15(+) "
   Else
       strExc(0) = "select cp162,los02 from caseprogress,lawofficesource where cp09='" & m_CP09 & "' and cp162=los15(+) "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
       m_LOS15 = "" & RsTemp.Fields("cp162")
       m_LOS02 = "" & RsTemp.Fields("los02")
   End If
   'end 2023/03/17
End Sub

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   'add by nickc 2006/01/26
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
   
   ' 本所案號
'   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   textTMKey.Text = m_TM01 & "-" & m_TM02 & "-" & IIf(Len("" & m_TM03) <= 0, "0", m_TM03) & "-" & IIf(Len("" & m_TM04) <= 0, "00", m_TM04)
   
   'add by nickc 2006/01/26
   tm(1) = m_TM01
   tm(2) = m_TM02
   tm(3) = m_TM03
   tm(4) = m_TM04
   'Modify By Sindy 2010/9/20 預設出名代理人,移到下面讀完CP再做
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   '2010/9/20 End
   
   '讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   'Modified by Lydia 2021/09/09 + Form 2.0 = True
   'PUB_SetOurAgent lstNameAgent, tm(), m_CP110
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True 'Modify By Sindy 2010/9/20
   
   ' 本案期限
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
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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
               grdList.TextMatrix(grdList.row, 1) = GetCaseTypeName(m_TM01, rsTmp.Fields("NP07"))
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
         'Added by Lydia 2023/10/17
         If grdList.Rows >= 2 Then
            grdList.FixedRows = 1
         End If
         'end 2023/10/17
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
            textDL02 = TAIWANDATE(rsTmp.Fields("DL02"))
         End If
      End If
   End If
   rsTmp.Close
   
   'Add By Sindy 2012/12/20 外商000台灣案所有案件性質加電子送件功能
   If m_TM01 = "FCT" And m_TM10 = "000" Then
      Label43.Visible = True
      textCP118.Visible = True
   Else
      Label43.Visible = False
      textCP118.Visible = False
   End If
   '2012/12/20 End

   Set rsTmp = Nothing
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2023/03/17
   'Add By Cheng 2002/07/19
   Set frm030202_11 = Nothing
End Sub

' 使用者點選GridList中的內容時
Private Sub grdList_Click()
Dim i As Integer
Dim Cancel As Boolean
   
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         If grdList.TextMatrix(grdList.row, 0) = "V" Then
            grdList.TextMatrix(grdList.row, 0) = Empty
         Else
            'Add By Sindy 2009/05/01
            For i = 1 To grdList.Rows - 1
               grdList.TextMatrix(i, 0) = Empty
            Next i
            '2009/05/01 End
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
               If grdList.TextMatrix(grdList.row, 7) <> "" Then 'Modify By Sindy 2012/5/17 +if
                  'Add By Sindy 2012/5/4 預設延期月數
                  m_CP07 = DBDATE(strExc(2))
                  Text9 = GetDelayMonth(grdList.TextMatrix(grdList.row, 7))
                  If Text9 = "" Then Text9.Text = "1"
                  Call Text9_Validate(Cancel)
                  '2012/5/4 End
               Else
                  'Add By Sindy 2009/05/01
                  Call CaculateNP08NP09(grdList.TextMatrix(grdList.row, 3), grdList.TextMatrix(grdList.row, 8))
               End If
            End If
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

'add by nickc 2006/01/26
'檢查並設定cp110資料
Private Sub lstNameAgent_Validate(Cancel As Boolean)
   Dim ii As Integer, bolCheck As Boolean
   bolCheck = False
   m_CP110 = ""
   For ii = 0 To lstNameAgent.ListCount - 1
      If lstNameAgent.Selected(ii) = True Then
         'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
         'm_CP110 = m_CP110 & "," & lstNameAgent.ItemData(ii)
         'Modified by Lydia 2021/09/09 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      Text7 = "N"
      MsgBox "未勾選代理人!", vbInformation, "必要欄位！"
      Cancel = True
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
Dim i As Integer
   
   If Val(Text9) > 0 Then
      'Modify By Sindy 2012/5/11
'      strExc(1) = m_TM01
'      strExc(2) = m_TM10
'      strExc(3) = CompDate("1", Val(Text9), m_CP07)
'      GetCtrlDT strExc
'      textCP07 = TransDate(strExc(3), 1)
'      textCP06 = TransDate(PUB_GetWorkDay1(strExc(0), True), 1)
      '延期月數
      textCP07 = TAIWANDATE(AddMonth(DBDATE(m_CP07), Val(Text9)))
      If Val(Text9) >= 2 Then
         i = -4
      Else
         i = -2
      End If
      textCP07 = PUB_FCTGetDelaySpecDay(textCP07, m_CP43) 'Modify By Sindy 2014/3/5
      If textCP07 <> "" Then
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
'      '2012/5/11 End
   End If
End Sub
'end 2009/12/29

Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的延期後本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextInverse textCP06
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/09
      End If
   End If
End Sub

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

'add by sonia 2017/10/25
Private Sub textCP18_GotFocus()
   InverseTextBox textCP06
End Sub
'end 2017/10/25

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
      'edit by nick 2004/08/31 系統日加一天
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'edit by nick 2004/08/31
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         GoTo EXITSUB
      End If
      'Modify By Sindy 2009/05/01
      'CaculateNP08NP09
      
   End If
EXITSUB:
End Sub

' 計算本所期限及法定期限
'Modify By Sindy 2009/05/01 增加傳入日期
Private Sub CaculateNP08NP09(strDate As String, strCF03 As String)
   'Modify By Sindy 2009/05/01
   'If IsEmptyText(textCP27) = False Then
      'strExc(0) = TransDate(textCP27.Text, 2)
   If IsEmptyText(strDate) = False Then
      strExc(0) = TransDate(strDate, 2)
      'edit by nickc 2007/02/06 不用 dll 了
      'If objLawDll.GetCaseFeeDelay(m_TM01, m_TM10, m_CP10, strExc) Then
      If ClsLawGetCaseFeeDelay(m_TM01, m_TM10, strCF03, strExc) Then
         textCP07 = TransDate(strExc(1), 1)
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(strExc(2), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

'edit by nickc 2006/01/27
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

Private Sub textDN_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否輸入D/N
Private Sub textDN_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textDN) = False Then
      Select Case textDN
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textDN_GotFocus
      End Select
   End If
End Sub

Private Sub textPetition_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否列印申請書
Private Sub textPetition_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textPetition) = False Then
      Select Case textPetition
         Case " ", "Y":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPetition_GotFocus
      End Select
   End If
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 列印定稿
Private Sub textPrint_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
      
   If IsEmptyText(textPrint) = False Then
      Select Case textPrint
         Case " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            'textPrint_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 代表人1(中)
   SetTMSPFieldNewData "TM47", textTM47
   ' 代表人1(英)
   SetTMSPFieldNewData "TM48", textTM48
   ' 代表人1(日)
   SetTMSPFieldNewData "TM49", textTM49
   ' 代表人2(中)
   SetTMSPFieldNewData "TM50", textTM50
   ' 代表人2(英)
   SetTMSPFieldNewData "TM51", textTM51
   ' 代表人2(日)
   SetTMSPFieldNewData "TM52", textTM52
   
   ' 發文日
   'SetCPFieldNewData "CP27", DBDATE(textCP27)
   ' 進度備註
   '910801 Sieg 602
'edit by nickc 2006/01/26
'   If textCP64_2 <> "" Then
'      If textCP64 = "" Then
'         textCP64 = textCP64_2
'      Else
'         textCP64 = textCP64 & "," & textCP64_2
'      End If
'   End If
   SetCPFieldNewData "CP64", textCP64
   
   'add by nickc 2006/01/26
   SetCPFieldNewData "CP110", m_CP110
   'add by nickc 2006/02/10
   SetCPFieldNewData "CP22", Text7
   ' Add By Sindy 98/03/11
   SetCPFieldNewData "CP113", textCP113
   
   'Add By Sindy 2012/12/20
   ' 是否電子送件
   SetCPFieldNewData "CP118", textCP118
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateTradeMark()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
      
   ' 更新案件進度檔
   strSql = "UPDATE TradeMark SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_TMSPCount - 1
      strTmp = Empty
      If m_TMSPList(nIndex).fiOldData <> m_TMSPList(nIndex).fiNewData Then
         If m_TMSPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_TMSPList(nIndex).fiName & " = '" & m_TMSPList(nIndex).fiNewData & "'"
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
End Sub

' 更新商標基本檔的相關欄位
Private Sub OnUpdateCaseProperty()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   Dim strDL01 As String
   Dim strDL02 As String
   Dim strDL03 As String
   Dim strDL04 As String
   
   ' 新增資料到延期記錄檔
   'strDL01 = Empty
   'strDL03 = Empty
   'strDL04 = Empty
   'If m_CP10 = "303" Then
   '   ' 案件性質為延期時, 總收文號, 本所期限及法定期限為未收文期限所選取的收文資料
   '   For nIndex = 1 To grdList.Rows - 1
   '      ' 判斷該列是否有被選取
   '      If grdList.TextMatrix(nIndex, 0) = "V" Then
   '         strDL01 = grdList.TextMatrix(nIndex, 7)
   '         strDL03 = DBDATE(grdList.TextMatrix(grdList.Row, 2))
   '         strDL04 = DBDATE(grdList.TextMatrix(grdList.Row, 3))
            
            'Modify By Cheng 2002/06/20
'            strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'                     "VALUES ('" & strDL01 & "'," & _
'                              DBDATE(textCP27) & "," & _
'                              strDL03 & "," & strDL04 & ")"
   '         strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
   '                  "VALUES ('" & strDL01 & "'," & _
   '                           DBDATE(textCP27) & "," & _
   '                           strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
   '         cnnConnection.Execute strSQL
   '      End If
   '   Next nIndex
   'Else
      ' 案件性質不為延期時, 總收文號, 本所期限及法定期限為該案本身
   '   strDL01 = m_CP09
   '   strDL03 = m_CP06
   '   strDL04 = m_CP07
   '   'Modify By Cheng 2002/06/20
'  '    strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04) " & _
'  '             "VALUES ('" & strDL01 & "'," & _
'  '                      DBDATE(textCP27) & "," & _
'  '                      strDL03 & "," & strDL04 & ")"
   '   strSQL = "INSERT INTO DateLimit (DL01, DL02, DL03, DL04, DL05) " & _
   '            "VALUES ('" & strDL01 & "'," & _
   '                     DBDATE(textCP27) & "," & _
   '                     strDL03 & "," & strDL04 & ",'" & IIf(m_CP10 = "303", "2", "1") & "')"
   '   cnnConnection.Execute strSQL
   'End If
   
   ' 更新案件進度檔
   'strSQL = "UPDATE CaseProgress SET "
   'bFirst = True
   'bDifference = False
   'For nIndex = 0 To m_CPCount - 1
   '   strTmp = Empty
   '   If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
   '      If m_CPList(nIndex).fiType = 0 Then
   '         ' 91.03.25 modify by louis (單引號)
   '         'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
   '         strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
   '      Else
   '         If m_CPList(nIndex).fiNewData = Empty Then
   '            strTmp = m_CPList(nIndex).fiName & " = " & 0
   '         Else
   '            strTmp = m_CPList(nIndex).fiName & " = " & m_CPList(nIndex).fiNewData
   '         End If
   '      End If
   '   End If
   '   If strTmp <> Empty Then
   '      bDifference = True
   '      If bFirst = True Then
   '         strSQL = strSQL & strTmp
   '         bFirst = False
   '      Else
   '         strSQL = strSQL & "," & strTmp
   '      End If
   '   End If
   'Next nIndex
   '' 設定SQL語法更新的條件
   'strSQL = strSQL & " " & _
   '               "WHERE CP09 = '" & m_CP09 & "' "
   '' 執行SQL指令
   'If bDifference = True Then: cnnConnection.Execute strSQL
  '
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String
   Dim strNP01 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP22 As String
   Dim strDL01 As String
   Dim strDL02 As String
   Dim strDL03 As String
   Dim strDL04 As String
   Dim strCP05 As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strCP27 As String
   Dim strCP30 As String 'Add by Morgan 2011/4/22
   Dim str303CP10 As String, str303CP09 As String 'Add By Sindy 2012/9/10
   Dim rsTmp As New ADODB.Recordset 'Add By Sindy 2012/9/10
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 更新商標基本檔
   OnUpdateTradeMark
      
   '911011 nick   因為下面 function 都已經update 了，所以不用執行
   ' 更新案件進度檔
   'OnUpdateCaseProperty
   
   ' 新增資料到延期記錄檔
   strDL01 = Empty
   strDL03 = Empty
   strDL04 = Empty
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
            
            ' 先刪除舊的資料
            strSql = "DELETE FROM DATELIMIT " & _
                     "WHERE DL01 = '" & strDL01 & "' AND " & _
                           "DL02 = " & DBDATE(textCP27) & " "
            cnnConnection.Execute strSql
            
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
   Else
      ' 案件性質不為延期時, 總收文號, 本所期限及法定期限為該案本身
      strDL01 = m_CP09
      strDL03 = m_CP06
      strDL04 = m_CP07
      'Add By Cheng 2002/08/19
      strNP22 = ""
      
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
   
   ' 案件性質為延期時
   If m_CP10 = "303" Then
      
      ' 更新原案件資料的發文日為延期日
      '911014 NICK 更新 CP64
      'strSQL = "UPDATE CaseProgress SET CP27 = " & DBDATE(textCP27) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      'modify by sonia 2016/10/7 CP22,CP110,CP113,CP118都沒存檔
      'strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(textCP27) & ",CP64='" & textCP64 & "' " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(textCP27) & ",CP64='" & textCP64 & "', " & _
               "CP22='" & Text7 & "',CP113='" & textCP113 & "',CP118='" & textCP118 & "',CP110='" & m_CP110 & "' " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
      
      ''''''''''''''''''''''''''''''''''''''''''''''''''''''
      
      ' 取得未收文資料的第一筆所選取的收文號
      strCP09 = Empty
      For nIndex = 1 To grdList.Rows - 1
         If grdList.TextMatrix(nIndex, 0) = "V" Then
            strCP09 = grdList.TextMatrix(nIndex, 7)
            'Add by Morgan 2011/4/22
            If grdList.TextMatrix(nIndex, 9) = "0" Then
               strCP30 = ""
            Else
               strCP30 = grdList.TextMatrix(nIndex, 9)
            End If
            'end 2011/4/22
            Exit For
         End If
      Next nIndex
      
      ' 若原收文資料的相關總收文號為空白時, 更新原收文資料的相關總收文號為所點選未收文資料的收文號
      'Modify by Morgan 2011/4/22 +CP30
      If IsEmptyText(strCP09) = False Then
         strSql = "UPDATE CaseProgress SET CP43 = '" & strCP09 & "',cp30='" & strCP30 & "' " & _
                  "WHERE CP09 = '" & m_CP09 & "' AND " & _
                        "(CP43 IS NULL OR CP43 = '')"
         cnnConnection.Execute strSql
      End If
      
      '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
      ' 更新使用者所選取的未收文期限資料
      'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
      Dim TheFirstV As String
      TheFirstV = ""
      For nIndex = 1 To grdList.Rows - 1
         grdList.row = nIndex
         ' 判斷該列是否有被選取
         'Modify By Sindy 2009/11/03
         'If grdList.Text = "V" Then
         If grdList.TextMatrix(nIndex, 0) = "V" Then
         '2009/11/03 End
            strNP01 = grdList.TextMatrix(grdList.row, 7)
            'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
            If TheFirstV = "" Then TheFirstV = strNP01
            strNP07 = grdList.TextMatrix(grdList.row, 8)
            strNP22 = grdList.TextMatrix(grdList.row, 9)
            'Modify by Morgan 2009/12/29 +更新CP
            If Val(strNP22) > 0 Then
               strSql = "UPDATE NextProgress SET NP08 = " & DBDATE(textCP06) & "," & _
                                             "NP09 = " & DBDATE(textCP07) & " " & _
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
            'Added by Lydia 2016/05/09 更新FCT之「申請」、「變更」、「移轉」、「授權」等案件性質之催審期限
            strExc(2) = ""
            strExc(1) = CompDate(1, 3, DBDATE(textCP27))
            If Left(strNP01, 1) >= "C" Then '欲延期期限為C類，以C類之CP43抓進度檔的CP09,ex:FCT-027981
               strSql = "select c2.cp09,c2.cp10 from caseprogress c1,caseprogress c2 where c1.cp09='" & strNP01 & "' and c1.cp43=c2.cp09(+) "
            Else
               '選取之欲延期期限為A或B類(資料來自於進度檔)時，以其CP43串進度檔若為C類，則再以C類之CP43再串進度檔，若其案件性質為變更301、移轉501、授權502、延展102案時，以延期的發文日+3個月更新變更、移轉、授權的催審期限；例：FCT-009891
               strSql = "select decode(substr(c1.cp43,1,1),'C',c3.cp09,c1.cp09) cp09,decode(substr(c1.cp43,1,1),'C',c3.cp10,c1.cp10) cp10 from caseprogress c1,caseprogress c2,caseprogress c3 where c1.cp09='" & strNP01 & "' and c1.cp43=c2.cp09(+) and c2.cp43=c3.cp09(+) "
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               'Modified by Lydia 2016/05/24 +申請案
               If InStr("101,102,301,501,502", "" & RsTemp.Fields("cp10")) > 0 And "" & RsTemp.Fields("cp09") <> "" Then
                  strExc(2) = RsTemp.Fields("cp09")
               End If
                'Added by Lydia 2016/06/08 申請案以發文日+6個月更新申請之催審期限
                If "" & RsTemp.Fields("cp10") = "101" And "" & RsTemp.Fields("cp09") <> "" Then
                   strExc(1) = CompDate(1, 6, DBDATE(textCP27))
                End If
            End If
            If strExc(2) <> "" Then
                'Modified by Lydia 2023/11/13 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天 +PUB_GetWorkDay1()
                'Modified by Lydia 2023/11/13 若依原設定規則更新之催審期限小於原催審期限，則不更新原催審期限。=>AND NP09<
                strSql = "UPDATE NEXTPROGRESS SET NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & _
                         " WHERE NP01='" & strExc(2) & "' AND NP07='305' AND NP06 IS NULL AND NP09 < " & strExc(1)
                cnnConnection.Execute strSql, intI
            End If
            'end 2016/05/09
         End If
      Next nIndex
      'edit by nick 2004/09/27 加入將本案期限選取的第一個收文號存入相關總收文號
      If TheFirstV <> "" Then
            strSql = " update caseprogress set cp43='" & TheFirstV & "' Where CP09='" & m_CP09 & "' "
            cnnConnection.Execute strSql
      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質非延期時時新增一筆資料到案件進度檔中
   ' 並更新原案件進度檔資料中的本所期限及法定期限
   If m_CP10 <> "303" Then
      ' 收文號
      strCP09 = Empty
      strCP09 = AutoNo("B", 6)
      
      'Add by Sindy 98/3/24 B類延期要補收文號
      If m_TM10 = "000" Then
         m_CP09s = strCP09 & m_CP09s
      End If
      
      ' 收文日
      strCP05 = DBDATE(SystemDate())
      ' 案件性質
      strCP10 = "303"
      ' 業務區別 91.8.26 MODIFY BY SONIA
      'strCP12 = GetStaffDepartment(m_CP13)
      ' 發文日
      strCP27 = DBDATE(textCP27)
      '911014 NICK 更新 CP64
      'strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
                       "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "') "
        'Modify By Cheng 2003/09/05
'      strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP64) " & _
'               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
'                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & m_CP13 & "','" & strUserNum & "'," & _
'                       "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','" & textCP64 & "') "
      'modify by sonia 2016/10/7 CP22,CP110,CP113,CP118都沒存檔
      'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP64) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                       "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','" & textCP64 & "') "
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP43,CP44,CP45,CP64,CP22,CP113,CP118,CP110) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strCP05 & "," & _
                       m_CP06 & "," & m_CP07 & ",'" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & strUserNum & "'," & _
                       "'" & "N" & "','" & "N" & "'," & strCP27 & ",'" & "N" & "','" & m_CP09 & "','" & m_CP44 & "','" & m_CP45 & "','" & textCP64 & "','" & Text7 & "','" & textCP113 & "','" & textCP118 & "','" & m_CP110 & "') "
      cnnConnection.Execute strSql
      
      'Add By Sindy 2014/10/31
      If m_CP10 = "201" And Val(textCP18) > 0 Then
         strSql = "UPDATE CASEPROGRESS SET CP16 = " & (Val(textCP18) * 1000) & ", " & _
                                          "CP18 = " & Val(textCP18) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
      '2014/10/31 END
      
      ' 更新原案件進度檔資料中的本所期限及法定期限
      strSql = "UPDATE CASEPROGRESS SET CP06 = " & DBDATE(textCP06) & ", " & _
                                       "CP07 = " & DBDATE(textCP07) & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
      
      'Added by Lydia 2016/05/09 更新FCT之「申請」、「變更」、「移轉」、「授權」等案件性質之催審期限
      strExc(2) = ""
      strExc(1) = CompDate(1, 3, DBDATE(textCP27))
      strSql = "select decode(substr(c1.cp43,1,1),'C',c3.cp09,c1.cp09) cp09,decode(substr(c1.cp43,1,1),'C',c3.cp10,c1.cp10) cp10 from caseprogress c1,caseprogress c2,caseprogress c3 where c1.cp09='" & m_CP09 & "' and c1.cp43=c2.cp09(+) and c2.cp43=c3.cp09(+)"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'Modified by Lydia 2016/05/24 +申請案
         If InStr("101,102,301,501,502", "" & RsTemp.Fields("cp10")) > 0 And "" & RsTemp.Fields("cp09") <> "" Then
            strExc(2) = RsTemp.Fields("cp09")
         End If
      End If
      If strExc(2) <> "" Then
          'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天 +PUB_GetWorkDay1()
          'Modified by Lydia 2023/11/13 若依原設定規則更新之催審期限小於原催審期限，則不更新原催審期限。=>AND NP09<
          strSql = "UPDATE NEXTPROGRESS SET NP08=" & PUB_GetWorkDay1(strExc(1), True) & ",NP09=" & strExc(1) & _
                 " WHERE NP01='" & strExc(2) & "' AND NP07='305' AND NP06 IS NULL AND NP09 < " & strExc(1)
          cnnConnection.Execute strSql, intI
      End If
      'end 2016/05/09
   End If
   
   'Add By Sindy 2012/9/10
   '按確定按鈕進入
   If m_CP10 = "303" Then
      str303CP10 = strNP07 '點未收文期限的那一筆案件性質
      str303CP09 = m_CP09
   '按延期按鈕進入
   Else
      str303CP10 = m_CP10
      str303CP09 = strCP09 'B類收文
   End If
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & m_TM01 & "' AND " & _
                  "CF02 = '" & m_TM10 & "' AND " & _
                  "CF03 = '" & str303CP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      'Modify By Sindy 2013/3/25 內外商的延期發文取消掛催審期限
'      If IsNull(rsTmp.Fields("CF05")) = False Then
'         strNP07 = "305"
'         strNP08 = GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27)
'         strNP22 = GetNextProgressNo()
'         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'                  "VALUES ('" & str303CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
'                            strNP08 & "," & strNP08 & ",'" & m_CP14 & "'," & strNP22 & ")"
'         cnnConnection.Execute strSql
'      End If
   End If
   rsTmp.Close
   '2012/9/10 End
   
   'add by nick 2004/08/13 更新實際發文規費
   If textCP84.Enabled = True Then
        'edit by nick 2004/08/31
        If m_CP10 = "303" Then
            'modify by sonia 2023/11/20 同時更新為要請款CP20=NULL
            strSql = "Update CaseProgress Set CP20=NULL,CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
        Else
            'modify by sonia 2023/11/20 同時更新為要請款CP20=NULL
            strSql = "Update CaseProgress Set CP20=NULL,CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
        End If
   End If
   
   'Add By Sindy 2012/12/20 若為電子送件則自動設定為不經發文室
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
    
   'Add by Sindy 98/3/24
   If m_TM10 = "000" Then
      'Modify By Sindy 2009/04/24
      'PUB_UpdateDispatch m_CP09s, m_CP123s
      PUB_UpdateDispatch m_CP09s, m_CP123s, m_CP130s
   End If
   
   'Add By Sindy 2010/7/8 檢查商品資料與基本檔商品類別是否一致
   Call CheckTMGoodsErr(m_TM01, m_TM02, m_TM03, m_TM04, False, True, m_CP14)
   
   'Added by Lydia 2023/03/17 若延期的相關總收文號為B2案源時，同時新增法務案之內部收文39延期
   If m_LOS15 <> "" And m_LOS02 = "B2" Then
       Call PUB_InsertLosBCP(m_LOS15, DBDATE(textCP27), DBDATE(textCP06), DBDATE(textCP07))
   End If
   'end 2023/03/17
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44no, m_CP116
    PUB_CheckEMail m_TM44, m_TM119
    If m_TM120 <> "" Then
       PUB_CheckEMail m_TM44, m_TM120
    End If
    'end 2008/02/22
  
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     OnSaveData = False
End Function

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
    'Add By Cheng 2002/11/29
    Dim bFind As Boolean
    Dim nIndex  As Integer
   
   CheckDataValid = False
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 本所期限的日期不可超過法定期限的日期
   If Val(textCP06) > Val(textCP07) Then
      strTit = "資料檢核"
      strMsg = "本所期限的日期不可超過法定期限的日期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
    'Add By Cheng 2002/11/29
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
   '外商(S)申請人1或FC代理人至少要輸入一個
   '其他的一定要輸入申請人1
   If m_TM01 = "S" Then
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
   
   'Add By Sindy 2014/10/31
   If m_CP10 = "201" And Val(textCP18) = 0 Then
      If MsgBox("補正已收文之延期但未輸延期點數，是否要請延期點數？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         textCP18.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2014/10/31 END
   
   'Added by Lydia 2021/09/09 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       GoTo EXITSUB
   End If

   CheckDataValid = True
EXITSUB:
End Function

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub

Private Sub textPetition_GotFocus()
   InverseTextBox textPetition
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textTM47_GotFocus()
   InverseTextBox textTM47
End Sub

Private Sub textTM48_GotFocus()
   InverseTextBox textTM48
End Sub

Private Sub textTM49_GotFocus()
   InverseTextBox textTM49
End Sub

Private Sub textTM50_GotFocus()
   InverseTextBox textTM50
End Sub

Private Sub textTM51_GotFocus()
   InverseTextBox textTM51
End Sub

Private Sub textTM52_GotFocus()
   InverseTextBox textTM52
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False
'add by nick 2004/08/13 發文規費，申請國家台灣才檢查
If Me.textCP84.Enabled = True Then
   Cancel = False
   textCP84_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
If textCP84.Enabled = True And m_TM10 = "000" Then
    If Val(textCP84.Text) <> Val(m_CP84) Then
        MsgBox "發文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", , "警告！"
        textCP84_GotFocus
        Exit Function
    End If
End If

'add by nickc 2005/07/29
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
If Me.textCP27.Enabled = True Then
   Cancel = False
   If IsEmptyText(textCP27) = False Then
      ' 發文日日期不正確
      If CheckIsTaiwanDate(textCP27, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的發文日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         Exit Function
      End If
      
      ' 發文日日期不可超過系統日
      'edit by nickc 2007/12/17
      'If Val(DBDATE(textCP27)) > Val(DBDATE(SystemDate())) Then
      If Val(DBDATE(textCP27)) > Val(DBDATE(PUB_GetWorkDay(2))) Then
         Cancel = True
         strTit = "資料檢核"
         'strMsg = "發文日不可超過系統日"
         strMsg = "發文日不可超過系統日加一天"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP27_GotFocus
         Exit Function
      End If
      
   End If
End If

'Add By Sindy 98/03/11
If Me.textCP113.Enabled = True Then
   Cancel = False
   textCP113_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'98/03/11 End

If Me.textDN.Enabled = True Then
   Cancel = False
   textDN_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textPetition.Enabled = True Then
   Cancel = False
   textPetition_Validate Cancel
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

'add by nickc 2006/01/27
'edit by nickc 2006/02/07
If m_TM01 = "FCT" Then
    If Me.lstNameAgent.Enabled = True Then
        Cancel = False
        lstNameAgent_Validate Cancel
        If Cancel = True Then
            lstNameAgent.SetFocus
            Exit Function
        End If
    End If
End If

TxtValidate = True
End Function

'add by nick 2004/08/13
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

'Add By Sindy 98/03/11
Private Sub textCP113_GotFocus()
   TextInverse textCP113
End Sub
Private Sub textCP113_Validate(Cancel As Boolean)
   If textCP113 <> "" Then
      If Not IsNumeric(textCP113) Then
         MsgBox "請輸入數字！", vbExclamation
         textCP113.SetFocus
         textCP113_GotFocus
         Cancel = True
         Exit Sub
      End If
   End If
   If GetPrjNation1(textTMKey) = "000" Then
      Cancel = Not PUB_CheckCP113(textCP113, m_TM01, m_CP10, m_CP14)
   End If
End Sub
'98/03/11 End

'Add By Sindy 2012/12/20
Private Sub textCP118_GotFocus()
   TextInverse textCP118
   CloseIme
End Sub

'Add By Sindy 2012/12/20
Private Sub textCP118_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
