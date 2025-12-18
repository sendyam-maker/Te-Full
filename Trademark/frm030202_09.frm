VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030202_09 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文(授權, 再授權, 終止授權, 終止再授權,徵求同意書)"
   ClientHeight    =   5760
   ClientLeft      =   5090
   ClientTop       =   2420
   ClientWidth     =   9150
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9150
   Begin VB.CommandButton cmdOK 
      Caption         =   "同時發文(&N)"
      Height          =   405
      Index           =   1
      Left            =   3510
      TabIndex        =   17
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   405
      Left            =   4710
      TabIndex        =   18
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   405
      Left            =   6870
      TabIndex        =   20
      Top             =   0
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   5910
      TabIndex        =   19
      Top             =   0
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   8070
      TabIndex        =   21
      Top             =   0
      Width           =   912
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1050
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1359
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1350
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   750
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   1056
      Width           =   2532
   End
   Begin VB.TextBox textTM20 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   753
      Width           =   2532
   End
   Begin VB.TextBox textCP12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   450
      Width           =   2532
   End
   Begin VB.TextBox textTM08 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H80000007&
      Height          =   285
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   2571
      Width           =   2532
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2540
      Left            =   120
      TabIndex        =   45
      Top             =   3210
      Width           =   8900
      _ExtentX        =   15699
      _ExtentY        =   4480
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   5
      TabHeight       =   420
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm030202_09.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label11"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label37"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label36"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label28"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label14"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label22"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label23"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(10)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label39"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblNameAgent"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label55"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label43"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "textCP64"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lstNameAgent"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "textAdd"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP18"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textDN"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP27"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textUargeDate"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textPrint"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP84"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Text7"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP113"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP118"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "授權資料"
      TabPicture(1)   =   "frm030202_09.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP53"
      Tab(1).Control(1)=   "textCP54"
      Tab(1).Control(2)=   "textCP72"
      Tab(1).Control(3)=   "textCP52"
      Tab(1).Control(4)=   "textCP51"
      Tab(1).Control(5)=   "textCP50"
      Tab(1).Control(6)=   "Label7"
      Tab(1).Control(7)=   "Label8"
      Tab(1).Control(8)=   "Label4"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(11)=   "Line1"
      Tab(1).ControlCount=   12
      Begin VB.TextBox textCP118 
         Height          =   270
         Left            =   3900
         MaxLength       =   1
         TabIndex        =   7
         Top             =   1260
         Width           =   375
      End
      Begin VB.TextBox textCP113 
         Height          =   270
         Left            =   6045
         MaxLength       =   4
         TabIndex        =   2
         Top             =   330
         Width           =   600
      End
      Begin VB.TextBox Text7 
         Height          =   288
         Left            =   5805
         MaxLength       =   1
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1275
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox textCP84 
         Alignment       =   1  '靠右對齊
         Height          =   285
         Left            =   3750
         TabIndex        =   1
         Top             =   330
         Width           =   1425
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1260
         Width           =   372
      End
      Begin VB.TextBox textUargeDate 
         Height          =   264
         Left            =   7620
         MaxLength       =   7
         TabIndex        =   3
         Top             =   330
         Width           =   1092
      End
      Begin VB.TextBox textCP27 
         Height          =   264
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   0
         Top             =   360
         Width           =   1092
      End
      Begin VB.TextBox textDN 
         Height          =   264
         Left            =   1440
         MaxLength       =   1
         TabIndex        =   4
         Top             =   660
         Width           =   492
      End
      Begin VB.TextBox textCP18 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   5520
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   660
         Width           =   2172
      End
      Begin VB.TextBox textAdd 
         Height          =   264
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   5
         Top             =   960
         Width           =   852
      End
      Begin VB.TextBox textCP53 
         Height          =   285
         Left            =   -73560
         MaxLength       =   7
         TabIndex        =   15
         Top             =   1590
         Width           =   852
      End
      Begin VB.TextBox textCP54 
         Height          =   285
         Left            =   -72390
         MaxLength       =   7
         TabIndex        =   16
         Top             =   1590
         Width           =   852
      End
      Begin VB.TextBox textCP72 
         Height          =   285
         Left            =   -73560
         MaxLength       =   9
         TabIndex        =   11
         Top             =   360
         Width           =   972
      End
      Begin MSForms.TextBox textCP52 
         Height          =   285
         Left            =   -73560
         TabIndex        =   14
         Top             =   1281
         Width           =   7215
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP51 
         Height          =   285
         Left            =   -73560
         TabIndex        =   13
         Top             =   974
         Width           =   7215
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP50 
         Height          =   285
         Left            =   -73560
         TabIndex        =   12
         Top             =   667
         Width           =   7215
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12726;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox lstNameAgent 
         Height          =   315
         Left            =   7200
         TabIndex        =   9
         Top             =   1170
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
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   990
         TabIndex        =   10
         Top             =   1560
         Width           =   6015
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "10610;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "是否電子送件:          (Y: 是)"
         Height          =   180
         Left            =   2730
         TabIndex        =   70
         Top             =   1320
         Width           =   2085
      End
      Begin VB.Label Label55 
         Caption         =   $"frm030202_09.frx":0038
         ForeColor       =   &H000000C0&
         Height          =   410
         Left            =   120
         TabIndex        =   69
         Top             =   2100
         Width           =   4070
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工作時數:"
         Height          =   180
         Index           =   12
         Left            =   5250
         TabIndex        =   68
         Top             =   390
         Width           =   765
      End
      Begin VB.Label lblNameAgent 
         AutoSize        =   -1  'True
         Caption         =   "出名代理人"
         Height          =   180
         Left            =   6120
         TabIndex        =   63
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "發文規費："
         Height          =   180
         Left            =   2820
         TabIndex        =   62
         Top             =   390
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "點　　數 :"
         Height          =   252
         Index           =   10
         Left            =   4560
         TabIndex        =   61
         Top             =   660
         Width           =   900
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   255
         Left            =   1590
         TabIndex        =   60
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   59
         Top             =   1260
         Width           =   972
      End
      Begin VB.Label Label14 
         Caption         =   "催審期限 :"
         Height          =   255
         Left            =   6750
         TabIndex        =   58
         Top             =   390
         Width           =   885
      End
      Begin VB.Label Label25 
         Caption         =   "發文日 :"
         Height          =   252
         Left            =   120
         TabIndex        =   57
         Top             =   360
         Width           =   852
      End
      Begin VB.Label Label28 
         Caption         =   "進度備註 :"
         Height          =   252
         Left            =   120
         TabIndex        =   56
         Top             =   1560
         Width           =   972
      End
      Begin VB.Label Label36 
         Caption         =   "是否輸入D/N :"
         Height          =   252
         Left            =   120
         TabIndex        =   55
         Top             =   660
         Width           =   1212
      End
      Begin VB.Label Label37 
         Caption         =   "(Y:輸入)"
         Height          =   252
         Left            =   2040
         TabIndex        =   54
         Top             =   660
         Width           =   852
      End
      Begin VB.Label Label11 
         Caption         =   "(1:授權人委任狀 2:被授權人委任狀 3:授權合約 4:註冊證)"
         Height          =   252
         Left            =   2640
         TabIndex        =   53
         Top             =   960
         Width           =   4812
      End
      Begin VB.Label Label17 
         Caption         =   "是否補件(可複選) :"
         Height          =   252
         Left            =   120
         TabIndex        =   52
         Top             =   960
         Width           =   1692
      End
      Begin VB.Label Label7 
         Caption         =   "被授權人代號 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   50
         Top             =   376
         Width           =   1212
      End
      Begin VB.Label Label8 
         Caption         =   "被授權人(中) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   49
         Top             =   683
         Width           =   1212
      End
      Begin VB.Label Label4 
         Caption         =   "被授權人(英) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   48
         Top             =   990
         Width           =   1212
      End
      Begin VB.Label Label10 
         Caption         =   "被授權人(日) :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   47
         Top             =   1297
         Width           =   1212
      End
      Begin VB.Label Label12 
         Caption         =   "授權期間 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   46
         Top             =   1606
         Width           =   972
      End
      Begin VB.Line Line1 
         X1              =   -72600
         X2              =   -72480
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   5610
      TabIndex        =   79
      Top             =   2880
      Width           =   3375
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "5953;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM44 
      Height          =   285
      Left            =   1170
      TabIndex        =   78
      Top             =   2850
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM79 
      Height          =   285
      Left            =   1170
      TabIndex        =   77
      Top             =   2250
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
      Left            =   5610
      TabIndex        =   76
      Top             =   2268
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
      Left            =   1170
      TabIndex        =   75
      Top             =   2550
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
      Left            =   1170
      TabIndex        =   74
      Top             =   1950
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
      Left            =   5610
      TabIndex        =   73
      Top             =   1965
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
      Left            =   1170
      TabIndex        =   72
      TabStop         =   0   'False
      Top             =   1650
      Width           =   2532
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5610
      TabIndex        =   71
      Top             =   1662
      Width           =   2532
      VariousPropertyBits=   671105055
      Size            =   "4466;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label18 
      Caption         =   "申請人5 :"
      Height          =   285
      Left            =   90
      TabIndex        =   67
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label Label16 
      Caption         =   "申請人4 :"
      Height          =   285
      Left            =   4650
      TabIndex        =   66
      Top             =   2250
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "申請人3 :"
      Height          =   285
      Left            =   90
      TabIndex        =   65
      Top             =   2250
      Width           =   855
   End
   Begin VB.Label Label13 
      Caption         =   "申請人2 :"
      Height          =   285
      Left            =   4650
      TabIndex        =   64
      Top             =   1950
      Width           =   855
   End
   Begin VB.Label Label5 
      Caption         =   "案件名稱 :"
      Height          =   285
      Left            =   4650
      TabIndex        =   44
      Top             =   2850
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   285
      Left            =   90
      TabIndex        =   43
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   285
      Index           =   11
      Left            =   4650
      TabIndex        =   42
      Top             =   1650
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   285
      Index           =   9
      Left            =   4650
      TabIndex        =   41
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   285
      Index           =   6
      Left            =   90
      TabIndex        =   40
      Top             =   1350
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   39
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   38
      Top             =   750
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號 :"
      Height          =   285
      Left            =   4650
      TabIndex        =   37
      Top             =   1050
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "發證日 :"
      Height          =   285
      Index           =   3
      Left            =   4650
      TabIndex        =   36
      Top             =   750
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "業務區別 :"
      Height          =   285
      Index           =   2
      Left            =   4650
      TabIndex        =   35
      Top             =   450
      Width           =   975
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   285
      Left            =   90
      TabIndex        =   34
      Top             =   1650
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商標種類 :"
      Height          =   285
      Index           =   4
      Left            =   4650
      TabIndex        =   33
      Top             =   2550
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "申請人1 :"
      Height          =   285
      Left            =   90
      TabIndex        =   32
      Top             =   1950
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "代理人 :"
      Height          =   285
      Left            =   90
      TabIndex        =   31
      Top             =   2850
      Width           =   975
   End
End
Attribute VB_Name = "frm030202_09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/07 改成Form2.0 ;cmbTM05、textCP13、textCP14、textCP64、textTM44、textTM23、textTM78~81、textCP50~52、lstNameAgent
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
' 收文號
Dim m_CP09 As String
' 申請國家
Dim m_TM10 As String
' 案件性質代號
Dim m_CP10 As String
' 原專用期間
Dim m_TM21 As String
Dim m_TM22 As String
'承辦人 Add By Sindy 98/03/11
Dim m_CP14 As String
Dim m_CP82 As String 'Added by Lydia 2018/08/10 發文時間

Dim m_textUargeDate As String   '2009/10/14 add by sonia

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
'add by nick 2004/08/13
Dim m_CP84 As String       '發文規費
'add by nickc 2006/01/26
Dim m_CP110 As String
'add by nickc 2008/02/22
Dim m_CP44 As String
Dim m_CP116 As String
Dim m_TM44 As String
Dim m_TM119 As String
Dim m_TM120 As String
Dim m_CP09s As String, m_CP123s As String 'Add by Sindy 98/3/24 收文號,是否算發文室案件
Dim m_CP130s As String 'Add by Sindy 2009/4/24 發文-主管機關
Dim m_IsSend As Boolean 'Add By Sindy 2012/8/10 是否經發文室發文


Private Sub cmdCancel_Click()
   frm030202_01.Show
   Unload Me
End Sub

Private Sub cmdExit_Click()
   Unload frm030202_01
   Unload Me
   'frm030202_01.Show
End Sub

Private Sub cmdok_Click(Index As Integer)
   'Modify By Sindy 2010/11/19 把「確定」及「同時發文」按鈕程式碼合併
   Select Case Index
      Case 0, 1
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
   '                  'Add By Sindy 2012/8/6 因有一文多案的問題，所以若經發文室且作業畫面上為列印定稿時,詢問使用者
   '                  If m_CP123s = "Y" Or (m_CP123s = "N" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "3") Then
   '                     m_IsSend = True
   '                  Else
   '                     m_IsSend = False
   '                  End If
   '                  If m_CP123s = "Y" And Trim(textPrint.Text) = "" And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then
   '                     If MsgBox("是否需要列印定稿？", vbExclamation + vbYesNo) = vbNo Then
   '                        textPrint.Text = "N"
   '                     End If
   '                  End If
   '                  '2012/8/6 End
                     If m_CP123s = "Y" Then
                        'modify by sonia 2014/6/23 加傳發文規費, P-108903
                        If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP84, textCP27) = False Then
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
            'Add By Sindy 2012/8/6
'            '判斷是列印定稿
'            If Me.textPrint.Text <> "N" Then
'               If m_IsSend = True Then 'Modify By Sindy 2012/8/3 +if 因有一文多案狀況，所以增加判斷經發文室時才需出定稿
'                  PrintLetter
'               End If
'            End If
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
'            Else
'               'Add By Cheng 2002/04/30
'               '若有未發文資料顯示警告
'               PUB_GetCPunIssueDatas "" & Me.textTMKey.Text
'
'               frm030202_01.Show
'               ' 90.12.07 modify by louis
'               frm030202_01.Clear1
            End If
            'Ken 91.04.09 -- End
            
            Call PUB_FCTSendRecvMail(m_CP09) 'Add By Sindy 2024/10/30 外商發文時,增加發Mail通知承辦人及副本給判發主管
            'Add By Sindy 2024/8/19
            If frm030202_01.bolIsEMPFlow = True Then
               frm090202_4.QueryData
            End If
            '2024/8/19 End
            If Index = 0 Then '確定鍵
               'Ken 91.04.09 -- Start
               If textDN <> "Y" Then
                  'Add By Cheng 2002/04/30
                  '若有未發文資料顯示警告
                  If PUB_GetCPunIssueDatas("" & Me.textTMKey.Text) = True Then
                     frm030202_01.Show
                     ' 90.12.07 modify by louis
                     frm030202_01.Clear1
                  Else
                     'Add By Sindy 2024/8/19
                     If frm030202_01.bolIsEMPFlow = True Then
                        Unload frm030202_01
                        frm090202_4.Show
                     Else
                     '2024/8/19 End
                        frm030202_01.Show
                        frm030202_01.Clear1
                     End If
                  End If
               End If
               'Ken 91.04.09 -- End
               Unload Me
            ElseIf Index = 1 Then '同時發文鍵
               If textDN <> "Y" Then
                  ' 呼叫第一個畫面
                  frm030202_01.SetData 0, m_TM01, True
                  frm030202_01.SetData 1, m_TM02, False
                  frm030202_01.SetData 2, m_TM03, False
                  frm030202_01.SetData 3, m_TM04, False
                  frm030202_01.SetQueryFromTM
                  Unload Me
                  frm030202_01.Show
                  frm030202_01.radio(1).Value = True
                  frm030202_01.radio_Click 1
                  frm030202_01.QueryData
               Else
                  Unload Me
               End If
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
'      'Add By Cheng 2002/07/15
'      '重新檢查欄位有效性
'      If TxtValidate = False Then Exit Sub
'
'      'Add by Sindy 98/3/24 設定是否算發文室案件
'      If m_TM10 = "000" Then
'         'Add by Sindy 2009/4/24
'         If ModifyDispatchCp130(textCP09, m_CP09s, m_CP123s, m_CP130s, textCP27) = False Then
'            Exit Sub
'         Else
'            If m_CP123s = "Y" Then
'               If ModifyDispatch(textCP09, m_CP09s, m_CP123s, textCP27) = False Then
'                   Exit Sub
'               End If
'            End If
'         End If
'      End If
'
'      ' 設定滑鼠游標為等待狀態
'      Screen.MousePointer = vbHourglass
'      ' 更新欄位輸入的內容
'      OnUpdateField
'      ' 存檔
'      'edit by nick 2004/11/03
'      'OnSaveData
'      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
'
'      ' 設定滑鼠游標為預設
'      Screen.MousePointer = vbDefault
'
'      ' 呼叫第一個畫面
'      frm030202_01.SetData 0, m_TM01, True
'      frm030202_01.SetData 1, m_TM02, False
'      frm030202_01.SetData 2, m_TM03, False
'      frm030202_01.SetData 3, m_TM04, False
'      frm030202_01.SetQueryFromTM
'      Unload Me
'      frm030202_01.Show
'      frm030202_01.radio(1).Value = True
'      frm030202_01.radio_Click 1
'      frm030202_01.QueryData
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
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM08.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM20.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   'add by nickc 2007/01/29
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
   
   MoveFormToCenter Me
   'Add by nickc 2006/01/26
   '台灣加出名代理人清單供勾選,原是否出名欄位不顯示
   Text7.Visible = False
   lstNameAgent.Clear
   lstNameAgent.Visible = True
   lblNameAgent.Visible = True
   'Added by Lydia 2021/09/07 如果一開始將ListBox拉到需要的大小，字型會自動放大；所以畫面預設為一列高度，Form_Load才放大到需要的大小
   lstNameAgent.Height = 930
   lstNameAgent.Width = 1500
   Me.SSTab1.Tab = 0
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
      textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then: m_TM10 = rsTmp.Fields("TM10")
      ' 專用期間
      If IsNull(rsTmp.Fields("TM21")) = False Then
         m_TM21 = rsTmp.Fields("TM21")
      End If
      If IsNull(rsTmp.Fields("TM22")) = False Then
         m_TM22 = rsTmp.Fields("TM22")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then: textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      'add by nickc 2007/01/29
      If IsNull(rsTmp.Fields("TM78")) = False Then: textTM78 = GetCustomerName("" & rsTmp.Fields("TM78"), 0)
      If IsNull(rsTmp.Fields("TM79")) = False Then: textTM79 = GetCustomerName("" & rsTmp.Fields("TM79"), 0)
      If IsNull(rsTmp.Fields("TM80")) = False Then: textTM80 = GetCustomerName("" & rsTmp.Fields("TM80"), 0)
      If IsNull(rsTmp.Fields("TM81")) = False Then: textTM81 = GetCustomerName("" & rsTmp.Fields("TM81"), 0)
      
      ' FC代理人
      If IsNull(rsTmp.Fields("TM44")) = False Then: textTM44 = GetFAgentName(rsTmp.Fields("TM44"))
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then: textTM45 = rsTmp.Fields("TM45")
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
   Dim strTemp As String
   Dim strDate As String
   Dim strCP27 As String
   Dim strCP44 As String
   Dim strCP45 As String
   Dim nIndex As Integer
   Dim bFind As Boolean
   
   ' 系統日
   strDate = DBDATE(SystemDate())
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
      m_CP44 = CheckStr(rsTmp.Fields("CP44"))
      m_CP82 = "" & rsTmp.Fields("CP82")  'Added by Lydia 2018/08/10 發文時間
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區別
      '910718 Sieg
      If IsNull(rsTmp.Fields("CP12")) = False Then
         textCP12 = GetDepartmentName(rsTmp.Fields("CP12"))
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then: textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      
      'Add By Sindy 98/03/11
      '工作時數
      textCP113 = "" & rsTmp.Fields("CP113")
      SetCPFieldOldData "CP113", textCP113, 1
      '承辦人
      m_CP14 = "" & rsTmp.Fields("CP14")
      '98/03/11 End
      
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then: textCP14 = GetStaffName(rsTmp.Fields("CP14"))
      ' 發文日(預設為系統日)
      strCP27 = Empty
      textCP27 = TAIWANDATE(strDate)
      If IsNull(rsTmp.Fields("CP27")) = False Then: strCP27 = rsTmp.Fields("CP27")
      SetCPFieldOldData "CP27", strCP27, 1
      ' 點數
      textCP18 = Empty
      If IsNull(rsTmp.Fields("CP18")) = False Then: textCP18 = rsTmp.Fields("CP18")
      ' 被授權人
      textCP72 = Empty
      If IsNull(rsTmp.Fields("CP72")) = False Then: textCP72 = rsTmp.Fields("CP72")
      SetCPFieldOldData "CP72", textCP72, 0
      ' 被授權人(中)
      textCP50 = Empty
      If IsNull(rsTmp.Fields("CP50")) = False Then: textCP50 = rsTmp.Fields("CP50")
      SetCPFieldOldData "CP50", textCP50, 0
      ' 被授權人(英)
      textCP51 = Empty
      If IsNull(rsTmp.Fields("CP51")) = False Then: textCP51 = rsTmp.Fields("CP51")
      SetCPFieldOldData "CP51", textCP51, 0
      ' 被授權人(日)
      textCP52 = Empty
      If IsNull(rsTmp.Fields("CP52")) = False Then: textCP52 = rsTmp.Fields("CP52")
      SetCPFieldOldData "CP52", textCP52, 0
      ' 授權期間(起)
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP53")) = False Then
         strTemp = rsTmp.Fields("CP53")
         textCP53 = TAIWANDATE(textCP53)
      End If
      SetCPFieldOldData "CP53", strTemp, 1
      ' 授權期間(迄)
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP54")) = False Then
         strTemp = rsTmp.Fields("CP54")
         textCP54 = TAIWANDATE(strTemp)
      End If
      SetCPFieldOldData "CP54", strTemp, 1
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
   'Modified by Lydia 2021/09/07 + Form 2.0 = True
   PUB_SetOurAgent lstNameAgent, tm(), m_CP110, m_CP10, True 'Modify By Sindy 2010/9/20
   
   ' 取得催審期限的日期
   textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
   Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
   
   '2009/10/14 ADD BY SONIA
   If m_CP10 = "724" Then
      textPrint = "N"
      '2010/12/9 ADD BY SONIA FCT-029769
      Label8.Caption = "徵求對象(中)"
      Label4.Caption = "徵求對象(英)"
      Label10.Caption = "徵求對象(日)"
      textUargeDate = TAIWANDATE(CompDate(1, 1, textCP27)) 'Add By Sindy 2013/12/30
   Else
      Label8.Caption = "被授權人(中)"
      Label4.Caption = "被授權人(英)"
      Label10.Caption = "被授權人(日)"
      '2010/12/9 end
   End If
   m_textUargeDate = textUargeDate 'Add By Sindy 2013/12/30
   
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

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/19
   Set frm030202_09 = Nothing
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
         'Modified by Lydia 2021/09/07 改模組
         'm_CP110 = m_CP110 & "," & PUB_Num2Id(lstNameAgent.ItemData(ii))
         m_CP110 = m_CP110 & "," & PUB_GetItemData(lstNameAgent.Tag, ii)
         bolCheck = True
      End If
   Next
   If Left(m_CP110, 1) = "," Then m_CP110 = Mid(m_CP110, 2)
   If bolCheck = True Then
      Text7 = ""
   Else
      'Add By Sindy 2019/5/8
      '724.徵求同意書;FCT「徵求同意書」發文時不須勾註「出名代理人」請取消設定 例:FCT-41917
      If m_CP10 <> "724" Then
      '2019/5/8 END
         Text7 = "N"
         MsgBox "未勾選代理人!", vbInformation, "必要欄位！"
         Cancel = True
      End If
   End If
End Sub
' 是否補件
Private Sub textAdd_Validate(Cancel As Boolean)
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nCount As Integer
   Dim nIndex As Integer
   Cancel = False
   
   ' 無資料時不做任何檢查
   If IsEmptyText(textAdd) = True Then
      GoTo EXITSUB
   End If
   
   nCount = GetSubStringCount(textAdd)
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      Select Case strTemp
         Case "1", "2", "3", "4":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "是否補件項目<" & strTemp & ">不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textAdd_GotFocus
            GoTo EXITSUB
      End Select
   Next nIndex
   
   For nIndex = 1 To nCount
      strTemp = GetSubString(textAdd, nIndex)
      For nCount = 1 To nCount
         If nIndex <> nCount Then
            If strTemp = GetSubString(textAdd, nCount) Then
               Cancel = True
               strTit = "檢核資料"
               strMsg = "是否補件項目<" & strTemp & ">不可重覆"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textAdd_GotFocus
               GoTo EXITSUB
            End If
         End If
      Next nCount
   Next nIndex
   
EXITSUB:
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
      
      ' 取得催審期限的日期
      '2009/10/14 add by sonia
      'Modified by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      'If textUargeDate = m_textUargeDate Then
      If Me.textCP27.Tag <> Me.textCP27.Text Then
         'Add By Sindy 2013/12/30
         If m_CP10 = "724" Then
            textUargeDate = TAIWANDATE(CompDate(1, 1, textCP27))
         Else
         '2013/12/30 END
            textUargeDate = TAIWANDATE(GetUrgeDate(m_TM01, m_TM10, m_CP10, textCP27))
         End If
         m_textUargeDate = textUargeDate
      End If
      Me.textCP27.Tag = Me.textCP27.Text 'Added by Lydia 2019/11/08 記錄發文日，有修改發文日再重新計算催審期限
      
   End If
EXITSUB:
End Sub

' 授權期間起日
Private Sub textCP53_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textCP53) = False Then
      If CheckIsTaiwanDate(textCP53, False) = False Then
         strTit = "資料檢核"
         strMsg = "授權期間起日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP53_GotFocus
         GoTo EXITSUB
      End If
      
      'Modify By Cheng 2002/06/14
      '改成不與商標及服務基本檔的專用期限Check
'      If IsEmptyText(m_TM21) = False Then
'         If Val(DBDATE(textCP53)) < Val(DBDATE(m_TM21)) Then
'            strTit = "資料檢核"
'            strMsg = "授權期間起日不可小於專用期間起日"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP53_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
   End If
EXITSUB:
End Sub

' 授權期間止日
Private Sub textCP54_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textCP54) = False Then
      If CheckIsTaiwanDate(textCP54, False) = False Then
         strTit = "資料檢核"
         strMsg = "授權期間止日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP54_GotFocus
         GoTo EXITSUB
      End If

      'Modify By Cheng 2002/06/14
      '改成不與商標及服務基本檔的專用期限Check
'      If IsEmptyText(m_TM22) = False Then
'         If Val(DBDATE(textCP54)) > Val(DBDATE(m_TM22)) Then
'            strTit = "資料檢核"
'            strMsg = "授權期間止日不可超過專用期間止日"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP54_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
   End If
EXITSUB:
End Sub

'edit by nickc 2006/01/27
'Private Sub textCP64_2_GotFocus()
'   TextInverse textCP64_2
'End Sub

' 被授權人
Private Sub textCP72_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   Cancel = False
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   If IsEmptyText(textCP72) = False Then
      textCP50 = Empty
      textCP51 = Empty
      textCP52 = Empty
      Set rsTmp = New ADODB.Recordset
      strTemp = textCP72
        'Modify By Cheng 2003/03/03
        '被授權人代號先取8碼
'      strTemp = strTemp & String(8 - Len(strTemp), "0")
      strTemp = Left(strTemp & "00000000", 8)
      strSql = "SELECT * FROM CUSTOMER " & _
               "WHERE CU01 = '" & strTemp & "' AND " & _
                     "CU02 = '0'"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         ' 中文名稱
         If IsNull(rsTmp.Fields("CU04")) = False Then
            textCP50 = rsTmp.Fields("CU04")
         End If
         ' 英文名稱
         If IsNull(rsTmp.Fields("CU05")) = False Then
            textCP51 = rsTmp.Fields("CU05")
            '92.5.30 add by sonia
            If IsNull(rsTmp.Fields("CU88")) = False Then
               textCP51 = textCP51 + " " + rsTmp.Fields("CU88")
            End If
            If IsNull(rsTmp.Fields("CU89")) = False Then
               textCP51 = textCP51 + " " + rsTmp.Fields("CU89")
            End If
            If IsNull(rsTmp.Fields("CU90")) = False Then
               textCP51 = textCP51 + " " + rsTmp.Fields("CU90")
            End If
            '92.5.30 end
         End If
         ' 日文名稱
         If IsNull(rsTmp.Fields("CU06")) = False Then
            textCP52 = rsTmp.Fields("CU06")
         End If
         If CheckStr(rsTmp.Fields("CU80").Value) = "不再使用" Then
                MsgBox "此授權人資料已不再使用，請確認！！", , MsgText(5)
                Cancel = True
                Exit Sub
         End If
      Else
         Cancel = True
         strTit = "資料檢核"
         strMsg = "被授權人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP72_GotFocus
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

' 被授權人中文名稱
Private Sub textCP50_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP50) = False Then
      If CheckLengthIsOK(textCP50, 60) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "被授權人中文名稱內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP50_GotFocus
      End If
   End If
End Sub

' 被授權人英文名稱
Private Sub textCP51_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP51) = False Then
      If CheckLengthIsOK(textCP51, 60) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "被授權人英文名稱內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP51_GotFocus
      End If
   End If
End Sub

' 被授權人日文名稱
Private Sub textCP52_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP52) = False Then
      If CheckLengthIsOK(textCP52, 60) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "被授權人日文名稱內容太長"
         'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP52_GotFocus
      End If
   End If
End Sub

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
            textPrint_GotFocus
      End Select
   End If
End Sub

' 更新欄位的內容
Private Sub OnUpdateField()
   ' 發文日
   SetCPFieldNewData "CP27", DBDATE(textCP27)
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
   ' 授權人代號
   If IsEmptyText(textCP72) = False Then
      SetCPFieldNewData "CP72", textCP72 & String(9 - Len(textCP72), "0")
   Else
      SetCPFieldNewData "CP72", textCP72
   End If
   ' 被授權人(中)
   SetCPFieldNewData "CP50", textCP50
   ' 被授權人(英)
   SetCPFieldNewData "CP51", textCP51
   ' 被授權人(日)
   SetCPFieldNewData "CP52", textCP52
   ' 授權期間(起)
   SetCPFieldNewData "CP53", DBDATE(textCP53)
   ' 授權期間(迄)
   SetCPFieldNewData "CP54", DBDATE(textCP54)
   
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
Private Sub OnUpdateCaseProperty()
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
   
   ' 更新案件進度檔
   strSql = "UPDATE CaseProgress SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_CPCount - 1
      strTmp = Empty
      If m_CPList(nIndex).fiOldData <> m_CPList(nIndex).fiNewData Then
         If m_CPList(nIndex).fiType = 0 Then
            ' 91.03.25 modify by louis (單引號)
            'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
            strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & 0
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
   
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim strSql As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP22 As String
      
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   ' 更新案件進度檔
   OnUpdateCaseProperty
   
   ' 若有輸入催審期限時, 新增一筆催審的記錄到下一程序檔
   If IsEmptyText(textUargeDate) = False Then
      'Add By Sindy 2023/5/5 FCT重新發文，若下一程序已有該收文號未續辦之催審期限，則更新期限即可，不要另新增期限
      strExc(0) = "SELECT NP01,NP22 from NextProgress" & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strSql = "UPDATE NextProgress SET NP08=" & PUB_GetWorkDay1(textUargeDate, True) & ",NP09=" & DBDATE(textUargeDate) & _
                  " Where NP01='" & m_CP09 & "' and NP07='305' and NP06 is null"
         cnnConnection.Execute strSql
      Else
      '2023/5/5 END
         strNP07 = "305"
         strNP22 = GetNextProgressNo()
           'Modify By Cheng 2003/09/05
   '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
   '               "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
   '                        DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & strUserNum & "'," & strNP22 & ")"
         'Modified by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         'strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           DBDATE(textUargeDate) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                  "VALUES ('" & m_CP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & strNP07 & "," & _
                           PUB_GetWorkDay1(textUargeDate, True) & "," & DBDATE(textUargeDate) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & strNP22 & ")"
         cnnConnection.Execute strSql
      End If
      
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case strNP07
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
            'Modify By Cheng 2002/01/15
            '取消外商FCT列印接洽結案單
'            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
      End Select
   End If
   

   'add by nick 2004/08/13 更新實際發文規費
   If textCP84.Enabled = True Then
      strSql = "Update CaseProgress Set CP84=" & Trim(Val(textCP84.Text)) & " Where CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
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
   
   'Add By Sindy 2012/9/26 檢查是否為一申請書多件並更新資料
   '授權案、再授權案
   'Modify By Sindy 2013/4/9 定稿語文是英文時才做一申請書多件
   'Modify By Sindy 2014/6/24 mark : 不管定稿語文
   'If (m_CP10 = "502" Or m_CP10 = "504") And GetLetterLanguage(m_TM01, m_TM02, m_TM03, m_TM04) = "2" Then
   If (m_CP10 = "502" Or m_CP10 = "504") Then
   '2013/4/9 End
   '2014/6/24 END
      Call PUB_UpdateCP148(m_TM01, m_TM02, m_TM03, m_TM04, m_CP10, textCP27)
   End If
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
     'Add by nickc 2008/02/22 檢查代理人Email(需考慮可能為FF案件)
    PUB_CheckEMail m_CP44, m_CP116
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
   
   CheckDataValid = False
   ' 發文日
   If IsEmptyText(textCP27) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入發文日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP27.SetFocus
      GoTo EXITSUB
   End If
   ' 被授權人(中, 英, 日) 不可同時空白
   If IsEmptyText(textCP50) = True And IsEmptyText(textCP51) = True And IsEmptyText(textCP52) = True Then
      strTit = "檢核資料"
      '2009/12/23 modify by sonia FCT-029281
      'strMsg = "請輸入被授權人(中, 英, 日)"
      If m_CP10 <> "724" Then
         strMsg = "請輸入被授權人(中, 英, 日)"
      Else
         strMsg = "請於授權資料頁籤輸入的徵求同意書對象(中/英/日)欄"
      End If
      '2009/12/23 end
'      'Add By Sindy 2013/12/20 724徵求同意書,改提醒即可,不必一定要輸入
'      If m_CP10 = "724" Then
'         If MsgBox("是否要輸入徵求同意書對象(中/英/日)欄？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
'            SSTab1.Tab = 1
'            textCP50.SetFocus
'            GoTo EXITSUB
'         End If
'      Else
'      '2013/12/20 END
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP50.SetFocus
         GoTo EXITSUB
'      End If
   End If
   If (m_CP10 = "502" Or m_CP10 = "504") Then  '2009/10/14 add by sonia
      ' 授權期間(起)
      If IsEmptyText(textCP53) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入授權期間起日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP53.SetFocus
         GoTo EXITSUB
      End If
      ' 授權期間(迄)
      If IsEmptyText(textCP54) = True Then
         strTit = "檢核資料"
         strMsg = "請輸入授權期間迄日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP54.SetFocus
         GoTo EXITSUB
      End If
      ' 授權期間的範圍
      If IsEmptyText(textCP53) = False And IsEmptyText(textCP54) = False Then
         If Val(textCP53) > Val(textCP54) Then
            strTit = "檢核資料"
            strMsg = "授權期間範圍不正確"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP53.SetFocus
         GoTo EXITSUB
         End If
      End If
   End If
   '2009/10/14 ADD BY SONIA
   If m_CP10 = "724" And textUargeDate = "" Then
'      'Add By Sindy 2013/12/20 724徵求同意書,改提醒即可,不必一定要輸入
'      If MsgBox("徵求同意書發文，是否要輸入催審期限？", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
'         SSTab1.Tab = 0
'         textUargeDate.SetFocus
'         GoTo EXITSUB
'      End If
'      '2013/12/20 END
      strTit = "檢核資料"
      strMsg = "徵求同意書發文, 催審期限不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textUargeDate.SetFocus
      GoTo EXITSUB
   End If
   '2009/10/14 END
   
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
   
   'Added by Lydia 2021/09/07 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
        GoTo EXITSUB
   End If
    
   CheckDataValid = True
EXITSUB:
End Function



' 催審期限
Private Sub textUargeDate_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textUargeDate) = False Then
      If CheckIsTaiwanDate(textUargeDate, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "催審期限日期不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textUargeDate_GotFocus
      End If
   End If
End Sub

Private Sub textCP27_GotFocus()
   InverseTextBox textCP27
End Sub

Private Sub textUargeDate_GotFocus()
   InverseTextBox textUargeDate
End Sub

Private Sub textDN_GotFocus()
   InverseTextBox textDN
End Sub

Private Sub textAdd_GotFocus()
   InverseTextBox textAdd
End Sub

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCP72_GotFocus()
   InverseTextBox textCP72
End Sub

Private Sub textCP50_GotFocus()
   InverseTextBox textCP50
End Sub

Private Sub textCP51_GotFocus()
   InverseTextBox textCP51
End Sub

Private Sub textCP52_GotFocus()
   InverseTextBox textCP52
End Sub

Private Sub textCP53_GotFocus()
   InverseTextBox textCP53
End Sub

Private Sub textCP54_GotFocus()
   InverseTextBox textCP54
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

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
      'add by sonia 2019/12/17 徵求同意書724僅提醒即可 FCT-043686
      If m_CP10 = "724" Then
         'modify by sonia 2022/3/3 改為不檢查
         'If MsgBox("發文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同，是否繼續發文？", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then
         '   textCP84_GotFocus
         '   Exit Function
         'End If
         'end 2022/3/3
      Else
      'end 2019/12/17
         MsgBox "發文規費[" & Trim(Val(m_CP84)) & "] 與實際發文規費[" & Trim(Val(textCP84.Text)) & "]不同", , "警告！"
         textCP84_GotFocus
         Exit Function
      End If
   End If
End If

If Me.textAdd.Enabled = True Then
   Cancel = False
   textAdd_Validate Cancel
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

If Me.textCP50.Enabled = True Then
   Cancel = False
   textCP50_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP51.Enabled = True Then
   Cancel = False
   textCP51_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP52.Enabled = True Then
   Cancel = False
   textCP52_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP53.Enabled = True Then
   Cancel = False
   textCP53_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP54.Enabled = True Then
   Cancel = False
   textCP54_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP72.Enabled = True Then
   Cancel = False
   textCP72_Validate Cancel
   If Cancel = True Then
      Exit Function
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

If Me.textPrint.Enabled = True Then
   Cancel = False
   textPrint_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textUargeDate.Enabled = True Then
   Cancel = False
   textUargeDate_Validate Cancel
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
