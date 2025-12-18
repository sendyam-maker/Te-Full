VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_05 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文"
   ClientHeight    =   6180
   ClientLeft      =   2352
   ClientTop       =   3096
   ClientWidth     =   9096
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   9096
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關案號(&F)"
      Height          =   400
      Left            =   4380
      TabIndex        =   35
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   5625
      TabIndex        =   34
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6855
      TabIndex        =   33
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7680
      TabIndex        =   32
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox textPAKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox textPA57_2 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000FF&
      Height          =   264
      Left            =   2820
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   540
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4092
      Left            =   72
      TabIndex        =   37
      Top             =   2052
      Width           =   8832
      _ExtentX        =   15579
      _ExtentY        =   7218
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm010012_05.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label37"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label8"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label4"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(43)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(12)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(17)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(18)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label12"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(3)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP13_2"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP14_2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "grdList"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPA08_2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP13"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP43"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP07"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP06"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textPA23"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP05"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textPA09"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textPA09_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP26"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP10"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP10_2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textPA08"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textPA46"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textPA48"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textPA57"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCP14"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP21"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "備註"
      TabPicture(1)   =   "frm010012_05.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdPriority"
      Tab(1).Control(1)=   "textPA47"
      Tab(1).Control(2)=   "textCP64"
      Tab(1).Control(3)=   "textPA91"
      Tab(1).Control(4)=   "Label1(2)"
      Tab(1).Control(5)=   "Label1(11)"
      Tab(1).Control(6)=   "Label2(3)"
      Tab(1).Control(7)=   "Label2(2)"
      Tab(1).Control(8)=   "Label2(1)"
      Tab(1).Control(9)=   "Label2(0)"
      Tab(1).Control(10)=   "Label1(9)"
      Tab(1).Control(11)=   "Label1(10)"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "讓與申請人"
      TabPicture(2)   =   "frm010012_05.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "textCP92"
      Tab(2).Control(1)=   "textCP91"
      Tab(2).Control(2)=   "textCP90"
      Tab(2).Control(3)=   "textCP89"
      Tab(2).Control(4)=   "textCP56"
      Tab(2).Control(5)=   "textCP92_2"
      Tab(2).Control(6)=   "textCP91_2"
      Tab(2).Control(7)=   "textCP90_2"
      Tab(2).Control(8)=   "textCP89_2"
      Tab(2).Control(9)=   "textCP56_2"
      Tab(2).Control(10)=   "Label92"
      Tab(2).Control(11)=   "Label91"
      Tab(2).Control(12)=   "Label90"
      Tab(2).Control(13)=   "Label89"
      Tab(2).Control(14)=   "Label36"
      Tab(2).ControlCount=   15
      Begin VB.TextBox textCP92 
         Height          =   264
         Left            =   -73470
         MaxLength       =   9
         TabIndex        =   23
         Top             =   1950
         Width           =   1095
      End
      Begin VB.TextBox textCP91 
         Height          =   264
         Left            =   -73470
         MaxLength       =   9
         TabIndex        =   22
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox textCP90 
         Height          =   264
         Left            =   -73470
         MaxLength       =   9
         TabIndex        =   21
         Top             =   1290
         Width           =   1095
      End
      Begin VB.TextBox textCP89 
         Height          =   264
         Left            =   -73470
         MaxLength       =   9
         TabIndex        =   20
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox textCP56 
         Height          =   264
         Left            =   -73470
         MaxLength       =   9
         TabIndex        =   19
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox textCP21 
         Height          =   264
         Left            =   1368
         MaxLength       =   20
         TabIndex        =   6
         Top             =   1320
         Width           =   372
      End
      Begin VB.CommandButton cmdPriority 
         Caption         =   "優先權資料(&P)"
         Height          =   285
         Left            =   -73560
         TabIndex        =   15
         Top             =   420
         Width           =   1300
      End
      Begin VB.TextBox textPA47 
         Height          =   270
         Left            =   -73560
         MaxLength       =   50
         TabIndex        =   18
         Top             =   2295
         Width           =   7275
      End
      Begin VB.TextBox textCP14 
         Height          =   270
         Left            =   1380
         MaxLength       =   6
         TabIndex        =   0
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox textPA57 
         Height          =   270
         Left            =   1380
         MaxLength       =   1
         TabIndex        =   11
         Top             =   2220
         Width           =   375
      End
      Begin VB.TextBox textPA48 
         Height          =   270
         Left            =   5340
         MaxLength       =   30
         TabIndex        =   14
         Top             =   2520
         Width           =   1935
      End
      Begin VB.TextBox textPA46 
         Height          =   270
         Left            =   5340
         MaxLength       =   1
         TabIndex        =   12
         Top             =   2220
         Width           =   375
      End
      Begin VB.TextBox textPA08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   1368
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2520
         Width           =   375
      End
      Begin VB.TextBox textCP10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   420
         Width           =   2475
      End
      Begin VB.TextBox textCP10 
         Height          =   264
         Left            =   5340
         MaxLength       =   6
         TabIndex        =   1
         Top             =   420
         Width           =   732
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   2
         Top             =   720
         Width           =   372
      End
      Begin VB.TextBox textPA09_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1020
         Width           =   1500
      End
      Begin VB.TextBox textPA09 
         Height          =   264
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   4
         Top             =   1020
         Width           =   612
      End
      Begin VB.TextBox textCP05 
         Height          =   264
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   5
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox textPA23 
         Height          =   264
         Left            =   5340
         MaxLength       =   20
         TabIndex        =   7
         Top             =   1320
         Width           =   372
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1368
         MaxLength       =   7
         TabIndex        =   8
         Top             =   1620
         Width           =   1215
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   9
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox textCP43 
         Height          =   264
         Left            =   1620
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox textCP13 
         Height          =   264
         Left            =   5340
         MaxLength       =   6
         TabIndex        =   3
         Top             =   720
         Width           =   852
      End
      Begin VB.TextBox textPA08_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   1800
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2115
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1164
         Left            =   72
         TabIndex        =   89
         Top             =   2856
         Width           =   8592
         _ExtentX        =   15155
         _ExtentY        =   2053
         _Version        =   393216
         AllowUserResizing=   3
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
      Begin MSForms.TextBox textCP92_2 
         Height          =   264
         Left            =   -72330
         TabIndex        =   87
         TabStop         =   0   'False
         Top             =   1950
         Width           =   2175
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP91_2 
         Height          =   264
         Left            =   -72330
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   1620
         Width           =   2175
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP90_2 
         Height          =   264
         Left            =   -72330
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   1290
         Width           =   2175
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP89_2 
         Height          =   264
         Left            =   -72330
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP56_2 
         Height          =   264
         Left            =   -72330
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   630
         Width           =   2175
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   735
         Left            =   -73560
         TabIndex        =   16
         Top             =   750
         Width           =   7275
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12832;1296"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA91 
         Height          =   735
         Left            =   -73560
         TabIndex        =   17
         Top             =   1560
         Width           =   7275
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12832;1296"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   2280
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   420
         Width           =   1635
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP13_2 
         Height          =   264
         Left            =   6240
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   720
         Width           =   1635
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label92 
         Caption         =   "讓與申請人5 :"
         Height          =   195
         Left            =   -74670
         TabIndex        =   88
         Top             =   1980
         Width           =   1095
      End
      Begin VB.Label Label91 
         Caption         =   "讓與申請人4 :"
         Height          =   195
         Left            =   -74670
         TabIndex        =   86
         Top             =   1650
         Width           =   1095
      End
      Begin VB.Label Label90 
         Caption         =   "讓與申請人3 :"
         Height          =   195
         Left            =   -74670
         TabIndex        =   84
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label89 
         Caption         =   "讓與申請人2 :"
         Height          =   195
         Left            =   -74670
         TabIndex        =   82
         Top             =   990
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "讓與申請人1 :"
         Height          =   195
         Left            =   -74670
         TabIndex        =   80
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Index           =   3
         Left            =   1980
         TabIndex        =   78
         Top             =   1320
         Width           =   465
      End
      Begin VB.Label Label12 
         Caption         =   "是否為多國案 :"
         Height          =   255
         Left            =   180
         TabIndex        =   77
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "優先權資料"
         Height          =   180
         Index           =   2
         Left            =   -74760
         TabIndex        =   76
         Top             =   480
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類"
         Height          =   180
         Index           =   18
         Left            =   180
         TabIndex        =   67
         Top             =   2520
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "(Y:是)"
         Height          =   180
         Index           =   17
         Left            =   6000
         TabIndex        =   66
         Top             =   2220
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否 PCT 案件"
         Height          =   180
         Index           =   12
         Left            =   4140
         TabIndex        =   65
         Top             =   2220
         Width           =   1125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "分所案號"
         Height          =   180
         Index           =   11
         Left            =   -74760
         TabIndex        =   64
         Top             =   2310
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   63
         Top             =   3624
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   62
         Top             =   2484
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   -72120
         TabIndex        =   61
         Top             =   1524
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   -72120
         TabIndex        =   60
         Top             =   504
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   59
         Top             =   450
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號"
         Height          =   180
         Index           =   4
         Left            =   4140
         TabIndex        =   58
         Top             =   2520
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷"
         Height          =   180
         Index           =   43
         Left            =   180
         TabIndex        =   57
         Top             =   2220
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   56
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   55
         Top             =   1530
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "案件性質 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   54
         Top             =   420
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   1980
         TabIndex        =   53
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   180
         TabIndex        =   52
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家 :"
         Height          =   255
         Index           =   8
         Left            =   180
         TabIndex        =   51
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   50
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "(1:申請 2:異議 3:舉發)"
         Height          =   255
         Left            =   5820
         TabIndex        =   49
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "卷宗性質 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   48
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   180
         TabIndex        =   47
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   46
         Top             =   1620
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "(Y : 取消閉卷)"
         Height          =   255
         Left            =   2040
         TabIndex        =   44
         Top             =   2280
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   255
         Index           =   1
         Left            =   4140
         TabIndex        =   43
         Top             =   720
         Width           =   975
      End
   End
   Begin MSForms.ComboBox cmbPA05 
      Height          =   300
      Left            =   930
      TabIndex        =   36
      Top             =   810
      Width           =   7935
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "13996;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA26_2 
      Height          =   264
      Left            =   960
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA27_2 
      Height          =   264
      Left            =   4980
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA28_2 
      Height          =   264
      Left            =   960
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA30_2 
      Height          =   264
      Left            =   960
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA29_2 
      Height          =   264
      Left            =   4980
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1440
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textPA75_2 
      Height          =   264
      Left            =   4980
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2475
      VariousPropertyBits=   671107103
      ForeColor       =   -2147483641
      Size            =   "3413;476"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人4"
      Height          =   180
      Index           =   35
      Left            =   4260
      TabIndex        =   75
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人3"
      Height          =   180
      Index           =   34
      Left            =   120
      TabIndex        =   74
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人2"
      Height          =   180
      Index           =   32
      Left            =   4260
      TabIndex        =   73
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人1"
      Height          =   180
      Index           =   31
      Left            =   120
      TabIndex        =   72
      Top             =   1170
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請人5"
      Height          =   180
      Index           =   36
      Left            =   120
      TabIndex        =   71
      Top             =   1740
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "代理人"
      Height          =   180
      Index           =   38
      Left            =   4260
      TabIndex        =   70
      Top             =   1740
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   69
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   68
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "frm010012_05"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/16 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Morgan 2021/5/12 改成Form2.0 (cmbPA05,textPA26_2...)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/22 日期欄已修改
Option Explicit

' 本所案號
Dim m_PA01 As String
Dim m_PA02 As String
Dim m_PA03 As String
Dim m_PA04 As String

' 收文日
Dim m_CP05 As String
' 專利種類
Dim m_PA08 As String
' 收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 國家代碼
Dim m_PA09 As String
' 是否閉卷
Dim m_PA57 As String
' 讓與申請人
Dim m_CP56 As String
'Add By Sindy 2009/10/19
Dim m_CP89 As String
Dim m_CP90 As String
Dim m_CP91 As String
Dim m_CP92 As String
Dim m_CP55 As String '讓與人
Dim m_CP93 As String
Dim m_CP94 As String
Dim m_CP95 As String
Dim m_CP96 As String
Dim m_PA26 As String '申請人
Dim m_PA27 As String
Dim m_PA28 As String
Dim m_PA29 As String
Dim m_PA30 As String
'2009/10/19 End
' 相關總收文號
Dim m_CP43 As String
' 是否PCT案件
Dim m_PA46 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiOldData As String
   fiNewData As String
   fiType As Integer
End Type

' 儲存專利基本檔或服務業務基本檔檔案欄位的串列
Dim m_PASPList() As FIELDITEM
Dim m_PASPCount As Integer
' 儲存案件進度檔檔案欄位的串列
Dim m_CPList() As FIELDITEM
Dim m_CPCount As Integer
' 儲存國家的字串
Dim m_strCountry As String
'
Dim m_CurrSel As Integer
' 優先權畫面所使用的變數
Dim m_Pa(1 To 4) As String '本所案號
Dim m_Priority(1 To 5) As String

Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限

'Add by Morgan 2004/2/18
'若承辦人是王協理且未發文則要發EMail通知
Dim stCP09 As String, stCP14 As String, stCP27 As String
'Add By Sindy 2018/2/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_PrevForm As Form '前一畫面
'2018/2/22 END
'Add by Sindy 2023/12/6 相關總收文號的資料
Dim m_CP43CP08 As String
Dim m_CP43CP64 As String
'2023/12/6 END
Dim m_CP44 As String, m_CP45 As String 'Added by Morgan 2024/3/8

Private Sub ClearAll()
   ClearPASPFieldList
   ClearCPFieldList
   
   textPAKey = Empty
   textPA08 = Empty
   textPA09 = Empty
   textPA23 = Empty
   textPA46 = Empty
   textPA47 = Empty
   textPA48 = Empty
   textPA57 = Empty
   textPA91 = Empty
   
   textCP05 = Empty
   textCP06 = Empty
   textCP07 = Empty
   textCP10 = Empty
   textCP10_2 = Empty
   textCP13 = Empty
   textCP13_2 = Empty
   textCP14 = Empty
   textCP14_2 = Empty
   '911113 nick 邱小姐說刪除
   'textCP16 = Empty
   'textCP17 = Empty
   'textCP18 = Empty
   textCP26 = Empty
   textCP43 = Empty
   textCP56 = Empty
   'Add By Sindy 2009/10/19
   textCP89 = Empty
   textCP90 = Empty
   textCP91 = Empty
   textCP92 = Empty
   '2009/10/19 End
   textCP64 = Empty
   
   m_strCountry = Empty
End Sub

Public Sub SetData(ByVal strData As String, ByVal nType As Integer, ByVal bClear As Boolean)
   If bClear Then
      m_PA01 = Empty
      m_PA02 = Empty
      m_PA03 = Empty
      m_PA04 = Empty
      m_CP10 = Empty
      m_CP56 = Empty
      'Add By Sindy 2009/10/19
      m_CP89 = Empty
      m_CP90 = Empty
      m_CP91 = Empty
      m_CP92 = Empty
      '2009/10/19 End
      '92.03.27 nick
      m_CP09 = Empty
   End If
   
   Select Case nType
      Case 0: m_PA01 = strData
      Case 1: m_PA02 = strData
      Case 2: m_PA03 = strData & String(1 - Len(strData), "0")
      Case 3: m_PA04 = strData & String(2 - Len(strData), "0")
      Case 4:
              m_CP10 = strData
              '911113 nick
              'Modify By Sindy 2009/10/19 增加案件性質708
                If textCP10 = "701" Or textCP10 = "708" Then
                   Label36.Visible = True
                   EnableTextBox textCP56, True
                   textCP56.Visible = True
                   textCP56_2.Visible = True
                   'Add By Sindy 2009/10/19
                   Label89.Visible = True
                   EnableTextBox textCP89, True
                   textCP89.Visible = True
                   textCP89_2.Visible = True
                   Label90.Visible = True
                   EnableTextBox textCP90, True
                   textCP90.Visible = True
                   textCP90_2.Visible = True
                   Label91.Visible = True
                   EnableTextBox textCP91, True
                   textCP91.Visible = True
                   textCP91_2.Visible = True
                   Label92.Visible = True
                   EnableTextBox textCP92, True
                   textCP92.Visible = True
                   textCP92_2.Visible = True
                   '2009/10/19 End
                Else
                   Label36.Visible = False
                   EnableTextBox textCP56, False
                   textCP56.Visible = False
                   textCP56_2.Visible = False
                   'Add By Sindy 2009/10/19
                   Label89.Visible = False
                   EnableTextBox textCP89, False
                   textCP89.Visible = False
                   textCP89_2.Visible = False
                   Label90.Visible = False
                   EnableTextBox textCP90, False
                   textCP90.Visible = False
                   textCP90_2.Visible = False
                   Label91.Visible = False
                   EnableTextBox textCP91, False
                   textCP91.Visible = False
                   textCP91_2.Visible = False
                   Label92.Visible = False
                   EnableTextBox textCP92, False
                   textCP92.Visible = False
                   textCP92_2.Visible = False
                   '2009/10/19 End
                End If
              
      Case 5:
         If Not IsEmptyText(strData) Then
            m_CP56 = strData & String(9 - Len(strData), "0")
         End If
      Case 6:
         m_CP43 = strData
         textCP43 = m_CP43
      Case 7:
         m_CP09 = strData
      'Add By Sindy 2009/10/19
      Case 8:
         If Not IsEmptyText(strData) Then
            m_CP89 = strData & String(9 - Len(strData), "0")
         End If
      Case 9:
         If Not IsEmptyText(strData) Then
            m_CP90 = strData & String(9 - Len(strData), "0")
         End If
      Case 10:
         If Not IsEmptyText(strData) Then
            m_CP91 = strData & String(9 - Len(strData), "0")
         End If
      Case 11:
         If Not IsEmptyText(strData) Then
            m_CP92 = strData & String(9 - Len(strData), "0")
         End If
      '2009/10/19 End
   End Select
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm010001.Show
End Sub

Private Sub cmdCaseProgress_Click()
   frm010012_03.SetData 0, m_PA01, True
   frm010012_03.SetData 1, m_PA02, False
   frm010012_03.SetData 2, m_PA03, False
   frm010012_03.SetData 3, m_PA04, False
   frm010012_03.SetData 4, m_CP09, False
   'Modified by Lydia 2020/04/21 改為Form型態
   'frm010012_03.SetParent "frm010012_05"
   frm010012_03.SetParent Me
   Me.Hide
   frm010012_03.Show
   frm010012_03.QueryData
End Sub

' 91.10.15 因加入此功能會牽扯太多的專利畫面, 因此拿掉此功能
'Private Sub cmdForeign_Click()
'   Me.Hide
'   frm050106_1.intWhereToGo = 2
'   frm050106_1.Show
'End Sub

' 91.10.15 因加入此功能會牽扯太多的專利畫面, 因此拿掉此功能
'Private Sub cmdIDS_Click()
'   Me.Hide
'   frm050107_1.intWhereToGo = 2
'   frm050107_1.Show
'   frm050107_1.txtCode(0) = m_PA01
'   frm050107_1.txtCode(1) = m_PA02
'   frm050107_1.txtCode(2) = m_PA03
'   frm050107_1.txtCode(3) = m_PA04
'End Sub

Private Sub cmdOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      '重新檢查欄位有效性
      If ValidateInput() = False Then
         Exit Sub
      End If
      'Added by Lydia 2015/02/04 所有內部收文, 若有輸入本所期限或法定期限者, 檢查期限不可小於系統日
      'Modified by Lydia 2017/07/31 改為預設和檢查
      'If PUB_CheckCP0607(0, textCP06.Text, textCP07.Text) = False Then Exit Sub
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Sub
      If PUB_CheckCP0607(0, textCP06, textCP07, "", textPA09, m_PA01, textCP10) = False Then Exit Sub
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      OnUpdateField
    'Modify By Cheng 2002/11/06
'      'OnSaveData

      'Add By Sindy 2022/7/1
      If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
         If PUB_ChkFileOpening2(m_PrevForm.m_strFullFileName, "後續才能一併歸卷！") = True Then
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
      End If
      '2022/7/1 END
      
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      
      PUB_SendMailCache 'Added by Morgan 2013/8/1
      
      'Add by Morgan 2004/2/18
      '若承辦人是王協理且未發文則要發EMail通知
      stCP14 = textCP14
      'Modify by Amy 2024/07/16 原:71011(王副總) 改李柏翰經理
      If stCP14 = "99050" Then
         stCP09 = m_CP09
         Call PUB_SendMail(strUserNum, "99050", stCP09, "分案通知")
      End If
           
      'Add By Sindy 2021/9/6 614.檢視核准版本, 內部收文後系統請自動發MAIL如下:
      '  P-XXXXXX-0-00 已內部收文檢視核准版本，請至卷宗區參看相關電子檔，並進行處理。
      If textCP10 = "614" And frm010001.intModifyKind = 0 Then
         PUB_SendMail strUserNum, textCP14, m_CP09, m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 = "000", "", "-" & m_PA03 & "-" & m_PA04) & " 已內部收文檢視核准版本，請至卷宗區參看核准通知相關電子檔，並進行處理。"
      End If
      '2021/9/6 END
      
      'Added by Lydia 2023/02/08 內部收文補收款，智權人員為SXX部門時，要發MAIL給杜協理及智權人員
      If (m_PA01 = "P" Or m_PA01 = "PS" Or m_PA01 = "CFP" Or m_PA01 = "CPS") And textCP10 <> "" And InStr(textCP10_2, "補收款") > 0 And Left(GetST15(textCP13), 1) = "S" Then
          strExc(0) = m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
          strExc(1) = "本所案號：" & strExc(0) & vbCrLf & _
                           "案件名稱：" & cmbPA05.List(0) & vbCrLf & _
                           "申請人1：" & m_PA26 & " " & textPA26_2 & vbCrLf & _
                           "申請國家：" & textPA09_2 & vbCrLf & _
                           "補收款費用：0" & vbCrLf & _
                           "補收款備註：" & Trim(textCP64)
          strExc(2) = Pub_GetSpecMan("全所智權部主管")
          If InStr(strExc(2), textCP13) = 0 Then
              strExc(2) = strExc(2) & ";" & textCP13
          End If
          PUB_SendMail strUserNum, strExc(2), "", strExc(0) & "內部收文補收款通知!", strExc(1)
      End If
      'end 2023/02/08
      
          'Add By Sindy 2023/12/4
      Dim m_StrTo As String, m_StrSub As String, m_StrCont As String, m_strCP09 As String
      If (textCP10 = "211" Or textCP10 = "212") And textPA09 = "000" Then
         If textCP06 <> "" And textCP07 <> "" Then
            strSql = "select CP08,CP64 from CASEPROGRESS where CP09='" & textCP43 & "'"
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount > 0 Then
               m_CP43CP08 = CheckStr(adoRecordset.Fields(0))
               m_CP43CP64 = CheckStr(adoRecordset.Fields(1))
            End If
            'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
            'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
            'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & textCP13
            m_StrTo = PUB_GetLosCL02list(m_PA01, m_PA02, m_PA03, m_PA04)
            m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & textCP13
            'end 2024/10/30
            
            m_StrSub = "開庭通知--內部收文案件：" & textPAKey
            m_StrCont = "本所案號：" & textPAKey & vbCrLf & _
                        "案件名稱：" & cmbPA05 & vbCrLf & _
                        "案件性質：" & textCP10_2 & vbCrLf & _
                        "申請人　：" & textPA26_2 & vbCrLf & _
                        "承辦人　：" & textCP14_2 & vbCrLf & _
                        "智權人員：" & textCP13_2 & vbCrLf & _
                        "法定期限：" & DBYEAR(textCP07.Text) - 1911 & "" & " 年 " & DBMONTH(textCP07.Text) & "" & " 月 " & DBDAY(textCP07.Text) & "" & " 日 " & vbCrLf & _
                        "時間地點：" & m_CP43CP64 & vbCrLf & _
                        "法院案號：" & m_CP43CP08
            m_strCP09 = PUB_GetLastABKindCP09(m_PA01, m_PA02, m_PA03, m_PA04)
            PUB_SendMail strUserNum, m_StrTo, m_strCP09, m_StrSub, m_StrCont
         End If
      End If
      '2023/12/4 END
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      'Modify By Sindy 2018/2/22 信件內部收文執行完畢後,關閉視窗
      If m_strIR01 <> "" Then
         'Modify By Sindy 2022/6/29
         If Pub_StrUserSt03 = "F23" Then
            Call PUB_RecvOutLookF23(m_strIR01, m_strIR02, m_strIR03, m_strIR04, "1", m_CP09)
         End If
         '2022/6/29 END
         Unload frm010001
         Unload Me
      Else
      '2018/2/22 End
         ' 回到收文的畫面
         frm010001.SetData m_CP09, 0, True
         frm010001.SetData m_PA01, 1, False
         frm010001.SetData m_PA02, 2, False
         frm010001.SetData m_PA03, 3, False
         frm010001.SetData m_PA04, 4, False
         frm010001.Show
         ClearAll
         Unload Me
      End If
   End If
End Sub

Private Sub cmdPriority_Click()
   ' 修改優先權資料
   'Modify by Amy 2014/04/18 +, m_Priority(5)
   'Modify by Sindy 2019/1/23 + m_PA01 & m_PA02 & m_PA03 & m_PA04
   ModifyPriority m_Priority(1), m_Priority(2), m_Priority(3), , , m_PA01 & m_PA02 & m_PA03 & m_PA04, , , m_Priority(4), m_Priority(5)
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_PA01, m_PA02, m_PA03, m_PA04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textPAKey.BackColor = &H8000000F
   textPA08_2.BackColor = &H8000000F
   textPA09_2.BackColor = &H8000000F
   textPA26_2.BackColor = &H8000000F
   textPA27_2.BackColor = &H8000000F
   textPA28_2.BackColor = &H8000000F
   textPA29_2.BackColor = &H8000000F
   textPA30_2.BackColor = &H8000000F
   textPA57_2.BackColor = &H8000000F
   textPA75_2.BackColor = &H8000000F
   
   textCP10_2.BackColor = &H8000000F
   textCP13_2.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP56_2.BackColor = &H8000000F
   'Add By Sindy 2009/10/19
   textCP89_2.BackColor = &H8000000F
   textCP90_2.BackColor = &H8000000F
   textCP91_2.BackColor = &H8000000F
   textCP92_2.BackColor = &H8000000F
   '2009/10/19 End
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
   
   'Add By Sindy 2018/2/22
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2018/2/22 END
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   GrdList.Clear
   GrdList.Rows = 1
   '911111 nick
   'grdList.Cols = 11
   GrdList.Cols = 12
   GrdList.ColWidth(0) = 300
   GrdList.row = 0
   GrdList.col = 1
   GrdList.Text = "下一程序"
   GrdList.ColWidth(1) = 1200
   GrdList.col = 2
   GrdList.Text = "本所期限"
   GrdList.ColWidth(2) = 1000
   GrdList.col = 3
   GrdList.Text = "法定期限"
   GrdList.ColWidth(3) = 1000
   GrdList.col = 4
   GrdList.Text = "機關文號"
   GrdList.ColWidth(4) = 1000
   GrdList.col = 5
   GrdList.Text = "相關人"
   GrdList.ColWidth(5) = 1200
   GrdList.col = 6
   GrdList.Text = "解除期限日"
   GrdList.ColWidth(6) = 1200
   GrdList.col = 7
   GrdList.Text = "備註"
   GrdList.ColWidth(7) = 1200
   GrdList.col = 8
   GrdList.Text = "收文號"
   GrdList.ColWidth(8) = 0
   GrdList.col = 9
   GrdList.Text = "下一程序代號"
   GrdList.ColWidth(9) = 0
   GrdList.col = 10
   GrdList.Text = "序號"
   GrdList.ColWidth(10) = 0
   '911111 nick add
   GrdList.col = 11
   GrdList.Text = "序號"
   GrdList.ColWidth(11) = 0
End Sub

Private Sub UpdateGrdList(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String)
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   
   'Modify by Morgan 2009/12/23 下一程序要排除程序管制的案件性質
   'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & strPA01 & "' AND " & _
                  "NP03 = '" & strPA02 & "' AND " & _
                  "NP04 = '" & strPA03 & "' AND " & _
                  "NP05 = '" & strPA04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') " & strNpSqlOfNoSalesDuty
   
   'Add by Morgan 2009/12/23 延期+AB類未發文未取消收文的程序
   If textCP10 = "404" Then
      textCP10.Enabled = False
      strSql = strSql & " UNION SELECT CP09,CP10,CP06,CP07,CP13,CP57,CP58,CP08,CP40,CP64,0" & _
         " FROM CASEPROGRESS WHERE CP01 = '" & strPA01 & "' AND CP02 = '" & strPA02 & "'" & _
         " AND CP03 = '" & strPA03 & "' AND CP04 = '" & strPA04 & "'" & _
         " AND CP09<'C' and cp10<>'404' and cp07>0 AND CP27 IS NULL AND CP57 IS NULL"
   End If
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While rsTmp.EOF = False
         GrdList.Rows = GrdList.Rows + 1
         nIndex = GrdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            GrdList.TextMatrix(nIndex, 8) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            '911111 nick 案件性質要依國家判斷
            'grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_PA01, rsTmp.Fields("NP07"))
            GrdList.TextMatrix(nIndex, 1) = GetPrjState4(strPA01 & "-" & strPA02 & "-" & strPA03 & "-" & strPA04, rsTmp.Fields("NP07"))
            
            GrdList.TextMatrix(nIndex, 9) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               GrdList.TextMatrix(nIndex, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               GrdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            GrdList.TextMatrix(nIndex, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            GrdList.TextMatrix(nIndex, 5) = rsTmp.Fields("NP14")
         End If
         ' 解除期限日期
         If IsNull(rsTmp.Fields("NP11")) = False Then
            GrdList.TextMatrix(nIndex, 6) = rsTmp.Fields("NP11")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            GrdList.TextMatrix(nIndex, 7) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            GrdList.TextMatrix(nIndex, 10) = rsTmp.Fields("NP22")
         End If
         '911111 nick 智權人員
         If IsNull(rsTmp.Fields("NP10")) = False Then
            GrdList.TextMatrix(nIndex, 11) = rsTmp.Fields("NP10")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/16
      If GrdList.Rows >= 2 Then
         GrdList.FixedRows = 1
      End If
      'end 2023/10/16
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 清除商標基本檔檔案欄位串列
Private Sub ClearPASPFieldList()
   If m_PASPCount > 0 Then
      Erase m_PASPList
   End If
   m_PASPCount = 0
End Sub

' 設定專利基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetPASPFieldOldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_PASPCount - 1
      If m_PASPList(nPos).fiName = strFieldName Then
         bFind = True
         m_PASPList(nPos).fiOldData = strFieldData
         m_PASPList(nPos).fiNewData = strFieldData
         m_PASPList(nPos).fiType = nFieldType
         Exit For
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_PASPList(m_PASPCount + 1)
      m_PASPList(m_PASPCount).fiName = strFieldName
      m_PASPList(m_PASPCount).fiOldData = strFieldData
      m_PASPList(m_PASPCount).fiNewData = strFieldData
      m_PASPList(m_PASPCount).fiType = nFieldType
      m_PASPCount = m_PASPCount + 1
   End If
End Sub

' 設定專利基本檔或服務業務基本檔欄位串列中的欄位內容
Private Sub SetPASPFieldNewData(ByVal strFieldName As String, ByVal strFieldData As String)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_PASPCount - 1
      If m_PASPList(nPos).fiName = strFieldName Then
         bFind = True
         m_PASPList(nPos).fiNewData = strFieldData
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

' 更新欄位的內容
Private Sub OnUpdateField()
   SetCPFieldNewData "CP01", m_PA01
   SetCPFieldNewData "CP02", m_PA02
   SetCPFieldNewData "CP03", m_PA03
   SetCPFieldNewData "CP04", m_PA04
   ' 收文日
   If IsEmptyText(textCP05) = False Then
      SetCPFieldNewData "CP05", DBDATE(textCP05)
   Else
      SetCPFieldNewData "CP05", Empty
   End If
   ' 本所期限
   If IsEmptyText(textCP06) = False Then
      SetCPFieldNewData "CP06", DBDATE(textCP06)
   Else
      SetCPFieldNewData "CP06", Empty
   End If
   ' 法定期限
   If IsEmptyText(textCP07) = False Then
      SetCPFieldNewData "CP07", DBDATE(textCP07)
   Else
      SetCPFieldNewData "CP07", Empty
   End If
   ' 收文號
    '若為新增
    If frm010001.intModifyKind = 0 Then
        m_CP09 = AutoNo("B", 6)
    End If
   SetCPFieldNewData "CP09", m_CP09
   ' 案件性質
   SetCPFieldNewData "CP10", textCP10
   
   '911113 nick
   SetCPFieldNewData "CP11", "90"
   SetCPFieldNewData "CP20", "N"
   SetCPFieldNewData "CP32", "N"
   '911113 nick 案件性質是  909 or 911 時，cp27 = 系統日
   'Modified by Morgan 2024/11/6 改內部收文都不自動上發文,若為智權的需求都會填接洽單(A類)--郭
   'If textCP10 = "909" Or textCP10 = "911" Then
   '    'Add by Amy 2018/04/10 +if 智權=承辦且為911補收款不上發文日
   '    If Not (textCP13 = textCP14 And textCP10 = "911") Then
   '       SetCPFieldNewData "CP27", ServerDate
   '    End If
   'Else
   '    SetCPFieldNewData "CP27", Empty
   'End If
       SetCPFieldNewData "CP27", Empty
   'end 2024/11/6
   
   ' 業務區
   SetCPFieldNewData "CP12", GetSalesArea(textCP13)
   ' 智權人員
   SetCPFieldNewData "CP13", textCP13
   ' 承辦人員
   SetCPFieldNewData "CP14", textCP14
   '911113 nick 邱小姐說刪除
   ' 費用
   'SetCPFieldNewData "CP16", textCP16
   ' 規費
   'SetCPFieldNewData "CP17", textCP17
   ' 點數
   'SetCPFieldNewData "CP18", textCP18
   ' 是否為多國案
   SetCPFieldNewData "CP21", textCP21
   ' 相關總收文號
   SetCPFieldNewData "CP43", textCP43
   ' 是否算案件數
   SetCPFieldNewData "CP26", textCP26
   
   ' 讓與申請人
   '911113 nick
   'Modify By Sindy 2009/10/19 增加案件性質708
   If textCP10 = "701" Or textCP10 = "708" Then
      If frm010001.intModifyKind = 0 Then 'Add By Sindy 2009/10/19
        '若讓與人與原申請人不同時
        If ChangeCustomerL(m_CP55) <> ChangeCustomerL(m_PA26) Then
            '更新進度檔讓與人
            SetCPFieldNewData "CP55", ChangeCustomerL(m_PA26)
        End If
        '若讓與人與原申請人不同時
        If ChangeCustomerL(m_CP93) <> ChangeCustomerL(m_PA27) Then
            '更新進度檔讓與人
            SetCPFieldNewData "CP93", ChangeCustomerL(m_PA27)
        End If
        '若讓與人與原申請人不同時
        If ChangeCustomerL(m_CP94) <> ChangeCustomerL(m_PA28) Then
            '更新進度檔讓與人
            SetCPFieldNewData "CP94", ChangeCustomerL(m_PA28)
        End If
        '若讓與人與原申請人不同時
        If ChangeCustomerL(m_CP95) <> ChangeCustomerL(m_PA29) Then
            '更新進度檔讓與人
            SetCPFieldNewData "CP95", ChangeCustomerL(m_PA29)
        End If
        '若讓與人與原申請人不同時
        If ChangeCustomerL(m_CP96) <> ChangeCustomerL(m_PA30) Then
            '更新進度檔讓與人
            SetCPFieldNewData "CP96", ChangeCustomerL(m_PA30)
        End If
      End If
   End If
   
   ' 讓與申請人1
   If Not IsEmptyText(textCP56) Then
      SetCPFieldNewData "CP56", textCP56 & String(9 - Len(textCP56), "0")
   Else
      SetCPFieldNewData "CP56", Empty
   End If
   'Add By Sindy 2009/10/19
   ' 讓與申請人2
   If Not IsEmptyText(textCP89) Then
      SetCPFieldNewData "CP89", textCP89 & String(9 - Len(textCP89), "0")
   Else
      SetCPFieldNewData "CP89", Empty
   End If
   ' 讓與申請人3
   If Not IsEmptyText(textCP90) Then
      SetCPFieldNewData "CP90", textCP90 & String(9 - Len(textCP90), "0")
   Else
      SetCPFieldNewData "CP90", Empty
   End If
   ' 讓與申請人4
   If Not IsEmptyText(textCP91) Then
      SetCPFieldNewData "CP91", textCP91 & String(9 - Len(textCP91), "0")
   Else
      SetCPFieldNewData "CP91", Empty
   End If
   ' 讓與申請人5
   If Not IsEmptyText(textCP92) Then
      SetCPFieldNewData "CP92", textCP92 & String(9 - Len(textCP92), "0")
   Else
      SetCPFieldNewData "CP92", Empty
   End If
   '2009/10/19 End
   
   ' 進度備註
   SetCPFieldNewData "CP64", textCP64
   
   'Added by Morgan 2024/3/8
   SetCPFieldNewData "CP44", m_CP44
   SetCPFieldNewData "CP45", m_CP45
   'end 2024/3/8

   Select Case m_PA01
      ' 系統類別為CFT的為更新商標基本檔
      Case "P", "CFP", "FCP":
         ' 專利種類
         SetPASPFieldNewData "PA08", textPA08
         ' 申請國家
         SetPASPFieldNewData "PA09", textPA09
         ' 卷宗性質
         SetPASPFieldNewData "PA23", textPA23
         ' 是否PCT案件
         SetPASPFieldNewData "PA46", textPA46
         ' 分所案號
         SetPASPFieldNewData "PA47", textPA47
         ' 客戶案件案號
         SetPASPFieldNewData "PA48", textPA48
         ' 案件備註
         SetPASPFieldNewData "PA91", textPA91
      Case Else:
         ' 申請國家
         SetPASPFieldNewData "SP09", textPA09
         ' 案件備註
         SetPASPFieldNewData "SP18", textPA91
         ' 分所案號
         SetPASPFieldNewData "SP28", textPA47
         ' 客戶案件案號
         SetPASPFieldNewData "SP29", textPA48
   End Select
End Sub

' 新增案件進度檔
'Modify By Cheng 2002/11/06
'Private Sub SaveNewCaseProgress()
Private Function SaveNewCaseProgress() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
SaveNewCaseProgress = True

   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         strSql = strSql & m_CPList(nIndex).fiName
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   For nIndex = 0 To m_CPCount - 1
      If Not IsEmptyText(m_CPList(nIndex).fiNewData) Then
         If nIndex <> 0 Then strSql = strSql & ","
         If m_CPList(nIndex).fiType = 0 Then
            '911028 nick 加 chgsql
            'strSQL = strSQL & "'" & m_CPList(nIndex).fiNewData & "'"
            strSql = strSql & "'" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
         Else
            strSql = strSql & m_CPList(nIndex).fiNewData
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    SaveNewCaseProgress = False
End Function

' 更新專利基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdatePatent()
Private Function OnUpdatePatent() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean
        
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdatePatent = True

   ' 更新案件進度檔
   strSql = "UPDATE PATENT SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_PASPCount - 1
      strTmp = Empty
      If m_PASPList(nIndex).fiOldData <> m_PASPList(nIndex).fiNewData Then
         bDifference = True
         If m_PASPList(nIndex).fiType = 0 Then
            If m_PASPList(nIndex).fiNewData = Empty Then
               strTmp = m_PASPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_PASPList(nIndex).fiName & " = '" & ChgSQL(m_PASPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_PASPList(nIndex).fiNewData = Empty Then
               strTmp = m_PASPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_PASPList(nIndex).fiName & " = " & m_PASPList(nIndex).fiNewData
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
   strSql = strSql & " " & _
                  "WHERE PA01 = '" & m_PA01 & "' AND " & _
                        "PA02 = '" & m_PA02 & "' AND " & _
                        "PA03 = '" & m_PA03 & "' AND " & _
                        "PA04 = '" & m_PA04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdatePatent = False
End Function

' 更新服務業務基本檔的相關欄位
'Modify By Cheng 2002/11/06
'Private Sub OnUpdateServicePractice()
Private Function OnUpdateServicePractice() As Boolean
   Dim strTmp As String
   Dim strSql As String
   Dim nIndex As Integer
   Dim bFirst As Boolean
   Dim bDifference As Boolean

'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnUpdateServicePractice = True

   ' 更新案件進度檔
   strSql = "UPDATE ServicePractice SET "
   bFirst = True
   bDifference = False
   For nIndex = 0 To m_PASPCount - 1
      strTmp = Empty
      If m_PASPList(nIndex).fiOldData <> m_PASPList(nIndex).fiNewData Then
         If m_PASPList(nIndex).fiType = 0 Then
            If m_PASPList(nIndex).fiNewData = Empty Then
               strTmp = m_PASPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_PASPList(nIndex).fiName & " = '" & ChgSQL(m_PASPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_PASPList(nIndex).fiNewData = Empty Then
               strTmp = m_PASPList(nIndex).fiName & " = " & "NULL"
            Else
               strTmp = m_PASPList(nIndex).fiName & " = " & m_PASPList(nIndex).fiNewData
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
                  "WHERE SP01 = '" & m_PA01 & "' AND " & _
                        "SP02 = '" & m_PA02 & "' AND " & _
                        "SP03 = '" & m_PA03 & "' AND " & _
                        "SP04 = '" & m_PA04 & "' "
   ' 執行SQL指令
   If bDifference = True Then: cnnConnection.Execute strSql
'Add By Cheng 2002/11/06
Exit Function
ErrorHandler:
    OnUpdateServicePractice = False
End Function

'Modify By Cheng 2002/11/06
'Private Function OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strCF13 As String
   Dim strCF14 As String
   Dim strDay As String
   Dim strDate As String
   Dim strNP07 As String
   Dim strNP22 As String
   Dim strTemp As String
   Dim nIndex As Integer
   Dim nSubIndex As Integer
   Dim strCountry As String
   Dim objCopyPA As ClsCopyPA
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新案件進度檔
    'Modify By Cheng 2002/11/06
'   SaveNewCaseProgress
    'Modify By Cheng 2003/03/28
    '若為新增
    If frm010001.intModifyKind = 0 Then
        If SaveNewCaseProgress = False Then GoTo ErrorHandler
    '若為修改
    ElseIf frm010001.intModifyKind = 1 Then
        OnUpdateCaseProgress
    End If
   Select Case m_PA01
      ' 更新專利基本檔
      Case "P", "CFP", "FCP":
        'Modify By Cheng 2002/11/06
'         OnUpdatePatent
         If OnUpdatePatent = False Then GoTo ErrorHandler
      ' 更新服務業務基本檔
      Case Else:
        'Modify By Cheng 2002/11/06
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 機關文號
   For nIndex = 1 To GrdList.Rows - 1
      If GrdList.TextMatrix(nIndex, 0) = "V" Then
         If Not IsEmptyText(GrdList.TextMatrix(nIndex, 4)) Then
            strSql = "UPDATE CASEPROGRESS SET CP08 = '" & GrdList.TextMatrix(nIndex, 4) & "' " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strSql
         End If
         Exit For
      End If
   Next nIndex
    
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 儲存優先權資料
   'Modify by Amy 2014/04/18 +, m_Priority(5)
   If ClsPDSavePriority(m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5)) = False Then GoTo ErrorHandler

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新基本檔是否閉卷, 閉卷日期, 閉卷原因
   If textPA57 = "Y" Then
      Select Case m_PA01
      ' 更新專利基本檔
         Case "P", "CFP", "FCP":
            strSql = "UPDATE PATENT SET PA57=NULL, PA58=NULL,PA59=NULL " & _
                     "WHERE PA01 = '" & m_PA01 & "' AND " & _
                           "PA02 = '" & m_PA02 & "' AND " & _
                           "PA03 = '" & m_PA03 & "' AND " & _
                           "PA04 = '" & m_PA04 & "' "
         Case Else:
            strSql = "UPDATE SERVICEPRACTICE SET SP15=NULL, SP16=NULL,SP17=NULL " & _
                     "WHERE SP01 = '" & m_PA01 & "' AND " & _
                           "SP02 = '" & m_PA02 & "' AND " & _
                           "SP03 = '" & m_PA03 & "' AND " & _
                           "SP04 = '" & m_PA04 & "' "
      End Select
      cnnConnection.Execute strSql
   End If

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '91.11.10 cancel by sonia
   ' 更新案件進度檔的標準價與底價
   'strCF13 = "0"
   'strCF14 = "0"
   'strSQL = "SELECT * FROM CASEFEE " & _
   '         "WHERE CF01 = '" & m_PA01 & "' AND " & _
   '               "CF02 = '" & textPA09 & "' AND " & _
   '               "CF03 = '" & textCP10 & "' "
   'rsTmp.CursorLocation = adUseClient
   'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
   'If rsTmp.RecordCount > 0 Then
   '   rsTmp.MoveFirst
   '   If IsNull(rsTmp.Fields("CF13")) = False Then
   '      strCF13 = rsTmp.Fields("CF13")
   '   End If
   '   If IsNull(rsTmp.Fields("CF14")) = False Then
   '      strCF14 = rsTmp.Fields("CF14")
   '   End If
   'End If
   'rsTmp.Close
   ' 更新案件進度檔的標準價及底價欄位
   'strSQL = "UPDATE CaseProgress SET CP33 = " & strCF13 & ", " & _
   '                                 "CP34 = " & strCF14 & " " & _
   '         "WHERE CP09 = '" & m_CP09 & "' "
   'cnnConnection.Execute strSQL
   '91.11.10 end
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若案件性質為救濟程序時或爭議程序更新基本檔的欄位
   Select Case Mid(textCP10, 1, 1)
      ' 救濟程序
      Case "5":
         Select Case m_PA01:
            Case "P", "CFP", "FCP":
               strSql = "UPDATE PATENT SET PA18 = 'Y' " & _
                        "WHERE PA01 = '" & m_PA01 & "' AND " & _
                              "PA02 = '" & m_PA02 & "' AND " & _
                              "PA03 = '" & m_PA03 & "' AND " & _
                              "PA04 = '" & m_PA04 & "' "
               cnnConnection.Execute strSql
            Case Else:
         End Select
      ' 爭議程序
      Case "8":
         Select Case m_PA01:
            Case "P", "CFP", "FCP":
               strSql = "UPDATE PATENT SET PA19 = 'Y' " & _
                        "WHERE PA01 = '" & m_PA01 & "' AND " & _
                              "PA02 = '" & m_PA02 & "' AND " & _
                              "PA03 = '" & m_PA03 & "' AND " & _
                              "PA04 = '" & m_PA04 & "' "
               cnnConnection.Execute strSql
            Case Else:
         End Select
   End Select
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新使用者所選取的本案期限資料
   For nIndex = 1 To GrdList.Rows - 1
      ' 判斷該列是否有被選取
      If GrdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = GrdList.TextMatrix(nIndex, 9)
         strNP22 = GrdList.TextMatrix(nIndex, 10)
         'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(m_CP09) &
         strSql = "UPDATE NextProgress SET NP06 = 'Y',np24=" & CNULL(m_CP09) & _
                  " WHERE NP02 = '" & m_PA01 & "' AND " & _
                        "NP03 = '" & m_PA02 & "' AND " & _
                        "NP04 = '" & m_PA03 & "' AND " & _
                        "NP05 = '" & m_PA04 & "' AND " & _
                        "NP07 = " & strNP07 & " AND " & _
                        "NP22 = " & strNP22 & " "
         Pub_SeekTbLog strSql 'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業，若畫面勾選下一程序期限且存檔有上續辦Y的都寫Log以便事後能追蹤
         cnnConnection.Execute strSql
      End If
   Next nIndex
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '911018 nick 當有相關總收文號時，要將總收文號該筆更新成續辦，因為只會有一筆時才會讀出來秀畫面，所以不用np22
   '91.11.10 MODIFY BY SONIA
   'If textCP43 <> "" Then
   '     strSQL = "update nextprogress set np06='Y' where np01='" & textCP43 & "' "
   '     cnnConnection.Execute strSQL
   'End If
   '91.11.10 END
   
   ' 若申請國家為EPC時, 若此案為母案時
   If m_PA09 = "221" And m_PA03 = "0" And m_PA04 = "00" Then
      ' 系統類別為TF類
      Select Case m_PA01
         Case "P", "CFP", "FCP":
            If IsEmptyText(m_strCountry) = False Then
               For nSubIndex = 1 To GetSubStringCount(m_strCountry)
                  strCountry = GetSubString(m_strCountry, nSubIndex)
                  Set objCopyPA = New ClsCopyPA
                  objCopyPA.SetSrc m_PA01, m_PA02, m_PA03, m_PA04
                  objCopyPA.SetDes m_PA01, m_PA02, m_PA03, Format(CStr(Val(m_PA04) + nSubIndex), "00")
                  objCopyPA.SetExtraField "TM10", strCountry
                  objCopyPA.CopyPatent
                  Set objCopyPA = Nothing
               Next nSubIndex
            End If
      End Select
   End If
   
'   'Added by Morgan 2013/8/1
'   'P或CFP案收文主動修正203或修正204時若有相關新案已齊備未發文則清除完稿日及會稿日並EMail通知承辦人
'   If (m_PA01 = "CFP" Or m_PA01 = "P") And (m_CP10 = "203" Or m_CP10 = "204") Then
'      strExc(0) = "select cp09,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNo,cp14,ep09,ep07" & _
'         " from (select cm01,cm02,cm03,cm04 from casemap where cm10='0' and cm05='" & m_PA01 & "' and cm06='" & m_PA02 & "' and cm07='" & m_PA03 & "' and cm08='" & m_PA04 & "'" & _
'         " union select cm05,cm06,cm07,cm08 from casemap where cm10='0' and cm01='" & m_PA01 & "' and cm02='" & m_PA02 & "' and cm03='" & m_PA03 & "' and cm04='" & m_PA04 & "'" & _
'         " union select cr01,cr02,cr03,cr04 from caserelation where cr05='" & m_PA01 & "' and cr06='" & m_PA02 & "' and cr07='" & m_PA03 & "' and cr08='" & m_PA04 & "'" & _
'         "),caseprogress,engineerprogress where cp01(+)=cm01 and cp02(+)=cm02 and cp03(+)=cm03 and cp04(+)=cm04 and instr('" & NewCasePtyList & "',cp10)>0" & _
'         " and cp27||cp57 is null and ep02(+)=cp09 and ep06>0"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Do While Not RsTemp.EOF
'            If RsTemp("ep09") > 0 Then
'               'Modify By Sindy 2014/1/10 +EP12
'               'strSql = "update engineerprogress set ep09=null,ep07=null where ep02='" & RsTemp("cp09") & "'"
'               strSql = "update engineerprogress set ep09=null,ep07=null,ep12='" & ChangeTStringToTDateString(strSrvDate(2)) & "因相關案" & m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 = "000", "", "-" & m_PA03 & "-" & m_PA04) & "已收文" & textCP10_2 & "固清除相關日期,原完稿日：'||ep09||'原會稿日：'||ep07||';'||ep12 where ep02='" & RsTemp("cp09") & "'"
'               cnnConnection.Execute strSql, intI
'            End If
'            strExc(1) = RsTemp("Cno") & " 的相關案 " & m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 = "000", "", "-" & m_PA03 & "-" & m_PA04) & " 已收文" & textCP10_2 & "，請於2日內確認預定修正的內容..."
'            If IsNull(RsTemp("ep09")) Then
'               strExc(2) = "無"
'            Else
'               strExc(2) = TranslateKeyWord(incCNV_CHINESE_MINKO, RsTemp("ep09"), "")
'            End If
'            If IsNull(RsTemp("ep07")) Then
'               strExc(3) = "無"
'            Else
'               strExc(3) = TranslateKeyWord(incCNV_CHINESE_MINKO, RsTemp("ep07"), "")
'            End If
'            strExc(4) = RsTemp("Cno") & " 的相關案 " & m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 = "000", "", "-" & m_PA03 & "-" & m_PA04) & " 已收文" & textCP10_2 & "，請於2日內確認預定修正的內容，" & _
'               "若修正內容已實質改變原齊備內容，請修改齊備日；若修正內容未實質改變原齊備內容，請將原完稿日及原會稿日(若有)填入系統。" & _
'               vbCrLf & "原完稿日：" & strExc(2) & _
'               vbCrLf & "原會稿日：" & strExc(3)
'            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values ('" & strUserNum & "','" & RsTemp("cp14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & strExc(1) & "','" & strExc(4) & "')"
'            cnnConnection.Execute strSql, intI
'            RsTemp.MoveNext
'         Loop
'      End If
'   End If
'   'end 2013/8/1
   'Modify By Sindy 2016/4/13 抽出來變共用Func
   Call PUB_UpdRelationCaseFixEP(m_PA01, m_PA02, m_PA03, m_PA04, m_CP10, textCP10_2)
   '2016/4/13 END
   
   'Add by Sindy 2018/2/22
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_CP09, "")
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm010001", IIf(Left(Pub_StrUserSt03, 2) = "F2", m_CP09, "")
   End If
   '2018/2/22 END
   
   'Add By Cheng 2002/11/06
   cnnConnection.CommitTrans
   Exit Function
   
ErrorHandler:
   cnnConnection.RollbackTrans
   OnSaveData = False
End Function

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nickI As Integer
   Dim strCP14 As String
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
      
   m_PA57 = Empty
   
   ' 先清除商標基本檔或服務業務基本檔欄位串列
   ClearPASPFieldList
   ' 先清除案件進度檔欄位串列
   ClearCPFieldList
   
   m_CP05 = TAIWANDATE(SystemDate())
   textCP05 = m_CP05
   textCP10 = m_CP10
   textCP10_Validate False
   textCP56 = m_CP56
   textCP56_Validate False
   'Add By Sindy 2009/10/19
   textCP89 = m_CP89
   textCP89_Validate False
   textCP90 = m_CP90
   textCP90_Validate False
   textCP91 = m_CP91
   textCP91_Validate False
   textCP92 = m_CP92
   textCP92_Validate False
   '2009/10/19 End
   
   '2008/5/30 add by sonia
   '智權人員預設收文日最大的’A’或’B’類收文資料的智權人員
   textCP13 = PUB_GetAKindSalesNo(m_PA01, m_PA02, m_PA03, m_PA04)
   textCP13_Validate False
   
   ' 本所案號
   textPAKey = m_PA01 & m_PA02 & m_PA03 & m_PA04
      
   Select Case m_PA01
      ' 系統類別為CFT的為讀取商標基本檔
      Case "P", "CFP", "FCP":
         QueryPatent
      Case Else:
         QueryServicePractice
   End Select

   ' 取得案件進度檔的欄位
   '92.03.27 nick 修正
   If frm010001.intModifyKind = 0 Then
        QueryCaseProgressWithNewCP
   Else
        QueryCaseProgress
   End If
   

   ' 是否閉卷
   If m_PA57 = "Y" Then
      EnableTextBox textPA57, True
      textPA57_2 = "本案已閉卷"
   Else
      EnableTextBox textPA57, False
      textPA57_2 = Empty
   End If

   ' 依讀取的是專利基本檔還是服務業務基本檔來更新控制項的狀態
   Select Case m_PA01
      Case "P", "CFP", "FCP":
         EnableTextBox textPA08, True
         'EnableTextBox textPA14, True
      Case Else:
         EnableTextBox textPA08, False
         'EnableTextBox textPA14, False
   End Select

   ' 讀取優先權資料
   m_Pa(1) = m_PA01
   m_Pa(2) = m_PA02
   m_Pa(3) = m_PA03
   m_Pa(4) = m_PA04

   'Modify by Amy 2014/04/18 +, m_Priority(5)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5)
   
   ' 更新本案期限的資料
   UpdateGrdList m_PA01, m_PA02, m_PA03, m_PA04
   
   '911018 nick 新增時要待下一程序資料     本所期限，法定期限，收文號==>相關總收文號，備註==>進度備註    #只有一筆時，且本所案號和案件性質都要輸入且找的到
   If frm010001.intModifyKind = 0 Then
      If m_PA01 <> "" And m_PA02 <> "" And m_PA03 <> "" And m_PA04 <> "" And m_CP10 <> "" Then
          Dim nick911018rs As New ADODB.Recordset
          Dim nickstrsql As String
          Set nick911018rs = New ADODB.Recordset
          '911111 nick 邱小姐說要加入 np06 is null  np06<>'Y'(包含 null) 同意義
          'nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07=" & m_CP10 & " "
          '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
          'nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07=" & m_CP10 & " and (np06 <>'Y' or np06 is null) "
          nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07=" & m_CP10 & " and  np06 is null "
          nick911018rs.CursorLocation = adUseClient
          nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
          If nick911018rs.RecordCount = 1 Then
              textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
              textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
              textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
              textCP64 = textCP64 & CheckStr(nick911018rs.Fields("np15").Value)
              '91.11.10 ADD BY SONIA
              textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
              textCP13_Validate False
              '91.11.10 END
              '911030 nick 自動上勾
              For nickI = 1 To GrdList.Rows - 1
                  'edit by nick 2004/09/08
                  'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                  If Trim(GrdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(GrdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(GrdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                      GrdList.TextMatrix(nickI, 0) = "V"
                  End If
              Next nickI
          Else
              '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
              If nick911018rs.RecordCount = 0 Then
                  Set nick911018rs = New ADODB.Recordset
                  nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07=" & m_CP10 & " and np06 <>'Y'  "
                  nick911018rs.CursorLocation = adUseClient
                  nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
                  If nick911018rs.RecordCount = 1 Then
                      textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                      textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                      textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                      textCP64 = textCP64 & CheckStr(nick911018rs.Fields("np15").Value)
                      textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                      textCP13_Validate False
                      For nickI = 1 To GrdList.Rows - 1
                          'edit by nick 2004/09/08
                          'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                          If Trim(GrdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(GrdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(GrdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                              GrdList.TextMatrix(nickI, 0) = "V"
                          End If
                      Next nickI
                  End If
              End If
          End If
          strDate = "" 'Added by Morgan 2021/10/28 全域變數要先清除，否則會抓到殘留值
          'Add By Sindy 2020/3/9 936.回覆委任代理人 957.詢問代理人,本所期限預設3個工作天
          'If Val(textCP06) = 0 Then
             If Trim(Left(textCP10.Text, 4)) = "936" Or _
                Trim(Left(textCP10.Text, 4)) = "957" Then
                '本所期限預設3個工作天
                'Modified by Morgan 2020/7/30 +3個工作天--玫音
                'textCP06 = PUB_GetWorkDay1(CompDate(2, 3, strSrvDate(1)), True) - 19110000
                strDate = TransDate(CompWorkDay(3, CompDate(2, 1, strSrvDate(1)), 0), 1)
                'end 2020/7/30
             'Add By Sindy 2021/9/6 本所期限為收文日起算3個工作日--雅娟
             ElseIf textCP10 = "614" Then
                'strDate = TransDate(PUB_GetWorkDay1(PUB_GetWorkDayAfterSysDate(TransDate(textCP05, 2), 3), True), 1)
                strDate = TransDate(CompWorkDay(3, CompDate(2, 1, TransDate(textCP05, 2)), 0), 1)
             '2021/9/6 END
             End If
             
             'Modified by Morgan 2021/10/28 有日期才要設定,否則所限會被清除
             'If Val(strDate) < Val(textCP06) Or Val(textCP06) = 0 Then
             If strDate <> "" And (Val(strDate) < Val(textCP06) Or Val(textCP06) = 0) Then
             'end 2021/10/28
                textCP06 = strDate
             End If
          'End If
          '2020/3/9 END
          
          'Add By Sindy 2021/9/6 614.檢視核准版本, 預設此程序之工程師為最近一道程序的工程師。
         If textCP10 = "614" Then
            If PUB_GetCP14_P11(m_Pa, strCP14) = True Then
               textCP14 = strCP14
               textCP14_Validate True
            End If
         End If
         '2021/9/6 END
       End If
   End If
   
   ' 設定輸入的位置
   SetInputEntry
   '92.03.27 nick 當查詢時，將確定 disabled
   If frm010001.intModifyKind = 2 Then
      cmdOK.Enabled = False
   End If

End Sub

Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
     
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         textCP10 = rsTmp.Fields("CP10")
         textCP10_Validate False
      End If
      SetCPFieldOldData "CP10", textCP10, 0
      ' 收文日
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP05")) = False Then
         strTemp = rsTmp.Fields("CP05")
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      SetCPFieldOldData "CP05", strTemp, 1
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
      End If
      SetCPFieldOldData "CP06", textCP06, 1
      'Add By Cheng 2002/06/12
      m_strCP06 = "" & rsTmp.Fields("CP06")
      
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      End If
      SetCPFieldOldData "CP07", textCP07, 1
      'Add By Cheng 2002/06/12
      m_strCP07 = "" & rsTmp.Fields("CP07")
      ' 業務區
      SetCPFieldOldData "CP12", rsTmp.Fields("CP12"), 0
      
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         textCP13 = rsTmp.Fields("CP13")
         textCP13_Validate False
      End If
      SetCPFieldOldData "CP13", textCP13, 0
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_Validate False
      End If
      SetCPFieldOldData "CP14", textCP14, 0
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         textCP43 = rsTmp.Fields("CP43")
      End If
      SetCPFieldOldData "CP43", textCP43, 0
      ' 是否算案件數
      If IsNull(rsTmp.Fields("CP26")) = False Then
         textCP26 = rsTmp.Fields("CP26")
      End If
      SetCPFieldOldData "CP26", textCP26, 0
      '911018 nick 不要此欄位
      ' 取消收文日期
      'If IsNull(rsTmp.Fields("CP57")) = False Then
      '   textCP57 = TAIWANDATE(rsTmp.Fields("CP57"))
      'End If
      
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         textCP64 = rsTmp.Fields("CP64")
      End If
      SetCPFieldOldData "CP64", textCP64, 0
      ' 對造案件中文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP37")) = False Then
         strTemp = rsTmp.Fields("CP37")
      End If
      SetCPFieldOldData "CP37", strTemp, 0
      ' 對造案件英文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP38")) = False Then
         strTemp = rsTmp.Fields("CP38")
      End If
      SetCPFieldOldData "CP38", strTemp, 0
      ' 對造案件日文名稱
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP39")) = False Then
         strTemp = rsTmp.Fields("CP39")
      End If
      SetCPFieldOldData "CP39", strTemp, 0
      
      'Add By Sindy 2009/10/19
      m_CP55 = CheckStr(rsTmp.Fields("CP55"))
      m_CP93 = CheckStr(rsTmp.Fields("CP93"))
      m_CP94 = CheckStr(rsTmp.Fields("CP94"))
      m_CP95 = CheckStr(rsTmp.Fields("CP95"))
      m_CP96 = CheckStr(rsTmp.Fields("CP96"))
      SetCPFieldOldData "CP55", m_CP55, 0
      SetCPFieldOldData "CP93", m_CP93, 0
      SetCPFieldOldData "CP94", m_CP94, 0
      SetCPFieldOldData "CP95", m_CP95, 0
      SetCPFieldOldData "CP96", m_CP96, 0
      m_CP56 = CheckStr(rsTmp.Fields("CP56"))
      m_CP89 = CheckStr(rsTmp.Fields("CP89"))
      m_CP90 = CheckStr(rsTmp.Fields("CP90"))
      m_CP91 = CheckStr(rsTmp.Fields("CP91"))
      m_CP92 = CheckStr(rsTmp.Fields("CP92"))
      textCP56 = m_CP56
      textCP56_Validate False
      textCP89 = m_CP89
      textCP89_Validate False
      textCP90 = m_CP90
      textCP90_Validate False
      textCP91 = m_CP91
      textCP91_Validate False
      textCP92 = m_CP92
      textCP92_Validate False
      SetCPFieldOldData "CP56", m_CP56, 0
      SetCPFieldOldData "CP89", m_CP89, 0
      SetCPFieldOldData "CP90", m_CP90, 0
      SetCPFieldOldData "CP91", m_CP91, 0
      SetCPFieldOldData "CP92", m_CP92, 0
      
      ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
      If IsEmptyText(m_CP10) = False Then
         'modify by sonia 2023/10/30 加入textPA23 <> ""條件，否則服務業務會當掉
         If textPA23.Text <> "1" And textPA23 <> "" Then
            cmbPA05.Clear
            Set rsSubTmp = New ADODB.Recordset
            strSubSQL = "SELECT * FROM CaseProgress " & _
                        "WHERE CP01 = '" & m_PA01 & "' AND " & _
                              "CP02 = '" & m_PA02 & "' AND " & _
                              "CP03 = '" & m_PA03 & "' AND " & _
                              "CP04 = '" & m_PA04 & "' AND " & _
                              "CP31 = 'Y' "
            rsSubTmp.CursorLocation = adUseClient
            rsSubTmp.Open strSubSQL, cnnConnection, adOpenStatic, adLockReadOnly
            If rsSubTmp.RecordCount > 0 Then
               rsSubTmp.MoveFirst
               ' 對造案件中文名稱
               If IsNull(rsSubTmp.Fields("CP37")) = False Then
                  cmbPA05.AddItem rsSubTmp.Fields("CP37"), 0
                  SetCPFieldOldData "CP37", rsSubTmp.Fields("CP37"), 0
               Else
                  cmbPA05.AddItem "", 0
                  SetCPFieldOldData "CP37", "", 0
               End If
               ' 對造案件英文名稱
               If IsNull(rsSubTmp.Fields("CP38")) = False Then
                  cmbPA05.AddItem rsSubTmp.Fields("CP38"), 1
                  SetCPFieldOldData "CP38", rsSubTmp.Fields("CP38"), 0
               Else
                  cmbPA05.AddItem "", 1
                  SetCPFieldOldData "CP38", "", 0
               End If
               ' 對造案件日文名稱
               If IsNull(rsSubTmp.Fields("CP39")) = False Then
                  cmbPA05.AddItem rsSubTmp.Fields("CP39"), 2
                  SetCPFieldOldData "CP39", rsSubTmp.Fields("CP39"), 0
               Else
                  cmbPA05.AddItem "", 2
                  SetCPFieldOldData "CP39", "", 0
               End If
            End If
            rsSubTmp.Close
            Set rsSubTmp = Nothing
         End If
      End If

   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryPatent()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   '顯示本所案號
   textPAKey = m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
   
   strSql = "SELECT * FROM Patent " & _
            "WHERE PA01 = '" & m_PA01 & "' AND " & _
                  "PA02 = '" & m_PA02 & "' AND " & _
                  "PA03 = '" & m_PA03 & "' AND " & _
                  "PA04 = '" & m_PA04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 專利名稱(中)
      If Not IsNull(rsTmp.Fields("PA05")) Then
         cmbPA05.AddItem rsTmp.Fields("PA05")
         'm_PA05 = rsTmp.Fields("PA05")
      End If
      ' 專利名稱(英)
      If Not IsNull(rsTmp.Fields("PA06")) Then
         cmbPA05.AddItem rsTmp.Fields("PA06")
         'm_PA06 = rsTmp.Fields("PA06")
      End If
      ' 專利名稱(日)
      If Not IsNull(rsTmp.Fields("PA07")) Then
         cmbPA05.AddItem rsTmp.Fields("PA07")
         'm_PA07 = rsTmp.Fields("PA07")
      End If
      ' 顯示專利名稱
      If cmbPA05.ListCount > 0 Then
         cmbPA05.ListIndex = 0
      End If
      ' 專利種類
      If Not IsNull(rsTmp.Fields("PA08")) Then
         textPA08 = rsTmp.Fields("PA08")
         m_PA08 = rsTmp.Fields("PA08")
         textPA08_Validate False
      End If
      SetPASPFieldOldData "PA08", textPA08, 0
      ' 申請國家
      If Not IsNull(rsTmp.Fields("PA09")) Then
         textPA09 = rsTmp.Fields("PA09")
         m_PA09 = rsTmp.Fields("PA09")
         textPA09_Validate False
      End If
      SetPASPFieldOldData "PA09", textPA09, 0
      ' 券宗性質
      If Not IsNull(rsTmp.Fields("PA23")) Then
         textPA23 = rsTmp.Fields("PA23")
      End If
      SetPASPFieldOldData "PA23", textPA23, 1
      ' 申請人一
      If Not IsNull(rsTmp.Fields("PA26")) Then
         textPA26_2 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      ' 申請人二
      If Not IsNull(rsTmp.Fields("PA27")) Then
         textPA27_2 = GetCustomerName(rsTmp.Fields("PA27"), 0)
      End If
      ' 申請人三
      If Not IsNull(rsTmp.Fields("PA28")) Then
         textPA28_2 = GetCustomerName(rsTmp.Fields("PA28"), 0)
      End If
      ' 申請人四
      If Not IsNull(rsTmp.Fields("PA29")) Then
         textPA29_2 = GetCustomerName(rsTmp.Fields("PA29"), 0)
      End If
      ' 申請人五
      If Not IsNull(rsTmp.Fields("PA30")) Then
         textPA30_2 = GetCustomerName(rsTmp.Fields("PA30"), 0)
      End If
      
      'Add By Sindy 2009/10/19
      m_PA26 = "" & rsTmp.Fields("PA26")
      m_PA27 = "" & rsTmp.Fields("PA27")
      m_PA28 = "" & rsTmp.Fields("PA28")
      m_PA29 = "" & rsTmp.Fields("PA29")
      m_PA30 = "" & rsTmp.Fields("PA30")
      
      ' 是否PCT案件
      If Not IsNull(rsTmp.Fields("PA46")) Then
         textPA46 = rsTmp.Fields("PA46")
         m_PA46 = rsTmp.Fields("PA46")
      End If
      SetPASPFieldOldData "PA46", textPA46, 0
      ' 分所案件
      If Not IsNull(rsTmp.Fields("PA47")) Then
         textPA47 = rsTmp.Fields("PA47")
      End If
      SetPASPFieldOldData "PA47", textPA47, 0
      ' 客戶案件案號
      If Not IsNull(rsTmp.Fields("PA48")) Then
         textPA48 = rsTmp.Fields("PA48")
      End If
      SetPASPFieldOldData "PA48", textPA48, 0
      ' 代理人
      If Not IsNull(rsTmp.Fields("PA75")) Then
         textPA75_2 = GetFAgentName(rsTmp.Fields("PA75"))
      End If
      ' 案件備註
      If Not IsNull(rsTmp.Fields("PA91")) Then
         textPA91 = rsTmp.Fields("PA91")
      End If
      SetPASPFieldOldData "PA91", textPA91, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得案件進度檔的欄位內容
Private Sub QueryCaseProgressWithNewCP()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTemp As String
   Dim strSubSQL As String
   Dim rsSubTmp As ADODB.Recordset
   
   SetCPFieldOldData "CP01", Empty, 0
   SetCPFieldOldData "CP02", Empty, 0
   SetCPFieldOldData "CP03", Empty, 0
   SetCPFieldOldData "CP04", Empty, 0
   SetCPFieldOldData "CP09", Empty, 0
   ' 案件性質
   SetCPFieldOldData "CP10", Empty, 0
   ' 收文日
   SetCPFieldOldData "CP05", Empty, 1
   ' 本所期限
   SetCPFieldOldData "CP06", Empty, 1
   ' 法定期限
   SetCPFieldOldData "CP07", Empty, 1
   ' 業務區
   SetCPFieldOldData "CP12", Empty, 0
   ' 承辦人員
   SetCPFieldOldData "CP14", Empty, 0
   '911113 nick 邱小姐說刪除
   ' 費用
   'SetCPFieldNewData "CP16", textCP16
   ' 規費
   'SetCPFieldNewData "CP17", textCP17
   ' 點數
   'SetCPFieldNewData "CP18", textCP18
   
   ' 相關總收文號
   SetCPFieldOldData "CP43", Empty, 0
   ' 是否算案件數
   SetCPFieldOldData "CP26", Empty, 0
   ' 讓與申請人1
   SetCPFieldOldData "CP56", Empty, 0
   'Add By Sindy 2009/10/19
   ' 讓與申請人2
   SetCPFieldOldData "CP89", Empty, 0
   ' 讓與申請人3
   SetCPFieldOldData "CP90", Empty, 0
   ' 讓與申請人4
   SetCPFieldOldData "CP91", Empty, 0
   ' 讓與申請人5
   SetCPFieldOldData "CP92", Empty, 0
   m_CP55 = Empty
   m_CP93 = Empty
   m_CP94 = Empty
   m_CP95 = Empty
   m_CP96 = Empty
   SetCPFieldOldData "CP55", m_CP55, 0
   SetCPFieldOldData "CP93", m_CP93, 0
   SetCPFieldOldData "CP94", m_CP94, 0
   SetCPFieldOldData "CP95", m_CP95, 0
   SetCPFieldOldData "CP96", m_CP96, 0
   '2009/10/19 End
   
   '收據編號
   SetCPFieldOldData "CP60", Empty, 0
   ' 進度備註
   SetCPFieldOldData "CP64", Empty, 0
   
   '911104 nick 因為會有些值沒有先定義，所以會沒有更新
   SetCPFieldOldData "CP11", Empty, 0
   SetCPFieldOldData "CP13", Empty, 0
   SetCPFieldOldData "CP16", 0, 1
   SetCPFieldOldData "CP17", 0, 1
   SetCPFieldOldData "CP18", 0, 1
   SetCPFieldOldData "CP20", Empty, 0
   SetCPFieldOldData "CP21", Empty, 0
   '911113 nick
   SetCPFieldOldData "CP32", Empty, 0
   SetCPFieldOldData "CP27", Empty, 1
   
   'Added by Morgan 2024/3/8
   SetCPFieldOldData "CP44", Empty, 0
   SetCPFieldOldData "CP45", Empty, 0
End Sub

' 讀取服務業務基本檔
Private Sub QueryServicePractice()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_PA01 & "' AND " & _
                  "SP02 = '" & m_PA02 & "' AND " & _
                  "SP03 = '" & m_PA03 & "' AND " & _
                  "SP04 = '" & m_PA04 & "' "
                  
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 專利名稱(中)
      If Not IsNull(rsTmp.Fields("SP05")) Then
         cmbPA05.AddItem rsTmp.Fields("SP05")
         'm_PA05 = rsTmp.Fields("SP05")
      End If
      ' 專利名稱(英)
      If Not IsNull(rsTmp.Fields("SP06")) Then
         cmbPA05.AddItem rsTmp.Fields("SP06")
         'm_PA05 = rsTmp.Fields("SP06")
      End If
      ' 專利名稱(日)
      If Not IsNull(rsTmp.Fields("SP07")) Then
         cmbPA05.AddItem rsTmp.Fields("SP07")
         'm_PA05 = rsTmp.Fields("SP07")
      End If
      ' 顯示專利名稱
      If cmbPA05.ListCount > 0 Then
         cmbPA05.ListIndex = 0
      End If
      ' 申請國家
      If Not IsNull(rsTmp.Fields("SP09")) Then
         textPA09 = rsTmp.Fields("SP09")
         m_PA09 = rsTmp.Fields("SP09")
         textPA09_Validate False
      End If
      SetPASPFieldOldData "SP09", textPA09, 0
      ' 申請人一
      If Not IsNull(rsTmp.Fields("SP08")) Then
         textPA26_2 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      ' 申請人二
      If Not IsNull(rsTmp.Fields("SP58")) Then
         textPA27_2 = GetCustomerName(rsTmp.Fields("SP58"), 0)
      End If
      ' 申請人三
      If Not IsNull(rsTmp.Fields("SP59")) Then
         textPA28_2 = GetCustomerName(rsTmp.Fields("SP59"), 0)
      End If
      ' 申請人四
      If Not IsNull(rsTmp.Fields("SP65")) Then
         textPA29_2 = GetCustomerName(rsTmp.Fields("SP65"), 0)
      End If
      ' 申請人五
      If Not IsNull(rsTmp.Fields("SP66")) Then
         textPA30_2 = GetCustomerName(rsTmp.Fields("SP66"), 0)
      End If
      
      'Add By Sindy 2009/10/19
      m_PA26 = "" & rsTmp.Fields("SP08")
      m_PA27 = "" & rsTmp.Fields("SP58")
      m_PA28 = "" & rsTmp.Fields("SP59")
      m_PA29 = "" & rsTmp.Fields("SP65")
      m_PA30 = "" & rsTmp.Fields("SP66")
      
      ' 客戶案件案號
      If Not IsNull(rsTmp.Fields("SP29")) Then
         textPA48 = rsTmp.Fields("SP29")
      End If
      SetPASPFieldOldData "SP29", textPA48, 0
      ' 代理人
      If Not IsNull(rsTmp.Fields("SP26")) Then
         textPA75_2 = GetFAgentName(rsTmp.Fields("SP26"))
      End If
      ' 案件備註
      If Not IsNull(rsTmp.Fields("SP18")) Then
         textPA91 = rsTmp.Fields("SP18")
      End If
      SetPASPFieldOldData "SP18", textPA91, 0
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2018/2/23
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
         Set m_PrevForm = Nothing
      End If
   End If
   '2018/2/23 END
   
   Set frm010012_05 = Nothing
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If GrdList.row > 0 Then
         GrdList.col = 0
         If GrdList.Text = "V" Then
            GrdList.Text = Empty
         Else
             'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
             If Pub_CheckNpTheSameShow(m_PA01, textCP10, Trim("" & GrdList.TextMatrix(GrdList.row, 9))) = False Then
                 Exit Sub
             End If
             'end 2021/08/31
            'Modify by Morgan 2009/12/23 延期只更新期限不可點選
            'grdList.Text = "V"
            If textCP10 <> "404" Then
               GrdList.Text = "V"
            End If
            'End 2009/12/23
            
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                GrdList.col = 2
                textCP06 = GrdList.Text
                GrdList.col = 3
                textCP07 = GrdList.Text
                GrdList.col = 8
                textCP43 = GrdList.Text
                GrdList.col = 7
                textCP64 = textCP64 & GrdList.Text
             End If
             If textCP13.Text = "" Then
                GrdList.col = 11
                textCP13 = GrdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
         End If
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   If GrdList.row > 0 Then
      GrdList.col = 0
      If GrdList.Text = "V" Then
         GrdList.Text = Empty
      Else
             'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
             If Pub_CheckNpTheSameShow(m_PA01, textCP10, Trim("" & GrdList.TextMatrix(GrdList.row, 9))) = False Then
                 Exit Sub
             End If
             'end 2021/08/31
            'Modify by Morgan 2009/12/23 延期只更新期限不可點選
            'grdList.Text = "V"
            If textCP10 <> "404" Then
               GrdList.Text = "V"
            End If
            'End 2009/12/23
            
            '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
            '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
            '            智權人員沒值時，以勾的該筆代智權人員
            'If grdList.Row = 1 Then
             If textCP06.Text = "" Then
                GrdList.col = 2
                textCP06 = GrdList.Text
                GrdList.col = 3
                textCP07 = GrdList.Text
                GrdList.col = 8
                textCP43 = GrdList.Text
                GrdList.col = 7
                textCP64 = textCP64 & GrdList.Text
             End If
             If textCP13.Text = "" Then
                GrdList.col = 11
                textCP13 = GrdList.Text
                '911115 nick
                textCP13_2 = GetStaffName(textCP13)
             End If
            'End If
      End If
   End If
   grdList_ShowSelection
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = GrdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = GrdList.row Then
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      If GrdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To GrdList.Cols - 1
            GrdList.col = nCol
            If GrdList.CellBackColor <> &H80000005 Then: GrdList.CellBackColor = &H80000005
            If GrdList.CellForeColor <> &H80000008 Then: GrdList.CellForeColor = &H80000008
         Next nCol
      End If
      GrdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < GrdList.Rows Then
      GrdList.row = m_CurrSel
      GrdList.col = 1
      For nCol = 1 To GrdList.Cols - 1
         GrdList.col = nCol
         GrdList.CellBackColor = &H8000000D
         GrdList.CellForeColor = &H80000005
      Next nCol
      GrdList.col = 0
   End If
EXITSUB:
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP13_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 智權人員
Private Sub textCP13_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   'Added by Lydia 2019/02/14
   Dim m_SalesST15 As String '畫面上智權人員的收文部門
   Dim m_Tuser As String '創新業務部預設收文人員
   
   Cancel = False
   textCP13_2 = Empty
   If IsEmptyText(textCP13) = False Then
      textCP13_2 = GetStaffName(textCP13)
      If IsEmptyText(textCP13_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "智權人員代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0 'Added by Lydia 2019/02/14
         '911111 nick
         textCP13.SetFocus
         
         textCP13_GotFocus
      'Added by Lydia 2019/02/14 創新業務部人員收文控管
      Else
         m_SalesST15 = GetST15(textCP13)
         'Added by Lydia 2020/04/08 檢查案件或智權人員是否為法務部
         If PUB_ChkSalesL(m_PA01, textCP13.Text) = False Then
             SSTab1.Tab = 0
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
         'end 2020/04/08
         If PUB_ChkIsT10T20("2", textCP13.Text, m_Tuser, strTit) = True Then
             SSTab1.Tab = 0
             textCP13.Text = m_Tuser
             textCP13_2.Text = strTit
             textCP13.SetFocus
             Call textCP13_GotFocus
             Cancel = True
             Exit Sub
         End If
      'end 2019/02/14
      End If
   End If
End Sub

'Add By Sindy 2010/11/25
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

' 承辦人
Private Sub textCP14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   textCP14_2 = Empty
   If IsEmptyText(textCP14) = False Then
      textCP14_2 = GetStaffName(textCP14)
      If IsEmptyText(textCP14_2) = True Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP14.SetFocus
         
         textCP14_GotFocus
      End If
   '911111 nick 邱小姐說承辦人不可空白
   Else
         Cancel = True
         strTit = "資料檢核"
         strMsg = "承辦人不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP14.SetFocus
         
         textCP14_GotFocus
      
   End If
End Sub

' 案件性質
Private Sub textCP10_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   textCP10_2 = Empty
   Cancel = False
   If IsEmptyText(textCP10) = False Then
      If m_PA09 < "010" Then
         ' 取得國內的案件性質名稱
         textCP10_2 = GetCaseTypeName(m_PA01, textCP10, 0)
      Else
         textCP10_2 = GetCaseTypeName(m_PA01, textCP10, 1)
      End If
      If IsEmptyText(textCP10_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP10.SetFocus
         textCP10_GotFocus
      End If
   End If
   ' 讓與申請人欄位
   If Cancel = False Then
      '911113 nick 應該是 701
      'If textCP10 = "501" Then
      'Modify By Sindy 2009/10/19 增加案件性質708
      If textCP10 = "701" Or textCP10 = "708" Then
         Label36.Visible = True
         EnableTextBox textCP56, True
         textCP56.Visible = True
         textCP56_2.Visible = True
         'Add By Sindy 2009/10/19
         Label89.Visible = True
         EnableTextBox textCP89, True
         textCP89.Visible = True
         textCP89_2.Visible = True
         Label90.Visible = True
         EnableTextBox textCP90, True
         textCP90.Visible = True
         textCP90_2.Visible = True
         Label91.Visible = True
         EnableTextBox textCP91, True
         textCP91.Visible = True
         textCP91_2.Visible = True
         Label92.Visible = True
         EnableTextBox textCP92, True
         textCP92.Visible = True
         textCP92_2.Visible = True
         '2009/10/19 End
      Else
         Label36.Visible = False
         EnableTextBox textCP56, False
         textCP56.Visible = False
         textCP56_2.Visible = False
         'Add By Sindy 2009/10/19
         Label89.Visible = False
         EnableTextBox textCP89, False
         textCP89.Visible = False
         textCP89_2.Visible = False
         Label90.Visible = False
         EnableTextBox textCP90, False
         textCP90.Visible = False
         textCP90_2.Visible = False
         Label91.Visible = False
         EnableTextBox textCP91, False
         textCP91.Visible = False
         textCP91_2.Visible = False
         Label92.Visible = False
         EnableTextBox textCP92, False
         textCP92.Visible = False
         textCP92_2.Visible = False
         '2009/10/19 End
      End If
   End If
End Sub

Private Sub textCP21_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否為多國案
Private Sub textCP21_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP21) = False Then
      Select Case textCP21
         Case " ", "Y":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textCP21.SetFocus
            
            textCP21_GotFocus
      End Select
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否算案件數
Private Sub textCP26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   '911111 nick 當承辦人是 86032 or 85037 or 90009 時只能 N
   If IsEmptyText(textCP14) = False Then
      If Val(textCP14) = &H15010 Or Val(textCP14) = &H14C2D Or Val(textCP14) = &H15F99 Then
          If textCP26 <> "N" Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入 N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textCP26.SetFocus
            
            textCP26_GotFocus
            Exit Sub
          End If
      End If
   End If

   If IsEmptyText(textCP26) = False Then
      Select Case textCP26
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textCP26.SetFocus
            
            textCP26_GotFocus
      End Select
   End If
End Sub

' 收文日
Private Sub textCP05_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP05) = False Then
      If CheckIsTaiwanDate(textCP05, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "收文日日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP05.SetFocus
         
         textCP05_GotFocus
      End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "本所期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP06.SetFocus
         
         textCP06_GotFocus
      'Add by Morgan 2006/12/26
      Else
         '本所期限若非工作天則抓最近工作天
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      End If
   End If
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "法定期限的日期格式不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP07.SetFocus
         
         textCP07_GotFocus
      End If
   End If
End Sub

' 相關總收文號
Private Sub textCP43_Validate(Cancel As Boolean)
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP43) = False Then
      If textCP43 = m_CP09 Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號不可為本身之收文號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP43.SetFocus
         
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      
      strSql = "SELECT * FROM CaseProgress " & _
               "WHERE CP01 = '" & m_PA01 & "' AND " & _
                     "CP02 = '" & m_PA02 & "' AND " & _
                     "CP03 = '" & m_PA03 & "' AND " & _
                     "CP04 = '" & m_PA04 & "' AND " & _
                     "CP09 = '" & textCP43 & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         rsTmp.Close
         Cancel = True
         strTit = "資料檢核"
         strMsg = "相關總收文號資料不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textCP43.SetFocus
         
         textCP43_GotFocus
         GoTo EXITSUB
      End If
      rsTmp.Close
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub textCP56_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 讓與申請人1
Private Sub textCP56_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCP56 As String
   Dim strTemp As String
   
   Cancel = False
   If Not IsEmptyText(textCP56) Then
      strCP56 = textCP56
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP56, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetCustomerAndState(strCP56, strTemp, , , , m_PA01) Then
         textCP56 = strCP56 & String(9 - Len(strCP56), "0")
         textCP56_2 = strTemp
      Else
         Cancel = True
         '911111 nick
         textCP56.SetFocus
         
         textCP56_GotFocus
      End If
   End If
End Sub

Private Sub textCP89_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 讓與申請人2
Private Sub textCP89_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCP89 As String
   Dim strTemp As String
   
   Cancel = False
   If Not IsEmptyText(textCP89) Then
      strCP89 = textCP89
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP89, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetCustomerAndState(strCP89, strTemp, , , , m_PA01) Then
         textCP89 = strCP89 & String(9 - Len(strCP89), "0")
         textCP89_2 = strTemp
      Else
         Cancel = True
         '911111 nick
         textCP89.SetFocus
         
         textCP89_GotFocus
      End If
   End If
End Sub

Private Sub textCP90_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 讓與申請人3
Private Sub textCP90_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCP90 As String
   Dim strTemp As String
   
   Cancel = False
   If Not IsEmptyText(textCP90) Then
      strCP90 = textCP90
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP90, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetCustomerAndState(strCP90, strTemp, , , , m_PA01) Then
         textCP90 = strCP90 & String(9 - Len(strCP90), "0")
         textCP90_2 = strTemp
      Else
         Cancel = True
         '911111 nick
         textCP90.SetFocus
         
         textCP90_GotFocus
      End If
   End If
End Sub

Private Sub textCP91_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 讓與申請人4
Private Sub textCP91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCP91 As String
   Dim strTemp As String
   
   Cancel = False
   If Not IsEmptyText(textCP91) Then
      strCP91 = textCP91
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP91, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetCustomerAndState(strCP91, strTemp, , , , m_PA01) Then
         textCP91 = strCP91 & String(9 - Len(strCP91), "0")
         textCP91_2 = strTemp
      Else
         Cancel = True
         '911111 nick
         textCP91.SetFocus
         
         textCP91_GotFocus
      End If
   End If
End Sub

Private Sub textCP92_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 讓與申請人5
Private Sub textCP92_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strCP92 As String
   Dim strTemp As String
   
   Cancel = False
   If Not IsEmptyText(textCP92) Then
      strCP92 = textCP92
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If objPublicData.GetCustomer(strCP92, strTemp) Then
      'Modify By Sindy 2015/8/27 +m_PA01
      If GetCustomerAndState(strCP92, strTemp, , , , m_PA01) Then
         textCP92 = strCP92 & String(9 - Len(strCP92), "0")
         textCP92_2 = strTemp
      Else
         Cancel = True
         '911111 nick
         textCP92.SetFocus
         
         textCP92_GotFocus
      End If
   End If
End Sub

Private Sub textCP64_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP64, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "進度備註內容太長"
      '911111 nick
      textCP64.SetFocus
      
      textCP64_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textCP64.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

Private Sub textPA47_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPA47, 50) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "分所號內容太長"
      '911111 nick
      textPA47.SetFocus
      
      textPA47_GotFocus
   End If
End Sub

Private Sub textPA91_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textPA91, 2000) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "案件備註內容太長"
      '911111 nick
      textPA91.SetFocus
      
      textPA91_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   'If Cancel = False Then textPA91.IMEMode = 2
   If Cancel = False Then CloseIme
End Sub

' 專利種類
Private Sub textPA08_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textPA08_2 = Empty
   If IsEmptyText(textPA08) = False Then
      textPA08_2 = GetPatentName(textPA08, 0)
      If IsEmptyText(textPA08_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "專利種類不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA08.SetFocus
         
         textPA08_GotFocus
      End If
   End If
End Sub

' 申請國家
Private Sub textPA09_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textPA09_2 = Empty
   If IsEmptyText(textPA09) = False Then
      '911111 nick 邱小姐說不能 001~009
      If textPA09 >= "001" And textPA09 <= "009" Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA09.SetFocus
         
         textPA09_GotFocus
         Exit Sub
      End If
      textPA09_2 = GetNationName(textPA09, 0)
      If IsEmptyText(textPA09_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請國家代號不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA09.SetFocus
         
         textPA09_GotFocus
      End If
   End If
End Sub

' 券宗性質
Private Sub textPA23_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA23) = False Then
      If IsEmptyText(textCP10) = False Then
         Select Case textCP10
            ' 異議
            Case "801":
               If textPA23 <> "2" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textPA23.SetFocus
                  
                  textPA23_GotFocus
               End If
            ' 舉發
            Case "803":
               If textPA23 <> "3" Then
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textPA23.SetFocus
                  
                  textPA23_GotFocus
               End If
            Case Else:
               '91.11.10 CANCEL BY SONIA
               'If textPA23 <> "1" Then
               '   Cancel = True
               '   strTit = "檢核資料"
               '   strMsg = "卷宗性質不正確"
               '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               '   textPA23_GotFocus
               'End If
               '91.11.10 END
         End Select
      End If
   End If
End Sub

Private Sub textPA46_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否PCT案件
Private Sub textPA46_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA46) = False Then
      Select Case textPA46
         Case "Y", " ":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textPA46.SetFocus
            
            textPA46_GotFocus
      End Select
   End If
End Sub

Private Sub textPA57_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否取消閉卷
Private Sub textPA57_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textPA57) = False Then
      Select Case textPA57
         Case "Y", " ":
         Case Else
            Cancel = True
            strTit = "檢核資料"
            strMsg = "請輸入空白或Y"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textPA57.SetFocus
            
            textPA57_GotFocus
      End Select
   End If
End Sub
'911113 nick 邱小姐說刪除
' 費用
'Private Sub textCP16_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP16) = False Then
'      If IsNumeric(textCP16) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "費用為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP16.SetFocus
'
'         textCP16_GotFocus
'      End If
'   End If
'End Sub
'911113 nick 邱小姐說刪除
' 規費
'Private Sub textCP17_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP17) = False Then
'      If IsNumeric(textCP17) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "規費為數值資料"
 '        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP17.SetFocus
'
'         textCP17_GotFocus
'      End If
'   End If
'End Sub
'911113 nick 邱小姐說刪除
' 點數
'Private Sub textCP18_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If IsEmptyText(textCP18) = False Then
'      If IsNumeric(textCP18) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "點數為數值資料"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textCP18.SetFocus
'
'         textCP18_GotFocus
'      End If
'   End If
'End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim strTemp As String
Dim nResponse
Dim ii As Integer
   
   CheckDataValid = False
   ' 案件性質不可為空白
   If IsEmptyText(textCP10) = True Then
      strTit = "檢核資料"
      strMsg = "案件性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP10.SetFocus
      GoTo EXITSUB
   End If
   ' 申請國家不可空白
   If IsEmptyText(textPA09) = True Then
      strTit = "檢核資料"
      strMsg = "申請國家不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPA09.SetFocus
      GoTo EXITSUB
   End If
   '911113 nick
   '案件性質是  答辯 選取 公開費 延期 請求繼續審查 訴願 領證 年費 維持費  必須要輸入本所期限和法定期限
   '***** strat
   ' 案件性質為年費或延期時本所期限及法定期限不可為空白
   'If textCP10 = "605" Or textCP10 = "404" Then
   'Modified by Lydia 2016/08/26 +126 期末拋棄, +438 再考量試行計畫(AFCP2.0)
   If textCP10 = "107" Or textCP10 = "126" Or textCP10 = "438" Or textCP10 = "208" Or textCP10 = "217" Or textCP10 = "404" Or textCP10 = "424" Or textCP10 = "501" Or textCP10 = "601" Or textCP10 = "606" Or textCP10 = "607" Then
      If IsEmptyText(textCP06) = True Then
         strTit = "檢核資料"
         '911113 nick
         strMsg = "案件性質為年費或延期時, 本所期限不可為空白"
         strMsg = "本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         '911113 nick
         'strMsg = "案件性質為年費或延期時, 本所期限不可為空白"
         strMsg = "法定所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         GoTo EXITSUB
      End If
   End If
   '***** end
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限與法定期限範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 收文日
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "收文日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   ' 卷宗性質
   If IsEmptyText(textPA23) = True And m_PA01 = "CFP" Then
      strTit = "檢核資料"
      strMsg = "卷宗性質不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPA23.SetFocus
      GoTo EXITSUB
   End If
   ' 智權人員 ADD BY SONIA 91.11.3
   If IsEmptyText(textCP13) = True Then
      strTit = "檢核資料"
      strMsg = "智權人員不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP13.SetFocus
      GoTo EXITSUB
   End If
   '91.11.3 END
    'Add By Cheng 2003/08/13
    '若案件性質為延期, 則不可點選本案期限
    If Me.textCP10.Text = "404" Then
        For ii = 1 To Me.GrdList.Rows - 1
            If Me.GrdList.TextMatrix(ii, 0) <> "" Then
                MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
                GoTo EXITSUB
            End If
        Next ii
    End If
    
    m_CP44 = "": m_CP45 = "" 'Added by Morgan 2024/3/8
    
    'Added by Morgan 2023/12/13
    'CFP美國發明案=936(回覆委任代理人)或957(詢問代理人)內部收文/分案時，若曾發文IDS時必須要設相關總收文號以確保後續預設代理人時不會抓錯
    If m_PA01 = "CFP" And m_PA09 = "101" And m_PA08 = "1" And (textCP10 = "936" Or textCP10 = "957") Then
      If textCP43 = "" Then
         strExc(0) = "select cp44 from caseprogress where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "' and cp10='214' and cp27>0 and cp159=0 order by cp27 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            MsgBox "CFP美國發明案有IDS已發文時，" & textCP10_2 & "的【相關總收文號】不可空白！", vbExclamation
            textCP43.SetFocus
            GoTo EXITSUB
         End If
      'Added by Morgan 2024/3/8 若相關收文號為IDS時將CF代理人及彼號帶入以便發文時能正確預設
      Else
         strExc(0) = "select cp44,cp45 from caseprogress where cp09='" & textCP43 & "' and cp10='214'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            m_CP44 = "" & RsTemp("cp44")
            m_CP45 = "" & RsTemp("cp45")
         End If
      'end 2024/3/8
      End If
    End If
    'end 2023/12/13
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPA08_GotFocus()
   InverseTextBox textPA08
End Sub

Private Sub textPA09_GotFocus()
   InverseTextBox textPA09
End Sub

Private Sub textPA23_GotFocus()
   InverseTextBox textPA23
End Sub

Private Sub textPA46_GotFocus()
   InverseTextBox textPA46
End Sub

Private Sub textPA47_GotFocus()
   InverseTextBox textPA47
End Sub

Private Sub textPA48_GotFocus()
   InverseTextBox textPA48
End Sub

Private Sub textPA57_GotFocus()
   InverseTextBox textPA57
End Sub

Private Sub textPA91_GotFocus()
   InverseTextBox textPA91
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textPA91.IMEMode = 1
   OpenIme
End Sub

Private Sub textCP05_GotFocus()
   InverseTextBox textCP05
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
End Sub

Private Sub textCP10_GotFocus()
   InverseTextBox textCP10
End Sub

Private Sub textCP13_GotFocus()
   InverseTextBox textCP13
End Sub

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub
'911113 nick 邱小姐說刪除
'Private Sub textCP16_GotFocus()
'   InverseTextBox textCP16
'End Sub
'911113 nick 邱小姐說刪除
'Private Sub textCP17_GotFocus()
'   InverseTextBox textCP17
'End Sub
'Private Sub textCP18_GotFocus()
'   InverseTextBox textCP18
'End Sub

Private Sub textCP21_GotFocus()
   InverseTextBox textCP21
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP43_GotFocus()
   InverseTextBox textCP43
End Sub

Private Sub textCP56_GotFocus()
   InverseTextBox textCP56
End Sub
'Add By Sindy 2009/10/19
Private Sub textCP89_GotFocus()
   InverseTextBox textCP89
End Sub
Private Sub textCP90_GotFocus()
   InverseTextBox textCP90
End Sub
Private Sub textCP91_GotFocus()
   InverseTextBox textCP91
End Sub
Private Sub textCP92_GotFocus()
   InverseTextBox textCP92
End Sub
'2009/10/19 End

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP64.IMEMode = 1
   OpenIme
End Sub

Private Sub SetInputEntry()
   textCP14.SetFocus
End Sub

' 確認使用者所輸入的都完全正確
Private Function ValidateInput() As Boolean
   Dim Cancel As Boolean

   ValidateInput = False
   
   If textCP05.Enabled = True Then
      Cancel = False
      textCP05_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP10.Enabled = True Then
      Cancel = False
      textCP10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP13.Enabled = True Then
      Cancel = False
      textCP13_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP14.Enabled = True Then
      Cancel = False
      textCP14_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '911113 nick 邱小姐說刪除
   '***** START
   'If textCP16.Enabled = True Then
   '   Cancel = False
   '   textCP16_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   'If textCP17.Enabled = True Then
   '   Cancel = False
   '   textCP17_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   'If textCP18.Enabled = True Then
   '   Cancel = False
   '   textCP18_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   '***** end
   If textCP21.Enabled = True Then
      Cancel = False
      textCP21_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP26.Enabled = True Then
      Cancel = False
      textCP26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP43.Enabled = True Then
      Cancel = False
      textCP43_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textCP56.Enabled = True Then
      Cancel = False
      textCP56_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add By Sindy 2009/10/19
   If textCP89.Enabled = True Then
      Cancel = False
      textCP89_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textCP90.Enabled = True Then
      Cancel = False
      textCP90_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textCP91.Enabled = True Then
      Cancel = False
      textCP91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If textCP92.Enabled = True Then
      Cancel = False
      textCP92_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2009/10/19 End
   
   If textCP64.Enabled = True Then
      Cancel = False
      textCP64_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA08.Enabled = True Then
      Cancel = False
      textPA08_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA09.Enabled = True Then
      Cancel = False
      textPA09_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA23.Enabled = True Then
      Cancel = False
      textPA23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA46.Enabled = True Then
      Cancel = False
      textPA46_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA57.Enabled = True Then
      Cancel = False
      textPA57_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If

   If textPA91.Enabled = True Then
      Cancel = False
      textPA91_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   'Add by Amy 2018/10/18 智權人員非國外部FXX且修改案件性質時,不可改為 902(回覆代理人)
   If (m_PA01 = "CPS" Or m_PA01 = "CFP") And textCP10 = "902" Then
        If Left(PUB_GetStaffST15(textCP13, 1), 1) <> "F" Then
            MsgBox "智權人員非國外部，案件性質不可改為902(回覆代理人)"
            textCP10.SetFocus
            Exit Function
        End If
   End If
   'end 2018/10/18
   ValidateInput = True
End Function

'Add By Cheng 2003/03/28
' 更新案件進度檔
Private Sub OnUpdateCaseProgress()
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
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
            Else
               ' 91.03.25 modify by louis (單引號)
               'strTmp = m_CPList(nIndex).fiName & " = '" & m_CPList(nIndex).fiNewData & "'"
               strTmp = m_CPList(nIndex).fiName & " = '" & ChgSQL(m_CPList(nIndex).fiNewData) & "'"
            End If
         Else
            If m_CPList(nIndex).fiNewData = Empty Then
               strTmp = m_CPList(nIndex).fiName & " = " & "NULL"
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
