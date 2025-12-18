VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010012_06 
   BorderStyle     =   1  '單線固定
   Caption         =   "內部收文"
   ClientHeight    =   6430
   ClientLeft      =   900
   ClientTop       =   2210
   ClientWidth     =   9100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6430
   ScaleWidth      =   9100
   Begin VB.CommandButton cmdPriority 
      Caption         =   "優先權資料(&P)"
      Height          =   400
      Left            =   3060
      TabIndex        =   36
      Top             =   60
      Width           =   1300
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關案號(&F)"
      Height          =   400
      Left            =   4380
      TabIndex        =   34
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdCaseProgress 
      Caption         =   "案件進度(&C)"
      Height          =   400
      Left            =   5625
      TabIndex        =   33
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6855
      TabIndex        =   32
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7680
      TabIndex        =   31
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox textPAKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   540
      Width           =   1935
   End
   Begin VB.TextBox textPA57_2 
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000FF&
      Height          =   264
      Left            =   2940
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   540
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5145
      Left            =   120
      TabIndex        =   37
      Top             =   1200
      Width           =   8835
      _ExtentX        =   15575
      _ExtentY        =   9084
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "基本資料"
      TabPicture(0)   =   "frm010012_06.frx":0000
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
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label23(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label4"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(43)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(39)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(0)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(31)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(32)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(34)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(35)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(36)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(38)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(18)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label23(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label26"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label23(2)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label22"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP13_2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP14_2"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textPA26_2"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textPA27_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textPA28_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textPA29_2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textPA30_2"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textPA75_2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "grdList"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCP13"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textCP43"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textCP07"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textCP06"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textPA23"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP05"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCP26"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP10"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP10_2"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textPA48"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textPA57"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "textCP20"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "textCP14"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "textPA26"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "textPA27"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "textPA29"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "textPA28"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "textPA30"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "textPA75"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "textPA08"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "textPA08_2"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "textCP17"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "textCP16"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "textCP18"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).ControlCount=   59
      TabCaption(1)   =   "備註"
      TabPicture(1)   =   "frm010012_06.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cboAddCP64"
      Tab(1).Control(1)=   "textCP64"
      Tab(1).Control(2)=   "textPA91"
      Tab(1).Control(3)=   "lblAddCP64"
      Tab(1).Control(4)=   "Label2(3)"
      Tab(1).Control(5)=   "Label2(2)"
      Tab(1).Control(6)=   "Label2(1)"
      Tab(1).Control(7)=   "Label2(0)"
      Tab(1).Control(8)=   "Label1(9)"
      Tab(1).Control(9)=   "Label1(10)"
      Tab(1).ControlCount=   10
      TabCaption(2)   =   "移轉申請人"
      TabPicture(2)   =   "frm010012_06.frx":0038
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
         Left            =   -73500
         MaxLength       =   9
         TabIndex        =   28
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox textCP91 
         Height          =   264
         Left            =   -73500
         MaxLength       =   9
         TabIndex        =   27
         Top             =   1590
         Width           =   1095
      End
      Begin VB.TextBox textCP90 
         Height          =   264
         Left            =   -73500
         MaxLength       =   9
         TabIndex        =   26
         Top             =   1260
         Width           =   1095
      End
      Begin VB.TextBox textCP89 
         Height          =   264
         Left            =   -73500
         MaxLength       =   9
         TabIndex        =   25
         Top             =   930
         Width           =   1095
      End
      Begin VB.TextBox textCP56 
         Height          =   264
         Left            =   -73500
         MaxLength       =   9
         TabIndex        =   24
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox textCP18 
         Height          =   288
         Left            =   7575
         TabIndex        =   10
         Top             =   1590
         Width           =   1092
      End
      Begin VB.TextBox textCP16 
         Height          =   288
         Left            =   1080
         TabIndex        =   8
         Top             =   1590
         Width           =   1092
      End
      Begin VB.TextBox textCP17 
         Height          =   288
         Left            =   5040
         TabIndex        =   9
         Top             =   1590
         Width           =   1092
      End
      Begin VB.TextBox textPA08_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   5760
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   2520
         Width           =   2115
      End
      Begin VB.TextBox textPA08 
         Enabled         =   0   'False
         Height          =   270
         Left            =   5340
         MaxLength       =   1
         TabIndex        =   15
         Top             =   2517
         Width           =   375
      End
      Begin VB.TextBox textPA75 
         Height          =   264
         Left            =   5028
         MaxLength       =   8
         TabIndex        =   21
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox textPA30 
         Height          =   264
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   20
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox textPA28 
         Height          =   264
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   18
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox textPA29 
         Height          =   264
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   19
         Top             =   3120
         Width           =   1095
      End
      Begin VB.TextBox textPA27 
         Height          =   264
         Left            =   5040
         MaxLength       =   9
         TabIndex        =   17
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox textPA26 
         Height          =   264
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   16
         Top             =   2820
         Width           =   1095
      End
      Begin VB.TextBox textCP14 
         Height          =   270
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   0
         Top             =   420
         Width           =   855
      End
      Begin VB.TextBox textCP20 
         Height          =   270
         Left            =   5550
         MaxLength       =   8
         TabIndex        =   7
         Top             =   1290
         Width           =   630
      End
      Begin VB.TextBox textPA57 
         Height          =   270
         Left            =   5340
         MaxLength       =   1
         TabIndex        =   13
         Top             =   2217
         Width           =   375
      End
      Begin VB.TextBox textPA48 
         Height          =   270
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1920
         Width           =   1935
      End
      Begin VB.TextBox textCP10_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   423
         Width           =   2475
      End
      Begin VB.TextBox textCP10 
         Height          =   264
         Left            =   5328
         MaxLength       =   6
         TabIndex        =   1
         Top             =   423
         Width           =   732
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   1380
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2220
         Width           =   372
      End
      Begin VB.TextBox textCP05 
         Height          =   264
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   14
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox textPA23 
         Height          =   264
         Left            =   5340
         MaxLength       =   20
         TabIndex        =   3
         Top             =   720
         Width           =   372
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Top             =   990
         Width           =   1215
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5340
         MaxLength       =   7
         TabIndex        =   5
         Top             =   990
         Width           =   1095
      End
      Begin VB.TextBox textCP43 
         Height          =   264
         Left            =   1380
         MaxLength       =   9
         TabIndex        =   6
         Top             =   1290
         Width           =   2295
      End
      Begin VB.TextBox textCP13 
         Height          =   264
         Left            =   1080
         MaxLength       =   6
         TabIndex        =   2
         Top             =   720
         Width           =   852
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1380
         Left            =   168
         TabIndex        =   94
         Top             =   3720
         Width           =   8592
         _ExtentX        =   15152
         _ExtentY        =   2434
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
      Begin MSForms.ComboBox cboAddCP64 
         Height          =   300
         Left            =   -72615
         TabIndex        =   92
         TabStop         =   0   'False
         Top             =   360
         Width           =   6315
         VariousPropertyBits=   679479323
         DisplayStyle    =   7
         Size            =   "11139;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP92_2 
         Height          =   264
         Left            =   -72360
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1920
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
         Left            =   -72360
         TabIndex        =   88
         TabStop         =   0   'False
         Top             =   1590
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
         Left            =   -72360
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1260
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
         Left            =   -72360
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   930
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
         Left            =   -72360
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   600
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
      Begin MSForms.TextBox textPA75_2 
         Height          =   264
         Left            =   6180
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   3420
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
         Left            =   2220
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   3420
         Width           =   1815
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
         Left            =   6180
         TabIndex        =   70
         TabStop         =   0   'False
         Top             =   3120
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
         Left            =   2220
         TabIndex        =   68
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1815
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
         Left            =   6180
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2820
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
      Begin MSForms.TextBox textPA26_2 
         Height          =   264
         Left            =   2220
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   2820
         Width           =   1815
         VariousPropertyBits=   671107103
         ForeColor       =   -2147483641
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
         TabIndex        =   22
         Top             =   675
         Width           =   7245
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12779;1296"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textPA91 
         Height          =   735
         Left            =   -73560
         TabIndex        =   23
         Top             =   1470
         Width           =   7245
         VariousPropertyBits=   -1467987941
         ScrollBars      =   2
         Size            =   "12779;1296"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   1980
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   423
         Width           =   1935
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
         Left            =   1980
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   720
         Width           =   1935
         VariousPropertyBits=   671107103
         MaxLength       =   20
         Size            =   "3413;476"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label lblAddCP64 
         AutoSize        =   -1  'True
         Caption         =   "新增補件內容至進度備註"
         Height          =   180
         Left            =   -74760
         TabIndex        =   93
         Top             =   390
         Width           =   1980
      End
      Begin VB.Label Label92 
         Caption         =   "移轉申請人5 :"
         Height          =   195
         Left            =   -74700
         TabIndex        =   91
         Top             =   1950
         Width           =   1095
      End
      Begin VB.Label Label91 
         Caption         =   "移轉申請人4 :"
         Height          =   195
         Left            =   -74700
         TabIndex        =   89
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label90 
         Caption         =   "移轉申請人3 :"
         Height          =   195
         Left            =   -74700
         TabIndex        =   87
         Top             =   1290
         Width           =   1095
      End
      Begin VB.Label Label89 
         Caption         =   "移轉申請人2 :"
         Height          =   195
         Left            =   -74700
         TabIndex        =   85
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label36 
         Caption         =   "移轉申請人1 :"
         Height          =   195
         Left            =   -74700
         TabIndex        =   83
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label Label22 
         Caption         =   "費用:"
         Height          =   255
         Left            =   180
         TabIndex        =   81
         Top             =   1605
         Width           =   615
      End
      Begin VB.Label Label23 
         Caption         =   "點數:"
         Height          =   255
         Index           =   2
         Left            =   6930
         TabIndex        =   80
         Top             =   1605
         Width           =   660
      End
      Begin VB.Label Label26 
         Caption         =   "規費:"
         Height          =   255
         Left            =   4140
         TabIndex        =   79
         Top             =   1605
         Width           =   735
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不收)"
         Height          =   255
         Index           =   1
         Left            =   6300
         TabIndex        =   78
         Top             =   1305
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "專利種類"
         Height          =   180
         Index           =   18
         Left            =   4140
         TabIndex        =   77
         Top             =   2562
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人 :"
         Height          =   180
         Index           =   38
         Left            =   4140
         TabIndex        =   75
         Top             =   3462
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人5 :"
         Height          =   180
         Index           =   36
         Left            =   180
         TabIndex        =   73
         Top             =   3462
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人4 :"
         Height          =   180
         Index           =   35
         Left            =   4140
         TabIndex        =   71
         Top             =   3162
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人3 :"
         Height          =   180
         Index           =   34
         Left            =   180
         TabIndex        =   69
         Top             =   3162
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人2 :"
         Height          =   180
         Index           =   32
         Left            =   4140
         TabIndex        =   67
         Top             =   2862
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人1 :"
         Height          =   180
         Index           =   31
         Left            =   180
         TabIndex        =   65
         Top             =   2862
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   3
         Left            =   -72000
         TabIndex        =   61
         Top             =   3624
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   2
         Left            =   -72120
         TabIndex        =   60
         Top             =   2700
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   1
         Left            =   -72120
         TabIndex        =   59
         Top             =   1740
         Width           =   45
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Height          =   180
         Index           =   0
         Left            =   -72120
         TabIndex        =   58
         Top             =   720
         Width           =   45
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "承辦人"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   57
         Top             =   465
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "客戶案件案號"
         Height          =   180
         Index           =   4
         Left            =   180
         TabIndex        =   56
         Top             =   1965
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否向客戶收款"
         Height          =   180
         Index           =   39
         Left            =   4140
         TabIndex        =   55
         Top             =   1335
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "是否取消閉卷"
         Height          =   180
         Index           =   43
         Left            =   4140
         TabIndex        =   54
         Top             =   2262
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "進度備註"
         Height          =   180
         Index           =   9
         Left            =   -74760
         TabIndex        =   53
         Top             =   660
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註"
         Height          =   180
         Index           =   10
         Left            =   -74760
         TabIndex        =   52
         Top             =   1440
         Width           =   720
      End
      Begin VB.Label Label4 
         Caption         =   "案件性質 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   51
         Top             =   428
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不算)"
         Height          =   255
         Index           =   0
         Left            =   1980
         TabIndex        =   50
         Top             =   2225
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   180
         TabIndex        =   49
         Top             =   2225
         Width           =   1215
      End
      Begin VB.Label Label37 
         Caption         =   "收文日 :"
         Height          =   255
         Left            =   180
         TabIndex        =   48
         Top             =   2525
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "(1:申請 2:異議 3:舉發)"
         Height          =   255
         Left            =   5820
         TabIndex        =   47
         Top             =   725
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "卷宗性質 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   46
         Top             =   725
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   180
         TabIndex        =   45
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4140
         TabIndex        =   44
         Top             =   990
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "相關總收文號 :"
         Height          =   255
         Left            =   180
         TabIndex        =   43
         Top             =   1305
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "(Y : 取消閉卷)"
         Height          =   255
         Left            =   6000
         TabIndex        =   42
         Top             =   2225
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "智權人員 :"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   41
         Top             =   725
         Width           =   975
      End
   End
   Begin MSForms.ComboBox cmbPA05 
      Height          =   300
      Left            =   930
      TabIndex        =   35
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號"
      Height          =   180
      Index           =   6
      Left            =   120
      TabIndex        =   63
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱"
      Height          =   180
      Index           =   7
      Left            =   120
      TabIndex        =   62
      Top             =   840
      Width           =   720
   End
End
Attribute VB_Name = "frm010012_06"
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
' 案件名稱
Dim m_PA05 As String
Dim m_PA06 As String
Dim m_PA07 As String
' 專利種類
Dim m_PA08 As String

' 收文日
Dim m_CP05 As String
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
Dim m_PA16 As String   '2010/8/17 add by sonia
Dim m_PA14 As String   '2010/8/17 add by sonia

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

'911113 nick
Dim strNickPa01 As String
Dim strNickPa09 As String
Dim m_strCP06 As String '原本所期限
Dim m_strCP07 As String '原法定期限
Dim m_str945CP09 As String '要發文的945收文號 Added by Morgan 2012/4/18
Dim m_CP118 As String  'Added by Lydia 2017/12/14 記錄是否電子送件
Dim m_strCPM34 As String 'Add By Sindy 2021/4/29
'Dim m_bolFMP As Boolean 'Add By Sindy 2022/3/31
'Add By Sindy 2022/6/28
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_PrevForm As Form '前一畫面
'2022/6/28 END
Dim m_PA177 As String 'Added by Lydia 2023/07/28 FCP專利連結通知
Dim m_strCP27 As String 'Add By Sindy 2025/7/29


Private Sub ClearAll()
   ClearPASPFieldList
   ClearCPFieldList
   
   textPAKey = Empty
   textPA08 = Empty
   '911113 nick 邱小姐說刪除
   'textPA14 = Empty
   textPA23 = Empty
   textPA26 = Empty
   textPA27 = Empty
   textPA28 = Empty
   textPA29 = Empty
   textPA30 = Empty
   textPA48 = Empty
   textPA57 = Empty
   textPA75 = Empty
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
   '911113 ncik add
   textCP20 = Empty
   
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
   
   'Add by Morgan 2004/8/10
   textCP16 = Empty
   textCP17 = Empty
   textCP18 = Empty
   'Add end

   m_strCountry = Empty
   
   m_PA177 = "" 'Added by Lydia 2023/07/28
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
         'add by sonia 2017/10/18B類其他翻譯927且承辦人為外翻編號且相關總收文號為C類,預設進度備註
         If Left(textCP43, 1) = "C" And textCP10 = "927" And Left(textCP14, 1) = "F" Then
            If textCP64 = "OA委外翻譯" Then
            Else
               textCP64 = "OA委外翻譯;" & textCP64
            End If
         End If
         'end 2017/10/18
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

Private Sub cboAddCP64_Click()
   textCP64 = textCP64 & IIf(textCP64 = "", "", ", ") & cboAddCP64
   textCP64.SelStart = Len(textCP64)
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
   'frm010012_03.SetParent "frm010012_06"
   frm010012_03.SetParent Me
   Me.Hide
   frm010012_03.Show
   frm010012_03.QueryData
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDCase(1 To 4)  As String  'Added by Lydia 2022/07/04
   
   If CheckDataValid = True Then
      If ValidateInput() = False Then
         Exit Sub
      End If
      'Added by Lydia 2015/02/04 所有內部收文, 若有輸入本所期限或法定期限者, 檢查期限不可小於系統日
      'Modified by Lydia 2017/07/31 改為預設和檢查
      'If PUB_CheckCP0607(0, textCP06.Text, textCP07.Text) = False Then Exit Sub
      'Modified by Lyddia 2023/11/08 傳入必需欄位
      'If PUB_CheckCP0607(0, textCP06, textCP07) = False Then Exit Sub
      If PUB_CheckCP0607(0, textCP06, textCP07, "", m_PA09, m_PA01, textCP10) = False Then Exit Sub
      
      'Add By Sindy 2021/4/29 主管機關期限
      CheckOC3
      m_strCPM34 = ""
      strSql = "select cpm34 from casepropertymap where cpm01='" & m_PA01 & "' and cpm02='" & textCP10 & "'"
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount > 0 Then
         m_strCPM34 = "" & AdoRecordSet3.Fields(0)
      End If
      '2021/4/29 END
      
      'Add By Sindy 2021/6/23
      If (m_PA01 = "FCP" Or m_PA01 = "FG") And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
         If m_strCPM34 = "Y" And Val(textCP07) = 0 Then
            If MsgBox("此案件性質屬有主管機關期限，確定沒有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               textCP07.SetFocus
               Exit Sub
            End If
         ElseIf m_strCPM34 = "N" And Val(textCP07) > 0 Then
            If MsgBox("此案件性質屬非主管機關期限，確定有法定期限嗎？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
               textCP07.SetFocus
               Exit Sub
            End If
         End If
      End If
      '2021/6/23 END
      
      'Added by Lydia 2022/07/04 一案兩請僅其中一案收文701~709、401變更時，於確認收文時彈提醒：為一案兩請，請確認發明案及新型案是否一併收文。
      If frm010001.intModifyKind = 0 And (m_PA01 = "FCP" Or m_PA01 = "P") And (Left(textCP10, 2) = "70" Or textCP10 = "401") Then
          strExc(0) = ""
          If m_PA01 = "P" Then
              If PUB_ChkIsFMP(m_PA01, m_PA02, m_PA03, m_PA04) = False Then
                  strExc(0) = "N"
              End If
          End If
          If strExc(0) <> "N" Then
              If PUB_IsDualApply(m_Pa, strDCase, , , , , , True) = True Then
                 If MsgBox(m_PA01 & "-" & m_PA02 & IIf(m_PA03 & m_PA04 <> "000", "-" & m_PA03 & "-" & m_PA04, "") & "為一案兩請，請確認發明案及新型案是否一併收文，" & vbCrLf & _
                       "另一案件：" & strDCase(1) & "-" & strDCase(2) & IIf(strDCase(3) & strDCase(4) <> "000", "-" & strDCase(3) & "-" & strDCase(4), "") & vbCrLf & vbCrLf & _
                        "選擇""是""會繼續作業，選擇""否""會中斷作業。", vbYesNo + vbInformation + vbDefaultButton2) = vbNo Then
                       Exit Sub
                 End If
              End If
          End If
      End If
      'end 2022/07/04
      
      'Add By Sindy 2025/7/29
      '針對轉案進來的案件，若內部收文日為111111'案件性質為[（107）再審申請]，則
      '按確定後'彈出訊息：請輸入再審查送件日期,將日期回寫到進度檔的[再審查申請]的發文日
      If m_PA01 = "FCP" And textCP10 = "107" And textCP05 = "111111" Then
input_CP27:
         m_strCP27 = InputBox("請輸入再審查送件日期！" & vbCrLf & vbCrLf & _
                            "※若不確定正確之送件日期，請輸入大概之日期即可（僅供判斷新、舊法）", , m_strCP27)
         If Trim(m_strCP27) = "" Then
            MsgBox "再審查送件日期不可空白！", , "檢核資料", vbInformation
            GoTo input_CP27
         Else
            If CheckIsTaiwanDate(m_strCP27, False) = False Then
               MsgBox "再審查送件日期(" & m_strCP27 & ")，日期格式不正確！", vbExclamation, "檢核資料"
               GoTo input_CP27
            End If
         End If
      End If
      '2025/7/29 END
      
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
      
      'Added by Lydia 2023/02/08 內部收文補收款，智權人員為SXX部門時，要發MAIL給杜協理及智權人員
      If (m_PA01 = "P" Or m_PA01 = "PS" Or m_PA01 = "CFP" Or m_PA01 = "CPS") And textCP10 <> "" And InStr(textCP10_2, "補收款") > 0 And Left(GetST15(textCP13), 1) = "S" Then
          strExc(0) = m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
          strExc(1) = "本所案號：" & strExc(0) & vbCrLf & _
                           "案件名稱：" & m_PA05 & vbCrLf & _
                           "申請人1：" & m_PA26 & " " & textPA26_2 & vbCrLf & _
                           "申請國家：" & GetPrjNationName(m_PA09) & vbCrLf & _
                           "補收款費用：" & Val(textCP16) & vbCrLf & _
                           "補收款備註：" & Trim(textCP64)
          strExc(2) = Pub_GetSpecMan("全所智權部主管")
          If InStr(strExc(2), textCP13) = 0 Then
              strExc(2) = strExc(2) & ";" & textCP13
          End If
          PUB_SendMail strUserNum, strExc(2), "", strExc(0) & "內部收文補收款通知!", strExc(1)
      End If
      'end 2023/02/08
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      PUB_SendMailCache 'Added by Lydia 2023/07/28
      
      'Modify By Sindy 2022/6/28 信件內部收文執行完畢後,關閉視窗
      If m_strIR01 <> "" Then
         'Modify By Sindy 2022/6/29
         If Pub_StrUserSt03 = "F23" Then
            Call PUB_RecvOutLookF23(m_strIR01, m_strIR02, m_strIR03, m_strIR04, "1", m_CP09)
         End If
         '2022/6/29 END
         Unload frm010001
         Unload Me
      'Added by Lydia 2018/02/01 FCP客戶提供文件處理要進入內部收文
      'Modified by Lydia 2021/02/22 改判斷
      'If TypeName(frm010001.mPrevForm) = "frm060121_1" Then
      '     frm010001.Tag = m_CP09
      ElseIf frm010001.m_GetB202CP09 <> "" Then
           frm010001.m_GetB202CP09 = m_CP09
      'end 2021/02/22
           Unload frm010001
           Unload Me
      Else
      'end 2018/02/01
            ' 回到收文的畫面
            frm010001.SetData m_CP09, 0, True
            frm010001.SetData m_PA01, 1, False
            frm010001.SetData m_PA02, 2, False
            frm010001.SetData m_PA03, 3, False
            frm010001.SetData m_PA04, 4, False
            frm010001.Show
            ClearAll
            Unload Me
      End If 'end 2018/02/01
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
   
   'add by sonia 2017/10/18
   textCP13 = PUB_GetFCPSalesNo(m_PA01, m_PA02, m_PA03, m_PA04)
   textCP13_Validate False
   'end 2017/10/18
   
   'Added by Lydia 2018/08/17 預設承辦人=操作人員(by 敏莉)
   textCP14 = strUserNum
   textCP14_Validate False
   'end 2018/08/17
   
   'Add By Sindy 2022/6/28
   m_strIR01 = frm010001.m_strIR01
   m_strIR02 = frm010001.m_strIR02
   m_strIR03 = frm010001.m_strIR03
   m_strIR04 = frm010001.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2022/6/28 END
End Sub


' 初始化 GridList
Private Sub InitialGrdList()
   grdList.Clear
   grdList.Rows = 1
   '911111 nick
   'grdList.Cols = 11
   grdList.Cols = 12
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
   grdList.Text = "解除期限日"
   grdList.ColWidth(6) = 1200
   grdList.col = 7
   grdList.Text = "備註"
   grdList.ColWidth(7) = 1200
   grdList.col = 8
   grdList.Text = "收文號"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "下一程序代號"
   grdList.ColWidth(9) = 0
   grdList.col = 10
   grdList.Text = "序號"
   grdList.ColWidth(10) = 0
   '911111 nick add
   grdList.col = 11
   grdList.Text = "序號"
   grdList.ColWidth(11) = 0
End Sub

Private Sub UpdateGrdList(ByVal strPA01 As String, ByVal strPA02 As String, ByVal strPA03 As String, ByVal strPA04 As String)
   Dim nIndex As Integer
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 本案期限
   InitialGrdList
   
   'Modify by Morgan 2009/12/23 下一程序要排除程序管制的案件性質
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   'Modify by Morgan 2011/6/10 排除程序管制的案件性質改用 strNpSqlOfNoSalesDuty 常數
   'Modified by Lydia 2016/01/08 FCP內部收文案件性質為補文件202時,畫面下方之下一程序資料只帶補文件的期限.
'   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
'            "WHERE NP02 = '" & strPA01 & "' AND " & _
'                  "NP03 = '" & strPA02 & "' AND " & _
'                  "NP04 = '" & strPA03 & "' AND " & _
'                  "NP05 = '" & strPA04 & "' AND " & _
'                  "(NP06 IS NULL OR NP06 <> 'Y') " & strNpSqlOfNoSalesDuty
   'Modified by Lydia 2023/12/18 202補文件包含231國外寄存與存活證明 AND NP07='202'=> AND NP07 in ('202','231') ; ex.FCP-70830客戶提供文件在處理時無法選擇下一程序沖銷
   strSql = "SELECT NP01,NP07,NP08,NP09,NP10,NP11,NP12,NP13,NP14,NP15,NP22 FROM NextProgress " & _
            "WHERE NP02 = '" & strPA01 & "' AND " & _
                  "NP03 = '" & strPA02 & "' AND " & _
                  "NP04 = '" & strPA03 & "' AND " & _
                  "NP05 = '" & strPA04 & "' AND " & _
                  "(NP06 IS NULL OR NP06 <> 'Y') " & IIf(m_PA01 = "FCP" And textCP10 = "202", "AND NP07 in ('202','231') ", strNpSqlOfNoSalesDuty)
   
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
         grdList.Rows = grdList.Rows + 1
         nIndex = grdList.Rows - 1
         
         ' 收文號
         If IsNull(rsTmp.Fields("NP01")) = False Then
            grdList.TextMatrix(nIndex, 8) = rsTmp.Fields("NP01")
         End If
         ' 下一程序
         If IsNull(rsTmp.Fields("NP07")) = False Then
            '911111 nick 案件性質要依國家判斷
            'grdList.TextMatrix(nIndex, 1) = GetCaseTypeName(m_PA01, rsTmp.Fields("NP07"))
            grdList.TextMatrix(nIndex, 1) = GetPrjState4(strPA01 & "-" & strPA02 & "-" & strPA03 & "-" & strPA04, rsTmp.Fields("NP07"))
            
            grdList.TextMatrix(nIndex, 9) = rsTmp.Fields("NP07")
         End If
         ' 本所期限
         If IsNull(rsTmp.Fields("NP08")) = False Then
            If IsEmptyText(rsTmp.Fields("NP08")) = False Then
               grdList.TextMatrix(nIndex, 2) = ChangeWStringToTString(rsTmp.Fields("NP08"))
            End If
         End If
         ' 法定期限
         If IsNull(rsTmp.Fields("NP09")) = False Then
            If IsEmptyText(rsTmp.Fields("NP09")) = False Then
               grdList.TextMatrix(nIndex, 3) = ChangeWStringToTString(rsTmp.Fields("NP09"))
            End If
         End If
         ' 機關文號
         If IsNull(rsTmp.Fields("NP13")) = False Then
            grdList.TextMatrix(nIndex, 4) = rsTmp.Fields("NP13")
         End If
         ' 相關人
         If IsNull(rsTmp.Fields("NP14")) = False Then
            grdList.TextMatrix(nIndex, 5) = rsTmp.Fields("NP14")
         End If
         ' 解除期限日期
         If IsNull(rsTmp.Fields("NP11")) = False Then
            grdList.TextMatrix(nIndex, 6) = rsTmp.Fields("NP11")
         End If
         ' 備註
         If IsNull(rsTmp.Fields("NP15")) = False Then
            grdList.TextMatrix(nIndex, 7) = rsTmp.Fields("NP15")
         End If
         ' 序號
         If IsNull(rsTmp.Fields("NP22")) = False Then
            grdList.TextMatrix(nIndex, 10) = rsTmp.Fields("NP22")
         End If
         '911111 nick 智權人員
         If IsNull(rsTmp.Fields("NP10")) = False Then
            grdList.TextMatrix(nIndex, 11) = rsTmp.Fields("NP10")
         End If
NextRecord:
         rsTmp.MoveNext
      Loop
      'Added by Lydia 2023/10/16
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
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
Dim strCP64 As String 'Add By Sindy 2021/4/29

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
   
   ' 收文號
   'Modify by Morgan 2004/2/18
   '新增才要重抓收文號
    If frm010001.intModifyKind = 0 Then
        m_CP09 = AutoNo("B", 6)
    End If
    
   SetCPFieldNewData "CP09", m_CP09
   ' 案件性質
   SetCPFieldNewData "CP10", textCP10
   
   '911113 nick
   '***** start
   SetCPFieldNewData "CP11", "90"
   SetCPFieldNewData "CP20", textCP20
   SetCPFieldNewData "CP32", textCP20 'Add By Sindy 2016/6/30
   
   'Modify by Morgan 2004/8/10
'   SetCPFieldNewData "CP16", Empty
'   SetCPFieldNewData "CP17", Empty
'   SetCPFieldNewData "CP18", Empty
   SetCPFieldNewData "CP16", textCP16
   SetCPFieldNewData "CP17", textCP17
   SetCPFieldNewData "CP18", textCP18
   
   '911113 nick 承辦期限
   Dim m_strCP48 As String
   m_strCP48 = strSrvDate(1)
    'Modify By Cheng 2003/09/01
'   m_strCP48 = DBDATE(Format(DateSerial(Val(DBYEAR(m_strCP48)), Val(DBMONTH(m_strCP48)), Val(DBDAY(m_strCP48)) + GetCF04(strNickPa01, strNickPa09, textCP10))))
'edit by nickc 2007/10/11 改抓有時效性的
'''''   m_strCP48 = DBDATE(DateAdd("d", GetCF04(strNickPa01, strNickPa09, textCP10), ChangeWStringToWDateString(DBDATE(m_strCP48))))
'''''   If DBDATE(textCP06) <> "" Then
'''''      If m_strCP48 > DBDATE(textCP07) Then m_strCP48 = DBDATE(textCP07)
'''''   End If
'   'Add By Sindy 2022/3/31
'   If m_bolFMP = True Then
'      m_strCP48 = Pub_GetHandleDay("FCP", strNickPa09, textCP10, , DBDATE(textCP06))
'   '2022/3/31 END
'   Else
   If Not (strNickPa01 = "FCP" And InStr(SkipCasePtyList, textCP10) > 0) Then
      'Added by Morgan 2012/8/1
      '加速審查要判斷已輸入通知實審日才掛承辦期限
      'Modified by Morgan 2024/11/18 +477再審查加速審查並改用專用模組判斷
      If strNickPa01 = "FCP" And (textCP10 = "422" Or textCP10 = "447") Then
         strExc(1) = m_PA01
         strExc(2) = m_PA02
         strExc(3) = m_PA03
         strExc(4) = m_PA04
         'If PUB_ChkCPExist(strExc(), "1204") Then
         If PUB_Chk1204(strExc()) Then
            'Modify By Sindy 2021/7/23 + , , , m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
            m_strCP48 = Pub_GetHandleDay(strNickPa01, strNickPa09, textCP10, , DBDATE(textCP07), , , m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04)
         Else
            m_strCP48 = ""
         End If
      Else
      'end 2012/8/1
         'Modify By Sindy 2021/7/23 + , , , m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04
         m_strCP48 = Pub_GetHandleDay(strNickPa01, strNickPa09, textCP10, , DBDATE(textCP07), , , m_PA01 & "-" & m_PA02 & "-" & m_PA03 & "-" & m_PA04)
         
      End If 'Added by Morgan 2012/8/1
   End If
   SetCPFieldNewData "CP48", m_strCP48
   '***** end
   
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
   
   'Add By Sindy 2021/4/29 不是主管機關期限
   strCP64 = textCP64
   If m_strCPM34 = "N" And m_str945CP09 = "" And strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
      '新增
      If frm010001.intModifyKind = 0 Then
         '(2)收文時無設本所期限，以承辦期限＋5個工作天為本所期限
         If Val(textCP06) = 0 Then
            textCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(m_strCP48), , , , "N"), 1)
            SetCPFieldNewData "CP06", DBDATE(textCP06)
         '(1)收文時有設本所期限，自動備註:本所期限為yyy/mm/dd(本所期限)
         Else
            strCP64 = "本所期限為" & ChangeWStringToTDateString(DBDATE(textCP06)) & ";" & strCP64
         End If
      Else
         '(2)收文時無設本所期限，以承辦期限＋5個工作天為本所期限
         If Val(textCP06) = 0 Then
            textCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(m_strCP48), , , , "N"), 1)
            SetCPFieldNewData "CP06", DBDATE(textCP06)
         '(1)收文時有設本所期限，自動備註:本所期限為yyy/mm/dd(本所期限)
         ElseIf Val(DBDATE(m_strCP06)) <> Val(DBDATE(textCP06)) Then '有異動時
            If InStr(strCP64, "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;") = 0 Then
               strCP64 = "原本所期限為" & ChangeWStringToTDateString(DBDATE(m_strCP06)) & "已修改;" & strCP64
            End If
         End If
      End If
   End If
   ' 進度備註
   SetCPFieldNewData "CP64", strCP64 'textCP64 Modify By Sindy 2021/4/29
   
   ' 業務區
   SetCPFieldNewData "CP12", GetSalesArea(textCP13)
   ' 智權人員
   SetCPFieldNewData "CP13", textCP13
   ' 承辦人員
   SetCPFieldNewData "CP14", textCP14
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
   
   'Added by Lydia 2017/12/14 預設是否電子送件
   SetCPFieldNewData "CP118", m_CP118
   
   Select Case m_PA01
      ' 系統類別為CFT的為更新商標基本檔
      Case "P", "CFP", "FCP":
         ' 專利種類
         SetPASPFieldNewData "PA08", textPA08
         ' 卷宗性質
         SetPASPFieldNewData "PA23", textPA23
         ' 申請人一
         If Not IsEmptyText(textPA26) Then
            SetPASPFieldNewData "PA26", textPA26 & String(9 - Len(textPA26), "0")
         Else
            SetPASPFieldNewData "PA26", Empty
         End If
         ' 申請人二
         If Not IsEmptyText(textPA27) Then
            SetPASPFieldNewData "PA27", textPA27 & String(9 - Len(textPA27), "0")
         Else
            SetPASPFieldNewData "PA27", Empty
         End If
         ' 申請人三
         If Not IsEmptyText(textPA28) Then
            SetPASPFieldNewData "PA28", textPA28 & String(9 - Len(textPA28), "0")
         Else
            SetPASPFieldNewData "PA28", Empty
         End If
         ' 申請人四
         If Not IsEmptyText(textPA29) Then
            SetPASPFieldNewData "PA29", textPA29 & String(9 - Len(textPA29), "0")
         Else
            SetPASPFieldNewData "PA29", Empty
         End If
         ' 申請人五
         If Not IsEmptyText(textPA30) Then
            SetPASPFieldNewData "PA30", textPA30 & String(9 - Len(textPA30), "0")
         Else
            SetPASPFieldNewData "PA30", Empty
         End If
         ' 客戶案件案號
         SetPASPFieldNewData "PA48", textPA48
         ' 代理人
         If Not IsEmptyText(textPA75) Then
            SetPASPFieldNewData "PA75", textPA75 & String(9 - Len(textPA75), "0")
         Else
            SetPASPFieldNewData "PA75", Empty
         End If
         ' 案件備註
         SetPASPFieldNewData "PA91", textPA91
      Case Else:
         ' 申請人一
         If Not IsEmptyText(textPA26) Then
            SetPASPFieldNewData "SP08", textPA26 & String(9 - Len(textPA26), "0")
         Else
            SetPASPFieldNewData "SP08", Empty
         End If
         ' 案件備註
         SetPASPFieldNewData "SP18", textPA91
         ' 代理人
         If Not IsEmptyText(textPA75) Then
            SetPASPFieldNewData "SP26", textPA75 & String(9 - Len(textPA75), "0")
         Else
            SetPASPFieldNewData "SP26", Empty
         End If
         ' 客戶案件案號
         SetPASPFieldNewData "SP29", textPA48
         ' 申請人二
         If Not IsEmptyText(textPA27) Then
            SetPASPFieldNewData "SP58", textPA27 & String(9 - Len(textPA27), "0")
         Else
            SetPASPFieldNewData "SP58", Empty
         End If
         ' 申請人三
         If Not IsEmptyText(textPA28) Then
            SetPASPFieldNewData "SP59", textPA28 & String(9 - Len(textPA28), "0")
         Else
            SetPASPFieldNewData "SP59", Empty
         End If
         ' 申請人四
         If Not IsEmptyText(textPA29) Then
            SetPASPFieldNewData "SP65", textPA29 & String(9 - Len(textPA29), "0")
         Else
            SetPASPFieldNewData "SP65", Empty
         End If
         ' 申請人五
         If Not IsEmptyText(textPA30) Then
            SetPASPFieldNewData "SP66", textPA30 & String(9 - Len(textPA30), "0")
         Else
            SetPASPFieldNewData "SP66", Empty
         End If
   End Select
End Sub
'edit by nickc 2007/10/11  取消，改抓有時效性的
'''''''''取得案件收費表的工作天數
''''''''Private Function GetCF04(strCF01 As String, strCF02 As String, strCF03 As String) As String
''''''''Dim rsA As New ADODB.Recordset
''''''''Dim strSQLa As String
''''''''
''''''''GetCF04 = "0"
''''''''If rsA.State <> adStateClosed Then rsA.Close
''''''''Set rsA = Nothing
''''''''strSQLa = "Select CF04 From CaseFee Where CF01='" & strCF01 & "' AND CF02='" & strCF02 & "' AND CF03='" & strCF03 & "' AND CF04 IS NOT NULL"
''''''''rsA.CursorLocation = adUseClient
''''''''rsA.Open strSQLa, cnnConnection, adOpenStatic, adLockReadOnly
''''''''If rsA.RecordCount > 0 Then
''''''''   GetCF04 = rsA.Fields(0).Value
''''''''End If
''''''''If rsA.State <> adStateClosed Then rsA.Close
''''''''Set rsA = Nothing
''''''''
''''''''End Function


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
   Dim rsTmp As ADODB.Recordset
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
        'Modfiy By Cheng 2002/11/06
'         OnUpdateServicePractice
         If OnUpdateServicePractice = False Then GoTo ErrorHandler
   End Select

   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 機關文號
   For nIndex = 1 To grdList.Rows - 1
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         If Not IsEmptyText(grdList.TextMatrix(nIndex, 4)) Then
            strSql = "UPDATE CASEPROGRESS SET CP08 = '" & grdList.TextMatrix(nIndex, 4) & "' " & _
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
   ' 對造案件名稱
   'modify by sonia 2023/10/30 FG-001325故加入textPA23 <> ""條件
   If textPA23 <> "1" And textPA23 <> "" Then
      strSql = "UPDATE CASEPROGRESS SET CP37 = '" & m_PA05 & "', " & _
                                       "CP38 = '" & ChgSQL(m_PA06) & "', " & _
                                       "CP39 = '" & m_PA07 & "' " & _
                "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strSql
   End If
   
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
   'Set rsTmp = New ADODB.Recordset
   'strSQL = "SELECT * FROM CASEFEE " & _
   '         "WHERE CF01 = '" & m_PA01 & "' AND " & _
   '               "CF02 = '" & m_PA09 & "' AND " & _
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
   'Set rsTmp = Nothing
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
            'Modified by Lydia 2017/03/15 拿掉CFP
            'Case "P", "CFP", "FCP":
            Case "P", "FCP":
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
            'Modified by Lydia 2017/03/15 拿掉CFP
            'Case "P", "CFP", "FCP":
            Case "P", "FCP":
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
   For nIndex = 1 To grdList.Rows - 1
      ' 判斷該列是否有被選取
      If grdList.TextMatrix(nIndex, 0) = "V" Then
         strNP07 = grdList.TextMatrix(nIndex, 9)
         strNP22 = grdList.TextMatrix(nIndex, 10)
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

   'add by nickc 2005/05/18
   strSql = "UPDATE engineerPROGRESS SET ep06 = " & ServerDate & ", " & _
                                    "ep14 = " & ServerDate & " " & _
             "WHERE ep02 = '" & m_CP09 & "' "
   cnnConnection.Execute strSql
   '911018 nick 當有相關總收文號時，要將總收文號該筆更新成續辦，因為只會有一筆時才會讀出來秀畫面，所以不用np22
   '91.11.10 MODIFY BY SONIA
   'If textCP43 <> "" Then
   '     strSQL = "update nextprogress set np06='Y' where np01='" & textCP43 & "' "
   '     cnnConnection.Execute strSQL
   'End If
   '91.11.10 END
   
   'Add By Sindy 2025/7/29
   '針對轉案進來的案件，若內部收文日為111111'案件性質為[（107）再審申請]，則
   '按確定後'彈出訊息：請輸入再審查送件日期,將日期回寫到進度檔的[再審查申請]的發文日
   If m_PA01 = "FCP" And textCP10 = "107" And textCP05 = "111111" And m_strCP27 <> "" Then
      strSql = "update caseprogress set cp27=" & DBDATE(m_strCP27) & " where cp09='" & m_CP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   '2025/7/29 END
   'Added by Morgan 2012/4/18
   If m_str945CP09 <> "" Then
      strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_str945CP09 & "'"
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/4/18
   
   'Add by Sindy 2022/6/28
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm010001", IIf(Left(Pub_StrUserSt03, 1) = "F", m_CP09, "")
   End If
   '2022/6/28 END
   
   'Added by Lydia 2018/06/25 當客戶提供文件處理英說做內部收文進度備註有註明"英文參考本"時，自動將補文件收文號寫入"英文本收文號"
   'Modified by Lydia 2021/02/22 改判斷
   'If TypeName(frm010001.mPrevForm) = "frm060121_1" And InStr(textCP64, "英文參考本") > 0 Then
   If frm010001.m_GetB202CP09 <> "" And InStr(textCP64, "英文參考本") > 0 Then
         strSql = "update transfee set tf30='" & m_CP09 & "' where tf01 in (select cp09 from caseprogress " & _
                     "where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "' and cp10='201' and cp159=0) "
         cnnConnection.Execute strSql, intI
   End If
   'end 2018/06/25
   
   'Added by Lyda 2023/07/28 外專-FCP專利連結案管制：收文特定案件性質, 自動收文「通知資訊變更961」,發一封Email給承辦工程師
   If m_PA01 = "FCP" And m_PA177 = "Y" Then
      strExc(1) = m_PA01: strExc(2) = m_PA02: strExc(3) = m_PA03: strExc(4) = m_PA04
      If PUB_GetFCPlinkMC("3", strSrvDate(1), strExc, m_CP09, textCP10) = True Then
      End If
   End If
   'end 2023/07/28
   
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    'add by nickc 2006/08/22
    OnSaveData = False
End Function

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
      
      'Add by Morgan 2008/8/28 承辦期限
      SetCPFieldOldData "CP48", "" & rsTmp.Fields("cp48"), 0
      
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
      'Added by Lydia 2017/12/14 預設是否電子送件
      strTemp = Empty
      If IsNull(rsTmp.Fields("CP118")) = False Then
         strTemp = rsTmp.Fields("CP118")
      End If
      m_CP118 = strTemp
      SetCPFieldOldData "CP118", strTemp, 0
      'end 2017/12/14
      
      ' 卷宗性質不為1時, 案件中英日文名稱從案件進度檔中帶入
      If IsEmptyText(m_CP10) = False Then
         If textPA23.Text <> "1" Then
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

' 讀取資料庫
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim nickI As Integer
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
      
   m_PA57 = Empty
   m_PA16 = "": m_PA14 = "" '2010/8/17 add by sonia
   
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
   
   ' 本所案號
   textPAKey = m_PA01 & m_PA02 & m_PA03 & m_PA04
      
   Select Case m_PA01
      Case "P", "CFP", "FCP":
         QueryPatent
      Case Else:
         QueryServicePractice
   End Select
   
'   'Add By Sindy 2022/3/31 是否為FMP案件
'   If PUB_ChkIsFMP(m_PA01, m_PA02, m_PA03, m_PA04) = True Then
'      m_bolFMP = True
'   Else
'      m_bolFMP = False
'   End If
'   '2022/3/31 END

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
         '911113 nick 邱小姐說刪除
         'EnableTextBox textPA14, True
      Case Else:
         EnableTextBox textPA08, False
         '911113 nick 邱小姐說刪除
         'EnableTextBox textPA14, False
   End Select

   ' 讀取優先權資料
   m_Pa(1) = m_PA01
   m_Pa(2) = m_PA02
   m_Pa(3) = m_PA03
   m_Pa(4) = m_PA04
   'Modify by Amy 2014/04/18 +, m_Priority(5)
   ClsPDReadPriority m_Pa, m_Priority(1), m_Priority(2), m_Priority(3), m_Priority(4), m_Priority(5)
   
'modify by sonia 2019/3/12 改先抓案件性質檔,其他僅內部收文不請款的才寫在程式裡
'   '2006/3/7 ADD BY SONIA
'   '2007/6/21 MODIFY BY SONIA 加928重新委任
'   '2009/11/9 modify by sonia 加935案件轉至本所
'   'Modify by Morgan 2011/1/7 +229
'   If m_PA01 = "FCP" And (textCP10 = "908" Or textCP10 = "419" Or textCP10 = "928" Or textCP10 = "935" Or textCP10 = "229") Then
'      textCP20.Text = "N"
'   End If
'   '2006/3/7 END
   textCP20.Text = PUB_GetCP20(m_PA01, textCP10)
'end 2019/3/12
    
   'Modified by Lydia 2024/05/28 改成模組
   ''Added by Lydia 2022/05/03  FCP-062174審定前不收費控制:補上是否向客戶收款=N
   'If m_PA16 = "" And InStr("FCP062174000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
   '   textCP20.Text = "N"
   'End If
   ''end 2022/05/03
   ''Added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
   'If m_PA16 <> "1" And InStr("FCP067004000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
   '   textCP20.Text = "N"
   'End If
   ''end 2022/05/03
   If PUB_GetCP20forSpec(m_PA01, m_PA02, m_PA03, m_PA04, m_PA16) = "N" Then
       textCP20.Text = "N"
   End If
   'end 2024/05/28
   
   'add by nick 2004/09/08 案件性質是 411 時，是否向客戶收款=Y
   If textCP10.Text = "411" Or textCP10 = "202" Or textCP10 = "404" Then
      textCP20.Text = "N"
   End If

   'add by nick 2005/02/15
   Dim Is411IsNotOne As Boolean
   Dim Is411 As Boolean
   ' 更新本案期限的資料
   UpdateGrdList m_PA01, m_PA02, m_PA03, m_PA04
   
   'Modify by Morgan 2006/5/3
   'FCP的補文件202不要做
   If Not (m_PA01 = "FCP" And m_CP10 = "202") Then
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
               'Add by Morgan 2007/1/18 台灣專利的申復或修正時下一程序兩個都要抓
               If m_PA09 = "000" And (textCP10 = "205" Or textCP10 = "204") Then
                  nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07 IN ('204','205') and  np06 is null "
               End If
               'end 2007/1/18
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
                   For nickI = 1 To grdList.Rows - 1
                       'edit by nick 2004/09/08
                       'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                       If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                           grdList.TextMatrix(nickI, 0) = "V"
                       End If
                   Next nickI
                   If Is411 = True Then
                        For nickI = 1 To grdList.Rows - 1
                           
                        Next nickI
                   End If
               Else
                   '911203 nick 改成先抓 null 若 0 筆，則再抓 <>'Y'，但如果大於一筆，則都不代
                   If nick911018rs.RecordCount = 0 Then
                       Set nick911018rs = New ADODB.Recordset
                       nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07=" & m_CP10 & " and np06 <>'Y'  "
                       'Add by Morgan 2007/1/18 台灣專利的申復或修正時下一程序兩個都要抓
                       If m_PA09 = "000" And (textCP10 = "205" Or textCP10 = "204") Then
                           nickstrsql = "select * from nextprogress where np02='" & m_PA01 & "' and np03='" & m_PA02 & "' and np04='" & m_PA03 & "' and np05='" & m_PA04 & "' and np07 IN ('204','205') and np06 <>'Y' "
                       End If
                       'end 2007/1/18
                       nick911018rs.CursorLocation = adUseClient
                       nick911018rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
                       If nick911018rs.RecordCount = 1 Then
                           textCP06 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np08").Value))
                           textCP07 = ChangeWStringToTString(CheckStr(nick911018rs.Fields("np09").Value))
                           textCP43 = CheckStr(nick911018rs.Fields("np01").Value)
                           textCP64 = textCP64 & CheckStr(nick911018rs.Fields("np15").Value)
                           textCP13 = CheckStr(nick911018rs.Fields("np10").Value)
                           textCP13_Validate False
                           For nickI = 1 To grdList.Rows - 1
                               'edit by nick 2004/09/08
                               'If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And grdList.TextMatrix(nickI, 2) = textCP06 And grdList.TextMatrix(nickI, 3) = textCP07 Then
                               If Trim(grdList.TextMatrix(nickI, 9)) = Trim(CheckStr(nick911018rs.Fields("np07").Value)) And Val(grdList.TextMatrix(nickI, 2)) = Val(textCP06) And Val(grdList.TextMatrix(nickI, 3)) = Val(textCP07) And textCP10.Text <> "411" Then
                                   grdList.TextMatrix(nickI, 0) = "V"
                               End If
                           Next nickI
                       End If
                   End If
                   'add by nickc 2005/02/15  ************************************
                   If textCP10.Text = "411" Then
                        Is411IsNotOne = False
                        Is411 = False
                        For nickI = 1 To grdList.Rows - 1
                           If Trim(grdList.TextMatrix(nickI, 9)) = "411" Then
                              If Is411IsNotOne = False Then
                                  Is411 = True
                                  Is411IsNotOne = True
                              ElseIf Is411IsNotOne = True Then
                                  Is411 = False
                                  Exit For
                              End If
                           End If
                         Next nickI
                        If Is411 = True Then
                             For nickI = 1 To grdList.Rows - 1
                                 If Trim(grdList.TextMatrix(nickI, 9)) = "411" Then
                                       textCP06 = grdList.TextMatrix(nickI, 2)
                                       textCP07 = grdList.TextMatrix(nickI, 3)
                                       textCP43 = grdList.TextMatrix(nickI, 8)
                                 End If
                             Next nickI
                        End If
                  End If
                  'add end ********************************************
               End If
           End If
      End If
   End If
   '2006/5/3 end
   
   ' 設定輸入的位置
   SetInputEntry

   ' 91.09.11 申請人不輸入
   EnableTextBox textPA26, False
   EnableTextBox textPA27, False
   EnableTextBox textPA28, False
   EnableTextBox textPA29, False
   EnableTextBox textPA30, False
   
   '92.03.27 nick 當查詢時，將確定 disabled
   If frm010001.intModifyKind = 2 Then
        cmdOK.Enabled = False
   End If
   
   'Added by Lydia 2017/12/14 預設是否電子送件
   'Modify By Sindy 2024/11/11 1=有主管機關者
   ' 依操作的案件性質檢查是否屬於有呈送主管機關(不管是否為經濟部智慧財產局)，則"電子送件"欄位，請自動上"Y"，以防人員當紙本送件
   If PUB_ChkhadCF10forEMP_46(m_PA01, m_PA09, Trim(textCP10)) = 1 _
      And m_PA09 = "000" And m_CP118 = "" Then
      'Modify By Sindy 2024/12/2 敏莉說803舉發預設的電子送件"Y"請拿掉
      If Trim(textCP10) <> "803" Then
      '2024/12/2 END
         MsgBox "本案為電子送件案，本程序將預設為電子送件！", vbExclamation
         m_CP118 = "Y"
      End If
   End If
'   'Modified by Lydia 2018/05/17  排除對象非智慧局(告知代理人901,會稿924,回覆代理人902,其他翻譯927)
'   If m_PA01 = "FCP" And m_PA09 = "000" And _
'      InStr("601,605,232,421,807,941,501,503,803,804,901,924,902,927", Trim(textCP10)) = 0 And _
'      InStr(NewCasePtyList, textCP10) = 0 And m_CP118 = "" Then
'        'Modified by Lydia 2018/03/05 EXE檔會出錯
'        'strExc(0) = "select 1 from caseprogress where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' AND CP03='" & m_PA03 & "' AND CP04='" & m_PA04 & "' AND CP10 IN (" & NewCasePtyList & ") and cp118 is not null"
'        strExc(0) = "select count(*) cnt from caseprogress where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' AND CP03='" & m_PA03 & "' AND CP04='" & m_PA04 & "' AND CP10 IN (" & GetAddStr(NewCasePtyList) & ") and cp118 is not null "
'        intI = 1
'        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'        If intI = 1 Then
'           If Val("" & RsTemp.Fields("cnt")) > 0 Then 'Added by Lydia 2018/03/05
'             MsgBox "本案為電子送件案，本程序將預設為電子送件！", vbExclamation
'             m_CP118 = "Y"
'           End If  'end 2018/03/05
'        End If
'    End If
'   'end 2017/12/14
   '2024/11/11 END
   
   Call SetCP43 'Added by Lydia 2019/06/14
   
   'Add by Morgan 2004/8/16
   OnUpdateFee
End Sub

'Add by Morgan 2004/8/16
Private Sub OnUpdateFee()
   Dim lngFee As Long
   
   'Modify By Sindy 2016/7/1 是否向客戶收款預設為N者,不預設費用
   'If m_PA01 = "FCP" Then
   If m_PA01 = "FCP" And textCP20.Text <> "N" Then
   '2016/7/1 END
      ' 規費
      '2010/8/17 modify by sonia
      'textCP17 = GetPatentOfficialFee(m_PA01, textCP10.Text, "", m_PA08, m_PA09, "")
      'Modified by Lydia 2017/12/14 +是否電子送件
      'textCP17 = GetPatentOfficialFee(m_PA01, textCP10.Text, "", m_PA08, m_PA09, m_PA16, m_PA14, m_PA02, m_PA03, m_PA04)
      textCP17 = GetPatentOfficialFee(m_PA01, textCP10.Text, "", m_PA08, m_PA09, m_PA16, m_PA14, m_PA02, m_PA03, m_PA04, m_CP118)
      lngFee = Val(GetFCPFee(m_PA01, textCP10.Text)) + Val(textCP17)
      ' 費用
      If lngFee > 0 Then
         textCP16 = Format(lngFee)
         '點數
         textCP18 = Format((Val(lngFee) - Val(textCP17)) / 1000, "0.0")
      End If
   End If
   
   'Modified by Lydia 2024/05/28 改成模組
   ''Added by Lydia 2020/03/27 FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
   'If m_PA16 = "" And InStr("FCP062174000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
   '     textCP16 = ""
   '     textCP17 = ""
   '     textCP18 = ""
   'End If
   ''end 2020/03/27
   ''Added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
   'If m_PA16 <> "1" And InStr("FCP067004000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
   '     textCP16 = ""
   '     textCP17 = ""
   '     textCP18 = ""
   'End If
   ''end 2022/05/03
   If PUB_GetCP20forSpec(m_PA01, m_PA02, m_PA03, m_PA04, m_PA16) = "N" Then
        textCP16 = ""
        textCP17 = ""
        textCP18 = ""
   End If
   'end 2024/05/28
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
         m_PA05 = rsTmp.Fields("PA05")
      End If
      ' 專利名稱(英)
      If Not IsNull(rsTmp.Fields("PA06")) Then
         cmbPA05.AddItem rsTmp.Fields("PA06")
         m_PA06 = rsTmp.Fields("PA06")
      End If
      ' 專利名稱(日)
      If Not IsNull(rsTmp.Fields("PA07")) Then
         cmbPA05.AddItem rsTmp.Fields("PA07")
         m_PA07 = rsTmp.Fields("PA07")
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
         m_PA09 = rsTmp.Fields("PA09")
      End If
      '911113 nick 邱小姐說刪除
      ' 公告日
      'If Not IsNull(rsTmp.Fields("PA14")) Then
      '   textPA14 = TAIWANDATE(rsTmp.Fields("PA14"))
      'End If
      'SetPASPFieldOldData "PA14", textPA14, 1
      ' 券宗性質
      If Not IsNull(rsTmp.Fields("PA23")) Then
         textPA23 = rsTmp.Fields("PA23")
      End If
      SetPASPFieldOldData "PA23", textPA23, 1
      ' 申請人一
      If Not IsNull(rsTmp.Fields("PA26")) Then
         textPA26 = rsTmp.Fields("PA26")
         textPA26_2 = GetCustomerName(rsTmp.Fields("PA26"), 0)
      End If
      SetPASPFieldOldData "PA26", textPA26, 0
      ' 申請人二
      If Not IsNull(rsTmp.Fields("PA27")) Then
         textPA27 = rsTmp.Fields("PA27")
         textPA27_2 = GetCustomerName(rsTmp.Fields("PA27"), 0)
      End If
      SetPASPFieldOldData "PA27", textPA27, 0
      ' 申請人三
      If Not IsNull(rsTmp.Fields("PA28")) Then
         textPA28 = rsTmp.Fields("PA28")
         textPA28_2 = GetCustomerName(rsTmp.Fields("PA28"), 0)
      End If
      SetPASPFieldOldData "PA28", textPA28, 0
      ' 申請人四
      If Not IsNull(rsTmp.Fields("PA29")) Then
         textPA29 = rsTmp.Fields("PA29")
         textPA29_2 = GetCustomerName(rsTmp.Fields("PA29"), 0)
      End If
      SetPASPFieldOldData "PA29", textPA29, 0
      ' 申請人五
      If Not IsNull(rsTmp.Fields("PA30")) Then
         textPA30 = rsTmp.Fields("PA30")
         textPA30_2 = GetCustomerName(rsTmp.Fields("PA30"), 0)
      End If
      SetPASPFieldOldData "PA30", textPA30, 0
      
      'Add By Sindy 2009/10/19
      m_PA26 = Trim(textPA26)
      m_PA27 = Trim(textPA27)
      m_PA28 = Trim(textPA28)
      m_PA29 = Trim(textPA29)
      m_PA30 = Trim(textPA30)
      
      ' 客戶案件案號
      If Not IsNull(rsTmp.Fields("PA48")) Then
         textPA48 = rsTmp.Fields("PA48")
      End If
      SetPASPFieldOldData "PA48", textPA48, 0
      ' 代理人
      If Not IsNull(rsTmp.Fields("PA75")) Then
         textPA75 = rsTmp.Fields("PA75")
         textPA75_2 = GetFAgentName(rsTmp.Fields("PA75"))
         textPA75_Validate False
      End If
      '911113 nick 原先缺  補上
      SetPASPFieldOldData "PA75", textPA75, 0
      ' 案件備註
      If Not IsNull(rsTmp.Fields("PA91")) Then
         textPA91 = rsTmp.Fields("PA91")
      End If
      SetPASPFieldOldData "PA91", textPA91, 0
      
      '911113 nick 暫存
      strNickPa01 = CheckStr(rsTmp.Fields("PA01").Value)
      strNickPa09 = CheckStr(rsTmp.Fields("PA09").Value)
      'Add By Cheng 2002/12/16
      m_PA57 = "" & rsTmp.Fields("PA57").Value
      '2010/8/17 add by sonia
      If Not IsNull(rsTmp.Fields("PA14")) Then
         m_PA14 = rsTmp.Fields("PA14")
      End If
      If Not IsNull(rsTmp.Fields("PA16")) Then
         m_PA16 = rsTmp.Fields("PA16")
      End If
      '2010/8/17 end
      m_PA177 = "" & rsTmp.Fields("PA177") 'Added by Lydia 2023/07/28 FCP專利連結通知
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
   
   ' 收據編號
   SetCPFieldOldData "CP60", Empty, 0
   ' 進度備註
   SetCPFieldOldData "CP64", Empty, 0
   
   '911108 nick 因為會有些值沒有先定義，所以會沒有更新
   SetCPFieldOldData "CP11", Empty, 0
   SetCPFieldOldData "CP13", Empty, 0
   SetCPFieldOldData "CP16", 0, 1
   SetCPFieldOldData "CP17", 0, 1
   SetCPFieldOldData "CP18", 0, 1
   SetCPFieldOldData "CP20", Empty, 0
   SetCPFieldOldData "CP21", Empty, 0
   '911113 nick
   SetCPFieldOldData "CP32", Empty, 0
   SetCPFieldOldData "CP48", Empty, 1
'cancel by sonia 2019/3/12 移至QueryData
'   '92.10.31 ADD BY SONIA
'   If textCP10 = "202" Or textCP10 = "404" Then
'      textCP20 = "N"
'   End If
'   '92.10.31 END
   
   'Added by Lydia 2017/12/14 預設是否電子送件
   m_CP118 = Empty
   SetCPFieldOldData "CP118", Empty, 0
   'end 2017/12/14
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
         m_PA05 = rsTmp.Fields("SP05")
      End If
      ' 專利名稱(英)
      If Not IsNull(rsTmp.Fields("SP06")) Then
         cmbPA05.AddItem rsTmp.Fields("SP06")
         m_PA05 = rsTmp.Fields("SP06")
      End If
      ' 專利名稱(日)
      If Not IsNull(rsTmp.Fields("SP07")) Then
         cmbPA05.AddItem rsTmp.Fields("SP07")
         m_PA05 = rsTmp.Fields("SP07")
      End If
      ' 顯示專利名稱
      If cmbPA05.ListCount > 0 Then
         cmbPA05.ListIndex = 0
      End If
      ' 申請國家
      If Not IsNull(rsTmp.Fields("SP09")) Then
         m_PA09 = rsTmp.Fields("SP09")
      End If
      ' 申請人一
      If Not IsNull(rsTmp.Fields("SP08")) Then
         textPA26 = rsTmp.Fields("SP08")
         textPA26_2 = GetCustomerName(rsTmp.Fields("SP08"), 0)
      End If
      SetPASPFieldOldData "SP08", textPA26, 0
      ' 申請人二
      If Not IsNull(rsTmp.Fields("SP58")) Then
         textPA27 = rsTmp.Fields("SP58")
         textPA27_2 = GetCustomerName(rsTmp.Fields("SP58"), 0)
      End If
      SetPASPFieldOldData "SP58", textPA27, 0
      ' 申請人三
      If Not IsNull(rsTmp.Fields("SP59")) Then
         textPA28 = rsTmp.Fields("SP59")
         textPA28_2 = GetCustomerName(rsTmp.Fields("SP59"), 0)
      End If
      SetPASPFieldOldData "SP59", textPA28, 0
      ' 申請人四
      If Not IsNull(rsTmp.Fields("SP65")) Then
         textPA29 = rsTmp.Fields("SP65")
         textPA29_2 = GetCustomerName(rsTmp.Fields("SP65"), 0)
      End If
      SetPASPFieldOldData "SP65", textPA29, 0
      ' 申請人五
      If Not IsNull(rsTmp.Fields("SP66")) Then
         textPA30 = rsTmp.Fields("SP66")
         textPA30_2 = GetCustomerName(rsTmp.Fields("SP66"), 0)
      End If
      SetPASPFieldOldData "SP66", textPA30, 0
      
      'Add By Sindy 2009/10/19
      m_PA26 = Trim(textPA26)
      m_PA27 = Trim(textPA27)
      m_PA28 = Trim(textPA28)
      m_PA29 = Trim(textPA29)
      m_PA30 = Trim(textPA30)
      
      ' 客戶案件案號
      If Not IsNull(rsTmp.Fields("SP29")) Then
         textPA48 = rsTmp.Fields("SP29")
      End If
      SetPASPFieldOldData "SP29", textPA48, 0
      ' 代理人
      If Not IsNull(rsTmp.Fields("SP26")) Then
         textPA75 = rsTmp.Fields("SP26")
         textPA75_2 = GetFAgentName(rsTmp.Fields("SP26"))
         textPA75_Validate False
      End If
      SetPASPFieldOldData "SP26", textPA75, 0
      ' 案件備註
      If Not IsNull(rsTmp.Fields("SP18")) Then
         textPA91 = rsTmp.Fields("SP18")
      End If
      SetPASPFieldOldData "SP18", textPA91, 0
        'Add By Cheng 2002/12/16
        m_PA57 = "" & rsTmp.Fields("SP15").Value
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub GetGridData()
   
   'Add by Morgan 2007/10/24 只有補文件才可以點
   'Modify by Morgan 2007/10/29 +404延期
   'If textCP10 <> "202" Then
   If textCP10 <> "202" And textCP10 <> "404" Then
      MsgBox "只有【補文件或延期】才可點選下一程序！"
      Exit Sub
   End If
   
   'Added by Lydia 2021/08/31 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
   If Pub_CheckNpTheSameShow(m_PA01, textCP10, Trim("" & grdList.TextMatrix(grdList.row, 9))) = False Then
       Exit Sub
   End If
   'end 2021/08/31
             
   grdList.col = 0
            
   If grdList.Text = "V" Then
      grdList.Text = Empty
   Else
      'Modify by Morgan 2009/12/23 延期只更新期限不可點選
      'grdList.Text = "V"
      If textCP10 <> "404" Then
         grdList.Text = "V"
      'Added by Morgan 2025/3/5
      Else
         strExc(1) = grdList.TextMatrix(grdList.row, 9)
         strExc(2) = grdList.TextMatrix(grdList.row, 7)
         strExc(3) = ""
         If strExc(1) = "416" Then
            strExc(3) = grdList.TextMatrix(grdList.row, 1)
         ElseIf strExc(1) = "202" And InStr(strExc(2), "優先權證明") > 0 Then
            strExc(3) = grdList.TextMatrix(grdList.row, 1) & "-優先權證明"
         End If
         If strExc(3) <> "" Then
            MsgBox "【延期】不可點選【" & strExc(3) & "】！", vbExclamation
            Exit Sub
         End If
      'end 2025/3/5
      End If
      'End 2009/12/23
   
      '911018 nick 當有勾選第一筆時，將本所期限，法定期限，備註，相關總收文號更新
      '911111 nick 邱小姐說改成若本所期限沒值時，以勾的該筆代 本所期限，法定期限，備註，相關總收文號 到上方
      '            智權人員沒值時，以勾的該筆代智權人員
      'If grdList.Row = 1 Then
      'Modify By Sindy 2016/8/22 內部收文補文件時,若點選數個下一程序之補文件,
      '                          無論其文件備註內容是委任書或優先權文件,若其本所及法定期限均不一致時,
      '                          請以最早之本所及法定期限為該補文件之期限
'      If textCP06.Text = "" Then
'         grdList.col = 2
'         textCP06 = grdList.Text
'         grdList.col = 3
'         textCP07 = grdList.Text
'         grdList.col = 8
'         textCP43 = grdList.Text
'         grdList.col = 7
'         textCP64 = textCP64 & grdList.Text
'      End If
      grdList.col = 2
      If Val(textCP06.Text) = 0 Or _
         (Val(textCP06.Text) > Val(grdList.Text)) Then
         grdList.col = 2
         textCP06 = grdList.Text
         grdList.col = 3
         textCP07 = grdList.Text
         grdList.col = 8
         textCP43 = grdList.Text
      End If
      grdList.col = 7 '備註
      If grdList.Text <> "" Then
         If InStr(Trim(textCP64), Trim(grdList.Text)) = 0 Then
            'Modify By Sindy 2016/9/10 備註串在一起時要加分號區隔
            If textCP64 = "" Then
               textCP64 = grdList.Text
            Else
               textCP64 = textCP64 & ";" & grdList.Text
            End If
            '2016/9/10 END
         End If
      End If
      '2016/8/22 END
      If textCP13.Text = "" Then
         grdList.col = 11
         'Modify By Sindy 2016/8/22 該案的目前智權人員
         'textCP13 = grdList.Text
         textCP13 = ShowCurrCP13(m_PA01, m_PA02, m_PA03, m_PA04, m_PA09)
         textCP13_2 = GetStaffName(textCP13)
         '2016/8/22 END
      End If
      'End If
   End If
   grdList_ShowSelection
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2022/6/28
   If m_strIR01 <> "" Then
      If Not m_PrevForm Is Nothing Then
         Call m_PrevForm.GoNext
         Set m_PrevForm = Nothing
      End If
   End If
   '2022/6/28 END
   
   Set frm010012_06 = Nothing
End Sub

Private Sub grdList_KeyDown(KeyCode As Integer, Shift As Integer)
   ' 案件性質必須為延期的才可以選取
   If KeyCode = vbKeySpace Then
      If grdList.row > 0 Then
         GetGridData
      End If
   End If
EXITSUB:
End Sub

Private Sub grdList_SelChange()
   If grdList.row > 0 Then
      GetGridData
   End If
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
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

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      If m_PA01 = "FCP" And textCP10 = "202" Then
         cboAddCP64.Enabled = True
         strExc(1) = ""
         If textCP43 <> "" Then
            strExc(0) = "select cp10 from caseprogress where cp09='" & textCP43 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strExc(1) = RsTemp(0)
            End If
         End If
         PUB_SetCombo202 cboAddCP64, strExc(1)
      Else
         cboAddCP64.Enabled = False
      End If
   End If
End Sub

Private Sub textCP10_LostFocus()
'modify by sonia 2019/3/12 改先抓案件性質檔,其他僅內部收文不請款的才寫在程式裡
'   '2006/3/8 MODIFY BY SONIA 加 請求逕行審查419
'   '2007/6/21 MODIFY BY SONIA 加928重新委任
'   'Modify by Morgan 2011/1/7 +229
'   If m_PA01 = "FCP" And (textCP10 = "908" Or textCP10 = "419" Or textCP10 = "928" Or textCP10 = "229") Then
'      textCP20.Text = "N"
'      'Add By Sindy 2016/7/1
'      textCP16.Text = ""
'      textCP17.Text = ""
'      textCP18.Text = ""
'      '2016/7/1 END
'   End If
   textCP20.Text = PUB_GetCP20(m_PA01, textCP10)
   If textCP10.Text = "411" Or textCP10 = "202" Or textCP10 = "404" Then
      textCP20.Text = "N"
   End If
   If textCP20.Text = "N" Then
      textCP16.Text = ""
      textCP17.Text = ""
      textCP18.Text = ""
   End If
   
'end 2019/3/12
    Call SetCP43 'Added by Lydia 2019/06/14
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
         
      'Add by Morgan 2011/1/6
      Else
         '重新核稿承辦人不可為原翻譯的核稿人
         If textCP10 = "229" Then
            If Left(textCP14, 1) = "F" Then
               strExc(1) = "select * from staff_idmap where sim02='" & textCP14 & "' and sim01=ep04"
            Else
               strExc(1) = "select * from staff_idmap where sim01='" & textCP14 & "' and sim02=ep04"
            End If
            strExc(0) = "select ep04 from caseprogress,engineerprogress" & _
               " where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "'" & _
               " and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "' and cp10='201' and ep02(+)=cp09" & _
               " and (ep04='" & textCP14 & "' or exists(" & strExc(1) & "))"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               MsgBox "重新核稿承辦人不可為原翻譯的核稿人!!", vbExclamation
               Cancel = True
               textCP14.SetFocus
               textCP14_GotFocus
            End If
         End If
         'add by sonia 2017/10/18B類其他翻譯927且承辦人為外翻編號且相關總收文號為C類,預設進度備註
         If Left(textCP43, 1) = "C" And textCP10 = "927" And Left(textCP14, 1) = "F" Then
            If textCP64 = "OA委外翻譯" Then
            Else
               textCP64 = "OA委外翻譯;" & textCP64
            End If
         End If
         'end 2017/10/18
      End If
   End If
   Set rsTmp = Nothing
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
  'If Cancel = False Then
  '          If (textCP10 = "601" Or textCP10 = "605") And Len("" & textCP14) > 0 Then
  '             If rs.State <> adStateClosed Then rs.Close
  '             Set rs = Nothing
  '             rs.CursorLocation = adUseClient
  '             rs.Open " Select ST03 From Staff Where ST01='" & textCP14 & "'", cnnConnection, adOpenStatic, adLockReadOnly
  '             If rs.RecordCount > 0 Then
  '                If rs.Fields(0).Value <> "F22" Then
  '                   MsgBox "承辦人必須為F22部門人員!!!", vbExclamation + vbOKOnly, "輸入錯誤"
  '                   textCP14.SetFocus
  '                   Exit Function
  '                End If
  '             Else
  '                MsgBox "承辦人輸入錯誤!!!", vbExclamation + vbOKOnly
  '                textCP14.SetFocus
  '                Exit Function
  '             End If
  '             If rs.State <> adStateClosed Then rs.Close
  '             Set rs = Nothing
  '          End If
  '   End If
End Sub

Private Sub textCP16_GotFocus()
   TextInverse textCP16
End Sub

Private Sub textCP16_Validate(Cancel As Boolean)
    If textCP16.Text <> "" Then
      'Modified by Lydia 2024/05/28 改成模組
      ''Added by Lydia 2020/03/27 FCP-062174審定前不收費控制: 判斷基本檔之目前准/駁PA16為空值時，不管任何案件性質都不必預設收文費用、規費、點數。
      'If m_PA16 = "" And InStr("FCP062174000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
      '       textCP16 = ""
      '       textCP17 = ""
      '       textCP18 = ""
      '       Exit Sub
      'End If
      ''end 2020/03/27
      ''Added by Lydia 2022/05/03 FCP-067004核准前不收費控制：申請至核准(暫不包含領證)不收任何收費 (包含規費及服務費、若客戶提AEP也不收費)
      'If m_PA16 <> "1" And InStr("FCP067004000", m_PA01 & m_PA02 & m_PA03 & m_PA04) > 0 Then
      '       textCP16 = ""
      '       textCP17 = ""
      '       textCP18 = ""
      '       Exit Sub
      'End If
      ''end 2022/05/03
      If PUB_GetCP20forSpec(m_PA01, m_PA02, m_PA03, m_PA04, m_PA16) = "N" Then
           textCP16 = ""
           textCP17 = ""
           textCP18 = ""
           Exit Sub
      End If
      'end 2024/05/28
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GetCaseFee(m_PA01, m_PA09, textCP10.Text, Val(textCP16.Text)) <> 1 Then
      'MODIFY BY SONIA 2014/7/16 +傳規費 CFP-027024
      If ClsPDGetCaseFee(m_PA01, m_PA09, textCP10.Text, Val(textCP16.Text), Val(textCP17.Text)) <> 1 Then
         textCP16_GotFocus
         Cancel = True
      End If
   End If
End Sub

Private Sub textCP17_GotFocus()
   TextInverse textCP17
End Sub

Private Sub textCP17_Validate(Cancel As Boolean)
   
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strFee As String
   
   If textCP17.Text <> "" Then
      If m_PA09 = "000" Then
         '2010/8/17 modify by sonia
         'strFee = GetPatentOfficialFee(m_PA01, textCP10.Text, "", m_PA08, m_PA09, "")
         strFee = GetPatentOfficialFee(m_PA01, textCP10.Text, "", m_PA08, m_PA09, m_PA16, m_PA14, m_PA02, m_PA03, m_PA04)
         If Val(strFee) > 0 Then
            If Val(textCP17.Text) <> Val(strFee) Then
               strTit = "檢核資料"
               strMsg = "規費應為<" & strFee & ">"
               nResponse = MsgBox(strMsg, vbOKCancel + vbCritical, strTit)
               textCP17_GotFocus
               Cancel = True
            End If
         End If
      End If
   End If
End Sub

Private Sub textCP18_GotFocus()
   TextInverse textCP18
End Sub

Private Sub textCP18_Validate(Cancel As Boolean)
   If textCP18.Text <> "" Then
      If textCP16.Text <> "" Or textCP17.Text <> "" Then
         If Format((Val(textCP16.Text) - Val(textCP17.Text)) / 1000, "0.0") <> Format(Val(textCP18.Text), "0.0") Then
            ShowMsg MsgText(1036)
            textCP18_GotFocus
            Cancel = True
         End If
      Else
         ShowMsg MsgText(1037)
         textCP18_GotFocus
         Cancel = True
      End If
   End If
   'Add by Morgan 2004/10/1 點數為負時是否向客戶請款上'N'
   If Cancel = False Then
      'Modify By Sindy 2016/7/1 + And Val(textCP16) = 0
      'If Val(textCP18.Text) < 0 Then textCP20.Text = "N"
      If Val(textCP18.Text) < 0 And Val(textCP16) = 0 Then textCP20.Text = "N"
      '2016/7/1 END
   End If
End Sub

'911113 nick add
Private Sub textCP20_GotFocus()
   InverseTextBox textCP20
   'edit by nickc 2007/06/06 切換輸入法改用API
   'textCP20.IMEMode = 2 'Add by Morgan 2004/9/10
   CloseIme
End Sub

'911113 nick add
Private Sub textCP20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'Add by Morgan 2004/9/10 不向客戶請款，費用不可輸入
Private Sub textCP20_KeyUp(KeyCode As Integer, Shift As Integer)
   
   If textCP20.Text = "N" Then
      'Modify by Morgan 2004/9/23 改預設為0
      textCP16.Text = "0"
      textCP16.Enabled = False
   Else
      textCP16.Text = ""
      textCP16.Enabled = True
   End If
   OnUpdateFee 'Added by Lydia 2017/12/14
End Sub

'911113 nick add
Private Sub textCP20_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textCP20) = False Then
      Select Case textCP20
         Case " ", "N":
         Case Else
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '911111 nick
            textCP20.SetFocus
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
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
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
      
      'add by sonia 2017/10/17 B類其他翻譯927且承辦人為外翻編號且相關總收文號為C類時,此為OA委外翻譯,會有帳單要扣點數,輸請款單時要一併勾選
      If textCP10 = "927" And Left(textCP14, 1) = "F" And Left(textCP43, 1) = "C" And textCP20 <> "" Then
         MsgBox "其他翻譯927且承辦人為外翻編號且相關總收文號為C類時,此為OA委外翻譯,會有帳單要扣點數,輸請款單時要一併勾選,所以自動設定為要請款!!!", vbExclamation + vbOKOnly
         textCP20 = ""
      End If
      'end 2017/10/17
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

'911113 nick 邱小姐說刪除
' 公告日
'Private Sub textPA14_Validate(Cancel As Boolean)
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'   If Not IsEmptyText(textPA14) Then
'      If Not CheckIsTaiwanDate(textPA14, False) Then
'         Cancel = True
'         strTit = "檢核資料"
'         strMsg = "公告日輸入不正確"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         '911111 nick
'         textPA14.SetFocus
'
'         textPA14_GotFocus
'      End If
'   End If
'End Sub

' 申請人一
Private Sub textPA26_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textPA26_2 = Empty
   If IsEmptyText(textPA26) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textPA26_2 = GetCustomerName(textPA26, 0)
      textPA26_2 = GetCustomerNameAndState(textPA26, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textPA26_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textPA26 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA26.SetFocus
         
         textPA26_GotFocus
      End If
   End If
End Sub

' 申請人二
Private Sub textPA27_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textPA27_2 = Empty
   If IsEmptyText(textPA27) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textPA27_2 = GetCustomerName(textPA27, 0)
      textPA27_2 = GetCustomerNameAndState(textPA27, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textPA27_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textPA27 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA27.SetFocus
         
         textPA27_GotFocus
      End If
   End If
End Sub

' 申請人三
Private Sub textPA28_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textPA28_2 = Empty
   If IsEmptyText(textPA28) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textPA28_2 = GetCustomerName(textPA28, 0)
      textPA28_2 = GetCustomerNameAndState(textPA28, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textPA28_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textPA28 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA28.SetFocus
         
         textPA28_GotFocus
      End If
   End If
End Sub

' 申請人四
Private Sub textPA29_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textPA29_2 = Empty
   If IsEmptyText(textPA29) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textPA29_2 = GetCustomerName(textPA29, 0)
      textPA29_2 = GetCustomerNameAndState(textPA29, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textPA29_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textPA29 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA29.SetFocus
         
         textPA29_GotFocus
      End If
   End If
End Sub

' 申請人五
Private Sub textPA30_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textPA30_2 = Empty
   If IsEmptyText(textPA30) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      Dim oState As Boolean
      'textPA30_2 = GetCustomerName(textPA30, 0)
      textPA30_2 = GetCustomerNameAndState(textPA30, 0, oState)
      If oState = False Then
        Cancel = True
        Exit Sub
      End If
      If textPA30_2 = Empty Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "申請人代碼<" & textPA29 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA30.SetFocus
         
         textPA30_GotFocus
      End If
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
               '91.11.10 MODIFY BY SONIA
               'If textPA23 <> "1" Then
               If m_PA01 = "FCP" And (textPA23 <= "1" And textPA23 >= "3") Then
               '91.11.10 END
                  Cancel = True
                  strTit = "檢核資料"
                  strMsg = "卷宗性質不正確"
                  nResponse = MsgBox(strMsg, vbOKOnly, strTit)
                  '911111 nick
                  textPA23.SetFocus
                  
                  textPA23_GotFocus
               End If
         End Select
      End If
   End If
End Sub

Private Sub textPA57_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/12/16
   If KeyAscii <> 89 And KeyAscii <> 8 Then
        KeyAscii = 0
   End If
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

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim strTemp As String
Dim nResponse
Dim ii As Integer
Dim nickI As Integer          ' add by nickc 2005/02/15
   
Dim Cancel As Boolean 'Add by Morgan 2004/9/10

   CheckDataValid = False
   
   'Add by Morgan 2004/8/10
   ' 承辦人不可空白
   'Modify by Morgan 2011/1/7 重新核稿可以先不輸承辦人--靜芳
   'If IsEmptyText(textCP14) = True Then
   If IsEmptyText(textCP14) = True And Not (m_PA01 = "FCP" And textCP10 = "229") Then
      strTit = "檢核資料"
      strMsg = "承辦人不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP14.SetFocus
      GoTo EXITSUB
   End If
   
   If textCP20.Text <> "N" Then
      textCP16 = Format(Val(textCP16.Text))
      textCP17 = Format(Val(textCP17.Text))
      textCP18 = Format(Val(textCP18.Text), "0.0")
   End If
   'Add end
   
   ' 案件性質不可為空白
   If IsEmptyText(textCP10) = True Then
      strTit = "檢核資料"
      strMsg = "案件性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP10.SetFocus
      GoTo EXITSUB
   End If
   ' 案件性質為年費或延期時本所期限及法定期限不可為空白
   If textCP10 = "605" Or textCP10 = "404" Then
      If IsEmptyText(textCP06) = True Then
         strTit = "檢核資料"
         strMsg = "案件性質為年費或延期時, 本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         strMsg = "案件性質為年費或延期時, 本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         GoTo EXITSUB
      End If
   End If
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限與法定期限範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   '911113 nick 邱小姐說刪除
   ' 案件性質為異議時
   'If textCP10 = "801" Then
   '   If IsEmptyText(textPA14) = True Then
   '      strTit = "檢核資料"
   '      strMsg = "案件性質為異議時公告日不可空白"
   '      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '      textPA14.SetFocus
   '      GoTo EXITSUB
   '   End If
   'End If
   ' 收文日
   If IsEmptyText(textCP05) = True Then
      strTit = "檢核資料"
      strMsg = "收文日不可空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP05.SetFocus
      GoTo EXITSUB
   End If
   ' 卷宗性質
   If IsEmptyText(textPA23) = True And m_PA01 = "FCP" Then
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
   ' 申請人及代理人
   If IsEmptyText(textPA26) And IsEmptyText(textPA27) And IsEmptyText(textPA28) And IsEmptyText(textPA29) And IsEmptyText(textPA30) And IsEmptyText(textPA75) Then
      strTit = "檢核資料"
      strMsg = "申請人及代理人不可同時空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textPA26.SetFocus
      GoTo EXITSUB
   End If
   
   'Add By Cheng 2003/08/13
   '若案件性質為延期, 則不可點選本案期限
   If Me.textCP10.Text = "404" Then
       For ii = 1 To Me.grdList.Rows - 1
           If Me.grdList.TextMatrix(ii, 0) <> "" Then
               MsgBox "此案僅收文<延期>，不可點選下一程序期限資料，" & vbCrLf & "否則無法管制下一程序的期限!!!", vbExclamation + vbOKOnly
               GoTo EXITSUB
           End If
       Next ii
   End If
   
   'Add By Sindy 2016/6/30 若有輸入費用時,是否向客戶收款'欄不可為N
   If Val(textCP16) > 0 And textCP20 = "N" Then
      strTit = "檢核資料"
      strMsg = "是否向客戶收款欄不可為N"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP20.SetFocus
      GoTo EXITSUB
   End If
   '2016/6/30 END
    
   'Add by Morgan 2004/9/10
   If Me.textCP20.Enabled = True Then
      Cancel = False
      textCP20_Validate Cancel
      If Cancel = True Then
         Me.textCP20.SetFocus
         textCP20_GotFocus
         Exit Function
      End If
   End If
   If Me.textCP16.Enabled = True Then
      Cancel = False
      textCP16_Validate Cancel
      If Cancel = True Then
         Me.textCP16.SetFocus
         textCP16_GotFocus
         Exit Function
      End If
   End If
   If Me.textCP17.Enabled = True Then
      Cancel = False
      textCP17_Validate Cancel
      If Cancel = True Then
         Me.textCP17.SetFocus
         textCP17_GotFocus
         Exit Function
      End If
   End If
   If Me.textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         Me.textCP18.SetFocus
         textCP18_GotFocus
         Exit Function
      End If
   End If
   'add by nickc 2005/02/15 催審時都不能勾
   If textCP10.Text = "411" Then
      For nickI = 1 To grdList.Rows - 1
          If Trim(grdList.TextMatrix(nickI, 0)) <> "" Then
               MsgBox "催審，要將點選的勾拿掉!", , "警告！"
               Exit Function
          End If
      Next nickI
   End If
   
   'Added by Morgan 2011/11/11 Ex.FCP-030936
   '當收文有規費但不請款的補收款時,若有再審延期已發文則提醒是否為延期規費繳交不足要補繳
   If textCP10 = "911" And textCP20 = "N" And Val(textCP17) > 0 Then
      strExc(0) = "select cp09 from caseprogress a where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "' and cp10='404' and cp27>0"
      strExc(0) = strExc(0) & " and (exists(select * from caseprogress b where b.cp09=a.cp43 and b.cp10='107' and b.cp27 is null)"
      strExc(0) = strExc(0) & " or exists(select * from nextprogress b where b.np01=a.cp43 and b.np07='107' and (b.np06 is null or b.np06='Y')))"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '相關總收文號不為延期之總收文號時才提醒
         If textCP43 <> RsTemp.Fields("cp09") Then
            If MsgBox("請注意！若為再審延期規費繳交不足要補繳而收文之補收款，則相關總收文號應掛延期之總收文號。是否確定要繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
      
   'Added by Lydia 2019/05/23 勘誤公報控管: 內部收文更正(402)及更改(403)需輸入相關總收文號,不限案件性質
   'Modified by Lydia 2019/06/14 +核准之後m_PA16
   'Remove by Lydia 2019/07/30 Sharon表示有公告公報才需要彈提示,其他更正不用輸入相關總收文號
   'If m_PA01 = "FCP" And m_PA16 = "1" And (textCP10 = 更正 Or textCP10 = 更改) And Trim(textCP43) = "" Then
   '    MsgBox "請輸入相關總收文號！", vbCritical
   '    textCP43.SetFocus
   '    textCP43_GotFocus
   '    Exit Function
   'End If
   
   CheckDataValid = True
EXITSUB:
End Function

' FC代理人
Private Sub textPA75_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTempName As String
   
   Cancel = False
   textPA75_2 = Empty
   If IsEmptyText(textPA75) = False Then
      'edit by nick 2004/07/21 檢查該申請人或代理人狀態，若為不再使用則停在原地
      'If PUB_GetAgentName(m_PA01, textPA75.Text, strTempName) Then
      If PUB_GetAgentNameAndState(m_PA01, textPA75.Text, strTempName) Then
         textPA75_2.Text = strTempName
      Else
         textPA75_2.Text = ""
         If strTempName <> "" Then
              Cancel = True
              Exit Sub
         End If
      End If
      If IsEmptyText(textPA75_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "FC代理人<" & textPA75 & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         '911111 nick
         textPA75.SetFocus
         
         textPA75_GotFocus
      End If
   End If
End Sub

Private Sub textPA08_GotFocus()
   InverseTextBox textPA08
End Sub
'911113 nick 邱小姐說刪除
'Private Sub textPA14_GotFocus()
'   InverseTextBox textPA14
'End Sub

Private Sub textPA23_GotFocus()
   InverseTextBox textPA23
End Sub

Private Sub textPA26_GotFocus()
   InverseTextBox textPA26
End Sub

Private Sub textPA27_GotFocus()
   InverseTextBox textPA27
End Sub

Private Sub textPA28_GotFocus()
   InverseTextBox textPA28
End Sub

Private Sub textPA29_GotFocus()
   InverseTextBox textPA29
End Sub

Private Sub textPA30_GotFocus()
   InverseTextBox textPA30
End Sub

Private Sub textPA48_GotFocus()
   InverseTextBox textPA48
End Sub

Private Sub textPA57_GotFocus()
   InverseTextBox textPA57
End Sub

Private Sub textPA75_GotFocus()
   InverseTextBox textPA75
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
   
   '911113 nick 邱小姐說刪除
   'If textPA14.Enabled = True Then
   '   Cancel = False
   '   textPA14_Validate Cancel
   '   If Cancel = True Then
   '      Exit Function
   '   End If
   'End If
   
   If textPA23.Enabled = True Then
      Cancel = False
      textPA23_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA26.Enabled = True Then
      Cancel = False
      textPA26_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA27.Enabled = True Then
      Cancel = False
      textPA27_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA28.Enabled = True Then
      Cancel = False
      textPA28_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA29.Enabled = True Then
      Cancel = False
      textPA29_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If textPA30.Enabled = True Then
      Cancel = False
      textPA30_Validate Cancel
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
   
   If textPA75.Enabled = True Then
      Cancel = False
      textPA75_Validate Cancel
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
   
   'Add by Morgan 2004/8/10
   If textCP16.Enabled = True Then
      Cancel = False
      textCP16_Validate Cancel
      If Cancel = True Then
         textCP16.SetFocus
         Exit Function
      End If
   End If
   If textCP16.Enabled = True Then
      Cancel = False
      textCP17_Validate Cancel
      If Cancel = True Then
         textCP17.SetFocus
         Exit Function
      End If
   End If
   If textCP18.Enabled = True Then
      Cancel = False
      textCP18_Validate Cancel
      If Cancel = True Then
         textCP18.SetFocus
         Exit Function
      End If
   End If
   'Add end
   
   'Added by Morgan 2012/4/18
   m_str945CP09 = ""
   strExc(0) = "select cp09 from caseprogress where cp01='" & m_PA01 & "' and cp02='" & m_PA02 & "' and cp03='" & m_PA03 & "' and cp04='" & m_PA04 & "' and cp10='945' and cp27||cp57 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("此案尚有電話聯絡單未發文，此次內部收文是否回覆該筆電話聯絡單？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
         m_str945CP09 = RsTemp(0)
         'Added by Morgan 2012/5/22 帶相關總收文號--郭
         If textCP43 = "" Then
            textCP43 = m_str945CP09
         ElseIf textCP43 <> m_str945CP09 Then
            If MsgBox("本收文已輸入相關總收文號但並非該筆電話聯絡單，是否仍要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
         'end 2012/5/22
      End If
   End If
   'end 2012/4/18
   
   'Added by Morgan 2023/9/11
   If m_PA01 = "FCP" And textCP10 = "911" Then
      If textCP43 = "" Then
         MsgBox textCP10_2 & "的相關總收文號不可空白！", vbCritical
         Exit Function
      Else
         strExc(0) = "select * from caseprogress where cp09='" & textCP43 & "' and cp09<'C' and cp27>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI <> 1 Then
            MsgBox textCP10_2 & "的相關總收文號必須是AB類已發文程序！", vbCritical
            Exit Function
         End If
      End If
   End If
   'end 2023/9/11
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

'Added by Lydia 2019/05/23
Private Sub textCP43_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'Added by Lydia 2019/06/14 核准之後內部收文更改403及更正402自動帶入公告公報之相關總收文號；
Private Sub SetCP43()
Dim stCP09 As String

   If m_PA01 = "FCP" And (textCP10 = 更正 Or textCP10 = 更改) Then
       If Trim(textCP43) = "" Then
            Call PUB_ChkCPExist(m_Pa, "1228", , stCP09)
            'Modified by Lydia 2019/06/19 Sharon說人員反應不一定是公告公報，所以用問的
            'If stCP09 <> "" Then textCP43 = stCP09
            If stCP09 <> "" Then
               If MsgBox("相關總收文號是否預設公告公報？", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
                  textCP43 = stCP09
               End If
            End If
            'end 2019/06/19
       End If
   ElseIf m_PA01 = "FCP" And (m_CP10 = 更正 Or m_CP10 = 更改) And m_CP10 <> textCP10 Then
       textCP43 = ""
   End If
End Sub
