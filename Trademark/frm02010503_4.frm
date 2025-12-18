VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010503_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "被異議/被評定/被撤銷/對方補充理由/對方延期/通知復審答辯"
   ClientHeight    =   6036
   ClientLeft      =   156
   ClientTop       =   996
   ClientWidth     =   9336
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6036
   ScaleWidth      =   9336
   Begin TabDlg.SSTab SSTab1 
      Height          =   3216
      Left            =   36
      TabIndex        =   56
      Top             =   2796
      Width           =   9288
      _ExtentX        =   16383
      _ExtentY        =   5673
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   420
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm02010503_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label23"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label15"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label16"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label10"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label28"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label29"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label32"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label25"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label26"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label31"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label8"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "textCP14_2"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "grdList"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "textCP49"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "textPrint"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "textCP26"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP14"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP48"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textWord"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "TextCP64_1"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Frame1"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP06"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP07"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP08"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCF15_2"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCF15"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).ControlCount=   32
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm02010503_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label12"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label17"
      Tab(1).Control(3)=   "Label18"
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label21"
      Tab(1).Control(7)=   "Label27"
      Tab(1).Control(8)=   "Label30"
      Tab(1).Control(9)=   "textCP37"
      Tab(1).Control(10)=   "textCP39"
      Tab(1).Control(11)=   "textCP40"
      Tab(1).Control(12)=   "textCP42"
      Tab(1).Control(13)=   "textCP64"
      Tab(1).Control(14)=   "textCP37_1"
      Tab(1).Control(15)=   "textCP36"
      Tab(1).Control(16)=   "textCP38"
      Tab(1).Control(17)=   "textCP41"
      Tab(1).ControlCount=   18
      Begin VB.ComboBox textCF15 
         Height          =   276
         Left            =   5580
         TabIndex        =   1
         Top             =   300
         Width           =   1332
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   6960
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   300
         Width           =   1932
      End
      Begin VB.TextBox textCP08 
         Height          =   264
         Left            =   1050
         MaxLength       =   40
         TabIndex        =   0
         Top             =   300
         Width           =   2532
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5580
         MaxLength       =   7
         TabIndex        =   11
         Top             =   1020
         Width           =   3372
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1050
         MaxLength       =   7
         TabIndex        =   10
         Top             =   1020
         Width           =   2532
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3780
         TabIndex        =   85
         Top             =   510
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   252
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   9
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   5
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   7
            Top             =   150
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   6
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   8
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1050
         TabIndex        =   84
         Top             =   510
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   3
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.TextBox TextCP64_1 
         Height          =   264
         Left            =   6990
         MaxLength       =   40
         TabIndex        =   13
         Top             =   1290
         Width           =   2205
      End
      Begin VB.TextBox textWord 
         Height          =   264
         Left            =   5520
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73200
         MaxLength       =   600
         TabIndex        =   24
         Top             =   1740
         Width           =   7092
      End
      Begin VB.TextBox textCP38 
         Height          =   264
         Left            =   -73200
         MaxLength       =   100
         TabIndex        =   21
         Top             =   912
         Width           =   7092
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73200
         MaxLength       =   200
         TabIndex        =   18
         Top             =   360
         Width           =   7092
      End
      Begin VB.TextBox textCP48 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   4590
         Locked          =   -1  'True
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1275
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   12
         Top             =   1290
         Width           =   732
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   8130
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   1050
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1860
         Width           =   372
      End
      Begin VB.TextBox textCP49 
         Height          =   264
         Left            =   1050
         MaxLength       =   300
         TabIndex        =   14
         Top             =   1560
         Width           =   8145
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   972
         Left            =   1056
         TabIndex        =   87
         Top             =   2160
         Width           =   8172
         _ExtentX        =   14415
         _ExtentY        =   1715
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
      Begin MSForms.TextBox textCP37_1 
         Height          =   792
         Left            =   -73200
         TabIndex        =   19
         Top             =   636
         Width           =   7092
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12509;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   588
         Left            =   -73200
         TabIndex        =   26
         Top             =   2292
         Width           =   7092
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12509;1037"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73200
         TabIndex        =   25
         Top             =   2016
         Width           =   7092
         VariousPropertyBits=   679493659
         MaxLength       =   600
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73200
         TabIndex        =   23
         Top             =   1464
         Width           =   7092
         VariousPropertyBits=   679493659
         MaxLength       =   600
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73200
         TabIndex        =   22
         Top             =   1188
         Width           =   7092
         VariousPropertyBits=   679493661
         MaxLength       =   100
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73200
         TabIndex        =   20
         Top             =   636
         Width           =   7092
         VariousPropertyBits=   679493661
         MaxLength       =   100
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontEffects     =   1073750016
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   1830
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1692
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         MaxLength       =   20
         Size            =   "2984;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   4650
         TabIndex        =   69
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   68
         Top             =   300
         Width           =   975
      End
      Begin VB.Label Label31 
         Caption         =   "收文文號 :"
         Height          =   255
         Left            =   6090
         TabIndex        =   83
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   3690
         TabIndex        =   67
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   1290
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4650
         TabIndex        =   71
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   86
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   82
         Top             =   660
         Width           =   1572
      End
      Begin VB.Label Label29 
         Caption         =   "(Y:Word)"
         Height          =   255
         Left            =   5940
         TabIndex        =   81
         Top             =   1860
         Width           =   765
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "是否修改定稿 :"
         Height          =   180
         Left            =   4320
         TabIndex        =   80
         Top             =   1860
         Width           =   1170
      End
      Begin VB.Label Label27 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74880
         TabIndex        =   79
         Top             =   2310
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "對造日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   78
         Top             =   2016
         Width           =   1572
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   77
         Top             =   1776
         Width           =   1572
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   76
         Top             =   1512
         Width           =   1572
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件日文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   75
         Top             =   1260
         Width           =   1572
      End
      Begin VB.Label Label17 
         Caption         =   "對造案件英文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   74
         Top             =   960
         Width           =   1572
      End
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   73
         Top             =   660
         Width           =   1572
      End
      Begin VB.Label Label12 
         Caption         =   "對造號數 :"
         Height          =   252
         Left            =   -74880
         TabIndex        =   72
         Top             =   360
         Width           =   972
      End
      Begin VB.Label Label10 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   1020
         Width           =   855
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   6870
         TabIndex        =   65
         Top             =   1860
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   8520
         TabIndex        =   64
         Top             =   1860
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "本案期限 :"
         Height          =   252
         Left            =   120
         TabIndex        =   63
         Top             =   2160
         Width           =   852
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1440
         TabIndex        =   62
         Top             =   1860
         Width           =   2745
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   252
         Left            =   120
         TabIndex        =   61
         Top             =   1860
         Width           =   972
      End
      Begin VB.Label Label14 
         Caption         =   "條款 :"
         Height          =   252
         Left            =   120
         TabIndex        =   60
         Top             =   1560
         Width           =   852
      End
   End
   Begin VB.TextBox textCP40_S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1596
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   696
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   396
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1896
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1296
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   2196
      Width           =   2532
   End
   Begin VB.TextBox textTM45 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1896
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1200
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   2196
      Width           =   2532
   End
   Begin VB.TextBox textTM15 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   396
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   8496
      TabIndex        =   29
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Left            =   6456
      TabIndex        =   27
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   350
      Left            =   7272
      TabIndex        =   28
      Top             =   0
      Width           =   1200
   End
   Begin MSForms.TextBox textCP14_Src 
      Height          =   264
      Left            =   1200
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   2496
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1200
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1296
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5760
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1596
      Width           =   2532
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      MaxLength       =   20
      Size            =   "4466;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1170
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   960
      Width           =   7995
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14102;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label11 
      Caption         =   "申請案號 :"
      Height          =   252
      Left            =   120
      TabIndex        =   55
      Top             =   696
      Width           =   972
   End
   Begin VB.Label Label4 
      Caption         =   "對照名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   54
      Top             =   1596
      Width           =   972
   End
   Begin VB.Label Label24 
      Caption         =   "承辦人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   53
      Top             =   2496
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   52
      Top             =   396
      Width           =   852
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   252
      Left            =   120
      TabIndex        =   51
      Top             =   996
      Width           =   972
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   252
      Left            =   120
      TabIndex        =   50
      Top             =   1296
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   49
      Top             =   1896
      Width           =   732
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   252
      Index           =   6
      Left            =   4800
      TabIndex        =   48
      Top             =   1296
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   252
      Index           =   8
      Left            =   4800
      TabIndex        =   47
      Top             =   2196
      Width           =   852
   End
   Begin VB.Label Label1 
      Caption         =   "彼所案號 :"
      Height          =   252
      Index           =   9
      Left            =   4800
      TabIndex        =   46
      Top             =   1896
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   252
      Index           =   10
      Left            =   120
      TabIndex        =   45
      Top             =   2196
      Width           =   1092
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   252
      Index           =   11
      Left            =   4800
      TabIndex        =   44
      Top             =   1596
      Width           =   972
   End
   Begin VB.Label Label2 
      Caption         =   "審定號數 :"
      Height          =   255
      Left            =   4800
      TabIndex        =   43
      Top             =   390
      Width           =   915
   End
End
Attribute VB_Name = "frm02010503_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/19 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2022/01/03 Form2.0已修改 CmbTM05/textTM23/textCP13/textCP14_Src/textCP14_2/textCP64/textCP37_1/textCP37/textCP39/textCP40/textCP42/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/6 日期欄已修改
Option Explicit

Const CB_SHOWDROPDOWN = &H14F
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
' 申請國家
Dim m_TM10 As String
' 來函收文日
Dim m_CP05 As String
' 機關文號
Dim m_CP08 As String
' 所選取的收文號
Dim m_CP09 As String
' 案件性質
Dim m_CP10 As String
' 智權人員
Dim m_CP13 As String
Dim m_CP12 As String
Dim m_CurrSel As Integer
'Add By Cheng 2003/01/06
Dim m_TM23 As String  ' 申請人
Dim strCP09 As String ' 新增的總收文號
Dim m_TM14 As String  ' 公告日 Add By Sindy 2009/06/16
Dim m_TM28 As String  ' 卷宗性質 2011/6/15 ADD BY SONIA
Dim strRvType As String 'Add By Sindy 2012/4/26
'Added by Morgan 2017/4/25 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/25
'Add By Sindy 2019/5/22
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/22 END
Dim strLD18 As String 'Add By Sindy 2019/12/20 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/20 FC代理人


'Add By Sindy 2019/5/22
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010503_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 25004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010503_3
   Unload frm02010503_2
   Unload frm02010503_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
    'Modify By Cheng 2002/11/07
'      'OnSaveData
        If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
        'Add By Cheng 2003/01/06
        ' 列印定稿
        If textPrint <> "N" Then
           PrintLetter
        End If
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010503_3
      Unload frm02010503_2
      'Add By Sindy 2019/5/22
      If Me.m_strIR01 <> "" Then
         Unload frm02010503_1
         If Not m_PrevForm Is Nothing Then
            Call m_PrevForm.GoNext
         End If
         Unload Me
         '2019/5/22 END
      'Modified by Morgan 2017/4/25 電子公文
      'frm02010503_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010503_1
         frm02010412.GoNext
      Else
         frm02010503_1.Show
         Unload Me
      End If
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM15.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM45.BackColor = &H8000000F
   
   textCP05S.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_Src.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCP40_S.BackColor = &H8000000F
   textCP48.BackColor = &H8000000F
   
   textCF15_2.BackColor = &H8000000F
   
   '2011/6/9 MODIFY BY SONIA
   'textCF15.AddItem "異議答辯"
   If m_TM01 = "TD" Then
      textCF15.AddItem "爭議答辯"
   ElseIf m_TM01 = "TM" Then
      textCF15.AddItem "侵權處理"
   Else
      textCF15.AddItem "異議答辯"
   End If
   '2011/6/9 END
   textCF15.AddItem "評定答辯"
   textCF15.AddItem "廢止答辯"
   textCF15.AddItem "補充答辯"
   textCF15.AddItem "復審答辯"
   
   SSTab1.Tab = 0
   
   MoveFormToCenter Me
   
   'Add By Sindy 2019/5/22
   m_strIR01 = frm02010503_1.m_strIR01
   m_strIR02 = frm02010503_1.m_strIR02
   m_strIR03 = frm02010503_1.m_strIR03
   m_strIR04 = frm02010503_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/22 END
End Sub

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP05 = Empty
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
   End Select
End Sub

' 取得商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   '2011/6/9 ADD BY SONIA TM,TD自其他來函移過來
   Select Case m_TM01
      Case "TD", "TM":
         ' 設定SQL語法
         strSql = "SELECT SP01 AS TM01,SP02 AS TM02,SP03 AS TM03,SP04 AS TM04,SP05 AS TM05,SP06 AS TM06,SP07 AS TM07,SP09 AS TM10 " & _
            ",'' AS TM12,'' AS TM15,'' AS TM14,'' AS TM28,SP08 AS TM23,SP27 AS TM45,SP72 AS TM77,SP26 AS TM44 FROM SERVICEPRACTICE " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
      Case Else
   '2011/6/9 END
         strSql = "SELECT * FROM TradeMark " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
   End Select  '2011/6/9 ADD BY SONIA
                        
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"))
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
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
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'Add By Cheng 2003/01/06
      m_TM23 = "" & rsTmp.Fields("TM23").Value
      
      'Add By Sindy 2019/12/20
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2019/12/20 END
      
      'Add By Sindy 2009/06/16
      m_TM14 = "" & rsTmp.Fields("TM14").Value
      '2009/06/16 End
      
      m_TM28 = "" & rsTmp.Fields("TM28").Value  '2011/6/15 ADD BY SONIA
      
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM15 = rsTmp.Fields("TM15")
      End If
      ' 彼所案號
      If IsNull(rsTmp.Fields("TM45")) = False Then
         textTM45 = rsTmp.Fields("TM45")
      End If
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("TM77"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
Dim strSql As String
Dim rsTmp As New ADODB.Recordset
Dim bCP40 As Boolean
Dim strDay As String
Dim strDate As String
Dim strTemp As String
Dim strCP10 As String
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP12 = Empty
    m_TM23 = Empty
    'Add By Cheng 2003/11/10
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        Me.Label13.Visible = False
        Me.textCP37.Visible = False
        Me.textCP37.Enabled = False
        Me.Label17.Visible = False
        Me.textCP38.Visible = False
        Me.textCP38.Enabled = False
        Me.Label18.Visible = False
        Me.textCP39.Visible = False
        Me.textCP39.Enabled = False
    Case Else
        Me.Label30.Visible = False
        Me.textCP37_1.Visible = False
        Me.textCP37_1.Enabled = False
    End Select
    'End
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09
   
   ' 取得商標基本檔的相關項目
   QueryTradeMark
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP01 = '" & m_TM01 & "' AND " & _
                  "CP02 = '" & m_TM02 & "' AND " & _
                  "CP03 = '" & m_TM03 & "' AND " & _
                  "CP04 = '" & m_TM04 & "' AND " & _
                  "CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 機關文號
      'Add By Cheng 2002/07/17
      m_CP08 = Empty
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
      End If
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
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_Src = GetStaffName(rsTmp.Fields("CP14"))
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      ' 對照名稱 (無中文取英文, 無英文取日文)
      bCP40 = False
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP40")) = False Then
            If IsEmptyText(rsTmp.Fields("CP40")) = False Then
               textCP40_S = rsTmp.Fields("CP40")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP41")) = False Then
            If IsEmptyText(rsTmp.Fields("CP41")) = False Then
               textCP40_S = rsTmp.Fields("CP41")
               bCP40 = True
            End If
         End If
      End If
      If bCP40 = False Then
         If IsNull(rsTmp.Fields("CP42")) = False Then
            If IsEmptyText(rsTmp.Fields("CP42")) = False Then
               textCP40_S = rsTmp.Fields("CP42")
               bCP40 = True
            End If
         End If
      End If
      Select Case frm02010503_3.GetSelectResult()
         Case "4", "5":
            ' 對造號數
            If IsNull(rsTmp.Fields("CP36")) = False Then
               textCP36 = rsTmp.Fields("CP36")
            End If
            Select Case m_TM01
            Case "T", "FCT", "CFT", "TF"
                ' 對造案件名稱
                If IsNull(rsTmp.Fields("CP37")) = False Then
                   textCP37_1 = rsTmp.Fields("CP37")
                End If
            Case Else
                ' 對造案件名稱(中)
                If IsNull(rsTmp.Fields("CP37")) = False Then
                   textCP37 = rsTmp.Fields("CP37")
                End If
                ' 對造案件名稱(英)
                If IsNull(rsTmp.Fields("CP38")) = False Then
                   textCP38 = rsTmp.Fields("CP38")
                End If
                ' 對造案件名稱(日)
                If IsNull(rsTmp.Fields("CP39")) = False Then
                   textCP39 = rsTmp.Fields("CP39")
                End If
            End Select
            ' 對造名稱(中)
            If IsNull(rsTmp.Fields("CP40")) = False Then
               textCP40 = rsTmp.Fields("CP40")
            End If
            ' 對造名稱(英)
            If IsNull(rsTmp.Fields("CP41")) = False Then
               textCP41 = rsTmp.Fields("CP41")
            End If
            ' 對造名稱(日)
            If IsNull(rsTmp.Fields("CP42")) = False Then
               textCP42 = rsTmp.Fields("CP42")
            End If
            '92.5.30 ADD BY SONIA
            If frm02010503_3.GetSelectResult() = "5" Then
               textCP14 = strUserNum
               textCP14_2 = GetStaffName(strUserNum)
            End If
            '92.5.30 END
      End Select
   End If
   rsTmp.Close
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   ' 本案期限
   InitialGrdList
   ' 取得下一程序檔案中的資料列表在 Grid List 中
   strSql = "SELECT * FROM NextProgress " & _
            "WHERE NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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
            If m_TM10 < "010" Then
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
      'Added by Lydia 2023/10/19
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/19
   End If
   rsTmp.Close
   
   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(勝訴)搜尋案件收費表的工作天數
   Select Case frm02010503_3.GetSelectResult()
      Case 1: strCP10 = "1602"
      Case 2: strCP10 = "1603"
      Case 3: strCP10 = "1605"
      Case 4: strCP10 = "1609"
      Case 5: strCP10 = "1611"
      Case 6: strCP10 = "1616"
   End Select
   textCP48 = Empty
''''edit by nickc 2007/10/12 改抓有時效的
''''   strDay = GetWorkDays(m_TM01, m_TM10, StrCp10)
''''   If IsEmptyText(strDay) = False Then
''''      strDate = DBDATE(m_CP05)
''''      ' 90.07.03 modify by louis (承辦期限以實際的工作天數來計算)
''''      'strTemp = DBDATE(Format(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(strDay))))
''''      strTemp = DBDATE(CompWorkDay(Val(strDay), DBDATE(strDate), 0))
''''      textCP48 = TAIWANDATE(strTemp)
''''   End If
   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, strCP10, DBDATE(m_CP05), DBDATE(textCP06), textCP09))
    
   ' 無法取得承辦期限的日期
   If IsEmptyText(textCP48) = True Then
      strTit = "資料檢核"
      '2010/12/16 modify by sonia T-168057
      'strMsg = "無法取得承辦期限, 請聯絡電腦中心！"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      ' 回到前一畫面
      'Unload Me
      'frm02010503_3.Show
      If m_CP05 = 111111 Then
         strMsg = "無法取得承辦期限, 請自行輸入！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.Locked = False
         textCP48.BorderStyle = 1
         textCP48.Enabled = True
      End If
   ElseIf textCP48 = m_CP05 Then
      strMsg = "無法取得承辦期限, 請聯絡電腦中心設定工作天！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      ' 回到前一畫面
      Unload Me
      frm02010503_3.Show
      Exit Sub
   Else
      textCP48.Locked = True
      textCP48.BorderStyle = 0
      textCP48.Enabled = False
      '2010/12/16 end
   End If
      
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
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
         textCP08 = "（" & strTmp & "）慧商字第號"
      End If
   End If
    'Marked By Cheng 2004/04/08
'    'Add By Cheng 2004/03/16
'    '預設來文字號
'    TextCP64_1 = "（" & strTmp & "）智商字第號"
'    'End

   'Added by Morgan 2017/4/25 電子公文
   If m_DocWord <> "" Then
      textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
   ElseIf m_DocNo <> "" Then
      textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
   End If
   '期限
   If m_DeadLine <> "" Then
      Option1(1).Value = True
      If Len(m_DeadLine) >= 7 Then
         Option4(2).Value = True
         Text12 = m_DeadLine
         Text12_Validate False
      ElseIf Right(m_DeadLine, 1) = "日" Then
         Option4(0).Value = True
         Text10 = Val(m_DeadLine)
         Text10_Validate False
      ElseIf Right(m_DeadLine, 1) = "月" Then
         Option4(1).Value = True
         Text11 = Val(m_DeadLine)
         Text11_Validate False
      End If
   End If
   'end 2017/4/25

   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
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
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2019/5/22
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010503_4 = Nothing
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

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim nIndex As Integer
   Dim strSql As String
   Dim bUpdate As Boolean
   Dim strCP06 As String
   Dim strCP07 As String
   Dim strCP12 As String
   Dim strCP27 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   Dim strCP64 As String
   'Add By Sindy 2009/06/16
   Dim strNP08 As String
   Dim strNP09 As String
   '2009/06/16 End
   'Add by Amy 2017/11/13
    Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
    Dim bolUpdCP As Boolean '是否更新進度檔
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 案件性質
   strRvType = "1602"
   Select Case frm02010503_3.GetSelectResult
      Case 1: '被異議
         If IsEmptyText(textCF15) = False Then
            strRvType = "1602"
         Else
            strRvType = "1601"
         End If
      Case 2: '被評定
         If IsEmptyText(textCF15) = False Then
            strRvType = "1604"
         Else
            strRvType = "1603"
         End If
      Case 3: '被廢止
         If IsEmptyText(textCF15) = False Then
            strRvType = "1606"
         Else
            strRvType = "1605"
         End If
      Case 4: '對方補充理由
         strRvType = "1609"
      Case 5: '對方延期
         strRvType = "1611"
      ' 通知復審答辯
      Case 6: '通知復審答辯
         'edit by nickc 2008/01/10 修正，因為早就併入1404
         'strRvType = "1616"
         strRvType = "1404"
   End Select
   
   'Add By Cheng 2002/01/03
   '若來函性質屬於爭議程序(16XX), 應更新商標基本檔是否有爭議程序欄(TM19)為"Y"
   '2011/6/9 modify by sonia
   'If Left(strRvType, 2) = "16" Then
   If Left(strRvType, 2) = "16" And m_TM01 <> "TD" And m_TM01 <> "TD" Then
   
      strSql = "UPDATE TradeMark SET TM19='Y'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
      strSql = "UPDATE TradeMark SET TM77='" & textPrint & "'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
      '2011/6/9 ADD BY SONIA
      strSql = "UPDATE servicepractice SET SP72 = '" & textPrint & "' " & _
               "WHERE SP01 = '" & m_TM01 & "' AND " & _
                     "SP02 = '" & m_TM02 & "' AND " & _
                     "SP03 = '" & m_TM03 & "' AND " & _
                     "SP04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
      '2011/6/9 END
   End If
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   strCP27 = "NULL"
   '91.10.30 MODIFY BY SONIA 宋說都不上發文日, 由承辦人作業上發文日
   'If IsEmptyText(textCF15) = True Then
   '   strCP27 = DBDATE(SystemDate())
   'End If
   '91.10.30 END
   '92.5.30 ADD BY SONIA
   If frm02010503_3.GetSelectResult() = "5" Then
      'edit by nickc 2006/03/17
      'strCP27 = DBDATE(Date)
      strCP27 = strSrvDate(1)
   End If
   '92.5.30 END
   strCP06 = Empty
   strCP07 = Empty
   If IsEmptyText(textCP06) = False Then: strCP06 = DBDATE(textCP06)
   If IsEmptyText(textCP07) = False Then: strCP07 = DBDATE(textCP07)
   
   'Add By Cheng 2004/03/16
    strCP64 = Trim(textCP64)
    If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
       strCP64 = strCP64 & ",收文文號：" & Trim(TextCP64_1)
    ElseIf Trim(TextCP64_1) <> "" Then
       strCP64 = "收文文號：" & Trim(TextCP64_1)
    End If
    'End
   ' 先新增一筆案件進度記錄再更新其本所期限及法定期限
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2002/11/27
    '承辦人為原程序承辦人, 不上發文日
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP49,CP64) " & _
'                  "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                          "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & m_CP13 & "','" & textCP14 & "'," & _
'                          "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
'                          "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "'," & _
'                          "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
'                          "'" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "')"
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    Select Case m_TM01
    Case "T", "FCT", "CFT", "TF"
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP40,CP41,CP42,CP43,CP49,CP64) " & _
'                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                               "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
'                               "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
'                               "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "'," & _
'                               "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
'                               "'" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "')"
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP40,CP41,CP42,CP43,CP49,CP64) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
                               "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                               "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
                               "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "'," & _
                               "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                               "'" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(strCP64) & "')"
        'End
    Case Else
        'Modify By Cheng 2004/02/03
        '業務區為最近收文A類接洽記錄單智權人員的業務區
'        strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP49,CP64) " & _
'                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
'                               "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
'                               "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
'                               "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "'," & _
'                               "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
'                               "'" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "')"
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32,CP36,CP37,CP38,CP39,CP40,CP41,CP42,CP43,CP49,CP64) " & _
                       "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & ChangeTStringToWString(m_CP05) & "," & _
                               "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))) & "','" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP14 & "'," & _
                               "'" & "N" & "','" & textCP26 & "'," & strCP27 & ",'" & "N" & "'," & _
                               "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37) & "','" & ChgSQL(textCP38) & "','" & ChgSQL(textCP39) & "'," & _
                               "'" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & textCP42 & "'," & _
                               "'" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(strCP64) & "')"
        'End
    End Select
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/20 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
      strLD18 = strCP09
      strExc(1) = "" 'Pub_GetSpecMan("內商程序客戶函發後補看人員")
      If Val(textCP06) > 0 Then '有期限者,為掛號
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), strExc(1), True, m_TM23, strRvType, m_TM44
      Else
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), strExc(1), False, m_TM23, strRvType, m_TM44
      End If
   End If
   '2019/12/20 END
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   ' 本所期限
   If IsEmptyText(strCP06) = False Then
      strSql = "UPDATE CaseProgress SET CP06 = " & strCP06 & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 法定期限
   If IsEmptyText(strCP07) = False Then
      strSql = "UPDATE CaseProgress SET CP07 = " & strCP07 & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   If Trim(Text11) <> "" Then
      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   ' 有承辦期限時
   If IsEmptyText(textCP48) = False Then
      strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   ElseIf m_TM01 = "FCT" Then
        If Trim(textCP07) = "" Then
            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                     "WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
        Else
            If DateDiff("d", ChangeWStringToWDateString(DBDATE(m_CP05)), ChangeWStringToWDateString(DBDATE(textCP07))) <= 30 Then    '無法與上句合併，因為沒有日期時，datediff  會發生  型態不符 的錯誤
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            Else
                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(6, DBDATE(m_CP05), 0)) & " " & _
                         "WHERE CP09 = '" & strCP09 & "' "
                cnnConnection.Execute strSql
            End If
        End If
   End If
   Select Case frm02010503_3.GetSelectResult()
      ' 對方補充理由及對方延期時
      Case "4", "5":
        ' 需帶入相關總收文號
         strSql = "UPDATE CaseProgress SET CP43 = '" & m_CP09 & "' " & _
               "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
        '需更新對造資料
        'Modify By Cheng 2002/12/23
        'CP37欄位重覆
'         strSQL = "UPDATE CaseProgress SET CP36 = '" & ChgSQL(Me.textCP36.Text) & "',CP37='" & ChgSQL(Me.textCP37.Text) & "',CP37='" & ChgSQL(Me.textCP37.Text) & "',CP38='" & ChgSQL(Me.textCP38.Text) & "',CP39='" & ChgSQL(Me.textCP39.Text) & "',CP40='" & ChgSQL(Me.textCP40.Text) & "',CP41='" & ChgSQL(Me.textCP41.Text) & "',CP42='" & ChgSQL(Me.textCP42.Text) & "' " & _
'               "WHERE CP09 = '" & strCP09 & "' "
        Select Case m_TM01
        Case "T", "FCT", "CFT", "TF"
            strSql = "UPDATE CaseProgress SET CP36 = '" & ChgSQL(Me.textCP36.Text) & "',CP37='" & ChgSQL(Me.textCP37_1.Text) & "',CP40='" & ChgSQL(Me.textCP40.Text) & "',CP41='" & ChgSQL(Me.textCP41.Text) & "',CP42='" & ChgSQL(Me.textCP42.Text) & "' " & _
                  "WHERE CP09 = '" & strCP09 & "' "
        Case Else
            strSql = "UPDATE CaseProgress SET CP36 = '" & ChgSQL(Me.textCP36.Text) & "',CP37='" & ChgSQL(Me.textCP37.Text) & "',CP38='" & ChgSQL(Me.textCP38.Text) & "',CP39='" & ChgSQL(Me.textCP39.Text) & "',CP40='" & ChgSQL(Me.textCP40.Text) & "',CP41='" & ChgSQL(Me.textCP41.Text) & "',CP42='" & ChgSQL(Me.textCP42.Text) & "' " & _
                  "WHERE CP09 = '" & strCP09 & "' "
        End Select
         cnnConnection.Execute strSql
   End Select
   
   
   'Add By Sindy 2009/06/15
   If frm02010503_3.GetSelectResult = "1" Then '被異議
      '大陸案被異議時,新增下一程序檔
      If m_TM01 = "T" And m_TM10 = "020" Then
         strNP07 = "109"
         ' 法定期限為公告日起3個月加10年-1天
         strNP09 = DateAdd("m", 3, ChangeWStringToWDateString(m_TM14))
         strNP09 = DateAdd("m", 120, strNP09)
         'strNP09 = DBDATE(DateAdd("d", -1, strNP09))  'cancel by sonia 2019/4/11 發現2007/12/1以後案件不必-1天
         strNP09 = DBDATE(strNP09) 'Add by Amy 2020/03/25 bug:日期格式未轉成DB用
         ' 本所期限為法定期限-2天
         strNP08 = DBDATE(DateAdd("d", -2, ChangeWStringToWDateString(DBDATE(strNP09))))
         strNP08 = PUB_GetWorkDay1(strNP08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         
         ' 2010/3/2 add BY SONIA 改為先檢查下一程序無109期限時才新增,有則更新期限,否則被異議二次會重覆掛T-148035
         strSql = "SELECT * FROM NEXTPROGRESS " & _
                  " WHERE NP02 = '" & m_TM01 & "'" & _
                  " And NP03 = '" & m_TM02 & "'" & _
                  " And NP04 = '" & m_TM03 & "'" & _
                  " And NP05 = '" & m_TM04 & "' AND NP06 IS NULL AND NP07='109'"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
         If rsTmp.RecordCount > 0 Then
            strSql = "UPDATE NEXTPROGRESS SET NP08=" & strNP08 & ",NP09=" & strNP09 & _
                     " WHERE NP02 = '" & m_TM01 & "'" & _
                     " And NP03 = '" & m_TM02 & "'" & _
                     " And NP04 = '" & m_TM03 & "'" & _
                     " And NP05 = '" & m_TM04 & "' AND NP06 IS NULL AND NP07='109'"
         Else
         '2010/3/2 END
            '智權人員存最近收文A類接洽記錄單的智權人員
            strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
                           "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & _
                           "'" & strNP07 & "'," & strNP08 & "," & strNP09 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "'," & GetNextProgressNo() & ")"
         End If
         cnnConnection.Execute strSql
         rsTmp.Close
         
         'Added by Lydia 2016/09/23 T大陸案1602被異議(理由),管制催註冊証(1701)先上N
         If strRvType = "1602" Then
            strSql = "UPDATE NEXTPROGRESS SET NP06='N' " & _
                     " WHERE NP02 = '" & m_TM01 & "'" & _
                     " And NP03 = '" & m_TM02 & "'" & _
                     " And NP04 = '" & m_TM03 & "'" & _
                     " And NP05 = '" & m_TM04 & "' AND NP07='1701'"
            cnnConnection.Execute strSql
         End If
         'end 2016/09/23
      End If
   End If
   '2009/06/15 End
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   strNP22 = GetNextProgressNo()
   If IsEmptyText(textCF15) = False Then
    'Modify by Amy 2017/11/13 +if 判斷進度檔已有相同未發文未取消收文之案件性質,則判斷是否更新本限及法限
    If ChkSameCaseProgress(m_TM01, m_TM02, m_TM03, m_TM04, textCF15, m_CP06, m_CP07, st_CP09, m_CP14) = True Then
      If m_CP06 = MsgText(601) Or m_CP07 = MsgText(601) Then
        If MsgBox("下一程序已收文但無期限，是否要代入新期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      ElseIf Val(textCP06) + 19110000 <> Val(m_CP06) Or Val(textCP07) + 19110000 <> Val(m_CP07) Then
        strMsg = "下一程序已收文且期限不同" & vbCrLf & _
                 "已收文本所期限：" & IIf(m_CP06 <> "", Val(m_CP06) - 19110000, "") & " 來函本所期限：" & textCP06 & vbCrLf & _
                 "已收文法定期限：" & IIf(m_CP07 <> "", Val(m_CP07) - 19110000, "") & " 來函法定期限：" & textCP07 & vbCrLf
        
        If MsgBox(strMsg & "是否要更新為來函期限？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
            bolUpdCP = True
        End If
      End If
    End If
      
    '更新進度檔,並發Mail通知承辦人
    If bolUpdCP = True Then
        strSql = "Update CaseProgress Set CP06=" & Val(textCP06) + 19110000 & ",CP07=" & Val(textCP07) + 19110000 & " Where CP09='" & st_CP09 & "'"
        cnnConnection.Execute strSql
        
        If m_CP14 = MsgText(601) Then m_CP14 = GetDeptMan("P20") '無承辦人發給P20部門之A0908
        strMsg = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "收到" & "" & GetCaseTypeName(m_TM01, textCF15, IIf(m_TM10 = "000", 0, 1)) & "前已收文,請辦理後續！"
        PUB_SendMail strUserNum, m_CP14, "", strMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
      
    '進度檔未有相同未發文未取消收文之案件性質或上述不更新期限,才新增下一程序
    Else
        strNP14 = Empty
        strNP14 = GetRelatedPerson(strCP09)
        'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/03
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modified by Lydia 2024/07/02 +ChgSql
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                            strCP06 & "," & strCP07 & ",'" & IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & ChgSQL(textCP64) & "'," & strNP22 & ")"
        cnnConnection.Execute strSql
    End If
    'end 2017/11/13
     
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
'      '92.6.8 SONIA 加 言詞辯論, 準備程序
      Select Case textCF15
'         Case "102", "105", "702", "708", "305", "998", "997", "204", "205":
         Case "102", "105", "702", "708", "305", "998", "997"
         Case Else:
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2004/04/08
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
            'add by nickc 2008/04/23  加入案件回覆單  延期有定稿，理由沒定稿
            If frm02010503_3.textResult.Text <> "5" Then
               '2008/5/20 MODIFY BY SONIA FCT案不印回覆單
               If m_TM01 <> "FCT" Then
                'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                If bolUpdCP = False Then
                    Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strCP07), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                End If
               End If
            End If
      End Select
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
   
   'Add By Sindy 2009/09/24
   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
   Dim strCP48 As String, strCP09B As String
   'modify by sonia 2020/7/9 對方延期1611不產生外商發文722,因為都不通知客戶
   If Left(GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 1) = "F" And strRvType <> "1611" And _
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
   
   'Added by Morgan 2017/4/25 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strRvType
   End If
   'end 2017/4/25
   
   'Add by Sindy 2019/5/22
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010503_1"
   End If
   '2019/5/22 END
   
   Set rsTmp = Nothing
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
    'Add By Cheng 2002/11/19
    If Me.SSTab1.Tab = 0 Then
        textCP08.SetFocus
    Else
        textCP36.SetFocus
    End If
End Sub

' 當使用者按向下鍵時, 將ComboBox顯示成下拉式的樣子
Private Sub textCF15_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyDown Then
      SendMessage textCF15.hWnd, CB_SHOWDROPDOWN, True, ByVal 0&
   End If
End Sub

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      Select Case textCF15.Text
         Case "異議答辯", "爭議答辯", "侵權處理":
            textCF15 = "602"
         Case "評定答辯":
            textCF15 = "604"
         Case "廢止答辯":
            textCF15 = "606"
         Case "補充答辯":
            textCF15 = "613"
         Case "復審答辯":
            textCF15 = "406"
      End Select
      
      ' 只取得國內的案件性質名稱
      If m_TM10 < "010" Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
      Else
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
      End If
      If IsEmptyText(textCF15_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
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
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/07
      End If
      'Add By Cheng 2002/03/11
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/19
        '按下確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP06_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP06_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
   End If
EXITSUB:
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
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/19
        '按下確定時才檢查
'      ' 申請國家為台灣時需檢查來函記錄檔
'      If m_TM10 < "010" Then
'         strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
'         If IsEmptyText(strDate) = False Then
'            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'               strTit = "資料檢核"
'               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'               If nResponse = vbCancel Then
'                  Cancel = True
'                  textCP07_GotFocus
'                  GoTo EXITSUB
'               End If
'            End If
'         Else
'            strTit = "資料檢核"
'            strMsg = "來函記錄中無該筆記錄"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP07_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      End If
   End If
EXITSUB:
End Sub

'Add By Sindy 2010/11/26
Private Sub textCP14_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
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
      strMsg = "對造案件名稱欄位內容太長"
      textCP37_1_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme
End Sub

Private Sub textCP37_LostFocus()
    'Add By Cheng 2002/12/03
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Me.textCP37.IMEMode = 2
    CloseIme
End Sub

' 對造案件名稱(中)
Private Sub textCP37_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP37, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件名稱(中)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP37_GotFocus
   End If
End Sub

' 對造案件名稱(英)
Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP38, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件名稱(英)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP38_GotFocus
   End If
End Sub

' 對造案件名稱(日)
Private Sub textCP39_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP39, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件名稱(日)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP39_GotFocus
   End If
End Sub

Private Sub textCP40_LostFocus()
    'Add By Cheng 2002/12/03
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Me.textCP40.IMEMode = 2
    CloseIme
End Sub

' 對造案件(中)
Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP40, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件(中)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If
End Sub

' 對造案件(英)
Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP41, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件(英)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If
End Sub

' 對造案件(日)
Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP42, 100) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造案件(日)欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
End Sub

Private Sub TextCP64_1_GotFocus()
    TextInverse Me.TextCP64_1
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
      strMsg = "進度備註欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
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
      'Modify By Cheng 2002/07/22
      '條款每項可輸入1~3碼
'      If Len(strTemp) > 3 Then
      'Modify By Sindy 2012/7/5
      'If Len(strTemp) > 3 Or Len(strTemp) < 1 Then
      If Len(strTemp) > 4 Or Len(strTemp) < 1 Then
      '2012/7/5 End
         Cancel = True
         strTit = "條款"
         strMsg = "條款內容<" & strTemp & ">不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP49_GotFocus
         GoTo EXITSUB
      End If
      
      ' 檢查主張內容分類表
      'strSQL = "SELECT * FROM ClaimContents " & _
      '         "WHERE CC01 = '" & Right(strTemp, 1) & "'"
      'rsTmp.CursorLocation = adUseClient
      'rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      'If rsTmp.RecordCount <= 0 Then
      '   Cancel = True
      '   strTit = "條款"
      '   strMsg = "條款內容<" & strTemp & ">不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP49_GotFocus
      '   rsTmp.Close
      '   GoTo ExitSub
      'End If
      'rsTmp.Close
      
      ' 檢查
      'Modify By Sindy 2012/7/5
'      strSql = "SELECT * FROM LAW " & _
'               "WHERE LW01 = '" & Mid(strTemp, 1, 3) & "' "
      strSql = "SELECT * FROM LAW " & _
               "WHERE LW01 = '" & Trim(strTemp) & "' "
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

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2022/01/03檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   'Add by Sindy 2010/02/12
'   '來函性質為1601~1606時控制對造名稱(中,英,日)不可全部空白
'   Select Case m_CP10
'   Case "1601", "1602", "1603", "1604", "1605", "1606"
      If RTrim(textCP40) = "" And RTrim(textCP41) = "" And RTrim(textCP42) = "" Then
         SSTab1.Tab = 1
         strTit = "檢核資料"
         strMsg = "對造中英日文名稱不可均為空白！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP40.SetFocus
         GoTo EXITSUB
      Else
         PUB_ChkCustNameExist textCP40, textCP41, textCP42
      End If
'   Case Else
'   End Select
   
   If m_TM10 < "010" Then
      ' 申請國家為台灣時, 機關文號不可為空白
      If IsEmptyText(textCP08) = True Then
         strTit = "檢核資料"
         strMsg = "申請國家為台灣時, 機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
      ' 申請國家為台灣時, 下一程序不可為空白
      'If IsEmptyText(textCF15) = True Then
      '   strTit = "檢核資料"
      '   strMsg = "申請國家為台灣時, 下一程序不可為空白"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   GoTo ExitSub
      'End If
   End If
   
   'Add By Sindy 2012/4/17
   '檢查來函期限--日期
   If m_TM10 = 台灣國家代號 Then
      If Me.Option4(2).Value = True Then
         If Me.Text12.Text = "" Then
            MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
            Me.Text12.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
   If IsEmptyText(textCF15) = False Then
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      'Add By Cheng 2002/03/11
      If Me.textCP06.Text <> "" Then
         If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
            MsgBox "本所期限不可小於系統日期!!!", vbExclamation
            Me.textCP06.SetFocus
            textCP06_GotFocus
            GoTo EXITSUB
         End If
      End If
      ' 本所期限必須小於法定期限
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If

   ' 承辦期限不可空白
   'If IsEmptyText(textCP48) = True Then
   '   strTit = "檢核資料"
   '   strMsg = "承辦期限不可為空白, 請聯絡電腦中心"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   GoTo ExitSub
   'End If

   ' 申請國家為台灣時, 且有下一程序時, 對造資料不可為空白
   If m_TM10 < "010" Then
      If IsEmptyText(textCF15) = False Then
         'Modify By Cheng 2002/01/15
         '取消對對造號數及對造案件號碼的檢查
'         ' 申請國家為台灣時, 對造號數不可為空白
'         If IsEmptyText(textCP36) = True Then
'            strTit = "檢核資料"
'            strMsg = "申請國家為台灣時, 對造號數不可為空白"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP36.SetFocus
'            GoTo EXITSUB
'         End If
'         ' 申請國家為台灣時, 對造案件中英日文名稱不可均為空白
'         If IsEmptyText(textCP37) = True And IsEmptyText(textCP38) = True And IsEmptyText(textCP39) = True Then
'            strTit = "檢核資料"
'            strMsg = "申請國家為台灣時, 對造案件中英日文名稱不可均為空白"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textCP37.SetFocus
'            GoTo EXITSUB
'         End If
         ' 申請國家為台灣時, 對造中英日文名稱不可均為空白
         If IsEmptyText(textCP40) = True And IsEmptyText(textCP41) = True And IsEmptyText(textCP42) = True Then
            strTit = "檢核資料"
            strMsg = "申請國家為台灣時, 對造中英日文名稱不可均為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP40.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

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
         textCP14_GotFocus
      End If
   End If
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
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
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

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

Private Sub textCP06_GotFocus()
   InverseTextBox textCP06
End Sub

Private Sub textCP07_GotFocus()
   InverseTextBox textCP07
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

Private Sub textCP14_GotFocus()
   InverseTextBox textCP14
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP36_GotFocus()
   InverseTextBox textCP36
End Sub

Private Sub textCP37_GotFocus()
   InverseTextBox textCP37
    'Add By Cheng 2002/12/03
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Me.textCP37.IMEMode = 1
    OpenIme
End Sub

Private Sub textCP38_GotFocus()
   InverseTextBox textCP38
End Sub

Private Sub textCP39_GotFocus()
   InverseTextBox textCP39
End Sub

Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
    'Add By Cheng 2002/12/03
    'edit by nickc 2007/06/06 切換輸入法改用API
    'Me.textCP40.IMEMode = 1
    OpenIme
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
   OpenIme
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/19
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False
If Me.textCF15.Enabled = True Then
   Cancel = False
   textCF15_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
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

If Me.textCP14.Enabled = True Then
   Cancel = False
   textCP14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textCP26.Enabled = True Then
   Cancel = False
   textCP26_Validate Cancel
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

If Me.textCP49.Enabled = True Then
   Cancel = False
   textCP49_Validate Cancel
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
    'Add By Cheng 2002/11/19
    ' 申請國家為台灣時需檢查來函記錄檔
    If m_TM10 < "010" Then
       strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR16")
       If IsEmptyText(strDate) = False Then
          If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
             strTit = "資料檢核"
             strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
             nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
             If nResponse = vbCancel Then
                Cancel = True
                textCP06_GotFocus
                Exit Function
             End If
          End If
       '2008/11/27 CANCEL BY SONIA
       'Else
       '   strTit = "資料檢核"
       '   strMsg = "來函記錄中無該筆記錄"
       '   nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
       '   If nResponse = vbCancel Then
       '      Cancel = True
       '      textCP06_GotFocus
       '      Exit Function
       '   End If
       '2008/11/27 END
       '2011/6/15 ADD BY SONIA
       Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
         Else
            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP06_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
       End If
    End If
    ' 申請國家為台灣時需檢查來函記錄檔
    If m_TM10 < "010" Then
       strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, DBDATE(m_CP05), "MR17")
       If IsEmptyText(strDate) = False Then
          If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
             strTit = "資料檢核"
             strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
             nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
             If nResponse = vbCancel Then
                Cancel = True
                textCP07_GotFocus
                Exit Function
             End If
          End If
       Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            'modify by sonia 2018/2/8 電子公文都不檢查來函記錄檔
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/25 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/25 電子公文
               strTit = "資料檢核"
               strMsg = "來函記錄中無該筆記錄"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         '2011/6/15 ADD BY SONIA
         Else
            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  Cancel = True
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
       End If
    End If

TxtValidate = True
End Function

'Add By Cheng 2003/01/06
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 列印定稿
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub PrintLetter()
Dim strTM23Nation As String
'Add By Sindy 2012/1/13
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/13 End
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
   InsExpField
   
   'Add By Sindy 2012/1/13
   ET01 = "14"
   ET02 = strCP09
   bolEdit = IIf(Me.textWord.Text = "Y", True, False)
   '2012/1/13 End
   
    ' 案件性質為對方補充理由
    If frm02010503_3.textResult.Text = "4" Then
        ' 申請國家為台灣
        If m_TM10 < "010" Then
            'Modify By Cheng 2003/03/04
            '不出定稿
'            ' 列印定稿
'            NowPrint m_CP09, "14", "04", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
        End If
    ' 案件性質為對方延期
    ElseIf frm02010503_3.textResult.Text = "5" Then
        ' 申請國家為台灣
        If m_TM10 < "010" Then
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                ' 列印定稿
                '2006/3/28 MODIFY BY SONIA m_CP09->strCP09
'                NowPrint strCP09, "14", "05", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "05" 'Modify By Sindy 2012/1/13
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
                ' 列印定稿
'                NowPrint strCP09, "14", "06", IIf(Me.textWord.Text = "Y", True, False), strUserNum, 0
               ET03 = "06" 'Modify By Sindy 2012/1/13
            End If
        End If
    End If
    
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, , , bolPlusPaper)
      If bolEmail Then
         '判斷是否EMail同時寄紙本
         If Not bolPlusPaper Then
            iCopy = 1
         End If
         'Add By Sindy 2020/1/7 + 信函總收文號
         If strSrvDate(1) >= T商標電子化第2階段啟用日 And Left(m_TM01, 1) = "T" Then
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
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
   '2021/1/5 EMD
   End If
   '2012/1/13 End
End Sub

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
Dim strTM23Nation As String
Dim strSql As String
Dim strNext As String    '2011/6/14 add by sonia
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
    ' 案件性質為對方補充理
    If frm02010503_3.textResult.Text = "4" Then
        ' 申請國家為台灣
        If m_TM10 < "010" Then
            'Modify By Cheng 2003/03/04
            '不出定稿
'            ' 清除定稿例外欄位檔原有資料
'            EndLetter "14", m_CP09, "04", strUserNum
        End If
    ' 案件性質為對方延期
    ElseIf frm02010503_3.textResult.Text = "5" Then
        ' 申請國家為台灣
        If m_TM10 < "010" Then
            '2011/6/14 add by sonia 依卷宗性質決定定稿下一程序
            If m_TM28 = "1" Then
               strNext = "補充理由"
            Else
               strNext = "答辯"
            End If
            '2011/6/14 end
            
            ' 申請人國籍為台灣
            'edit by nickc 2006/06/30
            'If strTM23Nation < "010" Then
            If textPrint = "1" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "14", m_CP09, "05", strUserNum
                'add by nickc 2008/04/25 案件回覆單
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "14" & "','" & strCP09 & "','" & "05" & "','" & strUserNum & "'," & _
                         "'" & "下一程序" & "','" & strNext & "')"
                cnnConnection.Execute strSql
            ' 申請人國籍非台灣
            'edit by nickc 2006/06/30
            'Else
            ElseIf textPrint = "2" Then
                ' 清除定稿例外欄位檔原有資料
                EndLetter "14", m_CP09, "06", strUserNum
                'add by nickc 2008/04/25 案件回覆單
                strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                         "VALUES ('" & "14" & "','" & strCP09 & "','" & "06" & "','" & strUserNum & "'," & _
                         "'" & "下一程序" & "','" & strNext & "')"
                cnnConnection.Execute strSql
            End If
        End If
    End If
End Sub

Private Sub textWord_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/01/06
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 89 Then
        KeyAscii = 0
    End If
End Sub

'Add By Sindy 2012/4/17
Private Sub Option1_Click(Index As Integer)
   If Me.Option4(0).Value Then
      Text10_Validate False
   ElseIf Me.Option4(1).Value Then
      Text11_Validate False
   ElseIf Me.Option4(2).Value Then
      Text12_Validate False
   End If
End Sub

Private Sub Text10_GotFocus()
   TextInverse Text10
   CloseIme
End Sub

Private Sub Text10_LostFocus()
   '非台灣"天"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text10_Validate(Cancel As Boolean)
   If Text10 <> "" Then GetTime
End Sub

Private Sub Text11_GotFocus()
   TextInverse Text11
   CloseIme
End Sub

Private Sub Text11_LostFocus()
   '非台灣"月"跳離時到"本所期限"欄位
   'If m_TM10 <> 台灣國家代號 Then
   '   If textCP06.Enabled = True Then textCP06.SetFocus
   'End If
End Sub

Private Sub Text11_Validate(Cancel As Boolean)
   If Text11 <> "" Then GetTime
End Sub

Private Sub Text12_GotFocus()
   TextInverse Text12
End Sub

Private Sub Text12_LostFocus()
   '非台灣"日"跳離時到"本所期限"欄位
   If m_TM10 <> 台灣國家代號 Then
      If textCP06.Enabled = True Then textCP06.SetFocus
   End If
End Sub

Private Sub Text12_Validate(Cancel As Boolean)
   If Option4(2).Value = False Then Exit Sub
   If Text12 = "" Then
   Else
      If ChkDate(Text12) Then
         If m_TM10 = 台灣國家代號 Then
            If Val(Text12) < Val(strSrvDate(2)) Then
               MsgBox "來函期限不可小於系統日 !", vbCritical
               Cancel = True
            Else
               textCP07 = Text12
               'Modify By Sindy 2014/10/6 台灣案之本所期限設定
               If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                  textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
               Else
               '2014/10/6 END
                  textCP06 = TransDate(CompDate(2, -2, TransDate(textCP07, 2)), 1)
               End If
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
            End If
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text12
End Sub

Private Sub GetTime()
   Dim i As Integer
   Dim strFromDate As String '期限起算日
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm02010503_1.textCP05)
   
   If m_TM10 = 台灣國家代號 Then
      '文到天數
      If Option4(0).Value = True Then
         textCP07 = TransDate(CompDate(2, Val(Text10), strFromDate), 1)
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text10) >= 60 Then
            i = -4
         Else
            i = -2
         End If
      '文到月數
      ElseIf Option4(1).Value = True Then
         textCP07 = TAIWANDATE(AddMonth(strFromDate, Val(Text11)))
         If Option1(0).Value = True Then textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
         If Val(Text11) >= 2 Then
            i = -4
         Else
            i = -2
         End If
      End If
      If textCP07 <> "" Then
         'Modify By Sindy 2014/10/6 台灣案之本所期限設定
         If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
            textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
         Else
         '2014/10/6 END
            textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
         End If
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm02010503_1.textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' 案件性質
   strRvType = "1602"
   Select Case frm02010503_3.GetSelectResult
      Case 1: '被異議
         strRvType = "1601"
      Case 2: '被評定
         strRvType = "1603"
      Case 3: '被廢止
         strRvType = "1605"
      Case 4: '對方補充理由
         strRvType = "1609"
      Case 5: '對方延期
         strRvType = "1611"
      ' 通知復審答辯
      Case 6: '通知復審答辯
         'edit by nickc 2008/01/10 修正，因為早就併入1404
         'strRvType = "1616"
         strRvType = "1404"
   End Select
   If strRvType = "" Then Exit Function
   
   If ClsPDGetCaseProperty(m_TM01, strRvType, strTempName, bolTmp) Then
      textCP06 = ""
      textCP07 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & strRvType & "'"
         If strExc(0) <> "" Then
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            With RsTemp
               If intI = 1 Then
                  If Not IsNull(.Fields(1)) Then
                     '文到天數
                     Option4(0).Value = True
                     Text10 = .Fields(1)
                     textCP07 = TransDate(CompDate(2, Text10, TransDate(strFromDate, 2)), 1)
                  ElseIf Not IsNull(.Fields(2)) Then
                     '文到月數
                     Option4(1).Value = True
                     Text11 = .Fields(2)
                     textCP07 = TransDate(CompDate(1, .Fields(2), TransDate(strFromDate, 2)), 1)
                  Else
                     '文到天數
                     Option4(0).Value = True
                     Text10 = ""
                     Text11 = ""
                  End If
                  If textCP07 <> "" And Not IsNull(.Fields(0)) Then
                     '文到當日
                     If .Fields(0) = "1" Then
                        Option1(0).Value = True
                        textCP07 = TransDate(CompDate(2, -1, TransDate(textCP07, 2)), 1)
                     '文到次日
                     Else
                        Option1(1).Value = True
                     End If
                  End If
                  '文到天數
                  If Text10 <> "" Then
                     If Val(Text10) >= 60 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  '文到月數
                  ElseIf Not IsNull(.Fields(2)) Then
                     If Val(.Fields(2)) >= 2 Then
                        i = -4
                     Else
                        i = -2
                     End If
                  End If
                  If textCP07 <> "" Then
                     'Modify By Sindy 2014/10/6 台灣案之本所期限設定
                     If m_TM10 = "000" And Val(strSrvDate(1)) >= 台灣案所限新規則啟用日 Then
                        textCP06 = TransDate(PUB_GetOurDeadline(DBDATE(textCP07)), 1)
                     Else
                     '2014/10/6 END
                        textCP06 = TransDate(CompDate(2, i, TransDate(textCP07, 2)), 1)
                     End If
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function
