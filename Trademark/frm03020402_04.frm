VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020402_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核駁輸入"
   ClientHeight    =   6852
   ClientLeft      =   -2688
   ClientTop       =   4596
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6852
   ScaleWidth      =   9144
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   390
      Width           =   2532
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "變更事項(R)"
      Height          =   400
      Left            =   4710
      TabIndex        =   0
      Top             =   -30
      Width           =   1152
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6885
      TabIndex        =   2
      Top             =   -30
      Width           =   1152
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5910
      TabIndex        =   1
      Top             =   -30
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8070
      TabIndex        =   3
      Top             =   -30
      Width           =   912
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1870
      Width           =   1125
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1574
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1574
      Width           =   2412
   End
   Begin VB.TextBox textTM22S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1278
      Width           =   1125
   End
   Begin VB.TextBox textTM27 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1278
      Width           =   2412
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   390
      Width           =   2412
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4600
      Left            =   60
      TabIndex        =   25
      Top             =   2190
      Width           =   9050
      _ExtentX        =   15960
      _ExtentY        =   8128
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm03020402_04.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "textCP35"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "textCP14_2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textCP64"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label25"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label9"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label32"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label28"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label23"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label22"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label21"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label20"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label19"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label18"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label13"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label12"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label11"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label10"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label8"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label24"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label26"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label14"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label15"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label16"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "grdList"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCF15"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCF15_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP06"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP07"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Frame1"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Frame2"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textPrint"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM17"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM16S"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTMBM07_2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTMBM07_1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM14"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP08"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textTM15"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "textCP14"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP48"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCP49"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCP26"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "Text10"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "Text11"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "Text12"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "關係案"
      TabPicture(1)   =   "frm03020402_04.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP40"
      Tab(1).Control(1)=   "textCP42"
      Tab(1).Control(2)=   "textCP37_1"
      Tab(1).Control(3)=   "Label39"
      Tab(1).Control(4)=   "Label35"
      Tab(1).Control(5)=   "Label34"
      Tab(1).Control(6)=   "Label36"
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(8)=   "Label38"
      Tab(1).Control(9)=   "textCP80"
      Tab(1).Control(10)=   "textCP36"
      Tab(1).Control(11)=   "textCP41"
      Tab(1).ControlCount=   12
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   6660
         MaxLength       =   7
         TabIndex        =   90
         Top             =   1410
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   5700
         MaxLength       =   2
         TabIndex        =   89
         Top             =   1410
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   88
         Top             =   1410
         Width           =   375
      End
      Begin VB.TextBox textCP26 
         Height          =   285
         Left            =   7620
         MaxLength       =   1
         TabIndex        =   60
         Top             =   2055
         Width           =   372
      End
      Begin VB.TextBox textCP49 
         Height          =   285
         Left            =   1176
         MaxLength       =   300
         TabIndex        =   59
         Top             =   2355
         Width           =   7695
      End
      Begin VB.TextBox textCP48 
         Height          =   285
         Left            =   4830
         MaxLength       =   8
         TabIndex        =   58
         Top             =   2055
         Width           =   1125
      End
      Begin VB.TextBox textCP14 
         Height          =   285
         Left            =   1200
         MaxLength       =   6
         TabIndex        =   57
         Top             =   2055
         Width           =   732
      End
      Begin VB.TextBox textTM15 
         Height          =   285
         Left            =   5610
         MaxLength       =   20
         TabIndex        =   56
         Top             =   660
         Width           =   2532
      End
      Begin VB.TextBox textCP08 
         Height          =   285
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   55
         Top             =   660
         Width           =   2412
      End
      Begin VB.TextBox textTM14 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   54
         Top             =   360
         Width           =   1125
      End
      Begin VB.TextBox textTMBM07_1 
         Height          =   285
         Left            =   5610
         MaxLength       =   2
         TabIndex        =   53
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox textTMBM07_2 
         Height          =   285
         Left            =   6810
         MaxLength       =   2
         TabIndex        =   52
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox textTM16S 
         Height          =   285
         Left            =   4260
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   51
         Top             =   2655
         Width           =   375
      End
      Begin VB.TextBox textTM17 
         Height          =   285
         Left            =   7620
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   50
         Top             =   2655
         Width           =   372
      End
      Begin VB.TextBox textPrint 
         Height          =   285
         Left            =   1200
         MaxLength       =   1
         TabIndex        =   49
         Text            =   "N"
         Top             =   2655
         Width           =   372
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3900
         TabIndex        =   45
         Top             =   1260
         Width           =   4215
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   48
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   47
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   46
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1200
         TabIndex        =   42
         Top             =   1260
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   44
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   43
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.TextBox textCP07 
         Height          =   285
         Left            =   4830
         MaxLength       =   8
         TabIndex        =   41
         Top             =   1755
         Width           =   1125
      End
      Begin VB.TextBox textCP06 
         Height          =   285
         Left            =   1200
         MaxLength       =   7
         TabIndex        =   40
         Top             =   1755
         Width           =   1125
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   2040
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   960
         Width           =   1572
      End
      Begin VB.TextBox textCF15 
         Height          =   285
         Left            =   1200
         MaxLength       =   4
         TabIndex        =   38
         Top             =   960
         Width           =   732
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73530
         TabIndex        =   29
         Top             =   1800
         Width           =   7092
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73530
         MaxLength       =   200
         TabIndex        =   26
         Top             =   390
         Width           =   7092
      End
      Begin VB.TextBox textCP80 
         Height          =   264
         Left            =   -73530
         MaxLength       =   39
         TabIndex        =   32
         Top             =   2370
         Width           =   3495
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   948
         Left            =   1140
         TabIndex        =   91
         Top             =   2976
         Width           =   7692
         _ExtentX        =   13568
         _ExtentY        =   1672
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
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   285
         Left            =   6390
         TabIndex        =   87
         Top             =   2055
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   285
         Left            =   8040
         TabIndex        =   86
         Top             =   2055
         Width           =   705
      End
      Begin VB.Label Label14 
         Caption         =   "條款 :"
         Height          =   285
         Left            =   120
         TabIndex        =   85
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   285
         Left            =   3930
         TabIndex        =   84
         Top             =   2055
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   285
         Left            =   120
         TabIndex        =   83
         Top             =   2055
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "審定號數 :"
         Height          =   285
         Left            =   4650
         TabIndex        =   82
         Top             =   660
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   285
         Left            =   120
         TabIndex        =   81
         Top             =   660
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "公告日 :"
         Height          =   285
         Left            =   120
         TabIndex        =   80
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "公報卷期 :"
         Height          =   285
         Left            =   4650
         TabIndex        =   79
         Top             =   330
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "卷"
         Height          =   285
         Left            =   6450
         TabIndex        =   78
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label13 
         Caption         =   "期"
         Height          =   285
         Left            =   7650
         TabIndex        =   77
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "案件目前准駁 :"
         Height          =   285
         Left            =   2970
         TabIndex        =   76
         Top             =   2655
         Width           =   2295
      End
      Begin VB.Label Label18 
         Caption         =   "(1:准 , 2:駁)"
         Height          =   285
         Left            =   4680
         TabIndex        =   75
         Top             =   2655
         Width           =   975
      End
      Begin VB.Label Label19 
         Caption         =   "專用權是否存在 :"
         Height          =   285
         Left            =   6150
         TabIndex        =   74
         Top             =   2655
         Width           =   1455
      End
      Begin VB.Label Label20 
         Caption         =   "(Y / N)"
         Height          =   285
         Left            =   8040
         TabIndex        =   73
         Top             =   2655
         Width           =   615
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   285
         Left            =   120
         TabIndex        =   71
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label23 
         Caption         =   "(N:不印)"
         Height          =   285
         Left            =   1620
         TabIndex        =   70
         Top             =   2655
         Width           =   795
      End
      Begin VB.Label Label28 
         Caption         =   "本案期限 :"
         Height          =   255
         Left            =   90
         TabIndex        =   69
         Top             =   2970
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   285
         Left            =   120
         TabIndex        =   68
         Top             =   1395
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   285
         Left            =   120
         TabIndex        =   67
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   285
         Left            =   120
         TabIndex        =   66
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "審查委員 :"
         Height          =   285
         Left            =   4650
         TabIndex        =   65
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   285
         Left            =   3930
         TabIndex        =   64
         Top             =   1755
         Width           =   855
      End
      Begin MSForms.TextBox textCP64 
         Height          =   525
         Left            =   1140
         TabIndex        =   63
         Top             =   3960
         Width           =   7695
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13573;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   1980
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2055
         Width           =   1785
         VariousPropertyBits=   671105055
         MaxLength       =   20
         Size            =   "3149;503"
         BorderColor     =   16777215
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP35 
         Height          =   285
         Left            =   5610
         TabIndex        =   61
         Top             =   960
         Width           =   2535
         VariousPropertyBits=   671105051
         MaxLength       =   32
         Size            =   "4471;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label38 
         Caption         =   "對方日文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   37
         Top             =   2070
         Width           =   1300
      End
      Begin VB.Label Label37 
         Caption         =   "對方英文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   36
         Top             =   1830
         Width           =   1575
      End
      Begin VB.Label Label36 
         Caption         =   "對方中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   35
         Top             =   1545
         Width           =   1300
      End
      Begin VB.Label Label34 
         Caption         =   "對方案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   34
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label35 
         Caption         =   "對方號數 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   33
         Top             =   390
         Width           =   975
      End
      Begin VB.Label Label39 
         Caption         =   "對方商品類別 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   31
         Top             =   2370
         Width           =   1575
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   795
         Left            =   -73530
         TabIndex        =   27
         Top             =   660
         Width           =   7095
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12515;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73530
         TabIndex        =   30
         Top             =   2070
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73530
         TabIndex        =   28
         Top             =   1500
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5670
      TabIndex        =   24
      Top             =   1870
      Width           =   2535
      VariousPropertyBits=   671105055
      Size            =   "4471;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1290
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   982
      Width           =   7245
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "12779;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1290
      TabIndex        =   22
      Top             =   686
      Width           =   7500
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13229;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblClose 
      Caption         =   "lblClose"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   3750
      TabIndex        =   21
      Top             =   390
      Width           =   645
   End
   Begin VB.Label Label27 
      Caption         =   "申請案號 :"
      Height          =   285
      Left            =   4710
      TabIndex        =   20
      Top             =   390
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   285
      Index           =   11
      Left            =   4710
      TabIndex        =   18
      Top             =   1870
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   285
      Index           =   10
      Left            =   180
      TabIndex        =   17
      Top             =   1870
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   285
      Index           =   7
      Left            =   4710
      TabIndex        =   16
      Top             =   1574
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   285
      Index           =   6
      Left            =   180
      TabIndex        =   15
      Top             =   1574
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "正商標專用期止日 :"
      Height          =   285
      Index           =   5
      Left            =   4710
      TabIndex        =   14
      Top             =   1278
      Width           =   1605
   End
   Begin VB.Label Label1 
      Caption         =   "正商標號數 :"
      Height          =   285
      Index           =   4
      Left            =   180
      TabIndex        =   13
      Top             =   1278
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   285
      Left            =   180
      TabIndex        =   12
      Top             =   982
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   285
      Left            =   180
      TabIndex        =   11
      Top             =   686
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   10
      Top             =   390
      Width           =   855
   End
End
Attribute VB_Name = "frm03020402_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、textCP13、textCP14_2、textCP35、textCP64、grdList改字型=新細明體-ExtB
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
' 來函收文日
Dim m_CP05 As String
' 收文號
Dim m_CP09 As String
' 原案件性質
Dim m_CP10 As String
' 原業務區
Dim m_CP12 As String
' 原智權人員代號
Dim m_CP13 As String
' 原移轉申請人代號
Dim m_CP56 As String
' 商標種類代碼
Dim m_TM08 As String
' 國家代碼
Dim m_TM10 As String
' 原專用期限起日
Dim m_TM21 As String
' 原專用期限止日
Dim m_TM22 As String
' 原申請人代號
Dim m_TM23 As String
' 申請國家的延展年度
Dim m_NA14 As Integer
Dim m_TM15 As String '2012/9/20 ADD BY SONIA
Dim m_CurrSel As Integer

'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String

'Add By Cheng 2002/02/01
Dim m_strLastTextTM14 As String
Dim m_strLastTextTMBM07_1 As String
Dim m_strLastTextTMBM07_2 As String
Dim m_strLastTextTM16S As String
Dim m_strLastTextTM17 As String
'add by nickc 2005/08/04
'Dim m_blnClkChgButton As Boolean '是否有按變更事項鈕
Public m_blnClkChgButton As Boolean '是否有按變更事項鈕 'Modify By Sindy 2012/2/6 Dim->Public
Dim strRvType As String 'Add By Sindy 2012/4/26
'Added by Morgan 2017/5/3 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/3
Dim m_NewCP09 As String 'Added by Lydia 2022/02/10 新增C類收文號

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm03020402_03.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm03020402_03
   Unload frm03020402_02
   Unload frm03020402_01
   Unload Me
End Sub

' 提供外部程式呼叫用來結束此項作業
Public Sub OnAppExit()
   cmdExit_Click
End Sub

Private Sub cmdMod_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'add by nickc 2005/08/04
   'Modify By Sindy 2012/2/6 Mark
'    m_blnClkChgButton = True
    
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   'edit by nickc 2005/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "無變更事項記錄"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   DisplayNextForm
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub cmdok_Click()
Dim strFilePath As String 'Added by Lydia 2022/02/10 掃瞄檔的路徑

   If CheckDataValid = True Then
      'Added by Lydia 2022/02/10 FCT紙本公文來函，同時將公文函FCT_OA_SCAN匯入卷宗區
      If frm03020402_03.GetSelectResult() = "1" Then
        If m_DocNo = "" Then
            If PUB_FCTCheckPDF(m_TM01, m_TM02, m_TM03, m_TM04, "1002", m_CP09, strFilePath) = False Then
                 Exit Sub
            End If
        End If
      End If
      'end 2022/02/10
      
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
      'edit by  nick 2004/11/03
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
       'Added by Lydia 2022/02/10 FCT紙本公文來函，同時將公文函FCT_OA_SCAN匯入卷宗區
       'Move by Lydia 2022/02/23 從frm03020402_01.Show上方移過來
       If strFilePath <> "" Then
           If Pub_AutoSavePdf2_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_NewCP09, strRvType, strFilePath) = False Then
               Exit Sub
           End If
       End If
       'end 2022/02/10
       
      Unload Me
      Unload frm03020402_03
      Unload frm03020402_02
      'Modified by Morgan 2017/5/3 電子公文
      'frm03020402_01.Show
      If m_DocNo <> "" Then
         Unload frm03020402_01
         frm02010412.GoNext
      Else
         frm03020402_01.Show
      End If
      'end 2017/5/3
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textTM27.BackColor = &H8000000F
   textTM22S.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
   SSTab1.Tab = 0 'Add by Amy 2022/09/26
    'Add By nickc 2005/08/04
'    m_blnClkChgButton = False
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
   End Select
End Sub

Public Sub QueryTradeMark()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   Dim strSub As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   'edit by nickc 2005/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      'Add By Cheng 2002/07/19
      m_TM10 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         textTM12 = rsTmp.Fields("TM12")
      End If
      ' 審定號數
      m_TM15 = Empty '2012/9/20 ADD BY SONIA
      If IsNull(rsTmp.Fields("TM15")) = False Then
         m_TM15 = rsTmp.Fields("TM15")  '2012/9/20 MODIFY BY SONIA
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
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      'Add By Cheng 2002/07/19
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      ' 正商標號數
      If IsNull(rsTmp.Fields("TM27")) = False Then
         textTM27 = rsTmp.Fields("TM27")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
      ' 正商標專用期止日
      Set rsSub = New ADODB.Recordset
      strSub = "SELECT * FROM TradeMark " & _
               "WHERE TM15 = '" & textTM27 & "' AND " & _
                     "TM10 = '" & m_TM10 & "' "
      rsSub.CursorLocation = adUseClient
      'edit by nickc 2005/08/04
      'rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly
      If rsSub.RecordCount > 0 Then
         rsSub.MoveFirst
         If IsNull(rsSub.Fields("TM22")) = False Then
            textTM22S = TAIWANDATE(rsSub.Fields("TM22"))
         End If
      End If
      rsSub.Close
      Set rsSub = Nothing
      'Modify By Cheng 2002/04/29
'      ' 公告日
'      If IsNull(rsTmp.Fields("TM14")) = False Then
'         textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
'      End If
      'Add By Cheng 2002/07/22
      Me.textTM16S.Text = "" & rsTmp.Fields("TM16").Value
            
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      
   End If
   rsTmp.Close
End Sub

' 讀取案件進度檔
Public Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   'edit by nickc 2005/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 機關文號
      'If IsNull(rsTmp.Fields("CP08")) = False Then
      '   textCP08 = rsTmp.Fields("CP08")
      'End If
      '91.7.9 end
      
      ' 案件性質
      'Add By Cheng 2002/07/19
      m_CP10 = Empty
      If IsNull(rsTmp.Fields("CP10")) = False Then
         m_CP10 = rsTmp.Fields("CP10")
         If m_TM10 < "010" Then
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 0)
         Else
            textCP10 = GetCaseTypeName(m_TM01, rsTmp.Fields("CP10"), 1)
         End If
      End If
      ' 業務區
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      'Add By Cheng 2002/07/19
      m_CP13 = Empty
      'Modified by Lydia 2021/08/03 改由PUB_GetFCTSalesNo帶出和產生的C類收文一致
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      '   textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      'End If
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13 = GetStaffName(m_CP13)
      'end 2021/08/03
      
      ' 核准通知日 91.4.29 modify by Cheng : CANCEL
      'If IsNull(rsTmp.Fields("CP25")) = False Then
      '   textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
      'End If
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 審查委員
      'If IsNull(rsTmp.Fields("CP35")) = False Then
      '   textCP35 = rsTmp.Fields("CP35")
      'End If
      ' 移轉申請人代號
      'Add By Cheng 2002/07/19
      m_CP56 = Empty
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      ' 下一程序
      textCF15 = GetNextProgress(m_TM01, m_TM10, m_CP10)
      If IsEmptyText(textCF15) = False Then
         textCF15_2 = GetCaseTypeName(m_TM01, textCF15)
      End If
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 本所期限
      'If IsNull(rsTmp.Fields("CP06")) = False Then
      '   If IsEmptyText(rsTmp.Fields("CP06")) = False Then
      '      textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
      '   End If
      'End If
      '91.7.9 end
      
      '91.7.9 modify by sonia 根本不必帶此欄
      ' 法定期限
      'If IsNull(rsTmp.Fields("CP07")) = False Then
      '   If IsEmptyText(rsTmp.Fields("CP07")) = False Then
      '      textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
      '   End If
      'End If
      '91.7.9 end
      
      '91.7.9 modify by sonia 承辦人預設為點選資料之智權人員
      ' 承辦人
      'If IsNull(rsTmp.Fields("CP14")) = False Then
      '   textCP14 = rsTmp.Fields("CP14")
      '   textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      'End If
'      If IsNull(rsTmp.Fields("CP13")) = False Then
'         textCP14 = rsTmp.Fields("CP13")
'         textCP14_2 = GetStaffName(rsTmp.Fields("CP13"))
'      End If
        '預設承辦人
        Me.textCP14.Text = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
        Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
      '91.7.9 end
      '91.7.9 modify by sonia 承辦人預設為點選資料之智權人員
      ' 條款
      'If IsNull(rsTmp.Fields("CP49")) = False Then
      '   textCP49 = rsTmp.Fields("CP49")
      'End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   '2012/9/21 add by sonia 延展申請駁回不必管變更事項 FCT-018077
   If frm03020402_03.textResult.Text = "3" And m_CP10 = "102" Then
      m_blnClkChgButton = True
   End If
   '2012/9/21 end
   
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If textCP08 = "" Then
      textCP08 = "（" & strTmp & "）智商字第號"
   End If
   
   'Add By Cheng 2002/01/15
   m_strNumBegin = "商"
   m_strNumEnd = "字"
   
   'Added by Morgan 2017/5/3 電子公文
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
   'end 2017/5/3
End Sub

Public Sub QueryData()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strDay As String
   
   ' 來函收文日
   textCP05S = m_CP05
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   ' 以案件性質"核駁"或"改變原處分"計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case frm03020402_03.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1002")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1002", DBDATE(m_CP05), DBDATE(textCP06)))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1403")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), DBDATE(textCP06)))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''      'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   
   ' 案件性質為申請時才可輸入公告日, 審定號, 公報卷期
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      EnableTextBox textTM14, True
      EnableTextBox textTM15, True
      EnableTextBox textTMBM07_1, True
      EnableTextBox textTMBM07_2, True
      textTM15 = m_TM15    '2012/9/20 add by sonia
   Else
      EnableTextBox textTM14, False
      EnableTextBox textTM15, False
      EnableTextBox textTMBM07_1, False
      EnableTextBox textTMBM07_2, False
      textCP08.SetFocus
   End If
   
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
      'Added by Lydia 2023/10/20
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/20
   End If
   rsTmp.Close
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   Set rsTmp = Nothing
   'Add By Cheng 2002/05/08
   '若前畫面(frm03020402_03)結果欄為"3"時, 預設"是否更新基本檔目前准駁"欄為"N"(不可修改),
   '並預設"進度備註"欄為"申請駁回"
   If frm03020402_03.GetSelectResult = "3" Then
      'Modify By Cheng 2002/07/22
'      Me.textTM16S.Text = "N"
'      Me.textTM16S.Enabled = False
      Me.textCP64.Text = "申請駁回"
   End If
   
   'Add By Cheng 2002/07/11
   If frm03020402_03.textResult.Text = "1" Or frm03020402_03.textResult.Text = "2" Then
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" Then
      If m_CP10 = "101" Or m_CP10 = "308" Then
         'Modify By Cheng 2002/07/22
'         Me.textTM16S.Text = "Y"
         Me.textTM16S.Text = "2"
      Else
         'Modify By Cheng 2002/07/22
'         Me.textTM16S.Text = "N"
      End If
   End If
   
End Sub

Private Sub DisplayNextForm()
   frm03020402_05.SetData 0, m_TM01, True
   frm03020402_05.SetData 1, m_TM02, False
   frm03020402_05.SetData 2, m_TM03, False
   frm03020402_05.SetData 3, m_TM04, False
   frm03020402_05.SetData 5, m_CP09, False
   Me.Hide
   frm03020402_05.Show
   frm03020402_05.QueryData
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
Dim nIndex As Integer
Dim strSql As String
Dim strCP09 As String
'Dim strCP12 As String
Dim strCP48 As String
Dim strNP07 As String
Dim strNP08 As String
Dim strNP09 As String
Dim strNP14 As String
Dim strNP22 As String
'Add By Cheng 2002/07/16
Dim rsTmp As New ADODB.Recordset
Dim bUpdate As Boolean
Dim strCP27 As String   '2010/3/15 add by sonia 已閉卷直接上發文日
   
 '911107 nick transation
   OnSaveData = True
On Error GoTo CheckingErr
cnnConnection.BeginTrans

   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   'edit by  nickc 2005/08/04
   'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then
      GoTo EXITSUB
   End If
   
   rsTmp.MoveFirst
   
   ' 當案件性質為延展, 且原實際結果或准駁日無資料時需 Update 實際結果,准駁日,審查委員的欄位
   'Modify By Cheng 2002/07/16
   '當前一畫面的結果欄為"1"或"3"
'   If frm03020402_03.GetSelectResult() = "1" Then
   If frm03020402_03.GetSelectResult() = "1" Or frm03020402_03.GetSelectResult() = "3" Then
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
         '91.4.29 MODIFY BY SONIA
         strSql = "UPDATE CaseProgress SET CP24 = '2', " & _
                                          "CP25=" & DBDATE(textCP05S) & ", " & _
                                          "CP35='" & textCP35 & "' " & _
                  "WHERE CP09 = '" & m_CP09 & "' AND " & _
                        "(CP24 IS null OR CP24 = '' OR CP24 = ' ')"
         'strSQL = "UPDATE CaseProgress SET CP24 = '2', " & _
         '                                 "CP25=" & DBDATE(textCP25) & ", " & _
         '                                 "CP35='" & textCP35 & "' " & _
         '         "WHERE CP09 = '" & m_CP09 & "' AND " & _
         '               "(CP24 = null OR CP24 = '' OR CP24 = ' ')"
         '91.4.29 END
         cnnConnection.Execute strSql
      End If
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'Modify By Cheng 2002/07/22
   '若案件性質為"延展"(102)
   If m_CP10 = "102" Then
      ' 更新商標基本檔的專用權是否存在
      strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "' "
      cnnConnection.Execute strSql
   End If
      
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 案件性質為申請時
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      If frm03020402_03.textResult.Text = "1" Or frm03020402_03.textResult.Text = "2" Then 'Add By Sindy 2013/7/15 +if
         ' 更新審定號, 來函收文日, 公告日
   '      StrSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                       "TM13 = " & DBDATE(textCP05S) & "," & _
                                       "TM14 = " & DBDATE(textTM14) & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         strSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                       "TM13 = " & DBDATE(textCP05S) & "," & _
                                       "TM14 = " & CNULL(DBDATE(textTM14)) & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         cnnConnection.Execute strSql
      End If
      'Modify By Cheng 2002/07/22
      ' 當使用者輸入要更新基本檔之准/駁時, 更新目前准/駁及審定來函日兩個欄位
'      If textTM16S = "Y" Then
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" And (frm03020402_03.textResult.Text = "1" Or frm03020402_03.textResult.Text = "2") Then
      If (m_CP10 = "101" Or m_CP10 = "308") And (frm03020402_03.textResult.Text = "1" Or frm03020402_03.textResult.Text = "2") Then
         '91.4.29 MODIFY BY SONIA
         strSql = "UPDATE TradeMark SET TM16='2'," & _
                                       "TM13=" & DBDATE(textCP05S) & " " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
         'strSQL = "UPDATE TradeMark SET TM16='2'," & _
         '                              "TM13=" & ChangeTStringToWString(textCP25) & " " & _
         '         "WHERE TM01 = '" & m_TM01 & "' AND " & _
         '               "TM02 = '" & m_TM02 & "' AND " & _
         '               "TM03 = '" & m_TM03 & "' AND " & _
         '               "TM04 = '" & m_TM04 & "' "
         '91.4.29 END
         cnnConnection.Execute strSql
      End If
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   m_NewCP09 = strCP09 'Added by Lydia 2022/02/10 新增C類收文號
   ' 案件性質
   strRvType = "1002"
   Select Case frm03020402_03.GetSelectResult
      Case "1": strRvType = "1002"
      Case "2": strRvType = "1403"
   End Select
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 組成SQL
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/07
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2003/09/05
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP35,CP43,CP49,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
   '2010/3/15 modify by sonia 已閉卷直接上發文日
   'strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP35,CP43,CP49,CP64) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
   strCP27 = ""
   If Me.lblClose.Caption <> "" Then strCP27 = "19221111"
   'Modify by Amy 2022/09/07 +CP36/37/40~42/80
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP35,CP43,CP49,CP64,CP27,CP36,CP37,CP40,CP41,CP42,CP80) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "','" & strCP27 & "'," & _
                    "'" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & ChgSQL(textCP42) & "'," & _
                    "'" & ChgSQL(textCP80) & "') "
   '2010/3/15 end
   cnnConnection.Execute strSql
   '92.11.20 ADD BY SONIA
   If strRvType = "1403" Then
       strSql = "Update CaseProgress Set CP24='2' Where CP09='" & strCP09 & "' "
       cnnConnection.Execute strSql
   End If
   '92.11.20 END
   ' 若有輸入承辦人時
   If IsEmptyText(textCP14) = False Then
      strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 有輸入本所期限時
   If IsEmptyText(textCP06) = False Then
      strSql = "UPDATE CASEPROGRESS SET CP06 = " & DBDATE(textCP06) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   ' 有輸入法定期限時
   If IsEmptyText(textCP07) = False Then
      strSql = "UPDATE CASEPROGRESS SET CP07 = " & DBDATE(textCP07) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   If Trim(Text11) <> "" Then
      strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
               "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
   End If
   
   'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   If Trim(textCP48) <> "" Then
            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
                     "WHERE CP09 = '" & strCP09 & "' "
            cnnConnection.Execute strSql
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
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新下一程序檔案件性質為催審的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "305"
   cnnConnection.Execute strSql
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 當使用者在前畫面選取2時, 更新下一程序檔案件性質為改變原處份的資料
   If frm03020402_03.GetSelectResult() = "2" Then
      strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "1403"
      cnnConnection.Execute strSql
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
      strNP08 = Empty
      If IsEmptyText(textCP06) = False Then: strNP08 = DBDATE(textCP06)
      strNP09 = Empty
      If IsEmptyText(textCP07) = False Then: strNP09 = DBDATE(textCP07)
      strNP14 = Empty
      strNP14 = GetRelatedPerson(m_CP09)
      ' 序號
      strNP22 = GetNextProgressNo()
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strNP08 & "," & strNP09 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/07
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2003/09/05
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
'                          strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                          strNP08 & "," & strNP09 & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
      cnnConnection.Execute strSql
      ' 延展, 使用宣誓, 刊登廣告, 繳年費, 催審, 提申, 收達不印接洽結案單
      Select Case textCF15
         Case "102", "105", "702", "708", "305", "998", "997":
         Case Else:
        'Modify By Cheng 2002/12/05
        '恢復列印接洽結案單
'            'Modify By Cheng 2002/01/15
'            '取消外商FCT列印接洽結案單
            ' 列印國內案件接洽及結案記錄單
'            g_PrtForm001.PrintForm strNP22, m_TM01, m_TM02, m_TM03, m_TM04
            'Add By Cheng 2003/06/23
            '新增列印接洽結案單資料
            pub_AddressListSN = pub_AddressListSN + 1
            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
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
   
   Dim SeekMonTM01 As String
   Dim SeekMonTM02 As String
   Dim SeekMonTM03 As String
   Dim SeekMonTM04 As String
   Dim rsA As New ADODB.Recordset
   'ADD BY nickc 2006/09/27 若是B類申請案，則代表是分割產生，要檢查分割的相關子案是否有准駁，若全都有，則將母案上閉卷
   If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" Then
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
   
   'Added by Morgan 2017/5/3 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strRvType
   End If
   'end 2017/5/3
   
 '911107 nick transation
  cnnConnection.CommitTrans
  
   Exit Function
EXITSUB:
If rsTmp.State <> adStateClosed Then rsTmp.Close
 Set rsTmp = Nothing
 
 '911107 nick transation
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
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

Private Sub Form_Unload(Cancel As Integer)
'   'Add By Cheng 2002/07/19
'   Set frm03020402_04 = Nothing
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

' 下一程序
Private Sub textCF15_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If IsEmptyText(textCF15) = False Then
      strSql = "SELECT * FROM CasePropertyMap " & _
               "WHERE CPM01 = '" & m_TM01 & "' AND " & _
                     "CPM02 = '" & textCF15 & "' "
      rsTmp.CursorLocation = adUseClient
      'edit by nickc 2005/08/04
      'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount <= 0 Then
         strTit = "檢核資料"
         strMsg = "案件性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCF15_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
         rsTmp.Close
         GoTo EXITSUB
      End If
      
      rsTmp.MoveFirst
      If m_TM10 < "010" Then
         If IsNull(rsTmp.Fields("CPM03")) = False Then
            textCF15_2 = rsTmp.Fields("CPM03")
         End If
      Else
         If IsNull(rsTmp.Fields("CPM04")) = False Then
            textCF15_2 = rsTmp.Fields("CPM04")
         End If
      End If
      rsTmp.Close
   End If
   
EXITSUB:
   Set rsTmp = Nothing
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

'Add By Sindy 2010/11/29
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
         textCP14_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
      End If
   End If
End Sub

' 核駁通知日 91.4.29 CANCEL
'Private Sub textCP25_Validate(Cancel As Boolean)
'   Dim sysDate As String
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'
'   Cancel = False
'   If IsEmptyText(textCP25) = False Then
'      ' 檢查是否為民國年
'      If CheckIsTaiwanDate(textCP25, False) = False Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "請輸入正確的核駁通知日"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP25_GotFocus
'      End If
'      ' 核准通知日不可超過系統日
'      sysDate = TAIWANDATE(Date)
'      If Val(textCP25) > Val(sysDate) Then
'         Cancel = True
'         strTit = "資料檢核"
'         strMsg = "核駁通知日不可超過系統日"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         textCP25_GotFocus
'      End If
'   End If
'End Sub

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
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
      End Select
   End If
End Sub

' 審查委員
Private Sub textCP35_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   If CheckLengthIsOK(textCP35, 32) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "審查委員資料內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP35_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
   End If

End Sub

'Add by Amy 2022/09/07 +對造頁籤
Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      textTM14.SetFocus
   Else
      textCP36.SetFocus
   End If
End Sub

Private Sub textCP36_GotFocus()
    InverseTextBox textCP36
End Sub

Private Sub textCP37_1_GotFocus()
    InverseTextBox textCP37_1
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label34, ":", "")
   Cancel = False
   'Modify by Amy 2025/01/17 原:140
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
End Sub

Private Sub textCP40_GotFocus()
    InverseTextBox textCP40
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label36, ":", "")
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      textCP40_GotFocus
   End If
End Sub

Private Sub textCP41_GotFocus()
    InverseTextBox textCP41
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label37, ":", "")
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      textCP41_GotFocus
   End If

End Sub

Private Sub textCP42_GotFocus()
    InverseTextBox textCP42
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label38, ":", "")
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP42, 600, True, strMsg) = False Then
      Cancel = True
      textCP42_GotFocus
   End If
End Sub

Private Sub textCP80_GotFocus()
    InverseTextBox textCP80
End Sub

Private Sub textCP80_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   Cancel = False
   If CheckLengthIsOK(textCP80, textCP80.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = Replace(Label39, ":", "") & "欄位內容太長"
      textCP80_GotFocus
   End If
End Sub
'end 2022/09/07

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         Cancel = True
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         textCP48_GotFocus
         Exit Sub
      End If
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
      ' 90.07.03 modify by louis (不檢查幾碼)
'      If Len(strTemp) > 4 Then
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
      
      ' 90.07.03 modify by louis (不檢查主張內容分類表)
      ' 檢查主張內容分類表
      'strSQL = "SELECT * FROM ClaimContents " & _
      '         "WHERE CC01 = '" & Right(strTemp, 1) & "'"
      'rsTmp.CursorLocation = adUseClient
      'rsTmp.Open strSQL, cnnConnection, adOpenDynamic
      'If rsTmp.RecordCount <= 0 Then
      '   Cancel = True
      '   strTit = "條款"
      '   strMsg = "條款內容<" & strTemp & ">不正確"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textCP49_GotFocus
      '   rsTmp.Close
      '   GoTo EXITSUB
      'End If
      'rsTmp.Close
      
      ' 檢查
      'Modify By Sindy 2012/7/5
'      strSql = "SELECT * FROM LAW " & _
'               "WHERE LW01 = '" & Left(strTemp, 3) & "' "
      strSql = "SELECT * FROM LAW " & _
               "WHERE LW01 = '" & Trim(strTemp) & "' "
      '2012/7/5 End
      rsTmp.CursorLocation = adUseClient
      'edit by nickc 2005/08/04
      'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
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
   
   If CheckLengthIsOK(textCP49, 300) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "條款資料內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP49_GotFocus
   End If
EXITSUB:
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
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
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
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textPrint_GotFocus
      End Select
   End If
End Sub

Private Sub textTM14_Change()
m_strLastTextTM14 = Me.textTM14.Text
End Sub

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim SysDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textTM14) = False Then
      ' 檢核是否為民國日期
      If CheckIsTaiwanDate(textTM14) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
      'sysDate = ChangeWStringToTString(ChangeWDateStringToWString(Date))
      'If Val(textTM14) > Val(sysDate) Then
      '   Cancel = True
      '   strTit = "資料檢核"
      '   strMsg = "公告日不可超過系統日"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textTM14_GotFocus
      'End If
   End If
End Sub

' 本所期限
Private Sub textCP06_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
         GoTo EXITSUB
      'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      Else
          textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06, True), 1)
      'end 2020/07/09
      End If
      'Add By Cheng 2002/03/11
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Cancel = True
         textCP06_GotFocus
         GoTo EXITSUB
      End If
      
      ' 申請國家為台灣時需檢查來函記錄檔
      If m_TM10 < "010" Then
         strSql = "SELECT * FROM MailRec " & _
                  "WHERE MR12 = '" & m_TM01 & "' AND " & _
                        "MR13 = '" & m_TM02 & "' AND " & _
                        "MR14 = '" & m_TM03 & "' AND " & _
                        "MR15 = '" & m_TM04 & "' AND " & _
                        "MR02 = " & DBDATE(m_CP05) & " "
         rsTmp.CursorLocation = adUseClient
         'edit by nickc 2005/08/04
         'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("MR16")) = False Then
               If IsEmptyText(rsTmp.Fields("MR16")) = False Then
                  If DBDATE(textCP06) <> rsTmp.Fields("MR16") Then
                     strTit = "資料檢核"
                     strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
                     nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
                     If nResponse = vbCancel Then
                        Cancel = True
                        textCP06_GotFocus
                        rsTmp.Close
                        GoTo EXITSUB
                     End If
                  End If
               End If
            End If
         'add by sonia 2018/2/8 非電子公文要顯示訊息
         ElseIf m_DocNo = "" Then
            strTit = "資料檢核"
            strMsg = "來函記錄中無該筆記錄"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP06_GotFocus
               rsTmp.Close
               GoTo EXITSUB
            End If
         'end 2018/2/8
         End If
         rsTmp.Close
      End If
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim rsTmp As New ADODB.Recordset
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         'add by nickc 2005/09/06
         Cancel = True
         GoTo EXITSUB
      End If
      ' 申請國家為台灣時需檢查來函記錄檔
      If m_TM10 < "010" Then
         strSql = "SELECT * FROM MailRec " & _
                  "WHERE MR12 = '" & m_TM01 & "' AND " & _
                        "MR13 = '" & m_TM02 & "' AND " & _
                        "MR14 = '" & m_TM03 & "' AND " & _
                        "MR15 = '" & m_TM04 & "' AND " & _
                        "MR02 = " & ChangeTStringToWString(m_CP05) & " "
         rsTmp.CursorLocation = adUseClient
         'edit by nickc 2005/08/04
         'rsTmp.Open StrSql, cnnConnection, adOpenDynamic
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("MR17")) = False Then
               If IsEmptyText(rsTmp.Fields("MR17")) = False Then
                  If ChangeTStringToWString(textCP07) <> rsTmp.Fields("MR17") Then
                     strTit = "資料檢核"
                     strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
                     nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
                     If nResponse = vbCancel Then
                        Cancel = True
                        textCP07_GotFocus
                        rsTmp.Close
                        GoTo EXITSUB
                     End If
                  End If
               End If
            End If
         'add by sonia 2018/2/8 非電子公文要顯示訊息
         ElseIf m_DocNo = "" Then
            strTit = "資料檢核"
            strMsg = "來函記錄中無該筆記錄"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               Cancel = True
               textCP07_GotFocus
               rsTmp.Close
               GoTo EXITSUB
            End If
         'end 2018/2/8
         End If
         rsTmp.Close
      End If
   End If
EXITSUB:
   Set rsTmp = Nothing
End Sub

' 審定號數
Private Sub textTM15_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   Cancel = False
            
   If IsEmptyText(textTM15) = False Then
      'Add By Sindy 2010/9/1
      '檢查審定號所輸入的長度是否正確
      If bolNewAppNoFormat Then
         '2011/1/14 MODIFY BY SONIA 台灣核駁審定號0+6碼數字
         'If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10) = False Then
         'Add By Sindy 2017/5/17 + strRetrunText
         If PUB_ChkTm12Tm15Length("2", textTM15, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, "2", , strRetrunText) = False Then
            Cancel = True
            textTM15_GotFocus
            Exit Sub
         'Add By Sindy 2017/5/17
         Else
            textTM15 = strRetrunText
         '2017/5/17 END
         End If
      Else
         If IsNumeric(Mid(textTM15, 1, 8)) = False Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "請輸入正確的審定號數"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM15_GotFocus
         End If
      End If
   End If
End Sub

Private Sub textTM16S_Change()
'Modify By Cheng 2002/07/22
'm_strLastTextTM16S = Me.textTM16S.Text
End Sub

Private Sub textTM16S_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否更新基本檔目前準駁
Private Sub textTM16S_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/07/22
   '取消檢查
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'
'   If IsEmptyText(textTM16S) = False Then
'      Select Case textTM16S
'         Case "Y", "N":
'         Case Else:
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM16S_GotFocus
'      End Select
'   End If
End Sub

Private Sub textTM17_Change()
'Modify By Cheng 2002/07/22
'm_strLastTextTM17 = Me.textTM17.Text
End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/07/22
'   Dim strTit As String
'   Dim strMsg As String
'   Dim nResponse
'   Cancel = False
'
'   If IsEmptyText(textTM17) = False Then
'      Select Case textTM17
'         Case "Y", "N":
'         Case Else:
'            Cancel = True
'            strTit = "資料檢核"
'            strMsg = "只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM17_GotFocus
'      End Select
'   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean
   
   CheckDataValid = False
      
      'Memo by Amy 2022/09/26 對造頁籤改為「關係案」,對造文字改「對方」
      'Modify by Amy 2022/09/07 +SSTab1.Tab = 0,因加對造頁籤
      'Add By Sindy 2010/12/24
      If Me.textTM15.Enabled = True Then
         Cancel = False
         textTM15_Validate Cancel
         If Cancel = True Then
            SSTab1.Tab = 0
            textTM15.SetFocus
            Exit Function
         End If
      End If
      
      'Add by Amy 2025/01/17 避免輸完直接按Enter鍵不會檢查
      If Me.textCP37_1.Enabled = True Then
         Cancel = False
         textCP37_1_Validate Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            textCP37_1.SetFocus
            Exit Function
         End If
      End If
      
      If Me.textCP40.Enabled = True Then
         Cancel = False
         textCP40_Validate Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            textCP40.SetFocus
            Exit Function
         End If
      End If
      
      If Me.textCP41.Enabled = True Then
         Cancel = False
         textCP41_Validate Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            textCP41.SetFocus
            Exit Function
         End If
      End If
      
      If Me.textCP42.Enabled = True Then
         Cancel = False
         textCP42_Validate Cancel
         If Cancel = True Then
            SSTab1.Tab = 1
            textCP42.SetFocus
            Exit Function
         End If
      End If
      'end 2025/01/17
      
'add by nickc 2005/08/04
      If m_blnClkChgButton = False Then
          MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
          SSTab1.Tab = 0
          Me.cmdMod.SetFocus
          GoTo EXITSUB
      End If
   
   
   ' 機關文號不可空白
   If IsEmptyText(textCP08) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入機關文號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      SSTab1.Tab = 0
      textCP08.SetFocus
      GoTo EXITSUB
   End If
   'Modify By Cheng 2002/05/08
   '若前一畫面(frm03020402_3)結果欄為"3"時, 不檢查審定號數是否空白
   If frm03020402_03.GetSelectResult <> "3" Then
      ' 審定號數不可空白
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" And IsEmptyText(textTM15) = True Then
      If (m_CP10 = "101" Or m_CP10 = "308") And IsEmptyText(textTM15) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入審定號數"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textTM15.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Modify By Cheng 2002/05/08
   '若前一畫面(frm03020402_3)結果欄為"3"時, 不檢查審定號數是否空白
   If frm03020402_03.GetSelectResult <> "3" Then
      ' 公告日不可空白
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" And IsEmptyText(textTM14) = True Then
      If (m_CP10 = "101" Or m_CP10 = "308") And IsEmptyText(textTM14) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textTM14.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 核駁通知日不可為空白 91.4.29 CANCEL
   'If IsEmptyText(textCP25) = True Then
   '   strTit = "資料檢核"
   '   strMsg = "請輸入核駁通知日"
   '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   '   textCP25.SetFocus
   '   GoTo EXITSUB
   'End If
   
   ' 案件性質為申請時下一程序不可空白
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      If IsEmptyText(textCF15) = True Then
         'add by sonia 2019/4/18 若前畫面選擇3.申請駁回可不輸下一程序,提醒就好
         If frm03020402_03.textResult = "3" Then
            If MsgBox("申請駁回，是否要輸下一程序及期限？", vbYesNo) = vbYes Then
               SSTab1.Tab = 0
               textCF15.SetFocus
               GoTo EXITSUB
            End If
         Else
         'end 2019/4/18
            strTit = "資料檢核"
            strMsg = "請輸入下一程序"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            SSTab1.Tab = 0
            textCF15.SetFocus
            GoTo EXITSUB
         End If  'add by sonia 2019/4/18
      End If
   End If
   
   'Add By Sindy 2012/4/17
   '檢查來函期限--日期
   If m_TM10 = 台灣國家代號 Then
      If Me.Option4(2).Value = True Then
         If Me.Text12.Text = "" Then
            MsgBox "請輸入來函期限!!!", vbExclamation + vbOKOnly
            SSTab1.Tab = 0
            Me.Text12.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   ' 有輸入下一程序
   If IsEmptyText(textCF15) = False Then
      ' 本所期限不可空白
      If IsEmptyText(textCP06) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      ' 法定期限不可空白
      If IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "請輸入法定期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCP07.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/03/11
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         SSTab1.Tab = 0
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   ' 本所期限的日期不可超過法定期限
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
      If Val(textCP06) > Val(textCP07) Then
         strTit = "資料檢核"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         SSTab1.Tab = 0
         textCP06.SetFocus
         GoTo EXITSUB
      End If
   End If
   'Add By Cheng 2002/05/07
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         SSTab1.Tab = 0
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2012/7/9 以防修改期限天數或月數,重新計算期限
   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         Text10.SetFocus
         Exit Function
      End If
   End If
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         Text11.SetFocus
         Exit Function
      End If
   End If
   '2012/7/9 End
   
   'add by sonia 2018/2/8
   If Me.textCP06.Enabled = True Then
      Cancel = False
      textCP06_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textCP06.SetFocus
         Exit Function
      End If
   End If
   If Me.textCP07.Enabled = True Then
      Cancel = False
      textCP07_Validate Cancel
      If Cancel = True Then
         SSTab1.Tab = 0
         textCP07.SetFocus
         Exit Function
      End If
   End If
   'end 2018/2/8
   'end 2022/09/07
   
    'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTMBM07_1_Change()
m_strLastTextTMBM07_1 = Me.textTMBM07_1.Text
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_Change()
m_strLastTextTMBM07_2 = Me.textTMBM07_2.Text
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

'91.4.29 CANCEL
'Private Sub textCP25_GotFocus()
'   InverseTextBox textCP25
'End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP35_GotFocus()
   InverseTextBox textCP35
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP49_GotFocus()
   InverseTextBox textCP49
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

'Add By Cheng 2002/02/01
'保留上一次輸入的資料
Public Sub SetLastData()
Me.textTM14.Text = "" & m_strLastTextTM14
Me.textTMBM07_1.Text = "" & m_strLastTextTMBM07_1
Me.textTMBM07_2.Text = "" & m_strLastTextTMBM07_2
'Modify By Cheng 2002/07/22
'Me.textTM16S.Text = "" & m_strLastTextTM16S
'Me.textTM17.Text = "" & m_strLastTextTM17
End Sub

'Add By Cheng 2002/02/01
'清空上一次輸入的資料
Public Sub ClearLastData()
m_strLastTextTM14 = Empty
m_strLastTextTMBM07_1 = Empty
m_strLastTextTMBM07_2 = Empty
'Modify By Cheng 2002/07/22
'm_strLastTextTM16S = Empty
'm_strLastTextTM17 = Empty
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
               textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
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
   strFromDate = DBDATE(frm03020402_01.textCP05)
   
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
         textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
      End If
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
Dim strFromDate As String '期限起算日
   
   'strFromDate = DBDATE(textCP05)
   strFromDate = DBDATE(frm03020402_01.textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
     
   ' 案件性質
   strRvType = "1002"
   Select Case frm03020402_03.GetSelectResult
      Case "1": strRvType = "1002"
      Case "2": strRvType = "1403"
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
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

