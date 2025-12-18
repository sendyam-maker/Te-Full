VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010408_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "其他來函輸入"
   ClientHeight    =   5748
   ClientLeft      =   -156
   ClientTop       =   972
   ClientWidth     =   9348
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   9348
   Begin TabDlg.SSTab SSTab1 
      Height          =   3645
      Left            =   60
      TabIndex        =   56
      Top             =   2100
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   6435
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm02010408_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label23"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label22"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label12"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(8)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label26"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label24"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label25"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label7"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label14"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label29"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label32"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label8"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label4"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textCP14_2"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "grdList"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textPrint"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textEditPrint"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textRvType_2"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textBTTM"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP25"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP26"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP48"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP14"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textTM29"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textRvType"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP06"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP07"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "TextCP64_1"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTM12_new"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Frame1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textCF15"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textCP08"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "textCF15_2"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "MaskEdBox1"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).ControlCount=   43
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm02010408_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label30"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label27"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label21"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label20"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label19"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label18"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label17"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label13"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label10"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label28"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "textCP42"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "textCP39"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "textCP37"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "textCP40"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "textCP37_1"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "textCP64"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "textCP41"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "textCP38"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "textCP36"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).ControlCount=   19
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   30
         Left            =   1350
         TabIndex        =   93
         Top             =   390
         Width           =   30
         _ExtentX        =   42
         _ExtentY        =   42
         _Version        =   393216
         PromptChar      =   "_"
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2340
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   615
         Width           =   1092
      End
      Begin VB.TextBox textCP08 
         Height          =   300
         Left            =   4740
         MaxLength       =   40
         TabIndex        =   4
         Top             =   615
         Width           =   3165
      End
      Begin VB.ComboBox textCF15 
         Height          =   276
         Left            =   990
         TabIndex        =   3
         Top             =   615
         Width           =   1332
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   990
         TabIndex        =   91
         Top             =   900
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   5
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3690
         TabIndex        =   90
         Top             =   900
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   252
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   12
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   10
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   8
            Top             =   150
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   11
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   9
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox textTM12_new 
         Height          =   264
         Left            =   4740
         MaxLength       =   20
         TabIndex        =   1
         Top             =   345
         Width           =   1725
      End
      Begin VB.TextBox TextCP64_1 
         Height          =   300
         Left            =   7440
         MaxLength       =   40
         TabIndex        =   2
         Top             =   345
         Width           =   1725
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   4590
         MaxLength       =   7
         TabIndex        =   14
         Top             =   1410
         Width           =   1155
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   990
         MaxLength       =   7
         TabIndex        =   13
         Top             =   1410
         Width           =   1155
      End
      Begin VB.TextBox textRvType 
         Height          =   264
         Left            =   990
         MaxLength       =   4
         TabIndex        =   0
         Top             =   345
         Width           =   732
      End
      Begin VB.TextBox textTM29 
         Height          =   264
         Left            =   7680
         MaxLength       =   1
         TabIndex        =   17
         Top             =   1680
         Width           =   732
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   990
         MaxLength       =   6
         TabIndex        =   18
         Top             =   1950
         Width           =   732
      End
      Begin VB.TextBox textCP48 
         Height          =   264
         Left            =   4350
         MaxLength       =   7
         TabIndex        =   19
         Top             =   1950
         Width           =   1155
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   1320
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2220
         Width           =   372
      End
      Begin VB.TextBox textCP25 
         Height          =   264
         Left            =   4350
         MaxLength       =   7
         TabIndex        =   22
         Top             =   2220
         Width           =   1155
      End
      Begin VB.TextBox textBTTM 
         Height          =   264
         Left            =   7020
         MaxLength       =   15
         TabIndex        =   20
         Top             =   1950
         Width           =   2055
      End
      Begin VB.TextBox textRvType_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   1740
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   345
         Width           =   1692
      End
      Begin VB.TextBox textEditPrint 
         Height          =   264
         Left            =   5580
         MaxLength       =   1
         TabIndex        =   16
         Top             =   1680
         Width           =   372
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   990
         MaxLength       =   1
         TabIndex        =   15
         Top             =   1680
         Width           =   372
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73065
         MaxLength       =   200
         TabIndex        =   23
         Top             =   345
         Width           =   6795
      End
      Begin VB.TextBox textCP38 
         Height          =   264
         Left            =   -73065
         TabIndex        =   25
         Top             =   900
         Width           =   6795
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73065
         TabIndex        =   29
         Top             =   1785
         Width           =   6795
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   1044
         Left            =   1008
         TabIndex        =   94
         Top             =   2496
         Width           =   8148
         _ExtentX        =   14372
         _ExtentY        =   1842
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
      Begin MSForms.TextBox textCP64 
         Height          =   1245
         Left            =   -73065
         TabIndex        =   31
         Top             =   2340
         Width           =   6795
         VariousPropertyBits=   -1467989989
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "11986;2205"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   825
         Left            =   -73065
         TabIndex        =   27
         Top             =   615
         Width           =   6795
         VariousPropertyBits=   679493659
         ScrollBars      =   2
         Size            =   "11986;1455"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73065
         TabIndex        =   28
         Top             =   1515
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   1740
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1950
         Width           =   1185
         VariousPropertyBits=   679493663
         ForeColor       =   -2147483641
         MaxLength       =   20
         Size            =   "2090;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37 
         Height          =   300
         Left            =   -73065
         TabIndex        =   24
         Top             =   615
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP39 
         Height          =   300
         Left            =   -73065
         TabIndex        =   26
         Top             =   1230
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73065
         TabIndex        =   30
         Top             =   2055
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   615
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   3690
         TabIndex        =   72
         Top             =   615
         Width           =   975
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   92
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label29 
         Caption         =   "新申請案號 :"
         Height          =   255
         Left            =   3690
         TabIndex        =   89
         Top             =   345
         Width           =   1005
      End
      Begin VB.Label Label28 
         Caption         =   "(開庭時間,第X法庭)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   88
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "收文文號 :"
         Height          =   255
         Left            =   6570
         TabIndex        =   87
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label10 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   86
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   85
         Top             =   645
         Width           =   1575
      End
      Begin VB.Label Label17 
         Caption         =   "對造案件英文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   84
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label Label18 
         Caption         =   "對造案件日文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   83
         Top             =   1260
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   82
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   81
         Top             =   1815
         Width           =   1575
      End
      Begin VB.Label Label21 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   80
         Top             =   2055
         Width           =   1575
      End
      Begin VB.Label Label27 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   -74745
         TabIndex        =   79
         Top             =   2355
         Width           =   975
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   675
         Width           =   1575
      End
      Begin VB.Label Label7 
         Caption         =   "本案期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   77
         Top             =   2490
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   3690
         TabIndex        =   76
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   75
         Top             =   1410
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "來函性質 :"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   73
         Top             =   345
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Left            =   8460
         TabIndex        =   71
         Top             =   1710
         Width           =   705
      End
      Begin VB.Label Label15 
         Caption         =   "是否閉卷 :"
         Height          =   255
         Left            =   6780
         TabIndex        =   70
         Top             =   1710
         Width           =   975
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   3480
         TabIndex        =   68
         Top             =   1950
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   1770
         TabIndex        =   67
         Top             =   2220
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   120
         TabIndex        =   66
         Top             =   2220
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   $"frm02010408_4.frx":0038
         Height          =   360
         Index           =   4
         Left            =   6330
         TabIndex        =   65
         Top             =   1950
         Width           =   570
      End
      Begin VB.Label Label1 
         Caption         =   "專用權消滅日 :"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   64
         Top             =   2220
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "(Y:修改)"
         Height          =   255
         Left            =   5970
         TabIndex        =   63
         Top             =   1710
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "是否修改定稿 :"
         Height          =   255
         Left            =   4320
         TabIndex        =   62
         Top             =   1710
         Width           =   1245
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   1440
         TabIndex        =   60
         Top             =   1710
         Width           =   2745
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7200
      TabIndex        =   34
      Top             =   15
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6375
      TabIndex        =   32
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8430
      TabIndex        =   35
      Top             =   15
      Width           =   800
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   1815
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1170
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   1290
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5610
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   435
      Width           =   2532
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   5145
      TabIndex        =   33
      Top             =   15
      Width           =   1200
   End
   Begin MSForms.ComboBox cmbTM05 
      Height          =   300
      Left            =   1140
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   720
      Width           =   7992
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "14097;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5610
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   1815
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
      Left            =   1170
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1020
      Width           =   7992
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "14097;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4650
      TabIndex        =   55
      Top             =   1815
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   90
      TabIndex        =   54
      Top             =   1815
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   90
      TabIndex        =   53
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4650
      TabIndex        =   52
      Top             =   1560
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   90
      TabIndex        =   51
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   90
      TabIndex        =   50
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   90
      TabIndex        =   49
      Top             =   435
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   90
      TabIndex        =   48
      Top             =   1290
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4650
      TabIndex        =   47
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Index           =   5
      Left            =   4650
      TabIndex        =   46
      Top             =   435
      Width           =   855
   End
End
Attribute VB_Name = "frm02010408_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/18 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo by Amy 2021/12/29 Form2.0已修改 cmbTm05/textTM23/textCP13/textCP14_2/textCP37/textCP37_1/textCP39/textCP40/textCP42/textCP64/grdList
'Memo By Sindy 2012/12/4 智權人員欄已修改
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
' 新增的收文號
Dim strCP09 As String

' 宣告欄位內容結構
Private Type FIELDITEM
   fiName As String
   fiNewData As String
   fiType As Integer
End Type
Dim m_FieldList() As FIELDITEM
Dim m_FieldCount As Integer

' 暫時存放 CF15
Dim m_CF15 As String
'
Dim m_CurrSel As Integer
'add by nickc 2006/09/08
Dim m_CP27 As String

Dim m_CP110 As String  '2008/11/12 add by sonia
Dim t_CP110 As String  '2008/11/12 add by sonia
'Add By Sindy 2012/3/5
Dim m_bolClose As String
Dim m_strCloseDT As String
Dim m_strCloseReason As String
'2012/3/5 End
'Added by Morgan 2017/4/24 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
Public m_NewCP10 As String
'end 2017/4/24
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim strLD18 As String 'Add By Sindy 2019/12/19 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/19 FC代理人
Dim m_TM23 As String 'Add By Sindy 2019/12/19 申請人1

'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 清除欄位串列
Private Sub ClearFieldList()
   If m_FieldCount > 0 Then
      Erase m_FieldList
   End If
   m_FieldCount = 0
End Sub
' 設定欄位的內容
Private Sub SetFieldData(ByVal strFieldName As String, ByVal strFieldData As String, ByVal nFieldType As Integer)
   Dim nPos As Integer
   Dim bFind As Boolean
   bFind = False
   For nPos = 0 To m_FieldCount - 1
      If m_FieldList(nPos).fiName = strFieldName Then
         bFind = True
         m_FieldList(nPos).fiNewData = strFieldData
         m_FieldList(nPos).fiType = nFieldType
      End If
   Next nPos
   If bFind = False Then
      ReDim Preserve m_FieldList(m_FieldCount + 1)
      m_FieldList(m_FieldCount).fiName = strFieldName
      m_FieldList(m_FieldCount).fiNewData = strFieldData
      m_FieldList(m_FieldCount).fiType = nFieldType
      m_FieldCount = m_FieldCount + 1
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
   frm02010408_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010408_3
   Unload frm02010408_2
   Unload frm02010408_1
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim m_StrTo As String, m_StrSub As String, m_StrCont As String 'Added by Lydia 2022/05/30 整理frm880005改用寄信模組
   
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
   
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
        'Modify By Cheng 2002/11/07
'      'OnSaveData
      cmdOK.Enabled = False 'Add By Sindy 2022/1/19 秀玲:Enter按了2次~~
      If OnSaveData = False Then
         cmdOK.Enabled = True 'Add By Sindy 2022/1/19
         MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         Screen.MousePointer = vbDefault
         Exit Sub
      End If
      
      'add by Toni 2008/10/27
      'Modify By Sindy 2023/3/28 控管台灣的才發Mail ex:TF-000870-1-06
      If (textCP10 = "準備程序" Or textCP10 = "言詞辯論") And m_TM10 = "000" Then
         If textRvType = "1203" Or textRvType = "1204" Then
            'Modified by Lydia 2022/05/30 整理frm880005改用寄信模組
            'Load frm880005
            ''2008/11/7 modify by sonia
            ''frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & textCP13 & "@taie.com.tw"
            ''Modify By Sindy 2012/8/16 開庭通知發mail對象,若為FCT案件再增加Pub_GetSpecMan("Q1")
            'frm880005.txtEmail(0).Text = Pub_GetSpecMan("Q") & ";" & _
                  IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & _
                  IIf(m_TM01 = "FCT", ";" & Pub_GetSpecMan("Q1"), "")
            ''2008/11/7 end
            ''2008/11/12 modify by sonia 再抓時間地點,法院案號,律師,案號加-
            'frm880005.txtEmail(1).Text = "開庭通知--來函案件：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
           ' frm880005.txtEmail(2).Text = "本所案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & vbCrLf & _
                                          "案件名稱：" & cmbTM05 & vbCrLf & _
                                          "案件性質：" & textRvType_2 & vbCrLf & _
                                          "申請人　：" & textTM23.Text & vbCrLf & _
                                          "承辦人　：" & textCP14_2.Text & vbCrLf & _
                                          "智權人員：" & textCP13.Text & vbCrLf & _
                                          "法定期限：" & Val(Mid(DBDATE(textCP07), 1, 4)) - 1911 & " 年 " & Mid(DBDATE(textCP07), 5, 2) & " 月 " & Mid(DBDATE(textCP07), 7, 2) & " 日 " & vbCrLf & _
                                          "時間地點：" & textCP64 & vbCrLf & _
                                          "法院案號：" & textCP08 & vbCrLf & _
                                          "律　　師：" & t_CP110
            'frm880005.Form_Activate: DoEvents
            'frm880005.cmdOK_Click 0: DoEvents
            'Modify By Sindy 2023/12/8 法律所調整內專行政訴訟開庭通知之系統通知信也請一併轉陳亮之; 商標一併調整
            'Modified by Lydia 2024/10/30 串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
            'm_StrTo = Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & _
            '      IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
            '      'IIf(m_TM01 = "FCT", ";" & Pub_GetSpecMan("Q1"), "") & IIf(textCP14 <> "", ";" & textCP14, "")
            m_StrTo = PUB_GetLosCL02list(m_TM01, m_TM02, m_TM03, m_TM04)
            m_StrTo = IIf(m_StrTo <> "", m_StrTo & ";", "") & Pub_GetSpecMan("Q") & ";" & Pub_GetSpecMan("Q1") & ";" & _
                  IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))
            'end 2024/10/30
            
            m_StrSub = "開庭通知--來函案件：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
            m_StrCont = "本所案號：" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & vbCrLf & _
                                          "案件名稱：" & cmbTM05 & vbCrLf & _
                                          "案件性質：" & textRvType_2 & vbCrLf & _
                                          "申請人　：" & textTM23.Text & vbCrLf & _
                                          "承辦人　：" & textCP14_2.Text & vbCrLf & _
                                          "智權人員：" & textCP13.Text & vbCrLf & _
                                          "法定期限：" & Val(Mid(DBDATE(textCP07), 1, 4)) - 1911 & " 年 " & Mid(DBDATE(textCP07), 5, 2) & " 月 " & Mid(DBDATE(textCP07), 7, 2) & " 日 " & vbCrLf & _
                                          "時間地點：" & textCP64 & vbCrLf & _
                                          "法院案號：" & textCP08 & vbCrLf & _
                                          "律　　師：" & t_CP110
            PUB_SendMail strUserNum, m_StrTo, m_CP09, m_StrSub, m_StrCont
            'end 2022/05/30
         End If
      End If
      'end 2008/10/27
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      Unload frm02010408_3
      Unload frm02010408_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010408_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/24 電子公文
      'frm02010408_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010408_1
         frm02010412.GoNext
      Else
         frm02010408_1.Show
         Unload Me
      End If
      'end 2017/4/24
   End If
End Sub

Private Sub cmdRelate_Click()
   Where1103ComeFrom Me, m_TM01, m_TM02, m_TM03, m_TM04
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
   textCP05S.BackColor = &H8000000F
   textCP05.BackColor = &H8000000F
   textCP09.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   textRvType_2.BackColor = &H8000000F
   
   ' 先設定專用權消滅日為不可輸入
   textCP25.Locked = True
   textCP25.BackColor = &H8000000F
   textCP25.TabStop = False
   
   textCF15.AddItem "補正"
   textCF15.AddItem "申請意見書"
   '2011/6/9 還原 BY SONIA TD,TM的1602移至商爭被異議輸入
   textCF15.AddItem "異議答辯"
   textCF15.AddItem "評定答辯"
   textCF15.AddItem "廢止答辯"
   textCF15.AddItem "補充答辯"
   textCF15.AddItem "補充理由"
   textCF15.AddItem "變更"
   textCF15.AddItem "領證"
   
   MoveFormToCenter Me
   
   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010408_1.m_strIR01
   m_strIR02 = frm02010408_1.m_strIR02
   m_strIR03 = frm02010408_1.m_strIR03
   m_strIR04 = frm02010408_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
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
   
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
   'Add By Cheng 2002/07/18
   Set frm02010408_4 = Nothing
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

' 取得商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
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
      ' 申請國家
      m_TM10 = ""
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
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
      m_TM23 = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23") 'Add By Sindy 2019/12/19
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'Add By Sindy 2019/12/19
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2019/12/19 END
      
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("TM77"))
      
      'Add By Sindy 2012/3/5
      m_bolClose = ""
      If IsNull(rsTmp.Fields("TM29")) = False Then
         textTM29 = rsTmp.Fields("TM29")
         m_bolClose = rsTmp.Fields("TM29")
      End If
      m_strCloseDT = ""
      If IsNull(rsTmp.Fields("TM30")) = False Then
         m_strCloseDT = rsTmp.Fields("TM30")
      End If
      m_strCloseReason = ""
      If IsNull(rsTmp.Fields("TM31")) = False Then
         m_strCloseReason = rsTmp.Fields("TM31")
      End If
      '2012/3/5 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 取得服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 取得商標基本檔的相關項目
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請國家
      m_TM10 = ""
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_TM10 = rsTmp.Fields("SP09")
      End If
      ' 案件名稱(中)
      If IsNull(rsTmp.Fields("SP05")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP05")
      End If
      ' 案件名稱(英)
      If IsNull(rsTmp.Fields("SP06")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP06")
      End If
      ' 案件名稱(日)
      If IsNull(rsTmp.Fields("SP07")) = False Then
         cmbTM05.AddItem rsTmp.Fields("SP07")
      End If
      ' 顯示商標名稱
      If cmbTM05.ListCount > 0 Then
         cmbTM05.ListIndex = 0
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("SP08")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("SP08"))
      End If
      
      'Add By Sindy 2019/12/25
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_TM44 = rsTmp.Fields("SP26")
      End If
      '2019/12/25 END
      
      ' BTTM modify by sonia 91.10.10
      If IsNull(rsTmp.Fields("SP50")) = False Then
         textBTTM = rsTmp.Fields("SP50")
      End If
      'add by nickc 2006/11/21
      textPrint = CheckStr(rsTmp.Fields("SP72"))
      
      'Add By Sindy 2012/3/5
      m_bolClose = ""
      If IsNull(rsTmp.Fields("SP15")) = False Then
         textTM29 = rsTmp.Fields("SP15")
         m_bolClose = rsTmp.Fields("SP15")
      End If
      m_strCloseDT = ""
      If IsNull(rsTmp.Fields("SP16")) = False Then
         m_strCloseDT = rsTmp.Fields("SP16")
      End If
      m_strCloseReason = ""
      If IsNull(rsTmp.Fields("SP17")) = False Then
         m_strCloseReason = rsTmp.Fields("SP17")
      End If
      '2012/3/5 End
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   m_TM10 = Empty
   m_CP13 = Empty
   m_CP12 = Empty
   'Add By Cheng 2002/07/17
   m_CP10 = Empty
   m_CP08 = Empty
      
   ' 本所案號
   textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   ' 收文號
   textCP09 = m_CP09

   Select Case m_TM01
      Case "T", "TF", "CFT", "FCT":
         'Add by Morgan 2003/11/25
         Me.Label13.Visible = False
         Me.textCP37.Visible = False
         Me.textCP37.Enabled = False
         Me.Label17.Visible = False
         Me.textCP38.Visible = False
         Me.textCP38.Enabled = False
         Me.Label18.Visible = False
         Me.textCP39.Visible = False
         Me.textCP39.Enabled = False
         '---End
         QueryTradeMark
      Case Else:
         'Add by Morgan 2003/11/25
         Me.Label30.Visible = False
         Me.textCP37_1.Visible = False
         Me.textCP37_1.Enabled = False
         '---End
         QueryServicePractice
   End Select
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔資料
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
      ' 收文日
      If IsNull(rsTmp.Fields("CP05")) = False Then
         textCP05 = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 機關文號
      If IsNull(rsTmp.Fields("CP08")) = False Then
         m_CP08 = rsTmp.Fields("CP08")
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
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區   nick 91.08.22
      If IsNull(rsTmp.Fields("CP12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      
      '2008/11/12 add by sonia 抓出庭律師,商標輸在出名代理人欄
      m_CP110 = CheckStr(rsTmp.Fields("cp110"))
      If m_CP110 <> "" Then
         strSql = "select st01,st02,OA03 from staff,ouragent where instr('" & m_CP110 & "',st01)>0 and oa01(+)='" & m_TM01 & "' and oa02(+)=st01 order by 3 , 1 "
         CheckOC
         adoRecordset.CursorLocation = adUseClient
         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If adoRecordset.RecordCount > 0 Then
            Do While Not adoRecordset.EOF
               If InStr(m_CP110, CheckStr(adoRecordset.Fields(0))) > 0 Then
                  t_CP110 = t_CP110 & CheckStr(adoRecordset.Fields(1)) & "、"
               End If
               adoRecordset.MoveNext
            Loop
         End If
         CheckOC
         If Right(t_CP110, 1) = "、" Then
            t_CP110 = Mid(t_CP110, 1, Len(t_CP110) - 1)
         End If
      End If
      '2008/11/12 end
      
      'Add by Morgan 2003/11/25
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
      
      '---End
         'add by nickc 2006/09/08
         m_CP27 = CheckStr(rsTmp.Fields("CP27"))
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
      'Added by Lydia 2023/10/18
      If grdList.Rows >= 2 Then
         grdList.FixedRows = 1
      End If
      'end 2023/10/18
   End If
   rsTmp.Close
   
   ' 系統類別為TM時, BTTM欄位才允許輸入
   If m_TM01 = "TM" Then
      EnableTextBox textBTTM, True
   Else
      EnableTextBox textBTTM, False
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
   
   'Added by Morgan 2017/4/24 電子公文
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
   'end 2017/4/24
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   'edit by nickc 2006/11/21
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
End Sub

'Modify By Cheng 2002/11/07
'Private sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP06 As String
   Dim strCP07 As String
   Dim strCP10 As String
   Dim strCP12 As String
   Dim strCP14 As String
   Dim strCP27 As String
   Dim strCP48 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strTmp As String
   Dim bFirst As Boolean
   Dim strNP22 As String
   Dim strCP64 As String
   'Add by Amy 2017/11/13
   Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
   Dim bolUpdCP As Boolean '是否更新進度檔
   Dim strMailMsg As String 'Add By Sindy 2022/1/19
   
'Add By Sindy 2012/1/12
Dim ET01 As String, ET02 As String, ET03 As String, bolEdit As Boolean
Dim bolEmail As Boolean, bolPlusPaper As Boolean, iCopy As Integer
'2012/1/12 End
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   
   ' 本所期限
   strCP06 = Empty
   If IsEmptyText(textCP06) = False Then
      strCP06 = DBDATE(textCP06)
      'add by Toni 2008/10/27
      textCP06.Tag = DBDATE(textCP06)
      'end 2008/10/27
   End If
   ' 法定期限
   strCP07 = Empty
   If IsEmptyText(textCP07) = False Then
      strCP07 = DBDATE(textCP07)
      'add by Toni 2008/10/27
      textCP07.Tag = DBDATE(textCP07)
      'end 2008/10/24
   End If
   
   ' 案件性質為來函性質
   strCP10 = Trim(textRvType)
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 發文日為系統日
   'strCP27 = DBDATE(Date)
   strCP27 = Empty
   '2009/2/26 MODIFY BY SONIA 剔除1709智慧局答辯函T-139214
   'If IsEmptyText(textCF15) = True Then
   'Modify By Sindy 2009/10/26
   'If IsEmptyText(textCF15) = True And textRvType <> "1709" Then
   '2012/9/20 modify by sonia 其他來函也由原承辦人分析 T-169807 不上發文日
   '2013/1/29 modify by sonia 1724通知已轉他所由原承辦人分析 T-180403 不上發文日
   If (IsEmptyText(textCF15) = True And textRvType <> "1709" And textRvType <> "1706" And textRvType <> "1724") Or textRvType = "1718" Then
      strCP27 = DBDATE(SystemDate())
   End If
   ' 有下一程序時, 承辦人為所輸入的承辦人, 否則為LogoOn的員工代號
   ' 承辦期限也是類似
   '2009/6/9 modify by sonia 1709智慧局答辯函改由原承辦人分析
   'If IsEmptyText(textCF15) = False Then
   'Modify By Sindy 2009/10/26 增加來函性質為1718變更申請案號
   '2012/9/20 modify by sonia 其他來函也由原承辦人分析 T-169807 不上發文日
   '2013/1/29 modify by sonia 1724通知已轉他所由原承辦人分析 T-180403 不上發文日
   If IsEmptyText(textCF15) = False Or textRvType = "1709" Or textRvType = "1718" Or textRvType = "1706" Or textRvType = "1724" Then
      strCP14 = textCP14
      'modify by sonia 2020/8/27 MCT已閉卷案之1724通知已轉他所直接上發文日
      'strCP48 = DBDATE(textCP48)
      If textRvType = "1724" And m_bolClose = "Y" And Left(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), 3) = "MCT" Then
         strCP27 = 19221111
      Else
         strCP48 = DBDATE(textCP48)
      End If
      'end 2020/8/27
   Else
      strCP14 = strUserNum
      strCP48 = "0"
   End If
   
   '2008/11/24 ADD BY SONIA TM-000042 TM其他來函下一程序繳費時自動上發文日
   If m_TM01 = "TM" And strCP10 = "1706" And textCF15 = "708" Then
      strCP27 = DBDATE(SystemDate())
      strCP14 = strUserNum
      strCP48 = "0"
   End If
   '2008/11/24 END
   
   ' 清除欄位串列
   ClearFieldList
   ' 設定欄位的內容
   SetFieldData "CP01", m_TM01, 0
   SetFieldData "CP02", m_TM02, 0
   SetFieldData "CP03", m_TM03, 0
   SetFieldData "CP04", m_TM04, 0
   SetFieldData "CP05", DBDATE(m_CP05), 1
   If IsEmptyText(strCP06) = False Then: SetFieldData "CP06", strCP06, 1
   If IsEmptyText(strCP07) = False Then: SetFieldData "CP07", strCP07, 1
   If IsEmptyText(textCP08) = False Then: SetFieldData "CP08", textCP08, 0
   SetFieldData "CP09", strCP09, 0
   SetFieldData "CP10", textRvType, 0
    'Modify By Cheng 2004/02/03
    '業務區為最近收文A類接洽記錄單智權人員的業務區
'   SetFieldData "CP12", m_CP12, 0
   SetFieldData "CP12", GetSalesArea(IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04))), 0
    'End
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
'    SetFieldData "CP13", m_CP13, 0
    SetFieldData "CP13", IIf(m_TM01 = "FCT", PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)), 0
    'Modify By Cheng 2002/11/27
    '承辦人為原程序承辦人
'   SetFieldData "CP14", strCP14, 0
   '2008/11/24 MODIFY BY SONIA 恢復有上發文日者CP14掛操作者
   'SetFieldData "CP14", Me.textCP14.Text, 0
   SetFieldData "CP14", strCP14, 0
   '2008/11/24 END
   SetFieldData "CP20", "N", 0
   If IsEmptyText(textCP25) = False Then: SetFieldData "CP25", DBDATE(textCP25), 1
   SetFieldData "CP26", textCP26, 0
    'Modify By Cheng 2002/11/27
    '不上發文日
    '2008/11/24 MODIFY BY SONIA 恢復
'   If IsEmptyText(strCP27) = False Then: SetFieldData "CP27", strCP27, 1
   If IsEmptyText(strCP27) = False Then: SetFieldData "CP27", strCP27, 1
   '2008/11/24 END
   SetFieldData "CP32", "N", 0
   SetFieldData "CP43", m_CP09, 0
    'Add By Cheng 2004/03/16
    strCP64 = Trim(textCP64)
    If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
       strCP64 = strCP64 & ",收文文號：" & Trim(TextCP64_1)
    ElseIf Trim(TextCP64_1) <> "" Then
       strCP64 = "收文文號：" & Trim(TextCP64_1)
    End If
    'End
   'modify by sonia 2021/4/21 此處沒改,根本沒存
   'SetFieldData "CP64", textCP64, 0
   SetFieldData "CP64", strCP64, 0
   'end 2021/4/21
   'Add by Morgan 2003/11/24
   SetFieldData "CP36", textCP36, 0
   Select Case m_TM01
      Case "T", "FCT", "CFT", "TF"
          ' 對造案件名稱
          SetFieldData "CP37", textCP37_1, 0
          
      Case Else
          ' 對造案件名稱(中)
          SetFieldData "CP37", textCP37, 0
          ' 對造案件名稱(英)
          SetFieldData "CP38", textCP38, 0
          ' 對造案件名稱(日)
          SetFieldData "CP39", textCP39, 0
   End Select
   SetFieldData "CP40", textCP40, 0
   SetFieldData "CP41", textCP41, 0
   SetFieldData "CP42", textCP42, 0
   '2018/4/19 add by sonia T-199865行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   If textRvType = "1709" And m_CP10 = "403" Then
     SetFieldData "CP35", textTM12_new, 0
   End If
   '2018/4/19 end
   '---End
   
   ' 設定SQL語法
   bFirst = True
   strSql = "INSERT INTO CaseProgress ("
   For nIndex = 0 To m_FieldCount - 1
      strTmp = m_FieldList(nIndex).fiName
      If strTmp <> Empty Then
         If bFirst = True Then
            strSql = strSql & strTmp
            bFirst = False
         Else
            strSql = strSql & "," & strTmp
         End If
      End If
   Next nIndex
   strSql = strSql & ") "
   strSql = strSql & "VALUES ("
   
   bFirst = True
   For nIndex = 0 To m_FieldCount - 1
      strTmp = Empty
      If m_FieldList(nIndex).fiType = 0 Then
         ' 91.03.25 modify by louis (單引號)
         strTmp = "'" & ChgSQL(m_FieldList(nIndex).fiNewData) & "'"
      Else
         strTmp = m_FieldList(nIndex).fiNewData
      End If
      If bFirst = True Then
         strSql = strSql & strTmp
         bFirst = False
      Else
         strSql = strSql & "," & strTmp
      End If
   Next nIndex
   strSql = strSql & ")"
   ' 存取資料庫
   cnnConnection.Execute strSql
   
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      '1702.通知修正 1706.其他來函:沒有公文
      If Val(textCP06) > 0 Then '有期限者,為掛號
         PUB_AddLetterProgress strLD18, IIf(strCP10 = "1702" Or strCP10 = "1706", 0, 1), IIf(textPrint = "N", False, True), "", True, m_TM23, strCP10, m_TM44
      Else
         PUB_AddLetterProgress strLD18, IIf(strCP10 = "1702" Or strCP10 = "1706", 0, 1), IIf(textPrint = "N", False, True), "", False, m_TM23, strCP10, m_TM44
      End If
   End If
   '2019/12/19 END
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   If Trim(Text11) <> "" Then
     strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
              "WHERE CP09='" & strCP09 & "' "
     cnnConnection.Execute strSql
   End If
   
   'add by nickc 2008/01/09 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
   If Val(strCP48) <> "0" Then
        strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(strCP48) & " " & _
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
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   ' 清除欄位串列
   ClearFieldList
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新商標基本檔 (是否閉卷欄位)
   'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'   strSql = "UPDATE TradeMark SET TM29='" & textTM29 & "' " & _
'            "WHERE TM01 = '" & m_TM01 & "' AND " & _
'                  "TM02 = '" & m_TM02 & "' AND " & _
'                  "TM03 = '" & m_TM03 & "' AND " & _
'                  "TM04 = '" & m_TM04 & "'"
'   cnnConnection.Execute strSql
   If textTM29 = "" Then
      strSql = "UPDATE TradeMark SET TM29=null,TM30=null,TM31=null " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   Else
      '原基本檔非閉卷時才更新
      If textTM29 = "Y" And m_bolClose <> "Y" Then
         strSql = "UPDATE TradeMark SET TM29='" & textTM29 & "',TM30=" & strSrvDate(1) & ",TM31='99' " & _
                  "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql
      End If
   End If
   
   'add by nickc 2006/11/21
   If textPrint <> "N" Then
        strSql = "UPDATE TradeMark SET TM77 = '" & textPrint & "' " & _
                 "WHERE TM01 = '" & m_TM01 & "' AND " & _
                       "TM02 = '" & m_TM02 & "' AND " & _
                       "TM03 = '" & m_TM03 & "' AND " & _
                       "TM04 = '" & m_TM04 & "'"
        cnnConnection.Execute strSql
        strSql = "UPDATE servicepractice SET SP72 = '" & textPrint & "' " & _
                 "WHERE SP01 = '" & m_TM01 & "' AND " & _
                       "SP02 = '" & m_TM02 & "' AND " & _
                       "SP03 = '" & m_TM03 & "' AND " & _
                       "SP04 = '" & m_TM04 & "'"
        cnnConnection.Execute strSql
   End If
   
   ' 更新服務業務基本檔 (是否閉卷及BTTM欄位) modify by sonia 91.10.10
   'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'   strSql = "UPDATE servicepractice SET SP15 = '" & textTM29 & "', SP50 = '" & textBTTM & "' " & _
'            "WHERE SP01 = '" & m_TM01 & "' AND " & _
'                  "SP02 = '" & m_TM02 & "' AND " & _
'                  "SP03 = '" & m_TM03 & "' AND " & _
'                  "SP04 = '" & m_TM04 & "'"
'   cnnConnection.Execute strSql
   strSql = "UPDATE servicepractice SET SP50='" & textBTTM & "' " & _
            "WHERE SP01 = '" & m_TM01 & "' AND " & _
                  "SP02 = '" & m_TM02 & "' AND " & _
                  "SP03 = '" & m_TM03 & "' AND " & _
                  "SP04 = '" & m_TM04 & "'"
   cnnConnection.Execute strSql
   If textTM29 = "" Then
      strSql = "UPDATE servicepractice SET SP15=null,SP16=null,SP17=null " & _
               "WHERE SP01 = '" & m_TM01 & "' AND " & _
                     "SP02 = '" & m_TM02 & "' AND " & _
                     "SP03 = '" & m_TM03 & "' AND " & _
                     "SP04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   Else
      '原基本檔非閉卷時才更新
      If textTM29 = "Y" And m_bolClose <> "Y" Then
         strSql = "UPDATE servicepractice SET SP15='" & textTM29 & "',SP16=" & strSrvDate(1) & ",SP17='99' " & _
                  "WHERE SP01 = '" & m_TM01 & "' AND " & _
                        "SP02 = '" & m_TM02 & "' AND " & _
                        "SP03 = '" & m_TM03 & "' AND " & _
                        "SP04 = '" & m_TM04 & "'"
         cnnConnection.Execute strSql
      End If
   End If
   
   'Add By Cheng 2002/01/03
   '若來函性質屬於爭議程序(16XX), 應更新商標基本檔是否有爭議程序欄(TM19)為"Y"
   If Left(strCP10, 2) = "16" Then
      strSql = "UPDATE TradeMark SET TM19='Y'" & _
               " WHERE TM01 = '" & m_TM01 & "'" & _
               " And TM02 = '" & m_TM02 & "'" & _
               " And TM03 = '" & m_TM03 & "'" & _
               " And TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 當來函性質為專用權消滅時, 更新商標基本檔的專用權是否存在欄位為N
   If Trim(textRvType) = "1704" Then
      strSql = "UPDATE TradeMark SET TM17 = '" & "N" & "' " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   
   'Add By Sindy 2009/05/11
   ' 當來函性質為變更申請案號時, 更新商標基本檔的申請案號為新申請案號
   ' 將原TM12存入此筆的C類來函之CP30
   If Trim(textRvType) = "1718" Then
      strSql = "UPDATE CaseProgress " & _
                              "SET CP30='" & Trim(textTM12.Text) & "' " & _
                       "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
      
      'modify by sonia 2020/11/16 同時清除審定來函日,公告日,審定號,目前准駁及專用權並加註備註,FCT040933
      strSql = "UPDATE TradeMark SET TM12='" & Trim(textTM12_new.Text) & "', " & _
               "TM13=NULL,TM14=NULL,TM15=NULL,TM16=NULL,TM17=NULL,TM20=NULL,TM21=NULL,TM22=NULL,TM58='" & ChangeTStringToTDateString(textCP05S) & "變更申請案號同時清除審定資料;'||TM58 " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   '2009/05/11 End
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   'add by nickc 2006/09/08 案件性質 1203、1204、1405 若收文未發文，不產生下一程序，改更新案件進度，但是因為上面已有程式做更新 cp 的動作，所以這裡不做更新 cp
   '2009/10/1 MODIFY BY SONIA 來函性質 1203、1204、1405 若收文未發文，不產生下一程序，改更新案件進度,原NICK未寫,因為上面沒有更新 CP 的動作
   'If Trim(textRvType) <> "1203" And Trim(textRvType) <> "1204" And Trim(textRvType) <> "1405" Or (m_CP27 <> "" And (Trim(textRvType) = "1203" Or Trim(textRvType) = "1204" Or Trim(textRvType) = "1405")) Then
   If m_CP27 = "" And (Trim(textRvType) = "1203" Or Trim(textRvType) = "1204" Or Trim(textRvType) = "1405") Then
      '2009/10/21 MODIFY BY SONIA 同時更新進度備註及機關文號T-121604
      'strSQL = "Update CaseProgress Set CP06=" & DBDATE(strCP06) & ", CP07=" & DBDATE(strCP07) & " Where CP09='" & m_CP09 & "'"
      'Modified by Lydia 2024/07/02 +ChgSql
      strSql = "Update CaseProgress Set CP06=" & DBDATE(strCP06) & ", CP07=" & DBDATE(strCP07) & ",CP08='" & textCP08 & "',CP64=DECODE(CP64,'','" & ChgSQL(textCP64) & "',CP64||';" & ChgSQL(textCP64) & "') Where CP09='" & m_CP09 & "'"
      cnnConnection.Execute strSql
   Else
   '2009/10/1
       ' 若有輸入下一程序時, 新增資料到下一程序檔
       strNP22 = GetNextProgressNo()
       If IsEmptyText(textCF15) = False Then
         '2008/11/24 ADD BY SONIA TM-000042 TM其他來函下一程序繳費時不產生下一程序,但新增B類進度檔
         If m_TM01 = "TM" And strCP10 = "1706" And textCF15 = "708" Then
            strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48,CP64) " & _
                     "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & DBDATE(strCP06) & "," & DBDATE(strCP07) & "," & _
                             "'" & AutoNo("B", 6) & "','708','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                             "'" & "N" & "','" & "N" & "','" & "N" & "','" & strCP09 & "'," & DBDATE(textCP48) & ",'" & ChgSQL(textCP64) & "')"
            cnnConnection.Execute strSql
         Else
         '2008/11/24 END
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
                strMailMsg = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "收到" & "" & GetCaseTypeName(m_TM01, textCF15, IIf(m_TM10 = "000", 0, 1)) & "前已收文,請辦理後續！"
                'Modify By Sindy 2022/1/19 改到commit後,再發信
                'PUB_SendMail strUserNum, m_CP14, "", strMailMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
              
             '進度檔未有相同未發文未取消收文之案件性質或上述不更新期限,才新增下一程序
             Else
                strNP14 = Empty
                strNP14 = GetRelatedPerson(m_CP09)
                'Modify By Cheng 2002/09/25
        '      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
        '                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & textCF15 & "'," & _
        '                          strCP06 & "," & strCP07 & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
                'Modify By Cheng 2003/04/03
                '智權人員存最近收文A類接洽記錄單的智權人員
                'Modified by Lydia 2024/07/02 +ChgSql
                strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                        "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & textCF15 & "'," & _
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
                'add by nickc 2008/05/09 補回覆單，之前沒有列到，龔說要印
                '2008/5/20 MODIFY BY SONIA FCT案不印回覆單
                If m_TM01 <> "FCT" Then
                    'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                    If bolUpdCP = False Then
                        Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strCP07), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                    End If
                End If
            End Select
         End If    '2008/11/24 add by sonia
       End If
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
   
   'add by sonia 2013/11/28 通知已轉他所, 下一程序除延展期限外其他期限都結案
   If textRvType = "1724" Then
      strSql = "UPDATE NextProgress SET NP06 = 'N' " & _
               "WHERE NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 <>'102' AND NP06 IS NULL "
      cnnConnection.Execute strSql
   End If
   '2013/11/28 end
   
   'Add By Sindy 2009/09/24
   '因為有些來函由內商輸入，內商有自行控管之承辦期限及發文日。改為內商輸入所有C類來函，
   '若業務區為F字頭者，除爭議受理外，自動產生B類收文，案件性質為外商發文722，不上發文日，不向客戶請款
   Dim strCP09B As String
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
   
   'Added by Morgan 2017/4/24 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/4/24
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010408_1"
   End If
   '2019/5/10 END
   
'Add By Cheng 2002/11/07
cnnConnection.CommitTrans
   
   'Modify By Sindy 2022/1/19 改到commit後,再發信
   '更新進度檔,並發Mail通知承辦人
   If bolUpdCP = True Then
      PUB_SendMail strUserNum, m_CP14, "", strMailMsg, "本所期限：" & textCP06 & "　　法定期限：" & textCP07
   End If
   '2022/1/19 END
   
   ' 列印定稿
   If Me.textPrint.Text <> "N" Then
      'Add By Sindy 2012/1/12
      ET01 = "12"
      ET02 = strCP09
      bolEdit = IIf(Me.textEditPrint.Text = "Y", True, False)
      '2012/1/12 End
      
      'add by nickc 2006/11/21
      If m_TM10 < "010" Then
        If textPrint = "1" Then
            'Add By Sindy 2009/10/26
            ' 先呼叫定稿程式的清除原定稿資料的函式去清除之前殘留在例外欄位檔中的資料
            InsExpField
            If textRvType = "1718" Then
'               NowPrint strCP09, "12", "02", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
               ET03 = "02" 'Modify By Sindy 2012/1/12
            '2009/10/26 End
            Else
               '2010/4/16 CANCEL BY SONIA 找不到定稿
               'NowPrint strCP09, "12", "01", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
            End If
        End If
      '2010/2/9 add by sonia
      ElseIf m_TM01 = "TF" Then
'         NowPrint strCP09, "12", "01", IIf(Me.textEditPrint.Text = "Y", True, False), strUserNum, 0
         ET03 = "01" 'Modify By Sindy 2012/1/12
      '2010/2/9 end
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
            'Add By Sindy 2019/12/19 + strLD18.信函總收文號
            NowPrint ET02, ET01, ET03, bolEdit, strUserNum, 0, , , , , , , , , , , , strLD18
         End If
      End If
      '2012/1/12 End
   End If
''Add By Cheng 2002/11/07
'cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnSaveData = False
End Function

'Add By Sindy 2009/10/26
' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   If textRvType = "1718" Then
      EndLetter "12", strCP09, "02", strUserNum
      'Modified by Lydia 2024/07/02 +ChgSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "12" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
               "'" & "收文文號" & "','" & ChgSQL(Trim(TextCP64_1.Text)) & "')"
      cnnConnection.Execute strSql
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
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   
   textCF15_2 = Empty
   If IsEmptyText(textCF15) = False Then
      Select Case textCF15.Text
         Case "補正":
            textCF15 = "201"
         Case "申請意見書":
            textCF15 = "202"
         Case "異議答辯", "爭議答辯", "侵權處理":
            textCF15 = "602"
         Case "評定答辯":
            textCF15 = "604"
         Case "廢止答辯":
            textCF15 = "606"
         Case "補充答辯":
            textCF15 = "613"
         Case "補充理由":
            textCF15 = "612"
         Case "變更":
            textCF15 = "301"
         Case "領證":
            textCF15 = "701"
      End Select
         
      'Add By Cheng 2002/01/10
      If Len(Me.textCF15.Text) <> 3 Then
         Cancel = True
         MsgBox "下一程序欄位值必須為三碼!!!", vbExclamation
         textCF15_GotFocus
         Exit Sub
      End If
   
      ' 91.01.22 modify by louis (案件性質名稱)
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
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2010/11/26
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
      End If
   End If
End Sub

'專用權消滅日
Private Sub textCP25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP25) = False Then
      If CheckIsTaiwanDate(textCP25) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的專用權消滅日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
      End If
   End If
End Sub

Private Sub textCP26_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textCP37_1_GotFocus()
'edit by nickc 2007/06/06 切換輸入法改用API
OpenIme
End Sub

Private Sub textCP37_LostFocus()
'edit by nickc 2007/06/06 切換輸入法改用API
CloseIme
End Sub

'Add by Amy 2025/01/17
Private Sub textCP37_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件中文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP37, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_GotFocus
   End If
End Sub

Private Sub textCP37_1_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件中文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
End Sub

Private Sub textCP38_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件英文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP38, 250, True, strMsg) = False Then
      Cancel = True
      textCP38_GotFocus
   End If
End Sub

Private Sub textCP39_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造案件日文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP39, 160, True, strMsg) = False Then
      Cancel = True
      textCP39_GotFocus
   End If
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造中文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      textCP40_GotFocus
   End If
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造英文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      textCP41_GotFocus
   End If
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造日文名稱"
   Cancel = False
   If CheckLengthIsOK(textCP42, 600) = False Then
      Cancel = True
      textCP42_GotFocus
   End If
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strDay As String
   Dim strTit As String
   Dim strMsg As String
   Dim strCF15 As String
   Dim nResponse
   
   Cancel = False
   
    'Modify By Cheng 2002/11/18
'   ' 案件性質依還函性質欄位來區分為"被禁止處分"或"被取消禁止處分"
'   strCF15 = "1614"
'   If textRvType = "2" Then
'      strCF15 = "1615"
'   End If
    strCF15 = "" & Me.textRvType.Text

   ' 承辦期限的日期應為來函收文日加上工作天數
   ' 工作天數由系統別+國家代碼+案件性質(下一程序代號)搜尋案件收費表的工作天數
   ' 若有值才做檢查
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為西元日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
         GoTo EXITSUB
      End If
   Else
      ' 承辦期限不可超過本所期限
''''edit by nickc 2007/10/12 改抓有時效的
''''      strDay = GetWorkDays(m_TM01, m_TM10, strCF15)
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis
''''         'strDate = DBDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
      strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCF15, DBDATE(m_CP05), DBDATE(textCP06), textCP09)
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP48) <> TAIWANDATE(strDate) Then
            Cancel = True
            strTit = "資料檢核"
            strMsg = "承辦期限日期應為<" & TAIWANDATE(strDate) & ">"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP48_GotFocus
         End If
      End If
   End If
EXITSUB:
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
   '2008/11/11 ADD BY SONIA
   If (Me.textRvType.Text = "1203" Or Me.textRvType.Text = "1204") And textCP64 = "上下午時分,第法庭" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請於進度備註欄輸入開庭時間及法庭"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   Else
   '2008/11/11 END
      If CheckLengthIsOK(textCP64, 2000) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "進度備註欄位內容太長"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP64_GotFocus
      End If
   End If
End Sub

Private Sub textEditPrint_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 89 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
End Sub

Private Sub textPrint_GotFocus()
TextInverse Me.textPrint
End Sub

Private Sub textPrint_KeyPress(KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
If KeyAscii <> 78 And KeyAscii <> 8 Then
   KeyAscii = 0
End If
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
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

Private Sub textRvType_LostFocus()
   'Add By Cheng 2003/01/14
   '若來函性質為專用權消滅(1704), 預設定稿要開Word修改
   If Me.textRvType.Text = "1704" Then
      Me.textEditPrint.Text = "Y"
      '92.6.8 ADD BY SONIA
      Me.textTM29.Text = "Y"
      '92.6.8 END
   End If
   '2008/11/11 ADD BY SONIA
   If (Me.textRvType.Text = "1203" Or Me.textRvType.Text = "1204") And textCP64 = "" Then
      textCP64 = "上下午時分,第法庭"
   ElseIf Me.textRvType.Text <> "1203" And Me.textRvType.Text <> "1204" And textCP64 = "上下午時分,第法庭" Then
      textCP64 = ""
   End If
   '2018/4/19 add by sonia T-199865行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   If textRvType = "1709" And m_CP10 = "403" Then
      textTM12_new = Val(Left(DBDATE(m_CP27), 4) - 1911) & "年度行商訴字第號"
   End If
   '2018/4/19 end
   '2008/11/11 END
   If m_DeadLine = "" Then 'Added by Morgan 2017/4/24 電子公文
      If IsEmptyText(textRvType) = False Then Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   End If
End Sub

' 來函性質
Private Sub textRvType_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'Add by Amy 2022/09/26 畫面來函性質輸「其他來函1706」則對造頁籤改為「關係案」,對造字樣改「對方」
   SSTab1.TabCaption(1) = "對造名稱"
   strExc(1) = "對造"
   If textRvType = "1706" Then
        SSTab1.TabCaption(1) = "關係案"
        strExc(1) = "對方"
        Label10.Caption = strExc(1) & Mid(Label10.Caption, 3)
        Label30.Caption = strExc(1) & Mid(Label30.Caption, 3)
        Label13.Caption = strExc(1) & Mid(Label13.Caption, 3)
        Label17.Caption = strExc(1) & Mid(Label17.Caption, 3)
        Label18.Caption = strExc(1) & Mid(Label18.Caption, 3)
        Label19.Caption = strExc(1) & Mid(Label19.Caption, 3)
        Label20.Caption = strExc(1) & Mid(Label20.Caption, 3)
        Label21.Caption = strExc(1) & Mid(Label21.Caption, 3)
   End If
   'end 2022/09/26
   textRvType_2 = Empty
   Cancel = False
   
   '若有輸入來函性質
   If IsEmptyText(textRvType) = False Then
      'Add By Cheng 2002/01/10
      If Len(Me.textRvType.Text) <> 4 Then
         Cancel = True
         MsgBox "來函性質欄位必須為四碼!!!", vbExclamation
         textRvType_GotFocus
         Exit Sub
      End If
      
      ' 取得案件性質名稱
      'Modify By Sindy 2009/10/26
      If m_TM10 = "000" Then
         textRvType_2 = GetCaseTypeName(m_TM01, textRvType, 0)
      Else
         textRvType_2 = GetCaseTypeName(m_TM01, textRvType, 1)
      End If
      '2009/10/26 End
      If IsEmptyText(textRvType_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRvType_GotFocus
      End If
      
      '2009/10/22 ADD BY SONIA
      '2011/6/9 MODIFY BY SONIA 加1602
      'modify by sonia 2019/10/9 +1716,1717,1799,1729,1728,1725,及FCT之1721,1722
      If InStr("1001,1002,1003,1004,1005,1006,1102,1403,1602,1716,1717,1799,1729,1728,1725,1721,1722", textRvType) > 0 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "此來函性質不可由此畫面輸入資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRvType_GotFocus
      End If
      '2009/10/22 END
   
   'Add By Cheng 2002/01/10
   '若沒輸入來函性質
   Else
      Cancel = True
      MsgBox "來函性質欄位必須輸入且為四碼!!!", vbExclamation
      textRvType_GotFocus
      Exit Sub
   End If
   
EXITSUB:
   ' 來函性質為專用權消滅時才可輸入專用權消滅日欄位
   If textRvType = "1704" Then
      EnableTextBox textCP25, True
   Else
      textCP25 = Empty
      EnableTextBox textCP25, False
   End If
   'Add By Sindy 2009/05/11
   ' 來函性質為變更申請案號時才可輸入新申請案號欄位
   If textRvType = "1718" Then
      Label29 = "新申請案號 :"
      EnableTextBox textTM12_new, True
      textTM12_new.MaxLength = 20
   '2018/4/19 add by sonia T-199865行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   ElseIf textRvType = "1709" And m_CP10 = "403" Then
      Label29 = "法院案號:"
      EnableTextBox textTM12_new, True
      textTM12_new.MaxLength = 32
   '2018/4/19 end
   Else
      textTM12_new = Empty
      EnableTextBox textTM12_new, False
   End If
   '2009/05/11 End
   '2009/6/9 add by sonia
   'Modify By Sindy 2009/10/26 增加1718變更申請案號
   If textRvType = "1709" Or textRvType = "1718" Then
      textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textRvType, DBDATE(m_CP05), DBDATE(textCP06), textCP09))
   End If
   '2009/6/9 end
   'Add By Sindy 2009/10/26 預設為操作人員
   If textRvType = "1718" Then
      textCP14 = strUserNum
      textCP14_2 = GetStaffName(strUserNum)
   End If
   '2009/10/26 End
   '2009/10/14 ADD BY SONIA
   'Modify By Sindy 2010/5/28
   'If textRvType = "1719" Then
   If textRvType <> "" And textCF15 = "" Then
      textCF15 = GetNextProgress(m_TM01, m_TM10, textRvType)   '取得下一程序
      If IsEmptyText(textCF15) = False Then
         If m_TM10 = "000" Then
            textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
         Else
            textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
         End If
      End If
   End If
   '2009/10/14 END
End Sub

Private Sub textTM12_new_GotFocus()
Dim intPos As Integer   'add by sonia 2018/4/19
   'modify by sonia 2018/4/19
   'InverseTextBox textTM12_new
   If textRvType = "1709" Then
      '將游標停在"號"的前面
      With Me.textTM12_new
         If Len("" & .Text) > 0 Then
            intPos = InStr("" & .Text, "號")
            If intPos - 1 >= 0 Then
               .SelStart = intPos - 1
               .SelLength = 0
            End If
         End If
      End With
   Else
      InverseTextBox textTM12_new
   End If
   '2018/4/19 end
End Sub

'Add By Sindy 2010/9/1
Private Sub textTM12_new_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   'modify by sonia 2018/4/19
   'If IsEmptyText(textTM12_new) = False Then
   If IsEmptyText(textTM12_new) = False And textRvType = "1718" Then
      '檢查申請案號所輸入的長度是否正確
      'Add By Sindy 2017/5/17 + strRetrunText
      If PUB_ChkTm12Tm15Length("1", textTM12_new, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM12_new_GotFocus
         Exit Sub
      'Add By Sindy 2017/5/17
      Else
         textTM12_new = strRetrunText
      '2017/5/17 END
      End If
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
      '2009/2/11 ADD BY SONIA 預設承辦期限
      If textCP48 = "" Then
         textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textRvType, DBDATE(m_CP05), DBDATE(textCP06)))
      End If
      '2009/2/11 END
        'Modify By Cheng 2002/11/18
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
        'Modify By Cheng 2002/11/18
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

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   'Add by Amy 2021/12/29檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

   ' 來函性質不可為空白
   If IsEmptyText(textRvType) = True Then
      strTit = "檢核資料"
      strMsg = "來函性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textRvType.SetFocus
      GoTo EXITSUB
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
   
   'Add By Cheng 2002/03/11
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
   ' 申請國家為台灣時, 機關文號不可為空白
   If m_TM10 < "010" Then
      If IsEmptyText(textCP08) = True Then
         'Modify By Cheng 2002/06/12
         '若來函性質為"=準備程序"或"言詞辨論"
         If Me.textRvType.Text <> "1203" And Me.textRvType.Text <> "1204" Then
            strTit = "檢核資料"
'            strMsg = "申請國家為台灣時, 機關文號不可為空白"
            strMsg = "機關文號不可為空白"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP08.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   ' 若來函性質為專用權消滅時一定要輸入專用權消滅日
   If textRvType = "1704" Then
      If IsEmptyText(textCP25) = True Then
         strTit = "檢核資料"
         strMsg = "來函性質為專用權消滅, 一定要輸入專用權消滅日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2009/05/11
   ' 若來函性質為變更申請案號時一定要輸入新申請案號
   If textRvType = "1718" Then
      If IsEmptyText(textTM12_new) = True Then
         strTit = "檢核資料"
         strMsg = "來函性質為變更申請案號, 一定要輸入新申請案號"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM12_new.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2009/05/11 End
   
   'Add By Sindy 2009/10/26
   If textRvType = "1718" Then
      If IsEmptyText(TextCP64_1) = True Then
         strTit = "檢核資料"
         strMsg = "來函性質為變更申請案號, 收文文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextCP64_1.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2009/10/26 End
   
   'Add by Morgan 2003/11/25
   '來函性質為 1404 或 1405 時, 對造中英日文名稱不可均為空白
   'modify by sonia 2021/12/2 +1619被部分廢止,1620被部分廢止(理由),1623被部分異議,1624被部分異議(理由),1625被部分評定,1626被部分評定(理由)
   'If ((textRvType.Text = "1404" Or textRvType.Text = "1405" Or textRvType.Text = "1619" Or textRvType.Text = "1620" Or textRvType.Text = "1623" Or textRvType.Text = "1624" Or textRvType.Text = "1625" Or textRvType.Text = "1626") And Trim(textCP40.Text) = "" And Trim(textCP41.Text) = "" And Trim(textCP42.Text) = "") Then
   '   MsgBox "來函性質為'通知1404、1405、1619、1620、1623、1624、1625、1626時，對造中、英、日名稱不可同時空白！", vbCritical
   '   SSTab1.Tab = 1
   '   textCP40.SetFocus
   '   GoTo EXITSUB
   'End If
   If Trim(textCP40.Text) = "" And Trim(textCP41.Text) = "" And Trim(textCP42.Text) = "" Then
      If (textRvType.Text = "1404" Or textRvType.Text = "1405" Or textRvType.Text = "1619" Or textRvType.Text = "1620" Or textRvType.Text = "1623" Or textRvType.Text = "1624" Or textRvType.Text = "1625" Or textRvType.Text = "1626") Then
         MsgBox "來函性質為'通知1404、1405、1619、1620、1623、1624、1625、1626時，對造中、英、日名稱不可同時空白！", vbCritical
         SSTab1.Tab = 1
         textCP40.SetFocus
         GoTo EXITSUB
      End If
   Else
      PUB_ChkCustNameExist textCP40, textCP41, textCP42
   End If
   'end 2021/12/2
   '---End
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textRvType_GotFocus()
   'Added by Morgan 2017/4/24 電子公文
   If textRvType = "" And m_NewCP10 <> "" Then
      textRvType = m_NewCP10
   End If
   'end 2014/4/17
   InverseTextBox textRvType
End Sub

Private Sub textBTTM_GotFocus()
   InverseTextBox textBTTM
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
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

Private Sub textCP25_GotFocus()
   InverseTextBox textCP25
End Sub

Private Sub textCP26_GotFocus()
   InverseTextBox textCP26
End Sub

Private Sub textCP48_GotFocus()
   InverseTextBox textCP48
End Sub

Private Sub textCP64_GotFocus()
   InverseTextBox textCP64
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
'Add By Cheng 2002/11/18
Dim strTit As String
Dim strMsg As String
Dim nResponse

TxtValidate = False

'Add By Sindy 2010/12/24
If Me.textTM12_new.Enabled = True Then
   Cancel = False
   textTM12_new_Validate Cancel
   If Cancel = True Then
      textTM12_new.SetFocus
      Exit Function
   End If
End If

'Add by Amy 2025/01/17 輸完直接按Enter鍵不會檢查
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
'end 2025/01/17

If Me.textRvType.Enabled = True Then
   Cancel = False
   textRvType_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

   '2009/10/14 自CheckDataValid移過來 BY SONIA
   ' 有輸入下一程序時, 本所期限與法定期限不可為空白
   If IsEmptyText(textCF15) = False Then
      If IsEmptyText(textCP06) = True Or IsEmptyText(textCP07) = True Then
         strTit = "檢核資料"
         strMsg = "有下一程序時, 本所期限與法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         Cancel = True
         Exit Function
      End If
      ' 本所期限必須小於法定期限
      If Val(textCP06) > Val(textCP07) Then
         strTit = "檢核資料"
         strMsg = "本所期限的日期不可超過法定期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         Cancel = True
         Exit Function
      End If
   End If
   
   ' 有本所期限時, 承辦期限不可為空白
   If IsEmptyText(textCP06) = False Then
      If IsEmptyText(textCP48) = True Then
         strTit = "檢核資料"
         strMsg = "本所期限有資料時承辦期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         Cancel = True
         Exit Function
      End If
      
    'Modify By Cheng 2002/11/18
'      If Val(textCP48) >= Val(textCP06) Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "檢核資料"
         strMsg = "承辦期限不可超過本所期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         Cancel = True
         Exit Function
      End If
   End If
   '2009/10/14 END
   
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

If Me.textCP25.Enabled = True Then
   Cancel = False
   textCP25_Validate Cancel
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

If Me.textCP48.Enabled = True Then
   Cancel = False
   textCP48_Validate Cancel
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

If Me.textTM29.Enabled = True Then
   Cancel = False
   textTM29_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
    'Modify By Cheng 2002/11/18
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
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/24 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/24 電子公文
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
    Me.textCP40.IMEMode = 1
End Sub

Private Sub textCP41_GotFocus()
   InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
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
   strFromDate = DBDATE(frm02010408_1.textCP05)
   
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
   strFromDate = DBDATE(frm02010408_1.textCP05)
   
   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   If textRvType = "" Then Exit Function
   If ClsPDGetCaseProperty(m_TM01, textRvType, strTempName, bolTmp) Then
      textCP06 = ""
      textCP07 = ""
      
      If m_TM10 = 台灣國家代號 Then
         strExc(0) = "SELECT CPM07,CPM08,CPM09 FROM CASEPROPERTYMAP WHERE CPM01='" & m_TM01 & "' AND CPM02='" & textRvType & "'"
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
