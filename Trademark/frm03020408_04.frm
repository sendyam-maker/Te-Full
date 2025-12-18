VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm03020408_04 
   BorderStyle     =   1  '單線固定
   Caption         =   "其它來函輸入"
   ClientHeight    =   6816
   ClientLeft      =   216
   ClientTop       =   936
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6816
   ScaleWidth      =   9144
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   1425
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5670
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   495
      Width           =   2532
   End
   Begin VB.CommandButton cmdRelate 
      Caption         =   "相關卷號(&F)"
      Height          =   400
      Left            =   4620
      TabIndex        =   0
      Top             =   30
      Width           =   1212
   End
   Begin VB.TextBox textCP09 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1425
      Width           =   2532
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   480
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1740
      Width           =   2532
   End
   Begin VB.TextBox textCP05S 
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2055
      Width           =   2412
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8100
      TabIndex        =   3
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   5880
      TabIndex        =   1
      Top             =   30
      Width           =   912
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   6840
      TabIndex        =   2
      Top             =   30
      Width           =   1212
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4450
      Left            =   30
      TabIndex        =   25
      Top             =   2370
      Width           =   9100
      _ExtentX        =   16066
      _ExtentY        =   7853
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm03020408_04.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label29"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "textCP14_2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "textCP64"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label32"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label4"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label7"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label16"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label15"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label24"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label26"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label2"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label9"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label21"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "grdList"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "TextCP64_1"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "textTM12_new"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "textCP25"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "textCP48"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP08"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP07"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Frame2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Frame1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP06"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCF15_2"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textRvType"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textTM29"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textCP14"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP26"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textRvType_2"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textCF15"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "Text10"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text11"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Text12"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).ControlCount=   39
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm03020408_04.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP41"
      Tab(1).Control(1)=   "textCP36"
      Tab(1).Control(2)=   "Label13"
      Tab(1).Control(3)=   "Label30"
      Tab(1).Control(4)=   "Label11"
      Tab(1).Control(5)=   "Label20"
      Tab(1).Control(6)=   "Label19"
      Tab(1).Control(7)=   "Label10"
      Tab(1).Control(8)=   "textCP42"
      Tab(1).Control(9)=   "textCP40"
      Tab(1).Control(10)=   "textCP37_1"
      Tab(1).ControlCount=   11
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73170
         TabIndex        =   70
         Top             =   2040
         Width           =   6795
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73170
         MaxLength       =   200
         TabIndex        =   69
         Top             =   390
         Width           =   6795
      End
      Begin VB.TextBox Text12 
         Height          =   285
         Left            =   6600
         MaxLength       =   7
         TabIndex        =   68
         Top             =   1140
         Width           =   975
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   5640
         MaxLength       =   2
         TabIndex        =   67
         Top             =   1140
         Width           =   375
      End
      Begin VB.TextBox Text10 
         Height          =   285
         Left            =   4620
         MaxLength       =   2
         TabIndex        =   66
         Top             =   1140
         Width           =   375
      End
      Begin VB.ComboBox textCF15 
         Height          =   276
         Left            =   1170
         TabIndex        =   46
         Top             =   675
         Width           =   1332
      End
      Begin VB.TextBox textRvType_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   1950
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   360
         Width           =   1692
      End
      Begin VB.TextBox textCP26 
         Height          =   285
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   44
         Top             =   1815
         Width           =   372
      End
      Begin VB.TextBox textCP14 
         Height          =   285
         Left            =   1170
         MaxLength       =   6
         TabIndex        =   43
         Top             =   2130
         Width           =   732
      End
      Begin VB.TextBox textTM29 
         Height          =   285
         Left            =   1170
         MaxLength       =   1
         TabIndex        =   42
         Top             =   1815
         Width           =   732
      End
      Begin VB.TextBox textRvType 
         Height          =   285
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   41
         Top             =   360
         Width           =   732
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   2550
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   690
         Width           =   1092
      End
      Begin VB.TextBox textCP06 
         Height          =   285
         Left            =   1170
         MaxLength       =   7
         TabIndex        =   39
         Top             =   1500
         Width           =   2532
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1170
         TabIndex        =   36
         Top             =   1005
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   38
            Top             =   180
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   37
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   3840
         TabIndex        =   32
         Top             =   1005
         Width           =   4215
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   35
            Top             =   180
            Width           =   1575
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   34
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   33
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.TextBox textCP07 
         Height          =   285
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   31
         Top             =   1500
         Width           =   2532
      End
      Begin VB.TextBox textCP08 
         Height          =   285
         Left            =   5640
         MaxLength       =   32
         TabIndex        =   30
         Top             =   360
         Width           =   2532
      End
      Begin VB.TextBox textCP48 
         Height          =   285
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   29
         Top             =   2130
         Width           =   2532
      End
      Begin VB.TextBox textCP25 
         Height          =   285
         Left            =   5640
         MaxLength       =   7
         TabIndex        =   28
         Top             =   690
         Width           =   1875
      End
      Begin VB.TextBox textTM12_new 
         Height          =   285
         Left            =   1170
         MaxLength       =   20
         TabIndex        =   27
         Top             =   2460
         Width           =   1725
      End
      Begin VB.TextBox TextCP64_1 
         Height          =   285
         Left            =   5640
         MaxLength       =   40
         TabIndex        =   26
         Top             =   2460
         Width           =   1725
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
         Height          =   984
         Left            =   1170
         TabIndex        =   80
         Top             =   2832
         Width           =   7680
         _ExtentX        =   13547
         _ExtentY        =   1736
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
      Begin VB.Label Label13 
         Caption         =   "對造案件中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   79
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   78
         Top             =   750
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   77
         Top             =   2310
         Width           =   1575
      End
      Begin VB.Label Label20 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   76
         Top             =   2070
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   75
         Top             =   1755
         Width           =   1575
      End
      Begin VB.Label Label10 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   74
         Top             =   400
         Width           =   975
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73170
         TabIndex        =   73
         Top             =   2310
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73170
         TabIndex        =   72
         Top             =   1710
         Width           =   6795
         VariousPropertyBits=   679493659
         Size            =   "11986;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   825
         Left            =   -73170
         TabIndex        =   71
         Top             =   720
         Width           =   6795
         VariousPropertyBits=   679493659
         ScrollBars      =   2
         Size            =   "11986;1455"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "專用權消滅日 :"
         Height          =   255
         Index           =   8
         Left            =   4350
         TabIndex        =   65
         Top             =   705
         Width           =   1245
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   90
         TabIndex        =   64
         Top             =   3864
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   4380
         TabIndex        =   63
         Top             =   1830
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   6060
         TabIndex        =   62
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   61
         Top             =   2145
         Width           =   855
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   90
         TabIndex        =   60
         Top             =   2145
         Width           =   855
      End
      Begin VB.Label Label15 
         Caption         =   "是否閉卷 :"
         Height          =   255
         Left            =   90
         TabIndex        =   59
         Top             =   1830
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "(Y:閉卷)"
         Height          =   255
         Left            =   1950
         TabIndex        =   58
         Top             =   1830
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   57
         Top             =   375
         Width           =   885
      End
      Begin VB.Label Label1 
         Caption         =   "來函性質 :"
         Height          =   255
         Index           =   2
         Left            =   90
         TabIndex        =   56
         Top             =   375
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "本案期限 :"
         Height          =   255
         Left            =   90
         TabIndex        =   55
         Top             =   2820
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   90
         TabIndex        =   54
         Top             =   705
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   90
         TabIndex        =   53
         Top             =   1515
         Width           =   855
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4710
         TabIndex        =   52
         Top             =   1515
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   90
         TabIndex        =   51
         Top             =   1125
         Width           =   1035
      End
      Begin MSForms.TextBox textCP64 
         Height          =   528
         Left            =   1170
         TabIndex        =   50
         Top             =   3864
         Width           =   7728
         VariousPropertyBits=   -1466941413
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "13631;931"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   285
         Left            =   1950
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2130
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
      Begin VB.Label Label29 
         Caption         =   "新申請案號 :"
         Height          =   255
         Left            =   90
         TabIndex        =   48
         Top             =   2490
         Width           =   1005
      End
      Begin VB.Label Label14 
         Caption         =   "收文文號 :"
         Height          =   255
         Left            =   4650
         TabIndex        =   47
         Top             =   2490
         Width           =   915
      End
   End
   Begin MSForms.TextBox textCP13 
      Height          =   285
      Left            =   5670
      TabIndex        =   24
      Top             =   2055
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
   Begin MSForms.ComboBox cmbTM05 
      Height          =   285
      Left            =   1200
      TabIndex        =   23
      Top             =   795
      Width           =   7485
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13203;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textTM23 
      Height          =   285
      Left            =   1200
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1110
      Width           =   7485
      VariousPropertyBits=   671105055
      MaxLength       =   20
      Size            =   "13203;503"
      BorderColor     =   16777215
      SpecialEffect   =   0
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
      Height          =   180
      Left            =   3780
      TabIndex        =   18
      Top             =   525
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號 :"
      Height          =   255
      Index           =   5
      Left            =   4740
      TabIndex        =   17
      Top             =   495
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4740
      TabIndex        =   16
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文號 :"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   495
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   810
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1125
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4740
      TabIndex        =   11
      Top             =   1755
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   10
      Top             =   1755
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "來函收文日 :"
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   9
      Top             =   2070
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4740
      TabIndex        =   8
      Top             =   2070
      Width           =   885
   End
End
Attribute VB_Name = "frm03020408_04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/20 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Lydia 2021/09/13 改成Form2.0 ; cmbTM05、textTM23、textCP13、textCP14_2、textCP64、grdList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
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
' 業務區
Dim m_CP12 As String
' 智權人員
Dim m_CP13 As String
' 暫時存放 CF15
Dim m_CF15 As String
'
Dim m_CurrSel As Integer
'Add By Sindy 2012/3/5
Dim m_bolClose As String
Dim m_strCloseDT As String
Dim m_strCloseReason As String
'2012/3/5 End
'Added by Morgan 2017/5/9 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/9
Dim m_CP27 As String
Dim m_NewCP09 As String 'Added by Lydia 2022/02/10 新增C類收文號

Private Sub cmdCancel_Click()
   Unload Me
   frm03020408_03.Show
End Sub

Private Sub cmdExit_Click()
   Unload frm03020408_03
   Unload frm03020408_02
   Unload frm03020408_01
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strFilePath As String 'Added by Lydia 2022/02/10 掃瞄檔的路徑
   
   If CheckDataValid() = True Then
      'Added by Lydia 2022/02/10 FCT紙本公文來函，同時將公文函FCT_OA_SCAN匯入卷宗區
      If m_DocNo = "" Then
          If PUB_FCTCheckPDF(m_TM01, m_TM02, m_TM03, m_TM04, textRvType, m_CP09, strFilePath) = False Then
               Exit Sub
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
        'Move by Lydia 2022/02/23 從frm03020408_01.Show上方移過來
        If strFilePath <> "" Then
            If Pub_AutoSavePdf2_FCT(m_TM01, m_TM02, m_TM03, m_TM04, m_NewCP09, textRvType, strFilePath) = False Then
                Exit Sub
            End If
        End If
        'end 2022/02/10
        
      Unload frm03020408_03
      Unload frm03020408_02
      Unload Me
      'Modified by Morgan 2017/5/9 電子公文
      'frm03020408_01.Show
      If m_DocNo <> "" Then
         Unload frm03020408_01
         frm02010412.GoNext
      Else
         frm03020408_01.Show
      End If
      'end 2017/5/9
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
   textCF15.AddItem "異議答辯"
   textCF15.AddItem "評定答辯"
   textCF15.AddItem "廢止答辯"
   textCF15.AddItem "補充答辯"
   textCF15.AddItem "補充理由"
   textCF15.AddItem "變更"
   textCF15.AddItem "領證"
   
   MoveFormToCenter Me
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
   'Add By Cheng 2002/07/18
   Set frm03020408_04 = Nothing
End Sub

Private Sub grdList_Click()
      If grdList.row > 0 Then
         grdList.col = 0
         If grdList.Text = "V" Then
            grdList.Text = Empty
         Else
            grdList.Text = "V"
            'add by sonia 2022/2/23 FCT-046540
            If IsEmptyText(Trim(textCF15)) = True Then
               textCF15 = grdList.TextMatrix(grdList.row, 8)
               textCF15_2 = grdList.TextMatrix(grdList.row, 1)
            End If
            'end 2022/2/23
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
      If IsNull(rsTmp.Fields("TM23")) = False Then
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"))
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      'add by nickc 2006/05/29 加入閉卷提示
      If IsNull(rsTmp.Fields("tm29")) Then
         Me.lblClose.Caption = ""
      Else
         Me.lblClose.Caption = "已閉卷"
      End If
      
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

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 取得案件進度檔A類資料的最後一筆
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
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         textCP09 = rsTmp.Fields("CP09")
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
      ' 業務區
      If IsNull(rsTmp.Fields("CP12")) = False Then
         m_CP12 = rsTmp.Fields("CP12")
      End If
      ' 智權人員
      'Modified by Lydia 2021/08/03 改由PUB_GetFCTSalesNo帶出和產生的C類收文一致
      'If IsNull(rsTmp.Fields("CP13")) = False Then
      '   m_CP13 = rsTmp.Fields("CP13")
      '   textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      'End If
      m_CP13 = Empty
      m_CP13 = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      textCP13 = GetStaffName(m_CP13)
      'end 2021/08/03
      ' 承辦人員
      If IsNull(rsTmp.Fields("CP14")) = False Then
'         textCP14 = rsTmp.Fields("CP14")
'         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
      
      'Add By Sindy 2021/9/27
      If IsNull(rsTmp.Fields("CP27")) = False Then
         m_CP27 = CheckStr("" & rsTmp.Fields("CP27"))
      End If
      
      '預設承辦人
      Me.textCP14.Text = PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)
      Me.textCP14_2.Text = GetStaffName(Me.textCP14.Text)
      
      'Add by Amy 2022/09/07 +對造頁籤
      ' 對造號數
      If IsNull(rsTmp.Fields("CP36")) = False Then
         textCP36 = rsTmp.Fields("CP36")
      End If
      ' 對造案件名稱
      If IsNull(rsTmp.Fields("CP37")) = False Then
          textCP37_1 = rsTmp.Fields("CP37")
      End If
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
      'end 2022/09/07
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   If textCP08 = "" Then
      textCP08 = "（" & strTmp & "）慧商字第號"
   End If
   
   
   'Added by Morgan 2017/5/9 電子公文
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
   'end 2017/5/9
   
End Sub

' 查詢資料庫取得資料
Public Sub QueryData()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bCP40 As Boolean
   Dim strDay As String
   m_TM10 = Empty
   m_CP13 = Empty
      
   ' 本所案號
   textTMKey = m_TM01 & m_TM02 & m_TM03 & m_TM04
   ' 來函收文日
   textCP05S = m_CP05
   
   ' 讀取商標基本檔
   QueryTradeMark
   
   ' 讀取案件進度檔
   QueryCaseProgress
   
   ' 以來函性質來計算承辦期限
   strDay = Empty
   If IsEmptyText(textRvType) = False Then
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = GetWorkDays(m_TM01, m_TM10, textRvType)
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      End If
        textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textRvType, DBDATE(m_CP05), DBDATE(textCP06), textCP09))
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
End Sub

'edit by nick 2004/11/03
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
OnSaveData = True
   Dim nIndex As Integer
   Dim strSql As String
   Dim strCP09 As String
   Dim strCP10 As String
   'Dim strCP12 As String
   Dim strNP07 As String
   Dim strNP14 As String
   Dim strNP22 As String
   Dim strCP64 As String, strCP35 As String
   
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
     
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 新增一筆資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   m_NewCP09 = strCP09  'Added by Lydia 2022/02/10 新增C類收文號
   ' 案件性質為來函性質
   strCP10 = Trim(textRvType)
   
   'Add By Sindy 2021/9/27
   strCP64 = Trim(textCP64)
   If strCP64 <> "" And Trim(TextCP64_1) <> "" Then
      strCP64 = strCP64 & ",收文文號：" & Trim(TextCP64_1)
   ElseIf Trim(TextCP64_1) <> "" Then
      strCP64 = "收文文號：" & Trim(TextCP64_1)
   End If
   '行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   If textRvType = "1709" And m_CP10 = "403" Then
      strCP35 = textTM12_new
   End If
   '2018/4/19 end
   '2021/9/27 END
   
   ' 業務區別 91.8.26 MODIFY BY SONIA
   'strCP12 = GetStaffDepartment(m_CP13)
   ' 組成SQL語法
   ' 91.03.25 modify by louis (單引號)
    'Modify By Cheng 2003/04/07
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2003/09/05
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP43,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & StrCp10 & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
'                    "'" & m_CP09 & "','" & ChgSQL(textCP64) & "') "
   'modify by sonia 2022/2/23 加入CP06,CP07(FCT-046540)
   'Modify by Amy 2022/09/07 +cp36/37/40~42
   strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10,CP12,CP13,CP20,CP26,CP32,CP43,CP64,CP35,CP06,CP07,CP36,CP37,CP40,CP41,CP42) " & _
            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strCP10 & "','" & m_CP12 & "','" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "'," & _
                    "'" & m_CP09 & "','" & ChgSQL(strCP64) & "','" & strCP35 & "','" & ChgSQL(DBDATE(textCP06)) & "','" & ChgSQL(DBDATE(textCP07)) & "'," & _
                    CNULL(ChgSQL(textCP36)) & "," & CNULL(ChgSQL(textCP37_1)) & "," & CNULL(ChgSQL(textCP40)) & "," & CNULL(ChgSQL(textCP41)) & "," & CNULL(ChgSQL(textCP42)) & ") "
   cnnConnection.Execute strSql
   
   'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
   
   If Trim(Text11) <> "" Then
     strSql = "UPDATE CaseProgress SET CP133=" & DBDATE(m_CP05) & ",CP134=" & Text11 & " " & _
              "WHERE CP09='" & strCP09 & "' "
     cnnConnection.Execute strSql
   End If
   
   ' 有輸入下一程序時 (發文日設定為空白), 否則為系統日
   'Modify By Sindy 2021/9/27 +　Or textRvType = "1709" Or textRvType = "1718"
   If IsEmptyText(textCF15) = False Or textRvType = "1709" Or textRvType = "1718" Then
      ' 有輸入承辦人時
      If IsEmptyText(textCP14) = False Then
         strSql = "UPDATE CaseProgress SET CP14 = '" & textCP14 & "' " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
      ' 有輸入承辦期限時
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
      
      'Add By Sindy 2021/9/27
      If textRvType = "1718" Then '變更申請案號
         strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & _
                  " WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
   Else
      'Modified by Lydia 2017/03/08 更新承辦人為操作人員 + ,CP14=" & CNULL(strUserNum)
      strSql = "UPDATE CaseProgress SET CP27 = " & DBDATE(SystemDate()) & " ,CP14=" & CNULL(strUserNum) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
      'add by nickc 2008/01/10 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
      'Remove by Lydia 2017/03/08 已上發文日就沒有承辦期限
'        If Trim(textCP07) = "" Then
'            strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
'                     "WHERE CP09 = '" & strCP09 & "' "
'            cnnConnection.Execute strSql
'        Else
'            If DateDiff("d", ChangeWStringToWDateString(DBDATE(m_CP05)), ChangeWStringToWDateString(DBDATE(textCP07))) <= 30 Then    '無法與上句合併，因為沒有日期時，datediff  會發生  型態不符 的錯誤
'                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(4, DBDATE(m_CP05), 0)) & " " & _
'                         "WHERE CP09 = '" & strCP09 & "' "
'                cnnConnection.Execute strSql
'            Else
'                strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(CompWorkDay(6, DBDATE(m_CP05), 0)) & " " & _
'                         "WHERE CP09 = '" & strCP09 & "' "
'                cnnConnection.Execute strSql
'            End If
'        End If
      'end 2017/03/08
   End If
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若來函性質為專用權消滅時, 准駁日設為專用權消滅日
   If Trim(textRvType) = "1704" Then
      If IsEmptyText(textCP25) = False Then
         strSql = "UPDATE CaseProgress SET CP25 = " & DBDATE(textCP25) & " " & _
                  "WHERE CP09 = '" & strCP09 & "' "
         cnnConnection.Execute strSql
      End If
   End If
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 更新商標基本檔 (是否閉卷欄位)
   'Modify By Sindy 2012/3/5 增加判斷及更新閉卷日期及閉卷原因
'   strSql = "UPDATE TradeMark SET TM29 = '" & textTM29 & "' " & _
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
   
   'Add By Sindy 2021/9/27
   ' 當來函性質為變更申請案號時, 更新商標基本檔的申請案號為新申請案號
   ' 將原TM12存入此筆的C類來函之CP30
   If Trim(textRvType) = "1718" Then
      strSql = "UPDATE CaseProgress " & _
                              "SET CP30='" & Trim(textTM12.Text) & "' " & _
                       "WHERE CP09='" & strCP09 & "' "
      cnnConnection.Execute strSql
      
      '同時清除審定來函日,公告日,審定號,目前准駁及專用權並加註備註,FCT040933
      strSql = "UPDATE TradeMark SET TM12='" & Trim(textTM12_new.Text) & "', " & _
               "TM13=NULL,TM14=NULL,TM15=NULL,TM16=NULL,TM17=NULL,TM20=NULL,TM21=NULL,TM22=NULL,TM58='" & ChangeTStringToTDateString(textCP05S) & "變更申請案號同時清除審定資料;'||TM58 " & _
               "WHERE TM01 = '" & m_TM01 & "' AND " & _
                     "TM02 = '" & m_TM02 & "' AND " & _
                     "TM03 = '" & m_TM03 & "' AND " & _
                     "TM04 = '" & m_TM04 & "'"
      cnnConnection.Execute strSql
   End If
   '2021/9/27 End
   
   '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   ' 若有輸入下一程序時, 新增資料到下一程序檔
   If IsEmptyText(textCF15) = False Then
      strNP22 = GetNextProgressNo()
      strNP14 = Empty
      strNP14 = GetRelatedPerson(m_CP09)
      'Modify By Cheng 2002/09/25
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & textCF15 & "'," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & m_CP13 & "','" & textCP08 & "','" & strNP14 & "','" & textCP64 & "'," & strNP22 & ")"
        'Modify By Cheng 2003/04/07
        '智權人員存最近收文A類接洽記錄單的智權人員
        'Modify By Cheng 2003/09/05
'      strSQL = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
'                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & textCF15 & "'," & _
'                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & textCP64 & "'," & strNP22 & ")"
      'Modified by Lydia 2024/07/02 +ChgSQL
      strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP15,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "','" & textCF15 & "'," & _
                          DBDATE(textCP06) & "," & DBDATE(textCP07) & ",'" & PUB_GetFCTSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "','" & ChgSQL(strCP64) & "'," & strNP22 & ")"
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
            'Modify By Cheng 2003/06/26
            '取消列印接洽結案單
'            'Add By Cheng 2003/06/23
'            '新增列印接洽結案單資料
'            pub_AddressListSN = pub_AddressListSN + 1
'            PUB_AddNewCaseCloseSheet strUserNum, "" & pub_AddressListSN, "" & strNP22, "" & m_TM01, "" & m_TM02, "" & m_TM03, "" & m_TM04
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
   
   'Added by Morgan 2017/5/9 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strCP10
   End If
   'end 2017/5/9
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
     'edit by nick 2004/11/03
     OnSaveData = False
End Function

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
   
   If IsEmptyText(textCF15) = False Then
      Select Case textCF15.Text
         Case "補正":
            textCF15 = "201"
         Case "申請意見書":
            textCF15 = "202"
         Case "異議答辯":
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

'Add by Amy 2022/09/07 +對造頁籤
Private Sub textCP36_GotFocus()
    InverseTextBox textCP36
End Sub

Private Sub textCP37_1_GotFocus()
    InverseTextBox textCP37_1
End Sub

'Add by Amy 2025/01/17
Private Sub textCP37_1_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label13, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
End Sub

Private Sub textCP40_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label19, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      textCP40_GotFocus
   End If
End Sub

Private Sub textCP41_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label20, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      textCP41_GotFocus
   End If
End Sub

Private Sub textCP42_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = Replace(Label11, ":", "")
   Cancel = False
   If CheckLengthIsOK(textCP42, 600, True, strMsg) = False Then
      Cancel = True
      textCP42_GotFocus
   End If
End Sub
'end 2025/01/17

Private Sub textCP40_GotFocus()
    InverseTextBox textCP40
End Sub

Private Sub textCP41_GotFocus()
    InverseTextBox textCP41
End Sub

Private Sub textCP42_GotFocus()
    InverseTextBox textCP42
End Sub
'end 2022/09/07

' 承辦人期限
Private Sub textCP48_Validate(Cancel As Boolean)
'   Dim StrCP10 As String
   Dim strDay As String
   Dim strTemp As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   'Modify By Sindy 2012/4/26 Mark
'   ' 案件性質為被禁止處分或取消禁止處分
'   StrCP10 = "1614"
'   Select Case textRvType
'      Case "1": StrCP10 = "1614"
'      Case "2": StrCP10 = "1615"
'   End Select
   
   Cancel = False
   If IsEmptyText(textCP48) = False Then
      ' 檢查是否為民國日期
      If CheckIsTaiwanDate(textCP48, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的承辦期限"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48_GotFocus
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
      strTit = "資料檢核"
      strMsg = "進度備註資料內容長度太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP64_GotFocus
   End If
End Sub

'92.6.8 ADD BY SONIA
Private Sub textRvType_LostFocus()
   '若來函性質為專用權消滅(1704), 預設為閉卷
   If Me.textRvType.Text = "1704" Then
      Me.textTM29.Text = "Y"
   End If
   
   If m_DeadLine = "" Then 'Added by Morgan 2017/5/9 電子公文
      If IsEmptyText(textRvType) = False Then Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   End If
   
   'Add By Sindy 2021/9/27 行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   If textRvType = "1709" And m_CP10 = "403" Then
      textTM12_new = IIf(m_CP27 <> "", Val(Left(DBDATE(m_CP27), 4) - 1911), "") & "年度行商訴字第號"
   End If
   '2021/9/27 end
End Sub
'92.6.8 END

' 來函性質
Private Sub textRvType_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strDay As String
   
   'Add by Amy 2022/09/26 畫面來函性質輸「其他來函1706」則對造頁籤改為「關係案」,對造字樣改「對方」
   SSTab1.TabCaption(1) = "對造名稱"
   strExc(1) = "對造"
   If textRvType = "1706" Then
        SSTab1.TabCaption(1) = "關係案"
        strExc(1) = "對方"
        Label10.Caption = strExc(1) & Mid(Label10.Caption, 3)
        Label13.Caption = strExc(1) & Mid(Label13.Caption, 3)
        Label19.Caption = strExc(1) & Mid(Label19.Caption, 3)
        Label20.Caption = strExc(1) & Mid(Label20.Caption, 3)
        Label11.Caption = strExc(1) & Mid(Label11.Caption, 3)
   End If
   'end 2022/09/26
   
   textRvType_2 = Empty
   Cancel = False
   '若有輸入來函性質
   If IsEmptyText(textRvType) = False Then
            
      'Add By Cheng 2002/01/10
      If Len(Me.textRvType.Text) <> 4 Then
         Cancel = True
         MsgBox "來函性質欄位值必須為四碼!!!", vbExclamation
         textRvType_GotFocus
         Exit Sub
      End If
      
      ' 只取得國內的案件性質名稱
      textRvType_2 = GetCaseTypeName(m_TM01, textRvType, 0)
      If IsEmptyText(textRvType_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "來函性質代號不存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRvType_GotFocus
      End If
      
      '2009/10/22 ADD BY SONIA
      'modify by sonia 2019/10/9 +1602,1716,1717,1799,1729,1728,1725,及FCT之1721,1722
      If InStr("1001,1002,1003,1004,1005,1006,1102,1403,1602,1716,1717,1799,1729,1728,1725,1721,1722", textRvType) > 0 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "此來函性質不可由此畫面輸入資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRvType_GotFocus
      End If
      '2009/10/22 END
      
      'Added by Lydia 2017/03/08
      'Modified by Lydia 2018/02/06 + 1726 更改來函期限
      'If textRvType = "1724" Then
      If InStr("1724,1726", textRvType) > 0 Then
         Cancel = True
         strTit = "檢核資料"
         'Modified by Lydia 2018/02/06
         'strMsg = "通知已轉他所已改獨立功能, 請由該程式進入 !"
         Select Case textRvType
              Case "1724": strExc(1) = "通知已轉他所"
              Case "1726": strExc(1) = "更改來函期限"
         End Select
         strMsg = strExc(1) & "已改獨立功能, 請由該程式進入 !"
         'end 2018/02/06
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textRvType_GotFocus
      End If
      'end 2017/03/08
      
      ' 以來函性質來計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''      strDay = Empty
''''      strDay = GetWorkDays(m_TM01, m_TM10, textRvType)
''''      If IsEmptyText(strDay) = False Then
''''         ' 90.07.03 modify by louis (承辦期限以實際工作天數來計算)
''''         'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''         textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''      End If
   textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, textRvType, DBDATE(m_CP05), DBDATE(textCP06), textCP09))
   
   'Add By Cheng 2002/01/10
   '若無輸入來函性質
   Else
      Cancel = True
      MsgBox "來函性質欄位值必須輸入且為四碼!!!", vbExclamation
      textRvType_GotFocus
      Exit Sub
   End If
   
EXITSUB:
   ' 來函性質為專用權消滅時才可輸入專用權消滅日欄位
   If Trim(textRvType) = "1704" Then
      textCP25.Locked = False
      textCP25.BackColor = &H80000005
      textCP25.TabStop = True
   Else
      textCP25 = Empty
      textCP25.Locked = True
      textCP25.BackColor = &H8000000F
      textCP25.TabStop = False
   End If
   
   'Add By Sindy 2021/9/27
   '來函性質為變更申請案號時才可輸入新申請案號欄位
   If textRvType = "1718" Then
      Label29 = "新申請案號 :"
      EnableTextBox textTM12_new, True
      textTM12_new.MaxLength = 20
   '行政訴訟之智慧局答辯函輸入(存在cp35以便來函可查詢)
   ElseIf textRvType = "1709" And m_CP10 = "403" Then
      Label29 = "法院案號:"
      EnableTextBox textTM12_new, True
      textTM12_new.MaxLength = 32
   Else
      textTM12_new = Empty
      EnableTextBox textTM12_new, False
   End If
   '2021/9/27 End
   
   '2009/10/14 ADD BY SONIA
   If textRvType = "1719" Then
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

'Add By Sindy 2021/9/27
Private Sub textTM12_new_GotFocus()
Dim intPos As Integer

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
End Sub
Private Sub textTM12_new_Validate(Cancel As Boolean)
Dim strRetrunText As String
   
   If IsEmptyText(textTM12_new) = False And textRvType = "1718" Then
      '檢查申請案號所輸入的長度是否正確
      If PUB_ChkTm12Tm15Length("1", textTM12_new, m_TM01, m_TM02, m_TM03, m_TM04, m_TM10, , , strRetrunText) = False Then
         Cancel = True
         textTM12_new_GotFocus
         Exit Sub
      Else
         textTM12_new = strRetrunText
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
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP06) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP06, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06_GotFocus
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
      
'2011/6/15 CANCEL BY SONIA
'按下確定時才檢查
'      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
'      If IsEmptyText(strDate) = False Then
'         If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
'            strTit = "資料檢核"
'            strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP06_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      Else
'         strTit = "資料檢核"
'         strMsg = "來函記錄中無該筆記錄"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP06_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
   End If
EXITSUB:
End Sub

' 法定期限
Private Sub textCP07_Validate(Cancel As Boolean)
   Dim strDate As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textCP07) = False Then
      ' 檢查日期的格式
      If CheckIsTaiwanDate(textCP07, False) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07_GotFocus
         GoTo EXITSUB
      End If
      
'2011/6/15 CANCEL BY SONIA
'按下確定時才檢查
'      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR17")
'      If IsEmptyText(strDate) = False Then
'         If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
'            strTit = "資料檢核"
'            strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
'            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'            If nResponse = vbCancel Then
'               Cancel = True
'               textCP07_GotFocus
'               GoTo EXITSUB
'            End If
'         End If
'      Else
'         strTit = "資料檢核"
'         strMsg = "來函記錄中無該筆記錄"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP07_GotFocus
'            GoTo EXITSUB
'         End If
'      End If
   End If
EXITSUB:
End Sub

Private Function CheckDataValid()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim Cancel As Boolean
   
   CheckDataValid = False
   ' 機關文號不可為空白
   If IsEmptyText(textCP08) = True Then
      strTit = "檢核資料"
      strMsg = "機關文號不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP08.SetFocus
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
   
   ' 有輸入下一程序時, 本所期限及法定期限不可為空白
   If IsEmptyText(textCF15) = False Then
      ' 有輸入下一程序時, 本所期限不可為空白
      If IsEmptyText(textCP06) = True Then
         strTit = "資料檢核"
         strMsg = "有輸入下一程序時, 本所期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP06.SetFocus
         GoTo EXITSUB
      End If
      ' 有輸入下一程序時, 法定期限不可為空白
      If IsEmptyText(textCP07) = True Then
         strTit = "資料檢核"
         strMsg = "有輸入下一程序時, 法定期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP07.SetFocus
         GoTo EXITSUB
      End If
      ' 本所期限不可大於法定期限
      If IsEmptyText(textCP06) = False And IsEmptyText(textCP07) = False Then
         If Val(textCP06) > Val(textCP07) Then
            strTit = "資料檢核"
            strMsg = "本所期限的日期不可超過法定期限的日期"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP06.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If

   ' 來函性質不可為空白
   If IsEmptyText(textRvType) = True Then
      strTit = "檢核資料"
      strMsg = "來函性質不可為空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textRvType.SetFocus
      GoTo EXITSUB
   End If
   'add by sonia 2022/2/23 FCT-046540
   ' 有本所期限時, 下一程序不可為空白
   If IsEmptyText(textCP06) = False Then
      If IsEmptyText(textCF15) = True Then
         strTit = "檢核資料"
         strMsg = "本所期限有資料時下一程序不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   'end 2022/2/23
   ' 有本所期限時, 承辦期限不可為空白
   If IsEmptyText(textCP06) = False Then
      If IsEmptyText(textCP48) = True Then
         strTit = "檢核資料"
         strMsg = "本所期限有資料時承辦期限不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 承辦期限不可超過本所期限
   If IsEmptyText(textCP06) = False And IsEmptyText(textCP48) = False Then
      If Val(textCP48) > Val(textCP06) Then
         strTit = "資料檢核"
         strMsg = "承辦期限的日期不可超過本所期限的日期"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   ' 來函性質為專用權消滅時, 專用權消滅日不可空白
   If textRvType = "1704" Then
      If IsEmptyText(textCP25) = True Then
         strTit = "資料檢核"
         strMsg = "來函性質為專用權消滅時, 專用權消滅日不可空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2021/9/27
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
   If textRvType = "1718" Then
      If IsEmptyText(TextCP64_1) = True Then
         strTit = "檢核資料"
         strMsg = "來函性質為變更申請案號, 收文文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         TextCP64_1.SetFocus
         GoTo EXITSUB
      End If
   End If
   '2021/9/27 End
   
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
   
   '2011/6/15 ADD BY SONIA 自VALIDATE移過來並調整
   ' 檢查來函記錄檔
      '本所期限
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR16")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               textCP06_GotFocus
               GoTo EXITSUB
            End If
         End If
      Else
         '2011/6/15 MODIFY BY SONIA
'         strTit = "資料檢核"
'         strMsg = "來函記錄中無該筆記錄"
'         nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
'         If nResponse = vbCancel Then
'            Cancel = True
'            textCP06_GotFocus
'            GoTo EXITSUB
'         End If
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then
         Else
            If TAIWANDATE(textCP06) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的本所期限與來函記錄中的本所期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  textCP06_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
      End If
      '法定期限
      strDate = GetMailRecField(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05, "MR17")
      If IsEmptyText(strDate) = False Then
         If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
            strTit = "資料檢核"
            strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
            nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
            If nResponse = vbCancel Then
               textCP07_GotFocus
               GoTo EXITSUB
            End If
         End If
      Else
         If IsMailRecNoTermExist(m_TM01, m_TM02, m_TM03, m_TM04, m_CP05) = False Then  '2011/6/15 ADD BY SONIA
            'modify by sonia 2018/2/9 電子公文都不檢查來函記錄檔
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/5/9 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/5/9 電子公文
               strTit = "資料檢核"
               strMsg = "來函記錄中無該筆記錄"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  textCP07_GotFocus
                  GoTo EXITSUB
               End If
            End If
         '2011/6/15 ADD BY SONIA
         Else
            If TAIWANDATE(textCP07) <> TAIWANDATE(strDate) Then
               strTit = "資料檢核"
               strMsg = "輸入的法定期限與來函記錄中的法定期限日期不同"
               nResponse = MsgBox(strMsg, vbOKCancel + vbQuestion, strTit)
               If nResponse = vbCancel Then
                  textCP07_GotFocus
                  Exit Function
               End If
            End If
         End If
         '2011/6/15 END
      End If
   '2011/6/15 END
   
   'Add By Sindy 2012/7/9 以防修改期限天數或月數,重新計算期限
   If Me.Text10.Enabled = True Then
      Cancel = False
      Text10_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.Text11.Enabled = True Then
      Cancel = False
      Text11_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2012/7/9 End
   
    'Added by Lydia 2021/09/13 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    If PUB_ChkUniText(Me, , True, "TextBox") = False Then
         GoTo EXITSUB
    End If
    
   'Add By Sindy 2021/9/27
   If Me.textTM12_new.Enabled = True Then
      Cancel = False
      textTM12_new_Validate Cancel
      If Cancel = True Then
         textTM12_new.SetFocus
         Exit Function
      End If
   End If
   
   'Add by Amy 2022/09/07 +對造頁籤
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
    
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textRvType_GotFocus()
   InverseTextBox textRvType
End Sub

Private Sub textTM29_GotFocus()
   InverseTextBox textTM29
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
   strFromDate = DBDATE(frm03020408_01.textCP05)
   
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
   strFromDate = DBDATE(frm03020408_01.textCP05)
   
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
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/09 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      End If
      ChgType = True
   End If
End Function

''Add By Sindy 2021/9/27
'' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
'Private Sub InsExpField()
'   If textRvType = "1718" Then
'      EndLetter "12", strCP09, "02", strUserNum
'      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'               "VALUES ('" & "12" & "','" & strCP09 & "','" & "02" & "','" & strUserNum & "'," & _
'               "'" & "收文文號" & "','" & Trim(TextCP64_1.Text) & "')"
'      cnnConnection.Execute strSql
'   End If
'End Sub
