VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm02010402_4 
   BorderStyle     =   1  '單線固定
   Caption         =   "非爭議案核駁輸入"
   ClientHeight    =   5736
   ClientLeft      =   132
   ClientTop       =   996
   ClientWidth     =   9144
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   9144
   Begin TabDlg.SSTab SSTab1 
      Height          =   3615
      Left            =   30
      TabIndex        =   48
      Top             =   2100
      Width           =   9045
      _ExtentX        =   15960
      _ExtentY        =   6371
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "一般"
      TabPicture(0)   =   "frm02010402_4.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label18"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label16"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label15"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label14"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label26"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label25"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label24"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label2"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label7"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label8"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label9"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label11"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label12"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label13"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label19"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label20"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label21"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label22"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label23"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label32"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "textCP14_2"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "textCP35"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "textCP64"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "textCP49"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "textCP48"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "textCP07"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "textCP14"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "textCP06"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "textTM15"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "textCP25"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "textCP08"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "textTM14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "textTMBM07_1"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "textTMBM07_2"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "textTM16S"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "textTM17"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "textPrint"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "Frame1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "Frame2"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "textCP26"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "textCF15_2"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "textCF15"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).ControlCount=   46
      TabCaption(1)   =   "對造名稱"
      TabPicture(1)   =   "frm02010402_4.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "textCP80"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "textCP36"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "textCP41"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "textCP40"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "textCP42"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "textCP37_1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label27"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label35"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label34"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label31"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label29"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label28"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "Label30"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.TextBox textCF15 
         Height          =   264
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   7
         Top             =   1290
         Width           =   732
      End
      Begin VB.TextBox textCF15_2 
         BorderStyle     =   0  '沒有框線
         Height          =   264
         Left            =   2160
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   1290
         Width           =   1692
      End
      Begin VB.TextBox textCP26 
         Height          =   264
         Left            =   6180
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1290
         Width           =   372
      End
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   4050
         TabIndex        =   88
         Top             =   1530
         Width           =   4215
         Begin VB.TextBox Text12 
            Height          =   252
            Left            =   2760
            MaxLength       =   7
            TabIndex        =   16
            Top             =   150
            Width           =   975
         End
         Begin VB.TextBox Text10 
            Height          =   270
            Left            =   840
            MaxLength       =   2
            TabIndex        =   12
            Top             =   150
            Width           =   375
         End
         Begin VB.TextBox Text11 
            Height          =   270
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   14
            Top             =   150
            Width           =   375
         End
         Begin VB.OptionButton Option4 
            Caption         =   "文到          天"
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "        月"
            Height          =   180
            Index           =   1
            Left            =   1560
            TabIndex        =   13
            Top             =   180
            Width           =   855
         End
         Begin VB.OptionButton Option4 
            Caption         =   "                      日"
            Height          =   225
            Index           =   2
            Left            =   2520
            TabIndex        =   15
            Top             =   180
            Width           =   1575
         End
      End
      Begin VB.Frame Frame1 
         Height          =   495
         Left            =   1320
         TabIndex        =   87
         Top             =   1530
         Width           =   2535
         Begin VB.OptionButton Option1 
            Caption         =   "文到當日"
            Height          =   180
            Index           =   0
            Left            =   144
            TabIndex        =   9
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "文到次日"
            Height          =   180
            Index           =   1
            Left            =   1320
            TabIndex        =   10
            Top             =   180
            Width           =   1095
         End
      End
      Begin VB.TextBox textCP80 
         Height          =   264
         Left            =   -73530
         MaxLength       =   39
         TabIndex        =   79
         Top             =   2430
         Width           =   3495
      End
      Begin VB.TextBox textCP36 
         Height          =   264
         Left            =   -73530
         MaxLength       =   200
         TabIndex        =   74
         Top             =   450
         Width           =   7092
      End
      Begin VB.TextBox textCP41 
         Height          =   264
         Left            =   -73530
         TabIndex        =   77
         Top             =   1860
         Width           =   7092
      End
      Begin VB.TextBox textPrint 
         Height          =   264
         Left            =   5640
         MaxLength       =   1
         TabIndex        =   25
         Top             =   3240
         Width           =   372
      End
      Begin VB.TextBox textTM17 
         Height          =   264
         Left            =   6120
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   22
         Top             =   2640
         Width           =   372
      End
      Begin VB.TextBox textTM16S 
         Height          =   264
         Left            =   1410
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   21
         Top             =   2640
         Width           =   372
      End
      Begin VB.TextBox textTMBM07_2 
         Height          =   264
         Left            =   7020
         MaxLength       =   4
         TabIndex        =   6
         Top             =   990
         Width           =   732
      End
      Begin VB.TextBox textTMBM07_1 
         Height          =   264
         Left            =   5820
         MaxLength       =   2
         TabIndex        =   5
         Top             =   990
         Width           =   732
      End
      Begin VB.TextBox textTM14 
         Height          =   264
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   4
         Top             =   990
         Width           =   2532
      End
      Begin VB.TextBox textCP08 
         Height          =   264
         Left            =   1320
         MaxLength       =   40
         TabIndex        =   2
         Top             =   690
         Width           =   2532
      End
      Begin VB.TextBox textCP25 
         Height          =   264
         Left            =   5820
         MaxLength       =   7
         TabIndex        =   1
         Top             =   390
         Width           =   2532
      End
      Begin VB.TextBox textTM15 
         Height          =   264
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   0
         Top             =   390
         Width           =   2532
      End
      Begin VB.TextBox textCP06 
         Height          =   264
         Left            =   1320
         MaxLength       =   7
         TabIndex        =   17
         Top             =   2040
         Width           =   2532
      End
      Begin VB.TextBox textCP14 
         Height          =   264
         Left            =   1320
         MaxLength       =   6
         TabIndex        =   19
         Top             =   2340
         Width           =   732
      End
      Begin VB.TextBox textCP07 
         Height          =   264
         Left            =   5820
         MaxLength       =   7
         TabIndex        =   18
         Top             =   2040
         Width           =   2532
      End
      Begin VB.TextBox textCP48 
         Height          =   264
         Left            =   5820
         MaxLength       =   7
         TabIndex        =   20
         Top             =   2340
         Width           =   2532
      End
      Begin VB.TextBox textCP49 
         Height          =   264
         Left            =   1320
         MaxLength       =   300
         TabIndex        =   23
         Top             =   2940
         Width           =   7692
      End
      Begin MSForms.TextBox textCP40 
         Height          =   300
         Left            =   -73530
         TabIndex        =   76
         Top             =   1560
         Width           =   7092
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP42 
         Height          =   300
         Left            =   -73530
         TabIndex        =   78
         Top             =   2130
         Width           =   7095
         VariousPropertyBits=   679493659
         Size            =   "12509;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP37_1 
         Height          =   792
         Left            =   -73530
         TabIndex        =   75
         Top             =   720
         Width           =   7095
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "12515;1397"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP64 
         Height          =   300
         Left            =   1320
         TabIndex        =   24
         Top             =   3240
         Width           =   2532
         VariousPropertyBits=   -1467989989
         MaxLength       =   128
         ScrollBars      =   2
         Size            =   "4466;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP35 
         Height          =   300
         Left            =   5820
         TabIndex        =   3
         Top             =   690
         Width           =   2532
         VariousPropertyBits=   679493659
         MaxLength       =   32
         Size            =   "4466;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox textCP14_2 
         Height          =   264
         Left            =   2160
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   2340
         Width           =   1692
         VariousPropertyBits=   679493663
         MaxLength       =   20
         Size            =   "2984;466"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label32 
         Caption         =   "來函期限:"
         Height          =   255
         Left            =   120
         TabIndex        =   89
         Top             =   1710
         Width           =   855
      End
      Begin VB.Label Label27 
         Caption         =   "對造商品類別 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   86
         Top             =   2430
         Width           =   1575
      End
      Begin VB.Label Label35 
         Caption         =   "對造號數 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   85
         Top             =   450
         Width           =   975
      End
      Begin VB.Label Label34 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   84
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label31 
         Caption         =   "對造中文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   83
         Top             =   1605
         Width           =   1575
      End
      Begin VB.Label Label29 
         Caption         =   "對造英文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   82
         Top             =   1890
         Width           =   1575
      End
      Begin VB.Label Label28 
         Caption         =   "對造日文名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   81
         Top             =   2130
         Width           =   1575
      End
      Begin VB.Label Label30 
         Caption         =   "對造案件名稱 :"
         Height          =   255
         Left            =   -74850
         TabIndex        =   80
         Top             =   750
         Width           =   1335
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "(N:不印;1:台->各國;2:外->台;3:英文)"
         Height          =   180
         Left            =   6210
         TabIndex        =   73
         Top             =   3270
         Width           =   2745
      End
      Begin VB.Label Label22 
         Caption         =   "列印定稿 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   72
         Top             =   3270
         Width           =   975
      End
      Begin VB.Label Label21 
         Caption         =   "進度備註 :"
         Height          =   255
         Left            =   120
         TabIndex        =   71
         Top             =   3270
         Width           =   1215
      End
      Begin VB.Label Label20 
         Caption         =   "(Y / N)"
         Height          =   255
         Left            =   6660
         TabIndex        =   70
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label19 
         Caption         =   "專用權是否存在 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   69
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "案件目前准駁 :"
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   2640
         Width           =   1170
      End
      Begin VB.Label Label13 
         Caption         =   "期"
         Height          =   255
         Left            =   7830
         TabIndex        =   67
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label12 
         Caption         =   "卷"
         Height          =   255
         Left            =   6660
         TabIndex        =   66
         Top             =   990
         Width           =   375
      End
      Begin VB.Label Label11 
         Caption         =   "公報卷期 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   65
         Top             =   990
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "公告日 :"
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   990
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "審查委員 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   63
         Top             =   690
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "機關文號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "核駁通知日 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   61
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "審定號 :"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   390
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "下一程序 :"
         Height          =   255
         Left            =   120
         TabIndex        =   59
         Top             =   1290
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "本所期限 :"
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "承辦人 :"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2340
         Width           =   1215
      End
      Begin VB.Label Label25 
         Caption         =   "法定期限 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   56
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label26 
         Caption         =   "承辦期限 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   55
         Top             =   2340
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "條款 :"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   2940
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "(N:不算)"
         Height          =   255
         Left            =   6660
         TabIndex        =   53
         Top             =   1290
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "是否算案件數 :"
         Height          =   255
         Left            =   4620
         TabIndex        =   52
         Top             =   1290
         Width           =   1455
      End
      Begin VB.Label Label18 
         Caption         =   "(1:准 , 2:駁)"
         Height          =   255
         Left            =   1920
         TabIndex        =   51
         Top             =   2640
         Width           =   1275
      End
   End
   Begin VB.TextBox textTMKey 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1350
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   210
      Width           =   2532
   End
   Begin VB.TextBox textTM12 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1830
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
   End
   Begin VB.TextBox textCP05 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5850
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   1170
      Width           =   2532
   End
   Begin VB.TextBox textCP10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1350
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2532
   End
   Begin VB.TextBox textTM09 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   5850
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   1500
      Width           =   2532
   End
   Begin VB.TextBox textTM10 
      BorderStyle     =   0  '沒有框線
      Height          =   264
      Left            =   1350
      Locked          =   -1  'True
      MaxLength       =   20
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2532
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8292
      TabIndex        =   29
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6240
      TabIndex        =   26
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   7068
      TabIndex        =   28
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdMod 
      Caption         =   "變更事項(R)"
      Height          =   400
      Left            =   4980
      TabIndex        =   27
      Top             =   70
      Width           =   1200
   End
   Begin MSForms.TextBox textTM23 
      Height          =   264
      Left            =   1350
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   840
      Width           =   7692
      VariousPropertyBits=   679493663
      ForeColor       =   -2147483641
      Size            =   "13568;466"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCP13 
      Height          =   264
      Left            =   5850
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   1800
      Width           =   2532
      VariousPropertyBits=   679493663
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
      Left            =   1350
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   510
      Width           =   7692
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13568;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號 :"
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   47
      Top             =   210
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "案件名稱 :"
      Height          =   255
      Left            =   150
      TabIndex        =   46
      Top             =   540
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "申請人 :"
      Height          =   255
      Left            =   150
      TabIndex        =   45
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "審定號數/申請案號 :"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   44
      Top             =   1170
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "收文日 :"
      Height          =   255
      Index           =   3
      Left            =   4650
      TabIndex        =   43
      Top             =   1170
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質 :"
      Height          =   255
      Index           =   6
      Left            =   150
      TabIndex        =   42
      Top             =   1500
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "商品類別 :"
      Height          =   255
      Index           =   7
      Left            =   4650
      TabIndex        =   41
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家 :"
      Height          =   255
      Index           =   8
      Left            =   150
      TabIndex        =   40
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員 :"
      Height          =   255
      Index           =   11
      Left            =   4650
      TabIndex        =   39
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "frm02010402_4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/28 Form2.0已修改 cmbTM05/textTM23/textTM13/textCP14_2/textCP35/textCP64/textCP37_1/textCP40/textCP42
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
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
' 原智權人員代號
Dim m_CP13 As String
Dim m_CP12 As String
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
'Add By Cheng 2002/01/15
Dim m_intNumBegin As Integer
Dim m_intNumEnd As Integer
Dim m_strNumBegin As String
Dim m_strNumEnd As String
Dim m_txtTM14 As String
Dim m_txtTMBM07_1 As String
Dim m_txtTMBM07_2 As String
Dim m_txtTM16S As String
Dim m_txtTM17 As String
Public m_blnNotFirst As Boolean
'add by nickc 2005/08/04
'Dim m_blnClkChgButton As Boolean '是否有按變更事項鈕
Public m_blnClkChgButton As Boolean '是否有按變更事項鈕 'Modify By Sindy 2012/2/6 Dim->Public
Dim m_TM15 As String 'Add By Sindy 2011/3/2 審定號
Dim m_TM12 As String 'Added by Lydia 2025/09/12 申請案號
Dim BolPrintCaseCheck As Boolean 'Add By Sindy 2012/4/16
Dim BolPrintLetterDemand As Boolean 'Add By Sindy 2012/4/16
Dim strRvType As String 'Add By Sindy 2012/4/26

'Added by Morgan 2017/4/20 電子公文
Public m_DocWord As String
Public m_DocNo As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/4/20
'Add By Sindy 2019/5/10
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Dim m_PrevForm As Form
'2019/5/10 END
Dim strFromDate As String 'Added by Lydia 2019/06/21 期限起算日
Dim strLD18 As String 'Add By Sindy 2019/12/19 信函總收文號
Dim m_TM44 As String 'Add By Sindy 2019/12/19 FC代理人
Public PreResult As String 'Add by Amy 2022/09/26 前畫面結果

'Add By Sindy 2019/5/13
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

' 原資料是否有實際結果
Private Sub cmdCancel_Click()
   Unload Me
   frm02010402_3.Show
End Sub

Private Sub cmdExit_Click()
    Me.Enabled = False
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   Unload frm02010402_3
   Unload frm02010402_2
   Unload frm02010402_1
   Unload Me
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
   'Add by Morgan 2003/11/21
   BolPrintCaseCheck = CaseCheck(m_TM01, m_TM02, m_TM03, m_TM04, m_TM10)
   '---end
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
          'add by nickc 2005/04/22
          'Pub_EndModCashMsg m_TM10        '2009/11/11 CANCEL BY SONIA取消結餘詢問
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 儲存資料
    'Modify By Cheng 2002/11/07
'      'OnSaveData
    If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
    'Add By Cheng 2002/11/08
    ' 列印定稿
    If textPrint <> "N" Then
         BolPrintLetterDemand = True 'Add By Sindy 2012/4/16
         PrintLetter
    Else
         BolPrintLetterDemand = False 'Add By Sindy 2012/4/16
    End If
    
      'Add By Sindy 2012/4/16 列印帳款未結清案件資料
      If BolPrintCaseCheck = True And BolPrintLetterDemand = False Then
          Call GetPrintCaseCheck(m_CP09)
      End If
      '2012/4/16 End
      Call PUB_ChkTemporaryReceipts(m_TM01, m_TM02, m_TM03, m_TM04) 'Add By Sindy 2014/5/28 檢查是否有暫收款
      
      'Add By Cheng 2002/01/15
      m_txtTM14 = Me.textTM14.Text
      m_txtTMBM07_1 = Me.textTMBM07_1.Text
      m_txtTMBM07_2 = Me.textTMBM07_2.Text
      'Modify By Cheng 2002/07/22
'      m_txtTM16S = Me.textTM16S.Text
'      m_txtTM17 = Me.textTM17.Text
      m_blnNotFirst = True
      
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      'Unload Me
      Unload frm02010402_3
      Unload frm02010402_2
      'Add By Sindy 2019/5/10
      If Me.m_strIR01 <> "" Then
        Unload frm02010402_1
        If Not m_PrevForm Is Nothing Then
           Call m_PrevForm.GoNext
        End If
        Unload Me
      '2019/5/10 END
      'Modified by Morgan 2017/4/20 電子公文
      'frm02010402_1.Show
      ElseIf m_DocNo <> "" Then
         Unload Me
         Unload frm02010402_1
         frm02010412.GoNext
      Else
         frm02010402_1.Show
         Unload Me
      End If
      'end 2017/4/20
   End If
End Sub

Private Sub Form_Activate()
   If SSTab1.Tab = 0 Then
      If textTM15.Enabled = True Then
         textTM15.SetFocus
      Else
         'modify by sonia 2022/4/22 大陸案不可輸審定號,游標改停在核駁通知日欄
         'textCP08.SetFocus
         If m_TM10 = "020" Then
            textCP25.SetFocus
         Else
            textCP08.SetFocus
         End If
         'end 2022/4/22
      End If
   Else
      textCP36.SetFocus
   End If
End Sub

Private Sub Form_Load()
   ' 設定控制項的背景顏色
   textTMKey.BackColor = &H8000000F
'   textTM08.BackColor = &H8000000F
   textTM09.BackColor = &H8000000F
   textTM10.BackColor = &H8000000F
   textTM23.BackColor = &H8000000F
'   textTM27.BackColor = &H8000000F
'   textTM22S.BackColor = &H8000000F
'   textTM45.BackColor = &H8000000F
'   textCP05.BackColor = &H8000000F
'   textCP05S.BackColor = &H8000000F
   textTM12.BackColor = &H8000000F
   textCP10.BackColor = &H8000000F
   textCP13.BackColor = &H8000000F
   textCP14_2.BackColor = &H8000000F
   textCF15_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
    'Add By nickc 2005/08/04
'    m_blnClkChgButton = False

   ' 顯示畫面為第一頁
   SSTab1.Tab = 0
   
   'Add By Sindy 2019/5/10
   m_strIR01 = frm02010402_1.m_strIR01
   m_strIR02 = frm02010402_1.m_strIR02
   m_strIR03 = frm02010402_1.m_strIR03
   m_strIR04 = frm02010402_1.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2019/5/10 END
  
   strFromDate = DBDATE(frm02010402_1.textCP05)  'Added by Lydia 2019/06/21
   'Add by Amy 2022/09/26 輸「核駁」,對造改為「關係案」,對造字樣改為「對方」
   If PreResult = "1" Then
        SSTab1.TabCaption(1) = "關係案"
        strExc(1) = "對方"
        Label35.Caption = strExc(1) & Mid(Label35.Caption, 3)
        Label34.Caption = strExc(1) & Mid(Label34.Caption, 3)
        Label31.Caption = strExc(1) & Mid(Label31.Caption, 3)
        Label29.Caption = strExc(1) & Mid(Label29.Caption, 3)
        Label28.Caption = strExc(1) & Mid(Label28.Caption, 3)
        Label27.Caption = strExc(1) & Mid(Label27.Caption, 3)
   End If
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

Public Sub QueryData()
   Dim strSql As String
   Dim strSub As String
   Dim rsTmp As New ADODB.Recordset
   Dim rsSub As ADODB.Recordset
   Dim strDay As String
   
   m_CP10 = Empty
   m_CP56 = Empty
   m_TM08 = Empty
   m_TM10 = Empty
   m_TM21 = Empty
   m_TM22 = Empty
   m_TM23 = Empty
   
   ' 來函收文日
   '因為畫面放不下取消2008/10/01 add by Toni
'   textCP05S = m_CP05
   'end 2008/10/01
   
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
      textTMKey = m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04
      ' 申請國家
      'Add By Cheng 2002/07/17
      m_NA14 = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_TM10 = rsTmp.Fields("TM10")
         textTM10 = GetNationName(rsTmp.Fields("TM10"), 0)
         m_NA14 = GetNationExtentYear(rsTmp.Fields("TM10"))
      End If
      ' 審定號
      m_TM15 = Empty 'Add By Sindy 2011/3/2
      If IsNull(rsTmp.Fields("TM15")) = False Then
         m_TM15 = rsTmp.Fields("TM15") 'Modify By Sindy 2011/3/2
      End If
      'Add By Sindy 2010/12/31
      ' 審定號數
      If IsNull(rsTmp.Fields("TM15")) = False Then
         textTM12 = rsTmp.Fields("TM15")
      Else
         ' 申請案號
         If IsNull(rsTmp.Fields("TM12")) = False Then
            textTM12 = rsTmp.Fields("TM12")
         End If
      End If
      '2010/12/31 End
      m_TM12 = "" & rsTmp.Fields("TM12") 'Added by Lydia 2025/09/12 申請案號
      
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
      ' 商標種類
         '因為畫面放不下取消2008/10/01 add by Toni
'      If IsNull(rsTmp.Fields("TM08")) = False Then
'         m_TM08 = rsTmp.Fields("TM08")
'         If m_TM10 < "010" Then
'            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 0)
'         Else
'            textTM08 = GetTradeMarkName(rsTmp.Fields("TM08"), 1)
'         End If
'      End If
      'END BY TONI 2008/10/01
      
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         textTM09 = rsTmp.Fields("TM09")
      End If
      ' 申請人
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_TM23 = rsTmp.Fields("TM23")
         textTM23 = GetCustomerName(rsTmp.Fields("TM23"), 0)
      End If
      'Add By Sindy 2019/12/19
      ' FC代理人
      m_TM44 = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_TM44 = rsTmp.Fields("TM44")
      End If
      '2019/12/19 END
      
'      ' 正商標號數
   '因為畫面放不下取消2008/10/01 add by Toni
'      If IsNull(rsTmp.Fields("TM27")) = False Then
'         textTM27 = rsTmp.Fields("TM27")
'      End If

      ' 彼所案號
         '因為畫面放不下取消2008/10/01 add by Toni
'      If IsNull(rsTmp.Fields("TM45")) = False Then
'         textTM45 = rsTmp.Fields("TM45")
'      End If
      'END ADD BY TONI 2008/10/01
      
      'add by nickc  2006/11/20
      textPrint = CheckStr(rsTmp.Fields("TM77"))
      
      ' 正商標專用期止日
         '因為畫面放不下取消2008/10/01 add by Toni
'      Set rsSub = New ADODB.Recordset
'      strSub = "SELECT * FROM TradeMark " & _
'               "WHERE TM15 = '" & textTM27 & "' AND " & _
'                     "TM10 = '" & m_TM10 & "' "
'      rsSub.CursorLocation = adUseClient
'      rsSub.Open strSub, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
'      If rsSub.RecordCount > 0 Then
'         rsSub.MoveFirst
'         If IsNull(rsSub.Fields("TM22")) = False Then
'            textTM22S = rsSub.Fields("TM22")
'         End If
'      End If
'      rsSub.Close
'      Set rsSub = Nothing
      'END ADD BY TONI 2008/10/01
      
      'Modify By Cheng 2002/04/29
'      ' 公告日
'      If IsNull(rsTmp.Fields("TM14")) = False Then
'         textTM14 = TAIWANDATE(rsTmp.Fields("TM14"))
'      End If
      'Add By Cheng 2002/07/22
      Me.textTM16S.Text = "" & rsTmp.Fields("tm16").Value
      
      ' 專用權是否存在
      If IsNull(rsTmp.Fields("TM17")) = False Then
         textTM17 = rsTmp.Fields("TM17")
      End If
      
      'Added by Lydia 2019/06/21 台-大核駁案期限管制:取消來函期限
      If m_TM10 <> "000" And frm02010402_3.GetSelectResult = "1" Then
          Label32.Caption = "來函類別:"
          Option1(0).Caption = "紙本公文"
          Option1(1).Caption = "電子公文"
          Frame2.Visible = False
      End If
      
   End If
   rsTmp.Close
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
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
         'Modify By Sindy 2012/5/31 Mark
         'textCP08 = rsTmp.Fields("CP08")
      End If
'      ' 收文號
'      If IsNull(rsTmp.Fields("CP09")) = False Then
'         textCP09 = rsTmp.Fields("CP09")
'      End If
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
      'Add By Cheng 2002/07/17
      m_CP13 = Empty
      If IsNull(rsTmp.Fields("CP13")) = False Then
         m_CP13 = rsTmp.Fields("CP13")
         textCP13 = GetStaffName(rsTmp.Fields("CP13"))
      End If
      '業務區       nick 91.08.22
      m_CP12 = Empty
      If IsNull(rsTmp.Fields("cp12")) = False Then
        m_CP12 = rsTmp.Fields("cp12")
      End If
      ' 核准通知日
      'Add/Modify By Cheng 2002/04/30
      '若申請國家為台灣時, 則核駁通知日隱藏且Disable, 並預設來函收文日; 其他則不隱藏且不預設
      If m_TM10 = 台灣國家代號 Then
         Me.textCP25.Text = frm02010402_1.textCP05.Text
         Me.textCP25.Visible = False
         Me.Label7.Visible = False
      End If
'      If IsNull(rsTmp.Fields("CP25")) = False Then
'         textCP25 = TAIWANDATE(rsTmp.Fields("CP25"))
'      End If

      ' 審查委員
      If IsNull(rsTmp.Fields("CP35")) = False Then
         textCP35 = rsTmp.Fields("CP35")
      End If
      ' 移轉申請人代號
      If IsNull(rsTmp.Fields("CP56")) = False Then
         m_CP56 = rsTmp.Fields("CP56")
      End If
      ' 下一程序
      textCF15 = GetNextProgress(m_TM01, m_TM10, m_CP10)
      If IsEmptyText(textCF15) = False Then
         If m_TM10 = "000" Then
            textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 0)
         Else
            textCF15_2 = GetCaseTypeName(m_TM01, textCF15, 1)
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("CP06")) = False Then
         If IsEmptyText(rsTmp.Fields("CP06")) = False Then
            textCP06 = TAIWANDATE(rsTmp.Fields("CP06"))
         End If
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("CP07")) = False Then
         If IsEmptyText(rsTmp.Fields("CP07")) = False Then
            textCP07 = TAIWANDATE(rsTmp.Fields("CP07"))
         End If
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         textCP14 = rsTmp.Fields("CP14")
         textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
      End If
   End If
   rsTmp.Close
   
   'ADD BY SONIA 2015/9/16 台灣案若申請或分割之核駁且曾有A類申請意見書發文者則不出定稿但要印回覆單,承辦人改掛申請意見書之承辦人,由承辦人做撰寫信函
   If m_TM10 = "000" And (m_CP10 = "101" Or m_CP10 = "308") Then
      strSql = "SELECT CP14 FROM TRADEMARK,CASEPROGRESS " & _
            "WHERE TM01 = '" & m_TM01 & "' AND " & _
                  "TM02 = '" & m_TM02 & "' AND " & _
                  "TM03 = '" & m_TM03 & "' AND " & _
                  "TM04 = '" & m_TM04 & "' AND TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND '202'=CP10 AND CP09<'B' AND NVL(CP27,0)>0"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
      If rsTmp.RecordCount > 0 Then
         textPrint = "N"
         textPrint.Locked = True
         ' 承辦人
         If IsNull(rsTmp.Fields("CP14")) = False Then
            textCP14 = rsTmp.Fields("CP14")
            textCP14_2 = GetStaffName(rsTmp.Fields("CP14"))
         End If
      End If
      rsTmp.Close
   End If
   'END 2015/9/16
   
   Call ChgType 'Add By Sindy 2012/4/17 讀取來函期限
   
   'Modify By Cheng 2002/04/29
'   ' 案件性質為申請, 申請國家為台灣時, 以審定號數+商標種類代號抓商標公報檔, 帶出卷期
'   If m_CP10 = "101" And m_TM10 < "010" Then
'      strSQL = "SELECT * FROM TMBULLETIN " & _
'               "WHERE TMBM01 = '" & textTM15 & "' AND " & _
'                     "TMBM02 = '" & m_TM08 & "' "
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
'      If rsTmp.RecordCount > 0 Then
'         rsTmp.MoveFirst
'         If IsNull(rsTmp.Fields("TMBM07")) = False Then
'            textTMBM07_1 = Mid(rsTmp.Fields("TMBM07"), 1, 2)
'            textTMBM07_2 = Mid(rsTmp.Fields("TMBM07"), 3, 3)
'         End If
'      End If
'      rsTmp.Close
'   End If
   
   ' 以案件性質"核駁"或"改變原處分"計算承辦期限
''''edit by nickc 2007/10/11 改抓有時效性的
''''   strDay = Empty
   Select Case frm02010402_3.GetSelectResult
      Case "1":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1002")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1002", DBDATE(m_CP05), DBDATE(textCP06), m_CP09))
      Case "2":
''''         strDay = GetWorkDays(m_TM01, m_TM10, "1403")
            textCP48 = TAIWANDATE(Pub_GetHandleDay(m_TM01, m_TM10, "1403", DBDATE(m_CP05), DBDATE(textCP06), m_CP09))
   End Select
''''   If IsEmptyText(strDay) = False Then
''''      ' 90.07.03 modify by louis (承辦期限以實際工作天數計算)
''''      'textCP48 = TAIWANDATE(DateSerial(Val(DBYEAR(m_CP05)), Val(DBMONTH(m_CP05)), Val(DBDAY(m_CP05)) + Val(strDay)))
''''      textCP48 = TAIWANDATE(CompWorkDay(Val(strDay), DBDATE(m_CP05), 0))
''''   End If
   
   ' 非A類收文其預設為不可算案件數
   textCP26 = "N"
   
   ' 90.11.19 modify by sonia
   Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   
   If m_TM10 < "010" Then
      If textCP08 = "" Then
         textCP08 = "（" & strTmp & "）智商字第號"
      End If
      'Add By Cheng 2002/01/15
      m_strNumBegin = "商"
      m_strNumEnd = "字"
   End If
      
   Set rsTmp = Nothing
   
   'Added by Morgan 2017/4/20 電子公文
   If m_DocNo <> "" Then
      If m_DocWord <> "" Then
         textCP08 = m_DocWord & "字第" & PUB_GetEDocNo(m_DocNo) & "號"
      Else
         textCP08 = Replace(textCP08, "第號", "第" & PUB_GetEDocNo(m_DocNo) & "號")
      End If
      textCP08_LostFocus
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
   'end 2017/4/17
   
   'Add By Cheng 2002/01/15
   If m_blnNotFirst Then
      Me.textTM14.Text = m_txtTM14
      Me.textTMBM07_1.Text = m_txtTMBM07_1
      Me.textTMBM07_2.Text = m_txtTMBM07_2
      'Modify By Cheng 2002/07/22
'      Me.textTM16S.Text = m_txtTM16S
'      Me.textTM17.Text = m_txtTM17
   End If
   
   'Modify By Cheng 2002/07/22
   '若案件性質為申請 "101", 且前一畫面(frm02010402_3)結果欄為"1","2"
'   'Add By Cheng 2002/07/11
'   '若案件性質為申請 "101"
'   If m_CP10 = "101" Then
   '若案件性質為申請 "101"
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
   'edit by nickc 2005/06/24 加入領土延伸
   'If (m_CP10 = "101" Or m_CP10 = "308") And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
   If (m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "104") And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
      'Modify By Cheng 2002/07/22
      '是否更新基本檔目前准駁預設為"2"(駁)
'      '是否更新基本檔目前准駁預設為"Y"
'      Me.textTM16S.Text = "Y"
      Me.textTM16S.Text = "2"
   '其他案件性質
   Else
'      '是否更新基本檔目前准駁預設為"N"
'      Me.textTM16S.Text = "N"
   End If
   
   'Add By Cheng 2002/05/08
   '若前一畫面(frm02010402_3)結果欄為"3"(申請駁回)時, 預設"是否更新基本檔目前准駁"欄為"N", 且不可修改;
   '並預設"進度備註"欄為"申請駁回"
   If frm02010402_3.GetSelectResult = "3" Then
      'Modify By Cheng 2002/07/22
'      Me.textTM16S.Text = "N"
'      Me.textTM16S.Enabled = False
      Me.textCP64.Text = "申請駁回"
   End If
   'Add By Cheng 2002/07/22
   '若案件性質為"延展"(102)時, "專用權是否存"預設為"N"
   If m_CP10 = "102" Then
      Me.textTM17.Text = "N"
   End If
    'Add By Cheng 2002/11/27
   '若案件性質為"申請"(101)時, "專用權是否存"預設為"N"
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      Me.textTM17.Text = "N"
   End If
   
   'add by nickc 2006/06/30 帶列印定稿預設值
   If textPrint = "" Then
        textPrint = GetTWordLng(m_TM01, m_TM02, m_TM03, m_TM04)
   End If
   
   'Add By Sindy 2011/3/2
   'modify by sonia 2022/4/22 大陸案不可輸審定號
   'If m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "104" Then
   If (m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "104") And m_TM10 <> "020" Then
   'end 2022/4/22
      textTM15.Enabled = True
      textTM15.Visible = True
      Label2.Visible = True
      textTM15 = m_TM15
      textTM15.SetFocus
   Else
      textTM15.Enabled = False
      textTM15.Visible = False
      Label2.Visible = False
      'modify by sonia 2022/4/22 大陸案不可輸審定號,游標改停在核駁通知日欄
      'textCP08.SetFocus
      If m_TM10 = "020" Then
         textCP25.SetFocus
      Else
         textCP08.SetFocus
      End If
      'end 2022/4/22
   End If
   '2011/3/2 End
   
   'Modify By Sindy 2013/12/30 Mark
'   'Modify By Sindy 2013/12/19 台灣案的申請則鎖住
'   If m_TM10 = "000" And m_CP10 = "101" Then
'      textTM14.Enabled = False
'      textTMBM07_1.Enabled = False
'      textTMBM07_2.Enabled = False
'   End If
'   '2013/12/19 END
End Sub

Private Sub DisplayNextForm()
   frm02010402_5.SetData 0, m_TM01, True
   frm02010402_5.SetData 1, m_TM02, False
   frm02010402_5.SetData 2, m_TM03, False
   frm02010402_5.SetData 3, m_TM04, False
   frm02010402_5.SetData 5, m_CP09, False
   Me.Hide
   frm02010402_5.Show
   frm02010402_5.QueryData
End Sub

'Modify By Cheng 2002/11/07
'Public sub OnSaveData()
Public Function OnSaveData() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bUpdate As Boolean
   Dim strSubTMSQL As String
   Dim strSubCPSQL As String
   Dim strCP09 As String
   Dim strCP12 As String
   Dim strCP48 As String
   Dim strNP07 As String
   Dim strNP08 As String
   Dim strNP09 As String
   Dim strNP14 As String
   Dim strNP22 As String
   'Add By Sindy 2012/4/26
   Dim strCP133 As String
   Dim strCP134 As String
   '2012/4/26 End
   'Add by Amy 2017/11/13
   Dim m_CP06 As String, m_CP07 As String, st_CP09 As String, m_CP14 As String, strMsg As String
   Dim bolUpdCP As Boolean '是否更新進度檔
   
   ' 取得案件進度檔檔案中欄位
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount <= 0 Then
      GoTo EXITSUB
   End If
   
'Add By Cheng 2002/11/07
On Error GoTo ErrorHandler
OnSaveData = True
cnnConnection.BeginTrans
   
   rsTmp.MoveFirst
   
   ' 設定SQL中Update TradeMark的語法
   strSubTMSQL = "WHERE TM01 = '" & m_TM01 & "' AND " & _
                        "TM02 = '" & m_TM02 & "' AND " & _
                        "TM03 = '" & m_TM03 & "' AND " & _
                        "TM04 = '" & m_TM04 & "' "
   ' 設定SQL中CaseProgress的語法
   strSubCPSQL = "WHERE CP09 = '" & m_CP09 & "' "
   
   ' 當案件性質為延展, 且原實際結果或准駁日無資料時需 Update 實際結果及准駁日的欄位
   'Modify By Cheng 2002/07/16
   '當前一畫面的結果欄為"1"或"3"
'   If frm02010402_3.GetSelectResult() = "1" Then
   If frm02010402_3.GetSelectResult() = "1" Or frm02010402_3.GetSelectResult() = "3" Then
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
         strSql = "UPDATE CaseProgress SET CP24 = '2', " & _
                                          "CP25=" & DBDATE(textCP25) & ", " & _
                                          "CP35='" & textCP35 & "' "
         strSql = strSql & strSubCPSQL
         cnnConnection.Execute strSql
      End If
   End If
   
   'Modify By Cheng 2002/07/22
   '若案件性質為延展(102), 更新商標基本檔的專用權是否存在
'   ' 更新商標基本檔的專用權是否存在
   If m_CP10 = "102" Then
      strSql = "UPDATE TradeMark SET TM17 = '" & textTM17 & "' "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
      
   'add by nickc 2006/11/20
   If textPrint <> "N" Then
      strSql = "UPDATE TradeMark SET TM77 = '" & textPrint & "' "
      strSql = strSql & strSubTMSQL
      cnnConnection.Execute strSql
   End If
   
   ' 案件性質為申請時
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   'edit by nickc 2005/06/24 加入領土延伸
   'If m_CP10 = "101" Or m_CP10 = "308" Then
   If m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "104" Then
      If frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2" Then 'Add By Sindy 2013/7/15 +if
         If IsEmptyText(textTM14) = True Then
            '因為畫面放不下取消2008/10/01 add by Toni
            '將textCP05S改為textCP05因為畫面放不下取消
   '         strSQL = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
   '                                    "TM13 = " & DBDATE(textCP05S) & "," & _
   '                                    "TM14 = " & "NULL" & " "
            'END BY TONI 2008/10/01
            strSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                       "TM13 = " & DBDATE(textCP05) & "," & _
                                       "TM14 = " & "NULL" & " "
         Else
            ' 更新審定號, 來函收文日, 公告日
            '因為畫面放不下取消2008/10/01 add by Toni
            '將textCP05S改為textCP05因為畫面放不下取消
   '         strSQL = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
   '                                       "TM13 = " & DBDATE(textCP05S) & "," & _
   '                                       "TM14 = " & DBDATE(textTM14) & " "
            'END BY TONI 2008/10/01
            strSql = "UPDATE TradeMark SET TM15 = '" & textTM15 & "'," & _
                                                   "TM13 = " & DBDATE(textCP05) & "," & _
                                                   "TM14 = " & DBDATE(textTM14) & " "
         End If
         strSql = strSql & strSubTMSQL
         cnnConnection.Execute strSql
      End If
      'Modify By Cheng 2002/07/22
      '當案件性質為商申(101)時, 且前一畫面(frm02010402_3)結果欄為"1","2", 更新目前准/駁及審定來函日兩個欄位
'      ' 當使用者輸入要更新基本檔之准/駁時, 更新目前准/駁及審定來函日兩個欄位
'      If textTM16S = "Y" Then
      'edit by nick 2004/12/23 分割與申請做相同的事情
      'If m_CP10 = "101" And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
      'edit by nickc 2005/06/24 加入領土延伸
      'If (m_CP10 = "101" Or m_CP10 = "308") And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
      If (m_CP10 = "101" Or m_CP10 = "308" Or m_CP10 = "104") And (frm02010402_3.GetSelectResult = "1" Or frm02010402_3.GetSelectResult = "2") Then
         strSql = "UPDATE TradeMark SET TM16='2'," & _
                                       "TM13=" & DBDATE(textCP25) & " "
         strSql = strSql & strSubTMSQL
         cnnConnection.Execute strSql
      End If
      ' 當系統別為"TF", 且案件性質為申請時, 將商標基本檔的專用期限清為空白
      If m_TM01 = "TF" Then
         strSql = "UPDATE TradeMark SET TM21=''," & _
                                       "TM22='' "
         strSql = strSql & strSubTMSQL
         cnnConnection.Execute strSql
      End If
   End If
   
   '  新增資料到案件進度檔
   ' 收文號
   strCP09 = Empty
   strCP09 = AutoNo("C", 6)
   ' 案件性質
   strRvType = "1002"
   Select Case frm02010402_3.GetSelectResult
      Case "1": strRvType = "1002"
      Case "2": strRvType = "1403"
   End Select
   ' 業務區別
   'strCP12 = GetST15(m_CP13)
   ' 91.03.25 modify by louis (單引號)
    '承辦人為原程序承辦人, 不上發文日
    'Modify By Cheng 2003/04/03
    '智權人員存最近收文A類接洽記錄單的智權人員
    'Modify By Cheng 2004/02/03
'   strSQL = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP43,CP49,CP64) " & _
'            "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
'                    CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & "," & _
'                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & m_CP12 & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
'                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "') "
    '業務區為最近收文A類接洽記錄單智權人員的業務區
    'add by toni 新增加對造號數CP36,對造案件名稱(中)CP37,對造名稱(中)CP40,對造名稱(英)CP41,對造名稱(日)CP42 2008/10/01
    'Add By Sindy 2012/4/26 儲存官方發文日及官方期限月數
    strCP133 = "": strCP134 = ""
    If Trim(Text11) <> "" Then
      strCP133 = DBDATE(m_CP05)
      strCP134 = Text11
    End If
    'Modify By Sindy 2012/4/26 +CP133,CP134
    strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP40,CP41,CP42,CP80,CP43,CP49,CP64,CP133,CP134) " & _
             "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                    CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & "," & _
                    "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                    "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & ChgSQL(textCP42) & "'," & _
                    "'" & ChgSQL(textCP80) & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "'," & CNULL(strCP133) & "," & CNULL(strCP134) & ") "
    'add by sonia 2018/10/29 台灣案若申請或分割沒有提出申復意見書之核駁審定,由程序人員直接以定稿函知客戶故改為系統直接上發文日
    'modify by sonia 2018/11/6 再加非MCTF案件T-212390
    If m_TM10 = "000" And (m_CP10 = "101" Or m_CP10 = "308") And textPrint.Locked = False And Left(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), 3) <> "MCT" Then
      strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP08,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP32,CP35,CP36,CP37,CP40,CP41,CP42,CP80,CP43,CP49,CP64,CP133,CP134,CP27) " & _
               "VALUES ('" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & DBDATE(m_CP05) & "," & _
                      CNULL(DBDATE(textCP06)) & "," & CNULL(DBDATE(textCP07)) & "," & _
                      "'" & textCP08 & "','" & strCP09 & "','" & strRvType & "','" & GetSalesArea(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04)) & "','" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP14 & "'," & _
                      "'" & "N" & "','" & textCP26 & "','" & "N" & "','" & textCP35 & "','" & ChgSQL(textCP36) & "','" & ChgSQL(textCP37_1) & "','" & ChgSQL(textCP40) & "','" & ChgSQL(textCP41) & "','" & ChgSQL(textCP42) & "'," & _
                      "'" & ChgSQL(textCP80) & "','" & m_CP09 & "','" & textCP49 & "','" & ChgSQL(textCP64) & "'," & CNULL(strCP133) & "," & CNULL(strCP134) & "," & CNULL(strSrvDate(1)) & ") "
    End If
    'End 2018/10/29
    cnnConnection.Execute strSql
    
   'Add By Sindy 2019/12/19 商標電子化
   If strSrvDate(1) >= T商標電子化第2階段啟用日 Then
      strLD18 = strCP09
      If Val(textCP06) > 0 Then '有期限者,為掛號
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", True, m_TM23, strRvType, m_TM44
      Else
         PUB_AddLetterProgress strLD18, 1, IIf(textPrint = "N", False, True), "", False, m_TM23, strRvType, m_TM44
      End If
   End If
   '2019/12/19 END
   
   'Added by Lydia 2025/09/12 TF基礎案號設定：基礎案狀態通知Email
   Dim strTFcase As String
   If m_TM01 = "T" And strRvType = "1002" Then
      strTFcase = PUB_GetTFbaseInfo(m_TM01, m_TM02, m_TM03, m_TM04, m_TM15, m_TM10, "2", m_TM12, strCP09)
   End If
   'end 2025/09/12
   
    'add by nick 2004/11/30  更新c類的代理人及彼所案號，要在新增c類之後
    Pub_UpdateFromMaxCP27 m_TM01, m_TM02, m_TM03, m_TM04
   
   '92.11.19 ADD BY SONIA
   If strRvType = "1403" Then
       strSql = "Update CaseProgress Set CP24='2' Where CP09='" & strCP09 & "' "
       cnnConnection.Execute strSql
   End If
   '92.11.19 END
   ' 更新下一程序檔案件性質為催審的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 = " & "305"
   cnnConnection.Execute strSql
   '2007/8/7 ADD BY SONIA更新下一程序檔案件性質為收達及提申的資料
   strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
            "WHERE NP01 = '" & m_CP09 & "' AND " & _
                  "NP02 = '" & m_TM01 & "' AND " & _
                  "NP03 = '" & m_TM02 & "' AND " & _
                  "NP04 = '" & m_TM03 & "' AND " & _
                  "NP05 = '" & m_TM04 & "' AND " & _
                  "NP07 IN (997,998) AND NP06 IS NULL"
   cnnConnection.Execute strSql
   '2007/8/7 END
   ' 有輸入承辦期限時
   If IsEmptyText(textCP48) = False Then
      strSql = "UPDATE CaseProgress SET CP48 = " & DBDATE(textCP48) & " " & _
               "WHERE CP09 = '" & strCP09 & "' "
      cnnConnection.Execute strSql
   'add by nickc 2008/01/09 FCT 加判斷，有期限用期限判斷(第三或第五個工作天)，無期限以第三個工作日(當日不算)，寫入承辦期限
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
   ' 當使用者在前畫面選取2時, 更新下一程序檔案件性質為改變原處分的資料
   If frm02010402_3.GetSelectResult() = "2" Then
      strSql = "UPDATE NextProgress SET NP06 = '" & "Y" & "' " & _
               "WHERE NP01 = '" & m_CP09 & "' AND " & _
                     "NP02 = '" & m_TM01 & "' AND " & _
                     "NP03 = '" & m_TM02 & "' AND " & _
                     "NP04 = '" & m_TM03 & "' AND " & _
                     "NP05 = '" & m_TM04 & "' AND " & _
                     "NP07 = " & "1403"
      cnnConnection.Execute strSql
   End If
   
   ' 有輸入下一程序時, 新增資料到下一程序檔
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
    'Modify By Cheng 2003/04/04
    '智權人員存最近收文A類接洽記錄單的智權人員
        strSql = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP14,NP22) " & _
                "VALUES ('" & strCP09 & "','" & m_TM01 & "','" & m_TM02 & "','" & m_TM03 & "','" & m_TM04 & "'," & textCF15 & "," & _
                            strNP08 & "," & strNP09 & ",'" & PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04) & "','" & textCP08 & "','" & ChgSQL(strNP14) & "'," & strNP22 & ")"
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
            '   2008/05/12 改 大陸再控制，台灣有定稿
            'MODIFY BY SONIA 2015/9/16 台灣案若申請或分割之核駁且曾有A類申請意見書發文者則不出定稿但要印回覆單
            If m_TM10 <> "000" Or textPrint.Locked = True Then
                'Modify by Amy 2017/11/16 未更新進度檔才印回覆單
                If bolUpdCP = False Then
                    Call g_PrtForm001.PrintReturnSheet(strCP09, textCF15, DBDATE(strNP09), False, , , , m_TM01 & m_TM02 & m_TM03 & m_TM04)
                End If
            End If
      End Select
   End If
   
   Dim SeekMonTM01 As String
   Dim SeekMonTM02 As String
   Dim SeekMonTM03 As String
   Dim SeekMonTM04 As String
   Dim rsA As New ADODB.Recordset
   'ADD BY nickc 2006/09/27 若是B類申請案，則代表是分割產生，要檢查分割的相關子案是否有准駁，若全都有，則將母案上閉卷
   'MODIFY BY SONIA 2015/8/6 大陸案母案不可閉卷(分割案為母案抽部分出來)分割案T-196252母案T-190094
   'If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" Then
   If Mid(m_CP09, 1, 1) = "B" And m_CP10 = "101" And m_TM10 = "000" Then
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
   '2009/09/24 End
   
   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Modify By Cheng 2002/11/08
'   ' 列印定稿
'   If textPrint <> "N" Then PrintLetter
          'add by nickc 2005/04/22
          'Pub_UpdateEndModCash m_TM01, m_TM02, m_TM03, m_TM04     '2009/11/11 CANCEL BY SONIA取消結餘更新
          
   'Added by Morgan 2017/4/20 電子公文
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, strCP09, m_TM01, m_TM02, m_TM03, m_TM04, strRvType
   End If
   'end 2017/4/20
   'Add by Sindy 2019/5/10
   Call PUB_TMFilePathToCPP(strTMCppFilePath, strCP09) '檢查是否有電子檔要存入卷宗區
   If m_strIR01 <> "" Then
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm02010402_1"
   End If
   '2019/5/10 END
   
   'Add By Cheng 2002/11/07
   cnnConnection.CommitTrans
   Exit Function

ErrorHandler:
   cnnConnection.RollbackTrans
   OnSaveData = False

EXITSUB:
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
   Call PUB_SendMailCache 'Added by Lydia 2025/09/12
   
'edit by nickc 2008/04/25 改整批印
'    '列印接洽接案單
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
   
   PreResult = "" 'Add by Amy 2022/09/26
   'Add By Sindy 2019/5/13
   If Not m_PrevForm Is Nothing Then
      Set m_PrevForm = Nothing
   End If
      
   'Add By Cheng 2002/07/18
   Set frm02010402_4 = Nothing
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 0 Then
      If textTM15.Enabled = True Then
         textTM15.SetFocus
      Else
         textCP08.SetFocus
      End If
   Else
      textCP36.SetFocus
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
         textCP06_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2002/11/18
        '按確定時才檢查
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
        '按下確定才檢查
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

Private Sub textCP25_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   '若有輸入核駁通知日
   If IsEmptyText(textCP25) = False Then
      ' 檢核是否為民國日期
      If CheckIsTaiwanDate(textCP25) = False Then
         Cancel = True
         textCP25_GotFocus
      End If
      If Val(TAIWANDATE(textCP25)) > Val(TAIWANDATE(Date)) Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "核駁通知日不可超過系統日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP25_GotFocus
      End If
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
         Case "", " ", "N":
         Case Else:
            Cancel = True
            strTit = "資料檢核"
            strMsg = "只可輸入空白或N"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textCP26_GotFocus
      End Select
   End If
End Sub

' 審查委員
Private Sub textCP35_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP35, 128) = False Then
      Cancel = True
      strTit = "資料檢核"
      strMsg = "審查委員欄位內容太長"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP35_GotFocus
   End If
End Sub

Private Sub textCP36_GotFocus()
    InverseTextBox textCP36
    CloseIme
End Sub

'Add by Amy 2022/09/29
Private Sub textCP36_Validate(Cancel As Boolean)
   Dim strTit As String, strMsg As String, nResponse
   strTit = "檢核資料"
   strMsg = "對造號數"
   Cancel = False
   If CheckLengthIsOK(textCP36, textCP36.MaxLength, True, strMsg) = False Then
      Cancel = True
      textCP36_GotFocus
   End If
End Sub

'ADD BY TONI 2008/10/01
Private Sub textCP37_1_GotFocus()
   TextInverse Me.textCP37_1
    OpenIme
End Sub
'END BY TONI 2008/10/01

'ADD BY TONI 2008/10/01
Private Sub textCP37_1_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   strTit = "檢核資料"
   strMsg = "對造案件名稱"
   Cancel = False
   'Modify by Amy 2025/01/17 原:140
   If CheckLengthIsOK(textCP37_1, 160, True, strMsg) = False Then
      Cancel = True
      textCP37_1_GotFocus
   End If
   'edit by nickc 2007/06/06 切換輸入法改用API
   If Cancel = False Then CloseIme

End Sub
'END BY TONI 2008/10/01

'ADD BY TONI 2008/10/01
Private Sub textCP40_GotFocus()
   InverseTextBox textCP40
    OpenIme
End Sub
'end 2008/10/01

'ADD BY TONI 2008/10/01
'對造案件 (中)
Private Sub textCP40_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   strTit = "檢核資料"
   strMsg = "對造案件(中)"
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP40, 600, True, strMsg) = False Then
      Cancel = True
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP40_GotFocus
   End If

End Sub
'end by Toni 2008/10/01

'Add by Toni 2008/10/01
Private Sub textCP41_GotFocus()
    InverseTextBox textCP41
    CloseIme
End Sub
'end 2008/10/01

'ADD BY TONI 2008/10/01
'對造案件(英)
Private Sub textCP41_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   strTit = "檢核資料"
   strMsg = "對造案件(英)"
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP41, 600, True, strMsg) = False Then
      Cancel = True
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP41_GotFocus
   End If

End Sub
'end by Toni 2008/10/01

Private Sub textCP42_GotFocus()
   InverseTextBox textCP42
   OpenIme
End Sub

'Add By Sindy 2009/07/02
Private Sub textCP80_GotFocus()
   InverseTextBox textCP80
   CloseIme
End Sub

'ADD BY TONI 2008/10/01
'對造案件(日)
Private Sub textCP42_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   strTit = "檢核資料"
   strMsg = "對造案件(日)"
   Cancel = False
   'Modify by Amy 2025/01/17 原:100
   If CheckLengthIsOK(textCP42, 600, True, strMsg) = False Then
      Cancel = True
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP42_GotFocus
   End If
End Sub
'end by Toni 2008/10/01

'Add By Sindy 2009/07/02
'對造商品類別
Private Sub textCP80_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   Cancel = False
   If CheckLengthIsOK(textCP80, textCP80.MaxLength) = False Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "對造商品類別欄位內容太長"
      textCP80_GotFocus
   End If
End Sub

' 承辦期限
Private Sub textCP48_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
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
         Exit Sub
      End If
   End If
   'Add By Cheng 2002/05/06
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         Cancel = True
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
      
      ' 90.08.12 modify by sonia
      ' 檢查主張內容分類表
      'StrSQL = "SELECT * FROM ClaimContents " & _
      '         "WHERE CC01 = '" & Right(strTemp, 1) & "'"
      'rsTmp.CursorLocation = adUseClient
      'rsTmp.Open StrSQL, cnnConnection, adOpenDynamic
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

Private Sub textPrint_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'add by nickc 2006/06/29
   If KeyAscii <> 78 And KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 51 And KeyAscii <> 8 And KeyAscii <> 13 Then
       KeyAscii = 0
   End If
End Sub

' 是否列印定稿
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

' 公告日
Private Sub textTM14_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   '若有輸入公告日
   If IsEmptyText(textTM14) = False Then
      ' 檢核是否為民國日期
      If CheckIsTaiwanDate(textTM14) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "請輸入正確的公告日"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTM14_GotFocus
      End If
      'If Val(TAIWANDATE(textTM14)) > Val(TAIWANDATE(Date)) Then
      '   Cancel = True
      '   strTit = "資料檢核"
      '   strMsg = "公告日不可超過系統日"
      '   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '   textTM14_GotFocus
      'End If
   End If
End Sub

'Add By Sindy 2010/8/31
Private Sub textTM15_Validate(Cancel As Boolean)
Dim strRetrunText As String 'Add By Sindy 2017/5/17
   
   If IsEmptyText(textTM15) = False Then
      '檢查審定號所輸入的長度是否正確
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
   End If
End Sub

Private Sub textTM16S_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' 是否更新基本檔目前准駁
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
'            strMsg = "是否更新基本檔目前准駁只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM16S_GotFocus
'      End Select
'   End If
End Sub

Private Sub textTM17_KeyPress(KeyAscii As Integer)
   'Modify By Cheng 2002/07/22
'   KeyAscii = UpperCase(KeyAscii)
End Sub

' 專用權是否存在
Private Sub textTM17_Validate(Cancel As Boolean)
   'Modify By Cheng 2002/07/22
   '取消檢查
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
'            strMsg = "專用權是否存在只可輸入Y或N"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM17_GotFocus
'      End Select
'   End If
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   
   'Add by Amy 2021/12/28檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me, , True) = False Then
        GoTo EXITSUB
   End If

'add by nickc 2005/08/04
   If m_blnClkChgButton = False Then
       MsgBox "請輸入變更事項!!!", vbExclamation + vbOKOnly
       Me.cmdMod.SetFocus
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
   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 檢查法限和所限
   ElseIf textCP07.Tag <> "" And frm02010402_3.GetSelectResult = "1" Then
       If textCP07.Text > textCP07.Tag Then
            MsgBox "法定期限不可早於" & textCP07.Tag, vbExclamation + vbOKOnly
            Me.textCP07.SetFocus
            GoTo EXITSUB
       End If
   End If
   
   ' 核駁通知日不可空白
   If IsEmptyText(textCP25) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入核駁通知日"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP25.SetFocus
      GoTo EXITSUB
   End If
    'Modify By Cheng 2003/07/31
    '若有輸入下一程序
    If Me.textCF15.Text <> "" Then
        ' 本所期限不可空白
        If IsEmptyText(textCP06) = True Then
           strTit = "資料檢核"
           strMsg = "請輸入本所期限"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP06.SetFocus
           GoTo EXITSUB
        End If
    End If
   'Add By Cheng 2002/03/11
   '若有輸入本所期限
   If Me.textCP06.Text <> "" Then
      If Val(Me.textCP06.Text) + 19110000 < ServerDate Then
         MsgBox "本所期限不可小於系統日期!!!", vbExclamation
         Me.textCP06.SetFocus
         textCP06_GotFocus
         GoTo EXITSUB
      End If
   End If
    'Modify By Cheng 2003/07/31
    '若有輸入下一程序
    If Me.textCF15.Text <> "" Then
        ' 法定期限不可空白
        If IsEmptyText(textCP07) = True Then
           strTit = "資料檢核"
           strMsg = "請輸入法定期限"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCP07.SetFocus
           GoTo EXITSUB
        End If
    End If
   ' 本所期限的日期不可超過法定期限的日期
   If Val(textCP06) > Val(textCP07) Then
      strTit = "資料檢核"
      strMsg = "本所期限的日期不可超過法定期限的日期"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textCP06.SetFocus
      GoTo EXITSUB
   End If
   'Modify By Cheng 2002/05/08
   '若前一畫面(frm02010402_3)結果欄不為"3"時, 才要檢查公告日是否要輸入
   If frm02010402_3.GetSelectResult <> "3" Then
      If IsEmptyText(textTM14) = True Then
         'Modify By Cheng 2002/06/12
         '申請國家為台灣時, 且案件性質為"申請"時, 才一定要輸入
'      ' 申請國家為大陸時, 公告日可以不輸入
'         If m_TM10 <> "020" Then
'            strTit = "資料檢核"
'            strMsg = "申請國家非大陸, 一定要輸入公告日"
'            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'            textTM14.SetFocus
'            GoTo EXITSUB
'         End If
         'edit by nick 2004/12/23 分割與申請做相同的事情
         'If m_TM10 = 台灣國家代號 And m_CP10 = "101" Then
         If m_TM10 = 台灣國家代號 And (m_CP10 = "101" Or m_CP10 = "308") Then
            strTit = "資料檢核"
            strMsg = "申請國家為台灣且案件性質為申請時, 一定要輸入公告日"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            textTM14.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   ' 機關文號(申請國家為台灣時不可為空白)
   If IsEmptyText(textCP08) = True Then
      If m_TM10 < "010" Then
         strTit = "資料檢核"
         strMsg = "申請國家為台灣時機關文號不可為空白"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textCP08.SetFocus
         GoTo EXITSUB
      End If
   End If
    'Add By Cheng 2003/07/31
    '若為大陸非申請案時
    'edit by nick 2004/12/23 分割與申請做相同的事情
    'If m_TM10 = 大陸國家代號 And m_CP10 <> "101" Then
    '2005/6/28 modify by sonia
    'If m_TM10 = 大陸國家代號 And m_CP10 <> "101" And m_CP10 <> "308" Then
    If m_TM10 <> 台灣國家代號 And m_CP10 <> "101" And m_CP10 <> "308" Then
    '2005/6/28 end
        '若未輸入下一程序
        If Me.textCF15.Text = "" Then
            If MsgBox("您未輸入下一程序，是否要繼續?", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                textCF15.SetFocus
                textCF15_GotFocus
                GoTo EXITSUB
            End If
        End If
    Else
        ' 下一程序不可空白
        '2009/10/14 MODIFY BY SONIA 加徵求同意書724可不輸下一程序
        'If IsEmptyText(textCF15) = True And frm02010402_3.textResult <> "3" Then
        If IsEmptyText(textCF15) = True And frm02010402_3.textResult <> "3" And m_CP10 <> "724" Then
           strTit = "資料檢核"
           strMsg = "請輸入下一程序"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           textCF15.SetFocus
           GoTo EXITSUB
        End If
    End If
   'Add By Cheng 2002/05/06
   '若本所期限及承辦期限皆有輸入時, 承辦期限不可大於本所期限
   If Len(Me.textCP06.Text) > 0 And Len(Me.textCP48.Text) > 0 Then
      If Val(Me.textCP06.Text) < Val(Me.textCP48.Text) Then
         MsgBox "承辦期限不得大於本所期限!!!", vbExclamation + vbOKOnly
         Me.textCP48.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Modify By Cheng 2002/07/22
'   ' 是否更新基本檔目前准駁
'   If IsEmptyText(textTM16S) = True Then
'      strTit = "資料檢核"
'      strMsg = "請輸入是否更新基本檔目前准駁"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM16S.SetFocus
'      GoTo EXITSUB
'   End If
'   ' 專用權是否存在
'   If IsEmptyText(textTM17) = True Then
'      strTit = "資料檢核"
'      strMsg = "請輸入專用權是否存在"
'      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'      textTM17.SetFocus
'      GoTo EXITSUB
'   End If
    'Add By Cheng 2002/11/12
   ' 申請國家為台灣者, 條款項目不可空白
    '2009/10/14 MODIFY BY SONIA 加徵求同意書724可不輸條款
    'If m_TM10 = 台灣國家代號 Then
    If m_TM10 = 台灣國家代號 And m_CP10 <> "724" Then
        If IsEmptyText(Me.textCP49) = True Then
           strTit = "資料檢核"
           strMsg = "請輸入條款項目"
           nResponse = MsgBox(strMsg, vbOKOnly, strTit)
           Me.textCP49.SetFocus
           GoTo EXITSUB
        End If
    End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub textPrint_GotFocus()
   InverseTextBox textPrint
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Sub textCF15_GotFocus()
   InverseTextBox textCF15
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

' 列印定稿前將例外欄位加入到列印定稿例外欄位檔案中
Private Sub InsExpField()
   Dim strTM23Nation As String
   Dim strSql As String
   Dim strTmp As String
   
   strTM23Nation = Empty
   If IsEmptyText(m_TM23) = False Then: strTM23Nation = GetCustomerNation(m_TM23)
   
   ' 案件性質為申請
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      ' 申請國家為台灣
      If m_TM10 < "010" Then
         ' 申請人國籍為台灣
         'edit by nickc 2006/06/30
         'If strTM23Nation < "010" Then
         If textPrint = "1" Then
            ' 清除定稿例外欄位檔原有資料
            EndLetter "04", m_CP09, "01", strUserNum
            ' 下一程序
            If m_TM10 = "000" Then
               strTmp = GetCaseTypeName(m_TM01, textCF15, 0)
            Else
               strTmp = GetCaseTypeName(m_TM01, textCF15, 1)
            End If
            'add by nickc 2008/04/25 案件回覆單
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "04" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                     "'" & "下一程序" & "','" & textCF15 & "')"
            cnnConnection.Execute strSql
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "04" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                     "'" & "下一程序名稱" & "','" & strTmp & "')"
            cnnConnection.Execute strSql
            ' 本所期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "04" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                     "'" & "本所期限" & "','" & DBDATE(textCP06) & "')"
            cnnConnection.Execute strSql
            ' 法定期限
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "04" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                     "'" & "法定期限" & "','" & DBDATE(textCP07) & "')"
            cnnConnection.Execute strSql
            ' 機關文號
            strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     "VALUES ('" & "04" & "','" & m_CP09 & "','" & "01" & "','" & strUserNum & "'," & _
                     "'" & "機關文號" & "','" & textCP08 & "')"
            cnnConnection.Execute strSql
         End If
      End If
   '2009/1/12 ADD BY SONIA 其他案件性質只印案件回覆單
   Else
      ' 清除定稿例外欄位檔原有資料
      EndLetter "04", m_CP09, "00", strUserNum
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "04" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
               "'" & "下一程序" & "','" & textCF15 & "')"
      cnnConnection.Execute strSql
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "04" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
               "'" & "下一程序名稱" & "','" & strTmp & "')"
      cnnConnection.Execute strSql
      ' 本所期限
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "04" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
               "'" & "本所期限" & "','" & DBDATE(textCP06) & "')"
      cnnConnection.Execute strSql
      ' 法定期限
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "04" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
               "'" & "法定期限" & "','" & DBDATE(textCP07) & "')"
      cnnConnection.Execute strSql
      ' 機關文號
      strSql = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               "VALUES ('" & "04" & "','" & m_CP09 & "','" & "00" & "','" & strUserNum & "'," & _
               "'" & "機關文號" & "','" & textCP08 & "')"
      cnnConnection.Execute strSql
   '2009/1/12 END
   End If
End Sub

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
   ET01 = "04"
   ET02 = m_CP09
   bolEdit = False
   '2012/1/13 End
   
   ' 案件性質為申請
   'edit by nick 2004/12/23 分割與申請做相同的事情
   'If m_CP10 = "101" Then
   If m_CP10 = "101" Or m_CP10 = "308" Then
      ' 申請國家為台灣
      If m_TM10 < "010" Then
        'add by nickc 2006/06/30
        If textPrint = "1" Then
            '93.8.10 cancel by sonia 取消國籍限制
            ' 申請人國籍為台灣
            'If strTM23Nation < "010" Then
            '93.8.10 end
               ' 列印定稿
'               NowPrint m_CP09, "04", "01", False, strUserNum, 0
            ET03 = "01" 'Modify By Sindy 2012/1/13
            'End If  '93.8.10 cancel by sonia
        End If
      End If
   '2009/1/12 ADD BY SONIA 其他案件性質只印案件回覆單
   Else
'      NowPrint m_CP09, "04", "00", False, strUserNum, 0
      'modify by sonia 2018/9/28 大->台案件不印案件回覆單
      'ET03 = "00" 'Modify By Sindy 2012/1/13
      If Left(PUB_GetAKindSalesNo(m_TM01, m_TM02, m_TM03, m_TM04), 3) <> "MCT" Then
         ET03 = "00"
      End If
      'end 2018/9/27
      
      BolPrintLetterDemand = False 'Add By Sindy 2012/4/16 列印帳款未結清案件資料
   '2009/1/12 END
   End If
   
   'Add By Sindy 2012/1/13
   If ET03 <> "" Then
      bolEmail = PUB_GetEMailFlag(m_TM01 & m_TM02 & m_TM03 & m_TM04, m_CP10 = "102", , bolPlusPaper)
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
   Else
      BolPrintLetterDemand = False 'Add By Sindy 2012/4/16
      
      'Add By Sindy 2021/1/5 沒有系統產出的定稿
      'Add By Sindy 2021/2/1 詢問有沒有客戶函
      If strLD18 <> "" Then
         Call PUB_TCaseAskIsPost_C(strLD18)
      End If
      '2021/1/5 EMD
   End If
   '2012/1/13 End
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
If Me.textTM15.Enabled = True Then
   Cancel = False
   textTM15_Validate Cancel
   If Cancel = True Then
      textTM15.SetFocus
      Exit Function
   End If
End If

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

If Me.textCP35.Enabled = True Then
   Cancel = False
   textCP35_Validate Cancel
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

If Me.textCP49.Enabled = True Then
   Cancel = False
   textCP49_Validate Cancel
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

If Me.textTM14.Enabled = True Then
   Cancel = False
   textTM14_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add by Amy 2025/01/17 輸完直接按Enter鍵不會檢查
If Me.textCP37_1.Enabled = True Then
   Cancel = False
   textCP37_1_Validate Cancel
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

'Modify By Cheng 2002/07/22
'If Me.textTM16S.Enabled = True Then
'   Cancel = False
'   textTM16S_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
'If Me.textTM17.Enabled = True Then
'   Cancel = False
'   textTM17_Validate Cancel
'   If Cancel = True Then
'      Exit Function
'   End If
'End If
    'Add By Cheng 2002/11/18
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
       '     Exit Function
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
            'If m_DocNo = "" Or textCP07 <> "" Then 'Added by Morgan 2017/4/17 電子公文
            If m_DocNo = "" And textCP07 <> "" Then 'Added by Morgan 2017/4/17 電子公文
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

'Add By Sindy 2012/4/17
Private Sub Option1_Click(Index As Integer)
   If m_TM10 = 台灣國家代號 Then 'Addecd by Lydia 2019/06/21
        If Me.Option4(0).Value Then
           Text10_Validate False
        ElseIf Me.Option4(1).Value Then
           Text11_Validate False
        ElseIf Me.Option4(2).Value Then
           Text12_Validate False
        End If
   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
   Else
        Call GetTime
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
   'Dim strFromDate As String '期限起算日 'Remove by Lydia 2019/06/21
   
   'Add By Sindy 2012/8/30
   If Option4(0).Value = False And Option4(1).Value = False Then Exit Sub
   '2012/8/30 End
   
   'strFromDate = DBDATE(textCP05)
   'strFromDate = DBDATE(frm02010402_1.textCP05) 'Remove by Lydia 2019/06/21
   
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

   'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
   ElseIf frm02010402_3.GetSelectResult = "1" Then
      textCP06.Tag = ""
      textCP07.Tag = ""
      If Option1(0).Value = True Then ' 紙本公文不可大於15日曆天
           i = 15
      Else  '電子公文不可大於30日曆天
           i = 30
      End If
      strExc(1) = CompDate(2, i, strFromDate)
      If strExc(1) <> "" Then
         strExc(2) = CompWorkDay(4, strExc(1), 1)
         If strExc(2) < strSrvDate(1) Then
             strExc(2) = strSrvDate(1)
         End If
      End If
      textCP06.Text = TransDate(strExc(2), 1)
      textCP06.Tag = textCP06.Text
      textCP07.Text = TransDate(strExc(1), 1)
      textCP07.Tag = textCP07.Text
   End If
End Sub

'讀取來函期限
Private Function ChgType() As Boolean
Dim strTempName As String, bolTmp As Boolean
Dim i As Integer
'Remove by Lydia 2019/06/21
'Dim strFromDate As String '期限起算日
   
'   'strFromDate = DBDATE(textCP05)
'   strFromDate = DBDATE(frm02010402_1.textCP05)
'end 2019/06/21

   ChgType = False
   If m_TM10 = 台灣國家代號 Then
      bolTmp = False
   Else
      bolTmp = True
   End If
   
   ' 案件性質
   strRvType = "1002"
   Select Case frm02010402_3.GetSelectResult
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
                     textCP06.Text = TransDate(PUB_GetWorkDay1(textCP06.Text, True), 1) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
                  End If
               End If
            End With
         End If
      'Added by Lydia 2019/06/21 台-大核駁案期限管制: 先預設法限和所限，可人工變更；所限=法限-3個工作天。
      Else
           Call GetTime
      End If
      ChgType = True
   End If
End Function

