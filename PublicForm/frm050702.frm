VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050702 
   BorderStyle     =   1  '單線固定
   Caption         =   "服務業務基本資料維護"
   ClientHeight    =   5850
   ClientLeft      =   130
   ClientTop       =   970
   ClientWidth     =   9150
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9150
   Begin TabDlg.SSTab SSTab1 
      Height          =   4845
      Left            =   120
      TabIndex        =   41
      Top             =   960
      Width           =   8895
      _ExtentX        =   15681
      _ExtentY        =   8555
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "其他業務"
      TabPicture(0)   =   "frm050702.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "textSP85"
      Tab(0).Control(1)=   "Combo2"
      Tab(0).Control(2)=   "cboContact"
      Tab(0).Control(3)=   "Text3"
      Tab(0).Control(4)=   "Text1(3)"
      Tab(0).Control(5)=   "Text1(2)"
      Tab(0).Control(6)=   "Text1(1)"
      Tab(0).Control(7)=   "Text1(0)"
      Tab(0).Control(8)=   "Text1(4)"
      Tab(0).Control(9)=   "Text1(5)"
      Tab(0).Control(10)=   "Text1(6)"
      Tab(0).Control(11)=   "Text1(7)"
      Tab(0).Control(12)=   "Text1(8)"
      Tab(0).Control(13)=   "Text1(9)"
      Tab(0).Control(14)=   "Text1(10)"
      Tab(0).Control(15)=   "Text1(11)"
      Tab(0).Control(16)=   "Text1(12)"
      Tab(0).Control(17)=   "Text1(13)"
      Tab(0).Control(18)=   "Text1(14)"
      Tab(0).Control(19)=   "Text1(16)"
      Tab(0).Control(20)=   "Label1(117)"
      Tab(0).Control(21)=   "Label1(172)"
      Tab(0).Control(22)=   "Label1(15)"
      Tab(0).Control(23)=   "Label1(16)"
      Tab(0).Control(24)=   "Label1(11)"
      Tab(0).Control(25)=   "Label1(10)"
      Tab(0).Control(26)=   "Label1(9)"
      Tab(0).Control(27)=   "Label1(8)"
      Tab(0).Control(28)=   "Label12"
      Tab(0).Control(29)=   "Label1(14)"
      Tab(0).Control(30)=   "Labeld1(3)"
      Tab(0).Control(31)=   "Label1(7)"
      Tab(0).Control(32)=   "Label1(13)"
      Tab(0).Control(33)=   "Labeld1(0)"
      Tab(0).Control(34)=   "Label1(5)"
      Tab(0).Control(35)=   "Label1(6)"
      Tab(0).Control(36)=   "Label1(4)"
      Tab(0).Control(37)=   "Label1(3)"
      Tab(0).Control(38)=   "Label1(2)"
      Tab(0).Control(39)=   "Labeld1(2)"
      Tab(0).Control(40)=   "Label1(1)"
      Tab(0).Control(41)=   "Label1(0)"
      Tab(0).ControlCount=   42
      TabCaption(1)   =   "代理人相關資料"
      TabPicture(1)   =   "frm050702.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Combo5"
      Tab(1).Control(1)=   "Combo4"
      Tab(1).Control(2)=   "txtSP(84)"
      Tab(1).Control(3)=   "Text1(26)"
      Tab(1).Control(4)=   "Text1(25)"
      Tab(1).Control(5)=   "Text1(20)"
      Tab(1).Control(6)=   "Text1(24)"
      Tab(1).Control(7)=   "Text2"
      Tab(1).Control(8)=   "Text1(17)"
      Tab(1).Control(9)=   "Text1(18)"
      Tab(1).Control(10)=   "Text1(19)"
      Tab(1).Control(11)=   "Text1(21)"
      Tab(1).Control(12)=   "Text1(22)"
      Tab(1).Control(13)=   "Text1(23)"
      Tab(1).Control(14)=   "Label49"
      Tab(1).Control(15)=   "Label11(0)"
      Tab(1).Control(16)=   "Label2"
      Tab(1).Control(17)=   "Label10"
      Tab(1).Control(18)=   "Label5"
      Tab(1).Control(19)=   "Labeld1(13)"
      Tab(1).Control(20)=   "Labeld1(6)"
      Tab(1).Control(21)=   "Labeld1(5)"
      Tab(1).Control(22)=   "Labeld1(4)"
      Tab(1).Control(23)=   "Label16"
      Tab(1).Control(24)=   "Label18"
      Tab(1).Control(25)=   "Label19"
      Tab(1).Control(26)=   "Label20"
      Tab(1).Control(27)=   "Label21"
      Tab(1).Control(28)=   "Label22"
      Tab(1).Control(29)=   "Label24"
      Tab(1).Control(30)=   "Label25"
      Tab(1).Control(31)=   "Label26"
      Tab(1).Control(32)=   "Label28"
      Tab(1).Control(33)=   "Label29"
      Tab(1).ControlCount=   34
      TabCaption(2)   =   "銷卷資料"
      TabPicture(2)   =   "frm050702.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lblSP(70)"
      Tab(2).Control(1)=   "lblSP(69)"
      Tab(2).Control(2)=   "lblSP(68)"
      Tab(2).Control(3)=   "lblSP(61)"
      Tab(2).Control(4)=   "Label81"
      Tab(2).Control(5)=   "Label80(0)"
      Tab(2).Control(6)=   "Label79"
      Tab(2).Control(7)=   "Label78"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "其他"
      TabPicture(3)   =   "frm050702.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label15"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label1(166)"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(165)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label1(164)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "txtSP(80)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "txtSP(83)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "txtSP(82)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "txtSP(81)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Frame1K"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).ControlCount=   9
      TabCaption(4)   =   "參考備註"
      TabPicture(4)   =   "frm050702.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "cmdIns"
      Tab(4).Control(1)=   "Text1(15)"
      Tab(4).ControlCount=   2
      Begin VB.Frame Frame1K 
         Height          =   280
         Left            =   30
         TabIndex        =   100
         Top             =   1650
         Width           =   4930
         Begin VB.CheckBox Chk1K 
            Caption         =   "月帳單"
            Height          =   180
            Index           =   2
            Left            =   3840
            TabIndex        =   40
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "上傳平台"
            Height          =   180
            Index           =   1
            Left            =   2790
            TabIndex        =   39
            Top             =   60
            Width           =   1030
         End
         Begin VB.CheckBox Chk1K 
            Caption         =   "帳單另寄"
            Height          =   180
            Index           =   0
            Left            =   1740
            TabIndex        =   38
            Top             =   60
            Width           =   1030
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            Caption         =   "請款單寄送類型："
            Height          =   180
            Index           =   26
            Left            =   150
            TabIndex        =   101
            Top             =   60
            Width           =   1440
         End
      End
      Begin VB.CommandButton cmdIns 
         Caption         =   "各項指示"
         Height          =   300
         Left            =   -74880
         TabIndex        =   42
         Top             =   420
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Combo5 
         Height          =   260
         ItemData        =   "frm050702.frx":008C
         Left            =   -70350
         List            =   "frm050702.frx":009F
         Style           =   2  '單純下拉式
         TabIndex        =   32
         Top             =   3090
         Width           =   1470
      End
      Begin VB.ComboBox Combo4 
         Height          =   260
         ItemData        =   "frm050702.frx":00D3
         Left            =   -73560
         List            =   "frm050702.frx":00D5
         Style           =   2  '單純下拉式
         TabIndex        =   31
         Top             =   3090
         Width           =   990
      End
      Begin VB.TextBox textSP85 
         Height          =   264
         Left            =   -73560
         MaxLength       =   1
         TabIndex        =   19
         Top             =   3990
         Width           =   315
      End
      Begin VB.TextBox txtSP 
         Height          =   270
         Index           =   84
         Left            =   -68985
         MaxLength       =   20
         TabIndex        =   22
         Top             =   732
         Width           =   2655
      End
      Begin VB.TextBox txtSP 
         Height          =   270
         Index           =   81
         Left            =   1670
         MaxLength       =   1
         TabIndex        =   35
         Top             =   750
         Width           =   255
      End
      Begin VB.TextBox txtSP 
         Height          =   270
         Index           =   82
         Left            =   1670
         MaxLength       =   1
         TabIndex        =   36
         Top             =   1050
         Width           =   255
      End
      Begin VB.TextBox txtSP 
         Height          =   270
         Index           =   83
         Left            =   1670
         MaxLength       =   1
         TabIndex        =   37
         Top             =   1350
         Width           =   255
      End
      Begin VB.TextBox txtSP 
         Height          =   270
         Index           =   80
         Left            =   1670
         MaxLength       =   1
         TabIndex        =   34
         Top             =   420
         Width           =   255
      End
      Begin VB.ComboBox Combo2 
         Height          =   276
         ItemData        =   "frm050702.frx":00D7
         Left            =   -70305
         List            =   "frm050702.frx":00E7
         TabIndex        =   14
         Top             =   2532
         Width           =   2505
      End
      Begin MSForms.ComboBox cboContact 
         Height          =   315
         Left            =   -68040
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3090
         Width           =   1830
         VariousPropertyBits=   679495711
         DisplayStyle    =   7
         Size            =   "3228;556"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text3 
         Height          =   555
         Left            =   -73560
         TabIndex        =   18
         Top             =   3390
         Width           =   7335
         VariousPropertyBits=   -1466941409
         BackColor       =   -2147483633
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12938;979"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   3975
         Index           =   15
         Left            =   -74940
         TabIndex        =   44
         Top             =   780
         Width           =   8745
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "15425;7011"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   -73560
         TabIndex        =   24
         Top             =   1320
         Width           =   7245
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12779;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   -73260
         TabIndex        =   30
         Top             =   2805
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   -73560
         TabIndex        =   28
         Top             =   2190
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   -73560
         TabIndex        =   29
         Top             =   2505
         Width           =   2655
         VariousPropertyBits=   671105051
         MaxLength       =   35
         Size            =   "4683;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   -72120
         TabIndex        =   3
         Top             =   435
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   -72360
         TabIndex        =   2
         Top             =   435
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "450;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -73080
         TabIndex        =   1
         Top             =   435
         Width           =   705
         VariousPropertyBits=   671105051
         MaxLength       =   6
         Size            =   "1244;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text2 
         Height          =   1365
         Left            =   -73560
         TabIndex        =   33
         Top             =   3390
         Width           =   7335
         VariousPropertyBits=   -1466941409
         BackColor       =   -2147483633
         MaxLength       =   2000
         ScrollBars      =   2
         Size            =   "12938;2408"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -73560
         TabIndex        =   0
         Top             =   435
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   -73560
         TabIndex        =   5
         Top             =   735
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12938;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   -73560
         TabIndex        =   6
         Top             =   1035
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   180
         Size            =   "12938;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   -73560
         TabIndex        =   7
         Top             =   1335
         Width           =   7335
         VariousPropertyBits=   671105051
         MaxLength       =   160
         Size            =   "12938;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   7
         Left            =   -73560
         TabIndex        =   8
         Top             =   1635
         Width           =   975
         VariousPropertyBits=   671105051
         MaxLength       =   9
         Size            =   "1720;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   -70440
         TabIndex        =   4
         Top             =   435
         Width           =   495
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "873;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -73560
         TabIndex        =   9
         Top             =   1935
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1508;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -70440
         TabIndex        =   10
         Top             =   1935
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   -73560
         TabIndex        =   11
         Top             =   2235
         Width           =   855
         VariousPropertyBits=   671105051
         MaxLength       =   7
         Size            =   "1508;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -70440
         TabIndex        =   12
         Top             =   2235
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   2
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -73560
         TabIndex        =   13
         Top             =   2535
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -73560
         TabIndex        =   15
         Top             =   2805
         Width           =   2670
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4710;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -73560
         TabIndex        =   16
         Top             =   3090
         Width           =   4575
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "8070;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   17
         Left            =   -73560
         TabIndex        =   20
         Top             =   420
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   -73560
         TabIndex        =   23
         Top             =   1020
         Width           =   7245
         VariousPropertyBits=   671105051
         MaxLength       =   60
         Size            =   "12779;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   -73560
         TabIndex        =   25
         Top             =   1605
         Width           =   1095
         VariousPropertyBits=   671105051
         MaxLength       =   8
         Size            =   "1931;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   -73560
         TabIndex        =   21
         Top             =   735
         Width           =   2655
         VariousPropertyBits=   671105051
         MaxLength       =   50
         Size            =   "4683;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   22
         Left            =   -73560
         TabIndex        =   26
         Top             =   1905
         Width           =   375
         VariousPropertyBits=   671105051
         MaxLength       =   3
         Size            =   "661;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   -69180
         TabIndex        =   27
         Top             =   1890
         Width           =   255
         VariousPropertyBits=   671105051
         MaxLength       =   1
         Size            =   "450;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label49 
         AutoSize        =   -1  'True
         Caption         =   "請款單列印幣別格式："
         Height          =   180
         Left            =   -72180
         TabIndex        =   98
         Top             =   3120
         Width           =   1800
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "請款幣別："
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   97
         Top             =   3120
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "特殊出名公司 :         ( J:智權公司 空白:系統預設)"
         Height          =   180
         Index           =   117
         Left            =   -74805
         TabIndex        =   96
         Top             =   3990
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "與他案合併計算結餘請於案件備註欄註明""與某案號合併計算結餘""！"
         ForeColor       =   &H000000FF&
         Height          =   720
         Index           =   172
         Left            =   -67650
         TabIndex        =   95
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CLIENT_MATTER_ID:"
         Height          =   180
         Left            =   -70725
         TabIndex        =   94
         Top             =   777
         Width           =   1725
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "定稿份數:"
         Height          =   180
         Index           =   164
         Left            =   180
         TabIndex        =   93
         Top             =   780
         Width           =   770
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "請款單份數:"
         Height          =   180
         Index           =   165
         Left            =   180
         TabIndex        =   92
         Top             =   1080
         Width           =   950
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Email同時寄紙本:          (Y:是)"
         Height          =   180
         Index           =   166
         Left            =   180
         TabIndex        =   91
         Top             =   1380
         Width           =   2270
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "是否以Email通知:          (Y:是   D:僅D/N）"
         Height          =   180
         Left            =   180
         TabIndex        =   90
         Top             =   480
         Width           =   3150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "工程師組別:"
         Height          =   180
         Index           =   15
         Left            =   -71310
         TabIndex        =   89
         Top             =   2535
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "接洽人:"
         Height          =   180
         Index           =   16
         Left            =   -68700
         TabIndex        =   88
         Top             =   3150
         Width           =   585
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "聯絡人部門:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   87
         Top             =   1320
         Width           =   945
      End
      Begin MSForms.Label lblSP 
         Height          =   285
         Index           =   70
         Left            =   -73650
         TabIndex        =   86
         Top             =   1350
         Width           =   7425
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "13097;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP 
         Height          =   285
         Index           =   69
         Left            =   -73830
         TabIndex        =   85
         Top             =   1050
         Width           =   2025
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3572;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP 
         Height          =   285
         Index           =   68
         Left            =   -73830
         TabIndex        =   84
         Top             =   750
         Width           =   2175
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3836;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblSP 
         Height          =   285
         Index           =   61
         Left            =   -73830
         TabIndex        =   83
         Top             =   450
         Width           =   2295
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "4048;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label81 
         AutoSize        =   -1  'True
         Caption         =   "北所銷卷日："
         Height          =   180
         Left            =   -74940
         TabIndex        =   82
         Top             =   450
         Width           =   1080
      End
      Begin VB.Label Label80 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷日："
         Height          =   180
         Index           =   0
         Left            =   -74940
         TabIndex        =   81
         Top             =   750
         Width           =   1080
      End
      Begin VB.Label Label79 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷員："
         Height          =   180
         Left            =   -74940
         TabIndex        =   80
         Top             =   1050
         Width           =   1080
      End
      Begin VB.Label Label78 
         AutoSize        =   -1  'True
         Caption         =   "分所銷卷備註："
         Height          =   180
         Left            =   -74940
         TabIndex        =   79
         Top             =   1350
         Width           =   1260
      End
      Begin VB.Label Label5 
         Caption         =   "D/N固定列印對象:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   78
         Top             =   2805
         Width           =   1545
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   13
         Left            =   -72150
         TabIndex        =   77
         Top             =   2820
         Width           =   5175
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   6
         Left            =   -72450
         TabIndex        =   75
         Top             =   2190
         Width           =   5295
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9340;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   5
         Left            =   -72450
         TabIndex        =   74
         Top             =   1590
         Width           =   5295
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9340;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   4
         Left            =   -72450
         TabIndex        =   73
         Top             =   420
         Width           =   5295
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "9340;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label16 
         Caption         =   "FC代理人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   72
         Top             =   420
         Width           =   852
      End
      Begin VB.Label Label18 
         Caption         =   "彼所案號:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   71
         Top             =   732
         Width           =   852
      End
      Begin VB.Label Label19 
         Caption         =   "聯絡人:"
         Height          =   252
         Left            =   -74760
         TabIndex        =   70
         Top             =   1032
         Width           =   732
      End
      Begin VB.Label Label20 
         Caption         =   "折扣:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   69
         Top             =   1905
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "%"
         Height          =   255
         Left            =   -73080
         TabIndex        =   68
         Top             =   1905
         Width           =   375
      End
      Begin VB.Label Label22 
         Caption         =   "固定請款對象:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   67
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label24 
         Caption         =   "D/N否列印申請人:"
         Height          =   255
         Left            =   -70680
         TabIndex        =   66
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label Label25 
         Caption         =   "(Y:印)"
         Height          =   255
         Left            =   -68820
         TabIndex        =   65
         Top             =   1890
         Width           =   615
      End
      Begin VB.Label Label26 
         Caption         =   "副本收受人:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   64
         Top             =   2205
         Width           =   1095
      End
      Begin VB.Label Label28 
         Caption         =   "副本聯絡人:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   63
         Top             =   2505
         Width           =   975
      End
      Begin VB.Label Label29 
         Caption         =   "代理人備註:"
         Height          =   255
         Left            =   -74760
         TabIndex        =   62
         Top             =   3390
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "客戶備註:"
         Height          =   255
         Index           =   11
         Left            =   -74805
         TabIndex        =   61
         Top             =   3390
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "客戶案件案號:"
         Height          =   255
         Index           =   10
         Left            =   -74805
         TabIndex        =   60
         Top             =   3090
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "分所案號:"
         Height          =   255
         Index           =   9
         Left            =   -74805
         TabIndex        =   59
         Top             =   2805
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "定稿語文:"
         Height          =   255
         Index           =   8
         Left            =   -74805
         TabIndex        =   58
         Top             =   2535
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "(1.中文 2英文 3.日文)"
         Height          =   252
         Left            =   -73200
         TabIndex        =   57
         Top             =   2532
         Width           =   1812
      End
      Begin VB.Label Label1 
         Caption         =   "閉卷原因:"
         Height          =   255
         Index           =   14
         Left            =   -71280
         TabIndex        =   56
         Top             =   2235
         Width           =   855
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   3
         Left            =   -70035
         TabIndex        =   55
         Top             =   2220
         Width           =   2295
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "閉卷日期:"
         Height          =   255
         Index           =   7
         Left            =   -74805
         TabIndex        =   54
         Top             =   2235
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "是否閉卷:             (Y:閉卷)"
         Height          =   255
         Index           =   13
         Left            =   -71280
         TabIndex        =   53
         Top             =   1920
         Width           =   3015
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   0
         Left            =   -72540
         TabIndex        =   52
         Top             =   1620
         Width           =   4815
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "5741;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請人:"
         Height          =   255
         Index           =   5
         Left            =   -74805
         TabIndex        =   51
         Top             =   1635
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "申請日:"
         Height          =   255
         Index           =   6
         Left            =   -74805
         TabIndex        =   50
         Top             =   1935
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱(日):"
         Height          =   255
         Index           =   4
         Left            =   -74805
         TabIndex        =   49
         Top             =   1335
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱(英):"
         Height          =   255
         Index           =   3
         Left            =   -74805
         TabIndex        =   48
         Top             =   1035
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "案件名稱(中):"
         Height          =   255
         Index           =   2
         Left            =   -74805
         TabIndex        =   47
         Top             =   720
         Width           =   1215
      End
      Begin MSForms.Label Labeld1 
         Height          =   285
         Index           =   2
         Left            =   -69900
         TabIndex        =   46
         Top             =   420
         Width           =   1875
         BackColor       =   16777215
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "3307;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "申請國家:"
         Height          =   255
         Index           =   1
         Left            =   -71280
         TabIndex        =   45
         Top             =   435
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "本所案號:"
         Height          =   255
         Index           =   0
         Left            =   -74805
         TabIndex        =   43
         Top             =   435
         Width           =   855
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8340
      Top             =   0
      _ExtentX        =   988
      _ExtentY        =   988
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":011C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":0438
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":0754
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":0930
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":0C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":0F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":1284
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":15A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":18BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":1BD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm050702.frx":1EF4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbar 
      Align           =   1  '對齊表單上方
      Height          =   520
      Left            =   0
      TabIndex        =   76
      Top             =   0
      Width           =   9150
      _ExtentX        =   16140
      _ExtentY        =   917
      ButtonWidth     =   1076
      ButtonHeight    =   882
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      HotImageList    =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "新增"
            Key             =   "keyInsert"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "修改"
            Key             =   "keyUpdate"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "刪除"
            Key             =   "keyDelete"
            ImageIndex      =   3
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "查詢"
            Key             =   "keyQuery"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "第一筆"
            Key             =   "keyFirst"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "前一筆"
            Key             =   "keyPrevious"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "後一筆"
            Key             =   "keyNext"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "最後筆"
            Key             =   "keyLast"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "確定"
            Key             =   "keyOk"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "取消"
            Key             =   "keyCancel"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   1
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "結束"
            Key             =   "keyExit"
            ImageIndex      =   11
         EndProperty
      EndProperty
   End
   Begin MSForms.TextBox textCUID 
      Height          =   285
      Left            =   2940
      TabIndex        =   99
      TabStop         =   0   'False
      Top             =   660
      Width           =   6075
      VariousPropertyBits=   16415
      BackColor       =   16777215
      Size            =   "5741;503"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
End
Attribute VB_Name = "frm050702"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/07 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim Data_Mission As Integer  'Memo by Lydia 2020/05/05 1-新增,2-刪除,3-修改,4-查詢
Dim Rs As ADODB.Recordset
'Modify by Morgan 2006/10/18
'Dim sp(25) As String
Dim sp(26) As String
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj0701 As Object
'Dim obj0702 As Object
Dim nRet As Boolean
Dim strNo As String
Dim strName As String
Dim strTemp As String
Dim ChkKey As Boolean
Dim Fld1 As String
Dim Fld2 As String
Dim Fld3 As String
Dim Fld4 As String
Dim Fld5 As String
Dim Fld6 As String
Dim Fld7 As String
Dim Fld8 As String
Dim DelFlg As Boolean
Dim RsCounts As Integer
Dim InitValue As Boolean
Dim GetNowData As Boolean
Dim ChkData As Boolean
Dim intWhere As Integer
Dim intSysnum As Integer
Dim strKind As String
Dim BlnULetter As Boolean
Dim blnKeypreview As Boolean

' 90.07.31 modify by louis
Dim m_SysKind As String

' 90.07.13 modify by louis (執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim strChkCuAreaMail As String, strChkCuAreaMailTo As String 'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員


'Added by Lydia 2016/11/24 各項指示
Private Sub cmdIns_Click()
   If Me.Text1(0).Text = "" Or Me.Text1(1).Text = "" Then
      MsgBox "請輸入本所案號", vbInformation
      Exit Sub
   End If
   'Added by Lydia 2020/05/05
   If Data_Mission <> 4 Then
      MsgBox IIf(Data_Mission = 1, "新增中", IIf(Data_Mission = 2, "刪除中", "修改中")) & "不可執行！", vbInformation
      Exit Sub
   End If
   'end 2020/05/05
   'Added by Lydia 2020/05/05 各項指示：檢查表單是否開啟中
   If PUB_CheckFormExist("frm12040159") Then
       MsgBox "請先關閉〔申請人/代理人/案件各項指示資料〕的畫面！", vbInformation
       Exit Sub
   End If
   'end 2020/05/05
   
   frm12040159.SetParent "E", Trim(Me.Text1(0).Text & Me.Text1(1).Text & Me.Text1(2).Text & Me.Text1(3).Text), Me
   frm12040159.Show

End Sub

'2010/1/8 ADD BY SONIA
Private Sub Combo2_Validate(Cancel As Boolean)
   If Combo2 <> "" Then
      Combo2 = Left(Combo2, 1) + "." + PUB_GetFCPGrpName(Left(Combo2, 1))
      If Combo2 = Left(Combo2, 1) + "." Then
         Combo2 = Left(Combo2, 1)
         Cancel = True
         Combo2.SetFocus
      End If
   End If
End Sub
'2010/1/8 end

'Add By Sindy 2016/11/23
Private Sub Combo4_Click()
   Call GetCurrType
End Sub
Private Sub Combo4_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub Combo4_Validate(Cancel As Boolean)
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   If ExistCheck("acc1y0", "a1y01", Combo4, Label11(0)) = False Then
      Cancel = True
      Combo4.SetFocus
   End If
   If Combo4 <> "USD" Then
      If ExistCheck("DebitNoteRate", "DNR01", Combo4, Label11(0) & "匯率") = False Then
         Cancel = True
         Combo4.SetFocus
         Exit Sub
      End If
   End If
   Call GetCurrType
End Sub
Private Sub GetCurrType()
Dim intType As Integer
   
   If Combo4 = MsgText(601) Then
      Combo4.Tag = Combo4.Text
      Combo5.ListIndex = 0
      Combo5.Enabled = False
      Exit Sub
   End If
   '若更改請款幣別
   If Me.Combo4.Text <> Me.Combo4.Tag Then
      Me.Combo4.Tag = Me.Combo4.Text
      '請款幣別變更要重新預設列印幣別
      '台幣
      If Me.Combo4.Text = "NTD" Then
         intType = 1 '純台幣
      '人民幣
      ElseIf Me.Combo4.Text = "RMB" Then
         intType = 4 '外幣+美金合計
      '其他幣別
      Else
         intType = 2 '台幣+外幣合計
      End If
      Combo5.ListIndex = intType
      '若為台幣時則格式欄位鎖住不可修改
      If Me.Combo4.Text = "NTD" Then
         Combo5.Enabled = False
      Else
         Combo5.Enabled = True
      End If
   End If
End Sub
'2016/11/23 END

'add by nickc 2006/11/13 Enter 事件，等於存檔，做完取消，不然 form 內其他物件有寫 keycode 或是 keyascii 事件的話，也會做到
Private Sub Form_KeyPress(KeyAscii As Integer)
   'Add By Sindy 2014/8/29 當focus在備註欄時按enter鍵維持換行功能而不是存檔功能
   If KeyAscii = 13 And UCase(Me.ActiveControl.Name) = UCase("Text1") Then
      If Me.ActiveControl.Index = 15 Then
         Exit Sub
      End If
   End If
   '2014/8/29 END
   Select Case KeyAscii
      Case 13:
         If Data_Mission <> 0 Then
            KeyAscii = 0
            UseDatamaintain (vbKeyF9)
         End If
      Case Else
         If BlnULetter Then KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub Form_Load()
   'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
   FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
   
   ' 90.07.13 modify by louis (取得使用者執行各項功能的權限)
   m_bInsert = IsUserHasRightOfFunction("frm050702", strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction("frm050702", strEdit, False)
   m_bDelete = IsUserHasRightOfFunction("frm050702", strDel, False)
   m_bQuery = IsUserHasRightOfFunction("frm050702", strFind, False)
   
   textCUID.BackColor = &H8000000F
   
   ' 90.07.31 modify by louis
   m_SysKind = strSysKind
   'Add By Cheng 2003/04/11
   '預設系統類別
   Me.Text1(0).Text = strSysKind
   MoveFormToCenter Me
   OpenTable
   'Add By Cheng 2003/04/11
   '若有資料先記錄其中一筆
   If Rs.RecordCount > 0 Then
       Fld1 = "" & Rs.Fields(0).Value
       Fld2 = "" & Rs.Fields(1).Value
       Fld3 = "" & Rs.Fields(2).Value
       Fld4 = "" & Rs.Fields(3).Value
   End If
   InitValue = True
   ChkData = True
   'ShowData
   GetSysInf
   DelFlg = False
   InitValue = False
   OnOffTxt False
   'Me.Text2.Enabled = False
   Me.SSTab1.Tab = 0
   blnKeypreview = True
    
   ' 90.07.13 modify by louis (更新按紐的狀態)
   UpdateToolbarButtonState
   
   'Add By Cheng 2002/01/04
   SetQueryStatus
   
   SetBackColor 'Add by Morgan 2009/9/15
   
   'Add By Sindy 2016/11/23
   '抓有輸入過匯率的請款幣別
   Combo4.Clear
   Combo4.AddItem ""
   Combo4.AddItem "USD"
   If RsTemp.State <> adStateClosed Then RsTemp.Close
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open "select distinct DNR01 from DebitNoteRate order by DNR01 asc", adoTaie, adOpenStatic, adLockReadOnly
   Do While RsTemp.EOF = False
      Combo4.AddItem RsTemp.Fields("DNR01").Value
      RsTemp.MoveNext
   Loop
   RsTemp.Close
   '2016/11/23 End
   
   'Added by Lydia 2020/05/05 各項指示：顯示按鈕
   If strSrvDate(1) >= 各項指示啟用日 Then
      cmdIns.Visible = True
   Else
      cmdIns.Visible = False
      Text1(15).Top = 390
      Text1(15).Height = 4185
   End If
   
   Frame1K.BorderStyle = 0 'Add By Sindy 2025/1/7
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm050702 = Nothing
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As ReturnInteger)
   If Data_Mission = 4 And (Index = 1 Or Index = 2 Or Index = 3) And KeyAscii = 13 Then UseDatamaintain vbKeyF9
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Select Case Index
    Case 3
      If Data_Mission = 1 Then
         If Not ChkCaseCode("SP", Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(3).Text) Then
            Text1(1).SetFocus
         End If
        'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        'Modified by Lydia 2017/06/22 +系統別
        If FMP2open = True And Text1(0) = "PS" Then
          If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(0).Text, Text1(1).Text, Text1(2).Text, Text1(3).Text) = False Then Text1(1).SetFocus
        End If
        
      End If
    Case 7, 10, 17, 19, 20, 23, 0, 25
        BlnULetter = False
    End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
Dim i As Integer
'edit by nickc 2007/02/06 不用 dll 了 Dim obj01 As Object
Dim strA As String
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj0702 As Object
Dim strTmp As String
   If Data_Mission <> 1 And Data_Mission <> 3 Then Exit Sub
   Select Case Index
      Case 0
          If Text1(0) = "" Then Exit Sub
          strKind = Trim$(Text1(0).Text)
          'edit by nickc 2007/02/02 不用 dll 了
          'If Not objPublicData.GetSystemKind(strKind, intSysnum, , intWhere) Then
          If Not ClsPDGetSystemKind(strKind, intSysnum, , intWhere) Then
             Cancel = True
             Exit Sub
          Else
             If intSysnum < 5 Then
                ShowMsg MsgText(9041)
                Cancel = True
                Exit Sub
             End If
          End If
      Case 1
          If Text1(1) = "" Then Exit Sub
          If Len(Text1(1)) <> 6 Then
              ShowMsg MsgText(9042)
              Cancel = True
          End If
      Case 3
          If Text1(3) = "" Then Exit Sub
          If Len(Text1(3)) <> 2 Then
              ShowMsg MsgText(9043)
              Cancel = True
          End If
          
      Case 4, 5, 6
          'edit by nickc 2007/06/06 切換輸入法改用API
          'Text1(Index).IMEMode = 2
          CloseIme
          'Add by Morgan 2007/4/30
          If CheckLengthIsOK(Text1(Index), Text1(Index).MaxLength) = False Then
            Text1(Index).SetFocus
            Cancel = True
          End If
          'end 2007/4/30
          
      Case 7
          If Text1(7) = "" Then Labeld1(0).Caption = "": Exit Sub
          strTmp = Text1(7).Text
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.GETCUSTOMER(strTmp, strA) Then
          If ClsPDGetCustomer(strTmp, strA) Then
             Labeld1(0).Caption = strA
             Text1(7).Text = strTmp
          Else
             Labeld1(0).Caption = ""
             Cancel = True
          End If
          Text3 = GetMemo(Text1(7).Text)
      Case 8
          If Text1(8) = "" Then Labeld1(2).Caption = "": Exit Sub
          'edit by nickc 2007/02/02 不用 dll 了
          'If objPublicData.GetNation(Text1(8), strA) Then
          If ClsPDGetNation(Text1(8), strA) Then
             Labeld1(2).Caption = strA
          Else
             Labeld1(2).Caption = ""
             Cancel = True
          End If
      Case 9, 11
          If Text1(Index) = "" Then Exit Sub
          Cancel = Not ChkDate(Text1(Index))
      Case 10, 23
          If Text1(Index) = "" Then Exit Sub
          If Text1(Index).Text <> "Y" Then
              Cancel = True
              ShowMsg MsgText(9034)
          End If
      
      Case 12
          If Text1(12) = "" Then Labeld1(3) = "": Exit Sub
          'edit by nickc 2007/02/05 不用 dll 了
          'Set obj0702 = CreateObject("prjtaiedll.class0702")
          'If obj0702.GetReasonOfRelief(Text1(12), strA) Then
          If Cls0702GetReasonOfRelief(Text1(12), strA) Then
             Labeld1(3).Caption = strA
          Else
             Labeld1(3).Caption = ""
             Cancel = True
          End If
          'edit by nickc 2007/02/06 不用 dll 了
          'Set obj0702 = Nothing
      Case 13
          If Text1(13) = "" Then Exit Sub
          Select Case Text1(13)
          Case "1", "2", "3"
          Case Else
            ShowMsg MsgText(9036)
            Cancel = True
          End Select
      'add by nickc 2005/10/06
      Case 14
            If CheckLengthIsOK(Text1(14), 50) = False Then
               Cancel = True
               Text1(14).SetFocus
               TextInverse Text1(14)
               Exit Sub
            End If
            
      Case 15
          'edit by nickc 2007/06/06 切換輸入法改用API
          'Text1(Index).IMEMode = 2
          CloseIme
      Case 17
         Labeld1(4).Caption = ""
         Text2.Text = ""
         If Text1(17) = "" Then Exit Sub
         strA = Text1(17)
         'Modify By Cheng 2002/07/09
'         If objPublicData.GetAgent(strA, strName) Then
         If PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
            Me.Labeld1(4).Caption = strName
            Text1(17).Text = strA
            Text2.Text = GetFa29(strA)
         Else
            Cancel = True
         End If
      'Add by Morgan 2006/10/18
      Case 18, 26
         If CheckLengthIsOK(Text1(Index), Text1(Index).MaxLength) = False Then
            Text1(Index).SetFocus
            Cancel = True
         End If
      
      Case 19
          If Text1(19) = "" Then Me.Labeld1(5).Caption = "": Exit Sub
          strA = Text1(19).Text
         'Modify By Cheng 2002/07/09
'          If objPublicData.GetAgent(strA, strName) Then
          'Modified by Morgan 2011/11/29 也要能輸入申請人
          'If PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
          If ClsLawLawGetName(strA, strName) Then
          'end 2011/11/29
             Me.Labeld1(5).Caption = strName
          Else
             Me.Labeld1(5).Caption = ""
             Cancel = True
          End If
      Case 20
          If Text1(20) = "" Then Me.Labeld1(6).Caption = "": Exit Sub
          strA = Text1(20).Text
         'Modify By Cheng 2002/07/09
'          If objPublicData.GetAgent(strA, strName) Then
          'Modified by Morgan 2011/11/29 也要能輸入申請人
          'If PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
          If ClsLawLawGetName(strA, strName) Then
          'end 2011/11/29
             Me.Labeld1(6).Caption = strName
          Else
             Me.Labeld1(6).Caption = ""
             Cancel = True
          End If
      Case 22
          If Text1(22) = "" Then Exit Sub
          If 0 > Val(Text1(22).Text) Then
             ShowMsg MsgText(9035)
             Cancel = True
          Else
             If 100 < Val(Text1(22).Text) Then
               ShowMsg MsgText(9035)
               Cancel = True
             End If
          End If
      Case 25
          If Text1(25) = "" Then Me.Labeld1(13).Caption = "": Exit Sub
          strA = Text1(25).Text
         'Modify By Cheng 2002/07/09
'          If objPublicData.GetAgent(strA, strName) Then
          'Modified by Morgan 2011/11/29 也要能輸入申請人
          'If PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
          If ClsLawLawGetName(strA, strName) Then
          'end 2011/11/29
             Me.Labeld1(13).Caption = strName
          Else
             Me.Labeld1(13).Caption = ""
             Cancel = True
          End If
   End Select
End Sub

Private Function GetFa29(ByVal strTmp As String) As String
   strTmp = ChangeCustomerL(strTmp)
   strExc(0) = "select fa29 from fagent where " & ChgFagent(strTmp)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 And Not IsNull(RsTemp.Fields(0)) Then
      GetFa29 = RsTemp.Fields(0)
   Else
      GetFa29 = ""
   End If
End Function

Private Sub Text1_GotFocus(Index As Integer)
    Select Case Index
    Case 7, 10, 23, 19, 20, 17, 0, 25
        BlnULetter = True
    Case 4, 6, 15
        'edit by nickc 2007/06/06 切換輸入法改用API
        'Text1(Index).IMEMode = 1
        OpenIme
    End Select
   TextInverse Text1(Index)
End Sub
Private Sub ShowData()
Dim i As Integer
   Rs.ReQuery
   RsCounts = 0
   Do Until Rs.EOF
       Rs.MoveNext
       RsCounts = RsCounts + 1
   Loop
   If RsCounts > 0 Then
     Rs.MoveFirst
   End If
    '若無資料
   If Rs.EOF Then
      Call Clear_AllTxtAry(Text1, 0, 3)
      tlbar.Buttons.Item(1).Enabled = True
      tlbar.Buttons.Item(2).Enabled = False
      tlbar.Buttons.Item(3).Enabled = False
      tlbar.Buttons.Item(4).Enabled = False
      tlbar.Buttons.Item(6).Enabled = False
      tlbar.Buttons.Item(7).Enabled = False
      tlbar.Buttons.Item(8).Enabled = False
      tlbar.Buttons.Item(9).Enabled = False
      tlbar.Buttons.Item(11).Enabled = False
      tlbar.Buttons.Item(12).Enabled = False
      tlbar.Buttons.Item(14).Enabled = True
      Exit Sub
   End If
'   tlbar.Buttons.Item(1).Enabled = True
'   tlbar.Buttons.Item(2).Enabled = True
'   tlbar.Buttons.Item(3).Enabled = True
'   tlbar.Buttons.Item(4).Enabled = True
'   tlbar.Buttons.Item(6).Enabled = True
'   tlbar.Buttons.Item(7).Enabled = True
'   tlbar.Buttons.Item(8).Enabled = True
'   tlbar.Buttons.Item(9).Enabled = True
'   tlbar.Buttons.Item(11).Enabled = False
'   tlbar.Buttons.Item(12).Enabled = False
'   tlbar.Buttons.Item(14).Enabled = True
    UpdateToolbarButtonState
   If RsCounts > 1 And Not InitValue Then
       If DelFlg Then
           QueryData "sp01=" + CNULL(Fld5), 4
       Else
           QueryData "sp01=" + CNULL(Fld1), 4
       End If
       Exit Sub
   Else
       ShowDetail
       Exit Sub
   End If
   
   ' 90.07.13 modify by louis (更新按紐的狀態)
   UpdateToolbarButtonState
End Sub

Private Function ChkInData() As Boolean
Dim i As Integer, j As Integer
Dim Rs As ADODB.Recordset
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj0702 As Object
Dim strA As String
Dim strTmp As String

   If Data_Mission = 1 Then
      For i = 0 To 1
         If Text1(i) = "" Then
            ChkInData = False
            ShowMsg MsgText(9014)
            Text1(i).SetFocus
            Exit Function
         End If
      Next
      strTmp = Text1(0)
      If InStr(1, m_SysKind, strTmp) <= 0 Then
         ChkInData = False
         ShowMsg MsgText(1107)
         Text1(0).SetFocus
         Exit Function
      End If
      'edit by nickc 2007/02/05 不用 dll 了
      'Set obj0702 = CreateObject("prjtaiedll.class0702")
      If Text1(2) = "" Then
         strExc(2) = "0"
      Else
         strExc(2) = Text1(2).Text
      End If
      If Text1(3) = "" Then
         strExc(3) = "00"
      Else
         strExc(3) = Text1(3).Text
      End If

      'edit by nickc 2007/02/05 不用 dll 了
      'If Not obj0702.ChkCaseNoWithAuNum0702(Trim$(Text1(0).Text), Trim$(Text1(1).Text), strExc(2), strExc(3), strKind) Then
      If Not Cls0702ChkCaseNoWithAuNum0702(Trim$(Text1(0).Text), Trim$(Text1(1).Text), strExc(2), strExc(3), strKind) Then
         ChkInData = False
         'edit by nickc 2007/02/05 不用 dll 了
         'Set obj0702 = Nothing
         Text1(0).SetFocus
         Exit Function
      End If
      'edit by nickc 2007/02/05 不用 dll 了
      'Set obj0702 = Nothing
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetSystemKind(Text1(0), intSysnum, , intWhere) Then
      If Not ClsPDGetSystemKind(Text1(0), intSysnum, , intWhere) Then
         ChkInData = False
         Text1(0).SetFocus
         Exit Function
      Else
         If intSysnum < 5 Then
            MsgBox "本系統類別非服務業務", vbCritical
            ChkInData = False
            Text1(0).SetFocus
            Exit Function
         End If
      End If
   End If
   
     'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
     '新增或修改都要判斷
     'Modifeid by Lydia 2017/06/22 +系統別
     If FMP2open = True And Text1(0) = "PS" Then
       If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1(0), Text1(1), strExc(3), strExc(4)) = False Then Exit Function
     End If

   If Text1(8) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If Not objPublicData.GetNation(Text1(8).Text, strA) Then
      If Not ClsPDGetNation(Text1(8).Text, strA) Then
         SSTab1.Tab = 0
         Text1(8).SetFocus
         ChkInData = False
         Exit Function
      End If
   Else
      ShowMsg "申請國家不可空白，請重新輸入 !"
      SSTab1.Tab = 0
      Text1(8).SetFocus
      ChkInData = False
      Exit Function
   End If
   
   If Text1(4).Text = "" And Text1(5).Text = "" And Text1(6).Text = "" Then
        ShowMsg "案件名稱不可同時空白，請重新輸入 !"
        SSTab1.Tab = 0
        Text1(4).SetFocus
        ChkInData = False
        Exit Function
   End If
        
   If Text1(7) <> "" Then
      strTmp = Text1(7).Text
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GETCUSTOMER(strTmp, strA) Then
      If ClsPDGetCustomer(strTmp, strA) Then
         Text1(7).Text = strTmp
         Labeld1(0).Caption = strA
      Else
         Labeld1(0).Caption = ""
         SSTab1.Tab = 0
         Text1(7).SetFocus
         ChkInData = False
         Exit Function
      End If
   Else
      'Modify by Morgan 2007/4/30 加判斷代理人
      If Text1(17) = "" Then
         'ShowMsg "申請人不可空白，請重新輸入 !"
         MsgBox "申請人和代理人不可同時空白 !", vbCritical
         SSTab1.Tab = 0
         Text1(7).SetFocus
         ChkInData = False
         Exit Function
      End If
      'end 2007/4/30
   End If
    
   If Text1(9) <> "" Then
      If Not ChkDate(Text1(9).Text) Then
         SSTab1.Tab = 0
         Text1(9).SetFocus
         ChkInData = False
         Exit Function
      End If
   End If
     
   If Text1(11) <> "" Then
      If Not ChkDate(Text1(11).Text) Then
         SSTab1.Tab = 0
         Text1(11).SetFocus
         ChkInData = False
         Exit Function
      End If
   End If
     
   If Text1(13) <> "" Then
      Select Case Text1(13)
         Case "1", "2", "3"
         
         Case Else
            SSTab1.Tab = 0
            Text1(13).SetFocus
            ShowMsg MsgText(9036)
            ChkInData = False
      End Select
   End If
    
   If Text1(12) <> "" Then
     'edit by nickc 2007/02/05 不用 dll 了
     'Set obj0702 = CreateObject("prjtaiedll.class0702")
       'If Not obj0702.GetReasonOfRelief(Text1(12), strA) Then
       If Not Cls0702GetReasonOfRelief(Text1(12), strA) Then
          SSTab1.Tab = 0
          Text1(12).SetFocus
          ChkInData = False
          'edit by nickc 2007/02/05 不用 dll 了
          'Set obj0702 = Nothing
          Exit Function
       End If
       'edit by nickc 2007/02/05 不用 dll 了
       'Set obj0702 = Nothing
   End If
    
   If Text1(17) <> "" Then
     strA = Text1(17)
      'Modify By Cheng 2002/07/09
'     If Not objPublicData.GetAgent(strA, strName) Then
     If Not PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
        SSTab1.Tab = 1
        Text1(17).SetFocus
        ChkInData = False
        Exit Function
     Else
        Text1(17) = strA
     End If
   End If

    If Text1(19) <> "" Then
      strA = Text1(19)
      'Modify By Cheng 2002/07/09
'      If Not objPublicData.GetAgent(strA, strName) Then
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If Not PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
      If Not ClsLawLawGetName(strA, strName) Then
      'end 2011/11/29
         SSTab1.Tab = 1
         Text1(19).SetFocus
         ChkInData = False
         Exit Function
      End If
    End If
'-------------------------------------------------------------------------
    If Text1(20) <> "" Then
      strA = Text1(20)
      'Modify By Cheng 2002/07/09
'      If Not objPublicData.GetAgent(strA, strName) Then
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If Not PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
      If Not ClsLawLawGetName(strA, strName) Then
      'end 2011/11/29
         SSTab1.Tab = 1
         Text1(20).SetFocus
         ChkInData = False
         Exit Function
      End If
    End If
    If Text1(25) <> "" Then
      strA = Text1(25)
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If Not PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
      If Not ClsLawLawGetName(strA, strName) Then
      'end 2011/11/29
         SSTab1.Tab = 1
         Text1(25).SetFocus
         ChkInData = False
         Exit Function
      End If
    End If
    
    'Add By Sindy 2016/11/23
   If Trim(Me.Combo4.Text) <> "" Then
      '若輸入幣別就一定要選格式
      If Trim(Me.Combo5.Text) = "" Then
         ShowMsg "請款單列印幣別格式不可空白 !"
         SSTab1.Tab = 1
         Me.Combo5.SetFocus
         ChkInData = False
         Exit Function
      End If
      '請款幣別<>NTD時不可輸入1
      If Trim(Me.Combo4.Text) <> "NTD" And Me.Combo5.ListIndex = 1 Then
         ShowMsg "請款幣別<>NTD時，請款單列印幣別格式不可選純台幣 !"
         SSTab1.Tab = 1
         Me.Combo5.SetFocus
         ChkInData = False
         Exit Function
      End If
   End If
   '2016/11/23 ENd
    
    ChkInData = True
End Function
'edit by nickc 2006/06/08
'Private Sub insertdata()
Private Function insertdata() As Boolean
   Dim i As Integer
   Dim oText As Object
   
   'Modify by Morgan 2009/9/15
   'For i = 0 To 25
   For Each oText In Text1
      i = oText.Index
   'end 2009/9/15
   
      sp(i) = Trim$(Text1(i).Text)
      Select Case i
         Case 9, 11
            sp(i) = TransDate(sp(i), 2)
         Case 7, 17, 19, 20, 25
            sp(i) = SPChangeCustomerL(sp(i))
      End Select
   Next
   
   If sp(2) = "" Then sp(2) = "0"
   If sp(3) = "" Then sp(3) = "00"
   
   
   'edit by nickc 2006/06/08
   'Set obj0702 = CreateObject("PrjTaieDll.Class0702")
   'nRet = obj0702.Adddata0702(sp)
   'Set obj0701 = Nothing
   insertdata = Adddata0702(sp)
End Function
'edit by nickc 2006/06/08
'Private Sub DeleteData()
Private Function DeleteData() As Boolean
 Dim i As Integer
 'add by nickc 2006/06/08
 DeleteData = False
 
   sp(0) = Trim$(Text1(0).Text)
   sp(1) = Trim$(Text1(1).Text)
   sp(2) = Trim$(Text1(2).Text)
   sp(3) = Trim$(Text1(3).Text)
   
   If sp(2) = "" Then sp(2) = "0"
   If sp(3) = "" Then sp(3) = "00"
   
   If ChkCaseCode("CP", sp(0), sp(1), sp(2), sp(3)) = False Then Exit Function
   'Add By Sindy 2010/7/1
   If ChkCaseCode("NP", sp(0), sp(1), sp(2), sp(3)) = False Then Exit Function
   '2010/7/1 End
   
   Select Case OnDataDeleteRecord(0, sp(0) & sp(1) & sp(2) & sp(3))
      Case 0
         'edit by nickc 2006/06/08
         'Set obj0702 = CreateObject("PrjTaieDll.Class0702")
         'nRet = obj0702.EraseData0702(sp)
         'Set obj0702 = Nothing
         DeleteData = EraseData0702(sp)
      Case -3
         
      Case Else
         MsgBox "新增資料至案件刪除記錄檔失敗，請洽系統管理員 !", vbCritical
   End Select

End Function
'edit by nickc 2006/06/08
'Private Sub UpdateData()
Private Function UpdateData() As Boolean
   
   Dim i As Integer
   'Modify by Morgan 2006/10/18
   'For i = 0 To 25
    For i = 0 To Me.Text1.Count - 1
      sp(i) = Trim$(Text1(i).Text)
      Select Case i
         Case 9, 11
            sp(i) = TransDate(sp(i), 2)
         Case 7, 17, 19, 20, 25
             sp(i) = SPChangeCustomerL(sp(i))
      End Select
   Next
   
   If sp(2) = "" Then sp(2) = "0"
   If sp(3) = "" Then sp(3) = "00"
   
   
   'edit by nickc 2006/06/08 紀錄 log
   'Set obj0702 = CreateObject("PrjTaieDll.Class0702")
   'nRet = obj0702.ModifyData0702(sp, True)
   'Set obj0702 = Nothing
   UpdateData = ModifyData0702(sp, True)
   
   'add by nickc 2005/08/23 紀錄修改案號
   pub_ModifyCaseNum = Trim(Text1(0)) & "-" & Trim(Text1(1)) & "-" & Trim(Text1(2)) & "-" & Trim(Text1(3))
End Function

Private Function GetMemo(ByVal strTmp As String) As String
   strExc(0) = "SELECT CU79 FROM CUSTOMER WHERE " & ChgCustomer(strTmp)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   GetMemo = ""
   If intI = 1 Then If Not IsNull(RsTemp.Fields(0)) Then GetMemo = RsTemp.Fields(0)
End Function

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("SP52")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP52")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SP52"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP53")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP53")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP53"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP54")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP54")) = False Then
         strTemp = rsSrcTmp.Fields("SP54")
         strCTime = Format(strTemp, "##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP55")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP55")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SP55"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP56")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP56")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SP56"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SP57")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SP57")) = False Then
         strTemp = rsSrcTmp.Fields("SP57")
         strUTime = Format(strTemp, "##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   'Modified by Morgan 2014/6/4 內容太長無法完全顯示,去掉欄位間的冒號
   textCUID = "CREATE : " & strCName & " " & _
              strCDate & " " & _
              strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              strUDate & " " & _
              strUTime
End Sub

Private Sub ShowDetail()
Dim i As Integer
'edit by nickc 2007/02/06 不用 dll 了 Dim obj01 As Object
Dim strA As String
'edit by nickc 2007/02/05 不用 dll 了
'Dim obj0702 As Object
Dim oText As Object 'Add by Morgan 2010/11/5
Dim arrID 'Add By Sindy 2025/1/7

   Clear_AllTxt
   For i = 0 To 25
      Text1(i).Text = IIf(IsNull(Rs(i)), "", Rs(i))
      Select Case i
         Case 9, 11
            If Not IsNull(Rs(i)) Then Text1(i).Text = TransDate(Rs(i), 1)
         Case 7, 17, 19, 20, 25
             Text1(i) = SPChangeCustomerS(Trim$(Text1(i)))
             Text1(i).Tag = Text1(i) 'Added by Lydia 2024/06/13
      End Select
   Next
   
   Text1(26).Text = "" & Rs("sp71") 'Add by Morgan 2006/10/18
    
    'Modify by Morgan 2009/9/15
    'For i = 0 To 13
    '  If i <> 1 Then Labeld1(i).Caption = ""
    'Next
    ClearLabel
    'end 2009/9/15
    
    'Add By Sindy 2021/12/7
    ' 更新CUID
    UpdateCUID Rs
    '2021/12/7 END
    
    'Create ID
    'edit by nickc 2007/02/02 不用 dll 了
    'If Not IsNull(rs(26)) Then If objPublicData.GetStaffN(rs(26), strA) Then Labeld1(7).Caption = strA
'    If Not IsNull(rs(26)) Then If ClsPDGetStaffN(rs(26), strA) Then Labeld1(7).Caption = strA
    'Updaet ID
    'edit by nickc 2007/02/02 不用 dll 了
    'If Not IsNull(rs(29)) Then If objPublicData.GetStaffN(rs(29), strA) Then Labeld1(10).Caption = strA
'    If Not IsNull(rs(29)) Then If ClsPDGetStaffN(rs(29), strA) Then Labeld1(10).Caption = strA
    'add by nickc 2006/07/12
    'edit by nickc 2007/02/02 不用 dll 了
    'If Not IsNull(rs.Fields("sp69")) Then If objPublicData.GetStaffN(rs.Fields("sp69"), strA) Then lblSP(69).Caption = strA
    If Not IsNull(Rs.Fields("sp69")) Then If ClsPDGetStaffN(Rs.Fields("sp69"), strA) Then lblSP(69).Caption = strA
    If Not IsNull(Rs.Fields("sp68")) Then lblSP(68) = ChangeWStringToTDateString(Rs.Fields("sp68"))
    If Not IsNull(Rs.Fields("sp61")) Then lblSP(61) = ChangeWStringToTDateString(Rs.Fields("sp61"))
    If Not IsNull(Rs.Fields("sp70")) Then lblSP(70) = Rs.Fields("sp70")
    If Not IsNull(Rs.Fields("SP85")) Then: textSP85 = Rs.Fields("SP85") 'Add By Sindy 2014/2/10
    
'    If Not IsNull(rs(27)) Then Labeld1(8).Caption = rs(27)
'    If Not IsNull(rs(30)) Then Labeld1(11).Caption = rs(30)
'    If Not IsNull(rs(28)) And rs(28) <> "" Then Labeld1(9).Caption = Format(rs(28), "##:##")
'    If Not IsNull(rs(31)) And rs(31) <> "" Then Labeld1(12).Caption = Format(rs(31), "##:##")
    PUB_AddContact "" & Rs("sp08"), cboContact, "" & Rs("sp78"), , True 'Add by Morgan 2008/8/4
    
    If Text1(7) <> "" Then
      'edit by nickc 2007/02/02 不用 dll 了
      'If objPublicData.GETCUSTOMER(Text1(7).Text, strA) Then
      If ClsPDGetCustomer(Text1(7).Text, strA) Then
         Labeld1(0).Caption = strA
      Else
         Labeld1(0).Caption = ""
      End If
      Text3 = GetMemo(Text1(7).Text)
    End If
    'edit by nickc 2007/02/02 不用 dll 了
    'If objPublicData.GetNation(Text1(8), strA) Then
    If ClsPDGetNation(Text1(8), strA) Then
       Labeld1(2).Caption = strA
    Else
       Labeld1(2).Caption = ""
    End If
    
    If Text1(12) <> "" Then
        'edit by nickc 2007/02/05 不用 dll 了
        'Set obj0702 = CreateObject("prjtaiedll.class0702")
        'If obj0702.GetReasonOfRelief(Text1(12), strA) Then
        If Cls0702GetReasonOfRelief(Text1(12), strA) Then
           Labeld1(3).Caption = strA
        Else
           Labeld1(3).Caption = ""
        End If
        'edit by nickc 2007/02/05 不用 dll 了
        'Set obj0702 = Nothing
    End If
     
    If Text1(17) <> "" Then
        strA = Text1(17)
         'Modify By Cheng 2002/07/09
'        If objPublicData.GetAgent(strA, strName) Then
        If PUB_GetAgentName(Me.Text1(0).Text, strA, strName) Then
           Me.Labeld1(4).Caption = strName
           Text2.Text = GetFa29(strA)
        Else
           Me.Labeld1(4).Caption = ""
        End If
    Else
      Text2.Text = ""
    End If
        
    If Text1(19) <> "" Then
      'Modify By Cheng 2002/07/09
'      If objPublicData.GetAgent(Text1(19).Text, strName) Then
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If PUB_GetAgentName(Me.Text1(0).Text, Text1(19).Text, strName) Then
      strA = Text1(19).Text
      If ClsLawLawGetName(strA, strName) Then
     'end 2011/11/29
         Me.Labeld1(5).Caption = strName
      Else
         Me.Labeld1(5).Caption = ""
      End If
    End If
       
    If Text1(20) <> "" Then
      'Modify By Cheng 2002/07/09
'      If objPublicData.GetAgent(Text1(20).Text, strName) Then
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If PUB_GetAgentName(Me.Text1(0).Text, Text1(20).Text, strName) Then
      strA = Text1(20).Text
      If ClsLawLawGetName(strA, strName) Then
     'end 2011/11/29
         Me.Labeld1(6).Caption = strName
      Else
         Me.Labeld1(6).Caption = ""
      End If
    End If
    If Text1(25) <> "" Then
      'Modify By Cheng 2002/07/09
'      If objPublicData.GetAgent(Text1(19).Text, strName) Then
      'Modified by Morgan 2011/11/29 也要能輸入申請人
      'If PUB_GetAgentName(Me.Text1(0).Text, Text1(25).Text, strName) Then
      strA = Text1(25).Text
      If ClsLawLawGetName(strA, strName) Then
     'end 2011/11/29
         Me.Labeld1(13).Caption = strName
      Else
         Me.Labeld1(13).Caption = ""
      End If
    End If
    
   'Add by Morgan 2009/9/7
   If IsNull(Rs.Fields("sp79")) Then
      Combo2 = ""
   Else
      Combo2 = Rs.Fields("sp79") + "." + PUB_GetFCPGrpName(Rs.Fields("sp79"))
   End If
   'Add by Morgan 2009/9/15
   For Each oText In txtSP
      oText = "" & Rs.Fields("SP" & oText.Index)
   Next
   
   'Add By Sindy 2016/11/23
   If IsNull(Rs.Fields("SP88")) = False Then
      For i = 0 To Combo4.ListCount - 1
         Combo4.ListIndex = i
         If InStr(Combo4.Text, Rs.Fields("SP88")) > 0 Then
            Exit For
         End If
      Next
   Else
      Combo4.ListIndex = 0
   End If
   If IsNull(Rs.Fields("SP89")) = False Then
      Combo5.ListIndex = Rs.Fields("SP89")
   Else
      Combo5.ListIndex = 0
   End If
   '2016/11/23 End
   
   'Add By Sindy 2025/1/7
   If IsNull(Rs.Fields("SP90")) = False Then
      arrID = Split(Rs.Fields("SP90"), ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         Chk1K(Val(arrID(intI)) - 1).Value = 1
      Next intI
   End If
   '2025/1/7 END
End Sub
Private Sub OpenTable()
   Dim strSys As String
   strSys = m_SysKind
   If Left(strSys, 1) = "'" Then
      strSys = Right(strSys, Len(strSys) - 1)
   End If
   If Right(strSys, 1) = "'" Then
      strSys = Left(strSys, Len(strSys) - 1)
   End If
   'Modify by Morgan 2008/8/4 +SP78
   'Modify by Morgan 2009/9/7 +SP79,SP80,SP81,SP82,SP83
   'Modify by Morgan 2010/11/5 +SP84
   'Modify by Sindy 2014/2/10 +SP85
   'Add by Lydia 2014/10/31 設別名f0+fmp2opensql
   'Modify by Sindy 2016/11/23 +SP88,SP89
   'Modified by Lydia 2017/06/22 因為外專只有FG,所以拿掉
   'strSql = "select sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16" & _
      ",sp17,sp34,sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27 " & _
      ",sp31,sp33,sp36, SP67,sp52," & SQLDate("SP53") & ",sp54,sp55," & SQLDate("SP56") & ",sp57 " & _
      ",sp61,sp68,sp69,sp70,sp71,sp78,sp79,sp80,sp81,sp82,sp83,sp84,SP85,SP88,SP89" & _
      " from servicepractice f0 where sp01 IN (" + CNULL(strSys) & ") " & FMP2openSQL
   'Modify By Sindy 2025/1/7 +,SP90
   strSql = "select sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16" & _
      ",sp17,sp34,sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27 " & _
      ",sp31,sp33,sp36, SP67,sp52,sp53,sp54,sp55,sp56,sp57 " & _
      ",sp61,sp68,sp69,sp70,sp71,sp78,sp79,sp80,sp81,sp82,sp83,sp84,SP85,SP88,SP89,SP90" & _
      " from servicepractice f0 where sp01 IN (" + CNULL(strSys) & ") "
    strSql = Replace(strSql, "f0.CP", "f0.SP")
    
    '依本所案號排序
    strSql = strSql & " Order By SP01, SP02, SP03, SP04 "
    'edit by nickc 2007/02/02 不用 dll 了
    'Set rs = objPublicData.ReadRst(strSQL, True)
    Set Rs = ClsPDReadRst(strSql, True)
End Sub

Private Sub QueryData(StrCriteria As String, i As Integer)
 On Error Resume Next
   Rs.Find StrCriteria
   i = i - 1
   If Rs.EOF Then
      ShowMsg MsgText(9007)
      DelFlg = False
      ShowData
      Exit Sub
   Else
      If i = 0 Then
         ShowDetail
         Call OnOff_Button(tlbar, True)
         ' 90.07.13 modify by louis (更新按紐的狀態)
         UpdateToolbarButtonState
         Exit Sub
      End If
      If i = 3 Then
         If DelFlg Then
           QueryData "sp02=" + CNULL(Fld6), i
        Else
           QueryData "sp02=" + CNULL(Fld2), i
        End If
      End If
      If i = 2 Then
        If DelFlg Then
           QueryData "sp03=" + CNULL(Fld7), i
        Else
           QueryData "sp03=" + CNULL(Fld3), i
        End If
      End If
      If i = 1 Then
         If DelFlg Then
           QueryData "sp04=" + CNULL(Fld8), i
        Else
           QueryData "sp04=" + CNULL(Fld4), i
        End If
      End If
   End If
   Exit Sub
ErrHand:
   MsgBox "錯誤 : " & Err.Description & "!", vbCritical
End Sub
Private Sub OnOffTxt(OnOffValue As Boolean)
Dim i As Integer
    'Modify by Morgan 2006/10/18
    'For i = 0 To 24
    For i = 0 To Me.Text1.Count - 1
        Text1(i).Locked = Not OnOffValue
    Next
    'Modify by Amy 2018/07/03 只有電腦中心才可改 特殊出名公司
    textSP85.Locked = True
    If Pub_StrUserSt03 = "M51" Then
        textSP85.Locked = Not OnOffValue 'Add By Sindy 2014/2/10
    End If
    'end 2018/07/03
   'Add By Sindy 2016/11/23
   Combo4.Locked = Not OnOffValue
   Combo5.Locked = Not OnOffValue
   '2016/11/23 End
   
   Frame1K.Enabled = OnOffValue 'Add By Sindy 2025/1/7
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim MsgAns As String
   Select Case KeyCode
      Case vbKeyEnd, vbKeyHome
         'If Data_Mission <> 1 And Data_Mission <> 3 And Data_Mission <> 4 Then
         '  UseDatamaintain (KeyCode)
         '  KeyCode = 0
         'End If
         If m_bQuery Then
            If Data_Mission <> 1 And Data_Mission <> 3 And Data_Mission <> 4 Then
              UseDatamaintain (KeyCode)
              KeyCode = 0
            End If
         End If
      'Case vbKeyF2, vbKeyF3, vbKeyF4, vbKeyF5, vbKeyF9, vbKeyF10, vbKeyEnd, vbKeyHome, vbKeyPageUp, vbKeyPageDown
      '    UseDatamaintain (KeyCode)
      '    KeyCode = 0
      Case vbKeyF2
         If m_bInsert Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyF3
         If m_bUpdate Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyF4
         If m_bQuery Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyF5
         If m_bDelete Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyF9, vbKeyF10
         UseDatamaintain (KeyCode)
         KeyCode = 0
      Case vbKeyEnd, vbKeyHome, vbKeyPageUp, vbKeyPageDown
         If m_bQuery Then
            UseDatamaintain (KeyCode)
            KeyCode = 0
         End If
      Case vbKeyEscape
         MsgAns = MsgBox("是否確定結束?", vbYesNo + vbCritical)
         If MsgAns = vbYes Then
            UseDatamaintain (KeyCode)
         End If
   End Select
    
End Sub
Private Sub UseDatamaintain(j As Integer)
Dim i As Integer
Dim MsgAns As Integer
   If Data_Mission <> 2 And Not GetNowData Then
        Fld1 = Trim$(Text1(0).Text)
        Fld2 = Trim$(Text1(1).Text)
        Fld3 = Trim$(Text1(2).Text)
        Fld4 = Trim$(Text1(3).Text)
   End If
   If Data_Mission = 2 And GetNowData Then
        Fld1 = Trim$(Text1(0).Text)
        Fld2 = Trim$(Text1(1).Text)
        Fld3 = Trim$(Text1(2).Text)
        Fld4 = Trim$(Text1(3).Text)
        If RsCounts > 1 Then
            Rs.MovePrevious
            If Rs.BOF Then
                Rs.MoveFirst
                Rs.MoveNext
            End If
            Fld5 = Rs(0)
            Fld6 = Rs(1)
            Fld7 = Rs(2)
            Fld8 = Rs(3)
        End If
   End If

    Select Case j
    Case vbKeyF2
            If blnKeypreview Then
            Data_Mission = 1
            ChkData = False
            'Modify by Morgan 2009/9/15 都要清除
            'Call Clear_AllTxtAry(Text1, 0, 24)
            Clear_AllTxt
            
            ' 90.07.13 modify by louis (更新按紐的狀態)
            UpdateToolbarButtonState
            
            'Modify by Morgan 2009/9/15
            'For i = 0 To 13
            '   Labeld1(i).Caption = ""
            'Next
            ClearLabel
            'end 2009/9/15
            
            Combo2 = "" 'Add by Morgan 2009/9/8
            
            GetNowData = True
            ChkData = True
            Call OnOff_Button(tlbar, False)
            OnOffTxt True
            blnKeypreview = False
            Text1(0).SetFocus
            End If
    Case vbKeyF3
            If blnKeypreview Then
            Data_Mission = 3
            Call OnOff_Button(tlbar, False)
            ' 90.07.13 modify by louis (更新按紐的狀態)
            UpdateToolbarButtonState
            OnOffTxt True
            Text1(1).Locked = True
            Text1(0).Locked = True
            Text1(2).Locked = True
            Text1(3).Locked = True
            blnKeypreview = False
            Text1(8).SetFocus
            End If
    Case vbKeyF5
            If blnKeypreview Then
            Data_Mission = 2
            GetNowData = True
            Call OnOff_Button(tlbar, False)
            ' 90.07.13 modify by louis (更新按紐的狀態)
            UpdateToolbarButtonState
            blnKeypreview = False
            End If
    Case vbKeyF4
            Rs.MoveFirst 'Add By Sindy 2021/12/7
            If blnKeypreview Then
            Data_Mission = 4
            GetNowData = True
            ChkData = False
            'Modify by Morgan 2009/9/15 都要清除
            'Call Clear_AllTxtAry(Text1, 0, 24)
            Clear_AllTxt
            
            'Modify by Morgan 2009/9/15
            'For i = 0 To 7
            '     Labeld1(i).Caption = ""
            'Next
            ClearLabel
            'end 2009/9/15
            
            ChkData = True
            Call OnOff_Button(tlbar, False)
            ' 90.07.13 modify by louis (更新按紐的狀態)
            UpdateToolbarButtonState
            Text1(1).Locked = False
            Text1(0).Locked = False
            Text1(2).Locked = False
            Text1(3).Locked = False
            blnKeypreview = False
            Text1(0).SetFocus
            End If
    Case vbKeyHome
            If blnKeypreview Then
                Rs.MoveFirst
                ShowDetail
                Text1(0).SetFocus
            End If
    Case vbKeyPageUp
            If blnKeypreview Then
                Rs.MovePrevious
                If Rs.BOF Then
                    Rs.MoveFirst
                    ShowMsg MsgText(9008)
                Else
                   ShowDetail
                   Text1(0).SetFocus
                End If
            End If
            
    Case vbKeyPageDown
            If blnKeypreview Then
                Rs.MoveNext
                If Rs.EOF Then
                    Rs.MoveLast
                    ShowMsg MsgText(9009)
                Else
                   ShowDetail
                   Text1(0).SetFocus
                End If
            End If
    Case vbKeyEnd
            If blnKeypreview Then
                Rs.MoveLast
                ShowDetail
                Text1(0).SetFocus
            End If
    Case vbKeyF9
            If Not blnKeypreview Then
            Select Case Data_Mission
            Case 1
                               
               If ChkInData Then
                  'Add By Cheng 2002/05/22
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Sub
                  
                  'edit by nickc 2006/06/08
                  'InsertData
                  If insertdata = False Then Exit Sub
                  DelFlg = False
               Else
                  Exit Sub
               End If
               Fld1 = Trim$(Text1(0).Text)
               Fld2 = Trim$(Text1(1).Text)
               Fld3 = Trim$(Text1(2).Text)
               Fld4 = Trim$(Text1(3).Text)
               ShowData
               
            Case 2
               If DelMsg Then
                  'edit by nickc 2006/06/08
                  'DeleteData
                  If DeleteData = False Then Exit Sub
                  
                  DelFlg = True
               End If
               ShowData
               DelFlg = False
               
            Case 3

               If ChkInData Then
                  'Add By Cheng 2002/05/22
                  '重新檢查欄位有效性
                  If TxtValidate = False Then Exit Sub
                  
                  'Added by Lydia 2017/06/19 (存檔前)檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
                  strChkCuAreaMail = PUB_ChkSameCustSales(Trim(Text1(0)), Trim(Text1(1)), Trim(Text1(2)), Trim(Text1(3)), "", Trim(Text1(7)), "", "", "", "", strChkCuAreaMailTo)
                  
                  'edit by nickc 2006/06/08
                  'UpdateData
                  If UpdateData = False Then Exit Sub
                  
                  'Added by Lydia 2017/06/19 檢查收文智權人員和客戶智權人員不同業務區要發mail的內文和通知人員
                  If strChkCuAreaMail <> "" Then
                     PUB_SendMail strUserNum, strChkCuAreaMailTo, "", "案件收文通知--此案收文非原智權人員(區)！", strChkCuAreaMail
                  End If
                  'end 2017/06/19
                  
                  ShowData
                  DelFlg = False
               Else
                  Exit Sub
               End If
              
          Case 4
                
               Fld5 = Trim$(Text1(0).Text)
               Fld6 = Trim$(Text1(1).Text)
               If Text1(2).Text = "" Then
                  Fld7 = "0"
               Else
                  Fld7 = Trim$(Text1(2).Text)
               End If
               If Text1(3).Text = "" Then
                  Fld8 = "00"
               Else
                  Fld8 = Trim$(Text1(3).Text)
               End If
               DelFlg = True
               If Rs.BOF <> Rs.EOF Then Rs.MoveFirst
               QueryData "sp01=" + CNULL(Fld5), 4
         End Select
         Data_Mission = 0
         Call OnOff_Button(tlbar, True)
         ' 90.07.13 modify by louis (更新按紐的狀態)
         UpdateToolbarButtonState
         OnOffTxt False
         GetNowData = False
         ChkData = True
         blnKeypreview = True
         Text1(0).SetFocus
         End If
    Case vbKeyF10
         If Not blnKeypreview Then
            If MsgBox("你並未存檔，確定離開嗎 ?", vbYesNo + vbCritical + vbDefaultButton2) = vbYes Then
               Call OnOff_Button(tlbar, True)
                'Add By Cheng 2003/04/11
                '設定按扭狀態
                UpdateToolbarButtonState
               DelFlg = False
               Data_Mission = 0
               OnOffTxt False
               ChkData = True
               ShowData
               GetNowData = False
               blnKeypreview = True
               Text1(0).SetFocus
            End If
         End If
    Case vbKeyEscape
        Unload Me
    End Select
End Sub
Private Sub GetSysInf()
'    Nret = objPublicData.GetSystemKind(strSysKind, intSysnum, , intWhere)
End Sub

'Add By Sindy 2014/2/10
Private Sub textSP85_GotFocus()
   InverseTextBox textSP85
End Sub
Private Sub textSP85_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
'特殊出名公司
Private Sub textSP85_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   textSP85 = Trim(textSP85)
   If textSP85 <> "" And textSP85 <> "J" Then
      Cancel = True
      strTit = "檢核資料"
      strMsg = "請輸入J或空白"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textSP85_GotFocus
   End If
End Sub

Private Sub tlbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        UseDatamaintain (vbKeyF2)
    Case 2
        UseDatamaintain (vbKeyF3)
    Case 3
        UseDatamaintain vbKeyF5
        UseDatamaintain vbKeyF9
    Case 4
        UseDatamaintain (vbKeyF4)
    Case 6
        UseDatamaintain (vbKeyHome)
    Case 7
        UseDatamaintain (vbKeyPageUp)
    Case 8
        UseDatamaintain (vbKeyPageDown)
    Case 9
        UseDatamaintain (vbKeyEnd)
    Case 11
        UseDatamaintain (vbKeyF9)
    Case 12
        UseDatamaintain (vbKeyF10)
    Case 14
        UseDatamaintain (vbKeyEscape)
    End Select
End Sub

Private Sub Clear_AllTxt()
   Dim txt As Object
   
   For Each txt In Text1
      txt.Text = ""
   Next
   textSP85 = Empty 'Add By Sindy 2014/2/10
   'add by nickc 2006/07/12
   lblSP(61) = ""
   lblSP(68) = ""
   lblSP(69) = ""
   lblSP(70) = ""
   cboContact.Clear 'Add by Morgan 2008/8/4
   Combo2 = "" 'Add by Morgan 2009/9/7
   'Add by Morgan 2009/9/15
   Text2 = ""
   Text3 = ""
   For Each txt In txtSP
      txt.Text = ""
   Next
   'Add By Sindy 2016/11/23
   If Combo4.Visible = True Then
      Me.Combo4.ListIndex = 0
      Me.Combo5.ListIndex = 0
   End If
   '2016/11/23 End
   textCUID = "" 'Add By Sindy 2021/12/7
   
   'Add By Sindy 2025/1/7
   For Each txt In Chk1K
      txt = Empty
   Next
   '2025/1/7 END
End Sub

' 90.07.12 modify by louis (更新按紐的狀態)
Private Sub UpdateToolbarButtonState()
   If Not m_bInsert Then
      tlbar.Buttons(1).Enabled = False
   End If
   If Not m_bUpdate Then
      tlbar.Buttons(2).Enabled = False
   End If
   If Not m_bQuery Then
        'Modify By Cheng 2003/04/11
'      tlbar.Buttons(3).Enabled = False
      tlbar.Buttons(4).Enabled = False
   End If
   If Not m_bDelete Then
        'Modify By Cheng 2003/04/11
'      tlbar.Buttons(4).Enabled = False
      tlbar.Buttons(3).Enabled = False
   End If
   If Not m_bQuery Then
      tlbar.Buttons(6).Enabled = False
      tlbar.Buttons(7).Enabled = False
      tlbar.Buttons(8).Enabled = False
      tlbar.Buttons(9).Enabled = False
   End If
End Sub

' 91.05.16 modify by louis
Private Sub SetQueryStatus()
   If blnKeypreview Then
      Dim i As Integer
      Data_Mission = 4
      GetNowData = True
      ChkData = False
      'Modify by Morgan 2009/9/15 都要清除
      'Call Clear_AllTxtAry(Text1, 0, 24)
      Clear_AllTxt
      
      'Modify by Morgan 2009/9/15
      'For i = 0 To 7
      '     Labeld1(i).Caption = ""
      'Next
      ClearLabel
      'end 2009/9/15
      
      ChkData = True
      Call OnOff_Button(tlbar, False)
      ' 90.07.13 modify by louis (更新按紐的狀態)
      UpdateToolbarButtonState
      Text1(1).Locked = False
      Text1(0).Locked = False
      Text1(2).Locked = False
      Text1(3).Locked = False
      blnKeypreview = False
   End If
End Sub

Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False

'Add by Sindy 2021/12/07 檢查畫面上的物件是否含有Unicode文字
If PUB_ChkUniText(Me, True, True) = False Then
   Exit Function
End If

For Each objTxt In Text1
   If objTxt.Enabled = True Then
      Cancel = False
      Text1_Validate objTxt.Index, Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
Next
'Add by Morgan 2007/5/10
If Not ((Text1(10).Text = "" And Text1(11).Text = "" And Text1(12).Text = "") Or (Text1(10).Text <> "" And Text1(11).Text <> "" And Text1(12).Text <> "")) Then
   MsgBox "是否閉卷、閉卷日期、閉卷原因三個欄位須同時空白或有值！", vbExclamation
   Exit Function
End If
'end 2007/5/10
'Add by Morgan 2009/9/9
If Text1(0) = "FG" And Text1(1) >= "000536" And Combo2 = "" Then
   MsgBox "請輸入FG工程師組別"
   Combo2.SetFocus
   Exit Function
End If
   
   '2010/1/8 ADD BY SONIA
   If Combo2.Enabled = True Then
      Cancel = False
      Combo2_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   '2010/1/8 END
   
'Add By Sindy 2014/2/10
If Me.textSP85.Enabled = True Then
   Cancel = False
   textSP85_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

'Add By Sindy 2016/11/23
If Me.Combo4.Enabled = True Then
   Cancel = False
   Combo4_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If
'2016/11/23 End
'Added by Lydia 2024/06/13 檢查更新代理人／申請人狀態排除「不得代理」
strExc(1) = ChangeCustomerL(Text1(7))
strExc(2) = ChangeCustomerL(Text1(7).Tag)
If strExc(1) <> "" And strExc(1) <> strExc(2) Then
   If GetCustomerAndState(strExc(1), strExc(3), , , , Text1(0), strExc(8), False, Me.Name, Text1(1), Text1(2), Text1(3)) = False Then
      Me.SSTab1.Tab = 0
      Text1(7).SetFocus
      Text1_GotFocus 7
      Exit Function
   End If
End If
strExc(1) = ChangeCustomerL(Text1(17))
strExc(2) = ChangeCustomerL(Text1(17).Tag)
If strExc(1) <> "" And strExc(1) <> strExc(2) Then
   If GetAgentAndState(strExc(1), strExc(3), , , , Text1(0), strExc(8), False) = False Then
      Me.SSTab1.Tab = 1
      Text1(17).SetFocus
      Text1_GotFocus 7
      Exit Function
   End If
End If
'end 2024/06/13
TxtValidate = True
End Function
'add by nickc 2006/06/08 dll 搬出
Private Function ModifyData0702(sp() As String, Optional BolWDB As Boolean = False) As Integer
   Dim i As Integer
   Dim j As Integer
   Dim BolTransOk As Boolean
   Dim oText As Object 'Add by Morgan 2009/9/15
   
   BolTransOk = True
    'Modified by Morgan 2021/8/25 修改案件名稱有單引號問題 Ex:FG-001323
    strSql = "update servicepractice set sp05=" + CNULL(ChgSQL(sp(4))) + ","
    strSql = strSql + "sp06=" + CNULL(ChgSQL(sp(5))) + ",sp07=" + CNULL(ChgSQL(sp(6))) + ","
    strSql = strSql + "sp08=" + CNULL(sp(7)) + ",sp09=" + CNULL(sp(8)) + ","
    strSql = strSql + "sp10=" + CNULL(sp(9)) + ",sp15=" + CNULL(sp(10)) + ","
    strSql = strSql + "sp16=" + CNULL(sp(11)) + ",sp17=" + CNULL(sp(12)) + ","
    strSql = strSql + "sp34=" + CNULL(sp(13)) + ",sp28=" + CNULL(sp(14)) + ","
    strSql = strSql + "sp18=" + CNULL(sp(15)) + ",sp29=" + CNULL(sp(16)) + ","
    strSql = strSql + "sp26=" + CNULL(sp(17)) + ",sp30=" + CNULL(sp(18)) + ","
    strSql = strSql + "sp37=" + CNULL(sp(19)) + ",sp35=" + CNULL(sp(20)) + ","
    strSql = strSql + "sp27=" + CNULL(sp(21)) + ",sp31=" + CNULL(sp(22)) + ","
    strSql = strSql + "sp33=" + CNULL(sp(23)) + ",sp36=" + CNULL(sp(24)) + ",sp67=" + CNULL(sp(25))
    strSql = strSql + ",sp71=" + CNULL(sp(26)) 'Add by Morgan 2006/10/18
    strSql = strSql + ",sp79='" & Left(Combo2, 1) & "'" 'Add by Morgan 2009/9/8
    
    'Add by Morgan 2009/9/15
    'SP80以後改用新陣列(index=欄位數字部份)以方便後續再新增欄位
    For Each oText In txtSP
      Select Case oText.Index
         '數字欄位
         Case 81, 82
            strSql = strSql & ",sp" & oText.Index & "=" & CNULL(oText.Text, True)
         '文字欄位
         Case Else
            strSql = strSql & ",sp" & oText.Index & "=" & CNULL(ChgSQL(oText.Text))
      End Select
    Next
    'end 2009/9/15
    strSql = strSql & ",sp85=" & CNULL(textSP85.Text) 'Add By Sindy 2014/2/10
    'Add By Sindy 2016/11/23
    strSql = strSql & ",sp88=" & CNULL(Combo4.Text)
    strSql = strSql & ",sp89=" & CNULL(IIf(Combo5.Text <> "", Combo5.ListIndex, ""))
    '2016/11/23 END
      'Add By Sindy 2025/1/7
      strExc(10) = ""
      For Each oText In Chk1K
         If oText.Value = 1 Then
            strExc(10) = strExc(10) & "," & oText.Index + 1
         End If
      Next
      If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
      strSql = strSql & ",sp90=" & CNULL(strExc(10))
      '2025/1/7 END
      
    strSql = strSql + " where sp01=" + CNULL(sp(0)) + " and sp02=" + CNULL(sp(1)) + " and sp03=" + CNULL(sp(2)) + "  and sp04=" + CNULL(sp(3))
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    'add by nickc 2006/06/08 紀錄分析語法
    Pub_SeekTbLog strSql
    '910910  nick tigger
    '***** start
    If BolWDB = True Then
        strSql = "begin user_data.user_enabled:=1; " & strSql & "; end;"
    End If
    '***** end
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    ModifyData0702 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    MsgBox Err.Description
    ModifyData0702 = False
    'ErrorLog
End Function

Private Function Adddata0702(sp() As String) As Boolean
   Dim i As Integer
   Dim j As Integer
   Dim BolTransOk As Boolean
   'Add by Morgan 2009/9/15
   Dim oText As Object
   Dim stValues As String
   
BolTransOk = True
    'Modify By Cheng 2003/09/24
'    strSQL = "insert into Servicepractice(sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16,sp17,sp34,"
'    strSQL = strSQL + "sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27,sp31,sp33,sp36"
'    strSQL = strSQL + ") values("
    'Modify by Morgan 2006/10/18 +sp71
    'Modify by Morgan 2009/9/8 +sp79
    strSql = "insert into Servicepractice (sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp09,sp10,sp15,sp16,sp17,sp34,"
    strSql = strSql + "sp28,sp18,sp29,sp26,sp30,sp37,sp35,sp27,sp31,sp33,sp36,SP67,sp71,sp79"
    
    'Modify by Morgan 2006/10/18
    'For i = 0 To 25
    For i = 0 To Me.Text1.Count - 1
        stValues = stValues + CNULL(sp(i)) + ","
    Next
    'Modify by Morgan 2009/9/8
    'strSQL = Mid$(strSQL, 1, Len(strSQL) - 1)
    stValues = stValues & "'" & Left(Combo2, 1) & "'"
    'end 2009/9/8
    
    'Add by Morgan 2009/9/15
    For Each oText In txtSP
      If Not IsEmpty(oText.Text) Then
         strSql = strSql & ",sp" & oText.Index
         Select Case oText.Index
            '數字欄位
            Case 81, 82
               stValues = stValues & "," & CNULL(oText.Text, True)
            '文字欄位
            Case Else
               stValues = stValues & "," & CNULL(ChgSQL(oText.Text))
         End Select
      End If
    Next
    'end 2009/9/15
    'Modify By Sindy 2014/2/10 +,sp85
    'Modify By Sindy 2016/11/23 +,sp88,sp89
    'Modify By Sindy 2025/1/7 +,sp90
    strSql = strSql + ",sp85,sp88,sp89,sp90"
    stValues = stValues + "," & CNULL(textSP85.Text)
    'Add By Sindy 2016/11/23
    stValues = stValues + "," & CNULL(Combo4.Text)
    stValues = stValues + "," & CNULL(IIf(Combo5.Text <> "", Combo5.ListIndex, ""))
    '2016/11/23 END
      'Add By Sindy 2025/1/7
      strExc(10) = ""
      For Each oText In Chk1K
         If oText.Value = 1 Then
            strExc(10) = strExc(10) & "," & oText.Index + 1
         End If
      Next
      If strExc(10) <> "" Then strExc(10) = Mid(strExc(10), 2)
      stValues = stValues + "," & CNULL(strExc(10))
      '2025/1/7 END
    strSql = strSql + ") values (" & stValues & ")"
    
   'Debug.Print strSQL
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    'add by nickc 2006/06/08 紀錄分析語法
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    Adddata0702 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    MsgBox Err.Description
    Adddata0702 = False
    'ErrorLog
End Function

Private Function EraseData0702(sp() As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
    
    'Debug.Print strSQL
On Error GoTo ErrHand
    cnnConnection.BeginTrans
    'add by nickc 2006/06/08 紀錄分析語法
    'Move by Lydia 2016/11/24 從GoTo ErrHand上方移過來
    strSql = "delete from servicepractice where sp01='" + sp(0) + "' and sp02='" + sp(1) + "' and sp03='" + sp(2) + "' and sp04='" + sp(3) + "'"
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    'Modify By Cheng 2002/11/14
    'Added by Lydia 2016/11/24 一併刪除各項指示
    strSql = "DELETE FROM INSTRUCTIONS WHERE ITS01=" & CNULL(Pub_GetITS01Type(sp(0))) & " AND ITS02=" & CNULL(sp(0) & sp(1) & sp(2) & sp(3))
    Pub_SeekTbLog strSql
    cnnConnection.Execute strSql
    'end 2016/11/24
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
    EraseData0702 = True
    Exit Function
ErrHand:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If
    
    cnnConnection.RollbackTrans
    EraseData0702 = False
    MsgBox Err.Description
    'ErrorLog
End Function
'Add by Morgan 2009/9/15
'清除內容
Private Sub ClearLabel()
   Dim oLabel As Object
   For Each oLabel In Labeld1
      oLabel.Caption = ""
   Next
End Sub
'Add by Morgan 2009/9/15
'還原底色
Private Sub SetBackColor()
   Dim oLabel As Object
   For Each oLabel In Labeld1
      oLabel.BackColor = &H8000000F '設定底色(設計時改白色較方便)
   Next
   For Each oLabel In lblSP
      oLabel.BackColor = &H8000000F '設定底色(設計時改白色較方便)
   Next
End Sub

Private Sub txtSP_GotFocus(Index As Integer)
   TextInverse txtSP(Index)
End Sub

Private Sub txtSP_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      'Added by Morgan 2014/6/4
      Case 80
         If KeyAscii <> 89 And KeyAscii <> 68 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      'Modified by Morgan 2014/6/4
      'Case 80, 83
      Case 83
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
      Case 81, 82
         If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         End If
   End Select
End Sub

'Move by Lydia 2022/09/06 從basQuery搬過來; 並且把Public改回Private
'檢查新增之本所案號是否大於現在AutoNum
Private Function Cls0702ChkCaseNoWithAuNum0702(strNo1 As String, strNo2 As String, strNo3 As String, strNo4 As String, strSysKind As String) As Boolean
'edit by nickc 2007/02/06 不用 dll 了
 'Dim rs As ADODB.Recordset, nRet As Boolean, strName As String, obj01 As Object
 Dim Rs As ADODB.Recordset, nRet As Boolean, strName As String
   strSql = "select * from servicepractice where sp01=" + CNULL(strNo1) + " and sp02=" + CNULL(strNo2) + " and sp03=" + CNULL(strNo3) + " and sp04=" + CNULL(strNo4)
   'edit by nickc 2007/02/06 不用 dll 了
   'Set obj01 = CreateObject("prjtaiedll.clspublicdata")
   'Set rs = obj01.ReadRst(strSQL)
   Set Rs = ClsPDReadRst(strSql)
   If Not Rs.EOF Then
      MsgBox "本所案號重覆 !", vbCritical
      Cls0702ChkCaseNoWithAuNum0702 = False
      Rs.Close
      Exit Function
   End If
   Rs.Close
   strName = ""
   strSql = "select * from autonumber where au01=" + CNULL(strSysKind) + " and au03 >" + CStr(Val(strNo2))
   'edit by nickc 2007/02/06 不用 dll 了
   'Set rs = obj01.ReadRst(strSQL)
   'Set obj01 = Nothing
   Set Rs = ClsPDReadRst(strSql)
   If Rs.EOF Then
       Rs.Close
       MsgBox "本所案號不合規則 !", vbCritical
       Cls0702ChkCaseNoWithAuNum0702 = False
   Else
       Cls0702ChkCaseNoWithAuNum0702 = True  '可以新增資料
       Rs.Close
   End If
End Function

