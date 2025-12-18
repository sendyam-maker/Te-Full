VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010602_3 
   BorderStyle     =   1  '³æ½u©T©w
   Caption         =   "®Ö­ã¨ç¿é¤J"
   ClientHeight    =   6324
   ClientLeft      =   -1020
   ClientTop       =   996
   ClientWidth     =   8940
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6324
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdMod 
      Caption         =   "ÅÜ§ó¨Æ¶µ(R)"
      Height          =   400
      Left            =   4710
      TabIndex        =   72
      Top             =   15
      Visible         =   0   'False
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4545
      Left            =   120
      TabIndex        =   44
      Top             =   1740
      Width           =   8655
      _ExtentX        =   15261
      _ExtentY        =   8022
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "®Ö­ã¸ê®Æ"
      TabPicture(0)   =   "frm06010602_3.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label9(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label10"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label11"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label13"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label14"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label15"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label9(1)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9(2)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label3(6)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblCP19"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label8"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label27(4)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label27(3)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label27(2)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label27(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label34"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label32"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label12"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label29"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label27(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label26(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Text9(0)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Text9(1)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Text9(2)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "LblFM2(1)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Text33(10)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text33(9)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Text33(13)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Text33(12)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "Text33(11)"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "lbl415Date"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "Text6"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "Text7"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "Text10(0)"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "Text10(1)"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "Text10(2)"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCP19"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Text16"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "Check1"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "Frame1"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt415Date"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Ápµ¸¤H¸ê®Æ"
      TabPicture(1)   =   "frm06010602_3.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label16"
      Tab(1).Control(1)=   "Label17"
      Tab(1).Control(2)=   "Label18"
      Tab(1).Control(3)=   "Label19"
      Tab(1).Control(4)=   "Label20"
      Tab(1).Control(5)=   "Label21"
      Tab(1).Control(6)=   "Label22"
      Tab(1).Control(7)=   "Label23"
      Tab(1).Control(8)=   "Label25"
      Tab(1).Control(9)=   "Label26(1)"
      Tab(1).Control(10)=   "Label28(0)"
      Tab(1).Control(11)=   "Label5"
      Tab(1).Control(12)=   "LblFM2(2)"
      Tab(1).Control(13)=   "Text33(5)"
      Tab(1).Control(14)=   "Text33(4)"
      Tab(1).Control(15)=   "Text33(3)"
      Tab(1).Control(16)=   "Text33(2)"
      Tab(1).Control(17)=   "Text33(1)"
      Tab(1).Control(18)=   "Text33(0)"
      Tab(1).Control(19)=   "Text33(6)"
      Tab(1).Control(20)=   "Text12"
      Tab(1).Control(21)=   "Text19"
      Tab(1).Control(22)=   "Text20"
      Tab(1).Control(23)=   "Text21"
      Tab(1).Control(24)=   "Text22"
      Tab(1).Control(25)=   "Combo1(0)"
      Tab(1).Control(26)=   "Combo1(1)"
      Tab(1).ControlCount=   27
      TabCaption(2)   =   "Àu¥ýÅv"
      TabPicture(2)   =   "frm06010602_3.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdDataList2"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt415Date 
         Height          =   300
         Left            =   6510
         MaxLength       =   7
         TabIndex        =   1
         Top             =   930
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   1
         Left            =   -68865
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   18
         Top             =   432
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   300
         Index           =   0
         Left            =   -71760
         Style           =   2  '³æ¯Â¤U©Ô¦¡
         TabIndex        =   19
         Top             =   432
         Width           =   2295
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  '¨S¦³®Ø½u
         Caption         =   "Frame1"
         Height          =   250
         Left            =   210
         TabIndex        =   94
         Top             =   1830
         Width           =   8205
         Begin VB.OptionButton Opt1 
            Caption         =   "±M§QÅvÅÜ§ó"
            Height          =   255
            Index           =   2
            Left            =   5880
            TabIndex        =   13
            Top             =   0
            Width           =   1425
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "§ó¥¿"
            Height          =   255
            Index           =   1
            Left            =   4935
            TabIndex        =   12
            Top             =   0
            Width           =   885
         End
         Begin VB.OptionButton Opt1 
            Caption         =   "°É»~"
            Height          =   255
            Index           =   0
            Left            =   3990
            TabIndex        =   11
            Top             =   0
            Width           =   885
         End
         Begin VB.TextBox txtCRC 
            Height          =   270
            Index           =   1
            Left            =   2820
            MaxLength       =   2
            TabIndex        =   10
            Top             =   0
            Width           =   555
         End
         Begin VB.TextBox txtCRC 
            Height          =   270
            Index           =   0
            Left            =   1380
            MaxLength       =   7
            TabIndex        =   9
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "²Ä¡@¡@¡@¡@´Á ¤§"
            Height          =   180
            Left            =   2550
            TabIndex        =   96
            Top             =   45
            Width           =   1305
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "°É»~¤é´Á¡G"
            Height          =   180
            Left            =   0
            TabIndex        =   95
            Top             =   45
            Width           =   900
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "¦³ÀË¯Á"
         Height          =   255
         Left            =   2940
         TabIndex        =   2
         Top             =   353
         Width           =   1155
      End
      Begin VB.TextBox Text16 
         Height          =   300
         Left            =   5880
         MaxLength       =   6
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox txtCP19 
         Height          =   300
         Left            =   5880
         MaxLength       =   7
         TabIndex        =   3
         Top             =   330
         Width           =   1095
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Index           =   2
         Left            =   1575
         MaxLength       =   1
         TabIndex        =   8
         Top             =   1530
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Index           =   1
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   6
         Top             =   930
         Width           =   255
      End
      Begin VB.TextBox Text10 
         Height          =   300
         Index           =   0
         Left            =   1575
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   4
         Top             =   630
         Width           =   255
      End
      Begin VB.TextBox Text22 
         Height          =   525
         Left            =   -72600
         MaxLength       =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  '««ª½±²¶b
         TabIndex        =   30
         Top             =   3060
         Width           =   6012
      End
      Begin VB.TextBox Text21 
         Height          =   525
         Left            =   -72600
         MaxLength       =   140
         MultiLine       =   -1  'True
         ScrollBars      =   2  '««ª½±²¶b
         TabIndex        =   29
         Top             =   2490
         Width           =   6012
      End
      Begin VB.TextBox Text20 
         Height          =   300
         Left            =   -73320
         MaxLength       =   35
         TabIndex        =   28
         Top             =   2147
         Width           =   1335
      End
      Begin VB.TextBox Text19 
         Height          =   300
         Left            =   -73320
         MaxLength       =   9
         TabIndex        =   27
         Top             =   1804
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Height          =   300
         Left            =   -73560
         TabIndex        =   17
         Top             =   432
         Width           =   1335
      End
      Begin VB.TextBox Text7 
         Height          =   300
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1230
         Width           =   6735
      End
      Begin VB.TextBox Text6 
         Height          =   300
         Left            =   1575
         MaxLength       =   7
         TabIndex        =   0
         Top             =   330
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList2 
         Height          =   4065
         Left            =   -74940
         TabIndex        =   93
         Top             =   360
         Width           =   8535
         _ExtentX        =   15050
         _ExtentY        =   7176
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         HighLight       =   0
         AllowUserResizing=   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé-ExtB"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin VB.Label lbl415Date 
         AutoSize        =   -1  'True
         Caption         =   "±M§QÅv´Á¶¡©µªø¦Ü                           ¤î"
         Height          =   180
         Left            =   5010
         TabIndex        =   100
         Top             =   990
         Width           =   2835
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   11
         Left            =   1575
         TabIndex        =   80
         Top             =   3580
         Width           =   1095
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   12
         Left            =   1575
         TabIndex        =   79
         Top             =   3880
         Width           =   1095
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   13
         Left            =   1575
         TabIndex        =   78
         Top             =   4170
         Width           =   1095
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   9
         Left            =   1575
         TabIndex        =   77
         Top             =   2980
         Width           =   1095
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   10
         Left            =   1575
         TabIndex        =   76
         Top             =   3280
         Width           =   1095
         VariousPropertyBits=   671105055
         BackColor       =   -2147483633
         MaxLength       =   9
         Size            =   "1931;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   6
         Left            =   -73320
         TabIndex        =   26
         Top             =   1461
         Width           =   6750
         VariousPropertyBits=   671105051
         Size            =   "11906;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   0
         Left            =   -73560
         TabIndex        =   20
         Top             =   775
         Width           =   1335
         VariousPropertyBits=   671105051
         Size            =   "2355;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   1
         Left            =   -70920
         TabIndex        =   21
         Top             =   775
         Width           =   1455
         VariousPropertyBits=   671105051
         Size            =   "2566;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   2
         Left            =   -68145
         TabIndex        =   22
         Top             =   775
         Width           =   1575
         VariousPropertyBits=   671105051
         Size            =   "2778;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   3
         Left            =   -73560
         TabIndex        =   23
         Top             =   1118
         Width           =   1335
         VariousPropertyBits=   671105051
         Size            =   "2355;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   4
         Left            =   -70920
         TabIndex        =   24
         Top             =   1118
         Width           =   1455
         VariousPropertyBits=   671105051
         Size            =   "2566;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text33 
         Height          =   300
         Index           =   5
         Left            =   -68145
         TabIndex        =   25
         Top             =   1118
         Width           =   1575
         VariousPropertyBits=   671105051
         Size            =   "2778;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblFM2 
         Height          =   255
         Index           =   2
         Left            =   -71940
         TabIndex        =   99
         Top             =   1827
         Width           =   4170
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "7355;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label LblFM2 
         Height          =   255
         Index           =   1
         Left            =   7020
         TabIndex        =   98
         Top             =   653
         Width           =   1275
         VariousPropertyBits=   27
         Caption         =   "LblFM2"
         Size            =   "2249;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   300
         Index           =   2
         Left            =   1575
         TabIndex        =   16
         Top             =   2680
         Width           =   6795
         VariousPropertyBits=   671105051
         Size            =   "11986;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   300
         Index           =   1
         Left            =   1575
         TabIndex        =   15
         Top             =   2380
         Width           =   6795
         VariousPropertyBits=   671105051
         Size            =   "11986;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox Text9 
         Height          =   300
         Index           =   0
         Left            =   1575
         TabIndex        =   14
         Top             =   2080
         Width           =   6795
         VariousPropertyBits=   671105051
         Size            =   "11986;529"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         Caption         =   "¥Ó½Ð¤H1:"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   90
         Top             =   3003
         Width           =   855
      End
      Begin MSForms.Label Label27 
         Height          =   255
         Index           =   0
         Left            =   2715
         TabIndex        =   89
         Top             =   3003
         Width           =   5500
         VariousPropertyBits=   27
         Caption         =   "Label27"
         Size            =   "9701;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label29 
         Caption         =   "¥Ó½Ð¤H2:"
         Height          =   255
         Left            =   600
         TabIndex        =   88
         Top             =   3303
         Width           =   855
      End
      Begin VB.Label Label12 
         Caption         =   "¥Ó½Ð¤H3:"
         Height          =   255
         Left            =   600
         TabIndex        =   87
         Top             =   3603
         Width           =   855
      End
      Begin VB.Label Label32 
         Caption         =   "¥Ó½Ð¤H4:"
         Height          =   255
         Left            =   600
         TabIndex        =   86
         Top             =   3903
         Width           =   855
      End
      Begin VB.Label Label34 
         Caption         =   "¥Ó½Ð¤H5:"
         Height          =   255
         Left            =   600
         TabIndex        =   85
         Top             =   4193
         Width           =   855
      End
      Begin MSForms.Label Label27 
         Height          =   255
         Index           =   1
         Left            =   2715
         TabIndex        =   84
         Top             =   3303
         Width           =   5500
         VariousPropertyBits=   27
         Caption         =   "Label27"
         Size            =   "9701;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   255
         Index           =   2
         Left            =   2715
         TabIndex        =   83
         Top             =   3603
         Width           =   5500
         VariousPropertyBits=   27
         Caption         =   "Label27"
         Size            =   "9701;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   255
         Index           =   3
         Left            =   2715
         TabIndex        =   82
         Top             =   3903
         Width           =   5500
         VariousPropertyBits=   27
         Caption         =   "Label27"
         Size            =   "9701;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label Label27 
         Height          =   255
         Index           =   4
         Left            =   2730
         TabIndex        =   81
         Top             =   4193
         Width           =   5505
         VariousPropertyBits=   27
         Caption         =   "Label27"
         Size            =   "9710;450"
         FontName        =   "·s²Ó©úÅé-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "©Ó¿ì¤H:"
         Height          =   180
         Left            =   5010
         TabIndex        =   75
         Top             =   690
         Width           =   585
      End
      Begin VB.Label lblCP19 
         AutoSize        =   -1  'True
         Caption         =   "°h¶Oª÷ÃB:"
         Height          =   180
         Left            =   5010
         TabIndex        =   74
         Top             =   390
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H³¡ªù(¤é):"
         Height          =   180
         Left            =   -74760
         TabIndex        =   73
         Top             =   1521
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   6
         Left            =   3030
         TabIndex        =   71
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "±M§QÅv¬O§_¦s¦b          (Y/N)"
         Height          =   180
         Index           =   2
         Left            =   240
         TabIndex        =   70
         Top             =   990
         Width           =   2145
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "®×¥ó¥Ø«e­ã»é:             (1:­ã , 2:»é)"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   69
         Top             =   690
         Width           =   2595
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹2:"
         Height          =   180
         Index           =   0
         Left            =   -74760
         TabIndex        =   61
         Top             =   3060
         Width           =   2115
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "¹êÅé°Æ¥»¦¬¨ü¤H©¼©Ò®×¸¹1:"
         Height          =   180
         Index           =   1
         Left            =   -74760
         TabIndex        =   60
         Top             =   2490
         Width           =   2115
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "¹êÅé°Æ¥»Ápµ¸¤H:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   59
         Top             =   2207
         Width           =   1305
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "¹êÅé°Æ¥»¦¬¨ü¤H:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   58
         Top             =   1864
         Width           =   1305
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤é):"
         Height          =   180
         Left            =   -69240
         TabIndex        =   57
         Top             =   1178
         Width           =   972
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(­^):"
         Height          =   180
         Left            =   -72000
         TabIndex        =   56
         Top             =   1178
         Width           =   972
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H2(¤¤):"
         Height          =   180
         Left            =   -74760
         TabIndex        =   55
         Top             =   1178
         Width           =   972
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤é):"
         Height          =   180
         Left            =   -69240
         TabIndex        =   54
         Top             =   835
         Width           =   972
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(­^):"
         Height          =   180
         Left            =   -72000
         TabIndex        =   53
         Top             =   835
         Width           =   972
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Ápµ¸¤H1(¤¤):"
         Height          =   180
         Left            =   -74760
         TabIndex        =   52
         Top             =   835
         Width           =   972
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "«È¤á®×¥ó®×¸¹:"
         Height          =   180
         Left            =   -74760
         TabIndex        =   51
         Top             =   492
         Width           =   1128
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "®×¥ó¦WºÙ(¤é):"
         Height          =   180
         Left            =   210
         TabIndex        =   50
         Top             =   2740
         Width           =   1065
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "®×¥ó¦WºÙ(­^):"
         Height          =   180
         Left            =   210
         TabIndex        =   49
         Top             =   2440
         Width           =   1065
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "®×¥ó¦WºÙ(¤¤):"
         Height          =   180
         Left            =   210
         TabIndex        =   48
         Top             =   2140
         Width           =   1065
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "¬O§_³¬¨÷:                    (Y:³¬¨÷)"
         Height          =   180
         Left            =   240
         TabIndex        =   47
         Top             =   1590
         Width           =   2370
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "¾÷Ãö¤å¸¹:"
         Height          =   180
         Left            =   240
         TabIndex        =   46
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "¥Ó½Ð®×®Ö­ã¤é:"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   45
         Top             =   390
         Width           =   1125
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      IntegralHeight  =   0   'False
      ItemData        =   "frm06010602_3.frx":0054
      Left            =   1110
      List            =   "frm06010602_3.frx":0061
      Style           =   2  '³æ¯Â¤U©Ô¦¡
      TabIndex        =   64
      Top             =   815
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "µ²§ô(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   7980
      TabIndex        =   33
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "½T©w(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   5928
      TabIndex        =   31
      Top             =   15
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "¦^«eµe­±(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   6756
      TabIndex        =   32
      Top             =   15
      Width           =   1200
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2670
      MaxLength       =   2
      TabIndex        =   38
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   2430
      MaxLength       =   1
      TabIndex        =   37
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1590
      MaxLength       =   6
      TabIndex        =   36
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   1110
      MaxLength       =   3
      TabIndex        =   35
      Top             =   480
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   300
      Left            =   4350
      TabIndex        =   34
      Top             =   480
      Width           =   1575
   End
   Begin MSForms.Label LblFM2 
      Height          =   255
      Index           =   0
      Left            =   1770
      TabIndex        =   97
      Top             =   838
      Width           =   7005
      VariousPropertyBits=   27
      Caption         =   "LblFM2"
      Size            =   "12356;450"
      FontName        =   "·s²Ó©úÅé-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label31 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð¤é:"
      Height          =   255
      Left            =   6390
      TabIndex        =   92
      Top             =   1150
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Index           =   8
      Left            =   7050
      TabIndex        =   91
      Top             =   1150
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   252
      Index           =   4
      Left            =   5016
      TabIndex        =   68
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Index           =   3
      Left            =   1110
      TabIndex        =   67
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Index           =   2
      Left            =   4050
      TabIndex        =   66
      Top             =   1150
      Width           =   480
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      Height          =   255
      Index           =   1
      Left            =   1110
      TabIndex        =   65
      Top             =   1150
      Width           =   480
   End
   Begin VB.Label Label30 
      AutoSize        =   -1  'True
      Caption         =   "«áÄò­ã»éÂ²³æ³ø§i:"
      Height          =   180
      Left            =   3396
      TabIndex        =   63
      Top             =   1440
      Width           =   1488
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "¨Ó¨ç¦¬¤å¤é:"
      Height          =   255
      Index           =   1
      Left            =   150
      TabIndex        =   62
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "¦¬¤å¸¹:"
      Height          =   255
      Left            =   3390
      TabIndex        =   43
      Top             =   1150
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "®×¥ó©Ê½è:"
      Height          =   255
      Left            =   150
      TabIndex        =   42
      Top             =   1150
      Width           =   765
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "¥»©Ò®×¸¹:"
      Height          =   255
      Left            =   150
      TabIndex        =   41
      Top             =   503
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "¥Ó½Ð®×¸¹:"
      Height          =   255
      Left            =   3390
      TabIndex        =   40
      Top             =   503
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "±M§Q¦WºÙ:"
      Height          =   255
      Left            =   150
      TabIndex        =   39
      Top             =   838
      Width           =   765
   End
End
Attribute VB_Name = "frm06010602_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 §ï¦¨Form2.0 ; grdDataList2§ï¦r«¬=·s²Ó©úÅé-ExtB¡BLabel3(0)=>LblFM2(0)¡BLabel3(7)=>LblFM2(1)¡BLabel3(5)=>LblFM2(2)¡BLabel27(index)¡BText33(index)
'Memo By Morgan 2012/12/10 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo by Morgan2010/12/27 ¥Ó½Ð®×¸¹Äæ¤w­×§ï
'2010/12/6 memo by sonia ­û¤u½s¸¹Äæ¤w­×§ï
'Memo by Morgan2010/8/12 ¤é´ÁÄæ¤w­×§ï
Option Explicit

Dim strReceiveNo As String, strTemp As String, strKind As String, cp(10) As String
'Modify by Morgan 2006/10/20 §ï°ÊºA
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim intWhere As Integer, strSales As String
' 90.06.27 modify by louis ®×¥ó©Ê½è
Dim m_CP10 As String
' 92.1.19 add by sonia
Dim m_CP14 As String
'Add By Cheng 2002/01/28
Dim m_NewReceiveNo As String 'Á`¦¬¤å¸¹
'Add by Morgan 2004/6/23
Dim stNP07 As String, stNP08 As String, stNP09 As String    '·s¥Ó½Ð®×»âÃÒ´Á­­
Dim m_BSheetNo As String 'Add by Morgan 2007/4/4 BÃþ±µ¬¢³æ¸¹

Dim m_928Upd As Boolean '¬O§_§ó·s­«·s©e¥ô­ã»é
Dim m_928CP09 As String '­«·s©e¥ô¦¬¤å¸¹

'Add by Morgan 2009/10/2
Dim m_bPrintFlowSheet As Boolean '¬O§_¦C¦L¬yµ{ªí
Dim m_bAddAcc1k0 As Boolean '¬O§_·s¼W½Ð´Ú³æ
Dim m_bNoDN As Boolean '°h¶O¬O§_½Ð´Ú
'Added by Morgan 2012/12/13
Dim m_bDivSugTextAlert As Boolean 'ªì¼f®Ö­ã¤À³Î«ØÄ³©w½Z®Ö­ã´£¿ô
Dim m_EditDivSugText As String '©|¥¼­×§ï¤À³Î«ØÄ³°T®§ Added by Morgan 2020/2/27
Dim m_PA162 As String
Dim m_bNewGrant As Boolean '¬O§_ªì¼f®Ö­ã Added by Morgan 2013/10/29
Dim m_bAgainGrant As Boolean 'Added by Lydia 2019/07/30 µo©ú¦A¼f®Ö­ã
Dim m_strMemo As String '¤À³Î´Á­­³Æµù Added by Morgan 2013/10/29
Dim m_926strMemo As String 'Added by Lydia 2022/08/02 ®Ö¹ï¤w­ã±M§Q³Æµù(¥u¥Î¨Ó¦C¦L)
Dim mAddSCalendar As Boolean 'Added by Lydia 2015/12/31 ¬O§_·s¼W¦æ¨Æ¾ä
Dim m_bHasDivCase As Boolean '¬O§_¦³¤À³Î®× 'Added by Morgan 2019/10/7
 
'Added by Morgan 2017/5/10 ¹q¤l¤½¤å
Public m_DocWord As String
Public m_DocNo As String
Public m_DocDate As String
Public m_AppNo As String
Public m_DeadLine As String
'end 2017/5/10
'Added by Morgan 2017/8/17
Dim m_bIsDualInvWithNoSelInform As Boolean '¬O§_¤@®×¨â½Ðµo©ú®×¥BµL¾Ü¤@¨ç
Dim m_bAdd1919 As Boolean '¬O§_·s¼W1919(«DÄÝ¬Û¦P³Ð§@)¨Ó¨ç
Dim m_st1919CP09 As String '1919¦¬¤å¸¹
Dim m_stUPA(4) As String '¤@®×¨â½Ð·s«¬®×¸¹
'Added by Lydia 2017/08/21 ¦æ¨Æ¾ä·s¼W2¦¸¶Ê¤À³Î´Á­­
Dim m_1stDate As String  '²Ä1¦¸¤À³Î´Á­­
Dim m_2ndDate As String  '²Ä2¦¸¤À³Î´Á­­
Dim bolTmp As Boolean 'Added by Lydia 2019/03/06
Dim m_bMiddleCase As Boolean '¤¤¶¡¨Ó©Ò®×¥ó Added by Morgan 2019/12/31
Dim m_pAgreeOnDate As String 'Modify By Sindy 2021/4/26

Private Sub cmdMod_Click()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   strSql = "SELECT * FROM ChangeEvent " & _
            "WHERE CE01 = '" & strReceiveNo & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount <= 0 Then
      rsTmp.Close
      strMsg = "µLÅÜ§ó¨Æ¶µ°O¿ý"
      strTit = "¸ê®ÆÀË®Ö"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   rsTmp.Close
   
   DisplayNextForm
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub DisplayNextForm()
   frm06010602_4.SetData 0, pa(1), True
   frm06010602_4.SetData 1, pa(2), False
   frm06010602_4.SetData 2, pa(3), False
   frm06010602_4.SetData 3, pa(4), False
   frm06010602_4.SetData 5, strReceiveNo, False
   Me.Hide
   frm06010602_4.Show
   frm06010602_4.QueryData
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim aKind As String 'Add by Lydia 2014/11/26

If frm06010602_2.Text6 = "1" Then
    'Added by Lydia 2015/10/02 ³¡¥÷®×¥ó©Ê½è¤§®Ö­ã1001§ï¬°®Öµo1008
    If InStr(Patent1001Display, m_CP10) > 0 Then
        aKind = "1008"
    Else
        aKind = "1001"  '®Ö­ã
    End If
    'end 2015/10/02

'Modified by Lydia 2015/01/05
'Else
'   aKind = "1503" '§ïÅÜ­ì³B¤À
End If

   Select Case Index
      Case 0
         ' 91.01.28 modify by louis
         If strKind >= "101" And strKind <= "105" Then
            If IsEmptyText(Text6) Then
               MsgBox "½Ð¿é¤J¥Ó½Ð®×®Ö­ã¤é", vbOKOnly + vbCritical, "ÀË®Ö¸ê®Æ"
               Exit Sub
            End If
         End If
         If Mid(strKind, 1, 1) = "3" Then
            If IsEmptyText(Text6) Then
               MsgBox "½Ð¿é¤J¥Ó½Ð®×®Ö­ã¤é", vbOKOnly + vbCritical, "ÀË®Ö¸ê®Æ"
               Exit Sub
            End If
         End If
         
         'Add by Morgan 2009/10/13
         If txtCP19.Visible = True Then
            If txtCP19 = "" Then
               MsgBox "½Ð¿é¤J°h¶Oª÷ÃB¡I", vbExclamation
               txtCP19.SetFocus
               Exit Sub
            ElseIf Val(txtCP19.Tag) > 0 And Val(txtCP19) <> Val(txtCP19.Tag) Then
               If MsgBox("¥»¦¸¿é¤Jªº°h¶Oª÷ÃB»P¥Ó½Ð®Ñªº¤£¦P¬O§_­nÄ~Äò¡H", vbYesNo + vbDefaultButton2) = vbNo Then
                  txtCP19.SetFocus
                  Exit Sub
               End If
            End If
         End If
         
         'Add By Cheng 2002/05/22
         '­«·sÀË¬dÄæ¦ì¦³®Ä©Ê
         If TxtValidate = False Then Exit Sub
         
         '2006/3/29 ADD BY SONIA ¤wµo¤å½Ð¨D­±¸ß407¦ýµL³qª¾­±¸ß1401¥BµL­±¸ß408¤§¦¬¤åªÌ´£¥Ü°T®§
         CHECKFCP407 pa(1), pa(2), pa(3), pa(4)
         '2006/2/29 END
         
         'Added by Morgan 2013/10/29
         m_strMemo = ""
         m_926strMemo = "" 'Added by Lydia 2022/08/02
         m_1stDate = "": m_2ndDate = "" 'Added by Lydia 2017/08/21
         'Modified by Lydia 2019/07/30 +¦A¼f®Ö­ã
         'If m_bNewGrant Or m_bAgainGrant Then 'ªì¼f®Ö­ã
         If m_bNewGrant = True Or m_bAgainGrant = True Then 'ªì¼f®Ö­ã+¦A¼f®Ö­ã
            'Modified by Lydia 2014/11/26 ±N³Æµù³]¬°¦@¥Îªº©T©w³ÆµùÀÉApprovalMemo2
'            intI = 0
'            Select Case Left(pa(75) & "000", 8)
'            Case "Y4514900", "Y4745300"
'               If Left(pa(26) & "000", 8) = "X4514900" Then
'                  intI = 1
'               End If
'            'Added by Morgan 2014/3/6 +Y51551,Y47901 --Susan
'            'Modified by Moragn 2014/10/9 +Y52798 --¦¿¦p¥É
'            Case "Y5155100", "Y4790100", "Y5279800"
'               intI = 1
'            End Select
'
'            If intI = 1 Then
           '¦sÀÉ«eMessage (ªì¼f®Ö­ã) 'Memo by Lydia 2019/07/30 »PSharon½T»{: ¦A¼f®Ö­ã¤]­n§ìªì¼f®Ö­ãªº³Æµù
           'Modified by Lydia 2015/01/05 §ï¬°¤Ä¿ï°T®§ºØÃþ ,ªì¼f=4
           ' m_strMemo = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), "Y")
            'Modified by Lydia 2019/03/06 ³vµ§§PÂ_Y¥N²z¤H+X¥Ó½Ð¤H1~5
            'm_strMemo = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), "4")
            strExc(1) = "": strExc(2) = ""
            'Modified by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡G­×§ï¤@¯ë³Æµù¡B®Ö¹ï¤w­ã³Æµù¬°½Æ¼Æ·s³W«h
            'For intI = 0 To 4
            '     If pa(26 + intI) <> "" Then
            '        strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)), "4", bolTmp)
            '        If strExc(1) <> "" Then
            '            If bolTmp = True Then '­Ó®×³Æµù
            '               m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
            '               Exit For
            '            ElseIf strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
            '               If m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0) Then
            '                    m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
            '               End If
            '               strExc(2) = strExc(2) & strExc(1) & "||" '§PÂ_¬O§_¦³­«½Æ³Æµù (ªì¼f®Ö­ãªºÀË¬d)
            '            End If
            '        End If
            '     End If
            'Next intI
            ''end 2019/03/06
            strExc(1) = PUB_GetApprMemo2("4", pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
            If strExc(1) <> "" And InStr(m_strMemo & ",", strExc(1)) = 0 Then
                m_strMemo = m_strMemo & strExc(1)
            End If
            'end 2022/08/02
            
            'Added by Lydia 2019/07/30 ¦]108.11.1­×ªk¤À³ÎºÞ¨î´Á­­³]©w
            '1. ©ó108.8.1¦¬¨ì¤§®Ö­ã¨ç¡G
            '¡@1.1. µo©úªì¼f®Ö­ã¡Gºû«ù­ì¦³³]©w¤§¤À³Î´Á­­
            '¡@1.2. µo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã¡G­ì¦³³]©w¤À³Î´Á­­¤§«È¤á½s¸¹¡A¼W¥[±±ºÞ¦æ¨Æ¾ä´Á­­¡A­ì«h·Óªì¼f®Ö­ã¡A´Á­­¬°¦¬¨ì®Ö­ã¨ç«á¢²­Ó¤ë´Á­­¡A¨Ã±a³Æµù¦Ü³qª¾§i­ã¤§¶i«×³Æµù¡C
            '2. ©ó108.10.1¦¬¨ì¤§®Ö­ã¨ç¡Gµo©úªì¼f®Ö­ã¡Bµo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã¡G¬Ò³]©w¦¬¨ì®Ö­ã¨ç«á¢²­Ó¤ë´Á­­¡C
            strExc(0) = DBDATE(Label3(3))
            If strExc(0) >= "20191001" Or _
                (strExc(0) >= "20190801" And m_bAgainGrant = True) Then  '108.8.1¦¬¨ì¤§µo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã
                 strExc(1) = CompWorkDay(1, CompDate(2, -7, CompDate(1, 3, strExc(0))), 1) '²Ä¤@¦¸¶Ê¤À³Î(ªk­­-7¤Ñ)
                 strExc(2) = CompWorkDay(1, CompDate(2, -1, CompDate(1, 3, strExc(0))), 1) '²Ä¤G¦¸¶Ê¤À³Î(ªk­­-1¤Ñ)
            Else  'ÂÂªk
                 strExc(1) = CompWorkDay(1, CompDate(2, 23, strExc(0)), 1) '²Ä¤@¦¸¶Ê¤À³Î(¦¬¤å¤é¦A¥[23¤é)
                 strExc(2) = CompWorkDay(1, CompDate(2, 29, strExc(0)), 1) '²Ä¤G¦¸¶Ê¤À³Î(¦¬¤å¤é¦A¥[29¤é)
            End If
            
            'Modified by Lydia 2017/08/21 ¼W¥["¦æ¨Æ¾ä¤wºÞ¨î2¦¸¶Ê¤À³Î´Á­­"
            'If Len(m_strMemo) > 0 And InStr(m_strMemo, "½ÐºÞ¨î¶Ê¤À³Î´Á­­") > 0 Then
            If Len(m_strMemo) > 0 And InStr(m_strMemo, "¦æ¨Æ¾ä¤wºÞ¨î2¦¸¶Ê¤À³Î´Á­­") > 0 Then
               'Modified by Lydia 2019/07/30 ¦]108.11.1­×ªk¤À³ÎºÞ¨î´Á­­­×§ï
'               strExc(1) = DBDATE(Label3(3))
'               'Modified by Lydia 2017/10/12 ­Yªâ±j½Õ¶Ê¤À³Î´Á­­¬°«D°²¤é,»P¤@¯ë¶Ê¤À³Î¦¬¤å¤é+23¤é¤£¦P
'               'm_1stDate = CompDate(2, 23, strExc(1))
'               m_1stDate = CompWorkDay(1, CompDate(2, 23, strExc(1)), 1)
'               '²Ä2¦¸¶Ê¤À³Î´Á­­ªº­pºâ¬°®Ö­ã¨çªº¥»©Ò¦¬¤å¤é¦A¥[29¤é, ­Y¹J°²¤é«h´£«e¦Ü«e¤@¤u§@¤é
'               m_2ndDate = CompWorkDay(1, CompDate(2, 29, strExc(1)), 1)
               m_1stDate = strExc(1)
               m_2ndDate = strExc(2)
               'end 2019/07/30
               m_strMemo = m_strMemo & ": " & ChangeWStringToTDateString(m_1stDate) & " ¤Î " & ChangeWStringToTDateString(m_2ndDate)
               '·s¼W¦æ¨Æ¾ä«á,¤~¼u°T®§
            'Modified by Lydia 2017/10/12 ¤@¨Ö²£¥Í¦æ¨Æ¾ä
            'ElseIf Len(m_strMemo) > 0 And InStr(m_strMemo, "½ÐºÞ¨î¶Ê¤À³Î´Á­­") > 0 Then
            ElseIf Len(m_strMemo) > 0 And InStr(m_strMemo, "¦æ¨Æ¾ä¤wºÞ¨î¶Ê¤À³Î´Á­­") > 0 Then
            'end 2017/08/21
               'Modified by Lydia 2019/07/30 ¦]108.11.1­×ªk¤À³ÎºÞ¨î´Á­­­×§ï
'               strExc(1) = DBDATE(Label3(3))
'               'Modified by Lydia 2017/10/12 »P±Ó²ú·¾³q: ¶Ê¤À³Î´Á­­­Y¹J°²¤é«h´£«e¦Ü«e¤@¤u§@¤é,¤@¨Ö²£¥Í¦æ¨Æ¾ú (¤ñ·Ó¤W¦CºÞ¨î2¦¸¶Ê¤À³Î´Á­­)
'               'strExc(2) = CompDate(2, 23, strExc(1))
'               ''m_strMemo = "½ÐºÞ¨î¶Ê¤À³Î´Á­­ " & ChangeWStringToTDateString(strExc(2)) & " !!!"
'               'm_strMemo = m_strMemo & " " & ChangeWStringToTDateString(strExc(2)) & " !!!"
'               m_1stDate = CompWorkDay(1, CompDate(2, 23, strExc(1)), 1)
               m_1stDate = strExc(1)
               'end 2019/07/30
               m_strMemo = m_strMemo & ": " & ChangeWStringToTDateString(m_1stDate)
               'end 2017/10/12
               'MsgBox m_strMemo, vbExclamation 'Remove by Lydia 2017/10/16 ³Ì«á·s¼W§¹¦æ¨Æ¾ä¤~¼u°T®§
            ElseIf Len(m_strMemo) > 0 Then
               MsgBox m_strMemo, vbExclamation
               'end  'Modified by Lydia 2014/11/26
            End If
'            End If
         End If
         'end 2013/10/29
         
         'Added by Morgan 2013/12/11
         'Modified by Morgan 2014/3/7 +Y51306
         'Modified by Morgan 2014/6/18 +Y28043
         'Modified by Morgan 2014/8/6 +Y52061
         'Modified by Morgan 2014/8/29 +Y47453--§d±mµÙ
         'Modified by Morgan 2014/10/8 +Y51622--§d±mµÙ
         'Modified by Lydia 2014/11/26 ±N³Æµù³]¬°¦@¥Îªº©T©w³ÆµùÀÉApprovalMemo2
'         Select Case Left(pa(75) & "000", 8)
'         Case "Y2004900", "Y5130600", "Y2804300", "Y5206100", "Y4745300", "Y5162200"
'            'Modified by Morgan 2014/8/6 »P²QµØ½T»{¹L³£¥]§t­ì¤å¤ÎÄ¶¤å¬G§ï¥Î¤U­±ªº¤å¥y
'            'strExc(1) = "§i­ã®É¡A¶·¦P®ÉFAX®Ö­ã¨ç!!!"
'            strExc(1) = "§i­ã®É¡A¶·¦P®ÉFAX®Ö­ã³qª¾­ì¤å¤ÎÄ¶¤å!!!"
'            'end 2014/8/6
'            MsgBox strExc(1), vbExclamation
'            If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'            m_strMemo = m_strMemo & strExc(1)
'         End Select
         '¦sÀÉ«eMessage,®×¥ó©Ê½è©T©w¬°1001
         '­nª`·NCU122(FCP¬O§_®Ö¹ï¤w­ã±M§Q)=N,±N¤£·|²£¥Í®Ö¹ï¤w­ã±M§Q¦¬¤å³æ(BÃþ³æ)->¤£¦C¦L
         'Modified by Lydia 2015/01/05 §ï¬°¤Ä¿ï°T®§ºØÃþ ,¤@¯ë=1
         'strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)))
         'Modified by Lydia 2019/03/06 ³vµ§§PÂ_Y¥N²z¤H+X¥Ó½Ð¤H1~5
'         strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), "1")
'         If Len(strExc(1)) > 0 Then
'            MsgBox strExc(1), vbExclamation
'            If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'            m_strMemo = m_strMemo & strExc(1)
'         End If
'          'end  'Modified by Lydia 2014/11/26
'         'end 2013/12/11
         
         'Modified by Lydia 2019/08/01 ¥u°w¹ï¥Ó½Ð®×¤§®Ö­ã³Æµù,½Ð±Æ°£«D¥Ó½Ð®×(¦pÅÜ§ó,Åý»P,§ó§ï,§ó¥¿¡K)¤§®Ö­ã³Æµù
         'If Frame1.Visible = False Then 'Added by Lydia 2019/07/10 §ó§ï®Ö­ã¤£¥Î§ì®Ö­ã³Æµù
         If InStr(NewCasePtyList & ",107", m_CP10) > 0 Or Left(m_CP10, 1) = "3" Then '·s¥Ó½Ð®×+¤À³Î307+¦A¼f107+§ï½Ð3¶}ÀY
            strExc(1) = "": strExc(2) = ""
            'Modified by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡G­×§ï¤@¯ë³Æµù¡B®Ö¹ï¤w­ã³Æµù¬°½Æ¼Æ·s³W«h
            'For intI = 0 To 4
            '    If pa(26 + intI) <> "" Then
            '        strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)), "1", bolTmp)
            '        If strExc(1) <> "" Then
            '           If bolTmp = True Then '­Ó®×³Æµù
            '              m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
            '              strExc(2) = strExc(2) & strExc(1) 'Added by Morgan 2020/3/4
            '              Exit For
            '           ElseIf strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
            '              If m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0) Then
            '                   m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
            '              End If
            '              strExc(2) = strExc(2) & strExc(1) & "||" '§PÂ_¬O§_¦³­«½Æ³Æµù (¤@¯ë®Ö­ãªºÀË¬d)
            '           End If
            '        End If
            '    End If
            'Next intI
            'If strExc(2) <> "" Then MsgBox Replace(strExc(2), "||", vbCrLf), vbExclamation
            strExc(1) = PUB_GetApprMemo2("1", pa(1) & pa(2) & pa(3) & pa(4), aKind, ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
            If strExc(1) <> "" And InStr(m_strMemo & ",", strExc(1)) = 0 Then
                m_strMemo = m_strMemo & IIf(Len(m_strMemo) = 0, "", vbCrLf) & strExc(1)
            End If
            If strExc(1) <> "" Then MsgBox strExc(1), vbExclamation, "®Ö­ã¨ç³Æµù"
            'end 2022/08/02
         End If 'end 2019/07/10
         'end 2019/03/06
            
         'Added by Morgan 2017/8/17
         '¤@®×¨â½Ð¦³µL«DÄÝ¬Û¦P³Ð§@±±¨î
         m_bIsDualInvWithNoSelInform = False
         m_st1919CP09 = ""
         'Modified by Morgan 2017/11/29 +§PÂ_¥Ó½Ð©Î¦A¼fµ{§Çªº®Ö­ã
         If pa(8) = "1" And (m_CP10 = "101" Or m_CP10 = "107") Then
            If PUB_IsDualApply(pa, m_stUPA) Then
               'Modified by Morgan 2019/7/17
               'If PUB_ChkCPExist(pa(), "1232") = False And PUB_ChkCPExist(pa(), "239", 2) = False Then
               '°ò¥»ÀÉ³]©w©ñ±óµo©ú
               If pa(60) = "N" Then
                  MsgBox "¾Ü¤@¥Ó´_¿ï¾Ü¡i©ñ±óµo©ú¡j¡A½Ð½T»{¡G" & vbCrLf & vbCrLf & _
                        "1.½T©w©ñ±óµo©ú --> Ápµ¸IPO" & vbCrLf & vbCrLf & _
                        "2.½T©w©ñ±ó·s«¬ --> ­×§ï°ò¥»ÀÉ¬O§_©ñ±ó·s«¬¬°""Y""" & vbCrLf & vbCrLf & _
                        "( Y: ©ñ±ó·s«¬  N: ©ñ±óµo©ú  ªÅ: ³£¤£©ñ±ó,2ªÌ¦s¦b )", vbExclamation
                  Exit Sub
               'µL¾Ü¤@¥Ó´_µo¤å or ¥¼¿ï¾Ü©ñ±óµo©ú©Î·s«¬
               ElseIf PUB_ChkCPExist(pa(), "239", 2) = False Or pa(60) = "" Then
               'end 2019/7/17
               
                  m_bIsDualInvWithNoSelInform = True
                  intI = MsgBox("®Ö­ã¨ç¬O§_¦³""«DÄÝ¬Û¦P³Ð§@""¡H", vbYesNoCancel + vbDefaultButton3 + vbExclamation, "¤@®×¨â½Ðµo©ú®×®Ö­ã´£¿ô")
                  If intI = vbYes Then
                     m_bAdd1919 = True
                  ElseIf intI = vbNo Then
                     'Added by Morgan 2019/7/17
                     MsgBox "½Ð¤uµ{®v½T»{¡G" & vbCrLf & vbCrLf & _
                           "1.¤@®×¤G½ÐµL¾Ü¤@¨ç,½Ð¤uµ{®v»P´¼¼z§½³sµ¸" & vbCrLf & vbCrLf & _
                           "2.­Y¥»©Ò¤w¾Ü¤@,½Ð³qª¾µ{§Ç¤H­û,¤º³¡¦¬¤å""¾Ü¤@¥Ó´_""", vbExclamation
                     'end 2019/7/17
                     m_bAdd1919 = False
                  Else
                     Exit Sub
                  End If
               End If
            End If
         End If
         'end 2017/8/17
         
         'Add by Sindy 2021/11/22 ÀË¬dµe­±¤Wªºª«¥ó¬O§_§t¦³Unicode¤å¦r
         If PUB_ChkUniText(Me, True, True) = False Then
            Exit Sub
         End If
         
         'Add by Morgan 2004/7/28
         '¥[º|¤æ
         Screen.MousePointer = vbHourglass
         
         If FormSave = False Then
            Screen.MousePointer = vbDefault
            MsgBox "¦sÀÉ¥¢±Ñ¡A½Ð¬¢¨t²ÎºÞ²z­û !", vbCritical
            Exit Sub
         End If
         Screen.MousePointer = vbDefault
         
         'Added by Morgan 2012/12/4
         'Modified by Morgan 2019/10/7
         'If m_bDivSugTextAlert = True Then
         '   MsgBox "¦¹®×­n¥t¨ç³qª¾ªì¼f®Ö­ã«á¤À³Î¡A½Ð±N¨÷©vÂà¥æ¤uµ{®v¡I", vbInformation
         If m_bHasDivCase Then
            MsgBox "¦¹®×¤w´£¤À³Î®×½Ð¤uµ{®v½T»{¬O§_¤´¥[µù¤À³Î«ØÄ³¡I", vbInformation
            
         ElseIf m_bDivSugTextAlert Then
            'Added by Morgan 2019/12/27
            If pa(162) = "" Then
               MsgBox "¨÷°h¤uµ{®v§PÂ_¬O§_¥[µù®Ö­ã¤À³Î«ØÄ³¡I", vbInformation
            Else
            'end 2019/12/27
               'Added by Morgan 2020/2/27
               If m_EditDivSugText <> "" Then
                  MsgBox "¨÷°h¤uµ{®v­×§ï¤À³Î«ØÄ³¤º®e!", vbInformation
               Else
               'end 2020/2/27
               
                  MsgBox "¦¹®×­n¥[µù®Ö­ã¤À³Î«ØÄ³¡A½Ð±N¨÷©vÂà¥æ¤uµ{®v¡I", vbInformation
               End If 'Added by Morgan 2020/2/27
            End If
         'end 2019/10/7
         
            'PUB_SendMailCache 'Removed by Morgan 2019/7/17 ²¾¨ì¤U­±
         End If
         'end 2012/12/4
         
         PUB_SendMailCache 'Added by Morgan 2019/7/17

         'Added by Morgan 2013/11/21' 'Modified by Lydia 2014/11/26 ¦]¬°±ø¥ó¤£¦P,¤£¦C¤J¦@¦P³Æµù
         If Left(pa(75) & "000", 8) = "Y4827900" And InStr(NewCasePtyList & ",107", m_CP10) > 0 Then
            MsgBox "¥»¨÷»Ý°h¤uµ{®v·Ç³Æ¤w­ãªº­^¤å±M§Q½d³ò!!", vbInformation
         End If
         'end 2013/11/7
         
         'Added by Lydia 2015/12/31
         'Remark by Lydia 2019/07/09 ¨ú®ø´£¿ô
         'If mAddSCalendar Then
         '   MsgBox "¤À³Îªk©w´Á­­¤Î°h¤uµ{®v1st®Ö¹ï¤w­ã!!", vbInformation
         'End If
         'end 2015/12/31
         
         'Added by Lydia 2017/08/21 ´£¿ôFCPºÞ¨î¤H©MÂ¾¥N
         If Right(m_1stDate, 1) = "Y" And Right(m_2ndDate, 1) = "Y" Then
            MsgBox "¦æ¨Æ¾ä¤wºÞ¨î2¦¸¶Ê¤À³Î´Á­­: " & ChangeWStringToTDateString(Replace(m_1stDate, "Y", "")) & " ¤Î " & ChangeWStringToTDateString(Replace(m_2ndDate, "Y", "")), vbInformation
         End If
         'end 2017/08/21
         
         'Added by Lydia 2017/10/12 ´£¿ôFCPºÞ¨î¤H©MÂ¾¥N
         If Right(m_1stDate, 1) = "Y" And m_2ndDate = "" Then
            MsgBox "¦æ¨Æ¾ä¤wºÞ¨î¶Ê¤À³Î´Á­­: " & ChangeWStringToTDateString(Replace(m_1stDate, "Y", "")), vbInformation
         End If
         'end 2017/08/21
         
         'Added by Lydia 2019/05/28 ¿é¤J·s¥Ó½Ð®×¡B§ï½Ð®×¡B¤À³Î¤§®Ö­ã¨ç®É¡A½Ð±±ºÞ¤U¤@µ{§Ç¥¼§¹¦¨©Î¶i«×ÀÉ¦³¥¼µo¤å¤§µ{§Ç¡A¼u´£¿ô:¤U¤@µ{§Çor ¶i«×ÀÉ¥¼§¹¦¨,½Ð½T»{¬O§_Äò¿ì
                                                '±Æ°£®Ö­ã¿é¤J²£¥Íªº¦¬¤å¸¹©M¤U¤@µ{§Ç
         If InStr(NewCasePtyList, m_CP10) > 0 Or Mid(m_CP10, 1, 1) = "3" Then
              'Modified by Lydia 2019/06/17 ¤U¤@µ{§Ç±Æ°£¶Ê¼f=> ¨q¬Â¡G­ç°£¤U¤@µ{§Ç¬°µ{§ÇºÞ±±¤§®×¥ó©Ê½è
              'strSql = "select '1' ord1,np01 as pno,nvl(cpm03,cpm04) cpm0304 from nextprogress,casepropertymap " & _
                           "where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null " & _
                           "and np01 not in (" & GetAddStr(m_NewReceiveNo & "," & m_BSheetNo) & ") and np02=cpm01(+) and np07=cpm02(+) "
              strSql = "select '1' ord1,np01 as pno,nvl(cpm03,cpm04) cpm0304 from nextprogress,casepropertymap " & _
                           "where np02='" & pa(1) & "' and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "' and np06 is null and np07 not in (" & PAnp07NotIn & ")" & _
                           "and np01 not in (" & GetAddStr(m_NewReceiveNo & "," & m_BSheetNo) & ") and np02=cpm01(+) and np07=cpm02(+) "
              strSql = strSql & "Union All " & _
                           "select '2' ord1,cp09 as pno,nvl(cpm03,cpm04) cpm0304 from caseprogress,casepropertymap " & _
                           "where cp01='" & pa(1) & "'and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp158=0 and cp159=0 " & _
                           "and cp09 not in (" & GetAddStr(m_NewReceiveNo & "," & m_BSheetNo) & ")  and cp43 not in (" & GetAddStr(m_NewReceiveNo & "," & m_BSheetNo) & ") " & _
                           "and cp01=cpm01(+) and cp10=cpm02(+)   order by ord1 "
              intI = 1
              Set RsTemp = ClsLawReadRstMsg(intI, strSql)
              If intI = 1 Then
                  strExc(1) = "": strExc(2) = ""
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                       If "" & RsTemp.Fields("ord1") = "1" Then
                            strExc(1) = strExc(1) & "¡B" & RsTemp.Fields("cpm0304")
                       ElseIf "" & RsTemp.Fields("ord1") = "2" Then
                            strExc(2) = strExc(2) & "¡B" & RsTemp.Fields("cpm0304")
                       End If
                       RsTemp.MoveNext
                  Loop
                  If strExc(1) & strExc(2) <> "" Then
                      MsgBox IIf(strExc(1) <> "", "¤U¤@µ{§Ç¡G" & Mid(strExc(1), 2) & vbCrLf, "") & IIf(strExc(2) <> "", "¶i«×ÀÉ¡G" & Mid(strExc(2), 2) & vbCrLf, "") & "¥¼§¹¦¨¡A½Ð½T»{¬O§_Äò¿ì¡I", vbExclamation, "¤U¤@µ{§Ç©Î¶i«×ÀÉ¥¼§¹¦¨"
                  End If
              End If
         End If
         
         'Added by Morgan 2017/8/18
         '«DÄÝ¬Û¦P³Ð§@CÃþ±µ¬¢³æ
         If m_st1919CP09 <> "" Then
            'Modified by Lydia 2018/12/17 FCP®×CÃþ±µ¬¢³æ¦P®É¦C¦L¨Ã¥B¤W¶Ç¨ì¨÷©v°Ï
            'g_PrtForm001.PrintCForm m_st1919CP09
            'Modified by Lydia 2019/03/18 §ï¦¨¶}±ÒWord
            'g_PrtForm001.PrintCForm m_st1919CP09, , , True
            g_PrtForm001.PrintCFormNew m_st1919CP09, , , True
         End If
         'end 2017/8/18
         
        '­Y·s¼W¦Ü®×¥ó¶i«×ÀÉªºCÃþ¸ê®Æ, ­Y®×¥ó©Ê½è¬°
        '1002,1201~1203,1210~1212,1301~1307,1401,1502,1504~1507,
        '1801,1802,1805~1808,1903, «h¦C¦LCÃþ±µ¬¢°O¿ý³æ
'         'Add By Cheng 2002/01/25
'         '­Y·s¼Wªº®×¥ó¶i«×ÀÉªº®×¥ó©Ê½è¬°®Ö­ã
'         If frm06010602_2.Text6 = "1" Then
'            '¦C¦LCÃþ±µ¬¢°O¿ý³æ
'            g_PrtForm001.PrintCForm m_NewReceiveNo
'         End If
            'Add By Cheng 2003/04/03
            '­YÂI¿ïªº®×¥ó©Ê½èÄÝ©óª§Ä³µ{§Ç(8¶}ÀY)
            If Left(m_CP10, 1) = "8" Then
                '¦C¦LCÃþ±µ¬¢°O¿ý³æ
                'Modified by Lydia 2018/12/17 FCP®×CÃþ±µ¬¢³æ¦P®É¦C¦L¨Ã¥B¤W¶Ç¨ì¨÷©v°Ï
                'g_PrtForm001.PrintCForm m_NewReceiveNo
                'Modified by Lydia 2019/03/18 §ï¦¨¶}±ÒWord
                'g_PrtForm001.PrintCForm m_NewReceiveNo, , , True
                g_PrtForm001.PrintCFormNew m_NewReceiveNo, , , True
            End If
            'Add by Morgan 2007/4/4
            If m_BSheetNo <> "" Then
               '¦C¦LBÃþ±µ¬¢°O¿ý³æ
                'Modified by Lydia 2018/12/17 FCP®×CÃþ±µ¬¢³æ¦P®É¦C¦L¨Ã¥B¤W¶Ç¨ì¨÷©v°Ï
                'g_PrtForm001.PrintCForm m_BSheetNo, m_strMemo
                'Modified by Lydia 2019/03/18 §ï¦¨¶}±ÒWord
                'g_PrtForm001.PrintCForm m_BSheetNo, m_strMemo, , True
                'Modified by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡G¥t¥~°O¿ý926®Ö¹ï¤w­ã±M§Q³Æµù
                g_PrtForm001.PrintCFormNew m_BSheetNo, m_strMemo & IIf(Len(m_strMemo) = 0, "", vbCrLf) & m_926strMemo, , True
            End If

         'Add by Morgan 2007/10/24
         '­Y·s¼Wªº®×¥ó¶i«×ÀÉªº®×¥ó©Ê½è¬°§ïÅÜ­ì³B¤À
         If frm06010602_2.Text6 = "2" Then
            '¦C¦LCÃþ±µ¬¢°O¿ý³æ
            'Modified by Lydia 2018/12/17 FCP®×CÃþ±µ¬¢³æ¦P®É¦C¦L¨Ã¥B¤W¶Ç¨ì¨÷©v°Ï
            'g_PrtForm001.PrintCForm m_NewReceiveNo
            'Modified by Lydia 2019/03/18 §ï¦¨¶}±ÒWord
            'g_PrtForm001.PrintCForm m_NewReceiveNo, , , True
            g_PrtForm001.PrintCFormNew m_NewReceiveNo, , , True
         End If
         'end 2007/10/24
         
         'Add by Morgan 2009/10/12
         '¦L°h¶O¬yµ{ªí
         If m_bPrintFlowSheet = True Then
            PrintFlowSheet strReceiveNo, m_NewReceiveNo
         End If
         
         Unload frm06010602_2
         Unload Me
         
         'Added by Morgan 2017/5/10 ¹q¤l¤½¤å
         'frm06010602_1.Show
         If m_DocNo <> "" Then
            Unload frm06010602_1
            frm060119.GoNext
         Else
            frm06010602_1.Show
         End If
         'end 2017/5/10

      Case 1
         frm06010602_2.Show
         Unload Me
      Case 2
         Unload frm06010602_1
         Unload frm06010602_2
         Unload Me
   End Select
End Sub

Private Function FormSave() As Boolean
   Dim i As Integer, intStep As Integer, strTxt(1 To 20) As String, j As Integer
   Dim strCe(99) As String, bolChk As Boolean
   Dim NewReceiveNo As String, lMax As Long
   Dim strTmp(1 To 5) As String, strTemp1 As String
   Dim strNP08 As String
   Dim rsA As New ADODB.Recordset
   Dim StrSQLa As String
   Dim strCP12 As String, strCP13 As String, strCP14 As String 'Add by Morgan 2007/4/3
   Dim strBCP48 As String 'Add by Morgan 2007/5/3
   Dim strCP20 As String, strCP16 As String
   'Add by Morgan 2009/10/12
   Dim stA1k01 As String, stA1k03 As String, stA1k05 As String, stA1k11 As String, stA1k08 As String, strA1K27 As String, strA1K28 As String
   Dim stA1L05 As String, stA1L07 As String
   Dim dblUSRate As Double
   Dim strPrintCust As String
   Dim strDisc As String
   Dim str926CP14 As String 'Add by Morgan 2010/6/3
   Dim dblXRate As Double 'Added by Morgan 2011/12/21 ½Ð´Ú¹ô§O¹ï¥x¹ô¶×²v
   Dim st307Msg As String 'Added by Morgan 2012/11/13
   Dim strNewCP09 As String, strCP48 As String 'Add By Sindy 2017/1/11
   Dim strCP10 As String 'Added by Morgan 2017/5/10
   Dim strDivState As String, m_CP64 As String 'Add By Sindy 2017/6/6
   Dim strCP64 As String 'Added by Lydia 2019/05/23
   Dim strMailText As String 'Add By Sindy 2020/2/14
   Dim strMailSubject As String 'Added by Lydia 2021/02/02
   Dim strLang As String 'Added by Morgan 2021/9/24
   
   'Add by Morgan 2007/7/17
   If m_CP10 <> "928" Then
      m_928Upd = PUB_928Check(pa, m_928CP09)
   End If

 On Error GoTo ErrHnd
 
   cnnConnection.BeginTrans
   
 On Error GoTo CheckingErr
 
   strCP13 = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
   strCP12 = GetSalesArea(strCP13)

   'Add by Morgan 2007/7/17
   If m_928Upd = True And m_928CP09 <> "" Then
      PUB_928Update pa, m_928CP09
   End If
   'end 2007/7/17

   intStep = 1
   lMax = GetNextProgressNo  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  objPublicData.GetNextProgressNo
   strExc(0) = Empty
   
   strExc(0) = strExc(0) & "PA17='" & Me.Text10(1).Text & "',"
   'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
   '2013/10/24 MODIFY BY SONIA ¦A¥[¤J¨÷©v©Ê½è§PÂ_pa(23) = "1",P-083407ªº503¤£¥i§ó·s,§_«h«áÄò§ïÅÜ­ì³B¤À¤]¤£·|§ó·s
   If pa(23) = "1" And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "308") Or m_CP10 = "802" Or m_CP10 = "804") Then
      strExc(0) = strExc(0) & "PA16='" & Me.Text10(0).Text & "',"
      'Modify by Morgan 2004/12/1 ª§Ä³µ{§Ç¤£§ó·s°ò¥»ÀÉ­ã»é¤é
      'If IsEmptyText(Text6.Text) = False Then
      If IsEmptyText(Text6.Text) = False And Not (Val(m_CP10) >= 801 And Val(m_CP10) <= 805) Then
         strExc(0) = strExc(0) & "PA20=" & CNULL(TransDate(Text6, 2)) & ","
      End If
   End If
   
   'Added by Morgan 2023/2/23
   If m_CP10 = "415" And txt415Date <> "" Then
      strExc(0) = strExc(0) & "PA25=" & DBDATE(txt415Date) & ","
   End If
   'end 2023/2/23
   
   lMax = GetNextProgressNo  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  objPublicData.GetNextProgressNo
   
   'Modify by Morgan 2006/10/20 ¥[³sµ¸¤H³¡ªù(¤é)-->PA139
   strTxt(intStep) = "UPDATE PATENT SET " & strExc(0) & "PA05=" & CNULL(ChgSQL(Text9(0))) & ",PA06=" & CNULL(ChgSQL(Text9(1))) & ",PA07=" & CNULL(ChgSQL(Text9(2))) & _
      ",PA51=" & CNULL(ChgSQL(Text33(0))) & ",PA52=" & CNULL(ChgSQL(Text33(1))) & ",PA53=" & CNULL(ChgSQL(Text33(2))) & ",PA54=" & CNULL(ChgSQL(Text33(3))) & _
      ",PA55=" & CNULL(ChgSQL(Text33(4))) & ",PA56=" & CNULL(ChgSQL(Text33(5))) & ",PA48=" & CNULL(ChgSQL(Text12)) & ",PA57=" & CNULL(Text10(2)) & _
      ",PA101=" & CNULL(Text19) & ",PA102=" & CNULL(ChgSQL(Text20)) & ",PA103=" & CNULL(Replace(Text21, "'", "''")) & ",PA104=" & CNULL(ChgSQL(Text22)) & _
      ",PA139=" & CNULL(ChgSQL(Text33(6))) & " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      
    cnnConnection.Execute strTxt(intStep)

   intStep = intStep + 1
   
   '1
   If frm06010602_2.Text6 = "1" Then
      If Left(strKind, 1) = "1" Or Left(strKind, 1) = "3" Then
         '2005/10/20 MODIFY BY SONIA ¤£§PÂ_CP25
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & DBNullDate(DBDATE(Text6)) & " WHERE " & _
         '   "CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & DBNullDate(DBDATE(Text6)) & " WHERE " & _
            "CP09='" & strReceiveNo & "' AND CP24 IS NULL"
         '2005/10/20 END
        cnnConnection.Execute strTxt(intStep)

         intStep = intStep + 1
      End If
      If Left(strKind, 1) <> "1" And Left(strKind, 1) <> "3" Then
         '2005/10/20 MODIFY BY SONIA ¤£§PÂ_CP25
         'strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & DBNullDate(DBDATE(Label3(3))) & " WHERE " & _
         '   "CP09='" & strReceiveNo & "' AND CP24 IS NULL AND CP25 IS NULL"
         strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1',CP25=" & DBNullDate(DBDATE(Label3(3))) & " WHERE " & _
            "CP09='" & strReceiveNo & "' AND CP24 IS NULL"
         '2005/10/20 END
        cnnConnection.Execute strTxt(intStep)

         intStep = intStep + 1
         If strKind = "701" Then
            strTxt(intStep) = "UPDATE PATENT SET PA23=1 WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
            
            cnnConnection.Execute strTxt(intStep)

            intStep = intStep + 1
         End If
      End If
      'Added by Lydia 2015/10/02 ³¡¥÷®×¥ó©Ê½è¤§®Ö­ã1001§ï¬°®Öµo1008
      If InStr(Patent1001Display, m_CP10) > 0 Then
          i = 1008
      'Added by Lydia 2025/02/12
      ElseIf m_CP10 = "245" Then
          i = 1924
      Else
          i = ®Ö­ã
      End If
      'end 2015/10/02
      
      strExc(1) = ""
      'µo¤å¤é
      strExc(2) = IIf(Left(m_CP10, 1) <> "8", strSrvDate(1), "Null")
   Else
      i = §ïÅÜ­ì³B¤À
      'Add by Morgan 2007/10/24 ©Ó¿ì´Á­­¹w³]6­Ó¤u§@¤Ñ
      strExc(1) = CompWorkDay(6, strSrvDate(1))
      'Added by Morgan 2013/4/29
      strExc(2) = IIf(Left(m_CP10, 1) <> "8", strSrvDate(1), "Null")
   End If
   
   strCP10 = i 'Added by Morgan 2017/5/10
   
      '3
   NewReceiveNo = AutoNo("C", 6)
   
   m_NewReceiveNo = NewReceiveNo
   
   'Added by Morgan 2012/12/4
   If m_PA162 <> pa(162) Then
      strSql = "update patent set pa162='" & m_PA162 & "' where " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
      Pub_SeekTbLog strSql 'Added by Morgan 2019/11/20
      cnnConnection.Execute strSql, intI
   End If
   '­n¥t¨ç³qª¾¦ý¥¼¿é¤J«ØÄ³©w½Z¤å¦r®É¤£¤Wµo¤å¨ÃºÞ¨î©Ó¿ì´Á­­,ÁÙ­nEmailµ¹¤uµ{®v¤Î¨ä¥DºÞ
   'Modified by Morgan 2019/10/7 +¦³¤À³Î®×
   'If m_bDivSugTextAlert = True Then
   If m_bDivSugTextAlert Or m_bHasDivCase Then
   'end 2019/10/7
       
      strExc(1) = CompWorkDay(4, strSrvDate(1)) '©Ó¿ì´Á­­
      strExc(2) = "Null" 'µo¤å¤é
      strCP14 = Text16
            
      strExc(0) = "'¥»©Ò®×¸¹¡G'||pa01||'-'||pa02||'-'||pa03||'-'||pa04||chr(13)||chr(10)" & _
            "||'®×¥ó¦WºÙ¡G'||pa05||chr(13)||chr(10)" & _
            "||'¥Ó½Ð¤H¡G'||cu04||chr(13)||chr(10)" & _
            "||'©Ó¿ì´Á­­¡G'||sqldatet(" & strExc(1) & ")||chr(13)||chr(10)"
      
      strLang = PUB_GetLanguage(pa(1), pa(2), pa(3), pa(4))  'Added by Morgan 2021/9/24
      
      'Added by Morgan 2019/12/31
      '­Y­ì¥¼³]©w¬O§_¥[µù®Ö­ã¤À³Î«ØÄ³«h¤º®e¤£¦P
      If pa(162) = "" Then
         strExc(3) = "¦¬¨ì®Ö­ã³qª¾¡A©|¥¼§PÂ_¬O§_¥[µù®Ö­ã¤À³Î«ØÄ³!"
         strExc(0) = strExc(0) & "||chr(13)||chr(10)||chr(13)||chr(10)" & _
            "||'¥»®×¤w¦¬¨ì®Ö­ã³qª¾¡A©|¥¼§PÂ_¬O§_¥[µù®Ö­ã¤À³Î«ØÄ³" & IIf(m_bHasDivCase, "¥B¦³¦¬¤å¤À³Î®×", "") & "¡A¨t²Î¥ý¹w³]¬°""Y""'||chr(13)||chr(10)"
         'Modified by Morgan 2021/9/24
         'strExc(0) = strExc(0) & _
            "||'1. ­Y¤£»Ý¥[µù¤À³Î«ØÄ³¡A½Ð¦Ü¤u§@¶i«×¸ê®ÆºûÅ@±NY§ï¬°N¡A¨÷ª½±µ°hµ{§Ç'||chr(13)||chr(10)" & _
            "||'2. ­Y»Ý¥[µù¤À³Î«ØÄ³¡A½Ð¥[µù¤º®e¡A¨÷°h¥DºÞ¤W§¹½Z¤é¡A¦A°hµ{§Ç'||chr(13)||chr(10)"
         strExc(0) = strExc(0) & _
            "||'1. ­Y¤£»Ý¥[µù¤À³Î«ØÄ³¡A½Ð¦Ü¤u§@¶i«×¸ê®ÆºûÅ@±NY§ï¬°N -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C'||chr(13)||chr(10)"
         'Modified by Morgan 2022/10/11
         'Modified by Morgan 2022/10/11 ¨ú®ø,§ï¤ñ·Ó­^¤å§@ªk
         'If strLang = "3" Then
         '   strExc(0) = strExc(0) & "||'2. ­Y»Ý¥[µù¤À³Î«ØÄ³(¤é¤å©w½Z)¡A½Ð³qª¾Bobbie´£¨Ñ§i­ã©w½Zµ¹¤uµ{®v¥[µù¤À³Î«ØÄ³¤º®e©ó©w½Z«á -> email³qª¾¥DºÞ¤W§¹½Z¤é -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C'||chr(13)||chr(10)"
         'Else
         'end 2022/10/11
            'Modified by Morgan 2024/5/13 --±Ó²ú
            'strExc(0) = strExc(0) & "||'2. ­Y»Ý¥[µù¤À³Î«ØÄ³¡A½Ð¥[µù¤º®e -> email³qª¾¥DºÞ¤W§¹½Z¤é -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C'||chr(13)||chr(10)"
            strExc(0) = strExc(0) & "||'2. ­Y»Ý¥[µù¤À³Î«ØÄ³¡A½ÐÂI¿ï""®Ö­ã""¿é¤J¥[µù¤º®e -> ¶]¾úµ{§@·~'||chr(13)||chr(10)"
            'end 2024/5/13
            
         'End If 'Removed by Morgan 2022/10/11
         'end 2021/9/24
      Else
      'end 2019/12/31
      
         'Modified by Morgan 2019/10/7
         'strExc(3) = "¦¬¨ìªì¼f®Ö­ã³qª¾¡A½Ð´£¨Ñ¤À³Î«ØÄ³!"
         If m_bHasDivCase Then
            strExc(3) = "¤w´£¤À³Î®×½Ð¤uµ{®v½T»{¬O§_¤´¥[µù¤À³Î«ØÄ³!"
            strExc(0) = strExc(0) & "||chr(13)||chr(10)||chr(13)||chr(10)" & _
               "||'¥»®×¤w´£¤À³Î®×¡A½Ð¤uµ{®v¦Ü¤u§@¶i«×¸ê®ÆºûÅ@½T»{¬O§_¤´¥[µù¤À³Î«ØÄ³¡H§_½Ð§ïN¡A­Y¦³¤À³Î«ØÄ³¤º®e¬O§_­×§ï¡H'"
         Else
         
            'Added by Morgan 2020/2/27
            If m_EditDivSugText <> "" Then
               strExc(3) = "¦¬¨ì®Ö­ã³qª¾¡A½Ð­×§ï¤À³Î«ØÄ³¤º®e!"
               strExc(0) = strExc(0) & "||chr(13)||chr(10)||chr(13)||chr(10)" & _
                  "||'" & ChgSQL(m_EditDivSugText) & "'"
            Else
            'end 2020/2/27
               strExc(3) = "¦¬¨ì®Ö­ã³qª¾¡A½Ð¥[µù¤À³Î«ØÄ³!"
               
               'Added by Morgan 2021/9/24 -- Bobbie
               'Modified by Morgan 2024/5/13 --±Ó²ú
               'strExc(0) = strExc(0) & "||'»Ý¥[µù¤À³Î«ØÄ³¡A½Ð¥[µù¤º®e -> email³qª¾¥DºÞ¤W§¹½Z¤é -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C'||chr(13)||chr(10)"
               strExc(0) = strExc(0) & "||'»Ý¥[µù¤À³Î«ØÄ³¡A½ÐÂI¿ï""®Ö­ã""¿é¤J¥[µù¤º®e -> ¶]¾úµ{§@·~'||chr(13)||chr(10)"
               'end 2024/5/13
               'end 2021/9/24
            End If 'Added by Morgan 2020/2/27
      
         End If
         'end 2019/10/7
         
      End If
      
      'Modified by Morgan 2019/12/27 °Æ¥»§ïµo²Ä¤G,¤T¯Å¥DºÞ
      'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
         " select '" & strUserNum & "' mc01,st01 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
         ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||'(" & m_NewReceiveNo & ")" & ChgSQL(strExc(3)) & "' mc07" & _
         "," & strExc(0) & " mc08,decode(oMan,st01,B0102,oMan) mc09" & _
         " from patent,customer,divsugtext,staff,SetSpecMan,ABS001" & _
         " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
         " and dst01(+)=pa01 and dst02(+)=pa02 and dst03(+)=pa03 and dst04(+)=pa04" & _
         " and st01='" & strCP14 & "' and OCODE=decode(st16,'1','T','2','R','3','S','4','T1') and B0101(+)=st01"
      'Modified by Lydia 2020/08/24 §ï¥Î¼Ò²Õ
      'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
         " select '" & strUserNum & "' mc01,st01 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
         ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||'(" & m_NewReceiveNo & ")" & ChgSQL(strExc(3)) & "' mc07" & _
         "," & strExc(0) & " mc08,st52||';'||st53 mc09" & _
         " from patent,customer,staff" & _
         " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
         " and st01='" & strCP14 & "'"
      'end 2019/12/27
      
      'Added by Morgan 2021/9/24 ¤é¤å©w½Z¦Û°Êµoemailµ¹ Bobbie,ccµ¹¦U°Ïµ{§Ç
      'Modified by Morgan 2022/10/11 ¨ú®ø,§ï¤ñ·Ó­^¤å§@ªk
'      If strLang = "3" And pa(162) = "Y" And m_bHasDivCase = False Then
'         strExc(3) = "¤é¤å©w½Z¶·¥[µù¤À³Î«ØÄ³¡A½Ð´£¨Ñ§i­ã©w½Zµ¹¤uµ{®v¥[µù"
'         strExc(4) = "'¥»©Ò®×¸¹¡G'||pa01||'-'||pa02||'-'||pa03||'-'||pa04||chr(13)||chr(10)" & _
'            "||'©Ó¿ì´Á­­¡G'||sqldatet(" & strExc(1) & ")||chr(13)||chr(10)" & _
'            "||'¥»®×" & strExc(3) & "'||chr(13)||chr(10)" & _
'            "||'½Ð¤uµ{®v¥[µù¤À³Î«ØÄ³¤º®e©ó©w½Z«á -> email³qª¾¥DºÞ¤W§¹½Z¤é -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C'"
'
'         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'            " select '" & strUserNum & "' mc01,'" & Pub_GetSpecMan("¥~±M§i­ãµ{§Ç") & "' mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
'            ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||'(" & m_NewReceiveNo & ")" & ChgSQL(strExc(3)) & "' mc07" & _
'            "," & strExc(4) & " mc08,'" & PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4)) & "' mc09" & _
'            " from patent where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
'
'      Else
      'end 2021/9/24
         'Add by Amy 2025/08/05 «áÄò­ã»éÂ²³æ³ø§i=Y,¿éCÃþ¨Ó¨ç[¥D¦®]³Ì«e­±¥[¡i½ÐÂ²³æ³ø§i¡j-Winfrey
         If pa(89) = "Y" Then strExc(3) = "¡i½ÐÂ²³æ³ø§i¡j" & strExc(3)
         
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " select '" & strUserNum & "' mc01,st01 mc02,to_char(sysdate,'yyyymmdd') mc03,to_char(sysdate,'hh24miss') mc04" & _
            ",pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)||'(" & m_NewReceiveNo & ")" & ChgSQL(strExc(3)) & "' mc07" & _
            "," & strExc(0) & " mc08,'" & PUB_GetFCPEngSup(strCP14) & "' mc09" & _
            " from patent,customer,staff" & _
            " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
            " and st01='" & strCP14 & "'"
            
      'End If 'Added by Morgan 2021/9/24 'Removed by Morgan 2022/10/11
      
      cnnConnection.Execute strSql, intI
   End If
   'end 2012/12/4
      
   If strCP14 = "" Then strCP14 = PUB_GetFCPPromoterNo(strReceiveNo, "" & i, m_CP14)
   
   'Added by Lydia 2019/05/23 °É»~¤½³ø±±ºÞ: ©Ó¿ì¤H±¾Sharon¡A¿é¤J¸ê®Æ±a¤J¶i«×³Æµù¡A¨Ã¥B³]µo¤å¤é¬°ªÅ¥Õ;
   If Frame1.Visible = True Then
       'Modified by Lydia 2019/06/19 §ï¦¨¯S®í³]©w
       'strCP14 = "86013"
       strCP14 = Pub_GetSpecMan("¥~±Mµ{§Ç-°É»~§¹³Æ")
       strExc(2) = "Null" 'µo¤å¤é
       strExc(1) = TransDate(txtCRC(0), 2) '©Ó¿ì´Á­­
       If Trim(txtCRC(0)) = "" Then
           strCP64 = strCP64 & "___¦~___¤ë___¤é"
       Else
           strCP64 = strCP64 & Mid(txtCRC(0), 1, 3) & "¦~" & Mid(txtCRC(0), 4, 2) & "¤ë" & Mid(txtCRC(0), 6, 2) & "¤é"
       End If
       If Trim(txtCRC(1)) = "" Then
           strCP64 = strCP64 & "²Ä___´Á¤§"
       Else
           strCP64 = strCP64 & "²Ä" & txtCRC(1) & "´Á¤§"
       End If
       'Added by Lydia 2023/08/25 ±M§QÅv©µªø415: ¹w³]¶µ¥Ø
       If m_CP10 = "415" Then
          strCP64 = "¤½§i¤é´Á:" & strCP64 & Label3(1).Caption & ";"
       Else
       'end 2023/08/25
          If Opt1(0).Value = True Then strCP64 = strCP64 & "°É»~"
          If Opt1(1).Value = True Then strCP64 = strCP64 & "§ó¥¿"
          If Opt1(2).Value = True Then strCP64 = strCP64 & "±M§QÅvÅÜ§ó"
          strCP64 = "°É»~¤é´Á:" & strCP64 & ";"
       End If 'Added by Lydia 2023/08/25
       
       '§ó¥¿¥i¥H¤£¿é¤J¤é´Á¤Î´Á§O¡A¶i«×³Æµù¥Î__±a¤J¤é´Á¤Î´Á¼Æ¡F¦Û°Ê²£¥Í14¤é¾ä¤Ñ«áªº¦æ¨Æ¾ä´£¿ôFCPµ{§Ç¦V´¼¼z§½¸ß°Ý«á¸É¤W¤é´Á¡B´Á¼Æ©M©Ó¿ì´Á­­¡C
       'Modified by Lydia 2023/08/25 ¦æ¨Æ¾ä°Ï¤À¬O§_ÃÄ«~³sµ²®×
       'If Opt1(1).Value = True And (Trim(txtCRC(0)) = "" Or Trim(txtCRC(1)) = "") Then
       If Not (pa(177) = "Y" And i = ®Ö­ã And (m_CP10 = "415" Or m_CP10 = "402")) And Opt1(1).Value = True And (Trim(txtCRC(0)) = "" Or Trim(txtCRC(1)) = "") Then
            strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
            If strExc(3) <> "" Then
                strExc(6) = CompDate(2, 14, strSrvDate(1))
                strExc(4) = "¦V´¼¼z§½¸ß°Ý°É»~ªíªº¤½§i¤é©M´Á§O"
                If PUB_AddFCPStaffCalendar(strExc(6), "1", strExc(3), strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)) Then
                End If
            End If
       End If
   End If
   
   'Added by Morgan 2024/5/17
   '®Öµo§Þ³N³ø§i:¤£­n¦Û°Ê¤Wµo¤å¡A©Ó¿ì´Á­­+5­Ó¤u§@¤Ñ(¤£§t·í¤é)¡A¥»©Ò´Á­­=©Ó¿ì´Á­­+5­Ó¤u§@¤Ñ(Trigger)
   If m_CP10 = "421" Then
      strExc(1) = CompWorkDay(6, strSrvDate(1)) '©Ó¿ì´Á­­
      strExc(2) = "Null" 'µo¤å¤é
   End If
   'end 2024/5/17
   
   'Modified by Lydia 2019/05/23
   'strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP13,CP12,CP20,CP26,CP32,CP27,CP43,CP14,CP48) VALUES ('" & Text2 & "','" & Text3 & "','" & _
      Text4 & "','" & Text5 & "'," & TransDate(Label3(3), 2) & "," & _
      CNULL(Text7) & ",'" & NewReceiveNo & "','" & i & "','" & strCP13 & "','" & strCP12 & "'" & _
      ",'N','N','N'," & strExc(2) & ",'" & strReceiveNo & "','" & strCP14 & "'," & CNULL(strExc(1), True) & ")"
   strTxt(intStep) = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP08,CP09,CP10," & _
      "CP13,CP12,CP20,CP26,CP32,CP27,CP43,CP14,CP48,CP64) VALUES ('" & Text2 & "','" & Text3 & "','" & _
      Text4 & "','" & Text5 & "'," & TransDate(Label3(3), 2) & "," & _
      CNULL(Text7) & ",'" & NewReceiveNo & "','" & i & "','" & strCP13 & "','" & strCP12 & "'" & _
      ",'N','N','N'," & strExc(2) & ",'" & strReceiveNo & "','" & strCP14 & "'," & CNULL(strExc(1), True) & "," & CNULL(strCP64) & ")"
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1
   
   'Add By Sindy 2017/1/11 ·s¥Ó½Ð®×©Î¦A¼f107®×¤§®Ö­ã®É¡A¦P®É·s¼W²£¥Í"³qª¾§i­ã"(1917)ªºDÃþ¶i«×
   'Modify By Sindy 2017/5/4 ®Ö­ã®ÉÂI¿ï3¦rÀYªº©Ò¦³§ï½Ð®×¥ó©Ê½è(§t¤À³Î)®É¡A¤]­n²£¥Í"³qª¾§i­ã"¶i«×
   If (InStr(NewCasePtyList & ",107", m_CP10) > 0 Or Left(m_CP10, 1) = "3") And i = ®Ö­ã Then
      '¦³µo¤å¤é¥B¦³¡u¤À³Î«ØÄ³¡v
      If Val(strExc(2)) > 0 And m_bDivSugTextAlert = False Then
         '"³qª¾§i­ã" DÃþ¶i«×ªº©Ó¿ì´Á­­¡×®Ö­ã¨çµo¤å¤é¡Ï8­Ó¤é¾ä¤Ñ(·í¤é¤£ºâ)
         strCP48 = Val(CompDate(2, 8, DBDATE(strExc(2))))
         strCP48 = CompWorkDay(1, strCP48, 1)   'add by sonia 2025/3/14 ­Y¹J°²¤é«h´£«e¦Ü«e¤@¤u§@¤é
      Else
         strCP48 = "Null"
      End If
      strNewCP09 = AutoNo("D", 6)
      'Add By Sindy 2017/6/6 µo©úªì¼f®Ö­ã¥[µù¤À³Îªk©w´Á­­
      strDivState = "N"
      
      'Modified by Morgan 2019/10/7 108.11.1 ·sªkµo©ú/·s«¬­ã«á3¤ë¤º³£¥i´£¤À³Î
      'If pa(8) = "1" Then 'µo©ú
      '   'Modified by Morgan 2012/12/26 +¦Ò¼{¤À³Î®×®Ö­ã
      '   strExc(0) = "SELECT pa162,cp10,cp09,pa163 FROM caseprogress,patent" & _
      '      " WHERE " & TransDate(Label3(3), 2) & ">=20121202 and cp09=" & CNULL(Label3(2)) & " and cp10 in ('101','307')" & _
      '      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04"
      '   intI = 1
      '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      '   If intI = 1 Then
      '
      '      strExc(1) = CompDate(2, 30, TransDate(Label3(3), 2))
      '      strExc(2) = CompDate(2, -2, strExc(1))
      '      'µo©ú¥Ó½Ð
      '      If RsTemp.Fields("cp10") = "101" Then
      '         strDivState = "Y"
      '      '¤À³Î
      '      ElseIf RsTemp.Fields("cp10") = "307" And RsTemp.Fields("pa163") = "Y" Then
      '         strDivState = "Y"
      '      End If
      '   End If
      '   If strDivState = "Y" Then
      '      m_CP64 = "¤À³Îªk©w´Á­­" & ChangeWStringToTDateString(strExc(1))
      '   End If
      'End If
      If pa(8) = "1" Or pa(8) = "2" Then
         strDivState = "Y"
         strExc(1) = CompDate(1, 3, TransDate(Label3(3), 2))
         m_CP64 = "¤À³Îªk©w´Á­­" & ChangeWStringToTDateString(strExc(1))
      End If
      'end 2019/10/7
      
      strTxt(intStep) = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
         "CP12,CP13,CP14,CP20,CP26,CP32,CP43,CP48,CP64) " & _
      "VALUES ('" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & ",'" & strNewCP09 & "','1917'" & _
               ",'" & strCP12 & "','" & strCP13 & "','" & Pub_GetSpecMan("¥~±M§i­ãµ{§Ç") & "'," & _
               "'N','N','N','" & NewReceiveNo & "'," & strCP48 & "," & CNULL(m_CP64) & ")"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '2017/1/11 END
   
   'Add By Sindy 2015/6/3
   If i = ®Ö­ã And Check1.Value = 1 Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP148='Y' WHERE CP09='" & NewReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '2015/6/3 END
   
   'Add by Morgan 2007/7/23 CP20§ï§ìCPMªº³]©w
   'Modify by Morgan 2008/3/27 +pa75
   'Modify by Morgan 2008/4/10 +¥»©Ò®×¸¹
   strCP20 = PUB_GetCP20(Text2, Format(i), strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
   If strCP20 = "" Then
      strSql = "update caseprogress set cp20=NULL,cp16=" & strCP16 & ",cp17=0,cp18=" & strCP16 / 1000 & _
         " where cp09='" & NewReceiveNo & "'"
      cnnConnection.Execute strSql
   End If
   'end 2007/7/23
   
   If pa(9) = ¥xÆW°ê®a¥N¸¹ And Val(Label3(3)) >= 930701 Then
      'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
      If InStr("101,102,103,104,105,107,125,301,302,303,304,305,306,307,308", m_CP10) > 0 Then
      
         'Modify by Morgan 2006/8/28 ´¼Åv¤H­û§ï¥ÎPUB_GetFCPSalesNo¤£¥i¥ÎPUB_GetAKindSalesNo§ì
         'Modify By Sindy 2021/4/26 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):¬ù©w´Á­­
         strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP13,NP22,NP23) " & _
            "VALUES ('" & NewReceiveNo & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & stNP07 & "," & _
            stNP08 & "," & stNP09 & "," & CNULL(PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))) & "," & CNULL(Text7.Text) & "," & _
            lMax & "," & CNULL(DBDATE(m_pAgreeOnDate)) & ")"
          cnnConnection.Execute strTxt(intStep)
         intStep = intStep + 1
          lMax = GetNextProgressNo  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  objPublicData.GetNextProgressNo
          
          strTxt(intStep) = "Update CASEPROGRESS SET CP06=" & stNP08 & ", CP07=" & stNP09 & " WHERE CP09='" & NewReceiveNo & "'"
          cnnConnection.Execute strTxt(intStep)
          intStep = intStep + 1
      End If
   End If
   
   If i = §ïÅÜ­ì³B¤À Then
      strTxt(intStep) = "UPDATE CASEPROGRESS SET CP24='1' WHERE CP09='" & NewReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   
   'Added by Lydia 2025/02/12 245©µ½w¼f¬d>>1924­ã¤©©µ½w¼f¬d
   If i = 1924 And txt415Date <> "" Then
      strSql = "Update CaseProgress Set cp71=" & DBDATE(txt415Date) & " where cp09='" & NewReceiveNo & "'"
      cnnConnection.Execute strSql
      '½Õ¾ã¦ÛÄò¦æ¼f¬d¤é´Á+¨t²Î¤º¹w¦ô­n¶Ê¼fªº¤Ñ¼Æ
      strExc(0) = "select np01,np22,cf05 from caseprogress,nextprogress,casefee" & _
         " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10 in (" & NewCasePtyList & ")" & _
         " and np01(+)=cp09 and np07='411' and np06 is null" & _
         " and cf01(+)=cp01 and cf02='" & pa(9) & "' and cf03(+)=cp10 and cf05>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strExc(1) = CompDate(2, RsTemp("cf05"), DBDATE(txt415Date))
         strExc(2) = PUB_GetWorkDay1(strExc(1), True)
         strSql = "update nextprogress set np08=" & strExc(2) & ",np09=" & strExc(1) & " where np01='" & RsTemp("np01") & "' and np07='411' and np22=" & RsTemp("np22")
         cnnConnection.Execute strSql, intI
      End If
      '³qª¾Email
      strExc(1) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
      strExc(2) = PUB_GetFCPProSup(strExc(1))
      strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
      strExc(0) = strExc(2) & ";" & strExc(3) & ";backup"
      strExc(4) = pa(1) & "-" & pa(2) & IIf(pa(3) = "0", "", "-" & pa(3)) & IIf(pa(4) = "00", "", "-" & pa(4))
      strExc(5) = "¥»®×" & Label3(1) & "¤w®Ö­ã¡AÄò¦æ¼f¬d¤é¬°" & ChangeTStringToTDateString(txt415Date) & "¡A¨÷©v°Ï¹q¤l¤½¤åÀÉ¦W¡G" & strExc(4) & ".pdf¡A½Ð³ø§i«È¤á¡C"
      'Modified by Lydia 2025/04/30 ¥D¦®«á­±+Our Ref: FCP-xxxxx [INCOM.1924]
      strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
         " values('" & strUserNum & "','" & strExc(1) & "',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strExc(4) & Label3(1) & "¤w®Ö­ã") & "Our Ref: " & strExc(4) & " [INCOM." & i & "]','" & ChgSQL(strExc(5)) & "','" & strExc(0) & "')"
      cnnConnection.Execute strSql, intI
   End If
   'end 2025/02/12
   
   
'Move by Lydia 2021/04/08 ¸g¹L½T»{¡u¤½§i¸¹:109021103¡vªºµ{¦¡¤£¸Ó©ñ¦b§PÂ_¦³®Ö¹ï¤w­ãªº¬q¸¨,²¾¨ì«e¤è
    If m_bNewGrant = True Then 'Added by Lydia 2021/04/14 ªì¼f®Ö­ã¤~­n³qª¾
            'Add By Sindy 2020/2/7
            '¥N²z¤H¬°: Y2776600 (MURATA MANUFACTURING CO., LTD. ¥BINTELLECTUAL PROPERTY DEPT.)
            '¥Ó½Ð¤H¬°: X2776600 (MURATA MANUFACTURING CO., LTD.)
            '¤~»Ý­n±HµoE-Mail³qª¾
            strMailText = ""
            'Modify By Sindy 2020/3/20 + Bobbie:¥H¤U4­ÓY½s¸¹
            '  Y20990 Murgitroyd & Company
            '  (¥]§tY2099001¡BY2099002¡BY2099003¡BY2099004¡BY2099005¡BY2099006¡BY20990B7¡BY20990B8)
            '  Y20372   ALFA-LAVAL CORPORATE AB
            '  Y5179901 Sandvik Intellectual Property
            '  Y4830904 Syngenta Participations AG
            'modify by sonia 2020/6/1 Y27766§ï°T®§¬G¿W¥ß¦b¤U¤è
            'Modified by Lydia 2021/02/02 §ï¦¨¯S®í³Æµù³]©w(³qª¾§i­ã¥[µù/EmailºûÅ@)
'            If Left(ChangeCustomerL(pa(75)), 6) = "Y20990" Or _
'               Left(ChangeCustomerL(pa(75)), 6) = "Y20372" Or _
'               ChangeCustomerL(pa(75)) = "Y51799010" Or _
'               ChangeCustomerL(pa(75)) = "Y48309040" Then
'               strMailText = "³Ì·sª©¥»¤§­ì¤å½Ð¨D¶µWORDÀÉ"
'            'add by sonia 2020/6/1
'            ElseIf (ChangeCustomerL(pa(75)) = "Y27766000" And _
'                (Text33(9) = "X27766000" Or Text33(10) = "X27766000" Or Text33(11) = "X27766000" _
'                 Or Text33(12) = "X27766000" Or Text33(13) = "X27766000")) Then
'               strMailText = "®Ö­ãª©¥»½Ð¨D¶µ¤éÄ¶¤åWORDÀÉ¡APH®×°£®Ö·Ç½d³ò¤éÄ¶¤å¡A¥ç»ÝªþºK­n´y­z½Ð¨D¶µ¸g¼f¬d²£¥Í¤§ÅÜ¤Æ¡C"
'            'end 2020/6/1
'            'Modify By Sindy 2020/2/14
'            'Y47778(AJU Kim Chang & Lee) + ªÚ¦p:X26046 (SK hynix Inc)
'            'Modify By Sindy 2020/3/20 + ªÚ¦p:¥N²z¤HY49053 (YUIL HIGHEST INTERNATIONAL PATENT AND LAW FIRM)
'            'Modify By Sindy 2020/4/14 + ªÚ¦p:Y47778¡ÏX77517000¤]­n³qª¾
'            ElseIf ChangeCustomerL(pa(75)) = "Y49053000" Or _
'                   (ChangeCustomerL(pa(75)) = "Y47778000" And _
'                    (Text33(9) = "X26046000" Or Text33(10) = "X26046000" Or Text33(11) = "X26046000" _
'                     Or Text33(12) = "X26046000" Or Text33(13) = "X26046000")) Or _
'                   (ChangeCustomerL(pa(75)) = "Y47778000" And _
'                    (Text33(9) = "X77517000" Or Text33(10) = "X77517000" Or Text33(11) = "X77517000" _
'                     Or Text33(12) = "X77517000" Or Text33(13) = "X77517000")) Then
'               strMailText = "¤w­ã½Ð¨D¶µªº¤¤¤å¥»+­^¤å¥»WORDÀÉ"
'            '2020/2/14 END
'            End If
'            If strMailText <> "" Then
            strMailSubject = ""
            'Modified by Lydia 2022/03/30 ±Æ°£³¬¨÷(¾P¨÷)
            If pa(8) <> "3" And Trim("" & pa(57) & pa(108)) = "" Then    'Added by Lydia 2021/09/03 ±Æ°£³]­p®×
                'Modified by Lydia 2023/03/22 ¾ã¦X¼Ò²Õ¦bPUB_GetApprovalPS
                'If GetApprovalPS(pa(1) & pa(2) & pa(3) & pa(4), ChangeCustomerL(pa(75)), Text33(9) & "," & Text33(10) & "," & Text33(11) & "," & Text33(12) & "," & Text33(13), strMailSubject, strMailText) = True Then
                If PUB_GetApprovalPS("2", pa(1) & pa(2) & pa(3) & pa(4), ChangeCustomerL(pa(75)), Text33(9) & "," & Text33(10) & "," & Text33(11) & "," & Text33(12) & "," & Text33(13), strMailSubject, strMailText) = True Then
                'end 2021/02/02
                   '¤uµ{®v¤wÂ÷Â¾®É,§ì¨ä¥DºÞ
                   'Modified by Lydia 2021/04/08
                   'If str926CP14 = "" Then
                   '   '¤uµ{®v¥DºÞ
                   '   str926CP14 = PUB_GetFCPEngSup(strCP14)
                   'End If
                   If GetStaffName(strCP14) = "" Then
                       str926CP14 = PUB_GetFCPEngSup(strCP14) '¤uµ{®v¥DºÞ
                   Else
                       str926CP14 = strCP14
                   End If
                   'end 2021/04/08
                   'Added by Lydia 2024/06/14 ­Y¤W¤@¹Dªº©Ó¿ì¤uµ{®v¬°¤º±M¤uµ{®v¡A «h¥D­n¦¬¥óªÌ¡A§ï¬°¹ï±µªº¥~±M¥DºÞ
                   If Mid(str926CP14, 4, 1) = "9" Then
                       str926CP14 = PUB_GetFCPEngSup(str926CP14, , , True)
                   End If
                   'end 2024/06/14
                   
                   '¥D¦®
                   'Modified by Lydia 2021/02/02
                   'strExc(4) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "½Ð´£¨Ñ" & strMailText
                   'Modified by Lydia 2021/03/02 debug : strMailText => strMailSubject
                   strExc(4) = pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & strMailSubject
                   'Added by Morgan 2024/4/17 ¾÷±ñ²Õ®×¥ó¥D¦®³£¥[¡i¾÷±ñ³]­p²Õ¡j--Sharon
                   If pa(150) = "4" Then
                     strExc(4) = "¡i¾÷±ñ³]­p²Õ¡j" & strExc(4)
                   End If
                   'end 2024/4/17
                   'Add by Amy 2025/08/05 «áÄò­ã»éÂ²³æ³ø§i=Y,¿éCÃþ¨Ó¨ç[¥D¦®]³Ì«e­±¥[¡i½ÐÂ²³æ³ø§i¡j-Winfrey
                   If pa(89) = "Y" Then strExc(4) = "¡i½ÐÂ²³æ³ø§i¡j" & strExc(4)
                   
                   '¤º¤å
                   'Modified by Lydia 2021/02/02
                   'strExc(0) = "Dear " & GetPrjSalesNM(str926CP14) & "¡A" & vbCrLf & vbCrLf & _
                               "¡@¡@¦¹®×¤w®Ö­ã¡A½Ð´£¨Ñ" & strMailText & "¡A" & vbCrLf & _
                               "¡@¡@¨Ã½ÐEmailµ¹" & GetPrjSalesNM(Pub_GetSpecMan("¥~±M§i­ãµ{§Ç")) & "(" & Pub_GetSpecMan("¥~±M§i­ãµ{§Ç") & ")¡A­Y»Ý­n¨÷¡A½Ð¦Aª¾·|" & strUserName & "(" & strUserNum & ")" & vbCrLf & _
                               "ÁÂÁÂ¡I"
                   'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      " values( '" & strUserNum & "','" & str926CP14 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                      ",'" & strExc(4) & "','" & strExc(0) & "','" & Pub_GetSpecMan("¥~±M§i­ãµ{§Ç") & "')"
                   strExc(5) = Pub_GetSpecMan("¥~±M§i­ãµ{§Ç")
                   'Modified by Lydia 2022/05/20 GetPrjSalesNM=>PUB_ReadUserData
                   strExc(0) = "Dear " & PUB_ReadUserData(str926CP14) & "¡A" & vbCrLf & vbCrLf & _
                               "¡@¡@" & strMailText & vbCrLf & _
                               "¡@¡@¨Ã½ÐEmailµ¹" & GetPrjSalesNM(strExc(5)) & "(" & strExc(5) & ")¡A­Y»Ý­n¨÷¡A½Ð¦Aª¾·|" & strUserName & "(" & strUserNum & ")" & vbCrLf & _
                               "ÁÂÁÂ¡I"
                   'Modified by Lydia 2021/03/02 ¼W¥[CCµ¹¾Þ§@ªÌ
                   'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      " values( '" & strUserNum & "','" & str926CP14 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                      ",'" & strExc(4) & "','" & strExc(0) & "','" & strExc(5) & "')"
                   strExc(6) = strExc(5)
                   If strUserNum <> strExc(6) Then strExc(6) = strExc(6) & ";" & strUserNum
                   strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      " values( '" & strUserNum & "','" & str926CP14 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                      ",'" & strExc(4) & "','" & strExc(0) & "','" & strExc(6) & "')"
                   'end 2021/02/02
                   cnnConnection.Execute strSql, intI
                End If
                '2020/2/7 END
            End If 'Added by Lydia 2021/09/03
    End If 'Added by Lydia 2021/04/14
'----end --Move by Lydia 2021/04/08

   '2007/10/12 modify by sonia ­ì¥u°µ®Ö­ã®É¦Û°Ê¤º³¡¦¬¤å926®Ö¹ï¤w­ã±M§Q,FCP-024010§ïÅÜ­ì³B¤À¤]­n
   'Add by Morgan 2007/4/9 ¿é¤J®Ö­ã®É¦Û°Ê¤º³¡¦¬¤å926®Ö¹ï¤w­ã±M§Q
   m_BSheetNo = ""
   'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
   'Memo by Lydia 2015/07/17 ®Ö­ãªº§PÂ_¦³ÅÜ§ó,½Ð¤@¨Ö­×§ïfrm075004_2.cmdPrintCForm_Click
   If DBDATE(Label3(3)) > 20070415 And ((m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "308")) Then
      If Not (pa(57) = "Y" And pa(89) = "") Then
         'Add by Morgan 2007/10/29 'ÀË¬d®Ö¹ï¤w­ã±M§Q³]©wpa141->fa85->cu122
         If PUB_CheckAuto926(pa) = True Then
         'end 2007/10/29
            'Modify by Morgan 2007/5/3 ¥[©Ó¿ì´Á­­=¦¬¤å¤é+12¤u§@¤Ñ
            '2008/8/27 modify by sonia §ï§ìc
            'strBCP48 = CompWorkDay(12, strSrvDate(1))
            strBCP48 = Pub_GetHandleDay(pa(1), pa(9), "926", strSrvDate(1))
            '2008/8/27 end
            m_BSheetNo = AutoNo("B", 6)
            '2008/11/20 MODIFY BY SONIA ¥[¹w³]½Ð´Úª÷ÃBCP16
            'strSQL = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
               "CP12,CP13,CP14,CP43,CP48) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
               pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & ",'" & m_BSheetNo & "'" & _
               ",'926','" & strCP12 & "','" & strCP13 & "','" & strCP14 & "','" & m_NewReceiveNo & "'," & strBCP48 & ")"
            strCP16 = Val(GetFCPFee(pa(1), "926"))
            strExc(3) = "" 'Added by Lydia 2024/03/11
            'Modify by Morgan 2010/6/3 Â÷Â¾¤£¹w³]
            'Added by Lydia 2024/04/15 «e¤@©Ó¿ì¤H¤uµ{®v¬°¤º±M¤H­û(¤£³B²z¤G®Ö)
            'If GetStaffName(strCP14) = "" Then
            If Mid(strCP14, 4, 1) = "9" Then
               strExc(3) = "­ì©Ó¿ì¤uµ{®v¬°¡G" & GetStaffName(strCP14, True) & ";"
               strExc(4) = PUB_GetFCPEngSup(strCP14, , , True)
               str926CP14 = PUB_GetFCPEngSup(strExc(4), , , True) '¤À§Oµ¹Wilison,Red
            ElseIf GetStaffName(strCP14) = "" Then
            'end 2024/04/15
               'Modified by Lydia 2024/03/11 ©Ó¿ì¤uµ{®v¤wÂ÷Â¾¡A¡i®Ö¹ï¤w­ã±M§Q¡j¶i«×©Ó¿ì¤H±¾¤uµ{®v¥DºÞ¡]°Æ²z¡^
               'str926CP14 = ""
               'frm060118»Ý­nµoEmail
               strExc(3) = "­ì©Ó¿ì¤uµ{®v¬°¡G" & GetStaffName(strCP14, True) & ";"
               str926CP14 = PUB_GetFCPEngSup(strCP14, True)
               'str926CP14 = PUB_SetEng(str926CP14) '¥~±M¾÷±ñ³]­p²Õ¤H­û²§°Ê½Õ¾ãµ{¦¡ 'Mark by Lydia 2024/04/15 ¤w¤£¾A¥Î---Morgan
               'end 2024/03/11
            Else
               str926CP14 = strCP14
            End If
            'Modified by Morgan 2015/10/2 Y4829203 ¹w³]¤£½Ð´Ú
            strExc(1) = ""
            'Modified by Morgan 2016/8/17 +Y54047,X45814,X67402,X6740201,X6740202,X60507,X60507001,X6050701,X70749,X71831,X71773 --³¯©É»T
            'If Left(pa(75) & "000", 8) = "Y4829203" Then
            'Modified by Morgan 2017/9/8 +Y22457,Y52322B10,Y48842,Y52322,Y48048,Y22457020,Y49562,X70406,X71137,X49346,X70197,X69605,X71927,X72756,X48049,X27727,X60507020,X48049C10 --¬x°ö³ó
            'Modified by Morgan 2019/2/26 +Y55199
            'Modified by Lydia 2019/04/08 +Y20438 (EATON)
            'Modified by Morgan 2019/4/24 +Y55240 (DuPont)--¬x°ö³ó
            'Modified by Morgan 2022/3/4 +Y55423 DuPont Toray Specialty Materials Kabushiki Kaisha -- Kimi
            'Modified by Morgan 2022/7/20 +X4581400,X7503800,X7181500,X8262500,X4720000,X7868700,Y2041200,Y5197100--Franny
            'Modified by Morgan 2022/10/19 +Y55020000 (Dow Chemical (China) Investment Company Ltd.)--¬x°ö³ó
            'Removed by Morgan 2025/8/8 -X2772700,X48049C1,X4934600,X6050700,X6050701,X6050702,X6960500,
            'X7019700,X7074900,X7113700,X7192700,X7275600,Y2245700,Y2245702,Y4804800,Y4884200,Y4956200,
            'Y5519900,Y2043800,Y5404700,X7181500,X4581400,X7503800,X7868700,X8262500,X4720000,Y5197100,
            'Y2041200,X4804900,X6740200,X6740201,X6740202,Y5524000,X7177300,X7183100,Y4829203--Anny
            '-X7040600,Y5232200,Y52322B1,Y5542300--Kimi
            '-Y5502000 --Tim
            'If InStr("Y4829203,Y5404700,Y2245700,Y52322B1,Y4884200,Y5232200,Y4804800,Y2245702,Y4956200,Y5519900,Y2043800,Y5524000,Y5542300,Y2041200,Y5197100", Left(pa(75) & "000", 8)) > 0 Or InStr("X4581400,X6740200,X6740201,X6740202,X6050700,X6050701,X7074900,X7183100,X7177300,X7040600,X7113700,X4934600,X7019700,X6960500,X7192700,X7275600,X4804900,X2772700,X6050702,X48049C1,X4581400,X7503800,X7181500,X8262500,X4720000,X7868700,Y5502000", Left(pa(26) & "000", 8)) > 0 Then
            '   strExc(1) = "N"
            'End If
            'end 2025/8/8
            
            'Modified by Lydia 2024/03/11 +CP64
            strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10," & _
               "CP12,CP13,CP14,CP16,CP18,CP20,CP43,CP48,CP64) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
               pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & ",'" & m_BSheetNo & "'" & _
               ",'926','" & strCP12 & "','" & strCP13 & "','" & str926CP14 & "','" & strCP16 & "','" & strCP16 / 1000 & "','" & strExc(1) & "','" & m_NewReceiveNo & "'," & strBCP48 & ",'" & ChgSQL(strExc(3)) & "')"
            cnnConnection.Execute strSql, intI
            'end 2015/10/2
            '2008/11/20 END
            
            'Memo by Lydia 2021/04/08 ¸g¹L½T»{¡u¤½§i¸¹:109021103¡vªºµ{¦¡¤£¸Ó©ñ¦b§PÂ_¦³®Ö¹ï¤w­ãªº¬q¸¨,²¾¨ì«e¤è

            
'Modified by Lydia 2014/11/26 ±N³Æµù³]¬°¦@¥Îªº©T©w³ÆµùÀÉApprovalMemo2
'BÃþ³æ(¤º³¡±µ¬¢³æ)¦sÀÉ,®×¥ó©Ê½è©T©w¬°926

'             'ADD BY SONIA 2014/5/9 Intersil¤Î¨ä¤l¤½¥qªº®×¥ó¦b®Ö¹ï¤w­ã±M§Qªº¤º³¡¦¬¤å³æ¥[¦L
'            Select Case Left(pa(26) & "000", 8)
'               Case "X6217700", "X5272200", "X5422700", "X5819500", "X6380100", "X6554500", "X6036001", "X4899100", "X4899101"
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "§i­ã®É½Ð¤@¨ÖCCµ¹Intersil ¡I"
'
'               Case "X5863100" 'Added by Morgan 2014/8/5 --Sharon
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "§i­ã®É¤@¨Ö±H®Ö­ã¤§­^¤åClaims¡I"
'
'               Case "X4779400" 'Added by Morgan 2014/9/26 --Joanne
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "§i­ã«á¡A¨÷½Ð¥ý°h©Ó¿ì¦¬¤å»âÃÒor ½Ðµ{§ÇºÞ¨î¦¬¤å»âÃÒ´Á­­¡C"
'            End Select
'            'END 2014/5/9
            
'            'Added by Morgan 2014/7/31
'            Select Case Left(pa(75) & "000", 8)
'               Case "Y4945600"
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "§i­ã®É»Ýªþ¤W³Ì·sª©¥»¤§¤é¤å½Ð¨D¶µ¡A½Ðµ{§Ç·|¤uµ{®v¼g«H¡C"
'
'               'Added by Morgan 2014/8/14 --Joanne
'               Case "Y2204600"
'                  If Left(pa(26) & "00", 8) = "X3429100" Then
'                     If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                     m_strMemo = m_strMemo & "§i­ã®É¡A§i­ã«H¤Î¨ä¥Lªþ¥ó¥HPDFÀÉ§Î¦¡±H¥X¡A¤£¶·¦A±N§i­ã¤º®e¶K©óE-mail¥»¤å¤¤¡C"
'                  End If
'
'               'Added by Morgan 2014/8/19 --Sharon
'               Case "Y5241800"
'                  If Left(pa(26) & "00", 8) = "X5603801" Then
'                     If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                     m_strMemo = m_strMemo & "ÃÒ®Ñ¥¿¥»»Ý¥t±H¦ÜY52418 OMYA International AG ¡C"
'                  End If
'               Case "Y4830900", "Y4830901", "Y4830902", "Y4830903", "Y4830904", "Y4830905", "Y5132600"
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "ÃÒ®Ñ¥¿¥»»Ý¥t±H¦ÜY48309080 Syngenta International AG ¡C"
'               Case "Y5336300", "Y5339200"
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "ÃÒ®Ñ¥¿¥»»Ý¥t±H¦ÜY48292030 Hewlett-Packard Company Intellectual Property " & vbCrLf & "Administration¡C"
'               'end 2014/8/19
'
'               'Added by Morgan 2014/9/9
'               Case "Y4880400"
'                  If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                  m_strMemo = m_strMemo & "§i­ã®É¡A¥H Email ¶Ç°e³ø§i«H + ªþ¥ó¡CEmail¶Ç°e«á¡A¶·¶Ç¯u³qª¾«È¤áEmail¤º®e" & vbCrLf
'                  m_strMemo = m_strMemo & "( µo¤å«á2¤é¤º¥¼Àò¤é¥NACKG¡A½Ð­«·s±H¤@¦¸ )"
'
'               'Added by Lydia 2014/10/28 °w¹ï¥N²z¤HY4835301¥B¥Ó½Ð¤H¬°NIKE(X55265,X72195) ªº®×¥ó,©ó±M§Q®Ö­ã¨ç¤§¤w­ã¦¬¤å³æªº³Æµù¦C¦L´£¥Ü
'               Case "Y4835301"
'                  If Left(pa(26) & "00", 8) = "X5526500" Or Left(pa(26) & "00", 8) = "X7219500" Then
'                     If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
'                     m_strMemo = m_strMemo & "§i­ã®É»Ý¤@¨Öªþ¤W¤w­ãªº­^¤å±M§Q½d³ò(¥Ñ¤uµ{®v´£¨Ñ)"
'                  End If
'            End Select
           'Modified by Lydia 2015/01/05 §ï¬°¤Ä¿ï°T®§ºØÃþ ,®Ö¹ï¤w­ã=2
           ' strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), "926", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)))
           'Modified by Lydia 2019/03/06 ³vµ§§PÂ_Y¥N²z¤H+X¥Ó½Ð¤H1~5
           'strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), "926,1001", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), "2")
           ' If Len(strExc(1)) > 0 Then
           '    If m_strMemo <> "" Then m_strMemo = m_strMemo & vbCrLf
           '       m_strMemo = m_strMemo & strExc(1)
           ' End If
           strExc(1) = "": strExc(2) = ""
           'Modified by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡G­×§ï¤@¯ë³Æµù¡B®Ö¹ï¤w­ã³Æµù¬°½Æ¼Æ·s³W«h
           'For intI = 0 To 4
           '    If pa(26 + intI) <> "" Then
           '         strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), "926,1001", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)), "2", bolTmp)
           '         If strExc(1) <> "" Then
           '             'Modified by Lydia 2022/07/29 ¦sÀÉ«e¤w¦³°O¿ý³Æµù; ex.FCP063282¦³­«ÂÐ³Æµù
           '             'If bolTmp = True Then '­Ó®×³Æµù
           '             '   m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
           '             '   Exit For
           '             'ElseIf strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
           '             If strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
           '             'end 2022/07/29
           '                If m_strMemo = "" Or (m_strMemo <> "" And InStr(m_strMemo, strExc(1)) = 0) Then
           '                      m_strMemo = m_strMemo & IIf(m_strMemo <> "", vbCrLf, "") & strExc(1)
           '                End If
           '                strExc(2) = strExc(2) & strExc(1) & "||" '§PÂ_¬O§_¦³­«½Æ³Æµù (®Ö¹ï¤w­ã±M§QªºÀË¬d)
           '             End If
           '         End If
           '    End If
           'Next intI
           ''end 2019/03/06
           '¦]¬°«e­±¤w§ì¤@¯ë®Ö­ã, ©Ò¥H­­©w¶Ç¤J®×¥ó©Ê½è,¥u§ì926®Ö¹ï¤w­ã±M§Q
           m_926strMemo = PUB_GetApprMemo2("2", pa(1) & pa(2) & pa(3) & pa(4), "926", ChangeCustomerL(pa(75)), pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30))
           'end 2022/08/02
'end 'Modified by Lydia 2014/11/26 ±N³Æµù³]¬°¦@¥Îªº©T©w³ÆµùÀÉApprovalMemo2

        'Modified by Lydia 2019/07/10 §ó§ï®Ö­ã¤£¥Î§ì®Ö­ã³Æµù
        'Else
        'Modified by Lydia 2019/08/01 ¥u°w¹ï¥Ó½Ð®×¤§®Ö­ã³Æµù,½Ð±Æ°£«D¥Ó½Ð®×(¦pÅÜ§ó,Åý»P,§ó§ï,§ó¥¿¡K)¤§®Ö­ã³Æµù
        'ElseIf Frame1.Visible = False Then
        ElseIf InStr(NewCasePtyList & ",107", m_CP10) > 0 Or Left(m_CP10, 1) = "3" Then '·s¥Ó½Ð®×+¤À³Î307+¦A¼f107+§ï½Ð3¶}ÀY
           'Modified by Lydia 2015/01/07 «D®Ö¹ï¤w­ã±M§Q=>¼u°T®§
           'Modified by Lydia 2019/03/06 ³vµ§§PÂ_Y¥N²z¤H+X¥Ó½Ð¤H1~5
           'strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), "926,1001", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26)), "2")
           ' If Len(strExc(1)) > 0 Then
           '    MsgBox strExc(1), vbExclamation, "¥»®×¤£¦C¦L®Ö¹ï¤w­ã±M§Q"
           ' End If
           'Modified by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡GFormSave¤§«e¤w¦³§ì³Æµù
           ' strExc(1) = "": strExc(2) = ""
           ' For intI = 0 To 4
           '      If pa(26 + intI) <> "" Then
           '         strExc(1) = PUB_GetApprMemo(pa(1) & pa(2) & pa(3) & pa(4), "926,1001", ChangeCustomerL(pa(75)), ChangeCustomerL(pa(26 + intI)), "2", bolTmp)
           '         If strExc(1) <> "" Then
           '             If bolTmp = True Then '­Ó®×³Æµù
           '                MsgBox strExc(1), vbExclamation, "¥»®×¤£¦C¦L®Ö¹ï¤w­ã±M§Q"
           '                Exit For
           '             ElseIf strExc(2) = "" Or (strExc(2) <> "" And InStr(strExc(2), strExc(1)) = 0) Then
           '                strExc(2) = strExc(2) & strExc(1) & "||" '§PÂ_¬O§_¦³­«½Æ³Æµù (®Ö¹ï¤w­ã±M§QªºÀË¬d)
           '             End If
           '         End If
           '      End If
           ' Next intI
           ' If strExc(2) <> "" Then MsgBox Replace(strExc(2), "||", vbCrLf), vbExclamation, "¥»®×¤£¦C¦L®Ö¹ï¤w­ã±M§Q"
           ''end 2019/03/06
           If m_strMemo <> "" Then MsgBox m_strMemo, vbExclamation, "¥»®×¤£¦C¦L®Ö¹ï¤w­ã±M§Q"
           'end 2022/08/02
        End If 'If PUB_CheckAuto926(pa) = True Then
      End If 'If Not (pa(57) = "Y" And pa(89) = "") Then
   End If
   'end 2007/4/9
   
   'Add By Sindy 2016/1/15 ¨Ï¥Î©ó©Ó¿ì³æ¦C¦L©ó³Æµù¤¤
   If Trim(m_strMemo) <> "" Then
      strTxt(intStep) = "update CASEPROGRESS set CP64='©Ó¿ì³æ³Æµù:" & m_strMemo & "|' where CP09='" & NewReceiveNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   '2016/1/15 END
   'Added by Lydia 2022/08/02 ¾ã¦X¼Ò²Õ¡G¥t¥~°O¿ý926®Ö¹ï¤w­ã±M§Q³Æµù
   If Trim(m_926strMemo) <> "" And m_BSheetNo <> "" Then
      strTxt(intStep) = "update CASEPROGRESS set CP64='©Ó¿ì³æ³Æµù:" & m_926strMemo & "|' where CP09='" & m_BSheetNo & "'"
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
   End If
   'end 2022/08/02
   
   '4
   strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & ¶Ê¼f & "'"
   cnnConnection.Execute strTxt(intStep)
   intStep = intStep + 1

   '5
   If frm06010602_2.Text6 = "2" Then
      'Modify by Morgan 2005/5/24 §ï§ì¥»©Ò¸¹
      'strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP01='" & strReceiveNo & "' AND NP07='" & §ïÅÜ­ì³B¤À & "'"
      strTxt(intStep) = "UPDATE NEXTPROGRESS SET NP06='Y' WHERE NP02='" & pa(1) & "' and NP03='" & pa(2) & "' AND NP04='" & pa(3) & "' AND NP05='" & pa(4) & "' AND NP06 IS NULL AND NP07='" & §ïÅÜ­ì³B¤À & "'"
      
    cnnConnection.Execute strTxt(intStep)

      intStep = intStep + 1
   End If
   
   Dim strOldData As String
   strOldData = Empty
   
   '6
   strExc(0) = "SELECT * FROM CHANGEEVENT WHERE CE01='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 1 To 99
            If IsNull(.Fields(i - 1)) Then
               strCe(i) = ""
            Else
               strCe(i) = .Fields(i - 1)
            End If
         Next
      End With
      strExc(1) = ""
      strExc(2) = ""
      strExc(3) = ""
      
      '¥Ó½Ð¤é 10
      If strCe(2) <> "" Then
         strExc(1) = strExc(1) & "¥Ó½Ð¤é : " & strCe(2) & ","
         strExc(2) = strExc(2) & "PA10=" & strCe(2) & ","
         strExc(3) = strExc(3) & "CE03='1',"
          strOldData = strOldData & "¥Ó½Ð¤é : " & pa(10) & " "
      End If
      
      '¥Ó½Ð¤H 26-30
      bolChk = False
      For i = 4 To 8
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strOldData = strOldData & "¥Ó½Ð¤H : "
         strExc(1) = strExc(1) & "¥Ó½Ð¤H : "
         For i = 4 To 8
            If strCe(i) <> "" Then
               strExc(1) = strExc(1) & strCe(i) & ","
               'edit by nickc 2007/02/02 ¤£¥Î dll ¤F
               'If objPublicData.GetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
               If ClsPDGetCustomerNameAndAddress(strCe(i), strTmp(5), strTmp(1), strTmp(2), strTmp(3)) Then
                  strExc(2) = strExc(2) & "PA" & i + 27 & "=" & CNULL(ChgSQL(strTmp(1))) & ",PA" & i + 32 & "=" & CNULL(ChgSQL(strTmp(2))) & ",PA" & i + 37 & "=" & CNULL(ChgSQL(strTmp(3))) & ","
               End If
            End If
            strExc(2) = strExc(2) & "PA" & i + 22 & "=" & CNULL(ChangeCustomerL(strCe(i))) & ","
         Next
         If IsEmptyText(strCe(4)) = False Then
            strOldData = strOldData & pa(26) & " "
         End If
         If IsEmptyText(strCe(5)) = False Then
            strOldData = strOldData & pa(27) & " "
         End If
         If IsEmptyText(strCe(6)) = False Then
            strOldData = strOldData & pa(28) & " "
         End If
         If IsEmptyText(strCe(7)) = False Then
            strOldData = strOldData & pa(29) & " "
         End If
         If IsEmptyText(strCe(8)) = False Then
            strOldData = strOldData & pa(30) & " "
         End If
         strExc(3) = strExc(3) & "CE09='1',"
      Else
         '¥Ó½Ð¦a§} 31-45
         bolChk = False
         For i = 23 To 37
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If bolChk = True Then
            strOldData = strOldData & "¥Ó½Ð¦a§} : "
            strExc(1) = strExc(1) & "¥Ó½Ð¦a§} : "
            For i = 23 To 37
               If strCe(i) <> "" Then
                  strExc(1) = strExc(1) & strCe(i) & ","
               End If
               strExc(2) = strExc(2) & "PA" & i + 8 & "=" & CNULL(strCe(i)) & ","
            Next
            strExc(3) = strExc(3) & "CE38='1',"
            ' 90.07.17 modify by louis (ÅÜ§ó¨Æ¶µÂÂ¸ê®Æ)
            If IsEmptyText(strCe(23)) = False Then
               strOldData = strOldData & pa(31) & " "
            End If
            If IsEmptyText(strCe(24)) = False Then
               strOldData = strOldData & pa(36) & " "
            End If
            If IsEmptyText(strCe(25)) = False Then
               strOldData = strOldData & pa(41) & " "
            End If
            If IsEmptyText(strCe(26)) = False Then
               strOldData = strOldData & pa(32) & " "
            End If
            If IsEmptyText(strCe(27)) = False Then
               strOldData = strOldData & pa(37) & " "
            End If
            If IsEmptyText(strCe(28)) = False Then
               strOldData = strOldData & pa(42) & " "
            End If
            If IsEmptyText(strCe(29)) = False Then
               strOldData = strOldData & pa(33) & " "
            End If
            If IsEmptyText(strCe(30)) = False Then
               strOldData = strOldData & pa(38) & " "
            End If
            If IsEmptyText(strCe(31)) = False Then
               strOldData = strOldData & pa(43) & " "
            End If
            If IsEmptyText(strCe(32)) = False Then
               strOldData = strOldData & pa(34) & " "
            End If
            If IsEmptyText(strCe(33)) = False Then
               strOldData = strOldData & pa(39) & " "
            End If
            If IsEmptyText(strCe(34)) = False Then
               strOldData = strOldData & pa(44) & " "
            End If
            If IsEmptyText(strCe(35)) = False Then
               strOldData = strOldData & pa(35) & " "
            End If
            If IsEmptyText(strCe(36)) = False Then
               strOldData = strOldData & pa(40) & " "
            End If
            If IsEmptyText(strCe(37)) = False Then
               strOldData = strOldData & pa(45) & " "
            End If
         End If
      End If

      '±M§Q°Ó¼ÐºØÃþ¥N¸¹ 08
      If strCe(39) <> "" Then
         strOldData = strOldData & "±M§Q°Ó¼ÐºØÃþ¥N¸¹ : " & pa(8) & " "
         strExc(1) = strExc(1) & "±M§Q°Ó¼ÐºØÃþ¥N¸¹ : " & strCe(39) & ","
         strExc(2) = strExc(2) & "PA08='" & strCe(39) & "',"
         strExc(3) = strExc(3) & "CE40='1',"
      End If
      
      '®×¥ó¦WºÙ 05-07
      bolChk = False
      For i = 41 To 43
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strOldData = strOldData & "®×¥ó¦WºÙ : "
         strExc(1) = strExc(1) & "®×¥ó¦WºÙ : "
         For i = 41 To 43
            If strCe(i) <> "" Then
               strExc(1) = strExc(1) & strCe(i) & ","
            End If
            strExc(2) = strExc(2) & "PA" & i - 36 & "=" & CNULL(strCe(i)) & ","
         Next
         strExc(3) = strExc(3) & "CE44='1',"
         If IsEmptyText(strCe(41)) = False Then
            strOldData = strOldData & pa(5) & " "
         End If
         If IsEmptyText(strCe(42)) = False Then
            strOldData = strOldData & pa(6) & " "
         End If
         If IsEmptyText(strCe(43)) = False Then
            strOldData = strOldData & pa(7) & " "
         End If
      End If
      
      '¥Nªí¤H 79-84
      bolChk = False
      For i = 10 To 15
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If Not bolChk Then
         For i = 68 To 91
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
      End If
      
      If bolChk Then
         strOldData = strOldData & "¥Nªí¤H : "
         strExc(1) = strExc(1) & "¥Nªí¤H : "
         For i = 10 To 15
            If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            strExc(2) = strExc(2) & "PA" & i + 69 & "=" & CNULL(strCe(i)) & ","
            If IsEmptyText(strCe(i)) Then
               strOldData = strOldData & pa(i + 69) & " "
            End If
         Next
         For i = 68 To 91
            If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            strExc(2) = strExc(2) & "PA" & i + 41 & "=" & CNULL(strCe(i)) & ","
            If IsEmptyText(strCe(i)) Then
               strOldData = strOldData & pa(i + 41) & " "
            End If
         Next
         strExc(3) = strExc(3) & "CE16='1',"
      End If
      
      '¥Nªí¤H¤¤Ä¶¤å
      If Not bolChk Then
         bolChk = False
         For i = 63 To 64
            If strCe(i) <> "" Then
               bolChk = True
               Exit For
            End If
         Next
         If Not bolChk Then
            For i = 92 To 99
               If strCe(i) <> "" Then
                  bolChk = True
                  Exit For
               End If
            Next
         End If
         If bolChk Then
            strExc(1) = strExc(1) & "¥Nªí¤H¤¤Ä¶¤å : "
            strExc(2) = strExc(2) & "PA79=" & CNULL(strCe(63)) & ",PA82=" & CNULL(strCe(64)) & "," & _
               "PA109=" & CNULL(strCe(92)) & ",PA112=" & CNULL(strCe(93)) & ",PA115=" & CNULL(strCe(94)) & "," & _
               "PA118=" & CNULL(strCe(95)) & ",PA121=" & CNULL(strCe(96)) & ",PA124=" & CNULL(strCe(97)) & "," & _
               "PA127=" & CNULL(strCe(98)) & ",PA130=" & CNULL(strCe(99)) & ","
            For i = 63 To 64
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            Next
            For i = 92 To 99
               If strCe(i) <> "" Then strExc(1) = strExc(1) & strCe(i) & ","
            Next
            strExc(3) = strExc(3) & "CE65='1',"
         End If
      End If
      
      ' 90.07.17 modify by louis
      ' ¥Ó½Ð¤H¤¤Ä³¤å
      bolChk = False
      For i = 17 To 21
         If strCe(i) <> "" Then
            bolChk = True
            Exit For
         End If
      Next
      If bolChk = True Then
         strExc(3) = strExc(3) & "CE22='1',"
      End If
      
      If strExc(1) <> "" Then
         For i = 2 To 3
            If Right(strExc(i), 1) = "," Then strExc(i) = Left(strExc(i), Len(strExc(i)) - 1)
         Next
         intStep = intStep + 1
         intStep = intStep + 1
         strTxt(intStep) = "UPDATE CHANGEEVENT SET " & strExc(3) & " WHERE CE01='" & strReceiveNo & "'"
         
        cnnConnection.Execute strTxt(intStep)
         
         intStep = intStep + 1
      End If
      
   End If
   
   lMax = GetNextProgressNo  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  objPublicData.GetNextProgressNo
   
   '2005/11/11 MODIFY BY SONIA
   'ElseIf (strKind = ²§Ä³µªÅG Or strKind = ³Q²§Ä³²z¥Ñ) And Text10(0) = "1" Then
   If (strKind = ²§Ä³µªÅG Or strKind = ³Q²§Ä³²z¥Ñ) And Text10(0) = "1" Then
      strTemp = CompDate(1, 3, TransDate(Label3(3).Caption, 2))
      'Modified by Morgan 2014/11/20 ¥~±M§ï¦^ÂÂ³W«h
      ''Added by Morgan 2014/10/29
      'If pa(9) = ¥xÆW°ê®a¥N¸¹ And strSrvDate(1) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
      '   strTemp1 = PUB_GetOurDeadline(strTemp)
      'Else
      ''end 2014/10/29
      
      'Added by Morgan 2019/7/11 ¥~±M¥xÆW®×©Ò­­¥H§ï¤u§@¤Ñ­pºâ
      If strSrvDate(1) >= ¥~±M¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
         'Modify By Sindy 2021/4/26 + m_pAgreeOnDate
         strTemp1 = PUB_GetFCPOurDeadline(strTemp, 2, , m_pAgreeOnDate)
      Else
      'end 2019/7/11
         
         strTemp1 = CompDate(2, -2, strTemp)
         
      End If 'Added by Morgan 2019/7/11
      
      'End If 'Added by Morgan 2014/10/29
      'end 2014/11/20
      
      lMax = GetNextProgressNo  'edit by nickc 2007/02/02 ¤£¥Î dll ¤F  objPublicData.GetNextProgressNo
      '2005/10/24 MODIFY BY SONIA
      'strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
      '   "NP07,NP08,NP09,NP10,NP13,NP14,NP22) VALUES ('" & NewReceiveNo & "','" & pa(1) & _
      '   "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & ³qª¾»âÃÒ & "," & _
      '   strTemp1 & "," & strTemp & "," & CNULL(cp(3)) & "," & CNULL(ChgSQL(Text7)) & "," & PUB_GetFCPSalesNo(Me.Text2.Text, Me.Text3.Text, Me.Text4.Text, Me.Text5.Text) & _
      '   "," & lMax & ")"
      'Modify by Morgan 2010/12/28 ¥Ó½Ð®×¸¹§ï½X¼Æ
      'If Mid(pa(11), 9, 1) = "U" Then
      If Mid(pa(11), 10, 1) = "U" Then
         stNP07 = ¥[µùÁp¦X '603
      'ElseIf Mid(pa(11), 9, 1) = "A" Then
      ElseIf Mid(pa(11), 10, 1) = "A" Then
         stNP07 = ¥[µù°l¥[ '602
      Else
         stNP07 = ³qª¾»âÃÒ '1601
      End If
      'Modify By Sindy 2021/4/26 + ,NP23=" & CNULL(DBDATE(m_pAgreeOnDate)):¬ù©w´Á­­
      strTxt(intStep) = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05," & _
         "NP07,NP08,NP09,NP10,NP13,NP14,NP22,NP23) VALUES ('" & NewReceiveNo & "','" & pa(1) & _
         "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'," & stNP07 & "," & _
         strTemp1 & "," & strTemp & "," & CNULL(cp(3)) & "," & CNULL(ChgSQL(Text7)) & "," & PUB_GetFCPSalesNo(Me.Text2.Text, Me.Text3.Text, Me.Text4.Text, Me.Text5.Text) & _
         "," & lMax & "," & CNULL(DBDATE(m_pAgreeOnDate)) & ")"
      '2005/10/24 END
      cnnConnection.Execute strTxt(intStep)
      intStep = intStep + 1
      lMax = lMax + 1
   End If
   
   'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
   If (m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "308") Or (m_CP10 >= "501" And m_CP10 <= "508") Or (m_CP10 >= "801" And m_CP10 <= "805") Then
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " And CP10>='203' AND CP10<='206' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         While Not rsA.EOF
            strTxt(intStep) = "Update NextProgress Set NP06 ='N' Where NP01='" & rsA.Fields("CP09").Value & "' AND " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND NP07='411' AND NP06 IS NULL "
            
            cnnConnection.Execute strTxt(intStep)
            
            intStep = intStep + 1
            rsA.MoveNext
         Wend
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   
   'Add by Morgan 2009/10/12
   If Val(txtCP19) > 0 Then
      strSql = "update caseprogress set cp19=" & Val(txtCP19) & " where cp09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql, intI
   End If
   
   If m_bAddAcc1k0 Then
      '¶}©l·s¼W°ê¥~½Ð´Ú¸ê®Æ
      '1:¥ý¥H"X"§ìACC1R0¤§°ê¥~½Ð´Ú³æªº¦Û°Ê½s¸¹, ¨Ã§ó·s¨ä¬y¤ô¸¹
      stA1k01 = AccAutoNo(MsgText(815), 5)
      AccSaveAutoNo MsgText(815), Right(stA1k01, 5)
      '2:·s¼WACC1K0
      '¥N²z¤H½s¸¹
      stA1k03 = PUB_GetA1K03(pa(1), pa(2), pa(3), pa(4))
      '¬üª÷¶×²v
'      dblUSRate = PUB_GetUSXRate
     
      '¦C¦L¹ï¶H
      strA1K27 = PUB_GetA1K27(pa(1), pa(2), pa(3), pa(4), m_CP10)
      If strA1K27 = "" Then strA1K27 = stA1k03
      '½Ð´Ú¹ï¶H
      strA1K28 = PUB_GetA1K28(pa(1), pa(2), pa(3), pa(4), m_CP10)
      If strA1K28 = "" Then strA1K28 = stA1k03

      '¬O§_¦C¦L¥Ó½Ð¤H
      strPrintCust = PUB_GetA1K04(pa(1), pa(2), pa(3), pa(4), strA1K28, m_CP10)
      
      'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
        Dim strA1K33 As String, strA1K18 As String
        'Modify By Sindy 2016/11/30
        'strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate)
        'Modified by Morgan 2018/4/27 +strA1K27
        strA1K33 = PUB_GetInitCurrPrintType(pa(1), strA1K28, strA1K18, dblUSRate, pa(2), pa(3), pa(4), strA1K27)
        '2016/11/30 END
        
      
      '§é¦©
      strDisc = 1 - (PUB_GetA1L07Disc(pa(1), pa(2), pa(3), pa(4), m_CP10, strSrvDate(2)) / 100)
      stA1L05 = 2500
      stA1L07 = Val(stA1L05) * strDisc
      stA1k11 = Fix(Val(stA1L05) - Val(stA1L07))
      If dblUSRate = 0 Then
         stA1k08 = stA1k11
      Else
         stA1k08 = Fix(Val(stA1k11) / dblUSRate)
      End If
      
      stA1k05 = PUB_GetDNRemark(strA1K28, pa(1), pa(2), pa(3), pa(4)) 'Added by Morgan 2017/3/22
      '¬üª÷¨ú¾ã¼Æ¦ì(µL±ø¥ó±Ë¥h)
      'Added by Lydia 2014/12/15 ½Ð´Ú³æ½Ð§ï¬°¨Ì¥N²z¤H©Î«È¤áÀÉ³]©wªº½Ð´Ú¹ô§O
      'strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K21,A1K19,A1K20 ) " & _
               " VALUES  ('" & stA1k01 & "'," & strSrvDate(2) & ",0,NULL,0," & dblUSRate & "," & stA1k11 & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
               ",'USD',0, " & stA1k08 & ",'" & stA1k03 & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
      strSql = "INSERT INTO ACC1K0 (A1K01,A1K02,A1K05,A1K06,A1K07,A1K09,A1K10,A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K18,A1K30,A1K08,A1K03,A1K27,A1K28,A1K04,A1K21,A1K19,A1K20,A1K33 ) " & _
               " VALUES  ('" & stA1k01 & "'," & strSrvDate(2) & ",'" & ChgSQL(stA1k05) & "',0,NULL,0," & dblUSRate & "," & stA1k11 & ",NULL,'" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "'" & _
               ",'" & strA1K18 & "',0, " & stA1k08 & ",'" & stA1k03 & "','" & strA1K27 & "','" & strA1K28 & "','" & strPrintCust & "','" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'),'" & strA1K33 & "')"
               
      cnnConnection.Execute strSql, intI
      '3:·s¼W¤@µ§ACC1L0
      strSql = "INSERT INTO ACC1L0 (A1L01,A1L03,A1L07,A1L02,A1L04,A1L05,A1L10,A1L08,A1L09) " & _
               " VALUES  ('" & stA1k01 & "','FCP'," & stA1L07 & ",'001','" & m_CP10 & "'," & stA1L05 & ",'" & strUserNum & "'," & strSrvDate(2) & ",to_char(sysdate,'hh24miss'))"
      cnnConnection.Execute strSql, intI
      
      PUB_UpdateA1k08 stA1k01 'Added by Morgan 2012/11/2 §ó·s½Ð´Ú³æ¥~¹ôª÷ÃB
      
      '4:·s¼WACC1W0
      strSql = "INSERT INTO ACC1W0 VALUES  ('" & stA1k01 & "','" & strReceiveNo & "')"
      cnnConnection.Execute strSql, intI
      '5:§ó·s·s¼WªºCÃþ¦¬¤å¸¹
      strSql = "UPDATE CASEPROGRESS SET CP60='" & stA1k01 & "' WHERE CP09='" & strReceiveNo & "'"
      cnnConnection.Execute strSql, intI
      
      PUB_PointAutoassign stA1k01, True   'Add by Morgan 2017/1/5 ¦Û°Ê¤À°tÂI¼Æ

'Removed by Morgan 2012/11/1 ¨ú®ø§ï¥Ñ°]°È³B¤H¤u³B²z--ÔÑÞ±
'      '6:­Y­ì½Ð´Ú³æ©|¥¼¦¬´Ú«h§éÅýª÷ÃB=¼f¬d³W¶O
'      'Modified by Morgan 2011/12/21 a1k31¤]­n§ó·s
'      'strSql = "update acc1k0 set a1k06=(select nvl(a1k06,0)+trunc(decode(a1k10,0,sum(a1l05),sum(a1l05)/a1k10)) from acc1l0 where a1l01=a1k01 and a1l04 in ('41699','10799')),a1k07=" & strSrvDate(2) & _
'         " where a1k29 is null and a1k01=(select c2.cp60 from caseprogress c1,caseprogress c2 where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and c2.cp10 in ('416','107'))"
'      'cnnConnection.Execute strSql, intI
'      strExc(0) = "select * from acc1k0 where a1k29 is null and a1k01 in (select c2.cp60 from caseprogress c1,caseprogress c2 where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and c2.cp10 in ('416','107'))"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If RsTemp("a1k18") = "USD" Then
'            dblXRate = Val("" & RsTemp("a1k10"))
'            If dblXRate = 0 Then dblXRate = 1
'         Else
'            dblXRate = PUB_GetUSXRate_1(RsTemp("a1k02"), RsTemp("a1k18"))
'         End If
'
'         strSql = "update acc1k0 set a1k31=(select nvl(a1k31,0)+trunc(sum(a1l05)/" & dblXRate & ") from acc1l0" & _
'            " where a1l01=a1k01 and a1l04 in ('41699','10799')),a1k07=" & strSrvDate(2) & _
'            " where a1k01='" & RsTemp("a1k01") & "'"
'         cnnConnection.Execute strSql, intI
'
'         If RsTemp("a1k18") = "USD" Then
'            dblUSRate = 1
'         Else
'            dblUSRate = PUB_GetDNRate(RsTemp("a1k02"), RsTemp("a1k18"))
'         End If
'
'         strSql = "update acc1k0 set a1k06=round(a1k31*" & dblUSRate & ",2) where a1k01='" & RsTemp("a1k01") & "'"
'         cnnConnection.Execute strSql, intI
'      'end 2011/12/21
'         PUB_PointAutoassign stA1k01, True 'Add by Morgan 2010/4/21 ¦Û°Ê¤À°tÂI¼Æ
'      End If 'Added by Morgan 2011/12/21

   End If
   'end 2009/10/12
   
   
   'Added by Morgan 2012/11/13 102·sªk
   '¥xÆW¥À®×ªì¼f®Ö­ã¥²¶·§ó·s¤À³Î®×´Á­­
   'Modifie by Morgan 2019/10/17 108.11.1 ·sªkµo©ú/·s«¬­ã«á3¤ë¤º³£¥i´£¤À³Î
   'If pa(9) = "000" And m_CP10 = "101" Then
   If (pa(8) = "1" Or pa(8) = "2") And (strKind = "101" Or strKind = "102" Or strKind = "107" Or strKind = "301" Or strKind = "302" Or strKind = "307") Then
   'end 2019/10/7
      If Val(DBDATE(Text6)) >= 20130101 Then
         strSql = "select cp09 from divisioncase,caseprogress" & _
            " where dc05='" & pa(1) & "' and dc06='" & pa(2) & "'" & _
            " and dc07='" & pa(3) & "' and dc08='" & pa(4) & "'" & _
            " and cp01(+)=dc01 and cp02(+)=dc02 and cp03(+)=dc03 and cp04(+)=dc04 and cp10='307' and cp27||cp57 is null"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            Do While Not RsTemp.EOF
               strExc(1) = PUB_Update307RefTw(RsTemp(0))
               If strExc(1) <> "" Then
                  st307Msg = st307Msg & strExc(1) & vbCrLf
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
   End If
   'end 2012/8/14
   
   'Add by Lydia 2014/12/24 (frm060104_3)¥N¿ì°h¶Oµo¤å®É,¶i«×ÀÉ¦Û°Ê²£¥Í¤@¹D¡¨¦Û½ÐºM¦^¡¨(413,BÃþ³æ),
        '·í¥N¿ì°h¶O¿é¤J®Ö­ã®É,¶i«×ÀÉ¦Û°Ê²£¥Í¤@¹D¡¨¦Û½ÐºM¦^-®Ö­ã¡¨(1001,CÃþ³æ),¦¬¤å¤é¤Îµo¤å¤é¬°¨t²Î¤é
   If m_CP10 = "908" And pa(57) = "Y" Then
        'Modified by Morgan 2013/6/6 +ÀË¬d¦A¼f©µ´Á
        'strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strCP09 & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107')"
        'Modified by Morgan 2022/10/12 +435Äò¦æ¥À®×¦A¼f
        strExc(0) = "select 1 from caseprogress a,caseprogress b where a.cp09='" & strReceiveNo & "' and b.cp09(+)=a.cp43 and b.cp10 in ('416','107','435')" & _
           " union select 2 from  caseprogress a,caseprogress b,nextprogress where a.cp09='" & strReceiveNo & "' and b.cp09(+)=a.cp43 and b.cp10='404' and np01(+)=b.cp43 and np07='107'" & _
           " union select 3 from  caseprogress a,caseprogress b,caseprogress c where a.cp09='" & strReceiveNo & "' and b.cp09(+)=a.cp43 and b.cp10='404' and c.cp09(+)=b.cp43 and c.cp10='107'"
          
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
   
             strSql = "select a.CP09 from caseprogress a ,caseprogress b where a.cp43=b.cp09(+) and a.cp01='" & pa(1) & "' " & _
                      "and a.cp02='" & pa(2) & "' and a.cp03='" & pa(3) & "' and a.cp04='" & pa(4) & "' and a.cp10='413' " & _
                      " and substr(a.cp09,1,1) = 'B' and a.cp24 is null and instr('" & NewCasePtyList & "',b.cp10)>0 "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strSql = "UPDATE CASEPROGRESS SET CP24='1',CP25='" & strSrvDate(1) & "' WHERE CP09='" & RsTemp!CP09 & "' and cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'"
               cnnConnection.Execute strSql, intI
               
                strExc(0) = AutoNo("C", 6)
                strExc(9) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
    
                strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,cp26,cp27,CP43) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
                   pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & ",'" & strExc(0) & "','1001','" & strCP12 & "','" & strCP13 & "','" & strExc(9) & "','N'," & strSrvDate(1) & ",'" & RsTemp!CP09 & "')"
                cnnConnection.Execute strSql, intI
            End If
        End If
   End If
   'end 2014/12/24
    'Added by Lydia 2015/12/31 ¥Ó½Ð¤H¬°X47794(¤T¬PÆp¥Û)¦b¿é¤Jµo©ú®×ªì¼f®Ö­ã®É,²£¥Í¤À³Îªk©w´Á­­(¤å¨ì¦¸¤é30¤Ñ)·í¤Ñªº¦æ¨Æ¾ä´£¿ô¸ê®Æ
    'Modified by Lydia 2019/07/09 §ï¦¨­­¨î®×¥ó¥N²z¤H¬°Y4779400¨Ã¥B¥u¦³¥Ó½Ð¤H¢°=X4779400¡A®Ö­ã¬°·s¥Ó½Ð®×©Î107¦A¼f¥Ó½Ð¬Ò²£¥Í¦æ¨Æ“ï¡F
    'If i = ®Ö­ã And InStr(NewCasePtyList, m_CP10) > 0 And InStr(pa(26) & "," & pa(27) & "," & pa(28) & "," & pa(29) & "," & pa(30), "X47794") > 0 Then
    If i = ®Ö­ã And InStr(NewCasePtyList & ",107", m_CP10) > 0 And ChangeCustomerL(pa(75)) = "Y47794000" And ChangeCustomerL(pa(26)) = "X47794000" And pa(27) & pa(28) & pa(29) & pa(30) = "" Then
       'Modified by Lydia 2016/05/25 ¤å¨ì¦¸¤é30¤Ñ=¦¬¥ó¤é+30¤Ñ(°Ñ¦Òfrm010002¥DºÞ¾÷Ãö¨Ó¨ç),´£«e2¤Ñ¼u¸õ¦æ¨Æ¾ä
       'strExc(1) = CompWorkDay(3, CompDate(2, 31, DBDATE(Label3(3))))
       'Modified by Lydia 2019/07/09 §ï¦¨»âÃÒªk­­«e1­Ó¤ë(¦¬¤å¤é+2­Ó¤ë)
       'strExc(1) = CompDate(2, 30, DBDATE(Label3(3)))
       strExc(1) = CompDate(1, 2, DBDATE(Label3(3)))
       strExc(3) = PUB_GetFCPSalesNo(pa(1), pa(2), pa(3), pa(4))
       If strExc(3) <> "" Then
          strExc(4) = "FCP" & Val(pa(2)) & IIf(Val(pa(3) & pa(4)) = 0, "", pa(3) & pa(4)) & "(¤T¬PÆp¥Û)¡A¥i¦¬¤å»âÃÒ"
          If PUB_AddFCPStaffCalendar(strExc(1), "1", strExc(3), strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)) Then
             'mAddSCalendar = True 'Mark by Lydia 2019/07/09
          End If
       End If
    End If
    'end 2015/12/31
    
   'Added by Lydia 2017/08/21 ¼W¥[2¦¸¶Ê¤À³Î¦æ¨Æ¾ä
   'Modified by Lydia 2019/07/30 §ï§PÂ_ªì¼f®Ö­ã+¦A¼f®Ö­ã
   'If i = ®Ö­ã And InStr(NewCasePtyList, m_CP10) > 0 And m_1stDate & m_2ndDate <> "" Then
   If i = ®Ö­ã And (m_bNewGrant = True Or m_bAgainGrant = True) And m_1stDate & m_2ndDate <> "" Then
       '°Ñ¦Ò¦æ¨Æ¾úªº17¶µ:¹w³]´£¿ô¤H­û¬°FCPºÞ¨î¤H,¥i¸Ñ°£¤H­û¬°FCPºÞ¨î¤H+²Ä1Â¾¥N
       strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
       strExc(5) = GetABS001_17(strExc(3))
       'Modified by Lydia 2017/10/20 strSql => strExc(6)
       strExc(6) = strExc(3) & IIf(strExc(5) <> "", "," & strExc(5), "")
       'Added by Lydia 2017/10/12 ¤@¯ë¶Ê¤À³Î´Á­­¼W¥[¦æ¨Æ¾ä (±Ó²ú)
       If m_2ndDate = "" Then
            strExc(4) = "¶Ê¤À³Î1¦¸(1st®Ö­ã30¤Ñ«e¤@¶g)"
            'Modified by Lydia 2017/10/20 strSql => strExc(6)
            If PUB_AddFCPStaffCalendar(m_1stDate, "1", strExc(3), strExc(4), strExc(6), "1", pa(1), pa(2), pa(3), pa(4)) Then
               m_1stDate = m_1stDate & "Y"
            End If
       Else
       'end 2017/10/12
            strExc(4) = "¶Ê¤À³Î2¦¸(1st®Ö­ã30¤Ñ«e¤@¶g)"
            'Modified by Lydia 2017/10/20 strSql => strExc(6)
            If PUB_AddFCPStaffCalendar(m_1stDate, "1", strExc(3), strExc(4), strExc(6), "1", pa(1), pa(2), pa(3), pa(4)) Then
               m_1stDate = m_1stDate & "Y"
            End If
            strExc(4) = "¶Ê¤À³Î2¦¸(2ndªk©w«e¤@¤Ñ)"
            'Modified by Lydia 2017/10/20 strSql => strExc(6)
            If PUB_AddFCPStaffCalendar(m_2ndDate, "1", strExc(3), strExc(4), strExc(6), "1", pa(1), pa(2), pa(3), pa(4)) Then
               m_2ndDate = m_2ndDate & "Y"
            End If
       End If 'end 2017/10/12
   End If
   'end 2017/08/21
   
   'Added by Morgan 2017/5/10 ¹q¤l¤½¤å
   If m_DocNo <> "" Then
      PUB_UpdateEdocRec m_DocNo, NewReceiveNo, pa(1), pa(2), pa(3), pa(4), strCP10, "1"
   'Added by Morgan 2021/6/11 ¯È¥»¤½¤å--¦ó²QµØ
   Else
      PUB_FCPOAInform NewReceiveNo, pa(1), pa(2), pa(3), pa(4), strCP10
   End If
   'end 2017/5/10
   
   'Added by Morgan 2017/8/17
   If m_bIsDualInvWithNoSelInform Then
      'ºÞ¨î´Á­­=¨t²Î¤é+3­Ó¤u§@¤Ñ=¥»©Ò´Á­­=©Ó¿ì´Á­­
      strExc(1) = CompWorkDay(3, strSrvDate(1))
      If m_bAdd1919 Then
         m_st1919CP09 = AutoNo("C", 6)
         'Modified by Morgan 2022/2/7 +CP16,CP17,CP18,CP20
         strCP16 = ""
         strCP20 = PUB_GetCP20(Text2, "1919", strCP16, pa(26) & pa(27) & pa(28) & pa(29) & pa(30), pa(75), pa(1) & pa(2) & pa(3) & pa(4))
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP09,CP10,CP12,CP13,CP14,CP16,CP17,CP18,CP20,cp26,CP43,CP48) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
            pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strExc(1) & ",'" & m_st1919CP09 & "','1919','" & strCP12 & "','" & strCP13 & "','" & strCP14 & "'," & Val(strCP16) & ",0," & (Val(strCP16) / 1000) & ",'" & strCP20 & "','N','" & NewReceiveNo & "'," & strExc(1) & ")"
         cnnConnection.Execute strSql, intI
      Else
         
         'ºÞ¨î¤H
         strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
         'Modified by Morgan 2019/7/17 Ex:FCP-50781,FCP-50782
         'strExc(4) = "¤uµ{®v»P´¼¼z§½Ápµ¸¤@®×¤G½ÐµL¾Ü¤@¨ç"
         strExc(4) = "¤uµ{®v¬O§_¤w½T»{¡G" & vbCrLf & _
            "1.¤@®×¤G½ÐµL¾Ü¤@¨ç,»P´¼¼z§½³sµ¸¤§µ²ªG" & vbCrLf & _
            "2. ­Y¥»©Ò¤w¥h¨ç°µ¾Ü¤@°Ê§@,½Ð³qª¾µ{§Ç¤H­û,¤º³¡¦¬¤å""¾Ü¤@¥Ó´_ """
         'end 2019/7/17
         PUB_AddFCPStaffCalendar strExc(1), "1", strExc(3) & "," & strCP14, strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4)
         
         
         'EMail ©Ó¿ì¤uµ{®v,°Æ¥»:¤uµ{®v¥DºÞ¡Bµ{§ÇºÞ¨î¤H­û¡Bµ{§Ç¥DºÞ
         '¥D¦®
         strExc(4) = "½Ð¤uµ{®v½T»{" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & "(µo©ú®×)¤@®×¤G½Ð¨Æ©y"
         'Add by Amy 2025/08/05 «áÄò­ã»éÂ²³æ³ø§i=Y,¿éCÃþ¨Ó¨ç[¥D¦®]³Ì«e­±¥[¡i½ÐÂ²³æ³ø§i¡j-Winfrey
         If pa(89) = "Y" Then strExc(4) = "¡i½ÐÂ²³æ³ø§i¡j" & strExc(4)
         
         '¤º¤å
         strExc(0) = "¤uµ{®v¡G1.¥»®×¬°¤@®×¤G½Ð(µo©ú®×)¦ýµL¾Ü¤@¨ç¡A¥B®Ö­ã¨çµL""«DÄÝ¬Û¦P³Ð§@""Án©ú¡A½Ð»P´¼¼z§½³sµ¸½T»{¡C" & vbCrLf & _
                     "¡@¡@¡@¡@2.­Y¥»©Ò¤w¥h¨ç°µ¾Ü¤@°Ê§@,½Ð¥Î¦¹Email ¦^ÂÐµ{§Ç¤H­û" & vbCrLf & vbCrLf & _
                     "µ{§Ç¤H­û¡G­Y¤uµ{®v¦^ÂÐµ²ªG¬°¥H¤W2ªÌ,½Ð¤º³¡¦¬¤å239""¾Ü¤@¥Ó´_"",¨Ã¤â°Êµo¤å(µo¤å¤é111111¡A¤£½Ð´Ú""N""),¿ï¾Ü©ñ±ó·s«¬,¤Î¸Ñ°£¦æ¨Æ¾ä´Á­­¡C"
            
         '¤uµ{®v¥DºÞ
         strExc(5) = PUB_GetFCPEngSup(strCP14)
         'µ{§Ç¥DºÞ
         strExc(6) = PUB_GetFCPProSup(strExc(3))
         
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
            " values( '" & strUserNum & "','" & strCP14 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strExc(4) & "','" & strExc(0) & "','" & strExc(3) & ";" & strExc(5) & ";" & strExc(6) & "')"

         cnnConnection.Execute strSql, intI
         'end 2019/7/17
      End If
   End If
   'end 2017/8/17
   
   'Added by Lydia 2022/04/29  FCP®×Key®Ö­ã(¬ÛÃö¦¬¤å¸¹¬O±¾·s®×101,102,103,107, 307,308)½T©w«á¡A§PÂ_¬O§_¤w¸g½Ð´Ú¦p¤U¡A¸Ô²Ó¤º®e¥i°Ñ¦Òªþ¥ó
   If i = ®Ö­ã And pa(1) = "FCP" And InStr("101,102,103,107,307,308", m_CP10) > 0 And strCP14 <> "" Then
      'Modified by Lydia 2023/02/02 cp43=> nvl(cp43,'N')
      'Modified by Lydia 2023/10/31 §ï¦¨¼Ò²Õ
'      strExc(0) = "select cp09,cp60,cp14 from caseprogress where cp09= (" & _
'                        "select max(cp09) mno from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 " & _
'                        "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' " & _
'                        "and cp05 = (select max(cp05) mdate from caseprogress, staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp159=0 " & _
'                        "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' and nvl(cp43,'N') <> '" & m_NewReceiveNo & "' )) "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'        strExc(9) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
'        strExc(6) = Pub_GetSpecMan("¥~±M§i­ãµ{§Ç") 'A4029 ¾Gµú¤ß
'        If "" & RsTemp.Fields("CP60") = "" Then
'             '1.¤W¤@¹D¤uµ{®v®×¥ó©Ê½è¥¼¦³½Ð´Ú³æ¸¹¡A«h¦Û°ÊµoMail
'             '¦¬¥óªÌ: ¤uµ{®v   °Æ¥»¦¬¨üªÌ: ¤uµ{®v¤§¥DºÞ;µ{§ÇºÞ¨î¤H­û(Key¨Ó¨ç¤H­û¤£¬OºÞ¨î¤H­û¤]¦C¤J¦¬¥óªÌ);¾Gµú¤ß;backup
'            '¥D¦®: ¥»®×¤w®Ö­ã¡A½Ð¤uµ{®v¾¨³t³B²z½Ð´Ú¡A¥H§Q«áÄò§i­ã¬yµ{Our Ref: FCP-060000 [INCOM.1001]
'             strExc(2) = PUB_GetFCPEngSup(RsTemp.Fields("CP14"))
'             '¥D¦®
'             strExc(4) = "¥»®×¤w®Ö­ã¡A½Ð¤uµ{®v¾¨³t³B²z½Ð´Ú¡A¥H§Q«áÄò§i­ã¬yµ{Our Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " [INCOM." & ®Ö­ã & "]"
'             'CC
'             strExc(6) = strExc(2) & ";" & strExc(9) & IIf(strExc(9) <> strUserNum, ";" & strUserNum, "") & IIf(strUserNum <> strExc(6) And strExc(9) <> strExc(6) And strExc(6) <> "", ";" & strExc(6), "") & ";backup"
'             strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                    " values( '" & strUserNum & "','" & RsTemp.Fields("CP14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                     ",'" & strExc(4) & "','¦p¦®','" & strExc(6) & "')"
'             cnnConnection.Execute strSql, intI
'         Else
'             '¤W¤@¹D¤uµ{®v®×¥ó©Ê½è¤w¦³½Ð´Ú³æ¸¹¦ý¨÷©v°ÏµLREPDN(±H½Ð´Ú¨ç) or DNUPL(½Ð´Ú³æ¤W¶Ç) (¦³¤@¶µ´N¤£µoemail)¡A«h¦Û°ÊµoMail:
'             '¦¬¥óªÌ: µ{§ÇºÞ¨î¤H­û (Key¨Ó¨ç¤H­û¤£¬OºÞ¨î¤H­û¤]¦C¤J¦¬¥óªÌ)  °Æ¥»¦¬¨üªÌ: µ{§ÇºÞ¨î¤H­û¥DºÞ;¾Gµú¤ß;backup
'             '¥D¦®: ¥»®×¤w®Ö­ã¡A½Ðµ{§Ç¾¨³t³B²z½Ð´Ú¡A¥H§Q«áÄò§i­ã¬yµ{Our Ref: FCP-060000 [INCOM.1001]
'             'Modified by Lydia 2022/05/06 ­×§ï¦¨:­Y¦P¤@­Ó½Ð´Ú³æ¸¹ªº¨÷©v°Ï¸Ì­±µLREPDN(±H½Ð´Ú¨ç) or DNUPL(½Ð´Ú³æ¤W¶Ç)  (¦³¤@¶µ´N¤£µoemail)¡A«h¦Û°ÊµoMail; ex.FCP-059520(AB1009611,AB1009612,CB1002423)
'             'strExc(0) = "SELECT CPP01, CPP02 FROM CASEPAPERPDF B " & _
'                               "WHERE CPP01='" & RsTemp.Fields("CP09") & "' AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.REPDN.%' OR UPPER(CPP02) LIKE '%.DNUPL.%' ) "
'             strExc(0) = "SELECT CPP01, CPP02 FROM CASEPAPERPDF B " & _
'                               "WHERE CPP01 in (select cp09 from caseprogress where cp60='" & RsTemp.Fields("CP60") & "')  AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.REPDN.%' OR UPPER(CPP02) LIKE '%.DNUPL.%' ) "
'             intI = 1
'             Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'             If intI = 0 Then
'                 strExc(2) = PUB_GetFCPProSup(strExc(9))
'                 '¥D¦®
'                 strExc(4) = "¥»®×¤w®Ö­ã¡A½Ðµ{§Ç¾¨³t³B²z½Ð´Ú¡A¥H§Q«áÄò§i­ã¬yµ{Our Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) = "000", "", "-" & cp(3) & "-" & cp(4)) & " [INCOM." & ®Ö­ã & "]"
'                 'CC
'                 strExc(6) = strExc(2) & IIf(strUserNum <> strExc(6) And strExc(9) <> strExc(6) And strExc(6) <> "", ";" & strExc(6), "") & ";backup"
'                 strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'                        " values( '" & strUserNum & "','" & strExc(9) & IIf(strExc(9) <> strUserNum, ";" & strUserNum, "") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                         ",'" & strExc(4) & "','¦p¦®','" & strExc(6) & "')"
'                 cnnConnection.Execute strSql, intI
'             End If
'          End If
'      End If
      Call PUB_ChkFCPtoDNUPL(pa(1), pa(2), pa(3), pa(4), i, m_NewReceiveNo)
      'end 2023/10/31
   End If
   'end 2022/04/29
   
   'Added by Lydia 2023/07/28 ¥~±M-FCP±M§Q³sµ²®×ºÞ¨î¡G±M§QÅv©µªø415¡B§ó¥¿402µoEmail³qª¾¤uµ{®v¡A¨Ã¥B¦Û°Ê³]¦æ¨Æ¾äºÞ±±¨â¤Ñ¡A·íµ{§Ç½T»{¤½³ø¥Z¸ü¤é´Á«á¸Ñ°£¦æ¨Æ¾ä¦Û°Ê¦¬¤å¡u³qª¾¸ê°TÅÜ§ó961¡v,µo¤@«ÊEmailµ¹©Ó¿ì¤uµ{®v
   If pa(177) = "Y" And i = ®Ö­ã And (m_CP10 = "415" Or m_CP10 = "402") Then
      '´Á­­: 2¤Ñ(¤é¾ä¤Ñ
      'Modified by Lydia 2023/08/25 §ï5¤Ñ¤é¾ä¤Ñ
      strExc(1) = CompDate(2, 5, strSrvDate(1))
      '´£¿ô¤H­û: µ{§Ç , ¤uµ{®v ; ¸Ñ°£¤H­û: µ{§Ç¡Bµ{§Ç®×¥óÂ¾¥N
      strExc(3) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
      '¨Æ¥Ñ
      strExc(4) = "½Ðµ{§Ç½T»{" & Label3(1) & "¤½§i¤½³ø¥Z¸ü¤é´Á"
      PUB_AddFCPStaffCalendar strExc(1), "1", strExc(3) & "," & strCP14, strExc(4), strExc(3), "1", pa(1), pa(2), pa(3), pa(4), , , , NewReceiveNo
      
      'Modified by Lydia 2023/12/11 ¦]¬°±M§QÅv©µªø415·|Åã¥ÜFrame1¡A³y¦¨¹w³]©Ó¿ì¤HÅÜ¦¨"¥~±Mµ{§Ç-°É»~§¹³Æ"¡A©Ò¥H§ï¦b¼Ò²Õ¤º¨ú±o³Ì·s¤@¹D¤§¤uµ{®v; ex.FCP-51563
      'If PUB_GetFCPlinkMC("2", TransDate(Label3(3).Caption, 2), pa, strReceiveNo, m_CP10, "" & i, strCP12, strCP13, strCP14) = True Then
      'Mark by Lydia 2024/01/04 µ{§Ç¸Ñ°£¦æ¨Æ¾ä«á¦A¦¬¤å¸ê°TÅÜ§ó¨Ãª½±µÅã¥Ü¥¿½Tªºªk­­
      'If PUB_GetFCPlinkMC("2", TransDate(Label3(3).Caption, 2), pa, strReceiveNo, m_CP10, "" & i, strCP12, strCP13) = True Then
      'End If
      'end 2024/01/04
      'Added by Lydia 2024/04/10 ¼W¥[Email³qª¾©M§i¥N901
      '1.®Ö­ã¨º¹Dªºµo¤å¤é¬°ªÅ¡A­×§ï©Ó¿ì¤H¬°µ{§Ç
      strExc(0) = Pub_GetSpecMan("¥~±Mµ{§Ç-°É»~§¹³Æ")
      If strExc(0) <> "" Then
         strSql = "Update CaseProgress set cp27=null,cp14='" & strExc(0) & "' where cp09='" & NewReceiveNo & "' "
         cnnConnection.Execute strSql
         '2.¦P®É¤º³¡¦¬¤å901¡A©Ó¿ì¤H±¾¤uµ{®v¡A©Ó¿ì´Á­­+1¶g¡A¥»©Ò´Á­­+2¶g
         strExc(0) = PUB_GetFCPPromoterNo(strReceiveNo, "1001")
         If strExc(0) = "" Then strExc(0) = m_CP14
         strExc(3) = CompWorkDay(1, CompDate(2, 14, strSrvDate(1)), 1) '¥»©Ò´Á­­+2¶g
         strExc(4) = CompWorkDay(1, CompDate(2, 7, strSrvDate(1)), 1) '©Ó¿ì´Á­­+1¶g
         
         strExc(1) = AutoNo("B", 6)
         strSql = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP43,CP48) VALUES ('" & pa(1) & "','" & pa(2) & "','" & _
            pa(3) & "','" & pa(4) & "'," & strSrvDate(1) & "," & strExc(3) & ",'" & strExc(1) & "','901','" & strCP12 & "','" & strCP13 & "','" & strExc(0) & "','N','N','" & NewReceiveNo & "'," & strExc(4) & ")"
         cnnConnection.Execute strSql, intI
         
         'CC: ¤uµ{®v¥DºÞ¡Bµ{§ÇºÞ¨î¤H­û¡Bµ{§Ç¥DºÞ¡Bbackup
         strExc(1) = PUB_GetFCPEngSup(strExc(0))
         strExc(2) = PUB_GetFCPHandler(pa(1), pa(2), pa(3), pa(4))
         strExc(3) = PUB_GetFCPProSup(strExc(2))
         strExc(6) = ";" & strExc(1) & ";" & strExc(2) & ";" & strExc(3)
         strExc(6) = Mid(strExc(6), 2) & ";backup"
   
         '¥D¦®
         strExc(4) = "¡i½Ð³ø§i®Ö­ã-" & Label3(1) & "¡jOur Ref:" & pa(1) & "-" & pa(2) & IIf(pa(3) & pa(4) <> "000", "-" & pa(3) & "-" & pa(4), "") & "[INCOM." & i & "] (¦³±M§Q³sµ²®×)"
         'Add by Amy 2025/08/05 «áÄò­ã»éÂ²³æ³ø§i=Y,¿éCÃþ¨Ó¨ç[¥D¦®]³Ì«e­±¥[¡i½ÐÂ²³æ³ø§i¡j-Winfrey
         If pa(89) = "Y" Then strExc(4) = "¡i½ÐÂ²³æ³ø§i¡j" & strExc(4)
         
         strExc(5) = "¤uµ{®v½Ð¶i¦æ¥H¤U¨Æ¶µ:" & vbCrLf & _
                     "¥D¦®: ³ø§i ®Ö­ã-" & Label3(1) & vbCrLf
         If txt415Date.Visible = True And Trim(txt415Date) <> "" Then 'Added by Lydia 2024/12/02 FCP-059682¤£¥Î¿é¤J±M§QÅv´Á¶¡©µªø
            strExc(5) = strExc(5) & "¤º®e: ±M§QÅv´Á¶¡­ã¤©©µªø" & Mid(DBDATE(txt415Date), 1, 4) - 1911 & "¦~" & Mid(DBDATE(txt415Date), 5, 2) & "¤ë" & Mid(DBDATE(txt415Date), 7, 2) & "¤é¤î¡C"
         End If
         strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                " values( '" & strUserNum & "','" & strExc(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                 ",'" & ChgSQL(strExc(4)) & "','" & ChgSQL(strExc(5)) & "','" & strExc(6) & "')"
         cnnConnection.Execute strSql, intI
         '·íµ{§Ç¸Ñ°£¦æ¨Æ¾ä´Á­­®É¡A¨t²Î·|¼uµøµ¡¿é¤J¤½§i¤é¡A½Ð¦Û°Ê±N¤½³ø¥Z¸ü¤é´Á¤@¨Ö±¾¦b®Ö­ã¨º¹Dªº©Ó¿ì´Á­­¡C
      End If
      'end 2024/04/10
   End If
   'end 2023/07/28
   
   cnnConnection.CommitTrans
   FormSave = True
   
    'Added by Lydia 2016/11/17 ¥H½Ð´Ú¹ï¶HÀË¬d¬O§_¦s¦b©ó°ê¥~©T©w±H¶Ê´Ú³æ¥N²z¤HÀÉ(ACC225)¥B¤U¦¸±Hµo¤é´Á¡Ö¨t²Î¤é¡A­Y¦s¦b«hÅã¥Ü°T®§´£¿ô¾Þ§@¤H­û
    If stA1k01 <> "" And strA1K28 <> "" Then
       If PUB_ChkAcc225MsgList(stA1k01, strA1K28, pa(1), pa(2), pa(3), pa(4)) Then
       End If
    End If
    'end 2016/11/17
       
   If st307Msg <> "" Then MsgBox st307Msg 'Add by Morgan 2012/11/13
   
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   
ErrHnd:
   FormSave = False
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
   
End Function

Private Sub Combo1_Click(Index As Integer)
 Dim i As Integer, strTmp As String
   If Combo1(Index) = "" Then
      For i = 0 To 2
         Text33(i + Index * 3) = ""
      Next
      Exit Sub
   End If
   
   strTmp = Mid(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") + 1, 1)
   Select Case Text2
      Case "FCP"
         If pa(75) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07,FA08,FA09"
               Case "2"
                  strExc(1) = "FA52,FA53,FA54"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58,CU59,CU60"
               Case "2"
                  strExc(1) = "CU61,CU62,CU63"
            End Select
         End If
      Case "FG"
         If pa(26) <> "" Then
            Select Case strTmp
               Case "1"
                  strExc(1) = "FA07"
               Case "2"
                  strExc(1) = "FA52"
            End Select
         Else
            Select Case strTmp
               Case "1"
                  strExc(1) = "CU58"
               Case "2"
                  strExc(1) = "CU61"
            End Select
         End If
   End Select
   
   strExc(2) = ChgFagent(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   strExc(3) = ChgCustomer(Left(Combo1(Index).Text, InStr(Combo1(Index).Text, "-") - 1))
   Select Case Text2
      Case "FCP"
         If pa(75) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
      Case "FG"
         If pa(26) <> "" Then
            strExc(0) = "SELECT " & strExc(1) & " FROM FAGENT WHERE " & strExc(2)
         Else
            strExc(0) = "SELECT " & strExc(1) & " FROM CUSTOMER WHERE " & strExc(3)
         End If
   End Select
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Select Case Text2
         Case "FCP"
            For i = 0 To 2
               If Not IsNull(RsTemp.Fields(i)) Then
                  Text33(i + Index * 3) = RsTemp.Fields(i)
               Else
                  Text33(i + Index * 3) = ""
               End If
            Next
         Case "FG"
            If Not IsNull(RsTemp.Fields(0)) Then Text33(0) = RsTemp.Fields(0)
      End Select
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = °ê¥~_FC
   ReDim pa(TF_PA)
   With frm06010602_2
      pa(1) = .Text2
      pa(2) = .Text3
      pa(3) = .Text4
      pa(4) = .Text5
      strReceiveNo = .Tag
      strSales = strExc(5)
      ReadPatent
      'Move by Lydia 2019/07/30 ±qcombo2¤U­±
      Label3(3) = frm06010602_1.Text5
      Label3(2) = strReceiveNo
      'end 2019/07/30
      mAddSCalendar = False 'Added by Lydia 2015/12/31
      SetDivSug 'Added by Morgan 2012/12/13
      
      'Added by Morgan 2024/5/17
      If m_CP10 = "421" Then
         Label8.Visible = True
         Text16.Visible = True
         LblFM2(1).Visible = True
         Text16 = PUB_GetFCPPromoterNo(strReceiveNo, "1008")
         Text16_Validate False
      End If
      'end 20224/5/17
   End With
   Combo2.ListIndex = 0
   
   Call GRIDHEAND 'Add By Sindy 2017/6/27
   SSTab1.Tab = 0 'Add By Sindy 2017/6/27
   Frame1.BackColor = &H8000000F 'Added by Lydia 2019/05/23
   
Dim strTmp As String
   If Len(strSrvDate(2)) = 6 Then
      strTmp = Left(strSrvDate(2), 2)
   Else
      strTmp = Left(strSrvDate(2), 3)
   End If
   Text7.Text = "¡]" & strTmp & "¡^´¼±M¤@¡]¤G¡^¦r²Ä¸¹"
   
   'Added by Morgan 2017/5/10 ¹q¤l¤½¤å
   If m_DocWord <> "" Then
      Text7 = m_DocWord & "¦r²Ä" & m_DocNo & "¸¹"
   ElseIf m_DocNo <> "" Then
      Text7 = Replace(Text7, "²Ä¸¹", "²Ä" & m_DocNo & "¸¹")
   End If
   If m_DocDate <> "" And Text6.Locked = False Then
      Text6 = TransDate(m_DocDate, 1)
   End If
   'end 2017/5/10
   
   Check908 pa 'Add by Morgan 2009/10/1
   
End Sub

Private Sub ReadPatent()
Dim Lbl As Control, i As Integer, j As Integer
Dim strTmp(0 To 5) As String
   
   For Each Lbl In Label3
      Lbl.Caption = ""
   Next
   'Added by Lydia 2021/10/01
   For Each Lbl In LblFM2
      Lbl.Caption = ""
   Next
   'end 2021/10/01
   
   Text2 = pa(1)
   Text3 = pa(2)
   Text4 = pa(3)
   Text5 = pa(4)
   'Modify by Morgan 2006/10/20
   'If clspdReadPatentDatabase(pA(), intWhere) Then 'edit by nickc 2007/02/02 ¤£¥Î dll ¤F If objPublicData.ReadPatentDatabase(pA(), intWhere) Then
   If PUB_ReadPatentDatabase(pa(), intWhere) Then
      LblFM2(0) = pa(5)
      For i = 5 To 7
         Text9(i - 5) = pa(i)
      Next
      Text1 = pa(11)
      'µe­±¤T PA(89),PA(17),PA(57)
      If pa(16) = "1" Then
         Label3(6) = "°ò¥»ÀÉ¥Ø«e­ã»é : ­ã"
      ElseIf pa(16) = "2" Then
         Label3(6) = "°ò¥»ÀÉ¥Ø«e­ã»é : »é"
      'Modified by Lydia 2019/05/23
      'ElseIf pa(16) = "2" Then
      Else
         Label3(6) = "°ò¥»ÀÉ¥Ø«e­ã»é : µL"
      End If
      Text10(1) = pa(17)
      Text10(2) = pa(57)
      Label3(4) = pa(89)
      'µe­±¥| 48, 51,52,53,54,55,56, 101,102,103,104
      Text12 = pa(48)
     
      If pa(101) <> "" Then
         Text19 = pa(101)
         ChgType (5)
      End If
      Text20 = pa(102)
      Text21 = pa(103)
      Text22 = pa(104)
      If Left(pa(26), 6) = "X27766" And pa(101) <> "" And pa(103) = "" And pa(104) = "" Then
         Text21 = "*Murata's reference number for the U.S. Patent application is"
         Text22 = "*Corresponding Japanese Patent Application number"
      End If
      'Add By Sindy 2017/6/27
      '¥Ó½Ð¤é
      If pa(10) <> "" Then
         Label3(8) = pa(10)
      End If
      '¥Ó½Ð¤H
      Text33(9) = "": Label27(0) = ""
      Text33(10) = "": Label27(1) = ""
      Text33(11) = "": Label27(2) = ""
      Text33(12) = "": Label27(3) = ""
      Text33(13) = "": Label27(4) = ""
      For i = 0 To 4
         If pa(i + 26) <> "" Then
            Text33(i + 9) = ChangeCustomerL(pa(i + 26))
            Label27(i).Caption = GetPrjPeople1(Text33(i + 9))
         End If
      Next
      '2017/6/27 END
      
      Combo1(0).Clear
      Combo1(1).Clear
      Combo1(0).AddItem ""
      Combo1(1).AddItem ""
      
      For i = 0 To 5
         Text33(i) = pa(i + 51)
      Next
      Text33(6) = pa(139) 'Add by Morgan 2006/10/20
      
      If pa(75) <> "" Then
         Select Case pa(85)
            Case 1
               strExc(0) = "FA07,FA52"
            Case 2
               strExc(0) = "FA08,FA53"
            Case 3
               strExc(0) = "FA09,FA54"
            Case Else
               strExc(0) = "FA08,FA53"
         End Select
         
         strExc(0) = "SELECT " & strExc(0) & " FROM FAGENT WHERE " & ChgFagent(pa(75))
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If IsNull(RsTemp.Fields(0)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(0)
            End If
            Combo1(0).AddItem pa(75) & "-1" & strExc(0)
            Combo1(1).AddItem pa(75) & "-1" & strExc(0)
            If IsNull(RsTemp.Fields(1)) Then
               strExc(0) = ""
            Else
               strExc(0) = "-" & RsTemp.Fields(1)
            End If
            Combo1(0).AddItem pa(75) & "-2" & strExc(0)
            Combo1(1).AddItem pa(75) & "-2" & strExc(0)
         End If
      Else
         For i = 26 To 30
            If pa(i) <> "" Then
               Select Case pa(85)
                  Case 1
                     strExc(0) = "CU58,CU61"
                  Case 2
                     strExc(0) = "CU59,CU62"
                  Case 3
                     strExc(0) = "CU60,CU63"
                  Case Else
                     strExc(0) = "CU59,CU62"
               End Select
               strExc(0) = "SELECT " & strExc(0) & " FROM CUSTOMER WHERE " & ChgCustomer(pa(i))
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  For j = 1 To 2
                     If IsNull(RsTemp.Fields(j - 1)) Then
                        strExc(0) = ""
                     Else
                        strExc(0) = "-" & RsTemp.Fields(j - 1)
                     End If
                     Combo1(0).AddItem pa(i) & "-" & j & strExc(0)
                     Combo1(1).AddItem pa(i) & "-" & j & strExc(0)
                  Next
               End If
            End If
         Next
      End If
   End If
   
   'Modified by Moran 2019/12/31 +CP05
   strExc(0) = "SELECT CP10,CPM03,CP12,CP13,CP14,CP54,CP50,cp19,CP05 FROM CASEPROGRESS,CASEPROPERTYMAP WHERE " & _
      "CP09='" & strReceiveNo & "' AND CP01=CPM01(+) AND CP10=CPM02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 ¤£¥Î dll ¤F objLawDll.ReadRstMsg(intI, strExc(0))
   With RsTemp
      If intI = 1 Then
         If Not IsNull(.Fields(0)) Then strKind = .Fields(0)
         If Not IsNull(.Fields(1)) Then Label3(1) = .Fields(1)
         For i = 2 To 6
            If Not IsNull(.Fields(i)) Then cp(i) = .Fields(i)
         Next
         txtCP19.Tag = "" & .Fields("cp19") 'Add by Morgan 2009/10/13
         If .Fields("cp05") = 19221111 Then m_bMiddleCase = True 'Added by Morgan 2019/12/31
      End If
   End With
   
   ' 90.06.26 modify by louis
   m_CP10 = Empty
   If IsNull(RsTemp.Fields("CP10")) = False Then
      m_CP10 = RsTemp.Fields("CP10")
   End If
   ' 92.1.19 add by sonia
      
   m_CP14 = Empty
   If IsNull(RsTemp.Fields("CP14")) = False Then
      m_CP14 = RsTemp.Fields("CP14")
   End If
   '­Y®×¥ó©Ê½è¬°Á|µoµª¿ë(804)
   If m_CP10 = "804" Then
      EnableTextBox Text10(1), True
      'Add By Cheng 2001/12/20
      'Åã¥Ü±M§QÅv¬O§_¦s¦b¶µ¥Ø
      Me.Label9(2).Visible = True
      Me.Text10(1).Visible = True
   Else
      EnableTextBox Text10(1), False
      'Add By Cheng 2001/12/20
      'ÁôÂÃ±M§QÅv¬O§_¦s¦b¶µ¥Ø
      Me.Label9(2).Visible = False
      Me.Text10(1).Visible = False
   End If
   
   'Added by Morgan 2023/2/23
   '±M§QÅv©µªø
   If m_CP10 = "415" Then
      lbl415Date.Visible = True
      txt415Date.Visible = True
   Else
      lbl415Date.Visible = False
      txt415Date.Visible = False
   End If
   'end 2023/2/23
   
   'Added by Lydia 2025/02/12
   If m_CP10 = "245" Then
      lbl415Date.Visible = True
      txt415Date.Visible = True
      lbl415Date.Caption = "Äò¦æ¼f¬d¤é´Á¡G"
   End If
   'end 2025/02/12
   
   ' 90.06.27 modify by louis «D¥Ó½Ð®×¤Î«D§ï½Ð®×¤£¿é¤J¥Ó½Ð®Ö­ã¤é®³±¼
   'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
   If (m_CP10 < "101" Or m_CP10 > "105") And Mid(m_CP10, 1, 1) <> "3" And m_CP10 <> "107" And m_CP10 <> "125" Then
      EnableTextBox Text6, False
   Else
      EnableTextBox Text6, True
   End If
   'Add By Cheng 2002/07/23
   EnableTextBox Text10(1), False
   Me.Text10(1).Text = "" & pa(17)
      
   'MODIFY BY SONIA 90.11.4
   EnableTextBox Text10(0), False
   'Modify By Cheng 2002/07/23
'   Text10(0) = ""
   Text10(0) = "" & pa(16)
   Select Case m_CP10
      Case µo©ú¥Ó½Ð, ·s«¬¥Ó½Ð, ³]­p¥Ó½Ð, °l¥[¥Ó½Ð, Áp¦X¥Ó½Ð, µªÅG
         'Modify By Cheng 2002/07/23
'         Text10(0) = "Y"
      Case §ï½Ðµo©ú, §ï½Ð·s«¬, §ï½Ð³]­p, §ï½Ð°l¥[, §ï½ÐÁp¦X, §ï½Ð¿W¥ß, ¤À³Î
'         Text10(0) = "Y"
      Case ²§Ä³_±M, Á|µo
'         Text10(0) = "Y"
      Case ²§Ä³µªÅG, Á|µoµªÅG
'         Text10(0) = "Y"
   End Select
   
   'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
   If (m_CP10 >= "101" And m_CP10 <= "105") Or m_CP10 = "107" Or m_CP10 = "125" Or (m_CP10 >= "301" And m_CP10 <= "308") Or m_CP10 = "802" Or m_CP10 = "804" Then
      Me.Text10(0).Text = "1"
   End If
   If m_CP10 = "804" Then
      Me.Text10(1).Text = "Y"
   End If
   'Add By Cheng 2002/07/03
   If m_CP10 = ÅÜ§ó Then
      Me.cmdMod.Visible = True
   End If
   
   'Add by Morgan 2009/10/12
   m_bPrintFlowSheet = False
   m_bAddAcc1k0 = False
   m_bNoDN = False 'Added by Morgan 2014/4/24
   'Modified by Morgan 2015/4/24 °h¶O¤£½Ð´Ú®É¤£²£¥ÍD/N--David
   If m_CP10 = "908" Then
      m_bPrintFlowSheet = True
      'Modified by Morgan 2022/10/12 +435Äò¦æ¥À®×¦A¼f
      strExc(0) = "select 1,c1.cp60,c1.cp20 from caseprogress c1,caseprogress c2" & _
         " where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and c2.cp10 in ('416','107','435')"
      'Added by Morgan 2013/6/28 +¦A¼f©µ´Á(¦A¼f¨S¦³¦¬¤å)
      strExc(0) = strExc(0) & " union select 2,c1.cp60,c1.cp20 from caseprogress c1,caseprogress c2,nextprogress" & _
         " where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and c2.cp10='404' and np01(+)=c2.cp43 and np07='107'"
      'end 2013/6/28
      'add by sonia 2015/4/7 +¦A¼f©µ´Á(¦A¼f¥ý¦¬¤å¤~©µ´Á)FCP-034520
      strExc(0) = strExc(0) & " union select 3,c1.cp60,c1.cp20 from caseprogress c1,caseprogress c2,caseprogress c3" & _
         " where c1.cp09='" & strReceiveNo & "' and c2.cp09(+)=c1.cp43 and c2.cp10='404' and c3.cp09(+)=c2.cp43 and c3.cp10='107'"
      'end 2015/4/7
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Added by Morgan 2015/4/24
         If RsTemp("cp20") = "N" Then
            m_bNoDN = True
         Else
         'end 2015/4/24
            If IsNull(RsTemp("cp60")) Then
               m_bAddAcc1k0 = True
            End If
         End If 'Added by Morgan 2015/4/24
      End If
      lblCP19.Visible = True
      txtCP19.Visible = True
   Else
      lblCP19.Visible = False
      txtCP19.Visible = False
   End If
   
   'Add By Sindy 2017/6/27
   strSql = "select PD05 AS  Àu¥ýÅv¤é,PD06 AS Àu¥ýÅv¸¹,NA03 AS Àu¥ýÅv°ê®a,PD09 as Àu¥ýÅv¦s¨ú½X,PA01||PA02||PA03||PA04 AS ¥»©Ò®×¸¹ " & _
            "From PRIDATE, Nation, PATENT " & _
            "WHERE PD01='" & pa(1) & "' AND PD02='" & pa(2) & "' AND PD03='" & pa(3) & "' AND PD04 ='" & pa(4) & "' AND PD07=NA01(+) " & _
            "AND PD06=PA11(+) AND PD05=PA10(+) AND PD07=PA09(+) ORDER BY PD01,PD02,PD03,PD04 "
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   Set grdDataList2.Recordset = adoRecordset
   CheckOC
   '2017/6/27 END
   
   'Added by Lydia 2019/05/23 °É»~¤½³ø±±ºÞ¡G¦³±¾¤½§i¤½³ø(1228)¤§§ó¥¿,§ó§ï(402,403)ªº ®Ö­ã ¿é¤J°É»~¤é´Á
   Frame1.Visible = False
   '¦]¬°¤é¤å²Õªº§ó¥¿(402)³£¥H®Ö­ã¨ç¦V«È¤á³ø§i¡A©Ò¥H±q®Ö­ã¿é¤J¬Ò¤£¨«°É»~¤½³ø±±ºÞ
   'Modified by Lydia 2024/05/30 ²{¦b¤é¤å²Õªº§ó¥¿(402)¤]¨Ï¥Î°É»~ªí¡A©Ò¥H®³±¼¡u¤£¨«°É»~±±ºÞ¡v³o¶µ±±¨î¡C
   'If frm06010602_2.Text6 = "1" And ((pa(150) <> "3" And (m_CP10 = §ó¥¿ Or m_CP10 = §ó§ï)) Or (pa(150) = "3" And m_CP10 = §ó§ï)) Then
   If frm06010602_2.Text6 = "1" And (m_CP10 = §ó¥¿ Or m_CP10 = §ó§ï) Then
       strSql = "select c2.cp09,c2.cp10 from caseprogress c1, caseprogress c2 where c1.cp09='" & strReceiveNo & "' and c1.cp43=c2.cp09(+) and c2.cp10='1228' "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            Frame1.Visible = True
        End If
   End If
   'Added by Lydia 2023/08/25 ±M§QÅv©µªø415: ¨S¦³¤½§i¤½³ø
   If frm06010602_2.Text6 = "1" And pa(150) <> "3" And m_CP10 = "415" Then
      Frame1.Visible = True
      Opt1(0).Visible = False: Opt1(1).Visible = False: Opt1(2).Visible = False
      Label35 = Label35 & " " & Label3(1)
      Label33 = "¤½§i¤é´Á:"
   End If

End Sub

'Add By Sindy 2017/6/27
Private Function GRIDHEAND()
   With grdDataList2
   .row = 0
   .col = 0
   .ColWidth(0) = 1000
   .Text = "Àu¥ýÅv¤é"
   .col = 1
   .ColWidth(1) = 3000
   .Text = "Àu¥ýÅv¸¹"
   .col = 2
   .ColWidth(2) = 1000
   .Text = "Àu¥ýÅv°ê®a"
   .col = 3
   .ColWidth(3) = 1300
   .Text = "Àu¥ýÅv¦s¨ú½X"
   .col = 4
   .ColWidth(4) = 1300
   .Text = "¥»©Ò®×¸¹"
   End With
End Function

Private Function ChgType(i As Integer) As Boolean
 Dim strTempName As String
   ChgType = False
   Select Case i
      Case 5
         strExc(0) = Text19.Text
         'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
         'If objLawDll.LawGetName(strExc(0), strTempName, 1) Then
         If ClsLawLawGetName(strExc(0), strTempName, 1) Then
            Text19 = strExc(0)
            LblFM2(2) = strTempName
            ChgType = True
         End If
   End Select
End Function

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Morgan 2021/6/11
   PUB_KillTempFile pa(1) & pa(2) & "*.*" 'Added by Lydia 2018/12/17 ²M°£¼È¦sÀÉ
   
   Set frm06010602_3 = Nothing
End Sub

Private Sub Combo2_Click()
   Select Case Combo2
      Case "¤¤"
         LblFM2(0) = pa(5)
      Case "­^"
         LblFM2(0) = pa(6)
      'Modified by Lydia 2022/04/25 ¡u¤é¤å¦WºÙ¡v§ï¬°¡u¥~¤å¦WºÙ¡v
      Case "¥~"
         LblFM2(0) = pa(7)
   End Select
End Sub

Private Sub Text10_GotFocus(Index As Integer)
   InverseTextBox Text10(Index)
End Sub

Private Sub Text10_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   Select Case Index
      'Modify By Cheng 2002/07/23
'      Case 0, 1
'         If KeyAscii <> 89 And KeyAscii <> 8 Then
'            KeyAscii = 0
'            Beep
'         End If
      Case 2
         If KeyAscii <> 89 And KeyAscii <> 8 Then
            KeyAscii = 0
            Beep
         ElseIf KeyAscii = 89 Then
            If MsgBox("¬O§_½T©w³¬¨÷ ?", vbQuestion + vbYesNo) = vbNo Then KeyAscii = 0
         End If
   End Select
End Sub

Private Sub Text12_GotFocus()
   InverseTextBox Text12
End Sub

Private Sub Text19_GotFocus()
   InverseTextBox Text19
End Sub

Private Sub Text19_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text19_Validate(Cancel As Boolean)
   If Text19 <> "" Then
      If ChgType(5) = False Then
         Cancel = True
         TextInverse Text19
      End If
   End If
End Sub

Private Sub Text20_GotFocus()
   InverseTextBox Text20
End Sub

Private Sub Text21_GotFocus()
   InverseTextBox Text21
End Sub

Private Sub Text22_GotFocus()
   InverseTextBox Text22
End Sub

Private Sub Text33_GotFocus(Index As Integer)
   InverseTextBox Text33(Index)
End Sub

Private Sub Text33_Validate(Index As Integer, Cancel As Boolean)
   'Added by Lydia 2017/06/14 ³]Äæ¦ìªø«×
    Dim iLen As Integer
    Select Case Index
    Case 0, 3 '±M§Q-Ápµ¸¤H¤¤¤å
         iLen = 30
    Case 1, 4 'Ápµ¸¤H­^¤å
         iLen = 35
    Case 2, 5, 6 'Ápµ¸¤H¤é¤å
         iLen = 60
    Case Else
         iLen = Text33(Index).MaxLength
    End Select
    'end 2017/06/14
    
   'Modified by Lydia 2017/06/14
   'If Not CheckLengthIsOK(Text33(Index), Text33(Index).MaxLength) Then
   If Not CheckLengthIsOK(Text33(Index), iLen) Then
      Cancel = True
   End If
End Sub

Private Sub Text6_GotFocus()
   InverseTextBox Text6
End Sub

Private Sub Text6_Validate(Cancel As Boolean)
   If Text6 = "" Then
      '2015/3/6 modify by sonia ¦]¶}©ñ124¦^´_Àu¥ýÅv¥D±i(FCP-051344)¦ý¤£¥²¿é®×¥ó¥Ø«e­ã»é
      'If Left(strKind, 1) = "1" Or Left(strKind, 1) = "3" Then
      If (Left(strKind, 1) = "1" Or Left(strKind, 1) = "3") And strKind <> "124" Then
         MsgBox "·s¥Ó½Ð®×©Î¦A¼f©Î§ï½Ðµ{§Ç®É¤£¥iªÅ¥Õ !", vbCritical
         Cancel = True
      End If
   Else
      If ChkDate(Text6) Then
         If Val(Text6) > Val(strSrvDate(2)) Then
            MsgBox "¥Ó½Ð®×®Ö­ã¤é¤£¥i¤j©ó¨t²Î¤é !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel = True Then TextInverse Text6
End Sub

Private Sub Text7_GotFocus()
'   InverseTextBox Text7
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'Text7.IMEMode = 1
   OpenIme
Dim intPos As Integer
'Modify By Cheng 2002/04/22
'±N´å¼Ð³]©w¦b¾÷Ãö¤å¸¹Äæªº"±M"ªº«á­±
With Me.Text7
   If Len("" & .Text) > 0 Then
      intPos = InStr("" & .Text, "±M")
      If intPos > 0 Then
         .SelStart = intPos
         .SelLength = 0
      End If
   End If
End With
End Sub

Private Sub Text7_LostFocus()
   'edit by nickc 2007/07/11 ¤Á´«¿é¤Jªk§ï¥ÎAPI
   'Text7.IMEMode = 1
   CloseIme
End Sub
'Add by Morgan 2011/1/5
Private Sub Text7_Validate(Cancel As Boolean)
   If CheckLengthIsOK(Text7, Text7.MaxLength) = False Then
      Cancel = True
   End If
End Sub

Private Sub Text9_GotFocus(Index As Integer)
   InverseTextBox Text9(Index)
End Sub

Private Sub Text9_Validate(Index As Integer, Cancel As Boolean)
   If Index = 2 Then
      If Text9(0) = "" And Text9(1) = "" And Text9(2) = "" Then
         MsgBox "®×¥ó¦WºÙ¤£¥i¦P®ÉªÅ¥Õ !", vbCritical
         Cancel = True
      End If
   End If
End Sub

'Add By Cheng 2002/05/22
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean
   
   TxtValidate = False
   
   'Added by Morgan 2012/12/13
   If Text16.Visible = True Then
      If Me.Text16.Enabled = True Then
         If Text16 = "" Then
            MsgBox "½Ð¿é¤J©Ó¿ì¤H¡I"
            Text16.SetFocus
            Exit Function
         Else
            Cancel = False
            Text16_Validate Cancel
            If Cancel = True Then
               Exit Function
            End If
         End If
      End If
   End If
   'end 2012/12/13
   
   If Me.Text19.Enabled = True Then
      Cancel = False
      Text19_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   If Me.Text6.Enabled = True Then
      Cancel = False
      Text6_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   
   For Each objTxt In Text9
      If objTxt.Enabled = True Then
         Cancel = False
         Text9_Validate objTxt.Index, Cancel
         If Cancel = True Then
            Exit Function
         End If
      End If
   Next
   
   'Add by Morgan 2004/9/2 ±qformsave ²¾¨Ó
   '¦¬¤å¤é¡Ö¡×93/7/1®Ö­ã®×¥ó¤w¦¬¤å¤é±¾¤T­Ó¤ëªº»âÃÒ´Á­­
   stNP07 = "": stNP08 = "": stNP09 = ""
   If pa(9) = ¥xÆW°ê®a¥N¸¹ And Val(Label3(3)) >= 930701 Then
      'Modified by Morgan 2012/12/24 +­l¥Í³]­p125,§ï½Ð­l¥Í³]­p308
      If InStr("101,102,103,104,105,107,125,301,302,303,304,305,306,307,308", m_CP10) > 0 Then
         stNP09 = Format(Val(Label3(3)) + 19110000)
         
         'Modify by Morgan 2004/9/9 ³£±¾¤T­Ó¤ëªº»âÃÒ´Á­­--ÀRªÚ
'         If Mid(pa(11), 9, 1) <> "" Then
'
'            stNP07 = ¥[µùÁp¦X '603
'            'ªk©w´Á­­=¦¬¤å¤é+30¤Ñ
'            stNP09 = CompDate(2, 30, stNP09)
'            '¥»©Ò´Á­­=ªk©w-2¤Ñ
'            stNP08 = CompDate(2, -2, stNP09)
'         Else
'            stNP07 = »âÃÒ¤ÎÃº¦~¶O '601
'            ªk©w´Á­­=¦¬¤å¤é+3­Ó¤ë
'            stNP09 = CompDate(1, 3, stNP09)
'            ¥»©Ò´Á­­=ªk©w-4¤Ñ
'            stNP08 = CompDate(2, -4, stNP09)
'         End If
         'Modify by Morgan 2010/12/28 ¥Ó½Ð®×¸¹§ï½X¼Æ
         'If Mid(pa(11), 9, 1) = "U" Then
         If Mid(pa(11), 10, 1) = "U" Then
            stNP07 = ¥[µùÁp¦X '603
         'ElseIf Mid(pa(11), 9, 1) = "A" Then
         ElseIf Mid(pa(11), 10, 1) = "A" Then
            stNP07 = ¥[µù°l¥[ '602
         Else
            stNP07 = »âÃÒ¤ÎÃº¦~¶O '601
         End If

            'ªk©w´Á­­=¦¬¤å¤é+3­Ó¤ë
            stNP09 = CompDate(1, 3, stNP09)
            'Modified by Morgan 2014/11/20 ¥~±M§ï¦^ÂÂ³W«h
            ''Added by Morgan 2014/10/9
            'If pa(9) = ¥xÆW°ê®a¥N¸¹ And strSrvDate(1) >= ¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
            '   stNP08 = PUB_GetOurDeadline(stNP09)
            'Else
            ''end 2014/10/19
            
            'Added by Morgan 2019/7/11 ¥~±M¥xÆW®×©Ò­­¥H§ï¤u§@¤Ñ­pºâ
            If strSrvDate(1) >= ¥~±M¥xÆW®×©Ò­­·s³W«h±Ò¥Î¤é Then
               'Modify By Sindy 2021/4/26 + m_pAgreeOnDate
               stNP08 = PUB_GetFCPOurDeadline(stNP09, 4, , m_pAgreeOnDate)
            Else
            'end 2019/7/11
      
               '¥»©Ò´Á­­=ªk©w-4¤Ñ
               stNP08 = CompDate(2, -4, stNP09)
               
            End If 'Added by Morgan 2019/7/11
            'End If 'Added by Morgan 2014/10/9
            'end 2014/11/20
            
         '2004/9/9 end
         'edit by nickc 2007/02/05 ¤£¥Î dll ¤F
         'If objLawDll.ChkMRec(TransDate(Label3(3).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
         If ClsLawChkMRec(TransDate(Label3(3).Caption, 2), pa(1) & pa(2) & pa(3) & pa(4), strExc(1), strExc(2)) Then
            If stNP08 <> strExc(1) Then
               If MsgBox("»PÂd¥x¤§¨Ó¨ç¦¬¤å°O¿ý¥»©Ò´Á­­ ( " & TransDate(strExc(1), 1) & ") ¤£²Å¡A½Ð½T»{ !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            ElseIf stNP09 <> strExc(2) Then
               If MsgBox("»PÂd¥x¤§¨Ó¨ç¦¬¤å°O¿ýªk©w´Á­­ ( " & TransDate(strExc(2), 1) & ") ¤£²Å¡A½Ð½T»{ !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                  Exit Function
               End If
            End If
            
         'Added by Morgan 2017/5/10 ¹q¤l¤½¤å
         ElseIf m_DocNo <> "" Then
            If m_DeadLine <> "" Then
               If Len(m_DeadLine) >= 7 Then
                  strExc(2) = m_DeadLine
               ElseIf Right(m_DeadLine, 1) = "¤é" Then
                  strExc(2) = CompDate(2, Val(m_DeadLine), Label3(3))
               ElseIf Right(m_DeadLine, 1) = "¤ë" Then
                  strExc(2) = CompDate(1, Val(m_DeadLine), Label3(3))
               End If
               If stNP09 <> strExc(2) Then
                  If MsgBox("»P¹q¤l¤½¤å¤§ªk©w´Á­­ ( " & TransDate(strExc(2), 1) & ") ¤£²Å¡A½Ð½T»{ !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
                     Exit Function
                  End If
               End If
            End If
         'end 2017/5/10
         Else
            If MsgBox("¨Ó¨ç°O¿ýÀÉµL¦¹°O¿ý¡A½Ð½T»{ !", vbCritical + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
      End If
   End If
   
   'Added by Morgan 2016/2/3
   If Left(pa(75), 8) = "Y4829203" Then
      MsgBox "¦Ü HP ¥­¥x¿é¤J¬ÛÃö¸ê®Æ!!", vbExclamation
   End If
   'end 2016/2/3
   
   'Added by Lydia 2019/05/23 °É»~¤½³ø±±ºÞ
   If Frame1.Visible = True Then
       'Modified by Lydia 2023/08/25 ±Æ°£±M§QÅv©µªø415
       If m_CP10 <> "415" And Opt1(0).Value = False And Opt1(1).Value = False And Opt1(2).Value = False Then
           MsgBox "½Ð¤Ä¿ï°É»~Ãþ«¬¡I", vbCritical
           Exit Function
       End If

       '¤Ä¿ï§ó¥¿402®É¥i¥H¤£¿é¤J¤é´Á¤Î´Á§O
       'Modified by Lydia 2023/08/25 ±Æ°£±M§QÅv©µªø415
       If Opt1(1).Value = False And m_CP10 <> "415" Then
           For Each objTxt In txtCRC
               If Trim(objTxt) = "" Then
                    MsgBox IIf(objTxt.Index = 0, "°É»~¤é´Á", "´Á§O") & "¤£¥iªÅ¥Õ¡I", vbCritical
                    objTxt.SetFocus
                    txtCRC_GotFocus objTxt.Index
                    Exit Function
               Else
                    Cancel = False
                    Call txtCRC_Validate(objTxt.Index, Cancel)
                    If Cancel = True Then
                        Exit Function
                    End If
               End If
           Next
       End If
   End If
   
   'Added by Morgan 2023/2/23
   'Modified by Lydia 2025/02/12 +245©µ½w¼f¬d
   If m_CP10 = "415" Or m_CP10 = "245" Then
      If txt415Date = "" Then
         'Added by Lydia 2025/02/12
         If m_CP10 = "245" Then
             MsgBox "½Ð¿é¤J©µ½w¼f¬d¤é´Á¡I", vbCritical
         Else
         'end 2025/02/12
             MsgBox "½Ð¿é¤J±M§QÅv´Á¶¡©µªø«á¤é´Á¡I", vbCritical
         End If
         txt415Date.SetFocus
         Exit Function
      Else
         Cancel = False
         Call txt415Date_Validate(Cancel)
         If Cancel = True Then
            txt415Date_GotFocus
            Exit Function
         End If
      End If
   End If
   'end 2023/2/23
   
   'Add by Sindy 2021/4/27 ÀË¬dµe­±ªº TextBox, ComboBox ¬O§_§t¦³Unicode¤å¦r
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/4/27 END
   
   TxtValidate = True
End Function

'Add By Cheng 2002/07/03
Private Function GetPromoterNO(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As String
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String
Dim strMaxCP09 As String
'92.1.19 modify by sonia ¶È¥Ó½Ð®×¸¹201,209,210¤§®Ö½Z¤H, µL®Ö½Z¤H§ì©Ó¿ì¤H,¨ä¥L®×¥ó©Ê½è§ì­ì©Ó¿ì¤H
GetPromoterNO = m_CP14
If m_CP10 = "101" Or m_CP10 = "102" Or m_CP10 = "103" Or m_CP10 = "104" Or m_CP10 = "105" Then
   strMaxCP09 = ""
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   StrSQLa = "Select CP09,CP14 From CaseProgress Where CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "' AND (CP10='201' OR CP10='209' OR CP10='210' ) ORDER BY CP09 DESC"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      rsA.MoveFirst
      strMaxCP09 = "" & rsA.Fields(0).Value
      GetPromoterNO = "" & rsA.Fields(1).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
   If strMaxCP09 <> "" Then
      StrSQLa = "SELECT EP04 FROM ENGINEERPROGRESS WHERE EP02='" & strMaxCP09 & "'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         If Not IsNull(rsA.Fields(0).Value) Then GetPromoterNO = "" & rsA.Fields(0).Value
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
End If
End Function

Public Sub PrintFlowSheet(pRecNo As String, pCRecNo As String)
   Dim iPrint As Long, xi As Long, yi As Long, Xa As Long, xb As Long, yb As Long
   Dim stSQL As String, iR As Integer, stVTB As String, strTmp As String
   Dim adoRst As ADODB.Recordset
   Dim iSel As Integer
   Dim iCopy As Integer
   Dim stSFee As String
   Dim bClear As Boolean 'Added by Morgan 2015/11/24
   
   Const Xo As Integer = 1500
   Const Yo As Integer = 1200
   Const LH As Integer = 300
   Const LW As Long = 10300
   Const LD As Integer = 150
   
   'Added by Morgan 2021/6/2
   Dim strPdfPath As String, strPdfName As String
   Dim oFileSys As New FileSystemObject
   Dim oFile
   'end 2021/6/2
   
   stVTB = "select a1l01,sum(a1l05) a1l05 from caseprogress c1,caseprogress c2,acc1l0 where c1.cp09='" & pRecNo & "' and c2.cp09(+)=c1.cp43 and a1l01(+)=c2.cp60 and a1l04 in ('41699','10799') group by a1l01"
   
   'C1:¥»©Ò®×¸¹,C2:°h´Ú¤H¦WºÙ,C3:·sD/N,C4:¬O§_¦P·N¦©ªA°È¶O,C5:¬ÛÃö¸¹®×¥ó©Ê½è,C6:­ìD/N
   'C7:¬üª÷,C8:¶×²v,C9:¥x¹ô,C10:¬O§_µ²²M,C11:°h¶Oª÷ÃB,C12:­ìD/N³W¶O,C13:§éÅý¬üª÷,C14:°h¶OªA°È¶O
   'Modified by Morgan 2013/6/28 ¦Ò¼{¦A¼f©µ´Á°h¶O(¨S¦³¦¬¤å)
   'Modified by Morgan 2022/10/12 +435Äò¦æ¥À®×¦A¼f
   stSQL = "select c1.cp01||'-'||c1.cp02||decode(c1.cp03||c1.cp04,'000','','-'||c1.cp03||'-'||c1.cp04) C1" & _
      ",c1.cp49 C2,c1.cp60 C3,c1.cp86 C4,decode(c2.cp10,'404',np07,c2.cp10) C5,C2.CP60 C6,k2.a1k08 C7,k2.a1k10 C8,k2.a1k11 C9,k2.a1k29 C10" & _
      ",c1.cp19 C11,a1l05 C12,k2.a1k06 C13,k1.a1k11 C14,st02,C2.CP27 as ExDate" & _
      " from  caseprogress c3,caseprogress c1,caseprogress c2,nextprogress,acc1k0 k1,acc1k0 k2,(" & stVTB & ") V1,staff" & _
      " where c3.cp09='" & pCRecNo & "' and c1.cp09(+)=c3.cp43 and c2.cp09(+)=c1.cp43 and k1.a1k01(+)=c1.cp60" & _
      " and k2.a1k01(+)=c2.cp60 and a1l01(+)=k2.a1k01" & _
      " and st01(+)=nvl(c3.cp65,'" & strUserNum & "') and c2.cp10 in ('416','107','407','605','404','435') and np01(+)=c2.cp43"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      With adoRst
      'Modified by Morgan 2015/11/24 ­Y°h¶Oªº¬ÛÃö¦¬¤å¸¹¬°°²¦¬¤å®É(«D¥»©Ò¿ì²z)µø¬°µ²²M
      'If .Fields("C10") = "Y" Then
      If .Fields("C10") = "Y" Or ("" & .Fields("ExDate") = "19221111") Then
      'end 2015/11/24
         bClear = True
      Else
         bClear = False
      End If
      
      stSFee = Format("" & .Fields("C14"), DDollar)
      iSel = 0
      Select Case "" & .Fields("C5")
         'Modified by Morgan 2022/10/12 +435Äò¦æ¥À®×¦A¼f
         Case "416", "107", "435"
            '¤wµ²²M(¤w¦¬´Ú)
            If bClear Then
               '¦³¬Û¤Ï«ü¥Ü
               If .Fields("C4") = "N" Then
                  iSel = 2
               Else
                  iSel = 1
               End If
            
            'Added by Morgan 2015/4/24
            ElseIf m_bNoDN = True Then
               iSel = 4
            'end 2015/4/24
            
            '¥¼µ²²M(¥¼¦¬´Ú)
            Else
               iSel = 3
            End If
         Case "407"
            If bClear Then
               iSel = 41
            Else
               iSel = 42
            End If
         Case "605"
            If bClear Then
               iSel = 51
            Else
               iSel = 52
            End If
      End Select
      
'Added by Morgan 2021/6/2
      strPdfPath = App.path & "\" & strUserNum
      strPdfName = PUB_CaseNo2FileName(Text2, Text3, Text4, Text5) & ".1001." & Format(Now, "yyyymmddhhmmss") & ".INCOM.PDF"
      Load frmPDF
      frmPDF.Show
      frmPDF.StartProcess strPdfPath, strPdfName
      
'Removed by Morgan 2021/6/2
'RePrint:
'   For iCopy = 1 To 2
'      If iCopy > 1 Then Printer.NewPage
'end 2021/6/2

      Printer.PaperSize = 9 'A4
      Printer.Orientation = 1 'ª½¦L
      'Printer.Copies = 2
      Printer.Font.Name = "²Ó©úÅé"
      Printer.Font.Size = 16
      Printer.Font.Bold = True
      Printer.Font.Underline = False
      yi = Yo
      xi = Xo
      Printer.CurrentY = yi
      Printer.CurrentX = xi
      strExc(0) = "°h¶O®Ö­ã¬yµ{ªí(¦C¦L2±i:¤@¥æ°]°È¤£ÀH¨÷;¤@¦s¨÷¸mµ{§Ç³B)"
      Printer.Print strExc(0)
      Printer.DrawWidth = 5
      yi = Printer.CurrentY + 50
      'Printer.Line (Xi, Yi)-(Xi + Printer.TextWidth(strExc(0)), Yi)
      
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      
      yi = Printer.CurrentY + LH
      xi = Xo
      Printer.CurrentY = yi
      Printer.CurrentX = xi
      Printer.Print "®×¸¹:¡@¡@¡@¡@¡@¡@¡@¡@¡@µ{§Ç¤H­û:"
      
      strExc(0) = "¦C¦L¤é´Á¡G" & ChangeTStringToTDateString(strSrvDate(2))
      Printer.CurrentY = yi
      Printer.CurrentX = LW - Printer.TextWidth(strExc(0))
      Printer.Print strExc(0)
      
      Xa = xi + Printer.TextWidth("®×¸¹:")
      Printer.CurrentX = Xa + 100
      Printer.CurrentY = yi
      Printer.Font.Bold = True
      Printer.Print .Fields("C1")
      
      Printer.Font.Bold = False
      xb = Xa + Printer.TextWidth("¡@¡@¡@¡@¡@¡@¡@¡@¡@µ{§Ç¤H­û:")
      Printer.CurrentX = xb + 100
      Printer.CurrentY = yi
      Printer.Font.Bold = True
      Printer.Print .Fields("st02")
      
      yi = Printer.CurrentY + 50
      'Printer.Line (Xa, Yi)-(Xa + Printer.TextWidth("¡@¡@¡@¡@¡@¡@¡@¡@"), Yi)
      'Printer.Line (Xb, Yi)-(Xb + Printer.TextWidth("¡@¡@¡@¡@"), Yi)
      
      yi = Printer.CurrentY + LH
      Printer.CurrentX = xi: Printer.CurrentY = yi
      Printer.Font.Bold = False
      Printer.Print "¯S©w°h´Ú¤H¦WºÙ(¥~¤å):"
      
      Xa = xi + Printer.TextWidth("¯S©w°h´Ú¤H¦WºÙ(¥~¤å):")
      Printer.CurrentY = yi
      Printer.CurrentX = Xa + 100
      Printer.Font.Bold = True
      Printer.Print "" & .Fields("C2")
      Printer.Font.Bold = False
      
      yi = Printer.CurrentY + 50
      'Printer.Line (Xa, Yi)-(LW, Yi)
      
      yi = Printer.CurrentY + LH
      Xa = Printer.TextWidth("¤@. ")
      
      strExc(0) = "¤@. ¦¬¨ì¥N²z¤H¦^ÂÐ(¦³¬Û¤Ï«ü¥Ü) ¤£§Æ±æ¥»©Òª½±µ¦©©è¥»©Ò°h³W¶O©Ò²£¥Í¤§ªA°È¶ONT$2500«h©Ó¿ì¶·¼gÁpµ¸³æ¥æµ{§Ç¹q¸£¬ö¿ý, ¥N²z¤H¥¼¦^ÂÐªÌ(µL¬Û¤Ï«ü¥Ü), «h±qIPO©Ò°h³W¶Oª½±µ¦©©è¥»©Ò¤§ªA°È¶ONT$2500"
      Printer.CurrentY = yi
      Printer.CurrentX = xi
      SmartPrint strExc(0), Xa, LW, LD
      
      strExc(0) = "¤G. ¦¬¨ìIPO¹ê¼f©Î¦A¼f°h¶O®Ö­ã ¦¹±i³øªí¦Û°Ê¤Ä¿ï(1)¡X(3)"
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      SmartPrint strExc(0), Xa, LW, LD
      
      strExc(0) = "¤T. °h¶Oµo¤å«e½Ðµ{§Ç½T»{¸Ó®×³W¶O¬O§_¤À¦¸Ãº¯Ç, ­Y¬O,µo¤å®É³W¶Oª÷ÃB¶·¥[Á`¥Ó½Ð°h¶O,¨Ã³Æµù³W¶OÁ`ª÷ÃB¥B¤é«á¦]µLªk¥H©w½Z(1)¡X(3)³B²z,½Ð¥æ©Ó¿ì"
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      SmartPrint strExc(0), Xa, LW, LD
      
      
      Xa = Printer.TextWidth("¡¼(1")
      yi = Printer.CurrentY + LD
      Printer.CurrentY = yi
      Printer.CurrentX = xi
      
      If iSel = 1 Then
         yb = yi - 5
         
         strTmp = "¡½(1)µL¬Û¤Ï«ü¥Ü,µL¤í´Ú(¦©ªA°È¶O)+½Ð´Ú³æ¸¹½X(NT$"
         Printer.Print strTmp
         xb = xi + Printer.TextWidth(strTmp)
         
         strTmp = stSFee
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         strTmp = " D/N No."
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = "" & .Fields("C3")
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print ")"
         'Modified by Morgan 2014/3/26
         'strExc(0) = "¡÷°h°]°È³B¨R±b¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(1)(¤£¦LD/N)+±H¤ä²¼(¶·¦©´îNT$" & stSFee & ")"
         strExc(0) = "¡÷°h°]°È³B¨R±b¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(1)(¤£¦LD/N)+(A)±H¤ä²¼©Î(B)±HC/N(¶·¦©´îNT$" & stSFee & ")"
      Else
         Printer.Print "¡¼(1)µL¬Û¤Ï«ü¥Ü,µL¤í´Ú(¦©ªA°È¶O)+½Ð´Ú³æ¸¹½X(NT$2500 D/N No.__________)"
         'Modified by Morgan 2014/3/26
         'strExc(0) = "¡÷°h°]°È³B¨R±b¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(1)(¤£¦LD/N)+±H¤ä²¼(¶·¦©´îNT$2500)"
         strExc(0) = "¡÷°h°]°È³B¨R±b¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(1)(¤£¦LD/N)+(A)±H¤ä²¼©Î(B)±HC/N(¶·¦©´îNT$2500)"
      End If
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      SmartPrint strExc(0), 0, LW, LD
      
      yi = Printer.CurrentY + LD
      Printer.CurrentY = yi
      Printer.CurrentX = xi
      If iSel = 2 Then
         yb = yi - 5
         
         strTmp = "¡½(2)¦³¬Û¤Ï«ü¥Ü,µL¤í´Ú(°h¥þÃB)+½Ð´Ú³æ¸¹½X(NT$"
         Printer.Print strTmp
         xb = xi + Printer.TextWidth(strTmp)
         
         strTmp = stSFee
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         strTmp = " ¦C¦LD/N No."
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = "" & .Fields("C3")
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print ")"
      Else
         Printer.Print "¡¼(2)¦³¬Û¤Ï«ü¥Ü,µL¤í´Ú(°h¥þÃB)+½Ð´Ú³æ¸¹½X(NT$2500 ¦C¦LD/N No.__________)"
      End If
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      'Modified by Morgan 2014/3/26
      'Printer.Print "¡÷°h°]°È³B°h¶O¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(2)¤ÎD/N(908)+±H¤ä²¼(¥þÃB³W¶O)"
      Printer.Print "¡÷°h°]°È³B°h¶O¤Î¶}²¼¡÷µ{§Ç¦C¦L©w½Z(2)¤ÎD/N(908)+(A)±H¤ä²¼©Î(B)±HC/N(¥þÃB³W¶O)"
      
      If iSel = 3 Then
         strTmp = "¡½(3)¤£½×¦³µL¬Û¤Ï«ü¥Ü, ¸Ó®×¹ê¼f©Î¦A¼fµ{§Ç©|¥¼¦¬´Ú(³W¶O¥¼¦¬)«h¤@«ß±H½Ð´Ú³æ¸¹½X(NT$" & stSFee & ") +Credit Note(³W¶Oª÷ÃB)"
      Else
         strTmp = "¡¼(3)¤£½×¦³µL¬Û¤Ï«ü¥Ü, ¸Ó®×¹ê¼f©Î¦A¼fµ{§Ç©|¥¼¦¬´Ú(³W¶O¥¼¦¬)«h¤@«ß±H½Ð´Ú³æ¸¹½X(NT$2500) +Credit Note(³W¶Oª÷ÃB)"
      End If
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      SmartPrint strTmp, Xa, LW, LD
      
      yi = Printer.CurrentY + LD
      Printer.CurrentY = yi
      Printer.CurrentX = xi + Xa
      If iSel = 3 Then
         yb = yi - 5
         strTmp = "¦¹³øªí¹q¸£±a¥X:¥¼¦¬´ÚD/N No."
         Printer.Print strTmp
         xb = xi + Xa + Printer.TextWidth(strTmp)
         
         strTmp = "" & .Fields("C6")
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         strTmp = " ª÷ÃBNT$"
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = Format("" & .Fields("C9"), DDollar)
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         strTmp = " (US$"
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = Format("" & .Fields("C7"), DDollar)
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print ")"
      Else
         Printer.Print "¦¹³øªí¹q¸£±a¥X:¥¼¦¬´ÚD/N No.___________ª÷ÃBNT$________(US$______)"
      End If
      
      yi = Printer.CurrentY + LD
      Printer.CurrentY = yi
      Printer.CurrentX = xi + Xa
      If iSel = 3 Then
         yb = yi - 5
         xb = xi + Xa
         strTmp = "+·s°h¶O½Ð´Ú³æ¸¹½X(908) (NT$"
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = stSFee
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         strTmp = " ¦C¦LD/N No."
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         strTmp = "" & .Fields("C3")
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
                  
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print ")"
      Else
         Printer.Print "+·s°h¶O½Ð´Ú³æ¸¹½X(908) (NT$2500 ¦C¦LD/N No.__________)"
      End If
      
      yi = Printer.CurrentY + LD
      Printer.CurrentY = yi
      Printer.CurrentX = xi + Xa
      If iSel = 3 Then
         yb = yi - 5
         xb = xi + Xa
         
         'Modified by Morgan 2018/3/16 --David
         'strTmp = "+¹q¸£¦C¦L Credit Note No."
         strTmp = "+©Ó¿ì»s§@ Credit Note No."
         'end 2018/3/16
         
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         'Modified by Morgan 2018/7/31 ³æ¸¹§ï¥Ñ°]°È³B´£¨Ñ(¥N²z¤H­n¨D¤£¥i»P½Ð´Ú³æ¸¹¬Û¦P)--Lina Ex:FCP-052224
         'strTmp = "" & .Fields("C6")
         strTmp = "__________"
         'end 2018/7/31
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
                  
         strTmp = " ª÷ÃBNT$"
         Printer.CurrentY = yi
         Printer.CurrentX = xb
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         
         'Modified by Morgan 2017/8/4 §éÅýª÷ÃBÀ³¸Ó­n¦L°h¶Oª÷ÃB
         'strTmp = Format("" & .Fields("C12"), DDollar)
         strTmp = Format("" & .Fields("C11"), DDollar)
         'end 2017/8/4
         Printer.CurrentY = yb
         Printer.CurrentX = xb
         Printer.FontBold = True
         Printer.Print strTmp
         xb = xb + Printer.TextWidth(strTmp)
         Printer.FontBold = False
         
         'Modified by Morgan 2018/3/16 §ï©Ó¿ì»s§@¦¹®É©|¥¼¿é¤J¤£·|¦³­È(¥B§éÅý¤]§ï¬°¥x¹ôª÷ÃB)
         'strTmp = " (US$"
         'Printer.CurrentY = yi
         'Printer.CurrentX = xb
         'Printer.Print strTmp
         'xb = xb + Printer.TextWidth(strTmp)
         
         'strTmp = Format("" & .Fields("C13"), DDollar)
         'Printer.CurrentY = yb
         'Printer.CurrentX = xb
         'Printer.FontBold = True
         'Printer.Print strTmp
         'xb = xb + Printer.TextWidth(strTmp)
         'Printer.FontBold = False
         
         'Printer.CurrentY = yi
         'Printer.CurrentX = xb
         'Printer.Print ")"
         'end 2018/3/16
         
      Else
         'Modified by Morgan 2018/3/16 --David
         'Printer.Print "+¹q¸£¦C¦L Credit Note No.___________ª÷ÃBNT$________(US$______)"
         Printer.Print "+©Ó¿ì»s§@ Credit Note No.___________ª÷ÃBNT$________(US$______)"
         
      End If
      
      'Added by Morgan 2018/3/16 --±Ó²ú
'Removed by Morgan 2019/4/9 ¤£¥²¿é§éÅý--°û²ñ
'      strTmp = "¡÷°hµ{§Ç(Key±b³æªº¤H)¿é¤J§éÅýª÷ÃB"
'      yi = Printer.CurrentY + LD
'      Printer.CurrentY = yi
'      Printer.CurrentX = xi + Xa
'      'Modified by Morgan 2018/7/31
'      'Printer.Print strTmp
'      If iSel = 3 And "" & .Fields("C6") <> "" Then
'         strTmp = strTmp & "(D/N No."
'         Printer.Print strTmp
'         xb = xi + Xa + Printer.TextWidth(strTmp)
'
'         strTmp = "" & .Fields("C6")
'         Printer.CurrentY = yi
'         Printer.CurrentX = xb
'         Printer.FontBold = True
'         Printer.Print strTmp
'
'         Printer.CurrentY = yi
'         Printer.CurrentX = xb + Printer.TextWidth(strTmp)
'         Printer.FontBold = False
'         Printer.Print ")"
'      Else
'         Printer.Print strTmp
'      End If
'end 2019/4/9
      'end 2018/7/31
      'end 2018/3/16
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      Printer.Print "¡÷°h°]°È³B¨R³W¶Oª÷ÃB¡÷µ{§Ç¦C¦L©w½Z(3)+ D/N(908)+C/N"
      
      'Added by Morgan 2015/4/24
      If iSel = 4 Then
         strTmp = "¡½(4)·|°]°È³B¨R³W¶O(°h¶O¤£½Ð´Ú)"
      Else
         strTmp = "¡¼(4)·|°]°È³B¨R³W¶O(°h¶O¤£½Ð´Ú)"
      End If
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      SmartPrint strTmp, Xa, LW, LD
      'end 2015/4/24
      
      Xa = Printer.TextWidth("¤T. ")
      Printer.CurrentY = Printer.CurrentY + LH
      Printer.CurrentX = xi + Xa
      Printer.Print "¡°±H°ê¥~¥N²z¤H/«È¤á (1)¡X(3) ´Ú³qª¾¨ç¤º®e¤Îªþ¥ó¬Ò¤£¦P"
      
     
      xb = Printer.TextWidth("(1) ")
      'Modified by Morgan 2014/3/26
      'strExc(0) = "(1) µ{§Ç¤H­û¦C¦L½Ð´Ú³qª¾¨ç(¿é¤J¹ê¼f/¦A¼f³W¶O¡B¤ä²¼¸¹½X¤Î¬üª÷ª÷ÃB)+±H¤ä²¼(¶·¦©´îNT$2500)"
      strExc(0) = "(1) µ{§Ç¤H­û¦C¦L½Ð´Ú³qª¾¨ç(¿é¤J¹ê¼f/¦A¼f³W¶O¡B¤ä²¼¸¹½X¤Î¬üª÷ª÷ÃB)+(A)±H¤ä²¼©Î(B)±HC/N(¶·¦©´îNT$2500)"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      SmartPrint strExc(0), xb, LW, LD
      
      'Modified by Morgan 2014/3/26
      'strExc(0) = "(2) µ{§Ç¤H­û¦C¦L½Ð´Ú³qª¾¨ç(¿é¤J¹ê¼f/¦A¼f³W¶O¡B¤ä²¼¸¹½X¤Î¬üª÷ª÷ÃB)+¦C¦LD/N+±H¤ä²¼(¥þÃB³W¶O)"
      strExc(0) = "(2) µ{§Ç¤H­û¦C¦L½Ð´Ú³qª¾¨ç(¿é¤J¹ê¼f/¦A¼f³W¶O¡B¤ä²¼¸¹½X¤Î¬üª÷ª÷ÃB)+¦C¦LD/N+(A)±H¤ä²¼©Î(B)±HC/N(¥þÃB³W¶O)"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      SmartPrint strExc(0), xb, LW, LD
      
      strExc(0) = "(3) µ{§Ç¤H­û¦C¦L½Ð´Ú³qª¾¨ç(¿é¤J¹ê¼f/¦A¼f³W¶O)+¦C¦LD/N+¦C¦LC/N(¹q¸£²£¥Í)"
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + Xa
      SmartPrint strExc(0), xb, LW, LD
            
      Printer.CurrentY = Printer.CurrentY + 100
      Printer.CurrentX = xi
      Printer.Print "============================================================================"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      Printer.Print "¨ä¥L°h¶O: ¦P¥H©¹§@·~"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      If iSel \ 10 = 4 Then
         If iSel = 41 Then
            Printer.Print "¡½(4)½Ð¨D­±¸ß°h¶O  ¡½¤w¥I´Ú ¡¼¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
         Else
            Printer.Print "¡½(4)½Ð¨D­±¸ß°h¶O  ¡¼¤w¥I´Ú ¡½¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
         End If
      Else
         Printer.Print "¡¼(4)½Ð¨D­±¸ß°h¶O  ¡¼¤w¥I´Ú ¡¼¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
      End If
      'Modified by Morgan 2018/3/16 ­×¥¿¤å¥y--±Ó²ú
      xb = Printer.TextWidth("¡¼(4)")
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + xb
      Printer.Print "(A) ¸Óµ§±b³æ¤w¥I´Ú ½Ð©Ó¿ì¼g«H¨Ã¶}C/N ¥æ°]°È¬ö¿ýC/Nª÷ÃB"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + xb
      Printer.Print "(B) ¸Óµ§±b³æ¥¼¥I´Ú ½Ð©Ó¿ì¼g«H¥Bª½±µ¥H¸ÓD/N¶}C/N"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi
      If iSel \ 10 = 5 Then
         If iSel = 51 Then
            Printer.Print "¡½(5)¦~¶O°h¶O      ¡½¤w¥I´Ú ¡¼¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
         Else
            Printer.Print "¡½(5)¦~¶O°h¶O      ¡¼¤w¥I´Ú ¡½¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
         End If
      Else
         Printer.Print "¡¼(5)¦~¶O°h¶O      ¡¼¤w¥I´Ú ¡¼¥¼¥I´Ú(¹q¸£¦Û°Ê¤Ä¿ï)"
      End If
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + xb
      Printer.Print "(A) ¸Óµ§±b³æ¤w¥I´Ú ½Ð©Ó¿ì¼g«H¨Ã¶}C/N ¥æ°]°È¬ö¿ýC/Nª÷ÃB"
      
      Printer.CurrentY = Printer.CurrentY + LD
      Printer.CurrentX = xi + xb
      Printer.Print "(B) ¸Óµ§±b³æ¥¼¥I´Ú ½Ð©Ó¿ì¼g«H¥Bª½±µ¥H¸ÓD/N¶}C/N"
      'end 2018/3/16
      
'Removed by Morgan 2021/6/2
'   Next
   
      Printer.EndDoc
      
'Added by Morgan 2021/6/2
      frmPDF.EndtProcess
      Unload frmPDF
      If Dir(strPdfPath & "\" & strPdfName) <> "" Then
         Set oFile = oFileSys.GetFile(strPdfPath & "\" & strPdfName)
         SaveAttFile_PDF pCRecNo, strPdfPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
      End If
      
      
RePrint:
      PUB_PrintPDF strPdfPath & "\" & strPdfName, , 2
'end 2021/6/2
      
      If MsgBox("°h¶O®Ö­ã¬yµ{ªí¦C¦L§¹²¦¡I¬O§_­n­«¦L¡H", vbYesNo + vbDefaultButton1) = vbYes Then
         GoTo RePrint
      End If
      
      End With
   End If
   Set adoRst = Nothing
End Sub

Private Sub SmartPrint(pStr As String, lPresv As Long, lMax As Long, iLSpace As Integer)
   Dim iPos As Integer, Xa As Long, xb As Long
   iPos = 1
   Xa = lMax
   xb = Printer.CurrentX
   Do
      If Printer.TextWidth(Left(pStr, iPos)) > (Xa - xb) Then
         Printer.Print Left(pStr, iPos - 1)
         pStr = Mid(pStr, iPos)
         iPos = 0
         Printer.CurrentY = Printer.CurrentY + iLSpace
         Printer.CurrentX = xb + lPresv
         Xa = lMax - lPresv
      End If
      If Printer.TextWidth(pStr) <= (Xa - xb) Then
         Printer.Print pStr
         Exit Do
      End If
      iPos = iPos + 1
   Loop
End Sub

Private Sub txt415Date_GotFocus()
   TextInverse txt415Date
End Sub

Private Sub txt415Date_Validate(Cancel As Boolean)
   If txt415Date <> "" Then
      Cancel = Not ChkDate(txt415Date)
      'Added by Lydia 2025/02/12 ©µ½w¼f¬d
      If m_CP10 = "245" Then
      Else
      'end 2025/02/12
         If DBDATE(txt415Date) <= DBDATE(pa(25)) Then
            MsgBox "©µªø«á±M¥Î´Á¥²¶·¤j©ó¥Ø«e±M¥Î´Á¡I", vbCritical
            Cancel = True
         End If
      End If 'Added by Lydia 2025/02/12
   End If
End Sub

Private Sub txtCP19_GotFocus()
   TextInverse txtCP19
   CloseIme
End Sub

Private Sub txtCP19_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text16_GotFocus()
   InverseTextBox Text16
End Sub

Private Sub Text16_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text16_Validate(Cancel As Boolean)
   Dim strTempName As String
   LblFM2(1) = ""
   If Text16 <> "" Then
      If ClsPDGetStaff(Text16.Text, strTempName) Then
         LblFM2(1) = strTempName
      Else
         Cancel = True
         TextInverse Text16
      End If
   End If
End Sub
'Added by Morgan 2012/12/13
'³]©wªì¼f®Ö­ã¤À³Î«ØÄ³±±¨î
Private Sub SetDivSug()
   m_PA162 = pa(162)
   m_bDivSugTextAlert = False
   m_EditDivSugText = "" 'Added by Morgan 2020/2/27
   m_bNewGrant = False 'Added by Morgan 2013/10/29
   m_bAgainGrant = False 'Added by Lydia 2019/07/30 µo©ú¦A¼f®Ö­ã
   m_bHasDivCase = False 'Added by Morgan 2019/10/7 ¬O§_¦³¤À³Î®×
   
   Label8.Visible = False: Text16.Visible = False: LblFM2(1).Visible = False
   
   'Modified by Morgan 2012/12/19
   '¤w³¬¨÷¤£¥²³qª¾¦A³qª¾ FCP-033631 -- ÀRªÚ
   'Modified by Morgan 2013/1/30 +ªì¼f´£¤À³Îªº®Ö­ã
   'Memo by Lydia 2015/07/17 ªì¼f®Ö­ãªº§PÂ_¦³ÅÜ§ó,½Ð¤@¨Ö­×§ïfrm075004_2.cmdPrintCForm_Click
   'Modified by Lydia 2019/07/30 ¦]108.11.1­×ªk¤À³ÎºÞ¨î´Á­­³]©w
    '1. ©ó108.8.1¦¬¨ì¤§®Ö­ã¨ç¡G
    '¡@1.1. µo©úªì¼f®Ö­ã¡Gºû«ù­ì¦³³]©w¤§¤À³Î´Á­­
    '¡@1.2. µo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã¡G­ì¦³³]©w¤À³Î´Á­­¤§«È¤á½s¸¹¡A¼W¥[±±ºÞ¦æ¨Æ¾ä´Á­­¡A­ì«h·Óªì¼f®Ö­ã¡A´Á­­¬°¦¬¨ì®Ö­ã¨ç«á¢²­Ó¤ë´Á­­¡A¨Ã±a³Æµù¦Ü³qª¾§i­ã¤§¶i«×³Æµù¡C
    '2. ©ó108.10.1¦¬¨ì¤§®Ö­ã¨ç¡Gµo©úªì¼f®Ö­ã¡Bµo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã¡G¬Ò³]©w¦¬¨ì®Ö­ã¨ç«á¢²­Ó¤ë´Á­­¡C
   'If frm06010602_2.Text6 = "1" And (strKind = "101" Or (strKind = "307" And pa(163) = "Y")) And pa(57) = "" Then
   '   m_bNewGrant = True 'Added by Morgan 2013/10/29
   strExc(1) = DBDATE(Label3(3))
   'Modified by Morgan 2024/6/25 +§PÂ_µo©ú·s«¬ Ex:FCP-070550 --±Ó²ú
   If frm06010602_2.Text6 = "1" And (pa(8) = "1" Or pa(8) = "2") Then
        'µo©úªì¼f®Ö­ã
        If (strKind = "101" Or (strKind = "307" And pa(163) = "Y")) And pa(57) = "" Then
           m_bNewGrant = True
        'µo©ú¦A¼f®Ö­ã¡B·s«¬®Ö­ã(©ó108.8.1¦¬¨ì)
        ElseIf strExc(1) >= "20190801" And (strKind = "102" Or (strKind = "107" And pa(8) = "1")) Then
           m_bAgainGrant = True
        End If
        
      'Added by Morgan 2019/10/7
      '§ïµo©ú/·s«¬ªº¥Ó½Ð¡B¦A¼f¡B§ï½Ð¡B¤À³Î®Ö­ã³£­n§PÂ_¬O§_¦³¤À³Î«ØÄ³
      If (pa(8) = "1" Or pa(8) = "2") Then
         If strKind = "101" Or strKind = "102" Or strKind = "107" Or strKind = "301" Or strKind = "302" Or strKind = "307" Then
            m_bNewGrant = True
         End If
      End If
      'end 2019/10/7
      
   End If
   If m_bNewGrant = True Then
   'end 2019/07/30
      If m_PA162 <> "N" Then
         If m_PA162 = "" Then
            'Modified by Morgan 2012/12/13 ¹w³]­n¿é¤À³Î«ØÄ³(¹ê¼fµo¤å«áµL¥Ó´_­×¥¿µo¤åªÌ°£¥~,­YµL¹ê¼fµo¤å(¤¤¶¡¨Ó©Ò)¤]¹w³]­n)
            'Modified by Morgan 2019/10/7 ¦A¼f­ã©Î¦³¥Ó´_¡B­×¥¿µo¤å©Î¹ê¼f/¦A¼fµo¤å«á¦³¥D°Ê­×¥¿µo¤åªÌ³]Y§_«hN
            'strExc(0) = "select 1 from caseprogress a WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " and cp10='416' and cp27>0" & _
               " and not exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10 in ('204','205') and b.cp27>a.cp27)"
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '   m_PA162 = "N"
            'Else
            '   m_PA162 = "Y"
            'End If
            If strKind = "107" Then
               m_PA162 = "Y"
            Else
               strExc(0) = "select 1 from caseprogress a WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & _
                  " and cp27>0 and (cp10 in ('204','205') or (cp10='203' and exists(select * from caseprogress b where b.cp01=a.cp01" & _
                  " and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10 in ('416','107') and b.cp27<a.cp27)))"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_PA162 = "Y"
               'Added by Morgan 2019/12/27
               '¥Ñ§O©ÒÂà¨Ó®×¤l¡A¦]µLªk½T»{¬O§_¦³´£­×¥¿¡A­Y¥¼³]©w®É¤]¹w³]¬°­n¤À³Î«ØÄ³
               ElseIf m_bMiddleCase Then
                  m_PA162 = "Y"
               'end 2019/12/27
               Else
                  m_PA162 = "N"
               End If
            End If
            'end 2019/10/7
         End If
         
         If m_PA162 = "Y" Then
         
            'Added by Morgan 2019/10/7
            '­Y¤w¦¬¤å¤À³Î®×¼u°T®§´£¿ô¤ÎEMail³qª¾©Ó¿ì¤uµ{®v
            strExc(0) = "select dc01,dc02,dc03,dc04 from divisioncase WHERE dc05='" & pa(1) & "' and dc06='" & pa(2) & "' and dc07='" & pa(3) & "' and dc08='" & pa(4) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_bHasDivCase = True
            Else
            'end 2019/10/7
            
               strExc(0) = "select dst09 from divsugtext WHERE dst01='" & pa(1) & "' and dst02='" & pa(2) & "' and dst03='" & pa(3) & "' and dst04='" & pa(4) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  m_bDivSugTextAlert = True
                  
               'Added by Morgan 2020/2/27
               ElseIf intI = 1 Then
                  strExc(1) = "" & RsTemp(0)
                  strExc(0) = "select cp09,sqldatet(cp27) dt,cpm03 from caseprogress,casepropertymap" & _
                     " WHERE cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
                     " and cp10 in ('107','203','204','205') and cp27>0 and cp57 is null" & _
                     " and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp27 desc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp(0) <> strExc(1) Then
                        m_bDivSugTextAlert = True
                        'Modified by Morgan 2021/9/24
                        'm_EditDivSugText = "¤w¦¬¨ì®Ö­ã³qª¾¡A«ØÄ³¤À³Î¤º®e«D" & RsTemp("dt") & "µo¤å¤§" & RsTemp("cpm03") & "­×§ï¤º®e¡A½Ð­×§ï¤À³Î«ØÄ³¤º®e«á¡A¨÷°h¥DºÞ¤W§¹½Z¤é¡A¦A°hµ{§Ç"
                        'Modified by Morgan 2024/5/13 --±Ó²ú
                        'm_EditDivSugText = "¤w¦¬¨ì®Ö­ã³qª¾¡A«ØÄ³¤À³Î¤º®e«D" & RsTemp("dt") & "µo¤å¤§" & RsTemp("cpm03") & "­×§ï¤º®e¡A½Ð­×§ï¤À³Î«ØÄ³¤º®e«á -> email³qª¾¥DºÞ¤W§¹½Z¤é -> email³qª¾¦U°Ïµ{§Ç¤W®Ö­ãµo¤å¡C"
                        m_EditDivSugText = "¤w¦¬¨ì®Ö­ã³qª¾¡A«ØÄ³¤À³Î¤º®e«D" & RsTemp("dt") & "µo¤å¤§" & RsTemp("cpm03") & "­×§ï¤º®e¡A½ÐÂI¿ï""®Ö­ã""­×§ï¤À³Î«ØÄ³¤º®e«á -> ¶]¾úµ{§@·~"
                        'end 2024/5/13
                     End If
                  End If
               'end 2020/2/27
               End If
            End If 'Added by Morgan 2019/10/7
            
            'Modified by Morgan 2019/10/7
            'If m_bDivSugTextAlert Then
            If m_bDivSugTextAlert Or m_bHasDivCase Then
            'end 2019/10/7
            
               Label8.Visible = True
               Text16.Visible = True
               LblFM2(1).Visible = True
               strExc(1) = PUB_GetFCPPromoterNo(strReceiveNo, "1001", m_CP14)
               strExc(0) = "select st04,st02 from staff where st01='" & strExc(1) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp(0) = "2" Then
                     MsgBox "¹w³]©Ó¿ì¤H¡i" & RsTemp(1) & "¡j¤wÂ÷Â¾¡I½Ð¸ß°Ý¤uµ{®v¥DºÞ«á¿é¤J¡C"
                  Else
                     Text16 = strExc(1)
                     LblFM2(1) = "" & RsTemp(1)
                  End If
               End If
            End If
         End If
      End If
   End If
End Sub

'Added by Lydia 2019/05/23
Private Sub txtCRC_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtCRC_GotFocus(Index As Integer)
   TextInverse txtCRC(Index)
End Sub

Private Sub txtCRC_Validate(Index As Integer, Cancel As Boolean)
   If Trim(txtCRC(Index).Text) = "" Then Exit Sub
   Select Case Index
       Case 0 '°É»~¤é´Á
           If PUB_CheckKeyInDate(txtCRC(Index)) = -1 Then
               GoTo JumpExit
           Else
               If InStr("01,11,21", Right(txtCRC(Index), 2)) = 0 Then
                   If MsgBox("´¼¼z§½ªº¤½§i¤é¬°¨C¤ë01,11,21¸¹¡A½Ð°Ý¿é¤J" & txtCRC(Index) & "¬O§_¥¿½T¡H", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
                       GoTo JumpExit
                   End If
               End If
           End If
       Case 1 '´Á§O
           If Len(Trim(txtCRC(Index))) < 2 Then
               MsgBox "´Á§O½Ð¿é¤J01~36´Á¡I", vbCritical
               GoTo JumpExit
           Else
               If Not (Val(txtCRC(Index)) >= 1 And Val(txtCRC(Index)) <= 36) Then
                    MsgBox "´Á§O½Ð¿é¤J01~36´Á¡I", vbCritical
                    GoTo JumpExit
               End If
           End If
   End Select
   
   Exit Sub
   
JumpExit:
   Cancel = True
   txtCRC(Index).SetFocus
   txtCRC_GotFocus Index
End Sub

'Mark by Lydia 2023/03/22 ¾ã¦X¼Ò²Õ¦bPUB_GetApprovalPS
''Added by Lydia 2019/03/11 ³qª¾§i­ã¥[µù(ApprvoalPS) ¼W¥[¡¨³qª¾¤uµ{®vEmail³]©w¡¨
'                                        '°Ñ¦Ò¼Ò²Õ¼g¦bfrm060316_1¡A­Y¦³ÅÜ§óµ{¦¡¨âÃä³£­nÀË¬d¤@¤U
'Private Function GetApprovalPS(dbCaseNo As String, dbFA As String, dbCu As String, Optional ByRef pSubject As String = "", Optional ByRef pContext As String = "") As Boolean
'Dim stSQL As String, iR As Integer
'Dim stCon As String
'Dim rsQuery As ADODB.Recordset
''³vµ§§PÂ_Y¥N²z¤H+X¥Ó½Ð¤H1~5;­Y¦³¤@µ§¥H¤W,¥u¨Ï¥Î²Ä¤@µ§²Å¦X
'Dim m_Subject As String
'Dim m_Context As String
'Dim iCall As Integer, iRound As Integer
'Dim tmpArr As Variant
'
'   '§PÂ_¦³´X­Ó¥Ó½Ð¤H
'   tmpArr = Split(dbCu, ",")
'   For iR = 0 To UBound(tmpArr)
'       If Trim(tmpArr(iR)) <> "" Then
'           iCall = iCall + 1
'       End If
'   Next iR
'
'   For iRound = 1 To iCall
'        '¶¶§Ç 1.¥»©Ò®×¸¹ 2.¥N²z¤H+¥Ó½Ð¤H 3.¥N²z¤H 4.¥Ó½Ð¤H
'        stSQL = "select 0 Od1, APS13, APS14 from ApprovalPS where APS03='" & dbCaseNo & "' " & stCon & _
'           " union select 1 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 2 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " union select 3 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05 is null" & stCon & _
'           " union select 4 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 5 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " union select 6 Od1, APS13, APS14 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05 is null" & stCon & _
'           " union select 7 Od1, APS13, APS14 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
'           " union select 8 Od1, APS13, APS14 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
'           " order by Od1, APS13"
'            iR = 1
'            Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
'            If iR = 1 Then
'               'Modified by Lydia 2021/03/09 ­«·s¾ã²z,³vµ§§PÂ_¥u¨Ï¥Î²Ä¤@µ§²Å¦X;
'               rsQuery.MoveFirst
'               Do While Not rsQuery.EOF
'                    If "" & rsQuery.Fields("APS13") <> "" And rsQuery.Fields("APS14") <> "" Then
'                         m_Subject = "" & rsQuery.Fields("APS13")
'                         m_Context = "" & rsQuery.Fields("APS14")
'                         GoTo JumpToEnd
'                    End If
'                    rsQuery.MoveNext
'               Loop
'               'end 2021/03/09
'            End If
'   Next iRound
'
'JumpToEnd:
'   pSubject = m_Subject
'   pContext = m_Context
'   If pSubject <> "" And pContext <> "" Then
'       GetApprovalPS = True
'   End If
'   Set rsQuery = Nothing
'End Function



