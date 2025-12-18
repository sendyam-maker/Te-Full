VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060301_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "國內公開公報資料維護"
   ClientHeight    =   6140
   ClientLeft      =   630
   ClientTop       =   2700
   ClientWidth     =   8960
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6140
   ScaleWidth      =   8960
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   2340
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2430
      Width           =   1815
   End
   Begin VB.TextBox textNA01 
      Height          =   270
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2430
      Width           =   852
   End
   Begin VB.TextBox txtTPG18 
      Height          =   270
      Left            =   8220
      MaxLength       =   1
      TabIndex        =   14
      Top             =   1260
      Width           =   315
   End
   Begin VB.TextBox text01 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   930
      Width           =   2772
   End
   Begin VB.TextBox txtTPG15 
      Height          =   270
      Left            =   5760
      MaxLength       =   15
      TabIndex        =   16
      Top             =   1845
      Width           =   2235
   End
   Begin VB.TextBox txtTPG16 
      Height          =   270
      Left            =   5760
      MaxLength       =   2
      TabIndex        =   17
      Top             =   2145
      Width           =   615
   End
   Begin VB.TextBox txtTPG17 
      Height          =   270
      Left            =   7800
      MaxLength       =   2
      TabIndex        =   18
      Top             =   2145
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   5760
      MaxLength       =   1
      TabIndex        =   13
      Top             =   1260
      Width           =   345
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1410
      MaxLength       =   12
      TabIndex        =   0
      Top             =   630
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一筆公告號(&N)"
      Height          =   405
      Index           =   1
      Left            =   1710
      TabIndex        =   55
      Top             =   120
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一筆公告號(&P)"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   54
      Top             =   120
      Width           =   1560
   End
   Begin VB.TextBox text06_2 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   2340
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2130
      Width           =   1812
   End
   Begin VB.TextBox text09 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   5760
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1560
      Width           =   2772
   End
   Begin VB.TextBox text07_1 
      Height          =   270
      Left            =   5760
      MaxLength       =   4
      TabIndex        =   10
      Top             =   630
      Width           =   852
   End
   Begin VB.TextBox text06_1 
      Height          =   270
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2115
      Width           =   852
   End
   Begin VB.TextBox text05 
      Height          =   270
      Left            =   2850
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1815
      Width           =   852
   End
   Begin VB.TextBox text04 
      Height          =   270
      Left            =   1410
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1815
      Width           =   852
   End
   Begin VB.TextBox text03 
      Height          =   270
      Left            =   1410
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1530
      Width           =   2772
   End
   Begin VB.TextBox text02 
      Height          =   270
      Left            =   1410
      MaxLength       =   9
      TabIndex        =   2
      Top             =   1230
      Width           =   2772
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6810
      TabIndex        =   43
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton buttonCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   7650
      TabIndex        =   44
      Top             =   120
      Width           =   1200
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3195
      Left            =   30
      TabIndex        =   62
      Top             =   2910
      Width           =   8895
      _ExtentX        =   15699
      _ExtentY        =   5644
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "中文名稱"
      TabPicture(0)   =   "frm04060301_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label17"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label18"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label19"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label20"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label21"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label23"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label24"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label25"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label26"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtcAppl(0)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtcAppl(1)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtcAppl(2)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtcAppl(3)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtcAppl(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtcAppl(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtcAppl(6)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtcAppl(7)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtcAppl(8)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtcAppl(9)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "英文名稱"
      TabPicture(1)   =   "frm04060301_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txteAppl(9)"
      Tab(1).Control(1)=   "txteAppl(8)"
      Tab(1).Control(2)=   "txteAppl(7)"
      Tab(1).Control(3)=   "txteAppl(6)"
      Tab(1).Control(4)=   "txteAppl(5)"
      Tab(1).Control(5)=   "txteAppl(4)"
      Tab(1).Control(6)=   "txteAppl(3)"
      Tab(1).Control(7)=   "txteAppl(2)"
      Tab(1).Control(8)=   "txteAppl(1)"
      Tab(1).Control(9)=   "txteAppl(0)"
      Tab(1).Control(10)=   "Label29"
      Tab(1).Control(11)=   "Label30"
      Tab(1).Control(12)=   "Label31"
      Tab(1).Control(13)=   "Label32"
      Tab(1).Control(14)=   "Label33"
      Tab(1).Control(15)=   "Label34"
      Tab(1).Control(16)=   "Label35"
      Tab(1).Control(17)=   "Label36"
      Tab(1).Control(18)=   "Label37"
      Tab(1).Control(19)=   "Label38"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "其他"
      TabPicture(2)   =   "frm04060301_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTPG39"
      Tab(2).Control(1)=   "txtTPG40"
      Tab(2).Control(2)=   "txtTPG41"
      Tab(2).Control(3)=   "txtTPG42"
      Tab(2).Control(4)=   "Label39"
      Tab(2).Control(5)=   "Label40"
      Tab(2).Control(6)=   "Label41"
      Tab(2).Control(7)=   "Label42"
      Tab(2).Control(8)=   "Label43"
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtTPG39 
         Height          =   264
         Left            =   -73530
         MaxLength       =   7
         TabIndex        =   39
         Top             =   420
         Width           =   1185
      End
      Begin VB.TextBox txtTPG40 
         Height          =   264
         Left            =   -73530
         MaxLength       =   7
         TabIndex        =   40
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtTPG41 
         Height          =   600
         Left            =   -73530
         MaxLength       =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   41
         Top             =   1020
         Width           =   7275
      End
      Begin VB.TextBox txtTPG42 
         Height          =   600
         Left            =   -73530
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   42
         Top             =   1650
         Width           =   7275
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   9
         Left            =   1380
         TabIndex        =   28
         Top             =   2820
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   8
         Left            =   1380
         TabIndex        =   27
         Top             =   2550
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   7
         Left            =   1380
         TabIndex        =   26
         Top             =   2280
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   6
         Left            =   1380
         TabIndex        =   25
         Top             =   2010
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   5
         Left            =   1380
         TabIndex        =   24
         Top             =   1740
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   4
         Left            =   1380
         TabIndex        =   23
         Top             =   1470
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   3
         Left            =   1380
         TabIndex        =   22
         Top             =   1200
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   2
         Left            =   1380
         TabIndex        =   21
         Top             =   930
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   20
         Top             =   660
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   19
         Top             =   390
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   9
         Left            =   -73650
         TabIndex        =   38
         Top             =   2820
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   8
         Left            =   -73650
         TabIndex        =   37
         Top             =   2550
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   7
         Left            =   -73650
         TabIndex        =   36
         Top             =   2280
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   6
         Left            =   -73650
         TabIndex        =   35
         Top             =   2010
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   5
         Left            =   -73650
         TabIndex        =   34
         Top             =   1740
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   4
         Left            =   -73650
         TabIndex        =   33
         Top             =   1470
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   3
         Left            =   -73650
         TabIndex        =   32
         Top             =   1200
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   2
         Left            =   -73650
         TabIndex        =   31
         Top             =   930
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   1
         Left            =   -73650
         TabIndex        =   30
         Top             =   660
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   0
         Left            =   -73650
         TabIndex        =   29
         Top             =   390
         Width           =   7425
         VariousPropertyBits=   671107099
         MaxLength       =   150
         Size            =   "13097;529"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱10 :"
         Height          =   180
         Left            =   120
         TabIndex        =   87
         Top             =   2850
         Width           =   1170
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱9 :"
         Height          =   180
         Left            =   120
         TabIndex        =   86
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱8 :"
         Height          =   180
         Left            =   120
         TabIndex        =   85
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱7 :"
         Height          =   180
         Left            =   120
         TabIndex        =   84
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱6 :"
         Height          =   180
         Left            =   120
         TabIndex        =   83
         Top             =   1770
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱5 :"
         Height          =   180
         Left            =   120
         TabIndex        =   82
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱4 :"
         Height          =   180
         Left            =   120
         TabIndex        =   81
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱3 :"
         Height          =   180
         Left            =   120
         TabIndex        =   80
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱2 :"
         Height          =   180
         Left            =   120
         TabIndex        =   79
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱1 :"
         Height          =   180
         Left            =   120
         TabIndex        =   78
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱10 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   77
         Top             =   2850
         Width           =   1170
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱9 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   76
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱8 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   75
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱7 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   74
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱6 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   73
         Top             =   1770
         Width           =   1080
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱5 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   72
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱4 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   71
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱3 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   70
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱2 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   69
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱1 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   68
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "申請日 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   67
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "最早優先權日 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   66
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "優先權號 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   65
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "優先權國家 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   64
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label43 
         AutoSize        =   -1  'True
         Caption         =   "優先權多國時以分號做區隔"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   -74790
         TabIndex        =   63
         Top             =   2310
         Width           =   2340
      End
   End
   Begin MSForms.TextBox text08 
      Height          =   300
      Left            =   5760
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   960
      Width           =   2775
      VariousPropertyBits=   671105055
      Size            =   "4895;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text07_2 
      Height          =   300
      Left            =   6660
      TabIndex        =   11
      Top             =   630
      Width           =   1815
      VariousPropertyBits=   671107099
      Size            =   "3201;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label28 
      Caption         =   "（臺灣要分到縣市統計用）"
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   90
      TabIndex        =   91
      Top             =   2700
      Width           =   2355
   End
   Begin VB.Label Label44 
      Caption         =   "地區名稱 :"
      Height          =   240
      Left            =   330
      TabIndex        =   90
      Top             =   2430
      Width           =   990
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "案件屬性  :"
      Height          =   180
      Left            =   7320
      TabIndex        =   89
      Top             =   1290
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "事務所名稱 :"
      Height          =   180
      Left            =   4500
      TabIndex        =   88
      Top             =   990
      Width           =   990
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "國際分類號  :"
      Height          =   180
      Left            =   4500
      TabIndex        =   61
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "IPC分類  :"
      Height          =   180
      Left            =   4500
      TabIndex        =   60
      Top             =   2190
      Width           =   765
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "產業別分類  :"
      Height          =   180
      Left            =   6660
      TabIndex        =   59
      Top             =   2190
      Width           =   1035
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請實體審查 :           (Y:有  N: 無 )"
      Height          =   180
      Left            =   4500
      TabIndex        =   58
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "新申請案號 :"
      Height          =   180
      Left            =   330
      TabIndex        =   57
      Top             =   630
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "上下筆移動時會儲存此筆記錄 !!!"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   3330
      TabIndex        =   56
      Top             =   180
      Width           =   3405
   End
   Begin VB.Label Label10 
      Caption         =   "期"
      Height          =   255
      Left            =   3840
      TabIndex        =   53
      Top             =   1830
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "卷"
      Height          =   255
      Left            =   2340
      TabIndex        =   52
      Top             =   1830
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "本所案號  :"
      Height          =   180
      Left            =   4500
      TabIndex        =   51
      Top             =   1620
      Width           =   855
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   4500
      TabIndex        =   50
      Top             =   630
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍 :"
      Height          =   180
      Left            =   330
      TabIndex        =   49
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公報 :"
      Height          =   180
      Left            =   330
      TabIndex        =   48
      Top             =   1860
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "公開日 :"
      Height          =   180
      Left            =   330
      TabIndex        =   47
      Top             =   1590
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公開號 :"
      Height          =   180
      Left            =   330
      TabIndex        =   46
      Top             =   1290
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   330
      TabIndex        =   45
      Top             =   960
      Width           =   810
   End
End
Attribute VB_Name = "frm04060301_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/24 改成Form2.0 (text07_2,text08,txtcAppl,txteAppl,Combo1
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

'Memo by Morgan 2008/11/18
'國內案公開不必控管是否有國外案未發文，若要管制則將於相同案之分案、發文、提申也加入國內案檢查

Dim m_EditMode As Integer
Dim m_DataKey As String
Dim m_CurrTPG02 As String
Dim m_CurrTPG03 As String
Dim m_CurrTPG04 As String
Dim m_CurrTPG05 As String


'使用者按下確定的按鍵
Private Sub buttonOK_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim nNo As Long
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strFreeAgentCode As String
   
   Select Case m_EditMode
      ' 新增或變更
      Case 0, 1:
         If CheckDataValid() = True Then
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            Select Case m_EditMode
               Case 0:
                  nNo = Val(text02.Text)
                  nNo = nNo + 1
                  m_CurrTPG02 = CStr(nNo)
                  m_CurrTPG03 = text03
                  m_CurrTPG04 = text04
                  m_CurrTPG05 = text05
            End Select
            '隱藏新申請案號欄位
            Me.Text1.Visible = False
            Me.Label12.Visible = False
            Me.Hide
            frm04060301_1.Show
            frm04060301_1.SetInputTPG01
         End If
      ' 刪除
      Case 3:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         If MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit) = vbYes Then
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            '隱藏新申請案號欄位
            Me.Text1.Visible = False
            Me.Label12.Visible = False
            Me.Hide
            frm04060301_1.Show
            frm04060301_1.SetInputTPG01
         End If
      Case Else:
        If m_EditMode = 1 Then
            Me.text02.SetFocus
        End If
        '隱藏新申請案號欄位
        Me.Text1.Visible = False
        Me.Label12.Visible = False
        Me.Hide
        frm04060301_1.Show
        frm04060301_1.SetInputTPG01
   End Select
   frm04060301_1.UpdateRecord m_DataKey
EXITSUB:
End Sub

' 使用者按下取消的按鍵
Private Sub buttonCancel_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
    strTit = "詢問"
    strMsg = "你並未存檔, 確定離開嗎?"
    'Add By Cheng 2002/11/21
    If m_EditMode = 1 Then
        Me.text02.SetFocus
    End If
    '隱藏新申請案號欄位
    Me.Text1.Visible = False
    Me.Label12.Visible = False
    Me.Hide
    frm04060301_1.Show
    frm04060301_1.SetInputTPG01
End Sub

' 設定控制項中初始的值 (申請案號的值)
Public Sub SetData(ByVal textKey As String)
    m_DataKey = textKey
    SSTab1.Tab = 0 'Add By Sindy 2018/11/12
End Sub

' 設定編輯資料的模式 (新增或修改)
Public Sub SetMode(ByVal nMode As Integer)
    Select Case nMode
        Case 0, 1, 2, 3:
            m_EditMode = nMode
        Case Else:
    End Select
End Sub

'Remove by Morgan 2011/5/17 改用公用函數
'Private Function GetFreeAgentCode() As String
'   Dim strLastAgent As String
'   Dim nNumber As Integer
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   strLastAgent = "01"
'   strSql = "SELECT * FROM TAgent WHERE TA01 = 'P'"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenDynamic
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Do While rsTmp.EOF = False
'         If IsNull(rsTmp.Fields("TA02")) = False Then
'            If Val(rsTmp.Fields("TA02")) > Val(strLastAgent) Then
'               strLastAgent = rsTmp.Fields("TA02")
'            End If
'         End If
'         rsTmp.MoveNext
'      Loop
'   End If
'
'   nNumber = Val(strLastAgent) + 1
'   Select Case Len(strLastAgent)
'      Case 1:
'         If Len(nNumber) > 1 Then
'            GetFreeAgentCode = Format(nNumber, "00")
'         Else
'            GetFreeAgentCode = Format(nNumber, "0")
'         End If
'      Case 2:
'         If Len(nNumber) > 2 Then
'            GetFreeAgentCode = Format(nNumber, "000")
'         Else
'            GetFreeAgentCode = Format(nNumber, "00")
'         End If
'      Case 3:
'         GetFreeAgentCode = Format(nNumber, "000")
'   End Select
'
'   Set rsTmp = Nothing
'End Function

Private Function IsTPG02Exist(ByVal strTPG02 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsTPG02Exist = False
   strSql = "SELECT * FROM TPGAZETTE " & _
            "WHERE TPG02 = '" & strTPG02 & "' AND " & _
                  "TPG01 <> '" & m_DataKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      IsTPG02Exist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 此模組在處理資料到資料庫的工作
Public Function OnWork() As Boolean
   Dim strSql As String
   Dim strDate As String
   Dim strFreeAgentCode As String
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim i As Integer, strcAppl(0 To 10) As String
   Dim streAppl(0 To 10) As String 'Add By Sindy 2018/11/12
   
On Error GoTo ErrorHandler
OnWork = True
cnnConnection.BeginTrans

   strDate = Empty
   If IsEmpty(text03) = False Then
      strDate = ChangeTStringToWString(text03)
   End If
   
   'Add By Sindy 2018/11/13
   For i = 0 To 9
      strcAppl(i) = ""
      streAppl(i) = ""
   Next i
   For i = 0 To 9
      If Trim(txtcAppl(i)) <> "" Then
         strcAppl(i) = txtcAppl(i).Text
      End If
      If Trim(txteAppl(i)) <> "" Then
         streAppl(i) = txteAppl(i).Text
      End If
   Next i
   '2018/11/13 END
   
   Select Case m_EditMode
      ' 新增資料到國內專利公報檔
      Case 0:
         'Modify By Sindy 2013/10/29 +TPG15,TPG16,TPG17
'         If strDate <> Empty Then
         'Modify By Sindy 2019/9/4 +,Trim(Mid(Combo1.Text, 4))
           'Modify by Amy 2021/09/17 +ChgSQL ex:申請案號110102309,申請人:日商大�曭悒鬫陪迨膝q 無法存檔
            strSql = "Insert into TPGAZETTE " & _
                     "(TPG01,TPG02" & IIf(strDate <> Empty, ",TPG03", "") & ",TPG04,TPG05,TPG06,TPG07,TPG08,TPG09,TPG15,TPG16,TPG17,TPG18" & _
                     ",TPG19,TPG20,TPG21,TPG22,TPG23,TPG24,TPG25,TPG26,TPG27,TPG28" & _
                     ",TPG29,TPG30,TPG31,TPG32,TPG33,TPG34,TPG35,TPG36,TPG37,TPG38" & _
                     ",TPG39,TPG40,TPG41,TPG42,TPG43" & _
                     ") Values ('" & text01 & "','" & text02 & "'" & IIf(strDate <> Empty, "," & strDate, "") & ",'" & Format(text04, "00") & "','" & _
                     Format(text05, "00") & "','" & text06_1 & "','" & text07_1 & "','" & text08 & "','" & Me.Text2.Text & "'" & _
                     "," & CNULL(txtTPG15) & "," & CNULL(Format(txtTPG16.Text, "00")) & "," & CNULL(txtTPG17) & "," & CNULL(txtTPG18) & _
                     "," & CNULL(ChgSQL(strcAppl(0))) & "," & CNULL(ChgSQL(strcAppl(1))) & "," & CNULL(ChgSQL(strcAppl(2))) & "," & CNULL(ChgSQL(strcAppl(3))) & "," & CNULL(ChgSQL(strcAppl(4))) & _
                     "," & CNULL(ChgSQL(strcAppl(5))) & "," & CNULL(ChgSQL(strcAppl(6))) & "," & CNULL(ChgSQL(strcAppl(7))) & "," & CNULL(ChgSQL(strcAppl(8))) & "," & CNULL(ChgSQL(strcAppl(9))) & _
                     "," & CNULL(ChgSQL(streAppl(0))) & "," & CNULL(ChgSQL(streAppl(1))) & "," & CNULL(ChgSQL(streAppl(2))) & "," & CNULL(ChgSQL(streAppl(3))) & "," & CNULL(ChgSQL(streAppl(4))) & _
                     "," & CNULL(ChgSQL(streAppl(5))) & "," & CNULL(ChgSQL(streAppl(6))) & "," & CNULL(ChgSQL(streAppl(7))) & "," & CNULL(ChgSQL(streAppl(8))) & "," & CNULL(ChgSQL(streAppl(9))) & _
                     "," & CNULL(DBDATE(txtTPG39), True) & "," & CNULL(DBDATE(txtTPG40), True) & "," & CNULL(txtTPG41) & "," & CNULL(txtTPG42) & "," & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                     ")"
'         Else
'            strSql = "Insert into TPGAZETTE " & _
'                     "(TPG01, TPG02, TPG04, TPG05, TPG06, TPG07, TPG08, TPG09,TPG15,TPG16,TPG17) " & _
'                     "Values ('" & text01 & "','" & text02 & "','" & Format(text04, "00") & "','" & _
'                     Format(text05, "00") & "','" & text06_1 & "','" & text07_1 & "','" & text08 & "','" & Me.Text2.Text & "'" & _
'                     "," & CNULL(txtTPG15) & "," & CNULL(txtTPG16) & "," & CNULL(txtTPG17) & ")"
'         End If
         cnnConnection.Execute strSql
         
          'Modify by Morgan 2004/8/4
          '本所案件才更新
         If text09.Text <> "" Then
            ' 更新專利基本檔的公開日及公開號
            strSql = "UPDATE Patent SET PA12 = " & ChangeTStringToWString(text03) & ", " & _
                                       "PA13 = '" & text02 & "' " & _
                     "WHERE PA11 = '" & m_DataKey & "'"
            cnnConnection.Execute strSql
         End If
      
      ' 修改專利公報檔的資料
      Case 1:
            '若非修改申請案號
            If Me.Text1.Visible = False Then
                'Modify By Sindy 2013/10/29 +TPG15,TPG16,TPG17
'                If strDate <> Empty Then
                 'Modify By Sindy 2019/9/4 +,Trim(Mid(Combo1.Text, 4))
                 'Modify by Amy 2021/09/17 +ChgSQL ex:申請案號110102309,申請人:日商大�曭悒鬫陪迨膝q 無法存檔
                  strSql = "Update TPGAZETTE " & _
                           "Set TPG01='" & text01 & "',TPG02='" & text02 & "'" & IIf(strDate <> Empty, ",TPG03=" & strDate, "") & ",TPG04='" & Format(text04, "00") & "'" & _
                           ",TPG05='" & Format(text05, "00") & "'," & "TPG06='" & text06_1 & "'," & "TPG07='" & text07_1 & "'," & "TPG08='" & text08 & "', TPG09='" & Me.Text2.Text & "'" & _
                           ",TPG15=" & CNULL(txtTPG15) & ",TPG16=" & CNULL(Format(txtTPG16.Text, "00")) & ",TPG17=" & CNULL(txtTPG17) & ",TPG18=" & CNULL(txtTPG18) & _
                           ",TPG19=" & CNULL(ChgSQL(strcAppl(0))) & ",TPG20=" & CNULL(ChgSQL(strcAppl(1))) & ",TPG21=" & CNULL(ChgSQL(strcAppl(2))) & ",TPG22=" & CNULL(ChgSQL(strcAppl(3))) & ",TPG23=" & CNULL(ChgSQL(strcAppl(4))) & _
                           ",TPG24=" & CNULL(ChgSQL(strcAppl(5))) & ",TPG25=" & CNULL(ChgSQL(strcAppl(6))) & ",TPG26=" & CNULL(ChgSQL(strcAppl(7))) & ",TPG27=" & CNULL(ChgSQL(strcAppl(8))) & ",TPG28=" & CNULL(ChgSQL(strcAppl(9))) & _
                           ",TPG29=" & CNULL(ChgSQL(streAppl(0))) & ",TPG30=" & CNULL(ChgSQL(streAppl(1))) & ",TPG31=" & CNULL(ChgSQL(streAppl(2))) & ",TPG32=" & CNULL(ChgSQL(streAppl(3))) & ",TPG33=" & CNULL(ChgSQL(streAppl(4))) & _
                           ",TPG34=" & CNULL(ChgSQL(streAppl(5))) & ",TPG35=" & CNULL(ChgSQL(streAppl(6))) & ",TPG36=" & CNULL(ChgSQL(streAppl(7))) & ",TPG37=" & CNULL(ChgSQL(streAppl(8))) & ",TPG38=" & CNULL(ChgSQL(streAppl(9))) & _
                           ",TPG39=" & CNULL(DBDATE(txtTPG39), True) & ",TPG40=" & CNULL(DBDATE(txtTPG40), True) & _
                           ",TPG41=" & CNULL(txtTPG41) & ",TPG42=" & CNULL(txtTPG42) & ",TPG43=" & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                           " Where TPG01='" & text01 & "'"
'                Else
'                   strSql = "Update TPGAZETTE " & _
'                            "Set TPG01='" & text01 & "'," & "TPG02='" & text02 & "'," & "TPG04='" & Format(text04, "00") & "'," & _
'                            "TPG05='" & Format(text05, "00") & "'," & "TPG06='" & text06_1 & "'," & "TPG07='" & text07_1 & "'," & "TPG08='" & text08 & "', TPG09='" & Me.Text2.Text & "'" & _
'                            ",TPG15=" & CNULL(txtTPG15) & ",TPG16=" & CNULL(txtTPG16) & ",TPG17=" & CNULL(txtTPG17) & _
'                            " Where TPG01='" & text01 & "'"
'                End If
                cnnConnection.Execute strSql
                
                'Modify by Morgan 2004/8/4
                '本所案件才更新
               If text09.Text <> "" Then
                  ' 更新專利基本檔的公開日及公開號
                  strSql = "UPDATE Patent SET PA12 = " & ChangeTStringToWString(text03) & ", " & _
                                             "PA13 = '" & text02 & "' " & _
                           "WHERE PA11 = '" & m_DataKey & "'"
                  cnnConnection.Execute strSql
               End If
            '修改申請案號
            Else
               '新增新申請案號資料
               'Modify By Sindy 2013/10/29 +TPG15,TPG16,TPG17
               'Modify By Sindy 2019/9/4 +,Trim(Mid(Combo1.Text, 4))
               'Modify by Amy 2021/09/17 +ChgSQL ex:申請案號110102309,申請人:日商大�曭悒鬫陪迨膝q 無法存檔
               strSql = "Insert Into TPGAZETTE (TPG01,TPG02,TPG03,TPG04,TPG05,TPG06,TPG07,TPG08,TPG09,TPG15,TPG16,TPG17,TPG18" & _
                        ",TPG19,TPG20,TPG21,TPG22,TPG23,TPG24,TPG25,TPG26,TPG27,TPG28" & _
                        ",TPG29,TPG30,TPG31,TPG32,TPG33,TPG34,TPG35,TPG36,TPG37,TPG38" & _
                        ",TPG39,TPG40,TPG41,TPG42,TPG43" & _
                        ") Values ('" & Me.Text1.Text & "','" & Me.text02.Text & "'," & IIf(Me.text03.Text = "", "NULL", DBDATE(Me.text03.Text)) & ",'" & Format(Me.text04.Text, "00") & "','" & _
                        Format(Me.text05.Text, "00") & "','" & Me.text06_1.Text & "','" & Me.text07_1.Text & "','" & Me.text08.Text & "','" & Me.Text2.Text & "'" & _
                        "," & CNULL(txtTPG15) & "," & CNULL(Format(txtTPG16.Text, "00")) & "," & CNULL(txtTPG17) & "," & CNULL(txtTPG18) & _
                        "," & CNULL(ChgSQL(strcAppl(0))) & "," & CNULL(ChgSQL(strcAppl(1))) & "," & CNULL(ChgSQL(strcAppl(2))) & "," & CNULL(ChgSQL(strcAppl(3))) & "," & CNULL(ChgSQL(strcAppl(4))) & _
                        "," & CNULL(ChgSQL(strcAppl(5))) & "," & CNULL(ChgSQL(strcAppl(6))) & "," & CNULL(ChgSQL(strcAppl(7))) & "," & CNULL(ChgSQL(strcAppl(8))) & "," & CNULL(ChgSQL(strcAppl(9))) & _
                        "," & CNULL(ChgSQL(streAppl(0))) & "," & CNULL(ChgSQL(streAppl(1))) & "," & CNULL(ChgSQL(streAppl(2))) & "," & CNULL(ChgSQL(streAppl(3))) & "," & CNULL(ChgSQL(streAppl(4))) & _
                        "," & CNULL(ChgSQL(streAppl(5))) & "," & CNULL(ChgSQL(streAppl(6))) & "," & CNULL(ChgSQL(streAppl(7))) & "," & CNULL(ChgSQL(streAppl(8))) & "," & CNULL(ChgSQL(streAppl(9))) & _
                        "," & CNULL(DBDATE(txtTPG39), True) & "," & CNULL(DBDATE(txtTPG40), True) & "," & CNULL(txtTPG41) & "," & CNULL(txtTPG42) & "," & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                        ")"
               cnnConnection.Execute strSql
               
               ' 更新專利基本檔的公開日及公開號
               strSql = "UPDATE Patent SET PA12 = " & IIf(Me.text03.Text = "", "NULL", ChangeTStringToWString(text03)) & ", " & _
                                          "PA13 = '" & text02 & "' " & _
                        "WHERE PA11 = '" & Me.Text1.Text & "'"
               cnnConnection.Execute strSql
               
               '刪除舊申請案號
               strSql = "Delete From TPGAZETTE Where TPG01='" & text01.Text & "'"
               cnnConnection.Execute strSql
                
               'Modify by Morgan 2004/8/4
               '本所案件才更新
               If text09.Text <> "" Then
                  ' 更新專利基本檔的公開日及公開號
                  strSql = "UPDATE Patent SET PA12 = NULL , PA13 = NULL " & _
                           "WHERE PA11 = '" & Me.text01.Text & "'"
                  cnnConnection.Execute strSql
               End If
            End If
      Case 3: '刪除
         strSql = "Delete From TPGAZETTE where TPG01 = '" & m_DataKey & "'"
         cnnConnection.Execute strSql
   End Select
cnnConnection.CommitTrans
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    OnWork = False
End Function

Public Function UpdateCtrlData(ByVal nAction As Integer) As Boolean
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   UpdateCtrlData = True
   Select Case nAction
      ' 依照申請案號帶出所有相關資料
      Case 0:
         If m_EditMode = 1 Or m_EditMode = 2 Or m_EditMode = 3 Then
            Set rsTmp = New ADODB.Recordset
            strSql = "Select * from TPGAZETTE where TPG01 = '" & m_DataKey & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("TPG02")) = False Then
                  text02 = rsTmp.Fields("TPG02")
               End If
               If IsNull(rsTmp.Fields("TPG03")) = False Then
                  text03 = ChangeWStringToTString(rsTmp.Fields("TPG03"))
               End If
               If IsNull(rsTmp.Fields("TPG04")) = False Then
                  text04 = rsTmp.Fields("TPG04")
               End If
               If IsNull(rsTmp.Fields("TPG05")) = False Then
                  text05 = rsTmp.Fields("TPG05")
               End If
               If IsNull(rsTmp.Fields("TPG06")) = False Then
                  text06_1 = rsTmp.Fields("TPG06")
               End If
               If IsNull(rsTmp.Fields("TPG07")) = False Then
                  text07_1 = rsTmp.Fields("TPG07")
               End If
               If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(1)
               End If
               If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(2)
               End If
               If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(3)
               End If
                '申請實體審查
               If IsNull(rsTmp.Fields("TPG09")) = False Then
                  Text2.Text = rsTmp.Fields("TPG09")
               End If
               'Add By Sindy 2013/10/29
               If IsNull(rsTmp.Fields("TPG15")) = False Then
                  txtTPG15.Text = rsTmp.Fields("TPG15")
               End If
               If IsNull(rsTmp.Fields("TPG16")) = False Then
                  txtTPG16.Text = rsTmp.Fields("TPG16")
               End If
               If IsNull(rsTmp.Fields("TPG17")) = False Then
                  txtTPG17.Text = rsTmp.Fields("TPG17")
               End If
               '2013/10/29 END
               
               'Add By Sindy 2018/11/13
               txtTPG18.Text = "" & rsTmp.Fields("TPG18")
               txtcAppl(0).Text = "" & rsTmp.Fields("TPG19")
               txtcAppl(1).Text = "" & rsTmp.Fields("TPG20")
               txtcAppl(2).Text = "" & rsTmp.Fields("TPG21")
               txtcAppl(3).Text = "" & rsTmp.Fields("TPG22")
               txtcAppl(4).Text = "" & rsTmp.Fields("TPG23")
               txtcAppl(5).Text = "" & rsTmp.Fields("TPG24")
               txtcAppl(6).Text = "" & rsTmp.Fields("TPG25")
               txtcAppl(7).Text = "" & rsTmp.Fields("TPG26")
               txtcAppl(8).Text = "" & rsTmp.Fields("TPG27")
               txtcAppl(9).Text = "" & rsTmp.Fields("TPG28")
               txteAppl(0).Text = "" & rsTmp.Fields("TPG29")
               txteAppl(1).Text = "" & rsTmp.Fields("TPG30")
               txteAppl(2).Text = "" & rsTmp.Fields("TPG31")
               txteAppl(3).Text = "" & rsTmp.Fields("TPG32")
               txteAppl(4).Text = "" & rsTmp.Fields("TPG33")
               txteAppl(5).Text = "" & rsTmp.Fields("TPG34")
               txteAppl(6).Text = "" & rsTmp.Fields("TPG35")
               txteAppl(7).Text = "" & rsTmp.Fields("TPG36")
               txteAppl(8).Text = "" & rsTmp.Fields("TPG37")
               txteAppl(9).Text = "" & rsTmp.Fields("TPG38")
               If IsNull(rsTmp.Fields("TPG39")) = False Then
                  txtTPG39 = ChangeWStringToTString(rsTmp.Fields("TPG39"))
               End If
               If IsNull(rsTmp.Fields("TPG40")) = False Then
                  txtTPG40 = ChangeWStringToTString(rsTmp.Fields("TPG40"))
               End If
               txtTPG41.Text = "" & rsTmp.Fields("TPG41")
               txtTPG42.Text = "" & rsTmp.Fields("TPG42")
               '2018/11/13 END
               
               'Add By Sindy 2019/9/4
               Combo1.Text = "" & rsTmp.Fields("TPG43")
               If IsEmptyText(Combo1.Text) = False Then
                  textNA01 = GetNationNo(Combo1.Text)
                  Combo1.Text = textNA01 & " " & Combo1.Text
               End If
               '2019/9/4 END
            Else
               UpdateCtrlData = False
            End If
            rsTmp.Close
         End If
         '2013/9/16 modify by sonia 加入m_EditMode = 1否則本所案件修改時不會抓101140032(FCP-046707)
         If m_EditMode = 0 Or m_EditMode = 1 Then
            UpdateCtrlData = UpdateCtrlData(3)
         End If
      ' 依照國籍代號帶出國家名稱
      Case 1:
         Set rsTmp = New ADODB.Recordset
         strSql = "Select * from NATION where NA01 = '" & text06_1 & "'"
         text06_2 = Empty
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            If IsNull(rsTmp.Fields("NA03")) = False Then
               text06_2 = rsTmp.Fields("NA03")
            End If
         Else
            UpdateCtrlData = False
         End If
         rsTmp.Close
      ' 依照代理人代號帶出代理人名稱及事務所名稱
      Case 2:
         If IsEmptyText(text07_1) = False Then
            Set rsTmp = New ADODB.Recordset
            'Modify by Morgan 2010/10/18
            'strSql = "Select * from TAGENT where TA02 = '" & text07_1 & "'"
            strSql = "Select * from TAGENT where TA01 = 'P' AND TA02 = '" & text07_1 & "'"
            text07_2 = Empty
            text08 = Empty
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("TA03")) = False Then
                  text07_2 = rsTmp.Fields("TA03")
               End If
               If IsNull(rsTmp.Fields("TA04")) = False Then
                  text08 = rsTmp.Fields("TA04")
               End If
            Else
               UpdateCtrlData = False
            End If
            rsTmp.Close
         End If
      ' 依照申請案號帶出本所案號(申請國家為台灣的資料才顯示本所案號)
      Case 3:
         Set rsTmp = New ADODB.Recordset
         'Modify by Morgan 2004/8/4
         '申請案才要帶
         strSql = "SELECT * FROM Patent " & _
                  "WHERE PA11 = '" & m_DataKey & "' AND " & _
                        "PA09 = '000' AND PA23='1'"
         text09 = Empty
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            text09 = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
            '若為新增狀況
            If m_EditMode = 0 Then
                Me.text07_1.Text = "01"
                text07_1_Validate False
            End If
         Else
            UpdateCtrlData = False
         End If
         rsTmp.Close
   End Select
   Set rsTmp = Nothing
End Function

Public Sub UpdateState()
   Select Case m_EditMode
      Case 2, 3:
         text02.Locked = True
         text03.Locked = True
         text04.Locked = True
         text05.Locked = True
         text06_1.Locked = True
         text07_1.Locked = True
         Text2.Locked = True
         'Add By Sindy 2013/10/29
         txtTPG15.Locked = True
         txtTPG16.Locked = True
         txtTPG17.Locked = True
         '2013/10/29 END
         'Add By Sindy 2018/11/12
         txtTPG18.Locked = True
         txtcAppl(0).Locked = True
         txtcAppl(1).Locked = True
         txtcAppl(2).Locked = True
         txtcAppl(3).Locked = True
         txtcAppl(4).Locked = True
         txtcAppl(5).Locked = True
         txtcAppl(6).Locked = True
         txtcAppl(7).Locked = True
         txtcAppl(8).Locked = True
         txtcAppl(9).Locked = True
         txteAppl(0).Locked = True
         txteAppl(1).Locked = True
         txteAppl(2).Locked = True
         txteAppl(3).Locked = True
         txteAppl(4).Locked = True
         txteAppl(5).Locked = True
         txteAppl(6).Locked = True
         txteAppl(7).Locked = True
         txteAppl(8).Locked = True
         txteAppl(9).Locked = True
         txtTPG39.Locked = True
         txtTPG40.Locked = True
         txtTPG41.Locked = True
         txtTPG42.Locked = True
         '2018/11/12 END
         Combo1.Locked = True: textNA01.Locked = True 'Add By Sindy 2019/9/4
      Case Else:
         text02.Locked = False
         text03.Locked = False
         text04.Locked = False
         text05.Locked = False
         text06_1.Locked = False
         text07_1.Locked = False
         Text2.Locked = False
         'Add By Sindy 2013/10/29
         txtTPG15.Locked = False
         txtTPG16.Locked = False
         txtTPG17.Locked = False
         '2013/10/29 END
         'Add By Sindy 2018/11/12
         txtTPG18.Locked = False
         txtcAppl(0).Locked = False
         txtcAppl(1).Locked = False
         txtcAppl(2).Locked = False
         txtcAppl(3).Locked = False
         txtcAppl(4).Locked = False
         txtcAppl(5).Locked = False
         txtcAppl(6).Locked = False
         txtcAppl(7).Locked = False
         txtcAppl(8).Locked = False
         txtcAppl(9).Locked = False
         txteAppl(0).Locked = False
         txteAppl(1).Locked = False
         txteAppl(2).Locked = False
         txteAppl(3).Locked = False
         txteAppl(4).Locked = False
         txteAppl(5).Locked = False
         txteAppl(6).Locked = False
         txteAppl(7).Locked = False
         txteAppl(8).Locked = False
         txteAppl(9).Locked = False
         txtTPG39.Locked = False
         txtTPG40.Locked = False
         txtTPG41.Locked = False
         txtTPG42.Locked = False
         '2018/11/12 END
         Combo1.Locked = False: textNA01.Locked = False 'Add By Sindy 2019/9/4
   End Select
   text01.BackColor = &H8000000F
   text06_2.BackColor = &H8000000F
   text08.BackColor = &H8000000F
   text09.BackColor = &H8000000F
End Sub

'Add By Sindy 2019/9/26
Private Sub Combo1_Click()
   If m_EditMode = 1 Or m_EditMode = 2 Then
      If Trim(Combo1.Text) <> "" Then
         textNA01 = Left(Trim(Combo1.Text), 3)
      End If
   End If
End Sub

'910709 Sieg 412
Private Sub Command1_Click(Index As Integer)
 Dim i As Integer
 Dim nNo As Long
   If CheckDataValid() = True Then
    If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
      Select Case m_EditMode
         Case 0:
            nNo = Val(text02.Text)
            nNo = nNo + 1
            m_CurrTPG02 = CStr(nNo)
            m_CurrTPG03 = text03
            m_CurrTPG04 = text04
            m_CurrTPG05 = text05
      End Select
      frm04060301_1.UpdateRecord m_DataKey
   Else
      Exit Sub
   End If
 
   If Index = 0 Then
      '上一筆
      i = frm04060301_1.GrdList.row - 1
      If i > 0 Then
         frm04060301_1.GrdList.row = frm04060301_1.GrdList.row - 1
         frm04060301_1.textQuery = frm04060301_1.GrdList.TextMatrix(i, 1)
      Else
         MsgBox "已是第一筆了 !", vbInformation
      End If
      
   Else
      i = frm04060301_1.GrdList.row + 1
      If i < frm04060301_1.GrdList.Rows Then
         frm04060301_1.GrdList.row = frm04060301_1.GrdList.row + 1
         frm04060301_1.textQuery = frm04060301_1.GrdList.TextMatrix(i, 1)
      Else
         MsgBox "已是最後一筆了 !", vbInformation
      End If
   End If
   Select Case m_EditMode
      Case "1"
         frm04060301_1.buttonMod_Click
        If Me.text02.Enabled Then Me.text02.SetFocus
      Case "2"
         frm04060301_1.buttonQuery_Click
      Case "3"
         frm04060301_1.buttonDel_Click
   End Select
End Sub

Private Sub Form_Activate()
    If Me.Text1.Visible Then
        Me.Text1.SetFocus
    End If
End Sub

Private Sub Form_Load()
Dim rsTmp As New ADODB.Recordset
   
   SSTab1.Tab = 0 'Add By Sindy 2018/11/12
   
   'Add By Sindy 2019/9/25
   Combo1.Clear
   strSql = "SELECT na01,na03 FROM NATION" & _
            " WHERE LENGTH(na01)=3 AND substr(na01,1,1)='A' AND substr(na01,1,2)<>'A9'" & _
            " group by na01,na03 order by na01 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         Combo1.AddItem rsTmp.Fields("na01") & " " & rsTmp.Fields("na03")
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   strSql = "SELECT na01,na03 FROM NATION" & _
            " WHERE LENGTH(na01)=3 AND substr(na01,1,1)<>'A' AND substr(na01,1,1)<>'B'" & _
            " and na01>'010'" & _
            " GROUP BY na01,na03" & _
            " order by na01 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         Combo1.AddItem rsTmp.Fields("na01") & " " & rsTmp.Fields("na03")
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   '2019/9/25 END
   
   Set rsTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '隱藏新申請案號欄位
    Me.Text1.Visible = False
    Me.Label12.Visible = False
    frm04060301_1.Show
    Set frm04060301_2 = Nothing
End Sub

Private Sub text02_Validate(Cancel As Boolean)
   Dim nLen As Integer
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(text02) = False Then
'      nLen = Len(text02)
'      text02 = String(9 - nLen, "0") & text02
      If IsTPG02Exist(text02) = True Then
         Cancel = True
         strTit = "公開號"
         strMsg = "公開號已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text02_GotFocus
      End If
   End If
End Sub

Private Sub text03_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(text03) = False Then
      If CheckIsTaiwanDate(text03, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的公開日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
         text03_GotFocus
         GoTo EXITSUB
      End If
        '公開日不能大於系統日
      If DBDATE(text03) > strSrvDate(1) Then
         Cancel = True
         strMsg = "公開日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
         text03_GotFocus
      End If
   End If
EXITSUB:
End Sub

'910709 Sieg 412
Private Sub text04_LostFocus()
   If text04 <> "" Then
      If Not ChkText04 Then text04.SetFocus: text04_GotFocus
   End If
End Sub

'910709 Sieg 412
Private Function ChkText04() As Boolean
 Dim strTmp As String
   ChkText04 = True
   If Len(text03) = 6 Then
      strTmp = Left(text03, 2)
   Else
      strTmp = Left(text03, 3)
   End If
   If Val(text04) + 91 <> Val(strTmp) Then
      MsgBox "公開日期與公報卷期不符，請重新輸入 !", vbCritical
      ChkText04 = False
   End If
End Function

'910709 Sieg 412
Private Function ChkText05() As Boolean
 Dim strTmp As String
 Dim i As Integer, j As Integer
   ChkText05 = True
   If Len(text03) = 6 Then
      j = Val(Mid(text03, 3, 2))
   Else
      j = Val(Mid(text03, 4, 2))
   End If
   i = (j - 1) * 2
   j = Val(Right(text03, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   End If
    'Modify By Cheng 2004/01/02
'   If Val(text05) <> i - 8 Then
   'Modify by Morgan 2005/9/23
   '92年公報從5月開始
   If Val(text03) < 930000 Then i = i - 8
   
   If Val(text05) <> i Then
    'End
      MsgBox "公開日期與公報卷期不符，請重新輸入 !", vbCritical
      ChkText05 = False
   End If
End Function

'910709 Sieg 412
Private Sub text05_LostFocus()
   If text05 <> "" Then
      If Not ChkText05 Then text05.SetFocus: text05_GotFocus
   End If
End Sub

Private Sub text06_1_Validate(Cancel As Boolean)
Dim strMsg As String
Dim strTit As String
Dim nResponse
    Cancel = False
    If IsEmptyText(text06_1) = False Then
        'Add By Cheng 2003/05/23
        If Me.text06_1.Text = "000" Then
            Cancel = True
            MsgBox "申請人國籍不可輸入 000 !!!"
            text06_1_GotFocus
        'Add By Sindy 2011/12/21
        ElseIf Me.text06_1.Text = "003" Then
            Cancel = True
            MsgBox "申請人國籍不可輸入 003 !!!"
            text06_1_GotFocus
        ElseIf UpdateCtrlData(1) = False Then
            Cancel = True
            strMsg = "無此國籍資料"
            strTit = "錯誤"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text06_1_GotFocus
        End If
    Else
        text06_2 = Empty
    End If
End Sub

Private Sub text07_1_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
    '欄位資料檢查有誤
    Cancel = False
    '若有輸入代理人代號
    If IsEmptyText(text07_1) = False Then
       If UpdateCtrlData(2) = False Then
            Cancel = True
            strMsg = "無此代理人資料"
            strTit = "錯誤"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            '欄位資料檢查有誤
            Cancel = True
            text07_1.SetFocus
            text07_1_GotFocus
       End If
    '若沒輸入代理人代號
    Else
       text08 = Empty
       text07_2 = Empty
    End If
End Sub

Public Sub UpdateData()
   ' 先清除欄位內容
   Clear
   ' 更新 Caption
   Dim strCap As String
   Dim strTmp As String
   strCap = "專利公報資料維護"
   
   Command1(0).Visible = True
   Command1(1).Visible = True
   Select Case m_EditMode
      Case 0:
         strTmp = " -- 新增"
         text02 = m_CurrTPG02
         text03 = m_CurrTPG03
         Command1(0).Visible = False
         Command1(1).Visible = False
      Case 1:
         strTmp = " -- 修改"
         text02 = ""
         text03 = ""
      Case 2:
         strTmp = " -- 查詢"
      Case 3:
         strTmp = " -- 刪除"
   End Select
   Caption = strCap & strTmp
   ' 更新第一個欄位
   text01 = m_DataKey
   MoveFormToCenter Me
   ' 更新內容
   UpdateCtrlData (0)
   UpdateState
   
   If m_EditMode = 0 Then
      SetPrevState
   End If
End Sub

Public Function CheckDataValid() As Boolean
   Dim strMsg As String
   Dim strTit As String
   Dim nRet As Boolean
   Dim nResponse
   Dim nLen As Integer
   Dim strFreeAgentCode As String
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
    Dim StrSQLa As String
    Dim rsA As New ADODB.Recordset
   
   CheckDataValid = False
   
   'Added by Morgan 2021/12/24 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   If PUB_ChkUniText(Me, , True, "ComboBox") = False Then
       Exit Function
   End If
   'end 2021/12/24
   
    '若為修改申請案號
    If m_EditMode = 1 And Me.Text1.Visible = True Then
        '若未輸入新申請案號
        If Me.Text1.Text = "" Then
            strMsg = "請輸入新申請案號"
            strTit = "資料檢核"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text1.SetFocus
            Text1_GotFocus
            GoTo EXITSUB
        End If
        '若有輸入新的申請案號
        If Me.Text1.Text <> "" Then
            StrSQLa = "Select * From TPGAZETTE Where TPG01='" & ChgSQL(Me.Text1.Text) & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                MsgBox "您輸入的新申請案號已存在, 請重新輸入!!!", vbExclamation + vbOKOnly
                Text1.SetFocus
                Text1_GotFocus
                If rsA.State <> adStateClosed Then rsA.Close
                Set rsA = Nothing
                GoTo EXITSUB
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
        End If
    End If
 
   If IsEmptyText(text02) = True Then
        strMsg = "請輸入公開號"
        strTit = "資料檢核"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        text02.SetFocus
        text02_GotFocus
        GoTo EXITSUB
   Else
'      nLen = Len(text02)
'      text02 = String(6 - nLen, "0") & text02
        If IsTPG02Exist(text02) = True Then
            strTit = "公開號"
            strMsg = "公開號已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text02.SetFocus
            text02_GotFocus
            GoTo EXITSUB
        End If
   End If
   
   If IsEmptyText(text03) = True Then
        strMsg = "請輸入公開日"
        strTit = "資料檢核"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        text03.SetFocus
        text03_GotFocus
        GoTo EXITSUB
   Else
      If CheckIsTaiwanDate(text03, False) = False Then
         strMsg = "請輸入正確的公開日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
        text03_GotFocus
         GoTo EXITSUB
      End If
        '公開日不能大於系統日
      If DBDATE(text03) > strSrvDate(1) Then
        '公開日不能大於系統日
         strMsg = "公開日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
         text03_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   If IsEmptyText(text04) = True Then
        strMsg = "請輸入公開卷期"
        strTit = "資料檢核"
        nResponse = MsgBox(strMsg, vbOKOnly, strTit)
        text04.SetFocus
        text04_GotFocus
        GoTo EXITSUB
   Else
        '910709 Sieg 412
        If Not ChkText04 Then
            text04.SetFocus
            text04_GotFocus
            GoTo EXITSUB
        End If
   End If
   
   If IsEmptyText(text05) = True Then
      strMsg = "請輸入公開卷期"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text05.SetFocus
      text05_GotFocus
      GoTo EXITSUB
   Else
      
      '910709 Sieg 412
      If Not ChkText05 Then
        text05.SetFocus
        text05_GotFocus
        GoTo EXITSUB
      End If
   End If
   
   If IsEmptyText(text06_1) = True Then
      strMsg = "請輸入申請人國籍"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text06_1.SetFocus
      text06_1_GotFocus
      GoTo EXITSUB
   Else
      If UpdateCtrlData(1) = False Then
         strMsg = "申請人國籍不正確"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text06_1.SetFocus
         text06_1_GotFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2019/9/4
   ' 地區編號
   If IsEmptyText(Combo1.Text) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入地區"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      Combo1.SetFocus
      GoTo EXITSUB
   End If
   
   ' 代理人
   If IsEmptyText(text07_1) = False Then
      If UpdateCtrlData(2) = False Then
         strMsg = "無此代理人資料"
         strTit = "錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text07_1.SetFocus
         text07_1_GotFocus
         GoTo EXITSUB
      '92.11.5 add by sonia
      Else
         If text07_1 = "01" And text09 = "" Then
            strMsg = "代理人為本所, 但無本所案件資料"
            strTit = "錯誤"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text07_1.SetFocus
            text07_1_GotFocus
            GoTo EXITSUB
         End If
      '92.11.5 end
      End If
   End If
   
   If IsEmptyText(text07_1) = True And IsEmptyText(text07_2) = False Then
      Set rsTmp = New ADODB.Recordset
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA01 = 'P' AND " & _
                     "TA03 = '" & Trim(text07_2) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("TA02")) = False Then
            text07_1 = rsTmp.Fields("TA02")
         End If
         If IsNull(rsTmp.Fields("TA04")) = False Then
            text08 = rsTmp.Fields("TA04")
         End If
      Else
      
        On Error GoTo ErrorHandler
        cnnConnection.BeginTrans
         strTit = "代理人"
         'Modif by Morgan 2011/5/17
         'strFreeAgentCode = GetFreeAgentCode
         strFreeAgentCode = PUB_GetFreeAgentCode("P")
         strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                        "　　　　　代理人名稱 <" & text07_2 & "> "
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
         If nResponse = vbYes Then
            strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('P','" & strFreeAgentCode & "','" & text07_2 & "','" & text07_2 & "')"
            cnnConnection.Execute strSql
            text07_1 = strFreeAgentCode
            Me.text08.Text = "" & Me.text07_2.Text
            ' 儲存公開日
            If IsEmpty(text03) = False Then
               strSql = "UPDATE TAgent SET TA05 = " & DBDATE(text03) & " " & _
                        "WHERE TA01 = 'P' AND " & _
                              "TA02 = '" & strFreeAgentCode & "' "
               cnnConnection.Execute strSql
            End If
            cnnConnection.CommitTrans
         '若不新增
         Else
            cnnConnection.RollbackTrans
            rsTmp.Close
            Set rsTmp = Nothing
            text07_2.SetFocus
            text07_2_GotFocus
            GoTo EXITSUB
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If

   ' 申請人國籍為外國時, 代理人一定要輸入
   If text06_1 > "010" Then
      If IsEmptyText(text07_1) = True Then
         strMsg = "申請人國籍為外國, 代理人一定要輸入"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text07_1.SetFocus
            text07_1_GotFocus
         GoTo EXITSUB
      End If
   End If
    'Add By Cheng 2003/05/16
    '檢查申請實體審查
    If Me.Text2.Text = "" Then
         strMsg = "請輸入有無申請實體審查"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text2.SetFocus
         Text2_GotFocus
         GoTo EXITSUB
    End If
    '若為本所案件
    If Me.text09.Text <> "" Then
        StrSQLa = "Select * From CaseProgress Where " & ChgCaseprogress(Replace(Me.text09.Text, "-", "")) & " And CP10='416' And CP27 Is Not Null "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        '若實體審查已發文
        If rsA.RecordCount > 0 And Me.Text2.Text = "N" Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            strMsg = "此案件已提實審，請重新輸入"
            strTit = "資料檢核"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            Text2.SetFocus
            Text2_GotFocus
            GoTo EXITSUB
        '若無實體審查或實體審查未發文
        ElseIf rsA.RecordCount <= 0 And Me.Text2.Text = "Y" Then
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            strTit = "資料檢核"
            'Modify by Morgan 2004/11/10
            'strMsg = "此案件未提實審，請重新輸入"
            'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            strMsg = "此案件未提實審，請依公報資料先行輸入，並通知專業部確認資料是否正確!!"
            nResponse = MsgBox(strMsg, vbYesNo + vbDefaultButton2, strTit)
            If nResponse = vbNo Then
            '2004/11/10 end
               Text2.SetFocus
               Text2_GotFocus
               GoTo EXITSUB
            End If
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    End If
    
   'Add By Sindy 2013/10/29 IPC分類資料開始於101年01月
   If Val(DBDATE(text03)) >= 20120101 Then
      If IsEmptyText(txtTPG15) = True Then
         strMsg = "請輸入國際分類號"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG15.SetFocus
         txtTPG15_GotFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtTPG16) = True Then
         strMsg = "請輸入IPC分類"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG16.SetFocus
         txtTPG16_GotFocus
         GoTo EXITSUB
      End If
      If IsEmptyText(txtTPG17) = True Then
         strMsg = "請輸入產業別分類 "
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG17.SetFocus
         txtTPG17_GotFocus
         GoTo EXITSUB
      End If
   End If
   '2013/10/29 END
   
   CheckDataValid = True
EXITSUB:
Exit Function
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "(" & Err.NUMBER & ")" & Err.Description
End Function

Public Sub Clear()
   text01 = Empty
   text02 = Empty
   text03 = Empty
   text04 = Empty
   text05 = Empty
   text06_1 = Empty
   text06_2 = Empty
   text07_1 = Empty
   text07_2 = Empty
   text08 = Empty
   text09 = Empty
   Text2.Text = Empty
   'Add By Sindy 2013/10/29
   txtTPG15.Text = Empty
   txtTPG16.Text = Empty
   txtTPG17.Text = Empty
   '2013/10/29 END
   'Add By Sindy 2018/11/12
   txtTPG18.Text = Empty
   txtcAppl(0).Text = ""
   txtcAppl(1).Text = ""
   txtcAppl(2).Text = ""
   txtcAppl(3).Text = ""
   txtcAppl(4).Text = ""
   txtcAppl(5).Text = ""
   txtcAppl(6).Text = ""
   txtcAppl(7).Text = ""
   txtcAppl(8).Text = ""
   txtcAppl(9).Text = ""
   txteAppl(0).Text = ""
   txteAppl(1).Text = ""
   txteAppl(2).Text = ""
   txteAppl(3).Text = ""
   txteAppl(4).Text = ""
   txteAppl(5).Text = ""
   txteAppl(6).Text = ""
   txteAppl(7).Text = ""
   txteAppl(8).Text = ""
   txteAppl(9).Text = ""
   txtTPG39.Text = ""
   txtTPG40.Text = ""
   txtTPG41.Text = ""
   txtTPG42.Text = ""
   '2018/11/12 END
   Combo1.Text = "": textNA01.Text = "" 'Add By Sindy 2019/9/4
End Sub

Private Sub SetPrevState()
   text02 = m_CurrTPG02
   text03 = m_CurrTPG03
   text04 = m_CurrTPG04
   text05 = m_CurrTPG05
   If text02 = Empty Then
      text02.SetFocus
   ElseIf text03 = Empty Then
      text03.SetFocus
   ElseIf text04 = Empty Then
      text04.SetFocus
   ElseIf text05 = Empty Then
      text05.SetFocus
   Else
      If text06_1.Locked = False Then
         text06_1.SetFocus
      End If
   End If
End Sub

Public Function IsEmpty(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmpty = False
   
   If Len(strData) <= 0 Then
      IsEmpty = True
   Else
      IsEmpty = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmpty = False
            Exit For
         End If
      Next nIndex
   End If
End Function

Private Sub text07_2_Validate(Cancel As Boolean)
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strFreeAgentCode As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ' 代理人是鍵入名稱時
   If IsEmpty(text07_1) = True And IsEmpty(text07_2) = False Then
      strSql = "SELECT * FROM TAGENT " & _
               "WHERE TA01 = 'P' AND " & _
                     "TA03 = '" & Trim(text07_2) & "' "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If IsNull(rsTmp.Fields("TA02")) = False Then
            text07_1 = rsTmp.Fields("TA02")
         End If
         If IsNull(rsTmp.Fields("TA04")) = False Then
            text08 = rsTmp.Fields("TA04")
         End If
      Else
        On Error GoTo ErrorHandler
        cnnConnection.BeginTrans
         strTit = "代理人"
        'Modif by Morgan 2011/5/17
         'strFreeAgentCode = GetFreeAgentCode
         strFreeAgentCode = PUB_GetFreeAgentCode("P")
         strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                        "　　　　　代理人名稱 <" & text07_2 & "> "
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
         If nResponse = vbYes Then
            strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('P','" & strFreeAgentCode & "','" & text07_2 & "','" & text07_2 & "')"
            cnnConnection.Execute strSql
            text07_1 = strFreeAgentCode
            Me.text08.Text = "" & Me.text07_2.Text
            ' 儲存公開日
            If IsEmpty(text03) = False Then
               strSql = "UPDATE TAgent SET TA05 = " & DBDATE(text03) & " " & _
                        "WHERE TA01 = 'P' AND " & _
                              "TA02 = '" & strFreeAgentCode & "' "
               cnnConnection.Execute strSql
            End If
            cnnConnection.CommitTrans
         '若不新增
         Else
            Cancel = True
            cnnConnection.RollbackTrans
         End If
      End If
      rsTmp.Close
   End If
   Set rsTmp = Nothing

EXITSUB:
    'edit by nickc 2007/07/11 切換輸入法改用API
    'If Cancel = False Then: text07_2.IMEMode = 2
    If Cancel = False Then CloseIme
    If Cancel = True Then text07_2.SetFocus: text07_2_GotFocus
Exit Sub
ErrorHandler:
    cnnConnection.RollbackTrans
    MsgBox "(" & Err.NUMBER & ")" & Err.Description
End Sub

' 將所有的文字反白
Private Sub InverseAll(ByRef tb As Object)
   tb.SelStart = 0
   tb.SelLength = Len(tb.Text)
End Sub

Private Sub text02_KeyPress(KeyAscii As Integer)
   KeyAscii = UCase(KeyAscii)
End Sub

Private Sub text02_GotFocus()
    If m_EditMode = 1 Then
        '最後一個字反白
        If Me.text02.Text <> "" Then
            Me.text02.SelStart = Len(Me.text02.Text) - 1
            Me.text02.SelLength = 1
        End If
    Else
        InverseAll text02
    End If
End Sub

Private Sub text03_GotFocus()
   InverseAll text03
End Sub

Private Sub text04_GotFocus()
   InverseAll text04
End Sub

Private Sub text05_GotFocus()
   InverseAll text05
End Sub

Private Sub text06_1_GotFocus()
   InverseAll text06_1
End Sub

Private Sub text07_1_GotFocus()
   InverseAll text07_1
End Sub

Private Sub text07_2_GotFocus()
   InverseAll text07_2
   'edit by nickc 2007/07/11 切換輸入法改用API
   'text07_2.IMEMode = 1
   OpenIme
End Sub

Private Sub Text1_GotFocus()
    TextInverse Me.Text1
End Sub

'Add by Morgan 2010/12/28
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = Asc(" ") Or KeyAscii = Asc("-") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Text2_GotFocus()
    'Add By Cheng 2003/05/16
    TextInverse Me.Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/05/16
    KeyAscii = UpperCase(KeyAscii)
    Select Case KeyAscii
    Case "78", "89", "8"
        KeyAscii = KeyAscii
    Case Else
        KeyAscii = 0
    End Select
End Sub

'Add By Sindy 2013/10/29
Private Sub txtTPG15_GotFocus()
   TextInverse Me.txtTPG15
End Sub
Private Sub txtTPG15_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtTPG16_GotFocus()
   TextInverse Me.txtTPG16
End Sub
Private Sub txtTPG17_GotFocus()
   TextInverse Me.txtTPG17
End Sub

'Add By Sindy 2018/11/12
Private Sub txtTPG18_GotFocus()
   TextInverse Me.txtTPG18
End Sub
Private Sub txtTPG39_GotFocus()
   InverseAll txtTPG39
End Sub
Private Sub txtTPG39_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtTPG39) = False Then
      If CheckIsTaiwanDate(txtTPG39, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的申請日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG39_GotFocus
         GoTo EXITSUB
      End If
      If DBDATE(txtTPG39) > DBDATE(SystemDate()) Then
         Cancel = True
         strMsg = "申請日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG39_GotFocus
      End If
   End If
EXITSUB:
End Sub
Private Sub txtTPG40_GotFocus()
   InverseAll txtTPG40
End Sub
Private Sub txtTPG40_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtTPG40) = False Then
      If CheckIsTaiwanDate(txtTPG40, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的最早優先權日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG40_GotFocus
         GoTo EXITSUB
      End If
      If DBDATE(txtTPG40) > DBDATE(SystemDate()) Then
         Cancel = True
         strMsg = "最早優先權日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPG40_GotFocus
      End If
   End If
EXITSUB:
End Sub
Private Sub txteAppl_GotFocus(Index As Integer)
   InverseTextBox txteAppl(Index)
End Sub
Private Sub txtcAppl_GotFocus(Index As Integer)
   '切換輸入法改用API
   OpenIme
   TextInverse txtcAppl(Index)
End Sub
Private Sub txtcAppl_Validate(Index As Integer, Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 0 And m_EditMode <> 1 Then Exit Sub
   If txtcAppl(Index).Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(txtcAppl(Index), txtcAppl(Index).MaxLength - 1) Then
      Cancel = True
   End If
End Sub
'2018/11/12 END

' 取得國家的代碼
Private Function GetNationNo(ByVal strData As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   GetNationNo = Empty
   'Modify By Sindy 2013/8/19 + AND length(na01)=3
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA03 = '" & strData & "' AND length(na01)=3 "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("NA01")) = False Then
         GetNationNo = rsTmp.Fields("NA01")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/9/4
Private Sub textNA01_Change()
   If IsEmptyText(textNA01) = False Then
      Combo1.TabStop = False
   Else
      Combo1.TabStop = True
   End If
End Sub
Private Sub textNA01_GotFocus()
   InverseTextBox textNA01
End Sub
Private Sub textNA01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
' 地區別
Private Sub textNA01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strTemp As String
   
   If IsEmptyText(textNA01) = False Then
      Combo1.TabStop = False
      If textNA01 <= "010" Then
         Combo1.Text = Empty
         strTit = "檢核資料"
         strMsg = "地區別不正確"
         Cancel = True
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNA01.SetFocus
      End If
      strTemp = GetNationName(textNA01, 0)
      If IsEmptyText(strTemp) = False Then
         Combo1.Text = textNA01 & " " & strTemp
      Else
         Select Case m_EditMode
            Case 1, 2:
               Combo1.Text = Empty
               strTit = "檢核資料"
               strMsg = "地區名稱不存在"
               Cancel = True
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textNA01.SetFocus
         End Select
      End If
   Else
      Combo1.TabStop = True
   End If
End Sub
